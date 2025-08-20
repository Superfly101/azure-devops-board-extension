import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";

import "./work-hub-group.scss";

import { Header, TitleSize } from "azure-devops-ui/Header";
import { Page } from "azure-devops-ui/Page";
import { Card } from "azure-devops-ui/Card";

import { showRootComponent } from "../../Common";
import {
  CommonServiceIds,
  IProjectPageService,
  getClient,
} from "azure-devops-extension-api";
import { WorkItemTrackingRestClient } from "azure-devops-extension-api/WorkItemTracking";

import { ObservableValue } from "azure-devops-ui/Core/Observable";
import { FormItem } from "azure-devops-ui/FormItem";
import { TextField, TextFieldWidth } from "azure-devops-ui/TextField";
import { AsyncEpicTagPicker } from "./epic-tag-picker";
import { Button } from "azure-devops-ui/Button";
import { ButtonGroup } from "azure-devops-ui/ButtonGroup";
import PfoBoardDropdown from "./pfo-board-dropdown";
import { IListBoxItem } from "azure-devops-ui/ListBox";
import { JsonPatchDocument } from "azure-devops-extension-api/WebApi";
import { MessageCard, MessageCardSeverity } from "azure-devops-ui/MessageCard";

const projectIdObservable = new ObservableValue<string | undefined>("");

interface IWorkHubGroup {
  projectContext: any;
  assignedWorkItems: any[];
  validationErrors: {
    projectId: boolean;
    epics: boolean;
    boards: boolean;
  };
  isCreating: boolean;
  currentProgress: number;
  totalOperations: number;
  toastMessage: { message: string; severity: MessageCardSeverity } | null;
}

interface TagItem {
  id: number;
  text: string;
}

class WorkHubGroup extends React.Component<{}, IWorkHubGroup> {
  private selectedEpics: TagItem[] = [];
  private selectedBoards: Array<IListBoxItem<{ areaPath: string }>> = [];

  constructor(props: {}) {
    super(props);
    this.state = { 
      projectContext: undefined, 
      assignedWorkItems: [],
      validationErrors: {
        projectId: false,
        epics: false,
        boards: false,
      },
      isCreating: false,
      currentProgress: 0,
      totalOperations: 0,
      toastMessage: null,
    };
  }

  public componentDidMount() {
    try {
      console.log("Component did mount, initializing SDK...");
      SDK.init();

      SDK.ready()
        .then(() => {
          console.log("SDK is ready, loading project context...");
          this.loadProjectContext();
        })
        .catch((error) => {
          console.error("SDK ready failed: ", error);
        });
    } catch (error) {
      console.error(
        "Error during SDK initialization or project context loading: ",
        error
      );
    }
  }

  public render(): JSX.Element {
    return (
      <Page className="dependency-epic-hub flex-grow">
        <Header
          title="Create Dependency Epic"
          titleSize={TitleSize.Large}
          description="Configure and create dependency epics across multiple PFO boards"
        />

        <div className="page-content">
          {this.state.toastMessage && (
            <div style={{ marginBottom: "16px", maxWidth: "800px", margin: "0 auto 16px auto" }}>
              <MessageCard
                severity={this.state.toastMessage.severity}
                onDismiss={() => this.setState({ toastMessage: null })}
              >
                {this.state.toastMessage.message}
              </MessageCard>
            </div>
          )}
          
          <div className="form-container">
            <Card className="form-card">
              <div className="form-content">
                <div className="form-header">
                  <h3 className="form-title">Epic Configuration</h3>
                  <p className="form-description">
                    Fill in the required information to create a new dependency
                    epic
                  </p>
                </div>

                <div className="form-fields">
                  <FormItem
                    label="Project ID"
                    message={this.state.validationErrors.projectId ? "Project ID is required" : "Enter the target project identifier"}
                    error={this.state.validationErrors.projectId}
                    className="form-field"
                  >
                    <TextField
                      ariaLabel="Project ID input field"
                      placeholder="Enter project ID..."
                      value={projectIdObservable}
                      onChange={(e) => {
                        projectIdObservable.value = e.target.value;
                        // Clear error when user starts typing
                        if (this.state.validationErrors.projectId && e.target.value.trim()) {
                          this.setState({
                            validationErrors: {
                              ...this.state.validationErrors,
                              projectId: false
                            }
                          });
                        }
                      }}
                      width={TextFieldWidth.auto}
                      disabled={this.state.isCreating}
                    />
                  </FormItem>

                  <FormItem
                    label="Source of Truth Epic ID"
                    message={this.state.validationErrors.epics ? "At least one epic must be selected" : "Select the epic that will serve as the source of truth"}
                    error={this.state.validationErrors.epics}
                    className="form-field"
                  >
                    <AsyncEpicTagPicker onSelect={this.handleEpicSelection} />
                  </FormItem>

                  <FormItem
                    label="PFO Boards"
                    message={this.state.validationErrors.boards ? "At least one board must be selected" : "Choose the PFO boards where dependency epics will be created"}
                    error={this.state.validationErrors.boards}
                    className="form-field"
                  >
                    <PfoBoardDropdown onSelect={this.handleBoardSelection} />
                  </FormItem>
                </div>

                <div className="form-actions">
                  <ButtonGroup>
                    <Button
                      text={this.state.isCreating ? `Creating ${this.state.currentProgress} of ${this.state.totalOperations}...` : "Create Epic"}
                      primary={true}
                      iconProps={this.state.isCreating ? { iconName: "Spinner" } : { iconName: "Add" }}
                      onClick={() => this.handleSubmit()}
                      disabled={this.state.isCreating}
                    />
                  </ButtonGroup>
                </div>
              </div>
            </Card>
          </div>
        </div>
      </Page>
    );
  }

  private validateForm = (): boolean => {
    const projectIdValue = projectIdObservable.value?.trim();
    const hasProjectId = Boolean(projectIdValue);
    const hasEpics = this.selectedEpics.length > 0;
    const hasBoards = this.selectedBoards.length > 0;

    const validationErrors = {
      projectId: !hasProjectId,
      epics: !hasEpics,
      boards: !hasBoards,
    };

    this.setState({ validationErrors });

    return hasProjectId && hasEpics && hasBoards;
  };

  private handleSubmit = async () => {
    console.log("Validating form...");
    
    const isValid = this.validateForm();
    
    if (!isValid) {
      console.log("Form validation failed");
      return;
    }

    // Clear any previous toast messages
    this.setState({ toastMessage: null });

    console.log("Creating dependency epic...");
    console.log("###### Form Values ######");

    console.log("Project ID:", projectIdObservable.value);
    console.log("Selected Epics:", this.selectedEpics);
    console.log("Select Boards:", this.selectedBoards);

    await this.createDependcyEpics();
  };

  private createDependcyEpics = async () => {
    const witClient = getClient(WorkItemTrackingRestClient);
    const projectId = projectIdObservable.value?.trim();
    
    const totalOperations = this.selectedBoards.length * this.selectedEpics.length;
    let currentProgress = 0;
    let successCount = 0;
    let errorCount = 0;

    // Initialize progress state
    this.setState({
      isCreating: true,
      currentProgress: 0,
      totalOperations: totalOperations,
    });

    for (const board of this.selectedBoards) {
      for (const leadEpic of this.selectedEpics) {
        currentProgress++;
        
        // Update progress
        this.setState({
          currentProgress: currentProgress,
        });

        try {
          console.log(
            `Creating dependency epic for Lead Epic ${leadEpic.id} on team ${board.text}`
          );
         console.log("Area Path:", board.data?.areaPath || board)
          const dependencyEpicTitle = `DEPENDENCY: ${leadEpic.text}`;
          const organization = SDK.getHost().name;
          const baseUrl = `https://dev.azure.com/${organization}/${this.state.projectContext.id}/_apis/wit/workItems`;
          const patchOperations: JsonPatchDocument = [
            {
              op: "add",
              path: "/fields/System.Title",
              value: dependencyEpicTitle,
            },
            {
              op: "add",
              path: "/fields/System.AreaPath",
              value: board.data?.areaPath || board.text,
            },
            {
              op: "add",
              path: "/relations/-",
              value: {
                rel: "System.LinkTypes.Hierarchy-Reverse",
                url: `${baseUrl}/${projectId}`,
                attributes: {
                  name: "Parent",
                  comment: "Parent project relationship",
                },
              },
            },
            {
              op: "add",
              path: "/relations/-",
              value: {
                rel: "System.LinkTypes.Dependency-Forward",
                url: `${baseUrl}/${leadEpic.id}`,
                attributes: {
                  name: "Successor",
                  comment: "Dependency lead epic relationship",
                },
              },
            },
          ];

          console.log("Patch Operations:", patchOperations);

          const dependencyEpic = await witClient.createWorkItem(
            patchOperations,
            this.state.projectContext.id,
            "Epic"
          );
          console.log("Dependency Epic Created:", dependencyEpic);
          successCount++;
        } catch (error) {
          console.log(`Failed to create dependency epic for Lead Epic ${leadEpic.id} on team ${board.text}`);
          errorCount++;
        }
      }
    }

    // Show completion feedback
    this.showCompletionToast(successCount, errorCount, totalOperations);
    
    // Reset creation state
    this.setState({
      isCreating: false,
      currentProgress: 0,
      totalOperations: 0,
    });
  };

  private showCompletionToast = (successCount: number, errorCount: number, totalOperations: number) => {
    let message: string;
    let severity: MessageCardSeverity;

    if (errorCount === 0) {
      // All successful
      message = `Successfully created ${successCount} dependency epic${successCount !== 1 ? 's' : ''}`;
      severity = MessageCardSeverity.Info;
    } else if (successCount === 0) {
      // All failed
      message = `Failed to create dependency epics. Check console for details.`;
      severity = MessageCardSeverity.Error;
    } else {
      // Partial success
      message = `Created ${successCount} of ${totalOperations} dependency epics. ${errorCount} failed. Check console for details.`;
      severity = MessageCardSeverity.Warning;
    }

    this.setState({
      toastMessage: { message, severity }
    });
  };

  private async loadProjectContext(): Promise<void> {
    try {
      const client = await SDK.getService<IProjectPageService>(
        CommonServiceIds.ProjectPageService
      );
      const context = await client.getProject();

      this.setState({ projectContext: context });

      SDK.notifyLoadSucceeded();
    } catch (error) {
      console.error("Failed to load project context: ", error);
    }
  }

  private handleEpicSelection = (epics: TagItem[]) => {
    this.selectedEpics = epics;
    
    // Clear error when epics are selected
    if (this.state.validationErrors.epics && epics.length > 0) {
      this.setState({
        validationErrors: {
          ...this.state.validationErrors,
          epics: false
        }
      });
    }
  };

  private handleBoardSelection = (boards: Array<IListBoxItem<{ areaPath: string }>>) => {
    this.selectedBoards = boards;
    
    // Clear error when boards are selected
    if (this.state.validationErrors.boards && boards.length > 0) {
      this.setState({
        validationErrors: {
          ...this.state.validationErrors,
          boards: false
        }
      });
    }
  };
}

showRootComponent(<WorkHubGroup />);