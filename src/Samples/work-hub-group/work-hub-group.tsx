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
import {
  WorkItemTrackingRestClient,
  Wiql,
} from "azure-devops-extension-api/WorkItemTracking";

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
  selectedEpics: TagItem[];
  selectedBoards: Array<IListBoxItem<{ areaPath: string }>>;
  validationErrors: {
    projectId: boolean;
    epics: boolean;
    boards: boolean;
  };
  isCreating: boolean;
  currentProgress: number;
  totalOperations: number;
  feedback: { message: string; severity: MessageCardSeverity } | null;
}

interface TagItem {
  id: number;
  text: string;
}

class WorkHubGroup extends React.Component<{}, IWorkHubGroup> {

  constructor(props: {}) {
    super(props);
    this.state = {
      projectContext: undefined,
      assignedWorkItems: [],
      selectedEpics: [],
      selectedBoards: [],
      validationErrors: {
        projectId: false,
        epics: false,
        boards: false,
      },
      isCreating: false,
      currentProgress: 0,
      totalOperations: 0,
      feedback: null,
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
    const {
      validationErrors,
      feedback,
      isCreating,
      currentProgress,
      totalOperations,
      selectedEpics,
      selectedBoards,
    } = this.state;

    return (
      <Page className="dependency-epic-hub flex-grow">
        <Header
          title="Create Dependency Epic"
          titleSize={TitleSize.Large}
          description="Configure and create dependency epics across multiple PFO boards"
        />

        <div className="page-content">
          <div>
            {feedback && (
              <MessageCard
                className="flex-self-stretch"
                onDismiss={() => this.setState({ feedback: null })}
                severity={feedback?.severity}
              >
                {feedback?.message}
              </MessageCard>
            )}
          </div>
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
                    message={validationErrors.projectId ? "Project Id cannot be empty and must be an integer" : "Enter the Project Id"}
                    error={validationErrors.projectId}
                    className="form-field"
                  >
                    <TextField
                      ariaLabel="Project ID input field"
                      placeholder="Enter project ID..."
                      value={projectIdObservable}
                      onChange={(e) => {
                        if (e.target.value && validationErrors.projectId) {
                          this.setState({
                            validationErrors: {
                              ...validationErrors,
                              projectId: false,
                            },
                          });
                        }
                        projectIdObservable.value = e.target.value;
                      }}
                      width={TextFieldWidth.auto}
                    />
                  </FormItem>

                  <FormItem
                    label="Source of Truth Epic ID"
                    message={
                      validationErrors.epics
                        ? "Atleast one epic must be selected"
                        : "Select the epic that will serve as the source of truth"
                    }
                    error={validationErrors.epics}
                    className="form-field"
                  >
                    <AsyncEpicTagPicker 
                      selectedTags={selectedEpics}
                      onSelect={this.handleEpicSelection} 
                    />
                  </FormItem>

                  <FormItem
                    label="PFO Boards"
                    message={
                      validationErrors.boards
                        ? "Atleast one board must be selected"
                        : "Choose the PFO boards where dependency epics will be created"
                    }
                    error={validationErrors.boards}
                    className="form-field"
                  >
                    <PfoBoardDropdown 
                      selectedBoards={selectedBoards}
                      onSelect={this.handleBoardSelection} 
                    />
                  </FormItem>
                </div>

                <div className="form-actions">
                  <ButtonGroup>
                    <Button
                      text={`${
                        isCreating
                          ? `Creating ${currentProgress} of ${totalOperations}...`
                          : "Create Epic"
                      }`}
                      primary={true}
                      iconProps={
                        isCreating
                          ? { iconName: "Spinner" }
                          : { iconName: "Add" }
                      }
                      onClick={() => this.handleSubmit()}
                      disabled={isCreating}
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

  private resetForm = () => {
    this.setState({
      selectedEpics: [],
      selectedBoards: [],
      validationErrors: {
        projectId: false,
        epics: false,
        boards: false,
      },
    });
    
    projectIdObservable.value = "";
  };

  private validateForm = () => {
    const { selectedEpics, selectedBoards } = this.state;

    const projectIdValue = projectIdObservable.value?.trim();
    const hasProjectId = /^\d+$/.test(projectIdValue || "");
    const hasEpics = selectedEpics.length > 0;
    const hasBoards = selectedBoards.length > 0;

    const isValid = hasProjectId && hasEpics && hasBoards;

    this.setState({
      validationErrors: {
        projectId: !hasProjectId,
        epics: !hasEpics,
        boards: !hasBoards,
      },
    });
    return isValid;
  };

  private handleSubmit = async () => {
    const { selectedEpics, selectedBoards } = this.state;

    console.log("Creating dependency epic...");
    console.log("###### Form Values ######");

    console.log("Project ID:", projectIdObservable.value);

    console.log("Selected Epics:", selectedEpics);

    console.log("Select Boards:", selectedBoards);

    const formIsvalid = this.validateForm();
    if (!formIsvalid) {
      console.log("Form Validation Failed!");
      return;
    }

    const projectId = projectIdObservable.value?.trim();

    if (!projectId) {
      return;
    }

    const projectExists = await this.ValidateProjectExists(projectId);

    if (!projectExists) {
      return;
    }

    this.setState({ feedback: null });

    await this.createDependcyEpics();
  };

  private ValidateProjectExists = async (projectId: string) => {
    try {
      const client = getClient(WorkItemTrackingRestClient);

      const workItem = await client.getWorkItem(
        parseInt(projectId),
        undefined,
        ["System.Id", "System.Title", "System.WorkItemType"]
      );

      const workItemType = workItem.fields["System.WorkItemType"];

      if (workItemType !== "Project") {
        this.setState({
          feedback: {
            message: `Work item with ID ${projectId} exists but not of type "Project". Found type "${workItemType}".`,
            severity: MessageCardSeverity.Error,
          },
        });
        return false;
      }

      return true;
    } catch (error) {
      console.log("Error validating Project Id:", error);

      if (error.status === 404) {
        this.setState({
          feedback: {
            message: `Project with ID ${projectId} doest not exist. Please verify the Project ID and try again.`,
            severity: MessageCardSeverity.Error,
          },
        });
      } else {
        this.setState({
          feedback: {
            message: `Failed to validate Project ID ${projectId}. Please check your permissions and try again.`,
            severity: MessageCardSeverity.Error,
          },
        });
      }
      return false;
    }
  };

  private createDependcyEpics = async () => {
    const { selectedBoards, selectedEpics } = this.state;

    const witClient = getClient(WorkItemTrackingRestClient);
    const projectId = projectIdObservable.value?.trim();

    const totalOperations = selectedBoards.length * selectedEpics.length;
    let currentProgress = 0;
    let successCount = 0;
    let errorCount = 0;

    this.setState({
      isCreating: true,
      currentProgress: 0,
      totalOperations: totalOperations,
    });

    for (const board of selectedBoards) {
      for (const leadEpic of selectedEpics) {
        try {
          console.log(
            `Creating dependency epic for Lead Epic ${leadEpic.id} on team ${board.text}`
          );
          console.log("Area Path:", board.data?.areaPath || board);

          currentProgress += 1;

          this.setState({
            currentProgress: currentProgress,
          });

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

          successCount += 1;
        } catch (error) {
          console.log(
            `Failed to create dependency epic for Lead Epic ${leadEpic.id} on team ${board.text}`
          );
          errorCount += 1;
        }
      }
    }

    this.showCompletionFeedback(successCount, errorCount, totalOperations);

    this.setState({
      isCreating: false,
      currentProgress: 0,
      totalOperations: 0,
    });
  };

  private showCompletionFeedback = (
    successCount: number,
    errorCount: number,
    totalOperations: number
  ) => {
    let message, severity;

    if (errorCount === 0) {
      message = `Successfully created ${successCount} dependency epic${
        successCount > 1 ? "s" : ""
      }`;
      severity = MessageCardSeverity.Info;
      
      // Reset form only on complete success
      this.resetForm();
    } else if (successCount === 0) {
      message = "Failed to create dependency epics. Check console for details.";
      severity = MessageCardSeverity.Error;
    } else {
      message = `Created ${successCount} of ${totalOperations} dependency epics. ${errorCount} failed. Check console for details.`;
      severity = MessageCardSeverity.Warning;
    }

    this.setState({
      feedback: {
        message,
        severity,
      },
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
    const { validationErrors } = this.state;

    if (validationErrors.epics && epics.length > 0) {
      this.setState({
        validationErrors: {
          ...validationErrors,
          epics: false,
        },
      });
    }
    this.setState({ selectedEpics: epics });
  };

  private handleBoardSelection = (
    boards: Array<IListBoxItem<{ areaPath: string }>>
  ) => {
    const { validationErrors } = this.state;

    if (validationErrors.boards && boards.length > 0) {
      this.setState({
        validationErrors: {
          ...validationErrors,
          boards: false,
        },
      });
    }
    this.setState({ selectedBoards: boards });
  };
}

showRootComponent(<WorkHubGroup />);