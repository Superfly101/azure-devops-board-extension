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
  WorkItem,
} from "azure-devops-extension-api/WorkItemTracking";

import { ObservableValue } from "azure-devops-ui/Core/Observable";
import { FormItem } from "azure-devops-ui/FormItem";
import { TextField, TextFieldWidth } from "azure-devops-ui/TextField";
import { AsyncEpicTagPicker } from "./epic-tag-picker";
import { Button } from "azure-devops-ui/Button";
import { ButtonGroup } from "azure-devops-ui/ButtonGroup";
import { IListBoxItem } from "azure-devops-ui/ListBox";
import PfoBoardDropdown from "./multi-select-dropdown";
import { MessageCard, MessageCardSeverity } from "azure-devops-ui/MessageCard";
import {
  JsonPatchOperation,
  Operation,
  JsonPatchDocument,
} from "azure-devops-extension-api/WebApi";

const projectIdObservable = new ObservableValue<string | undefined>("");

interface IWorkHubGroup {
  projectContext: any;
  assignedWorkItems: any[];
  isCreating: boolean;
  lastOperationMessage: string;
  lastOperationSeverity: MessageCardSeverity;
  validationErrors: ValidationErrors;
  hasAttemptedSubmit: boolean;
}

interface ValidationErrors {
  projectId?: string;
  selectedEpics?: string;
  selectedPfoBoards?: string;
}

interface TagItem {
  id: number;
  text: string;
}

interface EpicCreationResult {
  sourceEpic: TagItem;
  targetTeam: IListBoxItem<{}>;
  success: boolean;
  workItem?: WorkItem;
  error?: string;
}

class WorkHubGroup extends React.Component<{}, IWorkHubGroup> {
  private selectedEpics: TagItem[] = [];
  private selectedPfoBoards: Array<IListBoxItem<{}>> = [];

  constructor(props: {}) {
    super(props);
    this.state = {
      projectContext: undefined,
      assignedWorkItems: [],
      isCreating: false,
      lastOperationMessage: "",
      lastOperationSeverity: MessageCardSeverity.Info,
      validationErrors: {},
      hasAttemptedSubmit: false,
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
          this.loadAssignedWorkItems();
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

                {this.state.lastOperationMessage && (
                  <MessageCard
                    className="example-error-message"
                    severity={this.state.lastOperationSeverity}
                  >
                    {this.state.lastOperationMessage}
                  </MessageCard>
                )}

                <div className="form-fields">
                  <FormItem
                    label="Project ID"
                    message={
                      this.state.validationErrors.projectId ||
                      "Enter the target project identifier"
                    }
                    error={!!this.state.validationErrors.projectId}
                    className="form-field"
                  >
                    <TextField
                      ariaLabel="Project ID input field"
                      placeholder="Enter project ID..."
                      value={projectIdObservable}
                      onChange={(e) => {
                        projectIdObservable.value = e.target.value;
                        this.clearValidationError("projectId");
                      }}
                      width={TextFieldWidth.auto}
                    />
                  </FormItem>

                  <FormItem
                    label="Source of Truth Epic ID"
                    message={
                      this.state.validationErrors.selectedEpics ||
                      "Select the epic that will serve as the source of truth"
                    }
                    error={!!this.state.validationErrors.selectedEpics}
                    className="form-field"
                  >
                    <AsyncEpicTagPicker
                      onSelectionChange={(tags) => {
                        this.selectedEpics = tags;
                        this.clearValidationError("selectedEpics");
                      }}
                    />
                  </FormItem>

                  <FormItem
                    label="PFO Boards"
                    message={
                      this.state.validationErrors.selectedPfoBoards ||
                      "Choose the PFO boards where dependency epics will be created"
                    }
                    error={!!this.state.validationErrors.selectedPfoBoards}
                    className="form-field"
                  >
                    <PfoBoardDropdown
                      onSelectionChange={(boards) => {
                        this.selectedPfoBoards = boards;
                        this.clearValidationError("selectedPfoBoards");
                      }}
                    />
                  </FormItem>
                </div>

                <div className="form-actions">
                  <ButtonGroup>
                    <Button
                      text={
                        this.state.isCreating
                          ? "Creating Epics..."
                          : "Create Epic"
                      }
                      primary={true}
                      iconProps={{
                        iconName: this.state.isCreating ? "Sync" : "Add",
                      }}
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

  private clearValidationError = (field: keyof ValidationErrors) => {
    if (this.state.hasAttemptedSubmit && this.state.validationErrors[field]) {
      this.setState({
        validationErrors: {
          ...this.state.validationErrors,
          [field]: undefined,
        },
      });
    }
  };

  private validateForm = (): ValidationErrors => {
    const errors: ValidationErrors = {};

    if (!projectIdObservable.value?.trim()) {
      errors.projectId = "Project ID is required";
    }

    if (this.selectedEpics.length === 0) {
      errors.selectedEpics = "At least one source epic must be selected";
    }

    if (this.selectedPfoBoards.length === 0) {
      errors.selectedPfoBoards = "At least one PFO board must be selected";
    }

    return errors;
  };

  private handleSubmit = async () => {
    this.setState({ hasAttemptedSubmit: true });

    const validationErrors = this.validateForm();

    if (Object.keys(validationErrors).length > 0) {
      this.setState({
        validationErrors,
        lastOperationMessage: "Please fix the validation errors above",
        lastOperationSeverity: MessageCardSeverity.Error,
      });
      return;
    }

    // Clear any previous validation errors
    this.setState({
      validationErrors: {},
      lastOperationMessage: "",
      lastOperationSeverity: MessageCardSeverity.Info,
    });

    console.log("=== Form Values ===");
    console.log("Project ID:", projectIdObservable.value);
    console.log("Selected Epic Tags:", this.selectedEpics);
    console.log("Selected PFO Boards:", this.selectedPfoBoards);
    console.log("===================");

    this.setState({
      isCreating: true,
    });

    try {
      const results = await this.createDependencyEpics();
      this.handleCreationResults(results);
    } catch (error) {
      console.error("Error creating dependency epics:", error);
      this.setState({
        isCreating: false,
        lastOperationMessage: `Failed to create dependency epics: ${
          error.message || error
        }`,
        lastOperationSeverity: MessageCardSeverity.Error,
      });
    }
  };

  private createDependencyEpics = async (): Promise<EpicCreationResult[]> => {
    const results: EpicCreationResult[] = [];
    const witClient = getClient(WorkItemTrackingRestClient);
    const projectId = projectIdObservable.value?.trim();

    // For each source epic, create dependency epics on each PFO board
    for (const sourceEpic of this.selectedEpics) {
      for (const pfoBoard of this.selectedPfoBoards) {
        try {
          console.log(
            `Creating dependency epic for source ${sourceEpic.id} on team ${pfoBoard.text}`
          );

          // Create the dependency epic work item
          const workItem = await this.createDependencyEpic(
            witClient,
            sourceEpic,
            pfoBoard,
            projectId!
          );

          results.push({
            sourceEpic,
            targetTeam: pfoBoard,
            success: true,
            workItem,
          });

          console.log(
            `Successfully created dependency epic ${workItem.id} for source ${sourceEpic.id}`
          );
        } catch (error) {
          console.error(
            `Failed to create dependency epic for source ${sourceEpic.id} on team ${pfoBoard.text}:`,
            error
          );
          results.push({
            sourceEpic,
            targetTeam: pfoBoard,
            success: false,
            error: error.message || error.toString(),
          });
        }
      }
    }

    return results;
  };

  private createDependencyEpic = async (
    witClient: WorkItemTrackingRestClient,
    sourceEpic: TagItem,
    targetTeam: IListBoxItem<{}>,
    projectId: string
  ): Promise<WorkItem> => {
    const dependencyTitle = `DEPENDENCY: ${sourceEpic.text}`;

    // Create the JSON patch operations for the new work item
    const patchOperations: JsonPatchDocument = [
      {
        op: "add",
        path: "/fields/System.Title",
        value: dependencyTitle,
      },
    ];

    // const patchOperations: JsonPatchDocument = [
    //   {
    //     op: "add",
    //     path: "/fields/System.Title",
    //     value: dependencyTitle,
    //   },
    //   // Set area path to the target team (if available)
    //   {
    //     op: "add",
    //     path: "/fields/System.AreaPath",
    //     value: `${this.state.projectContext?.name}\\${targetTeam.text}`,
    //   },
    //   // Set iteration path to current project
    //   {
    //     op: "add",
    //     path: "/fields/System.IterationPath",
    //     value: this.state.projectContext?.name || "",
    //   },
    // ];

    console.log("Patch Operations:", patchOperations);

    console.log("Project Context:", this.state.projectContext);

    // Create the work item
    const newWorkItem = await witClient.createWorkItem(
      patchOperations,
      this.state.projectContext.id,
      "Epic"
    );

    console.log("New WorkItem:", newWorkItem);
    // Now add the relationships (Parent and Successor links)
    await this.addWorkItemRelations(
      witClient,
      newWorkItem.id,
      sourceEpic.id,
      parseInt(projectId)
    );

    return newWorkItem;
  };

  private addWorkItemRelations = async (
    witClient: WorkItemTrackingRestClient,
    dependencyEpicId: number,
    sourceEpicId: number,
    parentProjectId: number
  ): Promise<void> => {
    // const relationOperations: JsonPatchDocument = [
    //   // Add SUCCESSOR link to source epic
    //   {
    //     op: "add",
    //     path: "/relations/-",
    //     value: {
    //       rel: "Microsoft.VSTS.Common.Successor",
    //       url: `${this.getWorkItemUrl(sourceEpicId)}`,
    //       attributes: {
    //         comment: "Dependency relationship to source of truth epic",
    //       },
    //     },
    //   },
    //   // Add PARENT link to project (if valid work item ID)
    //   {
    //     op: "add",
    //     path: "/relations/-",
    //     value: {
    //       rel: "System.LinkTypes.Hierarchy-Reverse",
    //       url: `${this.getWorkItemUrl(parentProjectId)}`,
    //       attributes: {
    //         comment: "Parent project relationship",
    //       },
    //     },
    //   },
    // ];

    const relationOperations: JsonPatchDocument = [
      // Add PARENT link to project (if valid work item ID)
      {
        op: "add",
        path: "/relations/-",
        value: {
          rel: "System.LinkTypes.Hierarchy-Reverse",
          url: `${this.getWorkItemUrl(parentProjectId)}`,
          attributes: {
            isLocked: false,
            name: "Parent",
            comment: "Parent project relationship",
          },
        },
      },
    ];

    // Update the work item with relations
    await witClient.updateWorkItem(
      relationOperations,
      dependencyEpicId,
      this.state.projectContext.id
    );
  };

  private getWorkItemUrl = (workItemId: number): string => {
    // Get the base URL for work items in this Azure DevOps organization
    const organization = SDK.getHost().name;
    return `https://dev.azure.com/${organization}/${this.state.projectContext.id}/_apis/wit/workItems/${workItemId}`;
  };

  private handleCreationResults = (results: EpicCreationResult[]) => {
    const successCount = results.filter((r) => r.success).length;
    const failCount = results.filter((r) => !r.success).length;

    let message: string;
    let severity: MessageCardSeverity;

    if (failCount === 0) {
      message = `Successfully created ${successCount} dependency epic${
        successCount !== 1 ? "s" : ""
      }!`;
      severity = MessageCardSeverity.Info;
    } else if (successCount === 0) {
      message = `Failed to create all ${failCount} dependency epics. Check console for details.`;
      severity = MessageCardSeverity.Error;
    } else {
      message = `Created ${successCount} dependency epics successfully, but ${failCount} failed. Check console for details.`;
      severity = MessageCardSeverity.Warning;
    }

    // Log detailed results
    console.log("=== Epic Creation Results ===");
    results.forEach((result, index) => {
      console.log(
        `${index + 1}. Source Epic: ${result.sourceEpic.text} -> Team: ${
          result.targetTeam.text
        }`
      );
      if (result.success) {
        console.log(`Success: Created Epic ID ${result.workItem?.id}`);
      } else {
        console.log(`Failed: ${result.error}`);
      }
    });
    console.log("==============================");

    this.setState({
      isCreating: false,
      lastOperationMessage: message,
      lastOperationSeverity: severity,
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

  private async loadAssignedWorkItems(): Promise<void> {
    try {
      const client = getClient(WorkItemTrackingRestClient);

      const wiqlQuery = {
        query: `
                    SELECT [System.Id], [System.Title]
                    FROM WorkItems
                    WHERE [System.AssignedTo] = @Me
                    ORDER BY [System.ChangedDate] DESC
                    `,
      };

      const queryResult = await client.queryByWiql(wiqlQuery);
      const workItemIds = queryResult.workItems.map((wi) => wi.id);

      if (workItemIds.length > 0) {
        const workItems = await client.getWorkItems(workItemIds);
        this.setState({ assignedWorkItems: workItems });
      } else {
        this.setState({ assignedWorkItems: [] });
      }
    } catch (error) {
      console.error("Failed to load assigned work items: ", error);
    }
  }
}

showRootComponent(<WorkHubGroup />);
