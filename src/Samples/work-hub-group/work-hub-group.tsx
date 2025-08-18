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

const projectIdObservable = new ObservableValue<string | undefined>("");

interface IWorkHubGroup {
  projectContext: any;
  assignedWorkItems: any[];
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
    this.state = { projectContext: undefined, assignedWorkItems: [] };
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
                    message="Enter the target project identifier"
                    error={false}
                    className="form-field"
                  >
                    <TextField
                      ariaLabel="Project ID input field"
                      placeholder="Enter project ID..."
                      value={projectIdObservable}
                      onChange={(e) =>
                        (projectIdObservable.value = e.target.value)
                      }
                      width={TextFieldWidth.auto}
                    />
                  </FormItem>

                  <FormItem
                    label="Source of Truth Epic ID"
                    message="Select the epic that will serve as the source of truth"
                    error={false}
                    className="form-field"
                  >
                    <AsyncEpicTagPicker onSelect={this.handleEpicSelection} />
                  </FormItem>

                  <FormItem
                    label="PFO Boards"
                    message="Choose the PFO boards where dependency epics will be created"
                    error={false}
                    className="form-field"
                  >
                    <PfoBoardDropdown onSelect={this.handleBoardSelection} />
                  </FormItem>
                </div>

                <div className="form-actions">
                  <ButtonGroup>
                    <Button
                      text="Create Epic"
                      primary={true}
                      iconProps={{ iconName: "Add" }}
                      onClick={() => this.handleSubmit()}
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

  private handleSubmit = async () => {
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

    for (const board of this.selectedBoards) {
      for (const leadEpic of this.selectedEpics) {
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

          // const relationOperations: JsonPatchDocument = [
          //   {
          //     op: "add",
          //     path: "/relations/-",
          //     value: {
          //       rel: "System.LinkTypes.Hierarchy-Reverse",
          //       url: parentUrl,
          //       attributes: {
          //         isLocked: false,
          //         name: "Parent",
          //         comment: "Parent project relationship",
          //       },
          //     },
          //   },
          //   {
          //     op: "add",
          //     path: "/relations/-",
          //     value: {
          //       rel: "System.LinkTypes.Dependency-Forward",
          //       url: parentUrl,
          //       attributes: {
          //         isLocked: false,
          //         name: "Successor",
          //         comment: "Dependency lead epic relationship",
          //       },
          //     },
          //   },
          // ];

          // console.log("Relation Operations:", relationOperations);

          // await witClient.updateWorkItem(relationOperations, dependencyEpic.id, this.state.projectContext.id);
        } catch (error) {
          console.log(`Failed to create dependency epic for Lead Epic ${leadEpic.id} on team ${board.text}`)
        }
      }
    }
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
  };

  private handleBoardSelection = (boards: Array<IListBoxItem<{ areaPath: string }>>) => {
    this.selectedBoards = boards;
  };
}

showRootComponent(<WorkHubGroup />);
