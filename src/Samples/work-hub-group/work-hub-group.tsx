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

import { FormItem } from "azure-devops-ui/FormItem";
import { AsyncEpicTagPicker, TagItem } from "./epic-tag-picker";
import { Button } from "azure-devops-ui/Button";
import { ButtonGroup } from "azure-devops-ui/ButtonGroup";
import PfoBoardDropdown from "./multi-select-dropdown";
import { ProjectTagPicker, ProjectTagItem } from "./project-tag-picker";

// removed unused observable pattern in favor of controlled string value

interface IWorkHubGroupState {
  projectContext: any;
  assignedWorkItems: any[];
  projectIdInput: string;
  projectIdValid: boolean;
  projectSelected?: ProjectTagItem | undefined;
  epicTags: TagItem[];
  selectedPfoCount: number;
  submitting: boolean;
  touched: {
    projectId: boolean;
    epics: boolean;
    pfoBoards: boolean;
  };
}

class WorkHubGroup extends React.Component<{}, IWorkHubGroupState> {
  constructor(props: {}) {
    super(props);
    this.state = {
      projectContext: undefined,
      assignedWorkItems: [],
      projectIdInput: "",
      projectIdValid: false,
      projectSelected: undefined,
      epicTags: [],
      selectedPfoCount: 0,
      submitting: false,
      touched: { projectId: false, epics: false, pfoBoards: false },
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
    const { projectIdValid, epicTags, selectedPfoCount, submitting, touched } =
      this.state;
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
                    label="Project"
                    message={
                      touched.projectId && !projectIdValid
                        ? "Select a valid Project work item"
                        : "Search and select the Project work item"
                    }
                    error={touched.projectId && !projectIdValid}
                    className="form-field"
                  >
                    <ProjectTagPicker
                      onChange={(project) => {
                        const valid = !!project;
                        this.setState({
                          projectSelected: project,
                          projectIdValid: valid,
                          projectIdInput: project ? String(project.id) : "",
                        });
                      }}
                    />
                  </FormItem>

                  <FormItem
                    label="Source of Truth Epic ID"
                    message="Select the epic that will serve as the source of truth"
                    error={touched.epics && epicTags.length === 0}
                    className="form-field"
                  >
                    <AsyncEpicTagPicker
                      onChange={(tags) => this.setState({ epicTags: tags })}
                    />
                  </FormItem>

                  <FormItem
                    label="PFO Boards"
                    message="Choose the PFO boards where dependency epics will be created"
                    error={touched.pfoBoards && selectedPfoCount === 0}
                    className="form-field"
                  >
                    <PfoBoardDropdown
                      onSelectionChange={(count) =>
                        this.setState({ selectedPfoCount: count })
                      }
                    />
                  </FormItem>
                </div>

                <div className="form-actions">
                  <ButtonGroup>
                    <Button
                      text={submitting ? "Submitting..." : "Create Epic"}
                      primary={true}
                      iconProps={{ iconName: "Add" }}
                      disabled={submitting}
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

  private handleSubmit = () => {
    this.setState(
      {
        touched: { projectId: true, epics: true, pfoBoards: true },
      },
      () => {
        if (!this.isFormValid()) {
          return;
        }
        this.setState({ submitting: true });
        // Placeholder for future creation logic
        setTimeout(() => {
          this.setState({ submitting: false });
          alert("Validation passed. Epic creation implementation coming soon.");
        }, 500);
      }
    );
  };

  private isFormValid(): boolean {
    const { projectIdValid, epicTags, selectedPfoCount } = this.state;
    return projectIdValid && epicTags.length > 0 && selectedPfoCount > 0;
  }

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

  // project is selected via ProjectTagPicker; validation occurs on submit
}

showRootComponent(<WorkHubGroup />);
