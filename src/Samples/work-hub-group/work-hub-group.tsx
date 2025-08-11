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
import { IListBoxItem } from "azure-devops-ui/ListBox";
import PfoBoardDropdown from "./multi-select-dropdown";

const errorObservable = new ObservableValue<string | undefined>("");

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
  private selectedPfoBoards: Array<IListBoxItem<{}>> = [];

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
                    Fill in the required information to create a new dependency epic
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
                      value={errorObservable}
                      onChange={(e) => (errorObservable.value = e.target.value)}
                      width={TextFieldWidth.auto}
                    />
                  </FormItem>

                  <FormItem
                    label="Source of Truth Epic ID"
                    message="Select the epic that will serve as the source of truth"
                    error={false}
                    className="form-field"
                  >
                    <AsyncEpicTagPicker onSelectionChange={(tags) => this.selectedEpics = tags} />
                  </FormItem>

                  <FormItem
                    label="PFO Boards"
                    message="Choose the PFO boards where dependency epics will be created"
                    error={false}
                    className="form-field"
                  >
                    <PfoBoardDropdown onSelectionChange={(boards) => this.selectedPfoBoards = boards} />
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

  private handleSubmit = () => {
    console.log("=== Form Values ===");
    
    // Project ID from TextField
    console.log("Project ID:", errorObservable.value);
    
    // Selected Epic Tags
    console.log("Selected Epic Tags:", this.selectedEpics);
    
    // Selected PFO Boards
    console.log("Selected PFO Boards:", this.selectedPfoBoards);
    
    console.log("===================");
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