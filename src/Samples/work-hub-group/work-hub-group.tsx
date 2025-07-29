import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";

import "./work-hub-group.scss";

import { Header } from "azure-devops-ui/Header";
import { Page } from "azure-devops-ui/Page";
import { TextField, TextFieldWidth } from "azure-devops-ui/TextField";
import { FormItem } from "azure-devops-ui/FormItem";
import { SimpleTagPickerExample } from "./tag-picker";
import  DropdownMultiSelectExample  from "./multi-select-dropdown";

import { showRootComponent } from "../../Common";
import {
  CommonServiceIds,
  IProjectPageService,
  getClient,
} from "azure-devops-extension-api";
import { WorkItemTrackingRestClient } from "azure-devops-extension-api/WorkItemTracking";
import { ObservableValue } from "azure-devops-ui/Core/Observable";

interface IWorkHubGroup {
  projectContext: any;
  assignedWorkItems: any[];
}

const projectObservable = new ObservableValue<string | undefined>("");

class WorkHubGroup extends React.Component<{}, IWorkHubGroup> {
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
      <Page className="sample-hub flex-grow">
        <Header title="Custom Work Hub" />
        <div className="page-content">
          <FormItem label="Project Name" error={false} className="sample-form-section">
            <TextField
              ariaLabel="Project Name"
              value={projectObservable}
              onChange={(e, newValue) => {
                projectObservable.value = newValue;
              }}
              width={TextFieldWidth.auto}
            />
          </FormItem>
          
          <FormItem label="Source of truth Epic ID" error={true} className="sample-form-section">
            <SimpleTagPickerExample />
          </FormItem>

          <FormItem label="List of PFO boards" error={true} className="sample-form-section">
            <DropdownMultiSelectExample />
          </FormItem>



        </div>
      </Page>
    );
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
}

showRootComponent(<WorkHubGroup />);
