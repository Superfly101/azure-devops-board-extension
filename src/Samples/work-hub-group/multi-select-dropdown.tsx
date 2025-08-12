// multi-select-dropdown.tsx

import * as React from "react";
import { Dropdown } from "azure-devops-ui/Dropdown";
import { Observer } from "azure-devops-ui/Observer";
import { DropdownMultiSelection } from "azure-devops-ui/Utilities/DropdownSelection";
import { IListBoxItem } from "azure-devops-ui/ListBox";
import * as SDK from "azure-devops-extension-sdk";
import {
  CommonServiceIds,
  getClient,
  IProjectPageService,
} from "azure-devops-extension-api";
import {
  CoreRestClient,
  TeamContext,
  WebApiTeam,
} from "azure-devops-extension-api/Core";
import { WorkRestClient } from "azure-devops-extension-api/Work";

interface IMultiSelectState {
  teams: Array<IListBoxItem<{}>>;
}

interface PfoBoardDropdownProps {
  onSelectionChange?: (selectedItems: Array<IListBoxItem<{}>>) => void;
  disabled?: boolean;
}

export default class PfoBoardDropdown extends React.Component<
  PfoBoardDropdownProps,
  IMultiSelectState
> {
  private selection = new DropdownMultiSelection();

  constructor(props: PfoBoardDropdownProps) {
    super(props);
    this.state = { teams: [] };
  }

  public componentDidMount() {
    try {
      SDK.ready()
        .then(() => {
          console.log("SDK is ready, loading project context...");
          this.loadTeams();
          this.setupSelectionChangeListener();
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

  private setupSelectionChangeListener = () => {
    this.selection.subscribe(() => {
      if (this.props.onSelectionChange && !this.props.disabled) {
        // Map selection ranges back to actual items
        const selectedItems: Array<IListBoxItem<{}>> = [];
        this.selection.value.forEach(range => {
          for (let i = range.beginIndex; i <= range.endIndex; i++) {
            if (this.state.teams[i]) {
              selectedItems.push(this.state.teams[i]);
            }
          }
        });
        this.props.onSelectionChange(selectedItems);
      }
    });
  };

  public render(): JSX.Element {
    const { disabled = false } = this.props;
    
    return (
      <div style={{ margin: "4px" }}>
        <Observer selection={this.selection}>
          {() => {
            return (
              <Dropdown
                ariaLabel="Multiselect"
                actions={[
                  {
                    className: "bolt-dropdown-action-right-button",
                    disabled: disabled || this.selection.selectedCount === 0,
                    iconProps: { iconName: "Clear" },
                    text: "Clear",
                    onClick: () => {
                      if (!disabled) {
                        this.selection.clear();
                      }
                    },
                  },
                ]}
                className={`example-dropdown ${disabled ? 'disabled' : ''}`}
                items={this.state.teams}
                selection={this.selection}
                placeholder={disabled ? "Creating epics..." : "Select an Option"}
                showFilterBox={!disabled}
                disabled={disabled}
              />
            );
          }}
        </Observer>
      </div>
    );
  }

  private async loadTeams(): Promise<void> {
    try {
      const projectService = await SDK.getService<IProjectPageService>(
        CommonServiceIds.ProjectPageService
      );
      const project = await projectService.getProject();

      if (!project) {
        console.log("Unable to retrieve Project");
        return;
      }

      const coreClient = getClient(CoreRestClient);
      const workClient = getClient(WorkRestClient);

      const teams = await coreClient.getTeams(project.id);

      const pfoBoardsPromises = teams.map(async (team: WebApiTeam) => {
        try {
          const boards = await workClient.getBoards({
            projectId: project.id,
            project: project.name,
            teamId: team.id,
            team: team.name,
          });

          if (boards && boards.length > 0) {
            return { id: team.id, text: team.name };
          }
        } catch (error) {
          console.log(`Could not get board for team ${team.name}: ${error}`);
        }
        return null;
      });

      const result = await Promise.all(pfoBoardsPromises);
      const teamsWithBoards = result.filter((item) => item != null) as Array<
        IListBoxItem<{}>
      >;

      this.setState({ teams: teamsWithBoards });
    } catch (error) {
      console.error("Failed to load PFO Teams: ", error);
    }
  }
}