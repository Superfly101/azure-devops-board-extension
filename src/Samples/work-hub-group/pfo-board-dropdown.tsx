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
  teams: Array<IListBoxItem<{ areaPath: string }>>;
}

interface PfoBoardDropdownProps {
  selectedBoards: Array<IListBoxItem<{ areaPath: string }>>;
  onSelect: (teams: Array<IListBoxItem<{ areaPath: string}>>) => void;
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
          this.setupSelectionChangeLister();
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

  public componentDidUpdate(prevProps: PfoBoardDropdownProps) {
    // Sync internal selection with prop changes
    if (prevProps.selectedBoards !== this.props.selectedBoards) {
      this.updateSelectionFromProps();
    }
  }

  private isUpdatingFromProps = false;

  private updateSelectionFromProps = () => {
    this.isUpdatingFromProps = true;
    
    this.selection.clear();
    
    this.props.selectedBoards.forEach(selectedBoard => {
      const index = this.state.teams.findIndex(team => team.id === selectedBoard.id);
      if (index !== -1) {
        this.selection.select(index);
      }
    });
    
    this.isUpdatingFromProps = false;
  };

  private setupSelectionChangeLister = () => {
    this.selection.subscribe(() => {
      // Prevent infinite loop when updating from props
      if (this.isUpdatingFromProps) {
        return;
      }

      const selectedItems: Array<IListBoxItem<{ areaPath: string}>> = [];
      this.selection.value.forEach(range => {
        for (let i = range.beginIndex; i <= range.endIndex; i++) {
          if (this.state.teams[i]) {
            selectedItems.push(this.state.teams[i]);
          }
        }
      });
      this.props.onSelect(selectedItems);
    });
  };

  public render(): JSX.Element {
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
                    disabled: this.selection.selectedCount === 0,
                    iconProps: { iconName: "Clear" },
                    text: "Clear",
                    onClick: () => {
                      this.selection.clear();
                    },
                  },
                ]}
                className="example-dropdown"
                items={this.state.teams}
                selection={this.selection}
                placeholder="Select an Option"
                showFilterBox={true}
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
          const teamContext: TeamContext = {
            projectId: project.id,
            project: project.name,
            teamId: team.id,
            team: team.name,
          };

          const boards = await workClient.getBoards(teamContext);

          const teamFieldValues = await workClient.getTeamFieldValues(teamContext);

          if (boards && boards.length > 0) {
            return { 
              id: team.id, 
              text: team.name, 
              data: { areaPath: teamFieldValues.defaultValue } 
            };
          }
        } catch (error) {
          console.log(`Could not get board for team ${team.name}: ${error}`);
        }
        return null;
      });

      const result = await Promise.all(pfoBoardsPromises);
      const teamsWithBoards = result.filter((item) => item != null) as Array<
        IListBoxItem<{ areaPath: string }>
      >;

      this.setState({ teams: teamsWithBoards }, () => {
        // After teams are loaded, sync with selected boards from props
        this.updateSelectionFromProps();
      });
    } catch (error) {
      console.error("Failed to load PFO Teams: ", error);
    }
  }
}