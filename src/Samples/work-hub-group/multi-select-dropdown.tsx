import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";
import {
    CommonServiceIds,
    getClient,
    IProjectPageService,
} from "azure-devops-extension-api";
import { CoreRestClient, WebApiTeam } from "azure-devops-extension-api/Core";
import { WorkRestClient } from "azure-devops-extension-api/Work";
import { Dropdown } from "azure-devops-ui/Dropdown";
import { Observer } from "azure-devops-ui/Observer";
import { DropdownMultiSelection } from "azure-devops-ui/Utilities/DropdownSelection";
import { IListBoxItem } from "azure-devops-ui/ListBox";

interface IMultiSelectState {
    teams: IListBoxItem<{}>[];
    loading: boolean;
}

export default class DropdownMultiSelectExample extends React.Component<{}, IMultiSelectState> {
    private selection = new DropdownMultiSelection();

    constructor(props: {}) {
        super(props);
        this.state = {
            teams: [],
            loading: true
        };
    }

    public componentDidMount() {
        this.loadTeams();
    }

    private async loadTeams() {
        await SDK.ready();
        const projectService = await SDK.getService<IProjectPageService>(
            CommonServiceIds.ProjectPageService
        );
        const project = await projectService.getProject();
        if (!project) {
            this.setState({ loading: false });
            return;
        }

        const coreClient = getClient(CoreRestClient);
        const workClient = getClient(WorkRestClient);

        const teams = await coreClient.getTeams(project.id);
        const teamsWithBoardsPromises = teams.map(async (team: WebApiTeam) => {
            try {
                const boards = await workClient.getBoards({
                    project: project.name,
                    projectId: project.id,
                    team: team.name,
                    teamId: team.id
                });
                if (boards && boards.length > 0) {
                    return { id: team.id, text: team.name };
                }
            } catch (error) {
                console.log(`Could not get boards for team ${team.name}: ${error}`);
            }
            return null;
        });

        const results = await Promise.all(teamsWithBoardsPromises);
        const teamsWithBoards = results.filter(team => team !== null) as Array<IListBoxItem<{}>>;

        this.setState({ teams: teamsWithBoards, loading: false });
    }

    public render() {
        const { loading, teams } = this.state;

        return (
            <div style={{ margin: "8px" }}>
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
                                        }
                                    }
                                ]}
                                className="example-dropdown"
                                items={teams}
                                loading={loading}
                                selection={this.selection}
                                placeholder="Select a Team"
                                showFilterBox={true}
                            />
                        );
                    }}
                </Observer>
            </div>
        );
    }
}