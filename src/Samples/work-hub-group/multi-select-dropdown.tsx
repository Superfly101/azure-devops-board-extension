import * as React from "react";
import { Dropdown } from "azure-devops-ui/Dropdown";
import { Observer } from "azure-devops-ui/Observer";
import { DropdownMultiSelection } from "azure-devops-ui/Utilities/DropdownSelection";
import { IListBoxItem, ListBoxItemType } from "azure-devops-ui/ListBox";

export const dropdownItems: Array<IListBoxItem<{}>> = [
    { id: "header", text: "Header", type: ListBoxItemType.Header },
    { id: "first", text: "first", groupId: "header" },
    // { id: "divider", type: ListBoxItemType.Divider },
    // { id: "header2", text: "Header 2", type: ListBoxItemType.Header },
    { id: "second", text: "second", groupId: "header" },
    { id: "third", text: "third", groupId: "header" },
    { id: "fourth", text: "fourth", groupId: "header" },
    { id: "sixth", text: "sixth", groupId: "header" },
    { id: "seventh", text: "seventh", groupId: "header" },
    { id: "eighth", text: "eighth", groupId: "header" },
    { id: "ninth", text: "ninth", groupId: "header" },
    { id: "tenth", text: "tenth", groupId: "header" },
    { id: "eleventh", text: "eleventh", groupId: "header" },
    { id: "twelvth", text: "twelvth", groupId: "header" },
    { id: "thirteenth", text: "thirteenth", groupId: "header" },
    { id: "fourteenth", text: "fourteenth", groupId: "header" },
    { id: "fifteenth", text: "fifteenth", groupId: "header" }
];


export default class DropdownMultiSelectExample extends React.Component {
    private selection = new DropdownMultiSelection();

    public render() {
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
                                items={dropdownItems}
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
}