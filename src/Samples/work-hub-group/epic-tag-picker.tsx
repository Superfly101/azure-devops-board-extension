// epic-tag-picker.tsx

import * as React from "react";
import { TagPicker } from "azure-devops-ui/TagPicker";
import { useObservableArray, useObservable } from "azure-devops-ui/Core/Observable";
import { ISuggestionItemProps } from "azure-devops-ui/SuggestionsList";
import { getClient } from "azure-devops-extension-api";
import { WorkItemTrackingRestClient, Wiql } from "azure-devops-extension-api/WorkItemTracking";

interface TagItem {
    id: number;
    text: string;
}

interface AsyncEpicTagPickerProps {
    onSelectionChange?: (selectedTags: TagItem[]) => void;
    disabled?: boolean;
}

export const AsyncEpicTagPicker: React.FunctionComponent<AsyncEpicTagPickerProps> = ({ 
    onSelectionChange, 
    disabled = false 
}) => {
    const [tagItems, setTagItems] = useObservableArray<TagItem>([]);
    const [suggestions, setSuggestions] = useObservableArray<TagItem>([]);
    const [suggestionsLoading, setSuggestionsLoading] = useObservable<boolean>(true);
    const timeoutId = React.useRef<number>(0);

    const areTagsEqual = (a: TagItem, b: TagItem) => {
        return a.id === b.id;
    };

    const convertItemToPill = (tag: TagItem) => {
        return {
            content: `${tag.id}: ${tag.text}`,
            onClick: disabled ? undefined : () => alert(`Clicked tag "${tag.text}"`)
        };
    };

    const onSearchChanged = (searchValue: string) => {
        if (disabled) return;
        
        clearTimeout(timeoutId.current);

        if (!searchValue) {
            setSuggestions([]);
            setSuggestionsLoading(false);
            return;
        }

        timeoutId.current = window.setTimeout(async () => {
            setSuggestionsLoading(true);
            try {
                const witClient = getClient(WorkItemTrackingRestClient);
                const searchId = parseInt(searchValue, 10);
                let query: Wiql;

                if (!isNaN(searchId)) {
                    query = {
                        query: `
                            SELECT [System.Id], [System.Title]
                            FROM WorkItems
                            WHERE [System.WorkItemType] = 'Epic'
                            AND [System.Id] = ${searchId}
                        `
                    };
                } else {
                    query = {
                        query: `
                            SELECT [System.Id], [System.Title]
                            FROM WorkItems
                            WHERE [System.WorkItemType] = 'Epic'
                            AND [System.Title] CONTAINS '${searchValue}'
                        `
                    };
                }
                
                const queryResult = await witClient.queryByWiql(query);
                const workItemIds = queryResult.workItems.map(wi => wi.id);

                if (workItemIds.length > 0) {
                    const workItems = await witClient.getWorkItems(workItemIds, undefined, ["System.Id", "System.Title"]);
                    const suggestedItems = workItems.map(wi => ({ id: wi.id, text: wi.fields["System.Title"] }));
                    setSuggestions(suggestedItems.filter(
                        testItem =>
                            tagItems.value.findIndex(
                                testSuggestion => testSuggestion.id == testItem.id
                            ) === -1
                    ));
                } else {
                    setSuggestions([]);
                }
            } catch (error) {
                console.error("Error fetching epic suggestions:", error);
                setSuggestions([]);
            } finally {
                setSuggestionsLoading(false);
            }
        }, 500);
    };

    const onTagAdded = (tag: TagItem) => {
        if (disabled) return;
        
        const newTags = [...tagItems.value, tag];
        setTagItems(newTags);
        if (onSelectionChange) {
            onSelectionChange(newTags);
        }
    };

    const onTagRemoved = (tag: TagItem) => {
        if (disabled) return;
        
        const newTags = tagItems.value.filter(x => x.id !== tag.id);
        setTagItems(newTags);
        if (onSelectionChange) {
            onSelectionChange(newTags);
        }
    };

    const renderSuggestionItem = (tag: ISuggestionItemProps<TagItem>) => {
        return <div className="body-m">{tag.item.id}: {tag.item.text}</div>;
    };

    return (
        <div className={`flex-column ${disabled ? 'disabled' : ''}`}>
            <TagPicker
                areTagsEqual={areTagsEqual}
                convertItemToPill={convertItemToPill}
                noResultsFoundText={"No results found"}
                onSearchChanged={onSearchChanged}
                onTagAdded={onTagAdded}
                onTagRemoved={onTagRemoved}
                renderSuggestionItem={renderSuggestionItem}
                selectedTags={tagItems}
                suggestions={suggestions}
                suggestionsLoading={suggestionsLoading}
                ariaLabel={"Search for epics by ID or title"}
                
            />
        </div>
    );
};