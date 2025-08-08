import * as React from "react";
import { TagPicker } from "azure-devops-ui/TagPicker";
import {
  useObservableArray,
  useObservable,
} from "azure-devops-ui/Core/Observable";
import { ISuggestionItemProps } from "azure-devops-ui/SuggestionsList";
import { getClient } from "azure-devops-extension-api";
import {
  WorkItemTrackingRestClient,
  Wiql,
} from "azure-devops-extension-api/WorkItemTracking";

export interface ProjectTagItem {
  id: number;
  text: string;
}

interface ProjectTagPickerProps {
  onChange?: (project: ProjectTagItem | undefined) => void;
}

export const ProjectTagPicker: React.FunctionComponent<
  ProjectTagPickerProps
> = ({ onChange }) => {
  const [selected, setSelected] = useObservableArray<ProjectTagItem>([]);
  const [suggestions, setSuggestions] = useObservableArray<ProjectTagItem>([]);
  const [suggestionsLoading, setSuggestionsLoading] =
    useObservable<boolean>(false);
  const timeoutId = React.useRef<number>(0);

  const areTagsEqual = (a: ProjectTagItem, b: ProjectTagItem) => a.id === b.id;

  const convertItemToPill = (tag: ProjectTagItem) => ({
    content: `${tag.id}: ${tag.text}`,
  });

  const onSearchChanged = (searchValue: string) => {
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
                WHERE [System.WorkItemType] = 'Project'
                AND [System.Id] = ${searchId}
            `,
          };
        } else {
          query = {
            query: `
                SELECT [System.Id], [System.Title]
                FROM WorkItems
                WHERE [System.WorkItemType] = 'Project'
                AND [System.Title] CONTAINS '${searchValue}'
            `,
          };
        }
        const queryResult = await witClient.queryByWiql(query);
        const workItemIds = queryResult.workItems.map((wi) => wi.id);
        if (workItemIds.length > 0) {
          const workItems = await witClient.getWorkItems(
            workItemIds,
            undefined,
            ["System.Id", "System.Title"]
          );
          const suggested = workItems.map((wi) => ({
            id: wi.id,
            text: wi.fields["System.Title"],
          }));
          setSuggestions(suggested);
        } else {
          setSuggestions([]);
        }
      } catch (e) {
        // Fallback to no suggestions on error
        setSuggestions([]);
      } finally {
        setSuggestionsLoading(false);
      }
    }, 400);
  };

  const onTagAdded = (tag: ProjectTagItem) => {
    const next = [tag];
    setSelected(next);
    if (onChange) onChange(tag);
  };

  const onTagRemoved = (_tag: ProjectTagItem) => {
    setSelected([]);
    if (onChange) onChange(undefined);
  };

  const renderSuggestionItem = (tag: ISuggestionItemProps<ProjectTagItem>) => (
    <div className="body-m">
      {tag.item.id}: {tag.item.text}
    </div>
  );

  return (
    <div className="flex-column">
      <TagPicker
        areTagsEqual={areTagsEqual}
        convertItemToPill={convertItemToPill}
        noResultsFoundText={"No results found"}
        onSearchChanged={onSearchChanged}
        onTagAdded={onTagAdded}
        onTagRemoved={onTagRemoved}
        renderSuggestionItem={renderSuggestionItem}
        selectedTags={selected}
        suggestions={suggestions}
        suggestionsLoading={suggestionsLoading}
        ariaLabel={"Search and select the Project work item by ID or title"}
      />
    </div>
  );
};
