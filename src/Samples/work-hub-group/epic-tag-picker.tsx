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

interface TagItem {
  id: number;
  text: string;
}

interface TagPickerProps {
  onSelect: (selectedTags: TagItem[]) => void;
}


export const AsyncEpicTagPicker: React.FunctionComponent<TagPickerProps> = ({ onSelect }) => {
  const [tagItems, setTagItems] = useObservableArray<TagItem>([]);
  const [suggestions, setSuggestions] = useObservableArray<TagItem>([]);
  const [suggestionsLoading, setSuggestionsLoading] =
    useObservable<boolean>(false);

  const areTagsEqual = (a: TagItem, b: TagItem) => {
    return a.id === b.id;
  };

  const convertItemToPill = (tag: TagItem) => {
    return {
      content: `${tag.id}: ${tag.text}`,
      onClick: () => alert(`Clicked tag "${tag.text}"`),
    };
  };

  const onSearchChanged = async (searchValue: string) => {

    setSuggestionsLoading(true);

    if (!searchValue) {
        setSuggestions([])
        setSuggestionsLoading(false);
        return;
    }

      try {
        const client = getClient(WorkItemTrackingRestClient);
        const searchId = parseInt(searchValue);
        let wiqlQuery: Wiql;

        if (!isNaN(searchId)) {
          wiqlQuery = {
            query: `
                SELECT [System.Id], [System.Title]
                FROM WorkItems
                WHERE [System.WorkItemType] = 'Epic'
                AND [System.Id] = ${searchId}
            `,
          };
        } else {
            wiqlQuery = {
                query: `
                    SELECT [System.Id], [System.Title]
                    FROM WorkItems
                    WHERE [System.WorkItemType] = 'Epic'
                    AND [System.Title] CONTAINS '${searchValue}'
                `
            }
        }

        const queryResult = await client.queryByWiql(wiqlQuery);
        const workItemIds = queryResult.workItems.map((wi) => wi.id).slice(0, 200);

        if (workItemIds.length > 0) {
        const workItems = await client.getWorkItems(workItemIds, undefined, ["System.Id", "System.Title"]);
        const suggestedItems = workItems.map(wi => ({ id: wi.id, text: wi.fields["System.Title"]}));

         setSuggestions(
          suggestedItems
            .filter(
              (testItem) =>
                tagItems.value.findIndex(
                  (testSuggestion) => testSuggestion.id == testItem.id
                ) === -1
            )
        );
      } else {
       setSuggestions([]);
      }

        setSuggestionsLoading(false);
      } catch(error) {
          console.log("Error fetching epic suggestions:", error);
          setSuggestionsLoading(false);
        setSuggestions([]);
      }
  };

  const onTagAdded = (tag: TagItem) => {
    const newTags = [...tagItems.value, tag]
    setTagItems(newTags);
    onSelect(newTags)
  };

  const onTagRemoved = (tag: TagItem) => {
    const newTags = tagItems.value.filter((x) => x.id !== tag.id)
    setTagItems(newTags);
    onSelect(newTags)
  };

  const renderSuggestionItem = (tag: ISuggestionItemProps<TagItem>) => {
    return <div className="body-m">{tag.item.id}: {tag.item.text}</div>;
  };

  return (
    <div className="flex-column">
      <TagPicker
        areTagsEqual={areTagsEqual}
        convertItemToPill={convertItemToPill}
        noResultsFoundText={"No results found"}
        onSearchChanged={onSearchChanged}
        onSearchChangedDebounceWait={500}
        onTagAdded={onTagAdded}
        onTagRemoved={onTagRemoved}
        renderSuggestionItem={renderSuggestionItem}
        selectedTags={tagItems}
        suggestions={suggestions}
        suggestionsLoading={suggestionsLoading}
        ariaLabel={"Search for Epics by Id or Title"}
        placeholderText="Enter Epic ID or Title to search..."
      />
    </div>
  );
};
