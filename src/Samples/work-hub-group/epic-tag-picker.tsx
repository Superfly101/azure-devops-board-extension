import * as React from "react";
import { TagPicker } from "azure-devops-ui/TagPicker";
import { useObservableArray, useObservable } from "azure-devops-ui/Core/Observable";
import { ISuggestionItemProps } from "azure-devops-ui/SuggestionsList";

interface TagItem {
    id: number;
    text: string;
}

const tagData: TagItem[] = [
    {
        id: 1,
        text: "Item 1"
    },
    {
        id: 2,
        text: "Item 2"
    },
    {
        id: 3,
        text: "Item 3"
    }
];

export const AsyncEpicTagPicker: React.FunctionComponent<{}> = () => {
    const [tagItems, setTagItems] = useObservableArray<TagItem>(tagData.slice(0, 2));
    const [suggestions, setSuggestions] = useObservableArray<TagItem>([]);
    const [suggestionsLoading, setSuggestionsLoading] = useObservable<boolean>(false);
    const timeoutId = React.useRef<number>(0);

    const areTagsEqual = (a: TagItem, b: TagItem) => {
        return a.id === b.id;
    };

    const convertItemToPill = (tag: TagItem) => {
        return {
            content: tag.text,
            onClick: () => alert(`Clicked tag "${tag.text}"`)
        };
    };

    const onSearchChanged = (searchValue: string) => {
        clearTimeout(timeoutId.current);

        // Simulates a 1000ms round trip to retrieve values
        setSuggestionsLoading(true);

        timeoutId.current = window.setTimeout(() => {
            setSuggestions(
                tagData
                    .filter(
                        // Items not already included
                        testItem =>
                            tagItems.value.findIndex(
                                testSuggestion => testSuggestion.id == testItem.id
                            ) === -1
                    )
                    .filter(
                        testItem =>
                            testItem.text.toLowerCase().indexOf(searchValue.toLowerCase()) > -1
                    )
            );

            setSuggestionsLoading(false);
        }, 1000);
    };

    const onTagAdded = (tag: TagItem) => {
        setTagItems([...tagItems.value, tag]);
    };

    const onTagRemoved = (tag: TagItem) => {
        setTagItems(tagItems.value.filter(x => x.id !== tag.id));
    };

    const renderSuggestionItem = (tag: ISuggestionItemProps<TagItem>) => {
        return <div className="body-m">{tag.item.text}</div>;
    };

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
                selectedTags={tagItems}
                suggestions={suggestions}
                suggestionsLoading={suggestionsLoading}
                ariaLabel={"Search for additional tags"}
            />
        </div>
    );
};