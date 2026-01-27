import { type App, type FuzzyMatch, FuzzySuggestModal } from 'obsidian';
import { type TodoTaskList } from '@microsoft/microsoft-graph-types';

export interface ListSelectorResult {
    list: TodoTaskList | null; // null means "All Lists"
    cancelled: boolean;
}

/**
 * Modal for selecting a Microsoft To Do list.
 * Shows all available lists plus an "All Lists" option.
 */
export class ListSelectorModal extends FuzzySuggestModal<TodoTaskList | null> {
    private lists: TodoTaskList[];
    private resolve: (result: ListSelectorResult) => void;
    private result: ListSelectorResult = { list: null, cancelled: true };

    constructor(
        app: App,
        lists: TodoTaskList[],
        private readonly allowAllLists = true,
    ) {
        super(app);
        this.lists = lists;
        this.setPlaceholder('Select a list...');
    }

    /**
     * Opens the modal and returns a promise that resolves when user selects a list.
     */
    public openAndGetValue(): Promise<ListSelectorResult> {
        return new Promise((resolve) => {
            this.resolve = resolve;
            this.open();
        });
    }

    getItems(): (TodoTaskList | null)[] {
        // Add "All Lists" option at the beginning if allowed
        if (this.allowAllLists) {
            return [null, ...this.lists];
        }
        return this.lists;
    }

    getItemText(item: TodoTaskList | null): string {
        if (item === null) {
            return '📋 All Lists';
        }
        return item.displayName ?? 'Unnamed List';
    }

    selectSuggestion(value: FuzzyMatch<TodoTaskList | null>, evt: MouseEvent | KeyboardEvent): void {
        this.result = { list: value.item, cancelled: false };
        super.selectSuggestion(value, evt);
    }

    onChooseItem(item: TodoTaskList | null, _evt: MouseEvent | KeyboardEvent): void {
        // Handled in selectSuggestion to ensure it runs before onClose
        this.result = {
            list: item,
            cancelled: false,
        };
    }

    onClose(): void {
        if (this.resolve) {
            this.resolve(this.result);
        }
    }
}
