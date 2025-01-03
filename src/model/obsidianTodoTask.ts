/* eslint-disable unicorn/filename-case */
import {
    type AttachmentBase,
    type AttachmentSession,
    type ChecklistItem,
    type DateTimeTimeZone,
    type Extension,
    type Importance,
    type ItemBody,
    type LinkedResource,
    type NullableOption,
    type PatternedRecurrence,
    type TaskStatus,
    type TodoTask,
} from '@microsoft/microsoft-graph-types';
import {type ISettingsManager} from 'src/utils/settingsManager.js';
import {t} from '../lib/lang.js';
import {logging} from '../lib/logging.js';
import {IMPORTANCE_REGEX, STATUS_SYMBOL_REGEX, TASK_REGEX} from '../constants.js';

/**
 * Represents a task in Obsidian that can be synchronized with Microsoft To Do.
 * Implements the TodoTask interface.
 */
export class ObsidianTodoTask implements TodoTask {
    id: string;

    // The task body that typically contains information about the task.
    public body?: NullableOption<ItemBody>;
    /**
	 * The date and time when the task body was last modified. By default, it is in UTC. You can provide a custom time zone in
	 * the request header. The property value uses ISO 8601 format and is always in UTC time. For example, midnight UTC on Jan
	 * 1, 2020 would look like this: '2020-01-01T00:00:00Z'.
	 */
    public bodyLastModifiedDateTime?: string;
    /**
	 * The categories associated with the task. Each category corresponds to the displayName property of an outlookCategory
	 * that the user has defined.
	 */
    public categories?: NullableOption<string[]>;
    /**
     * The date and time in the specified time zone that the task was finished.
     */
    public completedDateTime?: NullableOption<DateTimeTimeZone>;
    /**
	 * The date and time when the task was created. By default, it is in UTC. You can provide a custom time zone in the
	 * request header. The property value uses ISO 8601 format. For example, midnight UTC on Jan 1, 2020 would look like this:
	 * '2020-01-01T00:00:00Z'.
	 */
    public createdDateTime?: string;
    // The date and time in the specified time zone that the task is to be finished.

    /**
     * The date and time in the specified time zone that the task is to be finished.
     */
    public dueDateTime?: NullableOption<DateTimeTimeZone>;

    /**
     * Indicates whether the task has attachments.
     */
    public hasAttachments?: NullableOption<boolean>;

    /**
     * The importance of the task. Possible values are: low, normal, high.
     */
    public importance?: Importance;

    /**
     * Set to true if an alert is set to remind the user of the task.
     */
    public isReminderOn?: boolean;

    /**
	 * The date and time when the task was last modified. By default, it is in UTC. You can provide a custom time zone in the
	 * request header. The property value uses ISO 8601 format and is always in UTC time. For example, midnight UTC on Jan 1,
	 * 2020 would look like this: '2020-01-01T00:00:00Z'.
	 */
    public lastModifiedDateTime?: string;

    /**
     * The recurrence pattern for the task.
     */
    public recurrence?: NullableOption<PatternedRecurrence>;

    /**
     * The date and time in the specified time zone for a reminder alert of the task to occur.
     */
    public reminderDateTime?: NullableOption<DateTimeTimeZone>;

    /**
     * The date and time in the specified time zone that the task is to start.
     */
    public startDateTime?: NullableOption<DateTimeTimeZone>;

    /**
	 * Indicates the state or progress of the task. Possible values are: notStarted, inProgress, completed, waitingOnOthers,
	 * deferred.
	 */
    public status?: TaskStatus;
    /**
     * A brief description of the task.
     */
    public title?: NullableOption<string>;

    /**
     * A collection of attachments linked to the task.
     */
    public attachments?: NullableOption<AttachmentBase[]>;

    /**
     * A collection of attachment sessions linked to the task.
     */
    public attachmentSessions?: NullableOption<AttachmentSession[]>;

    /**
     * A collection of checklist items linked to the task.
     */
    public checklistItems?: NullableOption<ChecklistItem[]>;

    /**
     * The collection of open extensions defined for the task. Nullable.
     */
    public extensions?: NullableOption<Extension[]>;

    /**
     * A collection of resources linked to the task.
     */
    public linkedResources?: NullableOption<LinkedResource[]>;

    /**
     * The block link associated with the task.
     */
    public blockLink?: string;

    /**
     * Logger instance for logging task-related information.
     */
    private readonly logger = logging.getLogger('mstodo-sync.ObsidianTodoTask');

    /**
     * The original title of the task.
     */
    private readonly originalTitle: string;

    /**
     * Creates an instance of ObsidianTodoTask.
     * @param settingsManager - The settings manager instance.
     * @param line - The line of text representing the task.
     * @param fileName - The name of the file where the task is located.
     */
    constructor(private readonly settingsManager: ISettingsManager, line: string, public fileName: string) {
        this.originalTitle = line;
        this.logger.debug(`Creating: '${this.originalTitle}'`);

        this.title = line.trim();

        // This will strip out the block link if it exists as
        // it is part of this plugin and not user specified.
        this.checkForBlockLink(line);

        // This will strip out the checkbox if in title.
        this.checkForStatus(line);

        // - [ ] Adding in updated linked resources updated from list dump  🔎[[2024-12-30]]  🔎[[2024-12-30]] ^MSTDa8de00053
        // This will strip out the created date if in title.
        if (this.title.includes(settingsManager.settings.displayOptions_TaskCreatedPrefix)) {
            this.title = this.title.replaceAll(/🔎\[\[.*]]/g, '').replaceAll(/🔎\d{4}-\d{2}-\d{2}/g, '');
        }

        this.checkForImportance(line);

        this.title = this.title
            .trim()
            .replaceAll(/(- \[([ /x])] )|\*|^> |^#* |- /gm, '')
            .trim();

        this.body = {
            content: `${t('displayOptions_CreatedInFile')} [[${this.fileName}]]`,
            contentType: 'text',
        };

        this.linkedResources ||= [];

        const redirectUrl = `http://192.168.0.137:8901/redirectpage.html?vault=brainstore&filepath=${encodeURIComponent(fileName)}&block=${this.blockLink}`;
        this.linkedResources.push({
            webUrl: redirectUrl,
            applicationName: 'Obsidian Microsoft To Do Sync',
            externalId: this.blockLink,
            displayName: `Tracking Block Link: ${this.blockLink}`,
        });

        this.logger.debug(`Created: '${this.title}'`);
    }

    /**
	 * Cache the ID internally and generate block link.
	 *
	 * @param {string} [id]
	 * @return {*}  {Promise<void>}
	 * @memberof ObsidianTodoTask
	 */
    public async cacheTaskId(id: string): Promise<void> {
        this.settingsManager.settings.taskIdIndex += 1;

        const index = `MSTD${Math.random().toString(20).slice(2, 6)}${this.settingsManager.settings.taskIdIndex
            .toString()
            .padStart(5, '0')}`;
        this.logger.debug(`id: ${id}, index: ${index}, taskIdIndex: ${this.settingsManager.settings.taskIdIndex}`);

        this.settingsManager.settings.taskIdLookup[index] = id ?? '';
        this.blockLink = index;
        this.id = id;

        this.settingsManager.saveSettings();
    }

    /**
     * Get the task as a TodoTask object.
     * @param withChecklist - Whether to include checklist items in the returned task.
     * @returns The task as a TodoTask object.
     */
    public getTodoTask(withChecklist = false): TodoTask {
        const toDo: TodoTask = {
            title: this.title,
        };

        if (this.body?.content && this.body.content.length > 0) {
            toDo.body = this.body;
        }

        if (this.status && this.status.length > 0) {
            toDo.status = this.status;
        }

        if (this.importance && this.importance.length > 0) {
            toDo.importance = this.importance;
        }

        if (withChecklist && this.checklistItems && this.checklistItems.length > 0) {
            toDo.checklistItems = this.checklistItems;
        }

        if (this.linkedResources && this.linkedResources.length > 0) {
            toDo.linkedResources = this.linkedResources;
        }

        return toDo;
    }

    /**
     * Get the task as a TodoTask object.
     * @param withChecklist - Whether to include checklist items in the returned task.
     * @returns The task as a TodoTask object.
     */
    public updateFromTodoTask(remoteTask: TodoTask) {
        this.title = remoteTask.title;

        if (remoteTask.body?.content && remoteTask.body.content.length > 0) {
            this.body = remoteTask.body;
        }

        if (remoteTask.status && remoteTask.status.length > 0) {
            this.status = remoteTask.status;
        }

        if (remoteTask.importance && remoteTask.importance.length > 0) {
            this.importance = remoteTask.importance;
        }

        if (remoteTask.linkedResources && remoteTask.linkedResources.length > 0) {
            this.linkedResources = remoteTask.linkedResources;
        }

        // Need to determine if we want to update the checklist items
        // if (withChecklist && remoteTask.checklistItems && remoteTask.checklistItems.length > 0) {
        //     this.checklistItems = remoteTask.checklistItems;
        // }
    }

    /**
     * Set the body content of the task.
     * @param body - The body content to set.
     */
    public setBody(body: string) {
        this.body = {
            content: body,
            contentType: 'text',
        };
    }

    /**
     * Add a checklist item to the task.
     * @param item - The checklist item to add.
     */
    public addChecklistItem(item: string) {
        this.checklistItems ||= [];

        this.checklistItems.push({
            displayName: item
                .trim()
                .replaceAll(/(- \[([ /x])] )|\*|^> |^#* |- /gm, '')
                .trim(),
        });
    }

    /**
	 * Return the task as a well formed markdown task.
	 *
	 * @return {*}  {string}
	 * @memberof ObsidianTodoTask
	 */
    public getMarkdownTask(singleLine: boolean): string {
        let output: string;

        // Format and display the task which is the first line.
        const format = this.settingsManager.settings.displayOptions_ReplacementFormat;
        const priorityIndicator = this.importance === 'normal' ? '' : this.getPriorityIndicator();

        output = format
            .replace(TASK_REGEX, this.title?.trim() ?? '')
            .replace(STATUS_SYMBOL_REGEX, this.getStatusIndicator());

        output = output.includes(priorityIndicator) ? output.replace(IMPORTANCE_REGEX, '') : output.replace(IMPORTANCE_REGEX, priorityIndicator);

        // Append block link at the end if it exists
        if (this.hasBlockLink && this.blockLink) {
            output = `${output.trim()} ^${this.blockLink}`;
        }

        this.logger.debug(`Updated task: '${output}'`);

        let formattedBody = '';
        let formattedChecklist = '';

        // Add in the body if it exists and indented by two spaces.
        if (this.body?.content && this.body.content.length > 0) {
            for (const bodyLine of this.body?.content.split('\n')) { // eslint-disable-line no-unsafe-optional-chaining
                if (bodyLine.trim().length > 0) {
                    formattedBody += '  ' + bodyLine + '\n';
                }
            }
        }
        // This.logger.debug(`formattedBody: '${formattedBody}'`);

        if (this.checklistItems && this.checklistItems.length > 0) {
            for (const item of this.checklistItems) {
                formattedChecklist += item.isChecked ? '  - [x] ' + item.displayName + '\n' : '  - [ ] ' + item.displayName + '\n';
            }
        }
        // This.logger.debug(`formattedChecklist: '${formattedChecklist}'`);

        output = singleLine ? `${output.trim()}` : `${output.trim()}\n${formattedBody}${formattedChecklist}`;
        // This.logger.debug(`output: '${output}'`);

        return output;
    }

    /**
     * Check the task title for a status indicator and update the status accordingly.
     * @param line - The line of text representing the task.
     */
    private checkForStatus(line: string) {
        const regex = /\[(.)]/;

        const m = regex.exec(line);
        if (m && m.length > 0) {
            this.status = m[1] === 'x' ? 'completed' : 'notStarted';
            this.title = this.title?.replace(regex, '').trim();
        } else {
            this.status = 'notStarted';
        }
    }

    /**
     * Check the task title for an importance indicator and update the importance accordingly.
     * @param line - The line of text representing the task.
     */
    private checkForImportance(line: string) {
        this.importance = 'normal';

        if (line.includes(this.settingsManager.settings.displayOptions_TaskImportance_Low)) {
            this.importance = 'low';
        }

        if (line.includes(this.settingsManager.settings.displayOptions_TaskImportance_High)) {
            this.importance = 'high';
        }
    }

    /**
     * Get the priority indicator based on the task's importance.
     * @returns The priority indicator as a string.
     */
    private getPriorityIndicator(): string {
        switch (this.importance) {
            case 'normal': {
                return this.settingsManager.settings.displayOptions_TaskImportance_Normal;
            }

            case 'low': {
                return this.settingsManager.settings.displayOptions_TaskImportance_Low;
            }

            case 'high': {
                return this.settingsManager.settings.displayOptions_TaskImportance_High;
            }

            default: {
                return '';
            }
        }
    }

    /**
     * Get the status indicator based on the task's status.
     * @returns The status indicator as a string.
     */
    private getStatusIndicator(): string {
        switch (this.status) {
            case 'notStarted': {
                return this.settingsManager.settings.displayOptions_TaskStatus_NotStarted;
            }

            case 'inProgress': {
                return this.settingsManager.settings.displayOptions_TaskStatus_InProgress;
            }

            case 'completed': {
                return this.settingsManager.settings.displayOptions_TaskStatus_Completed;
            }

            default: {
                return ' ';
            }
        }
    }

    /**
     * Check the task title for a block link and update the block link and ID accordingly.
     * @param line - The line of text representing the task.
     */
    private checkForBlockLink(line: string) {
        const blockLinkRegex = /\^(?!.*\^)([A-Za-z\d]+)/gm;
        const blockLinkMatch = blockLinkRegex.exec(line);
        if (blockLinkMatch) {
            this.blockLink = blockLinkMatch[1];

            // If there's a 'Created at xxxx' replaced line,
            // it's not enough to get a cleanTaskTitle after the next line.
            this.title = this.title?.replace(`^${this.blockLink}`, '');
        }

        if (this.hasBlockLink && this.blockLink) {
            this.id = this.settingsManager.settings.taskIdLookup[this.blockLink];
        }
    }

    /**
     * Get the clean title of the task, without any block links or status indicators.
     * @returns The clean title as a string.
     */
    public get cleanTitle(): string {
        return '';
    }

    /**
     * Check if the task has a block link.
     * @returns True if the task has a block link, false otherwise.
     */
    public get hasBlockLink(): boolean {
        return this.blockLink !== undefined && this.blockLink.length > 0;
    }

    /**
     * Check if the task has an id for the remote task.
     * @returns True if the task has a id set, false otherwise.
     */
    public get hasId(): boolean {
        return this.id !== undefined && this.id.length > 0;
    }
}