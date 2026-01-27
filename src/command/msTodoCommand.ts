import { type BlockCache, type DataAdapter, type Editor, type EditorPosition, MarkdownView } from 'obsidian';
import { ObsidianTodoTask } from 'src/model/obsidianTodoTask.js';
import { type TodoTask } from '@microsoft/microsoft-graph-types';
import { type SettingsManager } from 'src/utils/settingsManager.js';
import type MsTodoSync from '../main.js';
import { TasksDeltaCollection, type TodoApi } from '../api/todoApi.js';
import { t } from '../lib/lang.js';
import { log, logging } from '../lib/logging.js';
import { UserNotice } from 'src/lib/userNotice.js';

const userNotice = new UserNotice();

export function getTaskIdFromLine(line: string, plugin: MsTodoSync): string {
    const regex = /\^(?!.*\^)([A-Za-z\d]+)/gm;
    const blocklistMatch = regex.exec(line.trim());
    if (blocklistMatch) {
        const blockLink = blocklistMatch[1];
        const taskId = plugin.settings.taskIdLookup[blockLink];
        console.log(taskId);
        return taskId;
    }

    return '';
}

interface ISelection {
    start: EditorPosition;
    end?: EditorPosition;
    lines: number[];
}

/**
 * Retrieves the current lines from the editor based on the cursor position or selection.
 *
 * @param editor - The editor instance from which to get the current lines.
 * @returns A promise that resolves to a Selection object containing:
 * - `start`: The starting position of the cursor or selection.
 * - `end`: The ending position of the cursor or selection.
 * - `lines`: An array of line numbers that are currently selected or where the cursor is located.
 */
export async function getCurrentLinesFromEditor(editor: Editor): Promise<ISelection> {
    log('info', 'Getting current lines from editor', {
        from: editor.getCursor('from'),
        to: editor.getCursor('to'),
        anchor: editor.getCursor('anchor'),
        head: editor.getCursor('head'),
        general: editor.getCursor(),
    });

    // Const activeFile = this.app.workspace.getActiveFile();
    // const source = await this.app.vault.read(activeFile);

    let start: EditorPosition;
    let end: EditorPosition;
    // Let lines: string[] = [];
    let lines: number[] = [];
    if (editor.somethingSelected()) {
        start = editor.getCursor('from');
        end = editor.getCursor('to');
        // Lines = source.split('\n').slice(start.line, end.line + 1);
        lines = Array.from({ length: end.line + 1 - start.line }, (_v, k) => k + start.line);
    } else {
        start = editor.getCursor();
        end = editor.getCursor();
        // Lines = source.split('\n').slice(start.line, end.line + 1);
        lines.push(start.line);
    }

    return {
        start,
        end,
        lines,
    };
}

export async function cleanupCachedTaskIds(plugin: MsTodoSync) {
    const logger = logging.getLogger('mstodo-sync.command.lookupPluginBlocks');

    // Collect all the blocks and ids from the metadata cache under the app.
    const blockCache: Record<string, BlockCache> = populateBlockCache(plugin);

    // Iterate over all the internal cached task ids in settings. If the block is not found in the metadata cache
    // we will log it. The cache is a metadata hash and block id as block ids can be reused across pages.
    for (const blockId in plugin.settings.taskIdLookup) {
        if (Object.hasOwn(plugin.settings.taskIdLookup, blockId)) {
            // Check if the block is in the metadata cache.
            let found = false;
            let block;
            for (const key in blockCache) {
                if (key.includes(blockId.toLowerCase())) {
                    found = true;
                    block = blockCache[key];
                }
            }

            if (found) {
                logger.info(`Block found in metadata cache: ${blockId}`, block);
            } else {
                logger.info(`Block not found in metadata cache: ${blockId}`);
                // Clean up the block id from the settings.
                delete plugin.settings.taskIdLookup[blockId];
                await plugin.settingsManager.saveSettings();
            }
        }
    }

    logger.info('blockCache', blockCache);
}

/**
 * This will find all block references across all files.
 *
 * @param {MsTodoSync} plugin
 * @return {*}  {Record<string, BlockCache>}
 */
function populateBlockCache(plugin: MsTodoSync): Record<string, BlockCache> {
    const blockCache: Record<string, BlockCache> = {};
    const internalMetadataCache = plugin.app.metadataCache.metadataCache;
    for (const cacheKey in internalMetadataCache) {
        if (Object.hasOwn(internalMetadataCache, cacheKey) && internalMetadataCache[cacheKey].blocks) {
            const blocksCache = internalMetadataCache[cacheKey].blocks;
            for (const blockKey in blocksCache) {
                if (Object.hasOwn(internalMetadataCache, cacheKey)) {
                    const block = blocksCache[blockKey];
                    blockCache[`${cacheKey}-${blockKey}`] = block;
                }
            }
        }
    }

    return blockCache;
}

/**
 * Posts tasks to Microsoft To Do from the selected lines in the editor.
 *
 * @param todoApi - The TodoApi instance used to interact with Microsoft To Do.
 * @param listId - The ID of the list where the tasks will be posted. If undefined, a notice will be shown to set the list name.
 * @param editor - The editor instance from which the tasks will be extracted.
 * @param fileName - The name of the file being edited. If undefined, an empty string will be used.
 * @param plugin - The MsTodoSync plugin instance.
 * @param replace - Optional. If true, the original tasks in the editor will be replaced with the new tasks. Defaults to false.
 *
 * @returns A promise that resolves when the tasks have been posted and the file has been modified.
 */
export async function postTask(
    todoApi: TodoApi,
    listId: string | undefined,
    editor: Editor,
    _fileName: string | undefined,
    plugin: MsTodoSync,
    replace?: boolean,
) {
    const logger = logging.getLogger('mstodo-sync.command.post');

    if (!listId) {
        userNotice.showMessage(t('CommandNotice_SetListName'));
        return;
    }

    const activeFile = plugin.app.workspace.getActiveFile();
    if (activeFile === null) {
        return;
    }

    userNotice.showMessage(t('CommandNotice_CreatingToDo'), 3000);

    const source = await plugin.app.vault.read(activeFile);
    const { lines } = await getCurrentLinesFromEditor(editor);

    // Single call to update the cache using the delta link.
    await getTaskDelta(todoApi, listId, plugin);

    const split = source.split('\n');
    const modifiedPage = await Promise.all(
        split.map(async (line: string, index: number) => {
            // If the line is not in the selection, return the line as is.
            if (!lines.includes(index)) {
                return line;
            }

            // Create the to do task from the line that is in the selection.
            const todo = new ObsidianTodoTask(plugin.settingsManager, line);

            // If there is a block link in the line, we will try to find
            // the task id from the block link and update the task instead.
            // As a user can add a block link, not all tasks will be able to
            // lookup a id from the internal cache.
            if (todo.hasBlockLink && todo.hasId) {
                logger.debug(`Updating Task: ${todo.title}`);

                // Check for linked resource and update if there otherwise create.
                const cachedTasksDelta = await getDeltaCache(plugin);
                const cachedTask = cachedTasksDelta?.allTasks.find((task) => task.id === todo.id);
                if (cachedTask) {
                    const linkedResource = cachedTask.linkedResources?.first();
                    if (linkedResource && linkedResource.id) {
                        await todoApi.updateLinkedResource(
                            listId,
                            todo.id,
                            linkedResource.id,
                            todo.blockLink ?? '',
                            todo.getRedirectUrl(),
                        );
                    } else {
                        await todoApi.createLinkedResource(
                            listId,
                            todo.id,
                            todo.blockLink ?? '',
                            todo.getRedirectUrl(),
                        );
                    }
                }

                todo.linkedResources = cachedTask?.linkedResources;

                const returnedTask = await todoApi.updateTaskFromToDo(listId, todo.id, todo.getTodoTask());
                logger.debug(`blockLink: ${todo.blockLink}, taskId: ${todo.id}`);
                logger.debug(`updated: ${returnedTask.id}`);
            } else {
                logger.debug(`Creating Task: ${todo.title}`);
                logger.debug(`Creating Task: ${listId}`);

                const returnedTask = await todoApi.createTaskFromToDo(listId, todo.getTodoTask());

                todo.status = returnedTask.status;
                await todo.cacheTaskId(returnedTask.id ?? '');

                // Create linkedResource after task is created and blockLink is assigned
                // (blockLink didn't exist before cacheTaskId() was called)
                if (todo.blockLink && returnedTask.id) {
                    try {
                        await todoApi.createLinkedResource(
                            listId,
                            returnedTask.id,
                            todo.blockLink,
                            todo.getRedirectUrl(),
                        );
                        logger.debug(`Created linkedResource for task: ${returnedTask.id}`);
                    } catch (error) {
                        logger.warn(`Failed to create linkedResource: ${error}`);
                        // Don't fail the whole operation if linkedResource creation fails
                    }
                }

                logger.debug(`blockLink: ${todo.blockLink}, taskId: ${todo.id}`, todo);
            }

            // If false there will be a orphaned block id for this task.
            if (replace) {
                return todo.getMarkdownTask(true);
            }

            return line;
        }),
    );

    await plugin.app.vault.modify(activeFile, modifiedPage.join('\n'));
}

export async function getTask(
    todoApi: TodoApi,
    listId: string | undefined,
    editor: Editor,
    _fileName: string | undefined,
    plugin: MsTodoSync,
) {
    const logger = logging.getLogger('mstodo-sync.command.get');

    if (!listId) {
        userNotice.showMessage(t('CommandNotice_SetListName'));
        return;
    }

    const activeFile = plugin.app.workspace.getActiveFile();
    if (activeFile === null) {
        return;
    }

    userNotice.showMessage(t('CommandNotice_GettingToDo'), 3000);

    const source = await plugin.app.vault.read(activeFile);
    const { lines } = await getCurrentLinesFromEditor(editor);

    // Single call to update the cache using the delta link.
    await getTaskDelta(todoApi, listId, plugin);

    const split = source.split('\n');
    const modifiedPage = await Promise.all(
        split.map(async (line: string, index: number) => {
            // If the line is not in the selection, return the line as is.
            if (!lines.includes(index)) {
                return line;
            }

            // Create the to do task from the line that is in the selection.
            const todo = new ObsidianTodoTask(plugin.settingsManager, line);

            // If there is a block link in the line, we will try to find
            // the task id from the block link and update the task instead.
            // As a user can add a block link, not all tasks will be able to
            // lookup a id from the internal cache.
            if (todo.hasBlockLink && todo.hasId) {
                logger.debug(`Updating Task: ${todo.title}`);

                // Load from the delta cache file and pull the task from the cache.
                const cachedTasksDelta = await getDeltaCache(plugin);
                const returnedTask = cachedTasksDelta?.allTasks.find((task) => task.id === todo.id);

                if (returnedTask) {
                    todo.updateFromTodoTask(returnedTask);
                    logger.debug(`blockLink: ${todo.blockLink}, taskId: ${todo.id}`);
                    logger.debug(`updated: ${returnedTask.id}`);
                }

                return todo.getMarkdownTask(true);
            }

            return line;
        }),
    );

    await plugin.app.vault.modify(activeFile, modifiedPage.join('\n'));
}

async function getDeltaCache(plugin: MsTodoSync) {
    const cachePath = `${plugin.manifest.dir}/mstd-tasks-delta.json`;
    const adapter: DataAdapter = plugin.app.vault.adapter;
    let cachedTasksDelta: TasksDeltaCollection | undefined;

    if (await adapter.exists(cachePath)) {
        cachedTasksDelta = JSON.parse(await adapter.read(cachePath)) as TasksDeltaCollection;
    }

    return cachedTasksDelta;
}

export async function getTaskDelta(todoApi: TodoApi, listId: string | undefined, plugin: MsTodoSync, reset = false) {
    const logger = logging.getLogger('mstodo-sync.command.delta');

    if (!listId) {
        userNotice.showMessage(t('CommandNotice_SetListName'));
        return;
    }

    const cachePath = `${plugin.manifest.dir}/mstd-tasks-delta.json`;
    const adapter: DataAdapter = plugin.app.vault.adapter;
    if (reset) {
        await adapter.remove(cachePath);
    }

    let deltaLink = '';
    let cachedTasksDelta = await getDeltaCache(plugin);

    if (cachedTasksDelta) {
        deltaLink = cachedTasksDelta.deltaLink;
    } else {
        cachedTasksDelta = new TasksDeltaCollection([], '', '', '');
    }

    const returnedTask = await todoApi.getTasksDelta(listId, deltaLink);
    logger.info('deltaLink', deltaLink);
    logger.info('ReturnedDelta', returnedTask);

    if (cachedTasksDelta) {
        logger.info('cachedTasksDelta.allTasks', cachedTasksDelta.allTasks.length);
        logger.info('returnedTask.allTasks', returnedTask.allTasks.length);

        cachedTasksDelta.allTasks = mergeCollections(cachedTasksDelta.allTasks, returnedTask.allTasks);
        logger.info('cachedTasksDelta.allTasks', cachedTasksDelta.allTasks.length);

        cachedTasksDelta.deltaLink = returnedTask.deltaLink;
    } else {
        logger.info('First run, loading delta cache');

        cachedTasksDelta = returnedTask;
    }

    await adapter.write(cachePath, JSON.stringify(cachedTasksDelta));
}

// Function to merge collections
function mergeCollections(col1: TodoTask[], col2: TodoTask[]): TodoTask[] {
    const map = new Map<string, TodoTask>();

    // Helper function to add items to the map
    function addToMap(item: TodoTask) {
        if (item.id && item.lastModifiedDateTime) {
            const existingItem = map.get(item.id);
            // If there is no last modified then just use the current item.
            if (
                !existingItem ||
                new Date(item.lastModifiedDateTime) > new Date(existingItem.lastModifiedDateTime ?? 0)
            ) {
                map.set(item.id, item);
            }
        }
    }

    // Add items from both collections to the map
    for (const item of col1) {
        addToMap(item);
    }

    for (const item of col2) {
        addToMap(item);
    }

    // Convert map values back to an array
    return Array.from(map.values());
}

// Experimental
// Should handle the following cases:
// - [ ] Task
// - [ ] Task with indented note
//   note
// - [ ] Task with subtasks
//   - [ ] Task One
//   - [ ] Task Two
// - [ ] Task with subtasks and notes
//   Need to think about this one. Perhaps a task 3?
//   - [ ] Task One
//   - [ ] Task Two
// Lines are processed until the next line is blank or not indented by two spaces.
// Also EOF will stop processing.
// Allow variable depth or match column of first [
export async function postTaskAndChildren(
    todoApi: TodoApi,
    listId: string | undefined,
    editor: Editor,
    _fileName: string | undefined,
    plugin: MsTodoSync,
    push = true,
) {
    const logger = logging.getLogger('mstodo-sync.command.post');

    if (!listId) {
        userNotice.showMessage(t('CommandNotice_SetListName'));
        return;
    }

    userNotice.showMessage(t('CommandNotice_CreatingToDo'), 3000);

    const cursorLocation = editor.getCursor();
    const topLevelTask = editor.getLine(cursorLocation.line);
    logger.debug(`topLevelTask: ${topLevelTask}`);
    // Logger.debug(`cursorLocation: ${cursorLocation.line}`, cursorLocation);

    let body = '';
    const childTasks: string[] = [];

    // Get all lines including the line the cursor is on.
    const lines = editor.getValue().split('\n').slice(cursorLocation.line);
    // Logger.debug(`editor: ${cursorLocation}`, lines);

    // Find the end of section which a blank line or a line that is not indented by two spaces.
    const endLine = lines.findIndex(
        // (line, index) => !/[ ]{2,}- \[(.)\]/.test(line) && !line.startsWith('  ') && index > 0,
        (line, index) => line.length === 0 && index > 0,
    );
    logger.debug(`endLine: ${endLine}`);

    // Scan lines below task for sub tasks and body.
    for (const [_index, line] of lines.slice(1, endLine).entries()) {
        // Logger.debug(`processing line: ${index} -- ${line}`);

        if (line.startsWith('  - [')) {
            childTasks.push(line.trim());
        } else {
            // Remove the two spaces at the beginning of the line, will be added back on sync.
            // on sync the body will be indented by two spaces and the tasks will be appended at this point.
            body += line.trim() + '\n';
        }
    }

    logger.debug(`body: ${body}`);
    logger.debug(`childTasks: ${childTasks}`, childTasks);

    const todo = new ObsidianTodoTask(plugin.settingsManager, topLevelTask);
    todo.setBody(body);
    for (const childTask of childTasks) {
        todo.addChecklistItem(childTask);
    }

    logger.debug(`updated: ${todo.title}`, todo);

    if (todo.hasBlockLink && todo.id) {
        logger.debug(`Updating Task: ${todo.title}`, todo.getTodoTask());

        // Const currentTaskState = await todoApi.getTask(listId, todo.id);
        let returnedTask;
        if (push) {
            returnedTask = await todoApi.updateTaskFromToDo(listId, todo.id, todo.getTodoTask());
            // Push the checklist items...
            todo.checklistItems = returnedTask.checklistItems;
            todo.status = returnedTask.status;
            todo.body = returnedTask.body;
        } else {
            // Pull: Load task WITH details (checklistItems, linkedResources)
            returnedTask = await todoApi.getTask(listId, todo.id, true);
            if (returnedTask) {
                todo.checklistItems = returnedTask.checklistItems;
                todo.status = returnedTask.status;
                todo.body = returnedTask.body;
            }
        }

        logger.debug(`blockLink: ${todo.blockLink}, taskId: ${todo.id}`);
        logger.debug(`updated: ${returnedTask?.id}`, returnedTask);
    } else {
        logger.debug(`Creating Task: ${todo.title}`);

        const returnedTask = await todoApi.createTaskFromToDo(listId, todo.getTodoTask(true));

        todo.status = returnedTask.status;
        await todo.cacheTaskId(returnedTask.id ?? '');

        // Create linkedResource after task is created and blockLink is assigned
        // (blockLink didn't exist before cacheTaskId() was called)
        if (todo.blockLink && returnedTask.id) {
            try {
                await todoApi.createLinkedResource(listId, returnedTask.id, todo.blockLink, todo.getRedirectUrl());
                logger.debug(`Created linkedResource for task: ${returnedTask.id}`);
            } catch (error) {
                logger.warn(`Failed to create linkedResource: ${error}`);
                // Don't fail the whole operation if linkedResource creation fails
            }
        }

        logger.debug(`blockLink: ${todo.blockLink}, taskId: ${todo.id}`, todo);
    }

    // Update the task on the page.
    const start = getLineStartPos(cursorLocation.line);
    const end = getLineEndPos(cursorLocation.line + endLine, editor);

    editor.replaceRange(todo.getMarkdownTask(false), start, end);
}

function getLineStartPos(line: number): EditorPosition {
    return {
        line,
        ch: 0,
    };
}

function getLineEndPos(line: number, editor: Editor): EditorPosition {
    return {
        line,
        ch: editor.getLine(line).length,
    };
}

export async function getAllTasksInList(
    todoApi: TodoApi,
    listId: string | undefined,
    editor: Editor,
    plugin: MsTodoSync,
    withBody: boolean,
) {
    const now = globalThis.moment();
    const settings = plugin.settingsManager.settings;

    if (!listId) {
        userNotice.showMessage(t('CommandNotice_SetListName'));
        return;
    }

    // Single call to update the cache using the delta link.
    await getTaskDelta(todoApi, listId, plugin);
    const cachedTasksDelta = await getDeltaCache(plugin);

    cachedTasksDelta?.allTasks.sort((a, _b) => (a.status === 'completed' ? 1 : -1));

    const lines = cachedTasksDelta?.allTasks
        ?.filter((task) => task.status !== 'completed')
        .map((task) => {
            const formattedCreateDate = globalThis
                .moment(task.createdDateTime)
                .format(settings.displayOptions_DateFormat);
            const done = task.status === 'completed' ? 'x' : ' ';
            const createDate =
                formattedCreateDate === now.format(settings.displayOptions_DateFormat)
                    ? ''
                    : `${settings.displayOptions_TaskCreatedPrefix}[[${formattedCreateDate}]]`;

            let blockId = '';
            for (const key in settings.taskIdLookup) {
                if (Object.hasOwn(plugin.settings.taskIdLookup, key) && settings.taskIdLookup[key] === task.id) {
                    blockId = `^${key}`;
                }
            }

            if (blockId === '') {
                const newId = cacheTaskId(task.id ?? '', plugin.settingsManager);
                blockId = `^${newId}`;
            }

            if (task.body?.content && withBody) {
                // If the body has multiple lines then indent slightly on a new line.
                const bodyLines = task.body.content.split('\r\n');
                const newBody = bodyLines.map((line, _index) => `  ${stripHtml(line).trimEnd()}`);
                return `- [${done}] ${task.title}  ${createDate} ${blockId}\n${newBody.join('\n')}`.trimEnd();
            }

            return `- [${done}] ${task.title}  ${createDate} ${blockId}`.trimEnd();
        });

    const allTasks = lines?.join('\n');

    if (editor) {
        editor.replaceSelection(allTasks ?? '');
        const view = plugin.app.workspace.getActiveViewOfType(MarkdownView);
        view?.leaf.view.tree.setCollapseAll(true);

        // GetActiveViewOfType will return null if the active view is null, or if it's not a MarkdownView.
        if (view?.tree) {
            view.tree.setCollapseAll(true);
            // ...
        }
    }
}

/**
 * Cache the ID internally and generate block link.
 *
 * @param {string} [id]
 * @return {*}  {Promise<void>}
 * @memberof ObsidianTodoTask
 */
async function cacheTaskId(id: string, settingsManager: SettingsManager): Promise<string> {
    settingsManager.settings.taskIdIndex += 1;

    const index = `MSTD${Math.random().toString(20).slice(2, 6)}${settingsManager.settings.taskIdIndex
        .toString()
        .padStart(5, '0')}`;

    settingsManager.settings.taskIdLookup[index] = id ?? '';

    settingsManager.saveSettings().catch((error) => {
        console.error('Error saving settings', error);
    });

    return index;
}

function stripHtml(html: string): string {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    return doc.body.textContent || '';
}

/**
 * Inserts tasks from Microsoft To Do into the editor.
 *
 * @param todoApi - The TodoApi instance
 * @param plugin - The plugin instance (needed for settings and caching)
 * @param editor - Optional editor to insert into
 * @param filterListName - Optional: Only insert tasks from this specific list. If undefined, uses configured list or all lists.
 */
export async function createTodayTasks(todoApi: TodoApi, plugin: MsTodoSync, editor?: Editor, filterListName?: string) {
    const settings = plugin.settingsManager.settings;
    userNotice.showMessage('Getting Microsoft To Do tasks...', 3000);
    const now = globalThis.moment();
    const pattern = `status ne 'completed' or completedDateTime/dateTime ge '${now.format('YYYY-MM-DD')}'`;

    // Fetch lists WITHOUT tasks first to avoid 429 (Too Many Requests)
    // getLists() without pattern only fetches list metadata
    const allLists = await todoApi.getLists();

    if (!allLists || allLists.length === 0) {
        userNotice.showMessage('Task list is empty');
        return;
    }

    // Determine which list(s) to include
    // If filterListName is provided (string), use it.
    // If it is empty string "", it means "All Lists".
    // If it is undefined/null, fallback to settings default.
    let targetListName = filterListName;
    if (targetListName === undefined || targetListName === null) {
        targetListName = settings.todoListSync.listName;
    }

    // If targetListName is "", it means NO filter (All Lists)
    const useFilter = targetListName !== '';

    const filteredLists = allLists.filter((taskList) => {
        // If a target list is specified, only include that list
        if (useFilter) {
            return taskList.displayName === targetListName;
        }
        // Otherwise include all lists
        return true;
    });

    const segmentPromises = filteredLists.map(async (taskList) => {
        // Fetch tasks for this specific list
        // We do this inside the map/loop now, instead of fetching ALL tasks for ALL lists at start
        try {
            taskList.tasks = await todoApi.getListTasks(taskList.id, pattern);
        } catch (e) {
            console.error(`Error fetching tasks for list ${taskList.displayName}`, e);
            taskList.tasks = [];
        }

        if (!taskList.tasks || taskList.tasks.length === 0) {
            return;
        }

        taskList.tasks.sort((a, _b) => (a.status === 'completed' ? 1 : -1));
        const lines = await Promise.all(
            taskList.tasks?.map(async (task) => {
                const formattedCreateDate = globalThis
                    .moment(task.createdDateTime)
                    .format(settings.displayOptions_DateFormat);
                const done = task.status === 'completed' ? 'x' : ' ';
                const createDate =
                    formattedCreateDate === now.format(settings.displayOptions_DateFormat)
                        ? ''
                        : `${settings.displayOptions_TaskCreatedPrefix}[[${formattedCreateDate}]]`;
                const body = task.body?.content ? `${settings.displayOptions_TaskBodyPrefix}${task.body.content}` : '';

                // Generate subtasks
                let subtasks = '';
                if (task.checklistItems && task.checklistItems.length > 0) {
                    task.checklistItems.sort(
                        (a, b) =>
                            new Date(a.createdDateTime as string).getTime() -
                            new Date(b.createdDateTime as string).getTime(),
                    );
                    subtasks = task.checklistItems
                        .map((item) => {
                            const subDone = item.isChecked ? 'x' : ' ';
                            return `\n  - [${subDone}] ${item.displayName}`;
                        })
                        .join('');
                }

                // Generate block ID for tracking (enables two-way sync)
                let blockId = '';
                // Check if task is already tracked
                for (const key in settings.taskIdLookup) {
                    if (Object.hasOwn(settings.taskIdLookup, key) && settings.taskIdLookup[key] === task.id) {
                        blockId = `^${key}`;
                        break;
                    }
                }
                // If not tracked, create new tracking ID
                if (blockId === '' && task.id) {
                    const newId = await cacheTaskId(task.id, plugin.settingsManager);
                    blockId = `^${newId}`;
                }

                return `- [${done}] ${task.title}  ${createDate} ${blockId} ${body}${subtasks}`.trimEnd();
            }) ?? [],
        );

        return `**${taskList.displayName}**
${lines?.join('\n')}
`;
    });

    const segmentsArray = await Promise.all(segmentPromises);
    const segments = segmentsArray.filter((s) => s !== undefined).join('\n\n');

    if (editor) {
        editor.replaceSelection(segments);
    } else {
        return segments;
    }
}
