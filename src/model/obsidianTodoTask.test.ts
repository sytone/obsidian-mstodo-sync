/* eslint-disable no-undef */
import { ObsidianTodoTask } from './obsidianTodoTask';
import { type ISettingsManager } from 'src/utils/settingsManager';
import { TodoTask } from '@microsoft/microsoft-graph-types';

describe('ObsidianTodoTask', () => {
    let settingsManager: ISettingsManager;

    beforeEach(() => {
        settingsManager = {
            settings: {
                displayOptions_TaskCreatedPrefix: '🔎',
                displayOptions_TaskDuePrefix: '📅',
                displayOptions_TaskImportance_Low: '🔽',
                displayOptions_TaskImportance_Normal: '🔼',
                displayOptions_TaskImportance_High: '⏫',
                displayOptions_TaskStatus_NotStarted: ' ',
                displayOptions_TaskStatus_InProgress: '/',
                displayOptions_TaskStatus_Completed: 'x',
                displayOptions_ReplacementFormat:
                    '- [{{STATUS_SYMBOL}}] {{TASK}}{{IMPORTANCE}}{{TASK_LIST_NAME}}{{DUE_DATE}}{{CREATED_DATE}}',
                displayOptions_ListIndicator: '+',
                displayOptions_RegExToRunOnPushAgainstTitle: '',
                microsoftToDoApplication_RedirectUriBase: 'https://todo.microsoft.com',
                taskIdIndex: 0,
                taskIdLookup: {},
                todoListSync: {
                    listName: 'ToDo',
                    listId: undefined,
                },
                diary: {
                    folder: '',
                    format: '',
                    stayWithPN: false,
                },
                displayOptions_DateFormat: 'YYYY-MM-DD',
                displayOptions_TimeFormat: 'HH:mm',
                displayOptions_TaskStartPrefix: '🛫',
                displayOptions_TaskBodyPrefix: '💡',
                displayOptions_ReplaceAddCreatedAt: false,
                todo_OpenUsingApplicationProtocol: false,
                microsoft_AuthenticationClientId: '',
                microsoft_AuthenticationAuthority: '',
                loggingOptions: {
                    minLevels: {
                        '': 'debug',
                        'mstodo-sync': 'debug',
                    },
                },
                hackingEnabled: false,
                displayOptions_ListIndicator_UseSingleQuotes: false,
                todo_CreateToDoListIfMissing: false,
            },
            vaultName: 'test-vault',
            saveSettings: jest.fn(),
        };
    });

    it('should create an instance of ObsidianTodoTask', () => {
        const line = '- [ ] Test task';
        const task = new ObsidianTodoTask(settingsManager, line);
        expect(task).toBeInstanceOf(ObsidianTodoTask);
        expect(task.title).toBe('Test task');
    });

    it('should strip out block link from title', () => {
        const line = '- [ ] Test task ^blockLink';
        const task = new ObsidianTodoTask(settingsManager, line);
        expect(task.title).toBe('Test task');
        expect(task.blockLink).toBe('blockLink');
    });

    it('should strip out status from title', () => {
        const line = '- [ ] Test task';
        const task = new ObsidianTodoTask(settingsManager, line);
        expect(task.title).toBe('Test task');
        expect(task.status).toBe('notStarted');
    });

    it('should strip out created date from title', () => {
        const line = '- [ ] Test task 🔎2020-01-01';
        const task = new ObsidianTodoTask(settingsManager, line);
        expect(task.title).toBe('Test task');
    });

    it('should strip out due date from title and set dueDateTime', () => {
        const line = '- [ ] Test task 📅2020-01-01';
        const task = new ObsidianTodoTask(settingsManager, line);
        expect(task.title).toBe('Test task');
        expect(task.dueDateTime?.dateTime).toBe('2020-01-01');
    });

    it('should set importance based on title', () => {
        const line = '- [ ] Test task ⏫';
        const task = new ObsidianTodoTask(settingsManager, line);
        expect(task.importance).toBe('high');
    });

    it('should cache task ID and generate block link', async () => {
        const line = '- [ ] Test task';
        const task = new ObsidianTodoTask(settingsManager, line);
        await task.cacheTaskId('12345');
        expect(task.id).toBe('12345');
        expect(task.blockLink).toMatch(/^MSTD/);
        expect(settingsManager.saveSettings).toHaveBeenCalled();
    });

    it('should check equality of tasks', () => {
        const line = '- [ ] Test task';
        const task = new ObsidianTodoTask(settingsManager, line);
        const todoTask: TodoTask = {
            title: 'Test task',
            status: 'notStarted',
            dueDateTime: {
                dateTime: '2020-01-01',
                timeZone: 'UTC',
            },
            importance: 'normal',
        };
        task.dueDateTime = {
            dateTime: '2020-01-01',
            timeZone: 'UTC',
        };
        expect(task.equals(todoTask)).toBe(true);
    });

    it('should get TodoTask object', () => {
        const line = '- [ ] Test task';
        const task = new ObsidianTodoTask(settingsManager, line);
        const todoTask = task.getTodoTask();
        expect(todoTask.title).toBe('Test task');
    });

    it('should update from TodoTask object', () => {
        const line = '- [ ] Test task';
        const task = new ObsidianTodoTask(settingsManager, line);
        const remoteTask: TodoTask = {
            title: 'Updated task',
            body: {
                content: 'Updated content',
                contentType: 'text',
            },
            status: 'inProgress',
            importance: 'high',
            linkedResources: [],
            dueDateTime: {
                dateTime: '2020-01-01',
                timeZone: 'UTC',
            },
        };
        task.updateFromTodoTask(remoteTask);
        expect(task.title).toBe('Updated task');
        expect(task.body?.content).toBe('Updated content');
        expect(task.status).toBe('inProgress');
        expect(task.importance).toBe('high');
        expect(task.dueDateTime?.dateTime).toBe('2020-01-01');
    });

    it('should set body content', () => {
        const line = '- [ ] Test task';
        const task = new ObsidianTodoTask(settingsManager, line);
        task.setBody('New body content');
        expect(task.body?.content).toBe('New body content');
    });

    it('should add checklist item', () => {
        const line = '- [ ] Test task';
        const task = new ObsidianTodoTask(settingsManager, line);
        task.addChecklistItem('New checklist item');
        expect(task.checklistItems?.length).toBe(1);
        expect(task.checklistItems?.[0].displayName).toBe('New checklist item');
    });

    it('should return markdown task', () => {
        const line = '- [ ] Test task';
        const task = new ObsidianTodoTask(settingsManager, line);
        const markdownTask = task.getMarkdownTask(true);
        expect(markdownTask).toBe('- [ ] Test task 🔎2025-01-10');
    });

    const dataSet = [
        {
            taskLine: '- [ ] Test task +"Test list" 🔎2025-01-10',
            listName: 'Test list',
            renderedLine: '- [ ] Test task +"Test list" 🔎2025-01-10',
        },
        {
            taskLine: "- [ ] Test task +'Test list' 🔎2025-01-10",
            listName: 'Test list',
            renderedLine: '- [ ] Test task +"Test list" 🔎2025-01-10',
        },
        {
            taskLine: "- [ ] Test task +'Testlist' 🔎2025-01-10",
            listName: 'Testlist',
            renderedLine: '- [ ] Test task +Testlist 🔎2025-01-10',
        },
        {
            taskLine: '- [ ] Test task +🔧MyList 🔎2025-01-10',
            listName: '🔧MyList',
            renderedLine: '- [ ] Test task +🔧MyList 🔎2025-01-10',
        },
        {
            taskLine: "- [ ] Test task +'🔧 MyList' 🔎2025-01-10",
            listName: '🔧 MyList',
            renderedLine: '- [ ] Test task +"🔧 MyList" 🔎2025-01-10',
        },
    ];

    it.each(dataSet)(
        'should return markdown task with list indicator ($taskLine, $listName, $renderedLine)',
        ({ taskLine, listName, renderedLine }) => {
            const task = new ObsidianTodoTask(settingsManager, taskLine);
            task.listId = '12345';
            const markdownTask = task.getMarkdownTask(true);
            expect(task.listName).toBe(listName);
            expect(markdownTask).toContain('Test task');
            expect(markdownTask).toBe(`${renderedLine}`);
        },
    );

    it('should set list name to default value', () => {
        const line = '- [ ] Test task';
        const task = new ObsidianTodoTask(settingsManager, line);
        expect(task.listName).toBe('ToDo');
        const markdownTask = task.getMarkdownTask(true);
        expect(markdownTask).toContain('+ToDo');
    });

    it('should return markdown task with list indicator', () => {
        const line = '- [ ] Test task +"Test list"';
        const task = new ObsidianTodoTask(settingsManager, line);
        task.listId = '12345';
        const markdownTask = task.getMarkdownTask(true);
        expect(task.listName).toBe('Test list');
        expect(markdownTask).toContain('Test task');
        expect(markdownTask).toContain('+"Test list"');
    });

    it('should return markdown task with list indicator using single quotes', () => {
        const line = '- [ ] Test task +"Test list"';
        settingsManager.settings.displayOptions_ListIndicator_UseSingleQuotes = true;
        const task = new ObsidianTodoTask(settingsManager, line);
        task.listId = '12345';
        const markdownTask = task.getMarkdownTask(true);
        expect(task.listName).toBe('Test list');
        expect(markdownTask).toContain('Test task');
        expect(markdownTask).toContain("+'Test list'");
    });

    it('should return markdown task with list indicator unquoted from quoted word', () => {
        const line = '- [ ] Test task +"testListName"';
        settingsManager.settings.displayOptions_ListIndicator_UseSingleQuotes = true;
        const task = new ObsidianTodoTask(settingsManager, line);
        task.listId = '12345';
        const markdownTask = task.getMarkdownTask(true);
        expect(task.listName).toBe('testListName');
        expect(markdownTask).toContain('Test task');
        expect(markdownTask).toContain('+testListName');
    });

    it('should return markdown task with list indicator unquoted from unquoted word', () => {
        const line = '- [ ] Test task +testListName';
        settingsManager.settings.displayOptions_ListIndicator_UseSingleQuotes = true;
        const task = new ObsidianTodoTask(settingsManager, line);
        task.listId = '12345';
        const markdownTask = task.getMarkdownTask(true);
        expect(task.listName).toBe('testListName');
        expect(markdownTask).toContain('Test task');
        expect(markdownTask).toContain('+testListName');
    });
});
