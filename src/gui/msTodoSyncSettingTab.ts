import { type App, type PeriodicNotes, PluginSettingTab, Setting } from 'obsidian';
import type MsTodoSync from 'src/main.js';
import { t } from 'src/lib/lang.js';
import { type ILogOptions } from 'src/lib/logging.js';
import type { IUserNotice } from 'src/lib/userNotice.js';

export interface IMsTodoSyncSettings {
    todoListSync: {
        listName: string | undefined;
        listId: string | undefined;
    };

    todo_CreateToDoListIfMissing: boolean;

    diary: {
        folder: string;
        format: string;
        stayWithPN: boolean;
    };

    displayOptions_DateFormat: string;
    displayOptions_TimeFormat: string;
    displayOptions_RegExToRunOnPushAgainstTitle: string;
    displayOptions_TaskCreatedPrefix: string;
    displayOptions_TaskDuePrefix: string;
    displayOptions_TaskStartPrefix: string;
    displayOptions_TaskBodyPrefix: string;
    displayOptions_ReplaceAddCreatedAt: boolean;
    displayOptions_ReplacementFormat: string;

    // Importance
    // The importance of the task. Possible values
    // are: low, normal, high.
    // By default it is normal and the absence of a
    // indicator will also mean normal.
    displayOptions_TaskImportance_Low: string;
    displayOptions_TaskImportance_Normal: string;
    displayOptions_TaskImportance_High: string;

    displayOptions_TaskStatus_NotStarted: string;
    displayOptions_TaskStatus_InProgress: string;
    displayOptions_TaskStatus_Completed: string;

    displayOptions_ListIndicator: string;
    displayOptions_ListIndicator_UseSingleQuotes: boolean;

    // Microsoft To Do open handler.
    todo_OpenUsingApplicationProtocol: boolean;

    microsoft_AuthenticationClientId: string;
    microsoft_AuthenticationAuthority: string;

    // Logging options.
    loggingOptions: ILogOptions;

    // Private configuration updated by the plugin and not user.
    taskIdLookup: Record<string, string>;
    taskIdIndex: number;
    hackingEnabled: boolean;
    microsoftToDoApplication_RedirectUriBase: string;
}

export const DEFAULT_SETTINGS: IMsTodoSyncSettings = {
    todoListSync: {
        listName: undefined,
        listId: undefined,
    },
    todo_CreateToDoListIfMissing: true,
    diary: {
        folder: '',
        format: '',
        stayWithPN: false,
    },
    displayOptions_DateFormat: 'YYYY-MM-DD',
    displayOptions_TimeFormat: 'HH:mm',
    displayOptions_RegExToRunOnPushAgainstTitle: '',
    displayOptions_TaskCreatedPrefix: '🔎',
    displayOptions_TaskDuePrefix: '📅',
    displayOptions_TaskStartPrefix: '🛫',
    displayOptions_TaskBodyPrefix: '💡',
    displayOptions_ReplaceAddCreatedAt: false,
    displayOptions_ReplacementFormat:
        '- [{{STATUS_SYMBOL}}] {{TASK}}{{IMPORTANCE}}{{TASK_LIST_NAME}}{{DUE_DATE}}{{CREATED_DATE}}',

    displayOptions_TaskImportance_Low: '🔽',
    displayOptions_TaskImportance_Normal: '🔼',
    displayOptions_TaskImportance_High: '⏫',

    displayOptions_TaskStatus_NotStarted: ' ',
    displayOptions_TaskStatus_InProgress: '/',
    displayOptions_TaskStatus_Completed: 'x',

    displayOptions_ListIndicator: '+',
    displayOptions_ListIndicator_UseSingleQuotes: false,

    todo_OpenUsingApplicationProtocol: true,

    loggingOptions: {
        minLevels: {
            '': 'debug',
            'mstodo-sync': 'debug',
        },
    },
    taskIdLookup: { '0000ABCD': '0' },
    taskIdIndex: 0,
    microsoft_AuthenticationClientId: '',
    microsoft_AuthenticationAuthority: '',
    hackingEnabled: false,
    microsoftToDoApplication_RedirectUriBase: 'http://192.168.0.137:8901/redirectpage.html',
};

export class MsTodoSyncSettingTab extends PluginSettingTab {
    plugin: MsTodoSync;
    settings: IMsTodoSyncSettings;
    userNotice: IUserNotice;

    constructor(app: App, plugin: MsTodoSync, userNotice: IUserNotice) {
        super(app, plugin);
        this.plugin = plugin;
        this.settings = plugin.settings;
        this.userNotice = userNotice;
    }

    /**
     * Creates a setting entry in the settings form
     * for text based properties. If there is a update
     * it will save the new value.
     *
     * @param {HTMLElement} containerElement
     * @param {string} title
     * @param {string} description
     * @param {string} currentValue
     * @param {(value: string) => unknown} changeCallback
     * @memberof MsTodoSyncSettingTab
     */
    addTextSetting(
        containerElement: HTMLElement,
        title: string,
        description: string,
        currentValue: string,
        changeCallback: (_value: string) => unknown,
    ): void {
        new Setting(containerElement)
            .setName(t(title))
            .setDesc(t(description))
            .addText((text) =>
                text.setValue(currentValue).onChange(async (value) => {
                    changeCallback(value);
                    await this.plugin.saveSettings();
                }),
            );
    }

    display(): void {
        const { containerEl } = this;

        containerEl.empty();

        containerEl.createEl('h2', {
            text: `${this.plugin.manifest.name}`,
        });
        const span = containerEl.createSpan();
        span.style.fontSize = '0.8em';
        span.innerHTML = `Version ${this.plugin.manifest.version} <br /> ${this.plugin.manifest.description} created by ${this.plugin.manifest.author}`;

        new Setting(containerEl)
            .setName(t('Settings_Todo_DefaultListName'))
            .setDesc(t('Settings_Todo_DefaultListNameDescription'))
            .addText((text) =>
                text
                    // .setPlaceholder('输入Todo列表名称')
                    .setValue(this.settings.todoListSync.listName ?? '')
                    .onChange(async (value) => {
                        this.settings.todoListSync.listName = value;
                    }),
            );

        new Setting(containerEl)
            .setName(t('Settings_Todo_OpenUsingApplicationProtocolTitle'))
            .setDesc(t('Settings_Todo_OpenUsingApplicationProtocolDescription'))
            .addToggle((toggle) =>
                toggle.setValue(this.settings.todo_OpenUsingApplicationProtocol).onChange(async (value) => {
                    this.settings.todo_OpenUsingApplicationProtocol = value;
                    await this.plugin.saveSettings();
                }),
            );

        // Formatting Options that user can set
        containerEl.createEl('h2', {
            text: t('Settings_Todo_Display_Heading'),
        });

        new Setting(containerEl)
            .setName(t('Settings_Todo_Display_DateFormat'))
            .setDesc(t('Settings_Todo_Display_DateFormatDescription'))
            .addText((text) =>
                text.setValue(this.settings.displayOptions_DateFormat ?? '').onChange(async (value) => {
                    this.settings.displayOptions_DateFormat = value;
                    await this.plugin.saveSettings();
                }),
            );

        new Setting(containerEl)
            .setName(t('Settings_Todo_Display_TimeFormat'))
            .setDesc(t('Settings_Todo_Display_TimeFormatDescription'))
            .addText((text) =>
                text.setValue(this.settings.displayOptions_TimeFormat ?? '').onChange(async (value) => {
                    this.settings.displayOptions_TimeFormat = value;
                    await this.plugin.saveSettings();
                }),
            );

        this.addTextSetting(
            containerEl,
            'Settings_Todo_Display_RegExToRunOnPushAgainstTitle',
            'Settings_Todo_Display_RegExToRunOnPushAgainstTitle_Description',
            this.settings.displayOptions_RegExToRunOnPushAgainstTitle,
            async (value) => {
                this.settings.displayOptions_RegExToRunOnPushAgainstTitle = value;
            },
        );

        new Setting(containerEl)
            .setName(t('Settings_Todo_Display_AddCreatedAtOnReplace'))
            .setDesc(t('Settings_Todo_Display_AddCreatedAtOnReplaceDescription'))
            .addToggle((toggle) =>
                toggle.setValue(this.settings.displayOptions_ReplaceAddCreatedAt).onChange(async (value) => {
                    this.settings.displayOptions_ReplaceAddCreatedAt = value;
                    await this.plugin.saveSettings();
                }),
            );

        // Replacement Format: default: - [ ] {{TASK}}
        new Setting(containerEl)
            .setName(t('Settings_Todo_Display_ReplacementFormat'))
            .setDesc(t('Settings_Todo_Display_ReplacementFormatDescription'))
            .addText((text) =>
                text.setValue(this.settings.displayOptions_ReplacementFormat).onChange(async (value) => {
                    this.settings.displayOptions_ReplacementFormat = value;
                    await this.plugin.saveSettings();
                }),
            );

        // Task Importance Indicators - High
        this.addTextSetting(
            containerEl,
            'Settings_Todo_Display_Importance_HighName',
            'Settings_Todo_Display_Importance_HighDescription',
            this.settings.displayOptions_TaskImportance_High,
            async (value) => {
                this.settings.displayOptions_TaskImportance_High = value;
            },
        );

        // Task Importance Indicators - Normal
        this.addTextSetting(
            containerEl,
            'Settings_Todo_Display_Importance_NormalName',
            'Settings_Todo_Display_Importance_NormalDescription',
            this.settings.displayOptions_TaskImportance_Normal,
            async (value) => {
                this.settings.displayOptions_TaskImportance_Normal = value;
            },
        );

        // Task Importance Indicators - Low
        this.addTextSetting(
            containerEl,
            'Settings_Todo_Display_Importance_LowName',
            'Settings_Todo_Display_Importance_LowDescription',
            this.settings.displayOptions_TaskImportance_Low,
            async (value) => {
                this.settings.displayOptions_TaskImportance_Low = value;
            },
        );

        // Task Status Indicator - Not Started
        this.addTextSetting(
            containerEl,
            'Settings_Todo_Display_Status_NotStartedName',
            'Settings_Todo_Display_Status_NotStartedDescription',
            this.settings.displayOptions_TaskStatus_NotStarted,
            async (value) => {
                this.settings.displayOptions_TaskStatus_NotStarted = value;
            },
        );
        this.addTextSetting(
            containerEl,
            'Settings_Todo_Display_Status_InProgressName',
            'Settings_Todo_Display_Status_InProgressDescription',
            this.settings.displayOptions_TaskStatus_InProgress,
            async (value) => {
                this.settings.displayOptions_TaskStatus_InProgress = value;
            },
        );
        this.addTextSetting(
            containerEl,
            'Settings_Todo_Display_Status_CompletedName',
            'Settings_Todo_Display_Status_CompletedDescription',
            this.settings.displayOptions_TaskStatus_Completed,
            async (value) => {
                this.settings.displayOptions_TaskStatus_Completed = value;
            },
        );

        if (this.app.plugins.enabledPlugins.has('periodic-notes')) {
            containerEl.createEl('h2', { text: t('Settings_JournalFormatting') });
            new Setting(containerEl).setName(t('Settings_JournalFormatting_PeriodicNotes')).addToggle((toggle) =>
                toggle.setValue(this.settings.diary.stayWithPN).onChange(async (value) => {
                    if (value) {
                        const periodicNotesSettings = this.app.plugins.plugins['periodic-notes'] as PeriodicNotes;
                        if (periodicNotesSettings) {
                            const { format, folder } = periodicNotesSettings.settings.daily;
                            this.settings.diary = {
                                format,
                                folder,
                                stayWithPN: true,
                            };
                            await this.plugin.saveSettings();
                            this.display();
                        } else {
                            this.userNotice.showMessage('Periodic Notes 中未设置');
                            this.display();
                        }
                    } else {
                        this.settings.diary.stayWithPN = false;
                        await this.plugin.saveSettings();
                        this.display();
                    }
                }),
            );

            const dateFormat = new Setting(containerEl)
                .setName(t('Settings_JournalFormatting_DateFormat'))
                .setDesc(
                    `${t('Settings_JournalFormatting_DateFormatDescription')}  ${
                        this.settings.diary.format ? globalThis.moment().format(this.settings.diary.format) : ''
                    }`,
                )
                .addText((text) =>
                    text.setValue(this.settings.diary.format).onChange(async (value) => {
                        this.settings.diary.format = value;
                        dateFormat.setDesc(
                            `${t('Settings_JournalFormatting_DateFormatDescription')}  ${
                                this.settings.diary.format ? globalThis.moment().format(this.settings.diary.format) : ''
                            }`,
                        );
                        await this.plugin.saveSettings();
                    }),
                )
                .setDisabled(this.settings.diary.stayWithPN);

            new Setting(containerEl)
                .setName(t('Settings_JournalFormatting_Folder'))
                .setDesc(t('Settings_JournalFormatting_FolderDescription'))
                .addText((text) =>
                    text.setValue(this.settings.diary.folder).onChange(async (value) => {
                        this.settings.diary.format = value;
                        await this.plugin.saveSettings();
                    }),
                )
                .setDisabled(this.settings.diary.stayWithPN);
        }
        // Authentication Overrides
        containerEl.createEl('h2', {
            text: t('Settings_Authentication_Heading'),
        });
        containerEl.createEl('p', {
            text: t('Settings_Authentication_Heading_Description'),
        });

        this.addTextSetting(
            containerEl,
            'Settings_Authentication_ClientId',
            'Settings_Authentication_ClientId_Description',
            this.settings.microsoft_AuthenticationClientId,
            async (value) => {
                this.settings.microsoft_AuthenticationClientId = value;
            },
        );

        this.addTextSetting(
            containerEl,
            'Settings_Authentication_ClientId',
            'Settings_Authentication_ClientId_Description',
            this.settings.microsoft_AuthenticationClientId,
            async (value) => {
                this.settings.microsoft_AuthenticationClientId = value;
            },
        );

        // Application Redirect Settings
        containerEl.createEl('h2', {
            text: t('Settings_NativeApp_Heading'),
        });
        containerEl.createEl('p', {
            text: t('Settings_NativeApp_Heading_Description'),
        });

        this.addTextSetting(
            containerEl,
            'Settings_NativeApp_RedirectUriBase',
            'Settings_NativeApp_RedirectUriBase_Description',
            this.settings.microsoftToDoApplication_RedirectUriBase,
            async (value) => {
                this.settings.microsoftToDoApplication_RedirectUriBase = value;
            },
        );
    }

    async hide() {
        const listName = this.settings.todoListSync.listName;

        if (this.settings.todoListSync.listId !== undefined || !listName) {
            if (!listName) {
                this.userNotice.showMessage('General_NoListNameSet');
            }
        } else {
            let listId = await this.plugin.todoApi.getListIdByName(listName);
            const createdTaskList = await this.plugin.todoApi.createTaskList(listName);
            listId ||= createdTaskList?.id;

            if (listId) {
                this.settings.todoListSync = {
                    listName,
                    listId,
                };
                this.userNotice.showMessage('General_ListNameSet');
                await this.plugin.saveSettings();
            } else {
                this.userNotice.showMessage('General_FailedToCreateList');
            }
        }
    }
}
