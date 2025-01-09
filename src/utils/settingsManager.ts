import { type IMsTodoSyncSettings } from 'src/gui/msTodoSyncSettingTab';
import type MsTodoSync from 'src/main';

interface ISettingsManager {
    settings: IMsTodoSyncSettings;
    vaultName: string;
    saveSettings(): void;
}

class SettingsManager implements ISettingsManager {
    constructor(private readonly plugin: MsTodoSync) {}

    public get settings() {
        return this.plugin.settings;
    }

    public get vaultName() {
        return this.plugin.app.vault.getName();
    }

    public getTaskIdFromBlockId(blockId: string): string {
        return this.findKeyCaseInsensitive(this.plugin.settings.taskIdLookup, blockId);
    }

    async saveSettings(): Promise<void> {
        // Implementation to save settings
        await this.plugin.saveData(this.plugin.settings);
    }

    private findKeyCaseInsensitive(obj: Record<string, any>, key: string): any {
        const lowerCaseKey = key.toLowerCase();
        const foundKey = Object.keys(obj).find((k) => k.toLowerCase() === lowerCaseKey);
        return foundKey ? obj[foundKey] : undefined;
    }
}

export { type ISettingsManager, SettingsManager };
