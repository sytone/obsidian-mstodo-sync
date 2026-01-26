import en from './locale/en.json';
import zhCN from './locale/zh-cn.json';

export type Translations = Record<string, string>;

const localeMap: Record<string, Translations> = {
    en,
    'en-GB': en, // Yes I know it's not the same. But it's close enough.
    zh: zhCN,
};

export function getLocaleMap(): Record<string, Translations> {
    if (globalThis.localStorage.getItem('mstd_mock_localeMap')) {
        const mockLocaleMap = globalThis.localStorage.getItem('mstd_mock_localeMap');
        return mockLocaleMap ? JSON.parse(mockLocaleMap) : localeMap;
    }
    return localeMap;
}

function getLanguage(): string | null {
    return globalThis.localStorage.getItem('language');
}

function getLocale(language: string): Translations {
    const localeMap = getLocaleMap();
    return localeMap[language];
}

export function t(string_: string): string {
    const language = getLanguage();
    const locale = getLocale(language ?? 'en');
    if (!locale) {
        console.error('Error: locale not found', language);
    }

    return locale?.[string_] || string_;
}
