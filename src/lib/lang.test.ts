/* eslint-disable no-undef */
import { t } from './lang';

const mockedLangMap =
    '{"en": { "Notice_DeviceCodeCopiedToClipboard": "The device code has been copied to the clipboard" }, "zh": { "Notice_DeviceCodeCopiedToClipboard": "设备代码已复制到剪贴板" }}';

describe('t function', () => {
    it('mock should work', () => {
        globalThis.localStorage.setItem('language', 'nonexistent_locale');
        const result = globalThis.localStorage.getItem('language');
        expect(result).toBe('nonexistent_locale');
    });

    it('should return the translated string if it exists in the locale', () => {
        globalThis.localStorage.setItem('language', 'en');
        globalThis.localStorage.setItem('mstd_mock_localeMap', mockedLangMap);
        const result = t('Notice_DeviceCodeCopiedToClipboard');
        expect(result).toBe('The device code has been copied to the clipboard');
    });

    it('should return the input string if it does not exist in the locale', () => {
        const result = t('nonexistent_key');
        expect(result).toBe('nonexistent_key');
    });

    it('should log an error if the locale is not found', () => {
        globalThis.localStorage.setItem('language', 'nonexistent_locale');
        const consoleErrorSpy = jest.spyOn(console, 'error').mockImplementation(() => {});
        t('Notice_DeviceCodeCopiedToClipboard');
        expect(consoleErrorSpy).toHaveBeenCalledWith('Error: locale not found', 'nonexistent_locale');
        consoleErrorSpy.mockRestore();
    });

    it('should default to English if no language is set in localStorage', () => {
        globalThis.localStorage.removeItem('language');
        const result = t('Notice_DeviceCodeCopiedToClipboard');
        expect(result).toBe('The device code has been copied to the clipboard');
    });

    it('should use the correct locale based on the language set in localStorage', () => {
        globalThis.localStorage.setItem('language', 'zh');
        const result = t('Notice_DeviceCodeCopiedToClipboard');
        expect(result).toBe('设备代码已复制到剪贴板');
    });
});
