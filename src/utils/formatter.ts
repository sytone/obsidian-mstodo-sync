import { TASK_REGEX } from 'src/constants';
import type MsTodoSync from '../main.js';

export const formatTask = (plugin: MsTodoSync, line: string) => {
    let output: string;
    const format = plugin.settings.displayOptions_ReplacementFormat;
    // eslint-disable-next-line prefer-const
    output = format.replace(TASK_REGEX, line);
    return output;
};
