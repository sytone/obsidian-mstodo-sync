import { Platform, type Plugin } from 'obsidian';
/*
 * EventEmitter2 is an implementation of the EventEmitter module found in Node.js.
 * In addition to having a better benchmark performance than EventEmitter and being
 * browser-compatible, it also extends the interface of EventEmitter with many
 * additional non-breaking features.
 *
 * This has been added as EventEmitter in Node.JS is not available in the browser.
 * https://www.npmjs.com/package/eventemitter2
 */
import { EventEmitter2 } from 'eventemitter2';

/**
 * All possible log levels
 * @public
 */
export interface ILogLevel {
    1: 'trace';
    2: 'debug';
    3: 'info';
    4: 'warn';
    5: 'error';
}

/**
 * Logger class to handle consistency of logs across the plugin.
 *
 * @export
 * @interface ILogEntry
 */
export interface ILogEntry {
    traceId?: string;
    level: string;
    module: string;
    location?: string;
    message: string;
    objects: unknown;
}

/**
 * Logging options structure.
 *
 * @export
 * @interface ILogOptions
 */
export interface ILogOptions {
    minLevels: Record<string, string>;
}

/**
 * Log level IDs (1 - 5)
 * @public
 */
export type TLogLevelId = keyof ILogLevel;

/**
 * Log level names (trace - error)
 * @public
 */
export type TLogLevelName = ILogLevel[TLogLevelId];

/**
 * Logger class to handle consistency of logs across the plugin.
 *
 * @export
 * @class LogManager
 * @extends {EventEmitter2}
 */
export class LogManager extends EventEmitter2 {
    private options: ILogOptions = {
        minLevels: {
            '': 'debug',
            'mstodo-sync': 'debug',
        },
    };

    // Prevent the console logger from being added twice
    private consoleLoggerRegistered = false;

    /**
     * Set the minimum log levels for the module name or global.
     *
     * @param {ILogOptions} options
     * @return {*}  {LogManager}
     * @memberof LogManager
     */
    public configure(options: ILogOptions): this {
        this.options = Object.assign({}, this.options, options);
        return this;
    }

    /**
     * Returns a logger instance for the given module name.
     *
     * @param {string} module
     * @return {*}  {Logger}
     * @memberof LogManager
     */
    public getLogger(moduleName: string): Logger {
        let currentMinimumLevel = 'none';
        let match = '';

        for (const key in this.options.minLevels) {
            if (moduleName.startsWith(key) && key.length >= match.length) {
                currentMinimumLevel = this.options.minLevels[key];
                match = key;
            }
        }

        return new Logger(this, moduleName, currentMinimumLevel);
    }

    /**
     *
     *
     * @param {(logEntry: ILogEntry) => void} listener
     * @return {*}  {LogManager}
     * @memberof LogManager
     */
    public onLogEntry(listener: (logEntry: ILogEntry) => void): this {
        this.on('log', listener);
        return this;
    }

    // Private period: number = 0;
    arrAvg = (array: number[]) => array.reduce((a, b) => a + b, 0) / array.length;

    /**
     * Registers a logger that write to the console.
     *
     * @return {*}  {LogManager}
     * @memberof LogManager
     */
    public registerConsoleLogger(): this {
        if (this.consoleLoggerRegistered) {
            return this;
        }

        this.onLogEntry((logEntry) => {
            // 2024-12-19T22:53:37.000Z - >'2024-12-19 22:53:37'
            const messageDate = new Date().toISOString().slice(0, 19).replace('T', ' ');

            let message = `[${messageDate}][${logEntry.level}][${logEntry.module}]`;

            if (logEntry.traceId) {
                message += `[${logEntry.traceId}]`;
            }

            message += ` ${logEntry.message}`;
            if (logEntry.objects === undefined) {
                logEntry.objects = '';
            }

            switch (logEntry.level) {
                case 'trace': {
                    console.trace(message, logEntry.objects);
                    break;
                }

                case 'debug': {
                    console.debug(message, logEntry.objects);
                    break;
                }

                case 'info': {
                    console.info(message, logEntry.objects);
                    break;
                }

                case 'warn': {
                    console.warn(message, logEntry.objects);
                    break;
                }

                case 'error': {
                    console.error(message, logEntry.objects);
                    break;
                }

                default: {
                    console.log(`{${logEntry.level}} ${message}`, logEntry.objects);
                }
            }
        });

        this.consoleLoggerRegistered = true;
        return this;
    }
}

export const logging = new LogManager();

/**
 * Main logging library, to view the logs a logger listener must be added. The
 * Console Logger is already implemented for this project.
 *
 * @export
 * @class Logger
 */
export class Logger {
    private readonly minLevel: number;
    private readonly levels: Record<string, number> = {
        trace: 1,
        debug: 2,
        info: 3,
        warn: 4,
        error: 5,
    };

    /**
     * Creates an instance of Logger.
     * @param {EventEmitter2} logManager
     * @param {string} name
     * @param {string} minLevel
     * @memberof Logger
     */
    constructor(
        private readonly logManager: EventEmitter2,
        private readonly name: string,
        minLevel: string,
    ) {
        this.minLevel = this.levelToInt(minLevel);
    }

    /**
     * Central logging method.
     * @param logLevel
     * @param message
     */
    public log(logLevel: string, message: string, objects?: unknown): void {
        const level = this.levelToInt(logLevel);
        if (level < this.minLevel) {
            return;
        }

        const logEntry: ILogEntry = {
            level: logLevel,
            module: this.name,
            message,
            objects,
            traceId: undefined,
        };

        // Obtain the line/file through a thoroughly hacky method
        // This creates a new stack trace and pulls the caller from it.  If the caller
        // if .trace()
        // const error = new Error('');
        // if (error.stack) {
        //     const cla = error.stack.split('\n');
        //     let idx = 1;
        //     while (idx < cla.length && cla[idx].includes('at Logger.Object.')) idx++;
        //     if (idx < cla.length) {
        //         logEntry.location = cla[idx].slice(cla[idx].indexOf('at ') + 3, cla[idx].length);
        //     }
        // }

        this.logManager.emit('log', logEntry);
    }

    public trace(message: string, objects?: unknown): void {
        this.log('trace', message, objects);
    }

    public debug(message: string, objects?: unknown): void {
        this.log('debug', message, objects);
    }

    public info(message: string, objects?: unknown): void {
        this.log('info', message, objects);
    }

    public warn(message: string, objects?: unknown): void {
        this.log('warn', message, objects);
    }

    public error(message: string, objects?: unknown): void {
        this.log('error', message, objects);
    }

    /**
     * Central logging method with a trace ID to track calls between modules/components.
     * @param logLevel
     * @param message
     */
    public logWithId(logLevel: string, traceId: string, message: string, objects?: unknown): void {
        const level = this.levelToInt(logLevel);
        if (level < this.minLevel) {
            return;
        }

        const logEntry: ILogEntry = {
            level: logLevel,
            module: this.name,
            message,
            objects,
            traceId,
        };

        this.logManager.emit('log', logEntry);
    }

    public traceWithId(traceId: string, message: string, objects?: unknown): void {
        this.logWithId('trace', traceId, message, objects);
    }

    public debugWithId(traceId: string, message: string, objects?: unknown): void {
        this.logWithId('debug', traceId, message, objects);
    }

    public infoWithId(traceId: string, message: string, objects?: unknown): void {
        this.logWithId('info', traceId, message, objects);
    }

    public warnWithId(traceId: string, message: string, objects?: unknown): void {
        this.logWithId('warn', traceId, message, objects);
    }

    public errorWithId(traceId: string, message: string, objects?: unknown): void {
        this.logWithId('error', traceId, message, objects);
    }

    /**
     * Converts a string level (trace/debug/info/warn/error) into a number
     *
     * @param minLevel
     */
    private levelToInt(minLevel: string): number {
        if (minLevel.toLowerCase() in this.levels) {
            return this.levels[minLevel.toLowerCase()];
        }

        return 99;
    }
}

export function logCallDetails() {
    return function (target: unknown, propertyKey: string, descriptor: PropertyDescriptor) {
        const originalMethod = descriptor.value;
        const logger = logging.getLogger('mstodo-sync');

        descriptor.value = async function (...arguments_: unknown[]) {
            const startTime = new Date(Date.now());
            const result = await originalMethod.apply(this, arguments_);
            const endTime = new Date(Date.now());
            const elapsed = endTime.getTime() - startTime.getTime();

            logger.debug(
                `${typeof target}:${propertyKey} called with ${
                    arguments_.length
                } arguments. Took: ${elapsed}ms ${JSON.stringify(arguments_)}`,
            );
            return result;
        };

        return descriptor;
    };
}

/**
 * Provides a simple log function that can be used to log messages against default module.
 *
 * @export
 * @param {TLogLevelName} logLevel
 * @param {string} message
 */
export function log(logLevel: TLogLevelName, message: string, objects?: unknown) {
    const logger = logging.getLogger('mstodo-sync');

    switch (logLevel) {
        case 'trace': {
            logger.trace(message, objects);
            break;
        }

        case 'debug': {
            logger.debug(message, objects);
            break;
        }

        case 'info': {
            logger.info(message, objects);
            break;
        }

        case 'warn': {
            logger.warn(message, objects);
            break;
        }

        case 'error': {
            logger.error(message, objects);
            break;
        }
    }
}

/**
 * This allows the plugin to be debugged in a mobile application
 * add it when debugging on a device. Not meant to be used by
 * end users. Add it into main.ts and remove before you commit.
 *
 * @export
 * @param {Plugin} plugin
 * @return {*}
 */
export function monkeyPatchConsole(plugin: Plugin) {
    if (!Platform.isMobile) {
        return;
    }

    const logFile = `${plugin.manifest.dir}/mstodo-sync-logs.txt`;
    const logs: string[] = [];
    const logMessages =
        (prefix: string) =>
        (...messages: unknown[]) => {
            logs.push(`\n[${prefix}]`);
            for (const message of messages) {
                logs.push(String(message));
            }

            plugin.app.vault.adapter.write(logFile, logs.join(' '));
        };

    console.debug = logMessages('debug');
    console.error = logMessages('error');
    console.info = logMessages('info');
    console.log = logMessages('log');
    console.warn = logMessages('warn');
}
