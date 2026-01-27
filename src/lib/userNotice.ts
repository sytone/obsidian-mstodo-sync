import { Notice } from 'obsidian';
import { t } from './lang.js';
import { logging } from './logging.js';

/**
 * Represents an interface for displaying user-facing messages.
 *
 * @public
 */
export interface IUserNotice {
    showMessage(message: string, timeout?: number): void;
}

/**
 * A class providing functionality to show messages to the user.
 *
 * @public
 *
 * @remarks
 * This class uses Obsidian's Notice component to display a message with a default or user-defined timeout.
 */
export class UserNotice implements IUserNotice {
    /**
     * Internal default duration (in milliseconds) after which the message disappears if no timeout is specified.
     *
     * @private
     */
    private defaultTimeout: number = 5000;

    private readonly logger = logging.getLogger('mstodo-sync.UserNotice');

    /**
     * Displays a message to the user in the Obsidian interface for a specified duration.
     *
     * @param message - The message to be displayed.
     * @param timeout - An optional timeout (in milliseconds) for how long the message will be visible.
     */
    showMessage(message: string, timeout: number = this.defaultTimeout): void {
        const userMessage = new Notice(t(message), timeout);
        userMessage.setMessage(message);
        this.logger.debug(message);
    }
}
