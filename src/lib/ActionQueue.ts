export class ActionQueue {
    private queue: (() => Promise<void>)[] = [];
    private delay: number;
    private isProcessing: boolean = false;

    constructor(delay: number) {
        this.delay = delay;
    }

    public addAction(action: () => Promise<void>): void {
        this.queue.push(action);
        if (!this.isProcessing) {
            this.processQueue();
        }
    }

    private async processQueue(): Promise<void> {
        this.isProcessing = true;
        while (this.queue.length > 0) {
            const action = this.queue.shift();
            if (action) {
                await action();
                await this.sleep(this.delay);
            }
        }
        this.isProcessing = false;
    }

    private sleep(ms: number): Promise<void> {
        return new Promise((resolve) => setTimeout(resolve, ms));
    }
}
