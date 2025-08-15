export class CaseInsensitiveRecord<T> {
    private record: Record<string, T> = {};

    set(key: string, value: T): void {
        this.record[key.toLowerCase()] = value;
    }

    get(key: string): T | undefined {
        return this.record[key.toLowerCase()];
    }

    has(key: string): boolean {
        return key.toLowerCase() in this.record;
    }

    delete(key: string): void {
        delete this.record[key.toLowerCase()];
    }

    keys(): string[] {
        return Object.keys(this.record);
    }

    values(): T[] {
        return Object.values(this.record);
    }
}
