// require('react-native-reanimated').setUpTests();

// FIX:     ReferenceError: self is not defined
globalThis.self = globalThis.self || globalThis;

class LocalStorageMock {
    constructor() {
        this.store = {};
    }

    clear() {
        this.store = {};
    }

    getItem(key) {
        return this.store[key] || null;
    }

    setItem(key, value) {
        this.store[key] = String(value);
    }

    removeItem(key) {
        delete this.store[key];
    }
}

globalThis.localStorage = new LocalStorageMock();

// eslint-disable-next-line @typescript-eslint/no-require-imports, no-undef
globalThis.moment = require('moment');
