// jest.config.ts
import { pathsToModuleNameMapper } from 'ts-jest';
// In the following statement, replace `./tsconfig` with the path to your `tsconfig` file
// which contains the path mapping (ie the `compilerOptions.paths` option):
import { compilerOptions } from './tsconfig.json';
import type { JestConfigWithTsJest } from 'ts-jest';

const jestConfig: JestConfigWithTsJest = {
    preset: 'ts-jest',
    testEnvironment: 'node',
    testMatch: ['**/src/**/*.test.js', '**/src/**/*.test.ts'],
    transform: {
        '^.+\\.ts$': 'ts-jest',
    },
    setupFilesAfterEnv: ['./jest-setup.js'],
    moduleFileExtensions: ['ts', 'js', 'json', 'node'],
    roots: ['<rootDir>'],
    modulePaths: [compilerOptions.baseUrl, '/node_modules'], // <-- This will be set to 'baseUrl' value
    moduleNameMapper: pathsToModuleNameMapper(compilerOptions.paths /*, { prefix: '<rootDir>/' } */),
};

export default jestConfig;
