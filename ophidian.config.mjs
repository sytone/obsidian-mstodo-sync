import Builder from '@ophidian/build';

new Builder('src/main.ts') // <-- the path of your main module
    .withInstall() // Optional: publish to OBSIDIAN_TEST_VAULT on build
    .build();
