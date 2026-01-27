# Changelog

All notable changes to this project will be documented in this file.

## 2026-01-27

### Security
- Updated `esbuild` and other dependencies to fix security vulnerabilities reported by `pnpm audit`.

### Added
- **List Selector Modal**: Choose which To Do list to insert tasks from (or "All Lists")
- **Auto-Sync on Plugin Start**: Syncs 5 seconds after Obsidian loads
- **Auto-Sync on File Save**: Syncs when saving files containing tracked tasks (`^MSTD...`)
- **Block ID Generation**: "Insert summary" now generates tracking IDs for two-way sync
- **Subtasks Support**: "Sync Task with details (Pull)" now fetches checklistItems via `$expand`
- **Auto-Sync**: Background synchronization with configurable interval (default 30 mins).
- **Context Menu**: "Sync Task to specific list..." option to choose a target list when pushing.
- **Debug Logging**: Added a setting to enable/disable verbose debug logging to reduce console spam.

### Fixed
- **New Task Cache**: Newly created tasks (Right Click -> Send to To Do) are now immediately cached, ensuring subsequent updates (like completion) sync correctly.
- **Cache File Location**: Moved `Microsoft_cache.json` and `mstd-tasks-delta.json` to the plugin directory to ensure clean uninstallation.
- **400 Bad Request Error**: Fixed `webUrl` validation for linkedResources (local IPs excluded)
- **linkedResources Creation**: Now created separately after task creation when blockLink is available
- **List Selector**: Fixed modal not returning selection (onClose override issue)
- **Block ID Generation**: Fixed async cacheTaskId not being awaited
- **Sync Reliability**: Added retry mechanism for failed delta syncs (falls back to full sync).
- **Completion State**: Implemented hash-based sync to prevent tasks from being incorrectly marked undone when files are touched but not modified.

### Changed
- Modernized GitHub Actions workflows (Node 20, pnpm 9, updated actions)
- Updated README with correct Entra App setup instructions
- Removed redundant workflow files
- **Task Formatting**: Default replacement format is now cleaner (`- [ ] Task Name`) without list name or creation date. (Includes migration for existing settings).
- **Logging**: Reduced default log verbosity (moved many logs from info to debug).



## [1.0.1] - Previous

### Added
- Initial sync functionality
- Push/Pull tasks to Microsoft To Do
- Delta sync support
- Settings UI

## [1.0.0] - Initial Release

### Added
- Basic Microsoft To Do integration
- Task creation and synchronization
- Obsidian block ID linking
