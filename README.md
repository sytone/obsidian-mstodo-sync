# Microsoft To Do Sync for Obsidian

Sync your tasks between Obsidian and Microsoft To Do with two-way synchronization.

> [!IMPORTANT]
> **AI Assistance Disclaimer**
>
> Parts of this project were created with the assistance of AI tools. All generated code has been tested and verified by the author. This software is provided "as is", without any warranty or guarantee of correctness.

## Features

- **Push Tasks** - Create tasks in Microsoft To Do from your Obsidian notes
- **Pull Tasks** - Import tasks from Microsoft To Do into Obsidian
- **Two-Way Sync** - Automatic synchronization keeps both sides in sync
- **List Selection** - Choose which Microsoft To Do list to sync with
- **Subtasks Support** - Sync checklist items (subtasks) with "Sync with details"
- **Block ID Tracking** - Tasks are linked via block IDs (`^MSTD...`) for reliable sync
- **Auto-Sync** - Syncs automatically on plugin start and when saving files with tracked tasks

## Installation

### Via BRAT (Recommended for Beta Testing)

1. Install [BRAT](https://github.com/TfTHacker/obsidian42-brat) from Community Plugins
2. Open Command Palette → **BRAT: Add a beta plugin for testing**
3. Enter: `https://github.com/N0C1/obsidian-mstodo-sync/`
4. Enable the plugin in Settings → Community Plugins

### Manual Installation

1. Download `main.js`, `manifest.json`, and `styles.css` from the [latest release](https://github.com/N0C1/obsidian-mstodo-sync/releases)
2. Create folder: `YourVault/.obsidian/plugins/obsidian-mstodo-sync/`
3. Copy the downloaded files into this folder
4. Enable the plugin in Settings → Community Plugins

## Setup

### Using the Default App (Quick Start)

The plugin comes with a pre-configured Microsoft Entra App by the original author [Lumos](https://github.com/LumosLovegood/obsidian-mstodo-sync). Just:

1. Enable the plugin
2. Run any command (e.g., "Insert summary from Microsoft To Do")
3. A device code will be copied to your clipboard
4. Follow the authentication link and enter the code
5. Done! You're connected.

### Using Your Own Microsoft Entra App (Recommended)

For better security and control, create your own app:

#### Step 1: Create App Registration

1. Go to [Microsoft Entra App Registrations](https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Click **Register new application**
3. Configure:
   - **Name**: `Obsidian Microsoft To Do Sync` (or any name)
   - **Supported account types**: "Accounts in any organizational directory and personal Microsoft accounts"
   - **Redirect URI**: Leave empty for now
4. Click **Register**

#### Step 2: Enable Public Client Flows

1. In your new app, go to **Manage** → **Authentication**
2. Scroll down to **Advanced settings**
3. Set **Allow public client flows** to **Yes**
4. Click **Save**

#### Step 3: Add Platform (Optional)

1. Still in **Authentication**, click **Add a platform**
2. Select **Mobile and desktop applications**
3. Check: `https://login.microsoftonline.com/common/oauth2/nativeclient`
4. Click **Configure**

#### Step 4: Add API Permissions

1. Go to **Manage** → **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph** → **Delegated permissions**
4. Add these permissions:
   - `Tasks.ReadWrite`
   - `User.Read`
5. Click **Add permissions**

#### Step 5: Get Your Application ID

1. Go to **Overview**
2. Copy the **Application (client) ID**

#### Step 6: Configure the Plugin

1. In Obsidian, go to Settings → **Microsoft To Do Sync**
2. Ignore the first authentication try. This is the default Entra App.
3. Scroll to **Authentication**
4. Paste your **Client ID**
5. Restart Obsidian or reload the plugin

## Usage

### Commands (Ctrl+P)

| Command | Description |
|---------|-------------|
| **Push to Microsoft To Do** | Push current task to MS To Do (no replacement) |
| **Push to Microsoft To Do and replace** | Push and update the line with block ID |
| **Insert summary from Microsoft To Do** | Opens list selector, inserts tasks with tracking IDs |
| **Open To Do** | Opens the task in Microsoft To Do app/web |

### Context Menu (Right-click)

| Option | Description |
|--------|-------------|
| **Sync Task (Push)** | Push task to MS To Do |
| **Sync Task (Pull)** | Pull latest from MS To Do |
| **Sync Task to specific list...** | Pushes task to a user selected To Do list |
| **Sync Task with details (Push)** | Push with body and subtasks |
| **Sync Task with details (Pull)** | Pull with subtasks (checklistItems) |
| **Open To Do** | Open in MS To Do |

### Task Format

Tasks are tracked using Obsidian block IDs:

```markdown
- [ ] My task title ^MSTDa8de00001
  - [ ] Subtask 1
  - [ ] Subtask 2
  Notes about the task...
```

The `^MSTDa8de00001` links this task to the corresponding Microsoft To Do task.

## Auto-Sync Behavior

The plugin automatically syncs in these situations:

1. **On Plugin Start** - Syncs 5 seconds after Obsidian loads
2. **On File Save** - When you (auto)save a file containing tracked tasks (`^MSTD...`)
   - Uses 3-second debounce to avoid excessive syncing
   - 10-second cooldown per file to prevent loops

## Settings

| Setting | Description |
|---------|-------------|
| **Default List Name** | The MS To Do list to sync with (e.g., "Obsidian") |
| **Date Format** | Format for dates (default: `YYYY-MM-DD`) |
| **Priority Indicators** | Emojis for high/normal/low priority |
| **Status Indicators** | Characters for task status (`x`, `/`, ` `) |
| **Client ID** | Your own Entra App Client ID (optional) |

## Troubleshooting

### "The property 'webUrl' has an invalid value"
This error occurred with local IP addresses in the Redirect URI. Fixed in latest version - `webUrl` is now optional.

### Tasks not syncing
- Ensure tasks have a block ID (`^MSTD...`)
- Check if the list name in settings matches your MS To Do list
- Try "Reset Task Cache" from the Hacking menu (if enabled)

### Authentication issues
- Revoke permissions at [Microsoft Account](https://account.live.com/consent/Manage)
- Re-authenticate with the plugin

## Development

```bash
# Install dependencies
pnpm install

# Development build with watch
pnpm run dev

# Production build
pnpm run build

# Lint
pnpm run lint
pnpm run lint:fix
```

## License

MIT

## Credits

Originally created by [Lumos](https://github.com/LumosLovegood/obsidian-mstodo-sync).
Improved originally by [Sytone](https://github.com/sytone).
