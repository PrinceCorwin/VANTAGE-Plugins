# VANTAGE: Milestone Plugins

This repository contains plugins for the VANTAGE: Milestone construction progress tracking application.

## For Users

Plugins are managed through the **Plugin Manager** in VANTAGE:
1. Click the `...` (settings) button in the top-right toolbar
2. Select **Plugin Manager...**
3. Browse the **Available** tab to see plugins you can install
4. Select a plugin and click **Install Selected**

Plugins are automatically updated when VANTAGE starts.

## For Developers

See [CLAUDE.md](CLAUDE.md) for plugin development instructions.

## Repository Structure

```
VANTAGE-Plugins/
├── plugins-index.json    # Feed file listing all available plugins
├── src/                  # Plugin source code (one folder per plugin)
│   └── <plugin-id>/
│       ├── <PluginName>.csproj
│       ├── <PluginName>.cs
│       └── plugin.json
└── README.md
```

## Feed Format

The `plugins-index.json` file lists all available plugins:

```json
{
  "plugins": [
    {
      "id": "plugin-id",
      "name": "Plugin Display Name",
      "version": "1.0.0",
      "pluginType": "action",
      "project": "",
      "description": "What this plugin does",
      "packageUrl": "https://github.com/PrinceCorwin/VANTAGE-Plugins/releases/download/v1.0.0/plugin-id.1.0.0.zip",
      "sha256": ""
    }
  ]
}
```

## Plugin Types

- **action** - Creates UI elements (menu items, buttons) that users interact with
- **extension** - Passive features (themes, background services)
