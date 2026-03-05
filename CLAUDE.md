# VANTAGE Plugins Development Guide

This repository contains plugins for VANTAGE: Milestone. Plugins extend the application with custom functionality.

## Quick Reference

- **Main app repo:** `C:\Users\steve.amalfitano\source\repos\PrinceCorwin\VANTAGE`
- **Plugins repo:** `C:\Users\Steve.Amalfitano\source\repos\PrinceCorwin\VANTAGE-Plugins`
- **Feed URL:** `https://raw.githubusercontent.com/PrinceCorwin/VANTAGE-Plugins/main/plugins-index.json`
- **Local install path:** `%LocalAppData%\VANTAGE\Plugins\<plugin-id>\<version>\`

## Plugin Architecture

Each plugin must:
1. Reference `VANTAGE.exe` to access the plugin interfaces
2. Implement `IVantagePlugin` interface
3. Include a `plugin.json` manifest

### IVantagePlugin Interface

```csharp
public interface IVantagePlugin
{
    string Id { get; }
    string Name { get; }
    void Initialize(IPluginHost host);
    void Shutdown();
}
```

### IPluginHost (What the app provides)

```csharp
public interface IPluginHost
{
    void AddToolsMenuItem(string header, Action onClick, bool addSeparatorBefore = false);
    Window MainWindow { get; }
    string CurrentUsername { get; }
    void ShowInfo(string message, string title = "Information");
    void ShowError(string message, string title = "Error");
    bool ShowConfirmation(string message, string title = "Confirm");
    void LogInfo(string message, string source);
    void LogError(Exception ex, string source);
}
```

## Creating a New Plugin

### Step 1: Create Plugin Folder Structure

```
src/<plugin-id>/
├── <PluginName>.csproj
├── <PluginName>.cs
└── plugin.json
```

### Step 2: Create the .csproj File

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net8.0-windows</TargetFramework>
    <UseWPF>true</UseWPF>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <AssemblyName>PluginName</AssemblyName>
  </PropertyGroup>

  <ItemGroup>
    <!-- Reference the main VANTAGE app for plugin interfaces -->
    <Reference Include="VANTAGE">
      <HintPath>..\..\..\..\VANTAGE\bin\Debug\net8.0-windows\VANTAGE.exe</HintPath>
      <Private>false</Private>
    </Reference>
  </ItemGroup>
</Project>
```

### Step 3: Implement IVantagePlugin

```csharp
using System;
using System.Windows;
using VANTAGE.Services.Plugins;

namespace PluginNamespace
{
    public class PluginClassName : IVantagePlugin
    {
        private IPluginHost? _host;

        public string Id => "plugin-id";
        public string Name => "Plugin Display Name";

        public void Initialize(IPluginHost host)
        {
            _host = host;

            // Add a menu item under Tools
            host.AddToolsMenuItem("Menu Item Text", OnMenuClick, addSeparatorBefore: true);
        }

        public void Shutdown()
        {
            // Cleanup if needed
        }

        private void OnMenuClick()
        {
            // Plugin action here
            _host?.ShowInfo("Hello from plugin!", Name);
        }
    }
}
```

### Step 4: Create plugin.json Manifest

```json
{
  "id": "plugin-id",
  "name": "Plugin Display Name",
  "version": "1.0.0",
  "pluginType": "action",
  "project": "",
  "description": "What this plugin does",
  "assemblyFile": "PluginName.dll",
  "entryType": "PluginNamespace.PluginClassName",
  "minAppVersion": "1.0.0",
  "maxAppVersion": ""
}
```

**Field reference:**
- `id` - Unique plugin identifier (lowercase, hyphens OK)
- `name` - Display name shown in Plugin Manager
- `version` - Semantic version (X.Y.Z)
- `pluginType` - "action" (has UI) or "extension" (passive)
- `project` - Project scope (empty = global)
- `description` - Brief description
- `assemblyFile` - DLL filename
- `entryType` - Full type name implementing IVantagePlugin
- `minAppVersion` - Minimum VANTAGE version required
- `maxAppVersion` - Maximum VANTAGE version (empty = no limit)

## Building a Plugin

```bash
cd src/<plugin-id>
dotnet build -c Release
```

Output will be in `bin/Release/net8.0-windows/`

## Packaging a Plugin for Release

### Step 1: Build the Plugin

```bash
cd src/<plugin-id>
dotnet build -c Release
```

### Step 2: Create the ZIP Package

The ZIP must contain:
- `plugin.json` (manifest)
- `<PluginName>.dll` (the plugin assembly)
- Any additional dependencies (not from VANTAGE or .NET runtime)

```bash
# From the plugin's bin/Release/net8.0-windows directory
# Create ZIP with just the required files
```

Name the ZIP: `<plugin-id>.<version>.zip` (e.g., `ptp-updater.1.0.0.zip`)

### Step 3: Create GitHub Release

1. Go to https://github.com/PrinceCorwin/VANTAGE-Plugins/releases
2. Click "Draft a new release"
3. Tag: `<plugin-id>-v<version>` (e.g., `ptp-updater-v1.0.0`)
4. Title: `<Plugin Name> v<version>`
5. Upload the ZIP file
6. Publish the release

### Step 4: Update plugins-index.json

Add or update the plugin entry:

```json
{
  "id": "plugin-id",
  "name": "Plugin Display Name",
  "version": "1.0.0",
  "pluginType": "action",
  "project": "",
  "description": "What this plugin does",
  "packageUrl": "https://github.com/PrinceCorwin/VANTAGE-Plugins/releases/download/<tag>/<plugin-id>.<version>.zip",
  "sha256": ""
}
```

### Step 5: Commit and Push

```bash
git add plugins-index.json
git commit -m "Release <plugin-name> v<version>"
git push
```

## Testing Plugins Locally

### Manual Install (for development)

1. Build the plugin
2. Copy the output to: `%LocalAppData%\VANTAGE\Plugins\<plugin-id>\<version>\`
3. Ensure `plugin.json` is in that folder
4. Restart VANTAGE

### Via Plugin Manager (for release testing)

1. Update `plugins-index.json` with the new version
2. Push to GitHub
3. Wait for GitHub raw content cache (~5 minutes)
4. Open VANTAGE > Plugin Manager > Available tab
5. Install the plugin

## Updating a Plugin

1. Increment version in:
   - `plugin.json`
   - `plugins-index.json`
2. Build and create new ZIP
3. Create new GitHub Release with new tag
4. Update `packageUrl` in `plugins-index.json`
5. Commit and push

VANTAGE auto-updates plugins at startup, so users will get the new version automatically.

## Conventions

- Plugin IDs: lowercase, hyphens (e.g., `ptp-updater`)
- Version tags: `<plugin-id>-v<version>` (e.g., `ptp-updater-v1.0.0`)
- ZIP names: `<plugin-id>.<version>.zip`
- One plugin per folder in `src/`

## Accessing VANTAGE Services

Plugins can access app functionality through the `IPluginHost`. If you need additional capabilities not exposed by IPluginHost, the host interface can be extended in the main VANTAGE repo.

Common patterns:
- Show dialogs: Use `host.MainWindow` as owner
- File picker: Use standard WPF `OpenFileDialog` with `host.MainWindow` as owner
- Logging: Use `host.LogInfo()` and `host.LogError()`
- User feedback: Use `host.ShowInfo()`, `host.ShowError()`, `host.ShowConfirmation()`
