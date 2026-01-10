# Shared Ribbon Architecture

## Overview
This workspace uses a **Shared Ribbon Host** pattern where multiple independent add-in solutions can contribute commands to a single ribbon panel in Revit.

## Architecture

### Components

1. **Quoc_MEP.RibbonHost** (Shared Infrastructure)
   - Creates the main ribbon tab and panel
   - Dynamically discovers and loads commands from DLLs
   - Single `.addin` manifest that Revit loads
   - Location: `Quoc_MEP.RibbonHost/`

2. **ScheduleManager** (Independent Solution)
   - Separate solution for Schedule Manager features
   - Builds independently with SDK-style project
   - Commands automatically appear in shared ribbon
   - Location: `ScheduleManager/` + `ScheduleManager.sln`

3. **Quoc_MEP** (Main Solution)
   - Original add-in with export, connect, and other tools
   - Can remain as old-style project
   - Future: Add CommandInfo classes for ribbon discovery

## How It Works

### 1. Build Process
```powershell
# Build RibbonHost for Revit 2023
.\build-ribbonhost.ps1 -Versions @("2023")

# Build ScheduleManager for Revit 2023
.\build-schedule-manager.ps1 -Versions @("2023")
```

Both projects output to the same folder:
```
bin\Release\Revit2023\
├── Quoc_MEP.RibbonHost.dll    # Loaded by Revit
├── ScheduleManager.dll         # Discovered by RibbonHost
└── (other dependencies)
```

### 2. Deployment
Copy the `.addin` file to Revit addins folder:
```powershell
# For Revit 2023
Copy-Item "Quoc_MEP_SharedRibbon.addin" `
    -Destination "$env:APPDATA\Autodesk\Revit\Addins\2023\"
```

### 3. Command Discovery
RibbonHost automatically finds commands by looking for `*CommandInfo` classes:

```csharp
// ScheduleManagerCommandInfo.cs
public static class ScheduleManagerCommandInfo
{
    public static string Name => "ScheduleManager.Open";
    public static string Text => "Schedule\nManager";
    public static string Tooltip => "Open Schedule Manager";
    public static string CommandClass => "ScheduleManager.ScheduleManagerCommand";
}
```

### 4. Runtime Loading
1. Revit loads `Quoc_MEP.RibbonHost.dll` (from `.addin` manifest)
2. RibbonHost creates tab "Quoc MEP" and panel "MEP Tools"
3. RibbonHost scans all DLLs in same folder
4. For each DLL, finds `*CommandInfo` classes
5. Creates buttons using metadata from CommandInfo
6. Buttons execute commands from their respective DLLs

## Adding New Commands

### Option 1: Add to Existing Solution
1. Create your command class implementing `IExternalCommand`
2. Create a `CommandInfo` class with metadata
3. Build and deploy to shared bin folder

### Option 2: Create New Solution
1. Create new solution (e.g., `MyNewFeature.sln`)
2. Reference Nice3point packages
3. Create command + CommandInfo classes
4. Configure output to shared bin folder
5. Build script will automatically include it

## File Structure
```
RevitAPIMEP/
├── Quoc_MEP.sln                        # Original solution
├── ScheduleManager.sln                 # Schedule Manager solution
│
├── Quoc_MEP.RibbonHost/                # Shared ribbon infrastructure
│   ├── Application.cs                  # IExternalApplication
│   └── Quoc_MEP.RibbonHost.csproj
│
├── ScheduleManager/                    # Schedule Manager project
│   ├── ScheduleManagerCommand.cs       # IExternalCommand
│   ├── ScheduleManagerCommandInfo.cs   # Metadata for discovery
│   ├── ScheduleManagerViewModel.cs
│   └── ScheduleManager.csproj
│
├── bin/Release/Revit2023/              # Shared output folder
│   ├── Quoc_MEP.RibbonHost.dll
│   ├── ScheduleManager.dll
│   └── ...
│
├── build-ribbonhost.ps1                # Build RibbonHost
├── build-schedule-manager.ps1          # Build ScheduleManager
└── Quoc_MEP_SharedRibbon.addin         # Manifest file
```

## Benefits

### Independent Development
- Each solution builds separately
- No cross-solution dependencies
- Test and debug independently

### Shared UI
- Single ribbon tab and panel
- Consistent user experience
- No duplicate UI elements

### Flexible Deployment
- Add/remove features without rebuilding everything
- Different teams can work on different solutions
- Easy to enable/disable features (just remove DLL)

### Package Management
- SDK-style projects support Nice3point packages properly
- Each solution manages its own dependencies
- No package conflicts between solutions

## Testing

### Test RibbonHost
```powershell
.\build-ribbonhost.ps1 -Versions @("2023") -Configuration Debug
```

### Test ScheduleManager
```powershell
.\build-schedule-manager.ps1 -Versions @("2023") -Configuration Debug
```

### Deploy and Test in Revit
```powershell
# Copy manifest
Copy-Item "Quoc_MEP_SharedRibbon.addin" `
    -Destination "$env:APPDATA\Autodesk\Revit\Addins\2023\"

# Launch Revit 2023 and check ribbon
```

## Troubleshooting

### Commands not appearing in ribbon
1. Check DLL is in same folder as RibbonHost.dll
2. Verify CommandInfo class has correct static properties
3. Check RibbonHost debug output (use DebugView)

### Build errors
1. Ensure Nice3point packages are restoring correctly
2. Check RevitVersion parameter is set
3. Verify output paths match in all projects

### Multiple versions
Each Revit version needs its own build:
```powershell
# Build for multiple versions
.\build-ribbonhost.ps1 -Versions @("2020", "2021", "2022", "2023", "2024")
.\build-schedule-manager.ps1 -Versions @("2020", "2021", "2022", "2023", "2024")
```

## Future Enhancements
- Add icons to CommandInfo pattern
- Support for pulldown buttons (grouped commands)
- Dynamic command loading without restart
- Configuration file for command organization
