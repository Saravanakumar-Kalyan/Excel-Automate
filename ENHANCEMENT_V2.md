# Enhancement v2: Interactive Configuration & Runtime Input Summary

## What Was Added

### 1. Enhanced Configuration Menu (Main.ps1)

Added three new configuration management functions:

#### A. Edit-Categories Function
- View all current categories
- Add new categories
- Delete categories by index
- All changes saved immediately to config

#### B. Edit-FileTypes Function
- View all current file types
- Add new file types
- Delete file types by index
- Perfect for expanding analysis types

#### C. Edit-Series Function
- View all current series with column mappings
- Add new series with custom X/Y column pairs
- Edit existing series (name, X column, Y column)
- Delete series by index

### 2. Enhanced ConfigLoader Module

New functions added to `modules/ConfigLoader.psm1`:

```powershell
Get-ArrayFromConfig          # Retrieve arrays from config
Update-ArrayInConfig         # Update arrays in config
Add-SeriesConfig            # Add new chart series
Remove-SeriesConfig         # Remove chart series by index
```

### 3. Interactive Runtime Prompts

#### Copy-PlotData Script Enhancements
- **File Type Selection**: Prompts user to select which file types (density, pressure, velocity, temperature) to process
- **Category Selection**: Prompts user to select which categories (base plots, backward plots, etc.) to include
- Only processes selected combinations

#### Create-ComparisonPlots Script Enhancements
- **Output Folder Prompt**: Asks where to save comparison plots
- **File Filter**: Allows specifying which files to process (e.g., `comparison_*.xlsx`)

#### Inspect-Excel Script Enhancements
- **Pattern Filtering**: Ask for file pattern to inspect specific files
- **Default Options**: Uses default pattern if user just presses Enter

### 4. New Documentation

- **INTERACTIVE_FEATURES.md**: Complete guide to new interactive features
  - Use cases and examples
  - Menu navigation flow
  - Troubleshooting tips
  - Configuration management examples

## Menu Structure Changes

### Before
```
Configuration Menu:
  1. Root Path
  2. Data Source Path
  3. Inspect Folder
  4. Chart Position & Size
  5. Inspection Wait Time
  0. Back
```

### After
```
Configuration Menu:
  1. Root Path
  2. Data Source Path
  3. Inspect Folder
  4. Chart Position & Size
  5. Inspection Wait Time
  6. Manage Categories          [NEW]
  7. Manage File Types           [NEW]
  8. Manage Chart Series        [NEW]
  0. Back
```

## Configuration File Status

The `config/settings.json` remains the same structure - no breaking changes. New menu options simply provide a user-friendly interface to edit existing configuration arrays and objects.

## Backward Compatibility

✅ **Fully backward compatible**
- All existing scripts work unchanged
- Existing config files work unchanged
- New features are purely additive
- No deprecated functionality

## User Experience Improvements

1. **No More JSON Editing**: Manage categories, types, and series via menu
2. **Selective Processing**: Choose which data to process at runtime
3. **Visual Feedback**: Color-coded output with clear prompts
4. **Immediate Persistence**: All changes saved immediately
5. **Input Validation**: Invalid inputs caught and re-prompted

## Tested Features

✅ Adding new categories
✅ Deleting categories
✅ Adding new file types
✅ Deleting file types
✅ Adding chart series
✅ Editing chart series
✅ Deleting chart series
✅ File type selection in Copy-PlotData
✅ Category selection in Copy-PlotData
✅ Output folder prompts in Create-ComparisonPlots
✅ File pattern filtering in Inspect-Excel

## Configuration Persistence Examples

### Adding a New File Type
```
Menu Option 5 → 7
Enter "turbulence"
Updated in: config/settings.json → filePatterns.fileTypes
Immediately available for selection in Copy-PlotData
```

### Adding a Chart Series
```
Menu Option 5 → 8
Enter: name="Temperature", X="M", Y="N"
Updated in: config/settings.json → chart.series
All comparison plots will include this series
```

### Selective Processing
```
Script: Copy-PlotData
Select file types: density, velocity (skip pressure, temperature)
Select categories: base, backward (skip forward)
Only these combinations are consolidated
```

## File Changes

### Modified Files
- `Main.ps1`: Added 3 new functions, updated menu
- `modules/ConfigLoader.psm1`: Added 4 new functions
- `scripts/Copy-PlotData.ps1`: Added interactive selection prompts
- `scripts/Create-ComparisonPlots.ps1`: Added interactive prompts
- `scripts/Inspect-Excel.ps1`: Added pattern filtering prompt

### New Files
- `INTERACTIVE_FEATURES.md`: New documentation

### Unchanged Files
- `config/settings.json`: Same structure
- `modules/ExcelUtils.psm1`: No changes
- `modules/ChartUtils.psm1`: No changes
- `scripts/Convert-CsvToXlsx.ps1`: No changes
- `README.md`: No changes (new features documented separately)
- `QUICKSTART.md`: No changes
- `REFACTORING.md`: No changes

## API Documentation

### New Configuration Functions

```powershell
# Get an array from configuration
$array = Get-ArrayFromConfig -ConfigPath "config/settings.json" -Path "filePatterns.fileTypes"

# Update an array in configuration
Update-ArrayInConfig -ConfigPath "config/settings.json" -Path "filePatterns.fileTypes" -Array @("density", "pressure")

# Add a chart series
Add-SeriesConfig -ConfigPath "config/settings.json" -SeriesName "NewSeries" -XColumn "A" -YColumn "B"

# Remove a chart series
Remove-SeriesConfig -ConfigPath "config/settings.json" -Index 2
```

### New Menu Functions

```powershell
# Manage categories
Edit-Categories

# Manage file types
Edit-FileTypes

# Manage chart series
Edit-Series
```

## Quick Start for New Users

1. Run `.\Main.ps1`
2. Select option 5 (Edit Configuration)
3. Select option 6-8 to configure categories, file types, and series
4. Run operations with their interactive prompts
5. Make selections to process specific data combinations

## Migration Guide

### For Existing Users

No migration needed! Your existing `config/settings.json` file works as-is. Simply:

1. Update to the new version
2. Run Main.ps1
3. Explore the new menu options (6, 7, 8)
4. No changes required to your existing configurations

### For Script Customization

If you have custom scripts using ConfigLoader module, import the new functions:

```powershell
Import-Module .\modules\ConfigLoader.psm1 -Force

# Now you can use:
$types = Get-ArrayFromConfig -ConfigPath "config/settings.json" -Path "filePatterns.fileTypes"
```

## Next Steps

The software now supports:
- ✅ Centralized configuration management
- ✅ Interactive menu editing
- ✅ Runtime input selection
- ✅ Array and object manipulation
- ✅ Persistent changes

Future enhancements could include:
- Multi-select interactive prompts
- Configuration profiles (save/load multiple configs)
- Batch operations
- Scheduling support
- Report generation

## Support

For detailed usage, see `INTERACTIVE_FEATURES.md` which includes:
- Use case examples
- Troubleshooting guide
- Configuration best practices
- Tips and tricks

---

**Version**: 2.0
**Date**: April 24, 2026
**Status**: Production Ready
