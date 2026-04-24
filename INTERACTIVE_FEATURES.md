# Enhanced Menu System - User Input Features

## Overview

The Excel Automation Software has been enhanced with interactive configuration management and runtime input capabilities. Users can now configure categories, file types, and chart series directly from the menu, and provide input during script execution.

## New Features

### 1. Configuration Menu Enhancements

The configuration editor now includes three new sections:

#### Option 6: Manage Categories

Allows you to add, view, and delete data categories (e.g., "base plots", "backward plots", "forward plots").

**Features:**
- View all current categories
- Add new categories dynamically
- Delete categories by index
- Changes saved to `config/settings.json` immediately

**Example Usage:**
```
Current Categories:
  [0] base plots
  [1] backward plots
  [2] forward plots

Options:
  A - Add new category
  D - Delete category (by index)
  0 - Back to Configuration Menu

Enter your choice: A
Enter new category name: custom plots
Added: custom plots
```

#### Option 7: Manage File Types

Allows you to add, view, and delete file types (e.g., "density", "pressure", "velocity").

**Features:**
- View all current file types
- Add new file types dynamically
- Delete file types by index
- Perfect for adding new analysis types

**Example Usage:**
```
Current File Types:
  [0] density
  [1] pressure
  [2] velocity
  [3] temperature

Options:
  A - Add new file type
  D - Delete file type (by index)
  0 - Back to Configuration Menu

Enter your choice: A
Enter new file type (e.g., density, pressure): turbulence
Added: turbulence
```

#### Option 8: Manage Chart Series

Allows you to add, edit, and delete chart series configurations.

**Features:**
- View all current series (name and column mappings)
- Add new series with custom X and Y columns
- Edit existing series
- Delete series by index
- Changes reflected in all chart operations

**Example Usage:**
```
Current Series:
  [0] Backward - X:A Y:B
  [1] Strut A - X:D Y:E
  [2] Strut D - X:G Y:H

Options:
  A - Add new series
  D - Delete series (by index)
  E - Edit series (by index)
  0 - Back to Configuration Menu

Enter your choice: A
Enter series name: Custom Series
Enter X column letter (A-Z): J
Enter Y column letter (A-Z): K
Added series: Custom Series
```

### 2. Runtime Input Features

#### Copy-PlotData Script - Interactive Selection

When running the Copy-PlotData operation, the script now prompts you to select which data to process:

**Step 1: Select File Types**
```
═══════════════════════════════════════════
   Select File Types to Process
═══════════════════════════════════════════

Process 'density'? (Y/N): Y
Process 'pressure'? (Y/N): N
Process 'velocity'? (Y/N): Y
Process 'temperature'? (Y/N): N

Selected file types: density, velocity
```

**Step 2: Select Categories**
```
═══════════════════════════════════════════
   Select Categories to Process
═══════════════════════════════════════════

Include 'base plots'? (Y/N): Y
Include 'backward plots'? (Y/N): Y
Include 'forward plots'? (Y/N): N

Selected categories: base plots, backward plots
```

Only the selected file types and categories will be processed, giving you fine-grained control.

#### Create-ComparisonPlots Script - Interactive Configuration

When running the Create-ComparisonPlots operation:

```
Enter output folder path (or press Enter for current folder): C:\Output
Enter file filter pattern (default: *.xlsx): 
Output Folder: C:\Output
File Filter: *.xlsx
```

#### Inspect-Excel Script - Interactive Filtering

When running the Inspect-Excel operation:

```
Filter files by file pattern? (e.g., *.xlsx, default shows all) (Y/N or pattern): comparison_*.xlsx
Using filter: comparison_*.xlsx
```

This allows you to inspect only specific files instead of all files in a folder.

## Use Cases

### Use Case 1: Process Specific Analysis Types

**Scenario:** You have density, pressure, velocity, and temperature files, but only want to consolidate density and velocity data.

**Solution:**
```
Run Main.ps1 → Select 3 (Copy and Consolidate Plot Data)
Select File Types: Y for density, N for pressure, Y for velocity, N for temperature
Select Categories: Y for all
```

### Use Case 2: Add New Chart Series

**Scenario:** You need to add a new data series to your comparison charts representing a new measurement type.

**Solution:**
```
Run Main.ps1 → Select 5 (Edit Configuration) → Select 8 (Manage Chart Series)
Add new series with name, X column, and Y column
All future comparison plots will include this new series
```

### Use Case 3: Test New Category

**Scenario:** You have a new data category "experimental plots" and want to add it temporarily.

**Solution:**
```
Run Main.ps1 → Select 5 (Edit Configuration) → Select 6 (Manage Categories)
Add "experimental plots"
Run Copy-PlotData and select to include "experimental plots"
```

### Use Case 4: Compare Different File Versions

**Scenario:** You have multiple versions of files and want to inspect only the latest ones.

**Solution:**
```
Run Main.ps1 → Select 4 (Inspect Excel Files)
Enter folder path
Filter by pattern: latest_*.xlsx or v2_*.xlsx
```

## Menu Navigation Flow

```
Main Menu
├─ 1. Convert CSV to XLSX
├─ 2. Create Comparison Plots
├─ 3. Copy and Consolidate Plot Data
│   └─ Prompts for file types and categories
├─ 4. Inspect Excel Files
│   └─ Prompts for file pattern filter
├─ 5. Edit Configuration
│   ├─ 1-5. Path and simple settings
│   ├─ 6. Manage Categories (Add/Delete)
│   ├─ 7. Manage File Types (Add/Delete)
│   └─ 8. Manage Chart Series (Add/Edit/Delete)
├─ 6. Display Configuration
└─ 0. Exit
```

## Examples of Configuration Changes

### Adding a New Series for Additional Measurement

```
Enter your choice: 5
Enter your choice: 8
Enter your choice: A
Enter series name: Temperature Sensor 1
Enter X column letter (A-Z): M
Enter Y column letter (A-Z): N
Added series: Temperature Sensor 1
```

### Creating a New Data Category

```
Enter your choice: 5
Enter your choice: 6
Enter your choice: A
Enter new category name: validation data
Added: validation data
```

### Managing File Type for New Analysis

```
Enter your choice: 5
Enter your choice: 7
Enter your choice: A
Enter new file type (e.g., density, pressure): strain
Added: strain
```

## Technical Details

### Configuration Persistence

All changes made through the menu are immediately saved to `config/settings.json`. The changes persist across script runs and system restarts.

### Data Types

- **Categories**: String array
- **File Types**: String array  
- **Series**: Array of objects with `name`, `xColumn`, and `yColumn` properties

### Validation

- File type names and category names are trimmed of whitespace
- Column letters are validated to be single characters A-Z
- Series indices are validated to be within range
- File paths are validated for existence before use

### Error Handling

All interactive menus include error handling:
- Invalid indices show "Invalid index" message
- Invalid file types show "Invalid input" message
- Missing paths show "Path not found" message
- All errors are color-coded red for visibility

## Tips & Tricks

### Tip 1: Quick Configuration Review
Before running operations, always use option 6 (Display Configuration) to verify all settings are correct.

### Tip 2: Backup Configuration
Save a copy of `config/settings.json` before major configuration changes:
```powershell
Copy-Item config/settings.json config/settings_backup.json
```

### Tip 3: Selective Processing
Use the runtime selection prompts to process subsets of data without changing the main configuration.

### Tip 4: Series Organization
When managing series, maintain a consistent naming convention and column ordering for easier maintenance.

## Troubleshooting

### Issue: Changes not appearing in next run

**Solution:** Verify changes were saved by checking `config/settings.json` file modification time, or display configuration (option 6).

### Issue: Invalid category/file type selected

**Solution:** Ensure category and file type folders exist in the data source path. Use the option to add categories/types, then ensure corresponding folders exist.

### Issue: Series columns don't match data

**Solution:** Use option 8 to edit series and update column letters to match your actual data layout.

## Summary

The enhanced menu system provides complete flexibility to:
- Manage data categories without editing JSON
- Add new file types for analysis
- Configure custom chart series
- Make selective choices at runtime
- Maintain configurations persistently

All while keeping the main focus on automation and ease of use.
