# Quick Reference: Interactive Menu Features

## Main Menu
```
╔════════════════════════════════════════╗
║   Excel Automation Software v1.0       ║
╚════════════════════════════════════════╝

Select an operation:

  1. Convert CSV to XLSX (with charts)
  2. Create Comparison Plots
  3. Copy and Consolidate Plot Data        ← Runtime prompts for file types & categories
  4. Inspect Excel Files                   ← Runtime prompt for file filter
  5. Edit Configuration                    ← NEW OPTIONS 6, 7, 8
  6. Display Configuration
  0. Exit
```

## Configuration Menu (Option 5)
```
╔════════════════════════════════════════╗
║       Configuration Editor              ║
╚════════════════════════════════════════╝

Edit which setting?

  1. Root Path                             (existing)
  2. Data Source Path                      (existing)
  3. Inspect Folder                        (existing)
  4. Chart Position & Size                 (existing)
  5. Inspection Wait Time                  (existing)
  6. Manage Categories                     ★ NEW
  7. Manage File Types                     ★ NEW
  8. Manage Chart Series                   ★ NEW
  0. Back to Main Menu
```

## Manage Categories Menu (Option 6)
```
╔════════════════════════════════════════╗
║       Manage Categories                  ║
╚════════════════════════════════════════╝

Current Categories:
  [0] base plots
  [1] backward plots
  [2] forward plots

Options:
  A - Add new category
  D - Delete category (by index)
  0 - Back to Configuration Menu
```

### Actions:
- **Add (A)**: Input new category name → Saved to config
- **Delete (D)**: Input index → Category removed
- **Back (0)**: Return to Configuration Menu

## Manage File Types Menu (Option 7)
```
╔════════════════════════════════════════╗
║       Manage File Types                 ║
╚════════════════════════════════════════╝

Current File Types:
  [0] density
  [1] pressure
  [2] velocity
  [3] temperature

Options:
  A - Add new file type
  D - Delete file type (by index)
  0 - Back to Configuration Menu
```

### Actions:
- **Add (A)**: Input file type name (e.g., "strain") → Saved to config
- **Delete (D)**: Input index → File type removed
- **Back (0)**: Return to Configuration Menu

## Manage Chart Series Menu (Option 8)
```
╔════════════════════════════════════════╗
║       Manage Chart Series               ║
╚════════════════════════════════════════╝

Current Series:
  [0] Backward - X:A Y:B
  [1] Strut A - X:D Y:E
  [2] Strut D - X:G Y:H

Options:
  A - Add new series
  D - Delete series (by index)
  E - Edit series (by index)
  0 - Back to Configuration Menu
```

### Actions:
- **Add (A)**: Input series name + X column (A-Z) + Y column (A-Z) → Saved to config
- **Delete (D)**: Input index → Series removed
- **Edit (E)**: Input index, then update name/X/Y columns → Saved to config
- **Back (0)**: Return to Configuration Menu

## Copy and Consolidate Plot Data (Option 3) - Runtime Prompts
```
═══════════════════════════════════════════
   Select File Types to Process
═══════════════════════════════════════════

Process 'density'? (Y/N): Y
Process 'pressure'? (Y/N): Y
Process 'velocity'? (Y/N): N
Process 'temperature'? (Y/N): N

Selected file types: density, pressure

═══════════════════════════════════════════
   Select Categories to Process
═══════════════════════════════════════════

Include 'base plots'? (Y/N): Y
Include 'backward plots'? (Y/N): Y
Include 'forward plots'? (Y/N): Y

Selected categories: base plots, backward plots, forward plots
```

### Input Format:
- File Types: Answer Y/N for each configured type
- Categories: Answer Y/N for each configured category
- Only selected combinations are processed

## Create Comparison Plots (Option 2) - Runtime Prompts
```
Enter output folder path (or press Enter for current folder): 
C:\MyOutput

Enter file filter pattern (default: *.xlsx): 
comparison_*.xlsx

Output Folder: C:\MyOutput
File Filter: comparison_*.xlsx
```

### Input Format:
- Output folder: Full path or press Enter for current
- File filter: Wildcard pattern or press Enter for *.xlsx

## Inspect Excel Files (Option 4) - Runtime Prompts
```
Filter files by file pattern? (e.g., *.xlsx, default shows all) (Y/N or pattern): 
latest_*.xlsx

Using filter: latest_*.xlsx
```

### Input Format:
- Y/N for yes/no
- Direct pattern (e.g., latest_*.xlsx, v2_*)
- Press Enter for default *.xlsx

## Configuration Changes Persistence

All configuration changes made through the menu are automatically saved to:
```
config/settings.json
```

Structure:
```json
{
  "categories": ["base plots", "backward plots", ...],
  "filePatterns": {
    "fileTypes": ["density", "pressure", ...]
  },
  "chart": {
    "series": [
      { "name": "...", "xColumn": "A", "yColumn": "B" },
      ...
    ]
  }
}
```

## Common Workflows

### Workflow 1: Add New Analysis Type
1. Main Menu → 5 (Edit Configuration)
2. Configuration Menu → 7 (Manage File Types)
3. Select A (Add new file type)
4. Enter type name (e.g., "turbulence")
5. Done! Type available for selection in data consolidation

### Workflow 2: Create New Data Category
1. Main Menu → 5 (Edit Configuration)
2. Configuration Menu → 6 (Manage Categories)
3. Select A (Add new category)
4. Enter category name (e.g., "experimental data")
5. Done! Category available for selection in data consolidation

### Workflow 3: Add New Measurement Series
1. Main Menu → 5 (Edit Configuration)
2. Configuration Menu → 8 (Manage Chart Series)
3. Select A (Add new series)
4. Enter: name, X column, Y column
5. Done! Series included in all future comparison plots

### Workflow 4: Selective Data Consolidation
1. Main Menu → 3 (Copy and Consolidate Plot Data)
2. Script prompts for file types → Select only needed ones
3. Script prompts for categories → Select only needed ones
4. Only selected combinations are processed

### Workflow 5: Review Configuration
1. Main Menu → 6 (Display Configuration)
2. View complete JSON configuration
3. Verify all settings are correct before running operations

## Input Validation

The system validates all inputs:

| Input Type | Validation | Error Message |
|-----------|-----------|---|
| File Type Name | Non-empty | "Invalid input" |
| Category Name | Non-empty | "Invalid input" |
| Series Index | 0-based integer within range | "Invalid index" |
| Column Letter | Single A-Z character | "Invalid input" |
| Series Name | Non-empty | "Invalid input" |
| File Pattern | Any pattern | Uses as-is |
| Folder Path | Must exist on disk | "Folder not found" |

## Keyboard Shortcuts

```
Common Responses:
Y or y     → Yes/True
N or n     → No/False
0          → Back/Previous menu
A or a     → Add
D or d     → Delete
E or e     → Edit
Enter      → Use default
```

## Color Legend

```
Cyan       ♦ Headers and section titles
Green      ♦ Success messages
Yellow     ♦ Warnings
Red        ♦ Errors
White      ♦ Options and main text
Gray       ♦ Process details
```

## Tips & Best Practices

### ✓ Do's
- ✓ Verify configuration before major operations
- ✓ Use meaningful names for categories and series
- ✓ Backup config.json before bulk changes
- ✓ Test with small datasets first
- ✓ Use selective processing with runtime prompts

### ✗ Don'ts
- ✗ Don't manually edit JSON while Main.ps1 is running
- ✗ Don't delete all categories or file types
- ✗ Don't use special characters in names
- ✗ Don't forget to press Y or N for prompts

## Troubleshooting Quick Tips

| Issue | Solution |
|-------|----------|
| "Invalid index" | Check series/category count, use correct number |
| Changes not saved | Check config.json modification time |
| Files not found | Verify folder paths and file patterns |
| Series not appearing | Edit series to ensure columns exist in data |
| Prompt not responding | Press Y or N, not Enter alone |

## Where to Learn More

- `README.md` - Complete documentation
- `INTERACTIVE_FEATURES.md` - Detailed feature guide
- `QUICKSTART.md` - Getting started guide
- `ENHANCEMENT_V2.md` - What's new in v2

---

**Quick Reference Card v2.0**
Print this for quick access while using the software!
