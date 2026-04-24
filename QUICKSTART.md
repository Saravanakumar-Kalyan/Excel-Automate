# Quick Start Guide

## Installation

1. Ensure you have the Excel-Automate folder with all files
2. No additional installation needed - all scripts are ready to use

## First Time Setup

### Step 1: Configure Your Paths

Open `config/settings.json` and update these paths to match your environment:

```json
"paths": {
  "rootPath": "C:\\your\\project\\plots",
  "dataSourcePath": "C:\\your\\data\\source\\plots",
  "inspectFolder": "C:\\your\\inspection\\folder"
}
```

### Step 2: Run the Main Program

```powershell
cd C:\path\to\Excel-Automate
.\Main.ps1
```

You'll see an interactive menu. Choose option 6 to verify your configuration is correct.

## Common Tasks

### Task 1: Convert CSV Files to Excel with Charts

```powershell
.\Main.ps1
# Select: 1 (Convert CSV to XLSX)
```

**What it does:**
- Finds all `.csv` files in current directory and subdirectories
- Removes first 4 rows (configurable)
- Creates an XY scatter chart
- Saves as `.xlsx` file

### Task 2: Compare Multiple Data Series

```powershell
.\Main.ps1
# Select: 2 (Create Comparison Plots)
```

**What it does:**
- Creates charts comparing "Backward", "Strut A", and "Strut D" series
- One chart per file
- Removes old charts before creating new ones

### Task 3: Consolidate Data from Multiple Sources

```powershell
.\Main.ps1
# Select: 3 (Copy and Consolidate Plot Data)
```

**What it does:**
- Reads data from base plots, backward plots, and forward plots
- Merges them into single Excel files
- Creates output files like `Data_density_filename.xlsx`

### Task 4: Review Excel Files

```powershell
.\Main.ps1
# Select: 4 (Inspect Excel Files)
# Enter the folder path to inspect
```

**What it does:**
- Opens each Excel file for 10 seconds (configurable)
- Automatically closes and moves to next file
- Perfect for quality control

## Quick Configuration Changes

Without editing JSON, use the menu:

```powershell
.\Main.ps1
# Select: 5 (Edit Configuration)
```

### Common adjustments:
- **Root Path**: Where files are located
- **Data Source Path**: Where to read data from
- **Inspection Wait Time**: How long to display each file
- **Chart Position**: Where to place chart on sheet

## Troubleshooting Quick Tips

| Problem | Solution |
|---------|----------|
| "No files found" | Check `paths.rootPath` in config |
| Excel hangs | Set `excel.visible` to `false` in config |
| Chart not showing | Verify data exists in columns A & B |
| Script error on startup | Ensure all folders and files are present |

## File Structure Verification

Make sure you have this structure:

```
Excel-Automate/
├── config/
│   └── settings.json
├── modules/
│   ├── ExcelUtils.psm1
│   ├── ChartUtils.psm1
│   └── ConfigLoader.psm1
├── scripts/
│   ├── Convert-CsvToXlsx.ps1
│   ├── Create-ComparisonPlots.ps1
│   ├── Copy-PlotData.ps1
│   └── Inspect-Excel.ps1
├── Main.ps1
└── README.md
```

## Example Workflow

1. Place CSV files in a folder
2. Run Main.ps1 → Select 1 → Creates XLSX files with charts
3. Run Main.ps1 → Select 2 → Creates comparison plots
4. Run Main.ps1 → Select 4 → Review the output

## Direct Script Usage (Advanced)

Run scripts individually without the menu:

```powershell
# Convert CSV files
.\scripts\Convert-CsvToXlsx.ps1

# Create comparison charts
.\scripts\Create-ComparisonPlots.ps1 -OutputFolder "C:\MyOutput"

# Consolidate data
.\scripts\Copy-PlotData.ps1

# Inspect files
.\scripts\Inspect-Excel.ps1 -FolderPath "C:\MyFiles"
```

## Getting More Help

- See `README.md` for complete documentation
- Check individual script help: `Get-Help .\scripts\Script-Name.ps1 -Full`
- Review configuration options in `config/settings.json`

## Next Steps

1. Configure your paths in `config/settings.json`
2. Test with a small batch of files
3. Adjust chart and data settings as needed
4. Automate with Windows Task Scheduler (if desired)

Enjoy your Excel automation!
