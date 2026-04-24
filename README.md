# Excel Automation Software Documentation

## Overview

This is a refactored Excel automation suite that consolidates multiple PowerShell scripts into a modular, configurable system. All hardcoded values have been removed and replaced with a centralized configuration file.

## Project Structure

```
Excel-Automate/
├── config/
│   └── settings.json           # Main configuration file (ALL SETTINGS HERE)
├── modules/
│   ├── ConfigLoader.psm1       # Configuration management utilities
│   ├── ExcelUtils.psm1         # Core Excel operations
│   └── ChartUtils.psm1         # Chart creation and management
├── scripts/
│   ├── Convert-CsvToXlsx.ps1   # CSV to XLSX conversion with charts
│   ├── Create-ComparisonPlots.ps1  # Multi-series chart creation
│   ├── Copy-PlotData.ps1       # Data consolidation from multiple sources
│   └── Inspect-Excel.ps1       # Visual inspection of Excel files
├── Main.ps1                    # Main orchestrator with menu interface
└── README.md                   # This file
```

## Getting Started

### Prerequisites

- Windows PowerShell 5.1 or higher
- Microsoft Excel installed
- Appropriate file permissions for the directories you're working with

### Running the Software

1. Open PowerShell
2. Navigate to the `Excel-Automate` directory
3. Run:
   ```powershell
   .\Main.ps1
   ```

This will open an interactive menu with all available operations.

## Configuration

### Configuration File (`config/settings.json`)

All settings are centralized in `config/settings.json`. This file contains:

- **paths**: File and folder paths
  - `rootPath`: Default root directory for operations
  - `dataSourcePath`: Path for data consolidation source files
  - `inspectFolder`: Default folder for file inspection

- **filePatterns**: File filtering patterns
  - `csvPattern`: Pattern for CSV files (default: `*.csv`)
  - `xlsxPattern`: Pattern for XLSX files (default: `*.xlsx`)
  - `fileTypes`: List of file types to process

- **categories**: Data categories for consolidation
  - Examples: "base plots", "backward plots", "forward plots"

- **chart**: Chart configuration
  - `type`: Excel chart type constant (75 = XY Scatter Lines)
  - `style`: Chart style ID
  - `position`: Chart placement (left, top, width, height)
  - `series`: Data series definitions with column mappings
  - `axes`: Axis titles and configuration
  - `hasLegend`: Whether to show legend
  - `legendPosition`: Legend position constant

- **dataProcessing**: Data processing settings
  - `rowsToDelete`: Rows to remove from CSV
  - `headerRow`: Header row number
  - `dataStartRow`: First data row number
  - `columnFormat`: Number format for columns

- **excel**: Excel application settings
  - `visible`: Show Excel window (false recommended for automation)
  - `displayAlerts`: Show Excel alerts
  - `fileFormat`: Output format (51 = .xlsx)

- **inspection**: File inspection settings
  - `waitTime`: Seconds to display each file
  - `visible`: Show Excel window during inspection

- **logging**: Output formatting
  - `verboseOutput`: Enable verbose output
  - `colors`: Console colors for different message types

### Editing Configuration

You have three options:

#### Option 1: Using the Menu
```powershell
.\Main.ps1
# Select option 5 (Edit Configuration)
```

#### Option 2: Direct File Edit
Open `config/settings.json` in a text editor and modify values directly. Remember to maintain valid JSON syntax.

#### Option 3: PowerShell Command
```powershell
# First, import the module
Import-Module .\modules\ConfigLoader.psm1

# Update a single value
Update-ConfigValue -ConfigPath "config\settings.json" -Path "paths.rootPath" -Value "C:\MyPath"
```

## Features & Operations

### 1. Convert CSV to XLSX with Charts

Converts CSV files to Excel format and automatically creates scatter line charts.

**What it does:**
- Recursively finds all CSV files
- Removes the first 4 rows (configurable)
- Creates XY scatter line charts
- Saves as XLSX format

**To run:**
```powershell
.\Main.ps1
# Select option 1
```

Or directly:
```powershell
.\scripts\Convert-CsvToXlsx.ps1
```

**Configuration options:**
- `dataProcessing.rowsToDelete`: Which rows to remove
- `chart.*`: Chart appearance and settings
- `filePatterns.csvPattern`: CSV file pattern

### 2. Create Comparison Plots

Creates multi-series comparison charts from XLSX files (typically for comparing different scenarios like "backward", "strut A", "strut D").

**What it does:**
- Processes existing XLSX files
- Creates charts with multiple data series
- Each series comes from different column pairs
- Removes previous charts before adding new ones

**To run:**
```powershell
.\Main.ps1
# Select option 2
```

Or directly:
```powershell
.\scripts\Create-ComparisonPlots.ps1 -OutputFolder "C:\Path\To\Output"
```

**Configuration options:**
- `chart.series`: Define each series with name and columns
- `chart.position`: Chart size and placement

### 3. Copy and Consolidate Plot Data

Merges data from multiple categories (base plots, backward plots, forward plots) into consolidated Excel files.

**What it does:**
- Reads from multiple category folders
- Copies X and Y data from each source
- Creates a new consolidated XLSX file
- Preserves data structure for analysis

**To run:**
```powershell
.\Main.ps1
# Select option 3
```

Or directly:
```powershell
.\scripts\Copy-PlotData.ps1
```

**Configuration options:**
- `paths.dataSourcePath`: Source directory
- `categories`: Which categories to process
- `filePatterns.fileTypes`: Which file types to consolidate

### 4. Inspect Excel Files

Sequentially opens Excel files for visual inspection.

**What it does:**
- Opens each file in Excel
- Displays for configurable duration
- Closes without saving
- Great for quality control

**To run:**
```powershell
.\Main.ps1
# Select option 4
# Enter folder path when prompted
```

Or directly:
```powershell
.\scripts\Inspect-Excel.ps1 -FolderPath "C:\Path\To\Files"
```

**Configuration options:**
- `inspection.waitTime`: Seconds to display each file
- `inspection.visible`: Show/hide Excel window

## Using the Modules Directly

The modules provide reusable functions for custom automation:

```powershell
Import-Module .\modules\ExcelUtils.psm1
Import-Module .\modules\ChartUtils.psm1
Import-Module .\modules\ConfigLoader.psm1

# Initialize Excel
$excel = Initialize-ExcelApplication -Visible $false

# Open a workbook
$workbook = Open-ExcelWorkbook -FilePath "C:\file.xlsx" -ExcelApp $excel

# Get active sheet
$sheet = Get-ActiveSheet -Workbook $workbook

# Get last row with data
$lastRow = Get-LastRowWithData -Sheet $sheet -Column "A"

# Create a chart
$chart = New-ExcelChart -Sheet $sheet -Left 100 -Top 50

# Close properly
Close-ExcelWorkbook -Workbook $workbook -SaveChanges $false
Close-ExcelApplication -ExcelApp $excel
```

### Available Functions

#### ExcelUtils.psm1
- `Initialize-ExcelApplication`: Create Excel COM object
- `Close-ExcelApplication`: Properly close Excel and release resources
- `Open-ExcelWorkbook`: Open an Excel file
- `Close-ExcelWorkbook`: Close a workbook
- `Save-ExcelWorkbook`: Save with specified format
- `Get-LastRowWithData`: Find last row with data
- `Get-ActiveSheet`: Get the active worksheet
- `Delete-Rows`: Remove rows from worksheet
- `Set-ColumnFormat`: Format columns
- `Remove-ExistingCharts`: Clear all charts
- `Copy-Range`: Copy and paste ranges

#### ChartUtils.psm1
- `New-ExcelChart`: Create a new chart object
- `Add-ChartSeries`: Add a data series to chart
- `Clear-ChartSeries`: Remove all series
- `Set-ChartTitle`: Set chart title
- `Set-ChartAxes`: Configure axes
- `Set-ChartLegend`: Configure legend
- `Build-ChartFromConfig`: Create complete chart from config

#### ConfigLoader.psm1
- `Get-ConfigurationSettings`: Load config from JSON
- `Update-ConfigValue`: Modify config value
- `Get-ConfigValue`: Get specific config value
- `Validate-Configuration`: Verify required settings exist

## Examples

### Example 1: Custom CSV Processing

```powershell
Import-Module .\modules\ExcelUtils.psm1
Import-Module .\modules\ConfigLoader.psm1

$config = Get-ConfigurationSettings -ConfigPath "config\settings.json"
$excel = Initialize-ExcelApplication -Visible $false

$file = Get-Item "myfile.csv"
$workbook = Open-ExcelWorkbook -FilePath $file.FullName -ExcelApp $excel
$sheet = Get-ActiveSheet -Workbook $workbook

# Delete first 4 rows
Delete-Rows -Sheet $sheet -RowNumbers @(1, 2, 3, 4)

# Save
Save-ExcelWorkbook -Workbook $workbook -Path "output.xlsx" -FileFormat 51

Close-ExcelWorkbook -Workbook $workbook
Close-ExcelApplication -ExcelApp $excel
```

### Example 2: Batch Process with Custom Config

```powershell
$configPath = "config\settings.json"

# Update config temporarily
Update-ConfigValue -ConfigPath $configPath -Path "inspection.waitTime" -Value 15

# Run operation
.\scripts\Inspect-Excel.ps1 -FolderPath "C:\MyFiles"

# Restore config
Update-ConfigValue -ConfigPath $configPath -Path "inspection.waitTime" -Value 10
```

### Example 3: Custom Script Using Modules

```powershell
Import-Module .\modules\ExcelUtils.psm1
Import-Module .\modules\ChartUtils.psm1
Import-Module .\modules\ConfigLoader.psm1

$config = Get-ConfigurationSettings -ConfigPath "config\settings.json"
$excel = Initialize-ExcelApplication

# Get files
$files = Get-ChildItem -Filter "*.xlsx"

foreach ($file in $files) {
    $workbook = Open-ExcelWorkbook -FilePath $file.FullName -ExcelApp $excel
    $sheet = Get-ActiveSheet -Workbook $workbook
    
    $lastRow = Get-LastRowWithData -Sheet $sheet
    
    # Build chart from config
    $chart = Build-ChartFromConfig -Sheet $sheet `
        -Config $config.chart `
        -Title "My Chart" `
        -LastRow $lastRow
    
    Save-ExcelWorkbook -Workbook $workbook -Path "$($file.DirectoryName)\processed_$($file.Name)"
    Close-ExcelWorkbook -Workbook $workbook
}

Close-ExcelApplication -ExcelApp $excel
```

## Troubleshooting

### Excel Hangs or Freezes
- Set `excel.visible` to `false` in config
- Ensure proper cleanup of COM objects
- Restart PowerShell if issues persist

### Files Not Found
- Check that `rootPath` in config is correct
- Verify file patterns match your files
- Ensure paths are absolute or relative to script location

### Chart Not Appearing
- Verify `lastRow` is greater than or equal to 2
- Check that data exists in specified columns
- Confirm chart position values aren't outside sheet bounds

### Macro Security Issues
- You may need to adjust Excel macro security settings
- Run PowerShell as Administrator
- Check Group Policy settings on corporate machines

### COM Object Errors
- Don't have Excel open manually while scripts run
- Ensure proper cleanup with `Close-ExcelApplication`
- Check that Excel is fully updated

## Customization Guide

### Adding a New Operation

1. Create new script in `scripts/` folder
2. Import required modules at the top
3. Load configuration
4. Add menu option in `Main.ps1`

### Modifying Chart Settings

Edit `config/settings.json`:
```json
"chart": {
  "series": [
    {
      "name": "Your Series",
      "xColumn": "A",
      "yColumn": "B"
    }
  ],
  "axes": {
    "xTitle": "Your X Label",
    "yTitle": "Your Y Label"
  }
}
```

### Creating a Custom Module

1. Create `.psm1` file in `modules/` folder
2. Write functions following PowerShell best practices
3. Export functions at the end:
   ```powershell
   Export-ModuleMember -Function MyFunction1, MyFunction2
   ```
4. Import in your scripts

## Performance Tips

1. Set `excel.visible` to `false` for faster processing
2. Process files recursively in batches
3. Use `Get-LastRowWithData` to avoid processing empty rows
4. Close workbooks immediately after use
5. Use verbose logging only for debugging

## Version History

- **v1.0** - Initial refactored release
  - Modular architecture
  - Centralized configuration
  - No hardcoded values
  - Interactive menu system

## Support

For issues or questions, review:
1. Configuration file for incorrect settings
2. Log output for specific error messages
3. Troubleshooting section above
4. Module documentation in code comments

## License

This software is provided as-is for Excel automation tasks.
