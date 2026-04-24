<#
.SYNOPSIS
Convert CSV files to XLSX format with charts

.PARAMETER ConfigPath
Path to configuration file
#>

param(
    [string]$ConfigPath = "..\config\settings.json"
)

# Import modules
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulePath = Join-Path (Split-Path -Parent $scriptPath) "modules"
Import-Module (Join-Path $modulePath "ConfigLoader.psm1") -Force
Import-Module (Join-Path $modulePath "ExcelUtils.psm1") -Force
Import-Module (Join-Path $modulePath "ChartUtils.psm1") -Force

# Load configuration
$configPath = Join-Path (Split-Path -Parent $scriptPath) "config\settings.json"
$config = Get-ConfigurationSettings -ConfigPath $configPath
Validate-Configuration -Config $config

# Initialize logging
$logColor = $config.logging.colors

try {
    Write-Host "Starting CSV to XLSX Conversion..." -ForegroundColor $logColor.info
    
    # Initialize Excel
    $excel = Initialize-ExcelApplication `
        -Visible $config.excel.visible `
        -DisplayAlerts $config.excel.displayAlerts
    
    # Get all CSV files
    $files = Get-ChildItem -File -Filter $config.filePatterns.csvPattern -Recurse
    
    if ($files.Count -eq 0) {
        Write-Host "No CSV files found." -ForegroundColor $logColor.warning
    }
    else {
        Write-Host "Found $($files.Count) CSV file(s)" -ForegroundColor $logColor.info
    }
    
    foreach ($file in $files) {
        try {
            Write-Host "Processing: $($file.FullName)" -ForegroundColor $logColor.info
            
            # Open workbook
            $workbook = Open-ExcelWorkbook -FilePath $file.FullName -ExcelApp $excel
            $sheet = Get-ActiveSheet -Workbook $workbook
            
            # Remove existing charts
            Remove-ExistingCharts -Sheet $sheet
            
            # Delete header rows
            $rowsToDelete = $config.dataProcessing.rowsToDelete
            Delete-Rows -Sheet $sheet -RowNumbers $rowsToDelete
            
            # Set column formats
            Set-ColumnFormat -Sheet $sheet `
                -Columns @("A", "B") `
                -Format $config.dataProcessing.columnFormat
            
            # Get last row with data
            $lastRow = Get-LastRowWithData -Sheet $sheet -Column "A"
            
            if ($lastRow -lt $config.dataProcessing.headerRow + 1) {
                Write-Host "  Insufficient data after row deletion." -ForegroundColor $logColor.warning
                Close-ExcelWorkbook -Workbook $workbook -SaveChanges $false
                continue
            }
            
            # Get header for chart title
            $headerText = $sheet.Range("A1").Text
            
            # Create chart
            $chart = Build-ChartFromConfig -Sheet $sheet `
                -Config $config.chart `
                -Title "$headerText vs $($config.chart.axes.yTitle)" `
                -LastRow $lastRow
            
            # Save workbook
            $outputPath = Join-Path $file.DirectoryName ($file.BaseName + ".xlsx")
            Save-ExcelWorkbook -Workbook $workbook `
                -Path $outputPath `
                -FileFormat $config.excel.fileFormat
            
            Close-ExcelWorkbook -Workbook $workbook -SaveChanges $false
            Write-Host "  Success: $($file.BaseName).xlsx" -ForegroundColor $logColor.success
        }
        catch {
            Write-Host "  Failed to process $($file.Name): $_" -ForegroundColor $logColor.error
            if ($workbook) { Close-ExcelWorkbook -Workbook $workbook -SaveChanges $false }
        }
    }
    
    # Cleanup Excel
    Close-ExcelApplication -ExcelApp $excel
    
    Write-Host "Batch Processing Complete!" -ForegroundColor $logColor.success
}
catch {
    Write-Host "Error: $_" -ForegroundColor $logColor.error
    if ($excel) { Close-ExcelApplication -ExcelApp $excel }
    exit 1
}
