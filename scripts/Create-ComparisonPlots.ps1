<#
.SYNOPSIS
Create comparison plots with multiple series

.PARAMETER ConfigPath
Path to configuration file

.PARAMETER OutputFolder
Output folder for comparison plots
#>

param(
    [string]$ConfigPath = "..\config\settings.json",
    [string]$OutputFolder = (Get-Location)
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

$logColor = $config.logging.colors

try {
    Write-Host "Starting Comparison Plots Creation..." -ForegroundColor $logColor.info
    Write-Host ""
    
    # Ask user for output folder if not specified
    if ([string]::IsNullOrWhiteSpace($OutputFolder) -or $OutputFolder -eq (Get-Location)) {
        $OutputFolder = Read-Host "Enter output folder path (or press Enter for current folder)"
        if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
            $OutputFolder = Get-Location
        }
    }
    
    # Ask if user wants to filter files
    $fileFilter = Read-Host "Enter file filter pattern (default: *.xlsx)"
    if ([string]::IsNullOrWhiteSpace($fileFilter)) {
        $fileFilter = $config.filePatterns.xlsxPattern
    }
    
    Write-Host ""
    Write-Host "Output Folder: $OutputFolder" -ForegroundColor $logColor.process
    Write-Host "File Filter: $fileFilter" -ForegroundColor $logColor.process
    Write-Host ""
    
    # Initialize Excel
    $excel = Initialize-ExcelApplication `
        -Visible $config.excel.visible `
        -DisplayAlerts $config.excel.displayAlerts
    
    # Get all XLSX files with filter
    $files = Get-ChildItem -File -Filter $fileFilter -Recurse
    
    if ($files.Count -eq 0) {
        Write-Host "No XLSX files found." -ForegroundColor $logColor.warning
    }
    else {
        Write-Host "Found $($files.Count) XLSX file(s)" -ForegroundColor $logColor.info
    }
    
    foreach ($file in $files) {
        try {
            Write-Host "Processing: $($file.FullName)" -ForegroundColor $logColor.info
            
            # Open workbook
            $workbook = Open-ExcelWorkbook -FilePath $file.FullName -ExcelApp $excel
            $sheet = Get-ActiveSheet -Workbook $workbook
            
            # Remove existing charts
            Remove-ExistingCharts -Sheet $sheet
            Write-Host "  Cleared existing charts" -ForegroundColor $logColor.process
            
            # Set column formats
            Set-ColumnFormat -Sheet $sheet `
                -Columns @("A", "B", "D", "E", "G", "H") `
                -Format $config.dataProcessing.columnFormat
            
            # Get last row with data
            $lastRow = Get-LastRowWithData -Sheet $sheet -Column "A"
            
            if ($lastRow -lt 2) {
                Write-Host "  Insufficient data in $($file.Name)" -ForegroundColor $logColor.warning
                Close-ExcelWorkbook -Workbook $workbook -SaveChanges $false
                continue
            }
            
            # Get header for chart title
            $headerText = $sheet.Range("A1").Text
            
            # Create chart using config
            $chart = Build-ChartFromConfig -Sheet $sheet `
                -Config $config.chart `
                -Title "$headerText vs $($config.chart.axes.yTitle) - Comparison" `
                -LastRow $lastRow
            
            # Save workbook
            $outputPath = Join-Path $OutputFolder ($file.BaseName + ".xlsx")
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
    
    # Cleanup
    Close-ExcelApplication -ExcelApp $excel
    
    Write-Host "Comparison Plots Creation Complete!" -ForegroundColor $logColor.success
}
catch {
    Write-Host "Error: $_" -ForegroundColor $logColor.error
    if ($excel) { Close-ExcelApplication -ExcelApp $excel }
    exit 1
}
