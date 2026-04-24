<#
.SYNOPSIS
Copy and consolidate plot data from multiple categories

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

# Load configuration
$configPath = Join-Path (Split-Path -Parent $scriptPath) "config\settings.json"
$config = Get-ConfigurationSettings -ConfigPath $configPath
Validate-Configuration -Config $config

$logColor = $config.logging.colors

try {
    Write-Host "Starting Data Consolidation..." -ForegroundColor $logColor.info
    
    $basePath = $config.paths.dataSourcePath
    $categories = $config.categories
    $fileTypes = $config.filePatterns.fileTypes
    
    if (-not (Test-Path $basePath)) {
        throw "Base path not found: $basePath"
    }
    
    # Ask user to select which file types to process
    Clear-Host
    Write-Host "═══════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "   Select File Types to Process" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    
    $selectedTypes = @()
    for ($i = 0; $i -lt $fileTypes.Count; $i++) {
        $selection = Read-Host "Process '$($fileTypes[$i])'? (Y/N)" -ForegroundColor $logColor.info
        if ($selection -eq "Y" -or $selection -eq "y") {
            $selectedTypes += $fileTypes[$i]
        }
    }
    
    if ($selectedTypes.Count -eq 0) {
        Write-Host "No file types selected." -ForegroundColor $logColor.warning
        Write-Host "Press any key to continue..." -ForegroundColor $logColor.info
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }
    
    Write-Host ""
    Write-Host "Selected file types: $($selectedTypes -join ', ')" -ForegroundColor $logColor.success
    Write-Host ""
    
    # Ask user to select which categories to process
    Clear-Host
    Write-Host "═══════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "   Select Categories to Process" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    
    $selectedCategories = @()
    for ($i = 0; $i -lt $categories.Count; $i++) {
        $selection = Read-Host "Include '$($categories[$i])'? (Y/N)"
        if ($selection -eq "Y" -or $selection -eq "y") {
            $selectedCategories += $categories[$i]
        }
    }
    
    if ($selectedCategories.Count -eq 0) {
        Write-Host "No categories selected." -ForegroundColor $logColor.warning
        Write-Host "Press any key to continue..." -ForegroundColor $logColor.info
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }
    
    Write-Host ""
    Write-Host "Selected categories: $($selectedCategories -join ', ')" -ForegroundColor $logColor.success
    Write-Host ""
    
    # Initialize Excel
    $excel = Initialize-ExcelApplication `
        -Visible $config.excel.visible `
        -DisplayAlerts $config.excel.displayAlerts
    
    foreach ($type in $selectedTypes) {
        Write-Host "Processing type: $type" -ForegroundColor $logColor.info
        
        $typeBasePath = Join-Path $basePath $selectedCategories[0] $type
        
        if (-not (Test-Path $typeBasePath)) {
            Write-Host "  Path not found: $typeBasePath" -ForegroundColor $logColor.warning
            continue
        }
        
        $filesList = Get-ChildItem -Path $typeBasePath -Filter $config.filePatterns.xlsxPattern
        
        if ($filesList.Count -eq 0) {
            Write-Host "  No files found in $typeBasePath" -ForegroundColor $logColor.warning
            continue
        }
        
        foreach ($file in $filesList) {
            Write-Host "  Creating consolidated data for: $($file.BaseName)" -ForegroundColor $logColor.info
            
            try {
                # Create master workbook
                $masterWorkbook = $excel.Workbooks.Add()
                $masterSheet = Get-ActiveSheet -Workbook $masterWorkbook
                
                $colOffset = 1  # Start at column A
                
                foreach ($category in $selectedCategories) {
                    $sourcePath = Join-Path $basePath $category $type $($file.Name)
                    Write-Host "    Reading: $category\$type\$($file.Name)" -ForegroundColor $logColor.process
                    
                    if (Test-Path $sourcePath) {
                        $tempWorkbook = Open-ExcelWorkbook -FilePath $sourcePath -ExcelApp $excel
                        $tempSheet = Get-ActiveSheet -Workbook $tempWorkbook
                        $lastRow = Get-LastRowWithData -Sheet $tempSheet -Column "A"
                        
                        if ($lastRow -ge 2) {
                            # Copy header and data from column A
                            $rangeA = "A1:A$lastRow"
                            $tempSheet.Range($rangeA).Copy() | Out-Null
                            $targetCell = $masterSheet.Cells.Item(1, $colOffset)
                            $masterSheet.Paste($targetCell) | Out-Null
                            $colOffset++
                            
                            # Copy data from column B
                            $rangeB = "B1:B$lastRow"
                            $tempSheet.Range($rangeB).Copy() | Out-Null
                            $targetCell = $masterSheet.Cells.Item(1, $colOffset)
                            $masterSheet.Paste($targetCell) | Out-Null
                            $colOffset++
                        }
                        else {
                            Write-Host "      Insufficient data" -ForegroundColor $logColor.warning
                        }
                        
                        Close-ExcelWorkbook -Workbook $tempWorkbook -SaveChanges $false
                    }
                    else {
                        Write-Host "      File not found" -ForegroundColor $logColor.warning
                    }
                    
                    # Add spacing between categories
                    $colOffset++
                }
                
                # Save consolidated data
                $outputFile = Join-Path $basePath "Data_$($type)_$($file.BaseName).xlsx"
                Save-ExcelWorkbook -Workbook $masterWorkbook `
                    -Path $outputFile `
                    -FileFormat $config.excel.fileFormat
                
                Close-ExcelWorkbook -Workbook $masterWorkbook -SaveChanges $false
                Write-Host "  Created: $outputFile" -ForegroundColor $logColor.success
            }
            catch {
                Write-Host "  Failed to process $($file.Name): $_" -ForegroundColor $logColor.error
                if ($masterWorkbook) { Close-ExcelWorkbook -Workbook $masterWorkbook -SaveChanges $false }
            }
        }
    }
    
    # Cleanup
    Close-ExcelApplication -ExcelApp $excel
    
    Write-Host "Data Consolidation Complete!" -ForegroundColor $logColor.success
}
catch {
    Write-Host "Error: $_" -ForegroundColor $logColor.error
    if ($excel) { Close-ExcelApplication -ExcelApp $excel }
    exit 1
}
