<#
.SYNOPSIS
Visually inspect Excel files by opening them sequentially

.PARAMETER ConfigPath
Path to configuration file

.PARAMETER FolderPath
Folder containing Excel files to inspect
#>

param(
    [string]$ConfigPath = "..\config\settings.json",
    [string]$FolderPath = (Read-Host -Prompt "Enter folder path to inspect")
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
    if (-not (Test-Path $FolderPath)) {
        throw "Folder not found: $FolderPath"
    }
    
    Write-Host "Starting Visual Inspection..." -ForegroundColor $logColor.info
    Write-Host "Folder: $FolderPath" -ForegroundColor $logColor.info
    Write-Host ""
    
    # Ask if user wants to filter by file pattern
    $usePattern = Read-Host "Filter files by pattern? (e.g., *.xlsx, default shows all) (Y/N or pattern)"
    if ($usePattern -eq "N" -or $usePattern -eq "n") {
        $filePattern = "*"
    }
    elseif ($usePattern -eq "Y" -or $usePattern -eq "y" -or [string]::IsNullOrWhiteSpace($usePattern)) {
        $filePattern = $config.filePatterns.xlsxPattern
    }
    else {
        $filePattern = $usePattern
    }
    
    Write-Host "Using filter: $filePattern" -ForegroundColor $logColor.process
    Write-Host ""
    
    # Initialize Excel
    $excel = Initialize-ExcelApplication `
        -Visible $config.inspection.visible `
        -DisplayAlerts $config.excel.displayAlerts
    
    # Get all Excel files with pattern
    $files = Get-ChildItem -Path $FolderPath -Filter $filePattern
    
    if ($files.Count -eq 0) {
        Write-Host "No Excel files found in $FolderPath" -ForegroundColor $logColor.warning
    }
    else {
        Write-Host "Found $($files.Count) file(s) to inspect" -ForegroundColor $logColor.info
    }
    
    $waitSeconds = $config.inspection.waitTime
    
    foreach ($file in $files) {
        try {
            Write-Host "Opening: $($file.Name)" -ForegroundColor $logColor.info
            
            # Open workbook
            $workbook = Open-ExcelWorkbook -FilePath $file.FullName -ExcelApp $excel
            
            # Display for specified time
            Write-Host "  Holding for $waitSeconds seconds..." -ForegroundColor $logColor.process
            Start-Sleep -Seconds $waitSeconds
            
            # Close workbook
            Close-ExcelWorkbook -Workbook $workbook -SaveChanges $false
            Write-Host "  Closed: $($file.Name)" -ForegroundColor $logColor.success
        }
        catch {
            Write-Host "  Failed to process $($file.Name): $_" -ForegroundColor $logColor.error
            if ($workbook) { Close-ExcelWorkbook -Workbook $workbook -SaveChanges $false }
        }
    }
    
    # Cleanup
    Close-ExcelApplication -ExcelApp $excel
    
    Write-Host "Visual Inspection Complete!" -ForegroundColor $logColor.success
}
catch {
    Write-Host "Error: $_" -ForegroundColor $logColor.error
    if ($excel) { Close-ExcelApplication -ExcelApp $excel }
    exit 1
}
