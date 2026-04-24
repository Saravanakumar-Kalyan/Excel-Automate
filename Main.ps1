<#
.SYNOPSIS
Excel Automation Software - Main Orchestrator

This script provides a menu-driven interface for Excel automation tasks
including CSV conversion, comparison plots, data consolidation, and inspection.
#>

# Import modules
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulePath = Join-Path $scriptPath "modules"
Import-Module (Join-Path $modulePath "ConfigLoader.psm1") -Force

# Load configuration
$configPath = Join-Path $scriptPath "config\settings.json"
$config = Get-ConfigurationSettings -ConfigPath $configPath

$logColor = $config.logging.colors

function Show-Menu {
    Clear-Host
    Write-Host "╔════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   Excel Automation Software v1.0       ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Select an operation:" -ForegroundColor $logColor.info
    Write-Host ""
    Write-Host "  1. Convert CSV to XLSX (with charts)" -ForegroundColor White
    Write-Host "  2. Create Comparison Plots" -ForegroundColor White
    Write-Host "  3. Copy and Consolidate Plot Data" -ForegroundColor White
    Write-Host "  4. Inspect Excel Files" -ForegroundColor White
    Write-Host "  5. Edit Configuration" -ForegroundColor White
    Write-Host "  6. Display Configuration" -ForegroundColor White
    Write-Host "  0. Exit" -ForegroundColor White
    Write-Host ""
}

function Invoke-ScriptSafely {
    param(
        [string]$ScriptName,
        [string[]]$Parameters
    )
    
    $scriptFile = Join-Path $scriptPath "scripts\$ScriptName"
    
    if (-not (Test-Path $scriptFile)) {
        Write-Host "Script not found: $scriptFile" -ForegroundColor $logColor.error
        return $false
    }
    
    try {
        if ($Parameters) {
            & $scriptFile @Parameters
        }
        else {
            & $scriptFile
        }
        Write-Host ""
        Write-Host "Press any key to continue..." -ForegroundColor $logColor.info
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return $true
    }
    catch {
        Write-Host "Error executing $ScriptName : $_" -ForegroundColor $logColor.error
        Write-Host ""
        Write-Host "Press any key to continue..." -ForegroundColor $logColor.info
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return $false
    }
}

function Show-ConfigMenu {
    Clear-Host
    Write-Host "╔════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║       Configuration Editor              ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Current Configuration File: $configPath" -ForegroundColor $logColor.info
    Write-Host ""
    Write-Host "Edit which setting?" -ForegroundColor $logColor.info
    Write-Host ""
    Write-Host "  1. Root Path" -ForegroundColor White
    Write-Host "  2. Data Source Path" -ForegroundColor White
    Write-Host "  3. Inspect Folder" -ForegroundColor White
    Write-Host "  4. Chart Position & Size" -ForegroundColor White
    Write-Host "  5. Inspection Wait Time" -ForegroundColor White
    Write-Host "  6. Manage Categories" -ForegroundColor White
    Write-Host "  7. Manage File Types" -ForegroundColor White
    Write-Host "  8. Manage Chart Series" -ForegroundColor White
    Write-Host "  0. Back to Main Menu" -ForegroundColor White
    Write-Host ""
}

function Edit-Configuration {
    $continue = $true
    
    while ($continue) {
        Show-ConfigMenu
        $choice = Read-Host "Enter your choice"
        
        try {
            switch ($choice) {
                "1" {
                    $newPath = Read-Host "Enter new root path"
                    Update-ConfigValue -ConfigPath $configPath -Path "paths.rootPath" -Value $newPath
                    Write-Host "Updated: paths.rootPath = $newPath" -ForegroundColor $logColor.success
                }
                "2" {
                    $newPath = Read-Host "Enter new data source path"
                    Update-ConfigValue -ConfigPath $configPath -Path "paths.dataSourcePath" -Value $newPath
                    Write-Host "Updated: paths.dataSourcePath = $newPath" -ForegroundColor $logColor.success
                }
                "3" {
                    $newPath = Read-Host "Enter new inspect folder"
                    Update-ConfigValue -ConfigPath $configPath -Path "paths.inspectFolder" -Value $newPath
                    Write-Host "Updated: paths.inspectFolder = $newPath" -ForegroundColor $logColor.success
                }
                "4" {
                    $left = Read-Host "Enter chart left position"
                    $top = Read-Host "Enter chart top position"
                    Update-ConfigValue -ConfigPath $configPath -Path "chart.position.left" -Value ([int]$left)
                    Update-ConfigValue -ConfigPath $configPath -Path "chart.position.top" -Value ([int]$top)
                    Write-Host "Updated: chart position" -ForegroundColor $logColor.success
                }
                "5" {
                    $waitTime = Read-Host "Enter inspection wait time (seconds)"
                    Update-ConfigValue -ConfigPath $configPath -Path "inspection.waitTime" -Value ([int]$waitTime)
                    Write-Host "Updated: inspection.waitTime = $waitTime" -ForegroundColor $logColor.success
                }
                "6" {
                    Edit-Categories
                }
                "7" {
                    Edit-FileTypes
                }
                "8" {
                    Edit-Series
                }
                "0" {
                    $continue = $false
                }
                default {
                    Write-Host "Invalid choice. Please try again." -ForegroundColor $logColor.warning
                }
            }
        }
        catch {
            Write-Host "Error: $_" -ForegroundColor $logColor.error
        }
        
        if ($choice -notin @("0", "6", "7", "8")) {
            Start-Sleep -Seconds 2
        }
    }
}

function Show-Configuration {
    Clear-Host
    Write-Host "╔════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║       Current Configuration              ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    
    $config | ConvertTo-Json | Write-Host -ForegroundColor White
    
    Write-Host ""
    Write-Host "Press any key to continue..." -ForegroundColor $logColor.info
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Edit-Categories {
    $continue = $true
    
    while ($continue) {
        Clear-Host
        Write-Host "╔════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║       Manage Categories                  ║" -ForegroundColor Cyan
        Write-Host "╚════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host ""
        
        # Reload config to show latest
        $config = Get-ConfigurationSettings -ConfigPath $configPath
        
        Write-Host "Current Categories:" -ForegroundColor $logColor.info
        for ($i = 0; $i -lt $config.categories.Count; $i++) {
            Write-Host "  [$i] $($config.categories[$i])" -ForegroundColor White
        }
        Write-Host ""
        Write-Host "Options:" -ForegroundColor $logColor.info
        Write-Host "  A - Add new category" -ForegroundColor White
        Write-Host "  D - Delete category (by index)" -ForegroundColor White
        Write-Host "  0 - Back to Configuration Menu" -ForegroundColor White
        Write-Host ""
        
        $choice = Read-Host "Enter your choice"
        
        try {
            switch ($choice.ToUpper()) {
                "A" {
                    $newCategory = Read-Host "Enter new category name"
                    if (-not [string]::IsNullOrWhiteSpace($newCategory)) {
                        $newArray = @($config.categories) + $newCategory
                        Update-ArrayInConfig -ConfigPath $configPath -Path "filePatterns.fileTypes" -Array $newArray
                        Write-Host "Added: $newCategory" -ForegroundColor $logColor.success
                        Start-Sleep -Seconds 1
                    }
                }
                "D" {
                    $indexStr = Read-Host "Enter category index to delete"
                    if ([int]::TryParse($indexStr, [ref]$null)) {
                        $index = [int]$indexStr
                        if ($index -ge 0 -and $index -lt $config.categories.Count) {
                            $newArray = @($config.categories | Select-Object -Index 0..($config.categories.Count-1) | Where-Object { $config.categories.IndexOf($_) -ne $index })
                            Update-ArrayInConfig -ConfigPath $configPath -Path "categories" -Array $newArray
                            Write-Host "Deleted category at index $index" -ForegroundColor $logColor.success
                            Start-Sleep -Seconds 1
                        }
                        else {
                            Write-Host "Invalid index" -ForegroundColor $logColor.error
                            Start-Sleep -Seconds 1
                        }
                    }
                }
                "0" {
                    $continue = $false
                }
                default {
                    Write-Host "Invalid choice" -ForegroundColor $logColor.warning
                    Start-Sleep -Seconds 1
                }
            }
        }
        catch {
            Write-Host "Error: $_" -ForegroundColor $logColor.error
            Start-Sleep -Seconds 2
        }
    }
}

function Edit-FileTypes {
    $continue = $true
    
    while ($continue) {
        Clear-Host
        Write-Host "╔════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║       Manage File Types                 ║" -ForegroundColor Cyan
        Write-Host "╚════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host ""
        
        # Reload config to show latest
        $config = Get-ConfigurationSettings -ConfigPath $configPath
        
        Write-Host "Current File Types:" -ForegroundColor $logColor.info
        for ($i = 0; $i -lt $config.filePatterns.fileTypes.Count; $i++) {
            Write-Host "  [$i] $($config.filePatterns.fileTypes[$i])" -ForegroundColor White
        }
        Write-Host ""
        Write-Host "Options:" -ForegroundColor $logColor.info
        Write-Host "  A - Add new file type" -ForegroundColor White
        Write-Host "  D - Delete file type (by index)" -ForegroundColor White
        Write-Host "  0 - Back to Configuration Menu" -ForegroundColor White
        Write-Host ""
        
        $choice = Read-Host "Enter your choice"
        
        try {
            switch ($choice.ToUpper()) {
                "A" {
                    $newType = Read-Host "Enter new file type (e.g., density, pressure)"
                    if (-not [string]::IsNullOrWhiteSpace($newType)) {
                        $newArray = @($config.filePatterns.fileTypes) + $newType
                        Update-ArrayInConfig -ConfigPath $configPath -Path "filePatterns.fileTypes" -Array $newArray
                        Write-Host "Added: $newType" -ForegroundColor $logColor.success
                        Start-Sleep -Seconds 1
                    }
                }
                "D" {
                    $indexStr = Read-Host "Enter file type index to delete"
                    if ([int]::TryParse($indexStr, [ref]$null)) {
                        $index = [int]$indexStr
                        if ($index -ge 0 -and $index -lt $config.filePatterns.fileTypes.Count) {
                            $newArray = @($config.filePatterns.fileTypes | Select-Object -Index 0..($config.filePatterns.fileTypes.Count-1) | Where-Object { $config.filePatterns.fileTypes.IndexOf($_) -ne $index })
                            Update-ArrayInConfig -ConfigPath $configPath -Path "filePatterns.fileTypes" -Array $newArray
                            Write-Host "Deleted file type at index $index" -ForegroundColor $logColor.success
                            Start-Sleep -Seconds 1
                        }
                        else {
                            Write-Host "Invalid index" -ForegroundColor $logColor.error
                            Start-Sleep -Seconds 1
                        }
                    }
                }
                "0" {
                    $continue = $false
                }
                default {
                    Write-Host "Invalid choice" -ForegroundColor $logColor.warning
                    Start-Sleep -Seconds 1
                }
            }
        }
        catch {
            Write-Host "Error: $_" -ForegroundColor $logColor.error
            Start-Sleep -Seconds 2
        }
    }
}

function Edit-Series {
    $continue = $true
    
    while ($continue) {
        Clear-Host
        Write-Host "╔════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║       Manage Chart Series               ║" -ForegroundColor Cyan
        Write-Host "╚════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host ""
        
        # Reload config to show latest
        $config = Get-ConfigurationSettings -ConfigPath $configPath
        
        Write-Host "Current Series:" -ForegroundColor $logColor.info
        for ($i = 0; $i -lt $config.chart.series.Count; $i++) {
            $series = $config.chart.series[$i]
            Write-Host "  [$i] $($series.name) - X:$($series.xColumn) Y:$($series.yColumn)" -ForegroundColor White
        }
        Write-Host ""
        Write-Host "Options:" -ForegroundColor $logColor.info
        Write-Host "  A - Add new series" -ForegroundColor White
        Write-Host "  D - Delete series (by index)" -ForegroundColor White
        Write-Host "  E - Edit series (by index)" -ForegroundColor White
        Write-Host "  0 - Back to Configuration Menu" -ForegroundColor White
        Write-Host ""
        
        $choice = Read-Host "Enter your choice"
        
        try {
            switch ($choice.ToUpper()) {
                "A" {
                    $seriesName = Read-Host "Enter series name"
                    $xCol = (Read-Host "Enter X column letter (A-Z)").ToUpper()
                    $yCol = (Read-Host "Enter Y column letter (A-Z)").ToUpper()
                    
                    if (-not [string]::IsNullOrWhiteSpace($seriesName) -and $xCol.Length -eq 1 -and $yCol.Length -eq 1) {
                        Add-SeriesConfig -ConfigPath $configPath -SeriesName $seriesName -XColumn $xCol -YColumn $yCol
                        Write-Host "Added series: $seriesName" -ForegroundColor $logColor.success
                        Start-Sleep -Seconds 1
                    }
                    else {
                        Write-Host "Invalid input" -ForegroundColor $logColor.error
                        Start-Sleep -Seconds 1
                    }
                }
                "D" {
                    $indexStr = Read-Host "Enter series index to delete"
                    if ([int]::TryParse($indexStr, [ref]$null)) {
                        $index = [int]$indexStr
                        if ($index -ge 0 -and $index -lt $config.chart.series.Count) {
                            Remove-SeriesConfig -ConfigPath $configPath -Index $index
                            Write-Host "Deleted series at index $index" -ForegroundColor $logColor.success
                            Start-Sleep -Seconds 1
                        }
                        else {
                            Write-Host "Invalid index" -ForegroundColor $logColor.error
                            Start-Sleep -Seconds 1
                        }
                    }
                }
                "E" {
                    $indexStr = Read-Host "Enter series index to edit"
                    if ([int]::TryParse($indexStr, [ref]$null)) {
                        $index = [int]$indexStr
                        if ($index -ge 0 -and $index -lt $config.chart.series.Count) {
                            $series = $config.chart.series[$index]
                            Write-Host "Editing: $($series.name)" -ForegroundColor $logColor.info
                            
                            $newName = Read-Host "Enter new name (press Enter to keep '$($series.name)')"
                            $newXCol = (Read-Host "Enter new X column (press Enter to keep '$($series.xColumn)')").ToUpper()
                            $newYCol = (Read-Host "Enter new Y column (press Enter to keep '$($series.yColumn)')").ToUpper()
                            
                            if ([string]::IsNullOrWhiteSpace($newName)) { $newName = $series.name }
                            if ([string]::IsNullOrWhiteSpace($newXCol)) { $newXCol = $series.xColumn }
                            if ([string]::IsNullOrWhiteSpace($newYCol)) { $newYCol = $series.yColumn }
                            
                            $config.chart.series[$index].name = $newName
                            $config.chart.series[$index].xColumn = $newXCol
                            $config.chart.series[$index].yColumn = $newYCol
                            
                            $config | ConvertTo-Json -Depth 10 | Set-Content -Path $configPath
                            Write-Host "Updated series at index $index" -ForegroundColor $logColor.success
                            Start-Sleep -Seconds 1
                        }
                        else {
                            Write-Host "Invalid index" -ForegroundColor $logColor.error
                            Start-Sleep -Seconds 1
                        }
                    }
                }
                "0" {
                    $continue = $false
                }
                default {
                    Write-Host "Invalid choice" -ForegroundColor $logColor.warning
                    Start-Sleep -Seconds 1
                }
            }
        }
        catch {
            Write-Host "Error: $_" -ForegroundColor $logColor.error
            Start-Sleep -Seconds 2
        }
    }
}

# Main loop
$running = $true
while ($running) {
    Show-Menu
    $choice = Read-Host "Enter your choice"
    
    switch ($choice) {
        "1" {
            Invoke-ScriptSafely -ScriptName "Convert-CsvToXlsx.ps1"
        }
        "2" {
            Invoke-ScriptSafely -ScriptName "Create-ComparisonPlots.ps1"
        }
        "3" {
            Invoke-ScriptSafely -ScriptName "Copy-PlotData.ps1"
        }
        "4" {
            $folderPath = Read-Host "Enter folder path to inspect (or press Enter for current folder)"
            if ([string]::IsNullOrWhiteSpace($folderPath)) {
                $folderPath = Get-Location
            }
            Invoke-ScriptSafely -ScriptName "Inspect-Excel.ps1" -Parameters @($configPath, $folderPath)
        }
        "5" {
            Edit-Configuration
        }
        "6" {
            Show-Configuration
        }
        "0" {
            Write-Host "Exiting Excel Automation Software..." -ForegroundColor $logColor.success
            $running = $false
        }
        default {
            Write-Host "Invalid choice. Please try again." -ForegroundColor $logColor.warning
            Start-Sleep -Seconds 2
        }
    }
}
