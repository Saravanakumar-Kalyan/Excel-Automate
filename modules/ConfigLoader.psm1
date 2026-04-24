<#
.SYNOPSIS
Configuration management utilities
#>

function Get-ConfigurationSettings {
    <#
    .SYNOPSIS
    Load configuration from JSON file
    
    .PARAMETER ConfigPath
    Path to the configuration JSON file
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigPath
    )
    
    try {
        if (-not (Test-Path $ConfigPath)) {
            throw "Configuration file not found at: $ConfigPath"
        }
        
        $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        Write-Verbose "Configuration loaded from: $ConfigPath"
        return $config
    }
    catch {
        throw "Failed to load configuration: $_"
    }
}

function Update-ConfigValue {
    <#
    .SYNOPSIS
    Update a configuration value
    
    .PARAMETER ConfigPath
    Path to the configuration JSON file
    
    .PARAMETER Path
    Property path (dot notation, e.g., "paths.rootPath")
    
    .PARAMETER Value
    New value
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [Parameter(Mandatory=$true)]
        $Value
    )
    
    try {
        $config = Get-ConfigurationSettings -ConfigPath $ConfigPath
        
        # Navigate to the property and update it
        $parts = $Path -split '\.'
        $current = $config
        
        for ($i = 0; $i -lt $parts.Count - 1; $i++) {
            $current = $current.($parts[$i])
        }
        
        $current.($parts[-1]) = $Value
        
        # Save back to file
        $config | ConvertTo-Json -Depth 10 | Set-Content -Path $ConfigPath
        Write-Verbose "Configuration updated: $Path = $Value"
    }
    catch {
        throw "Failed to update configuration: $_"
    }
}

function Get-ConfigValue {
    <#
    .SYNOPSIS
    Get a specific configuration value
    
    .PARAMETER ConfigPath
    Path to the configuration JSON file
    
    .PARAMETER Path
    Property path (dot notation)
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    try {
        $config = Get-ConfigurationSettings -ConfigPath $ConfigPath
        
        $parts = $Path -split '\.'
        $current = $config
        
        foreach ($part in $parts) {
            $current = $current.$part
        }
        
        return $current
    }
    catch {
        throw "Failed to get configuration value: $_"
    }
}

function Validate-Configuration {
    <#
    .SYNOPSIS
    Validate that all required configuration values exist
    
    .PARAMETER Config
    Configuration object
    #>
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$Config
    )
    
    try {
        $requiredPaths = @(
            'paths.rootPath',
            'filePatterns.csvPattern',
            'filePatterns.xlsxPattern',
            'categories',
            'chart.type',
            'chart.style',
            'excel.visible',
            'excel.displayAlerts',
            'excel.fileFormat'
        )
        
        $missing = @()
        foreach ($path in $requiredPaths) {
            $parts = $path -split '\.'
            $current = $Config
            $found = $true
            
            foreach ($part in $parts) {
                if ($current.PSObject.Properties.Name -contains $part) {
                    $current = $current.$part
                }
                else {
                    $found = $false
                    break
                }
            }
            
            if (-not $found) {
                $missing += $path
            }
        }
        
        if ($missing.Count -gt 0) {
            throw "Missing required configuration: $($missing -join ', ')"
        }
        
        Write-Verbose "Configuration validation passed"
        return $true
    }
    catch {
        throw "Configuration validation failed: $_"
    }
}

function Get-ArrayFromConfig {
    <#
    .SYNOPSIS
    Get an array from configuration by path
    
    .PARAMETER ConfigPath
    Path to the configuration JSON file
    
    .PARAMETER Path
    Property path (dot notation)
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    try {
        $config = Get-ConfigurationSettings -ConfigPath $ConfigPath
        $parts = $Path -split '\.'
        $current = $config
        
        foreach ($part in $parts) {
            $current = $current.$part
        }
        
        return @($current)
    }
    catch {
        throw "Failed to get array from configuration: $_"
    }
}

function Update-ArrayInConfig {
    <#
    .SYNOPSIS
    Update an array in configuration
    
    .PARAMETER ConfigPath
    Path to the configuration JSON file
    
    .PARAMETER Path
    Property path (dot notation)
    
    .PARAMETER Array
    New array values
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [Parameter(Mandatory=$true)]
        [array]$Array
    )
    
    try {
        $config = Get-ConfigurationSettings -ConfigPath $ConfigPath
        $parts = $Path -split '\.'
        $current = $config
        
        for ($i = 0; $i -lt $parts.Count - 1; $i++) {
            $current = $current.($parts[$i])
        }
        
        $current.($parts[-1]) = @($Array)
        
        $config | ConvertTo-Json -Depth 10 | Set-Content -Path $ConfigPath
        Write-Verbose "Array updated: $Path"
    }
    catch {
        throw "Failed to update array in configuration: $_"
    }
}

function Add-SeriesConfig {
    <#
    .SYNOPSIS
    Add a new series to chart configuration
    
    .PARAMETER ConfigPath
    Path to the configuration JSON file
    
    .PARAMETER SeriesName
    Name of the series
    
    .PARAMETER XColumn
    X-axis column letter
    
    .PARAMETER YColumn
    Y-axis column letter
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory=$true)]
        [string]$SeriesName,
        
        [Parameter(Mandatory=$true)]
        [string]$XColumn,
        
        [Parameter(Mandatory=$true)]
        [string]$YColumn
    )
    
    try {
        $config = Get-ConfigurationSettings -ConfigPath $ConfigPath
        
        $newSeries = [PSCustomObject]@{
            name    = $SeriesName
            xColumn = $XColumn
            yColumn = $YColumn
        }
        
        $config.chart.series += $newSeries
        $config | ConvertTo-Json -Depth 10 | Set-Content -Path $ConfigPath
        Write-Verbose "Series added: $SeriesName"
    }
    catch {
        throw "Failed to add series: $_"
    }
}

function Remove-SeriesConfig {
    <#
    .SYNOPSIS
    Remove a series from chart configuration by index
    
    .PARAMETER ConfigPath
    Path to the configuration JSON file
    
    .PARAMETER Index
    Index of series to remove (0-based)
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory=$true)]
        [int]$Index
    )
    
    try {
        $config = Get-ConfigurationSettings -ConfigPath $ConfigPath
        
        if ($Index -lt 0 -or $Index -ge $config.chart.series.Count) {
            throw "Invalid series index: $Index"
        }
        
        $config.chart.series = @($config.chart.series | Where-Object { $_ } | Select-Object -Index 0..($config.chart.series.Count-2) | Where-Object { $config.chart.series.IndexOf($_) -ne $Index } )
        
        # Alternative simpler approach
        $newSeries = @()
        for ($i = 0; $i -lt $config.chart.series.Count; $i++) {
            if ($i -ne $Index) {
                $newSeries += $config.chart.series[$i]
            }
        }
        
        $config.chart.series = $newSeries
        $config | ConvertTo-Json -Depth 10 | Set-Content -Path $configPath
        Write-Verbose "Series removed at index: $Index"
    }
    catch {
        throw "Failed to remove series: $_"
    }
}

Export-ModuleMember -Function Get-ConfigurationSettings, Update-ConfigValue, `
    Get-ConfigValue, Validate-Configuration, Get-ArrayFromConfig, Update-ArrayInConfig, `
    Add-SeriesConfig, Remove-SeriesConfig
