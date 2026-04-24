<#
.SYNOPSIS
Chart creation and management utilities
#>

function New-ExcelChart {
    <#
    .SYNOPSIS
    Create a new chart in a worksheet
    
    .PARAMETER Sheet
    Worksheet object
    
    .PARAMETER Left
    Left position of the chart
    
    .PARAMETER Top
    Top position of the chart
    
    .PARAMETER Width
    Width of the chart
    
    .PARAMETER Height
    Height of the chart
    
    .PARAMETER ChartType
    Excel chart type constant
    
    .PARAMETER ChartStyle
    Chart style ID
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Sheet,
        
        [int]$Left = 350,
        [int]$Top = 20,
        [int]$Width = 360,
        [int]$Height = 216,
        [int]$ChartType = 75,
        [int]$ChartStyle = 240
    )
    
    try {
        $chartObj = $Sheet.ChartObjects().Add($Left, $Top, $Width, $Height)
        $chart = $chartObj.Chart
        $chart.ChartType = $ChartType
        $chart.ChartStyle = $ChartStyle
        
        Write-Verbose "Chart created at position ($Left, $Top)"
        return $chart
    }
    catch {
        throw "Failed to create chart: $_"
    }
}

function Add-ChartSeries {
    <#
    .SYNOPSIS
    Add a data series to a chart
    
    .PARAMETER Chart
    Chart object
    
    .PARAMETER Sheet
    Worksheet object
    
    .PARAMETER SeriesName
    Name of the series
    
    .PARAMETER XValuesRange
    Range for X values
    
    .PARAMETER YValuesRange
    Range for Y values
    
    .PARAMETER SeriesIndex
    Index of the series (for removal purposes)
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Chart,
        
        [Parameter(Mandatory=$true)]
        [Object]$Sheet,
        
        [Parameter(Mandatory=$true)]
        [string]$SeriesName,
        
        [Parameter(Mandatory=$true)]
        [string]$XValuesRange,
        
        [Parameter(Mandatory=$true)]
        [string]$YValuesRange
    )
    
    try {
        $series = $Chart.SeriesCollection().NewSeries()
        $series.Name = $SeriesName
        $series.XValues = $Sheet.Range($XValuesRange)
        $series.Values = $Sheet.Range($YValuesRange)
        
        Write-Verbose "Series '$SeriesName' added to chart"
        return $series
    }
    catch {
        throw "Failed to add series: $_"
    }
}

function Clear-ChartSeries {
    <#
    .SYNOPSIS
    Remove all series from a chart
    
    .PARAMETER Chart
    Chart object
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Chart
    )
    
    try {
        while ($Chart.SeriesCollection().Count -gt 0) {
            $Chart.SeriesCollection(1).Delete()
        }
        Write-Verbose "All chart series cleared"
    }
    catch {
        throw "Failed to clear series: $_"
    }
}

function Set-ChartTitle {
    <#
    .SYNOPSIS
    Set the chart title
    
    .PARAMETER Chart
    Chart object
    
    .PARAMETER Title
    Title text
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Chart,
        
        [Parameter(Mandatory=$true)]
        [string]$Title
    )
    
    try {
        $Chart.HasTitle = $true
        $Chart.ChartTitle.Text = $Title
        Write-Verbose "Chart title set to: $Title"
    }
    catch {
        throw "Failed to set chart title: $_"
    }
}

function Set-ChartAxes {
    <#
    .SYNOPSIS
    Configure chart axes
    
    .PARAMETER Chart
    Chart object
    
    .PARAMETER XAxisTitle
    X-axis title
    
    .PARAMETER YAxisTitle
    Y-axis title
    
    .PARAMETER HasGridlines
    Show gridlines
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Chart,
        
        [string]$XAxisTitle = "X-Axis",
        [string]$YAxisTitle = "Y-Axis",
        [bool]$HasGridlines = $true
    )
    
    try {
        # X-Axis (Axes constant 1)
        $xAxis = $Chart.Axes(1)
        $xAxis.HasMajorGridlines = $HasGridlines
        $xAxis.HasTitle = $true
        $xAxis.AxisTitle.Text = $XAxisTitle
        
        # Y-Axis (Axes constant 2)
        $yAxis = $Chart.Axes(2)
        $yAxis.HasMajorGridlines = $HasGridlines
        $yAxis.HasTitle = $true
        $yAxis.AxisTitle.Text = $YAxisTitle
        
        Write-Verbose "Chart axes configured: X='$XAxisTitle', Y='$YAxisTitle'"
    }
    catch {
        throw "Failed to configure axes: $_"
    }
}

function Set-ChartLegend {
    <#
    .SYNOPSIS
    Configure chart legend
    
    .PARAMETER Chart
    Chart object
    
    .PARAMETER ShowLegend
    Show or hide legend
    
    .PARAMETER Position
    Legend position constant
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Chart,
        
        [bool]$ShowLegend = $true,
        [int]$Position = -4152
    )
    
    try {
        $Chart.HasLegend = $ShowLegend
        if ($ShowLegend) {
            $Chart.Legend.Position = $Position
        }
        Write-Verbose "Chart legend configured"
    }
    catch {
        throw "Failed to configure legend: $_"
    }
}

function Build-ChartFromConfig {
    <#
    .SYNOPSIS
    Build a complete chart using configuration object
    
    .PARAMETER Sheet
    Worksheet object
    
    .PARAMETER Config
    Configuration object containing chart settings
    
    .PARAMETER Title
    Chart title
    
    .PARAMETER LastRow
    Last row with data
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Sheet,
        
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$Config,
        
        [Parameter(Mandatory=$true)]
        [string]$Title,
        
        [Parameter(Mandatory=$true)]
        [int]$LastRow
    )
    
    try {
        # Create chart
        $chart = New-ExcelChart -Sheet $Sheet `
            -Left $Config.position.left `
            -Top $Config.position.top `
            -Width $Config.position.width `
            -Height $Config.position.height `
            -ChartType $Config.type `
            -ChartStyle $Config.style
        
        # Clear any existing series
        Clear-ChartSeries -Chart $chart
        
        # Add series from config
        foreach ($seriesConfig in $Config.series) {
            $xRange = "$($seriesConfig.xColumn)2:$($seriesConfig.xColumn)$LastRow"
            $yRange = "$($seriesConfig.yColumn)2:$($seriesConfig.yColumn)$LastRow"
            Add-ChartSeries -Chart $chart -Sheet $Sheet `
                -SeriesName $seriesConfig.name `
                -XValuesRange $xRange `
                -YValuesRange $yRange
        }
        
        # Configure axes
        Set-ChartAxes -Chart $chart `
            -XAxisTitle $Config.axes.xTitle `
            -YAxisTitle $Config.axes.yTitle
        
        # Configure legend
        Set-ChartLegend -Chart $chart `
            -ShowLegend $Config.hasLegend `
            -Position $Config.legendPosition
        
        # Set title
        Set-ChartTitle -Chart $chart -Title $Title
        
        Write-Verbose "Chart configured successfully"
        return $chart
    }
    catch {
        throw "Failed to build chart: $_"
    }
}

Export-ModuleMember -Function New-ExcelChart, Add-ChartSeries, Clear-ChartSeries, `
    Set-ChartTitle, Set-ChartAxes, Set-ChartLegend, Build-ChartFromConfig
