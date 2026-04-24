<#
.SYNOPSIS
Excel utility functions for automation
#>

function Initialize-ExcelApplication {
    <#
    .SYNOPSIS
    Initialize Excel COM object with specified settings
    
    .PARAMETER Visible
    Set Excel visibility
    
    .PARAMETER DisplayAlerts
    Show or hide alerts
    #>
    param(
        [bool]$Visible = $false,
        [bool]$DisplayAlerts = $false
    )
    
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $Visible
        $excel.DisplayAlerts = $DisplayAlerts
        Write-Verbose "Excel application initialized"
        return $excel
    }
    catch {
        throw "Failed to initialize Excel: $_"
    }
}

function Close-ExcelApplication {
    <#
    .SYNOPSIS
    Properly close Excel and release COM objects
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$ExcelApp
    )
    
    try {
        $ExcelApp.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelApp) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Write-Verbose "Excel application closed and resources released"
    }
    catch {
        Write-Warning "Error closing Excel: $_"
    }
}

function Open-ExcelWorkbook {
    <#
    .SYNOPSIS
    Open an Excel workbook
    
    .PARAMETER FilePath
    Path to the Excel file
    
    .PARAMETER ExcelApp
    Excel application object
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$true)]
        [Object]$ExcelApp
    )
    
    try {
        $workbook = $ExcelApp.Workbooks.Open($FilePath)
        return $workbook
    }
    catch {
        throw "Failed to open workbook at $FilePath : $_"
    }
}

function Close-ExcelWorkbook {
    <#
    .SYNOPSIS
    Close an Excel workbook
    
    .PARAMETER Workbook
    Workbook object to close
    
    .PARAMETER SaveChanges
    Whether to save changes
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Workbook,
        
        [bool]$SaveChanges = $false
    )
    
    try {
        $Workbook.Close($SaveChanges)
    }
    catch {
        Write-Warning "Error closing workbook: $_"
    }
}

function Save-ExcelWorkbook {
    <#
    .SYNOPSIS
    Save an Excel workbook with specified format
    
    .PARAMETER Workbook
    Workbook object to save
    
    .PARAMETER Path
    Output file path
    
    .PARAMETER FileFormat
    Excel file format (51 = .xlsx)
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Workbook,
        
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [int]$FileFormat = 51
    )
    
    try {
        $Workbook.SaveAs($Path, $FileFormat)
        Write-Verbose "Workbook saved to $Path"
    }
    catch {
        throw "Failed to save workbook: $_"
    }
}

function Get-LastRowWithData {
    <#
    .SYNOPSIS
    Find the last row with data in a column
    
    .PARAMETER Sheet
    Worksheet object
    
    .PARAMETER Column
    Column letter (default 'A')
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Sheet,
        
        [string]$Column = "A"
    )
    
    return $Sheet.Cells($Sheet.Rows.Count, $Column).End(-4162).Row
}

function Get-ActiveSheet {
    <#
    .SYNOPSIS
    Get the active sheet from a workbook
    
    .PARAMETER Workbook
    Workbook object
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Workbook
    )
    
    $sheet = $Workbook.ActiveSheet
    if ($null -eq $sheet) {
        throw "Could not access ActiveSheet"
    }
    return $sheet
}

function Delete-Rows {
    <#
    .SYNOPSIS
    Delete specified rows from worksheet
    
    .PARAMETER Sheet
    Worksheet object
    
    .PARAMETER RowNumbers
    Array of row numbers to delete
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Sheet,
        
        [Parameter(Mandatory=$true)]
        [int[]]$RowNumbers
    )
    
    try {
        # Sort in descending order to delete from bottom to top
        $sortedRows = $RowNumbers | Sort-Object -Descending
        
        foreach ($row in $sortedRows) {
            $null = $Sheet.Rows.Item($row).Delete()
        }
        Write-Verbose "Deleted rows: $($RowNumbers -join ', ')"
    }
    catch {
        throw "Failed to delete rows: $_"
    }
}

function Set-ColumnFormat {
    <#
    .SYNOPSIS
    Set number format for columns
    
    .PARAMETER Sheet
    Worksheet object
    
    .PARAMETER Columns
    Array of column letters
    
    .PARAMETER Format
    Number format string
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Sheet,
        
        [Parameter(Mandatory=$true)]
        [string[]]$Columns,
        
        [string]$Format = "General"
    )
    
    foreach ($col in $Columns) {
        $Sheet.Columns.Item($col).NumberFormat = $Format
    }
}

function Remove-ExistingCharts {
    <#
    .SYNOPSIS
    Remove all existing charts from a worksheet
    
    .PARAMETER Sheet
    Worksheet object
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Sheet
    )
    
    try {
        $charts = $Sheet.ChartObjects()
        while ($charts.Count -gt 0) {
            $charts.Item(1).Delete()
        }
        Write-Verbose "Existing charts removed"
    }
    catch {
        Write-Warning "Error removing charts: $_"
    }
}

function Copy-Range {
    <#
    .SYNOPSIS
    Copy a range and paste it to a destination
    
    .PARAMETER Sheet
    Worksheet object
    
    .PARAMETER SourceRange
    Source range string (e.g., "A1:A10")
    
    .PARAMETER DestinationRange
    Destination range string or object
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Object]$Sheet,
        
        [Parameter(Mandatory=$true)]
        [string]$SourceRange,
        
        [Parameter(Mandatory=$true)]
        $DestinationRange
    )
    
    try {
        $Sheet.Range($SourceRange).Copy() | Out-Null
        $Sheet.Paste($DestinationRange) | Out-Null
        Write-Verbose "Copied range $SourceRange to destination"
    }
    catch {
        throw "Failed to copy/paste range: $_"
    }
}

Export-ModuleMember -Function Initialize-ExcelApplication, Close-ExcelApplication, `
    Open-ExcelWorkbook, Close-ExcelWorkbook, Save-ExcelWorkbook, Get-LastRowWithData, `
    Get-ActiveSheet, Delete-Rows, Set-ColumnFormat, Remove-ExistingCharts, Copy-Range
