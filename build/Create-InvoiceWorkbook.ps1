#requires -Version 5.1
<#
Builds a macro-enabled Excel workbook with three source modules.
Requires Microsoft Excel Desktop for Windows and temporary
"Trust access to the VBA project object model" permission.
Never disables security settings or overwrites output.
#>
[CmdletBinding()]
param([string]$OutputPath = '')
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
if (-not $OutputPath) { $OutputPath = Join-Path $root 'ExcelInvoiceAutomation.xlsm' }
if (Test-Path -LiteralPath $OutputPath) {
    throw "The output file already exists: $OutputPath. Back it up or rename it."
}
$sourceFiles = @(
    (Join-Path $root 'src\modCatalog.bas'),
    (Join-Path $root 'src\modSetup.bas'),
    (Join-Path $root 'src\modQuotations.bas')
)
foreach ($source in $sourceFiles) {
    if (-not (Test-Path -LiteralPath $source)) {
        throw "Missing source file: $source"
    }
}
$excel = $null
$workbook = $null
$success = $false
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    try {
        $components = $workbook.VBProject.VBComponents
        foreach ($source in $sourceFiles) {
            $imported = $components.Import($source)
            Write-Host "Imported $($imported.Name)"
        }
    } catch {
        throw "Cannot import VBA modules. Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings > temporarily allow Trust access to the VBA project object model. Details: $($_.Exception.Message)"
    }
    # 52 = macro-enabled workbook
    $workbook.SaveAs($OutputPath, 52)
    $excel.Run("'" + $workbook.Name.Replace("'", "''") + "'!InitializeInvoiceApp")
    # Delete only unused blank default worksheets.
    for ($i = $workbook.Worksheets.Count; $i -ge 1; $i--) {
        $sheet = $workbook.Worksheets.Item($i)
        if ($sheet.Name -notin @('Dashboard', 'Customers', 'Products', 'Audit Log', 'Quote Draft', 'Quotations', 'Quotation Lines')) {
            if ($sheet.UsedRange.Count -eq 1 -and [string]::IsNullOrWhiteSpace([string]$sheet.Range('A1').Value2)) {
                $sheet.Delete()
            }
        }
    }
    $workbook.Save()
    $success = $true
    Write-Host "CREATED: $OutputPath"
    Write-Host "Open in Excel, run Debug > Compile VBAProject, and test Add customer/product."
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($success) } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
