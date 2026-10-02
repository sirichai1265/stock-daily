param([Parameter(Mandatory=$true)][string]$Path)
$ErrorActionPreference = "Stop"
$full = (Resolve-Path $Path).Path
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($full)
    $xl.CalculateFullRebuild()
    $wb.Save()
    $wb.Close($true)
    "recalc OK: $full"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
