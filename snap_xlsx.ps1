param([string]$Path=".\Stock_Daily_9-10.xlsx",[string]$Sheet="Summary",[string]$Range="A1:AI42",[string]$Out=".\_slides\summary.png")
$ErrorActionPreference="Stop"
$full=(Resolve-Path $Path).Path
$out=Join-Path (Get-Location) $Out
$xl=New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false
try{
  $wb=$xl.Workbooks.Open($full)
  $ws=$wb.Sheets.Item($Sheet)
  $rng=$ws.Range($Range)
  $rng.CopyPicture(1,2)   # xlScreen, xlBitmap
  $co=$ws.ChartObjects().Add(0,0,$rng.Width,$rng.Height)
  $co.Chart.Paste()
  $co.Chart.Export($out,"PNG")
  $co.Delete()
  $wb.Close($false)
  "snap -> $out"
}finally{
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
