param(
  [Parameter(Mandatory=$true)][string]$InputPptx,
  [Parameter(Mandatory=$true)][string]$OutputDir
)
$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$ppt = $null
$presentation = $null
try {
  $ppt = New-Object -ComObject PowerPoint.Application
  $ppt.Visible = 1
  $presentation = $ppt.Presentations.Open((Resolve-Path $InputPptx).Path, $true, $true, $false)
  # ppSaveAsPNG = 18. PowerPoint itself is the renderer, avoiding HTML/browser drift.
  $presentation.SaveAs((Resolve-Path $OutputDir).Path, 18)
}
finally {
  if ($presentation) { $presentation.Close() }
  if ($ppt) { $ppt.Quit() }
  if ($presentation) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
  if ($ppt) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
