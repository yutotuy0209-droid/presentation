param(
  [Parameter(Mandatory=$true)][string]$Pptx,
  [int]$Iterations = 5,
  [switch]$Visual
)
$ErrorActionPreference="Stop"
$Root=Split-Path -Parent $PSScriptRoot
$ArgsList=@("$Root\src\roundtrip_loop.py",$Pptx,"--work","$Root\output\loop","--iterations","$Iterations")
if($Visual){$ArgsList += "--visual"}
python @ArgsList
