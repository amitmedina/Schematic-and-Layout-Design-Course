param(
  [Parameter(Mandatory=$true)]
  [string]$JsonPath,

  [Parameter(Mandatory=$false)]
  [string]$TemplatePath = "",

  [Parameter(Mandatory=$false)]
  [string]$OutXlsm = "",

  [Parameter(Mandatory=$false)]
  [string]$OutXlsx = "",

  [switch]$Visible
)

$ErrorActionPreference = 'Stop'

function Ensure-ParentDir([string]$Path) {
  $parent = Split-Path -Parent $Path
  if ($parent -and -not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent | Out-Null
  }
}

if (-not (Test-Path -LiteralPath $JsonPath)) {
  throw "JSON not found: $JsonPath"
}

# The file starting with "~$" is an Excel temp/lock file.
# Do not use it as input; use the real .xlsm template.

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptDir

if (-not $TemplatePath) {
  $TemplatePath = Join-Path $repoRoot "training\LM5148_LM25148_quickstart_calculator_A4.xlsm"
}
if (-not (Test-Path -LiteralPath $TemplatePath)) {
  throw "Template not found: $TemplatePath"
}

if (-not $OutXlsm) {
  $OutXlsm = Join-Path $scriptDir "LM5148_quickstart_filled_excel.xlsm"
}
if (-not $OutXlsx) {
  $OutXlsx = Join-Path $scriptDir "LM5148_quickstart_filled_excel.xlsx"
}

$payload = Get-Content -LiteralPath $JsonPath -Raw | ConvertFrom-Json
if (-not $payload.inputs) {
  throw "JSON must contain an 'inputs' object"
}

$inputs = $payload.inputs

# Inputs expected from the webapp JSON export
$vinNom = [double]$inputs.vinNom
$vinMax = [double]$inputs.vinMax
$vinMin = if ($inputs.PSObject.Properties.Name -contains 'vinMin' -and $inputs.vinMin) { [double]$inputs.vinMin } else { $vinNom }

$vout  = [double]$inputs.vout
$iout  = [double]$inputs.iout
$fswHz = [double]$inputs.fsw
$fswKHz = $fswHz / 1000.0

Ensure-ParentDir $OutXlsm
if ($OutXlsx) { Ensure-ParentDir $OutXlsx }

$excel = $null
$workbook = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = [bool]$Visible
  $excel.DisplayAlerts = $false
  $excel.ScreenUpdating = $false

  # Open the macro-enabled calculator (preserves shapes/macros)
  $workbook = $excel.Workbooks.Open($TemplatePath, 0, $false)

  $sheet = $workbook.Worksheets.Item('Design Regulator')

  # Known input cells (observed in the template)
  $sheet.Range('E6').Value2  = $vinMin
  $sheet.Range('E7').Value2  = $vinNom
  $sheet.Range('E8').Value2  = $vinMax
  $sheet.Range('E9').Value2  = $vout
  $sheet.Range('E10').Value2 = $iout
  $sheet.Range('E11').Value2 = $fswKHz

  # Force a full recalc so dependent sheets update
  $excel.CalculateFullRebuild()

  # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
  $workbook.SaveAs($OutXlsm, 52)

  if ($OutXlsx) {
    # 51 = xlOpenXMLWorkbook (.xlsx)
    $workbook.SaveAs($OutXlsx, 51)
  }

  Write-Host "Wrote: $OutXlsm"
  if ($OutXlsx) { Write-Host "Wrote: $OutXlsx" }
}
finally {
  if ($workbook) {
    try { $workbook.Close($false) } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
  }
  if ($excel) {
    try { $excel.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
