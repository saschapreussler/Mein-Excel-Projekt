param(
    [string]$WorkbookPath,
    [switch]$Visible
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$excelDir = Join-Path $repoRoot 'excel'

function Resolve-WorkbookPath {
    param([string]$InputPath)

    if ($InputPath) {
        if (-not (Test-Path $InputPath)) {
            throw "Workbook not found: $InputPath"
        }
        return (Resolve-Path $InputPath).Path
    }

    if (-not (Test-Path $excelDir)) {
        throw "Excel folder not found: $excelDir"
    }

    $candidates = Get-ChildItem -Path $excelDir -File -Filter '*.xlsm' |
        Where-Object { $_.Name -notlike '~$*' } |
        Sort-Object LastWriteTime -Descending

    if (-not $candidates -or $candidates.Count -eq 0) {
        throw "No .xlsm workbook found in $excelDir"
    }

    return $candidates[0].FullName
}

function Test-WorkbookLocked {
    param([string]$Path)

    try {
        $fs = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
        $fs.Close()
        return $false
    }
    catch {
        return $true
    }
}

$targetWorkbook = Resolve-WorkbookPath -InputPath $WorkbookPath
if (Test-WorkbookLocked -Path $targetWorkbook) {
    throw "Workbook is currently open/locked. Please close it first: $targetWorkbook"
}
Write-Host "Applying repo VBA to workbook:" -ForegroundColor Cyan
Write-Host "  $targetWorkbook" -ForegroundColor Cyan

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true

    $workbook = $excel.Workbooks.Open($targetWorkbook)

    Write-Host "Running VBA reimport macro..." -ForegroundColor Cyan
    $excel.Run('mod_Repo_Sync.SyncVBAVomRepository')

    Write-Host "Refreshing Start menu KPIs..." -ForegroundColor Cyan
    $excel.Run('mod_Startseite.InitialisiereStartseite')
    $excel.Run('mod_Startseite.AktualisiereParzellenAnzeigen')

    Write-Host "Refreshing dashboard..." -ForegroundColor Cyan
    try {
        $excel.Run('mod_Uebersicht_Dashboard.GeneriereUebersichtNeu', $true)
    }
    catch {
        Write-Warning "Dashboard refresh macro raised an error: $($_.Exception.Message)"
    }

    $workbook.Save()
    Write-Host "Workbook saved successfully." -ForegroundColor Green
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($true) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }

    if ($null -ne $workbook) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($null -ne $excel) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Done." -ForegroundColor Green
