param(
    [string]$WorkbookPath
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

Write-Host "Repair + Verify:" -ForegroundColor Cyan
Write-Host "  $targetWorkbook" -ForegroundColor Cyan

$excel = $null
$wb = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true

    $wb = $excel.Workbooks.Open($targetWorkbook)

    # Apply latest module hotfixes already in workbook (no import here), then force refresh
    $excel.Run('mod_Startseite.InitialisiereStartseite')
    $excel.Run('mod_Startseite.AktualisiereParzellenAnzeigen')
    $excel.Run('mod_Uebersicht_Dashboard.SynchronisiereDashboardKpiSofort')

    try { $excel.Run('mod_Uebersicht_Dashboard.GeneriereUebersichtNeu', $true) } catch {}

    $wsML = $wb.Worksheets.Item('Mitgliederliste')
    $wsStart = $wb.Worksheets.Item('Startmenü')
    $wsDash = $wb.Worksheets.Item('Dashboard Mitgliederzahlungen')

    $lastRow = $wsML.Cells($wsML.Rows.Count, 5).End(-4162).Row  # xlUp
    $countReal = 0

    for ($r = 6; $r -le $lastRow; $r++) {
        $vn = [string]$wsML.Cells($r, 6).Value2
        $nn = [string]$wsML.Cells($r, 5).Value2
        if ([string]::IsNullOrWhiteSpace($vn) -and [string]::IsNullOrWhiteSpace($nn)) { continue }

        $anrede = [string]$wsML.Cells($r, 4).Value2
        if ($anrede -and $anrede.Trim().ToUpper() -eq 'KGA') { continue }

        $funktion = [string]$wsML.Cells($r, 15).Value2
        if ($funktion -and $funktion.Trim().ToUpper() -eq 'EHEMALIGES MITGLIED') { continue }

        $countReal++
    }

    $startMembers = [string]$wsStart.Range('G8').Text
    $dashMembersText = [string]$wsDash.Cells(7, 1).Text

    $containsAktiv = $dashMembersText.ToLower().Contains('aktiv')

    Write-Host ("COUNT_REAL={0}" -f $countReal) -ForegroundColor Yellow
    Write-Host ("START_G8_TEXT={0}" -f $startMembers) -ForegroundColor Yellow
    Write-Host ("DASH_A7_TEXT={0}" -f $dashMembersText) -ForegroundColor Yellow
    Write-Host ("DASH_CONTAINS_AKTIV={0}" -f $containsAktiv) -ForegroundColor Yellow

    $wb.Save()
}
finally {
    if ($wb -ne $null) {
        try { $wb.Close($true) } catch {}
    }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
    }

    if ($wb -ne $null) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
    if ($excel -ne $null) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Done." -ForegroundColor Green
