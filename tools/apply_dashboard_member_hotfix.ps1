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

function Get-CodeBody {
    param([string]$Path)

    $raw = Get-Content -Path $Path -Raw -Encoding UTF8
    $match = [regex]::Match($raw, '(?im)^Option Explicit\b')
    if ($match.Success) {
        return $raw.Substring($match.Index)
    }

    # Fallback: remove export attributes if Option Explicit is missing
    $lines = $raw -split "`r?`n"
    $clean = $lines | Where-Object {
        $_ -notmatch '^Attribute\s+VB_' -and
        $_ -notmatch '^VERSION\s+\d+\.\d+\s+CLASS$' -and
        $_ -notmatch '^BEGIN$' -and
        $_ -notmatch '^\s+MultiUse\s*=.*$' -and
        $_ -notmatch '^END$'
    }
    return ($clean -join "`r`n")
}

$targetWorkbook = Resolve-WorkbookPath -InputPath $WorkbookPath
Write-Host "Applying hotfix modules to workbook:" -ForegroundColor Cyan
Write-Host "  $targetWorkbook" -ForegroundColor Cyan

$updates = @(
    @{ Component = 'mod_Startseite'; File = Join-Path $repoRoot 'vba/Modules/mod_Startseite.bas' },
    @{ Component = 'mod_Uebersicht_Daten'; File = Join-Path $repoRoot 'vba/Modules/mod_Uebersicht_Daten.bas' },
    @{ Component = 'mod_Uebersicht_Dashboard'; File = Join-Path $repoRoot 'vba/Modules/mod_Uebersicht_Dashboard.bas' },
    @{ Component = 'DieseArbeitsmappe'; File = Join-Path $repoRoot 'vba/Classes/DieseArbeitsmappe.cls' }
)

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Open($targetWorkbook)
    $vbProj = $workbook.VBProject

    foreach ($u in $updates) {
        if (-not (Test-Path $u.File)) {
            throw "Missing source file: $($u.File)"
        }

        $codeBody = Get-CodeBody -Path $u.File
        $comp = $vbProj.VBComponents.Item($u.Component)
        if ($null -eq $comp) {
            throw "VBA component not found in workbook: $($u.Component)"
        }

        $cm = $comp.CodeModule
        if ($cm.CountOfLines -gt 0) {
            $cm.DeleteLines(1, $cm.CountOfLines)
        }
        if ($codeBody.Trim().Length -gt 0) {
            [void]$cm.AddFromString($codeBody)
        }

        Write-Host "  Updated $($u.Component)" -ForegroundColor Green
    }

    $excel.EnableEvents = $true

    # Refresh key views and values with new code
    $excel.Run('mod_Startseite.InitialisiereStartseite')
    $excel.Run('mod_Startseite.AktualisiereParzellenAnzeigen')
    $excel.Run('mod_Uebersicht_Dashboard.GeneriereUebersichtNeu', $true)

    $workbook.Save()
    Write-Host "Workbook saved successfully." -ForegroundColor Green
}
finally {
    if ($workbook -ne $null) {
        try { $workbook.Close($true) } catch {}
    }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
    }

    if ($workbook -ne $null) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($excel -ne $null) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Hotfix applied." -ForegroundColor Green
