# ============================================================================
# convert_repo_to_utf8bom.ps1
# ----------------------------------------------------------------------------
# Konvertiert alle .bas/.cls Dateien im VBA-Repo von Windows-1252 (ANSI)
# zu UTF-8 mit BOM. .frm-Dateien werden absichtlich UEBERSPRUNGEN (die
# muessen im VBA-Import-Format bleiben, also ANSI ohne BOM).
#
# Erkennung pro Datei:
#   - hat schon UTF-8-BOM  -> unveraendert
#   - valides UTF-8 ohne BOM -> nur BOM hinzu
#   - sonst (ANSI) -> Windows-1252 lesen, UTF-8+BOM schreiben
#
# Verwendung:
#   powershell -ExecutionPolicy Bypass -File tools\convert_repo_to_utf8bom.ps1
#   powershell -ExecutionPolicy Bypass -File tools\convert_repo_to_utf8bom.ps1 -DryRun
# ============================================================================
param(
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)

function Get-Encoding {
    param([byte[]]$b)
    if ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) {
        return "UTF8_BOM"
    }
    # Try strict UTF-8 decode: if any invalid sequence -> not UTF-8
    try {
        $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)  # (BOM=false, throwOnInvalid=true)
        [void]$utf8Strict.GetString($b)
        # No exception -> valid UTF-8 (or pure ASCII)
        # Check whether file has any non-ASCII byte to decide
        $hasHigh = $false
        foreach ($by in $b) { if ($by -ge 0x80) { $hasHigh = $true; break } }
        if ($hasHigh) { return "UTF8_NoBom" } else { return "ASCII" }
    } catch {
        return "ANSI"  # Windows-1252 fallback
    }
}

$targets = @()
$targets += Get-ChildItem "vba\Modules" -Filter *.bas -File
$targets += Get-ChildItem "vba\Classes" -Filter *.cls -File

Write-Host "`n=== Repo -> UTF-8+BOM Konvertierung ===" -ForegroundColor Cyan
$modus = if ($DryRun) { 'DRY-RUN' } else { 'LIVE' }
Write-Host ("Modus: {0}  |  Ziele: {1}" -f $modus, $targets.Count)
Write-Host ""

$statConv=0; $statAdd=0; $statSkip=0
foreach ($f in $targets) {
    $b = [System.IO.File]::ReadAllBytes($f.FullName)
    $enc = Get-Encoding $b
    switch ($enc) {
        "UTF8_BOM"   { $statSkip++; continue }
        "ASCII"      {
            # Reinen ASCII-Dateien BOM hinzufuegen? Optional. Der Konsistenz halber ja.
            if (-not $DryRun) {
                $txt = [System.Text.Encoding]::ASCII.GetString($b)
                $utf8Bom = New-Object System.Text.UTF8Encoding($true)
                [System.IO.File]::WriteAllText($f.FullName, $txt, $utf8Bom)
            }
            $statAdd++
            Write-Host ("  BOM-add  ASCII    -> {0}" -f $f.Name) -ForegroundColor DarkCyan
        }
        "UTF8_NoBom" {
            if (-not $DryRun) {
                $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
                $txt = $utf8.GetString($b)
                $utf8Bom = New-Object System.Text.UTF8Encoding($true)
                [System.IO.File]::WriteAllText($f.FullName, $txt, $utf8Bom)
            }
            $statAdd++
            Write-Host ("  BOM-add  UTF-8    -> {0}" -f $f.Name) -ForegroundColor Green
        }
        "ANSI" {
            if (-not $DryRun) {
                $win1252 = [System.Text.Encoding]::GetEncoding(1252)
                $txt = $win1252.GetString($b)
                $utf8Bom = New-Object System.Text.UTF8Encoding($true)
                [System.IO.File]::WriteAllText($f.FullName, $txt, $utf8Bom)
            }
            $statConv++
            Write-Host ("  CONVERT  ANSI     -> {0}" -f $f.Name) -ForegroundColor Yellow
        }
    }
}

Write-Host "`n=== Zusammenfassung ===" -ForegroundColor Cyan
Write-Host ("ANSI->UTF8+BOM konvertiert: {0}" -f $statConv)
Write-Host ("Nur BOM hinzugefuegt:       {0}" -f $statAdd)
Write-Host ("Bereits UTF-8+BOM (skip):   {0}" -f $statSkip)

# --- .frm-Check: muessen ANSI ohne BOM sein --------------------------------
Write-Host "`n=== .frm-Dateien (sollen ANSI ohne BOM sein) ===" -ForegroundColor Cyan
$frms = Get-ChildItem "vba\UserForms" -Filter *.frm -File
foreach ($f in $frms) {
    $b = [System.IO.File]::ReadAllBytes($f.FullName)
    $enc = Get-Encoding $b
    $status = switch ($enc) {
        "UTF8_BOM" { "PROBLEM (BOM -> muss weg)" }
        "UTF8_NoBom" { "PROBLEM (UTF-8 -> muss ANSI)" }
        default { "OK ($enc)" }
    }
    Write-Host ("  {0,-40} {1}" -f $f.Name, $status)
}
