# ============================================================================
# fix_umlauts_dictionary.ps1
# ----------------------------------------------------------------------------
# Ersetzt ae/oe/ue durch echte Umlaute (ä/ö/ü) ausschliesslich in:
#   - Kommentaren (Zeilen die mit ' beginnen)
#   - String-Literalen (zwischen "...")
#
# Whitelist deutscher Woerter aus tools\umlaut_dictionary.txt
# (Format pro Zeile: "muster|ersatz", case-sensitive, laengere zuerst).
#
# AUSGESCHLOSSENE Dateien (enthalten absichtliche ae/oe/ue in Code):
#   mod_EntityKey_Normalize.bas
#   mod_Mapping_Tools.bas
#   mod_Repo_Sync.bas       (frisch v3.3)
#   mod_VBA_Export.bas      (frisch v1.2)
#
# Ausgabe: UTF-8 mit BOM.
# ============================================================================
param([switch]$DryRun)

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)

$excluded = @(
    "mod_EntityKey_Normalize.bas",
    "mod_Mapping_Tools.bas",
    "mod_Repo_Sync.bas",
    "mod_VBA_Export.bas"
)

# ---- Wörterliste einlesen -------------------------------------------------
$dictPath = "tools\umlaut_dictionary.txt"
if (-not (Test-Path $dictPath)) {
    Write-Error "Woerterliste fehlt: $dictPath"
    exit 1
}

$dictBytes = [System.IO.File]::ReadAllBytes($dictPath)
$dictText  = [System.Text.Encoding]::UTF8.GetString($dictBytes).TrimStart([char]0xFEFF)

$pairs = New-Object System.Collections.ArrayList
foreach ($line in ($dictText -split "`r?`n")) {
    $t = $line.Trim()
    if ($t -eq "" -or $t.StartsWith("#")) { continue }
    $ix = $t.IndexOf("|")
    if ($ix -lt 1) { continue }
    $key = $t.Substring(0, $ix)
    $val = $t.Substring($ix + 1)
    if ($key -eq "" -or $val -eq "") { continue }
    [void]$pairs.Add(@($key, $val))
}
Write-Host ("Wörterbuch geladen: {0} Muster" -f $pairs.Count) -ForegroundColor DarkCyan


# ---- Vorkompilierte Regex fuer Speed --------------------------------------
$regexes = New-Object System.Collections.ArrayList
foreach ($p in $pairs) {
    $pattern = '(?<![A-Za-z0-9_])' + [regex]::Escape($p[0]) + '(?![A-Za-z0-9_])'
    $rx = New-Object System.Text.RegularExpressions.Regex($pattern, [System.Text.RegularExpressions.RegexOptions]::Compiled)
    [void]$regexes.Add(@($rx, $p[1]))
}


function Fix-Text {
    param([string]$text)
    $result = $text
    foreach ($r in $regexes) {
        $result = $r[0].Replace($result, $r[1])
    }
    return $result
}


function Fix-Line {
    param([string]$line)
    $trim = $line.TrimStart()
    if ($trim.StartsWith("'")) {
        return Fix-Text $line
    }

    # Kein Kommentar -> nur "..." Substrings ersetzen
    $sb = New-Object System.Text.StringBuilder
    $i = 0
    while ($i -lt $line.Length) {
        $ch = $line[$i]
        if ($ch -eq '"') {
            $start = $i
            $j = $i + 1
            while ($j -lt $line.Length -and $line[$j] -ne '"') { $j++ }
            if ($j -lt $line.Length) {
                $str = $line.Substring($start, $j - $start + 1)
                $newStr = Fix-Text $str
                [void]$sb.Append($newStr)
                $i = $j + 1
            } else {
                [void]$sb.Append($ch)
                $i++
            }
        } else {
            [void]$sb.Append($ch)
            $i++
        }
    }
    return $sb.ToString()
}


$targets = @()
$targets += Get-ChildItem "vba\Modules" -Filter *.bas -File
$targets += Get-ChildItem "vba\Classes" -Filter *.cls -File

Write-Host "`n=== ae/oe/ue -> Umlaute (Kommentare + Strings) ===" -ForegroundColor Cyan
$modus = if ($DryRun) { 'DRY-RUN' } else { 'LIVE' }
Write-Host "Modus: $modus | Ziele: $($targets.Count) (davon ausgeschlossen: $($excluded.Count))"

$totalChanged = 0
$fileChanges = @{}
$utf8Bom = New-Object System.Text.UTF8Encoding($true)

foreach ($f in $targets) {
    if ($excluded -contains $f.Name) {
        Write-Host ("  SKIP  {0}" -f $f.Name) -ForegroundColor DarkGray
        continue
    }

    $originalBytes = [System.IO.File]::ReadAllBytes($f.FullName)
    $original = [System.Text.Encoding]::UTF8.GetString($originalBytes)
    $original = $original.TrimStart([char]0xFEFF)

    $lines = $original -split "`r?`n"
    $changedCount = 0
    $newLines = New-Object System.Collections.ArrayList
    foreach ($ln in $lines) {
        $new = Fix-Line -line $ln
        if ($new -ne $ln) { $changedCount++ }
        [void]$newLines.Add($new)
    }

    if ($changedCount -gt 0) {
        $totalChanged += $changedCount
        $fileChanges[$f.Name] = $changedCount
        if (-not $DryRun) {
            $result = ($newLines -join "`r`n")
            [System.IO.File]::WriteAllText($f.FullName, $result, $utf8Bom)
        }
        Write-Host ("  FIX   {0,4} Zeilen  {1}" -f $changedCount, $f.Name) -ForegroundColor Green
    }
}

Write-Host "`n=== Zusammenfassung ===" -ForegroundColor Cyan
Write-Host ("Geaenderte Zeilen gesamt: {0}" -f $totalChanged)
Write-Host ("Dateien mit Aenderungen : {0}" -f $fileChanges.Count)
