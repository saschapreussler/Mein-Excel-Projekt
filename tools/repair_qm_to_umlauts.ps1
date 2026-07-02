# ============================================================================
# repair_qm_to_umlauts.ps1
# ----------------------------------------------------------------------------
# Ersetzt "?" (ASCII 0x3F) durch die richtigen deutschen Umlaute (ä, ö, ü, ß)
# in .bas/.cls/.frm-Dateien -- ausschliesslich in:
#   - Kommentaren (Zeilen mit fuehrendem ')
#   - String-Literalen (zwischen "...")
#
# Basiert auf tools\umlaut_dictionary.txt. Fuer jede Zeile "key|value" wird
# aus der Umlaut-Version (value) ein "?"-Muster generiert:
#     value = "für"        -> pattern "f?r"        -> value
#     value = "Übersicht"  -> pattern "?bersicht"  -> value
#     value = "Größe"      -> pattern "Gr??e"      -> value
#
# Encoding-Handling:
#   .bas / .cls  -> UTF-8 mit BOM (wie im Repo Standard)
#   .frm         -> Windows-1252 (ANSI) ohne BOM (fuer VBA-Import)
#
# Ausschluss (absichtliche "?"/"ae"/"oe"/"ue" im Code):
#   mod_EntityKey_Normalize.bas
#   mod_Mapping_Tools.bas
#   mod_Repo_Sync.bas
#   mod_VBA_Export.bas
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

# ---- Woerterliste laden + "?"-Muster generieren ---------------------------
$dictPath = "tools\umlaut_dictionary.txt"
if (-not (Test-Path $dictPath)) {
    Write-Error "Woerterliste fehlt: $dictPath"
    exit 1
}
$dictBytes = [System.IO.File]::ReadAllBytes($dictPath)
$dictText  = [System.Text.Encoding]::UTF8.GetString($dictBytes).TrimStart([char]0xFEFF)

# Sammle einzigartige Werte (Umlaut-Version). Aus jedem Wert generieren wir
# ein "?"-Muster (ä/ö/ü/Ä/Ö/Ü/ß -> ?), keine Duplikate.
$umlautChars = @{ [char]'ä'=$true; [char]'ö'=$true; [char]'ü'=$true; [char]'Ä'=$true; [char]'Ö'=$true; [char]'Ü'=$true; [char]'ß'=$true }
$valueSet = New-Object System.Collections.Generic.HashSet[string]
foreach ($line in ($dictText -split "`r?`n")) {
    $t = $line.Trim()
    if ($t -eq "" -or $t.StartsWith("#")) { continue }
    $ix = $t.IndexOf("|")
    if ($ix -lt 1) { continue }
    $val = $t.Substring($ix + 1)
    if ($val -eq "") { continue }
    # Nur Werte mit mindestens einem Umlaut aufnehmen
    $hasUmlaut = $false
    foreach ($ch in $val.ToCharArray()) { if ($umlautChars.ContainsKey($ch)) { $hasUmlaut = $true; break } }
    if ($hasUmlaut) { [void]$valueSet.Add($val) }
}

# Baue Pairs: (qm_pattern, umlaut_replacement)
$pairs = New-Object System.Collections.ArrayList
foreach ($val in $valueSet) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $val.ToCharArray()) {
        if ($umlautChars.ContainsKey($ch)) {
            [void]$sb.Append('?')
        } else {
            [void]$sb.Append($ch)
        }
    }
    $qm = $sb.ToString()
    # Nur behalten wenn qm != val (also mind. ein Umlaut ersetzt)
    if ($qm -ne $val) {
        [void]$pairs.Add(@($qm, $val))
    }
}

# Nach Laenge sortieren (laengere zuerst -> laengere Match-Prioritaet)
$pairs = @($pairs | Sort-Object -Property @{Expression={($_[0]).Length}} -Descending)
Write-Host ("Woerter-Patterns: {0}" -f $pairs.Count) -ForegroundColor DarkCyan


# ---- Regex vorbauen -------------------------------------------------------
$regexList = New-Object System.Collections.ArrayList
foreach ($p in $pairs) {
    # Word-Boundary: davor / danach kein alphanum / underscore
    $pattern = '(?<![A-Za-z0-9_])' + [regex]::Escape($p[0]) + '(?![A-Za-z0-9_])'
    $rx = New-Object System.Text.RegularExpressions.Regex($pattern, [System.Text.RegularExpressions.RegexOptions]::Compiled)
    [void]$regexList.Add(@($rx, $p[1]))
}


function Fix-Text {
    param([string]$text)
    $result = $text
    foreach ($r in $regexList) { $result = $r[0].Replace($result, $r[1]) }
    return $result
}


function Fix-Line {
    param([string]$line, [bool]$isFrm)
    $trim = $line.TrimStart()
    if ($trim.StartsWith("'")) {
        return Fix-Text $line
    }

    # kein Kommentar -> nur "..." Substrings ersetzen
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
                [void]$sb.Append((Fix-Text $str))
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


# ---- Ziel-Dateien sammeln (bas/cls/frm) -----------------------------------
$targets = @()
$targets += Get-ChildItem "vba\Modules"   -Filter *.bas -File
$targets += Get-ChildItem "vba\Classes"   -Filter *.cls -File
$targets += Get-ChildItem "vba\UserForms" -Filter *.frm -File

Write-Host "`n=== '?' -> Umlaute Recovery ===" -ForegroundColor Cyan
$modus = if ($DryRun) { 'DRY-RUN' } else { 'LIVE' }
Write-Host "Modus: $modus | Ziele: $($targets.Count) | Ausgeschlossen: $($excluded.Count)"

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$win1252 = [System.Text.Encoding]::GetEncoding(1252)

$totalChanged = 0
$fileChanges  = @{}

foreach ($f in $targets) {
    if ($excluded -contains $f.Name) {
        Write-Host ("  SKIP  {0}" -f $f.Name) -ForegroundColor DarkGray
        continue
    }

    $isFrm = ($f.Extension -eq ".frm")

    # Encoding-aware lesen
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
    $original = $null

    if ($isFrm) {
        # .frm ist ANSI (Windows-1252)
        $original = $win1252.GetString($bytes)
    } else {
        # .bas/.cls: BOM-aware UTF-8, mit ANSI-Fallback wenn UTF-8 invalid
        $bomOk = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
        try {
            $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
            $test = $utf8Strict.GetString($bytes)
            $original = $test.TrimStart([char]0xFEFF)
        } catch {
            $original = $win1252.GetString($bytes)
        }
    }

    $lines = $original -split "`r?`n"
    $changedCount = 0
    $newLines = New-Object System.Collections.ArrayList
    foreach ($ln in $lines) {
        $new = Fix-Line -line $ln -isFrm $isFrm
        if ($new -ne $ln) { $changedCount++ }
        [void]$newLines.Add($new)
    }

    if ($changedCount -gt 0) {
        $totalChanged += $changedCount
        $fileChanges[$f.Name] = $changedCount
        if (-not $DryRun) {
            $result = ($newLines -join "`r`n")
            if ($isFrm) {
                # .frm: Windows-1252 (ANSI) ohne BOM
                $outBytes = $win1252.GetBytes($result)
                [System.IO.File]::WriteAllBytes($f.FullName, $outBytes)
            } else {
                # .bas/.cls: UTF-8 mit BOM
                [System.IO.File]::WriteAllText($f.FullName, $result, $utf8Bom)
            }
        }
        Write-Host ("  FIX   {0,5} Zeilen  {1}" -f $changedCount, $f.Name) -ForegroundColor Green
    }
}

Write-Host "`n=== Zusammenfassung ===" -ForegroundColor Cyan
Write-Host ("Geänderte Zeilen gesamt: {0}" -f $totalChanged)
Write-Host ("Dateien mit Änderungen : {0}" -f $fileChanges.Count)
