# =====================================================================
# fix_mixed_case_umlauts.ps1
# ---------------------------------------------------------------------
# KORREKTUR-SKRIPT: behebt die Folgen eines vorhergehenden Skript-
# Bugs (PowerShell-Variablen sind case-insensitive, daher wurde
# $ue (Lowercase) durch $UE (Uppercase) ueberschrieben).
#
# Wirkung: ueberall wo ein GROSSES Umlaut-Zeichen (Ä, Ö, Ü) unmittelbar
# neben einem KLEINBUCHSTABEN steht, wird es in das entsprechende
# kleine Umlaut-Zeichen (ä, ö, ü) konvertiert.
#
# Das laesst echte Grossbuchstaben-Sequenzen wie "FÜR", "ÄNDERUNG",
# "ÜBERSICHT" unangetastet.
# =====================================================================

$ErrorActionPreference = "Stop"
Set-Location -Path (Split-Path -Parent $PSScriptRoot)
$cp = [System.Text.Encoding]::GetEncoding(1252)

# Mapping: GrossUmlaut -> KleinUmlaut
$mapping = @{
    [char]196 = [char]228   # Ä -> ä
    [char]214 = [char]246   # Ö -> ö
    [char]220 = [char]252   # Ü -> ü
}

# Regex: Grossumlaut der entweder LINKS oder RECHTS einen
# Kleinbuchstaben (a-z) hat.
$regex = [regex]'(?<=[a-z])[\xC4\xD6\xDC]|[\xC4\xD6\xDC](?=[a-z])'

$files = @()
$files += Get-ChildItem -Path 'vba' -Recurse -Include *.bas,*.cls,*.frm |
          Where-Object { $_.FullName -notmatch '\\BackUp' }

$grand = 0
$touched = 0
foreach ($f in $files) {
    $txt = [System.IO.File]::ReadAllText($f.FullName, $cp)
    $cnt = 0
    $new = $regex.Replace($txt, {
        param($m)
        $script:cnt++
        $c = $m.Value[0]
        return [string]$mapping[$c]
    })
    if ($new -ne $txt) {
        [System.IO.File]::WriteAllText($f.FullName, $new, $cp)
        Write-Host ("  {0,-50}  -{1}" -f $f.Name, $cnt)
        $grand += $cnt
        $touched++
    }
}

Write-Host ""
Write-Host "===========================================" -ForegroundColor Green
Write-Host ("Geaenderte Dateien : {0}" -f $touched) -ForegroundColor Green
Write-Host ("Korrekturen total  : {0}" -f $grand) -ForegroundColor Green
Write-Host "===========================================" -ForegroundColor Green
