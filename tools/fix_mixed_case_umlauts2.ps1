# Simpler/robuster: zwei Pass-Regex mit static MatchEvaluators schwierig
# in PS - daher Replace mit Lookbehind/Lookahead pro Zeichen einzeln.
$ErrorActionPreference = "Stop"
Set-Location -Path (Split-Path -Parent $PSScriptRoot)
$cp = [System.Text.Encoding]::GetEncoding(1252)

$files = @()
$files += Get-ChildItem -Path 'vba' -Recurse -Include *.bas,*.cls,*.frm |
          Where-Object { $_.FullName -notmatch '\\BackUp' }

# Drei einzelne Regex-Patterns: Ae/Oe/Ue jeweils mit Lookbehind ODER
# Lookahead auf [a-z]. Replace mit fixem Lowercase-Char.
$patterns = @(
    @('(?<=[a-z])\xC4|\xC4(?=[a-z])', [char]228),
    @('(?<=[a-z])\xD6|\xD6(?=[a-z])', [char]246),
    @('(?<=[a-z])\xDC|\xDC(?=[a-z])', [char]252)
)

$grand = 0
$touched = 0
foreach ($f in $files) {
    $txt = [System.IO.File]::ReadAllText($f.FullName, $cp)
    $orig = $txt
    $local = 0
    foreach ($p in $patterns) {
        $rx = [regex]$p[0]
        $repl = [string]$p[1]
        $matches = $rx.Matches($txt)
        if ($matches.Count -gt 0) {
            $local += $matches.Count
            $txt = $rx.Replace($txt, $repl)
        }
    }
    if ($txt -cne $orig) {
        [System.IO.File]::WriteAllText($f.FullName, $txt, $cp)
        Write-Host ("  {0,-50}  -{1}" -f $f.Name, $local)
        $grand += $local
        $touched++
    }
}
Write-Host ""
Write-Host ("Geaenderte Dateien : {0}" -f $touched) -ForegroundColor Green
Write-Host ("Korrekturen total  : {0}" -f $grand) -ForegroundColor Green
