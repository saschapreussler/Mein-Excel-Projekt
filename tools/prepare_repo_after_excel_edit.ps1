param()

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

Write-Host "=== Prepare Repo After Excel Edit ===" -ForegroundColor Cyan
Write-Host "Step 1/3: Repair mojibake markers and umlauts" -ForegroundColor Cyan

$fixScript = Join-Path $repoRoot 'tools/repair_qm_to_umlauts.ps1'
if (-not (Test-Path $fixScript)) {
    throw "Missing script: $fixScript"
}
& $fixScript

Write-Host "Step 2/3: Normalize BOM policy" -ForegroundColor Cyan
$convScript = Join-Path $repoRoot 'tools/convert_repo_to_utf8bom.ps1'
if (-not (Test-Path $convScript)) {
    throw "Missing script: $convScript"
}
& $convScript

Write-Host "Step 3/3: Show Git status" -ForegroundColor Cyan
git status --short

Write-Host "`nStep 4/4: Quick byte-integrity check" -ForegroundColor Cyan
$root = Join-Path $repoRoot 'vba'
$files = Get-ChildItem $root -Recurse -File -Include *.bas,*.cls,*.frm | Where-Object { $_.FullName -notmatch '\\BackUp' }
$replFiles = 0
foreach($f in $files){
    $b = [System.IO.File]::ReadAllBytes($f.FullName)
    $hasRepl = $false
    for($i=0;$i -le $b.Length-3;$i++){
        if($b[$i]-eq 0xEF -and $b[$i+1]-eq 0xBF -and $b[$i+2]-eq 0xBD){ $hasRepl = $true; break }
    }
    if($hasRepl){ $replFiles++ }
}
Write-Host ("REPLACEMENT_BYTES_FILES={0}" -f $replFiles)

Write-Host "Done. Review diffs, then commit." -ForegroundColor Green
