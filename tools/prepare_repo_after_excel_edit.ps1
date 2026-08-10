param()

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

Write-Host "=== Prepare Repo After Excel Edit ===" -ForegroundColor Cyan
Write-Host "Step 1/3: Repair mojibake markers and umlauts" -ForegroundColor Cyan

$fixScript = Join-Path $repoRoot 'tools/fix_mojibake_safe.ps1'
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

Write-Host "Done. Review diffs, then commit." -ForegroundColor Green
