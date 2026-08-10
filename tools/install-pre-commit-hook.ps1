param()

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

$hooksPath = 'tools/git-hooks'
$hookFile = Join-Path $repoRoot 'tools/git-hooks/pre-commit'

if (-not (Test-Path $hookFile)) {
    throw "Hook file not found: $hookFile"
}

git config core.hooksPath $hooksPath

Write-Host "core.hooksPath set to: $hooksPath" -ForegroundColor Green
Write-Host "Pre-commit hook is now active for this local repository." -ForegroundColor Green
Write-Host ""
Write-Host "Verify:" -ForegroundColor Cyan
Write-Host "  git config --get core.hooksPath"
