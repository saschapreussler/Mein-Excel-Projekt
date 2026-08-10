param(
    [switch]$AutoCommit,
    [string]$CommitMessage = "Post-sync cleanup"
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

Write-Host "=== Prepare Repo After Sync ===" -ForegroundColor Cyan
Write-Host "Step 1/3: Repair mojibake and normalize encodings" -ForegroundColor Cyan
& (Join-Path $repoRoot 'tools/repair_qm_to_umlauts.ps1')
& (Join-Path $repoRoot 'tools/convert_repo_to_utf8bom.ps1')

Write-Host "Step 2/3: Integrity check" -ForegroundColor Cyan
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

Write-Host "Step 3/3: Git status" -ForegroundColor Cyan
git status --short

if($AutoCommit){
    $dirty = git status --porcelain
    if([string]::IsNullOrWhiteSpace($dirty)){
        Write-Host "Working tree already clean." -ForegroundColor Green
    } else {
        git add -A
        git commit -m $CommitMessage
        git push origin main
        Write-Host "Auto-commit and push completed." -ForegroundColor Green
    }
}
