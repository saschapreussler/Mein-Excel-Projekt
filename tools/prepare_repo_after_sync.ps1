param(
    [switch]$AutoCommit,
    [string]$CommitMessage = "Post-sync cleanup",
    [switch]$CheckOnly,
    [switch]$SkipRepair
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

Write-Host "=== Prepare Repo After Sync ===" -ForegroundColor Cyan
if(-not $CheckOnly -and -not $SkipRepair){
    Write-Host "Step 1/3: Repair mojibake and normalize encodings" -ForegroundColor Cyan
    & (Join-Path $repoRoot 'tools/repair_qm_to_umlauts.ps1')
    & (Join-Path $repoRoot 'tools/convert_repo_to_utf8bom.ps1')
} else {
    Write-Host "Step 1/3: Repair skipped" -ForegroundColor DarkYellow
}

Write-Host "Step 2/3: Integrity check" -ForegroundColor Cyan
$root = Join-Path $repoRoot 'vba'
$files = Get-ChildItem $root -Recurse -File -Include *.bas,*.cls,*.frm | Where-Object { $_.FullName -notmatch '\\BackUp' }
$replFiles = 0
$mojiFiles = 0
$bomIssues = 0
$frmUtf8Issues = 0
foreach($f in $files){
    $b = [System.IO.File]::ReadAllBytes($f.FullName)
    $ext = $f.Extension.ToLowerInvariant()

    $hasBom = ($b.Length -ge 3 -and $b[0]-eq 0xEF -and $b[1]-eq 0xBB -and $b[2]-eq 0xBF)
    if(($ext -eq '.bas' -or $ext -eq '.cls') -and -not $hasBom){ $bomIssues++ }
    if($ext -eq '.frm' -and $hasBom){ $bomIssues++ }

    $hasRepl = $false
    for($i=0;$i -le $b.Length-3;$i++){
        if($b[$i]-eq 0xEF -and $b[$i+1]-eq 0xBF -and $b[$i+2]-eq 0xBD){ $hasRepl = $true; break }
    }
    if($hasRepl){ $replFiles++ }

    $hasMoji = $false
    for($i=0;$i -le $b.Length-6;$i++){
        if($b[$i]-eq 0xC3 -and $b[$i+1]-eq 0xAF -and $b[$i+2]-eq 0xC2 -and $b[$i+3]-eq 0xBF -and $b[$i+4]-eq 0xC2 -and $b[$i+5]-eq 0xBD){ $hasMoji = $true; break }
    }
    if($hasMoji){ $mojiFiles++ }

    if($ext -eq '.frm'){
        $hasHigh = $false
        foreach($x in $b){ if($x -ge 128){ $hasHigh = $true; break } }
        if($hasHigh){
            $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
            $validUtf8 = $true
            try { [void]$utf8.GetString($b) } catch { $validUtf8 = $false }
            if($validUtf8){ $frmUtf8Issues++ }
        }
    }
}
Write-Host ("FILES_SCANNED={0}" -f $files.Count)
Write-Host ("BOM_ISSUES={0}" -f $bomIssues)
Write-Host ("REPLACEMENT_BYTES_FILES={0}" -f $replFiles)
Write-Host ("VISIBLE_MOJI_FILES={0}" -f $mojiFiles)
Write-Host ("FRM_UTF8_ISSUES={0}" -f $frmUtf8Issues)

$integrityOk = ($bomIssues -eq 0 -and $replFiles -eq 0 -and $mojiFiles -eq 0 -and $frmUtf8Issues -eq 0)
if(-not $integrityOk){
    throw "Integrity check failed. Sync cleanup stopped before commit."
}

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
