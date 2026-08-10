param(
    [string]$RepoRoot = "c:\Users\DELL Latitude 7490\Desktop\Mein Projekt",
    [string]$WorkbookPath = "c:\Users\DELL Latitude 7490\Desktop\Mein Projekt\excel\Programm Kassenbuch 2018_v2.7.6.xlsm"
)

$ErrorActionPreference = "Stop"

function Get-HashMap {
    param([string[]]$Paths)
    $map = @{}
    foreach ($p in $Paths) {
        $bytes = [System.IO.File]::ReadAllBytes($p)
        $sha = [System.Security.Cryptography.SHA256]::Create()
        try {
            $hash = [System.BitConverter]::ToString($sha.ComputeHash($bytes)).Replace("-", "")
        } finally {
            $sha.Dispose()
        }
        $map[$p] = $hash
    }
    return $map
}

function Has-Utf8Bom {
    param([byte[]]$Bytes)
    return ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF)
}

function Read-TextLikeSync {
    param([string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $utf8 = [System.Text.Encoding]::UTF8
    $cp1252 = [System.Text.Encoding]::GetEncoding(1252)

    if (Has-Utf8Bom -Bytes $bytes) {
        $txt = $utf8.GetString($bytes)
    } else {
        $txt = $utf8.GetString($bytes)
        if ($txt.Contains([char]0xFFFD)) {
            $txt = $cp1252.GetString($bytes)
        }
    }

    if ($txt.Length -gt 0 -and $txt[0] -eq [char]0xFEFF) {
        $txt = $txt.Substring(1)
    }

    return $txt
}

function Extract-FormCode {
    param([string]$FormText)

    $lines = $FormText -split "`r`n|`n"
    $lastAttr = -1
    for ($i = 0; $i -lt $lines.Length; $i++) {
        if ($lines[$i].TrimStart().StartsWith("Attribute ")) {
            $lastAttr = $i
        }
    }

    if ($lastAttr -lt 0) {
        return $FormText
    }

    if ($lastAttr + 1 -ge $lines.Length) {
        return ""
    }

    return ($lines[($lastAttr + 1)..($lines.Length - 1)] -join "`r`n")
}

function Replace-ComponentCodeFromFile {
    param(
        [object]$VbProj,
        [string]$ComponentName,
        [string]$SourcePath,
        [switch]$IsForm
    )

    $comp = $VbProj.VBComponents.Item($ComponentName)
    if ($null -eq $comp) {
        throw "Komponente nicht gefunden: $ComponentName"
    }

    $content = Read-TextLikeSync -Path $SourcePath
    if ($IsForm) {
        $content = Extract-FormCode -FormText $content
    }

    $cm = $comp.CodeModule
    $count = [int]$cm.CountOfLines
    if ($count -gt 0) {
        $cm.DeleteLines(1, $count)
    }
    if (-not [string]::IsNullOrEmpty($content)) {
        $cm.AddFromString($content)
    }
}

function Write-Utf8BomFromAnsiFile {
    param(
        [string]$AnsiPath,
        [string]$DestPath
    )

    $cp1252 = [System.Text.Encoding]::GetEncoding(1252)
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    $txt = [System.IO.File]::ReadAllText($AnsiPath, $cp1252)
    [System.IO.File]::WriteAllText($DestPath, $txt, $utf8Bom)
}

$targets = @(
    "vba/Modules/mod_Banking_Format.bas",
    "vba/Modules/mod_Formatierung.bas",
    "vba/Modules/mod_Mitglieder_UI.bas",
    "vba/Modules/mod_Mitglieder_Logik.bas",
    "vba/Modules/mod_Repo_Sync.bas",
    "vba/UserForms/frm_Mitgliedsdaten.frm"
)
$targetPaths = @($targets | ForEach-Object { Join-Path $RepoRoot $_ })

foreach ($p in $targetPaths) {
    if (-not (Test-Path -LiteralPath $p)) {
        throw "Datei fehlt: $p"
    }
}

Write-Host "ROUNDTRIP_SMOKE=START"
Write-Host "WORKBOOK=$WorkbookPath"

$beforeHashes = Get-HashMap -Paths $targetPaths

$excel = $null
$openedByScript = $false
$createdExcel = $false
$tempDir = Join-Path $env:TEMP ("roundtrip_smoke_" + [DateTime]::Now.ToString("yyyyMMdd_HHmmss"))

try {
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    } catch {
        $excel = New-Object -ComObject Excel.Application
        $createdExcel = $true
    }

    $wb = $null
    foreach ($w in $excel.Workbooks) {
        if ([string]::Equals($w.FullName, $WorkbookPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            $wb = $w
            break
        }
    }

    if ($null -eq $wb) {
        $wb = $excel.Workbooks.Open($WorkbookPath)
        $openedByScript = $true
    }

    $vbProj = $wb.VBProject

    Write-Host "STEP=REPO_TO_VBA"
    foreach ($rel in $targets) {
        $src = Join-Path $RepoRoot $rel
        $name = [System.IO.Path]::GetFileNameWithoutExtension($src)
        $isForm = ([System.IO.Path]::GetExtension($src).ToLowerInvariant() -eq ".frm")
        if ($isForm) {
            Replace-ComponentCodeFromFile -VbProj $vbProj -ComponentName $name -SourcePath $src -IsForm
        } else {
            Replace-ComponentCodeFromFile -VbProj $vbProj -ComponentName $name -SourcePath $src
        }
        Write-Host ("  IMPORTED=" + $rel)
    }

    [System.IO.Directory]::CreateDirectory($tempDir) | Out-Null

    Write-Host "STEP=VBA_TO_REPO"
    foreach ($rel in $targets) {
        $dst = Join-Path $RepoRoot $rel
        $name = [System.IO.Path]::GetFileNameWithoutExtension($dst)
        $ext = [System.IO.Path]::GetExtension($dst).ToLowerInvariant()
        $tempExport = Join-Path $tempDir ([System.IO.Path]::GetFileName($dst))

        $comp = $vbProj.VBComponents.Item($name)
        if ($null -eq $comp) {
            throw "Komponente nicht gefunden fuer Export: $name"
        }

        $comp.Export($tempExport)

        if ($ext -eq ".frm") {
            [System.IO.File]::Copy($tempExport, $dst, $true)
            $tempFrx = [System.IO.Path]::ChangeExtension($tempExport, ".frx")
            $dstFrx = [System.IO.Path]::ChangeExtension($dst, ".frx")
            if (Test-Path -LiteralPath $tempFrx) {
                [System.IO.File]::Copy($tempFrx, $dstFrx, $true)
            }
        } else {
            Write-Utf8BomFromAnsiFile -AnsiPath $tempExport -DestPath $dst
        }

        Write-Host ("  EXPORTED=" + $rel)
    }

    if ($openedByScript) {
        $wb.Close($true)
    }

    if ($createdExcel) {
        $excel.Quit()
    }
} finally {
    if ($null -ne $excel) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $tempDir) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force
    }
}

$afterHashes = Get-HashMap -Paths $targetPaths
$changed = New-Object System.Collections.Generic.List[string]
foreach ($p in $targetPaths) {
    if ($beforeHashes[$p] -ne $afterHashes[$p]) {
        $changed.Add($p)
    }
}

$files = Get-ChildItem (Join-Path $RepoRoot "vba") -Recurse -File -Include *.bas,*.cls,*.frm | Where-Object { $_.FullName -notmatch "\\BackUp" }
$bomIssues = 0
$replFiles = 0
$mojiFiles = 0

foreach ($f in $files) {
    $b = [System.IO.File]::ReadAllBytes($f.FullName)
    $ext = $f.Extension.ToLowerInvariant()
    $hasBom = Has-Utf8Bom -Bytes $b

    if (($ext -eq ".bas" -or $ext -eq ".cls") -and -not $hasBom) { $bomIssues++ }
    if ($ext -eq ".frm" -and $hasBom) { $bomIssues++ }

    $hasRepl = $false
    for ($i = 0; $i -le $b.Length - 3; $i++) {
        if ($b[$i] -eq 0xEF -and $b[$i + 1] -eq 0xBF -and $b[$i + 2] -eq 0xBD) {
            $hasRepl = $true
            break
        }
    }
    if ($hasRepl) { $replFiles++ }

    $hasMoji = $false
    for ($i = 0; $i -le $b.Length - 4; $i++) {
        if ($b[$i] -eq 0xC3 -and $b[$i + 1] -eq 0x83 -and $b[$i + 2] -eq 0xC2 -and ($b[$i + 3] -in 0xA4,0xB6,0xBC,0x84,0x96,0x9C,0x9F)) {
            $hasMoji = $true
            break
        }
    }
    if ($hasMoji) { $mojiFiles++ }
}

Write-Host "STEP=VERIFY"
Write-Host ("TARGET_FILES=" + $targets.Count)
Write-Host ("TARGET_HASH_CHANGES=" + $changed.Count)
foreach ($p in $changed) {
    Write-Host ("  HASH_CHANGED=" + $p)
}
Write-Host ("FILES_SCANNED=" + $files.Count)
Write-Host ("BOM_ISSUES=" + $bomIssues)
Write-Host ("REPLACEMENT_FILES=" + $replFiles)
Write-Host ("MOJIBAKE_FILES=" + $mojiFiles)
Write-Host "ROUNDTRIP_SMOKE=END"