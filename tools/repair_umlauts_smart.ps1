# ============================================================================
# repair_umlauts_smart.ps1
# ----------------------------------------------------------------------------
# Repariert FFFD (U+FFFD, Unicode Replacement Character) in VBA-Repo-Dateien
# durch kontext-basierte Umlaut-Wiederherstellung. Case-sensitive Ersetzung.
#
# Verwendung:
#   powershell -ExecutionPolicy Bypass -File tools\repair_umlauts_smart.ps1
#   powershell -ExecutionPolicy Bypass -File tools\repair_umlauts_smart.ps1 -DryRun
# ============================================================================
param(
    [switch]$DryRun,
    [string[]]$Files
)

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)

$F = [char]0xFFFD  # Kurz-Alias

# Reihenfolge WICHTIG: laengere/spezifischere Muster zuerst.
# Case-sensitive Ersetzung. Format: @(pattern, replacement)
$pairs = @(
    # ---- "ss/ß"-Faelle (immer VOR "?"-Vokal-Regeln!) ----
    ,@("gr${F}${F}er", 'größer')
    ,@("Gr${F}${F}er", 'Größer')
    ,@("gr${F}${F}",   'größ')
    ,@("Gr${F}${F}",   'Größ')
    ,@("Fu${F}",       'Fuß')
    ,@("Stra${F}e",    'Straße')
    ,@("wei${F}",      'weiß')
    ,@("Wei${F}",      'Weiß')
    ,@("hei${F}",      'heiß')
    ,@("Hei${F}",      'Heiß')
    ,@("gro${F}",      'groß')
    ,@("Gro${F}",      'Groß')

    # ---- "ueber" (Ü/ü) ----
    ,@("${F}bersicht", 'Übersicht')
    ,@("${F}bertrag",  'Übertrag')
    ,@("${F}berpr${F}f", 'überprüf')
    ,@("${F}berneh",   'überneh')
    ,@("${F}berge",    'überge')
    ,@("${F}berga",    'überga')
    ,@("${F}berl",     'überl')
    ,@("${F}bers",     'übers')
    ,@("${F}berk",     'überk')
    ,@("${F}berp",     'überp')
    ,@("${F}berst",    'überst')
    ,@("${F}bert",     'übert')
    ,@("${F}berw",     'überw')
    ,@("${F}berz",     'überz')
    ,@("${F}ber",      'über')

    # ---- "moecht"/"moegli" ----
    ,@("M${F}chten",   'Möchten')
    ,@("m${F}chten",   'möchten')
    ,@("M${F}chte",    'Möchte')
    ,@("m${F}chte",    'möchte')
    ,@("M${F}gli",     'Mögli')
    ,@("m${F}gli",     'mögli')
    ,@("M${F}g",       'Mög')
    ,@("m${F}g",       'mög')

    # ---- "koenn" ----
    ,@("K${F}nnen",    'Können')
    ,@("k${F}nnen",    'können')
    ,@("K${F}nnt",     'Könnt')
    ,@("k${F}nnt",     'könnt')
    ,@("K${F}nne",     'Könne')
    ,@("k${F}nne",     'könne')

    # ---- "muess" ----
    ,@("M${F}ss",      'Müss')
    ,@("m${F}ss",      'müss')

    # ---- "fuer" ----
    ,@("F${F}R",       'FÜR')
    ,@("F${F}r",       'Für')
    ,@("f${F}r",       'für')

    # ---- "wuensch" / "wuerde" ----
    ,@("W${F}nsch",    'Wünsch')
    ,@("w${F}nsch",    'wünsch')
    ,@("Gew${F}nsch",  'Gewünsch')
    ,@("gew${F}nsch",  'gewünsch')
    ,@("W${F}rd",      'Würd')
    ,@("w${F}rd",      'würd')

    # ---- "natuerlich" / "pruef" / "regulaer" ----
    ,@("Nat${F}rli",   'Natürli')
    ,@("nat${F}rli",   'natürli')
    ,@("Pr${F}f",      'Prüf')
    ,@("pr${F}f",      'prüf')
    ,@("Regul${F}r",   'Regulär')
    ,@("regul${F}r",   'regulär')

    # ---- "loesch" / "loesen" / "loesung" ----
    ,@("L${F}sch",     'Lösch')
    ,@("l${F}sch",     'lösch')
    ,@("Gel${F}scht",  'Gelöscht')
    ,@("gel${F}scht",  'gelöscht')
    ,@("L${F}sen",     'Lösen')
    ,@("l${F}sen",     'lösen')
    ,@("L${F}sung",    'Lösung')
    ,@("l${F}sung",    'lösung')
    ,@("l${F}st",      'löst')
    ,@("l${F}se",      'löse')

    # ---- "oeffn" / "geoeffnet" ----
    ,@("ge${F}ffnet",  'geöffnet')
    ,@("${F}ffnet",    'öffnet')
    ,@("${F}ffn",      'öffn')
    ,@("${F}ffent",    'öffent')

    # ---- "schoen" ----
    ,@("Sch${F}n",     'Schön')
    ,@("sch${F}n",     'schön')

    # ---- "hoer" / "hoeher" / "hoech" ----
    ,@("H${F}ren",     'Hören')
    ,@("h${F}ren",     'hören')
    ,@("h${F}rt",      'hört')
    ,@("H${F}her",     'Höher')
    ,@("h${F}her",     'höher')
    ,@("H${F}ch",      'Höch')
    ,@("h${F}ch",      'höch')

    # ---- "stoer" ----
    ,@("St${F}r",      'Stör')
    ,@("st${F}r",      'stör')

    # ---- "aender" / "aehnli" ----
    ,@("${F}nder",     'änder')
    ,@("${F}hnli",     'ähnli')
    ,@("${F}hnl",      'ähnl')
    ,@("${F}rzt",      'ärzt')
    ,@("${F}ltest",    'ältest')
    ,@("${F}rgerli",   'ärgerli')

    # ---- "naechst" / "naeher" ----
    ,@("N${F}chst",    'Nächst')
    ,@("n${F}chst",    'nächst')
    ,@("N${F}her",     'Näher')
    ,@("n${F}her",     'näher')

    # ---- "fuell" / "fuehl" / "fuehr" ----
    ,@("F${F}lle",     'Fülle')
    ,@("f${F}lle",     'fülle')
    ,@("F${F}ll",      'Füll')
    ,@("f${F}ll",      'füll')
    ,@("F${F}hr",      'Führ')
    ,@("f${F}hr",      'führ')
    ,@("F${F}hl",      'Fühl')
    ,@("f${F}hl",      'fühl')

    # ---- "gruen" ----
    ,@("Gr${F}n",      'Grün')
    ,@("gr${F}n",      'grün')

    # ---- "rueck" / "stueck" / "brueck" / "drueck" ----
    ,@("R${F}ck",      'Rück')
    ,@("r${F}ck",      'rück')
    ,@("St${F}ck",     'Stück')
    ,@("st${F}ck",     'stück')
    ,@("Br${F}ck",     'Brück')
    ,@("br${F}ck",     'brück')
    ,@("Dr${F}ck",     'Drück')
    ,@("dr${F}ck",     'drück')
    ,@("gedr${F}ckt",  'gedrückt')

    # ---- "waerme" / "waehl" / "waehrung" ----
    ,@("W${F}rme",     'Wärme')
    ,@("w${F}rme",     'wärme')
    ,@("W${F}hl",      'Wähl')
    ,@("w${F}hl",      'wähl')
    ,@("gew${F}hl",    'gewähl')
    ,@("W${F}hrend",   'Während')
    ,@("w${F}hrend",   'während')
    ,@("W${F}hrung",   'Währung')
    ,@("w${F}hrung",   'währung')

    # ---- "stand" / "unterstuetzt" ----
    ,@("Unterst${F}tz", 'Unterstütz')
    ,@("unterst${F}tz", 'unterstütz')
    ,@("Umst${F}nde",  'Umstände')
    ,@("St${F}nde",    'Stände')
    ,@("st${F}nde",    'stände')

    # ---- "koenig" / "koerper" / "kuerz" ----
    ,@("K${F}nig",     'König')
    ,@("k${F}nig",     'könig')
    ,@("K${F}rper",    'Körper')
    ,@("k${F}rper",    'körper')
    ,@("K${F}rzen",    'Kürzen')
    ,@("k${F}rzen",    'kürzen')
    ,@("k${F}rzer",    'kürzer')
    ,@("gek${F}rzt",   'gekürzt')

    # ---- "buendel" / "nuetz" ----
    ,@("B${F}ndel",    'Bündel')
    ,@("b${F}ndel",    'bündel')
    ,@("N${F}tz",      'Nütz')
    ,@("n${F}tz",      'nütz')
)

function Fix-File {
    param([string]$path)

    if (-not (Test-Path $path)) {
        Write-Warning "Nicht gefunden: $path"
        return $null
    }

    $original = [System.IO.File]::ReadAllText($path)
    $before = 0
    foreach ($ch in $original.ToCharArray()) { if ($ch -eq $F) { $before++ } }

    if ($before -eq 0) {
        Write-Host "  OK (keine FFFD): $path" -ForegroundColor DarkGray
        return @{ Path=$path; Before=0; After=0; Fixed=0 }
    }

    $fixed = $original
    foreach ($p in $pairs) {
        $fixed = $fixed.Replace($p[0], $p[1])
    }

    $after = 0
    foreach ($ch in $fixed.ToCharArray()) { if ($ch -eq $F) { $after++ } }
    $repaired = $before - $after

    if ($DryRun) {
        Write-Host ("  DRYRUN {0,4} -> {1,4}  ({2})" -f $before, $after, $path) -ForegroundColor Yellow
    } else {
        $enc = New-Object System.Text.UTF8Encoding($true)  # UTF-8 + BOM
        [System.IO.File]::WriteAllText($path, $fixed, $enc)
        Write-Host ("  FIXED  {0,4} -> {1,4}  ({2})" -f $before, $after, $path) -ForegroundColor Green
    }

    return @{ Path=$path; Before=$before; After=$after; Fixed=$repaired }
}

if ($null -eq $Files -or $Files.Count -eq 0) {
    $Files = @(
        "vba\Modules\mod_Repo_Sync.bas",
        "vba\Modules\mod_VBA_Export.bas",
        "vba\UserForms\frm_Mitgliederverwaltung.frm",
        "vba\UserForms\frm_Mitgliedsdaten.frm"
    )
}

Write-Host "`n=== Umlaut-Recovery ===" -ForegroundColor Cyan
$modus = if ($DryRun) { 'DRY-RUN' } else { 'LIVE' }
Write-Host ("Modus: {0}" -f $modus)
Write-Host ""

$results = @()
foreach ($f in $Files) {
    $r = Fix-File -path $f
    if ($null -ne $r) { $results += $r }
}

Write-Host "`n=== Zusammenfassung ===" -ForegroundColor Cyan
$total  = ($results | Measure-Object -Property Fixed -Sum).Sum
$remain = ($results | Measure-Object -Property After -Sum).Sum
Write-Host ("Repariert: {0}  |  Verbleibend: {1}" -f $total, $remain)

if ($remain -gt 0 -and -not $DryRun) {
    Write-Host "`nVerbleibende FFFD-Stellen (max. 5 je Datei):" -ForegroundColor Yellow
    foreach ($r in $results) {
        if ($r.After -gt 0) {
            $txt = [System.IO.File]::ReadAllText($r.Path)
            $lines = $txt -split "`r`n"
            $ln = 0
            $shown = 0
            foreach ($line in $lines) {
                $ln++
                if ($line.Contains($F) -and $shown -lt 5) {
                    $idx = $line.IndexOf($F)
                    $from = [Math]::Max(0, $idx-30)
                    $to   = [Math]::Min($line.Length, $idx+30)
                    $ctx  = $line.Substring($from, $to-$from)
                    Write-Host ("  {0,-40} L{1,-4} ...{2}..." -f (Split-Path $r.Path -Leaf), $ln, $ctx)
                    $shown++
                }
            }
        }
    }
}
