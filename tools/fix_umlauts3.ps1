$ErrorActionPreference = "Stop"
$ae = [char]228; $oe = [char]246; $ue = [char]252
$AE = [char]196; $OE = [char]214; $UE = [char]220
$ss = [char]223

$rules = @(
    @("F?lligkeitstypen", "F${ae}lligkeitstypen"),
    @("F?lligkeitsdatum", "F${ae}lligkeitsdatum"),
    @("F?lligkeit",   "F${ae}lligkeit"),
    @("f?lligkeit",   "f${ae}lligkeit"),
    @("f?llig",       "f${ae}llig"),
    @("F?llig",       "F${ae}llig"),
    @("gef?llt",      "gef${ue}llt"),
    @("bef?llt",      "bef${ue}llt"),
    @("aufgef?llt",   "aufgef${ue}llt"),
    @("Bef?llung",    "Bef${ue}llung"),
    @("bef?llen",     "bef${ue}llen"),
    @("F?llen",       "F${ue}llen"),
    @("F?lle",        "F${ae}lle"),
    @("F?LLBEREICHE", "F${UE}LLBEREICHE"),
    @("F?R",          "F${UE}R"),
    @("W?rter",       "W${oe}rter"),
    @("W?rtern",      "W${oe}rtern"),
    @("Bl?tter",      "Bl${ae}tter"),
    @("Bl?cke",       "Bl${oe}cke"),
    @("N?CHSTE",      "N${AE}CHSTE"),
    @("n?chst",       "n${ae}chst"),
    @("N?chst",       "N${ae}chst"),
    @("Nachp?chter",  "Nachp${ae}chter"),
    @("P?chter",      "P${ae}chter"),
    @("p?chter",      "p${ae}chter"),
    @("abh?ngig",     "abh${ae}ngig"),
    @("Abh?ngig",     "Abh${ae}ngig"),
    @("unabh?ngig",   "unabh${ae}ngig"),
    @("fallunabh?ngige", "fallunabh${ae}ngige"),
    @("systemunabh?ngiges", "systemunabh${ae}ngiges"),
    @("ausl?sen",     "ausl${oe}sen"),
    @("Ausl?sen",     "Ausl${oe}sen"),
    @("ausgel?st",    "ausgel${oe}st"),
    @("TEMPOR?RE",    "TEMPOR${AE}RE"),
    @("Tempor?rer",   "Tempor${ae}rer"),
    @("Tempor?res",   "Tempor${ae}res"),
    @("tempor?ren",   "tempor${ae}ren"),
    @("tempor?r",     "tempor${ae}r"),
    @("Regul?re",     "Regul${ae}re"),
    @("regul?r",      "regul${ae}r"),
    @("F?hrt",        "F${ue}hrt"),
    @("f?hrt",        "f${ue}hrt"),
    @("F?hr",         "F${ue}hr"),
    @("f?hr",         "f${ue}hr"),
    @("unn?tige",     "unn${oe}tige"),
    @("unn?tigen",    "unn${oe}tigen"),
    @("unn?tig",      "unn${oe}tig"),
    @("Bin?rdaten",   "Bin${ae}rdaten"),
    @("bin?r",        "bin${ae}r"),
    @("Zellf?rbungen","Zellf${ae}rbungen"),
    @("zellf?rb",     "zellf${ae}rb"),
    @("eingef?rbt",   "eingef${ae}rbt"),
    @("gef?rbt",      "gef${ae}rbt"),
    @("f?rben",       "f${ae}rben"),
    @("M?rz",         "M${ae}rz"),
    @("Zusammenf?hrung", "Zusammenf${ue}hrung"),
    @("zusammenf?hr", "zusammenf${ue}hr"),
    @("Gesch?ftslogik","Gesch${ae}ftslogik"),
    @("Gesch?ft",     "Gesch${ae}ft"),
    @("gesch?ft",     "gesch${ae}ft"),
    @("p?nktlich",    "p${ue}nktlich"),
    @("P?nktlich",    "P${ue}nktlich"),
    @("n?tig",        "n${oe}tig"),
    @("N?tig",        "N${oe}tig"),
    @("ben?tig",      "ben${oe}tig"),
    @("K?hltruhe",    "K${ue}hltruhe"),
    @("k?hl",         "k${ue}hl"),
    @("K?hl",         "K${ue}hl"),
    @("EINTR?GE",     "EINTR${AE}GE"),
    @("Kompatibilit?t","Kompatibilit${ae}t"),
    @("Priorit?ten",  "Priorit${ae}ten"),
    @("Priorit?t",    "Priorit${ae}t"),
    @("R?CKW",        "R${UE}CKW"),
    @("F?ge",         "F${ue}ge"),
    @("F?gen",        "F${ue}gen"),
    @("VERF?GBARE",   "VERF${UE}GBARE"),
    @("verf?gbar",    "verf${ue}gbar"),
    @("LL?",          "LL${OE}")  # rare placeholder
)

$cp = [System.Text.Encoding]::GetEncoding(1252)
$files = @()
$files += Get-ChildItem "vba\Modules" -Filter "*.bas"
$files += Get-ChildItem "vba\Classes" -Filter "*.cls"

$total = 0
foreach ($f in $files) {
    $txt = [System.IO.File]::ReadAllText($f.FullName, $cp)
    $orig = $txt
    $reps = 0
    foreach ($r in $rules) {
        if ($txt.Contains($r[0])) {
            $idx = 0; $cnt = 0
            while (($idx = $txt.IndexOf($r[0], $idx)) -ge 0) { $cnt++; $idx += $r[0].Length }
            $txt = $txt.Replace($r[0], $r[1])
            $reps += $cnt
        }
    }
    if ($txt -ne $orig) {
        [System.IO.File]::WriteAllText($f.FullName, $txt, $cp)
        $total += $reps
        Write-Host "  $($f.Name): $reps"
    }
}
Write-Host "Total: $total"

Write-Host ""
Write-Host "Verbleibende Patterns:" -ForegroundColor Yellow
$patterns = @{}
foreach ($f in $files) {
    $txt = [System.IO.File]::ReadAllText($f.FullName, $cp)
    foreach ($m in [regex]::Matches($txt, '\b\w*\?\w*\b')) {
        $w = $m.Value
        if (-not $patterns.ContainsKey($w)) { $patterns[$w] = 0 }
        $patterns[$w]++
    }
}
$patterns.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 30 | Format-Table -AutoSize
