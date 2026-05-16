# =====================================================================
# fix_umlauts_all.ps1
# ---------------------------------------------------------------------
# Repariert Umlaute in ALLEN VBA-Dateien (.bas / .cls / .frm) und
# behandelt drei Korruptionsarten:
#   1) Lossy "?"          (cp1252 0x3F, durch UTF-8->ASCII Konvertierung)
#   2) Mojibake "Ã¤" etc. (cp1252 als UTF-8 fehlinterpretiert)
#   3) Replacement "ï¿½"  (UTF-8 U+FFFD, dreifach Korruption)
# Ausschluss: BackUp*-Verzeichnisse.
# Ergebnis wird IMMER als cp1252 (Windows-1252) zurueckgeschrieben,
# damit Excel/VBE die Umlaute korrekt liest.
# =====================================================================

$ErrorActionPreference = "Stop"
Set-Location -Path (Split-Path -Parent $PSScriptRoot)

$cp     = [System.Text.Encoding]::GetEncoding(1252)
$ae=[char]228; $oe=[char]246; $ue=[char]252
$AE=[char]196; $OE=[char]214; $UE=[char]220
$ss=[char]223
$drei=[char]179   # cube ^3
$repl=[char]0xFFFD  # Unicode replacement char (single-char form)

# ---------------------------------------------------------------------
# Stufe A: Mojibake-Rekonstruktion (UTF-8-Sequenzen die als cp1252
#          gelesen wurden). MUSS vor "?"-Regeln laufen.
# Patterns werden hier aus Char-Codes gebaut, damit das Skript selbst
# in beliebigem Encoding gespeichert sein kann ohne dass die Patterns
# verfaelscht werden.
# ---------------------------------------------------------------------
$C3 = [char]0xC3; $C2 = [char]0xC2
$mojibake = @(
    @(($C3 + [char]0x84),  $AE),    # UE -> AE
    @(($C3 + [char]0x96),  $OE),
    @(($C3 + [char]0x9C),  $UE),
    @(($C3 + [char]0x9F),  $ss),
    @(($C3 + [char]0xA4),  $ae),
    @(($C3 + [char]0xB6),  $oe),
    @(($C3 + [char]0xBC),  $ue),
    @(($C2 + [char]0xB3),  $drei),
    @(($C2 + [char]0xB2),  [char]178),
    @(($C2 + [char]0xB0),  [char]176),
    @(($C2 + [char]0xA7),  [char]167),
    @(($C2 + [char]0xB4),  [char]180),
    @(($C2 + [char]0xA8),  [char]168)
)

# ---------------------------------------------------------------------
# Stufe B: Wort-Kontext-Regeln fuer "?" Korruption.
#          Reihenfolge: LAENGSTE Patterns zuerst, sonst werden
#          spezifische Treffer von generischen ueberdeckt.
# ---------------------------------------------------------------------
$rules = @(
    # ===== Sehr lange / spezifische Patterns ZUERST =====
    @('Zusammenf?hrung',"Zusammenf${ue}hrung"),
    @('zusammenf?hr',   "zusammenf${ue}hr"),
    @('Kompatibilit?t', "Kompatibilit${ae}t"),
    @('Gesch?ftslogik', "Gesch${ae}ftslogik"),
    @('Zellf?rbungen',  "Zellf${ae}rbungen"),
    @('Z?HLERWECHSEL',  "Z${AE}HLERWECHSEL"),
    @('F?LLBEREICHE',   "F${UE}LLBEREICHE"),
    @('VERF?GBARE',     "VERF${UE}GBARE"),
    @('Z?hlerhistorie', "Z${ae}hlerhistorie"),
    @('Z?hlerwechsel',  "Z${ae}hlerwechsel"),
    @('Z?hlerstand',    "Z${ae}hlerstand"),
    @('Hauptz?hler',    "Hauptz${ae}hler"),
    @('TEMPOR?RE',      "TEMPOR${AE}RE"),
    @('Tempor?rer',     "Tempor${ae}rer"),
    @('Tempor?res',     "Tempor${ae}res"),
    @('tempor?ren',     "tempor${ae}ren"),
    @('tempor?r',       "tempor${ae}r"),
    @('Bef?llung',      "Bef${ue}llung"),
    @('hinzugef?gt',    "hinzugef${ue}gt"),
    @('hinzuf?gen',     "hinzuf${ue}gen"),
    @('?berschrieben',  "${ue}berschrieben"),
    @('?berschreiben',  "${ue}berschreiben"),
    @('?berspringen',   "${ue}berspringen"),
    @('?bersprungen',   "${ue}bersprungen"),
    @('?bersichtsblatt',"${UE}bersichtsblatt"),
    @('Priorit?ten',    "Priorit${ae}ten"),
    @('Priorit?t',      "Priorit${ae}t"),
    @('Nachp?chter',    "Nachp${ae}chter"),
    @('Datens?tze',     "Datens${ae}tze"),
    @('vollst?ndig',    "vollst${ae}ndig"),
    @('eingef?rbt',     "eingef${ae}rbt"),
    @('Hellgr?n',       "Hellgr${ue}n"),
    @('ausgef?hrt',     "ausgef${ue}hrt"),
    @('Bin?rdaten',     "Bin${ae}rdaten"),
    @('tats?chlich',    "tats${ae}chlich"),
    @('p?nktlich',      "p${ue}nktlich"),
    @('P?nktlich',      "P${ue}nktlich"),
    @('zur?ckgesetzt',  "zur${ue}ckgesetzt"),
    @('zur?cksetzen',   "zur${ue}cksetzen"),
    @('aufger?umt',     "aufger${ae}umt"),
    @('aufr?umen',      "aufr${ae}umen"),
    @('unabh?ngig',     "unabh${ae}ngig"),
    @('abh?ngig',       "abh${ae}ngig"),
    @('Abh?ngig',       "Abh${ae}ngig"),
    @('fallunabh?ngig', "fallunabh${ae}ngig"),
    @('systemunabh?ngig',"systemunabh${ae}ngig"),
    @('K?hltruhe',      "K${ue}hltruhe"),
    @('F?lligkeitstypen',"F${ae}lligkeitstypen"),
    @('F?lligkeitsdatum',"F${ae}lligkeitsdatum"),
    @('F?lligkeit',     "F${ae}lligkeit"),
    @('f?lligkeit',     "f${ae}lligkeit"),
    @('best?tigt',      "best${ae}tigt"),
    @('gesch?tzt',      "gesch${ue}tzt"),
    @('Pr?fungen',      "Pr${ue}fungen"),
    @('Pr?fung',        "Pr${ue}fung"),
    @('PR?FUNG',        "PR${UE}FUNG"),
    @('pr?fung',        "pr${ue}fung"),
    @('Regul?re',       "Regul${ae}re"),
    @('regul?r',        "regul${ae}r"),
    @('S?umnis',        "S${ae}umnis"),

    # ===== Mittlere Laenge =====
    @('?bersicht',      "${UE}bersicht"),
    @('?nderung',       "${AE}nderung"),
    @('ge?ndert',       "ge${ae}ndert"),
    @('ge?nderten',     "ge${ae}nderten"),
    @('?ndern',         "${ae}ndern"),
    @('?ndere',         "${ae}ndere"),
    @('?ndert',         "${ae}ndert"),
    @('?ndere',         "${ae}ndere"),
    @('l?schen',        "l${oe}schen"),
    @('L?schen',        "L${oe}schen"),
    @('gel?scht',       "gel${oe}scht"),
    @('L?sche',         "L${oe}sche"),
    @('l?sche',         "l${oe}sche"),
    @('enth?lt',        "enth${ae}lt"),
    @('R?ckgabe',       "R${ue}ckgabe"),
    @('zur?ck',         "zur${ue}ck"),
    @('K?rzel',         "K${ue}rzel"),
    @('j?hrlich',       "j${ae}hrlich"),
    @('einf?gen',       "einf${ue}gen"),
    @('eingef?gt',      "eingef${ue}gt"),
    @('Geb?hren',       "Geb${ue}hren"),
    @('Geb?hr',         "Geb${ue}hr"),
    @('Eintr?gen',      "Eintr${ae}gen"),
    @('Eintr?ge',       "Eintr${ae}ge"),
    @('EINTR?GE',       "EINTR${AE}GE"),
    @('Betr?ge',        "Betr${ae}ge"),
    @('n?chsten',       "n${ae}chsten"),
    @('n?chste',        "n${ae}chste"),
    @('N?CHSTE',        "N${AE}CHSTE"),
    @('N?chst',         "N${ae}chst"),
    @('n?chst',         "n${ae}chst"),
    @('Bl?tter',        "Bl${ae}tter"),
    @('Bl?cke',         "Bl${oe}cke"),
    @('P?chter',        "P${ae}chter"),
    @('p?chter',        "p${ae}chter"),
    @('L?nge',          "L${ae}nge"),
    @('ausl?sen',       "ausl${oe}sen"),
    @('Ausl?sen',       "Ausl${oe}sen"),
    @('ausgel?st',      "ausgel${oe}st"),
    @('Bef?llt',        "Bef${ue}llt"),
    @('bef?llt',        "bef${ue}llt"),
    @('bef?llen',       "bef${ue}llen"),
    @('Z?hler',         "Z${ae}hler"),
    @('Z?HLER',         "Z${AE}HLER"),
    @('W?rter',         "W${oe}rter"),
    @('W?rtern',        "W${oe}rtern"),
    @('W?hrung',        "W${ae}hrung"),
    @('Pr?fen',         "Pr${ue}fen"),
    @('pr?fen',         "pr${ue}fen"),
    @('Pr?fe',          "Pr${ue}fe"),
    @('pr?ft',          "pr${ue}ft"),
    @('Pr?ft',          "Pr${ue}ft"),
    @('?ffnen',         "${oe}ffnen"),
    @('k?nnen',         "k${oe}nnen"),
    @('K?nnen',         "K${oe}nnen"),
    @('gef?rbt',        "gef${ae}rbt"),
    @('zellf?rb',       "zellf${ae}rb"),
    @('f?rben',         "f${ae}rben"),
    @('M?rz',           "M${ae}rz"),
    @('F?hrt',          "F${ue}hrt"),
    @('f?hrt',          "f${ue}hrt"),
    @('F?hr',           "F${ue}hr"),
    @('f?hr',           "f${ue}hr"),
    @('F?ge',           "F${ue}ge"),
    @('F?gen',          "F${ue}gen"),
    @('R?CKW',          "R${UE}CKW"),
    @('verf?gbar',      "verf${ue}gbar"),
    @('ben?tig',        "ben${oe}tig"),
    @('n?tige',         "n${oe}tige"),
    @('N?tig',          "N${oe}tig"),
    @('Gesch?ft',       "Gesch${ae}ft"),
    @('gesch?ft',       "gesch${ae}ft"),
    @('K?hl',           "K${ue}hl"),
    @('k?hl',           "k${ue}hl"),
    @('GR?N',           "GR${UE}N"),
    @('Gr?n',           "Gr${ue}n"),
    @('gr?ner',         "gr${ue}ner"),
    @('Bin?r',          "Bin${ae}r"),
    @('bin?r',          "bin${ae}r"),
    @('anh?ngen',       "anh${ae}ngen"),
    @('Anh?ngen',       "Anh${ae}ngen"),
    @('unn?tig',        "unn${oe}tig"),
    @('F?lle',          "F${ae}lle"),
    @('F?llen',         "F${ue}llen"),
    @('F?llt',          "F${ae}llt"),
    @('f?llig',         "f${ae}llig"),
    @('F?llig',         "F${ae}llig"),
    @('F?LLIG',         "F${AE}LLIG"),
    @('aufgef?llt',     "aufgef${ue}llt"),
    @('Gef?llt',        "Gef${ue}llt"),
    @('gef?llt',        "gef${ue}llt"),

    # ===== Kurze Patterns am Ende =====
    @('F?R',            "F${UE}R"),
    @('f?r',            "f${ue}r"),
    @('F?r',            "F${ue}r"),
    @('?ber',           "${ue}ber"),  # Achtung: "Über" am Satzanfang als ?ber
    @(' m?',            " m${drei}"),
    @('(m?)',           "(m${drei})"),
    @('m?]',            "m${drei}]"),
    @('m? ',            "m${drei} "),

    # ===== Zweite Welle (aus Diagnose) =====
    @('r?nglichen',     "r${ue}nglichen"),    # urspruenglichen
    @('g?ltigen',       "g${ue}ltigen"),
    @('g?ltiges',       "g${ue}ltiges"),
    @('g?ltige',        "g${ue}ltige"),
    @('g?ltig',         "g${ue}ltig"),
    @('s?tzliche',      "s${ae}tzliche"),     # zusaetzliche
    @('s?tzlich',       "s${ae}tzlich"),
    @('M?chten',        "M${oe}chten"),
    @('m?chten',        "m${oe}chten"),
    @('t?tspr',         "t${ae}tspr"),        # Prioritaetspruefung
    @('K?ndigung',      "K${ue}ndigung"),
    @('L?schung',       "L${oe}schung"),
    @('L?SCHEN',        "L${OE}SCHEN"),
    @('k?nnten',        "k${oe}nnten"),
    @('h?herer',        "h${oe}herer"),
    @('h?here',         "h${oe}here"),
    @('H?he',           "H${oe}he"),
    @('H?TZT',          "H${UE}TZT"),         # GESCHUETZT
    @('R?SSER',         "R${OE}SSER"),        # GROESSER
    @('f?gung',         "f${ue}gung"),
    @('b?ndig',         "b${ue}ndig"),
    @('w?hlen',         "w${ae}hlen"),
    @('w?hlt',          "w${ae}hlt"),
    @('W?hle',          "W${ae}hle"),
    @('w?rde',          "w${ue}rde"),
    @('d?rfen',         "d${ue}rfen"),
    @('r?fix',          "r${ae}fix"),         # Praefix
    @('l?ssige',        "l${ae}ssige"),       # zuverlaessige
    @('l?ssig',         "l${ae}ssig"),
    @('h?tzen',         "h${ue}tzen"),        # schuetzen
    @('h?tze',          "h${ue}tze"),
    @('t?nde',          "t${ae}nde"),         # Bestaende, Zustaende
    @('t?tzt',          "t${ue}tzt"),         # gestuetzt, unterstuetzt
    @('T?T',            "T${AE}T"),           # PRIORITAET
    @('R?GE',           "R${AE}GE"),          # BETRAEGE
    @('E?NDERT',        "E${AE}NDERT"),       # GEAENDERT
    @('h?ngend',        "h${ae}ngend"),
    @('h?lt',           "h${ae}lt"),
    @('R?FEN',          "R${UE}FEN"),         # PRUEFEN
    @('PR?F',           "PR${UE}F"),

    # ===== Dritte Welle (aus Diagnose) =====
    @('R?ckerstattung', "R${ue}ckerstattung"),
    @('R?ckzahlung',    "R${ue}ckzahlung"),
    @('H?ufigstes',     "H${ae}ufigstes"),
    @('S?TZLICHE',      "S${AE}TZLICHE"),     # ZUSAETZLICHE
    @('p?testens',      "p${ae}testens"),     # spaetestens
    @('h?chster',       "h${oe}chster"),
    @('h?chste',        "h${oe}chste"),
    @('j?ngstem',       "j${ue}ngstem"),
    @('j?ngste',        "j${ue}ngste"),
    @('W?RTS',          "W${AE}RTS"),         # RUECKWAERTS
    @('r?ckw',          "r${ue}ckw"),
    @('r?glich',        "r${ae}glich"),       # vertraeglich
    @('r?umen',         "r${ae}umen"),
    @('r?cken',         "r${ue}cken"),
    @('r?fe',           "r${ue}fe"),          # pruefe
    @('z?hlen',         "z${ae}hlen"),
    @('Z?hle',          "Z${ae}hle"),
    @('Z?hlt',          "Z${ae}hlt"),
    @('z?hlt',          "z${ae}hlt"),
    @('t?tigen',        "t${ae}tigen"),
    @('t?tige',         "t${ae}tige"),
    @('l?sung',         "l${oe}sung"),
    @('L?sung',         "L${oe}sung"),
    @('L?sch',          "L${oe}sch"),
    @('L?scht',         "L${oe}scht"),
    @('l?sst',          "l${ae}sst"),
    @('l?st',           "l${oe}st"),          # loest
    @('f?llen',         "f${ue}llen"),        # fuellen
    @('f?ge',           "f${ue}ge"),
    @('f?gbare',        "f${ue}gbare"),
    @('h?rige',         "h${oe}rige"),
    @('F?llung',        "F${ue}llung"),
    @('p?ter',          "p${ae}ter"),         # spaeter
    @('h?he',           "h${oe}he"),
    @('h?hter',         "h${oe}hter"),        # erhoehter
    @('r?n',            "r${ue}n"),           # gruen
    @('r?ren',          "r${ue}ren"),         # ggf. beruehren
    @('n?tig',          "n${oe}tig"),

    # ===== Vierte Welle (Restbestaende) =====
    @('Flie?komma',     "Flie${ss}komma"),
    @('st?rkere',       "st${ae}rkere"),
    @('st?rker',        "st${ae}rker"),
    @('sch?tzt',        "sch${ue}tzt"),
    @('Sch?tzt',        "Sch${ue}tzt"),
    @('Sch?tze',        "Sch${ue}tze"),
    @('sch?tze',        "sch${ue}tze"),
    @('Sch?tzen',       "Sch${ue}tzen"),
    @('ungesch?tzt',    "ungesch${ue}tzt"),
    @('erh?hten',       "erh${oe}hten"),
    @('erh?hte',        "erh${oe}hte"),
    @('R?ckkehr',       "R${ue}ckkehr"),
    @('abschlie?end',   "abschlie${ss}end"),
    @('auszuschlie?end',"auszuschlie${ss}end"),
    @('hei?t',          "hei${ss}t"),
    @('gedr?ckt',       "gedr${ue}ckt"),
    @('Bankgeb?hren',   "Bankgeb${ue}hren"),
    @('Qualit?ts',      "Qualit${ae}ts"),
    @('Priorit?ts',     "Priorit${ae}ts"),
    @('ZUS?TZLICHE',    "ZUS${AE}TZLICHE"),
    @('l?tter',         "l${ae}tter"),        # Blaetter (klein)
    @('w?rts',          "w${ae}rts")          # rueckwaerts
)

# ---------------------------------------------------------------------
# Sammeln aller Ziel-Dateien
# ---------------------------------------------------------------------
$files = @()
$files += Get-ChildItem -Path 'vba' -Recurse -Include *.bas,*.cls,*.frm |
          Where-Object { $_.FullName -notmatch '\\BackUp' }

Write-Host "Zu verarbeitende Dateien: $($files.Count)" -ForegroundColor Cyan

$grandTotal = 0
$fileTouched = 0

foreach ($f in $files) {
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
    $txt   = $cp.GetString($bytes)
    $orig  = $txt
    $reps  = 0

    # Stufe 0: U+FFFD (als Einzelchar nach cp1252-Read) -> "?"
    # (entsteht wenn Datei UTF-8 war; "ï¿½" Drei-Byte-Sequenz wird
    # von .NET cp1252-Decoder als ein U+FFFD geliefert).
    if ($txt.Contains($repl)) {
        $cnt = ([regex]::Matches($txt, [regex]::Escape($repl))).Count
        $txt = $txt.Replace($repl, '?')
        $reps += $cnt
    }
    # Zusaetzlich die seltene Drei-Char-Form (wenn jemand das
    # bereits einmal als cp1252 gespeichert hat).
    $three = [char]0xEF + [char]0xBF + [char]0xBD
    if ($txt.Contains($three)) {
        $cnt = ([regex]::Matches($txt, [regex]::Escape($three))).Count
        $txt = $txt.Replace($three, '?')
        $reps += $cnt
    }

    # Stufe A: Mojibake reparieren
    foreach ($r in $mojibake) {
        if ($txt.Contains($r[0])) {
            $cnt = ([regex]::Matches($txt, [regex]::Escape($r[0]))).Count
            $txt = $txt.Replace($r[0], $r[1])
            $reps += $cnt
        }
    }

    # Stufe B: Wort-Kontext Regeln fuer "?"
    foreach ($r in $rules) {
        if ($txt.Contains($r[0])) {
            $cnt = ([regex]::Matches($txt, [regex]::Escape($r[0]))).Count
            $txt = $txt.Replace($r[0], $r[1])
            $reps += $cnt
        }
    }

    if ($txt -ne $orig) {
        [System.IO.File]::WriteAllText($f.FullName, $txt, $cp)
        Write-Host ("  {0,-50}  +{1}" -f $f.Name, $reps)
        $grandTotal += $reps
        $fileTouched++
    }
}

Write-Host ""
Write-Host "===========================================" -ForegroundColor Green
Write-Host ("Geaenderte Dateien : {0}" -f $fileTouched) -ForegroundColor Green
Write-Host ("Ersetzungen total  : {0}" -f $grandTotal) -ForegroundColor Green
Write-Host "===========================================" -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------------------
# Diagnose: verbleibende "?"-Worte (echte Fragezeichen ausgeschlossen
# bleiben - das sind die mit ? am Ende eines ganzen Wortes/Satzes)
# ---------------------------------------------------------------------
Write-Host "Verbleibende verdaechtige '?'-Patterns (Top 40):" -ForegroundColor Yellow
$patterns = @{}
foreach ($f in $files) {
    $txt = [System.IO.File]::ReadAllText($f.FullName, $cp)
    # nur Patterns mit Buchstaben VOR und NACH dem ? (echtes Wort-? mitten drin)
    foreach ($m in [regex]::Matches($txt, '[A-Za-z]\?[A-Za-z]+')) {
        $w = $m.Value
        if (-not $patterns.ContainsKey($w)) { $patterns[$w] = 0 }
        $patterns[$w]++
    }
}
$patterns.GetEnumerator() | Sort-Object Value -Descending |
    Select-Object -First 40 | Format-Table -AutoSize
