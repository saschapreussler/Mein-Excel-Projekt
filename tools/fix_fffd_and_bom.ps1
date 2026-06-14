param(
    [Parameter(Mandatory=$true)]
    [string[]]$Files,
    [switch]$DryRun
)

# Replacement-Char
$FFFD = [char]0xFFFD
$ufffd = [string]$FFFD

# Dictionary: kontextspezifische Wortmuster mit FFFD -> Umlaut-Wort
# Wichtig: Reihenfolge zaehlt (laengste zuerst)
$wortMap = @(
    @('Pr{F}fung',           'Pruefung'),
    @('pr{F}ft',              'prueft'),
    @('Pr{F}fe',              'Pruefe'),
    @('pr{F}fen',             'pruefen'),
    @('m{F}ssen',             'muessen'),
    @('schlie{F}en',          'schliessen'),
    @('Schlie{F}en',          'Schliessen'),
    @('hei{F}t',              'heisst'),
    @('ge{F}ndert',           'geaendert'),
    @('ge{F}ffnet',           'geoeffnet'),
    @('{F}ndert',             'aendert'),
    @('{F}nderung',           'Aenderung'),
    @('{F}berschrieben',      'ueberschrieben'),
    @('{F}bersprungen',       'uebersprungen'),
    @('{F}berspringe',        'ueberspringe'),
    @('{F}bernommen',         'uebernommen'),
    @('{F}bersicht',          'Uebersicht'),
    @('{F}r ',                'fuer '),
    @('f{F}r ',               'fuer '),
    @('F{F}r ',               'Fuer '),
    @('w{F}rde',              'wuerde'),
    @('w{F}rden',             'wuerden'),
    @('k{F}nnen',             'koennen'),
    @('k{F}nnte',             'koennte'),
    @('k{F}nftig',            'kuenftig'),
    @('Tempor{F}r',           'Temporaer'),
    @('tempor{F}r',           'temporaer'),
    @('Z{F}hler',             'Zaehler'),
    @('z{F}hlt',              'zaehlt'),
    @('z{F}hlen',             'zaehlen'),
    @('aufr{F}umen',          'aufraeumen'),
    @('verkn{F}pft',          'verknuepft'),
    @('Regul{F}re',           'Regulaere'),
    @('regul{F}re',           'regulaere'),
    @('L{F}schen',            'Loeschen'),
    @('l{F}schen',            'loeschen'),
    @('F{F}llung',            'Fuellung'),
    @('F{F}ll',               'Fuell'),
    @('n{F}tig',              'noetig'),
    @('m{F}glich',            'moeglich'),
    @('m{F}glichkeit',        'Moeglichkeit'),
    @('zerst{F}rt',           'zerstoert'),
    @('zerst{F}ren',          'zerstoeren'),
    @('Unterst{F}tzte',       'Unterstuetzte'),
    @('unterst{F}tzt',        'unterstuetzt'),
    @('robuste, fallunabh{F}ngige', 'robuste, fallunabhaengige'),
    @('fallunabh{F}ngige',    'fallunabhaengige'),
    @('enth{F}lt',            'enthaelt'),
    @('z{F}sse',              'zuesse'),
    @('Fehlerschl{F}sse',     'Fehlerschluesse'),
    @('R{F}ckkehr',           'Rueckkehr'),
    @('R{F}ckbuchung',        'Rueckbuchung'),
    @('zust{F}ndig',          'zustaendig'),
    @('zust{F}ndige',         'zustaendige'),
    @('Stabilit{F}t',         'Stabilitaet'),
    @('STABILIT{F}T',         'STABILITAET'),
    @('Bl{F}tter',            'Blaetter'),
    @('Bl{F}ttern',           'Blaettern'),
    @('Bl{F}ttern',           'Blaettern'),
    @('hierf{F}r',            'hierfuer'),
    @('daf{F}r',              'dafuer'),
    @('da{F}r',               'dafuer'),
    @('verf{F}gbar',          'verfuegbar'),
    @('Hinzuf{F}gen',         'Hinzufuegen'),
    @('hinzuf{F}gen',         'hinzufuegen'),
    @('ausf{F}hren',          'ausfuehren'),
    @('Ausf{F}hrung',         'Ausfuehrung'),
    @('durchgef{F}hrt',       'durchgefuehrt'),
    @('aufger{F}umt',         'aufgeraeumt'),
    @('w{F}hrend',            'waehrend'),
    @('Diff{F}',              'Diff'),
    @('verz{F}gert',          'verzoegert'),
    @('ausgef{F}hrt',         'ausgefuehrt'),
    @('einf{F}gen',           'einfuegen'),
    @('zur{F}ck',             'zurueck'),
    @('R{F}ckgabe',           'Rueckgabe'),
    @('L{F}schung',           'Loeschung'),
    @('Bin{F}rdaten',         'Binaerdaten'),
    @('zugeh{F}rige',         'zugehoerige'),
    @('{F}berschrieben',      'ueberschrieben'),
    @('{F}berschreiben',      'ueberschreiben'),
    @('EXISTIERT {F} Code',   'EXISTIERT -> Code'),
    @('EXISTIERT NICHT {F} Neu', 'EXISTIERT NICHT -> Neu'),
    @('EXISTIERT {F}',        'EXISTIERT ->'),
    @('Umlaute \({F}, {F}, {F}, {F}\)', 'Umlaute (ae, oe, ue, ss)'),
    # Zweite Runde: Top 40 Wort-Pattern
    @('gr{F}n',               'gruen'),
    @('Gr{F}n',               'Gruen'),
    @('Nachp{F}chter',        'Nachpaechter'),
    @('Nachp{F}chterin',      'Nachpaechterin'),
    @('{F}ber',               'ueber'),
    @('{F}bergabe',           'Uebergabe'),
    @('S{F}umnis',            'Saeumnis'),
    @('j{F}hrlich',           'jaehrlich'),
    @('K{F}rzel',             'Kuerzel'),
    @('W{F}rter',             'Woerter'),
    @('vollst{F}ndig',        'vollstaendig'),
    @('Vollst{F}ndig',        'Vollstaendig'),
    @('gel{F}scht',           'geloescht'),
    @('Geb{F}hr',             'Gebuehr'),
    @('Geb{F}hren',           'Gebuehren'),
    @('N{F}CHSTE',            'NAECHSTE'),
    @('n{F}chste',            'naechste'),
    @('N{F}chste',            'Naechste'),
    @('Datens{F}tze',         'Datensaetze'),
    @('w{F}hlen',             'waehlen'),
    @('W{F}hlen',             'Waehlen'),
    @('w{F}hrend',            'waehrend'),
    @('W{F}hrung',            'Waehrung'),
    @('tats{F}chlich',        'tatsaechlich'),
    @('gesch{F}tzt',          'geschuetzt'),
    @('sch{F}tzen',           'schuetzen'),
    @('Betr{F}ge',            'Betraege'),
    @('ausl{F}sen',           'ausloesen'),
    @('anh{F}ngen',           'anhaengen'),
    @('Plausibilit{F}t',      'Plausibilitaet'),
    @('L{F}nge',              'Laenge'),
    @('abh{F}ngig',           'abhaengig'),
    @('ung{F}ltig',           'ungueltig'),
    @('zul{F}ssig',           'zulaessig'),
    @('unzul{F}ssige',        'unzulaessige'),
    @('Eintr{F}ge',           'Eintraege'),
    @('F{F}hrt',              'Fuehrt'),
    @('f{F}hrt',              'fuehrt'),
    @('ausgew{F}hlt',         'ausgewaehlt'),
    @('best{F}tigt',          'bestaetigt'),
    @('best{F}tigen',         'bestaetigen'),
    @('{F}ffnen',             'oeffnen'),
    @('P{F}chter',            'Paechter'),
    @('K{F}hltruhe',          'Kuehltruhe'),
    @('hellgr{F}n',           'hellgruen'),
    @('L{F}sche',             'Loesche'),
    @('gr{F}ner',             'gruener'),
    @('p{F}nktlich',          'puenktlich'),
    @('M{F}rz',               'Maerz'),
    @('Endg{F}ltig',          'Endgueltig'),
    @('endg{F}ltig',          'endgueltig'),
    @('Verm{F}gen',           'Vermoegen'),
    @('Spr{F}che',            'Sprueche'),
    @('Verg{F}tung',          'Verguetung'),
    @('Aufl{F}sung',          'Aufloesung'),
    @('aufgel{F}st',          'aufgeloest'),
    @('Tabellenbl{F}tter',    'Tabellenblaetter'),
    @('bl{F}tter',            'blaetter'),
    @('Bl{F}tter',            'Blaetter'),
    @('verkn{F}pft',          'verknuepft'),
    @('R{F}cksprung',         'Ruecksprung'),
    @('R{F}ck',               'Rueck'),
    @('Erkl{F}rung',          'Erklaerung'),
    @('Aufkl{F}rung',         'Aufklaerung'),
    @('h{F}ufig',             'haeufig'),
    @('h{F}ngt',              'haengt'),
    @('Sch{F}tzung',          'Schaetzung'),
    @('sp{F}ter',             'spaeter'),
    @('Sp{F}ter',             'Spaeter'),
    @('verz{F}gerung',        'verzoegerung'),
    @('V{F}llig',             'Voellig'),
    @('v{F}llig',             'voellig'),
    @('Anh{F}ngsel',          'Anhaengsel'),
    # Dritte Runde
    @('Pr{F}fix',             'Praefix'),
    @('f{F}rben',             'faerben'),
    @('gef{F}rbt',            'gefaerbt'),
    @('F{F}hre',              'Fuehre'),
    @('hinzugef{F}gt',        'hinzugefuegt'),
    @('Verf{F}gung',          'Verfuegung'),
    @('Sch{F}tze',            'Schuetze'),
    @('sch{F}tzt',            'schuetzt'),
    @('gew{F}hlt',            'gewaehlt'),
    @('GR{F}SSER',            'GROESSER'),
    @('Zusammenf{F}hrung',    'Zusammenfuehrung'),
    @('Startmen{F}',          'Startmenue'),
    @('Qualit{F}tsfaktor',    'Qualitaetsfaktor'),
    @('F{F}gen',              'Fuegen'),
    @('Priorit{F}t',          'Prioritaet'),
    @('L{F}sch',              'Loesch'),
    @('Flie{F}komma',         'Fliesskomma'),
    @('Gro{F}',               'Gross'),
    @('gro{F}',               'gross'),
    @('erh{F}ht',             'erhoeht'),
    @('wei{F}',               'weiss'),
    @('h{F}chst',             'hoechst'),
    @('h{F}her',              'hoeher'),
    @('St{F}nde',             'Staende'),
    @('Zaehlerst{F}nde',      'Zaehlerstaende'),
    @('St{F}nden',            'Staenden'),
    @('{F}ffentlich',         'oeffentlich'),
    @('{F}ffne',              'oeffne'),
    @('FUNKTIONALIT{F}T',     'FUNKTIONALITAET'),
    @('Zeilenh{F}he',         'Zeilenhoehe'),
    @('d{F}rfen',             'duerfen'),
    @('rechtsb{F}ndig',       'rechtsbuendig'),
    @('linksb{F}ndig',        'linksbuendig'),
    @('M{F}chten',            'Moechten'),
    @('L{F}scht',             'Loescht'),
    @('RueckW{F}RTS',         'RueckWAERTS'),
    @('R{F}ckW',              'RueckW'),
    @('zusammenh{F}ngend',    'zusammenhaengend'),
    @('Bl{F}cke',             'Bloecke'),
    @('l{F}st',               'loest'),
    @('Ausl{F}sung',          'Ausloesung'),
    @('zus{F}tzlich',         'zusaetzlich'),
    @('sp{F}testens',         'spaetestens'),
    @('st{F}rkere',           'staerkere'),
    @('unvertr{F}glich',      'unvertraeglich'),
    @('Gesch{F}ftslogik',     'Geschaeftslogik'),
    @('Pr{F}fixe',            'Praefixe'),
    @('m{F}',                 'mae')
)

foreach($f in $Files){
    if(-not (Test-Path $f)){
        Write-Warning "SKIP: $f - nicht gefunden"
        continue
    }
    $b = [System.IO.File]::ReadAllBytes($f)
    $hadBom = $b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF
    if($hadBom){
        $text = [System.Text.Encoding]::UTF8.GetString($b, 3, $b.Length - 3)
    } else {
        $text = [System.Text.Encoding]::UTF8.GetString($b)
    }
    $vorher = ([regex]::Matches($text, [regex]::Escape($ufffd))).Count

    foreach($pair in $wortMap){
        $pattern = $pair[0] -replace '\{F\}', [regex]::Escape($ufffd)
        $text = $text -replace $pattern, $pair[1]
    }

    $nachher = ([regex]::Matches($text, [regex]::Escape($ufffd))).Count

    # Schreiben mit BOM
    $bytes = New-Object byte[] 0
    $bom = [byte[]](0xEF, 0xBB, 0xBF)
    $payload = [System.Text.Encoding]::UTF8.GetBytes($text)
    $bytes = New-Object byte[] ($bom.Length + $payload.Length)
    [Array]::Copy($bom, 0, $bytes, 0, $bom.Length)
    [Array]::Copy($payload, 0, $bytes, $bom.Length, $payload.Length)

    if(-not $DryRun){
        [System.IO.File]::WriteAllBytes($f, $bytes)
    }

    "{0,-46} BOM:{1}->TRUE  FFFD:{2}->{3}" -f $f, $hadBom, $vorher, $nachher
}
