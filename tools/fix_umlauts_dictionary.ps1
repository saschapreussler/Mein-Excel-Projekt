# ========================================================================
# fix_umlauts_dictionary.ps1
# Repariert deutsche Umlaute in VBA-Quelldateien (.bas, .cls).
# Strategie: Wort-fuer-Wort-Ersetzung ueber ein explizites Woerterbuch.
# Datei selbst ist ASCII-clean -- Umlaute werden zur Laufzeit aus
# [char]-Codes gebaut, damit es KEINE Encoding-Probleme gibt.
# Geschrieben wird als UTF-8 MIT BOM (VBA-Editor erwartet das).
# ========================================================================

[CmdletBinding()]
param(
    [string]$Root   = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba",
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

# Umlaute als [char]  -- ACHTUNG: PowerShell-Variablen sind case-insensitive,
# darum eindeutige Namen mit Praefix "lo_" (lower) und "up_" (upper)!
$lo_a = [char]228   # klein ae  (ä)
$lo_o = [char]246   # klein oe  (ö)
$lo_u = [char]252   # klein ue  (ü)
$up_A = [char]196   # gross Ae  (Ä)
$up_O = [char]214   # gross Oe  (Ö)
$up_U = [char]220   # gross Ue  (Ü)
$ssch = [char]223   # ss-Ligatur (ß)

# Hilfs-Builder: nimmt einen Stringtemplate-Wert und ersetzt
# Platzhalter durch echte Umlaute.
#   {ae} -> ä klein   {oe} -> ö klein   {ue} -> ü klein
#   {AE} -> Ä gross   {OE} -> Ö gross   {UE} -> Ü gross   {ss} -> ß
function Build([string]$s) {
    $s.Replace('{ae}', $lo_a).Replace('{oe}', $lo_o).Replace('{ue}', $lo_u).
       Replace('{AE}', $up_A).Replace('{OE}', $up_O).Replace('{UE}', $up_U).
       Replace('{ss}', $ssch)
}

# Liste von Paaren als (kaputt, korrekt) - keine Duplikate, lange Wörter zuerst
$pairs = @(
    # ---- spezifische lange Woerter zuerst -------------------------------
    @('Unvollst?ndige',     'Unvollst{ae}ndige'),
    @('vervollst?ndigen',   'vervollst{ae}ndigen'),
    @('vervollst?ndigt',    'vervollst{ae}ndigt'),
    @('Vervollst?ndigung',  'Vervollst{ae}ndigung'),
    @('Mitgliederz?hlung',  'Mitgliederz{ae}hlung'),
    @('Z?hlerwechsel',      'Z{ae}hlerwechsel'),
    @('Z?hlerstand',        'Z{ae}hlerstand'),
    @('Z?hlerst?nde',       'Z{ae}hlerst{ae}nde'),
    @('Z?hlerhistorie',     'Z{ae}hlerhistorie'),
    @('Bankkontoausz?ge',   'Bankkontoausz{ue}ge'),
    @('Kontoausz?ge',       'Kontoausz{ue}ge'),
    @('Kontoausz?gen',      'Kontoausz{ue}gen'),
    @('Vereinskasseneintr?ge', 'Vereinskasseneintr{ae}ge'),
    @('Eintr?ge',           'Eintr{ae}ge'),
    @('Eintr?gen',          'Eintr{ae}gen'),
    @('Datens?tze',         'Datens{ae}tze'),
    @('Datens?tzen',        'Datens{ae}tzen'),

    # ---- ?bersicht / ?berschreiben (lang vor ?ber) ----------------------
    @('?berschreiben',      '{UE}berschreiben'),
    @('?berschrieben',      '{ue}berschrieben'),
    @('?bersichtsblatt',    '{UE}bersichtsblatt'),
    @('?bersichten',        '{UE}bersichten'),
    @('?bersicht',          '{UE}bersicht'),
    @('?berpr?fung',        '{UE}berpr{ue}fung'),
    @('?berpr?fen',         '{UE}berpr{ue}fen'),
    @('?bersetzung',        '{UE}bersetzung'),
    @('?bertragung',        '{UE}bertragung'),
    @('?bertragen',         '{UE}bertragen'),
    @('?bertr?gt',          '{ue}bertr{ae}gt'),
    @('?bertr?ge',          '{UE}bertr{ae}ge'),
    @('?bertrag',           '{UE}bertrag'),
    @('?bernehmen',         '{ue}bernehmen'),
    @('?bernahme',          '{UE}bernahme'),
    @('?bernimmt',          '{ue}bernimmt'),
    @('?bernommen',         '{ue}bernommen'),
    @('?bersprungen',       '{ue}bersprungen'),
    @('?berspringen',       '{ue}berspringen'),
    @('?berspringt',        '{ue}berspringt'),
    @('?berfl?ssig',        '{ue}berfl{ue}ssig'),
    @('?berraschung',       '{UE}berraschung'),
    @('?berf?hren',         '{ue}berf{ue}hren'),
    @('dar?ber',            'dar{ue}ber'),
    @('hier?ber',           'hier{ue}ber'),
    @('?bers',              '{ue}bers'),
    @('?ber',               '{ue}ber'),

    # ---- Pr?fung / Pr?fen ------------------------------------------------
    @('Pr?fungen',          'Pr{ue}fungen'),
    @('Pr?fung',            'Pr{ue}fung'),
    @('pr?fung',            'pr{ue}fung'),
    @('Pr?fen',             'Pr{ue}fen'),
    @('Pr?fer',             'Pr{ue}fer'),
    @('pr?fen',             'pr{ue}fen'),
    @('pr?ft',              'pr{ue}ft'),
    @('pr?fe',              'pr{ue}fe'),
    @('gepr?ft',            'gepr{ue}ft'),
    @('Endpr?fung',         'Endpr{ue}fung'),
    @('Pr?fdatum',          'Pr{ue}fdatum'),

    # ---- S?umnis / Geb?hr ------------------------------------------------
    @('S?umnisgeb?hr',      'S{ae}umnisgeb{ue}hr'),
    @('S?umnis-Geb?hr',     'S{ae}umnis-Geb{ue}hr'),
    @('S?umnis',            'S{ae}umnis'),
    @('Geb?hren',           'Geb{ue}hren'),
    @('Geb?hr',             'Geb{ue}hr'),
    @('geb?hr',             'geb{ue}hr'),
    @('Schl?sselgeb?hr',    'Schl{ue}sselgeb{ue}hr'),
    @('Strafgeb?hr',        'Strafgeb{ue}hr'),

    # ---- p?nktlich -------------------------------------------------------
    @('P?nktlichkeit',      'P{ue}nktlichkeit'),
    @('p?nktliche',         'p{ue}nktliche'),
    @('p?nktlich',          'p{ue}nktlich'),
    @('P?nktlich',          'P{ue}nktlich'),

    # ---- M?chten / m?glich / unm?glich ----------------------------------
    @('M?chten',            'M{oe}chten'),
    @('m?chten',            'm{oe}chten'),
    @('M?glichkeiten',      'M{oe}glichkeiten'),
    @('M?glichkeit',        'M{oe}glichkeit'),
    @('m?glicherweise',     'm{oe}glicherweise'),
    @('M?glich',            'M{oe}glich'),
    @('m?glich',            'm{oe}glich'),
    @('unm?glich',          'unm{oe}glich'),
    @('M?glichkeitsraum',   'M{oe}glichkeitsraum'),

    # ---- l?schen ---------------------------------------------------------
    @('L?schung',           'L{oe}schung'),
    @('L?schen',            'L{oe}schen'),
    @('l?schen',            'l{oe}schen'),
    @('L?scht',             'L{oe}scht'),
    @('l?scht',             'l{oe}scht'),
    @('Gel?scht',           'Gel{oe}scht'),
    @('gel?scht',           'gel{oe}scht'),
    @('l?schvorgang',       'l{oe}schvorgang'),

    # ---- ?ffnen ----------------------------------------------------------
    @('?ffentlich',         '{oe}ffentlich'),
    @('?ffnung',            '{OE}ffnung'),
    @('?ffnen',             '{OE}ffnen'),
    @('?ffnet',             '{oe}ffnet'),
    @('Ge?ffnet',           'Ge{oe}ffnet'),
    @('ge?ffnet',           'ge{oe}ffnet'),

    # ---- ?ndern ----------------------------------------------------------
    @('Unver?ndert',        'Unver{ae}ndert'),
    @('unver?ndert',        'unver{ae}ndert'),
    @('?nderungen',         '{AE}nderungen'),
    @('?nderung',           '{AE}nderung'),
    @('?ndern',             '{AE}ndern'),
    @('?ndere',             '{ae}ndere'),
    @('?ndert',             '{ae}ndert'),
    @('Ge?ndert',           'Ge{ae}ndert'),
    @('ge?ndertes',         'ge{ae}ndertes'),
    @('ge?nderte',          'ge{ae}nderte'),
    @('ge?ndert',           'ge{ae}ndert'),

    # ---- P?chter / Verp?chter -------------------------------------------
    @('Verp?chter',         'Verp{ae}chter'),
    @('P?chterin',          'P{ae}chterin'),
    @('P?chtern',           'P{ae}chtern'),
    @('P?chter',            'P{ae}chter'),

    # ---- f?r -------------------------------------------------------------
    @('hierf?r',            'hierf{ue}r'),
    @('daf?r',              'daf{ue}r'),
    @('wof?r',              'wof{ue}r'),
    @('f?r',                'f{ue}r'),
    @('F?r',                'F{ue}r'),

    # ---- ?blich ----------------------------------------------------------
    @('?blicherweise',      '{ue}blicherweise'),
    @('?blich',             '{ue}blich'),

    # ---- k?nnen ----------------------------------------------------------
    @('K?nnen',             'K{oe}nnen'),
    @('k?nnen',             'k{oe}nnen'),
    @('K?nnte',             'K{oe}nnte'),
    @('k?nnte',             'k{oe}nnte'),
    @('k?nnten',            'k{oe}nnten'),

    # ---- m?ssen ----------------------------------------------------------
    @('M?ssen',             'M{ue}ssen'),
    @('m?ssen',             'm{ue}ssen'),
    @('M?sste',             'M{ue}sste'),
    @('m?sste',             'm{ue}sste'),
    @('m?ssten',            'm{ue}ssten'),

    # ---- w?hlen / w?hrend / W?hrung -------------------------------------
    @('ausgew?hlt',         'ausgew{ae}hlt'),
    @('Ausgew?hlt',         'Ausgew{ae}hlt'),
    @('ausw?hlen',          'ausw{ae}hlen'),
    @('Ausw?hlen',          'Ausw{ae}hlen'),
    @('W?hlen',             'W{ae}hlen'),
    @('w?hlen',             'w{ae}hlen'),
    @('W?hle',              'W{ae}hle'),
    @('w?hle',              'w{ae}hle'),
    @('W?hrungs',           'W{ae}hrungs'),
    @('W?hrung',            'W{ae}hrung'),
    @('W?hrend',            'W{ae}hrend'),
    @('w?hrend',            'w{ae}hrend'),

    # ---- h?her / H?he / erh?ht ------------------------------------------
    @('Erh?hung',           'Erh{oe}hung'),
    @('erh?hter',           'erh{oe}hter'),
    @('erh?hen',            'erh{oe}hen'),
    @('erh?ht',             'erh{oe}ht'),
    @('H?chste',            'H{oe}chste'),
    @('h?chste',            'h{oe}chste'),
    @('H?her',              'H{oe}her'),
    @('h?her',              'h{oe}her'),
    @('H?he',               'H{oe}he'),

    # ---- w?re / h?tte ----------------------------------------------------
    @('W?re',               'W{ae}re'),
    @('w?re',               'w{ae}re'),
    @('w?ren',              'w{ae}ren'),
    @('H?tte',              'H{ae}tte'),
    @('h?tte',              'h{ae}tte'),
    @('h?tten',             'h{ae}tten'),

    # ---- f?llen / f?llig -------------------------------------------------
    @('F?lligkeitstermin',  'F{ae}lligkeitstermin'),
    @('F?lligkeit',         'F{ae}lligkeit'),
    @('f?lligkeit',         'f{ae}lligkeit'),
    @('F?llig',             'F{ae}llig'),
    @('f?llig',             'f{ae}llig'),
    @('Gef?llt',            'Gef{ue}llt'),
    @('gef?llt',            'gef{ue}llt'),
    @('F?llen',             'F{ue}llen'),
    @('f?llen',             'f{ue}llen'),
    @('f?llt',              'f{ae}llt'),
    @('Ausf?llen',          'Ausf{ue}llen'),
    @('ausf?llen',          'ausf{ue}llen'),
    @('ausgef?llt',         'ausgef{ue}llt'),

    # ---- t?glich ---------------------------------------------------------
    @('t?gliche',           't{ae}gliche'),
    @('t?glich',            't{ae}glich'),
    @('T?glich',            'T{ae}glich'),

    # ---- fr?her / sp?ter ------------------------------------------------
    @('fr?hzeitig',         'fr{ue}hzeitig'),
    @('fr?her',             'fr{ue}her'),
    @('Fr?her',             'Fr{ue}her'),
    @('fr?h',               'fr{ue}h'),
    @('sp?ter',             'sp{ae}ter'),
    @('Sp?ter',             'Sp{ae}ter'),

    # ---- ungef?hr / verf?gbar -------------------------------------------
    @('ungef?hre',          'ungef{ae}hre'),
    @('ungef?hr',           'ungef{ae}hr'),
    @('Ungef?hr',           'Ungef{ae}hr'),
    @('verf?gbar',          'verf{ue}gbar'),
    @('Verf?gbar',          'Verf{ue}gbar'),
    @('verf?gt',            'verf{ue}gt'),

    # ---- tempor?r --------------------------------------------------------
    @('tempor?res',         'tempor{ae}res'),
    @('tempor?re',          'tempor{ae}re'),
    @('tempor?r',           'tempor{ae}r'),
    @('Tempor?r',           'Tempor{ae}r'),

    # ---- gest?rt / St?rung ----------------------------------------------
    @('gest?rt',            'gest{oe}rt'),
    @('St?rung',            'St{oe}rung'),
    @('st?ren',             'st{oe}ren'),

    # ---- geh?rt ----------------------------------------------------------
    @('geh?rt',             'geh{oe}rt'),
    @('Geh?rt',             'Geh{oe}rt'),

    # ---- ?lter / k?rzel --------------------------------------------------
    @('?lteste',            '{ae}lteste'),
    @('?ltere',             '{ae}ltere'),
    @('?lter',              '{ae}lter'),
    @('K?rzung',            'K{ue}rzung'),
    @('K?rzel',             'K{ue}rzel'),
    @('k?rzel',             'k{ue}rzel'),
    @('k?rzen',             'k{ue}rzen'),
    @('k?rzlich',           'k{ue}rzlich'),

    # ---- gr??er / Gr??e --------------------------------------------------
    @('Gr??ere',            'Gr{oe}{ss}ere'),
    @('gr??ere',            'gr{oe}{ss}ere'),
    @('Gr??er',             'Gr{oe}{ss}er'),
    @('gr??er',             'gr{oe}{ss}er'),
    @('gr??te',             'gr{oe}{ss}te'),
    @('Gr??te',             'Gr{oe}{ss}te'),
    @('Gr??en',             'Gr{oe}{ss}en'),
    @('Gr??e',              'Gr{oe}{ss}e'),
    @('gr??e',              'gr{oe}{ss}e'),

    # ---- gr?n / GR?N -----------------------------------------------------
    @('Hellgr?n',           'Hellgr{ue}n'),
    @('hellgr?n',           'hellgr{ue}n'),
    @('Dunkelgr?n',         'Dunkelgr{ue}n'),
    @('Gr?n',               'Gr{ue}n'),
    @('gr?n',               'gr{ue}n'),
    @('GR?N',               'GR{UE}N'),

    # ---- Schl?ssel -------------------------------------------------------
    @('Schl?ssel',          'Schl{ue}ssel'),
    @('schl?ssel',          'schl{ue}ssel'),

    # ---- weitere ---------------------------------------------------------
    @('h?ufigkeit',         'h{ae}ufigkeit'),
    @('h?ufig',             'h{ae}ufig'),
    @('H?ufig',             'H{ae}ufig'),
    @('n?chste',            'n{ae}chste'),
    @('N?chste',            'N{ae}chste'),
    @('n?chst',             'n{ae}chst'),
    @('n?her',              'n{ae}her'),
    @('n?mlich',            'n{ae}mlich'),
    @('?hnliche',           '{ae}hnliche'),
    @('?hnlich',            '{ae}hnlich'),
    @('?hnlichkeit',        '{AE}hnlichkeit'),
    @('Tats?chlich',        'Tats{ae}chlich'),
    @('tats?chlich',        'tats{ae}chlich'),
    @('haupts?chlich',      'haupts{ae}chlich'),
    @('zus?tzliche',        'zus{ae}tzliche'),
    @('Zus?tzlich',         'Zus{ae}tzlich'),
    @('zus?tzlich',         'zus{ae}tzlich'),
    @('Abh?ngigkeit',       'Abh{ae}ngigkeit'),
    @('unabh?ngig',         'unabh{ae}ngig'),
    @('abh?ngig',           'abh{ae}ngig'),
    @('Abh?ngig',           'Abh{ae}ngig'),
    @('G?ltigkeitsdauer',   'G{ue}ltigkeitsdauer'),
    @('G?ltigkeit',         'G{ue}ltigkeit'),
    @('Ung?ltig',           'Ung{ue}ltig'),
    @('ung?ltig',           'ung{ue}ltig'),
    @('G?ltig',             'G{ue}ltig'),
    @('g?ltig',             'g{ue}ltig'),
    @('ausgef?hrt',         'ausgef{ue}hrt'),
    @('Ausf?hrung',         'Ausf{ue}hrung'),
    @('Ausf?hren',          'Ausf{ue}hren'),
    @('ausf?hren',          'ausf{ue}hren'),
    @('Eingef?gt',          'Eingef{ue}gt'),
    @('eingef?gt',          'eingef{ue}gt'),
    @('einf?gen',           'einf{ue}gen'),
    @('Einf?gen',           'Einf{ue}gen'),
    @('best?tigen',         'best{ae}tigen'),
    @('Best?tigen',         'Best{ae}tigen'),
    @('best?tigt',          'best{ae}tigt'),
    @('Best?tigung',        'Best{ae}tigung'),
    @('Erkl?rungen',        'Erkl{ae}rungen'),
    @('Erkl?rung',          'Erkl{ae}rung'),
    @('erkl?rung',          'erkl{ae}rung'),
    @('anschlie?end',       'anschlie{ss}end'),
    @('Anschlie?end',       'Anschlie{ss}end'),
    @('Schlie?en',          'Schlie{ss}en'),
    @('schlie?en',          'schlie{ss}en'),
    @('geschlie?en',        'geschlossen'),
    @('ausschlie?lich',     'ausschlie{ss}lich'),
    @('Stra?e',             'Stra{ss}e'),
    @('stra?e',             'stra{ss}e'),
    @('au?erdem',           'au{ss}erdem'),
    @('Au?er',              'Au{ss}er'),
    @('au?er',              'au{ss}er'),
    @('drau?en',            'drau{ss}en'),
    @('hei?en',             'hei{ss}en'),
    @('hei?t',              'hei{ss}t'),
    @('wei?',                'wei{ss}'),
    @('Wei?',                'Wei{ss}'),
    @('Pa?wort',            'Passwort'),

    # ---- f?hren / F?hrt --------------------------------------------------
    @('F?hrungs',           'F{ue}hrungs'),
    @('f?hren',             'f{ue}hren'),
    @('F?hrt',              'F{ue}hrt'),
    @('f?hrt',              'f{ue}hrt'),

    # ---- M?nner / K?rper / Fl?che ---------------------------------------
    @('K?rper',             'K{oe}rper'),
    @('M?nner',             'M{ae}nner'),
    @('Fl?che',             'Fl{ae}che'),
    @('fl?che',             'fl{ae}che'),

    # ---- Spalten?berschrift, verkn?pft, Vorg?nger -----------------------
    @('Spalten?berschrift', 'Spalten{ue}berschrift'),
    @('Hauptr?ckgabe',      'Hauptr{ue}ckgabe'),
    @('Verkn?pfung',        'Verkn{ue}pfung'),
    @('verkn?pft',          'verkn{ue}pft'),
    @('Vorg?nger',          'Vorg{ae}nger'),
    @('Vorh?ngen',          'Vorh{ae}ngen'),
    @('Vorjahres?berhang',  'Vorjahres{ue}berhang'),

    # ---- S?ule / Beitr?ge -----------------------------------------------
    @('S?ulen',             'S{ae}ulen'),
    @('S?ule',              'S{ae}ule'),
    @('Mitgliedsbeitr?ge',  'Mitgliedsbeitr{ae}ge'),
    @('Beitr?gen',          'Beitr{ae}gen'),
    @('Beitr?ge',           'Beitr{ae}ge'),

    # ---- ERWEITERUNG: aus Dry-Run-Restmenge -----------------------------
    # Pr?fe / Pr?ft / PR?FEN / PR?FUNG
    @('Pr?fe',              'Pr{ue}fe'),
    @('Pr?ft',              'Pr{ue}ft'),
    @('PR?FEN',             'PR{UE}FEN'),
    @('PR?FUNG',            'PR{UE}FUNG'),
    @('Pr?fix',             'Pr{ae}fix'),
    # enth?lt
    @('enth?lt',            'enth{ae}lt'),
    # Z?hler / Z?HLER (lange Wortvarianten zuerst!)
    @('Mitgliederz?hler',   'Mitgliederz{ae}hler'),
    @('PARZELLENZ?HLER',    'PARZELLENZ{AE}HLER'),
    @('Z?HLERWECHSEL',      'Z{AE}HLERWECHSEL'),
    @('Z?HLERLOGIK',        'Z{AE}HLERLOGIK'),
    @('Hauptz?hler',        'Hauptz{ae}hler'),
    @('Z?HLER',             'Z{AE}HLER'),
    @('Z?hlers',            'Z{ae}hlers'),
    @('Z?hler',             'Z{ae}hler'),
    @('Z?hlt',              'Z{ae}hlt'),
    @('z?hlen',             'z{ae}hlen'),
    @('Z?hlung',            'Z{ae}hlung'),
    # R?ck / R?ckgabe / R?ckerstattung / zur?ck
    @('R?ckerstattung',     'R{ue}ckerstattung'),
    @('R?ckgabe',           'R{ue}ckgabe'),
    @('zur?ckgesetzt',      'zur{ue}ckgesetzt'),
    @('zur?cksetzen',       'zur{ue}cksetzen'),
    @('zur?ck',             'zur{ue}ck'),
    @('R?CKW?RTS',          'R{UE}CKW{AE}RTS'),
    @('R?CKW',              'R{UE}CKW'),
    # ?NDERUNG / ?nderung schon vorhanden -> nur GROSS-Variante
    @('?NDERUNG',           '{AE}NDERUNG'),
    # F?R (gross)
    @('F?R',                'F{UE}R'),
    # j?hrlich / Halbj?hrlich
    @('Halbj?hrlich',       'Halbj{ae}hrlich'),
    @('j?hrlich',           'j{ae}hrlich'),
    # vollst?ndig
    @('vollst?ndigen',      'vollst{ae}ndigen'),
    @('vollst?ndig',        'vollst{ae}ndig'),
    # W?rter
    @('W?rter',             'W{oe}rter'),
    # gesch?tzt / sch?tzen / sch?tzt
    @('gesch?tzt',          'gesch{ue}tzt'),
    @('sch?tzen',           'sch{ue}tzen'),
    @('sch?tzt',            'sch{ue}tzt'),
    @('Sch?tze',            'Sch{ue}tze'),
    # Betr?ge
    @('Betr?ge',            'Betr{ae}ge'),
    # L?nge
    @('L?nge',              'L{ae}nge'),
    # Nachp?chter
    @('Nachp?chter',        'Nachp{ae}chter'),
    # ausl?sen / l?st / l?se / L?sche
    @('ausl?sen',           'ausl{oe}sen'),
    @('L?sche',             'L{oe}sche'),
    @('l?st',               'l{oe}st'),
    # Bl?tter
    @('Bl?tter',            'Bl{ae}tter'),
    # anh?ngen / hinzuf?gen / hinzugef?gt / F?gen
    @('anh?ngen',           'anh{ae}ngen'),
    @('hinzugef?gt',        'hinzugef{ue}gt'),
    @('hinzuf?gen',         'hinzuf{ue}gen'),
    @('F?gen',              'F{ue}gen'),
    # Regul?re
    @('Regul?re',           'Regul{ae}re'),
    # gew?hlt
    @('gew?hlt',            'gew{ae}hlt'),
    # unn?tige / unn?tigen / n?tig / ben?tigt
    @('unn?tigen',          'unn{oe}tigen'),
    @('unn?tige',           'unn{oe}tige'),
    @('ben?tigt',           'ben{oe}tigt'),
    @('n?tig',              'n{oe}tig'),
    # M?rz
    @('M?rz',               'M{ae}rz'),
    # aufr?umen
    @('aufr?umen',          'aufr{ae}umen'),
    # Zusammenf?hrung / Verf?gung
    @('Zusammenf?hrung',    'Zusammenf{ue}hrung'),
    @('Verf?gung',          'Verf{ue}gung'),
    # Bin?rdaten
    @('Bin?rdaten',         'Bin{ae}rdaten'),
    # ZELLF?RBUNGEN
    @('ZELLF?RBUNGEN',      'ZELLF{AE}RBUNGEN'),
    # Gesch?ftslogik
    @('Gesch?ftslogik',     'Gesch{ae}ftslogik'),
    # Unterst?tzte / unterdr?cken
    @('Unterst?tzte',       'Unterst{ue}tzte'),
    @('unterdr?cken',       'unterdr{ue}cken'),
    # F?lle
    @('F?lle',              'F{ae}lle'),
    # st?rkere
    @('st?rkere',           'st{ae}rkere'),
    # zusammenh?ngender
    @('zusammenh?ngender',  'zusammenh{ae}ngender'),
    # Priorit?ten
    @('Priorit?ten',        'Priorit{ae}ten'),
    # Zeilenh?he
    @('Zeilenh?he',         'Zeilenh{oe}he'),
    # TEMPOR?RE
    @('TEMPOR?RE',          'TEMPOR{AE}RE'),
    # Startmen?
    @('Startmen?',          'Startmen{ue}'),

    # ---- ERWEITERUNG 2: aus zweitem Dry-Run -----------------------------
    @('Enth?lt',            'Enth{ae}lt'),
    @('regul?re',           'regul{ae}re'),
    @('VERF?GBARE',         'VERF{UE}GBARE'),
    @('FUNKTIONALIT?T',     'FUNKTIONALIT{AE}T'),
    @('STABILIT?T',         'STABILIT{AE}T'),
    @('Kompatibilit?t',     'Kompatibilit{ae}t'),
    @('Qualit?tsfaktor',    'Qualit{ae}tsfaktor'),
    @('Priorit?tsbonus',    'Priorit{ae}tsbonus'),
    @('zugeh?rige',         'zugeh{oe}rige'),
    @('F?LLBEREICHE',       'F{UE}LLBEREICHE'),
    @('w?rde',              'w{ue}rde'),
    @('d?rfen',             'd{ue}rfen'),
    @('BETR?GE',            'BETR{AE}GE'),
    @('EINTR?GE',           'EINTR{AE}GE'),
    @('Flie?komma',         'Flie{ss}komma'),
    @('L?SCHEN',            'L{OE}SCHEN'),
    @('L?sch',              'L{oe}sch'),
    @('Ausl?sung',          'Ausl{oe}sung'),
    @('UNGESCH?TZT',        'UNGESCH{UE}TZT'),
    @('Best?tige',          'Best{ae}tige'),
    @('K?hltruhe',          'K{ue}hltruhe'),
    @('Bef?llung',          'Bef{ue}llung'),
    @('F?llung',            'F{ue}llung'),
    @('F?hre',              'F{ue}hre'),
    @('F?ge',               'F{ue}ge'),
    @('unvertr?glich',      'unvertr{ae}glich'),
    @('rechtsb?ndig',       'rechtsb{ue}ndig'),
    @('linksb?ndig',        'linksb{ue}ndig'),
    @('zusammenh?ngend',    'zusammenh{ae}ngend'),
    @('sp?testens',         'sp{ae}testens'),
    @('R?ckzahlung',        'R{ue}ckzahlung'),
    @('Bl?cke',             'Bl{oe}cke'),
    @('N?CHSTE',            'N{AE}CHSTE'),

    # ---- gro? / Schlu? ---------------------------------------------------
    @('gro?',                'gro{ss}'),
    @('Gro?',                'Gro{ss}'),
    @('Schlu?',              'Schluss')
)

# Konvertiere zu OrderedDictionary mit echten Umlauten
# Hinweis: PowerShell ordered hashtables sind case-insensitive beim Key-Lookup,
# darum iterieren wir spaeter direkt ueber $pairs (case-sensitiv).
$dict = [ordered]@{}
foreach ($p in $pairs) {
    $key = $p[0]
    $val = Build($p[1])
    if (-not $dict.Contains($key)) {
        $dict[$key] = $val
    }
}

# Suche aller .bas und .cls Dateien (BackUp-Ordner ausschliessen)
$files = Get-ChildItem -Path $Root -Recurse -File -Include *.bas, *.cls |
         Where-Object { $_.FullName -notmatch '\\BackUp' }

if (-not $files) {
    Write-Warning "Keine .bas/.cls Dateien unter $Root gefunden."
    return
}

Write-Host "Gefundene Dateien:        $($files.Count)" -ForegroundColor Cyan
Write-Host "Woerterbuch-Paare:        $($pairs.Count)" -ForegroundColor Cyan
if ($DryRun) {
    Write-Host "DRY-RUN -- keine Datei wird geschrieben" -ForegroundColor Yellow
}

$utf8WithBom    = New-Object System.Text.UTF8Encoding($true)
$totalReplaced  = 0
$report         = @()
$processedText  = @{}   # rel.Pfad -> verarbeiteter Text (fuer Leftover-Report im DryRun)

foreach ($f in $files) {
    $raw  = [System.IO.File]::ReadAllText($f.FullName)
    $orig = $raw
    $countInFile = 0

    # CASE-SENSITIV ueber $pairs iterieren (nicht ueber dict).
    foreach ($p in $pairs) {
        $k = $p[0]
        $v = Build($p[1])
        if ($k -ceq $v) { continue }
        if (-not $raw.Contains($k)) { continue }

        $occ = ([regex]::Matches($raw, [regex]::Escape($k))).Count
        if ($occ -gt 0) {
            $raw = $raw.Replace($k, $v)
            $countInFile += $occ
        }
    }

    if ($countInFile -gt 0) {
        $totalReplaced += $countInFile
        $relPath = $f.FullName.Substring($Root.Length).TrimStart('\')
        $report += [pscustomobject]@{
            File         = $relPath
            Replacements = $countInFile
        }

        if (-not $DryRun) {
            [System.IO.File]::WriteAllText($f.FullName, $raw, $utf8WithBom)
        }
    }
    $processedText[$f.FullName] = $raw
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host " Ergebnis: $totalReplaced Ersetzungen in $($report.Count) Dateien" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green

if ($report.Count -gt 0) {
    $report | Sort-Object -Property Replacements -Descending |
        Format-Table -AutoSize -Property File, Replacements
}

# Restliche '?' mitten in Woertern -> manuelle Pruefung
# (Trailing-? wie "Fortfahren?" sind legitime MsgBox-Strings und werden ignoriert.)
Write-Host ""
Write-Host "Verbleibende '?' MITTEN in Woertern -- sortiert nach Haeufigkeit:" -ForegroundColor Yellow
$wordRegex = "[A-Za-z$lo_a$lo_o$lo_u$up_A$up_O$up_U$ssch]+\?[A-Za-z$lo_a$lo_o$lo_u$up_A$up_O$up_U$ssch]+"
$wordCounts = @{}
$fileCounts = @{}
foreach ($f in $files) {
    $raw = $processedText[$f.FullName]
    if (-not $raw) { $raw = [System.IO.File]::ReadAllText($f.FullName) }
    $matches = [regex]::Matches($raw, $wordRegex)
    $relPath = $f.FullName.Substring($Root.Length).TrimStart('\')
    if ($matches.Count -gt 0) {
        $fileCounts[$relPath] = $matches.Count
    }
    foreach ($m in $matches) {
        $w = $m.Value
        if ($wordCounts.ContainsKey($w)) {
            $wordCounts[$w] = $wordCounts[$w] + 1
        } else {
            $wordCounts[$w] = 1
        }
    }
}

if ($wordCounts.Count -eq 0) {
    Write-Host "Keine Reste mehr." -ForegroundColor Green
} else {
    Write-Host "  -> Gesamt $($wordCounts.Count) verschiedene Woerter mit '?' in $($fileCounts.Count) Dateien"
    Write-Host ""
    $wordCounts.GetEnumerator() | Sort-Object -Property Value -Descending |
        Select-Object -First 80 |
        Format-Table -AutoSize @{N='Wort'; E={$_.Key}}, @{N='Vorkommen'; E={$_.Value}}
    Write-Host ""
    Write-Host "Pro Datei:" -ForegroundColor Yellow
    $fileCounts.GetEnumerator() | Sort-Object -Property Value -Descending |
        Format-Table -AutoSize @{N='Datei'; E={$_.Key}}, @{N='Resterst'; E={$_.Value}}
}
