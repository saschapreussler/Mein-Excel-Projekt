# Punkt 9: Umlaute in VBA .bas/.cls wiederherstellen
# Strategie: Wörterbuch deutscher Standard-Wörter (case-sensitive)
# Speichert als Windows-1252 (cp1252) - VBA-Standard

$ErrorActionPreference = "Stop"

# Wörterbuch: ASCII-Variante (mit ?) -> richtige deutsche Schreibweise
# Reihenfolge wichtig: längere Patterns zuerst (vermeidet Substring-Konflikte)
$dict = [ordered]@{
    # SS / ß
    'gr????'        = 'größ'   # placeholder - längste
    'au?ergew?hnlich' = 'außergewöhnlich'
    'au?erhalb'     = 'außerhalb'
    'au?erdem'      = 'außerdem'
    'Au?erdem'      = 'Außerdem'
    'gr??enordn'    = 'größenordn'
    'gr??te'        = 'größte'
    'gr??er'        = 'größer'
    'Gr??te'        = 'Größte'
    'Gr??er'        = 'Größer'
    'Gr??e'         = 'Größe'
    'gr??e'         = 'größe'
    'schlie?lich'   = 'schließlich'
    'Schlie?lich'   = 'Schließlich'
    'schlie?en'     = 'schließen'
    'Schlie?en'     = 'Schließen'
    'schlie?t'      = 'schließt'
    'Schlie?t'      = 'Schließt'
    'au?er'         = 'außer'
    'Au?er'         = 'Außer'
    'hei?t'         = 'heißt'
    'Hei?t'         = 'Heißt'
    'hei?en'        = 'heißen'
    'mu?'           = 'muß'
    'Stra?e'        = 'Straße'
    'stra?e'        = 'straße'
    'wei?'          = 'weiß'
    'Wei?'          = 'Weiß'
    'gewi?'         = 'gewiß'
    'gro?'          = 'groß'
    'Gro?'          = 'Groß'
    'lie?'          = 'ließ'
    'Lie?'          = 'Ließ'
    'flie?'         = 'fließ'
    'Flie?'         = 'Fließ'
    'sto?'          = 'stoß'
    'Sto?'          = 'Stoß'
    'pa?'           = 'paß'
    'Pa?'           = 'Paß'
    'sch?tz'        = 'schütz'   # context: Schutz dominiert über Schätz
    'Sch?tz'        = 'Schütz'
    'gesch?tz'      = 'geschütz'
    'Gesch?tz'      = 'Geschütz'
    'ungesch?tz'    = 'ungeschütz'

    # Ü
    '?bersicht'     = 'Übersicht'
    '?bertrag'      = 'Übertrag'
    '?berschrift'   = 'Überschrift'
    '?berschreib'   = 'Überschreib'
    '?berschritt'   = 'Überschritt'
    '?bernehm'      = 'Übernehm'
    '?bernimm'      = 'Übernimm'
    '?bernomm'      = 'Übernomm'
    '?bergeb'       = 'Übergeb'
    '?berge'        = 'Überge'
    '?berpr?f'      = 'Überprüf'
    '?ber'          = 'Über'
    '?brig'         = 'übrig'
    '?bung'         = 'Übung'
    'zur?ck'        = 'zurück'
    'Zur?ck'        = 'Zurück'
    'ZUR?CK'        = 'ZURÜCK'
    'gl?ck'         = 'glück'
    'Gl?ck'         = 'Glück'
    'br?ck'         = 'brück'
    'Br?ck'         = 'Brück'
    'st?ck'         = 'stück'
    'St?ck'         = 'Stück'
    'dr?ck'         = 'drück'
    'Dr?ck'         = 'Drück'
    'r?ck'          = 'rück'
    'R?ck'          = 'Rück'
    'g?ltig'        = 'gültig'
    'G?ltig'        = 'Gültig'
    'GR?N'          = 'GRÜN'
    'gr?n'          = 'grün'
    'Gr?n'          = 'Grün'
    'pr?f'          = 'prüf'
    'Pr?f'          = 'Prüf'
    'PR?F'          = 'PRÜF'
    'ausf?hr'       = 'ausführ'
    'Ausf?hr'       = 'Ausführ'
    'durchf?hr'     = 'durchführ'
    'Durchf?hr'     = 'Durchführ'
    'einf?hr'       = 'einführ'
    'Einf?hr'       = 'Einführ'
    'ausgef?hr'     = 'ausgeführ'
    'unterst?tz'    = 'unterstütz'
    'Unterst?tz'    = 'Unterstütz'
    'k?nstl'        = 'künstl'
    'K?nstl'        = 'Künstl'
    'k?rz'          = 'kürz'
    'K?rz'          = 'Kürz'
    'sp?t'          = 'spät'   # ä actually
    'm?glich'       = 'möglich'   # this is ö, careful with order
    'M?glich'       = 'Möglich'
    'Unm?glich'     = 'Unmöglich'
    'unm?glich'     = 'unmöglich'
    'erm?glich'     = 'ermöglich'
    'Erm?glich'     = 'Ermöglich'
    'F?r '          = 'Für '
    'f?r '          = 'für '
    ' f?r,'         = ' für,'
    ' f?r:'         = ' für:'
    'f?rs '         = 'fürs '
    'F?rs '         = 'Fürs '
    'b?ndig'        = 'bündig'
    'B?ndig'        = 'Bündig'
    'rechtsb?ndig'  = 'rechtsbündig'
    'linksb?ndig'   = 'linksbündig'
    'm?ssen'        = 'müssen'
    'M?ssen'        = 'Müssen'
    'm?sste'        = 'müsste'
    'M?sste'        = 'Müsste'
    'm?ssten'       = 'müssten'
    'm?ssen'        = 'müssen'
    'fl?ssig'       = 'flüssig'
    'Fl?ssig'       = 'Flüssig'
    'gr?nde'        = 'gründe'
    'Gr?nde'        = 'Gründe'
    'gr?nd'         = 'gründ'
    'Gr?nd'         = 'Gründ'
    'beg?nstig'     = 'begünstig'
    'Beg?nstig'     = 'Begünstig'
    'verf?gbar'     = 'verfügbar'
    'Verf?gbar'     = 'Verfügbar'
    'verf?g'        = 'verfüg'
    'Verf?g'        = 'Verfüg'
    'gen?g'         = 'genüg'
    'Gen?g'         = 'Genüg'
    'unverz?glich'  = 'unverzüglich'
    'urspr?nglich'  = 'ursprünglich'
    'Urspr?nglich'  = 'Ursprünglich'
    'gegen?ber'     = 'gegenüber'
    'Gegen?ber'     = 'Gegenüber'
    'einf?g'        = 'einfüg'
    'Einf?g'        = 'Einfüg'
    'eingef?g'      = 'eingefüg'
    'hinzuf?g'      = 'hinzufüg'
    'Hinzuf?g'      = 'Hinzufüg'
    'verkn?pf'      = 'verknüpf'
    'Verkn?pf'      = 'Verknüpf'
    'tr?b'          = 'trüb'
    'm?h'           = 'müh'
    'M?h'           = 'Müh'
    'fl?ch'         = 'flüch'
    'gef?hl'        = 'gefühl'
    'Gef?hl'        = 'Gefühl'
    'f?hl'          = 'fühl'
    'F?hl'          = 'Fühl'
    'fr?hst?ck'     = 'frühstück'
    'fr?h'          = 'früh'
    'Fr?h'          = 'Früh'
    'm?d'           = 'müd'
    'M?d'           = 'Müd'
    'urspr?ng'      = 'ursprüng'
    'b?ro'          = 'büro'
    'B?ro'          = 'Büro'
    'k?nf'          = 'künf'   # zukünftig
    'K?nf'          = 'Künf'
    'erm?dl'        = 'ermüdl'
    'm?nd'          = 'münd'
    'M?nd'          = 'Münd'
    'r?ckg'         = 'rückg'
    'R?ckg'         = 'Rückg'
    'br?d'          = 'brüd'
    'Br?d'          = 'Brüd'
    'st?tz'         = 'stütz'
    'St?tz'         = 'Stütz'
    'Sch?ler'       = 'Schüler'
    'sch?ler'       = 'schüler'
    'h?gel'         = 'hügel'
    'H?gel'         = 'Hügel'
    'h?bsch'        = 'hübsch'
    'H?bsch'        = 'Hübsch'
    'b?ndel'        = 'bündel'
    'B?ndel'        = 'Bündel'
    'erm?glich'     = 'ermöglich'
    'einf?hl'       = 'einfühl'
    'wir?'          = 'wirß'    # rare
    'r?hr'          = 'rühr'
    'R?hr'          = 'Rühr'
    'sp?l'          = 'spül'
    'Sp?l'          = 'Spül'
    'l?ck'          = 'lück'
    'L?ck'          = 'Lück'
    'h?lle'         = 'hülle'
    'H?lle'         = 'Hülle'

    # Ö
    '?ffnen'        = 'Öffnen'
    '?ffne'         = 'öffne'
    'ge?ffnet'      = 'geöffnet'
    '?konom'        = 'Ökonom'
    '?l'            = 'Öl'
    '?sterr'        = 'Österr'
    'k?nn'          = 'könn'
    'K?nn'          = 'Könn'
    'kann?'         = 'kannß'   # rare
    'l?sch'         = 'lösch'
    'L?sch'         = 'Lösch'
    'L?SCH'         = 'LÖSCH'
    'gel?sch'       = 'gelösch'
    'l?sung'        = 'lösung'
    'L?sung'        = 'Lösung'
    'aufl?s'        = 'auflös'
    'Aufl?s'        = 'Auflös'
    'st?r'          = 'stör'
    'St?r'          = 'Stör'
    'gest?rt'       = 'gestört'
    'h?h'           = 'höh'
    'H?h'           = 'Höh'
    'h?her'         = 'höher'
    'H?her'         = 'Höher'
    'h?chst'        = 'höchst'
    'H?chst'        = 'Höchst'
    'sch?n'         = 'schön'
    'Sch?n'         = 'Schön'
    'pers?nlich'    = 'persönlich'
    'Pers?nlich'    = 'Persönlich'
    'pl?tzlich'     = 'plötzlich'
    'Pl?tzlich'     = 'Plötzlich'
    'm?cht'         = 'möcht'
    'M?cht'         = 'Möcht'
    'gem?cht'       = 'gemöcht'
    'h?ren'         = 'hören'
    'H?ren'         = 'Hören'
    'geh?r'         = 'gehör'
    'Geh?r'         = 'Gehör'
    'angeh?r'       = 'angehör'
    'verm?gen'      = 'vermögen'
    'Verm?gen'      = 'Vermögen'
    'm?ge'          = 'möge'
    'M?ge'          = 'Möge'
    'gr??t'         = 'größt'   # already covered
    'r?t'           = 'röt'
    'R?T'           = 'RÖT'
    'ROT'           = 'ROT'   # Don't touch
    'gel?st'        = 'gelöst'
    'Gel?st'        = 'Gelöst'
    'erh?h'         = 'erhöh'
    'Erh?h'         = 'Erhöh'
    'verm?g'        = 'vermög'
    'Verm?g'        = 'Vermög'
    'k?rper'        = 'körper'
    'K?rper'        = 'Körper'
    'h?fl'          = 'höfl'
    'H?fl'          = 'Höfl'
    'erw?gung'      = 'erwögung'
    'erfo?'         = 'erfoß'
    'M?N'           = 'MÖN'   # rare
    'd?rfen'        = 'dürfen'   # NO, ü
    'D?rfen'        = 'Dürfen'
    'd?rf'          = 'dürf'
    'D?rf'          = 'Dürf'
    'fro?'          = 'froß'   # rare
    'kann?'         = 'kannß'

    # Ä
    '?nderung'      = 'Änderung'
    '?ndern'        = 'Ändern'
    '?ndert'        = 'Ändert'
    '?ndere'        = 'Ändere'
    'ge?ndert'      = 'geändert'
    '?nderbar'      = 'Änderbar'
    '?ndere'        = 'Ändere'
    '?hnlich'       = 'Ähnlich'
    '?hnlichkeit'   = 'Ähnlichkeit'
    '?hnel'         = 'Ähnel'
    '?u?er'         = 'Äußer'
    '?u?ern'        = 'Äußern'
    '?u?erst'       = 'Äußerst'
    '?lt'           = 'ält'   # älter
    '?ltere'        = 'ältere'
    '?lteste'       = 'älteste'
    'ber?cksichtig' = 'berücksichtig'
    'Ber?cksichtig' = 'Berücksichtig'
    'erkl?r'        = 'erklär'
    'Erkl?r'        = 'Erklär'
    'gekl?r'        = 'geklär'
    'aufkl?r'       = 'aufklär'
    'gef?hr'        = 'gefähr'
    'Gef?hr'        = 'Gefähr'
    'ungef?hr'      = 'ungefähr'
    'g?ng'          = 'gäng'   # gängig
    'G?ng'          = 'Gäng'
    'urspr?ng'      = 'ursprüng'   # already
    'l?ng'          = 'läng'   # länge
    'L?ng'          = 'Läng'
    'l?nge'         = 'länge'
    'L?nge'         = 'Länge'
    'verl?nge'      = 'verlänge'
    'Verl?nge'      = 'Verlänge'
    'l?nger'        = 'länger'
    'L?nger'        = 'Länger'
    'st?ndig'       = 'ständig'
    'St?ndig'       = 'Ständig'
    'stet?'         = 'stetß'
    'k?lt'          = 'kält'
    'K?lt'          = 'Kält'
    'erkl?rt'       = 'erklärt'   # already covered
    'verst?nd'      = 'verständ'
    'Verst?nd'      = 'Verständ'
    'best?tig'      = 'bestätig'
    'Best?tig'      = 'Bestätig'
    'best?nd'       = 'beständ'
    'Best?nd'       = 'Beständ'
    'unbest?ndig'   = 'unbeständig'
    'tats?chlich'   = 'tatsächlich'
    'Tats?chlich'   = 'Tatsächlich'
    'verf?ng'       = 'verfäng'
    'umf?ng'        = 'umfäng'
    'Umf?ng'        = 'Umfäng'
    'eintr?g'       = 'einträg'
    'Eintr?g'       = 'Einträg'
    'getr?g'        = 'geträg'
    'tr?g'          = 'träg'
    'Tr?g'          = 'Träg'
    'tr?gt'         = 'trägt'
    'Tr?gt'         = 'Trägt'
    'betr?g'        = 'beträg'
    'Betr?g'        = 'Beträg'
    'antr?g'        = 'anträg'
    'Antr?g'        = 'Anträg'
    'auftr?g'       = 'aufträg'
    'Auftr?g'       = 'Aufträg'
    'auftr?ge'      = 'aufträge'
    'k?m'           = 'käm'
    'K?m'           = 'Käm'
    'jeweil?'       = 'jeweilß'  # jeweils — actually has no umlaut
    'sp?ter'        = 'später'
    'Sp?ter'        = 'Später'
    'verz?gerung'   = 'verzögerung'   # ö
    'Verz?gerung'   = 'Verzögerung'
    'verz?ger'      = 'verzöger'
    'Verz?ger'      = 'Verzöger'
    'unverz?glich'  = 'unverzüglich'   # already
    'l?dt'          = 'lädt'
    'L?dt'          = 'Lädt'
    'gel?d'         = 'gelad'   # nope: geladen no umlaut
    'eingel?dt'     = 'eingelädt'
    'erh?ltlich'    = 'erhältlich'
    'Erh?ltlich'    = 'Erhältlich'
    'erh?lt'        = 'erhält'
    'Erh?lt'        = 'Erhält'
    'enth?lt'       = 'enthält'
    'Enth?lt'       = 'Enthält'
    'beh?lt'        = 'behält'
    'Beh?lt'        = 'Behält'
    'verh?lt'       = 'verhält'
    'Verh?lt'       = 'Verhält'
    'h?ufig'        = 'häufig'
    'H?ufig'        = 'Häufig'
    'unzul?ssig'    = 'unzulässig'
    'zul?ssig'      = 'zulässig'
    'Zul?ssig'      = 'Zulässig'
    'zus?tzlich'    = 'zusätzlich'
    'Zus?tzlich'    = 'Zusätzlich'
    's?mtlich'      = 'sämtlich'
    'S?mtlich'      = 'Sämtlich'
    'r?um'          = 'räum'
    'R?um'          = 'Räum'
    'aufr?um'       = 'aufräum'
    'erw?hn'        = 'erwähn'
    'Erw?hn'        = 'Erwähn'
    'erh?hen'       = 'erhöhen'   # ö, see Ö section
    'verm?ge'       = 'vermöge'
    'gez?hl'        = 'gezähl'
    'gez?hlt'       = 'gezählt'
    'z?hl'          = 'zähl'
    'Z?hl'          = 'Zähl'
    'Z?HL'          = 'ZÄHL'
    'z?hler'        = 'zähler'
    'Z?hler'        = 'Zähler'
    'z?hlerwechsel' = 'zählerwechsel'
    'Z?hlerwechsel' = 'Zählerwechsel'
    'aufz?hl'       = 'aufzähl'
    'Aufz?hl'       = 'Aufzähl'
    'z?hlt'         = 'zählt'
    'Z?hlt'         = 'Zählt'
    'erz?hl'        = 'erzähl'
    'Erz?hl'        = 'Erzähl'
    's?ulen'        = 'säulen'
    'S?ulen'        = 'Säulen'
    'h?lften'       = 'hälften'
    'H?lfte'        = 'Hälfte'
    'h?lfte'        = 'hälfte'
    'h?nde'         = 'hände'
    'H?nde'         = 'Hände'
    'h?nd'          = 'händ'
    'H?nd'          = 'Händ'
    'umh?nd'        = 'umhänd'
    'verb?nd'       = 'verbänd'
    'k?nnt'         = 'könnt'   # ö
    'K?nnt'         = 'Könnt'
    'h?lt'          = 'hält'
    'H?lt'          = 'Hält'
    'erh?lt'        = 'erhält'  # already
    'h?lte'         = 'hälte'   # rare
    'h?ltlich'      = 'hältlich'
    'h?ltigen'      = 'hältigen'

    # Specific phrases / domain
    'W?hrung'       = 'Währung'
    'w?hrend'       = 'während'
    'W?hrend'       = 'Während'
    'W?hr'          = 'Währ'
    'w?hr'          = 'währ'
    'erw?hl'        = 'erwähl'
    'Erw?hl'        = 'Erwähl'
    'gew?hl'        = 'gewähl'
    'gew?hlt'       = 'gewählt'
    'gew?hlte'      = 'gewählte'
    'w?hl'          = 'wähl'
    'W?hl'          = 'Wähl'
    'gem??'         = 'gemäß'
    'gem???'        = 'gemääß'   # rare error
    'M??ig'         = 'Mäßig'
    'm??ig'         = 'mäßig'
    'rechtm??ig'    = 'rechtmäßig'
    'regelm??ig'    = 'regelmäßig'
    'Regelm??ig'    = 'Regelmäßig'
    'm??ig'         = 'mäßig'
    'gleichm??ig'   = 'gleichmäßig'
    'unverh?ltnism??ig' = 'unverhältnismäßig'

    # Hilfs-/Funktions-Wörter
    'pr?fung'       = 'prüfung'
    'Pr?fung'       = 'Prüfung'
    'gepr?ft'       = 'geprüft'
    'gepr?fte'      = 'geprüfte'
    'überpr?fung'   = 'überprüfung'
    '?berpr?fung'   = 'Überprüfung'
    'M??'           = 'Mäß'    # häufig in "Maße"

    # Häufige Endungen / Spezielle
    'l?st'          = 'löst'
    'L?st'          = 'Löst'
    'l?sst'         = 'lässt'  # ä
    'L?sst'         = 'Lässt'
    'l?ssig'        = 'lässig'
    'L?ssig'        = 'Lässig'

    # Datensätze, Eintrag etc.
    'Datens?tze'    = 'Datensätze'
    'Datens?tzen'   = 'Datensätzen'
    'Eintr?ge'      = 'Einträge'
    'Eintr?gen'     = 'Einträgen'
    'Vortr?ge'      = 'Vorträge'
    'Auftr?ge'      = 'Aufträge'  # already
    'Bel?ge'        = 'Beläge'
    'Vors?tze'      = 'Vorsätze'
    'Aus?tze'       = 'Ausätze'
    'Anst??e'       = 'Anstöße'
    'Stra?en'       = 'Straßen'
    'Gr??en'        = 'Größen'

    # Abrechnung context
    'Abrechnungsj?hr' = 'Abrechnungsjähr'   # rare
    'j?hrlich'      = 'jährlich'
    'J?hrlich'      = 'Jährlich'
    'monatl'        = 'monatl'   # no umlaut
    'h?lfte'        = 'hälfte'

    # Misc nouns
    'Eintr?gt'      = 'Einträgt'
    'Ger?t'         = 'Gerät'
    'ger?t'         = 'gerät'
    'sp?t'          = 'spät'   # already covered
    'Sp?t'          = 'Spät'
    'ben?tig'       = 'benötig'
    'Ben?tig'       = 'Benötig'
    'gen?tig'       = 'genötig'
    'h?fl'          = 'höfl'
    'H?fl'          = 'Höfl'
    'tr?st'         = 'tröst'
    'Tr?st'         = 'Tröst'
    'gel?ufig'      = 'geläufig'
    'l?uf'          = 'läuf'
    'L?uf'          = 'Läuf'
    'r?uml'         = 'räuml'
    'R?uml'         = 'Räuml'
    'r?ucher'       = 'räucher'
    'R?ucher'       = 'Räucher'
    'br?u'          = 'bräu'
    'Br?u'          = 'Bräu'

    # Konsonanten + Umlaut bei Wortanfang
    'h?lt'          = 'hält'   # already
    'm?ssen'        = 'müssen' # already
    'k?nnte'        = 'könnte'
    'K?nnte'        = 'Könnte'
    'k?nnten'       = 'könnten'
    'K?nnten'       = 'Könnten'
    'k?nnen'        = 'können'
    'K?nnen'        = 'Können'
    'm?cht'         = 'möcht'  # already
    'sollt'         = 'sollt'  # no umlaut
    'l?st'          = 'löst'   # already
    
    # Verben / weitere
    'flie?en'       = 'fließen'
    'Flie?en'       = 'Fließen'
    'verflo?'       = 'verfloß'
    'genie?'        = 'genieß'
    'Genie?'        = 'Genieß'
    'gie?'          = 'gieß'
    'Gie?'          = 'Gieß'
    'verlie?'       = 'verließ'
    'beschlie?'     = 'beschließ'
    'Beschlie?'     = 'Beschließ'
    'erschlie?'     = 'erschließ'
    'aufschlie?'    = 'aufschließ'
    'gru?'          = 'gruß'
    'Gru?'          = 'Gruß'
    'gewi?'         = 'gewiß'

    # Spezial-Begriffe (Domäne)
    'Sch?tz'        = 'Schütz'   # already
    'Datentr?ger'   = 'Datenträger'
    'Tr?ger'        = 'Träger'
    'tr?ger'        = 'träger'
    'Drucker?nderung' = 'Druckeränderung'
    'Schl?ssel'     = 'Schlüssel'
    'schl?ssel'     = 'schlüssel'
    'Br?cke'        = 'Brücke'
    'br?cke'        = 'brücke'
    'gr?nde'        = 'gründe'   # already
    'Gr?nde'        = 'Gründe'
    'verg?nstigung' = 'vergünstigung'
    'Verg?nstigung' = 'Vergünstigung'
    'M?gen'         = 'Mögen'
    'm?gen'         = 'mögen'
    'm?gl'          = 'mögl'   # already covered
    'fro?'          = 'froß'

    # nochmal aufgeräumt: häufige in diesem Codebase
    'Geb?hr'        = 'Gebühr'
    'geb?hr'        = 'gebühr'
    'Geb?hren'      = 'Gebühren'
    'geb?hren'      = 'gebühren'
    'S?umnis'       = 'Säumnis'
    's?umnis'       = 'säumnis'
    'S?umni'        = 'Säumni'
    's?umni'        = 'säumni'
    'r?ck'          = 'rück'   # already
    'Pl?tze'        = 'Plätze'
    'pl?tz'         = 'plätz'  # plätzlich already covered separately
    'Bel?ge'        = 'Beläge'
    'erg?nz'        = 'ergänz'
    'Erg?nz'        = 'Ergänz'
    'k?rzlich'      = 'kürzlich'
    'K?rzlich'      = 'Kürzlich'
    'h?ngen'        = 'hängen'
    'H?ngen'        = 'Hängen'
    'aufh?ng'       = 'aufhäng'
    'umh?ng'        = 'umhäng'
    'unzug?nglich'  = 'unzugänglich'
    'zug?nglich'    = 'zugänglich'
    'Zug?nglich'    = 'Zugänglich'

    # ß placeholder rückgängig
    'gr????'        = 'größ'

    # Letzter Sweep für Einzelfälle die bisher fehlen
    'r?ckg?ng'      = 'rückgäng'   # rückgängig
    'r?ckverg?t'    = 'rückvergüt'
    'verg?t'        = 'vergüt'
    'Verg?t'        = 'Vergüt'
    'd?rfen'        = 'dürfen'  # already
    'D?rfen'        = 'Dürfen'
    'erw?nscht'     = 'erwünscht'
    'unerw?nscht'   = 'unerwünscht'
    'w?nschen'      = 'wünschen'
    'W?nschen'      = 'Wünschen'
    'gew?nscht'     = 'gewünscht'
    'Gew?nscht'     = 'Gewünscht'
    'ben?tzt'       = 'benützt'
    'M?lleimer'     = 'Mülleimer'
    'M?ll'          = 'Müll'
    'm?ll'          = 'müll'

    # === Letzte Säuberungen mit Word-Boundaries via Regex später ===
}

# Backup-Verzeichnis
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$backupRoot = "vba\BackUp_Punkt9_$ts"
New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null

$files = @()
$files += Get-ChildItem -Path "vba\Modules" -Filter "*.bas" -File
$files += Get-ChildItem -Path "vba\Classes" -Filter "*.cls" -File

$cp1252 = [System.Text.Encoding]::GetEncoding(1252)

$totalReplacements = 0
$filesChanged = 0

foreach ($f in $files) {
    # Backup
    Copy-Item $f.FullName -Destination (Join-Path $backupRoot $f.Name) -Force
    
    $content = [System.IO.File]::ReadAllText($f.FullName, $cp1252)
    $original = $content
    $fileReps = 0
    
    foreach ($key in $dict.Keys) {
        $val = $dict[$key]
        if ($content.Contains($key)) {
            # Count occurrences
            $count = ([regex]::Matches([regex]::Escape($content), [regex]::Escape($key))).Count
            $content = $content.Replace($key, $val)
            $fileReps += $count
        }
    }
    
    if ($content -ne $original) {
        [System.IO.File]::WriteAllText($f.FullName, $content, $cp1252)
        $filesChanged++
        $totalReplacements += $fileReps
        Write-Host "  $($f.Name): $fileReps Ersetzungen" -ForegroundColor Cyan
    }
}

Write-Host ""
Write-Host "FERTIG: $filesChanged Dateien geaendert, $totalReplacements Ersetzungen total." -ForegroundColor Green
Write-Host "Backup: $backupRoot" -ForegroundColor Yellow

# Verbleibende ?-Stellen zaehlen
Write-Host ""
Write-Host "Verbleibende '?'-Zeichen pro Datei (Top 15):" -ForegroundColor Yellow
$files | ForEach-Object {
    $c = ([System.IO.File]::ReadAllText($_.FullName, $cp1252) | Select-String -Pattern '\?' -AllMatches).Matches.Count
    [PSCustomObject]@{ File = $_.Name; Count = $c }
} | Sort-Object Count -Descending | Select-Object -First 15 | Format-Table -AutoSize
