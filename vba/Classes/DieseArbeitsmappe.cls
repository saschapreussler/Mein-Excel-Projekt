VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()

    ' Sicherheits-Reset: EnableEvents koennte durch Absturz
    ' einer vorherigen Makro-Ausfuehrung auf False stehen
    Application.EnableEvents = True

    Call mod_Banking_Report.Initialize_ImportReport_ListBox
    
    ' Bereinige eventuell verwaiste TEMP Worksheets
    On Error Resume Next
    Call mod_Hilfsfunktionen.BereinigeTempWorksheets
    On Error GoTo 0
    
    ' Stelle sicher dass alle Blaetter vertikal zentriert sind
    Call mod_Formatierung.ZentriereAlleZellenVertikal
    
    ' NEU v5.3: Hilfsspalten auf Daten-Blatt ausblenden
    Call mod_Format_Protection.BlendeDatenSpaltenAus

    ' Mitgliederliste: DropDowns + Spaltenschutz aktualisieren
    Call mod_Mitglieder_UI.ApplyMitgliederDropdowns
    
    ' v5.1: AutoFilter vorbereiten (wird am Ende erneut aufgerufen)
    ' (Platzhalter - Hauptaufruf am Ende von Workbook_Open)
    
    ' v7.1: Navigationszeilen migrieren (einmalig: 2 Zeilen oben einfuegen)
    On Error Resume Next
    Call mod_Navigation.MigriereNavigationszeilen
    On Error GoTo 0
    
    ' v7.2: Zahlungsuebersicht Monats-Register (Shapes) neu positionieren
    '        Notwendig damit Shapes nach Zeilenverschiebung korrekt sitzen
    On Error Resume Next
    Call mod_Uebersicht_Filter.ErstelleMonatsRegister
    On Error GoTo 0
    
    ' v6.0: Einstellungen-Layout migrieren (einmalig)
    On Error Resume Next
    Call mod_Einstellungen.MigriereEinstellungenLayout
    On Error GoTo 0
    
    ' v6.1: Abrechnungsjahr pruefen (ggf. Nutzer fragen)
    On Error Resume Next
    Call mod_Einstellungen.PruefeAbrechnungsjahr
    On Error GoTo 0
    
    ' v6.0: Kontostand Vorjahr pruefen (ggf. Nutzer fragen)
    On Error Resume Next
    Call mod_Einstellungen.PruefeKontostandVorjahr
    On Error GoTo 0
    
    ' v6.1: Vereinsdaten pruefen (ggf. Nutzer fragen)
    On Error Resume Next
    Call mod_Einstellungen.PruefeVereinsdaten
    On Error GoTo 0
    
    ' v6.0: Startseite initialisieren (KPIs, Buttons, Design)
    On Error Resume Next
    Call mod_Startseite.InitialisiereStartseite
    On Error GoTo 0
    
    ' v6.0: Home-Buttons auf allen Blaettern setzen
    On Error Resume Next
    Call mod_Navigation.SetzeHomeButtonsAufAllenBlaettern
    On Error GoTo 0
    
    ' v6.1: Vereinskasse ComboBox + Formeln initialisieren
    On Error Resume Next
    Call mod_Vereinskasse_Filter.InitialisiereVereinskasseComboBox
    Dim wsVK As Worksheet
    Set wsVK = ThisWorkbook.Worksheets(WS_VEREINSKASSE)
    If Not wsVK Is Nothing Then
        Call mod_Vereinskasse_Filter.SetzeVereinskasseFormeln(wsVK)
    End If
    On Error GoTo 0
    
    ' v6.1: Bankkonto-Formeln wiederherstellen (inkl. E2)
    On Error Resume Next
    Dim wsBK As Worksheet
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    If Not wsBK Is Nothing Then
        Call mod_Banking_Format.StelleFormelnWiederHer(wsBK)
    End If
    On Error GoTo 0

    ' v7.3 (alt): Aktiviere_BankkontoFilter wurde hier aufgerufen,
    '             aber StelleAutoFilterBereit weiter unten hat ihn
    '             wieder ueberschrieben. -> Aufruf nach unten verschoben.
    
    ' v6.0: Startseite als erstes Blatt anzeigen
    On Error Resume Next
    ThisWorkbook.Worksheets(WS_STARTMENUE()).Activate
    On Error GoTo 0
    
    ' v7.0: AutoFilter ZULETZT aufrufen, damit AllowFiltering
    '        nicht durch vorherige Protect-Aufrufe verloren geht
    On Error Resume Next
    Call mod_Formatierung.StelleAutoFilterBereit
    On Error GoTo 0
    
    ' v8.0: Standard-Zeilenhoehen auf Hauptblaettern setzen
    On Error Resume Next
    Call mod_Zeilenhoehen.FixiereStandardZeilenhoehen
    On Error GoTo 0
    
    ' v7.4: Bankkonto-AutoFilter ALLERLETZT setzen, weil
    '       StelleAutoFilterBereit ihn sonst clobbered hat.
    On Error Resume Next
    Call mod_Banking_Format.Aktiviere_BankkontoFilter
    On Error GoTo 0

    ' v8.2: Vereinskasse-AutoFilter B26:T26 (A26 bleibt bewusst leer)
    On Error Resume Next
    Call mod_Vereinskasse_Filter.Aktiviere_VereinskasseFilter
    On Error GoTo 0

    ' v8.2: Aktuelle Parzellen-Anzahl in Strom!B6 und Wasser!A4
    On Error Resume Next
    Call mod_Startseite.AktualisiereParzellenAnzeigen
    On Error GoTo 0

    ' v8.2: Mitgliederliste durchgaengig formatieren
    On Error Resume Next
    Call mod_Formatierung.FormatiereMitgliederlisteKomplett
    On Error GoTo 0

End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    On Error GoTo 0
End Sub


' ===============================================================
' SheetActivate: Gitternetzlinien je Blatt korrekt setzen
' Startmenue, Finanz- bersicht und Dashboard ohne Gitternetz,
' alle anderen Bl tter (BK, VK, Strom, Wasser,  bersicht etc.) MIT Gitternetz
'
' v8.0: Startmenue + Dashboard zus tzlich im Vollbildmodus (kein
'       Ribbon, keine Spalten-/Zeilennummern, keine Formelleiste).
'       Alle anderen Bl tter im normalen Anzeigemodus.
' ===============================================================
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    Dim shName As String
    shName = Sh.Name
    
    Dim ohneGitter As Boolean
    Dim vollbild As Boolean
    ohneGitter = False
    vollbild = False
    
    Select Case shName
        Case WS_STARTMENUE(), "Dashboard Mitgliederzahlungen"
            ohneGitter = True
            vollbild = True
        Case WS_FINANZ_UEBERSICHT()
            ohneGitter = True
            vollbild = False
    End Select
    
    If ActiveWindow Is Nothing Then Exit Sub
    ActiveWindow.DisplayGridlines = Not ohneGitter
    
    ' Vollbildmodus + Header/Formelleiste umschalten
    Call mod_Vollbild.SetzeVollbildModus(vollbild)
End Sub



' ===============================================================
' v8.1: Defensiver Cleanup beim Schliessen
' Verhindert, dass Excel nach dem Schliessen "haengt", indem
' ALS ERSTES alle Events deaktiviert werden, damit waehrend des
' Teardowns keine SheetActivate / SheetChange / Worksheet_Calculate
' Handler mehr feuern. Anschliessend werden Clipboard, offene
' UserForms und Application-States neutralisiert.
'
' Wichtig: ActiveWindow-Properties (DisplayHeadings/Gridlines) werden
' NICHT mehr beruehrt - das kann SheetActivate-Kaskaden ausloesen.
' Excel stellt den korrekten Anzeigezustand beim naechsten Oeffnen
' via Workbook_Open + Workbook_SheetActivate wieder her.
' ===============================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    ' (1) GANZ ZUERST: Events aus - keine Reentrancy mehr ab hier
    Application.EnableEvents = False
    
    ' (2) Laufende Berechnungen / Bildschirm einfrieren
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = True
    Application.Cursor = xlDefault
    Application.StatusBar = False
    Application.CutCopyMode = False    ' Clipboard freigeben
    
    ' (3) Vollbildmodus deaktivieren (vor allem fuer naechstes Oeffnen)
    Call mod_Vollbild.SetzeVollbildModus(False)
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    
    ' (4) Alle UserForms entladen (haengen sonst im Speicher)
    Dim frm As Object
    For Each frm In VBA.UserForms
        Unload frm
    Next frm
    
    ' (5) Calculation wieder auf Automatic - schadet Schliessen nicht
    '     und sorgt fuer sauberen State falls "Abbrechen" gedrueckt wird
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Events bleiben absichtlich AUS - Workbook_Open reaktiviert sie
    ' beim naechsten Start (siehe Workbook_Open Zeile 1).
End Sub


' ===============================================================
' v8.0: Vor Speichern - Vollbildmodus deaktivieren, damit beim
' naechsten Oeffnen keine inkonsistente Anzeige gespeichert wird.
' ===============================================================
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    ' Nichts erzwingen waehrend Save - nur sicherstellen, dass Events laufen
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub



' ===============================================================
' v4.5: Aenderungen auf dem Uebersicht-Blatt abfangen
' Wird ausgeloest wenn Nutzer einen Soll-Betrag manuell eintraegt
' v4.5b: Guard gegen Ausfuehrung waehrend GeneriereUebersicht
' ===============================================================
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    
    ' Nicht waehrend GeneriereUebersicht ausfuehren (Reentrancy-Schutz)
    If mod_Uebersicht_Generator.IsGenerating Then Exit Sub
    
    ' Nur auf dem Uebersicht-Blatt reagieren
    On Error Resume Next
    Dim wsName As String
    wsName = Sh.Name
    On Error GoTo 0
    
    If StrComp(wsName, WS_UEBERSICHT(), vbTextCompare) = 0 Then
        Call mod_Uebersicht_Event.VerarbeiteUebersichtAenderung(Target)
        ' v5.1: Dashboard bei manuellen Soll-Aenderungen aktualisieren
        On Error Resume Next
        Call mod_Uebersicht_Dashboard.GeneriereUebersichtNeu(stummModus:=True)
        On Error GoTo 0
    End If
    
End Sub


' ===============================================================
' Punkt 10: Snapshot der IST-Zelle bei Selektion auf Uebersicht
' Wird benoetigt fuer Rollback bei Abbruch der Datums-Inputbox
' ===============================================================
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    Dim shName As String
    shName = Sh.Name
    If StrComp(shName, WS_UEBERSICHT(), vbTextCompare) = 0 Then
        Call mod_Uebersicht_Event.MerkeAktuellenUebersichtZustand(Target)
    End If
End Sub



















