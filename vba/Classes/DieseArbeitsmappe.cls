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
    
    ' v6.0: Startseite als erstes Blatt anzeigen
    On Error Resume Next
    ThisWorkbook.Worksheets(WS_STARTMENUE()).Activate
    On Error GoTo 0
    
    ' v7.0: AutoFilter ZULETZT aufrufen, damit AllowFiltering
    '        nicht durch vorherige Protect-Aufrufe verloren geht
    On Error Resume Next
    Call mod_Formatierung.StelleAutoFilterBereit
    On Error GoTo 0
    
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    ' Neues Blatt automatisch vertikal zentrieren
    On Error Resume Next
    Call mod_Formatierung.FormatiereNeuesBlatt(Sh)
    On Error GoTo 0
End Sub


' ===============================================================
' SheetActivate: Gitternetzlinien je Blatt korrekt setzen
' Startmen?, Finanz-?bersicht und Dashboard ohne Gitternetz,
' alle anderen Bl?tter (BK, VK, Strom, Wasser, ?bersicht etc.) MIT Gitternetz
' ===============================================================
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    Dim shName As String
    shName = Sh.Name
    
    Dim ohneGitter As Boolean
    ohneGitter = False
    
    Select Case shName
        Case WS_STARTMENUE(), WS_FINANZ_UEBERSICHT(), "Dashboard Mitgliederzahlungen"
            ohneGitter = True
    End Select
    
    If ActiveWindow Is Nothing Then Exit Sub
    ActiveWindow.DisplayGridlines = Not ohneGitter
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














































