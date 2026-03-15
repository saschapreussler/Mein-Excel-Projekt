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
    
    ' Stelle sicher dass alle Bl?tter vertikal zentriert sind
    Call mod_Formatierung.ZentriereAlleZellenVertikal
    
    ' NEU v5.3: Hilfsspalten auf Daten-Blatt ausblenden
    Call mod_Format_Protection.BlendeDatenSpaltenAus

    ' Mitgliederliste: DropDowns + Spaltenschutz aktualisieren
    Call mod_Mitglieder_UI.ApplyMitgliederDropdowns
    
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    ' Neues Blatt automatisch vertikal zentrieren
    On Error Resume Next
    Call mod_Formatierung.FormatiereNeuesBlatt(Sh)
    On Error GoTo 0
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
    End If
    
End Sub














