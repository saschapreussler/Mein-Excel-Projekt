VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ===============================================================
' TABELLENBLATT-MODUL: Einstellungen
' Formatiert die Zahlungstermin-Tabelle bei Aenderungen
' NEU v4.2: Loest Re-Evaluierung der Bankkonto-Kategorien aus
' ===============================================================

Private Sub Worksheet_Activate()
    Call mod_Einstellungen.FormatiereZahlungsterminTabelle(Me)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    
    ' Nur reagieren auf Aenderungen in Spalten B-H, ab Zeile 4
    If Intersect(Target, Me.Range(Me.Cells(ES_START_ROW, ES_COL_START), _
                                   Me.Cells(Me.Rows.count, ES_COL_END))) Is Nothing Then
        Exit Sub
    End If
    
    If Target.Row < ES_START_ROW Then Exit Sub
    
    Application.EnableEvents = False
    
    ' Formatierung aktualisieren
    Call mod_Einstellungen.FormatiereZahlungsterminTabelle(Me)
    
    ' NEU v4.2: Bankkonto-Kategorien neu evaluieren
    ' (Betrags-/Zeitfenster-Scoring hat sich geaendert)
    ' Manuelle Eingaben des Nutzers werden NICHT ueberschrieben
    Call mod_KategorieEngine_Pipeline.ReEvaluiereAlleNichtManuellen
    
    Application.EnableEvents = True
    
End Sub

