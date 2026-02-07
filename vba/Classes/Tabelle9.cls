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
    
    ' Wenn eine Datenzeile komplett geleert wurde (Delete-Taste),
    ' wird sie beim Neu-Formatieren automatisch entfernt
    Call mod_Einstellungen.FormatiereZahlungsterminTabelle(Me)
    
    Application.EnableEvents = True
    
End Sub


