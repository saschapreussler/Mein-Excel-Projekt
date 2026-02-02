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

    Call Initialize_ImportReport_ListBox
    
    ' Bereinige eventuell verwaiste TEMP Worksheets
    On Error Resume Next
    Call mod_Hilfsfunktionen.BereinigeTempWorksheets
    On Error GoTo 0
    
    ' Stelle sicher dass alle Blaetter vertikal zentriert sind
    Call mod_Formatierung.ZentriereAlleZellenVertikal

End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    ' Neues Blatt automatisch vertikal zentrieren
    On Error Resume Next
    Call mod_Formatierung.FormatiereNeuesBlatt(Sh)
    On Error GoTo 0
End Sub
