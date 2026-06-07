Attribute VB_Name = "mod_Zeilenhoehen"
Option Explicit

' ===============================================================
' Modul: mod_Zeilenhoehen
' v8.0: Setzt definierte Zeilenhoehen auf den Hauptblaettern.
'       Wird beim Workbook_Open aufgerufen, kann auch manuell ausgeloest werden.
'
' Spezifikation:
'   Zahlungsuebersicht          : Zeile 2 = 35.00 (70 px)
'   Dashboard Mitgliederzahlungen: Zeile 2 = 35.00, Zeile 3 = 18.50, Zeile 4 = 18.50
'   Bankkonto                   : Zeile 2 unveraendert
'   Vereinskasse                : Zeile 1 = 50.00, Zeile 2 = 20.00, Zeile 3 = 20.00
'   Strom + Wasser              : Zeile 2 unveraendert
'   Mitgliederliste             : Zeile 2 = 18.00
'   Mitgliederhistorie + Einstellungen: unveraendert
'   Daten                       : Zeile 2 = 24.00
' ===============================================================

Public Sub FixiereStandardZeilenhoehen()
    On Error Resume Next
    
    Dim eventsAlt As Boolean
    eventsAlt = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Call SetzeZeilenhoehe(WS_UEBERSICHT(), 2, 35#)
    
    Call SetzeZeilenhoehe("Dashboard Mitgliederzahlungen", 2, 35#)
    Call SetzeZeilenhoehe("Dashboard Mitgliederzahlungen", 3, 18.5)
    Call SetzeZeilenhoehe("Dashboard Mitgliederzahlungen", 4, 18.5)
    
    Call SetzeZeilenhoehe(WS_VEREINSKASSE, 1, 50#)
    Call SetzeZeilenhoehe(WS_VEREINSKASSE, 2, 20#)
    Call SetzeZeilenhoehe(WS_VEREINSKASSE, 3, 20#)
    
    Call SetzeZeilenhoehe("Mitgliederliste", 2, 18#)
    
    Call SetzeZeilenhoehe(WS_DATEN, 2, 24#)
    
    Application.ScreenUpdating = True
    Application.EnableEvents = eventsAlt
End Sub

Private Sub SetzeZeilenhoehe(ByVal blattName As String, ByVal zeile As Long, ByVal hoehe As Double)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(blattName)
    If ws Is Nothing Then Exit Sub
    
    Dim warGeschuetzt As Boolean
    warGeschuetzt = ws.ProtectContents
    If warGeschuetzt Then ws.Unprotect PASSWORD:=PASSWORD
    
    ws.Rows(zeile).RowHeight = hoehe
    
    If warGeschuetzt Then ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
End Sub











