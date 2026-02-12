VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ==========================================================
' TABELLENBLATT-MODUL: Wasser (Tabelle6) - BEREINIGT
' ==========================================================
Option Explicit

' ==========================================================
' EVENTS
' ==========================================================

' Wird bei jedem Wechsel auf das Blatt ausgeführt
Private Sub Worksheet_Activate()
    ' Zwangsberechnung beim Aktivieren des Blattes
    Call UpdateWasserblatt
End Sub

' Wird bei jeder Änderung einer Zelle auf dem Blatt ausgeführt
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Relevant: Stand Anfang (B) und Stand Ende (C) für Parzellen (Z. 10-23) und Hauptzähler (Z. 29)
    ' Wir prüfen den gesamten relevanten Bereich für Wasser: B10:C29
    If Not Intersect(Target, Me.Range("B10:C29")) Is Nothing Then
        Application.EnableEvents = False
        Call UpdateWasserblatt
        Application.EnableEvents = True
    End If
End Sub

' ==========================================================
' ZÄHLERWECHSEL STARTEN (KORRIGIERT: Direkter Aufruf der UserForm)
' ==========================================================
Private Sub btn_neuerZaehler_Wasser_Click()
    On Error GoTo ErrHandler
    
    ' Ruft die Public Sub InitForm_Runtime("Wasser") im UserForm 'frm_Zaehlerwechsel' auf
    ' und zeigt es an.
    With frm_Zaehlerwechsel
        Call .InitForm_Runtime("Wasser")
        .Show
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Fehler beim Öffnen des Zählerwechsel-Formulars: " & Err.Description & vbLf & _
           "Stellen Sie sicher, dass das UserForm den Code-Namen 'frm_Zaehlerwechsel' hat.", vbExclamation
End Sub

' ==========================================================
' HAUPT-UPDATE-ROUTINE
' ==========================================================
Public Sub UpdateWasserblatt()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    ' Führt Berechnung und Formatierung durch (in mod_ZaehlerLogik)
    Call mod_ZaehlerLogik.CalculateAllZaehlerVerbrauch(Me)
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "FEHLER BEI BERECHNUNG Wasser: " & Err.Description, vbCritical
End Sub

' ==========================================================
' OPTIONALE ROUTINEN (Historie-Farben)
' ==========================================================
Public Sub KorrigiereHistorieFarben()
    Call mod_ZaehlerLogik.FarbeHistorieEintraege
End Sub

