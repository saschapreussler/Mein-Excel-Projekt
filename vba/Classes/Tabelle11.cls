VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Haelt Zebra/AutoFit/Ausblenden der Mitgliederhistorie nach manuellen
' Aenderungen konsistent.
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Fehler
    If Intersect(Target, Me.Range("A:J")) Is Nothing Then Exit Sub
    If Target.Row < H_START_ROW Then Exit Sub

    Application.EnableEvents = False
    Call mod_Formatierung.FormatiereMitgliederhistorie

Fehler:
    Application.EnableEvents = True
End Sub
