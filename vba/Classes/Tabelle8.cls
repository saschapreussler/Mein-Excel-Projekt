VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ***************************************************************
' WORKSHEET: Tabelle8 (Daten)
' ZWECK: Automatische Formatierung bei Datenaenderungen
' VERSION: 2.2 - 06.02.2026
' ***************************************************************

Private Sub Worksheet_Change(ByVal Target As Range)
    
    Dim changedCol As Long
    
    ' NEU: Dynamische Formatierung aufrufen
    Call mod_Formatierung.OnDatenChange(Target, Me)
    
    On Error GoTo ErrorHandler
    
    ' Nur Einzelzellen-Aenderungen verarbeiten
    If Target.Cells.CountLarge > 100 Then GoTo ErrorHandler
    If Target.Row < 4 Then GoTo ErrorHandler
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    changedCol = Target.Column
    
    ' === EINZELSPALTEN B, D, F, H ===
    If changedCol = 2 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 2)
        GoTo CleanUp
    End If
    
    If changedCol = 4 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 4)
        GoTo CleanUp
    End If
    
    If changedCol = 6 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 6)
        GoTo CleanUp
    End If
    
    If changedCol = 8 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 8)
        GoTo CleanUp
    End If
    
    ' === KATEGORIE-TABELLE J-P (10-16) ===
    If changedCol >= 10 And changedCol <= 16 Then
        Call mod_Formatierung.FormatKategorieTableComplete(Me)
        If changedCol = 10 Or changedCol = 11 Then
            Call mod_Formatierung.OnKategorieChange(Target)
        End If
        GoTo CleanUp
    End If
    
    ' === ENTITYKEY-TABELLE R-X (18-24) ===
    If changedCol >= 18 And changedCol <= 24 Then
        Call mod_Formatierung.FormatEntityKeyTableComplete(Me)
        If changedCol = 23 Then
            Call mod_EntityKey_Manager.VerarbeiteManuelleRoleAenderung(Target)
        End If
        GoTo CleanUp
    End If
    
    ' === HELPER-SPALTEN Z bis AH (26-34) ===
    If changedCol = 26 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 26)
        GoTo CleanUp
    End If
    
    If changedCol = 27 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 27)
        GoTo CleanUp
    End If
    
    If changedCol = 28 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 28)
        GoTo CleanUp
    End If
    
    If changedCol = 29 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 29)
        GoTo CleanUp
    End If
    
    If changedCol = 30 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 30)
        GoTo CleanUp
    End If
    
    If changedCol = 31 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 31)
        GoTo CleanUp
    End If
    
    If changedCol = 32 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 32)
        GoTo CleanUp
    End If
    
    If changedCol = 33 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 33)
        GoTo CleanUp
    End If
    
    If changedCol = 34 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 34)
        GoTo CleanUp
    End If

CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

