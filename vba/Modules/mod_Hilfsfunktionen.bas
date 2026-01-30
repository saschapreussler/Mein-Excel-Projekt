Attribute VB_Name = "mod_Hilfsfunktionen"
Option Explicit

' **********************************************************
' MODUL: mod_Hilfsfunktionen
' ZWECK: Generische Hilfsroutinen (Named Ranges, Listgenerierung)
' **********************************************************

Private Const TEMP_WS_NAME As String = "TEMP_LISTEN"

' **********************************************************
' PROZEDUR: AktualisiereNamedRange_MitgliederNamen
' Erstellt oder aktualisiert einen benannten Bereich
' mit den Namen aller aktiven Mitglieder.
' **********************************************************
Public Sub AktualisiereNamedRange_MitgliederNamen()
    
    Dim wsM As Worksheet
    Dim wsD As Worksheet
    Dim lastRow As Long
    Dim writeRow As Long
    Dim wasProtected As Boolean
    Dim rngTarget As Range
    Dim r As Long
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    
    ' TEMP_LISTEN falls vorhanden sofort entfernen
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(TEMP_WS_NAME).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    wasProtected = wsM.ProtectContents
    If wasProtected Then wsM.Unprotect PASSWORD:=PASSWORD
    
    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).row
    
    ' Zielbereich in Daten!AA (DATA_TEMP_COL_NAME) leeren
    wsD.Columns(DATA_TEMP_COL_NAME).ClearContents
    
    writeRow = 4 ' Startzeile in Daten
    
    If lastRow >= M_START_ROW Then
        For r = M_START_ROW To lastRow
            If Trim(wsM.Cells(r, M_COL_NACHNAME).value) <> "" Then
                If Trim(wsM.Cells(r, M_COL_PACHTENDE).value) = "" Then
                    wsD.Cells(writeRow, DATA_TEMP_COL_NAME).value = _
                        Trim(wsM.Cells(r, M_COL_NACHNAME).value) & ", " & Trim(wsM.Cells(r, M_COL_VORNAME).value)
                    writeRow = writeRow + 1
                End If
            End If
        Next r
    End If
    
    ' Named Range setzen
    On Error Resume Next
    ThisWorkbook.Names("rng_MitgliederNamen").Delete
    On Error GoTo 0
    
    If writeRow > 4 Then
        Set rngTarget = wsD.Range(wsD.Cells(4, DATA_TEMP_COL_NAME), wsD.Cells(writeRow - 1, DATA_TEMP_COL_NAME))
        ThisWorkbook.Names.Add Name:="rng_MitgliederNamen", RefersTo:=rngTarget
    End If
    
Cleanup:
    Application.ScreenUpdating = True
    If Not wsM Is Nothing Then
        If wasProtected Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler in AktualisiereNamedRange_MitgliederNamen: " & Err.Description, vbCritical
    Resume Cleanup

End Sub

' ***************************************************************
' HILFSFUNKTION: Prüfen, ob eine UserForm geladen ist (KORRIGIERT FÜR EXCEL)
' ***************************************************************
Private Function IsFormLoaded(ByVal FormName As String) As Boolean
    
    Dim i As Long
    
    ' Gehe alle geladenen UserForms in der VBA-Collection durch
    For i = 0 To VBA.UserForms.Count - 1
        ' Vergleiche den Klassennamen des geladenen Formulars mit dem gesuchten Namen
        If StrComp(VBA.UserForms.Item(i).Name, FormName, vbTextCompare) = 0 Then
            IsFormLoaded = True ' Formular gefunden und ist geladen
            Exit Function
        End If
    Next i
    
    IsFormLoaded = False ' Formular nicht gefunden
    
End Function

