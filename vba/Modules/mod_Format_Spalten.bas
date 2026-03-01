Attribute VB_Name = "mod_Format_Spalten"
Option Explicit

' ***************************************************************
' MODUL: mod_Format_Spalten
' ZWECK: Einzelspalten-Formatierung, Zebra-Streifen, Lueckenentfernung
' ABGELEITET AUS: mod_Formatierung (Modularisierung)
' VERSION: 1.0 - 01.03.2026
' FUNKTIONEN:
'   - FormatiereAlleDatenSpalten: Alle 14 Einzel-Spalten formatieren
'   - FormatiereSingleSpalte: Zebra + Rahmen fuer eine Spalte
'   - FormatSingleColumnComplete: Public Wrapper fuer Einzelspalte
'   - VerdichteSpalteOhneLuecken: Leere Zeilen entfernen
' ***************************************************************

Private Const ZEBRA_COLOR_1 As Long = &HFFFFFF  ' Weiss
Private Const ZEBRA_COLOR_2 As Long = &HDEE5E3  ' Hellgrau

' ===============================================================
' Formatiert ALLE Einzel-Spalten auf Daten-Blatt
' ===============================================================
Public Sub FormatiereAlleDatenSpalten(ByRef ws As Worksheet)
    
    Call FormatiereSingleSpalte(ws, 2, True)   ' Spalte B - Vereinsfunktionen
    Call FormatiereSingleSpalte(ws, 4, True)   ' Spalte D - Anredeformen
    Call FormatiereSingleSpalte(ws, 6, True)   ' Spalte F - Parzelle
    Call FormatiereSingleSpalte(ws, 8, True)   ' Spalte H - Seite
    
    Call FormatiereSingleSpalte(ws, 26, True)  ' Spalte Z - Einnahme/Ausgabe
    Call FormatiereSingleSpalte(ws, 27, True)  ' Spalte AA - Prioritaet
    Call FormatiereSingleSpalte(ws, 28, True)  ' Spalte AB - Ja/Nein
    Call FormatiereSingleSpalte(ws, 29, True)  ' Spalte AC - Faelligkeit
    Call FormatiereSingleSpalte(ws, 30, True)  ' Spalte AD - EntityRole
    Call FormatiereSingleSpalte(ws, 31, True)  ' Spalte AE - Hilfszelle
    Call FormatiereSingleSpalte(ws, 32, True)  ' Spalte AF - Kat Einnahmen
    Call FormatiereSingleSpalte(ws, 33, True)  ' Spalte AG - Kat Ausgaben
    Call FormatiereSingleSpalte(ws, 34, True)  ' Spalte AH - Monat/Periode
    
End Sub

' ===============================================================
' Formatiert eine einzelne Spalte mit Zebra + Rahmen
' ===============================================================
Public Sub FormatiereSingleSpalte(ByRef ws As Worksheet, ByVal colIndex As Long, ByVal mitZebra As Boolean)
    
    Dim lastRow As Long
    Dim rng As Range
    Dim r As Long
    Dim cleanEnd As Long
    
    lastRow = ws.Cells(ws.Rows.count, colIndex).End(xlUp).Row
    
    ' Bereich UNTERHALB der belegten Zeilen bereinigen
    cleanEnd = lastRow + 50
    If cleanEnd < DATA_START_ROW + 50 Then cleanEnd = DATA_START_ROW + 50
    If lastRow < DATA_START_ROW Then
        ws.Range(ws.Cells(DATA_START_ROW, colIndex), ws.Cells(cleanEnd, colIndex)).Interior.ColorIndex = xlNone
        ws.Range(ws.Cells(DATA_START_ROW, colIndex), ws.Cells(cleanEnd, colIndex)).Borders.LineStyle = xlNone
        Exit Sub
    Else
        ws.Range(ws.Cells(lastRow + 1, colIndex), ws.Cells(cleanEnd, colIndex)).Interior.ColorIndex = xlNone
        ws.Range(ws.Cells(lastRow + 1, colIndex), ws.Cells(cleanEnd, colIndex)).Borders.LineStyle = xlNone
    End If
    
    Set rng = ws.Range(ws.Cells(DATA_START_ROW, colIndex), ws.Cells(lastRow, colIndex))
    
    rng.Interior.ColorIndex = xlNone
    rng.Borders.LineStyle = xlNone
    rng.VerticalAlignment = xlCenter
    
    If colIndex = 26 Or colIndex = 27 Then
        rng.HorizontalAlignment = xlCenter
    End If
    
    If mitZebra Then
        For r = DATA_START_ROW To lastRow
            If (r - DATA_START_ROW) Mod 2 = 0 Then
                ws.Cells(r, colIndex).Interior.color = ZEBRA_COLOR_1
            Else
                ws.Cells(r, colIndex).Interior.color = ZEBRA_COLOR_2
            End If
        Next r
    End If
    
    With rng
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeTop).color = RGB(0, 0, 0)
        
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).color = RGB(0, 0, 0)
        
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeLeft).color = RGB(0, 0, 0)
        
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeRight).color = RGB(0, 0, 0)
    End With
    
    If lastRow > DATA_START_ROW Then
        With rng.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = RGB(0, 0, 0)
        End With
    End If
    
    ws.Columns(colIndex).AutoFit
    
End Sub

' ===============================================================
' PUBLIC WRAPPER: Formatiert eine einzelne Spalte komplett
' ===============================================================
Public Sub FormatSingleColumnComplete(ByRef ws As Worksheet, ByVal colIndex As Long)
    
    Dim lastRow As Long
    Dim cleanEnd As Long
    Dim rngClean As Range
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    Call VerdichteSpalteOhneLuecken(ws, colIndex, colIndex, colIndex)
    
    Call FormatiereSingleSpalte(ws, colIndex, True)
    
    lastRow = ws.Cells(ws.Rows.count, colIndex).End(xlUp).Row
    If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW - 1
    
    cleanEnd = lastRow + 50
    If cleanEnd > ws.Rows.count Then cleanEnd = ws.Rows.count
    
    If lastRow + 1 <= cleanEnd Then
        Set rngClean = ws.Range(ws.Cells(lastRow + 1, colIndex), ws.Cells(cleanEnd, colIndex))
        rngClean.Interior.ColorIndex = xlNone
        rngClean.Borders.LineStyle = xlNone
    End If
    
    Call mod_Format_Protection.EntspeerreEditierbareSpalten(ws)
    
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
End Sub

' ===============================================================
' Entfernt Luecken (leere Zeilen) in einer Tabelle/Spalte
' ===============================================================
Public Sub VerdichteSpalteOhneLuecken(ByRef ws As Worksheet, ByVal checkCol As Long, _
                                       ByVal startCol As Long, ByVal endCol As Long)
    
    Dim lastRow As Long
    Dim schreibZeile As Long
    Dim leseZeile As Long
    Dim numCols As Long
    Dim col As Long
    
    Dim maxRow As Long
    maxRow = DATA_START_ROW - 1
    For col = startCol To endCol
        lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
        If lastRow > maxRow Then maxRow = lastRow
    Next col
    
    If maxRow < DATA_START_ROW Then Exit Sub
    
    lastRow = maxRow
    numCols = endCol - startCol + 1
    
    Dim arrData() As Variant
    Dim arrResult() As Variant
    Dim totalRows As Long
    Dim resultCount As Long
    Dim isEmpty As Boolean
    
    totalRows = lastRow - DATA_START_ROW + 1
    ReDim arrData(1 To totalRows, 1 To numCols)
    ReDim arrResult(1 To totalRows, 1 To numCols)
    
    Dim r As Long, c As Long
    For r = 1 To totalRows
        For c = 1 To numCols
            arrData(r, c) = ws.Cells(DATA_START_ROW + r - 1, startCol + c - 1).value
        Next c
    Next r
    
    resultCount = 0
    Dim checkIdx As Long
    checkIdx = checkCol - startCol + 1
    
    For r = 1 To totalRows
        If Trim(CStr(arrData(r, checkIdx))) <> "" Then
            resultCount = resultCount + 1
            For c = 1 To numCols
                arrResult(resultCount, c) = arrData(r, c)
            Next c
        End If
    Next r
    
    If resultCount > 0 Then
        ws.Range(ws.Cells(DATA_START_ROW, startCol), _
                 ws.Cells(lastRow, endCol)).ClearContents
        
        For r = 1 To resultCount
            For c = 1 To numCols
                ws.Cells(DATA_START_ROW + r - 1, startCol + c - 1).value = arrResult(r, c)
            Next c
        Next r
    Else
        ws.Range(ws.Cells(DATA_START_ROW, startCol), _
                 ws.Cells(lastRow, endCol)).ClearContents
    End If
    
End Sub
