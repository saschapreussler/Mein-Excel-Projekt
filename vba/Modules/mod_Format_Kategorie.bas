Attribute VB_Name = "mod_Format_Kategorie"
Option Explicit

' ***************************************************************
' MODUL: mod_Format_Kategorie
' ZWECK: Kategorie-Tabelle (J-P) Formatierung, Sortierung, Zielspalte
' ABGELEITET AUS: mod_Formatierung (Modularisierung)
' VERSION: 1.0 - 01.03.2026
' FUNKTIONEN:
'   - FormatiereKategorieTabelle: Zebra + Rahmen fuer J-P
'   - SortiereKategorieTabelle: Sortierung nach Spalte J
'   - SetzeZielspalteDropdown: DropDown fuer Zielspalte N
' ***************************************************************

Private Const ZEBRA_COLOR_1 As Long = &HFFFFFF  ' Weiss
Private Const ZEBRA_COLOR_2 As Long = &HDEE5E3  ' Hellgrau

' ===============================================================
' Formatiert die Kategorie-Tabelle (Spalten J-P)
' ===============================================================
Public Sub FormatiereKategorieTabelle(Optional ByRef ws As Worksheet = Nothing)
    
    Dim lastRow As Long
    Dim lastRowMax As Long
    Dim rngTable As Range
    Dim rngLeeren As Range
    Dim r As Long
    Dim einAusWert As String
    
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    
    ' Maximale letzte Zeile ueber alle Spalten J-P ermitteln
    lastRowMax = lastRow
    Dim col As Long
    For col = DATA_CAT_COL_START To DATA_CAT_COL_END
        Dim colLastRow As Long
        colLastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
        If colLastRow > lastRowMax Then lastRowMax = colLastRow
    Next col
    
    ' Bereich UNTERHALB bereinigen
    If lastRowMax >= DATA_START_ROW Then
        Dim cleanStart As Long
        If lastRow < DATA_START_ROW Then
            cleanStart = DATA_START_ROW
        Else
            cleanStart = lastRow + 1
        End If
        
        If cleanStart <= lastRowMax + 50 Then
            Set rngLeeren = ws.Range(ws.Cells(cleanStart, DATA_CAT_COL_START), _
                                     ws.Cells(lastRowMax + 50, DATA_CAT_COL_END))
            rngLeeren.Interior.ColorIndex = xlNone
            rngLeeren.Borders.LineStyle = xlNone
        End If
    End If
    
    If lastRow < DATA_START_ROW Then Exit Sub
    
    Set rngTable = ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), _
                            ws.Cells(lastRow, DATA_CAT_COL_END))
    
    rngTable.Interior.ColorIndex = xlNone
    rngTable.Borders.LineStyle = xlNone
    
    ' Zebra-Formatierung
    For r = DATA_START_ROW To lastRow
        If (r - DATA_START_ROW) Mod 2 = 0 Then
            ws.Range(ws.Cells(r, DATA_CAT_COL_START), ws.Cells(r, DATA_CAT_COL_END)).Interior.color = ZEBRA_COLOR_1
        Else
            ws.Range(ws.Cells(r, DATA_CAT_COL_START), ws.Cells(r, DATA_CAT_COL_END)).Interior.color = ZEBRA_COLOR_2
        End If
    Next r
    
    ' Rahmenlinien
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    rngTable.VerticalAlignment = xlCenter
    
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KATEGORIE), _
             ws.Cells(lastRow, DATA_CAT_COL_KATEGORIE)).HorizontalAlignment = xlLeft
    
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_EINAUS), _
             ws.Cells(lastRow, DATA_CAT_COL_EINAUS)).HorizontalAlignment = xlCenter
    
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KEYWORD), _
             ws.Cells(lastRow, DATA_CAT_COL_KEYWORD)).HorizontalAlignment = xlLeft
    
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_PRIORITAET), _
             ws.Cells(lastRow, DATA_CAT_COL_PRIORITAET)).HorizontalAlignment = xlCenter
    
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_ZIELSPALTE), _
             ws.Cells(lastRow, DATA_CAT_COL_ZIELSPALTE)).HorizontalAlignment = xlLeft
    
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_FAELLIGKEIT), _
             ws.Cells(lastRow, DATA_CAT_COL_FAELLIGKEIT)).HorizontalAlignment = xlLeft
    
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KOMMENTAR), _
             ws.Cells(lastRow, DATA_CAT_COL_KOMMENTAR)).HorizontalAlignment = xlLeft
    
    For r = DATA_START_ROW To lastRow
        einAusWert = UCase(Trim(ws.Cells(r, DATA_CAT_COL_EINAUS).value))
        Call SetzeZielspalteDropdown(ws, r, einAusWert)
    Next r
    
    ' Spaltenbreiten
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KATEGORIE), _
             ws.Cells(lastRow, DATA_CAT_COL_KATEGORIE)).EntireColumn.AutoFit
    
    ws.Columns(DATA_CAT_COL_EINAUS).ColumnWidth = 12
    
    Dim autoFitCol As Long
    For autoFitCol = DATA_CAT_COL_KEYWORD To DATA_CAT_COL_END
        ws.Columns(autoFitCol).AutoFit
    Next autoFitCol
    
    ' Keyword-Spalte L
    If lastRow >= DATA_START_ROW Then
        ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KEYWORD), _
                 ws.Cells(lastRow, DATA_CAT_COL_KEYWORD)).ShrinkToFit = False
        ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KEYWORD), _
                 ws.Cells(lastRow, DATA_CAT_COL_KEYWORD)).WrapText = False
        ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KEYWORD), _
                 ws.Cells(lastRow, DATA_CAT_COL_KEYWORD)).Font.color = vbBlack
    End If
    
End Sub

' ===============================================================
' Sortiert die Kategorie-Tabelle nach Spalte J (A-Z)
' ===============================================================
Public Sub SortiereKategorieTabelle(Optional ByRef ws As Worksheet = Nothing)
    
    Dim lastRow As Long
    Dim sortRange As Range
    
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    
    Set sortRange = ws.Range(ws.Cells(DATA_START_ROW - 1, DATA_CAT_COL_START), _
                             ws.Cells(lastRow, DATA_CAT_COL_END))
    
    On Error Resume Next
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KATEGORIE), _
                                      ws.Cells(lastRow, DATA_CAT_COL_KATEGORIE)), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SetRange sortRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    On Error GoTo 0
    
End Sub

' ===============================================================
' ZIELSPALTE-DROPDOWN SETZEN (abhaengig von E/A)
' ===============================================================
Public Sub SetzeZielspalteDropdown(ByRef ws As Worksheet, ByVal zeile As Long, ByVal einAus As String)
    
    Dim dropdownSource As String
    
    On Error Resume Next
    ws.Cells(zeile, DATA_CAT_COL_ZIELSPALTE).Validation.Delete
    On Error GoTo 0
    
    Select Case einAus
        Case "E"
            dropdownSource = "=" & WS_BANKKONTO & "!$M$27:$S$27"
        Case "A"
            dropdownSource = "=" & WS_BANKKONTO & "!$T$27:$Z$27"
        Case Else
            dropdownSource = "=" & WS_BANKKONTO & "!$M$27:$Z$27"
    End Select
    
    On Error Resume Next
    With ws.Cells(zeile, DATA_CAT_COL_ZIELSPALTE).Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:=dropdownSource
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    On Error GoTo 0
    
End Sub


































































