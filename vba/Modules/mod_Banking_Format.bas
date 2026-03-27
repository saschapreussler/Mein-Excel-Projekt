Attribute VB_Name = "mod_Banking_Format"
Option Explicit

' ===============================================================
' MODUL: mod_Banking_Format
' Ausgelagert aus mod_Banking_Data
' Enth?lt: Zebra-Formatierung, Rahmen, allgemeine Formatierung,
'          Sortierung Bankkonto, Formel-Wiederherstellung
' ===============================================================

Private Const ZEBRA_COLOR As Long = &HDEE5E3


' ===============================================================
' ZEBRA-FORMATIERUNG (A-G und I-Z, Spalte H ausgenommen)
' ===============================================================
Public Sub Anwende_Zebra_Bankkonto(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim lRow As Long
    Dim rngPart1 As Range
    Dim rngPart2 As Range
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    For lRow = BK_START_ROW To lastRow
        Set rngPart1 = ws.Range(ws.Cells(lRow, 1), ws.Cells(lRow, 7))
        Set rngPart2 = ws.Range(ws.Cells(lRow, 9), ws.Cells(lRow, 26))
        
        If (lRow - BK_START_ROW) Mod 2 = 1 Then
            rngPart1.Interior.color = ZEBRA_COLOR
            rngPart2.Interior.color = ZEBRA_COLOR
        Else
            rngPart1.Interior.ColorIndex = xlNone
            rngPart2.Interior.ColorIndex = xlNone
        End If
    Next lRow
    
End Sub

' ===============================================================
' RAHMEN-FORMATIERUNG
' ===============================================================
Public Sub Anwende_Border_Bankkonto(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim rngPart1 As Range
    Dim rngPart2 As Range
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    Set rngPart1 = ws.Range(ws.Cells(BK_START_ROW, 1), ws.Cells(lastRow, 12))
    Set rngPart2 = ws.Range(ws.Cells(BK_START_ROW, 13), ws.Cells(lastRow, 26))
    
    Call SetBorders(rngPart1)
    Call SetBorders(rngPart2)
    
End Sub

Private Sub SetBorders(ByVal rng As Range)
    
    If rng Is Nothing Then Exit Sub
    
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
End Sub

' ===============================================================
' ALLGEMEINE FORMATIERUNG
' ===============================================================
Public Sub Anwende_Formatierung_Bankkonto(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim euroFormat As String
    
    If ws Is Nothing Then Exit Sub
    
    euroFormat = "#,##0.00 " & ChrW(8364)
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    ' Spalte B (Betrag): W?hrung + rechtsb?ndig
    With ws.Range(ws.Cells(BK_START_ROW, BK_COL_BETRAG), ws.Cells(lastRow, BK_COL_BETRAG))
        .NumberFormat = euroFormat
        .HorizontalAlignment = xlRight
    End With
    
    ' Spalten M-Z: W?hrung
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_MITGL_BEITR), ws.Cells(lastRow, BK_COL_AUSZAHL_KASSE)).NumberFormat = euroFormat
    
    With ws.Range(ws.Cells(BK_START_ROW, BK_COL_BEMERKUNG), ws.Cells(lastRow, BK_COL_BEMERKUNG))
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    
    ws.Cells.VerticalAlignment = xlCenter
    ws.Rows(BK_START_ROW & ":" & lastRow).AutoFit
    
End Sub


' ===============================================================
' SORTIERUNG NACH DATUM (AUFSTEIGEND - Januar oben)
' ===============================================================
Public Sub Sortiere_Bankkonto_nach_Datum()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim sortRange As Range
    
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        Exit Sub
    End If
    
    Set sortRange = ws.Range(ws.Cells(BK_START_ROW, 1), ws.Cells(lastRow, 26))
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(BK_START_ROW, BK_COL_DATUM), ws.Cells(lastRow, BK_COL_DATUM)), _
                           SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    With ws.Sort
        .SetRange sortRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
End Sub


' ===============================================================
' FORMEL-WIEDERHERSTELLUNG
' Stellt die Formeln auf dem Bankkonto-Blatt wieder her,
' die durch ClearContents oder Import verloren gehen k?nnen.
' Betrifft: C3, E8-E14, E16-E21, E23
' WICHTIG: Formeln werden 1:1 als FormulaLocal gesetzt!
' ===============================================================
Public Sub StelleFormelnWiederHer(ByVal ws As Worksheet)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    On Error Resume Next
    
    ' C3: Kontostand-Anzeige mit Monatsfilter
    ws.Range("C3").FormulaLocal = _
        "=WENN(Daten!$AE$4=0;WENN(ANZAHL(Bankkonto!$A$28:$A$3433)=0;"""";" & _
        """Kontostand nach der letzten Buchung im Monat am: "" & TEXT(MAX(Bankkonto!$A$28:$A$5000);""TT.MM.JJJJ""));" & _
        "WENN(Z" & ChrW(196) & "HLENWENNS(Bankkonto!$A$28:$A$5000;"">="" & DATUM(Startmen" & ChrW(252) & "!$F$1;Daten!$AE$4;1);" & _
        "Bankkonto!$A$28:$A$5000;""<="" & DATUM(Startmen" & ChrW(252) & "!$F$1;Daten!$AE$4+1;0))=0;"""";" & _
        """Kontostand nach der letzten Buchung im Monat am: "" & TEXT(MAXWENNS(Bankkonto!$A$28:$A$5000;" & _
        "Bankkonto!$A$28:$A$5000;"">="" & DATUM(Startmen" & ChrW(252) & "!$F$1;Daten!$AE$4;1);" & _
        "Bankkonto!$A$28:$A$5000;""<="" & DATUM(Startmen" & ChrW(252) & "!$F$1;Daten!$AE$4+1;0));""TT.MM.JJJJ""))))"
    
    ' E8-E14: Einnahmen (Spalten M-S) mit SUMMEWENNS + WENN=0 leer
    ws.Range("E8").FormulaLocal = _
        "=WENN(SUMMEWENNS(M28:M5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(M28:M5000;G28:G5000;WAHR))"
    ws.Range("E9").FormulaLocal = _
        "=WENN(SUMMEWENNS(N28:N5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(N28:N5000;G28:G5000;WAHR))"
    ws.Range("E10").FormulaLocal = _
        "=WENN(SUMMEWENNS(O28:O5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(O28:O5000;G28:G5000;WAHR))"
    ws.Range("E11").FormulaLocal = _
        "=WENN(SUMMEWENNS(P28:P5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(P28:P5000;G28:G5000;WAHR))"
    ws.Range("E12").FormulaLocal = _
        "=WENN(SUMMEWENNS(Q28:Q5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(Q28:Q5000;G28:G5000;WAHR))"
    ws.Range("E13").FormulaLocal = _
        "=WENN(SUMMEWENNS(R28:R5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(R28:R5000;G28:G5000;WAHR))"
    ws.Range("E14").FormulaLocal = _
        "=WENN(SUMMEWENNS(S28:S5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(S28:S5000;G28:G5000;WAHR))"
    
    ' E16-E21: Ausgaben (Spalten T-Y) mit SUMMEWENNS + WENN=0 leer
    ws.Range("E16").FormulaLocal = _
        "=WENN(SUMMEWENNS(T28:T5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(T28:T5000;G28:G5000;WAHR))"
    ws.Range("E17").FormulaLocal = _
        "=WENN(SUMMEWENNS(U28:U5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(U28:U5000;G28:G5000;WAHR))"
    ws.Range("E18").FormulaLocal = _
        "=WENN(SUMMEWENNS(V28:V5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(V28:V5000;G28:G5000;WAHR))"
    ws.Range("E19").FormulaLocal = _
        "=WENN(SUMMEWENNS(W28:W5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(W28:W5000;G28:G5000;WAHR))"
    ws.Range("E20").FormulaLocal = _
        "=WENN(SUMMEWENNS(X28:X5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(X28:X5000;G28:G5000;WAHR))"
    ws.Range("E21").FormulaLocal = _
        "=WENN(SUMMEWENNS(Y28:Y5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(Y28:Y5000;G28:G5000;WAHR))"
    
    ' E23: Auszahlung Kasse (Spalte Z)
    ws.Range("E23").FormulaLocal = _
        "=WENN(SUMMEWENNS(Z28:Z5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(Z28:Z5000;G28:G5000;WAHR))"
    
    On Error GoTo 0
    
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
End Sub






















































