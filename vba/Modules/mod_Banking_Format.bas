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
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
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
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    
End Sub


' ===============================================================
' FORMEL-WIEDERHERSTELLUNG
' Stellt die Formeln auf dem Bankkonto-Blatt wieder her,
' die durch ClearContents oder Import verloren gehen k?nnen.
' Betrifft: E4, C5, E10-E16, E18-E23, E25
' WICHTIG: Formeln werden 1:1 als FormulaLocal gesetzt!
' ===============================================================
Public Sub StelleFormelnWiederHer(ByVal ws As Worksheet)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    On Error Resume Next
    
    ' E4: Kontostand laufend mit Monatsfilter
    ' v6.1: Abrechnungsjahr aus Einstellungen!C6, Kontostand aus Einstellungen!C7
    ' Logik: Wenn Monat<=1 (ganzes Jahr/Jan) -> nur Kontostand Vorjahr
    '        Sonst: Kontostand Vorjahr + Summe aller Buchungen von Jan bis Filtermonat
    ws.Range("E4").FormulaLocal = _
        "=WENN(Daten!$AE$4<=1;Einstellungen!$C$7;" & _
        "Einstellungen!$C$7+SUMMEWENNS(Bankkonto!$B$30:$B$5000;" & _
        "Bankkonto!$A$30:$A$5000;"">=""&DATUM(Einstellungen!$C$6;1;1);" & _
        "Bankkonto!$A$30:$A$5000;""<""&DATUM(Einstellungen!$C$6;Daten!$AE$4;1)))"
    
    ' C5: Kontostand-Anzeige mit Monatsfilter
    ws.Range("C5").FormulaLocal = _
        "=WENN(Daten!$AE$4=0;WENN(ANZAHL(Bankkonto!$A$30:$A$3433)=0;"""";" & _
        """Kontostand nach der letzten Buchung im Monat am: "" & TEXT(MAX(Bankkonto!$A$30:$A$5000);""TT.MM.JJJJ""));" & _
        "WENN(Z" & ChrW(196) & "HLENWENNS(Bankkonto!$A$30:$A$5000;"">="" & DATUM(Einstellungen!$C$6;Daten!$AE$4;1);" & _
        "Bankkonto!$A$30:$A$5000;""<="" & DATUM(Einstellungen!$C$6;Daten!$AE$4+1;0))=0;"""";" & _
        """Kontostand nach der letzten Buchung im Monat am: "" & TEXT(MAXWENNS(Bankkonto!$A$30:$A$5000;" & _
        "Bankkonto!$A$30:$A$5000;"">="" & DATUM(Einstellungen!$C$6;Daten!$AE$4;1);" & _
        "Bankkonto!$A$30:$A$5000;""<="" & DATUM(Einstellungen!$C$6;Daten!$AE$4+1;0));""TT.MM.JJJJ""))))"
    
    ' E10-E16: Einnahmen (Spalten M-S) mit SUMMEWENNS + WENN=0 leer
    ws.Range("E10").FormulaLocal = _
        "=WENN(SUMMEWENNS(M30:M5000;G30:G5000;WAHR)=0;"""";SUMMEWENNS(M30:M5000;G30:G5000;WAHR))"
    ws.Range("E11").FormulaLocal = _
        "=WENN(SUMMEWENNS(N30:N5000;G30:G5000;WAHR)=0;"""";SUMMEWENNS(N30:N5000;G30:G5000;WAHR))"
    ws.Range("E12").FormulaLocal = _
        "=WENN(SUMMEWENNS(O30:O5000;G30:G5000;WAHR)=0;"""";SUMMEWENNS(O30:O5000;G30:G5000;WAHR))"
    ws.Range("E13").FormulaLocal = _
        "=WENN(SUMMEWENNS(P30:P5000;G30:G5000;WAHR)=0;"""";SUMMEWENNS(P30:P5000;G30:G5000;WAHR))"
    ws.Range("E14").FormulaLocal = _
        "=WENN(SUMMEWENNS(Q30:Q5000;G30:G5000;WAHR)=0;"""";SUMMEWENNS(Q30:Q5000;G30:G5000;WAHR))"
    ws.Range("E15").FormulaLocal = _
        "=WENN(SUMMEWENNS(R30:R5000;G30:G5000;WAHR)=0;"""";SUMMEWENNS(R30:R5000;G30:G5000;WAHR))"
    ws.Range("E16").FormulaLocal = _
        "=WENN(SUMMEWENNS(S30:S5000;G30:G5000;WAHR)=0;"""";SUMMEWENNS(S30:S5000;G30:G5000;WAHR))"
    
    ' E18-E23: Ausgaben (Spalten T-Y) mit SUMMEWENNS + WENN=0 leer
    ws.Range("E18").FormulaLocal = _
        "=WENN(SUMMEWENNS(T30:T5000;G30:G5000;WAHR)=0;"""";SUMMEWENNS(T30:T5000;G30:G5000;WAHR))"
    ws.Range("E19").FormulaLocal = _
        "=WENN(SUMMEWENNS(U30:U5000;G30:G5000;WAHR)=0;"""";SUMMEWENNS(U30:U5000;G30:G5000;WAHR))"
    ws.Range("E20").FormulaLocal = _
        "=WENN(SUMMEWENNS(V30:V5000;G30:G5000;WAHR)=0;"""";SUMMEWENNS(V30:V5000;G30:G5000;WAHR))"
    ws.Range("E21").FormulaLocal = _
        "=WENN(SUMMEWENNS(W30:W5000;G30:G5000;WAHR)=0;"""";SUMMEWENNS(W30:W5000;G30:G5000;WAHR))"
    ws.Range("E22").FormulaLocal = _
        "=WENN(SUMMEWENNS(X30:X5000;G30:G5000;WAHR)=0;"""";SUMMEWENNS(X30:X5000;G30:G5000;WAHR))"
    ws.Range("E23").FormulaLocal = _
        "=WENN(SUMMEWENNS(Y30:Y5000;G30:G5000;WAHR)=0;"""";SUMMEWENNS(Y30:Y5000;G30:G5000;WAHR))"
    
    ' E25: Auszahlung Kasse (Spalte Z)
    ws.Range("E25").FormulaLocal = _
        "=WENN(SUMMEWENNS(Z30:Z5000;G30:G5000;WAHR)=0;"""";SUMMEWENNS(Z30:Z5000;G30:G5000;WAHR))"
    
    On Error GoTo 0
    
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0
    
End Sub






















































































