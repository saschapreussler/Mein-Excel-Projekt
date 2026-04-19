Attribute VB_Name = "mod_FinanzUebersicht"
Option Explicit

' ===============================================================
' MODUL: mod_FinanzUebersicht
' VERSION: 2.0 - 21.04.2026
' ZWECK: Erstellt und pflegt das Blatt "Finanz-Uebersicht"
'        - Kategorien dynamisch aus Bankkonto Spalte H
'        - Sammelzahlungen: Aufschluesselung via Spalte L (Bemerkung)
'        - KPIs: Einnahmen, Ausgaben, Saldo, Kontostand, VK-Saldo
'        - Monatsfilter via DropDown
'        - Balkendiagramme Einnahmen / Ausgaben
' ===============================================================

' --- Farben ---
Private Const CLR_HEADER As Long = 2763306      ' RGB(26, 35, 42)
Private Const CLR_ACCENT As Long = 14521384     ' RGB(40, 167, 221)
Private Const CLR_WHITE As Long = 16777215
Private Const CLR_LIGHT_BG As Long = 15921906   ' RGB(242, 242, 242)
Private Const CLR_EINN As Long = 2573097        ' RGB(41, 69, 39)
Private Const CLR_EINN_LIGHT As Long = 14348258 ' RGB(226, 240, 217)
Private Const CLR_AUSG As Long = 4743219        ' RGB(163, 80, 72)
Private Const CLR_AUSG_LIGHT As Long = 13688301 ' RGB(237, 220, 209)
Private Const CLR_DARK_TEXT As Long = 2500134    ' RGB(38, 50, 56)
Private Const CLR_SUM_BG As Long = 14408667     ' RGB(219, 223, 219)
Private Const CLR_VK As Long = 7168108          ' RGB(108, 117, 109)

Private Const FILTER_DD_NAME As String = "dd_MonatFilter_FU"
Private Const KAT_SAMMELZAHLUNG As String = "Sammelzahlung"


' ===============================================================
' HAUPTPROZEDUR: Finanz-Uebersicht erstellen/aktualisieren
' ===============================================================
Public Sub ErstelleFinanzUebersicht()
    Dim ws As Worksheet
    Set ws = HoleOderErstelleBlatt()
    If ws Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    Call EntferneAlleObjekte(ws)
    Call BaueFinanzUebersicht(ws, 0)
    
    ws.Cells.Locked = True
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    
    ws.Activate
    ws.Range("A1").Select
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


' ===============================================================
' BLATT FINDEN ODER ERSTELLEN
' ===============================================================
Private Function HoleOderErstelleBlatt() As Worksheet
    Dim ws As Worksheet
    Dim blattName As String
    blattName = WS_FINANZ_UEBERSICHT()
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(blattName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        On Error GoTo ErrCreate
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = blattName
    End If
    
    Set HoleOderErstelleBlatt = ws
    Exit Function

ErrCreate:
    MsgBox "Blatt '" & blattName & "' konnte nicht erstellt werden." & vbLf & _
           Err.Description, vbExclamation, "Finanz-" & ChrW(220) & "bersicht"
    Set HoleOderErstelleBlatt = Nothing
End Function


' ===============================================================
' ALLE SHAPES/CHARTS/DROPDOWNS ENTFERNEN
' ===============================================================
Private Sub EntferneAlleObjekte(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        On Error Resume Next
        ws.Shapes(i).Delete
        Err.Clear
        On Error GoTo 0
    Next i
    
    Dim cht As ChartObject
    For Each cht In ws.ChartObjects
        On Error Resume Next
        cht.Delete
        Err.Clear
        On Error GoTo 0
    Next cht
End Sub


' ===============================================================
' KERNFUNKTION: Daten sammeln und Blatt aufbauen
' Wird bei Erstellung und bei Filterwechsel aufgerufen
' ===============================================================
Private Sub BaueFinanzUebersicht(ByVal ws As Worksheet, ByVal monatFilter As Long)
    
    ' --- 1. Daten sammeln ---
    Dim dictEinn As Object
    Dim dictAusg As Object
    Set dictEinn = CreateObject("Scripting.Dictionary")
    Set dictAusg = CreateObject("Scripting.Dictionary")
    
    Call SammleDaten(dictEinn, dictAusg, monatFilter)
    
    ' Sortierte Arrays erstellen
    Dim arrEinnKat() As String, arrEinnVal() As Double, cntEinn As Long
    Dim arrAusgKat() As String, arrAusgVal() As Double, cntAusg As Long
    
    Call DictToSortedArrays(dictEinn, arrEinnKat, arrEinnVal, cntEinn)
    Call DictToSortedArrays(dictAusg, arrAusgKat, arrAusgVal, cntAusg)
    
    ' --- 2. Blatt vorbereiten ---
    Dim maxRows As Long
    maxRows = 40 + cntEinn + cntAusg
    ws.Range("A1:K" & maxRows).ClearContents
    ws.Range("A1:K" & maxRows).ClearFormats
    ws.Range("A1:K" & maxRows).Interior.color = CLR_WHITE
    
    Dim wnd As Window
    For Each wnd In Application.Windows
        If wnd.Caption = ThisWorkbook.Name Then
            wnd.DisplayGridlines = False
        End If
    Next wnd
    
    ' Spaltenbreiten
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 4
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 4
    ws.Columns("G").ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 10
    ws.Columns("J").ColumnWidth = 4
    ws.Columns("K").ColumnWidth = 2
    
    ' --- 3. Titel-Banner ---
    With ws.Range("A3:K3")
        .Merge
        .value = "   FINANZ-" & ChrW(220) & "BERSICHT"
        .Font.Size = 18
        .Font.Bold = True
        .Font.color = CLR_WHITE
        .Interior.color = CLR_HEADER
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(3).RowHeight = 40
    
    ' Untertitel mit Filter-Info
    Dim abrJahr As Long
    abrJahr = HoleAbrechnungsjahr()
    Dim filterText As String
    If monatFilter = 0 Then
        filterText = "Gesamtjahr"
    Else
        filterText = Format$(DateSerial(2000, monatFilter, 1), "MMMM")
    End If
    
    With ws.Range("B4:F4")
        .Merge
        .value = "Abrechnungsjahr " & IIf(abrJahr > 0, CStr(abrJahr), "---") & "  |  Filter:"
        .Font.Size = 10
        .Font.color = RGB(200, 200, 200)
        .Interior.color = CLR_HEADER
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A4").Interior.color = CLR_HEADER
    ws.Range("G4:K4").Interior.color = CLR_HEADER
    ws.Rows(4).RowHeight = 24
    
    ' Akzentlinie
    ws.Range("A5:K5").Interior.color = CLR_ACCENT
    ws.Rows(5).RowHeight = 4
    
    ' Filterwert speichern (versteckt)
    ws.Range("A4").value = monatFilter
    ws.Range("A4").Font.color = CLR_HEADER
    
    ' --- 4. KPIs (Zeile 7-9) ---
    ws.Rows(6).RowHeight = 10
    ws.Range("A6:K10").Interior.color = CLR_LIGHT_BG
    
    With ws.Range("B7:I7")
        .Merge
        .value = ChrW(9473) & ChrW(9473) & "  KENNZAHLEN  " & ChrW(9473) & ChrW(9473)
        .Font.Size = 9
        .Font.Bold = True
        .Font.color = RGB(140, 140, 140)
        .Interior.color = CLR_LIGHT_BG
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(7).RowHeight = 18
    
    ' Summen berechnen
    Dim sumEinn As Double, sumAusg As Double
    Dim k As Variant
    For Each k In dictEinn.keys
        sumEinn = sumEinn + dictEinn(k)
    Next k
    For Each k In dictAusg.keys
        sumAusg = sumAusg + dictAusg(k)
    Next k
    
    ' Kontostand berechnen (Vorjahr + alle Buchungen)
    Dim kontostand As Double
    kontostand = HoleKontostandVorjahr()
    Dim wsBK As Worksheet
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo 0
    If Not wsBK Is Nothing Then
        Dim lr As Long
        lr = wsBK.Cells(wsBK.Rows.count, BK_COL_BETRAG).End(xlUp).Row
        If lr >= BK_START_ROW Then
            kontostand = kontostand + Application.WorksheetFunction.Sum( _
                wsBK.Range(wsBK.Cells(BK_START_ROW, BK_COL_BETRAG), wsBK.Cells(lr, BK_COL_BETRAG)))
        End If
    End If
    
    ' VK-Saldo berechnen
    Dim vkSaldo As Double
    Dim wsVK As Worksheet
    On Error Resume Next
    Set wsVK = ThisWorkbook.Worksheets(WS_VEREINSKASSE)
    On Error GoTo 0
    Dim lrVK As Long
    lrVK = 0
    If Not wsVK Is Nothing Then
        lrVK = wsVK.Cells(wsVK.Rows.count, VK_COL_BETRAG).End(xlUp).Row
        If lrVK >= VK_START_ROW Then
            vkSaldo = Application.WorksheetFunction.Sum( _
                wsVK.Range(wsVK.Cells(VK_START_ROW, VK_COL_BETRAG), wsVK.Cells(lrVK, VK_COL_BETRAG)))
        End If
    End If
    
    ' KPI-Karten schreiben
    ws.Rows(8).RowHeight = 42
    ws.Rows(9).RowHeight = 16
    
    Call SchreibeKPI(ws, "C", sumEinn, "Einnahmen", RGB(39, 174, 96))
    Call SchreibeKPI(ws, "E", sumAusg, "Ausgaben", RGB(231, 76, 60))
    Call SchreibeKPI(ws, "G", sumEinn - sumAusg, "Saldo", RGB(41, 128, 185))
    Call SchreibeKPI(ws, "I", kontostand, "Kontostand", RGB(142, 68, 173))
    
    ws.Rows(10).RowHeight = 10
    
    ' VK-Saldo als kleine Zusatzinfo
    With ws.Range("G10:I10")
        .Merge
        .value = "Vereinskasse: " & Format$(vkSaldo, "#,##0.00") & " " & ChrW(8364)
        .Font.Size = 8
        .Font.color = RGB(120, 120, 120)
        .Interior.color = CLR_LIGHT_BG
        .HorizontalAlignment = xlRight
    End With
    
    ' --- 5. Einnahmen-Tabelle ---
    Dim curRow As Long
    curRow = 11
    ws.Rows(curRow).RowHeight = 10
    curRow = curRow + 1
    
    With ws.Range("B" & curRow & ":I" & curRow)
        .Merge
        .value = ChrW(9650) & "  EINNAHMEN"
        .Font.Size = 11
        .Font.Bold = True
        .Font.color = CLR_WHITE
        .Interior.color = CLR_EINN
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    ws.Rows(curRow).RowHeight = 24
    curRow = curRow + 1
    
    Call SchreibeTabellenkopf(ws, curRow, CLR_EINN_LIGHT)
    curRow = curRow + 1
    
    Dim einnStartRow As Long
    einnStartRow = curRow
    
    If cntEinn > 0 Then
        Dim ie As Long
        For ie = 0 To cntEinn - 1
            ws.Range("C" & curRow).value = arrEinnKat(ie)
            ws.Range("D" & curRow).value = arrEinnVal(ie)
            ws.Range("D" & curRow).NumberFormat = "#,##0.00"
            ws.Range("D" & curRow).HorizontalAlignment = xlRight
            If sumEinn > 0 Then
                ws.Range("E" & curRow).value = arrEinnVal(ie) / sumEinn
            End If
            ws.Range("E" & curRow).NumberFormat = "0.0%"
            ws.Range("E" & curRow).HorizontalAlignment = xlRight
            ws.Range("C" & curRow & ":E" & curRow).Font.Size = 9
            If ie Mod 2 = 1 Then
                ws.Range("B" & curRow & ":I" & curRow).Interior.color = CLR_LIGHT_BG
            End If
            ws.Rows(curRow).RowHeight = 20
            curRow = curRow + 1
        Next ie
    Else
        ws.Range("C" & curRow).value = "(keine Einnahmen)"
        ws.Range("C" & curRow).Font.Size = 9
        ws.Range("C" & curRow).Font.Italic = True
        ws.Rows(curRow).RowHeight = 20
        curRow = curRow + 1
    End If
    
    Dim einnEndRow As Long
    einnEndRow = curRow - 1
    
    ' Summenzeile Einnahmen
    ws.Range("C" & curRow).value = "SUMME"
    ws.Range("C" & curRow).Font.Bold = True
    ws.Range("C" & curRow).Font.Size = 10
    ws.Range("D" & curRow).value = sumEinn
    ws.Range("D" & curRow).NumberFormat = "#,##0.00"
    ws.Range("D" & curRow).Font.Bold = True
    ws.Range("D" & curRow).Font.Size = 10
    ws.Range("D" & curRow).HorizontalAlignment = xlRight
    ws.Range("B" & curRow & ":I" & curRow).Interior.color = CLR_SUM_BG
    ws.Range("B" & curRow & ":I" & curRow).Borders(xlEdgeTop).Weight = xlThin
    ws.Range("B" & curRow & ":I" & curRow).Borders(xlEdgeBottom).Weight = xlMedium
    ws.Range("B" & curRow & ":I" & curRow).Borders(xlEdgeBottom).color = CLR_EINN
    ws.Rows(curRow).RowHeight = 22
    curRow = curRow + 1
    
    ' --- 6. Ausgaben-Tabelle ---
    ws.Rows(curRow).RowHeight = 10
    curRow = curRow + 1
    
    With ws.Range("B" & curRow & ":I" & curRow)
        .Merge
        .value = ChrW(9660) & "  AUSGABEN"
        .Font.Size = 11
        .Font.Bold = True
        .Font.color = CLR_WHITE
        .Interior.color = CLR_AUSG
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    ws.Rows(curRow).RowHeight = 24
    curRow = curRow + 1
    
    Call SchreibeTabellenkopf(ws, curRow, CLR_AUSG_LIGHT)
    curRow = curRow + 1
    
    Dim ausgStartRow As Long
    ausgStartRow = curRow
    
    If cntAusg > 0 Then
        Dim ia As Long
        For ia = 0 To cntAusg - 1
            ws.Range("C" & curRow).value = arrAusgKat(ia)
            ws.Range("D" & curRow).value = arrAusgVal(ia)
            ws.Range("D" & curRow).NumberFormat = "#,##0.00"
            ws.Range("D" & curRow).HorizontalAlignment = xlRight
            If sumAusg > 0 Then
                ws.Range("E" & curRow).value = arrAusgVal(ia) / sumAusg
            End If
            ws.Range("E" & curRow).NumberFormat = "0.0%"
            ws.Range("E" & curRow).HorizontalAlignment = xlRight
            ws.Range("C" & curRow & ":E" & curRow).Font.Size = 9
            If ia Mod 2 = 1 Then
                ws.Range("B" & curRow & ":I" & curRow).Interior.color = CLR_LIGHT_BG
            End If
            ws.Rows(curRow).RowHeight = 20
            curRow = curRow + 1
        Next ia
    Else
        ws.Range("C" & curRow).value = "(keine Ausgaben)"
        ws.Range("C" & curRow).Font.Size = 9
        ws.Range("C" & curRow).Font.Italic = True
        ws.Rows(curRow).RowHeight = 20
        curRow = curRow + 1
    End If
    
    Dim ausgEndRow As Long
    ausgEndRow = curRow - 1
    
    ' Summenzeile Ausgaben
    ws.Range("C" & curRow).value = "SUMME"
    ws.Range("C" & curRow).Font.Bold = True
    ws.Range("C" & curRow).Font.Size = 10
    ws.Range("D" & curRow).value = sumAusg
    ws.Range("D" & curRow).NumberFormat = "#,##0.00"
    ws.Range("D" & curRow).Font.Bold = True
    ws.Range("D" & curRow).Font.Size = 10
    ws.Range("D" & curRow).HorizontalAlignment = xlRight
    ws.Range("B" & curRow & ":I" & curRow).Interior.color = CLR_SUM_BG
    ws.Range("B" & curRow & ":I" & curRow).Borders(xlEdgeTop).Weight = xlThin
    ws.Range("B" & curRow & ":I" & curRow).Borders(xlEdgeBottom).Weight = xlMedium
    ws.Range("B" & curRow & ":I" & curRow).Borders(xlEdgeBottom).color = CLR_AUSG
    ws.Rows(curRow).RowHeight = 22
    curRow = curRow + 1
    
    ' --- 7. Vereinskasse-Bereich ---
    ws.Rows(curRow).RowHeight = 10
    curRow = curRow + 1
    
    With ws.Range("B" & curRow & ":I" & curRow)
        .Merge
        .value = ChrW(9830) & "  VEREINSKASSE"
        .Font.Size = 11
        .Font.Bold = True
        .Font.color = CLR_WHITE
        .Interior.color = CLR_VK
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    ws.Rows(curRow).RowHeight = 24
    curRow = curRow + 1
    
    ' VK-Einnahmen, -Ausgaben, -Saldo
    Dim vkEinn As Double, vkAusg As Double
    If Not wsVK Is Nothing And lrVK >= VK_START_ROW Then
        Dim rv As Long
        For rv = VK_START_ROW To lrVK
            Dim vkBetrag As Double
            vkBetrag = 0
            If IsNumeric(wsVK.Cells(rv, VK_COL_BETRAG).value) Then
                vkBetrag = CDbl(wsVK.Cells(rv, VK_COL_BETRAG).value)
            End If
            
            If monatFilter > 0 And IsDate(wsVK.Cells(rv, VK_COL_DATUM).value) Then
                If Month(CDate(wsVK.Cells(rv, VK_COL_DATUM).value)) <> monatFilter Then GoTo NextVK
            End If
            
            If vkBetrag > 0 Then
                vkEinn = vkEinn + vkBetrag
            ElseIf vkBetrag < 0 Then
                vkAusg = vkAusg + Abs(vkBetrag)
            End If
NextVK:
        Next rv
    End If
    
    ws.Range("C" & curRow).value = "Einnahmen"
    ws.Range("D" & curRow).value = "Ausgaben"
    ws.Range("E" & curRow).value = "Saldo"
    ws.Range("C" & curRow & ":E" & curRow).Font.Bold = True
    ws.Range("C" & curRow & ":E" & curRow).Font.Size = 9
    ws.Range("D" & curRow & ":E" & curRow).HorizontalAlignment = xlRight
    ws.Range("B" & curRow & ":I" & curRow).Interior.color = RGB(230, 235, 230)
    ws.Rows(curRow).RowHeight = 20
    curRow = curRow + 1
    
    ws.Range("C" & curRow).value = vkEinn
    ws.Range("D" & curRow).value = vkAusg
    ws.Range("E" & curRow).value = vkEinn - vkAusg
    ws.Range("C" & curRow & ":E" & curRow).NumberFormat = "#,##0.00"
    ws.Range("C" & curRow & ":E" & curRow).Font.Size = 10
    ws.Range("C" & curRow & ":E" & curRow).Font.Bold = True
    ws.Range("D" & curRow & ":E" & curRow).HorizontalAlignment = xlRight
    ws.Range("B" & curRow & ":I" & curRow).Interior.color = CLR_LIGHT_BG
    ws.Rows(curRow).RowHeight = 22
    curRow = curRow + 1
    
    ' --- 8. Filter-DropDown erstellen ---
    Call ErstelleFilterDropDown(ws, monatFilter)
    
    ' --- 9. Diagramme ---
    ws.Rows(curRow).RowHeight = 10
    curRow = curRow + 1
    
    Call ErstelleDiagramme(ws, curRow, _
        einnStartRow, einnEndRow, cntEinn, _
        ausgStartRow, ausgEndRow, cntAusg, _
        filterText)
End Sub


' ===============================================================
' DATEN SAMMELN: Bankkonto iterieren, nach Kategorie gruppieren
' ===============================================================
Private Sub SammleDaten(ByRef dictEinn As Object, _
                        ByRef dictAusg As Object, _
                        ByVal monatFilter As Long)
    
    Dim wsBK As Worksheet
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo 0
    If wsBK Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    Dim r As Long
    For r = BK_START_ROW To lastRow
        ' Monatsfilter pruefen
        If monatFilter > 0 Then
            If IsDate(wsBK.Cells(r, BK_COL_DATUM).value) Then
                If Month(CDate(wsBK.Cells(r, BK_COL_DATUM).value)) <> monatFilter Then GoTo nextRow
            Else
                GoTo nextRow
            End If
        End If
        
        Dim kategorie As String
        kategorie = Trim(CStr(wsBK.Cells(r, BK_COL_KATEGORIE).value))
        If kategorie = "" Then GoTo nextRow
        
        Dim betrag As Double
        betrag = 0
        If IsNumeric(wsBK.Cells(r, BK_COL_BETRAG).value) Then
            betrag = CDbl(wsBK.Cells(r, BK_COL_BETRAG).value)
        End If
        If betrag = 0 Then GoTo nextRow
        
        ' Sammelzahlung? -> Spalte L (Bemerkung) parsen
        If InStr(1, LCase(kategorie), LCase(KAT_SAMMELZAHLUNG)) > 0 Then
            Call VerteileSammelzahlung(wsBK, r, betrag, dictEinn, dictAusg)
        Else
            ' Normaler Eintrag: Betrag der Kategorie zuordnen
            If betrag > 0 Then
                If Not dictEinn.Exists(kategorie) Then dictEinn.Add kategorie, 0
                dictEinn(kategorie) = dictEinn(kategorie) + betrag
            Else
                If Not dictAusg.Exists(kategorie) Then dictAusg.Add kategorie, 0
                dictAusg(kategorie) = dictAusg(kategorie) + Abs(betrag)
            End If
        End If
nextRow:
    Next r
End Sub


' ===============================================================
' SAMMELZAHLUNG: Spalte L (Bemerkung) parsen
' Format: "SAMMEL:" & vbLf & "Kategorie: Betrag ?" & vbLf & ...
' Wird der Gesamtbetrag (positiv/negativ) beruecksichtigt um
' Einnahmen/Ausgaben korrekt zuzuordnen.
' ===============================================================
Private Sub VerteileSammelzahlung(ByVal wsBK As Worksheet, _
                                  ByVal zeile As Long, _
                                  ByVal gesamtBetrag As Double, _
                                  ByRef dictEinn As Object, _
                                  ByRef dictAusg As Object)
    
    Dim bemerkung As String
    bemerkung = Trim(CStr(wsBK.Cells(zeile, BK_COL_BEMERKUNG).value))
    
    ' Pruefen ob Spalte L das SAMMEL:-Format enthaelt
    If Left(UCase(bemerkung), 7) <> "SAMMEL:" Then
        ' Kein SAMMEL-Format -> Gesamtbetrag als "Sammelzahlung" buchen
        If gesamtBetrag > 0 Then
            If Not dictEinn.Exists(KAT_SAMMELZAHLUNG) Then dictEinn.Add KAT_SAMMELZAHLUNG, 0
            dictEinn(KAT_SAMMELZAHLUNG) = dictEinn(KAT_SAMMELZAHLUNG) + gesamtBetrag
        ElseIf gesamtBetrag < 0 Then
            If Not dictAusg.Exists(KAT_SAMMELZAHLUNG) Then dictAusg.Add KAT_SAMMELZAHLUNG, 0
            dictAusg(KAT_SAMMELZAHLUNG) = dictAusg(KAT_SAMMELZAHLUNG) + Abs(gesamtBetrag)
        End If
        Exit Sub
    End If
    
    ' SAMMEL:-Format parsen: Zeilen nach "SAMMEL:" lesen
    Dim zeilen() As String
    If InStr(bemerkung, vbLf) > 0 Then
        zeilen = Split(bemerkung, vbLf)
    ElseIf InStr(bemerkung, vbCrLf) > 0 Then
        zeilen = Split(bemerkung, vbCrLf)
    Else
        zeilen = Split(bemerkung, Chr(10))
    End If
    
    Dim hatTeilbetraege As Boolean
    hatTeilbetraege = False
    
    Dim z As Long
    For z = LBound(zeilen) To UBound(zeilen)
        Dim eineZeile As String
        eineZeile = Trim(zeilen(z))
        
        ' Erste Zeile "SAMMEL:" ueberspringen
        If UCase(eineZeile) = "SAMMEL:" Or eineZeile = "" Then GoTo NextZeile
        
        ' Format: "Kategorie: Betrag ?"
        Dim doppelPunkt As Long
        doppelPunkt = InStr(eineZeile, ":")
        If doppelPunkt > 1 Then
            Dim teilKat As String
            teilKat = Trim(Left(eineZeile, doppelPunkt - 1))
            
            ' Betrag extrahieren: alles nach dem Doppelpunkt
            Dim betragStr As String
            betragStr = Trim(Mid(eineZeile, doppelPunkt + 1))
            
            ' Euro-Zeichen und Leerzeichen entfernen
            betragStr = Replace(betragStr, ChrW(8364), "")
            betragStr = Replace(betragStr, "EUR", "")
            betragStr = Trim(betragStr)
            
            ' Komma durch Punkt ersetzen fuer CDbl
            betragStr = Replace(betragStr, ".", "")
            betragStr = Replace(betragStr, ",", ".")
            
            Dim teilBetrag As Double
            teilBetrag = 0
            On Error Resume Next
            teilBetrag = CDbl(betragStr)
            On Error GoTo 0
            
            If teilBetrag > 0 And teilKat <> "" Then
                hatTeilbetraege = True
                
                ' Richtung (Einnahme/Ausgabe) vom Gesamtbetrag ableiten
                If gesamtBetrag > 0 Then
                    If Not dictEinn.Exists(teilKat) Then dictEinn.Add teilKat, 0
                    dictEinn(teilKat) = dictEinn(teilKat) + teilBetrag
                Else
                    If Not dictAusg.Exists(teilKat) Then dictAusg.Add teilKat, 0
                    dictAusg(teilKat) = dictAusg(teilKat) + teilBetrag
                End If
            End If
        End If
NextZeile:
    Next z
    
    ' Falls Parsing fehlgeschlagen: Gesamtbetrag als "Sammelzahlung"
    If Not hatTeilbetraege Then
        If gesamtBetrag > 0 Then
            If Not dictEinn.Exists(KAT_SAMMELZAHLUNG) Then dictEinn.Add KAT_SAMMELZAHLUNG, 0
            dictEinn(KAT_SAMMELZAHLUNG) = dictEinn(KAT_SAMMELZAHLUNG) + gesamtBetrag
        ElseIf gesamtBetrag < 0 Then
            If Not dictAusg.Exists(KAT_SAMMELZAHLUNG) Then dictAusg.Add KAT_SAMMELZAHLUNG, 0
            dictAusg(KAT_SAMMELZAHLUNG) = dictAusg(KAT_SAMMELZAHLUNG) + Abs(gesamtBetrag)
        End If
    End If
End Sub


' ===============================================================
' DICTIONARY -> SORTIERTE ARRAYS (absteigend nach Betrag)
' ===============================================================
Private Sub DictToSortedArrays(ByVal dict As Object, _
                                ByRef arrKat() As String, _
                                ByRef arrVal() As Double, _
                                ByRef cnt As Long)
    cnt = dict.count
    If cnt = 0 Then Exit Sub
    
    ReDim arrKat(0 To cnt - 1)
    ReDim arrVal(0 To cnt - 1)
    
    Dim idx As Long
    idx = 0
    Dim k As Variant
    For Each k In dict.keys
        arrKat(idx) = CStr(k)
        arrVal(idx) = dict(k)
        idx = idx + 1
    Next k
    
    ' Bubble-Sort absteigend nach Betrag
    Dim i As Long, j As Long
    Dim tmpK As String, tmpV As Double
    For i = 0 To cnt - 2
        For j = i + 1 To cnt - 1
            If arrVal(j) > arrVal(i) Then
                tmpK = arrKat(i): arrKat(i) = arrKat(j): arrKat(j) = tmpK
                tmpV = arrVal(i): arrVal(i) = arrVal(j): arrVal(j) = tmpV
            End If
        Next j
    Next i
End Sub


' ===============================================================
' KPI-KARTE SCHREIBEN
' ===============================================================
Private Sub SchreibeKPI(ByVal ws As Worksheet, _
                         ByVal col As String, _
                         ByVal wert As Double, _
                         ByVal label As String, _
                         ByVal akzentFarbe As Long)
    With ws.Range(col & "8")
        .value = wert
        .NumberFormat = "#,##0.00 " & ChrW(8364)
        .Font.Size = 14
        .Font.Bold = True
        .Font.color = CLR_DARK_TEXT
        .Interior.color = CLR_WHITE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).color = akzentFarbe
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    With ws.Range(col & "9")
        .value = label
        .Font.Size = 8
        .Font.Bold = True
        .Font.color = RGB(120, 120, 120)
        .Interior.color = CLR_LIGHT_BG
        .HorizontalAlignment = xlCenter
    End With
End Sub


' ===============================================================
' TABELLENKOPF SCHREIBEN
' ===============================================================
Private Sub SchreibeTabellenkopf(ByVal ws As Worksheet, _
                                  ByVal zeile As Long, _
                                  ByVal bgColor As Long)
    With ws.Range("C" & zeile)
        .value = "Kategorie"
        .Font.Bold = True
        .Font.Size = 9
        .Interior.color = bgColor
    End With
    With ws.Range("D" & zeile)
        .value = "Betrag"
        .Font.Bold = True
        .Font.Size = 9
        .Interior.color = bgColor
        .HorizontalAlignment = xlRight
    End With
    With ws.Range("E" & zeile)
        .value = "Anteil"
        .Font.Bold = True
        .Font.Size = 9
        .Interior.color = bgColor
        .HorizontalAlignment = xlRight
    End With
    
    ws.Range("B" & zeile).Interior.color = bgColor
    ws.Range("F" & zeile & ":I" & zeile).Interior.color = bgColor
    ws.Rows(zeile).RowHeight = 20
End Sub


' ===============================================================
' FILTER-DROPDOWN
' ===============================================================
Private Sub ErstelleFilterDropDown(ByVal ws As Worksheet, ByVal aktuellerMonat As Long)
    Dim ddLeft As Double
    Dim ddTop As Double
    ddLeft = ws.Range("G4").Left
    ddTop = ws.Range("G4").Top + 2
    
    On Error GoTo DDErr
    
    Dim dd As DropDown
    Set dd = ws.DropDowns.Add(ddLeft, ddTop, 110, 18)
    
    With dd
        .Name = FILTER_DD_NAME
        .AddItem "Gesamtjahr"
        .AddItem "Januar"
        .AddItem "Februar"
        .AddItem "M" & ChrW(228) & "rz"
        .AddItem "April"
        .AddItem "Mai"
        .AddItem "Juni"
        .AddItem "Juli"
        .AddItem "August"
        .AddItem "September"
        .AddItem "Oktober"
        .AddItem "November"
        .AddItem "Dezember"
        .value = aktuellerMonat + 1
        .OnAction = "'mod_FinanzUebersicht.MonatFilterChanged'"
    End With
    
    Exit Sub

DDErr:
    Debug.Print "[FinanzUebersicht] DropDown-Fehler: " & Err.Description
    Err.Clear
End Sub


' ===============================================================
' FILTER-HANDLER: Komplette Neuberechnung bei Filterwechsel
' ===============================================================
Public Sub MonatFilterChanged()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_FINANZ_UEBERSICHT())
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    Dim dd As DropDown
    On Error Resume Next
    Set dd = ws.DropDowns(FILTER_DD_NAME)
    On Error GoTo 0
    If dd Is Nothing Then Exit Sub
    
    Dim monatWert As Long
    monatWert = dd.value - 1
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    Call EntferneAlleObjekte(ws)
    Call BaueFinanzUebersicht(ws, monatWert)
    
    ws.Cells.Locked = True
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


' ===============================================================
' DIAGRAMME ERSTELLEN
' ===============================================================
Private Sub ErstelleDiagramme(ByVal ws As Worksheet, _
                               ByVal startZeile As Long, _
                               ByVal einnVon As Long, _
                               ByVal einnBis As Long, _
                               ByVal cntEinn As Long, _
                               ByVal ausgVon As Long, _
                               ByVal ausgBis As Long, _
                               ByVal cntAusg As Long, _
                               ByVal zeitraum As String)
    
    Dim chartLeft As Double
    Dim chartTop As Double
    Dim chartWidth As Double
    Dim chartHeight As Long
    
    chartLeft = ws.Range("C" & startZeile).Left
    chartTop = ws.Range("C" & startZeile).Top
    chartWidth = ws.Range("C" & startZeile & ":I" & startZeile).Width
    chartHeight = 220
    
    On Error GoTo ChartErr
    
    ' --- Diagramm 1: Einnahmen ---
    If cntEinn > 0 Then
        Dim cht1 As ChartObject
        Set cht1 = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)
        cht1.Name = "cht_Einnahmen"
        
        With cht1.Chart
            .ChartType = xlColumnClustered
            
            Dim sr1 As Series
            Set sr1 = .SeriesCollection.NewSeries
            sr1.Name = "Einnahmen"
            sr1.values = ws.Range("D" & einnVon & ":D" & einnBis)
            sr1.XValues = ws.Range("C" & einnVon & ":C" & einnBis)
            sr1.Format.Fill.ForeColor.RGB = RGB(39, 174, 96)
            
            .HasTitle = True
            .ChartTitle.text = "Einnahmen nach Kategorie - " & zeitraum
            .ChartTitle.Font.Size = 11
            .ChartTitle.Font.Bold = True
            .HasLegend = False
            
            .Axes(xlCategory).TickLabels.Font.Size = 8
            .Axes(xlValue).TickLabels.Font.Size = 8
            .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        End With
        
        chartTop = chartTop + chartHeight + 15
    End If
    
    ' --- Diagramm 2: Ausgaben ---
    If cntAusg > 0 Then
        Dim cht2 As ChartObject
        Set cht2 = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)
        cht2.Name = "cht_Ausgaben"
        
        With cht2.Chart
            .ChartType = xlColumnClustered
            
            Dim sr2 As Series
            Set sr2 = .SeriesCollection.NewSeries
            sr2.Name = "Ausgaben"
            sr2.values = ws.Range("D" & ausgVon & ":D" & ausgBis)
            sr2.XValues = ws.Range("C" & ausgVon & ":C" & ausgBis)
            sr2.Format.Fill.ForeColor.RGB = RGB(231, 76, 60)
            
            .HasTitle = True
            .ChartTitle.text = "Ausgaben nach Kategorie - " & zeitraum
            .ChartTitle.Font.Size = 11
            .ChartTitle.Font.Bold = True
            .HasLegend = False
            
            .Axes(xlCategory).TickLabels.Font.Size = 8
            .Axes(xlValue).TickLabels.Font.Size = 8
            .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        End With
    End If
    
    Exit Sub

ChartErr:
    Debug.Print "[FinanzUebersicht] Diagramm-Fehler: " & Err.Description
    Err.Clear
End Sub






