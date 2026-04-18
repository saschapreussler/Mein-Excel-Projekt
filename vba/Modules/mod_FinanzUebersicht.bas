Attribute VB_Name = "mod_FinanzUebersicht"
Option Explicit

' ===============================================================
' MODUL: mod_FinanzUebersicht
' VERSION: 1.0 - 20.04.2026
' ZWECK: Erstellt und pflegt das Blatt "Finanz-Uebersicht"
'        - KPI-Bereich (Einnahmen, Ausgaben, Saldo, Kontostand)
'        - Einnahmen/Ausgaben-Tabellen nach Kategorie
'        - Vereinskasse-Zusammenfassung
'        - Monatsfilter via DropDown
'        - Balkendiagramme Einnahmen / Ausgaben
' ===============================================================

' --- Farben ---
Private Const CLR_HEADER As Long = 2763306      ' RGB(26, 35, 42)
Private Const CLR_ACCENT As Long = 14521384     ' RGB(40, 167, 221)
Private Const CLR_WHITE As Long = 16777215
Private Const CLR_LIGHT_BG As Long = 15921906   ' RGB(242, 242, 242)
Private Const CLR_EINN As Long = 2573097        ' RGB(41, 69, 39) - Einnahmen dunkelgruen
Private Const CLR_EINN_LIGHT As Long = 14348258 ' RGB(226, 240, 217) - Einnahmen hell
Private Const CLR_AUSG As Long = 4743219        ' RGB(163, 80, 72) - Ausgaben rot
Private Const CLR_AUSG_LIGHT As Long = 13688301 ' RGB(237, 220, 209) - Ausgaben hell
Private Const CLR_DARK_TEXT As Long = 2500134    ' RGB(38, 50, 56)
Private Const CLR_SUM_BG As Long = 14408667     ' RGB(219, 223, 219) - Summenzeile
Private Const CLR_VK As Long = 7168108          ' RGB(108, 117, 109) - Vereinskasse

' --- Layout-Zeilen ---
Private Const R_TITLE As Long = 1
Private Const R_SUBTITLE As Long = 2
Private Const R_ACCENT As Long = 3
Private Const R_KPI_HEADER As Long = 5
Private Const R_KPI_VALUE As Long = 6
Private Const R_KPI_LABEL As Long = 7
Private Const R_EINN_HEADER As Long = 9
Private Const R_EINN_COLHEAD As Long = 10
Private Const R_EINN_START As Long = 11
Private Const R_EINN_END As Long = 17
Private Const R_EINN_SUM As Long = 18
Private Const R_AUSG_HEADER As Long = 20
Private Const R_AUSG_COLHEAD As Long = 21
Private Const R_AUSG_START As Long = 22
Private Const R_AUSG_END As Long = 28
Private Const R_AUSG_SUM As Long = 29
Private Const R_VK_HEADER As Long = 31
Private Const R_VK_COLHEAD As Long = 32
Private Const R_VK_DATA As Long = 33
Private Const R_CHART_START As Long = 35

Private Const FILTER_DD_NAME As String = "dd_MonatFilter_FU"


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
    Call VorbereiteBlatt(ws)
    Call SchreibeTitel(ws)
    Call SchreibeKPIs(ws)
    Call SchreibeEinnahmenTabelle(ws)
    Call SchreibeAusgabenTabelle(ws)
    Call SchreibeVereinskasse(ws)
    Call ErstelleFilterDropDown(ws)
    Call ErstelleDiagramme(ws)
    
    ws.Cells.Locked = True
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
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
' BLATT VORBEREITEN
' ===============================================================
Private Sub VorbereiteBlatt(ByVal ws As Worksheet)
    ws.Cells.ClearContents
    ws.Cells.ClearFormats
    ws.Cells.Interior.color = CLR_WHITE
    
    ' Gitternetzlinien aus
    Dim wnd As Window
    For Each wnd In Application.Windows
        If wnd.Caption = ThisWorkbook.Name Then
            wnd.DisplayGridlines = False
        End If
    Next wnd
    
    ' Spaltenbreiten
    ws.Columns("A").ColumnWidth = 2      ' Rand
    ws.Columns("B").ColumnWidth = 4      ' Padding
    ws.Columns("C").ColumnWidth = 24     ' Kategorie / Label
    ws.Columns("D").ColumnWidth = 14     ' Betrag
    ws.Columns("E").ColumnWidth = 10     ' Anteil
    ws.Columns("F").ColumnWidth = 4      ' Luecke
    ws.Columns("G").ColumnWidth = 24     ' Label 2
    ws.Columns("H").ColumnWidth = 14     ' Wert 2
    ws.Columns("I").ColumnWidth = 10     ' Extra
    ws.Columns("J").ColumnWidth = 4      ' Padding
    ws.Columns("K").ColumnWidth = 2      ' Rand
    
    ' Zeilenhoehen
    ws.Rows(R_TITLE).RowHeight = 40
    ws.Rows(R_SUBTITLE).RowHeight = 24
    ws.Rows(R_ACCENT).RowHeight = 4
    ws.Rows(4).RowHeight = 10
    ws.Rows(R_KPI_HEADER).RowHeight = 18
    ws.Rows(R_KPI_VALUE).RowHeight = 42
    ws.Rows(R_KPI_LABEL).RowHeight = 16
    ws.Rows(8).RowHeight = 10
    ws.Rows(R_EINN_HEADER).RowHeight = 24
    ws.Rows(R_EINN_COLHEAD).RowHeight = 20
    
    Dim r As Long
    For r = R_EINN_START To R_EINN_END
        ws.Rows(r).RowHeight = 20
    Next r
    ws.Rows(R_EINN_SUM).RowHeight = 22
    
    ws.Rows(19).RowHeight = 10
    ws.Rows(R_AUSG_HEADER).RowHeight = 24
    ws.Rows(R_AUSG_COLHEAD).RowHeight = 20
    
    For r = R_AUSG_START To R_AUSG_END
        ws.Rows(r).RowHeight = 20
    Next r
    ws.Rows(R_AUSG_SUM).RowHeight = 22
    
    ws.Rows(30).RowHeight = 10
    ws.Rows(R_VK_HEADER).RowHeight = 24
    ws.Rows(R_VK_COLHEAD).RowHeight = 20
    ws.Rows(R_VK_DATA).RowHeight = 22
    ws.Rows(34).RowHeight = 10
    
    ' Filter-Monat Standardwert (0 = Gesamtjahr)
    ws.Range("A2").value = 0
    ws.Range("A2").Font.color = CLR_WHITE
End Sub


' ===============================================================
' TITEL-BANNER
' ===============================================================
Private Sub SchreibeTitel(ByVal ws As Worksheet)
    With ws.Range("A1:K1")
        .Merge
        .value = "   FINANZ-" & ChrW(220) & "BERSICHT"
        .Font.Size = 18
        .Font.Bold = True
        .Font.color = CLR_WHITE
        .Interior.color = CLR_HEADER
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    Dim abrJahr As Long
    abrJahr = HoleAbrechnungsjahr()
    
    With ws.Range("B2:F2")
        .Merge
        .value = "Abrechnungsjahr " & IIf(abrJahr > 0, CStr(abrJahr), "---") & "  |  Filter:"
        .Font.Size = 10
        .Font.color = RGB(200, 200, 200)
        .Interior.color = CLR_HEADER
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("A2").Interior.color = CLR_HEADER
    ws.Range("G2:K2").Interior.color = CLR_HEADER
    
    With ws.Range("A3:K3")
        .Interior.color = CLR_ACCENT
    End With
End Sub


' ===============================================================
' KPI-BEREICH: 4 Kennzahlen
' ===============================================================
Private Sub SchreibeKPIs(ByVal ws As Worksheet)
    ws.Range("A4:K8").Interior.color = CLR_LIGHT_BG
    
    With ws.Range("B5:I5")
        .Merge
        .value = ChrW(9473) & ChrW(9473) & "  KENNZAHLEN  " & ChrW(9473) & ChrW(9473)
        .Font.Size = 9
        .Font.Bold = True
        .Font.color = RGB(140, 140, 140)
        .Interior.color = CLR_LIGHT_BG
        .HorizontalAlignment = xlCenter
    End With
    
    ' KPI 1: Gesamteinnahmen (Summe Einnahmen-Tabelle)
    Call SchreibeKPIZelle(ws, "C", "=C" & R_EINN_SUM, "Einnahmen", RGB(39, 174, 96))
    
    ' KPI 2: Gesamtausgaben (ABS weil Ausgaben negativ)
    Call SchreibeKPIZelle(ws, "E", "=ABS(C" & R_AUSG_SUM & ")", "Ausgaben", RGB(231, 76, 60))
    
    ' KPI 3: Saldo (Einnahmen + Ausgaben, da Ausgaben negativ)
    Call SchreibeKPIZelle(ws, "G", "=C" & R_EINN_SUM & "+C" & R_AUSG_SUM, "Saldo", RGB(41, 128, 185))
    
    ' KPI 4: Kontostand aktuell (Vorjahr + alle Buchungen)
    Dim fKonto As String
    fKonto = "=Einstellungen!C" & ES_CFG_KONTOSTAND_ROW & _
             "+SUM(Bankkonto!B" & BK_START_ROW & ":B5000)"
    Call SchreibeKPIZelle(ws, "I", fKonto, "Kontostand", RGB(142, 68, 173))
End Sub


Private Sub SchreibeKPIZelle(ByVal ws As Worksheet, _
                              ByVal col As String, _
                              ByVal formel As String, _
                              ByVal label As String, _
                              ByVal akzentFarbe As Long)
    With ws.Range(col & R_KPI_VALUE)
        .Formula = formel
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
    
    With ws.Range(col & R_KPI_LABEL)
        .value = label
        .Font.Size = 8
        .Font.Bold = True
        .Font.color = RGB(120, 120, 120)
        .Interior.color = CLR_LIGHT_BG
        .HorizontalAlignment = xlCenter
    End With
End Sub


' ===============================================================
' EINNAHMEN-TABELLE (Bankkonto Spalten M-S)
' ===============================================================
Private Sub SchreibeEinnahmenTabelle(ByVal ws As Worksheet)
    ' Section Header
    With ws.Range("B" & R_EINN_HEADER & ":I" & R_EINN_HEADER)
        .Merge
        .value = ChrW(9650) & "  EINNAHMEN"
        .Font.Size = 11
        .Font.Bold = True
        .Font.color = CLR_WHITE
        .Interior.color = CLR_EINN
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    
    ' Spaltenkoepfe
    Call SchreibeTabellenkopf(ws, R_EINN_COLHEAD, CLR_EINN_LIGHT)
    
    ' 7 Einnahmen-Kategorien (Spalte M=13 bis S=19)
    Dim i As Long
    For i = 0 To 6
        Call SchreibeKategorieZeile(ws, R_EINN_START + i, _
            BK_COL_EINNAHMEN_START + i, R_EINN_SUM, (i Mod 2 = 1), False)
    Next i
    
    ' Summenzeile
    Call SchreibeSummenZeile(ws, R_EINN_SUM, R_EINN_START, R_EINN_END, CLR_EINN)
End Sub


' ===============================================================
' AUSGABEN-TABELLE (Bankkonto Spalten T-Z)
' ===============================================================
Private Sub SchreibeAusgabenTabelle(ByVal ws As Worksheet)
    With ws.Range("B" & R_AUSG_HEADER & ":I" & R_AUSG_HEADER)
        .Merge
        .value = ChrW(9660) & "  AUSGABEN"
        .Font.Size = 11
        .Font.Bold = True
        .Font.color = CLR_WHITE
        .Interior.color = CLR_AUSG
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    
    Call SchreibeTabellenkopf(ws, R_AUSG_COLHEAD, CLR_AUSG_LIGHT)
    
    ' 7 Ausgaben-Kategorien (Spalte T=20 bis Z=26)
    ' Ausgaben-Betraege sind negativ -> ABS fuer Anzeige
    Dim i As Long
    For i = 0 To 6
        Call SchreibeKategorieZeile(ws, R_AUSG_START + i, _
            BK_COL_AUSGABEN_START + i, R_AUSG_SUM, (i Mod 2 = 1), True)
    Next i
    
    Call SchreibeSummenZeile(ws, R_AUSG_SUM, R_AUSG_START, R_AUSG_END, CLR_AUSG)
End Sub


' ===============================================================
' TABELLEN-HILFSFUNKTIONEN
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
End Sub


Private Sub SchreibeKategorieZeile(ByVal ws As Worksheet, _
                                    ByVal zeile As Long, _
                                    ByVal bkSpalte As Long, _
                                    ByVal sumZeile As Long, _
                                    ByVal alteFarbe As Boolean, _
                                    ByVal istAusgabe As Boolean)
    ' Spaltenbuchstabe berechnen
    Dim bkColLetter As String
    bkColLetter = Split(Cells(1, bkSpalte).Address(True, False), "$")(0)
    
    ' Kategoriename aus Bankkonto-Header
    With ws.Range("C" & zeile)
        .Formula = "=Bankkonto!" & bkColLetter & BK_HEADER_ROW
        .Font.Size = 9
        .Font.color = CLR_DARK_TEXT
    End With
    
    ' Betrag: Gefiltert nach Monat (A2 = 0 fuer Gesamtjahr)
    Dim fBetrag As String
    If Not istAusgabe Then
        ' Einnahmen: Werte sind positiv
        fBetrag = "=IF($A$2=0," & _
                  "SUM(Bankkonto!" & bkColLetter & BK_START_ROW & ":" & bkColLetter & "5000)," & _
                  "SUMPRODUCT((MONTH(Bankkonto!$A$" & BK_START_ROW & ":$A$5000)=$A$2)*" & _
                  "(Bankkonto!" & bkColLetter & BK_START_ROW & ":" & bkColLetter & "5000)))"
    Else
        ' Ausgaben: Werte sind negativ -> ABS fuer positive Anzeige
        fBetrag = "=ABS(IF($A$2=0," & _
                  "SUM(Bankkonto!" & bkColLetter & BK_START_ROW & ":" & bkColLetter & "5000)," & _
                  "SUMPRODUCT((MONTH(Bankkonto!$A$" & BK_START_ROW & ":$A$5000)=$A$2)*" & _
                  "(Bankkonto!" & bkColLetter & BK_START_ROW & ":" & bkColLetter & "5000))))"
    End If
    
    With ws.Range("D" & zeile)
        .Formula = fBetrag
        .NumberFormat = "#,##0.00"
        .Font.Size = 9
        .HorizontalAlignment = xlRight
    End With
    
    ' Anteil an Summe
    With ws.Range("E" & zeile)
        .Formula = "=IF(D" & sumZeile & "=0,0,D" & zeile & "/D" & sumZeile & ")"
        .NumberFormat = "0.0%"
        .Font.Size = 9
        .HorizontalAlignment = xlRight
    End With
    
    ' Alternating row color
    If alteFarbe Then
        ws.Range("B" & zeile & ":I" & zeile).Interior.color = CLR_LIGHT_BG
    End If
End Sub


Private Sub SchreibeSummenZeile(ByVal ws As Worksheet, _
                                 ByVal zeile As Long, _
                                 ByVal vonZeile As Long, _
                                 ByVal bisZeile As Long, _
                                 ByVal farbe As Long)
    With ws.Range("C" & zeile)
        .value = "SUMME"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    With ws.Range("D" & zeile)
        .Formula = "=SUM(D" & vonZeile & ":D" & bisZeile & ")"
        .NumberFormat = "#,##0.00"
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlRight
    End With
    
    ws.Range("B" & zeile & ":I" & zeile).Interior.color = CLR_SUM_BG
    ws.Range("B" & zeile & ":I" & zeile).Borders(xlEdgeTop).Weight = xlThin
    ws.Range("B" & zeile & ":I" & zeile).Borders(xlEdgeBottom).Weight = xlMedium
    ws.Range("B" & zeile & ":I" & zeile).Borders(xlEdgeBottom).color = farbe
End Sub


' ===============================================================
' VEREINSKASSE-BEREICH
' ===============================================================
Private Sub SchreibeVereinskasse(ByVal ws As Worksheet)
    With ws.Range("B" & R_VK_HEADER & ":I" & R_VK_HEADER)
        .Merge
        .value = ChrW(9830) & "  VEREINSKASSE"
        .Font.Size = 11
        .Font.Bold = True
        .Font.color = CLR_WHITE
        .Interior.color = CLR_VK
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    
    ' Spaltenkoepfe
    ws.Range("C" & R_VK_COLHEAD).value = "Einnahmen"
    ws.Range("D" & R_VK_COLHEAD).value = "Ausgaben"
    ws.Range("E" & R_VK_COLHEAD).value = "Saldo"
    With ws.Range("C" & R_VK_COLHEAD & ":E" & R_VK_COLHEAD)
        .Font.Bold = True
        .Font.Size = 9
    End With
    ws.Range("D" & R_VK_COLHEAD & ":E" & R_VK_COLHEAD).HorizontalAlignment = xlRight
    ws.Range("B" & R_VK_COLHEAD & ":I" & R_VK_COLHEAD).Interior.color = RGB(230, 235, 230)
    
    Dim vs As Long
    vs = VK_START_ROW
    
    ' Einnahmen (positive Betraege)
    Dim fEinn As String
    fEinn = "=IF($A$2=0," & _
            "SUMPRODUCT((Vereinskasse!B" & vs & ":B5000>0)*Vereinskasse!B" & vs & ":B5000)," & _
            "SUMPRODUCT((MONTH(Vereinskasse!A" & vs & ":A5000)=$A$2)*" & _
            "(Vereinskasse!B" & vs & ":B5000>0)*Vereinskasse!B" & vs & ":B5000))"
    
    With ws.Range("C" & R_VK_DATA)
        .Formula = fEinn
        .NumberFormat = "#,##0.00"
        .Font.Size = 10
        .Font.Bold = True
    End With
    
    ' Ausgaben (negative Betraege -> ABS)
    Dim fAusg As String
    fAusg = "=ABS(IF($A$2=0," & _
            "SUMPRODUCT((Vereinskasse!B" & vs & ":B5000<0)*Vereinskasse!B" & vs & ":B5000)," & _
            "SUMPRODUCT((MONTH(Vereinskasse!A" & vs & ":A5000)=$A$2)*" & _
            "(Vereinskasse!B" & vs & ":B5000<0)*Vereinskasse!B" & vs & ":B5000)))"
    
    With ws.Range("D" & R_VK_DATA)
        .Formula = fAusg
        .NumberFormat = "#,##0.00"
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
    ' Saldo
    With ws.Range("E" & R_VK_DATA)
        .Formula = "=C" & R_VK_DATA & "-D" & R_VK_DATA
        .NumberFormat = "#,##0.00"
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
    ws.Range("B" & R_VK_DATA & ":I" & R_VK_DATA).Interior.color = CLR_LIGHT_BG
End Sub


' ===============================================================
' FILTER-DROPDOWN (Forms DropDown mit OnAction)
' ===============================================================
Private Sub ErstelleFilterDropDown(ByVal ws As Worksheet)
    Dim ddLeft As Double
    Dim ddTop As Double
    ddLeft = ws.Range("G2").Left
    ddTop = ws.Range("G2").Top + 2
    
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
        .value = 1  ' Index 1 = Gesamtjahr
        .OnAction = "'mod_FinanzUebersicht.MonatFilterChanged'"
    End With
    
    Exit Sub

DDErr:
    Debug.Print "[FinanzUebersicht] DropDown-Fehler: " & Err.Description
    Err.Clear
End Sub


' ===============================================================
' FILTER-HANDLER: Wird bei Auswahlaenderung aufgerufen
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
    
    ' Index 1 = Gesamtjahr (Monat 0), Index 2-13 = Jan-Dez (Monat 1-12)
    Dim monatWert As Long
    monatWert = dd.value - 1
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ws.Range("A2").value = monatWert
    
    ' Diagrammtitel aktualisieren
    Call AktualisiereDiagrammTitel(ws, monatWert)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
End Sub


' ===============================================================
' DIAGRAMME ERSTELLEN (Einnahmen + Ausgaben)
' ===============================================================
Private Sub ErstelleDiagramme(ByVal ws As Worksheet)
    Dim chartLeft As Double
    Dim chartTop As Double
    Dim chartWidth As Double
    Dim chartHeight As Long
    
    chartLeft = ws.Range("C" & R_CHART_START).Left
    chartTop = ws.Range("C" & R_CHART_START).Top
    chartWidth = ws.Range("C" & R_CHART_START & ":I" & R_CHART_START).Width
    chartHeight = 220
    
    On Error GoTo ChartErr
    
    ' --- Diagramm 1: Einnahmen ---
    Dim cht1 As ChartObject
    Set cht1 = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)
    cht1.Name = "cht_Einnahmen"
    
    With cht1.Chart
        .ChartType = xlColumnClustered
        
        Dim sr1 As Series
        Set sr1 = .SeriesCollection.NewSeries
        sr1.Name = "Einnahmen"
        sr1.values = ws.Range("D" & R_EINN_START & ":D" & R_EINN_END)
        sr1.XValues = ws.Range("C" & R_EINN_START & ":C" & R_EINN_END)
        sr1.Format.Fill.ForeColor.RGB = RGB(39, 174, 96)
        
        .HasTitle = True
        .ChartTitle.text = "Einnahmen nach Kategorie - Gesamtjahr"
        .ChartTitle.Font.Size = 11
        .ChartTitle.Font.Bold = True
        .HasLegend = False
        
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
    End With
    
    ' --- Diagramm 2: Ausgaben ---
    Dim cht2Top As Double
    cht2Top = chartTop + chartHeight + 15
    
    Dim cht2 As ChartObject
    Set cht2 = ws.ChartObjects.Add(chartLeft, cht2Top, chartWidth, chartHeight)
    cht2.Name = "cht_Ausgaben"
    
    With cht2.Chart
        .ChartType = xlColumnClustered
        
        Dim sr2 As Series
        Set sr2 = .SeriesCollection.NewSeries
        sr2.Name = "Ausgaben"
        sr2.values = ws.Range("D" & R_AUSG_START & ":D" & R_AUSG_END)
        sr2.XValues = ws.Range("C" & R_AUSG_START & ":C" & R_AUSG_END)
        sr2.Format.Fill.ForeColor.RGB = RGB(231, 76, 60)
        
        .HasTitle = True
        .ChartTitle.text = "Ausgaben nach Kategorie - Gesamtjahr"
        .ChartTitle.Font.Size = 11
        .ChartTitle.Font.Bold = True
        .HasLegend = False
        
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
    End With
    
    Exit Sub

ChartErr:
    Debug.Print "[FinanzUebersicht] Diagramm-Fehler: " & Err.Description
    Err.Clear
End Sub


' ===============================================================
' DIAGRAMMTITEL AKTUALISIEREN (bei Filterwechsel)
' ===============================================================
Private Sub AktualisiereDiagrammTitel(ByVal ws As Worksheet, ByVal monat As Long)
    Dim zeitraum As String
    If monat = 0 Then
        zeitraum = "Gesamtjahr"
    Else
        zeitraum = Format$(DateSerial(2000, monat, 1), "MMMM")
    End If
    
    On Error Resume Next
    Dim cht1 As ChartObject
    Set cht1 = ws.ChartObjects("cht_Einnahmen")
    If Not cht1 Is Nothing Then
        cht1.Chart.ChartTitle.text = "Einnahmen nach Kategorie - " & zeitraum
    End If
    
    Dim cht2 As ChartObject
    Set cht2 = ws.ChartObjects("cht_Ausgaben")
    If Not cht2 Is Nothing Then
        cht2.Chart.ChartTitle.text = "Ausgaben nach Kategorie - " & zeitraum
    End If
    On Error GoTo 0
End Sub
