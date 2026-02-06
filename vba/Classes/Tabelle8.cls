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
' MODUL: mod_Formatierung
' ZWECK: Formatierung und DropDown-Listen-Verwaltung
' VERSION: 2.1 - 06.02.2026
' ***************************************************************

Private Const ZEBRA_COLOR As Long = &HDEE5E3

Public Sub ZentriereAlleZellenVertikal()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect PASSWORD:=PASSWORD
        ws.Cells.VerticalAlignment = xlCenter
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Next ws
    On Error GoTo 0
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub FormatiereNeuesBlatt(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    ws.Cells.VerticalAlignment = xlCenter
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Public Sub Formatiere_Alle_Tabellen_Neu()
    Dim wsD As Worksheet, wsBK As Worksheet, wsM As Worksheet, ws As Worksheet
    Dim lastRowBK As Long, euroFormat As String
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    euroFormat = "#,##0.00 " & ChrW(8364)
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect PASSWORD:=PASSWORD
        ws.Cells.VerticalAlignment = xlCenter
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        On Error GoTo ErrorHandler
    Next ws
    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo ErrorHandler
    If Not wsD Is Nothing Then
        On Error Resume Next
        wsD.Unprotect PASSWORD:=PASSWORD
        On Error GoTo ErrorHandler
        Call FormatiereKategorieTabelle(wsD)
        Call FormatiereEntityKeyTabelleKomplett(wsD)
        Call AktualisiereKategorieDropdownListen(wsD)
        wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo ErrorHandler
    If Not wsBK Is Nothing Then
        On Error Resume Next
        wsBK.Unprotect PASSWORD:=PASSWORD
        On Error GoTo ErrorHandler
        lastRowBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
        If lastRowBK < BK_START_ROW Then lastRowBK = BK_START_ROW
        With wsBK.Range(wsBK.Cells(BK_START_ROW, BK_COL_BEMERKUNG), wsBK.Cells(lastRowBK, BK_COL_BEMERKUNG))
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        wsBK.Rows(BK_START_ROW & ":" & lastRowBK).AutoFit
        wsBK.Range(wsBK.Cells(BK_START_ROW, BK_COL_BETRAG), wsBK.Cells(lastRowBK, BK_COL_BETRAG)).NumberFormat = euroFormat
        wsBK.Range(wsBK.Cells(BK_START_ROW, BK_COL_MITGL_BEITR), wsBK.Cells(lastRowBK, BK_COL_AUSZAHL_KASSE)).NumberFormat = euroFormat
        wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    On Error Resume Next
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    On Error GoTo ErrorHandler
    If Not wsM Is Nothing Then
        On Error Resume Next
        wsM.Unprotect PASSWORD:=PASSWORD
        wsM.Cells.VerticalAlignment = xlCenter
        wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        On Error GoTo ErrorHandler
    End If
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not wsD Is Nothing Then wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    If Not wsBK Is Nothing Then wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    If Not wsM Is Nothing Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
End Sub

Public Sub FormatiereBlattDaten()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    ws.Cells.VerticalAlignment = xlCenter
    Call FormatiereKategorieTabelle(ws)
    Call FormatiereEntityKeyTabelleKomplett(ws)
    Call AktualisiereKategorieDropdownListen(ws)
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Formatierung des Daten-Blatts abgeschlossen!", vbInformation
    Exit Sub
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler bei der Formatierung: " & Err.Description, vbCritical
End Sub

Public Sub FormatiereKategorieTabelle(Optional ByRef ws As Worksheet = Nothing)
    Dim lastRow As Long, rngTable As Range, r As Long, einAusWert As String
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    Set rngTable = ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), ws.Cells(lastRow, DATA_CAT_COL_END))
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    rngTable.VerticalAlignment = xlCenter
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KATEGORIE), ws.Cells(lastRow, DATA_CAT_COL_KATEGORIE)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_EINAUS), ws.Cells(lastRow, DATA_CAT_COL_EINAUS)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KEYWORD), ws.Cells(lastRow, DATA_CAT_COL_KEYWORD)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_PRIORITAET), ws.Cells(lastRow, DATA_CAT_COL_PRIORITAET)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_ZIELSPALTE), ws.Cells(lastRow, DATA_CAT_COL_ZIELSPALTE)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_FAELLIGKEIT), ws.Cells(lastRow, DATA_CAT_COL_FAELLIGKEIT)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KOMMENTAR), ws.Cells(lastRow, DATA_CAT_COL_KOMMENTAR)).HorizontalAlignment = xlLeft
    For r = DATA_START_ROW To lastRow
        einAusWert = UCase(Trim(ws.Cells(r, DATA_CAT_COL_EINAUS).value))
        Call SetzeZielspalteDropdown(ws, r, einAusWert)
    Next r
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), ws.Cells(lastRow, DATA_CAT_COL_END)).EntireColumn.AutoFit
End Sub

Private Sub SetzeZielspalteDropdown(ByRef ws As Worksheet, ByVal zeile As Long, ByVal einAus As String)
    Dim dropdownSource As String
    On Error Resume Next
    ws.Cells(zeile, DATA_CAT_COL_ZIELSPALTE).Validation.Delete
    On Error GoTo 0
    Select Case einAus
        Case "E": dropdownSource = "=" & WS_BANKKONTO & "!$M$27:$S$27"
        Case "A": dropdownSource = "=" & WS_BANKKONTO & "!$T$27:$Z$27"
        Case Else: dropdownSource = "=" & WS_BANKKONTO & "!$M$27:$Z$27"
    End Select
    On Error Resume Next
    With ws.Cells(zeile, DATA_CAT_COL_ZIELSPALTE).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Formula1:=dropdownSource
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    On Error GoTo 0
End Sub

Public Sub AktualisiereKategorieDropdownListen(Optional ByRef ws As Worksheet = Nothing)
    Dim lastRow As Long, r As Long, kategorie As String, einAus As String
    Dim dictEinnahmen As Object, dictAusgaben As Object, key As Variant
    Dim nextRowE As Long, nextRowA As Long
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    Set dictEinnahmen = CreateObject("Scripting.Dictionary")
    Set dictAusgaben = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    For r = DATA_START_ROW To lastRow
        kategorie = Trim(ws.Cells(r, DATA_CAT_COL_KATEGORIE).value)
        einAus = UCase(Trim(ws.Cells(r, DATA_CAT_COL_EINAUS).value))
        If kategorie <> "" Then
            If einAus = "E" Then
                If Not dictEinnahmen.Exists(kategorie) Then dictEinnahmen.Add kategorie, kategorie
            ElseIf einAus = "A" Then
                If Not dictAusgaben.Exists(kategorie) Then dictAusgaben.Add kategorie, kategorie
            End If
        End If
    Next r
    On Error Resume Next
    ws.Range("AF4:AF1000").ClearContents
    ws.Range("AG4:AG1000").ClearContents
    ws.Range("AH4:AH1000").ClearContents
    On Error GoTo 0
    nextRowE = 4
    For Each key In dictEinnahmen.Keys
        ws.Cells(nextRowE, DATA_COL_KAT_EINNAHMEN).value = key
        nextRowE = nextRowE + 1
    Next key
    nextRowA = 4
    For Each key In dictAusgaben.Keys
        ws.Cells(nextRowA, DATA_COL_KAT_AUSGABEN).value = key
        nextRowA = nextRowA + 1
    Next key
    ws.Cells(4, DATA_COL_MONAT_PERIODE).value = "Januar"
    ws.Cells(5, DATA_COL_MONAT_PERIODE).value = "Februar"
    ws.Cells(6, DATA_COL_MONAT_PERIODE).value = "Maerz"
    ws.Cells(7, DATA_COL_MONAT_PERIODE).value = "April"
    ws.Cells(8, DATA_COL_MONAT_PERIODE).value = "Mai"
    ws.Cells(9, DATA_COL_MONAT_PERIODE).value = "Juni"
    ws.Cells(10, DATA_COL_MONAT_PERIODE).value = "Juli"
    ws.Cells(11, DATA_COL_MONAT_PERIODE).value = "August"
    ws.Cells(12, DATA_COL_MONAT_PERIODE).value = "September"
    ws.Cells(13, DATA_COL_MONAT_PERIODE).value = "Oktober"
    ws.Cells(14, DATA_COL_MONAT_PERIODE).value = "November"
    ws.Cells(15, DATA_COL_MONAT_PERIODE).value = "Dezember"
    Call ErstelleKategorieNamedRanges(ws, nextRowE - 1, nextRowA - 1)
End Sub

Private Sub ErstelleKategorieNamedRanges(ByRef ws As Worksheet, ByVal lastRowE As Long, ByVal lastRowA As Long)
    On Error Resume Next
    ThisWorkbook.Names("lst_KategorienEinnahmen").Delete
    ThisWorkbook.Names("lst_KategorienAusgaben").Delete
    ThisWorkbook.Names("lst_MonatPeriode").Delete
    If lastRowE >= 4 Then
        ThisWorkbook.Names.Add Name:="lst_KategorienEinnahmen", RefersTo:="=" & ws.Name & "!$AF$4:$AF$" & lastRowE
    Else
        ThisWorkbook.Names.Add Name:="lst_KategorienEinnahmen", RefersTo:="=" & ws.Name & "!$AF$4"
    End If
    If lastRowA >= 4 Then
        ThisWorkbook.Names.Add Name:="lst_KategorienAusgaben", RefersTo:="=" & ws.Name & "!$AG$4:$AG$" & lastRowA
    Else
        ThisWorkbook.Names.Add Name:="lst_KategorienAusgaben", RefersTo:="=" & ws.Name & "!$AG$4"
    End If
    ThisWorkbook.Names.Add Name:="lst_MonatPeriode", RefersTo:="=" & ws.Name & "!$AH$4:$AH$15"
    On Error GoTo 0
End Sub

Private Sub FormatiereEntityKeyTabelleKomplett(ByRef ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    If lastRow >= EK_START_ROW Then Call FormatiereEntityKeyTabelle(ws, lastRow)
End Sub

Private Sub FormatiereEntityKeyTabelle(ByRef ws As Worksheet, ByVal lastRow As Long)
    Dim rngTable As Range, rngZebra As Range, r As Long, currentRole As String
    If lastRow < EK_START_ROW Then Exit Sub
    Set rngTable = ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), ws.Cells(lastRow, EK_COL_DEBUG))
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    rngTable.VerticalAlignment = xlCenter
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), ws.Cells(lastRow, EK_COL_ENTITYKEY))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ENTITYKEY).ColumnWidth = 9
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_IBAN), ws.Cells(lastRow, EK_COL_IBAN))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_IBAN).ColumnWidth = 23
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_KONTONAME), ws.Cells(lastRow, EK_COL_KONTONAME))
        .WrapText = True
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_KONTONAME).ColumnWidth = 50
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ZUORDNUNG), ws.Cells(lastRow, EK_COL_ZUORDNUNG))
        .WrapText = True
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ZUORDNUNG).ColumnWidth = 30
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_PARZELLE), ws.Cells(lastRow, EK_COL_PARZELLE))
        .WrapText = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns(EK_COL_PARZELLE).ColumnWidth = 10
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ROLE), ws.Cells(lastRow, EK_COL_ROLE))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ROLE).AutoFit
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_DEBUG), ws.Cells(lastRow, EK_COL_DEBUG))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_DEBUG).AutoFit
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), ws.Cells(lastRow, EK_COL_KONTONAME)).Locked = True
    For r = EK_START_ROW To lastRow
        currentRole = Trim(ws.Cells(r, EK_COL_ROLE).value)
        Call SetzeZellschutzFuerZeile(ws, r, currentRole)
        Set rngZebra = ws.Range(ws.Cells(r, EK_COL_ENTITYKEY), ws.Cells(r, EK_COL_KONTONAME))
        If (r - EK_START_ROW) Mod 2 = 1 Then
            rngZebra.Interior.color = ZEBRA_COLOR
        Else
            rngZebra.Interior.ColorIndex = xlNone
        End If
    Next r
    ws.Rows(EK_START_ROW & ":" & lastRow).AutoFit
End Sub

Private Sub SetzeZellschutzFuerZeile(ByRef ws As Worksheet, ByVal zeile As Long, ByVal currentRole As String)
    Dim rngEditierbar As Range, rngGesperrt As Range
    On Error Resume Next
    Set rngGesperrt = ws.Range(ws.Cells(zeile, EK_COL_ENTITYKEY), ws.Cells(zeile, EK_COL_KONTONAME))
    rngGesperrt.Locked = True
    Set rngEditierbar = ws.Range(ws.Cells(zeile, EK_COL_ZUORDNUNG), ws.Cells(zeile, EK_COL_DEBUG))
    Select Case UCase(Trim(currentRole))
        Case "MITGLIED"
            ws.Cells(zeile, EK_COL_ZUORDNUNG).Locked = True
            ws.Cells(zeile, EK_COL_PARZELLE).Locked = True
            ws.Cells(zeile, EK_COL_ROLE).Locked = True
            ws.Cells(zeile, EK_COL_DEBUG).Locked = False
        Case "UNBEKANNT", ""
            rngEditierbar.Locked = False
        Case Else
            ws.Cells(zeile, EK_COL_ZUORDNUNG).Locked = False
            ws.Cells(zeile, EK_COL_PARZELLE).Locked = True
            ws.Cells(zeile, EK_COL_ROLE).Locked = True
            ws.Cells(zeile, EK_COL_DEBUG).Locked = False
    End Select
    On Error GoTo 0
End Sub




'--- Ende Teil 1 von 3 ---
'--- Anfang Teil 2 von 3 ---




Public Sub FormatiereBlattBankkonto()
    Dim ws As Worksheet, lastRow As Long, euroFormat As String
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    euroFormat = "#,##0.00 " & ChrW(8364)
    Application.ScreenUpdating = False
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    ws.Cells.VerticalAlignment = xlCenter
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then lastRow = BK_START_ROW
    With ws.Range(ws.Cells(BK_START_ROW, BK_COL_BEMERKUNG), ws.Cells(lastRow, BK_COL_BEMERKUNG))
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(BK_START_ROW & ":" & lastRow).AutoFit
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_BETRAG), ws.Cells(lastRow, BK_COL_BETRAG)).NumberFormat = euroFormat
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_MITGL_BEITR), ws.Cells(lastRow, BK_COL_AUSZAHL_KASSE)).NumberFormat = euroFormat
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Application.ScreenUpdating = True
    MsgBox "Formatierung des Bankkonto-Blatts abgeschlossen!", vbInformation
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler: " & Err.Description, vbCritical
End Sub

Public Function NamedRangeExists(ByVal rangeName As String) As Boolean
    Dim nm As Name
    NamedRangeExists = False
    On Error Resume Next
    Set nm = ThisWorkbook.Names(rangeName)
    If Not nm Is Nothing Then NamedRangeExists = True
    On Error GoTo 0
End Function

Public Sub OnKategorieChange(ByVal Target As Range)
    Dim ws As Worksheet, einAus As String
    Set ws = Target.Worksheet
    If Target.Column = DATA_CAT_COL_KATEGORIE Or Target.Column = DATA_CAT_COL_EINAUS Then
        Call AktualisiereKategorieDropdownListen(ws)
    End If
    If Target.Column = DATA_CAT_COL_EINAUS Then
        einAus = UCase(Trim(Target.value))
        Call SetzeZielspalteDropdown(ws, Target.Row, einAus)
    End If
End Sub

Public Sub ApplyDatenSheetFormatting(Optional ByVal formatType As String = "ALL")
    Dim ws As Worksheet, wsRef As Worksheet
    Dim lastRowB As Long, lastRowD As Long, lastRowF As Long, lastRowH As Long
    Dim lastRowKat As Long, lastRowEK As Long
    Dim lastRowZ As Long, lastRowAA As Long, lastRowAB As Long, lastRowAC As Long
    Dim lastRowAD As Long, lastRowAE As Long, lastRowAF As Long, lastRowAG As Long, lastRowAH As Long
    Dim r As Long, zebraColor1 As Long, zebraColor2 As Long, rngRow As Range
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    Set wsRef = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If wsRef Is Nothing Then Set wsRef = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    On Error GoTo ErrorHandler
    If Not wsRef Is Nothing Then
        zebraColor1 = wsRef.Range("A6").Interior.color
        zebraColor2 = wsRef.Range("A7").Interior.color
        If zebraColor1 = 16777215 Then zebraColor1 = RGB(255, 255, 255)
        If zebraColor2 = 16777215 Then zebraColor2 = ZEBRA_COLOR
    Else
        zebraColor1 = RGB(255, 255, 255)
        zebraColor2 = ZEBRA_COLOR
    End If
    Debug.Print "=== ApplyDatenSheetFormatting gestartet ==="
    Debug.Print "Zebra-Farbe 1 (gerade): " & Hex(zebraColor1)
    Debug.Print "Zebra-Farbe 2 (ungerade): " & Hex(zebraColor2)
    lastRowB = GetLastDataRowSafe(ws, 2)
    lastRowD = GetLastDataRowSafe(ws, 4)
    lastRowF = GetLastDataRowSafe(ws, 6)
    lastRowH = GetLastDataRowSafe(ws, 8)
    lastRowKat = GetLastDataRowSafe(ws, 10)
    lastRowEK = GetLastDataRowSafe(ws, 18)
    lastRowZ = GetLastDataRowSafe(ws, 26)
    lastRowAA = GetLastDataRowSafe(ws, 27)
    lastRowAB = GetLastDataRowSafe(ws, 28)
    lastRowAC = GetLastDataRowSafe(ws, 29)
    lastRowAD = GetLastDataRowSafe(ws, 30)
    lastRowAE = GetLastDataRowSafe(ws, 31)
    lastRowAF = GetLastDataRowSafe(ws, 32)
    lastRowAG = GetLastDataRowSafe(ws, 33)
    lastRowAH = GetLastDataRowSafe(ws, 34)
    Debug.Print "LastRows: B=" & lastRowB & ", D=" & lastRowD & ", F=" & lastRowF & ", H=" & lastRowH & ", Kat=" & lastRowKat & ", EK=" & lastRowEK
    Debug.Print "Helper: Z=" & lastRowZ & ", AA=" & lastRowAA & ", AB=" & lastRowAB & ", AC=" & lastRowAC & ", AD=" & lastRowAD & ", AE=" & lastRowAE & ", AF=" & lastRowAF & ", AG=" & lastRowAG & ", AH=" & lastRowAH
    If formatType = "ALL" Or formatType = "WIDTHS" Then
        Debug.Print "Setze Spaltenbreiten..."
        ws.Columns("R").ColumnWidth = 11#
        ws.Columns("T").ColumnWidth = 36#
        ws.Columns("U").ColumnWidth = 28#
        ws.Columns("V").ColumnWidth = 9#
        ws.Columns("X").ColumnWidth = 42#
        ws.Columns("B").AutoFit
        ws.Columns("D").AutoFit
        ws.Columns("F").AutoFit
        ws.Columns("H").AutoFit
        ws.Columns("J").AutoFit
        ws.Columns("K").AutoFit
        ws.Columns("L").AutoFit
        ws.Columns("M").AutoFit
        ws.Columns("N").AutoFit
        ws.Columns("O").AutoFit
        ws.Columns("P").AutoFit
        ws.Columns("S").AutoFit
        ws.Columns("W").AutoFit
        ws.Columns("Z").AutoFit
        ws.Columns("AA").AutoFit
        ws.Columns("AB").AutoFit
        ws.Columns("AC").AutoFit
        ws.Columns("AD").AutoFit
        ws.Columns("AE").AutoFit
        ws.Columns("AF").AutoFit
        ws.Columns("AG").AutoFit
        ws.Columns("AH").AutoFit
    End If
    If formatType = "ALL" Or formatType = "ALIGNMENT" Then
        Debug.Print "Setze Textausrichtung..."
        ws.Columns("F").HorizontalAlignment = xlCenter
        ws.Columns("K").HorizontalAlignment = xlCenter
        ws.Columns("M").HorizontalAlignment = xlCenter
        ws.Columns("V").HorizontalAlignment = xlCenter
        ws.Columns("B").HorizontalAlignment = xlLeft
        ws.Columns("D").HorizontalAlignment = xlLeft
        ws.Columns("H").HorizontalAlignment = xlLeft
        ws.Columns("J").HorizontalAlignment = xlLeft
        ws.Columns("L").HorizontalAlignment = xlLeft
        ws.Columns("N").HorizontalAlignment = xlLeft
        ws.Columns("O").HorizontalAlignment = xlLeft
        ws.Columns("P").HorizontalAlignment = xlLeft
        ws.Columns("R").HorizontalAlignment = xlLeft
        ws.Columns("S").HorizontalAlignment = xlLeft
        ws.Columns("T").HorizontalAlignment = xlLeft
        ws.Columns("U").HorizontalAlignment = xlLeft
        ws.Columns("W").HorizontalAlignment = xlLeft
        ws.Columns("X").HorizontalAlignment = xlLeft
        ws.Cells.VerticalAlignment = xlCenter
    End If
    If formatType = "ALL" Or formatType = "ZEBRA" Then
        Debug.Print "Wende Zebra-Streifen an..."
        If lastRowB >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowB
                If (r Mod 2) = 0 Then ws.Cells(r, 2).Interior.color = zebraColor1 Else ws.Cells(r, 2).Interior.color = zebraColor2
            Next r
        End If
        If lastRowD >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowD
                If (r Mod 2) = 0 Then ws.Cells(r, 4).Interior.color = zebraColor1 Else ws.Cells(r, 4).Interior.color = zebraColor2
            Next r
        End If
        If lastRowF >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowF
                If (r Mod 2) = 0 Then ws.Cells(r, 6).Interior.color = zebraColor1 Else ws.Cells(r, 6).Interior.color = zebraColor2
            Next r
        End If
        If lastRowH >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowH
                If (r Mod 2) = 0 Then ws.Cells(r, 8).Interior.color = zebraColor1 Else ws.Cells(r, 8).Interior.color = zebraColor2
            Next r
        End If
        If lastRowKat >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowKat
                Set rngRow = ws.Range(ws.Cells(r, 10), ws.Cells(r, 16))
                If (r Mod 2) = 0 Then rngRow.Interior.color = zebraColor1 Else rngRow.Interior.color = zebraColor2
            Next r
        End If
        If lastRowEK >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowEK
                Set rngRow = ws.Range(ws.Cells(r, 18), ws.Cells(r, 20))
                If (r Mod 2) = 0 Then rngRow.Interior.color = zebraColor1 Else rngRow.Interior.color = zebraColor2
            Next r
            Debug.Print "HINWEIS: Spalten U-X wurden NICHT mit Zebra formatiert (Ampel-Bereich)"
        End If
        If lastRowZ >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowZ
                If (r Mod 2) = 0 Then ws.Cells(r, 26).Interior.color = zebraColor1 Else ws.Cells(r, 26).Interior.color = zebraColor2
            Next r
            Debug.Print "Spalte Z formatiert bis Zeile " & lastRowZ
        End If
        If lastRowAA >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowAA
                If (r Mod 2) = 0 Then ws.Cells(r, 27).Interior.color = zebraColor1 Else ws.Cells(r, 27).Interior.color = zebraColor2
            Next r
            Debug.Print "Spalte AA formatiert bis Zeile " & lastRowAA
        End If
        If lastRowAB >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowAB
                If (r Mod 2) = 0 Then ws.Cells(r, 28).Interior.color = zebraColor1 Else ws.Cells(r, 28).Interior.color = zebraColor2
            Next r
            Debug.Print "Spalte AB formatiert bis Zeile " & lastRowAB
        End If
        If lastRowAC >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowAC
                If (r Mod 2) = 0 Then ws.Cells(r, 29).Interior.color = zebraColor1 Else ws.Cells(r, 29).Interior.color = zebraColor2
            Next r
            Debug.Print "Spalte AC formatiert bis Zeile " & lastRowAC
        End If
        If lastRowAD >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowAD
                If (r Mod 2) = 0 Then ws.Cells(r, 30).Interior.color = zebraColor1 Else ws.Cells(r, 30).Interior.color = zebraColor2
            Next r
            Debug.Print "Spalte AD formatiert bis Zeile " & lastRowAD
        End If
        If lastRowAE >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowAE
                If (r Mod 2) = 0 Then ws.Cells(r, 31).Interior.color = zebraColor1 Else ws.Cells(r, 31).Interior.color = zebraColor2
            Next r
            Debug.Print "Spalte AE formatiert bis Zeile " & lastRowAE
        End If
        If lastRowAF >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowAF
                If (r Mod 2) = 0 Then ws.Cells(r, 32).Interior.color = zebraColor1 Else ws.Cells(r, 32).Interior.color = zebraColor2
            Next r
            Debug.Print "Spalte AF formatiert bis Zeile " & lastRowAF
        End If
        If lastRowAG >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowAG
                If (r Mod 2) = 0 Then ws.Cells(r, 33).Interior.color = zebraColor1 Else ws.Cells(r, 33).Interior.color = zebraColor2
            Next r
            Debug.Print "Spalte AG formatiert bis Zeile " & lastRowAG
        End If
        If lastRowAH >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowAH
                If (r Mod 2) = 0 Then ws.Cells(r, 34).Interior.color = zebraColor1 Else ws.Cells(r, 34).Interior.color = zebraColor2
            Next r
            Debug.Print "Spalte AH formatiert bis Zeile " & lastRowAH
        End If
    End If
    If formatType = "ALL" Or formatType = "BORDERS" Then
        Debug.Print "Wende Rahmen an..."
        If lastRowB >= DATA_START_ROW Then Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowB, 2, 2)
        If lastRowD >= DATA_START_ROW Then Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowD, 4, 4)
        If lastRowF >= DATA_START_ROW Then Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowF, 6, 6)
        If lastRowH >= DATA_START_ROW Then Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowH, 8, 8)
        If lastRowKat >= DATA_START_ROW Then Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowKat, 10, 16)
        If lastRowEK >= DATA_START_ROW Then Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowEK, 18, 24)
        If lastRowZ >= DATA_START_ROW Then Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowZ, 26, 26)
        If lastRowAA >= DATA_START_ROW Then Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowAA, 27, 27)
        If lastRowAB >= DATA_START_ROW Then Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowAB, 28, 28)
        If lastRowAC >= DATA_START_ROW Then Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowAC, 29, 29)
        If lastRowAD >= DATA_START_ROW Then Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowAD, 30, 30)
        If lastRowAE >= DATA_START_ROW Then Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowAE, 31, 31)
        If lastRowAF >= DATA_START_ROW Then Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowAF, 32, 32)
        If lastRowAG >= DATA_START_ROW Then Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowAG, 33, 33)
        If lastRowAH >= DATA_START_ROW Then Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowAH, 34, 34)
    End If
    If formatType = "ALL" Or formatType = "FILTER" Then
        Debug.Print "Aktiviere AutoFilter..."
        On Error Resume Next
        ws.AutoFilterMode = False
        On Error GoTo ErrorHandler
        If lastRowKat >= DATA_START_ROW Then ws.Range(ws.Cells(DATA_HEADER_ROW, 10), ws.Cells(lastRowKat, 16)).AutoFilter
        If lastRowEK >= DATA_START_ROW Then ws.Range(ws.Cells(DATA_HEADER_ROW, 18), ws.Cells(lastRowEK, 24)).AutoFilter
    End If
    On Error Resume Next
    ws.Range("Y100").WrapText = True
    On Error GoTo ErrorHandler
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Debug.Print "=== ApplyDatenSheetFormatting erfolgreich beendet ==="
    MsgBox "Formatierung des Daten-Blatts abgeschlossen!", vbInformation, "Formatierung"
    Exit Sub
ErrorHandler:
    Debug.Print "FEHLER in ApplyDatenSheetFormatting: " & Err.Description
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Formatierungsfehler: " & Err.Description, vbCritical, "Fehler"
End Sub




'--- Ende Teil 2 von 3 ---
'--- Anfang Teil 3 von 3 ---



Private Function GetLastDataRowSafe(ByRef ws As Worksheet, ByVal col As Long) As Long
    Dim lastRow As Long
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
    On Error GoTo 0
    If lastRow < DATA_START_ROW Then lastRow = DATA_HEADER_ROW
    GetLastDataRowSafe = lastRow
End Function

Private Sub ApplyBordersToTableRange(ByRef ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal startCol As Long, ByVal endCol As Long)
    Dim rngTable As Range
    On Error Resume Next
    If endRow < startRow Then Exit Sub
    Set rngTable = ws.Range(ws.Cells(startRow, startCol), ws.Cells(endRow, endCol))
    With rngTable.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = RGB(0, 0, 0)
    End With
    With rngTable.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = RGB(0, 0, 0)
    End With
    With rngTable.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = RGB(0, 0, 0)
    End With
    With rngTable.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = RGB(0, 0, 0)
    End With
    With rngTable.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = RGB(0, 0, 0)
    End With
    With rngTable.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = RGB(0, 0, 0)
    End With
    On Error GoTo 0
End Sub

Public Sub FormatSingleColumn(ByRef ws As Worksheet, ByVal col As Long, Optional ByVal specificRow As Long = 0)
    Dim lastRow As Long, r As Long, startRow As Long, endRow As Long
    Dim zebraColor1 As Long, zebraColor2 As Long
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
    If lastRow < DATA_START_ROW Then
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        Exit Sub
    End If
    zebraColor1 = RGB(255, 255, 255)
    zebraColor2 = ZEBRA_COLOR
    startRow = DATA_START_ROW
    endRow = lastRow
    For r = startRow To endRow
        If (r Mod 2) = 0 Then ws.Cells(r, col).Interior.color = zebraColor1 Else ws.Cells(r, col).Interior.color = zebraColor2
        With ws.Cells(r, col).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = RGB(0, 0, 0)
        End With
    Next r
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Public Sub FormatHelperColumnIndependent(ByRef ws As Worksheet, ByVal col As Long)
    Dim lastRow As Long, r As Long, zebraColor1 As Long, zebraColor2 As Long
    On Error Resume Next
    If col < 26 Or col > 34 Then Exit Sub
    ws.Unprotect PASSWORD:=PASSWORD
    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
    If lastRow < DATA_START_ROW Then
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        Exit Sub
    End If
    zebraColor1 = RGB(255, 255, 255)
    zebraColor2 = ZEBRA_COLOR
    For r = DATA_START_ROW To lastRow
        If (r Mod 2) = 0 Then ws.Cells(r, col).Interior.color = zebraColor1 Else ws.Cells(r, col).Interior.color = zebraColor2
        With ws.Cells(r, col).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = RGB(0, 0, 0)
        End With
    Next r
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Public Sub SortKategorieTable(ByRef ws As Worksheet)
    Dim lastRow As Long, rngSort As Range
    On Error GoTo ErrorHandler
    lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    Set rngSort = ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), ws.Cells(lastRow, DATA_CAT_COL_END))
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KATEGORIE), ws.Cells(lastRow, DATA_CAT_COL_KATEGORIE)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange rngSort
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Exit Sub
ErrorHandler:
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
End Sub

Public Sub SortEntityKeyTable(ByRef ws As Worksheet)
    Dim lastRow As Long, rngData As Range, arrData() As Variant, arrKeys() As Long
    Dim i As Long, j As Long, entityKey As String, parzNum As Long
    Dim numRows As Long, numCols As Long, minIdx As Long, minKey As Long
    Dim tempKey As Long, tempVal As Variant
    On Error GoTo ErrorHandler
    lastRow = ws.Cells(ws.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    If lastRow < EK_START_ROW Then Exit Sub
    Set rngData = ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), ws.Cells(lastRow, EK_COL_DEBUG))
    arrData = rngData.value
    numRows = UBound(arrData, 1)
    numCols = UBound(arrData, 2)
    ReDim arrKeys(1 To numRows)
    For i = 1 To numRows
        entityKey = UCase(Trim(CStr(arrData(i, 1))))
        If Left(entityKey, 2) = "P-" Then
            parzNum = 0
            On Error Resume Next
            parzNum = CLng(Mid(entityKey, 3))
            On Error GoTo ErrorHandler
            If parzNum >= 1 And parzNum <= 14 Then arrKeys(i) = parzNum Else arrKeys(i) = 100 + parzNum
        ElseIf Left(entityKey, 3) = "EX-" Then
            arrKeys(i) = 15 + (Asc(Mid(entityKey, 4, 1)) Mod 15)
        ElseIf Left(entityKey, 5) = "VERS-" Then
            arrKeys(i) = 30 + (Asc(Mid(entityKey, 6, 1)) Mod 15)
        ElseIf Left(entityKey, 5) = "BANK-" Then
            arrKeys(i) = 45 + (Asc(Mid(entityKey, 6, 1)) Mod 15)
        Else
            arrKeys(i) = 60 + (Asc(Left(entityKey, 1)) Mod 40)
        End If
    Next i
    For i = 1 To numRows - 1
        minIdx = i
        minKey = arrKeys(i)
        For j = i + 1 To numRows
            If arrKeys(j) < minKey Then
                minKey = arrKeys(j)
                minIdx = j
            End If
        Next j
        If minIdx <> i Then
            tempKey = arrKeys(i)
            arrKeys(i) = arrKeys(minIdx)
            arrKeys(minIdx) = tempKey
            For j = 1 To numCols
                tempVal = arrData(i, j)
                arrData(i, j) = arrData(minIdx, j)
                arrData(minIdx, j) = tempVal
            Next j
        End If
    Next i
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    rngData.value = arrData
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Call RefreshEntityKeyZebra(ws, lastRow)
    Exit Sub
ErrorHandler:
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
End Sub

Public Sub RefreshEntityKeyZebra(ByRef ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long, rngZebra As Range, zebraColor1 As Long, zebraColor2 As Long
    On Error Resume Next
    zebraColor1 = RGB(255, 255, 255)
    zebraColor2 = ZEBRA_COLOR
    For r = EK_START_ROW To lastRow
        Set rngZebra = ws.Range(ws.Cells(r, EK_COL_ENTITYKEY), ws.Cells(r, EK_COL_KONTONAME))
        If (r Mod 2) = 0 Then rngZebra.Interior.color = zebraColor1 Else rngZebra.Interior.color = zebraColor2
    Next r
    On Error GoTo 0
End Sub

Public Sub UpdateEntityRoleDropdown(ByRef ws As Worksheet)
    Dim lastRowEK As Long, dictRoles As Object, r As Long, roleValue As String, nextRow As Long, key As Variant
    On Error Resume Next
    Set dictRoles = CreateObject("Scripting.Dictionary")
    dictRoles.Add "MITGLIED", "MITGLIED"
    dictRoles.Add "EXMITGLIED", "EXMITGLIED"
    dictRoles.Add "VERSORGER", "VERSORGER"
    dictRoles.Add "BANK", "BANK"
    dictRoles.Add "BEHOERDE", "BEHOERDE"
    dictRoles.Add "DIENSTLEISTER", "DIENSTLEISTER"
    dictRoles.Add "KUNDE", "KUNDE"
    dictRoles.Add "LIEFERANT", "LIEFERANT"
    dictRoles.Add "UNBEKANNT", "UNBEKANNT"
    lastRowEK = ws.Cells(ws.Rows.count, EK_COL_ROLE).End(xlUp).Row
    If lastRowEK >= EK_START_ROW Then
        For r = EK_START_ROW To lastRowEK
            roleValue = UCase(Trim(ws.Cells(r, EK_COL_ROLE).value))
            If roleValue <> "" Then
                If Not dictRoles.Exists(roleValue) Then dictRoles.Add roleValue, roleValue
            End If
        Next r
    End If
    ws.Range("AD4:AD100").ClearContents
    nextRow = 4
    For Each key In dictRoles.Keys
        ws.Cells(nextRow, DATA_COL_DD_ENTITYROLE).value = key
        nextRow = nextRow + 1
    Next key
    On Error GoTo 0
End Sub

Public Sub RefreshTableFormattingAfterChange(ByRef ws As Worksheet, ByVal changedCol As Long, ByVal changedRow As Long)
    On Error Resume Next
    If changedCol = 2 Or changedCol = 4 Or changedCol = 6 Or changedCol = 8 Then
        Call FormatSingleColumn(ws, changedCol, changedRow)
        Exit Sub
    End If
    If changedCol >= 10 And changedCol <= 16 Then
        Call FormatiereKategorieZeile(ws, changedRow)
        Exit Sub
    End If
    If changedCol >= 18 And changedCol <= 20 Then
        Call FormatiereEntityKeyZeileZebra(ws, changedRow)
        Exit Sub
    End If
    If changedCol >= 26 And changedCol <= 34 Then
        Call FormatHelperColumnIndependent(ws, changedCol)
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Public Sub FormatiereKategorieZeile(ByRef ws As Worksheet, ByVal zeile As Long)
    Dim rngRow As Range, zebraColor1 As Long, zebraColor2 As Long
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    zebraColor1 = RGB(255, 255, 255)
    zebraColor2 = ZEBRA_COLOR
    Set rngRow = ws.Range(ws.Cells(zeile, 10), ws.Cells(zeile, 16))
    If (zeile Mod 2) = 0 Then rngRow.Interior.color = zebraColor1 Else rngRow.Interior.color = zebraColor2
    With rngRow.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = RGB(0, 0, 0)
    End With
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Public Sub FormatiereEntityKeyZeileZebra(ByRef ws As Worksheet, ByVal zeile As Long)
    Dim rngRow As Range, zebraColor1 As Long, zebraColor2 As Long
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    zebraColor1 = RGB(255, 255, 255)
    zebraColor2 = ZEBRA_COLOR
    Set rngRow = ws.Range(ws.Cells(zeile, 18), ws.Cells(zeile, 20))
    If (zeile Mod 2) = 0 Then rngRow.Interior.color = zebraColor1 Else rngRow.Interior.color = zebraColor2
    With rngRow.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = RGB(0, 0, 0)
    End With
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub
