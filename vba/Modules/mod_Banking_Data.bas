Attribute VB_Name = "mod_Banking_Data"
Option Explicit

' ===============================================================
' MODUL: mod_Banking_Data (FINAL KONSOLIDIERT)
' VERSION: 2.1 - 02.02.2026
' KORREKTUR: Blattschutz vor Import aufheben
' ===============================================================

Private Const AMPEL_GRUEN As Long = 13561798
Private Const AMPEL_GELB As Long = 10025215
Private Const AMPEL_ROT As Long = 13551359
Private Const AMPEL_WEISS As Long = 16777215
Private Const ZEBRA_COLOR As Long = &HDEE5E3
Private Const RAHMEN_NAME As String = "ImportReport_Rahmen"

' ===============================================================
' 1. CSV-KONTOAUSZUG IMPORT
' ===============================================================
Public Sub Importiere_Kontoauszug()
    Const xlUTF8Value As Long = 65001
    Const xlDelimitedValue As Long = 1
    
    Dim wsZiel As Worksheet
    Dim wsTemp As Worksheet
    Dim dictUmsaetze As Object
    Dim strFile As Variant
    Dim lRowZiel As Long, i As Long
    Dim lRowTemp As Long, lastRowTemp As Long
    
    Dim sKey As String
    Dim dBetrag As Double
    Dim betragString As String
    Dim sIBAN As String, sText As String, sName As String, sVZ As String
    Dim tempSheetName As String
    Dim dDatum As Date
    Dim sFormelAuswertungsmonat As String
    
    Dim rowsProcessed As Long
    Dim rowsIgnoredDupe As Long
    Dim rowsIgnoredFilter As Long
    Dim rowsFailedImport As Long
    Dim rowsTotalInFile As Long
    
    tempSheetName = "TempImport"
    
    Set wsZiel = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    On Error Resume Next
    wsZiel.Unprotect PASSWORD:=PASSWORD
    Err.Clear
    On Error GoTo 0
    
    Set dictUmsaetze = CreateObject("Scripting.Dictionary")
    
    rowsProcessed = 0
    rowsIgnoredDupe = 0
    rowsIgnoredFilter = 0
    rowsFailedImport = 0
    rowsTotalInFile = 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error Resume Next
    ThisWorkbook.Worksheets(tempSheetName).Delete
    Err.Clear
    On Error GoTo 0
    
    Application.DisplayAlerts = True
    strFile = Application.GetOpenFilename("CSV (*.csv), *.csv")
    Application.DisplayAlerts = False
    
    If strFile = False Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        wsZiel.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        Call Initialize_ImportReport_ListBox
        Exit Sub
    End If
    
    lRowZiel = wsZiel.Cells(wsZiel.Rows.Count, BK_COL_BETRAG).End(xlUp).Row
    If lRowZiel < BK_START_ROW Then lRowZiel = BK_START_ROW - 1
    
    For i = BK_START_ROW To lRowZiel
        If wsZiel.Cells(i, BK_COL_BETRAG).value <> "" Then
            sKey = Format(wsZiel.Cells(i, BK_COL_DATUM).value, "YYYYMMDD") & "|" & _
                   CStr(wsZiel.Cells(i, BK_COL_BETRAG).value) & "|" & _
                   Replace(CStr(wsZiel.Cells(i, BK_COL_IBAN).value), " ", "") & "|" & _
                   CStr(wsZiel.Cells(i, BK_COL_VERWENDUNGSZWECK).value)
            dictUmsaetze(sKey) = True
        End If
    Next i
    
    On Error Resume Next
    Set wsTemp = ThisWorkbook.Worksheets.Add(After:=wsZiel)
    If Err.Number <> 0 Then
        MsgBox "Fehler beim Erstellen des Temp-Blatts: " & Err.Description, vbCritical
        Err.Clear
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        wsZiel.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        Exit Sub
    End If
    wsTemp.Name = tempSheetName
    Err.Clear
    On Error GoTo 0
    
    On Error Resume Next
    With wsTemp.QueryTables.Add(Connection:="TEXT;" & strFile, Destination:=wsTemp.Cells(1, 1))
        .Name = "CSV_Import"
        .FieldNames = True
        .TextFilePlatform = xlUTF8Value
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimitedValue
        .TextFileSemicolonDelimiter = True
        .Refresh BackgroundQuery:=False
    End With
    
    If Err.Number <> 0 Then
        MsgBox "Fehler beim Einlesen der CSV-Datei: " & Err.Description, vbCritical
        Err.Clear
        Application.DisplayAlerts = False
        wsTemp.Delete
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        wsZiel.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    
    lastRowTemp = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
    rowsTotalInFile = lastRowTemp - 1
    
    If lastRowTemp <= 1 Then
        rowsProcessed = 0
        GoTo ImportAbschluss
    End If
    
    On Error Resume Next
    wsTemp.QueryTables(1).Delete
    Err.Clear
    On Error GoTo 0
    
        
'--- Ende TEIL 1 ---
'--- Anfang Teil 2 ---


    For lRowTemp = 2 To lastRowTemp
        
        betragString = CStr(wsTemp.Cells(lRowTemp, CSV_COL_BETRAG).value)
        
        betragString = Replace(betragString, " EUR", "")
        betragString = Replace(betragString, "EUR", "")
        betragString = Trim(betragString)
        
        If betragString = "" Or Not IsNumeric(Replace(betragString, ",", ".")) Then
             rowsIgnoredFilter = rowsIgnoredFilter + 1
             GoTo NextRowImport
        End If
        
        On Error Resume Next
        dBetrag = CDbl(Replace(betragString, ",", Application.International(xlDecimalSeparator)))
        If Err.Number <> 0 Then
            rowsIgnoredFilter = rowsIgnoredFilter + 1
            Err.Clear
            GoTo NextRowImport
        End If
        On Error GoTo 0
        
        If IsDate(wsTemp.Cells(lRowTemp, CSV_COL_BUCHUNGSDATUM).value) Then
            dDatum = CDate(wsTemp.Cells(lRowTemp, CSV_COL_BUCHUNGSDATUM).value)
        Else
            rowsIgnoredFilter = rowsIgnoredFilter + 1
            GoTo NextRowImport
        End If
        
        sIBAN = Replace(Trim(wsTemp.Cells(lRowTemp, CSV_COL_IBAN).value), " ", "")
        sName = Trim(wsTemp.Cells(lRowTemp, CSV_COL_NAME).value)
        sVZ = Trim(wsTemp.Cells(lRowTemp, CSV_COL_VERWENDUNGSZWECK).value)
        sText = Trim(wsTemp.Cells(lRowTemp, CSV_COL_STATUS).value)
        
        sKey = Format(dDatum, "YYYYMMDD") & "|" & dBetrag & "|" & sIBAN & "|" & sVZ

        If dictUmsaetze.Exists(sKey) Then
            rowsIgnoredDupe = rowsIgnoredDupe + 1
            GoTo NextRowImport
        End If
        
        lRowZiel = wsZiel.Cells(wsZiel.Rows.Count, BK_COL_DATUM).End(xlUp).Row + 1
        dictUmsaetze.Add sKey, True
        
        wsZiel.Cells(lRowZiel, BK_COL_DATUM).value = dDatum
        wsZiel.Cells(lRowZiel, BK_COL_DATUM).NumberFormat = "DD.MM.YYYY"

        wsZiel.Cells(lRowZiel, BK_COL_BETRAG).value = dBetrag
        wsZiel.Cells(lRowZiel, BK_COL_BETRAG).NumberFormat = "#,##0.00 [$EUR]"

        wsZiel.Cells(lRowZiel, BK_COL_NAME).value = sName
        wsZiel.Cells(lRowZiel, BK_COL_IBAN).value = sIBAN
        wsZiel.Cells(lRowZiel, BK_COL_VERWENDUNGSZWECK).value = sVZ
        wsZiel.Cells(lRowZiel, BK_COL_BUCHUNGSTEXT).value = sText
        
        sFormelAuswertungsmonat = "=IF(A" & lRowZiel & "="""","""",IF(Daten!$AE$4=0,TRUE,MONTH(A" & lRowZiel & ")=Daten!$AE$4))"
        wsZiel.Cells(lRowZiel, BK_COL_IM_AUSWERTUNGSMONAT).Formula = sFormelAuswertungsmonat
        
        wsZiel.Cells(lRowZiel, BK_COL_STATUS).value = "Gebucht"
        
        rowsProcessed = rowsProcessed + 1

NextRowImport:
    Next lRowTemp

ImportAbschluss:
    
    rowsFailedImport = rowsIgnoredFilter
    
    Call Update_ImportReport_ListBox(rowsTotalInFile, rowsProcessed, rowsIgnoredDupe, rowsFailedImport)
    
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wsTemp Is Nothing Then wsTemp.Delete
    Application.DisplayAlerts = True
    Set wsTemp = Nothing
    Err.Clear
    On Error GoTo 0
    
    On Error Resume Next
    Call ImportiereIBANsAusBankkonto
    Call Sortiere_Bankkonto_nach_Datum
    Call Anwende_Zebra_Bankkonto(wsZiel)
    Call Anwende_Border_Bankkonto(wsZiel)
    Call Anwende_Formatierung_Bankkonto(wsZiel)
    If rowsProcessed > 0 Then Call KategorieEngine_Pipeline(wsZiel)
    Call Setze_Monat_Periode(wsZiel)
    Err.Clear
    On Error GoTo 0
    
    wsZiel.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    wsZiel.Activate

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    If rowsTotalInFile > 0 And rowsProcessed = 0 And rowsIgnoredDupe = rowsTotalInFile And rowsFailedImport = 0 Then
        MsgBox "Achtung: Die ausgewaehlte CSV-Datei enthaelt ausschliesslich Eintraege, " & _
               "die bereits in der Datenbank vorhanden sind (" & rowsIgnoredDupe & " Duplikate). " & _
               "Es wurden keine neuen Datensaetze importiert.", vbExclamation, "100% Duplikate erkannt"
    ElseIf rowsProcessed > 0 Then
        MsgBox "Import abgeschlossen! (" & rowsProcessed & " neue Zeilen hinzugefuegt)", vbInformation
    End If
    
End Sub

' ===============================================================
' 1b. ZEBRA-FORMATIERUNG
' ===============================================================
Private Sub Anwende_Zebra_Bankkonto(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim lRow As Long
    Dim rngRowPart1 As Range
    Dim rngRowPart2 As Range
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    For lRow = BK_START_ROW To lastRow
        Set rngRowPart1 = ws.Range(ws.Cells(lRow, 1), ws.Cells(lRow, 7))
        Set rngRowPart2 = ws.Range(ws.Cells(lRow, 13), ws.Cells(lRow, 26))
        
        If (lRow - BK_START_ROW) Mod 2 = 1 Then
            rngRowPart1.Interior.color = ZEBRA_COLOR
            rngRowPart2.Interior.color = ZEBRA_COLOR
        Else
            rngRowPart1.Interior.ColorIndex = xlNone
            rngRowPart2.Interior.ColorIndex = xlNone
        End If
    Next lRow
    
End Sub

' ===============================================================
' 1c. RAHMEN-FORMATIERUNG
' ===============================================================
Private Sub Anwende_Border_Bankkonto(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim rngPart1 As Range
    Dim rngPart2 As Range
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, BK_COL_DATUM).End(xlUp).Row
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
' 1d. ALLGEMEINE FORMATIERUNG
' ===============================================================
Private Sub Anwende_Formatierung_Bankkonto(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim euroFormat As String
    
    If ws Is Nothing Then Exit Sub
    
    euroFormat = "#,##0.00 " & ChrW(8364)
    
    lastRow = ws.Cells(ws.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_BETRAG), ws.Cells(lastRow, BK_COL_BETRAG)).NumberFormat = euroFormat
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_MITGL_BEITR), ws.Cells(lastRow, BK_COL_AUSZAHL_KASSE)).NumberFormat = euroFormat
    
    With ws.Range(ws.Cells(BK_START_ROW, BK_COL_BEMERKUNG), ws.Cells(lastRow, BK_COL_BEMERKUNG))
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    
    ws.Cells.VerticalAlignment = xlCenter
    ws.Rows(BK_START_ROW & ":" & lastRow).AutoFit
    
End Sub

' ===============================================================
' 2. SORTIERUNG NACH DATUM
' ===============================================================
Public Sub Sortiere_Bankkonto_nach_Datum()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim sortRange As Range
    
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        Exit Sub
    End If
    
    Set sortRange = ws.Range(ws.Cells(BK_START_ROW, 1), ws.Cells(lastRow, 26))
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(BK_START_ROW, BK_COL_DATUM), ws.Cells(lastRow, BK_COL_DATUM)), _
                           SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    
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
' 3. MONAT/PERIODE SETZEN
' ===============================================================
Private Sub Setze_Monat_Periode(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim r As Long
    Dim monatWert As Variant
    Dim datumWert As Variant
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    For r = BK_START_ROW To lastRow
        datumWert = ws.Cells(r, BK_COL_DATUM).value
        monatWert = ws.Cells(r, BK_COL_MONAT_PERIODE).value
        
        If IsDate(datumWert) And (IsEmpty(monatWert) Or monatWert = "") Then
            ws.Cells(r, BK_COL_MONAT_PERIODE).value = MonthName(Month(datumWert))
        End If
    Next r
    
End Sub

' ===============================================================
' 4. IMPORT REPORT LISTBOX
' ===============================================================
Public Sub Initialize_ImportReport_ListBox()
    
    Dim ws As Worksheet
    Dim shp As Shape
    Dim rahmenFound As Boolean
    
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    rahmenFound = False
    
    For Each shp In ws.Shapes
        If shp.Name = RAHMEN_NAME Then
            rahmenFound = True
            Exit For
        End If
    Next shp
    
    If Not rahmenFound Then Exit Sub
    
    shp.TextFrame2.TextRange.Characters.text = _
        "Import-Bericht:" & vbCrLf & _
        "----------------" & vbCrLf & _
        "Zeilen in CSV: 0" & vbCrLf & _
        "Importiert: 0" & vbCrLf & _
        "Duplikate: 0" & vbCrLf & _
        "Fehler: 0"
    
End Sub

Private Sub Update_ImportReport_ListBox(ByVal totalRows As Long, ByVal imported As Long, _
                                         ByVal dupes As Long, ByVal failed As Long)
    
    Dim ws As Worksheet
    Dim shp As Shape
    Dim rahmenFound As Boolean
    
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    rahmenFound = False
    
    For Each shp In ws.Shapes
        If shp.Name = RAHMEN_NAME Then
            rahmenFound = True
            Exit For
        End If
    Next shp
    
    If Not rahmenFound Then Exit Sub
    
    shp.TextFrame2.TextRange.Characters.text = _
        "Import-Bericht:" & vbCrLf & _
        "----------------" & vbCrLf & _
        "Zeilen in CSV: " & totalRows & vbCrLf & _
        "Importiert: " & imported & vbCrLf & _
        "Duplikate: " & dupes & vbCrLf & _
        "Fehler: " & failed
    
End Sub

' ===============================================================
' 5. IBAN IMPORT AUS BANKKONTO
' ===============================================================
Public Sub ImportiereIBANsAusBankkonto()
    
    Dim wsBK As Worksheet
    Dim wsD As Worksheet
    Dim dictIBANs As Object
    Dim r As Long
    Dim lastRowBK As Long
    Dim lastRowD As Long
    Dim nextRowD As Long
    Dim currentIBAN As String
    Dim currentKontoName As String
    Dim currentDatum As Variant
    Dim anzahlNeu As Long
    Dim anzahlBereitsVorhanden As Long
    
    On Error GoTo ErrorHandler
    
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    
    On Error Resume Next
    wsD.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    Set dictIBANs = CreateObject("Scripting.Dictionary")
    
    lastRowBK = wsBK.Cells(wsBK.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    If lastRowBK < BK_START_ROW Then lastRowBK = BK_START_ROW
    
    lastRowD = wsD.Cells(wsD.Rows.Count, EK_COL_IBAN).End(xlUp).Row
    If lastRowD < EK_START_ROW Then lastRowD = EK_START_ROW - 1
    
    For r = EK_START_ROW To lastRowD
        currentIBAN = Trim(wsD.Cells(r, EK_COL_IBAN).value)
        currentKontoName = Trim(wsD.Cells(r, EK_COL_KONTONAME).value)
        If currentIBAN <> "" Or currentKontoName <> "" Then
            dictIBANs(currentIBAN & "|" & currentKontoName) = True
        End If
    Next r
    
    anzahlBereitsVorhanden = dictIBANs.Count
    nextRowD = lastRowD + 1
    anzahlNeu = 0
    
    For r = BK_START_ROW To lastRowBK
        currentDatum = wsBK.Cells(r, BK_COL_DATUM).value
        If IsEmpty(currentDatum) Or currentDatum = "" Then GoTo NextRowIBAN
        
        currentIBAN = Trim(wsBK.Cells(r, BK_COL_IBAN).value)
        currentKontoName = Trim(wsBK.Cells(r, BK_COL_NAME).value)
        
        If currentIBAN = "" And currentKontoName = "" Then GoTo NextRowIBAN
        
        If Not dictIBANs.Exists(currentIBAN & "|" & currentKontoName) Then
            wsD.Cells(nextRowD, EK_COL_IBAN).value = currentIBAN
            wsD.Cells(nextRowD, EK_COL_KONTONAME).value = currentKontoName
            
            dictIBANs(currentIBAN & "|" & currentKontoName) = True
            nextRowD = nextRowD + 1
            anzahlNeu = anzahlNeu + 1
        End If
        
NextRowIBAN:
    Next r
    
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Debug.Print "Fehler in ImportiereIBANsAusBankkonto: " & Err.Description
End Sub

' ===============================================================
' 6. KATEGORIE ENGINE PIPELINE
' ===============================================================
Public Sub KategorieEngine_Pipeline(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim r As Long
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    For r = BK_START_ROW To lastRow
        If ws.Cells(r, BK_COL_KATEGORIE).value = "" Then
            Call VerarbeiteZeile(ws, r)
        End If
    Next r
    
End Sub

Private Sub VerarbeiteZeile(ByVal ws As Worksheet, ByVal zeile As Long)
    
    Dim suchText As String
    Dim betrag As Double
    Dim kategorie As String
    Dim zielspalte As Long
    
    suchText = LCase(Trim(ws.Cells(zeile, BK_COL_NAME).value) & " " & _
                     Trim(ws.Cells(zeile, BK_COL_VERWENDUNGSZWECK).value) & " " & _
                     Trim(ws.Cells(zeile, BK_COL_BUCHUNGSTEXT).value))
    
    betrag = ws.Cells(zeile, BK_COL_BETRAG).value
    
    kategorie = SucheKategorie(suchText, betrag)
    
    If kategorie <> "" Then
        ws.Cells(zeile, BK_COL_KATEGORIE).value = kategorie
        
        zielspalte = ErmittleZielspalte(kategorie, betrag)
        If zielspalte > 0 Then
            ws.Cells(zeile, zielspalte).value = Abs(betrag)
        End If
    End If
    
End Sub

Private Function SucheKategorie(ByVal suchText As String, ByVal betrag As Double) As String
    
    Dim wsD As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim keyword As String
    Dim kategorie As String
    Dim einAus As String
    Dim prioritaet As Long
    Dim bestePrioritaet As Long
    Dim besteKategorie As String
    
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    
    lastRow = wsD.Cells(wsD.Rows.Count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then
        SucheKategorie = ""
        Exit Function
    End If
    
    bestePrioritaet = 9999
    besteKategorie = ""
    
    For r = DATA_START_ROW To lastRow
        kategorie = Trim(wsD.Cells(r, DATA_CAT_COL_KATEGORIE).value)
        keyword = LCase(Trim(wsD.Cells(r, DATA_CAT_COL_KEYWORD).value))
        einAus = UCase(Trim(wsD.Cells(r, DATA_CAT_COL_EINAUS).value))
        prioritaet = Val(wsD.Cells(r, DATA_CAT_COL_PRIORITAET).value)
        
        If prioritaet = 0 Then prioritaet = 100
        
        If keyword <> "" And InStr(suchText, keyword) > 0 Then
            If (einAus = "E" And betrag > 0) Or (einAus = "A" And betrag < 0) Or einAus = "" Then
                If prioritaet < bestePrioritaet Then
                    bestePrioritaet = prioritaet
                    besteKategorie = kategorie
                End If
            End If
        End If
    Next r
    
    SucheKategorie = besteKategorie
    
End Function

Private Function ErmittleZielspalte(ByVal kategorie As String, ByVal betrag As Double) As Long
    
    Dim wsD As Worksheet
    Dim wsBK As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim kat As String
    Dim einAus As String
    Dim zielName As String
    Dim col As Long
    
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    ErmittleZielspalte = 0
    
    lastRow = wsD.Cells(wsD.Rows.Count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Function
    
    For r = DATA_START_ROW To lastRow
        kat = Trim(wsD.Cells(r, DATA_CAT_COL_KATEGORIE).value)
        einAus = UCase(Trim(wsD.Cells(r, DATA_CAT_COL_EINAUS).value))
        zielName = Trim(wsD.Cells(r, DATA_CAT_COL_ZIELSPALTE).value)
        
        If kat = kategorie Then
            If (einAus = "E" And betrag > 0) Or (einAus = "A" And betrag < 0) Then
                If zielName <> "" Then
                    For col = BK_COL_MITGL_BEITR To BK_COL_AUSZAHL_KASSE
                        If Trim(wsBK.Cells(BK_HEADER_ROW, col).value) = zielName Then
                            ErmittleZielspalte = col
                            Exit Function
                        End If
                    Next col
                End If
            End If
            Exit For
        End If
    Next r
    
End Function

' ===============================================================
' 7. HILFSFUNKTIONEN
' ===============================================================
Public Sub LoescheAlleBankkontoZeilen()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim antwort As VbMsgBoxResult
    
    antwort = MsgBox("ACHTUNG: Alle Daten auf dem Bankkonto-Blatt werden geloescht!" & vbCrLf & vbCrLf & _
                     "Fortfahren?", vbYesNo + vbCritical, "Alle Daten loeschen?")
    
    If antwort <> vbYes Then Exit Sub
    
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    
    If lastRow >= BK_START_ROW Then
        ws.Range(ws.Cells(BK_START_ROW, 1), ws.Cells(lastRow, 26)).ClearContents
        ws.Range(ws.Cells(BK_START_ROW, 1), ws.Cells(lastRow, 26)).Interior.ColorIndex = xlNone
    End If
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Call Initialize_ImportReport_ListBox
    
    MsgBox "Alle Daten wurden geloescht.", vbInformation
    
End Sub

Public Sub AktualisiereFormatierungBankkonto()
    
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    Call Anwende_Zebra_Bankkonto(ws)
    Call Anwende_Border_Bankkonto(ws)
    Call Anwende_Formatierung_Bankkonto(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    MsgBox "Formatierung aktualisiert!", vbInformation
    
End Sub

' ===============================================================
' 8. SORTIERE TABELLEN DATEN
' ===============================================================
Public Sub Sortiere_Tabellen_Daten()

    Dim ws As Worksheet
    Dim lr As Long
    
    Application.EnableEvents = False
    On Error GoTo ExitClean

    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ExitClean

    lr = ws.Cells(ws.Rows.Count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lr >= DATA_START_ROW Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add key:=ws.Cells(DATA_START_ROW, DATA_CAT_COL_KATEGORIE), _
                                 Order:=xlAscending
            .SetRange ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), ws.Cells(lr, DATA_CAT_COL_END))
            .Header = xlNo
            .Apply
        End With
    End If

    lr = ws.Cells(ws.Rows.Count, EK_COL_ENTITYKEY).End(xlUp).Row
    If lr >= EK_START_ROW Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add key:=ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                                 Order:=xlAscending
            .SetRange ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), ws.Cells(lr, EK_COL_DEBUG))
            .Header = xlNo
            .Apply
        End With
    End If
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True

ExitClean:
    Application.EnableEvents = True
End Sub


