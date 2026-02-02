Attribute VB_Name = "mod_Banking_Data"
Option Explicit

' ===============================================================
' MODUL: mod_Banking_Data (FINAL KONSOLIDIERT)
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
    
    Dim debugStep As String
    
    tempSheetName = "TempImport"
    
    Set wsZiel = ThisWorkbook.Worksheets(WS_BANKKONTO)
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
    
    strFile = Application.GetOpenFilename("CSV (*.csv), *.csv")
    If strFile = False Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Call Initialize_ImportReport_ListBox
        Exit Sub
    End If
    
    debugStep = "Schritt 1: Bestehende Umsaetze lesen"
    On Error GoTo ImportFehler
    
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
    
    debugStep = "Schritt 2: CSV-Datei einlesen"
    
    On Error Resume Next
    Set wsTemp = ThisWorkbook.Worksheets.Add(After:=wsZiel)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ImportFehler
        GoTo ImportFehler
    End If
    wsTemp.Name = tempSheetName
    Err.Clear
    On Error GoTo ImportFehler
    
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
        debugStep = "Schritt 2: CSV-Datei einlesen (QueryTable Fehler)"
        Dim errDesc As String
        Dim errNum As Long
        errDesc = Err.Description
        errNum = Err.Number
        Err.Clear
        On Error GoTo 0
        
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        
        If Not wsTemp Is Nothing Then
            Application.DisplayAlerts = False
            wsTemp.Delete
            Application.DisplayAlerts = True
        End If
        
        MsgBox "FEHLER beim Einlesen der CSV-Datei." & vbCrLf & vbCrLf & _
               "Schritt: " & debugStep & vbCrLf & _
               "Fehler: " & errDesc & vbCrLf & _
               "Fehler-Nr: " & errNum, vbCritical
        Exit Sub
    End If
    Err.Clear
    On Error GoTo ImportFehler
    
    lastRowTemp = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
    rowsTotalInFile = lastRowTemp - 1
    
    If lastRowTemp <= 1 Then
        rowsProcessed = 0
        GoTo ImportEnde
    End If
    
    debugStep = "Schritt 3: CSV-Zeilen verarbeiten"
    
    On Error Resume Next
    wsTemp.QueryTables(1).Delete
    Err.Clear
    On Error GoTo ImportFehler
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
            On Error GoTo ImportFehler
            GoTo NextRowImport
        End If
        On Error GoTo ImportFehler
        
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

ImportEnde:
    
    rowsFailedImport = rowsIgnoredFilter
    
    debugStep = "Schritt 4: Update_ImportReport_ListBox"
    Call Update_ImportReport_ListBox(rowsTotalInFile, rowsProcessed, rowsIgnoredDupe, rowsFailedImport)
    
    debugStep = "Schritt 5: TempSheet loeschen"
    If Not wsTemp Is Nothing Then
        On Error Resume Next
        Application.DisplayAlerts = False
        wsTemp.Delete
        Application.DisplayAlerts = True
        Err.Clear
        On Error GoTo ImportFehlerNachImport
        Set wsTemp = Nothing
    End If
    
    debugStep = "Schritt 6: ImportiereIBANsAusBankkonto"
    On Error GoTo ImportFehlerNachImport
    Call ImportiereIBANsAusBankkonto
    
    debugStep = "Schritt 7: Sortiere_Bankkonto_nach_Datum"
    Call Sortiere_Bankkonto_nach_Datum
    
    debugStep = "Schritt 8: Anwende_Zebra_Bankkonto"
    Call Anwende_Zebra_Bankkonto(wsZiel)
    
    debugStep = "Schritt 9: Anwende_Border_Bankkonto"
    Call Anwende_Border_Bankkonto(wsZiel)
    
    debugStep = "Schritt 10: Anwende_Formatierung_Bankkonto"
    Call Anwende_Formatierung_Bankkonto(wsZiel)
    
    debugStep = "Schritt 11: KategorieEngine_Pipeline"
    If rowsProcessed > 0 Then
        Call KategorieEngine_Pipeline(wsZiel)
    End If
    
    debugStep = "Schritt 12: Setze_Monat_Periode"
    Call Setze_Monat_Periode(wsZiel)
    
    debugStep = "Schritt 13: Fertig"
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
    
    Exit Sub

ImportFehlerNachImport:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "FEHLER nach CSV-Import bei: " & debugStep & vbCrLf & vbCrLf & _
           "Fehler: " & Err.Description & vbCrLf & _
           "Fehler-Nr: " & Err.Number, vbCritical, "Fehler nach Import"
    
    wsZiel.Activate
    Exit Sub

ImportFehler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    If rowsTotalInFile = 0 Then
        rowsFailedImport = 1
    Else
        rowsFailedImport = rowsFailedImport + 1
    End If
    
    Call Update_ImportReport_ListBox(rowsTotalInFile, rowsProcessed, rowsIgnoredDupe, rowsFailedImport)

    MsgBox "FATALER FEHLER beim Importieren der CSV-Datei." & vbCrLf & vbCrLf & _
           "Schritt: " & debugStep & vbCrLf & _
           "Fehler: " & Err.Description & vbCrLf & _
           "Fehler-Nr: " & Err.Number, vbCritical
    
    On Error Resume Next
    If Not wsTemp Is Nothing Then
        Application.DisplayAlerts = False
        wsTemp.Delete
        Application.DisplayAlerts = True
    End If
    wsZiel.Activate
    On Error GoTo 0
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
    
    On Error Resume Next
    ws.Range(ws.Cells(BK_START_ROW, 1), ws.Cells(lastRow, 7)).Interior.ColorIndex = xlNone
    ws.Range(ws.Cells(BK_START_ROW, 9), ws.Cells(lastRow, 26)).Interior.ColorIndex = xlNone
    On Error GoTo 0
    
    For lRow = BK_START_ROW To lastRow
        If ws.Cells(lRow, BK_COL_DATUM).value <> "" Then
            If (lRow - BK_START_ROW) Mod 2 = 1 Then
                Set rngRowPart1 = ws.Range(ws.Cells(lRow, 1), ws.Cells(lRow, 7))
                rngRowPart1.Interior.color = ZEBRA_COLOR
                
                Set rngRowPart2 = ws.Range(ws.Cells(lRow, 9), ws.Cells(lRow, 26))
                rngRowPart2.Interior.color = ZEBRA_COLOR
            End If
        End If
    Next lRow
    
End Sub


' ===============================================================
' 1c. BORDER-FORMATIERUNG
' ===============================================================
Private Sub Anwende_Border_Bankkonto(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim rngTable As Range
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    Set rngTable = ws.Range(ws.Cells(BK_START_ROW, 1), ws.Cells(lastRow, 26))
    
    On Error Resume Next
    
    With rngTable.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With rngTable.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With rngTable.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With rngTable.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With rngTable.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With rngTable.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    On Error GoTo 0
    
End Sub

'--- Ende Teil 1 ---
'--- Anfang Teil 2 ---

' ===============================================================
' 1d. FORMATIERUNG (Datum, Waehrung, DropDowns)
' ===============================================================
Private Sub Anwende_Formatierung_Bankkonto(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim rngMonatPeriode As Range
    Dim lRow As Long
    Dim euroFormat As String
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    euroFormat = "#,##0.00 " & ChrW(8364)
    
    Application.ScreenUpdating = False
    
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_DATUM), _
             ws.Cells(lastRow, BK_COL_DATUM)).NumberFormat = "DD.MM.YYYY"
    
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_BETRAG), _
             ws.Cells(lastRow, BK_COL_BETRAG)).NumberFormat = euroFormat
    
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_INTERNE_NR), _
             ws.Cells(lastRow, BK_COL_INTERNE_NR)).HorizontalAlignment = xlCenter
    
    With ws.Range(ws.Cells(BK_START_ROW, BK_COL_BEMERKUNG), _
                  ws.Cells(lastRow, BK_COL_BEMERKUNG))
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_MITGL_BEITR), _
             ws.Cells(lastRow, BK_COL_AUSZAHL_KASSE)).NumberFormat = euroFormat
    
    ws.Rows(BK_START_ROW & ":" & lastRow).AutoFit
    
    Dim hasEinnahmenList As Boolean
    Dim hasAusgabenList As Boolean
    hasEinnahmenList = NamedRangeExistsLocal("lst_KategorienEinnahmen")
    hasAusgabenList = NamedRangeExistsLocal("lst_KategorienAusgaben")
    
    If hasEinnahmenList And hasAusgabenList Then
        For lRow = BK_START_ROW To lastRow
            Dim betrag As Double
            betrag = ws.Cells(lRow, BK_COL_BETRAG).value
            
            On Error Resume Next
            ws.Cells(lRow, BK_COL_KATEGORIE).Validation.Delete
            On Error GoTo 0
            
            If betrag > 0 Then
                On Error Resume Next
                With ws.Cells(lRow, BK_COL_KATEGORIE).Validation
                    .Add Type:=xlValidateList, _
                         AlertStyle:=xlValidAlertWarning, _
                         Formula1:="=lst_KategorienEinnahmen"
                    .IgnoreBlank = True
                    .InCellDropdown = True
                End With
                On Error GoTo 0
            ElseIf betrag < 0 Then
                On Error Resume Next
                With ws.Cells(lRow, BK_COL_KATEGORIE).Validation
                    .Add Type:=xlValidateList, _
                         AlertStyle:=xlValidAlertWarning, _
                         Formula1:="=lst_KategorienAusgaben"
                    .IgnoreBlank = True
                    .InCellDropdown = True
                End With
                On Error GoTo 0
            End If
        Next lRow
    End If
    
    If NamedRangeExistsLocal("lst_MonatPeriode") Then
        Set rngMonatPeriode = ws.Range(ws.Cells(BK_START_ROW, BK_COL_MONAT_PERIODE), _
                                        ws.Cells(lastRow, BK_COL_MONAT_PERIODE))
        
        On Error Resume Next
        rngMonatPeriode.Validation.Delete
        On Error GoTo 0
        
        On Error Resume Next
        With rngMonatPeriode.Validation
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertWarning, _
                 Formula1:="=lst_MonatPeriode"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        On Error GoTo 0
    End If
    
    Application.ScreenUpdating = True
    
End Sub


Private Function NamedRangeExistsLocal(ByVal rangeName As String) As Boolean
    Dim nm As Name
    NamedRangeExistsLocal = False
    
    On Error Resume Next
    Set nm = ThisWorkbook.Names(rangeName)
    If Not nm Is Nothing Then
        NamedRangeExistsLocal = True
    End If
    On Error GoTo 0
End Function


' ===============================================================
' 1e. MONAT/PERIODE AUTOMATISCH SETZEN
' ===============================================================
Private Sub Setze_Monat_Periode(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim lRow As Long
    Dim buchungsDatum As Date
    Dim buchungsMonat As Long
    Dim periodeText As String
    
    Dim monate(1 To 12) As String
    monate(1) = "Januar"
    monate(2) = "Februar"
    monate(3) = "Maerz"
    monate(4) = "April"
    monate(5) = "Mai"
    monate(6) = "Juni"
    monate(7) = "Juli"
    monate(8) = "August"
    monate(9) = "September"
    monate(10) = "Oktober"
    monate(11) = "November"
    monate(12) = "Dezember"
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    For lRow = BK_START_ROW To lastRow
        If Trim(ws.Cells(lRow, BK_COL_MONAT_PERIODE).value) = "" Then
            If IsDate(ws.Cells(lRow, BK_COL_DATUM).value) Then
                buchungsDatum = ws.Cells(lRow, BK_COL_DATUM).value
                buchungsMonat = Month(buchungsDatum)
                periodeText = monate(buchungsMonat)
                ws.Cells(lRow, BK_COL_MONAT_PERIODE).value = periodeText
            End If
        End If
    Next lRow
    
End Sub


' ===============================================================
' 2. SORTIERUNG
' ===============================================================
Public Sub Sortiere_Bankkonto_nach_Datum()
    On Error GoTo SortError

    Dim ws As Worksheet
    Dim lr As Long
    
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    lr = ws.Cells(ws.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    
    If lr < BK_START_ROW Or IsEmpty(ws.Cells(BK_START_ROW, BK_COL_DATUM).value) Then
        Exit Sub
    End If

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Cells(BK_START_ROW, BK_COL_DATUM), _
                             SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(BK_START_ROW, 1), ws.Cells(lr, BK_COL_AUSZAHL_KASSE))
        .Header = xlNo
        .Apply
    End With
    
    Exit Sub

SortError:
    MsgBox "Sortierung konnte nicht durchgefuehrt werden: " & Err.Description, vbCritical
    
End Sub


Public Sub Sortiere_Tabellen_Daten()

    Dim ws As Worksheet
    Dim lr As Long
    
    Application.EnableEvents = False
    On Error GoTo ExitClean

    Set ws = ThisWorkbook.Worksheets(WS_DATEN)

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

    lr = ws.Cells(ws.Rows.Count, DATA_MAP_COL_ENTITYKEY).End(xlUp).Row
    If lr >= DATA_START_ROW Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add key:=ws.Cells(DATA_START_ROW, DATA_MAP_COL_ENTITYKEY), _
                                 Order:=xlAscending
            .SetRange ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_ENTITYKEY), ws.Cells(lr, DATA_MAP_COL_LAST))
            .Header = xlNo
            .Apply
        End With
    End If

ExitClean:
    Application.EnableEvents = True
End Sub


' ===============================================================
' 3. PROTOKOLLIERUNG (ListBox)
' ===============================================================
Private Function Get_Protocol_Temp_Sheet() As Worksheet
    
    On Error Resume Next
    Set Get_Protocol_Temp_Sheet = ThisWorkbook.Worksheets(WS_PROTOCOL_TEMP)
    On Error GoTo 0
    
    If Get_Protocol_Temp_Sheet Is Nothing Then
        Set Get_Protocol_Temp_Sheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Get_Protocol_Temp_Sheet.Name = WS_PROTOCOL_TEMP
        Get_Protocol_Temp_Sheet.Visible = xlSheetVeryHidden
    End If
    
    Get_Protocol_Temp_Sheet.Columns(1).NumberFormat = "@"
    
End Function


Private Sub SetzeListBoxHintergrundfarbe(ByVal wsZiel As Worksheet, ByVal farbe As Long)
    
    On Error Resume Next
    
    Dim oleObj As OLEObject
    Set oleObj = wsZiel.OLEObjects(FORM_LISTBOX_NAME)
    If Not oleObj Is Nothing Then
        oleObj.Object.BackColor = farbe
        If Err.Number = 0 Then Exit Sub
        Err.Clear
    End If
    
    Dim shp As Shape
    Set shp = wsZiel.Shapes(FORM_LISTBOX_NAME)
    If Not shp Is Nothing Then
        shp.DrawingObject.Interior.color = farbe
        If Err.Number = 0 Then Exit Sub
        Err.Clear
    End If
    
    Call SetzeListBoxRahmenFarbe(wsZiel, farbe)
    
    On Error GoTo 0
End Sub


Private Sub SetzeListBoxRahmenFarbe(ByVal wsZiel As Worksheet, ByVal farbe As Long)
    
    On Error Resume Next
    Dim shpRahmen As Shape
    Set shpRahmen = wsZiel.Shapes(RAHMEN_NAME)
    If Not shpRahmen Is Nothing Then
        shpRahmen.Fill.ForeColor.RGB = farbe
    End If
    On Error GoTo 0
    
End Sub


Private Function ErmittleAmpelFarbe(ByVal duplicates As Long, ByVal errors As Long) As Long
    If errors > 0 Then
        ErmittleAmpelFarbe = AMPEL_ROT
    ElseIf duplicates > 0 Then
        ErmittleAmpelFarbe = AMPEL_GELB
    Else
        ErmittleAmpelFarbe = AMPEL_GRUEN
    End If
End Function


Private Function ExtrahiereZahl(ByVal text As String) As Long
    Dim i As Long
    Dim numStr As String
    
    numStr = ""
    For i = 1 To Len(text)
        If Mid(text, i, 1) >= "0" And Mid(text, i, 1) <= "9" Then
            numStr = numStr & Mid(text, i, 1)
        End If
    Next i
    
    If numStr <> "" Then
        ExtrahiereZahl = CLng(numStr)
    Else
        ExtrahiereZahl = 0
    End If
End Function

'--- Ende Teil 2 ---
'--- Anfang Teil 3 ---

Public Sub Initialize_ImportReport_ListBox()
    
    Dim wsZiel As Worksheet
    Dim wsDaten As Worksheet
    Dim wsTemp As Worksheet
    Dim protocolRange As String
    Dim k As Long
    
    Const HISTORY_DELIMITER As String = "|REPORT_DELIMITER|"
    Const PART_DELIMITER As String = "|PART|"
    
    Set wsZiel = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsTemp = Get_Protocol_Temp_Sheet()

    Application.ScreenUpdating = False
    
    wsTemp.Cells.ClearContents
    
    If CStr(wsDaten.Range(CELL_IMPORT_PROTOKOLL).value) <> "" Then
        Dim historyString As String
        Dim reports() As String
        Dim reportParts() As String
        Dim i As Long
        Dim lastDuplicates As Long
        Dim lastErrors As Long
        
        historyString = CStr(wsDaten.Range(CELL_IMPORT_PROTOKOLL).value)
        reports = Split(historyString, HISTORY_DELIMITER)
        
        k = 1
        
        For i = 0 To UBound(reports)
            reportParts = Split(reports(i), PART_DELIMITER)
            
            If i = 0 Then
                If UBound(reportParts) >= 2 Then lastDuplicates = ExtrahiereZahl(reportParts(2))
                If UBound(reportParts) >= 3 Then lastErrors = ExtrahiereZahl(reportParts(3))
            End If
            
            If UBound(reportParts) >= 0 Then
                wsTemp.Cells(k, 1).value = Trim(reportParts(0))
                k = k + 1
            End If
            If UBound(reportParts) >= 1 Then
                wsTemp.Cells(k, 1).value = "  " & Trim(reportParts(1))
                k = k + 1
            End If
            If UBound(reportParts) >= 2 Then
                wsTemp.Cells(k, 1).value = "  " & Trim(reportParts(2))
                k = k + 1
            End If
            If UBound(reportParts) >= 3 Then
                wsTemp.Cells(k, 1).value = "  " & Trim(reportParts(3))
                k = k + 1
            End If
            
            wsTemp.Cells(k, 1).value = "--------------------------------"
            k = k + 1
            
            If k >= MAX_LISTBOX_LINES Then Exit For
        Next i
        
        Call SetzeListBoxHintergrundfarbe(wsZiel, ErmittleAmpelFarbe(lastDuplicates, lastErrors))
        
    Else
        wsTemp.Range(PROTOCOL_RANGE_START).value = "--------------------------------"
        wsTemp.Range(PROTOCOL_RANGE_START).Offset(1, 0).value = " Kein Import-Bericht verfuegbar."
        wsTemp.Range(PROTOCOL_RANGE_START).Offset(2, 0).value = " Fuehren Sie einen CSV-Import"
        wsTemp.Range(PROTOCOL_RANGE_START).Offset(3, 0).value = " durch, um den Bericht hier"
        wsTemp.Range(PROTOCOL_RANGE_START).Offset(4, 0).value = " anzuzeigen."
        wsTemp.Range(PROTOCOL_RANGE_START).Offset(5, 0).value = "--------------------------------"
        k = 7
        
        Call SetzeListBoxHintergrundfarbe(wsZiel, AMPEL_WEISS)
    End If
    
    On Error Resume Next
    If k > 1 Then
        protocolRange = wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(k - 1, 1)).Address(External:=False)
    Else
        protocolRange = wsTemp.Range("A1:A6").Address(External:=False)
    End If
    wsZiel.Shapes(FORM_LISTBOX_NAME).ControlFormat.ListFillRange = "'" & WS_PROTOCOL_TEMP & "'!" & protocolRange
    On Error GoTo 0
    
    Application.ScreenUpdating = True
End Sub


Public Sub Update_ImportReport_ListBox(ByVal totalEntries As Long, ByVal importedEntries As Long, ByVal duplicateEntries As Long, ByVal errorEntries As Long)

    Dim wsZiel As Worksheet
    Dim wsDaten As Worksheet
    Dim wsTemp As Worksheet
    Dim protocolRange As String
    
    Dim strDateTime As String
    Dim currentHistory() As String
    Dim historyString As String
    Dim newHistoryString As String
    Dim i As Long, k As Long
    
    Const HISTORY_DELIMITER As String = "|REPORT_DELIMITER|"
    Const PART_DELIMITER As String = "|PART|"
    
    Set wsZiel = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsTemp = Get_Protocol_Temp_Sheet()
    
    strDateTime = Format(Now, "dd.mm.yyyy hh:nn:ss")
    
    Application.ScreenUpdating = False
    
    Dim part1 As String: part1 = strDateTime
    Dim part2 As String: part2 = importedEntries & " / " & totalEntries & " Datensaetze importiert"
    Dim part3 As String: part3 = "Duplikate: " & duplicateEntries
    Dim part4 As String: part4 = "Fehler: " & errorEntries
    
    Dim newReportEntry As String
    newReportEntry = part1 & PART_DELIMITER & part2 & PART_DELIMITER & part3 & PART_DELIMITER & part4
    
    historyString = CStr(wsDaten.Range(CELL_IMPORT_PROTOKOLL).value)
    newHistoryString = newReportEntry & IIf(historyString <> "", HISTORY_DELIMITER & historyString, "")
    
    With wsDaten.Range(CELL_IMPORT_PROTOKOLL)
        .value = newHistoryString
        .WrapText = True
    End With

    wsTemp.Cells.ClearContents
    k = 1
    
    currentHistory = Split(newHistoryString, HISTORY_DELIMITER)
    
    For i = 0 To UBound(currentHistory)
        
        Dim reportParts() As String
        reportParts = Split(currentHistory(i), PART_DELIMITER)
        
        If UBound(reportParts) >= 0 Then
            wsTemp.Cells(k, 1).value = Trim(reportParts(0))
            k = k + 1
        End If
        If UBound(reportParts) >= 1 Then
            wsTemp.Cells(k, 1).value = "  " & Trim(reportParts(1))
            k = k + 1
        End If
        If UBound(reportParts) >= 2 Then
            wsTemp.Cells(k, 1).value = "  " & Trim(reportParts(2))
            k = k + 1
        End If
        If UBound(reportParts) >= 3 Then
            wsTemp.Cells(k, 1).value = "  " & Trim(reportParts(3))
            k = k + 1
        End If

        wsTemp.Cells(k, 1).value = "--------------------------------"
        k = k + 1
        
        If k >= MAX_LISTBOX_LINES Then Exit For
    Next i
    
    On Error Resume Next
    If Not wsZiel.Shapes(FORM_LISTBOX_NAME) Is Nothing Then
        protocolRange = wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(k - 1, 1)).Address(External:=False)
        wsZiel.Shapes(FORM_LISTBOX_NAME).ControlFormat.ListFillRange = "'" & WS_PROTOCOL_TEMP & "'!" & protocolRange
    End If
    On Error GoTo 0
    
    Call SetzeListBoxHintergrundfarbe(wsZiel, ErmittleAmpelFarbe(duplicateEntries, errorEntries))
    
    Application.ScreenUpdating = True
End Sub


' ===============================================================
' 4. KATEGORISIERUNG (ZENTRALE STEUERUNG)
' ===============================================================
Public Sub Kategorisiere_Umsaetze()
    
    Dim wsBK As Worksheet
    Dim lngLastRow As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo CategorizationError

    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    lngLastRow = wsBK.Cells(wsBK.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    
    If lngLastRow < BK_START_ROW Then
        MsgBox "Keine Banktransaktionen zum Kategorisieren gefunden.", vbInformation
        GoTo ExitClean
    End If
    
    Call ImportiereIBANsAusBankkonto
    
    Call KategorieEngine_Pipeline(wsBK)

    Call Sortiere_Bankkonto_nach_Datum

    MsgBox "Die Kategorisierung der Banktransaktionen wurde abgeschlossen.", vbInformation

ExitClean:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
CategorizationError:
    MsgBox "Ein Fehler ist bei der Kategorisierung aufgetreten: " & Err.Description, vbCritical
    Resume ExitClean
End Sub


