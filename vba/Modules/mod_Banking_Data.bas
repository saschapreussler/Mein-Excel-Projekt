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
' 1. IBAN-BASIERTES ENTITY MAPPING
' ===============================================================
Public Sub Aktualisiere_Parzellen_Mapping_Final()

    Dim wsBK As Worksheet, wsD As Worksheet, wsM As Worksheet
    Dim dictIBANsBank As Object, dictIBANsMapping As Object
    Dim rD As Long, r As Long, lastRowD As Long, lastRowBK As Long
    Dim currentIBAN As Variant, currentKontoName As String, tempIBAN As String
    Dim foundZuordnung As String, foundParzellenRange As Range
    Dim ktonames As String, fuzzyResultCode As Long
    Dim entityID As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set dictIBANsBank = CreateObject("Scripting.Dictionary")
    Set dictIBANsMapping = CreateObject("Scripting.Dictionary")

    lastRowBK = wsBK.Cells(wsBK.Rows.Count, BK_COL_IBAN).End(xlUp).Row
    lastRowD = wsD.Cells(wsD.Rows.Count, DATA_MAP_COL_ENTITYKEY).End(xlUp).Row

    Dim COLOR_GREEN As Long, COLOR_YELLOW As Long, COLOR_RED As Long, COLOR_WHITE As Long
    COLOR_GREEN = RGB(198, 224, 180)
    COLOR_YELLOW = RGB(255, 230, 153)
    COLOR_RED = RGB(255, 150, 150)
    COLOR_WHITE = RGB(255, 255, 255)

    ' SCHRITT 1: Bestehende IBANs merken
    For rD = DATA_START_ROW To lastRowD
        tempIBAN = Replace(Trim(wsD.Cells(rD, DATA_MAP_COL_IBAN_OLD).value), " ", "")
        If tempIBAN <> "" Then dictIBANsMapping(tempIBAN) = True
    Next rD

    ' SCHRITT 2: IBANs aus Bankkonto aggregieren
    For r = BK_START_ROW To lastRowBK
        tempIBAN = Replace(Trim(wsBK.Cells(r, BK_COL_IBAN).value), " ", "")
        currentKontoName = Trim(wsBK.Cells(r, BK_COL_NAME).value)

        If tempIBAN <> "" And tempIBAN <> "n.a." Then
            If dictIBANsBank.Exists(tempIBAN) Then
                If InStr(1, dictIBANsBank(tempIBAN), currentKontoName, vbTextCompare) = 0 Then
                    dictIBANsBank(tempIBAN) = dictIBANsBank(tempIBAN) & vbLf & currentKontoName
                End If
            Else
                dictIBANsBank(tempIBAN) = currentKontoName
            End If
        End If
    Next r

    ' SCHRITT 3: Neue IBANs anhaengen
    entityID = 1
    If lastRowD >= DATA_START_ROW Then
        entityID = Application.Max(wsD.Columns(DATA_MAP_COL_ENTITYKEY)) + 1
    End If

    rD = IIf(lastRowD < DATA_START_ROW, DATA_START_ROW, lastRowD + 1)

    For Each currentIBAN In dictIBANsBank.Keys
        If Not dictIBANsMapping.Exists(currentIBAN) Then
            wsD.Cells(rD, DATA_MAP_COL_ENTITYKEY).value = entityID
            wsD.Cells(rD, DATA_MAP_COL_IBAN_OLD).value = currentIBAN
            wsD.Cells(rD, DATA_MAP_COL_KTONAME).value = dictIBANsBank(currentIBAN)
            entityID = entityID + 1
            rD = rD + 1
        End If
    Next currentIBAN

    ' SCHRITT 4: FUZZY-SUCHE
    lastRowD = wsD.Cells(wsD.Rows.Count, DATA_MAP_COL_ENTITYKEY).End(xlUp).Row

    For rD = DATA_START_ROW To lastRowD

        fuzzyResultCode = 0
        foundZuordnung = ""
        tempIBAN = Replace(Trim(wsD.Cells(rD, DATA_MAP_COL_IBAN_OLD).value), " ", "")

        If dictIBANsBank.Exists(tempIBAN) Then
            wsD.Cells(rD, DATA_MAP_COL_KTONAME).value = dictIBANsBank(tempIBAN)
        End If

        wsD.Range(wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG), _
                  wsD.Cells(rD, DATA_MAP_COL_DEBUG)).Interior.color = COLOR_WHITE

        If Trim(wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG).value) <> "" Then
            wsD.Cells(rD, DATA_MAP_COL_DEBUG).value = "Manuell zugeordnet oder bestaetigt"
            wsD.Range(wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG), wsD.Cells(rD, DATA_MAP_COL_DEBUG)).Interior.color = COLOR_GREEN
            GoTo NextRow
        End If

        ktonames = wsD.Cells(rD, DATA_MAP_COL_KTONAME).value
        Set foundParzellenRange = wsD.Cells(rD, DATA_MAP_COL_PARZELLE)
        foundZuordnung = FuzzyMemberSearch(ktonames, wsM, foundParzellenRange)

        If Trim(foundZuordnung) <> "" Then

            Dim normFound As String, normLine As String, ln As Variant
            Dim partsFound() As String, partsLine() As String
            Dim foundOK As Boolean

            normFound = LCase(Replace(foundZuordnung, ",", " "))
            normFound = Application.WorksheetFunction.Trim(normFound)
            partsFound = Split(normFound, " ")

            For Each ln In Split(ktonames, vbLf)
                normLine = LCase(Replace(ln, ",", " "))
                normLine = Application.WorksheetFunction.Trim(normLine)
                partsLine = Split(normLine, " ")

                If UBound(partsFound) = 1 And UBound(partsLine) = 1 Then
                    If (partsFound(0) = partsLine(0) And partsFound(1) = partsLine(1)) _
                    Or (partsFound(0) = partsLine(1) And partsFound(1) = partsLine(0)) Then
                        foundOK = True
                        Exit For
                    End If
                End If
            Next ln

            If foundOK Then
                fuzzyResultCode = 2
            Else
                fuzzyResultCode = 1
            End If
        End If

        Select Case fuzzyResultCode
            Case 2
                wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG).value = foundZuordnung
                wsD.Cells(rD, DATA_MAP_COL_ENTITYROLE).value = IIf(InStr(1, wsD.Cells(rD, DATA_MAP_COL_PARZELLE).value, "Verein", vbTextCompare) > 0, "VEREIN", "MITGLIED")
                wsD.Cells(rD, DATA_MAP_COL_DEBUG).value = "Sicherer Treffer"
                wsD.Range(wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG), wsD.Cells(rD, DATA_MAP_COL_DEBUG)).Interior.color = COLOR_GREEN

            Case 1
                wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG).value = foundZuordnung
                wsD.Cells(rD, DATA_MAP_COL_ENTITYROLE).value = IIf(InStr(1, wsD.Cells(rD, DATA_MAP_COL_PARZELLE).value, "Verein", vbTextCompare) > 0, "VEREIN", "MITGLIED")
                wsD.Cells(rD, DATA_MAP_COL_DEBUG).value = "Unsicherer Treffer - bitte pruefen"
                wsD.Range(wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG), wsD.Cells(rD, DATA_MAP_COL_DEBUG)).Interior.color = COLOR_YELLOW

            Case Else
                wsD.Cells(rD, DATA_MAP_COL_DEBUG).value = "Kein Treffer - manuelle Zuordnung erforderlich"
                wsD.Range(wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG), wsD.Cells(rD, DATA_MAP_COL_DEBUG)).Interior.color = COLOR_RED
        End Select

NextRow:
    Next rD

    Call ApplyMappingTableFormatting(wsD, lastRowD)

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub


Private Sub ApplyMappingTableFormatting(ByVal ws As Worksheet, ByVal lastDataRow As Long)

    Const MAX_DROPDOWN_ROW As Long = 504
    Const COLOR_WHITE As Long = 16777215

    If lastDataRow < DATA_START_ROW Then Exit Sub

    Dim rngTable As Range
    Dim ddRange As Range
    Dim dropdownEndRow As Long

    dropdownEndRow = Application.WorksheetFunction.Max(lastDataRow, MAX_DROPDOWN_ROW)

    Set rngTable = ws.Range( _
        ws.Cells(DATA_START_ROW, DATA_MAP_COL_ENTITYKEY), _
        ws.Cells(lastDataRow, DATA_MAP_COL_LAST) _
    )

    With rngTable
        .ClearOutline
        .WrapText = True
        .VerticalAlignment = xlCenter

        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End With

    ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_ENTITYKEY), _
             ws.Cells(lastDataRow, DATA_MAP_COL_ENTITYKEY)).HorizontalAlignment = xlCenter

    ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_PARZELLE), _
             ws.Cells(lastDataRow, DATA_MAP_COL_PARZELLE)).HorizontalAlignment = xlCenter

    ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_IBAN_OLD), _
             ws.Cells(lastDataRow, DATA_MAP_COL_KTONAME)).HorizontalAlignment = xlLeft

    ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_ZUORDNUNG), _
             ws.Cells(lastDataRow, DATA_MAP_COL_DEBUG)).HorizontalAlignment = xlLeft

    ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_ENTITYKEY), _
             ws.Cells(lastDataRow, DATA_MAP_COL_LAST)).EntireColumn.AutoFit

    ws.Rows(DATA_START_ROW & ":" & lastDataRow).AutoFit

    With ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_KTONAME), _
                  ws.Cells(lastDataRow, DATA_MAP_COL_KTONAME))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ws.Rows(DATA_START_ROW & ":" & lastDataRow).AutoFit

    ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_ENTITYKEY), _
             ws.Cells(lastDataRow, DATA_MAP_COL_KTONAME)).Interior.color = COLOR_WHITE

    Set ddRange = ws.Range( _
        ws.Cells(DATA_START_ROW, DATA_MAP_COL_ENTITYROLE), _
        ws.Cells(dropdownEndRow, DATA_MAP_COL_ENTITYROLE) _
    )

    On Error Resume Next
    ddRange.Validation.Delete
    On Error GoTo 0

    With ddRange.Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="=Daten!$AF$4:$AF$8"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Rolle waehlen"
        .ErrorTitle = "Ungueltige Rolle"
        .ErrorMessage = "Bitte waehlen Sie eine gueltige Rolle aus der Liste."
    End With

End Sub


' ===============================================================
' 2. CSV-KONTOAUSZUG IMPORT
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
    On Error GoTo 0
    
    strFile = Application.GetOpenFilename("CSV (*.csv), *.csv")
    If strFile = False Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
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
    
    On Error GoTo ImportFehler

    Set wsTemp = ThisWorkbook.Worksheets.Add(After:=wsZiel)
    wsTemp.Name = tempSheetName
    
    With wsTemp.QueryTables.Add(Connection:="TEXT;" & strFile, Destination:=wsTemp.Cells(1, 1))
        .Name = "CSV_Import"
        .FieldNames = True
        .TextFilePlatform = xlUTF8Value
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimitedValue
        .TextFileSemicolonDelimiter = True
        .Refresh BackgroundQuery:=False
    End With
    
    lastRowTemp = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
    rowsTotalInFile = lastRowTemp - 1
    
    If lastRowTemp <= 1 Then
        rowsProcessed = 0
        GoTo ImportEnde
    End If
    
    wsTemp.QueryTables(1).Delete
    
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
        
        sFormelAuswertungsmonat = "=IF(A" & lRowZiel & "="""","""",IF(Daten!$AG$4=0,TRUE,MONTH(A" & lRowZiel & ")=Daten!$AG$4))"
        wsZiel.Cells(lRowZiel, BK_COL_IM_AUSWERTUNGSMONAT).Formula = sFormelAuswertungsmonat
        
        wsZiel.Cells(lRowZiel, BK_COL_STATUS).value = "Gebucht"
        
        rowsProcessed = rowsProcessed + 1

NextRowImport:
    Next lRowTemp

ImportEnde:
    
    rowsFailedImport = rowsIgnoredFilter
    
    Call Update_ImportReport_ListBox(rowsTotalInFile, rowsProcessed, rowsIgnoredDupe, rowsFailedImport)
    
    If Not wsTemp Is Nothing Then
        wsTemp.Delete
        Set wsTemp = Nothing
    End If
    
    ' =====================================================
    ' WICHTIG: EntityKey-Mapping VOR Kategorisierung!
    ' =====================================================
    Call Aktualisiere_Parzellen_Mapping_Final
    
    Call Sortiere_Bankkonto_nach_Datum
    Call Anwende_Zebra_Bankkonto(wsZiel)
    Call Anwende_Border_Bankkonto(wsZiel)
    Call Anwende_Formatierung_Bankkonto(wsZiel)
    
    ' Kategorisierung NACH EntityKey-Update
    If rowsProcessed > 0 Then
        Call KategorieEngine_Pipeline(wsZiel)
    End If
    
    ' Monat/Periode automatisch setzen
    Call Setze_Monat_Periode(wsZiel)
    
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

ImportFehler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    If rowsTotalInFile = 0 Then
        rowsFailedImport = 1
    Else
        rowsFailedImport = rowsFailedImport + 1
    End If
    
    Call Update_ImportReport_ListBox(rowsTotalInFile, rowsProcessed, rowsIgnoredDupe, rowsFailedImport)

    MsgBox "FATALER FEHLER beim Importieren der CSV-Datei. Fehler: " & Err.Description, vbCritical
    
    On Error Resume Next
    If Not wsTemp Is Nothing Then wsTemp.Delete
    wsZiel.Activate
    On Error GoTo 0
End Sub


' ===============================================================
' 2b. ZEBRA-FORMATIERUNG
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
' 2c. BORDER-FORMATIERUNG
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


' ===============================================================
' 2d. FORMATIERUNG (Datum, Waehrung, DropDowns)
' ===============================================================
Private Sub Anwende_Formatierung_Bankkonto(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim rngKategorieE As Range
    Dim rngKategorieA As Range
    Dim rngMonatPeriode As Range
    Dim lRow As Long
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ' --- Spalte A: Datum-Format ---
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_DATUM), _
             ws.Cells(lastRow, BK_COL_DATUM)).NumberFormat = "DD.MM.YYYY"
    
    ' --- Spalte B: Waehrung Euro ---
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_BETRAG), _
             ws.Cells(lastRow, BK_COL_BETRAG)).NumberFormat = "#,##0.00 [$EUR]"
    
    ' --- Spalte J: Zentriert ---
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_INTERNE_NR), _
             ws.Cells(lastRow, BK_COL_INTERNE_NR)).HorizontalAlignment = xlCenter
    
    ' --- Spalten M-Z: Waehrung Euro ---
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_MITGL_BEITR), _
             ws.Cells(lastRow, BK_COL_AUSZAHL_KASSE)).NumberFormat = "#,##0.00 [$EUR]"
    
    ' --- Spalte H: DropDown fuer Kategorie (abhaengig von E/A) ---
    For lRow = BK_START_ROW To lastRow
        Dim betrag As Double
        betrag = ws.Cells(lRow, BK_COL_BETRAG).value
        
        On Error Resume Next
        ws.Cells(lRow, BK_COL_KATEGORIE).Validation.Delete
        On Error GoTo 0
        
        If betrag > 0 Then
            ' Einnahmen-Kategorien
            With ws.Cells(lRow, BK_COL_KATEGORIE).Validation
                .Add Type:=xlValidateList, _
                     AlertStyle:=xlValidAlertWarning, _
                     Formula1:="=lst_KategorienEinnahmen"
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
        ElseIf betrag < 0 Then
            ' Ausgaben-Kategorien
            With ws.Cells(lRow, BK_COL_KATEGORIE).Validation
                .Add Type:=xlValidateList, _
                     AlertStyle:=xlValidAlertWarning, _
                     Formula1:="=lst_KategorienAusgaben"
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
        End If
    Next lRow
    
    ' --- Spalte I: DropDown fuer Monat/Periode ---
    Set rngMonatPeriode = ws.Range(ws.Cells(BK_START_ROW, BK_COL_MONAT_PERIODE), _
                                    ws.Cells(lastRow, BK_COL_MONAT_PERIODE))
    
    On Error Resume Next
    rngMonatPeriode.Validation.Delete
    On Error GoTo 0
    
    With rngMonatPeriode.Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:="=lst_MonatPeriode"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    Application.ScreenUpdating = True
    
End Sub


' ===============================================================
' 2e. MONAT/PERIODE AUTOMATISCH SETZEN
' ===============================================================
Private Sub Setze_Monat_Periode(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim lRow As Long
    Dim buchungsDatum As Date
    Dim buchungsMonat As Long
    Dim periodeText As String
    
    ' Monatsbezeichnungen
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
        ' Nur wenn Monat/Periode noch leer ist
        If Trim(ws.Cells(lRow, BK_COL_MONAT_PERIODE).value) = "" Then
            
            If IsDate(ws.Cells(lRow, BK_COL_DATUM).value) Then
                buchungsDatum = ws.Cells(lRow, BK_COL_DATUM).value
                buchungsMonat = Month(buchungsDatum)
                
                ' Standard: Buchungsmonat verwenden
                periodeText = monate(buchungsMonat)
                
                ' OPTIONAL: Hier kann spaeter eine komplexere Logik eingefuegt werden
                ' z.B. basierend auf Einstellungen!-Blatt oder Kategorie-Faelligkeit
                
                ws.Cells(lRow, BK_COL_MONAT_PERIODE).value = periodeText
            End If
        End If
    Next lRow
    
End Sub


' ===============================================================
' 3. SORTIERUNG
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
' 4. PROTOKOLLIERUNG (ListBox)
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
' 5. KATEGORISIERUNG (ZENTRALE STEUERUNG)
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
    
    ' WICHTIG: Erst EntityKey-Mapping aktualisieren!
    Call Aktualisiere_Parzellen_Mapping_Final
    
    ' Dann Kategorisierung
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




