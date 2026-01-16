Attribute VB_Name = "mod_Banking_Data"
Option Explicit
' ***************************************************************
' MODUL: mod_Banking_Data (FINAL KONSOLIDIERT & KORRIGIERT)
' KONSOLIDIERUNG: Mapping, Import, Sortierung und Protokollierung
' ***************************************************************

' ===============================================================
' 1. IBAN-BASIERTES ENTITY MAPPING (mit sauberer Fuzzy-Logik)
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

    ' Farben (NICHT Const!)
    Dim COLOR_GREEN As Long, COLOR_YELLOW As Long, COLOR_RED As Long, COLOR_WHITE As Long
    COLOR_GREEN = RGB(198, 224, 180)
    COLOR_YELLOW = RGB(255, 230, 153)
    COLOR_RED = RGB(255, 150, 150)
    COLOR_WHITE = RGB(255, 255, 255)

    ' ------------------------------------------------
    ' SCHRITT 1: Bestehende IBANs merken
    ' ------------------------------------------------
    For rD = DATA_START_ROW To lastRowD
        tempIBAN = Replace(Trim(wsD.Cells(rD, DATA_MAP_COL_IBAN_OLD).Value), " ", "")
        If tempIBAN <> "" Then dictIBANsMapping(tempIBAN) = True
    Next rD

    ' ------------------------------------------------
    ' SCHRITT 2: IBANs aus Bankkonto aggregieren
    ' ------------------------------------------------
    For r = BK_START_ROW To lastRowBK
        tempIBAN = Replace(Trim(wsBK.Cells(r, BK_COL_IBAN).Value), " ", "")
        currentKontoName = Trim(wsBK.Cells(r, BK_COL_NAME).Value)

        If tempIBAN <> "" And tempIBAN <> "n.a." Then
            If dictIBANsBank.exists(tempIBAN) Then
                If InStr(1, dictIBANsBank(tempIBAN), currentKontoName, vbTextCompare) = 0 Then
                    dictIBANsBank(tempIBAN) = dictIBANsBank(tempIBAN) & vbLf & currentKontoName
                End If
            Else
                dictIBANsBank(tempIBAN) = currentKontoName
            End If
        End If
    Next r

    ' ------------------------------------------------
    ' SCHRITT 3: Neue IBANs anhängen
    ' ------------------------------------------------
    entityID = 1
    If lastRowD >= DATA_START_ROW Then
        entityID = Application.Max(wsD.Columns(DATA_MAP_COL_ENTITYKEY)) + 1
    End If

    rD = IIf(lastRowD < DATA_START_ROW, DATA_START_ROW, lastRowD + 1)

    For Each currentIBAN In dictIBANsBank.Keys
        If Not dictIBANsMapping.exists(currentIBAN) Then
            wsD.Cells(rD, DATA_MAP_COL_ENTITYKEY).Value = entityID
            wsD.Cells(rD, DATA_MAP_COL_IBAN_OLD).Value = currentIBAN
            wsD.Cells(rD, DATA_MAP_COL_KTONAME).Value = dictIBANsBank(currentIBAN)
            entityID = entityID + 1
            rD = rD + 1
        End If
    Next currentIBAN

    ' ------------------------------------------------
    ' SCHRITT 4: FUZZY-SUCHE + SAUBERE AUSWERTUNG
    ' ------------------------------------------------
    lastRowD = wsD.Cells(wsD.Rows.Count, DATA_MAP_COL_ENTITYKEY).End(xlUp).Row

    For rD = DATA_START_ROW To lastRowD

        fuzzyResultCode = 0
        foundZuordnung = ""
        tempIBAN = Replace(Trim(wsD.Cells(rD, DATA_MAP_COL_IBAN_OLD).Value), " ", "")

        If dictIBANsBank.exists(tempIBAN) Then
            wsD.Cells(rD, DATA_MAP_COL_KTONAME).Value = dictIBANsBank(tempIBAN)
        End If

        wsD.Range(wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG), _
                  wsD.Cells(rD, DATA_MAP_COL_DEBUG)).Interior.color = COLOR_WHITE

        ' --- MANUELL HAT VORRANG ---
        If Trim(wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG).Value) <> "" Then
            wsD.Cells(rD, DATA_MAP_COL_DEBUG).Value = "Manuell zugeordnet oder bestätigt"
            wsD.Range(wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG), wsD.Cells(rD, DATA_MAP_COL_DEBUG)).Interior.color = COLOR_GREEN
            GoTo NextRow
        End If

        ktonames = wsD.Cells(rD, DATA_MAP_COL_KTONAME).Value
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
                wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG).Value = foundZuordnung
                wsD.Cells(rD, DATA_MAP_COL_ENTITYROLE).Value = IIf(InStr(1, wsD.Cells(rD, DATA_MAP_COL_PARZELLE).Value, "Verein", vbTextCompare) > 0, "VEREIN", "MITGLIED")
                wsD.Cells(rD, DATA_MAP_COL_DEBUG).Value = "Sicherer Treffer (Vor- und Nachname eindeutig)"
                wsD.Range(wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG), wsD.Cells(rD, DATA_MAP_COL_DEBUG)).Interior.color = COLOR_GREEN

            Case 1
                wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG).Value = foundZuordnung
                wsD.Cells(rD, DATA_MAP_COL_ENTITYROLE).Value = IIf(InStr(1, wsD.Cells(rD, DATA_MAP_COL_PARZELLE).Value, "Verein", vbTextCompare) > 0, "VEREIN", "MITGLIED")
                wsD.Cells(rD, DATA_MAP_COL_DEBUG).Value = "Unsicherer Treffer (keine eindeutige Namensgleichheit)"
                wsD.Range(wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG), wsD.Cells(rD, DATA_MAP_COL_DEBUG)).Interior.color = COLOR_YELLOW

            Case Else
                wsD.Cells(rD, DATA_MAP_COL_DEBUG).Value = "Kein Treffer – manuelle Zuordnung erforderlich"
                wsD.Range(wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG), wsD.Cells(rD, DATA_MAP_COL_DEBUG)).Interior.color = COLOR_RED
        End Select

NextRow:
    Next rD

    Call ApplyMappingTableFormatting(wsD, lastRowD)

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub



' ***************************************************************
' HILFSPROZEDUR: Formatiert die Entity Key Tabelle (FINAL KORRIGIERT)
' ***************************************************************
Private Sub ApplyMappingTableFormatting(ByVal ws As Worksheet, ByVal lastDataRow As Long)

    Const MAX_DROPDOWN_ROW As Long = 504
    Const COLOR_WHITE As Long = 16777215   ' RGB(255,255,255)

    If lastDataRow < DATA_START_ROW Then Exit Sub

    Dim rngTable As Range
    Dim ddRange As Range
    Dim dropdownEndRow As Long

    ' Dropdown soll mindestens bis Zeile 504 gehen
    dropdownEndRow = Application.WorksheetFunction.Max(lastDataRow, MAX_DROPDOWN_ROW)

    ' Gesamter Tabellenbereich (nur EntityKey-Tabelle)
    Set rngTable = ws.Range( _
        ws.Cells(DATA_START_ROW, DATA_MAP_COL_ENTITYKEY), _
        ws.Cells(lastDataRow, DATA_MAP_COL_LAST) _
    )

    ' ============================================================
    ' 1. GRUNDLEGENDES ALIGNMENT & RAHMEN
    ' ============================================================
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

    ' ============================================================
    ' 2. HORIZONTALE AUSRICHTUNG
    ' ============================================================
    ' Spalte S – EntityKey
    ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_ENTITYKEY), _
             ws.Cells(lastDataRow, DATA_MAP_COL_ENTITYKEY)).HorizontalAlignment = xlCenter

    ' Spalte W – Parzelle(n)
    ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_PARZELLE), _
             ws.Cells(lastDataRow, DATA_MAP_COL_PARZELLE)).HorizontalAlignment = xlCenter

    ' Restliche Spalten linksbündig
    ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_IBAN_OLD), _
             ws.Cells(lastDataRow, DATA_MAP_COL_KTONAME)).HorizontalAlignment = xlLeft

    ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_ZUORDNUNG), _
             ws.Cells(lastDataRow, DATA_MAP_COL_DEBUG)).HorizontalAlignment = xlLeft

    ' ============================================================
    ' 3. AUTOFIT – REIHENFOLGE ENTSCHEIDEND
    ' ============================================================
    ' Erst Spaltenbreite
    ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_ENTITYKEY), _
             ws.Cells(lastDataRow, DATA_MAP_COL_LAST)).EntireColumn.AutoFit

    ' Dann Zeilenhöhe
    ws.Rows(DATA_START_ROW & ":" & lastDataRow).AutoFit

    ' ============================================================
    ' 4. SPEZIALBEHANDLUNG SPALTE U (Kontoname)
    ' ============================================================
    ' Links, vertikal zentriert, mehrzeilig – stabil gegen AutoFit-Bug
    With ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_KTONAME), _
                  ws.Cells(lastDataRow, DATA_MAP_COL_KTONAME))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' Zeilenhöhe NACH Spalte-U-Korrektur nochmals sauber setzen
    ws.Rows(DATA_START_ROW & ":" & lastDataRow).AutoFit

    ' ============================================================
    ' 5. FARBRESET (NUR STRUKTURELL)
    ' ============================================================
    ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_ENTITYKEY), _
             ws.Cells(lastDataRow, DATA_MAP_COL_KTONAME)).Interior.color = COLOR_WHITE

    ' ============================================================
    ' 6. DROPDOWN – SPALTE X (EntityRole) BIS ZEILE 504+
    ' ============================================================
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
        .InputTitle = "Rolle wählen"
        .ErrorTitle = "Ungültige Rolle"
        .ErrorMessage = "Bitte wählen Sie eine gültige Rolle aus der Liste."
    End With

End Sub


' ===============================================================
' 2. CSV-KONTOAUSZUG IMPORT
' ***************************************************************
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

    ' Robustes Bereinigen
    On Error Resume Next
    ThisWorkbook.Worksheets(tempSheetName).Delete
    On Error GoTo 0

    ' 1. Datei auswählen
    strFile = Application.GetOpenFilename("CSV (*.csv), *.csv")
    If strFile = False Then
        Application.ScreenUpdating = True
        Call Initialize_ImportReport_ListBox ' Initialisiert die Listbox bei Abbruch
        Exit Sub
    End If

    ' --- Vorhandene Umsätze indexieren ---
    lRowZiel = wsZiel.Cells(wsZiel.Rows.Count, BK_COL_BETRAG).End(xlUp).Row
    If lRowZiel < BK_START_ROW Then lRowZiel = BK_START_ROW - 1

    For i = BK_START_ROW To lRowZiel
        If wsZiel.Cells(i, BK_COL_BETRAG).Value <> "" Then
            ' Schlüssel: Datum | Betrag | IBAN | Verwendungszweck
            sKey = Format(wsZiel.Cells(i, BK_COL_DATUM).Value, "YYYYMMDD") & "|" & CStr(wsZiel.Cells(i, BK_COL_BETRAG).Value) & "|" & Replace(CStr(wsZiel.Cells(i, BK_COL_IBAN).Value), " ", "") & ""
            dictUmsaetze(sKey) = True
        End If
    Next i

    ' 2. Temporäres Blatt erstellen
    On Error GoTo ImportFehler

    Set wsTemp = ThisWorkbook.Worksheets.Add(After:=wsZiel)
    wsTemp.Name = tempSheetName

    ' 3. CSV ROBUST IMPORTIEREN MIT QUERYTABLES
    With wsTemp.QueryTables.Add(Connection:="TEXT;" & strFile, Destination:=wsTemp.Cells(1, 1))
        .Name = "CSV_Import"
        .FieldNames = True
        .TextFilePlatform = xlUTF8Value
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimitedValue
        .TextFileSemicolonDelimiter = True
        .Refresh BackgroundQuery:=False
    End With

    ' 4. Bereinigen und vorbereiten
    lastRowTemp = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
    rowsTotalInFile = lastRowTemp - 1

    If lastRowTemp <= 1 Then
        rowsProcessed = 0
        GoTo ImportEnde
    End If

    wsTemp.QueryTables(1).Delete

    ' 5. Daten abgleichen und in "Bankkonto" schreiben

    For lRowTemp = 2 To lastRowTemp

        betragString = CStr(wsTemp.Cells(lRowTemp, CSV_COL_BETRAG).Value)

        ' Zahlensäuberung
        betragString = Replace(betragString, " EUR", "")
        betragString = Replace(betragString, "EUR", "")
        betragString = Trim(betragString)

        ' --- Betragsprüfung (Filterlogik) ---
        If betragString = "" Or Not IsNumeric(Replace(betragString, ",", ".")) Then
             rowsIgnoredFilter = rowsIgnoredFilter + 1
             GoTo NextRow
        End If

        ' Betrag konvertieren
        On Error Resume Next
        dBetrag = CDbl(Replace(betragString, ",", Application.International(xlDecimalSeparator)))
        If Err.Number <> 0 Then
            rowsIgnoredFilter = rowsIgnoredFilter + 1
            Err.Clear
            On Error GoTo ImportFehler
            GoTo NextRow
        End If
        On Error GoTo ImportFehler

        ' Datum extrahieren und konvertieren
        If IsDate(wsTemp.Cells(lRowTemp, CSV_COL_BUCHUNGSDATUM).Value) Then
            dDatum = CDate(wsTemp.Cells(lRowTemp, CSV_COL_BUCHUNGSDATUM).Value)
        Else
            rowsIgnoredFilter = rowsIgnoredFilter + 1
            GoTo NextRow
        End If

        sIBAN = Replace(Trim(wsTemp.Cells(lRowTemp, CSV_COL_IBAN).Value), " ", "")
        sName = Trim(wsTemp.Cells(lRowTemp, CSV_COL_NAME).Value)
        sVZ = Trim(wsTemp.Cells(lRowTemp, CSV_COL_VERWENDUNGSZWECK).Value)
        sText = Trim(wsTemp.Cells(lRowTemp, CSV_COL_STATUS).Value)

        ' Duplikatsprüfung mit dem stabilen Schlüssel
        sKey = Format(dDatum, "YYYYMMDD") & "|" & dBetrag & "|" & sIBAN & "|" & sVZ

        If dictUmsaetze.exists(sKey) Then
            rowsIgnoredDupe = rowsIgnoredDupe + 1
            GoTo NextRow
        End If

        ' Neue Zeile schreiben
        lRowZiel = wsZiel.Cells(wsZiel.Rows.Count, BK_COL_DATUM).End(xlUp).Row + 1
        dictUmsaetze.Add sKey, True

        wsZiel.Cells(lRowZiel, BK_COL_DATUM).Value = dDatum
        wsZiel.Cells(lRowZiel, BK_COL_DATUM).NumberFormat = "DD.MM.YYYY"

        wsZiel.Cells(lRowZiel, BK_COL_BETRAG).Value = dBetrag
        wsZiel.Cells(lRowZiel, BK_COL_BETRAG).NumberFormat = "#,##0.00 [$€-de-DE]"

        wsZiel.Cells(lRowZiel, BK_COL_NAME).Value = sName
        wsZiel.Cells(lRowZiel, BK_COL_IBAN).Value = sIBAN
        wsZiel.Cells(lRowZiel, BK_COL_VERWENDUNGSZWECK).Value = sVZ
        wsZiel.Cells(lRowZiel, BK_COL_BUCHUNGSTEXT).Value = sText
        wsZiel.Cells(lRowZiel, BK_COL_STATUS).Value = "Gebucht"

        rowsProcessed = rowsProcessed + 1

NextRow:
    Next lRowTemp

ImportEnde:

    ' Die Filtereinträge werden als Fehler im Protokoll gewertet
    rowsFailedImport = rowsIgnoredFilter

    ' 6. Protokollierung in ListBox-Historie
    Call Update_ImportReport_ListBox(rowsTotalInFile, rowsProcessed, rowsIgnoredDupe, rowsFailedImport)

    ' 7. Temporäres Blatt löschen
    If Not wsTemp Is Nothing Then
        wsTemp.Delete
        Set wsTemp = Nothing
    End If

    ' 8. Sortieren und Oberfläche aktualisieren
    Call Sortiere_Bankkonto_nach_Datum

    wsZiel.Activate

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ' Erfolgsmeldung
    If rowsTotalInFile > 0 And rowsProcessed = 0 And rowsIgnoredDupe = rowsTotalInFile And rowsFailedImport = 0 Then
          MsgBox "Achtung: Die ausgewählte CSV-Datei enthält ausschließlich Einträge, die bereits in der Datenbank vorhanden sind (" & rowsIgnoredDupe & " Duplikate). Es wurden keine neuen Daten hinzugefügt.", vbInformation
    ElseIf rowsProcessed > 0 Then
        MsgBox "Import abgeschlossen! (" & rowsProcessed & " neue Zeilen hinzugefügt)", vbInformation
    End If

    Exit Sub

ImportFehler:
    ' Bei fatalem Fehler
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ' Fehlerzähler setzen
    If rowsTotalInFile = 0 Then
        rowsFailedImport = 1
    Else
        rowsFailedImport = rowsFailedImport + 1
    End If

    ' Protokollierung versuchen
    Call Update_ImportReport_ListBox(rowsTotalInFile, rowsProcessed, rowsIgnoredDupe, rowsFailedImport)

    MsgBox "FATALER FEHLER beim Importieren der CSV-Datei. Fehler: " & Err.Description, vbCritical

    On Error Resume Next
    If Not wsTemp Is Nothing Then wsTemp.Delete
    wsZiel.Activate
    On Error GoTo 0
End Sub


' ===============================================================
' 3. HILFSPROZEDUREN ZUR SORTIERUNG
' ***************************************************************

Public Sub Sortiere_Bankkonto_nach_Datum()
    On Error GoTo SortError

    Dim ws As Worksheet
    Dim lr As Long

    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)

    lr = ws.Cells(ws.Rows.Count, BK_COL_DATUM).End(xlUp).Row

    ' Sicherstellen, dass Daten ab der Startzeile vorhanden sind
    If lr < BK_START_ROW Or IsEmpty(ws.Cells(BK_START_ROW, BK_COL_DATUM).Value) Then
        Exit Sub
    End If

    ' Sortierung des Datenbereichs von der Startzeile bis zur letzten belegten Zeile
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Cells(BK_START_ROW, BK_COL_DATUM), _
                             SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(BK_START_ROW, 1), ws.Cells(lr, BK_COL_STATUS))
        .Header = xlNo
        .Apply
    End With

    Exit Sub

SortError:
    MsgBox "Achtung: Die automatische Sortierung konnte nicht durchgeführt werden! Bitte prüfen Sie das Tabellenblatt '" & WS_BANKKONTO & "' manuell auf ungültige Datumswerte. Fehler: " & Err.Description, vbExclamation

End Sub

Public Sub Sortiere_Tabellen_Daten()

    Dim ws As Worksheet
    Dim lr As Long

    Application.EnableEvents = False
    On Error GoTo ExitClean

    Set ws = ThisWorkbook.Worksheets(WS_DATEN)

    ' --- Sortierung Kategorie-Regeln ---
    lr = ws.Cells(ws.Rows.Count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lr >= DATA_START_ROW Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Cells(DATA_START_ROW, DATA_CAT_COL_KATEGORIE), _
                                 Order:=xlAscending
            .SetRange ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), ws.Cells(lr, DATA_CAT_COL_END))
            .Header = xlNo
            .Apply
        End With
    End If

    ' --- Sortierung EntityKeys ---
    ' WICHTIG: Die Sortierung des Mappings soll nach EntityKey (ID) erfolgen
    lr = ws.Cells(ws.Rows.Count, DATA_MAP_COL_ENTITYKEY).End(xlUp).Row
    If lr >= DATA_START_ROW Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Cells(DATA_START_ROW, DATA_MAP_COL_ENTITYKEY), _
                                 Order:=xlAscending
            ' Sortierbereich muss alle Mapping-Spalten (bis Y) umfassen
            .SetRange ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_ENTITYKEY), ws.Cells(lr, DATA_MAP_COL_LAST))
            .Header = xlNo
            .Apply
        End With
    End If

ExitClean:
    Application.EnableEvents = True
End Sub


' ===============================================================
' 4. HILFSPROZEDUREN ZUR PROTOKOLLIERUNG
' ***************************************************************

' Stellt sicher, dass das temporäre Protokoll-Blatt existiert und gibt es zurück
Private Function Get_Protocol_Temp_Sheet() As Worksheet

    Dim ws As Worksheet
    Dim pwd As String
    Dim bProtected As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_PROTOCOL_TEMP)
    On Error GoTo 0

    ' Wenn ein Objekt mit dem Namen existiert, sicherstellen, dass es ein Worksheet ist
    If Not ws Is Nothing Then
        If TypeName(ws) <> "Worksheet" Then
            On Error Resume Next
            ' Versuche, das fehlerhafte Objekt zu entfernen, damit ein echtes Worksheet angelegt werden kann
            ThisWorkbook.Worksheets(WS_PROTOCOL_TEMP).Delete
            On Error GoTo 0
            Set ws = Nothing
        End If
    End If

    If ws Is Nothing Then
        ' Neues sehr-verborgenes Blatt anlegen (kein sichtbares Temp-Blatt für den Benutzer)
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        If Err.Number <> 0 Then
            Debug.Print "Get_Protocol_Temp_Sheet: Fehler beim Anlegen eines Blattes: " & Err.Description
            Err.Clear
            Application.DisplayAlerts = True
            Application.ScreenUpdating = True
            Exit Function
        End If
        On Error GoTo 0

        On Error Resume Next
        ws.Name = WS_PROTOCOL_TEMP
        If Err.Number <> 0 Then
            Debug.Print "Get_Protocol_Temp_Sheet: Name '" & WS_PROTOCOL_TEMP & "' konnte nicht gesetzt: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0

        ' Damit der Benutzer das Blatt nicht sieht
        On Error Resume Next
        ws.Visible = xlSheetVeryHidden
        On Error GoTo 0

        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    End If

    ' Protokollblatt-Spalte A formatieren, damit es Text aufnehmen kann
    ' Dieses Setzen ist robust: Fehler werden gefangen und führen nicht zu Laufzeitabbrüchen
    pwd = ""
    On Error Resume Next
    pwd = PASSWORD
    On Error GoTo 0

    On Error Resume Next
    If Not ws Is Nothing Then
        bProtected = ws.ProtectContents
        If bProtected Then
            ' Versuche temporär zu entsperren (falls Passwort vorhanden)
            On Error Resume Next
            If Len(Trim$(pwd)) > 0 Then
                ws.Unprotect PASSWORD:=pwd
            Else
                ws.Unprotect
            End If
            On Error GoTo 0
        End If

        ' Setze NumberFormat tolerant – falls das fehlschlägt, nur Debug-Ausgabe
        On Error Resume Next
        ws.Range("A:A").NumberFormat = "@"
        If Err.Number <> 0 Then
            Debug.Print "Get_Protocol_Temp_Sheet: NumberFormat konnte nicht gesetzt: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0

        ' Wenn wir vorher entsperrt haben, wieder schützen (UserInterfaceOnly=True für Makros)
        If bProtected Then
            On Error Resume Next
            If Len(Trim$(pwd)) > 0 Then
                ws.Protect PASSWORD:=pwd, UserInterfaceOnly:=True
            Else
                ws.Protect UserInterfaceOnly:=True
            End If
            On Error GoTo 0
        End If
    End If

    Set Get_Protocol_Temp_Sheet = ws

End Function

' Initialisiert die Form-Control ListBox beim Start oder Abbruch
Public Sub Initialize_ImportReport_ListBox()

    Dim wsZiel As Worksheet
    Dim wsDaten As Worksheet
    Dim wsTemp As Worksheet

    Set wsZiel = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsTemp = Get_Protocol_Temp_Sheet() ' Stellt das Hilfsblatt bereit

    Application.ScreenUpdating = False

    ' Nur initialisieren, wenn die Historie leer ist (d.h. noch kein Import durchgeführt)
    If CStr(wsDaten.Range(CELL_IMPORT_PROTOKOLL).Value) = "" Then

        ' 1. Temporären Bereich leeren
        If Not wsTemp Is Nothing Then wsTemp.Cells.ClearContents

        ' 2. Standardtext in den temporären Bereich schreiben
        If Not wsTemp Is Nothing Then
            wsTemp.Range(PROTOCOL_RANGE_START).Value = "--------------------------------------------------------"
            wsTemp.Range(PROTOCOL_RANGE_START).Offset(1, 0).Value = " Kein Import-Bericht verfügbar."
            wsTemp.Range(PROTOCOL_RANGE_START).Offset(2, 0).Value = " Führen Sie einen CSV-Import durch, um den Bericht hier anzuzeigen."
            wsTemp.Range(PROTOCOL_RANGE_START).Offset(3, 0).Value = "--------------------------------------------------------"
        End If

        ' 3. ListBox mit dem temporären Bereich verknüpfen
        On Error Resume Next
        If Not wsTemp Is Nothing Then
            wsZiel.Shapes(FORM_LISTBOX_NAME).ControlFormat.ListFillRange = "'" & WS_PROTOCOL_TEMP & "'!" & wsTemp.Range("A1:A4").Address(External:=False)
        End If
        On Error GoTo 0

    End If

    Application.ScreenUpdating = True
End Sub

' Aktualisiert die ListBox nach einem Import und speichert die Historie
Public Sub Update_ImportReport_ListBox(ByVal totalEntries As Long, ByVal importedEntries As Long, ByVal duplicateEntries As Long, ByVal errorEntries As Long)

    Dim wsZiel As Worksheet
    Dim wsDaten As Worksheet
    Dim wsTemp As Worksheet
    Dim protocolRange As String ' Der Bereich, der die ListBox füllt

    Dim strDateTime As String
    Dim currentHistory() As String
    Dim historyString As String
    Dim newHistoryString As String
    Dim i As Long, k As Long

    ' Konfiguration für die Speicherung der Historie
    Const HISTORY_DELIMITER As String = "|REPORT_DELIMITER|"
    Const PART_DELIMITER As String = "|PART|"
    Const MAX_REPORTS As Long = 12
    Const MAX_LISTBOX_LINES As Long = 50 ' Begrenzt die Anzahl der angezeigten Zeilen

    Set wsZiel = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsTemp = Get_Protocol_Temp_Sheet() ' Stellt das Hilfsblatt bereit

    strDateTime = Format(Now, "dd.mm.yyyy hh:nn:ss")

    Application.ScreenUpdating = False

    ' --- I. Historie Speichern und Kürzen ---

    ' BERICHTSTEILE FÜR DIE SPEICHERUNG
    Dim part1 As String: part1 = strDateTime
    Dim part2 As String: part2 = importedEntries & " / " & totalEntries & " importiert"
    Dim part3 As String: part3 = "Duplikate: " & duplicateEntries
    Dim part4 As String: part4 = "Fehler: " & errorEntries

    Dim newReportEntry As String
    newReportEntry = part1 & PART_DELIMITER & part2 & PART_DELIMITER & part3 & PART_DELIMITER & part4

    ' Historie laden und neue Historie erstellen
    historyString = CStr(wsDaten.Range(CELL_IMPORT_PROTOKOLL).Value)
    newHistoryString = newReportEntry & IIf(historyString <> "", HISTORY_DELIMITER & historyString, "")

    ' Historie auf maximale Größe trimmen und speichern
    currentHistory = Split(newHistoryString, HISTORY_DELIMITER)

    newHistoryString = ""
    For i = 0 To UBound(currentHistory)
        If i < MAX_REPORTS Then
            If i > 0 And newHistoryString <> "" Then newHistoryString = newHistoryString & HISTORY_DELIMITER
            newHistoryString = newHistoryString & currentHistory(i)
        Else
            Exit For
        End If
    Next i

    wsDaten.Range(CELL_IMPORT_PROTOKOLL).Value = newHistoryString

    ' --- II. Anzeige im Temporären Blatt und Verknüpfung ---

    If Not wsTemp Is Nothing Then wsTemp.Cells.ClearContents ' Vorherige Anzeige löschen
    k = 1 ' Zeilenzähler für das temporäre Blatt

    ' Wir durchlaufen die Historie und schreiben die Zeilen in das Temp-Blatt
    For i = 0 To UBound(currentHistory)

        Dim reportParts() As String
        reportParts = Split(currentHistory(i), PART_DELIMITER)

        ' 1. Zeile: Datum/Uhrzeit
        If UBound(reportParts) >= 0 Then wsTemp.Cells(k, 1).Value = Trim(reportParts(0)): k = k + 1
        ' 2. Zeile: Importierte Einträge
        If UBound(reportParts) >= 1 Then wsTemp.Cells(k, 1).Value = " * " & Trim(reportParts(1)): k = k + 1
        ' 3. Zeile: Duplikate
        If UBound(reportParts) >= 2 Then wsTemp.Cells(k, 1).Value = " * " & Trim(reportParts(2)): k = k + 1
        ' 4. Zeile: Fehler
        If UBound(reportParts) >= 3 Then wsTemp.Cells(k, 1).Value = " * " & Trim(reportParts(3)): k = k + 1

        ' 5. Trennlinie (wenn nicht der letzte Eintrag)
        If i < UBound(currentHistory) And UBound(currentHistory) > 0 Then
            wsTemp.Cells(k, 1).Value = "--------------------------------------------------------"
            k = k + 1
        End If

        ' Begrenzung der Zeilen für die ListBox (aus Performancegründen)
        If k >= MAX_LISTBOX_LINES Then Exit For
    Next i

    ' 3. ListBox verknüpfen
    On Error Resume Next
    ' Prüft, ob das Form Control existiert und verknüpft es mit dem gefüllten Range
    If Not wsZiel.Shapes(FORM_LISTBOX_NAME) Is Nothing Then
        protocolRange = wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(k - 1, 1)).Address(External:=False)
        wsZiel.Shapes(FORM_LISTBOX_NAME).ControlFormat.ListFillRange = "'" & WS_PROTOCOL_TEMP & "'!" & protocolRange
    End If

    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub


' ===============================================================
' 5. KATEGORISIERUNGS-LOGIK (ZENTRALE STEUERUNG)
' ***************************************************************
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

    ' -------------------------------------------------------------
    ' ZENTRALE STEUERUNG der KATEGORISIERUNGS-ENGINE
    ' Ruft die Public Sub 'KategorieEngine_Pipeline' auf (im separaten Modul).
    ' -------------------------------------------------------------

    ' Aufruf der neuen, optimierten Engine
    Call KategorieEngine_Pipeline(wsBK)

    ' -------------------------------------------------------------

    ' Nach der Kategorisierung: Tabelle erneut sortieren
    Call Sortiere_Bankkonto_nach_Datum

    MsgBox "Die Kategorisierung der Banktransaktionen wurde abgeschlossen.", vbInformation

ExitClean:
    Application.ScreenUpdating = True
    Application.EnableEvents = True

CategorizationError:
    If Err.Number <> 0 Then
        MsgBox "Ein Fehler ist bei der Kategorisierung aufgetreten: " & Err.Description, vbCritical
        GoTo ExitClean
    End If
End Sub


' ===============================================================
' Ende Modul
' ===============================================================

