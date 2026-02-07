Attribute VB_Name = "mod_Banking_Data"
Option Explicit

' ===============================================================
' MODUL: mod_Banking_Data
' VERSION: 3.2 - 07.02.2026
' AENDERUNG: ListBox als Formularsteuerelement korrekt,
'            Y500 als einzige Speicherzelle (serialisiert),
'            MsgBox mit vollstaendigen Import-Details,
'            Farbcodierung ueber Rahmen-Shape
' ===============================================================

Private Const ZEBRA_COLOR As Long = &HDEE5E3
Private Const RAHMEN_NAME As String = "ImportReport_Rahmen"

' Farb-Konstanten fuer Rahmen-Hintergrund
' VBA nutzt BGR-Reihenfolge bei Long-Werten!
Private Const RAHMEN_COLOR_GRUEN As Long = 11534528    ' RGB(0,176,80) dunkelgruen
Private Const RAHMEN_COLOR_GELB As Long = 65535        ' RGB(255,255,0) gelb
Private Const RAHMEN_COLOR_ROT As Long = 255           ' RGB(255,0,0) rot
Private Const RAHMEN_COLOR_WEISS As Long = 16777215    ' RGB(255,255,255) weiss

' Trennzeichen fuer Serialisierung in Zelle Y500
Private Const PROTO_SEP As String = "||"

' Hilfsbereich fuer ListFillRange: Spalte Y ab Zeile 501
' (wird bei jedem Initialize/Update neu beschrieben)
Private Const LB_FILL_START_ROW As Long = 501
Private Const LB_FILL_SPALTE As Long = 25             ' Spalte Y

' Maximale Anzahl Import-Bloecke im Speicher (je 5 Zeilen)
Private Const MAX_BLOECKE As Long = 12
' 12 x 5 = 60 Zeilen maximal
Private Const MAX_ZEILEN As Long = 60

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
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    On Error Resume Next
    ThisWorkbook.Unprotect PASSWORD:=PASSWORD
    Err.Clear
    On Error GoTo 0
    
    Set wsZiel = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    On Error Resume Next
    wsZiel.Unprotect PASSWORD:=PASSWORD
    Err.Clear
    On Error GoTo 0
    
    On Error Resume Next
    ThisWorkbook.Worksheets(tempSheetName).Delete
    Err.Clear
    On Error GoTo 0
    
    Set dictUmsaetze = CreateObject("Scripting.Dictionary")
    
    rowsProcessed = 0
    rowsIgnoredDupe = 0
    rowsIgnoredFilter = 0
    rowsFailedImport = 0
    rowsTotalInFile = 0
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    strFile = Application.GetOpenFilename("CSV (*.csv), *.csv")
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If strFile = False Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        wsZiel.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        Call Initialize_ImportReport_ListBox
        Exit Sub
    End If
    
    lRowZiel = wsZiel.Cells(wsZiel.Rows.count, BK_COL_BETRAG).End(xlUp).Row
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
    Set wsTemp = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    If Err.Number <> 0 Then
        MsgBox "Fehler beim Erstellen des Temp-Blatts: " & Err.Description & vbCrLf & vbCrLf & _
           "Bitte pruefen Sie ob die Arbeitsmappe geschuetzt ist.", vbCritical
        Err.Clear
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.EnableEvents = True
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
        Application.EnableEvents = True
        wsZiel.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    
    lastRowTemp = wsTemp.Cells(wsTemp.Rows.count, 1).End(xlUp).Row
    rowsTotalInFile = lastRowTemp - 1
    
    If lastRowTemp <= 1 Then
        rowsProcessed = 0
        GoTo ImportAbschluss
    End If
    
    On Error Resume Next
    wsTemp.QueryTables(1).Delete
    Err.Clear
    On Error GoTo 0
    
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
        
        lRowZiel = wsZiel.Cells(wsZiel.Rows.count, BK_COL_DATUM).End(xlUp).Row + 1
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
    
    ' ListBox und Protokoll-Speicher aktualisieren
    Call Update_ImportReport_ListBox(rowsTotalInFile, rowsProcessed, rowsIgnoredDupe, rowsFailedImport)
    
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wsTemp Is Nothing Then wsTemp.Delete
    Application.DisplayAlerts = True
    Set wsTemp = Nothing
    Err.Clear
    On Error GoTo 0
    
    ' ============================================================
    ' WICHTIG: Reihenfolge der Nachbearbeitung nach CSV-Import
    ' EXPLIZITE Modulangabe um Mehrdeutigkeiten zu vermeiden!
    ' ============================================================
    On Error Resume Next
    
    ' 1. IBANs aus Bankkonto in EntityKey-Tabelle importieren
    Call mod_EntityKey_Manager.ImportiereIBANsAusBankkonto
    
    ' 2. EntityKeys aktualisieren (GUIDs, Zuordnungen, Ampel, Formatierung)
    Call mod_EntityKey_Manager.AktualisiereAlleEntityKeys
    
    ' 3. Bankkonto sortieren (AUFSTEIGEND - Januar oben)
    Call Sortiere_Bankkonto_nach_Datum
    
    ' 4. Formatierungen anwenden
    Call Anwende_Zebra_Bankkonto(wsZiel)
    Call Anwende_Border_Bankkonto(wsZiel)
    Call Anwende_Formatierung_Bankkonto(wsZiel)
    
    ' 5. Kategorie-Engine nur bei neuen Zeilen
    If rowsProcessed > 0 Then Call KategorieEngine_Pipeline(wsZiel)
    
    ' 6. Monat/Periode setzen
    Call Setze_Monat_Periode(wsZiel)
    
    Err.Clear
    On Error GoTo 0
    
    wsZiel.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    wsZiel.Activate

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' ============================================================
    ' ERWEITERTE MsgBox mit vollstaendigen Import-Details
    ' ============================================================
    Dim msgIcon As VbMsgBoxStyle
    Dim msgTitle As String
    Dim msgText As String
    
    If rowsFailedImport > 0 Then
        msgIcon = vbCritical
        msgTitle = "Import mit Fehlern"
    ElseIf rowsIgnoredDupe > 0 And rowsProcessed = 0 Then
        msgIcon = vbExclamation
        msgTitle = "100% Duplikate erkannt"
    ElseIf rowsIgnoredDupe > 0 Then
        msgIcon = vbExclamation
        msgTitle = "Import mit Duplikaten"
    ElseIf rowsProcessed > 0 Then
        msgIcon = vbInformation
        msgTitle = "Import erfolgreich"
    Else
        msgIcon = vbInformation
        msgTitle = "Import abgeschlossen"
    End If
    
    msgText = "CSV-Import Ergebnis:" & vbCrLf & _
              "==============================" & vbCrLf & vbCrLf & _
              "Datensaetze in CSV:     " & rowsTotalInFile & vbCrLf & _
              "Importiert:                    " & rowsProcessed & " / " & rowsTotalInFile & vbCrLf & _
              "Duplikate:                     " & rowsIgnoredDupe & vbCrLf & _
              "Fehler:                          " & rowsFailedImport & vbCrLf & vbCrLf
    
    If rowsFailedImport > 0 Then
        msgText = msgText & "ACHTUNG: " & rowsFailedImport & " Zeilen konnten nicht verarbeitet werden!"
    ElseIf rowsProcessed = 0 And rowsIgnoredDupe > 0 Then
        msgText = msgText & "Alle Eintraege waren bereits in der Datenbank vorhanden."
    ElseIf rowsProcessed > 0 And rowsIgnoredDupe = 0 Then
        msgText = msgText & "Alle Datensaetze wurden erfolgreich importiert."
    ElseIf rowsProcessed > 0 And rowsIgnoredDupe > 0 Then
        msgText = msgText & rowsProcessed & " neue Datensaetze importiert, " & _
                  rowsIgnoredDupe & " Duplikate uebersprungen."
    End If
    
    MsgBox msgText, msgIcon, msgTitle
    
End Sub

' ===============================================================
' 2. ZEBRA-FORMATIERUNG (A-G und I-Z, Spalte H ausgenommen)
' ===============================================================
Private Sub Anwende_Zebra_Bankkonto(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim lRow As Long
    Dim rngPart1 As Range
    Dim rngPart2 As Range
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    For lRow = BK_START_ROW To lastRow
        ' Teil 1: Spalten A-G (1-7)
        Set rngPart1 = ws.Range(ws.Cells(lRow, 1), ws.Cells(lRow, 7))
        ' Teil 2: Spalten I-Z (9-26) - Spalte H (8) ausgenommen!
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
' 3. RAHMEN-FORMATIERUNG
' ===============================================================
Private Sub Anwende_Border_Bankkonto(ByVal ws As Worksheet)
    
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
' 4. ALLGEMEINE FORMATIERUNG
' ===============================================================
Private Sub Anwende_Formatierung_Bankkonto(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim euroFormat As String
    
    If ws Is Nothing Then Exit Sub
    
    euroFormat = "#,##0.00 " & ChrW(8364)
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
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



'--- Ende Teil 1 ---
'--- Anfang Teil 2 ---



' ===============================================================
' 5. SORTIERUNG NACH DATUM (AUFSTEIGEND - Januar oben)
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
' 6. MONAT/PERIODE SETZEN
' ===============================================================
Private Sub Setze_Monat_Periode(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim r As Long
    Dim monatWert As Variant
    Dim datumWert As Variant
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    For r = BK_START_ROW To lastRow
        datumWert = ws.Cells(r, BK_COL_DATUM).value
        monatWert = ws.Cells(r, BK_COL_MONAT_PERIODE).value
        
        If IsDate(datumWert) And (isEmpty(monatWert) Or monatWert = "") Then
            ws.Cells(r, BK_COL_MONAT_PERIODE).value = MonthName(Month(datumWert))
        End If
    Next r
    
End Sub

' ===============================================================
' 7. IMPORT REPORT LISTBOX (FORMULARSTEUERELEMENT)
'    -----------------------------------------------
'    Architektur:
'    - Speicher: Daten!Y500 (eine einzige Zelle, serialisiert
'      mit "||" als Trennzeichen zwischen Zeilen)
'    - Anzeige: Daten!Y501:Y5xx als Hilfsbereich, der aus Y500
'      deserialisiert wird. ListFillRange zeigt darauf.
'    - Farbe: Rahmen-Shape "ImportReport_Rahmen" wird eingefaerbt.
'    - Pro Import-Vorgang: 5 Zeilen (Datum, X/Y, Dupes, Fehler, ----)
'    - Max 12 Bloecke = 60 Zeilen Historie
' ===============================================================

' ---------------------------------------------------------------
' 7a. Initialize: Liest Y500, deserialisiert in Y501ff,
'     setzt ListFillRange und Rahmen-Farbe.
'     Aufruf: Workbook_Open, Worksheet_Activate, nach Loeschen
' ---------------------------------------------------------------
Public Sub Initialize_ImportReport_ListBox()
    
    Dim wsBK As Worksheet
    Dim wsDaten As Worksheet
    Dim shpLB As Shape
    Dim gespeichert As String
    Dim zeilen() As String
    Dim anzahl As Long
    Dim i As Long
    Dim letzteZeile As Long
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsBK Is Nothing Or wsDaten Is Nothing Then Exit Sub
    
    ' Formularsteuerelement-ListBox finden
    Set shpLB = HoleFormListBox(wsBK)
    If shpLB Is Nothing Then Exit Sub
    
    ' Daten-Blatt entsperren
    On Error Resume Next
    wsDaten.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' Hilfsbereich Y501:Y560 komplett leeren
    Call LeereHilfsbereich(wsDaten)
    
    ' Gespeichertes Protokoll aus Y500 lesen
    gespeichert = CStr(wsDaten.Cells(500, LB_FILL_SPALTE).value)
    
    If gespeichert = "" Or gespeichert = "0" Then
        ' Kein Protokoll vorhanden - Standardtext
        wsDaten.Cells(LB_FILL_START_ROW, LB_FILL_SPALTE).value = _
            "Es wurden noch keine Daten importiert."
        
        On Error Resume Next
        wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        On Error GoTo 0
        
        ' ListFillRange auf die eine Zeile setzen
        shpLB.ControlFormat.ListFillRange = _
            "'" & WS_DATEN & "'!Y" & LB_FILL_START_ROW & ":Y" & LB_FILL_START_ROW
        
        ' Rahmen weiss
        Call FaerbeRahmen(wsBK, RAHMEN_COLOR_WEISS)
        Exit Sub
    End If
    
    ' Protokoll-Zeilen aus Y500 deserialisieren
    zeilen = Split(gespeichert, PROTO_SEP)
    anzahl = UBound(zeilen) + 1
    If anzahl > MAX_ZEILEN Then anzahl = MAX_ZEILEN
    
    ' In Hilfsbereich Y501ff schreiben
    For i = 0 To anzahl - 1
        wsDaten.Cells(LB_FILL_START_ROW + i, LB_FILL_SPALTE).value = zeilen(i)
    Next i
    
    On Error Resume Next
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
    ' ListFillRange setzen
    letzteZeile = LB_FILL_START_ROW + anzahl - 1
    shpLB.ControlFormat.ListFillRange = _
        "'" & WS_DATEN & "'!Y" & LB_FILL_START_ROW & ":Y" & letzteZeile
    
    ' Farbe aus juengstem Block bestimmen
    Call FaerbeRahmenAusProtokoll(wsBK, zeilen)
    
End Sub

' ---------------------------------------------------------------
' 7b. Update: Neuen 5-Zeilen-Block OBEN einfuegen,
'     in Y500 serialisiert speichern, Hilfsbereich aktualisieren,
'     ListFillRange setzen, Rahmen einfaerben.
' ---------------------------------------------------------------
Private Sub Update_ImportReport_ListBox(ByVal totalRows As Long, ByVal imported As Long, _
                                         ByVal dupes As Long, ByVal failed As Long)
    
    Dim wsBK As Worksheet
    Dim wsDaten As Worksheet
    Dim shpLB As Shape
    Dim altGespeichert As String
    Dim neuerBlock As String
    Dim gesamt As String
    Dim zeilen() As String
    Dim anzahl As Long
    Dim i As Long
    Dim letzteZeile As Long
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsBK Is Nothing Or wsDaten Is Nothing Then Exit Sub
    
    ' --- 5-Zeilen-Block zusammenbauen ---
    neuerBlock = "Import: " & Format(Now, "DD.MM.YYYY  HH:MM:SS") & _
                 PROTO_SEP & _
                 imported & " / " & totalRows & " Datensaetze importiert" & _
                 PROTO_SEP & _
                 dupes & " Duplikate erkannt" & _
                 PROTO_SEP & _
                 failed & " Fehler" & _
                 PROTO_SEP & _
                 "--------------------------------------"
    
    ' --- Daten-Blatt entsperren ---
    On Error Resume Next
    wsDaten.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' --- Alten Inhalt aus Y500 laden ---
    altGespeichert = CStr(wsDaten.Cells(500, LB_FILL_SPALTE).value)
    
    ' Wenn alter Inhalt leer oder Default-Text, ignorieren
    If altGespeichert = "" Or altGespeichert = "0" Or _
       altGespeichert = "Es wurden noch keine Daten importiert." Then
        gesamt = neuerBlock
    Else
        gesamt = neuerBlock & PROTO_SEP & altGespeichert
    End If
    
    ' --- Auf MAX_ZEILEN begrenzen ---
    zeilen = Split(gesamt, PROTO_SEP)
    anzahl = UBound(zeilen) + 1
    If anzahl > MAX_ZEILEN Then
        ' Abschneiden und neu zusammensetzen
        gesamt = zeilen(0)
        For i = 1 To MAX_ZEILEN - 1
            gesamt = gesamt & PROTO_SEP & zeilen(i)
        Next i
        anzahl = MAX_ZEILEN
    End If
    
    ' --- In Y500 speichern (eine einzige Zelle!) ---
    wsDaten.Cells(500, LB_FILL_SPALTE).value = gesamt
    
    ' --- Hilfsbereich Y501ff leeren und neu beschreiben ---
    Call LeereHilfsbereich(wsDaten)
    
    zeilen = Split(gesamt, PROTO_SEP)
    For i = 0 To anzahl - 1
        wsDaten.Cells(LB_FILL_START_ROW + i, LB_FILL_SPALTE).value = zeilen(i)
    Next i
    
    ' --- Daten-Blatt schuetzen ---
    On Error Resume Next
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
    ' --- ListBox aktualisieren ---
    Set shpLB = HoleFormListBox(wsBK)
    If Not shpLB Is Nothing Then
        letzteZeile = LB_FILL_START_ROW + anzahl - 1
        shpLB.ControlFormat.ListFillRange = _
            "'" & WS_DATEN & "'!Y" & LB_FILL_START_ROW & ":Y" & letzteZeile
    End If
    
    ' --- Farbcodierung anwenden ---
    Call FaerbeRahmenNachImport(wsBK, imported, dupes, failed)
    
End Sub

' ---------------------------------------------------------------
' 7c. Hilfsbereich Y501:Y560 komplett leeren
' ---------------------------------------------------------------
Private Sub LeereHilfsbereich(ByVal wsDaten As Worksheet)
    Dim i As Long
    For i = 0 To MAX_ZEILEN - 1
        wsDaten.Cells(LB_FILL_START_ROW + i, LB_FILL_SPALTE).ClearContents
    Next i
End Sub

' ---------------------------------------------------------------
' 7d. Formularsteuerelement-ListBox finden (ueber Shapes)
' ---------------------------------------------------------------
Private Function HoleFormListBox(ByVal ws As Worksheet) As Shape
    
    Dim shp As Shape
    
    On Error Resume Next
    Set shp = ws.Shapes(FORM_LISTBOX_NAME)
    On Error GoTo 0
    
    If shp Is Nothing Then
        Set HoleFormListBox = Nothing
    Else
        Set HoleFormListBox = shp
    End If
    
End Function

' ---------------------------------------------------------------
' 7e. Farbcodierung nach Import-Ergebnis
'     GRUEN  = Alles OK (dupes = 0, failed = 0)
'     GELB   = Duplikate vorhanden (dupes > 0, failed = 0)
'     ROT    = Fehler vorhanden (failed > 0)
' ---------------------------------------------------------------
Private Sub FaerbeRahmenNachImport(ByVal ws As Worksheet, _
                                    ByVal imported As Long, _
                                    ByVal dupes As Long, _
                                    ByVal failed As Long)
    Dim farbe As Long
    
    If failed > 0 Then
        farbe = RAHMEN_COLOR_ROT
    ElseIf dupes > 0 Then
        farbe = RAHMEN_COLOR_GELB
    Else
        farbe = RAHMEN_COLOR_GRUEN
    End If
    
    Call FaerbeRahmen(ws, farbe)
    
End Sub

' ---------------------------------------------------------------
' 7f. Farbcodierung aus gespeichertem Protokoll bestimmen
'     (fuer Initialize beim Oeffnen der Arbeitsmappe)
'     Liest Zeile 3 (Index 2): "X Duplikate erkannt"
'     Liest Zeile 4 (Index 3): "X Fehler"
' ---------------------------------------------------------------
Private Sub FaerbeRahmenAusProtokoll(ByVal wsBK As Worksheet, ByRef zeilen() As String)
    
    Dim dupes As Long
    Dim failed As Long
    
    If UBound(zeilen) < 3 Then
        Call FaerbeRahmen(wsBK, RAHMEN_COLOR_WEISS)
        Exit Sub
    End If
    
    dupes = ExtrahiereZahl(CStr(zeilen(2)))
    failed = ExtrahiereZahl(CStr(zeilen(3)))
    
    If failed > 0 Then
        Call FaerbeRahmen(wsBK, RAHMEN_COLOR_ROT)
    ElseIf dupes > 0 Then
        Call FaerbeRahmen(wsBK, RAHMEN_COLOR_GELB)
    Else
        Call FaerbeRahmen(wsBK, RAHMEN_COLOR_GRUEN)
    End If
    
End Sub

' ---------------------------------------------------------------
' 7g. Zahl am Anfang eines Strings extrahieren
'     "123 Duplikate erkannt" -> 123
' ---------------------------------------------------------------
Private Function ExtrahiereZahl(ByVal text As String) As Long
    
    Dim i As Long
    Dim zahlStr As String
    
    zahlStr = ""
    For i = 1 To Len(text)
        If Mid(text, i, 1) >= "0" And Mid(text, i, 1) <= "9" Then
            zahlStr = zahlStr & Mid(text, i, 1)
        Else
            If zahlStr <> "" Then Exit For
        End If
    Next i
    
    If zahlStr <> "" Then
        ExtrahiereZahl = CLng(zahlStr)
    Else
        ExtrahiereZahl = 0
    End If
    
End Function

' ---------------------------------------------------------------
' 7h. Rahmen-Shape einfaerben (Hintergrund)
'     Das Shape "ImportReport_Rahmen" liegt visuell hinter/um
'     die ListBox und dient als farblicher Indikator.
' ---------------------------------------------------------------
Private Sub FaerbeRahmen(ByVal ws As Worksheet, ByVal farbe As Long)
    
    Dim shp As Shape
    
    On Error Resume Next
    Set shp = ws.Shapes(RAHMEN_NAME)
    On Error GoTo 0
    
    If shp Is Nothing Then Exit Sub
    
    On Error Resume Next
    With shp.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = farbe
        .Transparency = 0
    End With
    On Error GoTo 0
    
End Sub

' ===============================================================
' 8. HILFSFUNKTIONEN
' ===============================================================
Public Sub LoescheAlleBankkontoZeilen()
    
    Dim ws As Worksheet
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim antwort As VbMsgBoxResult
    
    antwort = MsgBox("ACHTUNG: Alle Daten auf dem Bankkonto-Blatt werden geloescht!" & vbCrLf & vbCrLf & _
                     "Fortfahren?", vbYesNo + vbCritical, "Alle Daten loeschen?")
    
    If antwort <> vbYes Then Exit Sub
    
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    If lastRow >= BK_START_ROW Then
        ws.Range(ws.Cells(BK_START_ROW, 1), ws.Cells(lastRow, 26)).ClearContents
        ws.Range(ws.Cells(BK_START_ROW, 1), ws.Cells(lastRow, 26)).Interior.ColorIndex = xlNone
    End If
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' Protokoll-Speicher leeren: Y500 + Hilfsbereich Y501:Y560
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    If Not wsDaten Is Nothing Then
        wsDaten.Unprotect PASSWORD:=PASSWORD
        wsDaten.Cells(500, LB_FILL_SPALTE).ClearContents
        Call LeereHilfsbereich(wsDaten)
        wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    On Error GoTo 0
    
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
' 9. SORTIERE TABELLEN DATEN
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

    lr = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
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

    lr = ws.Cells(ws.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
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

