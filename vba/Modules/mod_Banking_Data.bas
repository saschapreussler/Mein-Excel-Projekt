Attribute VB_Name = "mod_Banking_Data"
Option Explicit

' ===============================================================
' MODUL: mod_Banking_Data
' VERSION: 3.7 - 08.02.2026
' AENDERUNG: Redundante Funktion HoleActiveXListBox entfernt
'            (wurde nirgends aufgerufen)
' ===============================================================

Private Const ZEBRA_COLOR As Long = &HDEE5E3

' Farb-Konstanten fuer ListBox-Hintergrund (OLE_COLOR / BGR)
Private Const LB_COLOR_GRUEN As Long = &HC0FFC0     ' hellgruen
Private Const LB_COLOR_GELB As Long = &HC0FFFF      ' hellgelb
Private Const LB_COLOR_ROT As Long = &HC0C0FF       ' hellrot
Private Const LB_COLOR_WEISS As Long = &HFFFFFF     ' weiss

' Trennzeichen fuer Serialisierung in Zelle Y500
Private Const PROTO_SEP As String = "||"

' Protokoll-Speicher: Zelle Y500 auf dem Daten-Blatt
Private Const PROTO_ZEILE As Long = 500
Private Const PROTO_SPALTE As Long = 25              ' Spalte Y

' Maximale Anzahl Import-Bloecke im Speicher (je 5 Zeilen)
Private Const MAX_BLOECKE As Long = 100
' 100 x 5 = 500 Zeilen maximal
Private Const MAX_ZEILEN As Long = 500


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
    
    Err.Clear
    On Error GoTo 0
    
    ' 5. Kategorie-Engine nur bei neuen Zeilen
    ' WICHTIG: On Error GoTo 0 MUSS vorher stehen,
    ' damit die Pipeline ihr eigenes Error-Handling nutzen kann
    ' und nicht das "On Error Resume Next" von oben erbt!
    If rowsProcessed > 0 Then Call KategorieEngine_Pipeline(wsZiel)
    
    ' 6. Monat/Periode setzen
    On Error Resume Next
    Call Setze_Monat_Periode(wsZiel)
    Err.Clear
    On Error GoTo 0
    
    ' Blattschutz wird von der Pipeline selbst verwaltet (Protect am Ende).
    ' Hier nochmals sicherstellen falls Pipeline nicht lief:
    On Error Resume Next
    wsZiel.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
    ' 7. Formeln wiederherstellen (koennten durch Import/Sort ueberschrieben sein)
    Call StelleFormelnWiederHer(wsZiel)
    
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
              String(30, "=") & vbCrLf & vbCrLf & _
              "Datens" & ChrW(228) & "tze in CSV:" & vbTab & rowsTotalInFile & vbCrLf & _
              "Importiert:" & vbTab & vbTab & rowsProcessed & " / " & rowsTotalInFile & vbCrLf & _
              "Duplikate:" & vbTab & vbTab & rowsIgnoredDupe & vbCrLf & _
              "Fehler:" & vbTab & vbTab & vbTab & rowsFailedImport & vbCrLf & vbCrLf
    
    If rowsFailedImport > 0 Then
        msgText = msgText & "ACHTUNG: " & rowsFailedImport & " Zeilen konnten nicht verarbeitet werden!"
    ElseIf rowsProcessed = 0 And rowsIgnoredDupe > 0 Then
        msgText = msgText & "Alle Eintr" & ChrW(228) & "ge waren bereits in der Datenbank vorhanden."
    ElseIf rowsProcessed > 0 And rowsIgnoredDupe = 0 Then
        msgText = msgText & "Alle Datens" & ChrW(228) & "tze wurden erfolgreich importiert."
    ElseIf rowsProcessed > 0 And rowsIgnoredDupe > 0 Then
        msgText = msgText & rowsProcessed & " neue Datens" & ChrW(228) & "tze importiert," & vbCrLf & _
                  rowsIgnoredDupe & " Duplikate " & ChrW(252) & "bersprungen."
    End If
    
    MsgBox msgText, msgIcon, msgTitle
    
    ' ============================================================
    ' ENTITYKEY-PRUEFUNG: Spalte W (EntityRole) vollstaendig?
    ' Nur pruefen wenn tatsaechlich neue Datensaetze importiert wurden
    ' ============================================================
    If rowsProcessed > 0 Then
        Call PruefeUnvollstaendigeEntityKeys
    End If
    
End Sub


' ===============================================================
' 1b. ENTITYKEY-PRUEFUNG NACH IMPORT
'     Prueft ob alle IBANs in der EntityKey-Tabelle (Daten! R-X)
'     eine vollstaendige Zuordnung in Spalte W (EntityRole) haben.
'     Bei fehlenden Eintraegen: MsgBox mit Angebot zur Navigation.
' ===============================================================
Private Sub PruefeUnvollstaendigeEntityKeys()
    
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim ersteLeereZeile As Long
    Dim anzahlOhneRole As Long
    Dim ibanOhneRole As String
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    If wsDaten Is Nothing Then Exit Sub
    
    lastRow = wsDaten.Cells(wsDaten.Rows.count, EK_COL_IBAN).End(xlUp).Row
    If lastRow < EK_START_ROW Then Exit Sub
    
    ersteLeereZeile = 0
    anzahlOhneRole = 0
    ibanOhneRole = ""
    
    For r = EK_START_ROW To lastRow
        ' Nur Zeilen pruefen die eine IBAN haben
        If Trim(CStr(wsDaten.Cells(r, EK_COL_IBAN).value)) <> "" Then
            ' Spalte W (EntityRole) leer?
            If Trim(CStr(wsDaten.Cells(r, EK_COL_ROLE).value)) = "" Then
                anzahlOhneRole = anzahlOhneRole + 1
                
                ' Erste leere Zeile merken
                If ersteLeereZeile = 0 Then ersteLeereZeile = r
                
                ' Maximal 5 IBANs fuer die Anzeige sammeln
                If anzahlOhneRole <= 5 Then
                    Dim kontoname As String
                    kontoname = Trim(CStr(wsDaten.Cells(r, EK_COL_KONTONAME).value))
                    If kontoname <> "" Then
                        ibanOhneRole = ibanOhneRole & vbCrLf & "  " & ChrW(8226) & " " & _
                            Left(CStr(wsDaten.Cells(r, EK_COL_IBAN).value), 12) & "...  (" & kontoname & ")"
                    Else
                        ibanOhneRole = ibanOhneRole & vbCrLf & "  " & ChrW(8226) & " " & _
                            CStr(wsDaten.Cells(r, EK_COL_IBAN).value)
                    End If
                End If
            End If
        End If
    Next r
    
    ' Keine fehlenden Eintraege -> nichts tun
    If anzahlOhneRole = 0 Then Exit Sub
    
    ' MsgBox zusammenbauen
    Dim hinweis As String
    hinweis = "Nach dem Import wurden " & anzahlOhneRole & _
              " IBAN-Zuordnung(en) ohne EntityRole (Spalte W) gefunden:" & _
              vbCrLf & ibanOhneRole
    
    If anzahlOhneRole > 5 Then
        hinweis = hinweis & vbCrLf & "  ... und " & (anzahlOhneRole - 5) & " weitere"
    End If
    
    hinweis = hinweis & vbCrLf & vbCrLf & _
              "Ohne diese Zuordnung kann die Kategorie-Engine die Buchungen " & _
              "nicht korrekt verarbeiten." & vbCrLf & vbCrLf & _
              "M" & ChrW(246) & "chten Sie die fehlenden Angaben jetzt vervollst" & ChrW(228) & "ndigen?"
    
    Dim antwort As VbMsgBoxResult
    antwort = MsgBox(hinweis, vbYesNo + vbExclamation, _
                     "Unvollst" & ChrW(228) & "ndige IBAN-Zuordnungen")
    
    If antwort = vbYes Then
        ' Zum Daten-Blatt wechseln und erste leere Zelle in Spalte W anwaehlen
        wsDaten.Activate
        
        On Error Resume Next
        wsDaten.Unprotect PASSWORD:=PASSWORD
        On Error GoTo 0
        
        wsDaten.Cells(ersteLeereZeile, EK_COL_ROLE).Select
        
        wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    
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
    
    ' Spalte B (Betrag): Waehrung + rechtsbuendig
    With ws.Range(ws.Cells(BK_START_ROW, BK_COL_BETRAG), ws.Cells(lastRow, BK_COL_BETRAG))
        .NumberFormat = euroFormat
        .HorizontalAlignment = xlRight
    End With
    
    ' Spalten M-Z: Waehrung
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_MITGL_BEITR), ws.Cells(lastRow, BK_COL_AUSZAHL_KASSE)).NumberFormat = euroFormat
    
    With ws.Range(ws.Cells(BK_START_ROW, BK_COL_BEMERKUNG), ws.Cells(lastRow, BK_COL_BEMERKUNG))
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    
    ws.Cells.VerticalAlignment = xlCenter
    ws.Rows(BK_START_ROW & ":" & lastRow).AutoFit
    
End Sub



'--- Ende Teil 1 von 3 ---
'--- Anfang Teil 2 von 3 ---




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
' 6. MONAT/PERIODE SETZEN (intelligent ueber Einstellungen)
' ===============================================================
Private Sub Setze_Monat_Periode(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim r As Long
    Dim monatWert As Variant
    Dim datumWert As Variant
    Dim kategorie As String
    Dim faelligkeit As String
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    ' Faelligkeit aus Kategorie-Tabelle vorladen
    Dim wsDaten As Worksheet
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    
    For r = BK_START_ROW To lastRow
        datumWert = ws.Cells(r, BK_COL_DATUM).value
        monatWert = ws.Cells(r, BK_COL_MONAT_PERIODE).value
        
        If IsDate(datumWert) And (isEmpty(monatWert) Or monatWert = "") Then
            kategorie = Trim(ws.Cells(r, BK_COL_KATEGORIE).value)
            
            If kategorie <> "" Then
                ' Faelligkeit aus Kategorie-Tabelle holen (Spalte O)
                faelligkeit = HoleFaelligkeitFuerKategorie(wsDaten, kategorie)
                ' Intelligente Monat/Periode-Ermittlung
                ws.Cells(r, BK_COL_MONAT_PERIODE).value = _
                    ErmittleMonatPeriode(kategorie, CDate(datumWert), faelligkeit)
            Else
                ' Keine Kategorie: Fallback auf Buchungsmonat
                ws.Cells(r, BK_COL_MONAT_PERIODE).value = MonthName(Month(datumWert))
            End If
        End If
    Next r
    
End Sub

' ---------------------------------------------------------------
' 6a. Intelligente Monat/Periode-Ermittlung
'     Bestimmt den Periodennamen anhand der Faelligkeit:
'     - "monatlich"     -> Monatsname (z.B. "Januar")
'     - "quartalsweise" -> Quartal (z.B. "Q1 2026")
'     - "halbjaehrlich" -> Halbjahr (z.B. "H1 2026")
'     - "jaehrlich"     -> Jahr (z.B. "2026")
'     - sonst           -> Monatsname (Fallback)
' ---------------------------------------------------------------
Private Function ErmittleMonatPeriode(ByVal kategorie As String, _
                                       ByVal buchungsDatum As Date, _
                                       ByVal faelligkeit As String) As String
    
    Dim m As Long
    m = Month(buchungsDatum)
    
    Select Case LCase(Trim(faelligkeit))
        Case "quartalsweise", "quartal"
            Dim q As Long
            q = Int((m - 1) / 3) + 1
            ErmittleMonatPeriode = "Q" & q & " " & Year(buchungsDatum)
            
        Case "halbjaehrlich", "halbjahr"
            If m <= 6 Then
                ErmittleMonatPeriode = "H1 " & Year(buchungsDatum)
            Else
                ErmittleMonatPeriode = "H2 " & Year(buchungsDatum)
            End If
            
        Case "jaehrlich", "jahr", "j?hrlich"
            ErmittleMonatPeriode = CStr(Year(buchungsDatum))
            
        Case Else
            ' "monatlich" oder unbekannt -> Monatsname
            ErmittleMonatPeriode = MonthName(m)
    End Select
    
End Function

' ---------------------------------------------------------------
' 6b. Faelligkeit aus Kategorie-Tabelle (Spalte O) holen
' ---------------------------------------------------------------
Private Function HoleFaelligkeitFuerKategorie(ByVal wsDaten As Worksheet, _
                                               ByVal kategorie As String) As String
    Dim lastRow As Long
    Dim r As Long
    
    lastRow = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    
    For r = DATA_START_ROW To lastRow
        If Trim(wsDaten.Cells(r, DATA_CAT_COL_KATEGORIE).value) = kategorie Then
            HoleFaelligkeitFuerKategorie = LCase(Trim(wsDaten.Cells(r, DATA_CAT_COL_FAELLIGKEIT).value))
            Exit Function
        End If
    Next r
    
    HoleFaelligkeitFuerKategorie = "monatlich"
End Function

' ===============================================================
' 7. IMPORT REPORT LISTBOX (ACTIVEX STEUERELEMENT)
'    -----------------------------------------------
'    Architektur:
'    - ActiveX ListBox "lst_ImportReport" auf Bankkonto-Blatt
'    - Speicher: Daten!Y500 (eine einzige Zelle, serialisiert
'      mit "||" als Trennzeichen zwischen Zeilen)
'    - Befuellung: .Clear / .AddItem (ActiveX-Methoden)
'    - Hintergrundfarbe: .BackColor direkt auf der ListBox
'    - Pro Import-Vorgang: 5 Zeilen (Datum, X/Y, Dupes, Fehler, ----)
'    - Max 100 Bloecke = 500 Zeilen Historie
'    - WICHTIG: EnableEvents=False beim Schreiben in Daten!Y500
'      um Worksheet_Change-Kaskade zu verhindern
'    - WICHTIG: Position/Groesse werden VOR .Clear gesichert
'      und NACH .AddItem wiederhergestellt, da ActiveX-ListBox
'      .AddItem die OLE-Container-Groesse veraendern kann.
'      Der Designer bestimmt die Ausgangsgroesse.
' ===============================================================

' ---------------------------------------------------------------
' 7a. Initialize: Liest Y500, befuellt ActiveX ListBox,
'     setzt Hintergrundfarbe.
'     Aufruf: Workbook_Open, Worksheet_Activate, nach Loeschen
' ---------------------------------------------------------------
Public Sub Initialize_ImportReport_ListBox()
    
    Dim wsBK As Worksheet
    Dim wsDaten As Worksheet
    Dim lb As MSForms.ListBox
    Dim oleObj As OLEObject
    Dim gespeichert As String
    Dim zeilen() As String
    Dim anzahl As Long
    Dim i As Long
    Dim savLeft As Double, savTop As Double
    Dim savWidth As Double, savHeight As Double
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsBK Is Nothing Or wsDaten Is Nothing Then Exit Sub
    
    ' OLEObject holen und Position/Groesse VORHER sichern
    On Error Resume Next
    Set oleObj = wsBK.OLEObjects(FORM_LISTBOX_NAME)
    On Error GoTo 0
    If oleObj Is Nothing Then Exit Sub
    
    savLeft = oleObj.Left
    savTop = oleObj.Top
    savWidth = oleObj.Width
    savHeight = oleObj.Height
    
    ' Placement auf freifliegend setzen
    On Error Resume Next
    oleObj.Placement = xlFreeFloating
    On Error GoTo 0
    
    ' ActiveX ListBox holen
    On Error Resume Next
    Set lb = oleObj.Object
    On Error GoTo 0
    If lb Is Nothing Then Exit Sub
    
    ' ListBox leeren
    lb.Clear
    
    ' Gespeichertes Protokoll aus Y500 lesen
    gespeichert = CStr(wsDaten.Cells(PROTO_ZEILE, PROTO_SPALTE).value)
    
    If gespeichert = "" Or gespeichert = "0" Then
        ' Kein Protokoll vorhanden - Standardtext
        lb.AddItem "Kein Status Report"
        lb.AddItem "vorhanden."
        lb.BackColor = LB_COLOR_WEISS
    Else
        ' Protokoll-Zeilen aus Y500 deserialisieren und einfuegen
        zeilen = Split(gespeichert, PROTO_SEP)
        anzahl = UBound(zeilen) + 1
        If anzahl > MAX_ZEILEN Then anzahl = MAX_ZEILEN
        
        For i = 0 To anzahl - 1
            lb.AddItem zeilen(i)
        Next i
        
        ' Farbe aus juengstem Block bestimmen
        Call FaerbeListBoxAusProtokoll(lb, zeilen)
    End If
    
    ' Position und Groesse WIEDERHERSTELLEN (AddItem kann sie aendern)
    On Error Resume Next
    oleObj.Left = savLeft
    oleObj.Top = savTop
    oleObj.Width = savWidth
    oleObj.Height = savHeight
    On Error GoTo 0
    
End Sub

' ---------------------------------------------------------------
' 7b. Update: Neuen 5-Zeilen-Block OBEN einfuegen,
'     in Y500 serialisiert speichern, ListBox aktualisieren.
' ---------------------------------------------------------------
Private Sub Update_ImportReport_ListBox(ByVal totalRows As Long, ByVal imported As Long, _
                                         ByVal dupes As Long, ByVal failed As Long)
    
    Dim wsBK As Worksheet
    Dim wsDaten As Worksheet
    Dim lb As MSForms.ListBox
    Dim oleObj As OLEObject
    Dim altGespeichert As String
    Dim neuerBlock As String
    Dim gesamt As String
    Dim zeilen() As String
    Dim anzahl As Long
    Dim i As Long
    Dim eventsWaren As Boolean
    Dim savLeft As Double, savTop As Double
    Dim savWidth As Double, savHeight As Double
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsBK Is Nothing Or wsDaten Is Nothing Then Exit Sub
    
    ' OLEObject holen und Position/Groesse VORHER sichern
    On Error Resume Next
    Set oleObj = wsBK.OLEObjects(FORM_LISTBOX_NAME)
    On Error GoTo 0
    If oleObj Is Nothing Then Exit Sub
    
    savLeft = oleObj.Left
    savTop = oleObj.Top
    savWidth = oleObj.Width
    savHeight = oleObj.Height
    
    ' Placement auf freifliegend setzen
    On Error Resume Next
    oleObj.Placement = xlFreeFloating
    On Error GoTo 0
    
    ' --- 5-Zeilen-Block zusammenbauen ---
    neuerBlock = "Import: " & Format(Now, "DD.MM.YYYY  HH:MM:SS") & _
                 PROTO_SEP & _
                 imported & " / " & totalRows & " Datens" & ChrW(228) & "tze importiert" & _
                 PROTO_SEP & _
                 dupes & " Duplikate erkannt" & _
                 PROTO_SEP & _
                 failed & " Fehler" & _
                 PROTO_SEP & _
                 "--------------------------------------"
    
    ' --- WICHTIG: Events deaktivieren BEVOR in Daten geschrieben wird ---
    eventsWaren = Application.EnableEvents
    Application.EnableEvents = False
    
    ' --- Daten-Blatt entsperren ---
    On Error Resume Next
    wsDaten.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' --- Alten Inhalt aus Y500 laden ---
    altGespeichert = CStr(wsDaten.Cells(PROTO_ZEILE, PROTO_SPALTE).value)
    
    If altGespeichert = "" Or altGespeichert = "0" Then
        gesamt = neuerBlock
    Else
        gesamt = neuerBlock & PROTO_SEP & altGespeichert
    End If
    
    ' --- Auf MAX_ZEILEN begrenzen ---
    zeilen = Split(gesamt, PROTO_SEP)
    anzahl = UBound(zeilen) + 1
    If anzahl > MAX_ZEILEN Then
        gesamt = zeilen(0)
        For i = 1 To MAX_ZEILEN - 1
            gesamt = gesamt & PROTO_SEP & zeilen(i)
        Next i
        anzahl = MAX_ZEILEN
    End If
    
    ' --- In Y500 speichern (eine einzige Zelle!) ---
    wsDaten.Cells(PROTO_ZEILE, PROTO_SPALTE).value = gesamt
    
    ' --- Daten-Blatt schuetzen ---
    On Error Resume Next
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
    ' --- Events wieder herstellen ---
    Application.EnableEvents = eventsWaren
    
    ' --- ActiveX ListBox aktualisieren ---
    On Error Resume Next
    Set lb = oleObj.Object
    On Error GoTo 0
    
    If Not lb Is Nothing Then
        lb.Clear
        zeilen = Split(gesamt, PROTO_SEP)
        For i = 0 To anzahl - 1
            lb.AddItem zeilen(i)
        Next i
        
        ' Farbcodierung
        Call FaerbeListBoxNachImport(lb, imported, dupes, failed)
    End If
    
    ' Position und Groesse WIEDERHERSTELLEN (AddItem kann sie aendern)
    On Error Resume Next
    oleObj.Left = savLeft
    oleObj.Top = savTop
    oleObj.Width = savWidth
    oleObj.Height = savHeight
    On Error GoTo 0
    
End Sub

' ---------------------------------------------------------------
' 7c. Farbcodierung nach Import-Ergebnis (direkt auf ListBox)
'     GRUEN  = Alles OK (dupes = 0, failed = 0)
'     GELB   = Duplikate vorhanden (dupes > 0, failed = 0)
'     ROT    = Fehler vorhanden (failed > 0)
' ---------------------------------------------------------------
Private Sub FaerbeListBoxNachImport(ByVal lb As MSForms.ListBox, _
                                     ByVal imported As Long, _
                                     ByVal dupes As Long, _
                                     ByVal failed As Long)
    
    If failed > 0 Then
        lb.BackColor = LB_COLOR_ROT
    ElseIf dupes > 0 Then
        lb.BackColor = LB_COLOR_GELB
    Else
        lb.BackColor = LB_COLOR_GRUEN
    End If
    
End Sub

' ---------------------------------------------------------------
' 7d. Farbcodierung aus gespeichertem Protokoll bestimmen
'     Liest Index 2: "X Duplikate erkannt"
'     Liest Index 3: "X Fehler"
' ---------------------------------------------------------------
Private Sub FaerbeListBoxAusProtokoll(ByVal lb As MSForms.ListBox, ByRef zeilen() As String)
    
    Dim dupes As Long
    Dim failed As Long
    
    If UBound(zeilen) < 3 Then
        lb.BackColor = LB_COLOR_WEISS
        Exit Sub
    End If
    
    dupes = ExtrahiereZahl(CStr(zeilen(2)))
    failed = ExtrahiereZahl(CStr(zeilen(3)))
    
    If failed > 0 Then
        lb.BackColor = LB_COLOR_ROT
    ElseIf dupes > 0 Then
        lb.BackColor = LB_COLOR_GELB
    Else
        lb.BackColor = LB_COLOR_GRUEN
    End If
    
End Sub

' ---------------------------------------------------------------
' 7e. Zahl am Anfang eines Strings extrahieren
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




'--- Ende Teil 2 von 3 ---
'--- Anfang Teil 3 von 3 ---




' ===============================================================
' 8. HILFSFUNKTIONEN
' ===============================================================

' ===============================================================
' Stellt die Formeln auf dem Bankkonto-Blatt wieder her,
' die durch ClearContents oder Import verloren gehen koennen.
' Betrifft: C3, E8-E14, E16-E21, E23
' WICHTIG: Formeln werden 1:1 als FormulaLocal gesetzt!
' ===============================================================
Private Sub StelleFormelnWiederHer(ByVal ws As Worksheet)
    
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
        "Bankkonto!$A$28:$A$5000;""<="" & DATUM(Startmen" & ChrW(252) & "!$F$1;Daten!$AE$4+1;0));""TT.MM.JJJJ"")))"
    
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

' ---------------------------------------------------------------
' 8b. Alle Bankkontozeilen loeschen
' ---------------------------------------------------------------
Public Sub LoescheAlleBankkontoZeilen()
    
    Dim ws As Worksheet
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim antwort As VbMsgBoxResult
    Dim eventsWaren As Boolean
    
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
    
    ' Formeln wiederherstellen (wurden durch ClearContents geloescht)
    Call StelleFormelnWiederHer(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' Protokoll-Speicher leeren (Events aus!)
    eventsWaren = Application.EnableEvents
    Application.EnableEvents = False
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    If Not wsDaten Is Nothing Then
        wsDaten.Unprotect PASSWORD:=PASSWORD
        wsDaten.Cells(PROTO_ZEILE, PROTO_SPALTE).ClearContents
        wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    On Error GoTo 0
    
    Application.EnableEvents = eventsWaren
    
    Call Initialize_ImportReport_ListBox
    
    MsgBox "Alle Daten wurden geloescht.", vbInformation
    
End Sub

' ---------------------------------------------------------------
' 8c. Formatierung Bankkonto aktualisieren
' ---------------------------------------------------------------
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


