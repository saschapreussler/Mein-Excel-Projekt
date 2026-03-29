Attribute VB_Name = "mod_Banking_Data"
Option Explicit

' ===============================================================
' MODUL: mod_Banking_Data (ORCHESTRATOR)
' VERSION: 5.0 - Modularisiert
' ?NDERUNG v5.0:
'   - Formatierung ausgelagert nach mod_Banking_Format
'   - Import-Report ausgelagert nach mod_Banking_Report
'   - Dieses Modul: Import-Logik, Pr?fungen, L?sch-/Aktualisierung
' ?NDERUNG v4.0:
'   - NEU: Schritt 7 in Importiere_Kontoauszug:
'     ?bersicht generieren nach CSV-Import (nur bei neuen Daten)
'     Aufruf: mod_Uebersicht_Generator.GeneriereUebersicht
' ?NDERUNG v3.9:
'   - Setze_Monat_Periode ENTFERNT (verschoben nach
'     mod_Zahlungspruefung.SetzeMonatPeriode)
'   - HoleFaelligkeitFuerKategorie ENTFERNT (verschoben nach
'     mod_Zahlungspruefung.HoleFaelligkeitFuerKategorie)
'   - Aufruf in Importiere_Kontoauszug ge?ndert auf
'     mod_Zahlungspruefung.SetzeMonatPeriode
' ===============================================================


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
        Call mod_Banking_Report.Initialize_ImportReport_ListBox
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
           "Bitte pr?fen Sie ob die Arbeitsmappe gesch?tzt ist.", vbCritical
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
    
    ' ============================================================
    ' v5.1: Jahrespr?fung - CSV-Daten mit Startmen?!F1 abgleichen
    ' VOR dem Eintragen ins Bankkonto pr?fen ob die Jahre ?bereinstimmen
    ' ============================================================
    Dim wsStartImport As Worksheet
    On Error Resume Next
    Set wsStartImport = ThisWorkbook.Worksheets("Startmen" & ChrW(252))
    On Error GoTo 0
    
    Dim jahrF1Import As Long
    jahrF1Import = 0
    If Not wsStartImport Is Nothing Then
        If IsNumeric(wsStartImport.Range("F1").value) Then
            jahrF1Import = CLng(wsStartImport.Range("F1").value)
        End If
    End If
    
    ' H?ufigstes Jahr in CSV ermitteln
    Dim jahrCSV As Long
    Dim jahrDict As Object
    Set jahrDict = CreateObject("Scripting.Dictionary")
    
    Dim lRowScan As Long
    For lRowScan = 2 To lastRowTemp
        If IsDate(wsTemp.Cells(lRowScan, CSV_COL_BUCHUNGSDATUM).value) Then
            Dim scanJahr As String
            scanJahr = CStr(Year(CDate(wsTemp.Cells(lRowScan, CSV_COL_BUCHUNGSDATUM).value)))
            If jahrDict.Exists(scanJahr) Then
                jahrDict(scanJahr) = jahrDict(scanJahr) + 1
            Else
                jahrDict.Add scanJahr, 1
            End If
        End If
    Next lRowScan
    
    jahrCSV = 0
    If jahrDict.count > 0 Then
        Dim maxCSVAnzahl As Long
        Dim keyCSV As Variant
        maxCSVAnzahl = 0
        For Each keyCSV In jahrDict.keys
            If jahrDict(keyCSV) > maxCSVAnzahl Then
                maxCSVAnzahl = jahrDict(keyCSV)
                jahrCSV = CLng(keyCSV)
            End If
        Next keyCSV
    End If
    Set jahrDict = Nothing
    
    ' Vergleich und Nutzer-Abfrage bei Abweichung
    If jahrF1Import > 0 And jahrCSV > 0 And jahrF1Import <> jahrCSV Then
        ' Events kurz aktivieren damit MsgBox angezeigt wird
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        
        Dim jahrAntwort As VbMsgBoxResult
        jahrAntwort = MsgBox( _
            "Das Abrechnungsjahr in Startmen" & ChrW(252) & "!F1 ist " & jahrF1Import & "," & vbLf & _
            "aber die CSV-Kontoausz" & ChrW(252) & "ge stammen " & ChrW(252) & "berwiegend aus " & jahrCSV & "." & vbLf & vbLf & _
            "Was ist korrekt?" & vbLf & vbLf & _
            "  Ja = Abrechnungsjahr auf " & jahrCSV & " anpassen und importieren" & vbLf & _
            "  Nein = Import abbrechen (Daten werden NICHT eingetragen)", _
            vbExclamation + vbYesNo, "Abrechnungsjahr - Widerspruch erkannt")
        
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        If jahrAntwort = vbYes Then
            ' Nutzer will Jahr anpassen -> Startmen?!F1 aktualisieren
            On Error Resume Next
            wsStartImport.Unprotect PASSWORD:=PASSWORD
            On Error GoTo 0
            wsStartImport.Range("F1").value = jahrCSV
            On Error Resume Next
            wsStartImport.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
            On Error GoTo 0
            Debug.Print "[Import] Abrechnungsjahr angepasst: " & jahrF1Import & " -> " & jahrCSV
        Else
            ' Nutzer will NICHT importieren -> Abbruch
            On Error Resume Next
            Application.DisplayAlerts = False
            If Not wsTemp Is Nothing Then wsTemp.Delete
            Application.DisplayAlerts = True
            Set wsTemp = Nothing
            On Error GoTo 0
            
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True
            Application.EnableEvents = True
            wsZiel.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
            
            MsgBox "Import abgebrochen." & vbLf & _
                   "Die CSV-Daten wurden NICHT in das Bankkonto eingetragen.", _
                   vbInformation, "Import abgebrochen"
            Exit Sub
        End If
    End If
    
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
    Call mod_Banking_Report.Update_ImportReport_ListBox(rowsTotalInFile, rowsProcessed, rowsIgnoredDupe, rowsFailedImport)
    
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
    Call mod_Banking_Format.Sortiere_Bankkonto_nach_Datum
    
    ' 4. Formatierungen anwenden
    Call mod_Banking_Format.Anwende_Zebra_Bankkonto(wsZiel)
    Call mod_Banking_Format.Anwende_Border_Bankkonto(wsZiel)
    Call mod_Banking_Format.Anwende_Formatierung_Bankkonto(wsZiel)
    
    Err.Clear
    On Error GoTo 0
    
    ' 5. Kategorie-Engine nur bei neuen Zeilen
    ' WICHTIG: On Error GoTo 0 MUSS vorher stehen,
    ' damit die Pipeline ihr eigenes Error-Handling nutzen kann
    ' und nicht das "On Error Resume Next" von oben erbt!
    If rowsProcessed > 0 Then Call KategorieEngine_Pipeline(wsZiel)
    
    ' 6. Monat/Periode setzen (v3.9: verschoben nach mod_Zahlungspruefung)
    On Error Resume Next
    Call mod_ZP_Periode.SetzeMonatPeriode(wsZiel)
    Err.Clear
    On Error GoTo 0
    
    ' 7. ?bersicht IMMER aktualisieren (fasst ALLE vorhandenen Daten zusammen)
    '    v4.0: NEU - ?bersichtsblatt nach jedem Import generieren
    '    v4.1: stummModus=True da Import bereits eigene Erfolgsmeldung zeigt
    '    v4.2: On Error Resume Next ENTFERNT - GeneriereUebersicht hat eigenen ErrorHandler
    '    v4.3: Bedingung rowsProcessed>0 ENTFERNT - ?bersicht zeigt ALLE Daten,
    '          nicht nur neu importierte. Auch bei 100% Duplikaten aktualisieren!
    Debug.Print "[Import] Starte " & ChrW(220) & "bersicht-Generierung..."
    Call mod_Uebersicht_Generator.GeneriereUebersicht(stummModus:=True)
    
    ' Dashboard (neues Blatt) aktualisieren
    On Error Resume Next
    Call mod_Uebersicht_Dashboard.GeneriereUebersichtNeu(stummModus:=True)
    On Error GoTo 0
    
    ' Blattschutz wird von der Pipeline selbst verwaltet (Protect am Ende).
    ' Hier nochmals sicherstellen falls Pipeline nicht lief:
    On Error Resume Next
    wsZiel.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
    ' 8. Formeln wiederherstellen (k?nnten durch Import/Sort ?berschrieben sein)
    Call mod_Banking_Format.StelleFormelnWiederHer(wsZiel)
    
    wsZiel.Activate
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' ============================================================
    ' ERWEITERTE MsgBox mit vollst?ndigen Import-Details
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
              "Datens?tze in CSV:" & vbTab & rowsTotalInFile & vbCrLf & _
              "Importiert:" & vbTab & vbTab & rowsProcessed & " / " & rowsTotalInFile & vbCrLf & _
              "Duplikate:" & vbTab & vbTab & rowsIgnoredDupe & vbCrLf & _
              "Fehler:" & vbTab & vbTab & vbTab & rowsFailedImport & vbCrLf & vbCrLf
    
    If rowsFailedImport > 0 Then
        msgText = msgText & "ACHTUNG: " & rowsFailedImport & " Zeilen konnten nicht verarbeitet werden!"
    ElseIf rowsProcessed = 0 And rowsIgnoredDupe > 0 Then
        msgText = msgText & "Alle Eintr?ge waren bereits in der Datenbank vorhanden."
    ElseIf rowsProcessed > 0 And rowsIgnoredDupe = 0 Then
        msgText = msgText & "Alle Datens?tze wurden erfolgreich importiert."
    ElseIf rowsProcessed > 0 And rowsIgnoredDupe > 0 Then
        msgText = msgText & rowsProcessed & " neue Datens?tze importiert," & vbCrLf & _
                  rowsIgnoredDupe & " Duplikate ?bersprungen."
    End If
    
    MsgBox msgText, msgIcon, msgTitle
    
    ' ============================================================
    ' ENTITYKEY-PR?FUNG: Spalte W (EntityRole) vollst?ndig?
    ' Nur pr?fen wenn tats?chlich neue Datens?tze importiert wurden
    ' ============================================================
    If rowsProcessed > 0 Then
        Call PruefeUnvollstaendigeEntityKeys
    End If
    
End Sub


' ===============================================================
' 1b. ENTITYKEY-PR?FUNG NACH IMPORT
'     Pr?ft ob alle IBANs in der EntityKey-Tabelle (Daten! R-X)
'     eine vollst?ndige Zuordnung in Spalte W (EntityRole) haben.
'     Bei fehlenden Eintr?gen: MsgBox mit Angebot zur Navigation.
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
        ' Nur Zeilen pr?fen die eine IBAN haben
        If Trim(CStr(wsDaten.Cells(r, EK_COL_IBAN).value)) <> "" Then
            ' Spalte W (EntityRole) leer?
            If Trim(CStr(wsDaten.Cells(r, EK_COL_ROLE).value)) = "" Then
                anzahlOhneRole = anzahlOhneRole + 1
                
                ' Erste leere Zeile merken
                If ersteLeereZeile = 0 Then ersteLeereZeile = r
                
                ' Maximal 5 IBANs f?r die Anzeige sammeln
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
    
    ' Keine fehlenden Eintr?ge -> nichts tun
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
              "M?chten Sie die fehlenden Angaben jetzt vervollst?ndigen?"
    
    Dim antwort As VbMsgBoxResult
    antwort = MsgBox(hinweis, vbYesNo + vbExclamation, _
                     "Unvollst?ndige IBAN-Zuordnungen")
    
    If antwort = vbYes Then
        ' Zum Daten-Blatt wechseln und erste leere Zelle in Spalte W anw?hlen
        wsDaten.Activate
        
        On Error Resume Next
        wsDaten.Unprotect PASSWORD:=PASSWORD
        On Error GoTo 0
        
        wsDaten.Cells(ersteLeereZeile, EK_COL_ROLE).Select
        
        wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    
End Sub


' ===============================================================
' 8b. Alle Bankkontozeilen l?schen
' ===============================================================
Public Sub LoescheAlleBankkontoZeilen()
    
    Dim ws As Worksheet
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim antwort As VbMsgBoxResult
    Dim eventsWaren As Boolean
    
    antwort = MsgBox("ACHTUNG: Alle Daten auf dem Bankkonto-Blatt werden gel?scht!" & vbCrLf & vbCrLf & _
                     "Fortfahren?", vbYesNo + vbCritical, "Alle Daten l?schen?")
    
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
    
    ' Formeln wiederherstellen (wurden durch ClearContents gel?scht)
    Call mod_Banking_Format.StelleFormelnWiederHer(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' Protokoll-Speicher leeren (Events aus!)
    eventsWaren = Application.EnableEvents
    Application.EnableEvents = False
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    If Not wsDaten Is Nothing Then
        wsDaten.Unprotect PASSWORD:=PASSWORD
        wsDaten.Cells(500, 25).ClearContents
        wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    On Error GoTo 0
    
    Application.EnableEvents = eventsWaren
    
    Call mod_Banking_Report.Initialize_ImportReport_ListBox
    
    MsgBox "Alle Daten wurden gel?scht.", vbInformation
    
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
    
    Call mod_Banking_Format.Anwende_Zebra_Bankkonto(ws)
    Call mod_Banking_Format.Anwende_Border_Bankkonto(ws)
    Call mod_Banking_Format.Anwende_Formatierung_Bankkonto(ws)
    
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





























































