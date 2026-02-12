Attribute VB_Name = "mod_Testing"
Option Explicit

' ***************************************************************
' MODUL: mod_Testing
' ZWECK: Umfangreiche Test-Prozedur für Mitgliederliste
' ***************************************************************

Public Sub TesteMitgliederliste_Komplett()
    
    Debug.Print "=== MITGLIEDERLISTE TEST-PROTOKOLL ==="
    Debug.Print ""
    
    ' TEST 1: Blattstruktur prüfen
    Call Test_1_BlattStruktur
    Debug.Print ""
    
    ' TEST 2: DropDown-Listen prüfen
    Call Test_2_DropdownListen
    Debug.Print ""
    
    ' TEST 3: Zebra-Formatierung prüfen
    Call Test_3_ZebraFormatierung
    Debug.Print ""
    
    ' TEST 4: Verein-Parzelle prüfen
    Call Test_4_VereinsParzelleIntakt
    Debug.Print ""
    
    ' TEST 5: Blattschutz prüfen
    Call Test_5_BlattSchutz
    Debug.Print ""
    
    ' TEST 6: Neues Mitglied anlegen
    Call Test_6_NeuesMitgliedAnlegen
    Debug.Print ""
    
    ' TEST 7: Mitglied bearbeiten
    Call Test_7_MitgliedBearbeiten
    Debug.Print ""
    
    ' TEST 8: Mitglied austritt
    Call Test_8_MitgliedAustritt
    Debug.Print ""
    
    ' TEST 9: Mitgliederhistorie
    Call Test_9_MitgliederhistorieIntakt
    Debug.Print ""
    
    ' TEST 10: Validierungslogik
    Call Test_10_ValidierungsLogik
    Debug.Print ""
    
    Debug.Print "=== TEST-PROTOKOLL ABGESCHLOSSEN ==="
    
End Sub

' ***************************************************************
' TEST 1: Blattstruktur prüfen
' ***************************************************************
Private Sub Test_1_BlattStruktur()
    
    Debug.Print "TEST 1: BLATTSTRUKTUR"
    
    On Error Resume Next
    
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    If wsM Is Nothing Then
        Debug.Print "? Mitgliederliste nicht gefunden"
        Exit Sub
    End If
    
    ' Prüfe Header
    If wsM.Cells(M_HEADER_ROW, M_COL_MEMBER_ID).value = "Member ID" Then
        Debug.Print "? Spalte A (Member ID) Header OK"
    Else
        Debug.Print "? Spalte A Header falsch oder leer"
    End If
    
    If wsM.Cells(M_HEADER_ROW, M_COL_PARZELLE).value = "Parzelle" Then
        Debug.Print "? Spalte B (Parzelle) Header OK"
    Else
        Debug.Print "? Spalte B Header falsch oder leer"
    End If
    
    If wsM.Cells(M_HEADER_ROW, M_COL_FUNKTION).value = "Funktion" Then
        Debug.Print "? Spalte O (Funktion) Header OK"
    Else
        Debug.Print "? Spalte O Header falsch oder leer"
    End If
    
    ' Prüfe Startzeile
    Dim lastRow As Long
    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
    
    Debug.Print "? Datenbereich: Zeile " & M_START_ROW & " bis " & lastRow
    Debug.Print "? Anzahl Mitglieder: " & (lastRow - M_START_ROW + 1)
    
    On Error GoTo 0
    
End Sub

' ***************************************************************
' TEST 2: DropDown-Listen prüfen
' ***************************************************************
Private Sub Test_2_DropdownListen()
    
    Debug.Print "TEST 2: DROPDOWN-LISTEN"
    
    On Error Resume Next
    
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    ' Prüfe Spalte B (Parzelle)
    If wsM.Range("B6").Validation.Type = xlValidateList Then
        Debug.Print "? Spalte B (Parzelle) hat Validierung"
    Else
        Debug.Print "? Spalte B (Parzelle) hat KEINE Validierung"
    End If
    
    ' Prüfe Spalte C (Seite)
    If wsM.Range("C6").Validation.Type = xlValidateList Then
        Debug.Print "? Spalte C (Seite) hat Validierung"
    Else
        Debug.Print "? Spalte C (Seite) hat KEINE Validierung"
    End If
    
    ' Prüfe Spalte D (Anrede)
    If wsM.Range("D6").Validation.Type = xlValidateList Then
        Debug.Print "? Spalte D (Anrede) hat Validierung"
    Else
        Debug.Print "? Spalte D (Anrede) hat KEINE Validierung"
    End If
    
    ' Prüfe Spalte O (Funktion)
    If wsM.Range("O6").Validation.Type = xlValidateList Then
        Debug.Print "? Spalte O (Funktion) hat Validierung"
    Else
        Debug.Print "? Spalte O (Funktion) hat KEINE Validierung"
    End If
    
    On Error GoTo 0
    
End Sub

' ***************************************************************
' TEST 3: Zebra-Formatierung prüfen
' ***************************************************************
Private Sub Test_3_ZebraFormatierung()
    
    Debug.Print "TEST 3: ZEBRA-FORMATIERUNG"
    
    On Error Resume Next
    
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    Dim lastRow As Long
    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
    
    ' Prüfe auf bedingte Formatierung
    Dim hasZebra As Boolean
    hasZebra = False
    
    Dim cell As Range
    Set cell = wsM.Range("A6")
    
    If cell.FormatConditions.Count > 0 Then
        hasZebra = True
        Debug.Print "? Bedingte Formatierung vorhanden (" & cell.FormatConditions.Count & " Regeln)"
    Else
        Debug.Print "? Keine bedingte Formatierung vorhanden"
    End If
    
    ' Prüfe auf Farben in Zeilen
    Dim row6Color As Long
    Dim row7Color As Long
    
    row6Color = wsM.Range("A6").Interior.color
    row7Color = wsM.Range("A7").Interior.color
    
    If row6Color <> row7Color Then
        Debug.Print "? Zebrafarben unterschiedlich (alternierend)"
    Else
        Debug.Print "? Zebrafarben gleich oder nicht sichtbar"
    End If
    
    On Error GoTo 0
    
End Sub

' ***************************************************************
' TEST 4: Verein-Parzelle intakt prüfen (nicht überschrieben)
' ***************************************************************
Private Sub Test_4_VereinsParzelleIntakt()
    
    Debug.Print "TEST 4: VEREIN-PARZELLE SCHUTZ"
    
    On Error Resume Next
    
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    Dim lRow As Long
    Dim lastRow As Long
    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_PARZELLE).End(xlUp).Row
    
    Dim vereinFound As Boolean
    vereinFound = False
    Dim vereinRow As Long
    
    For lRow = M_START_ROW To lastRow
        If Trim(wsM.Cells(lRow, M_COL_PARZELLE).value) = PARZELLE_VEREIN Then
            vereinFound = True
            vereinRow = lRow
            
            ' Prüfe dass Parzelle nicht überschrieben wurde
            Dim vereinName As String
            vereinName = Trim(wsM.Cells(lRow, M_COL_NACHNAME).value)
            
            If vereinName <> "" Then
                Debug.Print "? Verein-Parzelle existiert mit Daten (Zeile " & lRow & ")"
                Debug.Print "  Name: " & vereinName
                Debug.Print "? Verein-Parzelle ist NICHT überschrieben"
            Else
                Debug.Print "? Verein-Parzelle existiert aber ist leer (Zeile " & lRow & ")"
            End If
            Exit For
        End If
    Next lRow
    
    If Not vereinFound Then
        Debug.Print "? Verein-Parzelle nicht gefunden"
    End If
    
    On Error GoTo 0
    
End Sub
' ***************************************************************
' TEST 5: Blattschutz prüfen
' ***************************************************************
Private Sub Test_5_BlattSchutz()
    
    Debug.Print "TEST 5: BLATTSCHUTZ"
    
    On Error Resume Next
    
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    If wsM.ProtectContents Then
        Debug.Print "? Blatt ist geschützt"
        
        ' Prüfe Schutzoptionen
        Dim protection As protection
        Set protection = wsM.protection
        
        If Not protection Is Nothing Then
            If protection.AllowEditRanges.Count > 0 Then
                Debug.Print "? Bearbeitbare Bereiche definiert"
            Else
                Debug.Print "? Keine speziellen Bearbeitungsbereiche"
            End If
        End If
    Else
        Debug.Print "? Blatt ist NICHT geschützt"
    End If
    
    On Error GoTo 0
    
End Sub

' ***************************************************************
' TEST 6: Neues Mitglied anlegen (Validierung)
' ***************************************************************
Private Sub Test_6_NeuesMitgliedAnlegen()
    
    Debug.Print "TEST 6: NEUES MITGLIED ANLEGEN (VALIDIERUNG)"
    Debug.Print "? MANUELLER TEST ERFORDERLICH:"
    Debug.Print "1. Klicke 'Neues Mitglied' in frm_Mitgliederverwaltung"
    Debug.Print "2. Wähle Funktion 'Mitglied mit Pacht'"
    Debug.Print "3. Gib Name + Parzelle ein"
    Debug.Print "4. Label sollten 'Pachtbeginn' anzeigen"
    Debug.Print "5. Pachtbeginn mit aktuellem Datum vorbefüllt?"
    Debug.Print "6. Klick 'Anlegen'"
    Debug.Print ""
    Debug.Print "NACH TEST: Berichte ob Fehler auftraten"
    
End Sub

' ***************************************************************
' TEST 7: Mitglied bearbeiten
' ***************************************************************
Private Sub Test_7_MitgliedBearbeiten()
    
    Debug.Print "TEST 7: MITGLIED BEARBEITEN (VALIDIERUNG)"
    Debug.Print "? MANUELLER TEST ERFORDERLICH:"
    Debug.Print "1. Doppelklick auf Mitglied in der Liste"
    Debug.Print "2. Klick 'Bearbeiten'"
    Debug.Print "3. Ändere einen Eintrag (z.B. Telefon)"
    Debug.Print "4. Klick 'Übernehmen'"
    Debug.Print ""
    Debug.Print "NACH TEST: Berichte ob Änderung gespeichert wurde"
    
End Sub

' ***************************************************************
' TEST 8: Mitglied austritt
' ***************************************************************
Private Sub Test_8_MitgliedAustritt()
    
    Debug.Print "TEST 8: MITGLIED AUSTRITT (VALIDIERUNG)"
    Debug.Print "? MANUELLER TEST ERFORDERLICH:"
    Debug.Print "1. Öffne bestehendes Mitglied"
    Debug.Print "2. Klick 'Entfernen'"
    Debug.Print "3. Wähle 'Austritt'"
    Debug.Print "4. Bestätige Austrittsdatum"
    Debug.Print ""
    Debug.Print "NACH TEST:"
    Debug.Print "- Ist Mitglied aus Mitgliederliste verschwunden?"
    Debug.Print "- Ist Eintrag in Mitgliederhistorie?"
    Debug.Print "- Zebra-Formatierung noch OK?"
    
End Sub

' ***************************************************************
' TEST 9: Mitgliederhistorie intakt
' ***************************************************************
Private Sub Test_9_MitgliederhistorieIntakt()
    
    Debug.Print "TEST 9: MITGLIEDERHISTORIE"
    
    On Error Resume Next
    
    Dim wsH As Worksheet
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    
    If wsH Is Nothing Then
        Debug.Print "? Mitgliederhistorie Blatt nicht gefunden"
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = wsH.Cells(wsH.Rows.Count, H_COL_NACHNAME).End(xlUp).Row
    
    If lastRow >= H_START_ROW Then
        Debug.Print "? Mitgliederhistorie hat " & (lastRow - H_START_ROW + 1) & " Einträge"
        Debug.Print "? Blatt existiert und ist befüllt"
    Else
        Debug.Print "? Mitgliederhistorie ist leer"
    End If
    
    ' Prüfe Zebra-Formatierung
    If wsH.Range("A" & H_START_ROW).FormatConditions.Count > 0 Then
        Debug.Print "? Zebra-Formatierung vorhanden"
    Else
        Debug.Print "? Zebra-Formatierung fehlt"
    End If
    
    On Error GoTo 0
    
End Sub

' ***************************************************************
' TEST 10: Validierungslogik
' ***************************************************************
Private Sub Test_10_ValidierungsLogik()
    
    Debug.Print "TEST 10: VALIDIERUNGSLOGIK (MANUELL)"
    Debug.Print "? FOLGENDE SZENARIEN PRÜFEN:"
    Debug.Print ""
    Debug.Print "1. MITGLIED OHNE PACHT, KEINE PARZELLE:"
    Debug.Print "   ? Sollte erlaubt sein"
    Debug.Print ""
    Debug.Print "2. MITGLIED OHNE PACHT, FREIE PARZELLE:"
    Debug.Print "   ? Sollte FEHLER geben"
    Debug.Print ""
    Debug.Print "3. MITGLIED OHNE PACHT, PARZELLE MIT MITGLIED MIT PACHT:"
    Debug.Print "   ? Sollte erlaubt sein"
    Debug.Print ""
    Debug.Print "4. DUPLIZIERTER VORSITZENDER:"
    Debug.Print "   ? Sollte WARNUNG geben"
    Debug.Print ""
    Debug.Print "5. LABEL-CAPTIONS BEIM FUNKTIONSWECHSEL:"
    Debug.Print "   'Pachtbeginn' <-> 'Mitgliedsbeginn'"
    
End Sub


Public Sub TestClearFormats()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Mitgliederliste")
    
    Debug.Print "Vor ClearFormats..."
    ws.Range("A6:Q27").ClearFormats
    ws.Range("A6:Q27").Interior.ColorIndex = xlNone
    Debug.Print "Nach ClearFormats - schau in Excel ob A6 weiß ist"
    
End Sub

Public Sub TestZebraDebug()
    
    Dim ws As Worksheet
    Dim lRow As Long
    
    Set ws = ThisWorkbook.Worksheets("Mitgliederliste")
    
    Debug.Print "=== ZEBRA DEBUG TEST ==="
    
    ' Zeile 7 gelb färben
    ws.Range("A7:Q7").Interior.color = &HFFFF
    Debug.Print "Zeile 7 gefärbt mit &H00FFFF (Gelb)"
    
    ' Farbe auslesen
    Dim color As Long
    color = ws.Range("A7").Interior.color
    Debug.Print "Farbe in A7: " & Hex(color)
    
    ' Alle Zeilen durchgehen und Farben prüfen
    Debug.Print ""
    Debug.Print "Alle Zeilen nach Formatiere_Alle_Tabellen_Neu:"
    For lRow = 6 To 15
        color = ws.Range("A" & lRow).Interior.color
        Debug.Print "Zeile " & lRow & ": " & Hex(color)
    Next lRow
    
End Sub

Public Sub DebugZebraFormatierung()
    
    Dim ws As Worksheet
    Dim lRow As Long
    Dim fc As FormatCondition
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("Mitgliederliste")
    
    Debug.Print "=== UMFASSENDER ZEBRA-DEBUG TEST ==="
    Debug.Print ""
    
    ' 1. Blattschutz-Status
    Debug.Print "1. BLATTSCHUTZ-STATUS:"
    Debug.Print "   ProtectContents: " & ws.ProtectContents
    Debug.Print ""
    
    ' 2. Bedingte Formatierungen prüfen
    Debug.Print "2. BEDINGTE FORMATIERUNGEN im Bereich A6:Q27:"
    Dim rngCheck As Range
    Set rngCheck = ws.Range("A6:Q27")
    Debug.Print "   Anzahl FormatConditions: " & rngCheck.FormatConditions.Count
    
    If rngCheck.FormatConditions.Count > 0 Then
        For i = 1 To rngCheck.FormatConditions.Count
            Set fc = rngCheck.FormatConditions(i)
            Debug.Print "   FC " & i & ":"
            Debug.Print "      Typ: " & fc.Type
            Debug.Print "      Formel: " & fc.Formula1
            Debug.Print "      Farbe Interior: " & Hex(fc.Interior.color)
            Debug.Print "      StopIfTrue: " & fc.StopIfTrue
            Debug.Print "      Priority: " & fc.Priority
        Next i
    Else
        Debug.Print "   ??  KEINE FormatConditions gefunden!"
    End If
    Debug.Print ""
    
    ' 3. Zellfärbungen prüfen (direkte Farben)
    Debug.Print "3. DIREKTE ZELLFÄRBUNGEN (A6:Q15):"
    For lRow = 6 To 15
        Dim cellColor As Long
        cellColor = ws.Range("A" & lRow).Interior.color
        Debug.Print "   Zeile " & lRow & ": " & Hex(cellColor)
    Next lRow
    Debug.Print ""
    
    ' 4. Test: Manuelle bedingte Formatierung hinzufügen
    Debug.Print "4. TEST: Füge manuelle BF hinzu..."
    On Error Resume Next
    ws.Range("B6:B8").FormatConditions.Delete
    On Error GoTo 0
    
    ws.Range("B6:B8").FormatConditions.Add Type:=xlExpression, Formula1:="=MOD(ZEILE();2)=0"
    ws.Range("B6:B8").FormatConditions(1).Interior.color = &HFF0000  ' ROT
    Debug.Print "   ? ROT FormatCondition hinzugefügt (Zeilen 6-8, Spalte B)"
    Debug.Print "   Schau in Excel - sollten Zeilen 6 und 8 rot sein"
    Debug.Print ""
    
    ' 5. Formel-Test
    Debug.Print "5. FORMEL-TEST:"
    Dim testFormula As String
    testFormula = "=UND(NICHT(ISTLEER($E$6)); MOD(ZEILE()-6;2)=1)"
    Debug.Print "   Test-Formel: " & testFormula
    Debug.Print "   Für Zeile 6: MOD(6-6;2)=1 -> MOD(0;2)=1 -> FALSE"
    Debug.Print "   Für Zeile 7: MOD(7-6;2)=1 -> MOD(1;2)=1 -> TRUE ? (sollte gefärbt sein)"
    Debug.Print "   Für Zeile 8: MOD(8-6;2)=1 -> MOD(2;2)=1 -> FALSE"
    
End Sub
