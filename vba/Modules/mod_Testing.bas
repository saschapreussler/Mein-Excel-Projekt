Attribute VB_Name = "mod_Testing"
Option Explicit

' ***************************************************************
' MODUL: mod_Testing
' VERSION: 3.0 - 07.06.2026
'
' ZWECK: Komplettes Test-Framework fuer das Kassenbuch.
'
' Vereinigt die frueheren Module mod_Testing + mod_TestReset.
'
' SEKTIONEN:
'   1. RESET        Komplett-Reset aller Importdaten vor CSV-Tests
'   2. VALIDIERUNG  Mitgliederliste-Konsistenz-Checks
'   3. CSV-GENERATOR  Test-CSV-Dateien erzeugen (2024-2026)
'   4. STATUS       Aktuellen Testfortschritt anzeigen
'   5. DEBUG        Zebra-/Format-Debug-Helfer
'   6. HILFSFUNKTIONEN (privat)
'
' AUFRUFE:
'   Alt+F8 > TestReset_VorCSVImport
'   Alt+F8 > GeneriereTestCSVDateien
'   Alt+F8 > ZeigeTestStatus
'   Alt+F8 > TesteMitgliederliste_Komplett
' ***************************************************************

' ===============================================================
' SEKTION 1: RESET
' ===============================================================
'
' TEST-SZENARIEN (in den generierten CSVs):
'   A: Fehlende MB-Zahlung Jan 2024 -> ROT (Nutzer verneint)
'   B: Fehlende MB-Zahlung Jan 2024 -> GRUEN (Nutzer bestaetigt)
'   C: Fehlende Brauchwasser Jan 2024 -> GRUEN (bestaetigt)
'   D: Vorauszahlung MB Dez 2024 -> auto-GRUEN Jan 2025
'
Public Sub TestReset_VorCSVImport()

    Dim wsBank As Worksheet
    Dim wsUeb As Worksheet
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim eventsWaren As Boolean
    Dim antwort As VbMsgBoxResult

    ' --- Sicherheitsabfrage ---
    antwort = MsgBox("Alle importierten Kontoausz" & ChrW(252) & "ge, die " & ChrW(220) & "bersicht " & _
                     "und das Import-Protokoll werden gel" & ChrW(246) & "scht." & vbCrLf & vbCrLf & _
                     "Die Einstellungen, Mitgliederliste, Kategorie- und " & _
                     "EntityKey-Tabellen bleiben erhalten." & vbCrLf & vbCrLf & _
                     "Fortfahren?", vbYesNo + vbQuestion, "Test-Reset vor CSV-Import")

    If antwort <> vbYes Then Exit Sub

    On Error GoTo ErrorHandler

    eventsWaren = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' =============================================================
    ' 1. BANKKONTO leeren (ab Zeile 28, Spalten A-Z)
    ' =============================================================
    Set wsBank = ThisWorkbook.Worksheets(WS_BANKKONTO)
    wsBank.Unprotect PASSWORD:=PASSWORD

    If wsBank.AutoFilterMode Then wsBank.AutoFilterMode = False

    lastRow = wsBank.Cells(wsBank.Rows.count, BK_COL_DATUM).End(xlUp).Row

    If lastRow >= BK_START_ROW Then
        wsBank.Rows(BK_START_ROW & ":" & lastRow).Clear
    End If

    ' Formeln wiederherstellen (Spalte G, Zusammenfassungen)
    Call mod_Banking_Format.StelleFormelnWiederHer(wsBank)

    wsBank.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True

    Debug.Print "[TestReset] Bankkonto: " & _
        IIf(lastRow >= BK_START_ROW, (lastRow - BK_START_ROW + 1) & " Zeilen", "keine Daten") & _
        " gel" & ChrW(246) & "scht."

    ' =============================================================
    ' 2. UEBERSICHT leeren (ab Zeile 4, Spalten A-H)
    ' =============================================================
    Set wsUeb = ThisWorkbook.Worksheets(WS_UEBERSICHT())
    wsUeb.Unprotect PASSWORD:=PASSWORD

    If wsUeb.AutoFilterMode Then wsUeb.AutoFilterMode = False

    lastRow = wsUeb.Cells(wsUeb.Rows.count, 1).End(xlUp).Row

    If lastRow >= 4 Then
        wsUeb.Rows("4:" & lastRow).Clear
        ' Auch Spalte I (Summe Ist) leeren, falls Zeilen weiter reichen
        wsUeb.Range(wsUeb.Cells(4, 9), wsUeb.Cells(lastRow, 9)).Clear
    End If

    wsUeb.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True

    Debug.Print "[TestReset] " & ChrW(220) & "bersicht: " & _
        IIf(lastRow >= 4, (lastRow - 3) & " Zeilen", "keine Daten") & _
        " gel" & ChrW(246) & "scht."

    ' =============================================================
    ' 2b. VEREINSKASSE leeren (ab Zeile 27, Spalten A-T)
    ' =============================================================
    Dim wsVK As Worksheet
    On Error Resume Next
    Set wsVK = ThisWorkbook.Worksheets(WS_VEREINSKASSE)
    On Error GoTo ErrorHandler

    If Not wsVK Is Nothing Then
        wsVK.Unprotect PASSWORD:=PASSWORD
        If wsVK.AutoFilterMode Then wsVK.AutoFilterMode = False

        Dim vkLastRow As Long
        vkLastRow = wsVK.Cells(wsVK.Rows.count, VK_COL_DATUM).End(xlUp).Row
        If vkLastRow < VK_START_ROW Then
            vkLastRow = wsVK.Cells(wsVK.Rows.count, 1).End(xlUp).Row
        End If

        If vkLastRow >= VK_START_ROW Then
            wsVK.Range(wsVK.Cells(VK_START_ROW, 1), _
                       wsVK.Cells(vkLastRow, 20)).Clear
            Debug.Print "[TestReset] Vereinskasse: " & _
                (vkLastRow - VK_START_ROW + 1) & " Zeilen gel" & ChrW(246) & "scht."
        Else
            Debug.Print "[TestReset] Vereinskasse: keine Daten."
        End If

        wsVK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, _
                     AllowFiltering:=True, AllowSorting:=True
    Else
        Debug.Print "[TestReset] Vereinskasse: Blatt nicht gefunden."
    End If

    ' =============================================================
    ' 2a. DASHBOARD MITGLIEDERZAHLUNGEN loeschen (ganzes Blatt)
    ' =============================================================
    Dim wsDash As Worksheet
    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets("Dashboard Mitgliederzahlungen")
    On Error GoTo ErrorHandler

    If Not wsDash Is Nothing Then
        Application.DisplayAlerts = False
        wsDash.Delete
        Application.DisplayAlerts = True
        Debug.Print "[TestReset] Dashboard Mitgliederzahlungen gel" & ChrW(246) & "scht."
    Else
        Debug.Print "[TestReset] Dashboard Mitgliederzahlungen: nicht vorhanden."
    End If

    ' =============================================================
    ' 3. IMPORT-PROTOKOLL leeren (Daten Y500)
    ' =============================================================
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    wsDaten.Unprotect PASSWORD:=PASSWORD

    wsDaten.Range(CELL_IMPORT_PROTOKOLL).ClearContents

    Debug.Print "[TestReset] Import-Protokoll (Y500) gel" & ChrW(246) & "scht."

    ' =============================================================
    ' 4. VORJAHR-SPEICHER (optional per MsgBox)
    ' =============================================================
    Dim vorjahrLoeschen As Boolean
    vorjahrLoeschen = False
    lastRow = wsDaten.Cells(wsDaten.Rows.count, VJ_COL_DATUM).End(xlUp).Row

    If lastRow >= VJ_START_ROW Then
        Application.ScreenUpdating = True
        Dim vjAntwort As VbMsgBoxResult
        vjAntwort = MsgBox("Der Vorjahr-Speicher (Daten CA-CF) enth" & ChrW(228) & "lt " & _
                           (lastRow - VJ_START_ROW + 1) & " Zeilen." & vbCrLf & vbCrLf & _
                           "Vorjahr-Speicher ebenfalls l" & ChrW(246) & "schen?", _
                           vbYesNo + vbQuestion, "Vorjahr-Speicher")
        Application.ScreenUpdating = False

        If vjAntwort = vbYes Then
            wsDaten.Range(wsDaten.Cells(VJ_START_ROW, VJ_COL_DATUM), _
                          wsDaten.Cells(lastRow, VJ_COL_ENTITYKEY)).Clear
            vorjahrLoeschen = True
            Debug.Print "[TestReset] Vorjahr-Speicher: " & _
                (lastRow - VJ_START_ROW + 1) & " Zeilen gel" & ChrW(246) & "scht."
        Else
            Debug.Print "[TestReset] Vorjahr-Speicher: beibehalten (Benutzerauswahl)."
        End If
    Else
        Debug.Print "[TestReset] Vorjahr-Speicher: keine Daten."
    End If

    ' =============================================================
    ' 4a. ENTITYKEY-TABELLE loeschen (optional per MsgBox)
    ' =============================================================
    Dim ekLoeschen As Boolean
    ekLoeschen = False
    Dim ekLastRow As Long
    ekLastRow = wsDaten.Cells(wsDaten.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row

    If ekLastRow >= EK_START_ROW Then
        Application.ScreenUpdating = True
        Dim ekAntwort As VbMsgBoxResult
        ekAntwort = MsgBox("Die EntityKey-Tabelle (Daten R-X) enth" & ChrW(228) & "lt " & _
                           (ekLastRow - EK_START_ROW + 1) & " Eintr" & ChrW(228) & "ge." & vbCrLf & vbCrLf & _
                           "EntityKey-Tabelle ebenfalls l" & ChrW(246) & "schen?" & vbCrLf & _
                           "(Empfohlen f" & ChrW(252) & "r kompletten Neustart)", _
                           vbYesNo + vbQuestion, "EntityKey-Tabelle")
        Application.ScreenUpdating = False

        If ekAntwort = vbYes Then
            wsDaten.Range(wsDaten.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                          wsDaten.Cells(ekLastRow, EK_COL_DEBUG)).Clear
            ekLoeschen = True
            Debug.Print "[TestReset] EntityKey-Tabelle: " & _
                (ekLastRow - EK_START_ROW + 1) & " Eintr" & ChrW(228) & "ge gel" & ChrW(246) & "scht."
        Else
            Debug.Print "[TestReset] EntityKey-Tabelle: beibehalten (Benutzerauswahl)."
        End If
    Else
        Debug.Print "[TestReset] EntityKey-Tabelle: keine Daten."
    End If

    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True

    ' =============================================================
    ' 5. Import-Report ListBox aktualisieren (falls sichtbar)
    ' =============================================================
    On Error Resume Next
    Call mod_Banking_Report.Initialize_ImportReport_ListBox
    On Error GoTo ErrorHandler

    ' =============================================================
    ' FERTIG
    ' =============================================================
    Application.ScreenUpdating = True
    Application.EnableEvents = eventsWaren

    MsgBox "Test-Reset abgeschlossen." & vbCrLf & vbCrLf & _
           "Gel" & ChrW(246) & "scht:" & vbCrLf & _
           "  " & ChrW(8226) & " Bankkonto (alle Kontoausz" & ChrW(252) & "ge)" & vbCrLf & _
           "  " & ChrW(8226) & " " & ChrW(220) & "bersicht (alle Eintr" & ChrW(228) & "ge)" & vbCrLf & _
           "  " & ChrW(8226) & " Vereinskasse (Daten ab Zeile 27)" & vbCrLf & _
           "  " & ChrW(8226) & " Dashboard Mitgliederzahlungen" & vbCrLf & _
           "  " & ChrW(8226) & " Import-Protokoll (Y500)" & vbCrLf & _
           "  " & ChrW(8226) & " Vorjahr-Speicher: " & _
           IIf(vorjahrLoeschen, "gel" & ChrW(246) & "scht", "beibehalten") & vbCrLf & _
           "  " & ChrW(8226) & " EntityKey-Tabelle: " & _
           IIf(ekLoeschen, "gel" & ChrW(246) & "scht", "beibehalten") & vbCrLf & vbCrLf & _
           "N" & ChrW(228) & "chste Schritte:" & vbCrLf & _
           "1. GeneriereTestCSVDateien aufrufen (Alt+F8)" & vbCrLf & _
           "2. Test-CSVs nacheinander importieren", _
           vbInformation, "Test-Reset"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = eventsWaren
    MsgBox "Fehler beim Test-Reset:" & vbCrLf & _
           "Nr. " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' Kurzbefehl
Public Sub TestReset_Kurz()
    Call TestReset_VorCSVImport
End Sub


' ===============================================================
' SEKTION 2: VALIDIERUNG (Mitgliederliste-Konsistenz)
' ===============================================================
Public Sub TesteMitgliederliste_Komplett()

    Debug.Print "=== MITGLIEDERLISTE TEST-PROTOKOLL ==="
    Debug.Print ""

    Call Test_1_BlattStruktur:           Debug.Print ""
    Call Test_2_DropdownListen:          Debug.Print ""
    Call Test_3_ZebraFormatierung:       Debug.Print ""
    Call Test_4_VereinsParzelleIntakt:   Debug.Print ""
    Call Test_5_BlattSchutz:             Debug.Print ""
    Call Test_6_NeuesMitgliedAnlegen:    Debug.Print ""
    Call Test_7_MitgliedBearbeiten:      Debug.Print ""
    Call Test_8_MitgliedAustritt:        Debug.Print ""
    Call Test_9_MitgliederhistorieIntakt: Debug.Print ""
    Call Test_10_ValidierungsLogik:      Debug.Print ""

    Debug.Print "=== TEST-PROTOKOLL ABGESCHLOSSEN ==="

End Sub

Private Sub Test_1_BlattStruktur()

    Debug.Print "TEST 1: BLATTSTRUKTUR"

    On Error Resume Next

    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)

    If wsM Is Nothing Then
        Debug.Print "  [FEHLER] Mitgliederliste nicht gefunden"
        Exit Sub
    End If

    If wsM.Cells(M_HEADER_ROW, M_COL_MEMBER_ID).value = "Member ID" Then
        Debug.Print "  [OK] Spalte A (Member ID) Header"
    Else
        Debug.Print "  [FEHLER] Spalte A Header falsch oder leer"
    End If

    If wsM.Cells(M_HEADER_ROW, M_COL_PARZELLE).value = "Parzelle" Then
        Debug.Print "  [OK] Spalte B (Parzelle) Header"
    Else
        Debug.Print "  [FEHLER] Spalte B Header falsch oder leer"
    End If

    If wsM.Cells(M_HEADER_ROW, M_COL_FUNKTION).value = "Funktion" Then
        Debug.Print "  [OK] Spalte O (Funktion) Header"
    Else
        Debug.Print "  [FEHLER] Spalte O Header falsch oder leer"
    End If

    Dim lastRow As Long
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row

    Debug.Print "  [INFO] Datenbereich: Zeile " & M_START_ROW & " bis " & lastRow
    Debug.Print "  [INFO] Anzahl Mitglieder: " & (lastRow - M_START_ROW + 1)

    On Error GoTo 0

End Sub

Private Sub Test_2_DropdownListen()

    Debug.Print "TEST 2: DROPDOWN-LISTEN"

    On Error Resume Next

    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)

    If wsM.Range("B6").Validation.Type = xlValidateList Then
        Debug.Print "  [OK] Spalte B (Parzelle) hat Validierung"
    Else
        Debug.Print "  [FEHLER] Spalte B (Parzelle) hat KEINE Validierung"
    End If

    If wsM.Range("C6").Validation.Type = xlValidateList Then
        Debug.Print "  [OK] Spalte C (Seite) hat Validierung"
    Else
        Debug.Print "  [FEHLER] Spalte C (Seite) hat KEINE Validierung"
    End If

    If wsM.Range("D6").Validation.Type = xlValidateList Then
        Debug.Print "  [OK] Spalte D (Anrede) hat Validierung"
    Else
        Debug.Print "  [FEHLER] Spalte D (Anrede) hat KEINE Validierung"
    End If

    If wsM.Range("O6").Validation.Type = xlValidateList Then
        Debug.Print "  [OK] Spalte O (Funktion) hat Validierung"
    Else
        Debug.Print "  [FEHLER] Spalte O (Funktion) hat KEINE Validierung"
    End If

    On Error GoTo 0

End Sub

Private Sub Test_3_ZebraFormatierung()

    Debug.Print "TEST 3: ZEBRA-FORMATIERUNG"

    On Error Resume Next

    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)

    Dim cell As Range
    Set cell = wsM.Range("A6")

    If cell.FormatConditions.count > 0 Then
        Debug.Print "  [OK] Bedingte Formatierung vorhanden (" & cell.FormatConditions.count & " Regeln)"
    Else
        Debug.Print "  [FEHLER] Keine bedingte Formatierung vorhanden"
    End If

    Dim row6Color As Long
    Dim row7Color As Long
    row6Color = wsM.Range("A6").Interior.color
    row7Color = wsM.Range("A7").Interior.color

    If row6Color <> row7Color Then
        Debug.Print "  [OK] Zebrafarben unterschiedlich (alternierend)"
    Else
        Debug.Print "  [HINWEIS] Zebrafarben gleich oder nicht sichtbar"
    End If

    On Error GoTo 0

End Sub

Private Sub Test_4_VereinsParzelleIntakt()

    Debug.Print "TEST 4: VEREIN-PARZELLE SCHUTZ"

    On Error Resume Next

    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)

    Dim lRow As Long
    Dim lastRow As Long
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_PARZELLE).End(xlUp).Row

    Dim vereinFound As Boolean
    vereinFound = False

    For lRow = M_START_ROW To lastRow
        If Trim(wsM.Cells(lRow, M_COL_PARZELLE).value) = PARZELLE_VEREIN Then
            vereinFound = True

            Dim vereinName As String
            vereinName = Trim(wsM.Cells(lRow, M_COL_NACHNAME).value)

            If vereinName <> "" Then
                Debug.Print "  [OK] Verein-Parzelle existiert mit Daten (Zeile " & lRow & ")"
                Debug.Print "       Name: " & vereinName
                Debug.Print "  [OK] Verein-Parzelle ist NICHT " & ChrW(252) & "berschrieben"
            Else
                Debug.Print "  [WARNUNG] Verein-Parzelle existiert aber ist leer (Zeile " & lRow & ")"
            End If
            Exit For
        End If
    Next lRow

    If Not vereinFound Then
        Debug.Print "  [FEHLER] Verein-Parzelle nicht gefunden"
    End If

    On Error GoTo 0

End Sub

Private Sub Test_5_BlattSchutz()

    Debug.Print "TEST 5: BLATTSCHUTZ"

    On Error Resume Next

    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)

    If wsM.ProtectContents Then
        Debug.Print "  [OK] Blatt ist gesch" & ChrW(252) & "tzt"

        Dim protection As protection
        Set protection = wsM.protection

        If Not protection Is Nothing Then
            If protection.AllowEditRanges.count > 0 Then
                Debug.Print "  [OK] Bearbeitbare Bereiche definiert"
            Else
                Debug.Print "  [INFO] Keine speziellen Bearbeitungsbereiche"
            End If
        End If
    Else
        Debug.Print "  [WARNUNG] Blatt ist NICHT gesch" & ChrW(252) & "tzt"
    End If

    On Error GoTo 0

End Sub

Private Sub Test_6_NeuesMitgliedAnlegen()

    Debug.Print "TEST 6: NEUES MITGLIED ANLEGEN (manuell)"
    Debug.Print "  1. Klicke 'Neues Mitglied' in frm_Mitgliederverwaltung"
    Debug.Print "  2. W" & ChrW(228) & "hle Funktion 'Mitglied mit Pacht'"
    Debug.Print "  3. Gib Name + Parzelle ein"
    Debug.Print "  4. Label sollten 'Pachtbeginn' anzeigen"
    Debug.Print "  5. Pachtbeginn mit aktuellem Datum vorbelegt?"
    Debug.Print "  6. Klick 'Anlegen'"
    Debug.Print "  NACH TEST: Berichte ob Fehler auftraten"

End Sub

Private Sub Test_7_MitgliedBearbeiten()

    Debug.Print "TEST 7: MITGLIED BEARBEITEN (manuell)"
    Debug.Print "  1. Doppelklick auf Mitglied in der Liste"
    Debug.Print "  2. Klick 'Bearbeiten'"
    Debug.Print "  3. " & ChrW(196) & "ndere einen Eintrag (z.B. Telefon)"
    Debug.Print "  4. Klick 'Speichern'"
    Debug.Print "  NACH TEST: Berichte ob " & ChrW(196) & "nderung gespeichert wurde"

End Sub

Private Sub Test_8_MitgliedAustritt()

    Debug.Print "TEST 8: MITGLIED AUSTRITT (manuell)"
    Debug.Print "  1. " & ChrW(214) & "ffne bestehendes Mitglied"
    Debug.Print "  2. Klick 'Entfernen'"
    Debug.Print "  3. W" & ChrW(228) & "hle 'Austritt'"
    Debug.Print "  4. Gib Austrittsdatum ein"
    Debug.Print "  NACH TEST:"
    Debug.Print "  - Ist Mitglied aus Mitgliederliste entfernt"
    Debug.Print "  - Ist Eintrag in Mitgliederhistorie vorhanden"
    Debug.Print "  - Zebra-Formatierung noch intakt"

End Sub

Private Sub Test_9_MitgliederhistorieIntakt()

    Debug.Print "TEST 9: MITGLIEDERHISTORIE"

    On Error Resume Next

    Dim wsH As Worksheet
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)

    If wsH Is Nothing Then
        Debug.Print "  [FEHLER] Mitgliederhistorie Blatt nicht gefunden"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsH.Cells(wsH.Rows.count, H_COL_NACHNAME).End(xlUp).Row

    If lastRow >= H_START_ROW Then
        Debug.Print "  [OK] Mitgliederhistorie hat " & (lastRow - H_START_ROW + 1) & " Eintr" & ChrW(228) & "ge"
    Else
        Debug.Print "  [INFO] Mitgliederhistorie ist leer"
    End If

    If wsH.Range("A" & H_START_ROW).FormatConditions.count > 0 Then
        Debug.Print "  [OK] Zebra-Formatierung vorhanden"
    Else
        Debug.Print "  [HINWEIS] Zebra-Formatierung fehlt"
    End If

    On Error GoTo 0

End Sub

Private Sub Test_10_ValidierungsLogik()

    Debug.Print "TEST 10: VALIDIERUNGSLOGIK (manuell)"
    Debug.Print "  Folgende Szenarien pr" & ChrW(252) & "fen:"
    Debug.Print ""
    Debug.Print "  1. Mitglied ohne Pacht, keine Parzelle"
    Debug.Print "     -> Sollte erlaubt sein"
    Debug.Print "  2. Mitglied ohne Pacht, freie Parzelle"
    Debug.Print "     -> Sollte FEHLER geben"
    Debug.Print "  3. Mitglied ohne Pacht, Parzelle mit Mitglied mit Pacht"
    Debug.Print "     -> Sollte erlaubt sein"
    Debug.Print "  4. Duplizierter Vorsitzender"
    Debug.Print "     -> Sollte WARNUNG geben"
    Debug.Print "  5. Label-Captions beim Funktionswechsel:"
    Debug.Print "     'Pachtbeginn' <-> 'Mitgliedsbeginn'"

End Sub


' ===============================================================
' SEKTION 3: TEST-CSV GENERATOR
' ===============================================================
' Erzeugt monatliche CSV-Dateien im Sparkasse-Format.
' Liest Mitglieder und Kategorien aus der Arbeitsmappe.
' Enth" & ChrW(228) & "lt Test-Szenarien f" & ChrW(252) & "r Vorjahr-Dezember-Zahlungen.
'
' CSV-Dateien: KTO_2024_01.csv bis KTO_2026_01.csv
'   - Jede Datei enth" & ChrW(228) & "lt alle Zahlungen eines Monats
'   - Format: Sparkasse (16 Spalten, Semikolon, UTF-8)
' ===============================================================

Public Sub GeneriereTestCSVDateien()

    Dim wsMitgl As Worksheet
    Dim wsEinst As Worksheet
    Dim r As Long, m As Long, k As Long
    Dim lastRow As Long
    Dim ordnerPfad As String

    ' --- Ordnerauswahl ---
    Dim shellApp As Object
    Set shellApp = CreateObject("Shell.Application")
    Dim oFolder As Object
    Set oFolder = shellApp.BrowseForFolder(0, _
        "Ordner f" & ChrW(252) & "r Test-CSV-Dateien w" & ChrW(228) & "hlen:", 0)
    If oFolder Is Nothing Then
        MsgBox "Abgebrochen.", vbInformation
        Exit Sub
    End If
    ordnerPfad = oFolder.Self.Path
    Set oFolder = Nothing
    Set shellApp = Nothing

    ' --- Mitglieder lesen ---
    Set wsMitgl = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    lastRow = wsMitgl.Cells(wsMitgl.Rows.count, M_COL_NACHNAME).End(xlUp).Row

    If lastRow < M_START_ROW Then
        MsgBox "Keine Mitglieder in der Mitgliederliste gefunden." & vbCrLf & _
               "Bitte zuerst Mitglieder anlegen.", vbExclamation
        Exit Sub
    End If

    Dim anzMitgl As Long
    anzMitgl = 0
    Dim mNachname() As String
    Dim mVorname() As String
    Dim mParzelle() As String
    Dim mIBAN() As String

    ReDim mNachname(1 To lastRow)
    ReDim mVorname(1 To lastRow)
    ReDim mParzelle(1 To lastRow)
    ReDim mIBAN(1 To lastRow)

    Dim dictParzellen As Object
    Set dictParzellen = CreateObject("Scripting.Dictionary")

    For r = M_START_ROW To lastRow
        If Trim(CStr(wsMitgl.Cells(r, M_COL_NACHNAME).value)) <> "" And _
           Trim(CStr(wsMitgl.Cells(r, M_COL_ANREDE).value)) <> ANREDE_KGA Then
            anzMitgl = anzMitgl + 1
            mNachname(anzMitgl) = Trim(CStr(wsMitgl.Cells(r, M_COL_NACHNAME).value))
            mVorname(anzMitgl) = Trim(CStr(wsMitgl.Cells(r, M_COL_VORNAME).value))
            mParzelle(anzMitgl) = Trim(CStr(wsMitgl.Cells(r, M_COL_PARZELLE).value))
            mIBAN(anzMitgl) = GeneriereTestIBAN(anzMitgl)

            If Not dictParzellen.exists(mParzelle(anzMitgl)) Then
                dictParzellen.Add mParzelle(anzMitgl), anzMitgl
            End If
        End If
    Next r

    If anzMitgl = 0 Then
        MsgBox "Keine aktiven Mitglieder gefunden.", vbExclamation
        Exit Sub
    End If

    ' --- Kategorien aus Einstellungen lesen ---
    Set wsEinst = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    lastRow = wsEinst.Cells(wsEinst.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row

    Dim anzKat As Long
    anzKat = 0
    Dim kName() As String
    Dim kBetrag() As Double
    Dim kMonate() As String
    Dim kIstMB() As Boolean

    ReDim kName(1 To lastRow)
    ReDim kBetrag(1 To lastRow)
    ReDim kMonate(1 To lastRow)
    ReDim kIstMB(1 To lastRow)

    For r = ES_START_ROW To lastRow
        If Trim(CStr(wsEinst.Cells(r, ES_COL_KATEGORIE).value)) <> "" Then
            anzKat = anzKat + 1
            kName(anzKat) = Trim(CStr(wsEinst.Cells(r, ES_COL_KATEGORIE).value))
            kBetrag(anzKat) = wsEinst.Cells(r, ES_COL_SOLL_BETRAG).value
            kMonate(anzKat) = CStr(wsEinst.Cells(r, ES_COL_SOLL_MONATE).value)
            kIstMB(anzKat) = (InStr(1, LCase(kName(anzKat)), "mitgliedsbeitrag") > 0)
        End If
    Next r

    If anzKat = 0 Then
        MsgBox "Keine Kategorien in der Zahlungstermin-Tabelle gefunden." & vbCrLf & _
               "Bitte Einstellungen pr" & ChrW(252) & "fen.", vbExclamation
        Exit Sub
    End If

    ' --- Test-Szenarien: Parzellen identifizieren ---
    Dim uniqueParz() As Variant
    uniqueParz = dictParzellen.keys

    Dim testParz1 As String: testParz1 = ""
    Dim testParz2 As String: testParz2 = ""
    Dim testParz3 As String: testParz3 = ""

    If UBound(uniqueParz) >= 0 Then testParz1 = CStr(uniqueParz(0))
    If UBound(uniqueParz) >= 1 Then testParz2 = CStr(uniqueParz(1))
    If UBound(uniqueParz) >= 2 Then testParz3 = CStr(uniqueParz(2))

    ' Brauchwasser-Kategorie suchen (fuer Szenario C)
    Dim bwIdx As Long: bwIdx = 0
    For k = 1 To anzKat
        If InStr(1, LCase(kName(k)), "brauchwasser") > 0 Or _
           InStr(1, LCase(kName(k)), "wasser") > 0 Then
            bwIdx = k
            Exit For
        End If
    Next k

    ' --- CSV-Dateien generieren ---
    Application.ScreenUpdating = False

    Dim dateiZaehler As Long: dateiZaehler = 0
    Dim jahr As Long, monat As Long
    Dim startMonat As Long, endMonat As Long
    Dim csvInhalt As String
    Dim csvPfad As String
    Dim buchDatum As String
    Dim zeileNr As Long
    Dim skipZeile As Boolean

    For jahr = 2024 To 2026
        If jahr = 2026 Then
            startMonat = 1: endMonat = 1
        Else
            startMonat = 1: endMonat = 12
        End If

        For monat = startMonat To endMonat
            csvInhalt = CSVHeaderZeile() & vbCrLf
            zeileNr = 0
            buchDatum = Format(DateSerial(jahr, monat, 15), "DD.MM.YYYY")

            For m = 1 To anzMitgl
                For k = 1 To anzKat
                    If Not IstMonatFaellig(kMonate(k), monat) Then GoTo NaechsteKat
                    If kBetrag(k) <= 0 Then GoTo NaechsteKat

                    ' Parzelle-basiert: nur erster Mieter zahlt
                    If Not kIstMB(k) Then
                        If dictParzellen.exists(mParzelle(m)) Then
                            If CLng(dictParzellen(mParzelle(m))) <> m Then GoTo NaechsteKat
                        End If
                    End If

                    skipZeile = False

                    ' === SZENARIO A: Parzelle 2 kein MB Jan 2024 (ROT) ===
                    If testParz2 <> "" And mParzelle(m) = testParz2 Then
                        If kIstMB(k) And jahr = 2024 And monat = 1 Then skipZeile = True
                    End If

                    ' === SZENARIO B: Parzelle 3 kein MB Jan 2024 (GRUEN) ===
                    If testParz3 <> "" And mParzelle(m) = testParz3 Then
                        If kIstMB(k) And jahr = 2024 And monat = 1 Then skipZeile = True
                    End If

                    ' === SZENARIO C: Parzelle 3 kein Brauchwasser Jan 2024 ===
                    If testParz3 <> "" And mParzelle(m) = testParz3 And bwIdx > 0 Then
                        If k = bwIdx And jahr = 2024 And monat = 1 Then skipZeile = True
                    End If

                    ' === SZENARIO D/F: Parzelle 1 kein MB Jan 2025 (auto) ===
                    If testParz1 <> "" And mParzelle(m) = testParz1 Then
                        If kIstMB(k) And jahr = 2025 And monat = 1 Then skipZeile = True
                    End If

                    ' === SZENARIO E: Parzelle 1 kein Brauchwasser Jan 2025 ===
                    If testParz1 <> "" And mParzelle(m) = testParz1 And bwIdx > 0 Then
                        If k = bwIdx And jahr = 2025 And monat = 1 Then skipZeile = True
                    End If

                    If skipZeile Then GoTo NaechsteKat

                    ' === SZENARIO D: Vorauszahlung Dez 2024 fuer Jan 2025 ===
                    If testParz1 <> "" And mParzelle(m) = testParz1 Then
                        If jahr = 2024 And monat = 12 Then
                            If kIstMB(k) Or (bwIdx > 0 And k = bwIdx) Then
                                csvInhalt = csvInhalt & CSVZeile( _
                                    Format(DateSerial(2024, 12, 28), "DD.MM.YYYY"), _
                                    FormatBetragCSV(kBetrag(k)), _
                                    mNachname(m) & " " & mVorname(m), _
                                    mIBAN(m), _
                                    kName(k) & " Vorauszahlung Januar 2025 Parz " & mParzelle(m), _
                                    "GUTSCHR. UEBERW.") & vbCrLf
                                zeileNr = zeileNr + 1
                            End If
                        End If
                    End If

                    ' Regulaere Zahlung
                    csvInhalt = csvInhalt & CSVZeile( _
                        buchDatum, _
                        FormatBetragCSV(kBetrag(k)), _
                        mNachname(m) & " " & mVorname(m), _
                        mIBAN(m), _
                        kName(k) & " " & MonatsName(monat) & " " & jahr & " Parz " & mParzelle(m), _
                        "GUTSCHR. UEBERW.") & vbCrLf
                    zeileNr = zeileNr + 1
NaechsteKat:
                Next k
            Next m

            If zeileNr > 0 Then
                csvPfad = ordnerPfad & "\KTO_" & jahr & "_" & Format(monat, "00") & ".csv"
                SchreibeUTF8Datei csvPfad, csvInhalt
                dateiZaehler = dateiZaehler + 1
                Debug.Print "[TestCSV] " & csvPfad & " (" & zeileNr & " Zeilen)"
            End If
        Next monat
    Next jahr

    Application.ScreenUpdating = True

    ' --- Zusammenfassung ---
    Dim szInfo As String
    szInfo = ""
    If testParz2 <> "" Then
        szInfo = szInfo & vbCrLf & ChrW(8226) & " Szenario A (ROT): Parzelle " & testParz2 & vbCrLf & _
            "  Kein MB im Jan 2024 - Vorjahr-Dialog: Nein " & ChrW(8594) & " ROT"
    End If
    If testParz3 <> "" Then
        szInfo = szInfo & vbCrLf & ChrW(8226) & " Szenario B (GR" & ChrW(220) & "N): Parzelle " & testParz3 & vbCrLf & _
            "  Kein MB im Jan 2024 - Vorjahr-Dialog: Ja " & ChrW(8594) & " GR" & ChrW(220) & "N"
    End If
    If testParz3 <> "" And bwIdx > 0 Then
        szInfo = szInfo & vbCrLf & ChrW(8226) & " Szenario C (Brauchwasser): Parzelle " & testParz3 & vbCrLf & _
            "  Kein Brauchwasser im Jan 2024 - Vorjahr-Dialog: Ja"
    End If
    If testParz1 <> "" Then
        szInfo = szInfo & vbCrLf & ChrW(8226) & " Szenario D (Auto-Vorjahr): Parzelle " & testParz1 & vbCrLf & _
            "  Vorauszahlung Dez 2024 " & ChrW(8594) & " Jan 2025"
    End If

    MsgBox dateiZaehler & " CSV-Dateien generiert in:" & vbCrLf & _
           ordnerPfad & vbCrLf & vbCrLf & _
           "Test-Szenarien:" & szInfo & vbCrLf & vbCrLf & _
           "N" & ChrW(228) & "chste Schritte:" & vbCrLf & _
           "1. Abrechnungsjahr = 2024 auf Einstellungen setzen" & vbCrLf & _
           "2. CSV-Import auf Bankkonto starten" & vbCrLf & _
           "3. KTO_2024_01.csv zuerst importieren" & vbCrLf & _
           "4. Dann KTO_2024_02.csv usw. der Reihe nach", _
           vbInformation, "Test-CSV generiert"

End Sub


' ===============================================================
' SEKTION 4: TEST-STATUS
' ===============================================================
Public Sub ZeigeTestStatus()

    Dim wsBank As Worksheet
    Dim wsDash As Worksheet
    Dim wsDaten As Worksheet

    Set wsBank = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Dim bankZeilen As Long
    Dim lr As Long
    lr = wsBank.Cells(wsBank.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lr < BK_START_ROW Then bankZeilen = 0 Else bankZeilen = lr - BK_START_ROW + 1

    Dim wsUeb As Worksheet
    Set wsUeb = ThisWorkbook.Worksheets(WS_UEBERSICHT())
    Dim uebZeilen As Long
    lr = wsUeb.Cells(wsUeb.Rows.count, 1).End(xlUp).Row
    If lr < 4 Then uebZeilen = 0 Else uebZeilen = lr - 3

    Dim dashExistiert As Boolean
    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets("Dashboard Mitgliederzahlungen")
    dashExistiert = (Not wsDash Is Nothing)
    On Error GoTo 0

    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    Dim vjZeilen As Long
    lr = wsDaten.Cells(wsDaten.Rows.count, VJ_COL_DATUM).End(xlUp).Row
    If lr < VJ_START_ROW Then vjZeilen = 0 Else vjZeilen = lr - VJ_START_ROW + 1

    Dim ekZeilen As Long
    lr = wsDaten.Cells(wsDaten.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    If lr < EK_START_ROW Then ekZeilen = 0 Else ekZeilen = lr - EK_START_ROW + 1

    Dim abrJahr As Long
    abrJahr = HoleAbrechnungsjahr()

    MsgBox "=== TEST-STATUS ===" & vbCrLf & vbCrLf & _
           "Abrechnungsjahr: " & IIf(abrJahr > 0, CStr(abrJahr), "(nicht gesetzt)") & vbCrLf & _
           "Bankkonto: " & bankZeilen & " Zeilen" & vbCrLf & _
           ChrW(220) & "bersicht: " & uebZeilen & " Zeilen" & vbCrLf & _
           "Dashboard: " & IIf(dashExistiert, "vorhanden", "nicht vorhanden") & vbCrLf & _
           "EntityKeys: " & ekZeilen & " Eintr" & ChrW(228) & "ge" & vbCrLf & _
           "Vorjahr-Speicher: " & vjZeilen & " Zeilen", _
           vbInformation, "Test-Status"

End Sub


' ===============================================================
' SEKTION 5: DEBUG (Zebra- und Format-Helfer)
' ===============================================================
Public Sub TestClearFormats()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Mitgliederliste")

    Debug.Print "[Debug] Vor ClearFormats..."
    ws.Range("A6:Q27").ClearFormats
    ws.Range("A6:Q27").Interior.ColorIndex = xlNone
    Debug.Print "[Debug] Nach ClearFormats - schau in Excel ob A6 leer ist"

End Sub

Public Sub TestZebraDebug()

    Dim ws As Worksheet
    Dim lRow As Long

    Set ws = ThisWorkbook.Worksheets("Mitgliederliste")

    Debug.Print "=== ZEBRA DEBUG TEST ==="

    ' Zeile 7 gelb faerben
    ws.Range("A7:Q7").Interior.color = &HFFFF
    Debug.Print "Zeile 7 gef" & ChrW(228) & "rbt mit &H00FFFF (Gelb)"

    Dim color As Long
    color = ws.Range("A7").Interior.color
    Debug.Print "Farbe in A7: " & Hex(color)

    Debug.Print ""
    Debug.Print "Alle Zeilen 6-15:"
    For lRow = 6 To 15
        color = ws.Range("A" & lRow).Interior.color
        Debug.Print "  Zeile " & lRow & ": " & Hex(color)
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

    Debug.Print "1. BLATTSCHUTZ-STATUS:"
    Debug.Print "   ProtectContents: " & ws.ProtectContents
    Debug.Print ""

    Debug.Print "2. BEDINGTE FORMATIERUNGEN im Bereich A6:Q27:"
    Dim rngCheck As Range
    Set rngCheck = ws.Range("A6:Q27")
    Debug.Print "   Anzahl FormatConditions: " & rngCheck.FormatConditions.count

    If rngCheck.FormatConditions.count > 0 Then
        For i = 1 To rngCheck.FormatConditions.count
            Set fc = rngCheck.FormatConditions(i)
            Debug.Print "   FC " & i & ":"
            Debug.Print "      Typ: " & fc.Type
            Debug.Print "      Formel: " & fc.Formula1
            Debug.Print "      Farbe Interior: " & Hex(fc.Interior.color)
            Debug.Print "      StopIfTrue: " & fc.StopIfTrue
            Debug.Print "      Priority: " & fc.Priority
        Next i
    Else
        Debug.Print "   KEINE FormatConditions gefunden!"
    End If
    Debug.Print ""

    Debug.Print "3. DIREKTE Hintergrundfarben (A6:Q15):"
    For lRow = 6 To 15
        Dim cellColor As Long
        cellColor = ws.Range("A" & lRow).Interior.color
        Debug.Print "   Zeile " & lRow & ": " & Hex(cellColor)
    Next lRow
    Debug.Print ""

    Debug.Print "4. TEST: F" & ChrW(252) & "ge manuelle BF hinzu..."
    On Error Resume Next
    ws.Range("B6:B8").FormatConditions.Delete
    On Error GoTo 0

    ws.Range("B6:B8").FormatConditions.Add Type:=xlExpression, Formula1:="=MOD(ZEILE();2)=0"
    ws.Range("B6:B8").FormatConditions(1).Interior.color = &HFF0000
    Debug.Print "   ROT FormatCondition 1 (Zeilen 6-8, Spalte B)"
    Debug.Print ""

    Debug.Print "5. FORMEL-TEST:"
    Dim testFormula As String
    testFormula = "=UND(NICHT(ISTLEER($E$6)); MOD(ZEILE()-6;2)=1)"
    Debug.Print "   Test-Formel: " & testFormula
    Debug.Print "   F" & ChrW(252) & "r Zeile 6: MOD(6-6;2)=1 -> MOD(0;2)=1 -> FALSE"
    Debug.Print "   F" & ChrW(252) & "r Zeile 7: MOD(7-6;2)=1 -> MOD(1;2)=1 -> TRUE"
    Debug.Print "   F" & ChrW(252) & "r Zeile 8: MOD(8-6;2)=1 -> MOD(2;2)=1 -> FALSE"

End Sub


' ===============================================================
' SEKTION 6: HILFSFUNKTIONEN (privat, fuer CSV-Generator)
' ===============================================================

Private Function Feld(ByVal wert As String) As String
    Feld = Chr(34) & wert & Chr(34)
End Function

Private Function CSVHeaderZeile() As String
    CSVHeaderZeile = _
        Feld("Auftragskonto") & ";" & _
        Feld("Buchungstag") & ";" & _
        Feld("Valutadatum") & ";" & _
        Feld("Buchungstext") & ";" & _
        Feld("Verwendungszweck") & ";" & _
        Feld("Glaeubiger ID") & ";" & _
        Feld("Mandatsreferenz") & ";" & _
        Feld("Kundenreferenz (End-to-End)") & ";" & _
        Feld("Sammlerreferenz") & ";" & _
        Feld("Lastschrift Ursprungsbetrag") & ";" & _
        Feld("Auslagenersatz Ruecklastschrift") & ";" & _
        Feld("Beguenstigter/Zahlungspflichtiger") & ";" & _
        Feld("Kontonummer/IBAN") & ";" & _
        Feld("BIC (SWIFT-Code)") & ";" & _
        Feld("Betrag") & ";" & _
        Feld("Waehrung")
End Function

Private Function CSVZeile( _
    ByVal buchungsDatum As String, _
    ByVal betrag As String, _
    ByVal personName As String, _
    ByVal personIBAN As String, _
    ByVal verwendungszweck As String, _
    ByVal buchungstext As String) As String

    CSVZeile = _
        Feld("DE89370400440532013000") & ";" & _
        Feld(buchungsDatum) & ";" & _
        Feld(buchungsDatum) & ";" & _
        Feld(buchungstext) & ";" & _
        Feld(verwendungszweck) & ";" & _
        Feld("") & ";" & _
        Feld("") & ";" & _
        Feld("") & ";" & _
        Feld("") & ";" & _
        Feld("") & ";" & _
        Feld("") & ";" & _
        Feld(personName) & ";" & _
        Feld(personIBAN) & ";" & _
        Feld("COBADEFFXXX") & ";" & _
        Feld(betrag) & ";" & _
        Feld("EUR")
End Function

Private Function GeneriereTestIBAN(ByVal index As Long) As String
    GeneriereTestIBAN = "DE8937040044053201" & Format(3000 + index, "0000")
End Function

Private Function MonatsName(ByVal monat As Long) As String
    Select Case monat
        Case 1: MonatsName = "Januar"
        Case 2: MonatsName = "Februar"
        Case 3: MonatsName = "M" & ChrW(228) & "rz"
        Case 4: MonatsName = "April"
        Case 5: MonatsName = "Mai"
        Case 6: MonatsName = "Juni"
        Case 7: MonatsName = "Juli"
        Case 8: MonatsName = "August"
        Case 9: MonatsName = "September"
        Case 10: MonatsName = "Oktober"
        Case 11: MonatsName = "November"
        Case 12: MonatsName = "Dezember"
    End Select
End Function

Private Function IstMonatFaellig(ByVal SollMonate As String, ByVal monat As Long) As Boolean
    If Trim(SollMonate) = "" Then
        IstMonatFaellig = True
        Exit Function
    End If

    Dim teile() As String
    teile = Split(SollMonate, ",")
    Dim i As Long
    For i = LBound(teile) To UBound(teile)
        If val(Trim(teile(i))) = monat Then
            IstMonatFaellig = True
            Exit Function
        End If
    Next i
    IstMonatFaellig = False
End Function

Private Function FormatBetragCSV(ByVal betrag As Double) As String
    FormatBetragCSV = Replace(Format(betrag, "0.00"), ".", ",")
End Function

Private Sub SchreibeUTF8Datei(ByVal pfad As String, ByVal inhalt As String)
    ' UTF-8 ohne BOM schreiben
    Dim utf8Stream As Object
    Set utf8Stream = CreateObject("ADODB.Stream")
    utf8Stream.Type = 2
    utf8Stream.Charset = "UTF-8"
    utf8Stream.Open
    utf8Stream.WriteText inhalt

    ' BOM entfernen: In Binary umschalten, 3 Bytes ueberspringen
    utf8Stream.Position = 0
    utf8Stream.Type = 1
    utf8Stream.Position = 3

    Dim byteData() As Byte
    byteData = utf8Stream.Read
    utf8Stream.Close

    Dim binStream As Object
    Set binStream = CreateObject("ADODB.Stream")
    binStream.Type = 1
    binStream.Open
    binStream.Write byteData
    binStream.SaveToFile pfad, 2
    binStream.Close

    Set utf8Stream = Nothing
    Set binStream = Nothing
End Sub





















