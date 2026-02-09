Attribute VB_Name = "mod_Einstellungen"
Option Explicit

' ===============================================================
' MODUL: mod_Einstellungen
' VERSION: 1.2 - 09.02.2026
' ZWECK: Formatierung, DropDowns, Schutz/Entsperrung für
'        die Zahlungstermin-Tabelle auf Blatt Einstellungen
'        (Spalten B-H, ab Zeile 4, Header Zeile 3)
' ÄNDERUNG v1.2: Spalte D zeigt ". Tag" hinter der Zahl,
'        Spalten F und G zeigen " Tage" hinter der Zahl.
'        Über benutzerdefiniertes Zahlenformat (NumberFormat)
'        bleibt der Zellwert eine reine Zahl, sodass alle
'        Formeln und CLng()-Zugriffe weiterhin funktionieren.
' ÄNDERUNG v1.1: Formatierung exakt wie FormatiereKategorieTabelle
'            auf dem Daten-Blatt (gleiche Zebra-Farben, gleiche
'            Rahmenlogik mit .ColorIndex = xlAutomatic)
' ===============================================================

Private Const ZEBRA_COLOR_1 As Long = &HFFFFFF  ' Weiß
Private Const ZEBRA_COLOR_2 As Long = &HDEE5E3  ' Hellgrau


' ===============================================================
' 1. HAUPTPROZEDUR: Komplette Formatierung der Tabelle
'    Aufruf: Worksheet_Activate, nach Löschen, nach Einfügen
' ===============================================================
Public Sub FormatiereZahlungsterminTabelle(Optional ByVal ws As Worksheet)
    
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
    End If
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 1. Header einmalig prüfen (nur setzen wenn leer)
    Call PruefeHeader(ws)
    
    ' 2. Leerzeilen entfernen (Daten verdichten)
    Call VerdichteDaten(ws)
    
    ' 3. Formatierung anwenden (Zebra + Rahmen in einem Schritt)
    Call FormatiereTabelle(ws)
    
    ' 4. Spaltenformate und Ausrichtung
    Call AnwendeSpaltenformate(ws)
    
    ' 5. DropDown-Listen setzen
    Call SetzeDropDowns(ws)
    
    ' 6. Zellen sperren/entsperren
    Call SperreUndEntsperre(ws)
    
    ' 7. Spaltenbreiten
    Call SetzeSpaltenbreiten(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub


' ===============================================================
' 2. HEADER PRÜFEN (Zeile 3) - nur setzen wenn leer
'    Der Designer kann die Header manuell anpassen,
'    sie werden nicht bei jedem Aufruf überschrieben
' ===============================================================
Private Sub PruefeHeader(ByVal ws As Worksheet)
    
    ' Nur setzen wenn die erste Header-Zelle leer ist
    If Trim(ws.Cells(ES_HEADER_ROW, ES_COL_KATEGORIE).value) <> "" Then Exit Sub
    
    ws.Cells(ES_HEADER_ROW, ES_COL_KATEGORIE).value = "Referenz Kategorie (Leistungsart)"
    ws.Cells(ES_HEADER_ROW, ES_COL_SOLL_BETRAG).value = "Soll-Betrag"
    ws.Cells(ES_HEADER_ROW, ES_COL_SOLL_TAG).value = "Soll-Tag (des Monats)"
    ws.Cells(ES_HEADER_ROW, ES_COL_STICHTAG_FIX).value = "Soll-Stichtag (Fix) TT.MM."
    ws.Cells(ES_HEADER_ROW, ES_COL_VORLAUF).value = "Vorlauf-Toleranz (Tage)"
    ws.Cells(ES_HEADER_ROW, ES_COL_NACHLAUF).value = "Nachlauf-Toleranz (Tage)"
    ws.Cells(ES_HEADER_ROW, ES_COL_SAEUMNIS).value = "Säumnis-Gebühr"
    
    Dim rngHeader As Range
    Set rngHeader = ws.Range(ws.Cells(ES_HEADER_ROW, ES_COL_START), _
                             ws.Cells(ES_HEADER_ROW, ES_COL_END))
    
    With rngHeader
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Locked = True
    End With
    
End Sub


' ===============================================================
' 3. LEERZEILEN ENTFERNEN (Daten verdichten)
'    Exakt gleiche Logik wie VerdichteSpalteOhneLuecken
'    in mod_Formatierung
' ===============================================================
Private Sub VerdichteDaten(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim r As Long
    Dim resultCount As Long
    Dim numCols As Long
    Dim arrResult() As Variant
    Dim c As Long
    
    lastRow = LetzteZeile(ws)
    If lastRow < ES_START_ROW Then Exit Sub
    
    numCols = ES_COL_END - ES_COL_START + 1
    ReDim arrResult(1 To lastRow - ES_START_ROW + 1, 1 To numCols)
    resultCount = 0
    
    For r = ES_START_ROW To lastRow
        If Trim(ws.Cells(r, ES_COL_KATEGORIE).value) <> "" Then
            resultCount = resultCount + 1
            For c = 1 To numCols
                arrResult(resultCount, c) = ws.Cells(r, ES_COL_START + c - 1).value
            Next c
        End If
    Next r
    
    ' Bereich löschen und verdichtete Daten zurückschreiben
    ws.Range(ws.Cells(ES_START_ROW, ES_COL_START), _
             ws.Cells(lastRow, ES_COL_END)).ClearContents
    
    If resultCount > 0 Then
        For r = 1 To resultCount
            For c = 1 To numCols
                ws.Cells(ES_START_ROW + r - 1, ES_COL_START + c - 1).value = arrResult(r, c)
            Next c
        Next r
    End If
    
End Sub


' ===============================================================
' 4. FORMATIERUNG: Zebra + Rahmen
'    Exakt gleiche Logik wie FormatiereKategorieTabelle:
'    - Zuerst Bereich unterhalb bereinigen
'    - Dann belegten Bereich zurücksetzen
'    - Zebra-Farben zeilenweise
'    - Rahmen auf gesamten belegten Bereich
' ===============================================================
Private Sub FormatiereTabelle(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim lastRowMax As Long
    Dim rngTable As Range
    Dim rngLeeren As Range
    Dim r As Long
    Dim col As Long
    Dim colLastRow As Long
    Dim cleanStart As Long
    
    lastRow = LetzteZeile(ws)
    
    ' Maximale letzte Zeile über alle Spalten B-H ermitteln (für Bereinigung)
    lastRowMax = lastRow
    For col = ES_COL_START To ES_COL_END
        colLastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
        If colLastRow > lastRowMax Then lastRowMax = colLastRow
    Next col
    
    ' Bereich UNTERHALB der belegten Zeilen bereinigen (Rahmen + Farbe entfernen)
    If lastRowMax >= ES_START_ROW Then
        If lastRow < ES_START_ROW Then
            cleanStart = ES_START_ROW
        Else
            cleanStart = lastRow + 1
        End If
        
        If cleanStart <= lastRowMax + 50 Then
            Set rngLeeren = ws.Range(ws.Cells(cleanStart, ES_COL_START), _
                                     ws.Cells(lastRowMax + 50, ES_COL_END))
            rngLeeren.Interior.ColorIndex = xlNone
            rngLeeren.Borders.LineStyle = xlNone
        End If
    End If
    
    If lastRow < ES_START_ROW Then Exit Sub
    
    Set rngTable = ws.Range(ws.Cells(ES_START_ROW, ES_COL_START), _
                            ws.Cells(lastRow, ES_COL_END))
    
    ' Zuerst alles zurücksetzen im belegten Bereich
    rngTable.Interior.ColorIndex = xlNone
    rngTable.Borders.LineStyle = xlNone
    
    ' Zebra-Formatierung NUR für belegte Zeilen
    For r = ES_START_ROW To lastRow
        If (r - ES_START_ROW) Mod 2 = 0 Then
            ws.Range(ws.Cells(r, ES_COL_START), ws.Cells(r, ES_COL_END)).Interior.color = ZEBRA_COLOR_1
        Else
            ws.Range(ws.Cells(r, ES_COL_START), ws.Cells(r, ES_COL_END)).Interior.color = ZEBRA_COLOR_2
        End If
    Next r
    
    ' Rahmenlinien NUR für belegte Zeilen
    ' (exakt wie FormatiereKategorieTabelle: .ColorIndex = xlAutomatic)
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    rngTable.VerticalAlignment = xlCenter
    
End Sub


' ===============================================================
' 5. SPALTENFORMATE UND AUSRICHTUNG
'    Spalte D: Benutzerdefiniertes Format  0". Tag"
'             -> Zellwert bleibt Zahl, Anzeige z.B. "15. Tag"
'    Spalte F: Benutzerdefiniertes Format  0" Tage"
'             -> Zellwert bleibt Zahl, Anzeige z.B. "5 Tage"
'    Spalte G: Benutzerdefiniertes Format  0" Tage"
'             -> Zellwert bleibt Zahl, Anzeige z.B. "10 Tage"
'    Formeln und CLng()-Zugriffe sehen nur die reine Zahl.
' ===============================================================
Private Sub AnwendeSpaltenformate(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim endRow As Long
    
    lastRow = LetzteZeile(ws)
    endRow = lastRow + 50
    If endRow < ES_START_ROW + 50 Then endRow = ES_START_ROW + 50
    
    ' Spalte B: Linksbündig, kein Textumbruch
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_KATEGORIE), _
                  ws.Cells(endRow, ES_COL_KATEGORIE))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
    
    ' Spalte C: Währung, rechtsbündig
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_SOLL_BETRAG), _
                  ws.Cells(endRow, ES_COL_SOLL_BETRAG))
        .NumberFormat = "#,##0.00 " & ChrW(8364)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    
    ' Spalte D: ". Tag" hinter der Zahl (Wert bleibt reine Zahl!)
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_SOLL_TAG), _
                  ws.Cells(endRow, ES_COL_SOLL_TAG))
        .NumberFormat = "0"". Tag"""
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Spalte E: Text TT.MM., zentriert
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_STICHTAG_FIX), _
                  ws.Cells(endRow, ES_COL_STICHTAG_FIX))
        .NumberFormat = "@"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Spalte F: " Tage" hinter der Zahl (Wert bleibt reine Zahl!)
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_VORLAUF), _
                  ws.Cells(endRow, ES_COL_VORLAUF))
        .NumberFormat = "0"" Tage"""
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Spalte G: " Tage" hinter der Zahl (Wert bleibt reine Zahl!)
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_NACHLAUF), _
                  ws.Cells(endRow, ES_COL_NACHLAUF))
        .NumberFormat = "0"" Tage"""
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Spalte H: Währung, rechtsbündig
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_SAEUMNIS), _
                  ws.Cells(endRow, ES_COL_SAEUMNIS))
        .NumberFormat = "#,##0.00 " & ChrW(8364)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    
End Sub


' ===============================================================
' 6. DROPDOWN-LISTEN SETZEN
' ===============================================================
Private Sub SetzeDropDowns(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim NextRow As Long
    Dim r As Long
    Dim kategorienListe As String
    Dim tagListe As String
    Dim toleranzListe As String
    
    lastRow = LetzteZeile(ws)
    NextRow = lastRow + 1
    If NextRow < ES_START_ROW Then NextRow = ES_START_ROW
    
    ' --- Kategorie-Liste aus Daten!J erstellen (nicht redundant) ---
    kategorienListe = HoleKategorienAlsListe()
    
    ' --- Spalte B: Kategorie-DropDown für alle Datenzeilen + Eingabezeile ---
    For r = ES_START_ROW To NextRow
        ws.Cells(r, ES_COL_KATEGORIE).Validation.Delete
        If kategorienListe <> "" Then
            With ws.Cells(r, ES_COL_KATEGORIE).Validation
                .Add Type:=xlValidateList, _
                     AlertStyle:=xlValidAlertStop, _
                     Formula1:=kategorienListe
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = False
                .ShowError = True
            End With
        End If
    Next r
    
    ' --- Spalte D: Tag 1-31 ---
    tagListe = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31"
    
    For r = ES_START_ROW To NextRow
        ws.Cells(r, ES_COL_SOLL_TAG).Validation.Delete
        With ws.Cells(r, ES_COL_SOLL_TAG).Validation
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Formula1:=tagListe
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = True
        End With
    Next r
    
    ' --- Spalte F: Vorlauf 0-31 ---
    toleranzListe = "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31"
    
    For r = ES_START_ROW To NextRow
        ws.Cells(r, ES_COL_VORLAUF).Validation.Delete
        With ws.Cells(r, ES_COL_VORLAUF).Validation
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Formula1:=toleranzListe
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = True
        End With
    Next r
    
    ' --- Spalte G: Nachlauf 0-31 ---
    For r = ES_START_ROW To NextRow
        ws.Cells(r, ES_COL_NACHLAUF).Validation.Delete
        With ws.Cells(r, ES_COL_NACHLAUF).Validation
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Formula1:=toleranzListe
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = True
        End With
    Next r
    
End Sub


' ===============================================================
' 7. SPERREN UND ENTSPERREN
'    - Bestehende Datenzeilen B-H: entsperrt (editierbar)
'    - Genau 1 nächste freie Zeile: entsperrt (Neuanlage)
'    - Alles darunter + außerhalb: gesperrt
' ===============================================================
Private Sub SperreUndEntsperre(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim NextRow As Long
    Dim lockEnd As Long
    
    lastRow = LetzteZeile(ws)
    NextRow = lastRow + 1
    If NextRow < ES_START_ROW Then NextRow = ES_START_ROW
    lockEnd = NextRow + 50
    
    ' Gesamtes Blatt sperren
    ws.Cells.Locked = True
    
    ' Bestehende Daten B-H entsperren (editierbar)
    If lastRow >= ES_START_ROW Then
        ws.Range(ws.Cells(ES_START_ROW, ES_COL_START), _
                 ws.Cells(lastRow, ES_COL_END)).Locked = False
    End If
    
    ' Genau 1 nächste freie Zeile B-H entsperren (Neuanlage)
    ws.Range(ws.Cells(NextRow, ES_COL_START), _
             ws.Cells(NextRow, ES_COL_END)).Locked = False
    
    ' Bereich darunter explizit sperren (Sicherheitspuffer)
    ws.Range(ws.Cells(NextRow + 1, ES_COL_START), _
             ws.Cells(lockEnd, ES_COL_END)).Locked = True
    
End Sub


' ===============================================================
' 8. SPALTENBREITEN
' ===============================================================
Private Sub SetzeSpaltenbreiten(ByVal ws As Worksheet)
    
    ws.Columns(ES_COL_KATEGORIE).AutoFit              ' B: AutoFit nach Datenbreite
    If ws.Columns(ES_COL_KATEGORIE).ColumnWidth < 24 Then
        ws.Columns(ES_COL_KATEGORIE).ColumnWidth = 24 ' Mindestbreite 24
    End If
    ws.Columns(ES_COL_SOLL_BETRAG).ColumnWidth = 14   ' C
    ws.Columns(ES_COL_SOLL_TAG).ColumnWidth = 12      ' D
    ws.Columns(ES_COL_STICHTAG_FIX).ColumnWidth = 14  ' E
    ws.Columns(ES_COL_VORLAUF).ColumnWidth = 14       ' F
    ws.Columns(ES_COL_NACHLAUF).ColumnWidth = 14      ' G
    ws.Columns(ES_COL_SAEUMNIS).ColumnWidth = 14      ' H
    
End Sub


' ===============================================================
' 9. HILFSFUNKTIONEN
' ===============================================================

' Letzte Zeile im Bereich B ermitteln
Private Function LetzteZeile(ByVal ws As Worksheet) As Long
    Dim lr As Long
    lr = ws.Cells(ws.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    If lr < ES_START_ROW Then lr = ES_START_ROW - 1
    LetzteZeile = lr
End Function


' Kategorien aus Daten!J als nicht-redundante kommaseparierte Liste
Private Function HoleKategorienAlsListe() As String
    
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim kat As String
    Dim dict As Object
    Dim result As String
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    If wsDaten Is Nothing Then
        HoleKategorienAlsListe = ""
        Exit Function
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    
    For r = DATA_START_ROW To lastRow
        kat = Trim(wsDaten.Cells(r, DATA_CAT_COL_KATEGORIE).value)
        If kat <> "" Then
            If Not dict.Exists(kat) Then
                dict.Add kat, True
            End If
        End If
    Next r
    
    If dict.count = 0 Then
        HoleKategorienAlsListe = ""
        Exit Function
    End If
    
    result = Join(dict.keys, ",")
    
    ' Excel DropDown Limit: max 255 Zeichen in Formula1
    If Len(result) > 255 Then
        HoleKategorienAlsListe = "='" & WS_DATEN & "'!$J$" & DATA_START_ROW & _
                                  ":$J$" & lastRow
    Else
        HoleKategorienAlsListe = result
    End If
    
End Function


' ===============================================================
' 10. ZEILE LÖSCHEN
' ===============================================================
Public Sub LoescheZahlungsterminZeile(ByVal ws As Worksheet, ByVal zeile As Long)
    
    If zeile < ES_START_ROW Then Exit Sub
    
    Dim lastRow As Long
    lastRow = LetzteZeile(ws)
    If zeile > lastRow Then Exit Sub
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ws.Range(ws.Cells(zeile, ES_COL_START), _
             ws.Cells(zeile, ES_COL_END)).ClearContents
    
    Call FormatiereZahlungsterminTabelle(ws)
    
End Sub

