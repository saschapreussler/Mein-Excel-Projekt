Attribute VB_Name = "mod_Einstellungen"
Option Explicit

' ===============================================================
' MODUL: mod_Einstellungen
' VERSION: 1.7 - 09.02.2026
' ZWECK: Formatierung, DropDowns, Schutz/Entsperrung für
'        die Zahlungstermin-Tabelle auf Blatt Einstellungen
'        (Spalten B-I, ab Zeile 4, Header Zeile 3)
' ÄNDERUNG v1.7:
'   - DropDown Spalte B: Variablen-Reset pro Schleifendurchlauf
'     (eigeneKat, zeilenListe) — behebt das Problem, dass
'     verwendete Kategorien weiterhin im DropDown erschienen.
'   - Fehlerbehandlung beim Validation.Add verbessert
'   - AutoFit für alle Spalten (B-I) inkl. Header
' ===============================================================

Private Const ZEBRA_COLOR_1 As Long = &HFFFFFF  ' Weiß
Private Const ZEBRA_COLOR_2 As Long = &HDEE5E3  ' Hellgrau


' ===============================================================
' 1. HAUPTPROZEDUR: Komplette Formatierung der Tabelle
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
    
    ' 3. Formatierung anwenden (Zebra + Rahmen)
    Call FormatiereTabelle(ws)
    
    ' 4. Spaltenformate und Ausrichtung
    Call AnwendeSpaltenformate(ws)
    
    ' 5. DropDown-Listen setzen
    Call SetzeDropDowns(ws)
    
    ' 6. Zellen sperren/entsperren
    Call SperreUndEntsperre(ws)
    
    ' 7. Spaltenbreiten (AutoFit für alle Spalten)
    Call SetzeSpaltenbreiten(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub


' ===============================================================
' 2. HEADER PRÜFEN (Zeile 3)
' ===============================================================
Private Sub PruefeHeader(ByVal ws As Worksheet)
    
    If Trim(ws.Cells(ES_HEADER_ROW, ES_COL_KATEGORIE).value) <> "" Then Exit Sub
    
    ws.Cells(ES_HEADER_ROW, ES_COL_KATEGORIE).value = "Referenz Kategorie (Leistungsart)"
    ws.Cells(ES_HEADER_ROW, ES_COL_SOLL_BETRAG).value = "Soll-Betrag"
    ws.Cells(ES_HEADER_ROW, ES_COL_SOLL_TAG).value = "Soll-Tag (des Monats)"
    ws.Cells(ES_HEADER_ROW, ES_COL_SOLL_MONATE).value = "Soll-Monat(e)"
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
' 3b. ALPHABETISCH SORTIEREN (Spalte B, A-Z)
'     Öffentlich aufrufbar aus Tabelle9.cls
' ===============================================================
Public Sub SortiereAlphabetisch(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    lastRow = LetzteZeile(ws)
    If lastRow < ES_START_ROW + 1 Then Exit Sub
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    Dim rngSort As Range
    Set rngSort = ws.Range(ws.Cells(ES_START_ROW, ES_COL_START), _
                           ws.Cells(lastRow, ES_COL_END))
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add2 key:=ws.Range(ws.Cells(ES_START_ROW, ES_COL_KATEGORIE), _
                                           ws.Cells(lastRow, ES_COL_KATEGORIE)), _
                             SortOn:=xlSortOnValues, Order:=xlAscending, _
                             DataOption:=xlSortNormal
    
    With ws.Sort
        .SetRange rngSort
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
End Sub


' ===============================================================
' 4. FORMATIERUNG: Zebra + Rahmen
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
    
    lastRowMax = lastRow
    For col = ES_COL_START To ES_COL_END
        colLastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
        If colLastRow > lastRowMax Then lastRowMax = colLastRow
    Next col
    
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
    
    rngTable.Interior.ColorIndex = xlNone
    rngTable.Borders.LineStyle = xlNone
    
    For r = ES_START_ROW To lastRow
        If (r - ES_START_ROW) Mod 2 = 0 Then
            ws.Range(ws.Cells(r, ES_COL_START), ws.Cells(r, ES_COL_END)).Interior.color = ZEBRA_COLOR_1
        Else
            ws.Range(ws.Cells(r, ES_COL_START), ws.Cells(r, ES_COL_END)).Interior.color = ZEBRA_COLOR_2
        End If
    Next r
    
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    rngTable.VerticalAlignment = xlCenter
    
End Sub


' ===============================================================
' 5. SPALTENFORMATE UND AUSRICHTUNG
'    Spalte D: ". Tag" (immer numerisch, 1-31)
'    Spalte E: Text (Monate), zentriert — freie Eingabe
'    Spalte F: Text (TT.MM.), zentriert — freie Eingabe
'    Spalte G: " Tage" hinter der Zahl
'    Spalte H: " Tage" hinter der Zahl
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
    
    ' Spalte D: Soll-Tag — immer numerisch mit ". Tag"
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_SOLL_TAG), _
                  ws.Cells(endRow, ES_COL_SOLL_TAG))
        .NumberFormat = "0"". Tag"""
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Spalte E: Soll-Monat(e) — Text, zentriert (freie Eingabe)
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_SOLL_MONATE), _
                  ws.Cells(endRow, ES_COL_SOLL_MONATE))
        .NumberFormat = "@"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Spalte F: Stichtag TT.MM. — Text, zentriert (freie Eingabe)
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_STICHTAG_FIX), _
                  ws.Cells(endRow, ES_COL_STICHTAG_FIX))
        .NumberFormat = "@"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Spalte G: Vorlauf " Tage" hinter der Zahl
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_VORLAUF), _
                  ws.Cells(endRow, ES_COL_VORLAUF))
        .NumberFormat = "0"" Tage"""
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Spalte H: Nachlauf " Tage" hinter der Zahl
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_NACHLAUF), _
                  ws.Cells(endRow, ES_COL_NACHLAUF))
        .NumberFormat = "0"" Tage"""
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Spalte I: Währung, rechtsbündig
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_SAEUMNIS), _
                  ws.Cells(endRow, ES_COL_SAEUMNIS))
        .NumberFormat = "#,##0.00 " & ChrW(8364)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    
End Sub


' ===============================================================
' 6. DROPDOWN-LISTEN SETZEN
'    Spalte B: Kategorie-DropDown
'              = ALLE Kategorien aus Daten!J (dedupliziert)
'                MINUS bereits in Einstellungen!B belegte
'                PLUS eigene Kategorie der jeweiligen Zeile
'    Spalte D: Tag 1-31 (DropDown)
'    Spalte E: KEIN DropDown — freie Monatseingabe
'    Spalte F: KEIN DropDown — freie Datumseingabe TT.MM.
'    Spalte G: Vorlauf 0-31 (DropDown)
'    Spalte H: Nachlauf 0-31 (DropDown)
' ===============================================================
Private Sub SetzeDropDowns(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim nextRow As Long
    Dim r As Long
    Dim tagListe As String
    Dim toleranzListe As String
    Dim eigeneKat As String
    Dim zeilenListe As String
    
    lastRow = LetzteZeile(ws)
    nextRow = lastRow + 1
    If nextRow < ES_START_ROW Then nextRow = ES_START_ROW
    
    ' ===================================================================
    ' SPALTE B: Kategorie-DropDown (pro Zeile individuell berechnet)
    ' ===================================================================
    
    ' 1. Alle Kategorien aus Daten!J holen (dedupliziert, case-insensitive)
    Dim alleKategorien As Object
    Set alleKategorien = HoleAlleKategorien()
    
    ' 2. Alle bereits in Einstellungen!B verwendeten Kategorien sammeln
    Dim verwendete As Object
    Set verwendete = CreateObject("Scripting.Dictionary")
    verwendete.CompareMode = vbTextCompare
    
    Dim tmpKat As String
    For r = ES_START_ROW To lastRow
        tmpKat = Trim(CStr(ws.Cells(r, ES_COL_KATEGORIE).value))
        If tmpKat <> "" Then
            If Not verwendete.Exists(tmpKat) Then
                verwendete.Add tmpKat, r
            End If
        End If
    Next r
    
    ' 3. Verfügbare Kategorien = Alle aus Daten!J MINUS bereits in Einstellungen!B verwendete
    Dim verfuegbar As Object
    Set verfuegbar = CreateObject("Scripting.Dictionary")
    verfuegbar.CompareMode = vbTextCompare
    
    Dim k As Variant
    For Each k In alleKategorien.keys
        If Not verwendete.Exists(CStr(k)) Then
            verfuegbar.Add CStr(k), True
        End If
    Next k
    
    ' 4. Basisliste als String (für leere Zeilen / nächste freie Zeile)
    Dim basisListe As String
    basisListe = ""
    If verfuegbar.count > 0 Then
        basisListe = Join(verfuegbar.keys, ",")
    End If
    
    ' 5. Pro Zeile das DropDown setzen
    '    WICHTIG: eigeneKat und zeilenListe MÜSSEN pro Durchlauf
    '    zurückgesetzt werden, sonst bleibt der Wert vom Vordurchlauf!
    For r = ES_START_ROW To nextRow
    
        ' Variablen explizit zurücksetzen (VBA-Dim in Schleifen resettet NICHT!)
        eigeneKat = ""
        zeilenListe = ""
        
        ' Alte Validation löschen
        On Error Resume Next
        ws.Cells(r, ES_COL_KATEGORIE).Validation.Delete
        On Error GoTo 0
        
        eigeneKat = Trim(CStr(ws.Cells(r, ES_COL_KATEGORIE).value))
        
        If eigeneKat = "" Then
            ' Leere Zeile / nächste freie: nur verfügbare (nicht verwendete) Kategorien
            zeilenListe = basisListe
        Else
            ' Belegte Zeile: eigene Kategorie steht an erster Stelle + alle verfügbaren
            If basisListe <> "" Then
                zeilenListe = eigeneKat & "," & basisListe
            Else
                zeilenListe = eigeneKat
            End If
        End If
        
        ' DropDown setzen (nur wenn Liste nicht leer)
        If zeilenListe <> "" Then
            If Len(zeilenListe) <= 255 Then
                On Error Resume Next
                With ws.Cells(r, ES_COL_KATEGORIE).Validation
                    .Add Type:=xlValidateList, _
                         AlertStyle:=xlValidAlertStop, _
                         Formula1:=zeilenListe
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = False
                    .ShowError = True
                End With
                On Error GoTo 0
            Else
                ' Fallback bei >255 Zeichen: Zellbereichsreferenz auf Daten!J
                ' (Duplikat-Schutz in Tabelle9.cls greift zusätzlich)
                Dim wsDaten As Worksheet
                Dim datenLastRow As Long
                On Error Resume Next
                Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
                On Error GoTo 0
                If Not wsDaten Is Nothing Then
                    datenLastRow = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
                    If datenLastRow >= DATA_START_ROW Then
                        On Error Resume Next
                        With ws.Cells(r, ES_COL_KATEGORIE).Validation
                            .Add Type:=xlValidateList, _
                                 AlertStyle:=xlValidAlertStop, _
                                 Formula1:="='" & WS_DATEN & "'!$J$" & DATA_START_ROW & _
                                           ":$J$" & datenLastRow
                            .IgnoreBlank = True
                            .InCellDropdown = True
                            .ShowInput = False
                            .ShowError = True
                        End With
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
    Next r
    
    ' ===================================================================
    ' SPALTE D: Tag 1-31
    ' ===================================================================
    tagListe = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31"
    
    For r = ES_START_ROW To nextRow
        On Error Resume Next
        ws.Cells(r, ES_COL_SOLL_TAG).Validation.Delete
        On Error GoTo 0
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
    
    ' ===================================================================
    ' SPALTE E: KEIN DropDown — alte Validierung explizit löschen
    ' ===================================================================
    For r = ES_START_ROW To nextRow
        On Error Resume Next
        ws.Cells(r, ES_COL_SOLL_MONATE).Validation.Delete
        On Error GoTo 0
    Next r
    
    ' ===================================================================
    ' SPALTE F: KEIN DropDown — alte Validierung explizit löschen
    ' ===================================================================
    For r = ES_START_ROW To nextRow
        On Error Resume Next
        ws.Cells(r, ES_COL_STICHTAG_FIX).Validation.Delete
        On Error GoTo 0
    Next r
    
    ' ===================================================================
    ' SPALTE G: Vorlauf 0-31
    ' ===================================================================
    toleranzListe = "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31"
    
    For r = ES_START_ROW To nextRow
        On Error Resume Next
        ws.Cells(r, ES_COL_VORLAUF).Validation.Delete
        On Error GoTo 0
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
    
    ' ===================================================================
    ' SPALTE H: Nachlauf 0-31
    ' ===================================================================
    For r = ES_START_ROW To nextRow
        On Error Resume Next
        ws.Cells(r, ES_COL_NACHLAUF).Validation.Delete
        On Error GoTo 0
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
' ===============================================================
Private Sub SperreUndEntsperre(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim nextRow As Long
    Dim lockEnd As Long
    
    lastRow = LetzteZeile(ws)
    nextRow = lastRow + 1
    If nextRow < ES_START_ROW Then nextRow = ES_START_ROW
    lockEnd = nextRow + 50
    
    ws.Cells.Locked = True
    
    If lastRow >= ES_START_ROW Then
        ws.Range(ws.Cells(ES_START_ROW, ES_COL_START), _
                 ws.Cells(lastRow, ES_COL_END)).Locked = False
    End If
    
    ws.Range(ws.Cells(nextRow, ES_COL_START), _
             ws.Cells(nextRow, ES_COL_END)).Locked = False
    
    ws.Range(ws.Cells(nextRow + 1, ES_COL_START), _
             ws.Cells(lockEnd, ES_COL_END)).Locked = True
    
End Sub


' ===============================================================
' 8. SPALTENBREITEN (AutoFit für alle Spalten B-I)
'    Berücksichtigt Header (Zeile 3) UND Datenzeilen.
'    Mindestbreiten stellen sicher, dass DropDowns
'    und Formatierungen (". Tag", " Tage", "€") lesbar bleiben.
' ===============================================================
Private Sub SetzeSpaltenbreiten(ByVal ws As Worksheet)
    
    Dim col As Long
    Dim lastRow As Long
    Dim endRow As Long
    Dim minBreite As Double
    
    lastRow = LetzteZeile(ws)
    endRow = lastRow + 1
    If endRow < ES_START_ROW Then endRow = ES_START_ROW
    
    For col = ES_COL_START To ES_COL_END
        ws.Range(ws.Cells(ES_HEADER_ROW, col), _
                 ws.Cells(endRow, col)).Columns.AutoFit
        
        Select Case col
            Case ES_COL_KATEGORIE:      minBreite = 24
            Case ES_COL_SOLL_BETRAG:    minBreite = 12
            Case ES_COL_SOLL_TAG:       minBreite = 12
            Case ES_COL_SOLL_MONATE:    minBreite = 18
            Case ES_COL_STICHTAG_FIX:   minBreite = 12
            Case ES_COL_VORLAUF:        minBreite = 12
            Case ES_COL_NACHLAUF:       minBreite = 12
            Case ES_COL_SAEUMNIS:       minBreite = 12
            Case Else:                  minBreite = 10
        End Select
        
        If ws.Columns(col).ColumnWidth < minBreite Then
            ws.Columns(col).ColumnWidth = minBreite
        End If
    Next col
    
End Sub


' ===============================================================
' 9. HILFSFUNKTIONEN
' ===============================================================

Private Function LetzteZeile(ByVal ws As Worksheet) As Long
    Dim lr As Long
    lr = ws.Cells(ws.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    If lr < ES_START_ROW Then lr = ES_START_ROW - 1
    LetzteZeile = lr
End Function


' ---------------------------------------------------------------
' Alle Kategorien aus Daten!J als Dictionary holen
' (Keys = Kategorie-Namen, dedupliziert, case-insensitive)
' Liest IMMER frisch aus dem Tabellenblatt — erfasst somit
' auch nachträglich vom Nutzer hinzugefügte Kategorien.
' ---------------------------------------------------------------
Private Function HoleAlleKategorien() As Object
    
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim kat As String
    Dim dict As Object
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then
        Set HoleAlleKategorien = dict
        Exit Function
    End If
    
    lastRow = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    
    For r = DATA_START_ROW To lastRow
        kat = Trim(CStr(wsDaten.Cells(r, DATA_CAT_COL_KATEGORIE).value))
        If kat <> "" Then
            If Not dict.Exists(kat) Then
                dict.Add kat, True
            End If
        End If
    Next r
    
    Set HoleAlleKategorien = dict
    
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

