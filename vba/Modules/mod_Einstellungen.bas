Attribute VB_Name = "mod_Einstellungen"
Option Explicit

' ===============================================================
' MODUL: mod_Einstellungen
' VERSION: 2.1 - 10.02.2026
' ZWECK: Formatierung, DropDowns, Schutz/Entsperrung für
'        die Zahlungstermin-Tabelle auf Blatt Einstellungen
'        (Spalten B-I, ab Zeile 4, Header Zeile 3)
' ÄNDERUNG v2.1:
'   - FIX: Hilfsspalte für Fallback-DropDown von
'     Einstellungen!A nach Daten!BA (Spalte 53) verschoben.
'     Spalte A auf Einstellungen wird NICHT mehr beschrieben
'     oder ausgeblendet.
'   - Daten!BA wird von BlendeDatenSpaltenAus mit ausgeblendet.
' ÄNDERUNG v2.0:
'   - FIX: DropDown Spalte B bei >255 Zeichen Fallback
'     auf Hilfsspalte statt Daten!J-Referenz.
' ÄNDERUNG v1.9:
'   - Spalte B DropDown: Nur noch NICHT verwendete Kategorien
'     aus Daten!J anbieten (dedupliziert, case-insensitive).
' ÄNDERUNG v1.8:
'   - Debug-Prozedur DebugDropDownLogik hinzugefügt
'   - Validation.Add mit explizitem Fehler-Check
'   - HoleAlleKategorien als Public für Debug-Zugriff
' ===============================================================

Private Const ZEBRA_COLOR_1 As Long = &HFFFFFF  ' Weiß
Private Const ZEBRA_COLOR_2 As Long = &HDEE5E3  ' Hellgrau


' ===============================================================
' DEBUG: DropDown-Logik Schritt für Schritt prüfen
' Aufruf: Im VBA-Editor über F5 oder im Direktfenster:
'         Call DebugDropDownLogik
' Gibt eine MsgBox mit dem vollständigen Ergebnis aus.
' ===============================================================
Public Sub DebugDropDownLogik()

    Dim ws As Worksheet
    Dim wsDaten As Worksheet
    Dim msg As String
    Dim r As Long
    Dim lastRow As Long
    Dim kat As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Blatt '" & WS_EINSTELLUNGEN & "' nicht gefunden!", vbCritical
        Exit Sub
    End If
    If wsDaten Is Nothing Then
        MsgBox "Blatt '" & WS_DATEN & "' nicht gefunden!", vbCritical
        Exit Sub
    End If
    
    msg = "=== DEBUG DropDown Spalte B ===" & vbLf & vbLf
    
    ' -----------------------------------------------------------
    ' SCHRITT 1: Konstanten prüfen
    ' -----------------------------------------------------------
    msg = msg & "--- KONSTANTEN ---" & vbLf
    msg = msg & "WS_DATEN = """ & WS_DATEN & """" & vbLf
    msg = msg & "WS_EINSTELLUNGEN = """ & WS_EINSTELLUNGEN & """" & vbLf
    msg = msg & "DATA_CAT_COL_KATEGORIE = " & DATA_CAT_COL_KATEGORIE & " (Spalte " & _
          Split(ws.Cells(1, DATA_CAT_COL_KATEGORIE).Address, "$")(1) & ")" & vbLf
    msg = msg & "DATA_START_ROW = " & DATA_START_ROW & vbLf
    msg = msg & "DATA_COL_ES_HILF = " & DATA_COL_ES_HILF & " (Spalte " & _
          Split(wsDaten.Cells(1, DATA_COL_ES_HILF).Address, "$")(1) & ")" & vbLf
    msg = msg & "ES_COL_KATEGORIE = " & ES_COL_KATEGORIE & " (Spalte " & _
          Split(ws.Cells(1, ES_COL_KATEGORIE).Address, "$")(1) & ")" & vbLf
    msg = msg & "ES_START_ROW = " & ES_START_ROW & vbLf
    msg = msg & vbLf
    
    ' -----------------------------------------------------------
    ' SCHRITT 2: Alle Kategorien aus Daten!J lesen
    ' -----------------------------------------------------------
    Dim alleKat As Object
    Set alleKat = HoleAlleKategorien()
    
    Dim datenLastRow As Long
    datenLastRow = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    
    msg = msg & "--- ALLE KATEGORIEN aus Daten!J (Zeile " & DATA_START_ROW & _
          " bis " & datenLastRow & ") ---" & vbLf
    msg = msg & "Anzahl (dedupliziert): " & alleKat.count & vbLf
    
    Dim k As Variant
    Dim idx As Long
    idx = 0
    For Each k In alleKat.keys
        idx = idx + 1
        msg = msg & "  " & idx & ". """ & CStr(k) & """" & vbLf
    Next k
    msg = msg & vbLf
    
    ' -----------------------------------------------------------
    ' SCHRITT 3: Verwendete Kategorien in Einstellungen!B
    ' -----------------------------------------------------------
    lastRow = ws.Cells(ws.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    If lastRow < ES_START_ROW Then lastRow = ES_START_ROW - 1
    
    msg = msg & "--- VERWENDETE in Einstellungen!B (Zeile " & ES_START_ROW & _
          " bis " & lastRow & ") ---" & vbLf
    
    Dim verwendete As Object
    Set verwendete = CreateObject("Scripting.Dictionary")
    verwendete.CompareMode = vbTextCompare
    
    For r = ES_START_ROW To lastRow
        kat = Trim(CStr(ws.Cells(r, ES_COL_KATEGORIE).value))
        If kat <> "" Then
            msg = msg & "  Zeile " & r & ": """ & kat & """ (Len=" & Len(kat) & ")"
            ' Unsichtbare Zeichen prüfen
            Dim ascDump As String
            ascDump = ""
            Dim ci As Long
            For ci = 1 To Len(kat)
                ascDump = ascDump & Asc(Mid(kat, ci, 1)) & " "
            Next ci
            msg = msg & " [ASC: " & Trim(ascDump) & "]" & vbLf
            
            If Not verwendete.Exists(kat) Then
                verwendete.Add kat, r
            End If
        Else
            msg = msg & "  Zeile " & r & ": (leer)" & vbLf
        End If
    Next r
    msg = msg & "Anzahl verwendete: " & verwendete.count & vbLf & vbLf
    
    ' -----------------------------------------------------------
    ' SCHRITT 4: Verfügbare = Alle MINUS Verwendete
    ' -----------------------------------------------------------
    Dim verfuegbar As Object
    Set verfuegbar = CreateObject("Scripting.Dictionary")
    verfuegbar.CompareMode = vbTextCompare
    
    For Each k In alleKat.keys
        If Not verwendete.Exists(CStr(k)) Then
            verfuegbar.Add CStr(k), True
        Else
            msg = msg & "  ENTFERNT (bereits verwendet): """ & CStr(k) & """" & vbLf
        End If
    Next k
    
    msg = msg & vbLf & "--- VERFÜGBARE Kategorien (für leere Zeilen) ---" & vbLf
    msg = msg & "Anzahl: " & verfuegbar.count & vbLf
    
    Dim basisListe As String
    basisListe = ""
    If verfuegbar.count > 0 Then
        basisListe = Join(verfuegbar.keys, ",")
        idx = 0
        For Each k In verfuegbar.keys
            idx = idx + 1
            msg = msg & "  " & idx & ". """ & CStr(k) & """" & vbLf
        Next k
    Else
        msg = msg & "  (keine — alle Kategorien sind bereits verwendet)" & vbLf
    End If
    
    msg = msg & vbLf & "basisListe Länge: " & Len(basisListe) & " Zeichen"
    If Len(basisListe) > 255 Then
        msg = msg & " >>> ACHTUNG: >255 Zeichen! Fallback auf Daten!BA!" & vbLf
    Else
        msg = msg & " (OK, <= 255)" & vbLf
    End If
    
    ' -----------------------------------------------------------
    ' SCHRITT 5: Validation der ersten Zeile prüfen
    ' -----------------------------------------------------------
    msg = msg & vbLf & "--- AKTUELLE VALIDATION in Zeile " & ES_START_ROW & " ---" & vbLf
    On Error Resume Next
    Dim valFormula As String
    valFormula = ws.Cells(ES_START_ROW, ES_COL_KATEGORIE).Validation.Formula1
    If Err.Number <> 0 Then
        msg = msg & "  Keine Validation vorhanden (Err " & Err.Number & ")" & vbLf
        Err.Clear
    Else
        msg = msg & "  Formula1 = """ & valFormula & """" & vbLf
        msg = msg & "  Länge = " & Len(valFormula) & vbLf
    End If
    On Error GoTo 0
    
    ' -----------------------------------------------------------
    ' SCHRITT 6: Nächste freie Zeile prüfen
    ' -----------------------------------------------------------
    Dim nextRow As Long
    nextRow = lastRow + 1
    If nextRow < ES_START_ROW Then nextRow = ES_START_ROW
    
    msg = msg & vbLf & "--- NÄCHSTE FREIE ZEILE: " & nextRow & " ---" & vbLf
    On Error Resume Next
    valFormula = ws.Cells(nextRow, ES_COL_KATEGORIE).Validation.Formula1
    If Err.Number <> 0 Then
        msg = msg & "  Keine Validation vorhanden (Err " & Err.Number & ")" & vbLf
        Err.Clear
    Else
        msg = msg & "  Formula1 = """ & valFormula & """" & vbLf
        msg = msg & "  Länge = " & Len(valFormula) & vbLf
    End If
    On Error GoTo 0
    
    ' -----------------------------------------------------------
    ' SCHRITT 7: Hilfsspalte BA auf Daten prüfen
    ' -----------------------------------------------------------
    msg = msg & vbLf & "--- HILFSSPALTE Daten!BA (Fallback-Quelle) ---" & vbLf
    Dim hilfLastRow As Long
    hilfLastRow = wsDaten.Cells(wsDaten.Rows.count, DATA_COL_ES_HILF).End(xlUp).Row
    If hilfLastRow < DATA_START_ROW Then
        msg = msg & "  (leer — kein Fallback aktiv)" & vbLf
    Else
        For r = DATA_START_ROW To hilfLastRow
            msg = msg & "  BA" & r & ": """ & Trim(CStr(wsDaten.Cells(r, DATA_COL_ES_HILF).value)) & """" & vbLf
        Next r
    End If
    
    ' Ausgabe
    Debug.Print msg
    MsgBox msg, vbInformation, "Debug DropDown Spalte B"
    
End Sub


' ===============================================================
' DEBUG: Validation einer bestimmten Zeile prüfen
' Aufruf: Call DebugValidationZeile(4)
' ===============================================================
Public Sub DebugValidationZeile(ByVal zeile As Long)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    
    Dim msg As String
    msg = "Validation Zeile " & zeile & ", Spalte B:" & vbLf & vbLf
    
    On Error Resume Next
    Dim valType As Long
    valType = ws.Cells(zeile, ES_COL_KATEGORIE).Validation.Type
    If Err.Number <> 0 Then
        msg = msg & "Keine Validation vorhanden!" & vbLf
        Err.Clear
    Else
        msg = msg & "Type = " & valType & vbLf
        msg = msg & "Formula1 = """ & ws.Cells(zeile, ES_COL_KATEGORIE).Validation.Formula1 & """" & vbLf
        msg = msg & "InCellDropdown = " & ws.Cells(zeile, ES_COL_KATEGORIE).Validation.InCellDropdown & vbLf
    End If
    On Error GoTo 0
    
    msg = msg & vbLf & "Zellwert = """ & Trim(CStr(ws.Cells(zeile, ES_COL_KATEGORIE).value)) & """"
    
    Debug.Print msg
    MsgBox msg, vbInformation, "Debug Validation Zeile " & zeile
    
End Sub


' ===============================================================
' DEBUG: SetzeDropDowns isoliert aufrufen und Ergebnis prüfen
' Aufruf: Call DebugSetzeDropDownsUndPruefe
' ===============================================================
Public Sub DebugSetzeDropDownsUndPruefe()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' DropDowns neu setzen
    Call SetzeDropDowns(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' Ergebnis sofort prüfen
    Dim lastRow As Long
    lastRow = LetzteZeile(ws)
    Dim nextRow As Long
    nextRow = lastRow + 1
    If nextRow < ES_START_ROW Then nextRow = ES_START_ROW
    
    Dim msg As String
    msg = "=== ERGEBNIS nach SetzeDropDowns ===" & vbLf & vbLf
    
    Dim r As Long
    For r = ES_START_ROW To nextRow
        Dim zellWert As String
        zellWert = Trim(CStr(ws.Cells(r, ES_COL_KATEGORIE).value))
        
        Dim valFormula As String
        valFormula = ""
        On Error Resume Next
        valFormula = ws.Cells(r, ES_COL_KATEGORIE).Validation.Formula1
        If Err.Number <> 0 Then
            valFormula = "(KEINE VALIDATION!)"
            Err.Clear
        End If
        On Error GoTo 0
        
        msg = msg & "Zeile " & r & ": "
        If zellWert = "" Then
            msg = msg & "(leer)"
        Else
            msg = msg & """" & zellWert & """"
        End If
        msg = msg & vbLf & "  -> Formula1 = """ & valFormula & """" & vbLf
        
        ' Prüfen ob zellWert in Formula1 doppelt vorkommt
        If zellWert <> "" And valFormula <> "(KEINE VALIDATION!)" Then
            Dim teile() As String
            teile = Split(valFormula, ",")
            Dim anzTreffer As Long
            anzTreffer = 0
            Dim t As Long
            For t = LBound(teile) To UBound(teile)
                If StrComp(Trim(teile(t)), zellWert, vbTextCompare) = 0 Then
                    anzTreffer = anzTreffer + 1
                End If
            Next t
            If anzTreffer > 1 Then
                msg = msg & "  >>> WARNUNG: Kategorie kommt " & anzTreffer & "x in der Liste vor!" & vbLf
            End If
        End If
    Next r
    
    Debug.Print msg
    MsgBox msg, vbInformation, "Debug SetzeDropDowns Ergebnis"
    
End Sub


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
' ===============================================================
Private Sub AnwendeSpaltenformate(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim endRow As Long
    
    lastRow = LetzteZeile(ws)
    endRow = lastRow + 50
    If endRow < ES_START_ROW + 50 Then endRow = ES_START_ROW + 50
    
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_KATEGORIE), _
                  ws.Cells(endRow, ES_COL_KATEGORIE))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
    
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_SOLL_BETRAG), _
                  ws.Cells(endRow, ES_COL_SOLL_BETRAG))
        .NumberFormat = "#,##0.00 " & ChrW(8364)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_SOLL_TAG), _
                  ws.Cells(endRow, ES_COL_SOLL_TAG))
        .NumberFormat = "0"". Tag"""
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_SOLL_MONATE), _
                  ws.Cells(endRow, ES_COL_SOLL_MONATE))
        .NumberFormat = "@"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_STICHTAG_FIX), _
                  ws.Cells(endRow, ES_COL_STICHTAG_FIX))
        .NumberFormat = "@"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_VORLAUF), _
                  ws.Cells(endRow, ES_COL_VORLAUF))
        .NumberFormat = "0"" Tage"""
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_NACHLAUF), _
                  ws.Cells(endRow, ES_COL_NACHLAUF))
        .NumberFormat = "0"" Tage"""
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_SAEUMNIS), _
                  ws.Cells(endRow, ES_COL_SAEUMNIS))
        .NumberFormat = "#,##0.00 " & ChrW(8364)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    
End Sub


' ===============================================================
' 6. DROPDOWN-LISTEN SETZEN
'    Spalte B: Kategorie-DropDown (nur nicht-verwendete)
'    FIX v2.1: Hilfsspalte ist jetzt Daten!BA (DATA_COL_ES_HILF)
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
    
    ' 5. Hilfsspalte BA auf Daten vorbereiten: Alte Werte löschen, dann
    '    verfügbare Kategorien hineinschreiben (für Fallback bei >255 Zeichen)
    Dim wsDaten As Worksheet
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If Not wsDaten Is Nothing Then
        On Error Resume Next
        wsDaten.Unprotect PASSWORD:=PASSWORD
        On Error GoTo 0
        
        ' Spalte BA leeren
        wsDaten.Range(wsDaten.Cells(1, DATA_COL_ES_HILF), _
                      wsDaten.Cells(wsDaten.Rows.count, DATA_COL_ES_HILF)).ClearContents
        
        ' Header setzen
        wsDaten.Cells(DATA_HEADER_ROW, DATA_COL_ES_HILF).value = "ES-Hilf"
        
        ' Verfügbare Kategorien schreiben
        Dim hilfZeile As Long
        hilfZeile = DATA_START_ROW
        For Each k In verfuegbar.keys
            wsDaten.Cells(hilfZeile, DATA_COL_ES_HILF).value = CStr(k)
            hilfZeile = hilfZeile + 1
        Next k
        
        ' Spalte BA ausblenden
        wsDaten.Columns(DATA_COL_ES_HILF).Hidden = True
        
        wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    
    Dim hilfLetzte As Long
    hilfLetzte = hilfZeile - 1
    If hilfLetzte < DATA_START_ROW Then hilfLetzte = DATA_START_ROW
    
    ' 6. Pro Zeile das DropDown setzen
    For r = ES_START_ROW To nextRow
    
        ' Variablen explizit zurücksetzen
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
                ' --- Normalfall: Kommaseparierte Liste passt in 255 Zeichen ---
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
                If Err.Number <> 0 Then
                    Debug.Print "FEHLER Validation.Add Zeile " & r & ": " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
            Else
                ' --- Fallback bei >255 Zeichen ---
                ' Referenz auf Daten!BA (enthält NUR verfügbare Kategorien)
                If eigeneKat <> "" And Not wsDaten Is Nothing Then
                    ' Eigenen Wert temporär an Position hilfLetzte+1 schreiben
                    On Error Resume Next
                    wsDaten.Unprotect PASSWORD:=PASSWORD
                    On Error GoTo 0
                    wsDaten.Cells(hilfLetzte + 1, DATA_COL_ES_HILF).value = eigeneKat
                    On Error Resume Next
                    With ws.Cells(r, ES_COL_KATEGORIE).Validation
                        .Add Type:=xlValidateList, _
                             AlertStyle:=xlValidAlertStop, _
                             Formula1:="='" & WS_DATEN & "'!$BA$" & DATA_START_ROW & ":$BA$" & (hilfLetzte + 1)
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .ShowInput = False
                        .ShowError = True
                    End With
                    If Err.Number <> 0 Then
                        Debug.Print "FEHLER Validation.Add Fallback+ Zeile " & r & ": " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo 0
                    ' Temporären Wert wieder löschen
                    wsDaten.Cells(hilfLetzte + 1, DATA_COL_ES_HILF).ClearContents
                    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
                ElseIf Not wsDaten Is Nothing Then
                    ' Leere Zeile: nur verfügbare aus Daten!BA
                    On Error Resume Next
                    With ws.Cells(r, ES_COL_KATEGORIE).Validation
                        .Add Type:=xlValidateList, _
                             AlertStyle:=xlValidAlertStop, _
                             Formula1:="='" & WS_DATEN & "'!$BA$" & DATA_START_ROW & ":$BA$" & hilfLetzte
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .ShowInput = False
                        .ShowError = True
                    End With
                    If Err.Number <> 0 Then
                        Debug.Print "FEHLER Validation.Add Fallback Zeile " & r & ": " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo 0
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
' 8. SPALTENBREITEN
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


Public Function HoleAlleKategorien() As Object
    
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

