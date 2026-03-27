Attribute VB_Name = "mod_Einstellungen_Debug"
Option Explicit

' ===============================================================
' MODUL: mod_Einstellungen_Debug
' Ausgelagert aus mod_Einstellungen
' Enth?lt: Diagnose-Prozeduren f?r DropDown-Logik
' Aufruf: Im VBA-Editor ?ber F5 oder Direktfenster
' ===============================================================


' ===============================================================
' DEBUG: DropDown-Logik Schritt f?r Schritt pr?fen
' Aufruf: Call DebugDropDownLogik
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
    
    Dim alleKat As Object
    Set alleKat = mod_Einstellungen_DropDowns.HoleAlleKategorien()
    
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
    
    msg = msg & vbLf & "--- VERF?GBARE Kategorien (f?r leere Zeilen) ---" & vbLf
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
        msg = msg & "  (keine - alle Kategorien sind bereits verwendet)" & vbLf
    End If
    
    msg = msg & vbLf & "basisListe L?nge: " & Len(basisListe) & " Zeichen"
    If Len(basisListe) > 255 Then
        msg = msg & " >>> ACHTUNG: >255 Zeichen! Fallback auf Daten!BA!" & vbLf
    Else
        msg = msg & " (OK, <= 255)" & vbLf
    End If
    
    msg = msg & vbLf & "--- AKTUELLE VALIDATION in Zeile " & ES_START_ROW & " ---" & vbLf
    On Error Resume Next
    Dim valFormula As String
    valFormula = ws.Cells(ES_START_ROW, ES_COL_KATEGORIE).Validation.Formula1
    If Err.Number <> 0 Then
        msg = msg & "  Keine Validation vorhanden (Err " & Err.Number & ")" & vbLf
        Err.Clear
    Else
        msg = msg & "  Formula1 = """ & valFormula & """" & vbLf
        msg = msg & "  L?nge = " & Len(valFormula) & vbLf
    End If
    On Error GoTo 0
    
    Dim nextRow As Long
    nextRow = lastRow + 1
    If nextRow < ES_START_ROW Then nextRow = ES_START_ROW
    
    msg = msg & vbLf & "--- N?CHSTE FREIE ZEILE: " & nextRow & " ---" & vbLf
    On Error Resume Next
    valFormula = ws.Cells(nextRow, ES_COL_KATEGORIE).Validation.Formula1
    If Err.Number <> 0 Then
        msg = msg & "  Keine Validation vorhanden (Err " & Err.Number & ")" & vbLf
        Err.Clear
    Else
        msg = msg & "  Formula1 = """ & valFormula & """" & vbLf
        msg = msg & "  L?nge = " & Len(valFormula) & vbLf
    End If
    On Error GoTo 0
    
    msg = msg & vbLf & "--- HILFSSPALTE Daten!BA (Fallback-Quelle) ---" & vbLf
    Dim hilfLastRow As Long
    hilfLastRow = wsDaten.Cells(wsDaten.Rows.count, DATA_COL_ES_HILF).End(xlUp).Row
    If hilfLastRow < DATA_START_ROW Then
        msg = msg & "  (leer - kein Fallback aktiv)" & vbLf
    Else
        For r = DATA_START_ROW To hilfLastRow
            msg = msg & "  BA" & r & ": """ & Trim(CStr(wsDaten.Cells(r, DATA_COL_ES_HILF).value)) & """" & vbLf
        Next r
    End If
    
    Debug.Print msg
    MsgBox msg, vbInformation, "Debug DropDown Spalte B"
    
End Sub


' ===============================================================
' DEBUG: Validation einer bestimmten Zeile pr?fen
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
' DEBUG: SetzeDropDowns isoliert aufrufen und Ergebnis pr?fen
' Aufruf: Call DebugSetzeDropDownsUndPruefe
' ===============================================================
Public Sub DebugSetzeDropDownsUndPruefe()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' DropDowns neu setzen
    Call mod_Einstellungen_DropDowns.SetzeDropDowns(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    If lastRow < ES_START_ROW Then lastRow = ES_START_ROW - 1
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


















































