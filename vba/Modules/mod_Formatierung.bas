Attribute VB_Name = "mod_Formatierung"
Option Explicit

' ***************************************************************
' MODUL: mod_Formatierung
' ZWECK: Zentrale Verwaltung aller Tabellenformatierungen
' ***************************************************************

' ***************************************************************
' PROZEDUR: Anwende_Zebra_Formatierung_Direkt (Mit direkter Zellfärbung, ohne BF)
' ***************************************************************
Public Sub Anwende_Zebra_Formatierung_Direkt(ByVal ws As Worksheet, ByVal startCol As Long, ByVal endCol As Long, ByVal startRow As Long, ByVal dataCheckCol As Long)
    
    Const ZEBRA_COLOR As Long = &HDEE5E3
    
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    Dim lRow As Long
    Dim rngRow As Range
    Dim checkColValue As Variant
    
    ' 1. Letzte gefüllte Zeile in der Prüfspalte ermitteln
    lastRow = ws.Cells(ws.Rows.Count, dataCheckCol).End(xlUp).Row
    If lastRow < startRow Then
        Exit Sub ' Keine Daten vorhanden
    End If
    
    ' 2. Existierende Formatierungen (Farben) löschen
    On Error Resume Next
    ws.Range(ws.Cells(startRow, startCol), ws.Cells(lastRow, endCol)).Interior.ColorIndex = xlNone
    On Error GoTo 0
    
    ' 3. Direkte Zellenf?bung mit MOD-Logik
    For lRow = startRow To lastRow
        ' Prüfe ob Zelle in der Prüfspalte gefüllt ist
        checkColValue = ws.Cells(lRow, dataCheckCol).value
        If checkColValue <> "" And Not IsEmpty(checkColValue) Then
            ' Ungerade Zeilen (ab startRow) färben
            If (lRow - startRow) Mod 2 = 1 Then
                Set rngRow = ws.Range(ws.Cells(lRow, startCol), ws.Cells(lRow, endCol))
                rngRow.Interior.color = ZEBRA_COLOR
            End If
        End If
    Next lRow

End Sub


' ***************************************************************
' PROZEDUR: Entferne_Zebra_Formatierung (Löscht alle Zellenfarben vor Neuerstellung)
' ***************************************************************
Public Sub Entferne_Zebra_Formatierung()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    
    On Error Resume Next
    
    ' 1. Mitgliederliste (WS_MITGLIEDER)
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
        If lastRow >= M_START_ROW Then
            ws.Range(ws.Cells(M_START_ROW, M_COL_MEMBER_ID), ws.Cells(lastRow, M_COL_PACHTENDE)).Interior.ColorIndex = xlNone
        End If
    End If
    
    ' 2. Mitgliederhistorie (WS_MITGLIEDER_HISTORIE)
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, H_COL_NACHNAME).End(xlUp).Row
        If lastRow >= H_START_ROW Then
            ws.Range(ws.Cells(H_START_ROW, H_COL_PARZELLE), ws.Cells(lastRow, H_COL_SYSTEMZEIT)).Interior.ColorIndex = xlNone
        End If
    End If
    
    ' 3. Datenblatt - Tabelle 1 (WS_DATEN)
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, DATA_CAT_COL_START).End(xlUp).Row
        If lastRow >= DATA_START_ROW Then
            ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), ws.Cells(lastRow, DATA_CAT_COL_END)).Interior.ColorIndex = xlNone
        End If
        
        ' Datenblatt - Tabelle 2 (WS_DATEN)
        lastRow = ws.Cells(ws.Rows.Count, DATA_MAP_COL_ENTITYKEY).End(xlUp).Row
        If lastRow >= DATA_START_ROW Then
            ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_ENTITYKEY), ws.Cells(lastRow, 21)).Interior.ColorIndex = xlNone
        End If
    End If
    
    On Error GoTo 0
    
End Sub


' ***************************************************************
' PROZEDUR: FormatiereMitgliedertabelle (Komplette Tabellenformatierung)
' ***************************************************************
Public Sub FormatiereMitgliedertabelle()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rngData As Range
    Dim wasProtected As Boolean
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If ws Is Nothing Then Exit Sub
    
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect PASSWORD:=PASSWORD
    
    lastRow = ws.Cells(ws.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
    If lastRow < M_START_ROW Then GoTo Cleanup
    
    ' Datenbereich für Formatierung
    Set rngData = ws.Range(ws.Cells(M_START_ROW, M_COL_MEMBER_ID), ws.Cells(lastRow, M_COL_PACHTENDE))
    
    ' --- 1. RAHMENLINIE (dünne schwarze Linien) ---
    With rngData.Borders
        .LineStyle = xlContinuous
        .color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
    
    ' --- 2. SPALTE A (Member ID) - zentrisch ---
    With ws.Range(ws.Cells(M_START_ROW, M_COL_MEMBER_ID), ws.Cells(lastRow, M_COL_MEMBER_ID))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' --- 3. SPALTE B (Parzelle) - zentrisch ---
    With ws.Range(ws.Cells(M_START_ROW, M_COL_PARZELLE), ws.Cells(lastRow, M_COL_PARZELLE))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' --- 4. SPALTE H (Nummer/Hausnummer) - zentrisch ---
    With ws.Range(ws.Cells(M_START_ROW, M_COL_NUMMER), ws.Cells(lastRow, M_COL_NUMMER))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' --- 5. SPALTE I (PLZ) - zentrisch ---
    With ws.Range(ws.Cells(M_START_ROW, M_COL_PLZ), ws.Cells(lastRow, M_COL_PLZ))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' --- 6. SPALTE C (Seite) - zentrisch ---
    With ws.Range(ws.Cells(M_START_ROW, M_COL_SEITE), ws.Cells(lastRow, M_COL_SEITE))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' --- 7. DATUMSFORMATE ---
    With ws.Range(ws.Cells(M_START_ROW, M_COL_GEBURTSTAG), ws.Cells(lastRow, M_COL_GEBURTSTAG))
        .NumberFormat = "dd.mm.yyyy"
    End With
    
    With ws.Range(ws.Cells(M_START_ROW, M_COL_PACHTANFANG), ws.Cells(lastRow, M_COL_PACHTANFANG))
        .NumberFormat = "dd.mm.yyyy"
    End With
    
    With ws.Range(ws.Cells(M_START_ROW, M_COL_PACHTENDE), ws.Cells(lastRow, M_COL_PACHTENDE))
        .NumberFormat = "dd.mm.yyyy"
    End With
    
Cleanup:
    If wasProtected Then ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Exit Sub
ErrorHandler:
    If wasProtected Then ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler beim Formatieren der Mitgliedertabelle: " & Err.Description, vbCritical
End Sub

' ***************************************************************
' PROZEDUR: Formatiere_Mitgliederhistorie (Formatierung für Historientabelle)
' ***************************************************************
Public Sub Formatiere_Mitgliederhistorie()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rngData As Range
    Dim wasProtected As Boolean
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    If ws Is Nothing Then Exit Sub
    
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect PASSWORD:=PASSWORD
    
    lastRow = ws.Cells(ws.Rows.Count, H_COL_NACHNAME).End(xlUp).Row
    If lastRow < H_START_ROW Then GoTo Cleanup
    
    ' Datenbereich für Formatierung
    Set rngData = ws.Range(ws.Cells(H_START_ROW, H_COL_PARZELLE), ws.Cells(lastRow, H_COL_SYSTEMZEIT))
    
    ' --- 1. RAHMENLINIE (dünne schwarze Linien) ---
    With rngData.Borders
        .LineStyle = xlContinuous
        .color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
    
    ' --- 2. SPALTEN AUSRICHTUNG ---
    With ws.Range(ws.Cells(H_START_ROW, H_COL_PARZELLE), ws.Cells(lastRow, H_COL_PARZELLE))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' --- 3. DATUMSFORMATE ---
    With ws.Range(ws.Cells(H_START_ROW, H_COL_AUST_DATUM), ws.Cells(lastRow, H_COL_AUST_DATUM))
        .NumberFormat = "dd.mm.yyyy"
    End With
    
    With ws.Range(ws.Cells(H_START_ROW, H_COL_SYSTEMZEIT), ws.Cells(lastRow, H_COL_SYSTEMZEIT))
        .NumberFormat = "dd.mm.yyyy hh:mm:ss"
    End With
    
    ' --- 4. ZEBRA-FORMATIERUNG (NACH anderen Formatierungen!) ---
    Call Anwende_Zebra_Formatierung_Direkt(ws, H_COL_PARZELLE, H_COL_SYSTEMZEIT, H_START_ROW, H_COL_NACHNAME)

Cleanup:
    If wasProtected Then ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Exit Sub
ErrorHandler:
    If wasProtected Then ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler beim Formatieren der Mitgliederhistorie: " & Err.Description, vbCritical
End Sub

' ***************************************************************
' PROZEDUR: Formatiere_Daten_Tabellen (Kategorie-Regeln und Mapping)
' ***************************************************************
Public Sub Formatiere_Daten_Tabellen()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rngData As Range
    Dim wasProtected As Boolean
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    If ws Is Nothing Then Exit Sub
    
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect PASSWORD:=PASSWORD
    
    ' --- TABELLE 1: KATEGORIE-REGELN (J bis Q) ---
    lastRow = ws.Cells(ws.Rows.Count, DATA_CAT_COL_START).End(xlUp).Row
    If lastRow >= DATA_START_ROW Then
        Set rngData = ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), ws.Cells(lastRow, DATA_CAT_COL_END))
        
        ' Rahmenlinie
        With rngData.Borders
            .LineStyle = xlContinuous
            .color = RGB(0, 0, 0)
            .Weight = xlThin
        End With
        
        ' Zebra-Formatierung (NACH Rahmenlinie!)
        Call Anwende_Zebra_Formatierung_Direkt(ws, DATA_CAT_COL_START, DATA_CAT_COL_END, DATA_START_ROW, DATA_CAT_COL_START)
    End If
    
    ' --- TABELLE 2: ENTITYKEY/MAPPING (S bis U) ---
    lastRow = ws.Cells(ws.Rows.Count, DATA_MAP_COL_ENTITYKEY).End(xlUp).Row
    If lastRow >= DATA_START_ROW Then
        Set rngData = ws.Range(ws.Cells(DATA_START_ROW, DATA_MAP_COL_ENTITYKEY), ws.Cells(lastRow, 21))
        
        ' Rahmenlinie
        With rngData.Borders
            .LineStyle = xlContinuous
            .color = RGB(0, 0, 0)
            .Weight = xlThin
        End With
        
        ' Zebra-Formatierung (NACH Rahmenlinie!)
        Call Anwende_Zebra_Formatierung_Direkt(ws, DATA_MAP_COL_ENTITYKEY, 21, DATA_START_ROW, DATA_MAP_COL_ENTITYKEY)
    End If

Cleanup:
    If wasProtected Then ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Exit Sub
ErrorHandler:
    If wasProtected Then ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler beim Formatieren der Datentabellen: " & Err.Description, vbCritical
End Sub

' ***************************************************************
' PROZEDUR: Formatiere_Alle_Tabellen_Neu (Zentrale Formatierungs-Koordination)
' WICHTIG: Zebra-Formatierung wird NACH FormatiereMitgliedertabelle angewendet!
' ***************************************************************
Public Sub Formatiere_Alle_Tabellen_Neu()

    Dim wsM As Worksheet
    Dim wsD As Worksheet
    Dim wsH As Worksheet

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' 1. Mitgliederliste (WS_MITGLIEDER)
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If Not wsM Is Nothing Then
        Dim wasProtectedM As Boolean
        wasProtectedM = wsM.ProtectContents
        If wasProtectedM Then wsM.Unprotect PASSWORD:=PASSWORD
        
        ' Spezielle Formatierungen (Borders, Alignment, etc.) ZUERST
        Call FormatiereMitgliedertabelle
        
        ' Zebra-Formatierung NACH anderen Formatierungen! (A bis Q, Prüfspalte: Nachname)
        Call Anwende_Zebra_Formatierung_Direkt(wsM, M_COL_MEMBER_ID, M_COL_PACHTENDE, M_START_ROW, M_COL_NACHNAME)
        
        If wasProtectedM Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    
    ' 2. Mitgliederhistorie (WS_MITGLIEDER_HISTORIE)
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    If Not wsH Is Nothing Then
        Dim wasProtectedH As Boolean
        wasProtectedH = wsH.ProtectContents
        If wasProtectedH Then wsH.Unprotect PASSWORD:=PASSWORD
        
        ' Formatierung für Historientabelle (Zebra wird dort bereits NACH anderen Formatierungen angewendet)
        Call Formatiere_Mitgliederhistorie
        
        If wasProtectedH Then wsH.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    
    ' 3. Datenblatt (WS_DATEN)
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    If Not wsD Is Nothing Then
        Dim wasProtectedD As Boolean
        wasProtectedD = wsD.ProtectContents
        If wasProtectedD Then wsD.Unprotect PASSWORD:=PASSWORD
        
        ' Formatierung für beide Tabellen (Zebra wird dort bereits NACH Rahmenlinie angewendet)
        Call Formatiere_Daten_Tabellen
        
        If wasProtectedD Then wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    
    ' BANKKONTO: Wird NICHT formatiert - nutzt bedingte Formatierung stattdessen

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "FEHLER beim Formatieren der Tabellen: " & Err.Description, vbCritical
End Sub

' ***************************************************************
' PROZEDUR: DebugNachFormatierung (Debug-Funktion für Zebra-Formatierung)
' ***************************************************************
Public Sub DebugNachFormatierung()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lRow As Long
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If ws Is Nothing Then Exit Sub
    
    Debug.Print ""
    Debug.Print "===== DEBUG VOR Formatiere_Alle_Tabellen_Neu ====="
    Debug.Print "Zeile 28 Spalte 1 Farbe (VOR): " & Hex(ws.Cells(28, 1).Interior.color)
    
    Debug.Print ""
    Debug.Print "--- FARBEN VOR FORMATIERUNG (letzte 5 Zeilen) ---"
    For lRow = 24 To 28
        Debug.Print "Zeile " & lRow & ": " & Hex(ws.Cells(lRow, 1).Interior.color)
    Next lRow
    
    Debug.Print ""
    Debug.Print "===== Formatiere_Alle_Tabellen_Neu START ====="
    Debug.Print ""
    
    Call Formatiere_Alle_Tabellen_Neu
    
    Debug.Print ""
    Debug.Print "===== Formatiere_Alle_Tabellen_Neu ENDE ====="
    Debug.Print ""
    Debug.Print "===== DEBUG NACH Formatiere_Alle_Tabellen_Neu ====="
    Debug.Print "Zeile 28 Spalte 1 Farbe (NACH): " & Hex(ws.Cells(28, 1).Interior.color)
    
    Debug.Print ""
    Debug.Print "--- FARBEN NACH FORMATIERUNG (letzte 5 Zeilen) ---"
    For lRow = 24 To 28
        Debug.Print "Zeile " & lRow & ": " & Hex(ws.Cells(lRow, 1).Interior.color)
    Next lRow
    
    Debug.Print ""
    Debug.Print "===== END DEBUG ====="
    
End Sub
