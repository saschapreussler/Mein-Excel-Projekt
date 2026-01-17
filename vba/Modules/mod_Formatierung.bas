Attribute VB_Name = "mod_Formatierung"
Option Explicit

' ***************************************************************
' MODUL: mod_Formatierung
' ZWECK: Zentrale Verwaltung aller Tabellenformatierungen
' ***************************************************************

' ***************************************************************
' PROZEDUR: Anwende_Zebra_Formatierung (Universelle BF mit Prüfspalte)
' ***************************************************************
Public Sub Anwende_Zebra_Formatierung(ByVal ws As Worksheet, ByVal startCol As Long, ByVal endCol As Long, ByVal startRow As Long, ByVal dataCheckCol As Long)
    
    Const ZEBRA_COLOR As Long = &HDEE5E3
    
    If ws Is Nothing Then Exit Sub

    Dim rngFullData As Range
    Dim sFormula As String
    
    ' 1. Zielbereich definieren
    Set rngFullData = ws.Range(ws.Cells(startRow, startCol), ws.Cells(1000, endCol))
    
    ' 2. Bestehende Regeln im BF-Bereich LÖSCHEN
    On Error Resume Next
    rngFullData.FormatConditions.Delete
    On Error GoTo 0
    
    ' 3. Explizites Entfernen aller manuellen Zellfüllungen im Bereich
    rngFullData.Interior.color = xlNone
    
    ' 4. Formel erstellen: =UND(NICHT(ISTLEER($[Prüfspalte][Startzeile])); REST(ZEILE();2)=0)
    Dim checkColLetter As String
    checkColLetter = Split(ws.Columns(dataCheckCol).Address(False, False), ":")(0)
    
    sFormula = "=UND(NICHT(ISTLEER($" & checkColLetter & startRow & ")); REST(ZEILE();2)=0)"
    
    With rngFullData.FormatConditions.Add(Type:=xlExpression, Formula1:=sFormula)
        .Interior.color = ZEBRA_COLOR
        .StopIfTrue = True
        .Priority = 1
    End With

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
' PROZEDUR: Formatiere_Alle_Tabellen_Neu (Zentrale Formatierungs-Koordination)
' ***************************************************************
Public Sub Formatiere_Alle_Tabellen_Neu()

    Dim wsM As Worksheet
    Dim wsD As Worksheet
    Dim wasProtectedM As Boolean
    Dim wasProtectedD As Boolean

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' 1. Mitgliederliste (WS_MITGLIEDER)
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If Not wsM Is Nothing Then
        wasProtectedM = wsM.ProtectContents
        If wasProtectedM Then wsM.Unprotect PASSWORD:=PASSWORD
        
        ' Zebra-Formatierung (A bis Q)
        Call Anwende_Zebra_Formatierung(wsM, M_COL_MEMBER_ID, M_COL_PACHTENDE, M_START_ROW, M_COL_NACHNAME)
        
        ' Spezielle Formatierungen (Borders, Alignment, etc.)
        Call FormatiereMitgliedertabelle
        
        If wasProtectedM Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    
    ' 2. Datenblatt (WS_DATEN) - Kategorie-Regeln
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    If Not wsD Is Nothing Then
        wasProtectedD = wsD.ProtectContents
        If wasProtectedD Then wsD.Unprotect PASSWORD:=PASSWORD
        
        ' BF 1: Kategorie-Regeln (J bis Q, Startzeile 4, Prüfspalte J)
        Call Anwende_Zebra_Formatierung(wsD, DATA_CAT_COL_START, DATA_CAT_COL_END, DATA_START_ROW, DATA_CAT_COL_START)
        
        ' BF 2: EntityKey/Mapping-Tabelle (S bis U, Startzeile 4, Prüfspalte S)
        ' WICHTIG: Nur bis Spalte U (21) wegen Ampel-Logik für die Farben!
        Call Anwende_Zebra_Formatierung(wsD, DATA_MAP_COL_ENTITYKEY, 21, DATA_START_ROW, DATA_MAP_COL_ENTITYKEY)
        
        If wasProtectedD Then wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
ErrorHandler:
    MsgBox "FEHLER beim Formatieren der Tabellen: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
