Attribute VB_Name = "mod_Format_Protection"
Option Explicit

' ***************************************************************
' MODUL: mod_Format_Protection
' ZWECK: Blattschutz, Ent-/Sperren editierbarer Spalten, Ausblenden
' ABGELEITET AUS: mod_Formatierung (Modularisierung)
' VERSION: 1.0 - 01.03.2026
' FUNKTIONEN:
'   - EntspeerreEditierbareSpalten: Unlock/Lock Logik fuer Daten-Blatt
'   - BlendeDatenSpaltenAus: Hilfsspalten ausblenden
' ***************************************************************

' ===============================================================
' Blendet Hilfsspalten aus (D-I, Z-AB, AE-AH, BA)
' ===============================================================
Public Sub BlendeDatenSpaltenAus()
    
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ws.Range("D:I").EntireColumn.Hidden = True
    ws.Range("Z:AB").EntireColumn.Hidden = True
    ws.Range("AE:AH").EntireColumn.Hidden = True
    ws.Columns(DATA_COL_ES_HILF).Hidden = True
    
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
End Sub

' ===============================================================
' Entsperrt bestehende Daten + genau 1 naechste freie Zeile
' Sperrt einen Puffer darunter
' Betrifft: B, D, F, H, J-P, R-X (via W), AB, AC, AD, AH
' ===============================================================
Public Sub EntspeerreEditierbareSpalten(ByRef ws As Worksheet)
    
    Dim lastRow As Long
    Dim nextRow As Long
    Dim lockEnd As Long
    Dim r As Long
    Dim lastRowDD As Long
    
    On Error Resume Next
    
    ' === EINZELSPALTEN: B (2), D (4), F (6), H (8) ===
    Dim singleCols As Variant
    Dim c As Long
    singleCols = Array(2, 4, 6, 8)
    
    For c = LBound(singleCols) To UBound(singleCols)
        lastRow = ws.Cells(ws.Rows.count, singleCols(c)).End(xlUp).Row
        If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW - 1
        nextRow = lastRow + 1
        lockEnd = nextRow + 50
        
        If lastRow >= DATA_START_ROW Then
            ws.Range(ws.Cells(DATA_START_ROW, singleCols(c)), _
                     ws.Cells(lastRow, singleCols(c))).Locked = False
        End If
        
        ws.Cells(nextRow, singleCols(c)).Locked = False
        
        ws.Range(ws.Cells(nextRow + 1, singleCols(c)), _
                 ws.Cells(lockEnd, singleCols(c))).Locked = True
    Next c
    
    ' === KATEGORIE-TABELLE: J-P (10-16) ===
    lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW - 1
    nextRow = lastRow + 1
    lockEnd = nextRow + 50
    
    If lastRow >= DATA_START_ROW Then
        ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), _
                 ws.Cells(lastRow, DATA_CAT_COL_END)).Locked = False
    End If
    
    ws.Range(ws.Cells(nextRow, DATA_CAT_COL_START), _
             ws.Cells(nextRow, DATA_CAT_COL_END)).Locked = False
    
    ws.Range(ws.Cells(nextRow + 1, DATA_CAT_COL_START), _
             ws.Cells(lockEnd, DATA_CAT_COL_END)).Locked = True
    
    ' DropDowns fuer die Eingabezeile
    Call mod_Format_Kategorie.SetzeZielspalteDropdown(ws, nextRow, "")
    
    ' Dropdown K (E/A)
    ws.Cells(nextRow, DATA_CAT_COL_EINAUS).Validation.Delete
    With ws.Cells(nextRow, DATA_CAT_COL_EINAUS).Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:="E,A"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    
    ' Dropdown M (Prioritaet)
    ws.Cells(nextRow, DATA_CAT_COL_PRIORITAET).Validation.Delete
    With ws.Cells(nextRow, DATA_CAT_COL_PRIORITAET).Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:="=" & WS_DATEN & "!$AA$4:$AA$" & _
                        ws.Cells(ws.Rows.count, DATA_COL_DD_PRIORITAET).End(xlUp).Row
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    
    ' Dropdown O (Faelligkeit)
    ws.Cells(nextRow, DATA_CAT_COL_FAELLIGKEIT).Validation.Delete
    With ws.Cells(nextRow, DATA_CAT_COL_FAELLIGKEIT).Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:="=" & WS_DATEN & "!$AC$4:$AC$" & _
                        ws.Cells(ws.Rows.count, DATA_COL_DD_FAELLIGKEIT).End(xlUp).Row
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    
    ' === ENTITYKEY-TABELLE: R-X (18-24) ===
    lastRow = ws.Cells(ws.Rows.count, EK_COL_IBAN).End(xlUp).Row
    Dim lastRowR As Long
    lastRowR = ws.Cells(ws.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    If lastRowR > lastRow Then lastRow = lastRowR
    
    If lastRow < EK_START_ROW Then lastRow = EK_START_ROW - 1
    nextRow = lastRow + 1
    lockEnd = nextRow + 50
    
    ws.Range(ws.Cells(nextRow, EK_COL_ENTITYKEY), _
             ws.Cells(nextRow, EK_COL_DEBUG)).Locked = False
    
    ws.Range(ws.Cells(nextRow + 1, EK_COL_ENTITYKEY), _
             ws.Cells(lockEnd, EK_COL_DEBUG)).Locked = True
    
    ' EntityRole-Dropdown (W) Quelle: Spalte AD
    lastRowDD = ws.Cells(ws.Rows.count, DATA_COL_DD_ENTITYROLE).End(xlUp).Row
    If lastRowDD < DATA_START_ROW Then lastRowDD = DATA_START_ROW
    
    ws.Cells(nextRow, EK_COL_ROLE).Validation.Delete
    With ws.Cells(nextRow, EK_COL_ROLE).Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:="=" & WS_DATEN & "!$AD$" & DATA_START_ROW & ":$AD$" & lastRowDD
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    
    ' EntityRole-Dropdown und Parzellen-Dropdown fuer ALLE Zeilen
    Dim lastRowParzelle As Long
    lastRowParzelle = ws.Cells(ws.Rows.count, DATA_COL_DD_PARZELLE).End(xlUp).Row
    If lastRowParzelle < DATA_START_ROW Then lastRowParzelle = DATA_START_ROW
    
    If lastRow >= EK_START_ROW Then
        For r = EK_START_ROW To lastRow
            Dim currentRole As String
            currentRole = UCase(Trim(ws.Cells(r, EK_COL_ROLE).value))
            
            ws.Cells(r, EK_COL_ROLE).Validation.Delete
            With ws.Cells(r, EK_COL_ROLE).Validation
                .Add Type:=xlValidateList, _
                     AlertStyle:=xlValidAlertWarning, _
                     Formula1:="=" & WS_DATEN & "!$AD$" & DATA_START_ROW & ":$AD$" & lastRowDD
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = False
                .ShowError = True
            End With
            ws.Cells(r, EK_COL_ROLE).Locked = False
            
            ws.Cells(r, EK_COL_ZUORDNUNG).Locked = False
            ws.Cells(r, EK_COL_DEBUG).Locked = False
            
            If currentRole = "EHEMALIGES MITGLIED" Or currentRole = "SONSTIGE" Then
                ws.Cells(r, EK_COL_PARZELLE).Validation.Delete
                With ws.Cells(r, EK_COL_PARZELLE).Validation
                    .Add Type:=xlValidateList, _
                         AlertStyle:=xlValidAlertWarning, _
                         Formula1:="=" & WS_DATEN & "!$F$" & DATA_START_ROW & ":$F$" & lastRowParzelle
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = False
                    .ShowError = True
                End With
                ws.Cells(r, EK_COL_PARZELLE).Locked = False
            End If
        Next r
    End If
    
    ' === HELPER-SPALTEN: AB (28), AC (29), AD (30), AH (34) ===
    Dim helperCols As Variant
    helperCols = Array(28, 29, 30, 34)
    
    For c = LBound(helperCols) To UBound(helperCols)
        lastRow = ws.Cells(ws.Rows.count, helperCols(c)).End(xlUp).Row
        If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW - 1
        nextRow = lastRow + 1
        lockEnd = nextRow + 50
        
        If lastRow >= DATA_START_ROW Then
            ws.Range(ws.Cells(DATA_START_ROW, helperCols(c)), _
                     ws.Cells(lastRow, helperCols(c))).Locked = False
        End If
        
        ws.Cells(nextRow, helperCols(c)).Locked = False
        
        ws.Range(ws.Cells(nextRow + 1, helperCols(c)), _
                 ws.Cells(lockEnd, helperCols(c))).Locked = True
    Next c
    
    On Error GoTo 0
    
End Sub






































































