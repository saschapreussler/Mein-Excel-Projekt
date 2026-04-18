Attribute VB_Name = "mod_Format_EntityKey"
Option Explicit

' ***************************************************************
' MODUL: mod_Format_EntityKey
' ZWECK: EntityKey-Tabelle (R-X) Formatierung, Sortierung, Zellschutz
' ABGELEITET AUS: mod_Formatierung (Modularisierung)
' VERSION: 1.0 - 01.03.2026
' FUNKTIONEN:
'   - FormatiereEntityKeyTabelleKomplett: Wrapper mit lastRow
'   - FormatiereEntityKeyTabelle: Core-Formatierung R-X
'   - SortiereEntityKeyTabelle: Bubble-Sort nach Parzelle/Prefix
'   - VergleicheEntityKeyZeilen: Vergleich zweier EK-Zeilen
'   - SetzeZellschutzFuerZeile: Zellschutz pro Zeile
' ***************************************************************

Private Const ZEBRA_COLOR_1 As Long = &HFFFFFF  ' Weiss
Private Const ZEBRA_COLOR_2 As Long = &HDEE5E3  ' Hellgrau

' ===============================================================
' WRAPPER: Formatiert die EntityKey-Tabelle komplett
' ===============================================================
Public Sub FormatiereEntityKeyTabelleKomplett(ByRef ws As Worksheet)
    
    Dim lastRow As Long
    Dim lastRowIBAN As Long
    
    lastRow = ws.Cells(ws.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    lastRowIBAN = ws.Cells(ws.Rows.count, EK_COL_IBAN).End(xlUp).Row
    If lastRowIBAN > lastRow Then lastRow = lastRowIBAN
    
    If lastRow >= EK_START_ROW Then
        Call FormatiereEntityKeyTabelle(ws, lastRow)
    End If
    
End Sub

' ===============================================================
' Core EntityKey-Formatierung
' R-T mit Zebra, U-X NUR Rahmen (Ampelfarben bleiben erhalten!)
' ===============================================================
Private Sub FormatiereEntityKeyTabelle(ByRef ws As Worksheet, ByVal lastRow As Long)
    
    Dim rngTable As Range
    Dim rngZebra As Range
    Dim rngAmpel As Range
    Dim r As Long
    Dim currentRole As String
    Dim kontoWert As String
    Dim zuordnungWert As String
    
    If lastRow < EK_START_ROW Then Exit Sub
    
    Set rngTable = ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                            ws.Cells(lastRow, EK_COL_DEBUG))
    
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    rngTable.VerticalAlignment = xlCenter
    
    ' Spalte R (EntityKey)
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                  ws.Cells(lastRow, EK_COL_ENTITYKEY))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ENTITYKEY).ColumnWidth = 11
    
    ' Spalte S (IBAN)
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_IBAN), _
                  ws.Cells(lastRow, EK_COL_IBAN))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_IBAN).AutoFit
    
    ' Spalte T: WrapText NUR wenn vbLf im Wert
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_KONTONAME), _
             ws.Cells(lastRow, EK_COL_KONTONAME)).HorizontalAlignment = xlLeft
    
    For r = EK_START_ROW To lastRow
        kontoWert = CStr(ws.Cells(r, EK_COL_KONTONAME).value)
        If InStr(kontoWert, vbLf) > 0 Then
            ws.Cells(r, EK_COL_KONTONAME).WrapText = True
        Else
            ws.Cells(r, EK_COL_KONTONAME).WrapText = False
        End If
    Next r
    ws.Columns(EK_COL_KONTONAME).ColumnWidth = 36
    
    ' Spalte U (Zuordnung)
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_ZUORDNUNG), _
             ws.Cells(lastRow, EK_COL_ZUORDNUNG)).HorizontalAlignment = xlLeft
    
    For r = EK_START_ROW To lastRow
        zuordnungWert = CStr(ws.Cells(r, EK_COL_ZUORDNUNG).value)
        If InStr(zuordnungWert, vbLf) > 0 Then
            ws.Cells(r, EK_COL_ZUORDNUNG).WrapText = True
        Else
            ws.Cells(r, EK_COL_ZUORDNUNG).WrapText = False
        End If
    Next r
    ws.Columns(EK_COL_ZUORDNUNG).ColumnWidth = 28
    
    ' Spalte V (Parzelle)
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_PARZELLE), _
                  ws.Cells(lastRow, EK_COL_PARZELLE))
        .WrapText = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns(EK_COL_PARZELLE).ColumnWidth = 9
    
    ' Spalte W (Role)
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ROLE), _
                  ws.Cells(lastRow, EK_COL_ROLE))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ROLE).AutoFit
    
    ' Spalte X (Debug) - WrapText erlaubt, Breite 65
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_DEBUG), _
                  ws.Cells(lastRow, EK_COL_DEBUG))
        .WrapText = True
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_DEBUG).ColumnWidth = 65
    
    ' R-T immer gesperrt
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
             ws.Cells(lastRow, EK_COL_KONTONAME)).Locked = True
    
    For r = EK_START_ROW To lastRow
        currentRole = Trim(ws.Cells(r, EK_COL_ROLE).value)
        
        Call SetzeZellschutzFuerZeile(ws, r, currentRole)
        
        ' Zebra NUR fuer R-T
        Set rngZebra = ws.Range(ws.Cells(r, EK_COL_ENTITYKEY), ws.Cells(r, EK_COL_KONTONAME))
        
        If (r - EK_START_ROW) Mod 2 = 0 Then
            rngZebra.Interior.color = ZEBRA_COLOR_1
        Else
            rngZebra.Interior.color = ZEBRA_COLOR_2
        End If
    Next r
    
    ws.Rows(EK_START_ROW & ":" & lastRow).AutoFit
    
End Sub

' ===============================================================
' Sortiert die EntityKey-Tabelle
' ===============================================================
Public Sub SortiereEntityKeyTabelle(Optional ByRef ws As Worksheet = Nothing)
    
    Dim lastRow As Long
    Dim r As Long
    
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    lastRow = ws.Cells(ws.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    Dim lastRowIBAN As Long
    lastRowIBAN = ws.Cells(ws.Rows.count, EK_COL_IBAN).End(xlUp).Row
    If lastRowIBAN > lastRow Then lastRow = lastRowIBAN
    
    If lastRow < EK_START_ROW Then Exit Sub
    
    Dim arrData() As Variant
    Dim i As Long, j As Long
    Dim numRows As Long
    Dim swap As Boolean
    Dim tempRow As Variant
    
    numRows = lastRow - EK_START_ROW + 1
    If numRows < 1 Then Exit Sub
    
    ReDim arrData(1 To numRows, 1 To 7)
    
    For r = EK_START_ROW To lastRow
        arrData(r - EK_START_ROW + 1, 1) = ws.Cells(r, EK_COL_ENTITYKEY).value
        arrData(r - EK_START_ROW + 1, 2) = ws.Cells(r, EK_COL_IBAN).value
        arrData(r - EK_START_ROW + 1, 3) = ws.Cells(r, EK_COL_KONTONAME).value
        arrData(r - EK_START_ROW + 1, 4) = ws.Cells(r, EK_COL_ZUORDNUNG).value
        arrData(r - EK_START_ROW + 1, 5) = ws.Cells(r, EK_COL_PARZELLE).value
        arrData(r - EK_START_ROW + 1, 6) = ws.Cells(r, EK_COL_ROLE).value
        arrData(r - EK_START_ROW + 1, 7) = ws.Cells(r, EK_COL_DEBUG).value
    Next r
    
    For i = 1 To numRows - 1
        swap = False
        For j = 1 To numRows - i
            If VergleicheEntityKeyZeilen(arrData(j, 1), arrData(j, 5), arrData(j + 1, 1), arrData(j + 1, 5)) > 0 Then
                ReDim tempRow(1 To 7)
                Dim k As Long
                For k = 1 To 7
                    tempRow(k) = arrData(j, k)
                    arrData(j, k) = arrData(j + 1, k)
                    arrData(j + 1, k) = tempRow(k)
                Next k
                swap = True
            End If
        Next j
        If Not swap Then Exit For
    Next i
    
    For r = EK_START_ROW To lastRow
        ws.Cells(r, EK_COL_ENTITYKEY).value = arrData(r - EK_START_ROW + 1, 1)
        ws.Cells(r, EK_COL_IBAN).value = arrData(r - EK_START_ROW + 1, 2)
        ws.Cells(r, EK_COL_KONTONAME).value = arrData(r - EK_START_ROW + 1, 3)
        ws.Cells(r, EK_COL_ZUORDNUNG).value = arrData(r - EK_START_ROW + 1, 4)
        ws.Cells(r, EK_COL_PARZELLE).value = arrData(r - EK_START_ROW + 1, 5)
        ws.Cells(r, EK_COL_ROLE).value = arrData(r - EK_START_ROW + 1, 6)
        ws.Cells(r, EK_COL_DEBUG).value = arrData(r - EK_START_ROW + 1, 7)
    Next r
    
    ' Ampelfarben NACH Sortierung neu berechnen
    Call mod_EntityKey_Ampel.SetzeAlleAmpelfarbenNachSortierung(ws)
    
End Sub

' ===============================================================
' Vergleicht zwei EntityKey-Zeilen fuer Sortierung
' ===============================================================
Private Function VergleicheEntityKeyZeilen(entityKey1 As Variant, parzelle1 As Variant, _
                                            entityKey2 As Variant, parzelle2 As Variant) As Long
    
    Dim order1 As Long
    Dim order2 As Long
    Dim parzelleStr1 As String
    Dim parzelleStr2 As String
    Dim entityStr1 As String
    Dim entityStr2 As String
    
    parzelleStr1 = Trim(CStr(parzelle1))
    parzelleStr2 = Trim(CStr(parzelle2))
    entityStr1 = Trim(CStr(entityKey1))
    entityStr2 = Trim(CStr(entityKey2))
    
    If InStr(parzelleStr1, ",") > 0 Then
        parzelleStr1 = Trim(Left(parzelleStr1, InStr(parzelleStr1, ",") - 1))
    End If
    If InStr(parzelleStr2, ",") > 0 Then
        parzelleStr2 = Trim(Left(parzelleStr2, InStr(parzelleStr2, ",") - 1))
    End If
    
    If IsNumeric(parzelleStr1) And parzelleStr1 <> "" Then
        order1 = CLng(parzelleStr1)
    ElseIf Left(UCase(entityStr1), 3) = "EX-" Then
        order1 = 100
    ElseIf Left(UCase(entityStr1), 5) = "VERS-" Then
        order1 = 200
    ElseIf Left(UCase(entityStr1), 5) = "BANK-" Then
        order1 = 300
    Else
        order1 = 400
    End If
    
    If IsNumeric(parzelleStr2) And parzelleStr2 <> "" Then
        order2 = CLng(parzelleStr2)
    ElseIf Left(UCase(entityStr2), 3) = "EX-" Then
        order2 = 100
    ElseIf Left(UCase(entityStr2), 5) = "VERS-" Then
        order2 = 200
    ElseIf Left(UCase(entityStr2), 5) = "BANK-" Then
        order2 = 300
    Else
        order2 = 400
    End If
    
    If order1 < order2 Then
        VergleicheEntityKeyZeilen = -1
    ElseIf order1 > order2 Then
        VergleicheEntityKeyZeilen = 1
    Else
        VergleicheEntityKeyZeilen = 0
    End If
    
End Function

' ===============================================================
' Setzt Zellschutz basierend auf EntityRole
' ===============================================================
Private Sub SetzeZellschutzFuerZeile(ByRef ws As Worksheet, ByVal zeile As Long, ByVal currentRole As String)
    
    On Error Resume Next
    
    ws.Range(ws.Cells(zeile, EK_COL_ENTITYKEY), ws.Cells(zeile, EK_COL_KONTONAME)).Locked = True
    ws.Cells(zeile, EK_COL_ZUORDNUNG).Locked = False
    ws.Cells(zeile, EK_COL_ROLE).Locked = False
    ws.Cells(zeile, EK_COL_DEBUG).Locked = False
    
    Dim roleUpper As String
    roleUpper = UCase(Trim(currentRole))
    
    If roleUpper = "EHEMALIGES MITGLIED" Or roleUpper = "SONSTIGE" Or roleUpper = "" Or roleUpper = "UNBEKANNT" Then
        ws.Cells(zeile, EK_COL_PARZELLE).Locked = False
    Else
        ws.Cells(zeile, EK_COL_PARZELLE).Locked = True
    End If
    
    On Error GoTo 0
    
End Sub












































































