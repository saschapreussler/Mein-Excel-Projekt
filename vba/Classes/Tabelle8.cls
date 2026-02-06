VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' TABELLE8 (DATEN) - WORKSHEET EVENT HANDLER
' VERSION: 2.0 - 06.02.2026
' ZWECK: Automatische Wartung aller Tabellen bei Datenaenderungen
'
' BEHANDELTE BEREICHE:
'   - B (col 2): Vereinsfunktionen - Zebra + Rahmen
'   - D (col 4): Anredeformen - Zebra + Rahmen
'   - F (col 6): Parzelle - Zebra + Rahmen
'   - H (col 8): Seite - Zebra + Rahmen
'   - J-P (10-16): Kategorie-Tabelle - Dropdown-Updates + Sortierung
'   - R-X (18-24): EntityKey-Tabelle - Role-Verarbeitung + Sortierung
'   - Z-AH (26-34): Helper-Spalten - Unabhaengige Formatierung pro Spalte
'
' WICHTIG:
'   - Spalten U-X erhalten NIE Zebra-Formatierung (Ampel-Bereich!)
'   - Helper-Spalten Z-AH werden als 9 unabhaengige Einheiten behandelt
'================================================================================

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rng_KategorieRegeln As Range
    Dim cat As String
    Dim r As Range
    Dim refValue As Variant
    Dim targetValue As Variant
    Dim zeile As Long
    Dim spalte As Long
    Dim needsKategorieSortAndFormat As Boolean
    Dim needsEntityKeySortAndFormat As Boolean

    On Error GoTo ExitPoint
    If Application.EnableEvents = False Then Exit Sub
    If Target.Cells.CountLarge <> 1 Then Exit Sub

    zeile = Target.Row
    spalte = Target.Column
    
    ' Nur Datenzeilen bearbeiten (ab Zeile 4)
    If zeile < DATA_START_ROW Then Exit Sub
    
    needsKategorieSortAndFormat = False
    needsEntityKeySortAndFormat = False
    
    ' =====================================================================
    ' ABSCHNITT 1: EINZELSPALTEN B, D, F, H (Zebra + Rahmen)
    ' =====================================================================
    If spalte = 2 Or spalte = 4 Or spalte = 6 Or spalte = 8 Then
        Application.EnableEvents = False
        Call mod_Formatierung.RefreshTableFormattingAfterChange(Me, spalte, zeile)
        Application.EnableEvents = True
        GoTo ExitPoint
    End If
    
    ' =====================================================================
    ' ABSCHNITT 2: ENTITYKEY-TABELLE (Spalten R-X = 18-24)
    ' =====================================================================
    If spalte >= EK_COL_ENTITYKEY And spalte <= EK_COL_DEBUG Then
        
        Application.EnableEvents = False
        
        ' Bei Role-Aenderung (Spalte W = 23): EntityKey anpassen
        If spalte = EK_COL_ROLE Then
            Call mod_EntityKey_Manager.VerarbeiteManuelleRoleAenderung(Target)
            ' Nach Role-Aenderung: Dropdown aktualisieren
            Call mod_Formatierung.UpdateEntityRoleDropdown(Me)
            needsEntityKeySortAndFormat = True
        End If
        
        ' Zeile formatieren (Stub in mod_EntityKey_Manager)
        Call mod_EntityKey_Manager.FormatiereEntityKeyZeile(zeile, Me)
        
        ' Bei EntityKey-Aenderung (Spalte R = 18): Sortierung triggern
        If spalte = EK_COL_ENTITYKEY Then
            needsEntityKeySortAndFormat = True
        End If
        
        ' Sortierung und Formatierung wenn erforderlich
        If needsEntityKeySortAndFormat Then
            Call mod_Formatierung.SortEntityKeyTable(Me)
        End If
        
        Application.EnableEvents = True
        GoTo ExitPoint
    End If
    
    ' =====================================================================
    ' ABSCHNITT 3: KATEGORIE-TABELLE (Spalten J-P = 10-16)
    ' =====================================================================
    If spalte >= DATA_CAT_COL_START And spalte <= DATA_CAT_COL_END Then
        
        Application.EnableEvents = False
        
        ' Bei Kategorie oder E/A Aenderung: Dropdowns aktualisieren
        If spalte = DATA_CAT_COL_KATEGORIE Or spalte = DATA_CAT_COL_EINAUS Then
            Call mod_Formatierung.OnKategorieChange(Target)
            needsKategorieSortAndFormat = True
        End If
        
        ' Sortierung und Formatierung wenn erforderlich
        If needsKategorieSortAndFormat Then
            Call mod_Formatierung.SortKategorieTable(Me)
            Call mod_Formatierung.FormatiereKategorieTabelle(Me)
        End If
        
        Application.EnableEvents = True
        
        ' === KATEGORIE-REGELN VALIDIERUNG ===
        Application.EnableEvents = True  ' Fuer MsgBox
        
        On Error Resume Next
        Set rng_KategorieRegeln = Me.Range("rng_KategorieRegeln")
        On Error GoTo ExitPoint
        
        If rng_KategorieRegeln Is Nothing Then GoTo ExitPoint
        
        Select Case spalte
            Case DATA_CAT_COL_EINAUS, DATA_CAT_COL_ZIELSPALTE, DATA_CAT_COL_FAELLIGKEIT
                If Intersect(Target, rng_KategorieRegeln) Is Nothing Then
                    GoTo ExitPoint
                End If
            Case Else
                GoTo ExitPoint
        End Select
        
        cat = Trim(CStr(Me.Cells(Target.Row, DATA_CAT_COL_KATEGORIE).value))
        If cat = "" Then GoTo ExitPoint
        
        targetValue = Me.Cells(Target.Row, Target.Column).value
        refValue = ""
        
        For Each r In Intersect(Me.Columns(DATA_CAT_COL_KATEGORIE), rng_KategorieRegeln).Cells
            
            If Trim(CStr(r.value)) = cat Then
                
                Dim currentValue As Variant
                currentValue = Me.Cells(r.Row, Target.Column).value
                
                If Not IsEmpty(currentValue) And Len(Trim(CStr(currentValue))) > 0 Then
                    If r.Row <> Target.Row Then
                        refValue = currentValue
                        Exit For
                    End If
                End If
            End If
        Next r
        
        If Not IsEmpty(refValue) And Len(Trim(CStr(refValue))) > 0 Then
            If Trim(CStr(targetValue)) <> Trim(CStr(refValue)) Then
                
                Application.EnableEvents = False
                Application.Undo
                
                MsgBox "Die Kategorie '" & cat & "' ist bereits mit dem Wert '" & refValue & "' definiert." & vbCrLf & _
                       "Unterschiedliche Werte in Spalte " & ColName(Target.Column) & " sind fuer diese Kategorie nicht erlaubt.", _
                       vbCritical, "Kategorie-Regel verletzt"
            End If
        End If
        
        GoTo ExitPoint
    End If
    
    ' =====================================================================
    ' ABSCHNITT 4: HELPER-SPALTEN Z-AH (26-34) - UNABHAENGIG!
    ' =====================================================================
    If spalte >= 26 And spalte <= 34 Then
        Application.EnableEvents = False
        ' WICHTIG: Jede Spalte wird als unabhaengige Einheit behandelt!
        Call mod_Formatierung.RefreshTableFormattingAfterChange(Me, spalte, zeile)
        Application.EnableEvents = True
        GoTo ExitPoint
    End If

ExitPoint:
    Application.EnableEvents = True
End Sub

'================================================================================
' Hilfsfunktion: Spaltenname aus Spaltennummer
'================================================================================
Private Function ColName(ByVal colNum As Long) As String
    Dim temp As String
    Do
        temp = Chr(((colNum - 1) Mod 26) + 65) & temp
        colNum = (colNum - 1) \ 26
    Loop While colNum > 0
    ColName = temp
End Function

'================================================================================
' ============================================================================
' ABSCHNITT: AUTOMATISCHE TABELLENWARTUNG
' Hinzugefuegt: 06.02.2026
' Zweck: Automatische Formatierung und Sortierung bei Datenaenderungen
' ============================================================================
'================================================================================

'--------------------------------------------------------------------------------
' FormatSingleColumn
' Zweck: Formatiert eine einzelne Spalte (B, D, F, H) mit Zebra/Rahmen
'--------------------------------------------------------------------------------
Public Sub FormatSingleColumn(ByRef ws As Worksheet, _
                               ByVal col As Long, _
                               Optional ByVal specificRow As Long = 0)
    
    Dim lastRow As Long
    Dim r As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim zebraColor1 As Long
    Dim zebraColor2 As Long
    
    On Error Resume Next
    
    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    
    zebraColor1 = RGB(255, 255, 255)
    zebraColor2 = ZEBRA_COLOR
    
    If specificRow > 0 And specificRow >= DATA_START_ROW Then
        startRow = specificRow
        endRow = specificRow
    Else
        startRow = DATA_START_ROW
        endRow = lastRow
    End If
    
    For r = startRow To endRow
        If (r Mod 2) = 0 Then
            ws.Cells(r, col).Interior.color = zebraColor1
        Else
            ws.Cells(r, col).Interior.color = zebraColor2
        End If
        
        With ws.Cells(r, col).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = RGB(0, 0, 0)
        End With
    Next r
    
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' FormatHelperColumnIndependent
'--------------------------------------------------------------------------------
Public Sub FormatHelperColumnIndependent(ByRef ws As Worksheet, ByVal col As Long)
    
    Dim lastRow As Long
    Dim r As Long
    Dim zebraColor1 As Long
    Dim zebraColor2 As Long
    
    On Error Resume Next
    
    If col < 26 Or col > 34 Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    
    zebraColor1 = RGB(255, 255, 255)
    zebraColor2 = ZEBRA_COLOR
    
    For r = DATA_START_ROW To lastRow
        If (r Mod 2) = 0 Then
            ws.Cells(r, col).Interior.color = zebraColor1
        Else
            ws.Cells(r, col).Interior.color = zebraColor2
        End If
        
        With ws.Cells(r, col).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = RGB(0, 0, 0)
        End With
    Next r
    
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' SortKategorieTable
'--------------------------------------------------------------------------------
Public Sub SortKategorieTable(ByRef ws As Worksheet)
    
    Dim lastRow As Long
    Dim rngSort As Range
    
    On Error GoTo ErrorHandler
    
    lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    
    Set rngSort = ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), _
                           ws.Cells(lastRow, DATA_CAT_COL_END))
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KATEGORIE), _
                                          ws.Cells(lastRow, DATA_CAT_COL_KATEGORIE)), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal
    
    With ws.Sort
        .SetRange rngSort
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
End Sub

'--------------------------------------------------------------------------------
' SortEntityKeyTable
'--------------------------------------------------------------------------------
Public Sub SortEntityKeyTable(ByRef ws As Worksheet)
    
    Dim lastRow As Long
    Dim rngData As Range
    Dim arrData() As Variant
    Dim arrKeys() As Long
    Dim i As Long, j As Long
    Dim entityKey As String
    Dim parzNum As Long
    Dim numRows As Long, numCols As Long
    Dim minIdx As Long, minKey As Long
    Dim tempKey As Long, tempVal As Variant
    
    On Error GoTo ErrorHandler
    
    lastRow = ws.Cells(ws.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    If lastRow < EK_START_ROW Then Exit Sub
    
    Set rngData = ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                           ws.Cells(lastRow, EK_COL_DEBUG))
    arrData = rngData.value
    
    numRows = UBound(arrData, 1)
    numCols = UBound(arrData, 2)
    
    ReDim arrKeys(1 To numRows)
    
    For i = 1 To numRows
        entityKey = UCase(Trim(CStr(arrData(i, 1))))
        
        If Left(entityKey, 2) = "P-" Then
            parzNum = 0
            On Error Resume Next
            parzNum = CLng(Mid(entityKey, 3))
            On Error GoTo ErrorHandler
            If parzNum >= 1 And parzNum <= 14 Then
                arrKeys(i) = parzNum
            Else
                arrKeys(i) = 100 + parzNum
            End If
        ElseIf Left(entityKey, 3) = "EX-" Then
            arrKeys(i) = 15 + (Asc(Mid(entityKey, 4, 1)) Mod 15)
        ElseIf Left(entityKey, 5) = "VERS-" Then
            arrKeys(i) = 30 + (Asc(Mid(entityKey, 6, 1)) Mod 15)
        ElseIf Left(entityKey, 5) = "BANK-" Then
            arrKeys(i) = 45 + (Asc(Mid(entityKey, 6, 1)) Mod 15)
        Else
            arrKeys(i) = 60 + (Asc(Left(entityKey, 1)) Mod 40)
        End If
    Next i
    
    For i = 1 To numRows - 1
        minIdx = i
        minKey = arrKeys(i)
        
        For j = i + 1 To numRows
            If arrKeys(j) < minKey Then
                minKey = arrKeys(j)
                minIdx = j
            End If
        Next j
        
        If minIdx <> i Then
            tempKey = arrKeys(i)
            arrKeys(i) = arrKeys(minIdx)
            arrKeys(minIdx) = tempKey
            
            For j = 1 To numCols
                tempVal = arrData(i, j)
                arrData(i, j) = arrData(minIdx, j)
                arrData(minIdx, j) = tempVal
            Next j
        End If
    Next i
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    rngData.value = arrData
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Call RefreshEntityKeyZebra(ws, lastRow)
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
End Sub

'--------------------------------------------------------------------------------
' RefreshEntityKeyZebra
'--------------------------------------------------------------------------------
Public Sub RefreshEntityKeyZebra(ByRef ws As Worksheet, ByVal lastRow As Long)
    
    Dim r As Long
    Dim rngZebra As Range
    Dim zebraColor1 As Long
    Dim zebraColor2 As Long
    
    On Error Resume Next
    
    zebraColor1 = RGB(255, 255, 255)
    zebraColor2 = ZEBRA_COLOR
    
    For r = EK_START_ROW To lastRow
        Set rngZebra = ws.Range(ws.Cells(r, EK_COL_ENTITYKEY), ws.Cells(r, EK_COL_KONTONAME))
        
        If (r Mod 2) = 0 Then
            rngZebra.Interior.color = zebraColor1
        Else
            rngZebra.Interior.color = zebraColor2
        End If
    Next r
    
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' UpdateEntityRoleDropdown
'--------------------------------------------------------------------------------
Public Sub UpdateEntityRoleDropdown(ByRef ws As Worksheet)
    
    Dim lastRowEK As Long
    Dim dictRoles As Object
    Dim r As Long
    Dim roleValue As String
    Dim nextRow As Long
    Dim key As Variant
    
    On Error Resume Next
    
    Set dictRoles = CreateObject("Scripting.Dictionary")
    
    dictRoles.Add "MITGLIED", "MITGLIED"
    dictRoles.Add "EXMITGLIED", "EXMITGLIED"
    dictRoles.Add "VERSORGER", "VERSORGER"
    dictRoles.Add "BANK", "BANK"
    dictRoles.Add "BEHOERDE", "BEHOERDE"
    dictRoles.Add "DIENSTLEISTER", "DIENSTLEISTER"
    dictRoles.Add "KUNDE", "KUNDE"
    dictRoles.Add "LIEFERANT", "LIEFERANT"
    dictRoles.Add "UNBEKANNT", "UNBEKANNT"
    
    lastRowEK = ws.Cells(ws.Rows.count, EK_COL_ROLE).End(xlUp).Row
    
    If lastRowEK >= EK_START_ROW Then
        For r = EK_START_ROW To lastRowEK
            roleValue = UCase(Trim(ws.Cells(r, EK_COL_ROLE).value))
            If roleValue <> "" Then
                If Not dictRoles.Exists(roleValue) Then
                    dictRoles.Add roleValue, roleValue
                End If
            End If
        Next r
    End If
    
    ws.Range("AD4:AD100").ClearContents
    
    nextRow = 4
    For Each key In dictRoles.Keys
        ws.Cells(nextRow, DATA_COL_DD_ENTITYROLE).value = key
        nextRow = nextRow + 1
    Next key
    
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' RefreshTableFormattingAfterChange
'--------------------------------------------------------------------------------
Public Sub RefreshTableFormattingAfterChange(ByRef ws As Worksheet, _
                                              ByVal changedCol As Long, _
                                              ByVal changedRow As Long)
    
    On Error Resume Next
    
    If changedCol = 2 Or changedCol = 4 Or changedCol = 6 Or changedCol = 8 Then
        Call FormatSingleColumn(ws, changedCol, changedRow)
        Exit Sub
    End If
    
    If changedCol >= 26 And changedCol <= 34 Then
        Call FormatHelperColumnIndependent(ws, changedCol)
        Exit Sub
    End If
    
    On Error GoTo 0
End Sub

