Attribute VB_Name = "mod_Formatierung"
Option Explicit

' ***************************************************************
' MODUL: mod_Formatierung
' ZWECK: Formatierung und DropDown-Listen-Verwaltung
' VERSION: 2.8 - 07.02.2026
' FIX: Inside-Borders, Z/AA zentriert, Cleanup unter lastRow
' ***************************************************************

Private Const ZEBRA_COLOR_1 As Long = &HFFFFFF  ' Weiss
Private Const ZEBRA_COLOR_2 As Long = &HDEE5E3  ' Hellgrau

' ===============================================================
' NEU: Zentriert ALLE Zellen auf ALLEN Blaettern vertikal
' ===============================================================
Public Sub ZentriereAlleZellenVertikal()
    
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error Resume Next
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect PASSWORD:=PASSWORD
        ws.Cells.VerticalAlignment = xlCenter
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Next ws
    
    On Error GoTo 0
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

' ===============================================================
' NEU: Wird aufgerufen wenn ein neues Blatt erstellt wird
' ===============================================================
Public Sub FormatiereNeuesBlatt(ByVal ws As Worksheet)
    
    On Error Resume Next
    
    ws.Unprotect PASSWORD:=PASSWORD
    ws.Cells.VerticalAlignment = xlCenter
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    On Error GoTo 0
    
End Sub

' ===============================================================
' HAUPTPROZEDUR: Formatiert ALLE relevanten Tabellen neu
' ===============================================================
Public Sub Formatiere_Alle_Tabellen_Neu()
    
    Dim wsD As Worksheet
    Dim wsBK As Worksheet
    Dim wsM As Worksheet
    Dim ws As Worksheet
    Dim lastRowD As Long
    Dim lastRowBK As Long
    Dim euroFormat As String
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    euroFormat = "#,##0.00 " & ChrW(8364)
    
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect PASSWORD:=PASSWORD
        ws.Cells.VerticalAlignment = xlCenter
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        On Error GoTo ErrorHandler
    Next ws
    
    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo ErrorHandler
    
    If Not wsD Is Nothing Then
        On Error Resume Next
        wsD.Unprotect PASSWORD:=PASSWORD
        On Error GoTo ErrorHandler
        
        Call FormatiereAlleDatenSpalten(wsD)
        Call FormatiereKategorieTabelle(wsD)
        Call FormatiereEntityKeyTabelleKomplett(wsD)
        Call AktualisiereKategorieDropdownListen(wsD)
        Call SortiereKategorieTabelle(wsD)
        Call SortiereEntityKeyTabelle(wsD)
        
        wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo ErrorHandler
    
    If Not wsBK Is Nothing Then
        On Error Resume Next
        wsBK.Unprotect PASSWORD:=PASSWORD
        On Error GoTo ErrorHandler
        
        lastRowBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
        If lastRowBK < BK_START_ROW Then lastRowBK = BK_START_ROW
        
        With wsBK.Range(wsBK.Cells(BK_START_ROW, BK_COL_BEMERKUNG), _
                        wsBK.Cells(lastRowBK, BK_COL_BEMERKUNG))
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        
        wsBK.Rows(BK_START_ROW & ":" & lastRowBK).AutoFit
        
        wsBK.Range(wsBK.Cells(BK_START_ROW, BK_COL_BETRAG), _
                   wsBK.Cells(lastRowBK, BK_COL_BETRAG)).NumberFormat = euroFormat
        
        wsBK.Range(wsBK.Cells(BK_START_ROW, BK_COL_MITGL_BEITR), _
                   wsBK.Cells(lastRowBK, BK_COL_AUSZAHL_KASSE)).NumberFormat = euroFormat
        
        wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    
    On Error Resume Next
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    On Error GoTo ErrorHandler
    
    If Not wsM Is Nothing Then
        On Error Resume Next
        wsM.Unprotect PASSWORD:=PASSWORD
        wsM.Cells.VerticalAlignment = xlCenter
        wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        On Error GoTo ErrorHandler
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not wsD Is Nothing Then wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    If Not wsBK Is Nothing Then wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    If Not wsM Is Nothing Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Debug.Print "Fehler in Formatiere_Alle_Tabellen_Neu: " & Err.Description
End Sub

' ===============================================================
' HAUPTPROZEDUR: Formatiert das gesamte Daten-Blatt
' ===============================================================
Public Sub FormatiereBlattDaten()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    ws.Cells.VerticalAlignment = xlCenter
    
    Call FormatiereAlleDatenSpalten(ws)
    Call FormatiereKategorieTabelle(ws)
    Call FormatiereEntityKeyTabelleKomplett(ws)
    Call AktualisiereKategorieDropdownListen(ws)
    Call SortiereKategorieTabelle(ws)
    Call SortiereEntityKeyTabelle(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Formatierung des Daten-Blatts abgeschlossen!" & vbCrLf & vbCrLf & _
           "- Alle Zellen vertikal zentriert" & vbCrLf & _
           "- Alle Spalten mit Zebra-Formatierung" & vbCrLf & _
           "- Kategorie-Tabelle formatiert und sortiert" & vbCrLf & _
           "- EntityKey-Tabelle formatiert und sortiert" & vbCrLf & _
           "- DropDown-Listen aktualisiert", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler bei der Formatierung: " & Err.Description, vbCritical
End Sub

' ===============================================================
' NEU: Formatiert ALLE Einzel-Spalten auf Daten-Blatt
' Spalten: B, D, F, H, Z, AA, AB, AC, AD, AE, AF, AG, AH
' ===============================================================
Private Sub FormatiereAlleDatenSpalten(ByRef ws As Worksheet)
    
    ' Einzelspalten mit Zebra
    Call FormatiereSingleSpalte(ws, 2, True)   ' Spalte B - Vereinsfunktionen
    Call FormatiereSingleSpalte(ws, 4, True)   ' Spalte D - Anredeformen
    Call FormatiereSingleSpalte(ws, 6, True)   ' Spalte F - Parzelle
    Call FormatiereSingleSpalte(ws, 8, True)   ' Spalte H - Seite
    
    ' Helper-Spalten mit Zebra
    Call FormatiereSingleSpalte(ws, 26, True)  ' Spalte Z - Einnahme/Ausgabe
    Call FormatiereSingleSpalte(ws, 27, True)  ' Spalte AA - Prioritaet
    Call FormatiereSingleSpalte(ws, 28, True)  ' Spalte AB - Ja/Nein
    Call FormatiereSingleSpalte(ws, 29, True)  ' Spalte AC - Faelligkeit
    Call FormatiereSingleSpalte(ws, 30, True)  ' Spalte AD - EntityRole
    Call FormatiereSingleSpalte(ws, 31, True)  ' Spalte AE - Hilfszelle
    Call FormatiereSingleSpalte(ws, 32, True)  ' Spalte AF - Kat Einnahmen
    Call FormatiereSingleSpalte(ws, 33, True)  ' Spalte AG - Kat Ausgaben
    Call FormatiereSingleSpalte(ws, 34, True)  ' Spalte AH - Monat/Periode
    
End Sub

' ===============================================================
' NEU: Formatiert eine einzelne Spalte mit Zebra + Rahmen
' FIX: Inside-Borders + Zentrierung Z/AA
' ===============================================================
Private Sub FormatiereSingleSpalte(ByRef ws As Worksheet, ByVal colIndex As Long, ByVal mitZebra As Boolean)
    
    Dim lastRow As Long
    Dim rng As Range
    Dim r As Long
    
    ' Letzte Zeile mit Daten ermitteln
    lastRow = ws.Cells(ws.Rows.count, colIndex).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    
    ' Datenbereich festlegen
    Set rng = ws.Range(ws.Cells(DATA_START_ROW, colIndex), ws.Cells(lastRow, colIndex))
    
    ' Formatierung loeschen
    rng.Interior.ColorIndex = xlNone
    rng.Borders.LineStyle = xlNone
    
    ' Vertikale Zentrierung
    rng.VerticalAlignment = xlCenter
    
    ' Horizontale Zentrierung fuer Spalte Z (26) und AA (27)
    If colIndex = 26 Or colIndex = 27 Then
        rng.HorizontalAlignment = xlCenter
    End If
    
    ' Zebra-Formatierung anwenden
    If mitZebra Then
        For r = DATA_START_ROW To lastRow
            If (r - DATA_START_ROW) Mod 2 = 0 Then
                ws.Cells(r, colIndex).Interior.color = ZEBRA_COLOR_1
            Else
                ws.Cells(r, colIndex).Interior.color = ZEBRA_COLOR_2
            End If
        Next r
    End If
    
    ' Aussere Rahmen setzen
    With rng
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeTop).color = RGB(0, 0, 0)
        
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).color = RGB(0, 0, 0)
        
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeLeft).color = RGB(0, 0, 0)
        
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeRight).color = RGB(0, 0, 0)
    End With
    
    ' Innere horizontale Trennlinien setzen (zwischen jeder Zelle)
    If lastRow > DATA_START_ROW Then
        With rng.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = RGB(0, 0, 0)
        End With
    End If
    
    ' Spaltenbreite anpassen
    ws.Columns(colIndex).AutoFit
    
End Sub

' ===============================================================
' PUBLIC WRAPPER FUER TABELLE8: Formatiert Kategorie-Tabelle komplett
' ===============================================================
Public Sub FormatKategorieTableComplete(ByRef ws As Worksheet)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    Call FormatiereKategorieTabelle(ws)
    Call SortiereKategorieTabelle(ws)
    
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
End Sub

' ===============================================================
' PUBLIC WRAPPER FUER TABELLE8: Formatiert EntityKey-Tabelle komplett
' ===============================================================
Public Sub FormatEntityKeyTableComplete(ByRef ws As Worksheet)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    Call FormatiereEntityKeyTabelleKomplett(ws)
    Call SortiereEntityKeyTabelle(ws)
    
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
End Sub

' ===============================================================
' PUBLIC WRAPPER FUER TABELLE8: Formatiert eine einzelne Spalte
' FIX: Raeumt Formatierung unterhalb lastRow auf (geloeschte Eintraege)
' ===============================================================
Public Sub FormatSingleColumnComplete(ByRef ws As Worksheet, ByVal colIndex As Long)
    
    Dim lastRow As Long
    Dim cleanEnd As Long
    Dim rngClean As Range
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    Call FormatiereSingleSpalte(ws, colIndex, True)
    
    ' Aufraumen: Formatierung unterhalb der letzten Datenzeile entfernen
    lastRow = ws.Cells(ws.Rows.count, colIndex).End(xlUp).Row
    If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW - 1
    
    cleanEnd = lastRow + 50
    If cleanEnd > ws.Rows.count Then cleanEnd = ws.Rows.count
    
    If lastRow + 1 <= cleanEnd Then
        Set rngClean = ws.Range(ws.Cells(lastRow + 1, colIndex), ws.Cells(cleanEnd, colIndex))
        rngClean.Interior.ColorIndex = xlNone
        rngClean.Borders.LineStyle = xlNone
    End If
    
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
End Sub




'--- Ende Teil 1 von 3 ---
'--- Anfang Teil 2 von 3 ---




' ===============================================================
' KATEGORIE-TABELLE FORMATIEREN (J-P) mit ZEBRA
' ===============================================================
Public Sub FormatiereKategorieTabelle(Optional ByRef ws As Worksheet = Nothing)
    
    Dim lastRow As Long
    Dim rngTable As Range
    Dim r As Long
    Dim einAusWert As String
    
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    
    Set rngTable = ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), _
                            ws.Cells(lastRow, DATA_CAT_COL_END))
    
    ' Formatierung loeschen
    rngTable.Interior.ColorIndex = xlNone
    rngTable.Borders.LineStyle = xlNone
    
    ' Zebra-Formatierung fuer gesamte Zeile J-P
    For r = DATA_START_ROW To lastRow
        If (r - DATA_START_ROW) Mod 2 = 0 Then
            ws.Range(ws.Cells(r, DATA_CAT_COL_START), ws.Cells(r, DATA_CAT_COL_END)).Interior.color = ZEBRA_COLOR_1
        Else
            ws.Range(ws.Cells(r, DATA_CAT_COL_START), ws.Cells(r, DATA_CAT_COL_END)).Interior.color = ZEBRA_COLOR_2
        End If
    Next r
    
    ' Rahmen setzen
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    rngTable.VerticalAlignment = xlCenter
    
    ' Horizontale Ausrichtungen
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KATEGORIE), _
             ws.Cells(lastRow, DATA_CAT_COL_KATEGORIE)).HorizontalAlignment = xlLeft
    
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_EINAUS), _
             ws.Cells(lastRow, DATA_CAT_COL_EINAUS)).HorizontalAlignment = xlCenter
    
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KEYWORD), _
             ws.Cells(lastRow, DATA_CAT_COL_KEYWORD)).HorizontalAlignment = xlLeft
    
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_PRIORITAET), _
             ws.Cells(lastRow, DATA_CAT_COL_PRIORITAET)).HorizontalAlignment = xlCenter
    
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_ZIELSPALTE), _
             ws.Cells(lastRow, DATA_CAT_COL_ZIELSPALTE)).HorizontalAlignment = xlLeft
    
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_FAELLIGKEIT), _
             ws.Cells(lastRow, DATA_CAT_COL_FAELLIGKEIT)).HorizontalAlignment = xlLeft
    
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KOMMENTAR), _
             ws.Cells(lastRow, DATA_CAT_COL_KOMMENTAR)).HorizontalAlignment = xlLeft
    
    ' Dropdowns setzen
    For r = DATA_START_ROW To lastRow
        einAusWert = UCase(Trim(ws.Cells(r, DATA_CAT_COL_EINAUS).value))
        Call SetzeZielspalteDropdown(ws, r, einAusWert)
    Next r
    
    ' Spaltenbreite anpassen
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), _
             ws.Cells(lastRow, DATA_CAT_COL_END)).EntireColumn.AutoFit
    
End Sub

' ===============================================================
' NEU: Sortiert die Kategorie-Tabelle nach Spalte J (A-Z)
' ===============================================================
Public Sub SortiereKategorieTabelle(Optional ByRef ws As Worksheet = Nothing)
    
    Dim lastRow As Long
    Dim sortRange As Range
    
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    
    Set sortRange = ws.Range(ws.Cells(DATA_START_ROW - 1, DATA_CAT_COL_START), _
                             ws.Cells(lastRow, DATA_CAT_COL_END))
    
    On Error Resume Next
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KATEGORIE), _
                                      ws.Cells(lastRow, DATA_CAT_COL_KATEGORIE)), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SetRange sortRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    On Error GoTo 0
    
End Sub

' ===============================================================
' NEU: Sortiert die EntityKey-Tabelle
' Sortierung: Parzelle 1-14, dann EX-, VERS-, BANK-, Rest
' NUR Spalten R bis X - keine anderen Spalten!
' ===============================================================
Public Sub SortiereEntityKeyTabelle(Optional ByRef ws As Worksheet = Nothing)
    
    Dim lastRow As Long
    Dim r As Long
    Dim helperCol As Long
    Dim parzelleWert As String
    Dim entityKey As String
    Dim sortOrder As Long
    Dim sortRange As Range
    
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    lastRow = ws.Cells(ws.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    If lastRow < EK_START_ROW Then Exit Sub
    
    Dim arrData() As Variant
    Dim arrSorted() As Variant
    Dim i As Long, j As Long
    Dim numRows As Long
    Dim swap As Boolean
    Dim tempRow As Variant
    
    numRows = lastRow - EK_START_ROW + 1
    
    If numRows < 1 Then Exit Sub
    
    ' Daten aus R-X in Array einlesen (7 Spalten: R, S, T, U, V, W, X)
    ReDim arrData(1 To numRows, 1 To 7)
    
    For r = EK_START_ROW To lastRow
        arrData(r - EK_START_ROW + 1, 1) = ws.Cells(r, EK_COL_ENTITYKEY).value       ' R
        arrData(r - EK_START_ROW + 1, 2) = ws.Cells(r, EK_COL_IBAN).value            ' S
        arrData(r - EK_START_ROW + 1, 3) = ws.Cells(r, EK_COL_KONTONAME).value       ' T
        arrData(r - EK_START_ROW + 1, 4) = ws.Cells(r, EK_COL_ZUORDNUNG).value       ' U
        arrData(r - EK_START_ROW + 1, 5) = ws.Cells(r, EK_COL_PARZELLE).value        ' V
        arrData(r - EK_START_ROW + 1, 6) = ws.Cells(r, EK_COL_ROLE).value            ' W
        arrData(r - EK_START_ROW + 1, 7) = ws.Cells(r, EK_COL_DEBUG).value           ' X
    Next r
    
    ' Bubble Sort mit benutzerdefinierter Sortierlogik
    For i = 1 To numRows - 1
        swap = False
        For j = 1 To numRows - i
            ' Vergleiche Zeile j mit Zeile j+1
            If VergleicheEntityKeyZeilen(arrData(j, 1), arrData(j, 5), arrData(j + 1, 1), arrData(j + 1, 5)) > 0 Then
                ' Tausche Zeilen
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
        If Not swap Then Exit For ' Bereits sortiert
    Next i
    
    ' Sortierte Daten zurueckschreiben NUR in Spalten R-X
    For r = EK_START_ROW To lastRow
        ws.Cells(r, EK_COL_ENTITYKEY).value = arrData(r - EK_START_ROW + 1, 1)
        ws.Cells(r, EK_COL_IBAN).value = arrData(r - EK_START_ROW + 1, 2)
        ws.Cells(r, EK_COL_KONTONAME).value = arrData(r - EK_START_ROW + 1, 3)
        ws.Cells(r, EK_COL_ZUORDNUNG).value = arrData(r - EK_START_ROW + 1, 4)
        ws.Cells(r, EK_COL_PARZELLE).value = arrData(r - EK_START_ROW + 1, 5)
        ws.Cells(r, EK_COL_ROLE).value = arrData(r - EK_START_ROW + 1, 6)
        ws.Cells(r, EK_COL_DEBUG).value = arrData(r - EK_START_ROW + 1, 7)
    Next r
    
End Sub

' ===============================================================
' HILFSFUNKTION: Vergleicht zwei EntityKey-Zeilen fuer Sortierung
' Rueckgabe: <0 wenn zeile1 < zeile2, 0 wenn gleich, >0 wenn zeile1 > zeile2
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
    
    ' Sortier-Order fuer Zeile 1 bestimmen
    If IsNumeric(parzelleStr1) And parzelleStr1 <> "" Then
        order1 = CLng(parzelleStr1)  ' Parzellen 1-14
    ElseIf Left(UCase(entityStr1), 3) = "EX-" Then
        order1 = 100  ' Ehemalige Mitglieder
    ElseIf Left(UCase(entityStr1), 5) = "VERS-" Then
        order1 = 200  ' Versorger
    ElseIf Left(UCase(entityStr1), 5) = "BANK-" Then
        order1 = 300  ' Banken
    Else
        order1 = 400  ' Rest
    End If
    
    ' Sortier-Order fuer Zeile 2 bestimmen
    If IsNumeric(parzelleStr2) And parzelleStr2 <> "" Then
        order2 = CLng(parzelleStr2)  ' Parzellen 1-14
    ElseIf Left(UCase(entityStr2), 3) = "EX-" Then
        order2 = 100  ' Ehemalige Mitglieder
    ElseIf Left(UCase(entityStr2), 5) = "VERS-" Then
        order2 = 200  ' Versorger
    ElseIf Left(UCase(entityStr2), 5) = "BANK-" Then
        order2 = 300  ' Banken
    Else
        order2 = 400  ' Rest
    End If
    
    ' Vergleich
    If order1 < order2 Then
        VergleicheEntityKeyZeilen = -1
    ElseIf order1 > order2 Then
        VergleicheEntityKeyZeilen = 1
    Else
        VergleicheEntityKeyZeilen = 0
    End If
    
End Function

' ===============================================================
' ZIELSPALTE-DROPDOWN SETZEN (abhaengig von E/A)
' ===============================================================
Private Sub SetzeZielspalteDropdown(ByRef ws As Worksheet, ByVal zeile As Long, ByVal einAus As String)
    
    Dim dropdownSource As String
    
    On Error Resume Next
    ws.Cells(zeile, DATA_CAT_COL_ZIELSPALTE).Validation.Delete
    On Error GoTo 0
    
    Select Case einAus
        Case "E"
            dropdownSource = "=" & WS_BANKKONTO & "!$M$27:$S$27"
        Case "A"
            dropdownSource = "=" & WS_BANKKONTO & "!$T$27:$Z$27"
        Case Else
            dropdownSource = "=" & WS_BANKKONTO & "!$M$27:$Z$27"
    End Select
    
    On Error Resume Next
    With ws.Cells(zeile, DATA_CAT_COL_ZIELSPALTE).Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:=dropdownSource
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    On Error GoTo 0
    
End Sub




'--- Ende Teil 2 von 3 ---
'--- Anfang Teil 3 von 3 ---




' ===============================================================
' DROPDOWN-LISTEN FUER KATEGORIEN AKTUALISIEREN (AF + AG + AH)
' ===============================================================
Public Sub AktualisiereKategorieDropdownListen(Optional ByRef ws As Worksheet = Nothing)
    
    Dim lastRow As Long
    Dim r As Long
    Dim kategorie As String
    Dim einAus As String
    Dim dictEinnahmen As Object
    Dim dictAusgaben As Object
    Dim key As Variant
    Dim nextRowE As Long
    Dim nextRowA As Long
    
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    Set dictEinnahmen = CreateObject("Scripting.Dictionary")
    Set dictAusgaben = CreateObject("Scripting.Dictionary")
    
    lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    
    For r = DATA_START_ROW To lastRow
        kategorie = Trim(ws.Cells(r, DATA_CAT_COL_KATEGORIE).value)
        einAus = UCase(Trim(ws.Cells(r, DATA_CAT_COL_EINAUS).value))
        
        If kategorie <> "" Then
            If einAus = "E" Then
                If Not dictEinnahmen.Exists(kategorie) Then
                    dictEinnahmen.Add kategorie, kategorie
                End If
            ElseIf einAus = "A" Then
                If Not dictAusgaben.Exists(kategorie) Then
                    dictAusgaben.Add kategorie, kategorie
                End If
            End If
        End If
    Next r
    
    On Error Resume Next
    ws.Range("AF4:AF1000").ClearContents
    ws.Range("AG4:AG1000").ClearContents
    ws.Range("AH4:AH1000").ClearContents
    On Error GoTo 0
    
    nextRowE = 4
    For Each key In dictEinnahmen.Keys
        ws.Cells(nextRowE, DATA_COL_KAT_EINNAHMEN).value = key
        nextRowE = nextRowE + 1
    Next key
    
    nextRowA = 4
    For Each key In dictAusgaben.Keys
        ws.Cells(nextRowA, DATA_COL_KAT_AUSGABEN).value = key
        nextRowA = nextRowA + 1
    Next key
    
    ws.Cells(4, DATA_COL_MONAT_PERIODE).value = "Januar"
    ws.Cells(5, DATA_COL_MONAT_PERIODE).value = "Februar"
    ws.Cells(6, DATA_COL_MONAT_PERIODE).value = "Maerz"
    ws.Cells(7, DATA_COL_MONAT_PERIODE).value = "April"
    ws.Cells(8, DATA_COL_MONAT_PERIODE).value = "Mai"
    ws.Cells(9, DATA_COL_MONAT_PERIODE).value = "Juni"
    ws.Cells(10, DATA_COL_MONAT_PERIODE).value = "Juli"
    ws.Cells(11, DATA_COL_MONAT_PERIODE).value = "August"
    ws.Cells(12, DATA_COL_MONAT_PERIODE).value = "September"
    ws.Cells(13, DATA_COL_MONAT_PERIODE).value = "Oktober"
    ws.Cells(14, DATA_COL_MONAT_PERIODE).value = "November"
    ws.Cells(15, DATA_COL_MONAT_PERIODE).value = "Dezember"
    
    Call ErstelleKategorieNamedRanges(ws, nextRowE - 1, nextRowA - 1)
    
    ' Formatiere Helper-Spalten nach Update
    Call FormatiereSingleSpalte(ws, 32, True)  ' AF
    Call FormatiereSingleSpalte(ws, 33, True)  ' AG
    Call FormatiereSingleSpalte(ws, 34, True)  ' AH
    
End Sub

' ===============================================================
' NAMED RANGES FUER KATEGORIEN ERSTELLEN
' ===============================================================
Private Sub ErstelleKategorieNamedRanges(ByRef ws As Worksheet, ByVal lastRowE As Long, ByVal lastRowA As Long)
    
    On Error Resume Next
    
    ThisWorkbook.Names("lst_KategorienEinnahmen").Delete
    ThisWorkbook.Names("lst_KategorienAusgaben").Delete
    ThisWorkbook.Names("lst_MonatPeriode").Delete
    
    If lastRowE >= 4 Then
        ThisWorkbook.Names.Add Name:="lst_KategorienEinnahmen", _
            RefersTo:="=" & ws.Name & "!$AF$4:$AF$" & lastRowE
    Else
        ThisWorkbook.Names.Add Name:="lst_KategorienEinnahmen", _
            RefersTo:="=" & ws.Name & "!$AF$4"
    End If
    
    If lastRowA >= 4 Then
        ThisWorkbook.Names.Add Name:="lst_KategorienAusgaben", _
            RefersTo:="=" & ws.Name & "!$AG$4:$AG$" & lastRowA
    Else
        ThisWorkbook.Names.Add Name:="lst_KategorienAusgaben", _
            RefersTo:="=" & ws.Name & "!$AG$4"
    End If
    
    ThisWorkbook.Names.Add Name:="lst_MonatPeriode", _
        RefersTo:="=" & ws.Name & "!$AH$4:$AH$15"
    
    On Error GoTo 0
    
End Sub

' ===============================================================
' WRAPPER: Formatiert die EntityKey-Tabelle komplett
' ===============================================================
Private Sub FormatiereEntityKeyTabelleKomplett(ByRef ws As Worksheet)
    
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    
    If lastRow >= EK_START_ROW Then
        Call FormatiereEntityKeyTabelle(ws, lastRow)
    End If
    
End Sub

' ===============================================================
' HILFSPROZEDUR: Formatiert die EntityKey-Tabelle
' R-T mit Zebra, U-X nur Rahmen (Ampel-Farben bleiben)
' ===============================================================
Private Sub FormatiereEntityKeyTabelle(ByRef ws As Worksheet, ByVal lastRow As Long)
    
    Dim rngTable As Range
    Dim rngZebra As Range
    Dim rngBorderOnly As Range
    Dim r As Long
    Dim currentRole As String
    
    If lastRow < EK_START_ROW Then Exit Sub
    
    Set rngTable = ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                            ws.Cells(lastRow, EK_COL_DEBUG))
    
    ' Rahmen fuer gesamte Tabelle
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    rngTable.VerticalAlignment = xlCenter
    
    ' Spaltenformatierung
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                  ws.Cells(lastRow, EK_COL_ENTITYKEY))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ENTITYKEY).ColumnWidth = 11
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_IBAN), _
                  ws.Cells(lastRow, EK_COL_IBAN))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_IBAN).AutoFit
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_KONTONAME), _
                  ws.Cells(lastRow, EK_COL_KONTONAME))
        .WrapText = True
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_KONTONAME).ColumnWidth = 36
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ZUORDNUNG), _
                  ws.Cells(lastRow, EK_COL_ZUORDNUNG))
        .WrapText = True
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ZUORDNUNG).ColumnWidth = 28
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_PARZELLE), _
                  ws.Cells(lastRow, EK_COL_PARZELLE))
        .WrapText = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns(EK_COL_PARZELLE).ColumnWidth = 9
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ROLE), _
                  ws.Cells(lastRow, EK_COL_ROLE))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ROLE).AutoFit
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_DEBUG), _
                  ws.Cells(lastRow, EK_COL_DEBUG))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_DEBUG).ColumnWidth = 42
    
    ' Zellschutz fuer R-T
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
             ws.Cells(lastRow, EK_COL_KONTONAME)).Locked = True
    
    ' Zebra-Formatierung NUR fuer R-T, U-X bleiben ohne Zebra (Ampel)
    For r = EK_START_ROW To lastRow
        currentRole = Trim(ws.Cells(r, EK_COL_ROLE).value)
        
        Call SetzeZellschutzFuerZeile(ws, r, currentRole)
        
        ' Zebra nur fuer R-T (EntityKey, IBAN, Kontoname)
        Set rngZebra = ws.Range(ws.Cells(r, EK_COL_ENTITYKEY), ws.Cells(r, EK_COL_KONTONAME))
        
        If (r - EK_START_ROW) Mod 2 = 0 Then
            rngZebra.Interior.color = ZEBRA_COLOR_1
        Else
            rngZebra.Interior.color = ZEBRA_COLOR_2
        End If
        
        ' U-X: NUR Rahmen, KEINE Zebra-Formatierung (Ampel bleibt)
        Set rngBorderOnly = ws.Range(ws.Cells(r, EK_COL_ZUORDNUNG), ws.Cells(r, EK_COL_DEBUG))
        ' Keine Interior-Formatierung fuer U-X
    Next r
    
    ws.Rows(EK_START_ROW & ":" & lastRow).AutoFit
    
End Sub

' ===============================================================
' HILFSPROZEDUR: Setzt Zellschutz basierend auf EntityRole
' ===============================================================
Private Sub SetzeZellschutzFuerZeile(ByRef ws As Worksheet, ByVal zeile As Long, ByVal currentRole As String)
    
    Dim rngEditierbar As Range
    Dim rngGesperrt As Range
    
    On Error Resume Next
    
    Set rngGesperrt = ws.Range(ws.Cells(zeile, EK_COL_ENTITYKEY), ws.Cells(zeile, EK_COL_KONTONAME))
    rngGesperrt.Locked = True
    
    Set rngEditierbar = ws.Range(ws.Cells(zeile, EK_COL_ZUORDNUNG), ws.Cells(zeile, EK_COL_DEBUG))
    
    Select Case UCase(Trim(currentRole))
        Case "MITGLIED"
            ws.Cells(zeile, EK_COL_ZUORDNUNG).Locked = True
            ws.Cells(zeile, EK_COL_PARZELLE).Locked = True
            ws.Cells(zeile, EK_COL_ROLE).Locked = True
            ws.Cells(zeile, EK_COL_DEBUG).Locked = False
            
        Case "UNBEKANNT", ""
            rngEditierbar.Locked = False
            
        Case Else
            ws.Cells(zeile, EK_COL_ZUORDNUNG).Locked = False
            ws.Cells(zeile, EK_COL_PARZELLE).Locked = True
            ws.Cells(zeile, EK_COL_ROLE).Locked = True
            ws.Cells(zeile, EK_COL_DEBUG).Locked = False
    End Select
    
    On Error GoTo 0
    
End Sub

' ===============================================================
' BANKKONTO-BLATT FORMATIEREN
' ===============================================================
Public Sub FormatiereBlattBankkonto()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim euroFormat As String
    
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    euroFormat = "#,##0.00 " & ChrW(8364)
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    ws.Cells.VerticalAlignment = xlCenter
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then lastRow = BK_START_ROW
    
    With ws.Range(ws.Cells(BK_START_ROW, BK_COL_BEMERKUNG), _
                  ws.Cells(lastRow, BK_COL_BEMERKUNG))
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    
    ws.Rows(BK_START_ROW & ":" & lastRow).AutoFit
    
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_BETRAG), _
             ws.Cells(lastRow, BK_COL_BETRAG)).NumberFormat = euroFormat
    
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_MITGL_BEITR), _
             ws.Cells(lastRow, BK_COL_AUSZAHL_KASSE)).NumberFormat = euroFormat
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.ScreenUpdating = True
    
    MsgBox "Formatierung des Bankkonto-Blatts abgeschlossen!" & vbCrLf & vbCrLf & _
           "- Alle Zellen vertikal zentriert" & vbCrLf & _
           "- Spalte L mit Textumbruch" & vbCrLf & _
           "- Zeilenhoehe angepasst" & vbCrLf & _
           "- Waehrung mit Euro-Zeichen", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler bei der Formatierung: " & Err.Description, vbCritical
End Sub

' ===============================================================
' HILFSFUNKTION: Prueft ob Named Range existiert
' ===============================================================
Public Function NamedRangeExists(ByVal rangeName As String) As Boolean
    Dim nm As Name
    NamedRangeExists = False
    
    On Error Resume Next
    Set nm = ThisWorkbook.Names(rangeName)
    If Not nm Is Nothing Then
        NamedRangeExists = True
    End If
    On Error GoTo 0
End Function

' ===============================================================
' WORKSHEET_CHANGE HELPER: Aktualisiert Listen bei Aenderung
' ===============================================================
Public Sub OnKategorieChange(ByVal Target As Range)
    
    Dim ws As Worksheet
    Set ws = Target.Worksheet
    
    If Target.Column = DATA_CAT_COL_KATEGORIE Or Target.Column = DATA_CAT_COL_EINAUS Then
        Call AktualisiereKategorieDropdownListen(ws)
    End If
    
    If Target.Column = DATA_CAT_COL_EINAUS Then
        Dim einAus As String
        einAus = UCase(Trim(Target.value))
        Call SetzeZielspalteDropdown(ws, Target.Row, einAus)
    End If
    
End Sub

' ===============================================================
' NEU: WORKSHEET_CHANGE HANDLER fuer dynamische Formatierung
' Wird von Tabelle8 (Daten) Worksheet_Change aufgerufen
' ===============================================================
Public Sub OnDatenChange(ByVal Target As Range, ByVal ws As Worksheet)
    
    On Error GoTo ErrorHandler
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Einzelspalten B, D, F, H
    If Not Intersect(Target, ws.Columns(2)) Is Nothing Then
        Call FormatiereSingleSpalte(ws, 2, True)
    End If
    
    If Not Intersect(Target, ws.Columns(4)) Is Nothing Then
        Call FormatiereSingleSpalte(ws, 4, True)
    End If
    
    If Not Intersect(Target, ws.Columns(6)) Is Nothing Then
        Call FormatiereSingleSpalte(ws, 6, True)
    End If
    
    If Not Intersect(Target, ws.Columns(8)) Is Nothing Then
        Call FormatiereSingleSpalte(ws, 8, True)
    End If
    
    ' Kategorie-Tabelle J-P
    If Not Intersect(Target, ws.Range("J:P")) Is Nothing Then
        Call FormatiereKategorieTabelle(ws)
        Call SortiereKategorieTabelle(ws)
        Call OnKategorieChange(Target)
    End If
    
    ' EntityKey-Tabelle R-X
    If Not Intersect(Target, ws.Range("R:X")) Is Nothing Then
        Call FormatiereEntityKeyTabelleKomplett(ws)
        Call SortiereEntityKeyTabelle(ws)
    End If
    
    ' Helper-Spalten Z-AH
    If Not Intersect(Target, ws.Columns(26)) Is Nothing Then
        Call FormatiereSingleSpalte(ws, 26, True)
    End If
    
    If Not Intersect(Target, ws.Columns(27)) Is Nothing Then
        Call FormatiereSingleSpalte(ws, 27, True)
    End If
    
    If Not Intersect(Target, ws.Columns(28)) Is Nothing Then
        Call FormatiereSingleSpalte(ws, 28, True)
    End If
    
    If Not Intersect(Target, ws.Columns(29)) Is Nothing Then
        Call FormatiereSingleSpalte(ws, 29, True)
    End If
    
    If Not Intersect(Target, ws.Columns(30)) Is Nothing Then
        Call FormatiereSingleSpalte(ws, 30, True)
    End If
    
    If Not Intersect(Target, ws.Columns(31)) Is Nothing Then
        Call FormatiereSingleSpalte(ws, 31, True)
    End If
    
    If Not Intersect(Target, ws.Columns(32)) Is Nothing Then
        Call FormatiereSingleSpalte(ws, 32, True)
    End If
    
    If Not Intersect(Target, ws.Columns(33)) Is Nothing Then
        Call FormatiereSingleSpalte(ws, 33, True)
    End If
    
    If Not Intersect(Target, ws.Columns(34)) Is Nothing Then
        Call FormatiereSingleSpalte(ws, 34, True)
    End If
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Debug.Print "Fehler in OnDatenChange: " & Err.Description
    End If
    
End Sub

