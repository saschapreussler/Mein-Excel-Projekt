Attribute VB_Name = "mod_Formatierung"
Option Explicit

' ***************************************************************
' MODUL: mod_Formatierung
' ZWECK: Formatierung und DropDown-Listen-Verwaltung
' VERSION: 5.3 - 06.02.2026
' NEU: BlendeDatenSpaltenAus (D-I, Z-AB, AE-AH automatisch ausblenden)
'      OnDatenChange formatiert bei JEDER Aenderung ALLE Spalten
'      OnKategorieChange: Spalte K Konsistenzpruefung, Spalte L Duplikat
'      FormatiereKategorieTabelle: Bereinigung unterhalb, K Breite 12
'      FormatiereSingleSpalte: Bereinigung unterhalb
'      FormatiereEntityKeyTabelle: Spalte X Breite 65
' ***************************************************************

Private Const ZEBRA_COLOR_1 As Long = &HFFFFFF  ' Weiss
Private Const ZEBRA_COLOR_2 As Long = &HDEE5E3  ' Hellgrau

' ===============================================================
' NEU v5.3: Blendet Hilfsspalten aus (D-I, Z-AB, AE-AH)
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
    
    ' Spalten D-I (4-9) ausblenden
    ws.Range("D:I").EntireColumn.Hidden = True
    
    ' Spalten Z-AB (26-28) ausblenden
    ws.Range("Z:AB").EntireColumn.Hidden = True
    
    ' Spalten AE-AH (31-34) ausblenden
    ws.Range("AE:AH").EntireColumn.Hidden = True
    
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
End Sub

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
' FIX v5.3: BlendeDatenSpaltenAus am Ende
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
        Call EntspeerreEditierbareSpalten(wsD)
        
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
    
    ' NEU v5.3: Hilfsspalten ausblenden
    Call BlendeDatenSpaltenAus
    
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
' FIX v5.3: BlendeDatenSpaltenAus vor MsgBox
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
    Call EntspeerreEditierbareSpalten(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' NEU v5.3: Hilfsspalten ausblenden
    Call BlendeDatenSpaltenAus
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Formatierung des Daten-Blatts abgeschlossen!" & vbCrLf & vbCrLf & _
           "- Alle Zellen vertikal zentriert" & vbCrLf & _
           "- Alle Spalten mit Zebra-Formatierung" & vbCrLf & _
           "- Kategorie-Tabelle formatiert und sortiert" & vbCrLf & _
           "- EntityKey-Tabelle formatiert und sortiert" & vbCrLf & _
           "- DropDown-Listen aktualisiert" & vbCrLf & _
           "- Editierbare Spalten und Eingabezeile entsperrt", vbInformation
    
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
' ===============================================================
Private Sub FormatiereAlleDatenSpalten(ByRef ws As Worksheet)
    
    Call FormatiereSingleSpalte(ws, 2, True)   ' Spalte B - Vereinsfunktionen
    Call FormatiereSingleSpalte(ws, 4, True)   ' Spalte D - Anredeformen
    Call FormatiereSingleSpalte(ws, 6, True)   ' Spalte F - Parzelle
    Call FormatiereSingleSpalte(ws, 8, True)   ' Spalte H - Seite
    
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
' Formatiert eine einzelne Spalte mit Zebra + Rahmen
' FIX v5.3: Bereinigt leere Zeilen unterhalb (Rahmen + Farbe entfernen)
' ===============================================================
Private Sub FormatiereSingleSpalte(ByRef ws As Worksheet, ByVal colIndex As Long, ByVal mitZebra As Boolean)
    
    Dim lastRow As Long
    Dim rng As Range
    Dim r As Long
    Dim cleanEnd As Long
    
    lastRow = ws.Cells(ws.Rows.count, colIndex).End(xlUp).Row
    
    ' Bereich UNTERHALB der belegten Zeilen bereinigen
    cleanEnd = lastRow + 50
    If cleanEnd < DATA_START_ROW + 50 Then cleanEnd = DATA_START_ROW + 50
    If lastRow < DATA_START_ROW Then
        ' Keine Daten -> alles ab DATA_START_ROW bereinigen
        ws.Range(ws.Cells(DATA_START_ROW, colIndex), ws.Cells(cleanEnd, colIndex)).Interior.ColorIndex = xlNone
        ws.Range(ws.Cells(DATA_START_ROW, colIndex), ws.Cells(cleanEnd, colIndex)).Borders.LineStyle = xlNone
        Exit Sub
    Else
        ' Unterhalb der letzten belegten Zeile bereinigen
        ws.Range(ws.Cells(lastRow + 1, colIndex), ws.Cells(cleanEnd, colIndex)).Interior.ColorIndex = xlNone
        ws.Range(ws.Cells(lastRow + 1, colIndex), ws.Cells(cleanEnd, colIndex)).Borders.LineStyle = xlNone
    End If
    
    Set rng = ws.Range(ws.Cells(DATA_START_ROW, colIndex), ws.Cells(lastRow, colIndex))
    
    rng.Interior.ColorIndex = xlNone
    rng.Borders.LineStyle = xlNone
    rng.VerticalAlignment = xlCenter
    
    If colIndex = 26 Or colIndex = 27 Then
        rng.HorizontalAlignment = xlCenter
    End If
    
    If mitZebra Then
        For r = DATA_START_ROW To lastRow
            If (r - DATA_START_ROW) Mod 2 = 0 Then
                ws.Cells(r, colIndex).Interior.color = ZEBRA_COLOR_1
            Else
                ws.Cells(r, colIndex).Interior.color = ZEBRA_COLOR_2
            End If
        Next r
    End If
    
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
    
    If lastRow > DATA_START_ROW Then
        With rng.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = RGB(0, 0, 0)
        End With
    End If
    
    ws.Columns(colIndex).AutoFit
    
End Sub

' ===============================================================
' PUBLIC WRAPPER: Formatiert Kategorie-Tabelle komplett
' ===============================================================
Public Sub FormatKategorieTableComplete(ByRef ws As Worksheet)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    Call VerdichteSpalteOhneLuecken(ws, DATA_CAT_COL_KATEGORIE, DATA_CAT_COL_START, DATA_CAT_COL_END)
    Call FormatiereKategorieTabelle(ws)
    Call SortiereKategorieTabelle(ws)
    Call EntspeerreEditierbareSpalten(ws)
    
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
End Sub

' ===============================================================
' PUBLIC WRAPPER: Formatiert EntityKey-Tabelle komplett
' FIX: Verdichtung prueft IBAN (S) statt EntityKey (R)
' ===============================================================
Public Sub FormatEntityKeyTableComplete(ByRef ws As Worksheet)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' WICHTIG: Pruefung ueber IBAN (Spalte S=19), NICHT EntityKey (R=18)
    ' Denn nach IBAN-Import ist R oft noch leer, S aber gefuellt!
    Call VerdichteSpalteOhneLuecken(ws, EK_COL_IBAN, EK_COL_ENTITYKEY, EK_COL_DEBUG)
    Call FormatiereEntityKeyTabelleKomplett(ws)
    Call SortiereEntityKeyTabelle(ws)
    Call EntspeerreEditierbareSpalten(ws)
    
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
End Sub

' ===============================================================
' PUBLIC WRAPPER: Formatiert eine einzelne Spalte
' ===============================================================
Public Sub FormatSingleColumnComplete(ByRef ws As Worksheet, ByVal colIndex As Long)
    
    Dim lastRow As Long
    Dim cleanEnd As Long
    Dim rngClean As Range
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' Luecken entfernen (Einzelspalte: Start=End=colIndex)
    Call VerdichteSpalteOhneLuecken(ws, colIndex, colIndex, colIndex)
    
    Call FormatiereSingleSpalte(ws, colIndex, True)
    
    ' Aufraumen unterhalb lastRow
    lastRow = ws.Cells(ws.Rows.count, colIndex).End(xlUp).Row
    If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW - 1
    
    cleanEnd = lastRow + 50
    If cleanEnd > ws.Rows.count Then cleanEnd = ws.Rows.count
    
    If lastRow + 1 <= cleanEnd Then
        Set rngClean = ws.Range(ws.Cells(lastRow + 1, colIndex), ws.Cells(cleanEnd, colIndex))
        rngClean.Interior.ColorIndex = xlNone
        rngClean.Borders.LineStyle = xlNone
    End If
    
    Call EntspeerreEditierbareSpalten(ws)
    
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
End Sub

' ===============================================================
' NEU: Entfernt Luecken (leere Zeilen) in einer Tabelle/Spalte
' Verschiebt alle Daten nach oben, sodass keine Leerzeilen bleiben
' checkCol = Spalte die auf Leerheit geprueft wird
' startCol/endCol = Bereich der verschoben wird
' ===============================================================
Private Sub VerdichteSpalteOhneLuecken(ByRef ws As Worksheet, ByVal checkCol As Long, _
                                       ByVal startCol As Long, ByVal endCol As Long)
    
    Dim lastRow As Long
    Dim schreibZeile As Long
    Dim leseZeile As Long
    Dim numCols As Long
    Dim col As Long
    
    ' Letzte Zeile im gesamten Bereich ermitteln (ueber alle Spalten)
    Dim maxRow As Long
    maxRow = DATA_START_ROW - 1
    For col = startCol To endCol
        lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
        If lastRow > maxRow Then maxRow = lastRow
    Next col
    
    If maxRow < DATA_START_ROW Then Exit Sub
    
    lastRow = maxRow
    numCols = endCol - startCol + 1
    
    ' Daten in Array einlesen
    Dim arrData() As Variant
    Dim arrResult() As Variant
    Dim totalRows As Long
    Dim resultCount As Long
    Dim isEmpty As Boolean
    
    totalRows = lastRow - DATA_START_ROW + 1
    ReDim arrData(1 To totalRows, 1 To numCols)
    ReDim arrResult(1 To totalRows, 1 To numCols)
    
    ' Daten einlesen
    Dim r As Long, c As Long
    For r = 1 To totalRows
        For c = 1 To numCols
            arrData(r, c) = ws.Cells(DATA_START_ROW + r - 1, startCol + c - 1).value
        Next c
    Next r
    
    ' Nicht-leere Zeilen sammeln (geprueft ueber checkCol)
    resultCount = 0
    Dim checkIdx As Long
    checkIdx = checkCol - startCol + 1
    
    For r = 1 To totalRows
        If Trim(CStr(arrData(r, checkIdx))) <> "" Then
            resultCount = resultCount + 1
            For c = 1 To numCols
                arrResult(resultCount, c) = arrData(r, c)
            Next c
        End If
    Next r
    
    ' Bereich loeschen und verdichtete Daten zurueckschreiben
    If resultCount > 0 Then
        ws.Range(ws.Cells(DATA_START_ROW, startCol), _
                 ws.Cells(lastRow, endCol)).ClearContents
        
        For r = 1 To resultCount
            For c = 1 To numCols
                ws.Cells(DATA_START_ROW + r - 1, startCol + c - 1).value = arrResult(r, c)
            Next c
        Next r
    Else
        ' Alles leer - Bereich leeren
        ws.Range(ws.Cells(DATA_START_ROW, startCol), _
                 ws.Cells(lastRow, endCol)).ClearContents
    End If
    
End Sub

' ===============================================================
' NEU: Entsperrt bestehende Daten + genau 1 naechste freie Zeile
' Sperrt einen Puffer darunter um ungewolltes Editieren zu verhindern
' Betrifft: B, D, F, H, J-P, R-X (via W), AB, AC, AD, AH
' NEU v4.0: EntityRole-Dropdown fuer ALLE bestehenden Zeilen,
'           lastRow fuer EntityKey-Bereich ueber IBAN ermittelt
' ===============================================================
Private Sub EntspeerreEditierbareSpalten(ByRef ws As Worksheet)
    
    Dim lastRow As Long
    Dim NextRow As Long
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
        NextRow = lastRow + 1
        lockEnd = NextRow + 50
        
        ' Bestehende Daten entsperren (editierbar)
        If lastRow >= DATA_START_ROW Then
            ws.Range(ws.Cells(DATA_START_ROW, singleCols(c)), _
                     ws.Cells(lastRow, singleCols(c))).Locked = False
        End If
        
        ' Genau 1 naechste freie Zeile entsperren
        ws.Cells(NextRow, singleCols(c)).Locked = False
        
        ' Bereich darunter sperren (50 Zeilen Sicherheitspuffer)
        ws.Range(ws.Cells(NextRow + 1, singleCols(c)), _
                 ws.Cells(lockEnd, singleCols(c))).Locked = True
    Next c
    
    ' === KATEGORIE-TABELLE: J-P (10-16) ===
    lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW - 1
    NextRow = lastRow + 1
    lockEnd = NextRow + 50
    
    ' Bestehende Daten J-P entsperren (editierbar)
    If lastRow >= DATA_START_ROW Then
        ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), _
                 ws.Cells(lastRow, DATA_CAT_COL_END)).Locked = False
    End If
    
    ' Genau 1 naechste freie Zeile J-P entsperren
    ws.Range(ws.Cells(NextRow, DATA_CAT_COL_START), _
             ws.Cells(NextRow, DATA_CAT_COL_END)).Locked = False
    
    ' Bereich darunter sperren
    ws.Range(ws.Cells(NextRow + 1, DATA_CAT_COL_START), _
             ws.Cells(lockEnd, DATA_CAT_COL_END)).Locked = True
    
    ' DropDowns fuer die Eingabezeile setzen
    Call SetzeZielspalteDropdown(ws, NextRow, "")
    
    ' Dropdown K (E/A)
    ws.Cells(NextRow, DATA_CAT_COL_EINAUS).Validation.Delete
    With ws.Cells(NextRow, DATA_CAT_COL_EINAUS).Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:="E,A"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    
    ' Dropdown M (Prioritaet)
    ws.Cells(NextRow, DATA_CAT_COL_PRIORITAET).Validation.Delete
    With ws.Cells(NextRow, DATA_CAT_COL_PRIORITAET).Validation
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
    ws.Cells(NextRow, DATA_CAT_COL_FAELLIGKEIT).Validation.Delete
    With ws.Cells(NextRow, DATA_CAT_COL_FAELLIGKEIT).Validation
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
    ' WICHTIG: lastRow ueber IBAN (Spalte S) ermitteln, nicht EntityKey (R)
    ' Denn nach IBAN-Import ist R oft noch leer
    lastRow = ws.Cells(ws.Rows.count, EK_COL_IBAN).End(xlUp).Row
    ' Auch R pruefen fuer den Fall dass R weiter reicht als S
    Dim lastRowR As Long
    lastRowR = ws.Cells(ws.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    If lastRowR > lastRow Then lastRow = lastRowR
    
    If lastRow < EK_START_ROW Then lastRow = EK_START_ROW - 1
    NextRow = lastRow + 1
    lockEnd = NextRow + 50
    
    ' Genau 1 naechste freie Zeile R-X entsperren
    ws.Range(ws.Cells(NextRow, EK_COL_ENTITYKEY), _
             ws.Cells(NextRow, EK_COL_DEBUG)).Locked = False
    
    ' Bereich darunter sperren
    ws.Range(ws.Cells(NextRow + 1, EK_COL_ENTITYKEY), _
             ws.Cells(lockEnd, EK_COL_DEBUG)).Locked = True
    
    ' EntityRole-Dropdown (W) Quelle: Spalte AD
    lastRowDD = ws.Cells(ws.Rows.count, DATA_COL_DD_ENTITYROLE).End(xlUp).Row
    If lastRowDD < DATA_START_ROW Then lastRowDD = DATA_START_ROW
    
    ' Dropdown W fuer die Eingabezeile
    ws.Cells(NextRow, EK_COL_ROLE).Validation.Delete
    With ws.Cells(NextRow, EK_COL_ROLE).Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:="=" & WS_DATEN & "!$AD$" & DATA_START_ROW & ":$AD$" & lastRowDD
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    
    ' NEU v5.1: EntityRole-Dropdown und Parzellen-Dropdown fuer ALLE Zeilen
    Dim lastRowParzelle As Long
    lastRowParzelle = ws.Cells(ws.Rows.count, DATA_COL_DD_PARZELLE).End(xlUp).Row
    If lastRowParzelle < DATA_START_ROW Then lastRowParzelle = DATA_START_ROW
    
    If lastRow >= EK_START_ROW Then
        For r = EK_START_ROW To lastRow
            Dim currentRole As String
            currentRole = UCase(Trim(ws.Cells(r, EK_COL_ROLE).value))
            
            ' EntityRole-Dropdown fuer ALLE Zeilen (W immer editierbar)
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
            
            ' U und X immer editierbar
            ws.Cells(r, EK_COL_ZUORDNUNG).Locked = False
            ws.Cells(r, EK_COL_DEBUG).Locked = False
            
            ' Parzellen-Dropdown fuer EHEMALIGES MITGLIED und SONSTIGE
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
        NextRow = lastRow + 1
        lockEnd = NextRow + 50
        
        ' Bestehende Daten entsperren (editierbar)
        If lastRow >= DATA_START_ROW Then
            ws.Range(ws.Cells(DATA_START_ROW, helperCols(c)), _
                     ws.Cells(lastRow, helperCols(c))).Locked = False
        End If
        
        ' Genau 1 naechste freie Zeile entsperren
        ws.Cells(NextRow, helperCols(c)).Locked = False
        
        ' Bereich darunter sperren
        ws.Range(ws.Cells(NextRow + 1, helperCols(c)), _
                 ws.Cells(lockEnd, helperCols(c))).Locked = True
    Next c
    
    On Error GoTo 0
    
End Sub

' ===============================================================
' Formatiert die Kategorie-Tabelle (Spalten J-P)
' FIX v5.3: Spalte K Breite 12, Zebra+Rahmen nur fuer belegte Zeilen,
'           leere Zeilen unterhalb werden bereinigt
' ===============================================================
Public Sub FormatiereKategorieTabelle(Optional ByRef ws As Worksheet = Nothing)
    
    Dim lastRow As Long
    Dim lastRowMax As Long
    Dim rngTable As Range
    Dim rngLeeren As Range
    Dim r As Long
    Dim einAusWert As String
    
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    
    ' Maximale letzte Zeile ueber alle Spalten J-P ermitteln (fuer Bereinigung)
    lastRowMax = lastRow
    Dim col As Long
    For col = DATA_CAT_COL_START To DATA_CAT_COL_END
        Dim colLastRow As Long
        colLastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
        If colLastRow > lastRowMax Then lastRowMax = colLastRow
    Next col
    
    ' Bereich UNTERHALB der belegten Zeilen bereinigen (Rahmen + Farbe entfernen)
    If lastRowMax >= DATA_START_ROW Then
        Dim cleanStart As Long
        If lastRow < DATA_START_ROW Then
            cleanStart = DATA_START_ROW
        Else
            cleanStart = lastRow + 1
        End If
        
        If cleanStart <= lastRowMax + 50 Then
            Set rngLeeren = ws.Range(ws.Cells(cleanStart, DATA_CAT_COL_START), _
                                     ws.Cells(lastRowMax + 50, DATA_CAT_COL_END))
            rngLeeren.Interior.ColorIndex = xlNone
            rngLeeren.Borders.LineStyle = xlNone
        End If
    End If
    
    If lastRow < DATA_START_ROW Then Exit Sub
    
    Set rngTable = ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), _
                            ws.Cells(lastRow, DATA_CAT_COL_END))
    
    ' Zuerst alles zuruecksetzen im belegten Bereich
    rngTable.Interior.ColorIndex = xlNone
    rngTable.Borders.LineStyle = xlNone
    
    ' Zebra-Formatierung NUR fuer belegte Zeilen
    For r = DATA_START_ROW To lastRow
        If (r - DATA_START_ROW) Mod 2 = 0 Then
            ws.Range(ws.Cells(r, DATA_CAT_COL_START), ws.Cells(r, DATA_CAT_COL_END)).Interior.color = ZEBRA_COLOR_1
        Else
            ws.Range(ws.Cells(r, DATA_CAT_COL_START), ws.Cells(r, DATA_CAT_COL_END)).Interior.color = ZEBRA_COLOR_2
        End If
    Next r
    
    ' Rahmenlinien NUR fuer belegte Zeilen
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    rngTable.VerticalAlignment = xlCenter
    
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
    
    For r = DATA_START_ROW To lastRow
        einAusWert = UCase(Trim(ws.Cells(r, DATA_CAT_COL_EINAUS).value))
        Call SetzeZielspalteDropdown(ws, r, einAusWert)
    Next r
    
    ' Spaltenbreiten
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KATEGORIE), _
             ws.Cells(lastRow, DATA_CAT_COL_KATEGORIE)).EntireColumn.AutoFit
    
    ' FIX v5.3: Spalte K (E/A) feste Breite 12
    ws.Columns(DATA_CAT_COL_EINAUS).ColumnWidth = 12
    
    ' Restliche Spalten AutoFit
    Dim autoFitCol As Long
    For autoFitCol = DATA_CAT_COL_KEYWORD To DATA_CAT_COL_END
        ws.Columns(autoFitCol).AutoFit
    Next autoFitCol
    
End Sub

' ===============================================================
' Sortiert die Kategorie-Tabelle nach Spalte J (A-Z)
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
' Sortiert die EntityKey-Tabelle
' FIX v5.0: Ampelfarben werden NACH Sortierung neu gesetzt
'           durch mod_EntityKey_Manager.SetzeAlleAmpelfarbenNachSortierung
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
    Dim lastRowIBAN As Long
    lastRowIBAN = ws.Cells(ws.Rows.count, EK_COL_IBAN).End(xlUp).Row
    If lastRowIBAN > lastRow Then lastRow = lastRowIBAN
    
    If lastRow < EK_START_ROW Then Exit Sub
    
    Dim arrData() As Variant
    Dim arrSorted() As Variant
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
    
    ' ===== FIX v5.0: Ampelfarben NACH Sortierung neu berechnen =====
    Call mod_EntityKey_Manager.SetzeAlleAmpelfarbenNachSortierung(ws)
    
End Sub

' ===============================================================
' Vergleicht zwei EntityKey-Zeilen fuer Sortierung
' FIX v5.3.1: Bei Mehrfach-Parzellen (z.B. "13, 14") nur erste
'             Parzelle fuer Sortierung verwenden
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
    
    ' FIX v5.3.1: Bei Mehrfach-Parzellen nur erste Parzelle nutzen
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

' ===============================================================
' DROPDOWN-LISTEN FUER KATEGORIEN AKTUALISIEREN (AF + AG + AH)
' FIX v5.2: "M" & ChrW(228) & "rz" statt "Maerz"
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
    For Each key In dictEinnahmen.keys
        ws.Cells(nextRowE, DATA_COL_KAT_EINNAHMEN).value = key
        nextRowE = nextRowE + 1
    Next key
    
    nextRowA = 4
    For Each key In dictAusgaben.keys
        ws.Cells(nextRowA, DATA_COL_KAT_AUSGABEN).value = key
        nextRowA = nextRowA + 1
    Next key
    
    ws.Cells(4, DATA_COL_MONAT_PERIODE).value = "Januar"
    ws.Cells(5, DATA_COL_MONAT_PERIODE).value = "Februar"
    ws.Cells(6, DATA_COL_MONAT_PERIODE).value = "M" & ChrW(228) & "rz"
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
' FIX v5.0: lastRow ueber IBAN UND EntityKey ermitteln
' ===============================================================
Private Sub FormatiereEntityKeyTabelleKomplett(ByRef ws As Worksheet)
    
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
' Formatiert die EntityKey-Tabelle
' R-T mit Zebra, U-X NUR Rahmen (Ampelfarben bleiben erhalten!)
' FIX v5.1: Spalte T WrapText nur bei vbLf-Inhalt
'           Spalten U, W, X IMMER editierbar
'           Spalte V editierbar bei EHEMALIGES MITGLIED / SONSTIGE
' FIX v5.2: Spalte U WrapText per Zelle (nur bei 2+ Namen mit vbLf)
' FIX v5.3: Spalte X Breite 65, WrapText erlaubt
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
    
    ' Spalte T: WrapText NUR wenn vbLf im Wert (mehrere Namen)
    ' Sonst kein Zeilenumbruch
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
    
    ' FIX v5.2: Spalte U (Zuordnung) - WrapText per Zelle (nur bei vbLf-Inhalt)
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
    
    ' FIX v5.3: Spalte X (Debug) - WrapText erlaubt, Breite 65
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_DEBUG), _
                  ws.Cells(lastRow, EK_COL_DEBUG))
        .WrapText = True
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_DEBUG).ColumnWidth = 65
    
    ' R-T immer gesperrt (EntityKey, IBAN, Kontoname)
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
             ws.Cells(lastRow, EK_COL_KONTONAME)).Locked = True
    
    For r = EK_START_ROW To lastRow
        currentRole = Trim(ws.Cells(r, EK_COL_ROLE).value)
        
        Call SetzeZellschutzFuerZeile(ws, r, currentRole)
        
        ' Zebra NUR fuer R-T (Spalten 18-20)
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
' Setzt Zellschutz basierend auf EntityRole
' FIX v5.1: U (Zuordnung), W (Role), X (Debug) IMMER editierbar
'           V (Parzelle) editierbar bei EHEMALIGES MITGLIED / SONSTIGE
' ===============================================================
Private Sub SetzeZellschutzFuerZeile(ByRef ws As Worksheet, ByVal zeile As Long, ByVal currentRole As String)
    
    On Error Resume Next
    
    ' R-T immer gesperrt (EntityKey, IBAN, Kontoname)
    ws.Range(ws.Cells(zeile, EK_COL_ENTITYKEY), ws.Cells(zeile, EK_COL_KONTONAME)).Locked = True
    
    ' U (Zuordnung) = IMMER editierbar
    ws.Cells(zeile, EK_COL_ZUORDNUNG).Locked = False
    
    ' W (Role) = IMMER editierbar (Dropdown)
    ws.Cells(zeile, EK_COL_ROLE).Locked = False
    
    ' X (Debug) = IMMER editierbar
    ws.Cells(zeile, EK_COL_DEBUG).Locked = False
    
    ' V (Parzelle) = nur bei EHEMALIGES MITGLIED oder SONSTIGE editierbar
    Dim roleUpper As String
    roleUpper = UCase(Trim(currentRole))
    
    If roleUpper = "EHEMALIGES MITGLIED" Or roleUpper = "SONSTIGE" Or roleUpper = "" Or roleUpper = "UNBEKANNT" Then
        ws.Cells(zeile, EK_COL_PARZELLE).Locked = False
    Else
        ws.Cells(zeile, EK_COL_PARZELLE).Locked = True
    End If
    
    On Error GoTo 0
    
End Sub

' ===============================================================
' BANKKONTO-BLATT FORMATIEREN
' FIX v5.2: Umlaute in MsgBox
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
           "- Zeilenh" & ChrW(246) & "he angepasst" & vbCrLf & _
           "- W" & ChrW(228) & "hrung mit Euro-Zeichen", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler bei der Formatierung: " & Err.Description, vbCritical
End Sub

' ===============================================================
' Prueft ob Named Range existiert
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
' NEU v5.3: Validierung und Reaktion bei Aenderung in Kategorie-Tabelle
' Prueft: Spalte K (E/A-Konsistenz bei gleicher Kategorie)
'         Spalte L (Keyword-Duplikat bei gleicher Kategorie)
' ===============================================================
Public Sub OnKategorieChange(ByVal Target As Range)
    
    Dim ws As Worksheet
    Dim zeile As Long
    Dim kategorie As String
    Dim r As Long
    Dim lastRow As Long
    
    Set ws = Target.Worksheet
    zeile = Target.Row
    
    If zeile < DATA_START_ROW Then Exit Sub
    
    ' --- Spalte J oder K geaendert: Dropdown-Listen aktualisieren ---
    If Target.Column = DATA_CAT_COL_KATEGORIE Or Target.Column = DATA_CAT_COL_EINAUS Then
        Call AktualisiereKategorieDropdownListen(ws)
    End If
    
    ' --- Spalte K (E/A) geaendert: Konsistenzpruefung ---
    If Target.Column = DATA_CAT_COL_EINAUS Then
        Dim einAus As String
        einAus = UCase(Trim(Target.value))
        
        kategorie = Trim(ws.Cells(zeile, DATA_CAT_COL_KATEGORIE).value)
        
        If kategorie <> "" And einAus <> "" Then
            lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
            
            Dim bestehenderTyp As String
            bestehenderTyp = ""
            
            ' Suche ob diese Kategorie bereits mit einem Typ existiert
            For r = DATA_START_ROW To lastRow
                If r <> zeile Then
                    If StrComp(Trim(ws.Cells(r, DATA_CAT_COL_KATEGORIE).value), kategorie, vbTextCompare) = 0 Then
                        Dim vorhandenerEA As String
                        vorhandenerEA = UCase(Trim(ws.Cells(r, DATA_CAT_COL_EINAUS).value))
                        If vorhandenerEA <> "" Then
                            bestehenderTyp = vorhandenerEA
                            Exit For
                        End If
                    End If
                End If
            Next r
            
            ' Wenn gleiche Kategorie mit anderem Typ existiert -> korrigieren
            If bestehenderTyp <> "" And bestehenderTyp <> einAus Then
                Application.EnableEvents = False
                Target.value = bestehenderTyp
                Application.EnableEvents = True
                
                Dim typBeschreibung As String
                If bestehenderTyp = "E" Then
                    typBeschreibung = "Einnahme (E)"
                Else
                    typBeschreibung = "Ausgabe (A)"
                End If
                
                MsgBox "Die Kategorie """ & kategorie & """ ist bereits als " & typBeschreibung & " eingetragen." & vbCrLf & vbCrLf & _
                       "Bei gleicher Kategorie kann nur einheitlich zwischen Einnahme oder Ausgabe gew" & ChrW(228) & "hlt werden." & vbCrLf & _
                       "Gemischte Angaben sind nicht gestattet." & vbCrLf & vbCrLf & _
                       "Der Wert wurde automatisch auf """ & bestehenderTyp & """ korrigiert.", _
                       vbInformation, "Kategorie-Konsistenz"
            End If
        End If
        
        ' Zielspalte-Dropdown aktualisieren
        einAus = UCase(Trim(Target.value))
        Call SetzeZielspalteDropdown(ws, zeile, einAus)
    End If
    
    ' --- Spalte L (Keyword) geaendert: Duplikat-Pruefung ---
    If Target.Column = DATA_CAT_COL_KEYWORD Then
        Dim neuesKeyword As String
        neuesKeyword = Trim(Target.value)
        
        If neuesKeyword <> "" Then
            kategorie = Trim(ws.Cells(zeile, DATA_CAT_COL_KATEGORIE).value)
            
            If kategorie <> "" Then
                lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
                
                For r = DATA_START_ROW To lastRow
                    If r <> zeile Then
                        If StrComp(Trim(ws.Cells(r, DATA_CAT_COL_KATEGORIE).value), kategorie, vbTextCompare) = 0 Then
                            If StrComp(Trim(ws.Cells(r, DATA_CAT_COL_KEYWORD).value), neuesKeyword, vbTextCompare) = 0 Then
                                MsgBox "F" & ChrW(252) & "r die Kategorie """ & kategorie & """ gibt es bereits das gew" & ChrW(228) & "hlte Schl" & ChrW(252) & "sselwort """ & neuesKeyword & """.", _
                                       vbExclamation, "Doppeltes Schl" & ChrW(252) & "sselwort"
                                Exit For
                            End If
                        End If
                    End If
                Next r
            End If
        End If
    End If
    
End Sub

' ===============================================================
' WORKSHEET_CHANGE HANDLER fuer dynamische Formatierung
' FIX v5.3: Bei jeder Aenderung ALLE Spalten neu formatieren,
'           damit benachbarte Rahmenlinien nicht verloren gehen
'           BlendeDatenSpaltenAus am Ende
' ===============================================================
Public Sub OnDatenChange(ByVal Target As Range, ByVal ws As Worksheet)
    
    On Error GoTo ErrorHandler
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Alle Einzelspalten neu formatieren (B, D, F, H, Z-AH)
    Call FormatiereAlleDatenSpalten(ws)
    
    ' Kategorie-Tabelle (J-P)
    Call FormatiereKategorieTabelle(ws)
    Call SortiereKategorieTabelle(ws)
    
    ' EntityKey-Tabelle (R-X)
    Call FormatiereEntityKeyTabelleKomplett(ws)
    Call SortiereEntityKeyTabelle(ws)
    
    ' Spezifische Change-Logik nur wenn in Kategorie-Bereich
    If Not Intersect(Target, ws.Range("J:P")) Is Nothing Then
        Call OnKategorieChange(Target)
    End If
    
    ' NEU v5.3: Hilfsspalten ausblenden
    Call BlendeDatenSpaltenAus
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Debug.Print "Fehler in OnDatenChange: " & Err.Description
    End If
    
End Sub



