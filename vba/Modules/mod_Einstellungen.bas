Attribute VB_Name = "mod_Einstellungen"
Option Explicit

' ===============================================================
' MODUL: mod_Einstellungen
' VERSION: 1.0 - 07.02.2026
' ZWECK: Formatierung, DropDowns, Schutz/Entsperrung fuer
'        die Zahlungstermin-Tabelle auf Blatt Einstellungen
'        (Spalten B-H, ab Zeile 4, Header Zeile 3)
' ===============================================================

Private Const ZEBRA_COLOR_1 As Long = &HFFFFFF  ' Weiss
Private Const ZEBRA_COLOR_2 As Long = &HDEE5E3  ' Hellgrau (gleich wie Daten/Bankkonto)


' ===============================================================
' 1. HAUPTPROZEDUR: Komplette Formatierung der Tabelle
'    Aufruf: Worksheet_Activate, nach Loeschen, nach Einfuegen
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
    
    ' 1. Header formatieren
    Call FormatiereHeader(ws)
    
    ' 2. Leerzeilen entfernen (Daten verdichten)
    Call VerdichteDaten(ws)
    
    ' 3. Zebra-Formatierung
    Call AnwendeZebra(ws)
    
    ' 4. Rahmenlinien
    Call AnwendeRahmen(ws)
    
    ' 5. Spaltenformate und Ausrichtung
    Call AnwendeSpaltenformate(ws)
    
    ' 6. DropDown-Listen setzen
    Call SetzeDropDowns(ws)
    
    ' 7. Zellen sperren/entsperren
    Call SperreUndEntsperre(ws)
    
    ' 8. Spaltenbreiten
    Call SetzeSpaltenbreiten(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub


' ===============================================================
' 2. HEADER FORMATIEREN (Zeile 3)
' ===============================================================
Private Sub FormatiereHeader(ByVal ws As Worksheet)
    
    Dim rngHeader As Range
    Set rngHeader = ws.Range(ws.Cells(ES_HEADER_ROW, ES_COL_START), _
                             ws.Cells(ES_HEADER_ROW, ES_COL_END))
    
    ' Ueberschriften setzen (falls leer oder falsch)
    ws.Cells(ES_HEADER_ROW, ES_COL_KATEGORIE).value = "Referenz Kategorie" & vbLf & "(Leistungsart)"
    ws.Cells(ES_HEADER_ROW, ES_COL_SOLL_BETRAG).value = "Soll-Betrag"
    ws.Cells(ES_HEADER_ROW, ES_COL_SOLL_TAG).value = "Soll-Tag" & vbLf & "(des Monats)"
    ws.Cells(ES_HEADER_ROW, ES_COL_STICHTAG_FIX).value = "Soll-Stichtag" & vbLf & "(Fix) TT.MM."
    ws.Cells(ES_HEADER_ROW, ES_COL_VORLAUF).value = "Vorlauf-Toleranz" & vbLf & "(Tage)"
    ws.Cells(ES_HEADER_ROW, ES_COL_NACHLAUF).value = "Nachlauf-Toleranz" & vbLf & "(Tage)"
    ws.Cells(ES_HEADER_ROW, ES_COL_SAEUMNIS).value = "Saeumnis-" & vbLf & "Gebuehr"
    
    With rngHeader
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Interior.color = RGB(68, 114, 196)  ' Dunkelblau
        .Font.color = RGB(255, 255, 255)     ' Weiss
        .Locked = True
        
        ' Rahmen um Header
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlThin
    End With
    
    ws.Rows(ES_HEADER_ROW).RowHeight = 36
    
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
        ' Zeile ist nicht leer wenn Spalte B (Kategorie) gefuellt
        If Trim(ws.Cells(r, ES_COL_KATEGORIE).value) <> "" Then
            resultCount = resultCount + 1
            For c = 1 To numCols
                arrResult(resultCount, c) = ws.Cells(r, ES_COL_START + c - 1).value
            Next c
        End If
    Next r
    
    ' Bereich loeschen und verdichtete Daten zurueckschreiben
    ws.Range(ws.Cells(ES_START_ROW, ES_COL_START), _
             ws.Cells(lastRow, ES_COL_END)).ClearContents
    ws.Range(ws.Cells(ES_START_ROW, ES_COL_START), _
             ws.Cells(lastRow, ES_COL_END)).Interior.ColorIndex = xlNone
    ws.Range(ws.Cells(ES_START_ROW, ES_COL_START), _
             ws.Cells(lastRow, ES_COL_END)).Borders.LineStyle = xlNone
    
    If resultCount > 0 Then
        For r = 1 To resultCount
            For c = 1 To numCols
                ws.Cells(ES_START_ROW + r - 1, ES_COL_START + c - 1).value = arrResult(r, c)
            Next c
        Next r
    End If
    
End Sub


' ===============================================================
' 4. ZEBRA-FORMATIERUNG
' ===============================================================
Private Sub AnwendeZebra(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim r As Long
    Dim rngRow As Range
    
    lastRow = LetzteZeile(ws)
    If lastRow < ES_START_ROW Then Exit Sub
    
    For r = ES_START_ROW To lastRow
        Set rngRow = ws.Range(ws.Cells(r, ES_COL_START), ws.Cells(r, ES_COL_END))
        
        If (r - ES_START_ROW) Mod 2 = 0 Then
            rngRow.Interior.color = ZEBRA_COLOR_1
        Else
            rngRow.Interior.color = ZEBRA_COLOR_2
        End If
    Next r
    
    ' Zeilen unterhalb der Daten: Formatierung entfernen
    If lastRow + 1 <= lastRow + 50 Then
        ws.Range(ws.Cells(lastRow + 1, ES_COL_START), _
                 ws.Cells(lastRow + 50, ES_COL_END)).Interior.ColorIndex = xlNone
        ws.Range(ws.Cells(lastRow + 1, ES_COL_START), _
                 ws.Cells(lastRow + 50, ES_COL_END)).Borders.LineStyle = xlNone
    End If
    
End Sub


' ===============================================================
' 5. RAHMENLINIEN (duenne schwarze Linien innen und aussen)
' ===============================================================
Private Sub AnwendeRahmen(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim rngDaten As Range
    
    lastRow = LetzteZeile(ws)
    If lastRow < ES_START_ROW Then Exit Sub
    
    Set rngDaten = ws.Range(ws.Cells(ES_START_ROW, ES_COL_START), _
                            ws.Cells(lastRow, ES_COL_END))
    
    With rngDaten.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = vbBlack
    End With
    With rngDaten.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = vbBlack
    End With
    With rngDaten.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = vbBlack
    End With
    With rngDaten.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = vbBlack
    End With
    With rngDaten.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = vbBlack
    End With
    With rngDaten.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = vbBlack
    End With
    
End Sub


' ===============================================================
' 6. SPALTENFORMATE UND AUSRICHTUNG
' ===============================================================
Private Sub AnwendeSpaltenformate(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim endRow As Long
    
    lastRow = LetzteZeile(ws)
    endRow = lastRow + 50  ' Pufferbereich mit formatieren
    If endRow < ES_START_ROW + 50 Then endRow = ES_START_ROW + 50
    
    ' Spalte B: Linkbuendig, kein Textumbruch
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_KATEGORIE), _
                  ws.Cells(endRow, ES_COL_KATEGORIE))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
    
    ' Spalte C: Waehrung, rechtsbuendig
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_SOLL_BETRAG), _
                  ws.Cells(endRow, ES_COL_SOLL_BETRAG))
        .NumberFormat = "#,##0.00 " & ChrW(8364)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    
    ' Spalte D: Zentriert, ganze Zahlen
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_SOLL_TAG), _
                  ws.Cells(endRow, ES_COL_SOLL_TAG))
        .NumberFormat = "0"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Spalte E: Text TT.MM., zentriert
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_STICHTAG_FIX), _
                  ws.Cells(endRow, ES_COL_STICHTAG_FIX))
        .NumberFormat = "@"  ' Text-Format damit 01.03 nicht als Datum interpretiert wird
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Spalte F: Zentriert, ganze Zahlen
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_VORLAUF), _
                  ws.Cells(endRow, ES_COL_VORLAUF))
        .NumberFormat = "0"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Spalte G: Zentriert, ganze Zahlen
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_NACHLAUF), _
                  ws.Cells(endRow, ES_COL_NACHLAUF))
        .NumberFormat = "0"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Spalte H: Waehrung, rechtsbuendig
    With ws.Range(ws.Cells(ES_START_ROW, ES_COL_SAEUMNIS), _
                  ws.Cells(endRow, ES_COL_SAEUMNIS))
        .NumberFormat = "#,##0.00 " & ChrW(8364)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    
End Sub


' ===============================================================
' 7. DROPDOWN-LISTEN SETZEN
' ===============================================================
Private Sub SetzeDropDowns(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim nextRow As Long
    Dim r As Long
    Dim kategorienListe As String
    
    lastRow = LetzteZeile(ws)
    nextRow = lastRow + 1
    If nextRow < ES_START_ROW Then nextRow = ES_START_ROW
    
    ' --- Kategorie-Liste aus Daten!J erstellen (nicht redundant) ---
    kategorienListe = HoleKategorienAlsListe()
    
    ' --- Spalte B: Kategorie-DropDown fuer alle Datenzeilen + Eingabezeile ---
    For r = ES_START_ROW To nextRow
        ws.Cells(r, ES_COL_KATEGORIE).Validation.Delete
        If kategorienListe <> "" Then
            With ws.Cells(r, ES_COL_KATEGORIE).Validation
                .Add Type:=xlValidateList, _
                     AlertStyle:=xlValidAlertStop, _
                     Formula1:=kategorienListe
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = False
                .ShowError = True
            End With
        End If
    Next r
    
    ' --- Spalte D: Tag 1-31 ---
    Dim tagListe As String
    tagListe = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31"
    
    For r = ES_START_ROW To nextRow
        ws.Cells(r, ES_COL_SOLL_TAG).Validation.Delete
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
    
    ' --- Spalte F: Vorlauf 0-31 ---
    Dim toleranzListe As String
    toleranzListe = "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31"
    
    For r = ES_START_ROW To nextRow
        ws.Cells(r, ES_COL_VORLAUF).Validation.Delete
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
    
    ' --- Spalte G: Nachlauf 0-31 ---
    For r = ES_START_ROW To nextRow
        ws.Cells(r, ES_COL_NACHLAUF).Validation.Delete
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
' 8. SPERREN UND ENTSPERREN
'    - Bestehende Datenzeilen B-H: entsperrt (editierbar)
'    - Genau 1 naechste freie Zeile: entsperrt (Neuanlage)
'    - Alles darunter: gesperrt
'    - Alles ausserhalb B-H: gesperrt
' ===============================================================
Private Sub SperreUndEntsperre(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim nextRow As Long
    Dim lockEnd As Long
    
    lastRow = LetzteZeile(ws)
    nextRow = lastRow + 1
    If nextRow < ES_START_ROW Then nextRow = ES_START_ROW
    lockEnd = nextRow + 50
    
    ' Gesamtes Blatt sperren
    ws.Cells.Locked = True
    
    ' Bestehende Daten B-H entsperren (editierbar)
    If lastRow >= ES_START_ROW Then
        ws.Range(ws.Cells(ES_START_ROW, ES_COL_START), _
                 ws.Cells(lastRow, ES_COL_END)).Locked = False
    End If
    
    ' Genau 1 naechste freie Zeile B-H entsperren (Neuanlage)
    ws.Range(ws.Cells(nextRow, ES_COL_START), _
             ws.Cells(nextRow, ES_COL_END)).Locked = False
    
    ' Bereich darunter explizit sperren (Sicherheitspuffer)
    ws.Range(ws.Cells(nextRow + 1, ES_COL_START), _
             ws.Cells(lockEnd, ES_COL_END)).Locked = True
    
End Sub


' ===============================================================
' 9. SPALTENBREITEN
' ===============================================================
Private Sub SetzeSpaltenbreiten(ByVal ws As Worksheet)
    
    ws.Columns(ES_COL_KATEGORIE).ColumnWidth = 24    ' B: Kategorie
    ws.Columns(ES_COL_SOLL_BETRAG).ColumnWidth = 14  ' C: Soll-Betrag
    ws.Columns(ES_COL_SOLL_TAG).ColumnWidth = 12     ' D: Soll-Tag
    ws.Columns(ES_COL_STICHTAG_FIX).ColumnWidth = 14 ' E: Stichtag Fix
    ws.Columns(ES_COL_VORLAUF).ColumnWidth = 14      ' F: Vorlauf
    ws.Columns(ES_COL_NACHLAUF).ColumnWidth = 14     ' G: Nachlauf
    ws.Columns(ES_COL_SAEUMNIS).ColumnWidth = 14     ' H: Saeumnis
    
End Sub


' ===============================================================
' 10. HILFSFUNKTIONEN
' ===============================================================

' Letzte Zeile im Bereich B ermitteln
Private Function LetzteZeile(ByVal ws As Worksheet) As Long
    
    Dim lr As Long
    lr = ws.Cells(ws.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    If lr < ES_START_ROW Then lr = ES_START_ROW - 1
    LetzteZeile = lr
    
End Function


' Kategorien aus Daten!J als nicht-redundante kommaseparierte Liste
Private Function HoleKategorienAlsListe() As String
    
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim kat As String
    Dim dict As Object
    Dim result As String
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    If wsDaten Is Nothing Then
        HoleKategorienAlsListe = ""
        Exit Function
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    
    For r = DATA_START_ROW To lastRow
        kat = Trim(wsDaten.Cells(r, DATA_CAT_COL_KATEGORIE).value)
        If kat <> "" Then
            If Not dict.Exists(kat) Then
                dict.Add kat, True
            End If
        End If
    Next r
    
    If dict.count = 0 Then
        HoleKategorienAlsListe = ""
        Exit Function
    End If
    
    ' Sortiert zusammenbauen (Dictionary-Keys sind Einfuegereihenfolge)
    result = Join(dict.keys, ",")
    
    ' Excel DropDown Limit: max 255 Zeichen in Formula1
    ' Bei Ueberschreitung: Named Range verwenden
    If Len(result) > 255 Then
        ' Fallback: Direkt auf den Bereich in Daten!J verweisen
        ' (kann Duplikate enthalten, aber funktioniert immer)
        HoleKategorienAlsListe = "='" & WS_DATEN & "'!$J$" & DATA_START_ROW & _
                                  ":$J$" & lastRow
    Else
        HoleKategorienAlsListe = result
    End If
    
End Function


' ===============================================================
' 11. ZEILE LOESCHEN (Aufruf aus Worksheet_Change oder Button)
' ===============================================================
Public Sub LoescheZahlungsterminZeile(ByVal ws As Worksheet, ByVal zeile As Long)
    
    If zeile < ES_START_ROW Then Exit Sub
    
    Dim lastRow As Long
    lastRow = LetzteZeile(ws)
    If zeile > lastRow Then Exit Sub
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' Zeile leeren (nicht loeschen, damit Blattstruktur erhalten bleibt)
    ws.Range(ws.Cells(zeile, ES_COL_START), _
             ws.Cells(zeile, ES_COL_END)).ClearContents
    
    ' Tabelle neu formatieren (verdichtet automatisch Leerzeilen)
    Call FormatiereZahlungsterminTabelle(ws)
    
End Sub


