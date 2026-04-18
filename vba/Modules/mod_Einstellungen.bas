Attribute VB_Name = "mod_Einstellungen"
Option Explicit

' ===============================================================
' MODUL: mod_Einstellungen (Orchestrator)
' VERSION: 4.0 - 18.04.2026
' ZWECK: Formatierung, Schutz/Entsperrung fuer
'        die Zahlungstermin-Tabelle auf Blatt Einstellungen
'        (Spalten B-I, ab Zeile 21, Header Zeile 20)
'        NEU v4.0: Konfigurationsbereich Zeilen 1-19
'          - Abrechnungsjahr, Kontostand Vorjahr
'          - Mitgliedsbeitrag, Miete, Grundsteuer, Pacht
'          - Vereinsadresse (Name, Strasse, PLZ, Ort)
'          - Migration des alten Layouts (Tabelle von Zeile 3 nach 20)
' AUSGELAGERT:
'   - mod_Einstellungen_DropDowns: SetzeDropDowns, HoleAlleKategorien
'   - mod_Einstellungen_Debug: DebugDropDownLogik, DebugValidation,
'                              DebugSetzeDropDownsUndPruefe
' ===============================================================

' Farben
Private Const ZEBRA_COLOR_1 As Long = &HFFFFFF  ' Weiss
Private Const CLR_CFG_HEADER As Long = 2894892   ' RGB(44, 62, 80) - Dunkel Blau-Grau
Private Const CLR_CFG_SECTION As Long = 6182740  ' RGB(52, 73, 94) - Mittel Blau-Grau
Private Const CLR_CFG_LABEL As Long = 15853804   ' RGB(236, 240, 241) - Helles Grau
Private Const CLR_CFG_CALC As Long = 14408667    ' RGB(219, 234, 219) - Helles Gruen (berechnet)
Private Const ZEBRA_COLOR_2 As Long = &HDEE5E3  ' Hellgrau


' ===============================================================
' 0. MIGRATION: Altes Layout (Header Zeile 3) nach neues (Zeile 20)
'    Wird einmalig bei Workbook_Open aufgerufen
' ===============================================================
Public Sub MigriereEinstellungenLayout()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    ' Pruefen ob Migration schon durchgefuehrt
    If Trim(CStr(ws.Cells(ES_CFG_TITEL_ROW, ES_CFG_LABEL_COL).value)) = "Konfiguration" Then
        Exit Sub
    End If
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Pruefen ob altes Layout vorhanden (Header in Zeile 3)
    Dim altesLayout As Boolean
    altesLayout = (InStr(1, CStr(ws.Cells(3, ES_COL_KATEGORIE).value), "Kategorie", vbTextCompare) > 0)
    
    If altesLayout Then
        ' 17 Zeilen oben einfuegen - verschiebt alle Daten automatisch
        ws.Rows("1:17").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
        ' Eingefuegte Zeilen bereinigen (keine Formatierung uebernehmen)
        ws.Range("A1:Z17").Clear
    End If
    
    ' Konfigurationsbereich schreiben
    Call SchreibeKonfigurationsBereich(ws)
    
    ' Parzellen-Anzahl automatisch ermitteln
    Call AktualisiereParzellen(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


' ===============================================================
' 0b. KONFIGURATIONSBEREICH: Labels, Formeln, Formatierung
' ===============================================================
Public Sub SchreibeKonfigurationsBereich(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
    End If
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    Dim euroFmt As String
    euroFmt = "#,##0.00 " & ChrW(8364)
    
    ' --- TITEL ---
    With ws.Range(ws.Cells(ES_CFG_TITEL_ROW, ES_CFG_LABEL_COL), _
                  ws.Cells(ES_CFG_TITEL_ROW, ES_COL_END))
        .Merge
        .value = "Konfiguration"
        .Font.Size = 16
        .Font.Bold = True
        .Font.color = RGB(255, 255, 255)
        .Interior.color = CLR_CFG_HEADER
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' --- SECTION: Kassenbuch ---
    Call SchreibeSectionHeader(ws, ES_CFG_KASSENBUCH_ROW, "Kassenbuch")
    
    ' Abrechnungsjahr
    Call SchreibeCfgLabel(ws, ES_CFG_ABRECHNUNGSJAHR_ROW, "Abrechnungsjahr:")
    With ws.Cells(ES_CFG_ABRECHNUNGSJAHR_ROW, ES_CFG_VALUE_COL)
        .NumberFormat = "0"
        .HorizontalAlignment = xlCenter
        .Locked = False
    End With
    
    ' Kontostand Vorjahr
    Call SchreibeCfgLabel(ws, ES_CFG_KONTOSTAND_ROW, "Kontostand Vorjahr (31.12.):")
    With ws.Cells(ES_CFG_KONTOSTAND_ROW, ES_CFG_VALUE_COL)
        .NumberFormat = euroFmt
        .HorizontalAlignment = xlRight
        .Locked = False
    End With
    
    ' --- SECTION: Beitraege & Pacht ---
    Call SchreibeSectionHeader(ws, ES_CFG_BEITRAEGE_ROW, "Beitr" & ChrW(228) & "ge & Pacht")
    
    ' Mitgliedsbeitrag
    Call SchreibeCfgLabel(ws, ES_CFG_MITGLIEDSBEITRAG_ROW, "Mitgliedsbeitrag (j" & ChrW(228) & "hrlich):")
    With ws.Cells(ES_CFG_MITGLIEDSBEITRAG_ROW, ES_CFG_VALUE_COL)
        .NumberFormat = euroFmt
        .HorizontalAlignment = xlRight
        .Locked = False
    End With
    
    ' Miete
    Call SchreibeCfgLabel(ws, ES_CFG_MIETE_ROW, "Miete (Pacht vom Verein):")
    With ws.Cells(ES_CFG_MIETE_ROW, ES_CFG_VALUE_COL)
        .NumberFormat = euroFmt
        .HorizontalAlignment = xlRight
        .Locked = False
    End With
    
    ' Grundsteuer
    Call SchreibeCfgLabel(ws, ES_CFG_GRUNDSTEUER_ROW, "Grundsteuer:")
    With ws.Cells(ES_CFG_GRUNDSTEUER_ROW, ES_CFG_VALUE_COL)
        .NumberFormat = euroFmt
        .HorizontalAlignment = xlRight
        .Locked = False
    End With
    
    ' Summe (berechnet)
    Call SchreibeCfgLabel(ws, ES_CFG_SUMME_ROW, "Summe (Miete + Grundsteuer):")
    With ws.Cells(ES_CFG_SUMME_ROW, ES_CFG_VALUE_COL)
        .FormulaLocal = "=C" & ES_CFG_MIETE_ROW & "+C" & ES_CFG_GRUNDSTEUER_ROW
        .NumberFormat = euroFmt
        .HorizontalAlignment = xlRight
        .Interior.color = CLR_CFG_CALC
        .Font.Bold = True
        .Locked = True
    End With
    
    ' Verpachtete Parzellen (Dropdown 1-14)
    Call SchreibeCfgLabel(ws, ES_CFG_PARZELLEN_ROW, "Verpachtete Parzellen:")
    With ws.Cells(ES_CFG_PARZELLEN_ROW, ES_CFG_VALUE_COL)
        .NumberFormat = "0"
        .HorizontalAlignment = xlCenter
        .Locked = False
        ' Dropdown 1-14
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="1,2,3,4,5,6,7,8,9,10,11,12,13,14"
            .IgnoreBlank = False
            .InCellDropdown = True
            .ErrorTitle = "Ung" & ChrW(252) & "ltige Eingabe"
            .ErrorMessage = "Bitte eine Zahl zwischen 1 und 14 w" & ChrW(228) & "hlen."
        End With
    End With
    
    ' Pacht pro Parzelle (berechnet)
    Call SchreibeCfgLabel(ws, ES_CFG_PACHT_ROW, "Pacht pro Parzelle:")
    With ws.Cells(ES_CFG_PACHT_ROW, ES_CFG_VALUE_COL)
        .FormulaLocal = "=WENN(C" & ES_CFG_PARZELLEN_ROW & ">0;C" & _
                        ES_CFG_SUMME_ROW & "/C" & ES_CFG_PARZELLEN_ROW & ";0)"
        .NumberFormat = euroFmt
        .HorizontalAlignment = xlRight
        .Interior.color = CLR_CFG_CALC
        .Font.Bold = True
        .Locked = True
    End With
    
    ' --- SECTION: Vereinsadresse ---
    Call SchreibeSectionHeader(ws, ES_CFG_ADRESSE_ROW, "Vereinsadresse")
    
    ' Vereinsname (Wert ueber C-I gemergt)
    Call SchreibeCfgLabel(ws, ES_CFG_VEREINSNAME_ROW, "Vereinsname:")
    With ws.Range(ws.Cells(ES_CFG_VEREINSNAME_ROW, ES_CFG_VALUE_COL), _
                  ws.Cells(ES_CFG_VEREINSNAME_ROW, ES_COL_END))
        .Merge
        .HorizontalAlignment = xlLeft
        .Locked = False
    End With
    
    ' Strasse
    Call SchreibeCfgLabel(ws, ES_CFG_STRASSE_ROW, "Stra" & ChrW(223) & "e:")
    With ws.Range(ws.Cells(ES_CFG_STRASSE_ROW, ES_CFG_VALUE_COL), _
                  ws.Cells(ES_CFG_STRASSE_ROW, ES_COL_END))
        .Merge
        .HorizontalAlignment = xlLeft
        .Locked = False
    End With
    
    ' PLZ + Ort
    Call SchreibeCfgLabel(ws, ES_CFG_PLZ_ORT_ROW, "PLZ:")
    With ws.Cells(ES_CFG_PLZ_ORT_ROW, ES_CFG_VALUE_COL)
        .NumberFormat = "@"
        .HorizontalAlignment = xlLeft
        .Locked = False
    End With
    ws.Cells(ES_CFG_PLZ_ORT_ROW, 4).value = "Ort:"
    With ws.Cells(ES_CFG_PLZ_ORT_ROW, 4)
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .Interior.color = CLR_CFG_LABEL
    End With
    With ws.Range(ws.Cells(ES_CFG_PLZ_ORT_ROW, 5), _
                  ws.Cells(ES_CFG_PLZ_ORT_ROW, ES_COL_END))
        .Merge
        .HorizontalAlignment = xlLeft
        .Locked = False
    End With
    
    ' Separator-Zeile
    With ws.Range(ws.Cells(ES_CFG_SEPARATOR_ROW, ES_CFG_LABEL_COL), _
                  ws.Cells(ES_CFG_SEPARATOR_ROW, ES_COL_END))
        .Interior.color = CLR_CFG_HEADER
        .RowHeight = 4
    End With
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
End Sub


' --- Hilfs-Sub: Section-Header schreiben ---
Private Sub SchreibeSectionHeader(ByVal ws As Worksheet, ByVal zeile As Long, ByVal titel As String)
    With ws.Range(ws.Cells(zeile, ES_CFG_LABEL_COL), _
                  ws.Cells(zeile, ES_COL_END))
        .Merge
        .value = titel
        .Font.Size = 11
        .Font.Bold = True
        .Font.color = RGB(255, 255, 255)
        .Interior.color = CLR_CFG_SECTION
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
End Sub


' --- Hilfs-Sub: Label in Spalte B schreiben ---
Private Sub SchreibeCfgLabel(ByVal ws As Worksheet, ByVal zeile As Long, ByVal text As String)
    With ws.Cells(zeile, ES_CFG_LABEL_COL)
        .value = text
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Interior.color = CLR_CFG_LABEL
        .Locked = True
    End With
End Sub


' ===============================================================
' 0c. PARZELLEN-ANZAHL: Automatisch aus Mitgliederliste ermitteln
' ===============================================================
Public Sub AktualisiereParzellen(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
    End If
    
    Dim aktParzellen As Long
    aktParzellen = mod_Startseite.ZaehleBelegteParzellen()
    
    If aktParzellen > 0 And aktParzellen <= 14 Then
        On Error Resume Next
        ws.Unprotect PASSWORD:=PASSWORD
        On Error GoTo 0
        
        ws.Cells(ES_CFG_PARZELLEN_ROW, ES_CFG_VALUE_COL).value = aktParzellen
        
        On Error Resume Next
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        On Error GoTo 0
    End If
End Sub


' ===============================================================
' 0d. KONTOSTAND VORJAHR: Pruefen und ggf. Nutzer fragen
' ===============================================================
Public Sub PruefeKontostandVorjahr()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    Dim wert As Variant
    wert = ws.Cells(ES_CFG_KONTOSTAND_ROW, ES_CFG_VALUE_COL).value
    
    ' Wenn gueltig -> nichts tun
    If IsNumeric(wert) And wert <> "" Then
        If CDbl(wert) <> 0 Then Exit Sub
    End If
    
    ' Nutzer nach Kontostand fragen
    Dim eingabe As String
    eingabe = InputBox("Der Kontostand vom 31.12. des Vorjahres fehlt " & _
                       "oder ist ung" & ChrW(252) & "ltig." & vbLf & vbLf & _
                       "Bitte den letzten bekannten Kontostand eingeben:", _
                       "Kontostand Vorjahr", "0,00")
    
    If eingabe = "" Then Exit Sub
    
    ' Komma durch Punkt ersetzen fuer CDbl
    eingabe = Replace(eingabe, ".", "")
    eingabe = Replace(eingabe, ",", ".")
    eingabe = Replace(eingabe, ChrW(8364), "")
    eingabe = Trim(eingabe)
    
    If IsNumeric(eingabe) Then
        On Error Resume Next
        ws.Unprotect PASSWORD:=PASSWORD
        On Error GoTo 0
        
        ws.Cells(ES_CFG_KONTOSTAND_ROW, ES_CFG_VALUE_COL).value = CDbl(eingabe)
        
        On Error Resume Next
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        On Error GoTo 0
    Else
        MsgBox "Die Eingabe """ & eingabe & """ ist keine g" & ChrW(252) & "ltige Zahl.", _
               vbExclamation, "Kontostand"
    End If
End Sub


' ===============================================================
' 0e. SYNC: Mitgliedsbeitrag -> Zahlungstermine-Tabelle
' ===============================================================
Public Sub SyncMitgliedsbeitragZuTabelle(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
    End If
    
    Dim mbWert As Variant
    mbWert = ws.Cells(ES_CFG_MITGLIEDSBEITRAG_ROW, ES_CFG_VALUE_COL).value
    If Not IsNumeric(mbWert) Or mbWert = "" Then Exit Sub
    
    ' In Zahlungstermine-Tabelle suchen: Spalte B = "Mitgliedsbeitrag"
    Dim r As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    
    For r = ES_START_ROW To lastRow
        If StrComp(Trim(CStr(ws.Cells(r, ES_COL_KATEGORIE).value)), _
                   "Mitgliedsbeitrag", vbTextCompare) = 0 Then
            ws.Cells(r, ES_COL_SOLL_BETRAG).value = CDbl(mbWert)
            Exit For
        End If
    Next r
End Sub


' ===============================================================
' 0f. SYNC: Pacht pro Parzelle -> Zahlungstermine-Tabelle
' ===============================================================
Public Sub SyncPachtZuTabelle(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
    End If
    
    Dim pachtWert As Variant
    pachtWert = ws.Cells(ES_CFG_PACHT_ROW, ES_CFG_VALUE_COL).value
    If Not IsNumeric(pachtWert) Or pachtWert = "" Then Exit Sub
    If CDbl(pachtWert) <= 0 Then Exit Sub
    
    ' In Zahlungstermine-Tabelle suchen: Spalte B = "Pacht Mitgliederzahlung"
    Dim r As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    
    For r = ES_START_ROW To lastRow
        If StrComp(Trim(CStr(ws.Cells(r, ES_COL_KATEGORIE).value)), _
                   "Pacht Mitgliederzahlung", vbTextCompare) = 0 Then
            ws.Cells(r, ES_COL_SOLL_BETRAG).value = CDbl(pachtWert)
            Exit For
        End If
    Next r
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
    
    ' Zustand sichern und wiederherstellen (nicht bedingungslos True setzen!)
    Dim eventsWaren As Boolean
    eventsWaren = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 1. Header einmalig pr?fen (nur setzen wenn leer)
    Call PruefeHeader(ws)
    
    ' 2. Leerzeilen entfernen (Daten verdichten)
    Call VerdichteDaten(ws)
    
    ' 3. Formatierung anwenden (Zebra + Rahmen)
    Call FormatiereTabelle(ws)
    
    ' 4. Spaltenformate und Ausrichtung
    Call AnwendeSpaltenformate(ws)
    
    ' 5. DropDown-Listen setzen (ausgelagert)
    Call mod_Einstellungen_DropDowns.SetzeDropDowns(ws)
    
    ' 6. Zellen sperren/entsperren
    Call SperreUndEntsperre(ws)
    
    ' 7. Spaltenbreiten (AutoFit f?r alle Spalten)
    Call SetzeSpaltenbreiten(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = eventsWaren
    Application.ScreenUpdating = True
    
End Sub


' ===============================================================
' 2. HEADER PR?FEN (Zeile 3)
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
    ws.Cells(ES_HEADER_ROW, ES_COL_SAEUMNIS).value = "S?umnis-Geb?hr"
    
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
'     ?ffentlich aufrufbar aus Tabelle9.cls
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


' ===============================================================
' 10. ZEILE L?SCHEN
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





































































