Attribute VB_Name = "mod_Formatierung"
Option Explicit

' ***************************************************************
' MODUL: mod_Formatierung
' ZWECK: Formatierung und DropDown-Listen-Verwaltung
' VERSION: 1.2 - 01.02.2026
' KORREKTUR: Euro-Zeichen Encoding + Formatiere_Alle_Tabellen_Neu
' ***************************************************************

Private Const ZEBRA_COLOR As Long = &HDEE5E3

' ===============================================================
' HAUPTPROZEDUR: Formatiert ALLE relevanten Tabellen neu
' Diese Prozedur wird von mod_Mitglieder_UI und frm_Mitgliedsdaten aufgerufen
' ===============================================================
Public Sub Formatiere_Alle_Tabellen_Neu()
    
    Dim wsD As Worksheet
    Dim wsBK As Worksheet
    Dim lastRowD As Long
    Dim lastRowBK As Long
    Dim euroFormat As String
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Euro-Format mit ChrW fuer korrektes Unicode-Zeichen
    euroFormat = "#,##0.00 " & ChrW(8364)
    
    ' === DATEN-BLATT FORMATIEREN ===
    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo ErrorHandler
    
    If Not wsD Is Nothing Then
        On Error Resume Next
        wsD.Unprotect PASSWORD:=PASSWORD
        On Error GoTo ErrorHandler
        
        ' Gesamtes Blatt: Vertikal zentriert
        wsD.Cells.VerticalAlignment = xlCenter
        
        ' Kategorie-Tabelle formatieren
        Call FormatiereKategorieTabelle(wsD)
        
        ' EntityKey-Tabelle formatieren
        Call FormatiereEntityKeyTabelleKomplett(wsD)
        
        ' DropDown-Listen aktualisieren
        Call AktualisiereKategorieDropdownListen(wsD)
        
        wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    
    ' === BANKKONTO-BLATT FORMATIEREN ===
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo ErrorHandler
    
    If Not wsBK Is Nothing Then
        On Error Resume Next
        wsBK.Unprotect PASSWORD:=PASSWORD
        On Error GoTo ErrorHandler
        
        lastRowBK = wsBK.Cells(wsBK.Rows.Count, BK_COL_DATUM).End(xlUp).Row
        If lastRowBK < BK_START_ROW Then lastRowBK = BK_START_ROW
        
        ' Spalte L (Bemerkung): Textumbruch
        With wsBK.Range(wsBK.Cells(BK_START_ROW, BK_COL_BEMERKUNG), _
                        wsBK.Cells(lastRowBK, BK_COL_BEMERKUNG))
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        
        ' Zeilenhoehe AutoFit
        wsBK.Rows(BK_START_ROW & ":" & lastRowBK).AutoFit
        
        ' Waehrungsformat mit Euro-Zeichen
        ' Spalte B (Betrag)
        wsBK.Range(wsBK.Cells(BK_START_ROW, BK_COL_BETRAG), _
                   wsBK.Cells(lastRowBK, BK_COL_BETRAG)).NumberFormat = euroFormat
        
        ' Spalten M-Z (Einnahmen und Ausgaben)
        wsBK.Range(wsBK.Cells(BK_START_ROW, BK_COL_MITGL_BEITR), _
                   wsBK.Cells(lastRowBK, BK_COL_AUSZAHL_KASSE)).NumberFormat = euroFormat
        
        wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
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
    
    ' 1. Gesamtes Blatt: Vertikal zentriert
    ws.Cells.VerticalAlignment = xlCenter
    
    ' 2. Kategorie-Tabelle formatieren
    Call FormatiereKategorieTabelle(ws)
    
    ' 3. EntityKey-Tabelle formatieren (falls vorhanden)
    Call FormatiereEntityKeyTabelleKomplett(ws)
    
    ' 4. DropDown-Listen fuer Kategorien aktualisieren
    Call AktualisiereKategorieDropdownListen(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Formatierung des Daten-Blatts abgeschlossen!" & vbCrLf & vbCrLf & _
           "- Alle Zellen vertikal zentriert" & vbCrLf & _
           "- Kategorie-Tabelle formatiert" & vbCrLf & _
           "- EntityKey-Tabelle formatiert" & vbCrLf & _
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
' KATEGORIE-TABELLE FORMATIEREN (J-P)
' ===============================================================
Public Sub FormatiereKategorieTabelle(Optional ByRef ws As Worksheet = Nothing)
    
    Dim lastRow As Long
    Dim rngTable As Range
    Dim r As Long
    Dim einAusWert As String
    
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    lastRow = ws.Cells(ws.Rows.Count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    
    Set rngTable = ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), _
                            ws.Cells(lastRow, DATA_CAT_COL_END))
    
    ' Rahmen fuer gesamte Tabelle (innen und aussen)
    With rngTable.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rngTable.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rngTable.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rngTable.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rngTable.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rngTable.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' Vertikal zentriert
    rngTable.VerticalAlignment = xlCenter
    
    ' Spalte J (Kategorie): Linksbuendig
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KATEGORIE), _
             ws.Cells(lastRow, DATA_CAT_COL_KATEGORIE)).HorizontalAlignment = xlLeft
    
    ' Spalte K (Einnahme/Ausgabe): Zentriert
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_EINAUS), _
             ws.Cells(lastRow, DATA_CAT_COL_EINAUS)).HorizontalAlignment = xlCenter
    
    ' Spalte L (Keyword): Linksbuendig
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KEYWORD), _
             ws.Cells(lastRow, DATA_CAT_COL_KEYWORD)).HorizontalAlignment = xlLeft
    
    ' Spalte M (Prioritaet): Zentriert
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_PRIORITAET), _
             ws.Cells(lastRow, DATA_CAT_COL_PRIORITAET)).HorizontalAlignment = xlCenter
    
    ' Spalte N (Zielspalte): Linksbuendig + DropDown
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_ZIELSPALTE), _
             ws.Cells(lastRow, DATA_CAT_COL_ZIELSPALTE)).HorizontalAlignment = xlLeft
    
    ' Spalte O (Faelligkeit): Linksbuendig
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_FAELLIGKEIT), _
             ws.Cells(lastRow, DATA_CAT_COL_FAELLIGKEIT)).HorizontalAlignment = xlLeft
    
    ' Spalte P (Kommentar): Linksbuendig
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_KOMMENTAR), _
             ws.Cells(lastRow, DATA_CAT_COL_KOMMENTAR)).HorizontalAlignment = xlLeft
    
    ' DropDown fuer Zielspalte (N) basierend auf E/A in Spalte K
    For r = DATA_START_ROW To lastRow
        einAusWert = UCase(Trim(ws.Cells(r, DATA_CAT_COL_EINAUS).value))
        Call SetzeZielspalteDropdown(ws, r, einAusWert)
    Next r
    
    ' AutoFit Spaltenbreiten
    ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), _
             ws.Cells(lastRow, DATA_CAT_COL_END)).EntireColumn.AutoFit
    
End Sub

' ===============================================================
' ZIELSPALTE-DROPDOWN SETZEN (abhaengig von E/A)
' ===============================================================
Private Sub SetzeZielspalteDropdown(ByRef ws As Worksheet, ByVal zeile As Long, ByVal einAus As String)
    
    Dim dropdownSource As String
    
    On Error Resume Next
    ws.Cells(zeile, DATA_CAT_COL_ZIELSPALTE).Validation.Delete
    On Error GoTo 0
    
    ' Erstelle DropDown basierend auf E/A
    Select Case einAus
        Case "E"
            ' Einnahmen: Spalten M-S auf Bankkonto! (Zeile 27)
            dropdownSource = "=" & WS_BANKKONTO & "!$M$27:$S$27"
        Case "A"
            ' Ausgaben: Spalten T-Z auf Bankkonto! (Zeile 27)
            dropdownSource = "=" & WS_BANKKONTO & "!$T$27:$Z$27"
        Case Else
            ' Alles: Spalten M-Z auf Bankkonto! (Zeile 27)
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
    
    lastRow = ws.Cells(ws.Rows.Count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    
    ' Sammle eindeutige Kategorien nach E/A
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
    
    ' Loesche alte Listen (AF ab Zeile 4, AG ab Zeile 4, AH ab Zeile 4)
    On Error Resume Next
    ws.Range("AF4:AF1000").ClearContents
    ws.Range("AG4:AG1000").ClearContents
    ws.Range("AH4:AH1000").ClearContents
    On Error GoTo 0
    
    ' Schreibe Einnahmen-Kategorien nach AF (Spalte 32)
    nextRowE = 4
    For Each key In dictEinnahmen.Keys
        ws.Cells(nextRowE, DATA_COL_EINNAHMEN).value = key
        nextRowE = nextRowE + 1
    Next key
    
    ' Schreibe Ausgaben-Kategorien nach AG (Spalte 33)
    nextRowA = 4
    For Each key In dictAusgaben.Keys
        ws.Cells(nextRowA, DATA_COL_AUSGABEN).value = key
        nextRowA = nextRowA + 1
    Next key
    
    ' Schreibe Monat/Periode nach AH (Spalte 34)
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
    
    ' Erstelle/aktualisiere Named Ranges
    Call ErstelleKategorieNamedRanges(ws, nextRowE - 1, nextRowA - 1)
    
End Sub

' ===============================================================
' NAMED RANGES FUER KATEGORIEN ERSTELLEN
' ===============================================================
Private Sub ErstelleKategorieNamedRanges(ByRef ws As Worksheet, ByVal lastRowE As Long, ByVal lastRowA As Long)
    
    On Error Resume Next
    
    ' Alte Named Ranges loeschen falls vorhanden
    ThisWorkbook.Names("lst_KategorienEinnahmen").Delete
    ThisWorkbook.Names("lst_KategorienAusgaben").Delete
    ThisWorkbook.Names("lst_MonatPeriode").Delete
    
    ' lst_KategorienEinnahmen
    If lastRowE >= 4 Then
        ThisWorkbook.Names.Add Name:="lst_KategorienEinnahmen", _
            RefersTo:="=" & ws.Name & "!$AF$4:$AF$" & lastRowE
    Else
        ThisWorkbook.Names.Add Name:="lst_KategorienEinnahmen", _
            RefersTo:="=" & ws.Name & "!$AF$4"
    End If
    
    ' lst_KategorienAusgaben
    If lastRowA >= 4 Then
        ThisWorkbook.Names.Add Name:="lst_KategorienAusgaben", _
            RefersTo:="=" & ws.Name & "!$AG$4:$AG$" & lastRowA
    Else
        ThisWorkbook.Names.Add Name:="lst_KategorienAusgaben", _
            RefersTo:="=" & ws.Name & "!$AG$4"
    End If
    
    ' lst_MonatPeriode
    ThisWorkbook.Names.Add Name:="lst_MonatPeriode", _
        RefersTo:="=" & ws.Name & "!$AH$4:$AH$15"
    
    On Error GoTo 0
    
End Sub

' ===============================================================
' ENTITYKEY-TABELLE KOMPLETT FORMATIEREN
' ===============================================================
Public Sub FormatiereEntityKeyTabelleKomplett(Optional ByRef ws As Worksheet = Nothing)
    
    Dim lastRow As Long
    Dim rngTable As Range
    Dim rngZebra As Range
    Dim r As Long
    
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    lastRow = ws.Cells(ws.Rows.Count, EK_COL_IBAN).End(xlUp).Row
    If lastRow < EK_START_ROW Then Exit Sub
    
    Set rngTable = ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                            ws.Cells(lastRow, EK_COL_DEBUG))
    
    ' Rahmen (innen und aussen)
    With rngTable.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rngTable.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rngTable.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rngTable.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rngTable.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rngTable.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' Vertikal zentriert
    rngTable.VerticalAlignment = xlCenter
    
    ' Textumbruch fuer Spalten S-X (IBAN bis Debug)
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_IBAN), _
             ws.Cells(lastRow, EK_COL_DEBUG)).WrapText = True
    
    ' Spalte R (EntityKey): Kein Textumbruch, linksbuendig
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                  ws.Cells(lastRow, EK_COL_ENTITYKEY))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ENTITYKEY).ColumnWidth = 14
    
    ' Spalte V + W: Zentriert
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_PARZELLE), _
             ws.Cells(lastRow, EK_COL_PARZELLE)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_ROLE), _
             ws.Cells(lastRow, EK_COL_ROLE)).HorizontalAlignment = xlCenter
    
    ' Zebra-Formatierung fuer Spalten R-T
    For r = EK_START_ROW To lastRow
        Set rngZebra = ws.Range(ws.Cells(r, EK_COL_ENTITYKEY), ws.Cells(r, EK_COL_KONTONAME))
        
        If (r - EK_START_ROW) Mod 2 = 1 Then
            rngZebra.Interior.color = ZEBRA_COLOR
        Else
            rngZebra.Interior.ColorIndex = xlNone
        End If
    Next r
    
    ' AutoFit
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_IBAN), _
             ws.Cells(lastRow, EK_COL_DEBUG)).EntireColumn.AutoFit
    ws.Rows(EK_START_ROW & ":" & lastRow).AutoFit
    
End Sub

' ===============================================================
' BANKKONTO-BLATT FORMATIEREN (Punkt 6 + 7)
' ===============================================================
Public Sub FormatiereBlattBankkonto()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim euroFormat As String
    
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    ' Euro-Format mit ChrW fuer korrektes Unicode-Zeichen
    euroFormat = "#,##0.00 " & ChrW(8364)
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    lastRow = ws.Cells(ws.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then lastRow = BK_START_ROW
    
    ' Spalte L (Bemerkung): Textumbruch ab Zeile 28
    With ws.Range(ws.Cells(BK_START_ROW, BK_COL_BEMERKUNG), _
                  ws.Cells(lastRow, BK_COL_BEMERKUNG))
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    
    ' Zeilenhoehe AutoFit
    ws.Rows(BK_START_ROW & ":" & lastRow).AutoFit
    
    ' Waehrungsformat mit Euro-Zeichen
    ' Spalte B
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_BETRAG), _
             ws.Cells(lastRow, BK_COL_BETRAG)).NumberFormat = euroFormat
    
    ' Spalten M-Z
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_MITGL_BEITR), _
             ws.Cells(lastRow, BK_COL_AUSZAHL_KASSE)).NumberFormat = euroFormat
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.ScreenUpdating = True
    
    MsgBox "Formatierung des Bankkonto-Blatts abgeschlossen!" & vbCrLf & vbCrLf & _
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
    
    ' Nur reagieren wenn Aenderung in Kategorie-Tabelle (Spalte J oder K)
    If Target.Column = DATA_CAT_COL_KATEGORIE Or Target.Column = DATA_CAT_COL_EINAUS Then
        ' DropDown-Listen aktualisieren
        Call AktualisiereKategorieDropdownListen(ws)
    End If
    
    ' Wenn E/A geaendert wurde, Zielspalte-DropDown aktualisieren
    If Target.Column = DATA_CAT_COL_EINAUS Then
        Dim einAus As String
        einAus = UCase(Trim(Target.value))
        Call SetzeZielspalteDropdown(ws, Target.Row, einAus)
    End If
    
End Sub

