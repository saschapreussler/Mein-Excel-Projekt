Attribute VB_Name = "mod_Formatierung"
Option Explicit

' ***************************************************************
' MODUL: mod_Formatierung
' ZWECK: Formatierung und DropDown-Listen-Verwaltung
' VERSION: 1.5 - 02.02.2026
' KORREKTUR: SetzeZellschutzFuerZeile hinzugefuegt
' ***************************************************************

Private Const ZEBRA_COLOR As Long = &HDEE5E3

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
        
        Call FormatiereKategorieTabelle(wsD)
        Call FormatiereEntityKeyTabelleKomplett(wsD)
        Call AktualisiereKategorieDropdownListen(wsD)
        
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
    
    Call FormatiereKategorieTabelle(ws)
    Call FormatiereEntityKeyTabelleKomplett(ws)
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
    
    lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    
    Set rngTable = ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), _
                            ws.Cells(lastRow, DATA_CAT_COL_END))
    
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
' Ermittelt lastRow automatisch und ruft FormatiereEntityKeyTabelle auf
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
' ===============================================================
Private Sub FormatiereEntityKeyTabelle(ByRef ws As Worksheet, ByVal lastRow As Long)
    
    Dim rngTable As Range
    Dim rngZebra As Range
    Dim r As Long
    Dim currentRole As String
    
    If lastRow < EK_START_ROW Then Exit Sub
    
    Set rngTable = ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                            ws.Cells(lastRow, EK_COL_DEBUG))
    
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    rngTable.VerticalAlignment = xlCenter
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                  ws.Cells(lastRow, EK_COL_ENTITYKEY))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ENTITYKEY).ColumnWidth = 9
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_IBAN), _
                  ws.Cells(lastRow, EK_COL_IBAN))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_IBAN).ColumnWidth = 23
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_KONTONAME), _
                  ws.Cells(lastRow, EK_COL_KONTONAME))
        .WrapText = True
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_KONTONAME).ColumnWidth = 50
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ZUORDNUNG), _
                  ws.Cells(lastRow, EK_COL_ZUORDNUNG))
        .WrapText = True
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ZUORDNUNG).ColumnWidth = 30
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_PARZELLE), _
                  ws.Cells(lastRow, EK_COL_PARZELLE))
        .WrapText = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns(EK_COL_PARZELLE).ColumnWidth = 10
    
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
    ws.Columns(EK_COL_DEBUG).AutoFit
    
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
             ws.Cells(lastRow, EK_COL_KONTONAME)).Locked = True
    
    For r = EK_START_ROW To lastRow
        currentRole = Trim(ws.Cells(r, EK_COL_ROLE).value)
        
        Call SetzeZellschutzFuerZeile(ws, r, currentRole)
        
        Set rngZebra = ws.Range(ws.Cells(r, EK_COL_ENTITYKEY), ws.Cells(r, EK_COL_KONTONAME))
        
        If (r - EK_START_ROW) Mod 2 = 1 Then
            rngZebra.Interior.color = ZEBRA_COLOR
        Else
            rngZebra.Interior.ColorIndex = xlNone
        End If
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
    
    ' Spalten R-T (EntityKey, IBAN, Kontoname) sind immer gesperrt
    Set rngGesperrt = ws.Range(ws.Cells(zeile, EK_COL_ENTITYKEY), ws.Cells(zeile, EK_COL_KONTONAME))
    rngGesperrt.Locked = True
    
    ' Spalten U-X (Zuordnung, Parzelle, Role, Debug) abhaengig von Role
    Set rngEditierbar = ws.Range(ws.Cells(zeile, EK_COL_ZUORDNUNG), ws.Cells(zeile, EK_COL_DEBUG))
    
    Select Case UCase(Trim(currentRole))
        Case "MITGLIED"
            ' Bei MITGLIED: Zuordnung, Parzelle, Role gesperrt; nur Debug editierbar
            ws.Cells(zeile, EK_COL_ZUORDNUNG).Locked = True
            ws.Cells(zeile, EK_COL_PARZELLE).Locked = True
            ws.Cells(zeile, EK_COL_ROLE).Locked = True
            ws.Cells(zeile, EK_COL_DEBUG).Locked = False
            
        Case "UNBEKANNT", ""
            ' Bei UNBEKANNT: Alles editierbar fuer manuelle Zuordnung
            rngEditierbar.Locked = False
            
        Case Else
            ' Bei anderen Roles (DIENSTLEISTER, BEHOERDE, etc.): Role gesperrt, Rest editierbar
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



'================================================================================
' Prozedur: ApplyDatenSheetFormatting
' Zweck:    Wendet einheitliche Zebra-Streifen, Rahmen, Spaltenbreiten und
'           AutoFilter auf ALLE Tabellen des "Daten"-Blatts an
' Autor:    [Platzhalter]
' Datum:    06.02.2026
' Hinweise:
'   - Erhält Ampel-Formatierung in Spalten U-X (KEIN Zebra dort!)
'   - Liest Zebra-Farben aus Mitgliederliste (Zeile 6 und 7)
'   - Formatiert: B, D, F, H, J-P, R-T, Z-AH
'   - Spalten U-X: NUR Rahmen, KEIN Zebra
'   - Y100 WrapText wird erhalten
'
' Tabellen-Bereiche:
'   - Einzelspalten: B, D, F, H
'   - Kategorie: J-P (7 Spalten)
'   - EntityKey: R-X (7 Spalten, aber U-X ohne Zebra)
'   - Helper: Z-AH (9 Spalten)
'
' Spaltenbreiten (aus Dokumentation):
'   - AutoFit: B, D, F, H, J, K, L, M, N, O, P, S, W, Z-AH
'   - Fest 11.00: R
'   - Fest 36.00: T
'   - Fest 28.00: U
'   - Fest 9.00: V
'   - Fest 42.00: X
'
' Ausrichtung:
'   - Zentriert: F, K, M, V
'   - Links: Alle anderen
'================================================================================
Public Sub ApplyDatenSheetFormatting(Optional ByVal formatType As String = "ALL")
    
    ' === VARIABLENDEKLARATION ===
    Dim ws As Worksheet
    Dim wsRef As Worksheet
    Dim lastRowB As Long, lastRowD As Long, lastRowF As Long, lastRowH As Long
    Dim lastRowKat As Long, lastRowEK As Long, lastRowHelper As Long
    Dim r As Long
    Dim zebraColor1 As Long  ' Farbe für gerade Zeilen (4, 6, 8...)
    Dim zebraColor2 As Long  ' Farbe für ungerade Zeilen (5, 7, 9...)
    Dim rngRow As Range
    
    On Error GoTo ErrorHandler
    
    ' === PERFORMANCE-OPTIMIERUNG ===
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' === ARBEITSBLÄTTER REFERENZIEREN ===
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    ' Zebra-Farben aus Mitgliederliste lesen
    On Error Resume Next
    Set wsRef = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If wsRef Is Nothing Then
        Set wsRef = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    End If
    On Error GoTo ErrorHandler
    
    ' Farben aus Referenz-Blatt ermitteln (Zeile 6 = gerade, Zeile 7 = ungerade)
    If Not wsRef Is Nothing Then
        zebraColor1 = wsRef.Range("A6").Interior.color  ' Gerade Zeilen
        zebraColor2 = wsRef.Range("A7").Interior.color  ' Ungerade Zeilen
        ' Falls keine Farbe gesetzt, Standardwerte verwenden
        If zebraColor1 = 16777215 Then zebraColor1 = RGB(255, 255, 255)  ' Weiß
        If zebraColor2 = 16777215 Then zebraColor2 = ZEBRA_COLOR         ' Hellgrün
    Else
        ' Fallback-Farben
        zebraColor1 = RGB(255, 255, 255)  ' Weiß
        zebraColor2 = ZEBRA_COLOR          ' Hellgrün (&HDEE5E3)
    End If
    
    Debug.Print "=== ApplyDatenSheetFormatting gestartet ==="
    Debug.Print "Zebra-Farbe 1 (gerade): " & Hex(zebraColor1)
    Debug.Print "Zebra-Farbe 2 (ungerade): " & Hex(zebraColor2)
    
    ' === BLATTSCHUTZ AUFHEBEN ===
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    ' === LETZTE ZEILEN ERMITTELN ===
    lastRowB = GetLastDataRowSafe(ws, 2)        ' Spalte B
    lastRowD = GetLastDataRowSafe(ws, 4)        ' Spalte D
    lastRowF = GetLastDataRowSafe(ws, 6)        ' Spalte F
    lastRowH = GetLastDataRowSafe(ws, 8)        ' Spalte H
    lastRowKat = GetLastDataRowSafe(ws, 10)     ' Spalte J (Kategorie)
    lastRowEK = GetLastDataRowSafe(ws, 19)      ' Spalte S (EntityKey IBAN)
    lastRowHelper = GetLastDataRowSafe(ws, 26)  ' Spalte Z (Helper)
    
    Debug.Print "LastRows: B=" & lastRowB & ", D=" & lastRowD & ", F=" & lastRowF & _
                ", H=" & lastRowH & ", Kat=" & lastRowKat & ", EK=" & lastRowEK & ", Helper=" & lastRowHelper
    
    ' ===================================================================
    ' SCHRITT 1: SPALTENBREITEN SETZEN
    ' ===================================================================
    If formatType = "ALL" Or formatType = "WIDTHS" Then
        Debug.Print "Setze Spaltenbreiten..."
        
        ' Feste Breiten (laut Dokumentation)
        ws.Columns("R").ColumnWidth = 11#       ' EntityKey
        ws.Columns("T").ColumnWidth = 36#       ' Zahler/Empfänger
        ws.Columns("U").ColumnWidth = 28#       ' Zuordnung
        ws.Columns("V").ColumnWidth = 9#        ' Parzelle
        ws.Columns("X").ColumnWidth = 42#       ' Debug
        
        ' AutoFit Spalten
        ws.Columns("B").AutoFit    ' Vereinsfunktionen
        ws.Columns("D").AutoFit    ' Anredeformen
        ws.Columns("F").AutoFit    ' Parzelle
        ws.Columns("H").AutoFit    ' Seite
        ws.Columns("J").AutoFit    ' Kategorie
        ws.Columns("K").AutoFit    ' Einnahme/Ausgabe
        ws.Columns("L").AutoFit    ' Keyword
        ws.Columns("M").AutoFit    ' Priorität
        ws.Columns("N").AutoFit    ' Zielspalte
        ws.Columns("O").AutoFit    ' Fälligkeit
        ws.Columns("P").AutoFit    ' Kommentar
        ws.Columns("S").AutoFit    ' IBAN
        ws.Columns("W").AutoFit    ' EntityRole
        ws.Columns("Z").AutoFit    ' Helper Z
        ws.Columns("AA").AutoFit   ' Helper AA
        ws.Columns("AB").AutoFit   ' Helper AB
        ws.Columns("AC").AutoFit   ' Helper AC
        ws.Columns("AD").AutoFit   ' Helper AD
        ws.Columns("AE").AutoFit   ' Helper AE
        ws.Columns("AF").AutoFit   ' Kat Einnahme
        ws.Columns("AG").AutoFit   ' Kat Ausgabe
        ws.Columns("AH").AutoFit   ' Monat/Periode
    End If
    
    ' ===================================================================
    ' SCHRITT 2: TEXTAUSRICHTUNG SETZEN
    ' ===================================================================
    If formatType = "ALL" Or formatType = "ALIGNMENT" Then
        Debug.Print "Setze Textausrichtung..."
        
        ' Zentrierte Spalten: F, K, M, V
        ws.Columns("F").HorizontalAlignment = xlCenter
        ws.Columns("K").HorizontalAlignment = xlCenter
        ws.Columns("M").HorizontalAlignment = xlCenter
        ws.Columns("V").HorizontalAlignment = xlCenter
        
        ' Linksbündige Spalten (explizit setzen)
        ws.Columns("B").HorizontalAlignment = xlLeft
        ws.Columns("D").HorizontalAlignment = xlLeft
        ws.Columns("H").HorizontalAlignment = xlLeft
        ws.Columns("J").HorizontalAlignment = xlLeft
        ws.Columns("L").HorizontalAlignment = xlLeft
        ws.Columns("N").HorizontalAlignment = xlLeft
        ws.Columns("O").HorizontalAlignment = xlLeft
        ws.Columns("P").HorizontalAlignment = xlLeft
        ws.Columns("R").HorizontalAlignment = xlLeft
        ws.Columns("S").HorizontalAlignment = xlLeft
        ws.Columns("T").HorizontalAlignment = xlLeft
        ws.Columns("U").HorizontalAlignment = xlLeft
        ws.Columns("W").HorizontalAlignment = xlLeft
        ws.Columns("X").HorizontalAlignment = xlLeft
        
        ' Vertikale Zentrierung für alle Zellen
        ws.Cells.VerticalAlignment = xlCenter
    End If
    
    ' ===================================================================
    ' SCHRITT 3: ZEBRA-STREIFEN ANWENDEN
    ' ===================================================================
    If formatType = "ALL" Or formatType = "ZEBRA" Then
        Debug.Print "Wende Zebra-Streifen an..."
        
        ' --- Tabelle B (Vereinsfunktionen) ---
        If lastRowB >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowB
                If (r Mod 2) = 0 Then
                    ws.Cells(r, 2).Interior.color = zebraColor1
                Else
                    ws.Cells(r, 2).Interior.color = zebraColor2
                End If
            Next r
        End If
        
        ' --- Tabelle D (Anredeformen) ---
        If lastRowD >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowD
                If (r Mod 2) = 0 Then
                    ws.Cells(r, 4).Interior.color = zebraColor1
                Else
                    ws.Cells(r, 4).Interior.color = zebraColor2
                End If
            Next r
        End If
        
        ' --- Tabelle F (Parzelle) ---
        If lastRowF >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowF
                If (r Mod 2) = 0 Then
                    ws.Cells(r, 6).Interior.color = zebraColor1
                Else
                    ws.Cells(r, 6).Interior.color = zebraColor2
                End If
            Next r
        End If
        
        ' --- Tabelle H (Seite) ---
        If lastRowH >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowH
                If (r Mod 2) = 0 Then
                    ws.Cells(r, 8).Interior.color = zebraColor1
                Else
                    ws.Cells(r, 8).Interior.color = zebraColor2
                End If
            Next r
        End If
        
        ' --- Kategorie-Tabelle J-P (7 Spalten) ---
        If lastRowKat >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowKat
                Set rngRow = ws.Range(ws.Cells(r, 10), ws.Cells(r, 16))  ' J-P
                If (r Mod 2) = 0 Then
                    rngRow.Interior.color = zebraColor1
                Else
                    rngRow.Interior.color = zebraColor2
                End If
            Next r
        End If
        
        ' --- EntityKey-Tabelle R-T NUR (3 Spalten) - NICHT U-X! ---
        If lastRowEK >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowEK
                Set rngRow = ws.Range(ws.Cells(r, 18), ws.Cells(r, 20))  ' R-T nur
                If (r Mod 2) = 0 Then
                    rngRow.Interior.color = zebraColor1
                Else
                    rngRow.Interior.color = zebraColor2
                End If
            Next r
            ' WICHTIG: Spalten U-X (21-24) erhalten KEIN Zebra!
            ' Diese haben Ampel-Formatierung
            Debug.Print "HINWEIS: Spalten U-X wurden NICHT mit Zebra formatiert (Ampel-Bereich)"
        End If
        
        ' --- Helper-Spalten Z-AH (9 Spalten) ---
        If lastRowHelper >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowHelper
                Set rngRow = ws.Range(ws.Cells(r, 26), ws.Cells(r, 34))  ' Z-AH
                If (r Mod 2) = 0 Then
                    rngRow.Interior.color = zebraColor1
                Else
                    rngRow.Interior.color = zebraColor2
                End If
            Next r
        End If
    End If
    
    ' ===================================================================
    ' SCHRITT 4: RAHMEN ANWENDEN
    ' ===================================================================
    If formatType = "ALL" Or formatType = "BORDERS" Then
        Debug.Print "Wende Rahmen an..."
        
        ' Tabelle B
        If lastRowB >= DATA_START_ROW Then
            Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowB, 2, 2)
        End If
        
        ' Tabelle D
        If lastRowD >= DATA_START_ROW Then
            Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowD, 4, 4)
        End If
        
        ' Tabelle F
        If lastRowF >= DATA_START_ROW Then
            Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowF, 6, 6)
        End If
        
        ' Tabelle H
        If lastRowH >= DATA_START_ROW Then
            Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowH, 8, 8)
        End If
        
        ' Kategorie-Tabelle J-P
        If lastRowKat >= DATA_START_ROW Then
            Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowKat, 10, 16)
        End If
        
        ' EntityKey-Tabelle R-X (ALLE 7 Spalten bekommen Rahmen!)
        If lastRowEK >= DATA_START_ROW Then
            Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowEK, 18, 24)
        End If
        
        ' Helper-Spalten Z-AH
        If lastRowHelper >= DATA_START_ROW Then
            Call ApplyBordersToTableRange(ws, DATA_START_ROW, lastRowHelper, 26, 34)
        End If
    End If
    
    ' ===================================================================
    ' SCHRITT 5: AUTOFILTER AKTIVIEREN
    ' ===================================================================
    If formatType = "ALL" Or formatType = "FILTER" Then
        Debug.Print "Aktiviere AutoFilter..."
        
        ' Bestehende AutoFilter entfernen
        On Error Resume Next
        ws.AutoFilterMode = False
        On Error GoTo ErrorHandler
        
        ' AutoFilter für jede Tabelle separat (nur wenn Daten vorhanden)
        
        ' Tabelle B
        If lastRowB >= DATA_START_ROW Then
            ws.Range(ws.Cells(DATA_HEADER_ROW, 2), ws.Cells(lastRowB, 2)).AutoFilter
        End If
        
        ' Tabelle D
        If lastRowD >= DATA_START_ROW Then
            ws.Range(ws.Cells(DATA_HEADER_ROW, 4), ws.Cells(lastRowD, 4)).AutoFilter
        End If
        
        ' Tabelle F
        If lastRowF >= DATA_START_ROW Then
            ws.Range(ws.Cells(DATA_HEADER_ROW, 6), ws.Cells(lastRowF, 6)).AutoFilter
        End If
        
        ' Tabelle H
        If lastRowH >= DATA_START_ROW Then
            ws.Range(ws.Cells(DATA_HEADER_ROW, 8), ws.Cells(lastRowH, 8)).AutoFilter
        End If
        
        ' Kategorie J-P
        If lastRowKat >= DATA_START_ROW Then
            ws.Range(ws.Cells(DATA_HEADER_ROW, 10), ws.Cells(lastRowKat, 16)).AutoFilter
        End If
        
        ' EntityKey R-X
        If lastRowEK >= DATA_START_ROW Then
            ws.Range(ws.Cells(DATA_HEADER_ROW, 18), ws.Cells(lastRowEK, 24)).AutoFilter
        End If
        
        ' Helper Z-AH
        If lastRowHelper >= DATA_START_ROW Then
            ws.Range(ws.Cells(DATA_HEADER_ROW, 26), ws.Cells(lastRowHelper, 34)).AutoFilter
        End If
    End If
    
    ' ===================================================================
    ' SCHRITT 6: SPEZIELLE FORMATIERUNGEN ERHALTEN
    ' ===================================================================
    ' Y100 WrapText erhalten
    On Error Resume Next
    ws.Range("Y100").WrapText = True
    On Error GoTo ErrorHandler
    
    ' ===================================================================
    ' CLEANUP
    ' ===================================================================
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Debug.Print "=== ApplyDatenSheetFormatting erfolgreich beendet ==="
    
    MsgBox "Formatierung des Daten-Blatts abgeschlossen!" & vbCrLf & vbCrLf & _
           "Angewendet:" & vbCrLf & _
           "- Spaltenbreiten (fest + AutoFit)" & vbCrLf & _
           "- Textausrichtung (zentriert/links)" & vbCrLf & _
           "- Zebra-Streifen (außer U-X)" & vbCrLf & _
           "- Rahmen für alle Tabellen" & vbCrLf & _
           "- AutoFilter auf Zeile 3", vbInformation, "Formatierung"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "FEHLER in ApplyDatenSheetFormatting: " & Err.Description
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Formatierungsfehler: " & Err.Description, vbCritical, "Fehler"
End Sub

'================================================================================
' Hilfsfunktion: Ermittelt die letzte Datenzeile einer Spalte (sicher)
'================================================================================
Private Function GetLastDataRowSafe(ByRef ws As Worksheet, ByVal col As Long) As Long
    Dim lastRow As Long
    
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
    On Error GoTo 0
    
    ' Wenn nur Header oder leer, DATA_HEADER_ROW zurückgeben
    If lastRow < DATA_START_ROW Then lastRow = DATA_HEADER_ROW
    
    GetLastDataRowSafe = lastRow
End Function

'================================================================================
' Hilfsprozedur: Wendet Rahmen auf einen Tabellenbereich an
'================================================================================
Private Sub ApplyBordersToTableRange(ByRef ws As Worksheet, _
                                      ByVal startRow As Long, _
                                      ByVal endRow As Long, _
                                      ByVal startCol As Long, _
                                      ByVal endCol As Long)
    
    Dim rngTable As Range
    
    On Error Resume Next
    
    If endRow < startRow Then Exit Sub
    
    Set rngTable = ws.Range(ws.Cells(startRow, startCol), ws.Cells(endRow, endCol))
    
    ' Alle Rahmenlinien setzen
    With rngTable.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = RGB(0, 0, 0)
    End With
    
    With rngTable.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = RGB(0, 0, 0)
    End With
    
    With rngTable.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = RGB(0, 0, 0)
    End With
    
    With rngTable.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = RGB(0, 0, 0)
    End With
    
    With rngTable.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = RGB(0, 0, 0)
    End With
    
    With rngTable.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = RGB(0, 0, 0)
    End With
    
    On Error GoTo 0
End Sub




