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





