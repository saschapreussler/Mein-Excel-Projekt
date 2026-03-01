Attribute VB_Name = "mod_ZaehlerLogik"
Option Explicit

' ==========================================================
' MODUL: mod_ZaehlerLogik (ORCHESTRATOR)
' VERSION: 2.0 - Modularisiert
' ÄNDERUNG v2.0:
'   - Berechnung ausgelagert nach mod_Zaehler_Berechnung
'   - Historie ausgelagert nach mod_Zaehler_Historie
'   - Dieses Modul: Konstanten, Hilfsfunktionen, Initialisierung, KPIs
' ==========================================================

' ==========================================================
' KONSTANTEN & VARIABLEN
' ==========================================================
Public IsZaehlerLogicRunning As Boolean
Private Const HIST_SHEET_NAME As String = "Zählerhistorie"
Private Const HIST_TABLE_NAME As String = "Tabelle_Zaehlerhistorie"
Private Const PASSWORD As String = ""
Private Const STR_HISTORY_SEPARATOR As String = "--- Zählerhistorie Makro-Eintrag ---"

Private Const COL_STAND_ANFANG As String = "B"
Private Const COL_STAND_ENDE As String = "C"
Private Const COL_VERBRAUCH_GESAMT As String = "D"
Private Const COL_BEMERKUNG As String = "E"

Private Const COL_HIST_ID As Long = 1
Private Const COL_HIST_DATUM As Long = 2
Private Const COL_HIST_PARZELLE As Long = 3
Private Const COL_HIST_MEDIUM As Long = 4
Private Const COL_HIST_ZAEHLER_ALT As Long = 5
Private Const COL_HIST_STAND_ALT_ANFANG As Long = 6
Private Const COL_HIST_STAND_ALT_ENDE As Long = 7
Private Const COL_HIST_ZAEHLER_NEU As Long = 8
Private Const COL_HIST_STAND_NEU_START As Long = 9
Private Const COL_HIST_VERBRAUCH_ALT As Long = 10
Private Const COL_HIST_BEMERKUNG As Long = 11

Private Const RGB_STROM As Long = 86271
Private Const RGB_WASSER As Long = 16737792
Private Const RGB_EINGABE_ERLAUBT As Long = 7592334
Private Const RGB_GEWECHSELT As Long = 4980735
Private Const RGB_HEADER_BG As Long = 13619148


' ==========================================================
' 0. HILFSFUNKTIONEN (Public für Sub-Module)
' ==========================================================

' Zeilenhöhe auf Minimum sicherstellen
Public Sub EnsureMinRowHeight(ws As Worksheet, targetRow As Long)
    Const MIN_HEIGHT As Double = 50
    On Error Resume Next
    If ws.Rows(targetRow).RowHeight < MIN_HEIGHT Then
        ws.Rows(targetRow).RowHeight = MIN_HEIGHT
    End If
    On Error GoTo 0
End Sub

' Hole Namen für Parzelle
Public Function HoleNamenFuerParzelle(ws As Worksheet, suchParzelle As String, maxZ As Long) As String
    Dim r As Long
    Dim sTemp As String
    Dim PVal As String
    
    sTemp = ""
    
    For r = 6 To maxZ
        PVal = Trim(CStr(ws.Cells(r, 2).value))
        
        If StrComp(PVal, suchParzelle, vbTextCompare) = 0 Then
            If sTemp <> "" Then sTemp = sTemp & vbLf
            sTemp = sTemp & Trim(ws.Cells(r, 6).value) & " " & Trim(ws.Cells(r, 5).value)
        End If
    Next r
    
    HoleNamenFuerParzelle = sTemp
End Function

' Ermittlung der Zielzeile
Public Function GetTargetRow(ByVal ZaehlerName As String, ByVal Medium As String) As Long
    Dim idx As Long
    
    Select Case ZaehlerName
        Case "Clubwagen"
            GetTargetRow = 22
        Case "Kühltruhe"
            GetTargetRow = 23
        Case "Hauptzähler"
            GetTargetRow = IIf(Medium = "Strom", 26, 29)
        Case Else
            If Left(ZaehlerName, 8) = "Parzelle" Then
                idx = Val(Mid(ZaehlerName, 10))
                GetTargetRow = IIf(Medium = "Strom", idx + 7, idx + 9)
            Else
                GetTargetRow = 0
            End If
    End Select
End Function

' Deterministische Zahlenbereinigung
Public Function CleanNumber(ByVal v As Variant) As String
    Dim s As String
    Dim decSep As String
    
    decSep = Application.International(xlDecimalSeparator)

    If Not IsNumeric(v) Then
        CleanNumber = ""
        Exit Function
    End If

    s = CStr(v)

    If InStr(s, decSep) > 0 Then
        If Val(Mid(s, InStr(s, decSep) + 1)) = 0 Then
            s = Left(s, InStr(s, decSep) - 1)
        End If
    End If

    CleanNumber = s
End Function


' ==========================================================
' 1. DATENSTRUKTUR & INITIALISIERUNG
' ==========================================================

Sub ErzeugeParzellenUebersicht()
    Dim wsQuelle As Worksheet
    Dim wsZiel As Worksheet
    Dim lastRow As Long
    Dim maxParzelle As Long
    Dim i As Long
    Dim zielZeile As Long
    Dim nameGefunden As String
    
    Set wsQuelle = ThisWorkbook.Worksheets("Mitgliederliste")
    Set wsZiel = ThisWorkbook.Worksheets("Übersicht")
    
    On Error Resume Next
    wsZiel.Unprotect
    On Error GoTo 0
    
    lastRow = wsQuelle.Cells(wsQuelle.Rows.count, 2).End(xlUp).Row
    maxParzelle = 14
    
    For i = 6 To lastRow
        If IsNumeric(wsQuelle.Cells(i, 2).value) Then
            If wsQuelle.Cells(i, 2).value > maxParzelle Then
                maxParzelle = wsQuelle.Cells(i, 2).value
            End If
        End If
    Next i
    
    With wsZiel.Range("B5:C" & wsZiel.Rows.count)
        .UnMerge
        .ClearContents
        .Locked = True
    End With

    wsZiel.Columns("C").WrapText = True
    
    zielZeile = 5
    
    For i = 1 To maxParzelle
        nameGefunden = HoleNamenFuerParzelle(wsQuelle, CStr(i), lastRow)
        
        With wsZiel.Range(wsZiel.Cells(zielZeile, 2), wsZiel.Cells(zielZeile + 7, 2))
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .value = "Parzelle " & i
            .Font.Bold = True
        End With
        
      With wsZiel.Range(wsZiel.Cells(zielZeile, 3), wsZiel.Cells(zielZeile + 7, 3))
            .Merge
            .WrapText = True
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
            .value = nameGefunden
        End With
        
        zielZeile = zielZeile + 8
    Next i
    
    nameGefunden = HoleNamenFuerParzelle(wsQuelle, "Verein", lastRow)
    
    If nameGefunden <> "" Then
        With wsZiel.Range(wsZiel.Cells(zielZeile, 2), wsZiel.Cells(zielZeile + 7, 2))
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .value = "Parzelle Verein"
            .Font.Bold = True
        End With
        
        With wsZiel.Range(wsZiel.Cells(zielZeile, 3), wsZiel.Cells(zielZeile + 7, 3))
            .Merge
            .WrapText = True
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
            .value = nameGefunden
        End With
    End If
    
    wsZiel.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Sub AktualisiereZaehlerTabellenSpalteA()

    Dim wsQuelle As Worksheet
    Dim wsStrom As Worksheet
    Dim wsWasser As Worksheet
    Dim lastRowQuelle As Long
    Dim ParzellenID As Long
    Dim zielZeileStrom As Long
    Dim zielZeileWasser As Long
    Dim nameGefunden As String
    Dim strParzelleTitel As String
    Dim GesamtText As String
    Dim LaengeParzelleTitel As Long

    Set wsQuelle = ThisWorkbook.Worksheets("Mitgliederliste")
    Set wsStrom = ThisWorkbook.Worksheets("Strom")
    Set wsWasser = ThisWorkbook.Worksheets("Wasser")

    zielZeileStrom = 8
    zielZeileWasser = 10

    lastRowQuelle = wsQuelle.Cells(wsQuelle.Rows.count, 2).End(xlUp).Row
    
    On Error Resume Next
    wsStrom.Unprotect
    wsWasser.Unprotect
    On Error GoTo 0

    For ParzellenID = 1 To 14
        
        nameGefunden = HoleNamenFuerParzelle(wsQuelle, CStr(ParzellenID), lastRowQuelle)
        
        strParzelleTitel = "Parzelle " & ParzellenID
        
        If Len(nameGefunden) > 0 Then
            GesamtText = strParzelleTitel & Chr(10) & nameGefunden
        Else
            GesamtText = strParzelleTitel
        End If
        
        LaengeParzelleTitel = Len(strParzelleTitel)

        ' A. Verarbeitung für Strom
        With wsStrom.Cells(zielZeileStrom, "A")
            .value = GesamtText
            .WrapText = True
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
            
            With .Characters(Start:=1, Length:=LaengeParzelleTitel).Font
                .Size = 11
                .Bold = True
            End With
            
            If Len(nameGefunden) > 0 Then
                With .Characters(Start:=LaengeParzelleTitel + 2, Length:=Len(nameGefunden)).Font
                    .Size = 10
                    .Bold = False
                End With
            End If
        End With

        ' B. Verarbeitung für Wasser
        With wsWasser.Cells(zielZeileWasser, "A")
            .value = GesamtText
            .WrapText = True
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
            
            With .Characters(Start:=1, Length:=LaengeParzelleTitel).Font
                .Size = 11
                .Bold = True
            End With
            
            If Len(nameGefunden) > 0 Then
                With .Characters(Start:=LaengeParzelleTitel + 2, Length:=Len(nameGefunden)).Font
                    .Size = 10
                    .Bold = False
                End With
            End If
        End With

        zielZeileStrom = zielZeileStrom + 1
        zielZeileWasser = zielZeileWasser + 1
    Next ParzellenID
    
    wsStrom.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    wsWasser.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
End Sub

' HISTORIE PRÜFUNG & ERSTELLUNG
Public Sub PruefeUndErstelleZaehlerhistorie()
    Dim ws As Worksheet
    Dim foundSheet As Boolean
    Dim lo As ListObject
    Dim wsBefore As Worksheet
    Dim wsTab6 As Worksheet
    
    Const COLUMN_COUNT As Long = 11
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HIST_SHEET_NAME)
    If Not ws Is Nothing Then foundSheet = True
    Set wsBefore = ActiveSheet
    Set wsTab6 = ThisWorkbook.Worksheets("Wasser")
    On Error GoTo 0
    
    If Not foundSheet Then
        
        If Not wsTab6 Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=wsTab6)
        Else
            Set ws = ThisWorkbook.Worksheets.Add
        End If
        
        ws.Name = HIST_SHEET_NAME
        
        ws.Range("A1:K1").value = Array( _
            "lfd. Nr. (ID)", _
            "Datum (Wechsel)", _
            "Parzelle/Zähler", _
            "Medium", _
            "Zähler-Nr. (ID) alt", _
            "Zählerstand (alt) aus der letzten Ablesung", _
            "Stand alt (Ende)", _
            "Zähler-Nr. (ID) neu", _
            "Stand neu (Start)", _
            "Verbrauch", _
            "Bemerkungen")

        Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, _
                                     Source:=ws.Range("A1:K1"), _
                                     XlListObjectHasHeaders:=xlYes)
        
        lo.Name = HIST_TABLE_NAME
        lo.TableStyle = "TableStyleMedium9"
        
        With ws
            .Columns("A").ColumnWidth = 6.5
            .Columns("B").ColumnWidth = 12.5
            .Columns("C").ColumnWidth = 14
            .Columns("D").ColumnWidth = 9
            .Columns("E").ColumnWidth = 17
            .Columns("F").ColumnWidth = 12
            .Columns("G").ColumnWidth = 12
            .Columns("H").ColumnWidth = 17
            .Columns("I").ColumnWidth = 12
            .Columns("J").ColumnWidth = 12
            .Columns("K").ColumnWidth = 40
            
            .Range("B:B").NumberFormat = "dd.mm.yyyy"
            .Range("F:G, I:J").NumberFormat = "General"
            .Range("C:K").HorizontalAlignment = xlLeft
            .Range("C:K").VerticalAlignment = xlCenter
            .Columns("A").HorizontalAlignment = xlCenter
            
            With .Range("A1:K1")
                .NumberFormat = "General"
                .Font.color = RGB(0, 0, 0)
                .Font.Bold = True
                .WrapText = True
                .ShrinkToFit = False
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Interior.color = RGB_HEADER_BG
                .Borders.LineStyle = xlContinuous
                .Borders.color = RGB(0, 0, 0)
            End With
            
            .Rows(1).AutoFit
        End With
        
    ElseIf foundSheet Then
        On Error Resume Next
        If Not wsTab6 Is Nothing Then
            If ws.Index <> wsTab6.Index + 1 Then ws.Move After:=wsTab6
        End If
        
        If ws.ProtectContents Then ws.Unprotect PASSWORD
        
        Set lo = ws.ListObjects(HIST_TABLE_NAME)
        If Not lo Is Nothing Then
            If lo.Range.Columns.count <> COLUMN_COUNT Then
                lo.Resize ws.Range("A1:K" & lo.Range.Rows.count)
            End If
            
            ws.Range("E1").value = "Zähler-Nr. (ID) alt"
            ws.Range("F1").value = "Zählerstand (alt) aus der letzten Ablesung"
            ws.Range("H1").value = "Zähler-Nr. (ID) neu"
            ws.Range("K1").value = "Bemerkungen"
            ws.Range("J1").value = "Verbrauch"
            
            With ws.Range("A1:K1")
                .NumberFormat = "General"
                .WrapText = True
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Interior.color = RGB_HEADER_BG
                .Font.color = RGB(0, 0, 0)
                .Font.Bold = True
                .ShrinkToFit = False
            End With
            
            ws.Rows(1).AutoFit
            ws.Range("F:G, I:J").NumberFormat = "General"
            ws.Range("C:K").HorizontalAlignment = xlLeft
            ws.Range("C:K").VerticalAlignment = xlCenter
            ws.Columns("A").HorizontalAlignment = xlCenter
        End If
        If ws.ProtectContents Then ws.Protect PASSWORD, AllowFormattingCells:=True
        
        On Error GoTo 0
    End If
    
    If Not wsBefore Is Nothing Then wsBefore.Activate
End Sub


' ==========================================================
' ZÄHLERWECHSEL-FORMULAR START
' ==========================================================
Public Sub Start_Zaehlerwechsel(ByVal Medium As String)
    ' HINWEIS: Hier wird das UserForm 'frm_Zaehlerwechsel' benötigt.
End Sub


' ==========================================================
' DATUM PRÜFUNG
' ==========================================================
Public Function PlausiDatum(datString As String) As Boolean
    On Error GoTo ErrHandler
    Dim d As Date
    
    Dim tempString As String
    tempString = Trim(datString)
    tempString = Replace(tempString, "/", ".")
    tempString = Replace(tempString, "-", ".")
    
    d = CDate(tempString)
    
    If Year(d) < 1900 Or Year(d) > 2100 Then
        PlausiDatum = False
    Else
        PlausiDatum = True
    End If
    Exit Function
ErrHandler:
    PlausiDatum = False
End Function


' ==========================================================
' 4. KENNZAHLEN FÜR STARTSEITE
' ==========================================================
Sub Ermittle_Kennzahlen_Mitgliederliste()

    Dim wsQuelle As Worksheet
    Dim wsStrom As Worksheet
    Dim wsWasser As Worksheet
    Dim wsStart As Worksheet
    
    Dim lastRowQuelle As Long
    Dim ParzellenID As Long
    Dim ZaehlerBelegteParzellen As Long
    Dim ZaehlerMitgliederGesamt As Long
    Dim Namenliste As String
    Dim colMitglieder As New Collection
    Dim r As Long
    Dim PVal As Variant
    Dim strMitgliedKey As String

    Set wsQuelle = ThisWorkbook.Worksheets("Mitgliederliste")
    Set wsStrom = ThisWorkbook.Worksheets("Strom")
    Set wsWasser = ThisWorkbook.Worksheets("Wasser")
    Set wsStart = ThisWorkbook.Worksheets("Startseite")
    
    On Error Resume Next
    wsStrom.Unprotect
    wsWasser.Unprotect
    wsStart.Unprotect
    On Error GoTo 0
    
    lastRowQuelle = wsQuelle.Cells(wsQuelle.Rows.count, 2).End(xlUp).Row
    
    ZaehlerBelegteParzellen = 0
    ZaehlerMitgliederGesamt = 0
    
    For ParzellenID = 1 To 14
        Namenliste = HoleNamenFuerParzelle(wsQuelle, CStr(ParzellenID), lastRowQuelle)
        If Len(Namenliste) > 0 Then
            ZaehlerBelegteParzellen = ZaehlerBelegteParzellen + 1
        End If
    Next ParzellenID
    
    For r = 6 To lastRowQuelle
        PVal = wsQuelle.Cells(r, 2).value
        
        If IsNumeric(PVal) Then
            Dim lNum As Long
            lNum = Val(PVal)
            
            If lNum >= 1 And lNum <= 14 Then
                
                strMitgliedKey = Trim(wsQuelle.Cells(r, 6).value) & " " & Trim(wsQuelle.Cells(r, 5).value)
                
                If Len(strMitgliedKey) > 1 Then
                    
                    On Error Resume Next
                    colMitglieder.Add item:=strMitgliedKey, key:=strMitgliedKey
                    
                    If Err.Number = 0 Then
                        ZaehlerMitgliederGesamt = ZaehlerMitgliederGesamt + 1
                    End If
                    
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo 0
                End If
            End If
        End If
    Next r
    
    wsStrom.Range("B4").value = ZaehlerBelegteParzellen
    wsWasser.Range("A2").value = ZaehlerBelegteParzellen
    wsStart.Range("F2").value = ZaehlerBelegteParzellen
    
    wsStart.Range("F3").value = ZaehlerMitgliederGesamt

    wsStrom.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    wsWasser.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    wsStart.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
End Sub


