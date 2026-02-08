Attribute VB_Name = "mod_ZaehlerLogik"
Option Explicit

' ==========================================================
' KONSTANTEN & VARIABLEN
' ==========================================================
Public IsZaehlerLogicRunning As Boolean ' Flag zur Unterdrückung von Worksheet Events
Private Const HIST_SHEET_NAME As String = "Zählerhistorie"
Private Const HIST_TABLE_NAME As String = "Tabelle_Zaehlerhistorie" ' Name des ListObjects
Private Const PASSWORD As String = "" ' Passwort (leer)
Private Const STR_HISTORY_SEPARATOR As String = "--- Zählerhistorie Makro-Eintrag ---" ' Neu: Zentralisiert

' --- SPALTEN FÜR DIE HAUPTSEITE (Tabelle5/Tabelle6) ---
Private Const COL_STAND_ANFANG As String = "B"
Private Const COL_STAND_ENDE As String = "C"
Private Const COL_VERBRAUCH_GESAMT As String = "D"
Private Const COL_BEMERKUNG As String = "E"

' --- SPALTEN FÜR DIE ZÄHLERHISTORIE (KORRIGIERT: NUR 11 SPALTEN A bis K) ---
Private Const COL_HIST_ID As Long = 1 ' A
Private Const COL_HIST_DATUM As Long = 2 ' B
Private Const COL_HIST_PARZELLE As Long = 3 ' C
Private Const COL_HIST_MEDIUM As Long = 4 ' D
Private Const COL_HIST_ZAEHLER_ALT As Long = 5 ' E
Private Const COL_HIST_STAND_ALT_ANFANG As Long = 6 ' F <-- Stand Alt Beginn (max 4 Dezimalstellen, dynamisch)
Private Const COL_HIST_STAND_ALT_ENDE As Long = 7 ' G <-- Stand Alt Ende (max 4 Dezimalstellen, dynamisch)
Private Const COL_HIST_ZAEHLER_NEU As Long = 8 ' H
Private Const COL_HIST_STAND_NEU_START As Long = 9 ' I <-- Stand Neu Start (max 4 Dezimalstellen, dynamisch)
Private Const COL_HIST_VERBRAUCH_ALT As Long = 10 ' J <-- Verbrauch Alt (JETZT MIT KOMMASTELLEN)
Private Const COL_HIST_BEMERKUNG As Long = 11 ' K (NEU: Index 11)

' --- FARBEN ---
Private Const RGB_STROM As Long = 86271 ' Hellgrün/Lime
Private Const RGB_WASSER As Long = 16737792 ' Dunkelblau
Private Const RGB_EINGABE_ERLAUBT As Long = 7592334 ' Farbe für ungesperrt (C-Spalte)
Private Const RGB_GEWECHSELT As Long = 4980735 ' Farbe für gesperrt (B-Spalte)
Private Const RGB_HEADER_BG As Long = 13619148 ' RGB(208, 208, 208) - Gewünschtes Grau

' ==========================================================
' 0. WICHTIGE HILFSFUNKTIONEN
' ==========================================================

' HILFSPROZEDUR: Zeilenhöhe auf Minimum sicherstellen
Public Sub EnsureMinRowHeight(ws As Worksheet, targetRow As Long)
    Const MIN_HEIGHT As Double = 50 ' Ihre Mindesthöhe
    On Error Resume Next
    If ws.Rows(targetRow).RowHeight < MIN_HEIGHT Then
        ws.Rows(targetRow).RowHeight = MIN_HEIGHT
    End If
    On Error GoTo 0
End Sub

' HILFSFUNKTION: Hole Namen für Parzelle
Function HoleNamenFuerParzelle(ws As Worksheet, suchParzelle As String, maxZ As Long) As String
    Dim r As Long
    Dim sTemp As String
    Dim PVal As String
    
    sTemp = ""
    
    For r = 6 To maxZ
        PVal = Trim(CStr(ws.Cells(r, 2).value))
        
        If StrComp(PVal, suchParzelle, vbTextCompare) = 0 Then
            ' Vorname (F) + Nachname (E)
            If sTemp <> "" Then sTemp = sTemp & vbLf
            sTemp = sTemp & Trim(ws.Cells(r, 6).value) & " " & Trim(ws.Cells(r, 5).value)
        End If
    Next r
    
    HoleNamenFuerParzelle = sTemp
End Function

' HILFSFUNKTION: Ermittlung der Zielzeile
Private Function GetTargetRow(ByVal ZaehlerName As String, ByVal Medium As String) As Long
    Dim idx As Long
    
    Select Case ZaehlerName
        Case "Clubwagen"
            GetTargetRow = 22
        Case "Kühltruhe"
            GetTargetRow = 23
        Case "Hauptzähler"
            ' Achtung: Hier müssen Sie die tatsächlichen Zeilen in Ihrem Sheet prüfen!
            GetTargetRow = IIf(Medium = "Strom", 26, 29)
        Case Else
            If Left(ZaehlerName, 8) = "Parzelle" Then
                ' Parzellen 1-14
                idx = Val(Mid(ZaehlerName, 10))
                ' Zeilen: Strom (Parzelle 1) = Zeile 8. Wasser (Parzelle 1) = Zeile 10.
                GetTargetRow = IIf(Medium = "Strom", idx + 7, idx + 9)
            Else
                GetTargetRow = 0
            End If
    End Select
End Function

' NEU: DETERMINISTISCHE ZAHLENBEREINIGUNG (UDF)
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
        ' Prüft, ob der Dezimalteil nur Nullen enthält
        If Val(Mid(s, InStr(s, decSep) + 1)) = 0 Then
            s = Left(s, InStr(s, decSep) - 1)
        End If
    End If

    CleanNumber = s
End Function

' ==========================================================
' 1. DATENSTRUKTUR & INITIALISIERUNG
' ==========================================================

' ... (ErzeugeParzellenUebersicht - Unverändert, verwendet Sheet-Namen) ...
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

' ... (AktualisiereZaehlerTabellenSpalteA - MUSS Sheet-Namen verwenden) ...
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

    ' HINWEIS: Hier müssen die CodeNamen (Tabelle5/6) durch die Sheet-Namen ersetzt werden,
    ' ODER die CodeNamen müssen in den Deklarationen oben verwendet werden.
    Set wsQuelle = ThisWorkbook.Worksheets("Mitgliederliste")
    Set wsStrom = ThisWorkbook.Worksheets("Strom") ' <- ANPASSEN
    Set wsWasser = ThisWorkbook.Worksheets("Wasser") ' <- ANPASSEN

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
    Dim wsTab6 As Worksheet ' CodeName Referenz
    
    Const COLUMN_COUNT As Long = 11
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HIST_SHEET_NAME)
    If Not ws Is Nothing Then foundSheet = True
    Set wsBefore = ActiveSheet
    ' ANPASSEN: CodeName Tabelle6 durch Sheet-Namen ersetzen, wenn möglich
    Set wsTab6 = ThisWorkbook.Worksheets("Wasser") ' <- ANPASSEN
    On Error GoTo 0
    
    If Not foundSheet Then
        
        If Not wsTab6 Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=wsTab6)
        Else
            Set ws = ThisWorkbook.Worksheets.Add
        End If
        
        ws.name = HIST_SHEET_NAME
        
        ' Header-Werte setzen
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

        ' Excel-Tabelle (ListObject) erstellen
        Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, _
                                     Source:=ws.Range("A1:K1"), _
                                     XlListObjectHasHeaders:=xlYes)
        
        lo.name = HIST_TABLE_NAME
        lo.TableStyle = "TableStyleMedium9"
        
        ' Formatierung der Tabelle
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
        
    ' Wenn Blatt existiert: Nur ListObject-Größe anpassen (falls notwendig)
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
' 2. ZÄHLER-BERECHNUNG (HAUPTPROZEDUR)
' ==========================================================

' Funktion: Ruft Einzelberechnung für alle Zähler einer Seite auf
Public Sub CalculateAllZaehlerVerbrauch(wsTarget As Worksheet)
    
    ' WICHTIG: Alle Variablen hier nur einmal deklarieren
    Dim wsHist As Worksheet
    Dim r As Long ' Wird für die For-Schleife der Parzellen verwendet
    Dim wasProtected As Boolean
    
    If wsTarget Is Nothing Then Exit Sub
    
    ' --- Blattschutz aufheben (zum Formatieren/Schreiben) ---
    wasProtected = wsTarget.ProtectContents
    If wasProtected Then
        On Error Resume Next
        wsTarget.Unprotect PASSWORD
        On Error GoTo 0 ' Deaktiviert Resume Next für regulären Code
    End If
    
    On Error GoTo Fehler_Handler_Berechnung
    
    ' --- Formatierung ---
    wsTarget.Range("8:23").RowHeight = 50

    With wsTarget.Range("B8:D23, F8:I23")
        .ShrinkToFit = True
        .WrapText = False
    End With
    
    With wsTarget.Range("A8:A23")
        .ShrinkToFit = False
        .WrapText = True
    End With

    With wsTarget.Range(COL_BEMERKUNG & "8:" & COL_BEMERKUNG & "23")
        .ShrinkToFit = False
        .WrapText = True
    End With
    
    ' --- Historie laden ---
    On Error Resume Next
    Set wsHist = ThisWorkbook.Worksheets(HIST_SHEET_NAME)
    On Error GoTo Fehler_Handler_Berechnung

    ' ==========================================================
    ' 1. PARZELLENZÄHLER & UNTERZÄHLER
    ' ==========================================================
    
    If LCase(wsTarget.name) = "strom" Then
        ' STROM: Parzelle 1 bis 12 (Zeilen 8 bis 19)
        For r = 1 To 12
            Call CalculateSingleZaehler(wsTarget, wsHist, "Strom", "Parzelle " & r, r + 7)
            wsTarget.Rows(r + 7).AutoFit
            Call EnsureMinRowHeight(wsTarget, r + 7)
        Next r
        
        ' STROM: Feste Zähler (Clubwagen, Kühltruhe, P13/P14 Unterzähler)
        Call CalculateSingleZaehler(wsTarget, wsHist, "Strom", "Clubwagen", 22)
        wsTarget.Rows(22).AutoFit
        Call EnsureMinRowHeight(wsTarget, 22)
        
        Call CalculateSingleZaehler(wsTarget, wsHist, "Strom", "Kühltruhe", 23)
        wsTarget.Rows(23).AutoFit
        Call EnsureMinRowHeight(wsTarget, 23)
        
        Call CalculateSingleZaehler(wsTarget, wsHist, "Strom", "Parzelle 13", 20) ' P13 Unterzähler
        wsTarget.Rows(20).AutoFit
        Call EnsureMinRowHeight(wsTarget, 20)
        
        Call CalculateSingleZaehler(wsTarget, wsHist, "Strom", "Parzelle 14", 21) ' P14 Unterzähler
        wsTarget.Rows(21).AutoFit
        Call EnsureMinRowHeight(wsTarget, 21)
        
    ElseIf LCase(wsTarget.name) = "wasser" Then
        ' WASSER: Parzelle 1 bis 14 (Zeilen 10 bis 23)
        For r = 1 To 14
            ' Hier wurde der Startindex (10 statt 8) verwendet, was zu Ihrer Tabellenstruktur passen muss.
            Call CalculateSingleZaehler(wsTarget, wsHist, "Wasser", "Parzelle " & r, r + 9)
            wsTarget.Rows(r + 9).AutoFit
            Call EnsureMinRowHeight(wsTarget, r + 9)
        Next r
    End If
    
    ' ******************************************************
    ' LOGIK FÜR A22 (FORMATIERTER TEXT) - DIESE LOGIK FEHLT NOCH HIER
    ' ******************************************************

    ' ==========================================================
    ' 2. HAUPTZÄHLER
    ' ==========================================================
    
    If LCase(wsTarget.name) = "strom" Then
        ' Hauptzähler Strom (Zeile 26)
        Call CalculateSingleZaehler(wsTarget, wsHist, "Strom", "Hauptzähler", 26)
        wsTarget.Rows(26).AutoFit
        Call EnsureMinRowHeight(wsTarget, 26)
    End If
    
    If LCase(wsTarget.name) = "wasser" Then
        ' Hauptzähler Wasser (Zeile 29)
        Call CalculateSingleZaehler(wsTarget, wsHist, "Wasser", "Hauptzähler", 29)
        wsTarget.Rows(29).AutoFit
        Call EnsureMinRowHeight(wsTarget, 29)
    End If
    
    ' 3. SUMMENZEILEN (Wenn relevant)
    ' ... (Hier könnte die Logik für Summenberechnungen folgen, falls vorhanden) ...

Cleanup_Berechnung:
    ' BLATTSCHUTZ WIEDERHERSTELLEN
    If wasProtected Then
        wsTarget.Protect PASSWORD, AllowFormattingCells:=True
    End If
    
    Exit Sub

Fehler_Handler_Berechnung:
    ' Im Fehlerfall (bevor wir zum Cleanup springen):
    MsgBox "Ein schwerwiegender Fehler ist während der Zählerberechnung aufgetreten. " & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description, vbCritical, "Fehler in CalculateAllZaehlerVerbrauch"

    ' Jetzt springen wir zum Cleanup, um den Blattschutz wiederherzustellen und die Prozedur zu verlassen.
    Resume Cleanup_Berechnung

End Sub

' Einzelberechnung (Kernlogik)
Private Sub CalculateSingleZaehler( _
    wsTarget As Worksheet, _
    wsHist As Worksheet, _
    ByVal ZaehlerTyp As String, _
    ByVal ZaehlerName As String, _
    ByVal targetRow As Long)

    Dim startCell As Range, endCell As Range
    Dim targetCellD As Range
    Dim targetBemerkung As Range
    
    Dim standAnfangCurrent As Double
    Dim standEndeCurrent As Double
    Dim VerbrauchGesamt As Double
    Dim verbrauchAltHistorie_Summe As Double
    Dim verbrauchNeuAktuell As Double
    Dim einheit As String
    Dim f As Range, firstAddr As String
    Dim currentRow As Long
    Dim zyklen As Long
    Dim lastDate As Date
    Dim snNeu_last As String
    Dim standNeuStart_last As Double
    
    Set startCell = wsTarget.Cells(targetRow, COL_STAND_ANFANG)
    Set endCell = wsTarget.Cells(targetRow, COL_STAND_ENDE)
    Set targetCellD = wsTarget.Cells(targetRow, COL_VERBRAUCH_GESAMT)
    Set targetBemerkung = wsTarget.Cells(targetRow, COL_BEMERKUNG)

    einheit = IIf(LCase(ZaehlerTyp) = "strom", "kWh", "m³")
    
    ' 0. Startwerte lesen
    If IsNumeric(startCell.value) And Not isEmpty(startCell.value) Then
        standAnfangCurrent = CDbl(startCell.value)
    Else
        standAnfangCurrent = 0
    End If
    
    If IsNumeric(endCell.value) And Not isEmpty(endCell.value) Then
        standEndeCurrent = CDbl(endCell.value)
    Else
        standEndeCurrent = 0
    End If
    
    ' 1. Vorabprüfung (Fehler)
    If standEndeCurrent < standAnfangCurrent Then
        targetBemerkung.value = "FEHLER: Endstand (" & Format(standEndeCurrent, "#,##0.00") & ") < Startstand (" & Format(standAnfangCurrent, "#,##0.00") & ")."
        targetCellD.ClearContents
        
        ' Freigabe beider Felder zur Korrektur
        startCell.Interior.color = RGB_EINGABE_ERLAUBT
        startCell.Locked = False
        endCell.Interior.color = RGB_EINGABE_ERLAUBT
        endCell.Locked = False
        
        With targetBemerkung
            .ShrinkToFit = False
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        
        Exit Sub
    End If
    
    
    ' 2. HISTORIE DURCHSUCHEN: SUMMIERE ALLE WECHSEL
    verbrauchAltHistorie_Summe = 0
    zyklen = 0
    lastDate = 0
    standNeuStart_last = 0
    snNeu_last = ""
    
    If Not wsHist Is Nothing Then
        Set f = wsHist.Columns(COL_HIST_PARZELLE).Find( _
            What:=ZaehlerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

        If Not f Is Nothing Then
            firstAddr = f.Address
            Do
                currentRow = f.Row
                If StrComp(Trim(CStr(wsHist.Cells(currentRow, COL_HIST_MEDIUM).value)), ZaehlerTyp, vbTextCompare) = 0 Then
                    
                    zyklen = zyklen + 1
                    If IsNumeric(wsHist.Cells(currentRow, COL_HIST_VERBRAUCH_ALT).value) Then
                        verbrauchAltHistorie_Summe = verbrauchAltHistorie_Summe + CDbl(wsHist.Cells(currentRow, COL_HIST_VERBRAUCH_ALT).value)
                    End If
                    
                    ' Finde den neuesten Eintrag
                    If IsDate(wsHist.Cells(currentRow, COL_HIST_DATUM).value) Then
                        If CDate(wsHist.Cells(currentRow, COL_HIST_DATUM).value) >= lastDate Then
                            lastDate = CDate(wsHist.Cells(currentRow, COL_HIST_DATUM).value)
                            snNeu_last = CStr(wsHist.Cells(currentRow, COL_HIST_ZAEHLER_NEU).value)
                            If IsNumeric(wsHist.Cells(currentRow, COL_HIST_STAND_NEU_START).value) Then
                                standNeuStart_last = CDbl(wsHist.Cells(currentRow, COL_HIST_STAND_NEU_START).value)
                            End If
                        End If
                    End If
                End If
                Set f = wsHist.Columns(COL_HIST_PARZELLE).FindNext(f)
            Loop While Not f Is Nothing And f.Address <> firstAddr
        End If
    End If


    ' 3. BERECHNUNG UND SCHREIBEN IN D, E
    If zyklen > 0 Then ' FALL A: Mindestens ein Zählerwechsel gefunden (B muss gesperrt werden)
        
        If standAnfangCurrent <> standNeuStart_last Then
            startCell.value = CleanNumber(standNeuStart_last)
            standAnfangCurrent = standNeuStart_last
        End If
        
        verbrauchNeuAktuell = Round(CDec(standEndeCurrent) - CDec(standAnfangCurrent), 2)
        VerbrauchGesamt = CDec(verbrauchAltHistorie_Summe) + CDec(verbrauchNeuAktuell)
        
        ' Spalte D: Gesamtverbrauch
        If targetRow = 22 And ZaehlerName = "Clubwagen" Then
              targetCellD.value = Round(VerbrauchGesamt, 0)
              targetCellD.NumberFormat = "0;[Red]-0;;"
        Else
              targetCellD.value = VerbrauchGesamt
              targetCellD.NumberFormat = "#,##0.00;[Red]-#,##0.00;;"
        End If
        
        ' ***************************************************************
        ' LOGIK FÜR SPALTE E (BEMERKUNG BEI ZÄHLERWECHSEL)
        ' ***************************************************************
        Dim oldBemerkung As String
        Dim newHistoryText As String
        Dim posSeparator As Long
        
        ' 1. Neuen, vereinfachten Historien-Block erstellen
        newHistoryText = "Letzter Zählerwechsel am: " & Format(lastDate, "dd.mm.yyyy") & vbLf & _
                             "Anzahl der Wechsel: " & zyklen & vbLf & _
                             "Gesamtverbrauch gewechselte Zähler: " & Format(verbrauchAltHistorie_Summe, "#,##0.00") & " " & einheit & vbLf & _
                             "Verbrauch derzeitiger Zähler: " & Format(verbrauchNeuAktuell, "#,##0.00") & " " & einheit
        
        oldBemerkung = Trim(CStr(targetBemerkung.value))
        
        ' 2. Prüfen, ob Makro-Eintrag bereits existiert
        posSeparator = InStr(1, oldBemerkung, STR_HISTORY_SEPARATOR, vbTextCompare)
        
        If posSeparator > 0 Then
            Dim userText As String
            
            userText = Trim(Left(oldBemerkung, posSeparator - 1))
            
            If Len(userText) > 0 Then
                targetBemerkung.value = userText & vbLf & STR_HISTORY_SEPARATOR & vbLf & newHistoryText
            Else
                targetBemerkung.value = STR_HISTORY_SEPARATOR & vbLf & newHistoryText
            End If
            
        Else ' Eintrag existiert noch nicht
            
            If Len(oldBemerkung) > 0 Then
                targetBemerkung.value = oldBemerkung & vbLf & STR_HISTORY_SEPARATOR & vbLf & newHistoryText
            Else
                targetBemerkung.value = STR_HISTORY_SEPARATOR & vbLf & newHistoryText
            End If
        End If
        
        ' 3. Formatierung von E sicherstellen
        With targetBemerkung
            .ShrinkToFit = False
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        
        ' Sicherstellung der Formatierung/Sperrung
        startCell.Interior.color = RGB_GEWECHSELT
        startCell.Locked = True
        endCell.Interior.color = RGB_EINGABE_ERLAUBT
        endCell.Locked = False
        
    Else ' FALL B: Kein Wechsel gefunden (Standardfall)
        
        verbrauchNeuAktuell = Round(CDec(standEndeCurrent) - CDec(standAnfangCurrent), 2)
        VerbrauchGesamt = verbrauchNeuAktuell
        
        ' Spalte D: Gesamtverbrauch
        targetCellD.value = VerbrauchGesamt
        targetCellD.NumberFormat = "#,##0.00;[Red]-#,##0.00;;"
        
        ' 1. Formatierung von E sicherstellen
        With targetBemerkung
            .ShrinkToFit = False
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        
        ' Formatierung und Sperrung für den Standardfall (beide sind änderbar)
        With startCell
            .Locked = False
            .Interior.color = RGB_EINGABE_ERLAUBT
        End With
        
        With endCell
            .Locked = False
            .Interior.color = RGB_EINGABE_ERLAUBT
        End With
        
    End If
    
End Sub

' ==========================================================
' 3. HISTORIE & HELFER
' ==========================================================

' Historie schreiben
Public Sub SchreibeHistorie( _
    ByVal parzelle As String, _
    ByVal DatumW As Date, _
    ByVal AltEnde As Double, _
    ByVal neuStart As Double, _
    ByVal snNeu As String, _
    ByVal snAlt As String, _
    Optional ByVal bem As String = "", _
    Optional ByVal Medium As String)
    
    Dim ws As Worksheet, lo As ListObject, newRow As ListRow
    Dim lngColor As Long
    Dim wsTarget As Worksheet
    Dim targetRow As Long
    Dim standAnfangAlt As Double, verbrauchAltHistorie As Double
    Dim wasTargetProtected As Boolean, wasHistoryProtected As Boolean
    Dim AltEnde_Geprueft As Double, neuStart_Geprueft As Double
    
    ' I. WERTE PRÜFEN & RUNDEN
    If AltEnde = Int(AltEnde) Then AltEnde_Geprueft = AltEnde Else AltEnde_Geprueft = Round(AltEnde, 4)
    If neuStart = Int(neuStart) Then neuStart_Geprueft = neuStart Else neuStart_Geprueft = Round(neuStart, 4)
    
    ' Zuerst Zielblatt und Schutzstatus ermitteln
    ' ANPASSUNG DER LOGIK: Verwendung von Select Case für saubere Zuweisung
    Select Case Medium
        Case "Strom"
            Set wsTarget = ThisWorkbook.Worksheets("Strom") ' <- ANPASSEN
        Case "Wasser"
            Set wsTarget = ThisWorkbook.Worksheets("Wasser") ' <- ANPASSEN
        Case Else
            Exit Sub
    End Select
    
    targetRow = GetTargetRow(parzelle, Medium)
    
    On Error GoTo Fehler_Handler
    
    If IsZaehlerLogicRunning Then Err.Raise 9998, "SchreibeHistorie", "Logik ist bereits aktiv (Rekursion). Vorgang abgebrochen."
    IsZaehlerLogicRunning = True
    Application.EnableEvents = False
    
    ' BLATTSCHUTZ AUFHEBEN (Zielblatt)
    If Not wsTarget Is Nothing Then
        wasTargetProtected = wsTarget.ProtectContents
        If wasTargetProtected Then
            On Error Resume Next
            wsTarget.Unprotect PASSWORD
            On Error GoTo Fehler_Handler
        End If
    End If

    ' 1. SICHERSTELLEN, DASS DAS BLATT/LISTOBJECT EXISTIERT
    Call PruefeUndErstelleZaehlerhistorie
    
    Set ws = ThisWorkbook.Worksheets(HIST_SHEET_NAME)
    Set lo = ws.ListObjects(HIST_TABLE_NAME)
    
    If lo Is Nothing Then Err.Raise 9999, "mod_ZaehlerLogik.SchreibeHistorie", "ListObject wurde nicht gefunden/erstellt."
    
    ' Historienblatt für das Schreiben entsperren
    wasHistoryProtected = ws.ProtectContents
    If wasHistoryProtected Then
        On Error Resume Next
        ws.Unprotect PASSWORD
        On Error GoTo Fehler_Handler
    End If
    
    ' 2. Lese den Startstand des ALTEN Zählers
    If targetRow > 0 And Not wsTarget Is Nothing Then
        If IsNumeric(wsTarget.Cells(targetRow, COL_STAND_ANFANG).value) Then
            standAnfangAlt = CDbl(wsTarget.Cells(targetRow, COL_STAND_ANFANG).value)
            If standAnfangAlt <> Int(standAnfangAlt) Then standAnfangAlt = Round(standAnfangAlt, 4)
        Else
            standAnfangAlt = 0
        End If
    Else
        standAnfangAlt = 0
    End If
    
    verbrauchAltHistorie = Round(CDec(AltEnde_Geprueft) - CDec(standAnfangAlt), 4)
    
    ' 3. Daten in Historie speichern
    Set newRow = lo.ListRows.Add(AlwaysInsert:=True)
    
    With newRow.Range
        .Cells(1, COL_HIST_ID).value = lo.ListRows.count
        .Cells(1, COL_HIST_DATUM).value = DatumW
        .Cells(1, COL_HIST_PARZELLE).value = parzelle
        .Cells(1, COL_HIST_MEDIUM).value = Medium
        
        .Cells(1, COL_HIST_ZAEHLER_ALT).value = snAlt
        .Cells(1, COL_HIST_STAND_ALT_ANFANG).value = CleanNumber(standAnfangAlt)
        .Cells(1, COL_HIST_STAND_ALT_ENDE).value = CleanNumber(AltEnde_Geprueft)
        .Cells(1, COL_HIST_ZAEHLER_NEU).value = snNeu
        .Cells(1, COL_HIST_STAND_NEU_START).value = CleanNumber(neuStart_Geprueft)
        .Cells(1, COL_HIST_VERBRAUCH_ALT).value = CleanNumber(verbrauchAltHistorie)
        .Cells(1, COL_HIST_BEMERKUNG).value = bem
    End With
    
    ' 4. ZIELBLATT-UPDATE (Spalten B, C)
    If targetRow > 0 And Not wsTarget Is Nothing Then
        
        wsTarget.Cells(targetRow, COL_STAND_ANFANG).value = CleanNumber(neuStart_Geprueft)
        wsTarget.Cells(targetRow, COL_STAND_ENDE).value = CleanNumber(neuStart_Geprueft)
        
        With wsTarget.Cells(targetRow, COL_STAND_ANFANG)
            .Interior.color = RGB_GEWECHSELT
            .Locked = True
        End With
        
        With wsTarget.Cells(targetRow, COL_STAND_ENDE)
            .Interior.color = RGB_EINGABE_ERLAUBT
            .Locked = False
        End With
        
    End If
    
    ' 5. Farben für Historie setzen & Update-Call
    ' KORREKTUR: Umwandlung der fehlerhaften einzeiligen If-ElseIf-Struktur in Select Case.
    Select Case Medium
        Case "Strom"
            lngColor = RGB_STROM
        Case "Wasser"
            lngColor = RGB_WASSER
        Case Else
            lngColor = xlNone
    End Select
    
    If lngColor <> xlNone Then newRow.Range.Interior.color = lngColor
    
    If Not wsTarget Is Nothing Then Call CalculateAllZaehlerVerbrauch(wsTarget)
    Call FarbeHistorieEintraege
    
CleanUp:
    IsZaehlerLogicRunning = False
    Application.EnableEvents = True
    
    If Not wsTarget Is Nothing Then
        If wasTargetProtected Then wsTarget.Protect PASSWORD, AllowFormattingCells:=True
    End If
    If Not ws Is Nothing Then
        If wasHistoryProtected Then ws.Protect PASSWORD, AllowFormattingCells:=True
    End If
    
    Exit Sub

Fehler_Handler:
    Dim errNum As Long
    Dim errDesc As String
    If Err.Number <> 0 Then
        errNum = Err.Number
        errDesc = Err.Description
    Else
        errNum = 9997
        errDesc = "Unbekannter Fehler im Fehler-Handler"
    End If

    Resume CleanUp

    Err.Raise errNum, "mod_ZaehlerLogik.SchreibeHistorie", errDesc
End Sub
' FUNKTION: Färbt die Historien-Einträge
Public Sub FarbeHistorieEintraege()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lr As ListRow
    Dim lngColor As Long
    Dim wasProtected As Boolean
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HIST_SHEET_NAME)
    If ws Is Nothing Then Exit Sub
    Set lo = ws.ListObjects(HIST_TABLE_NAME)
    If lo Is Nothing Then Exit Sub
    
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect PASSWORD
    On Error GoTo 0
    
    For Each lr In lo.ListRows
        If StrComp(Trim(CStr(lr.Range.Cells(1, COL_HIST_MEDIUM).value)), "Strom", vbTextCompare) = 0 Then
            lngColor = RGB_STROM
        ElseIf StrComp(Trim(CStr(lr.Range.Cells(1, COL_HIST_MEDIUM).value)), "Wasser", vbTextCompare) = 0 Then
            lngColor = RGB_WASSER
        Else
            lngColor = xlNone
        End If
        
        lr.Range.Interior.color = lngColor
    Next lr
    
    lo.Range.Borders.LineStyle = xlContinuous
    lo.Range.Borders.color = RGB(0, 0, 0)
    
    If wasProtected Then ws.Protect PASSWORD, AllowFormattingCells:=True
End Sub

' ZÄHLERWECHSEL-FORMULAR START (Unverändert - benötigt UserForm)
Public Sub Start_Zaehlerwechsel(ByVal Medium As String)
    ' HINWEIS: Hier wird das UserForm 'frm_Zaehlerwechsel' benötigt.
    ' Dim FormInstanz As frm_Zaehlerwechsel
    ' ... (Code) ...
End Sub

' DATUM PRÜFUNG (Unverändert)
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
    ' ANPASSEN: CodeNamen durch Sheet-Namen ersetzen
    Set wsStrom = ThisWorkbook.Worksheets("Strom")
    Set wsWasser = ThisWorkbook.Worksheets("Wasser")
    Set wsStart = ThisWorkbook.Worksheets("Startseite") ' <- ANPASSEN
    
    On Error Resume Next
    wsStrom.Unprotect
    wsWasser.Unprotect
    wsStart.Unprotect
    On Error GoTo 0
    
    lastRowQuelle = wsQuelle.Cells(wsQuelle.Rows.count, 2).End(xlUp).Row
    
    ZaehlerBelegteParzellen = 0
    ZaehlerMitgliederGesamt = 0
    
    ' 1. ZÄHLUNG DER BELEGTEN PARZELLEN
    For ParzellenID = 1 To 14
        Namenliste = HoleNamenFuerParzelle(wsQuelle, CStr(ParzellenID), lastRowQuelle)
        If Len(Namenliste) > 0 Then
            ZaehlerBelegteParzellen = ZaehlerBelegteParzellen + 1
        End If
    Next ParzellenID
    
    ' 2. ZÄHLUNG DER EINZELNEN MITGLIEDER
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
    
    ' 3. ERGEBNISSE IN DIE ZIELZELLEN SCHREIBEN
    wsStrom.Range("B4").value = ZaehlerBelegteParzellen
    wsWasser.Range("A2").value = ZaehlerBelegteParzellen
    wsStart.Range("F2").value = ZaehlerBelegteParzellen
    
    wsStart.Range("F3").value = ZaehlerMitgliederGesamt

    wsStrom.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    wsWasser.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    wsStart.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
End Sub

