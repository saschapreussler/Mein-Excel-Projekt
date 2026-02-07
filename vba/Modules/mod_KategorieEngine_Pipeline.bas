Attribute VB_Name = "mod_KategorieEngine_Pipeline"
Option Explicit

' ===============================================================
' KATEGORIEENGINE PIPELINE
' VERSION: 3.0 - 09.02.2026
' FIX: ApplyBetragsZuordnung nur bei GRÜN aufrufen
' NEU: Dynamische DropDown-Listen in Spalte H (Kategorie)
'      basierend auf Einnahme/Ausgabe
' NEU: AktualisierKategorieListen füllt Daten! AF/AG
' ===============================================================

Public Sub KategorieEngine_Pipeline(Optional ByVal wsBK As Worksheet)

    Dim wsData As Worksheet
    Dim rngRules As Range
    Dim lastRowBK As Long
    Dim r As Long

    If wsBK Is Nothing Then Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsData = ThisWorkbook.Worksheets(WS_DATEN)
    Set rngRules = wsData.Range(RANGE_KATEGORIE_REGELN)

    lastRowBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRowBK < BK_START_ROW Then Exit Sub

    ' Kategorie-Listen auf Daten! aktualisieren (für DropDowns)
    AktualisierKategorieListen

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For r = BK_START_ROW To lastRowBK

        Dim normText As String
        normText = NormalizeBankkontoZeile(wsBK, r)
        If normText = "" Then GoTo nextRow

        ' Kategorie ermitteln
        EvaluateKategorieEngineRow wsBK, r, rngRules

        ' Betrag nur zuordnen wenn Kategorie GRÜN ist
        ' GELB (Sammelzahlung) und ROT dürfen NICHT überschrieben werden!
        If wsBK.Cells(r, BK_COL_KATEGORIE).Interior.color = RGB(198, 239, 206) Then
            ApplyBetragsZuordnung wsBK, r
        End If
        
        ' DropDown für ROT und GELB setzen
        Dim katFarbe As Long
        katFarbe = wsBK.Cells(r, BK_COL_KATEGORIE).Interior.color
        If katFarbe = RGB(255, 199, 206) Or katFarbe = RGB(255, 235, 156) Then
            SetzeKategorieDropDown wsBK, r
        End If

nextRow:
    Next r

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub


' ===============================================================
' Kategorie-Listen auf Daten! AF + AG befüllen
' (Eindeutige Kategorienamen, getrennt nach E und A)
' ===============================================================
Private Sub AktualisierKategorieListen()
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets(WS_DATEN)
    
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    
    Dim dictE As Object, dictA As Object
    Set dictE = CreateObject("Scripting.Dictionary")
    Set dictA = CreateObject("Scripting.Dictionary")
    
    Dim r As Long
    Dim kat As String
    Dim ea As String
    
    For r = DATA_START_ROW To lastRow
        kat = Trim(wsData.Cells(r, DATA_CAT_COL_KATEGORIE).value)
        ea = UCase(Trim(wsData.Cells(r, DATA_CAT_COL_EINAUS).value))
        If kat = "" Then GoTo NextListRow
        
        ' Sammelzahlung NICHT in die Listen!
        If LCase(kat) Like "*sammelzahlung*" Then GoTo NextListRow
        
        If ea = "E" Then
            If Not dictE.Exists(kat) Then dictE.Add kat, True
        ElseIf ea = "A" Then
            If Not dictA.Exists(kat) Then dictA.Add kat, True
        End If
NextListRow:
    Next r
    
    ' Alte Listen löschen
    Dim clearEnd As Long
    clearEnd = Application.WorksheetFunction.Max(50, lastRow + 10)
    wsData.Range(wsData.Cells(DATA_START_ROW, DATA_COL_KAT_EINNAHMEN), _
                 wsData.Cells(clearEnd, DATA_COL_KAT_EINNAHMEN)).ClearContents
    wsData.Range(wsData.Cells(DATA_START_ROW, DATA_COL_KAT_AUSGABEN), _
                 wsData.Cells(clearEnd, DATA_COL_KAT_AUSGABEN)).ClearContents
    
    ' Einnahmen-Kategorien eintragen (Spalte AF = 32)
    Dim rowIdx As Long
    rowIdx = DATA_START_ROW
    Dim k As Variant
    For Each k In dictE.keys
        wsData.Cells(rowIdx, DATA_COL_KAT_EINNAHMEN).value = CStr(k)
        rowIdx = rowIdx + 1
    Next k
    
    ' Ausgaben-Kategorien eintragen (Spalte AG = 33)
    rowIdx = DATA_START_ROW
    For Each k In dictA.keys
        wsData.Cells(rowIdx, DATA_COL_KAT_AUSGABEN).value = CStr(k)
        rowIdx = rowIdx + 1
    Next k
    
End Sub


' ===============================================================
' DropDown-Validierung für Spalte H (Kategorie) setzen
' basierend auf Betrag-Vorzeichen (Einnahme/Ausgabe)
' ===============================================================
Private Sub SetzeKategorieDropDown(ByVal wsBK As Worksheet, ByVal rowBK As Long)
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets(WS_DATEN)
    
    Dim betrag As Double
    betrag = wsBK.Cells(rowBK, BK_COL_BETRAG).value
    
    ' Welche Liste verwenden?
    Dim listCol As Long
    If betrag >= 0 Then
        listCol = DATA_COL_KAT_EINNAHMEN   ' AF = Einnahmen
    Else
        listCol = DATA_COL_KAT_AUSGABEN    ' AG = Ausgaben
    End If
    
    ' Letzten gefüllten Wert in der Liste finden
    Dim lastListRow As Long
    lastListRow = wsData.Cells(wsData.Rows.count, listCol).End(xlUp).Row
    If lastListRow < DATA_START_ROW Then Exit Sub
    
    ' Validierungs-Formel als Bereichsreferenz
    Dim listRange As String
    listRange = "=" & WS_DATEN & "!" & _
                wsData.Cells(DATA_START_ROW, listCol).Address(True, True) & ":" & _
                wsData.Cells(lastListRow, listCol).Address(True, True)
    
    ' Alte Validierung löschen
    wsBK.Cells(rowBK, BK_COL_KATEGORIE).Validation.Delete
    
    ' Neue DropDown-Validierung setzen
    With wsBK.Cells(rowBK, BK_COL_KATEGORIE).Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:=listRange
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
        .ErrorTitle = "Kategorie"
        .ErrorMessage = "Bitte eine Kategorie aus der Liste wählen."
    End With
    
End Sub


' ===============================================================
' NEU: Re-Evaluierung nach EntityRole-Änderung
' Nur Zeilen neu bewerten die:
'   - automatisch zugeordnet wurden (nicht manuell geändert)
'   - ROT sind (keine Kategorie gefunden)
'   - GELB sind (Sammelzahlung durch fehlende EntityRole)
' Manuell korrigierte Zeilen werden NICHT angefasst!
' ===============================================================
Public Sub ReEvaluiereNachEntityRoleAenderung(ByVal geaenderteIBAN As String)

    Dim wsBK As Worksheet
    Dim wsData As Worksheet
    Dim rngRules As Range
    Dim lastRowBK As Long
    Dim r As Long
    Dim ibanClean As String
    Dim zeilenIBAN As String
    Dim kategorieFarbe As Long
    Dim anzahlNeu As Long
    
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsData = ThisWorkbook.Worksheets(WS_DATEN)
    Set rngRules = wsData.Range(RANGE_KATEGORIE_REGELN)
    
    ibanClean = UCase(Replace(geaenderteIBAN, " ", ""))
    If ibanClean = "" Then Exit Sub
    
    lastRowBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRowBK < BK_START_ROW Then Exit Sub
    
    ' Listen aktualisieren (falls neue Kategorien hinzugekommen sind)
    AktualisierKategorieListen
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error Resume Next
    wsBK.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    anzahlNeu = 0
    
    For r = BK_START_ROW To lastRowBK
        ' Nur Zeilen mit der geänderten IBAN betrachten
        zeilenIBAN = UCase(Replace(Trim(CStr(wsBK.Cells(r, BK_COL_IBAN).value)), " ", ""))
        If zeilenIBAN <> ibanClean Then GoTo NextRowReEval
        
        ' Prüfen ob die Zeile manuell geändert wurde
        kategorieFarbe = wsBK.Cells(r, BK_COL_KATEGORIE).Interior.color
        
        ' GRÜN = erfolgreich zugeordnet ? NICHT anfassen
        If kategorieFarbe = RGB(198, 239, 206) Then GoTo NextRowReEval
        
        ' Prüfen ob der Nutzer manuell etwas in den Betragsspalten eingetragen hat
        If HatManuelleBetragseingabe(wsBK, r) Then GoTo NextRowReEval
        
        ' Diese Zeile darf neu evaluiert werden!
        ' Alte Kategorie, Bemerkung und Validierung löschen
        wsBK.Cells(r, BK_COL_KATEGORIE).value = ""
        wsBK.Cells(r, BK_COL_KATEGORIE).Interior.ColorIndex = xlNone
        wsBK.Cells(r, BK_COL_KATEGORIE).Font.color = vbBlack
        wsBK.Cells(r, BK_COL_BEMERKUNG).value = ""
        On Error Resume Next
        wsBK.Cells(r, BK_COL_KATEGORIE).Validation.Delete
        On Error GoTo 0
        
        ' Neu evaluieren
        EvaluateKategorieEngineRow wsBK, r, rngRules
        
        ' Betrag nur zuordnen wenn GRÜN
        If wsBK.Cells(r, BK_COL_KATEGORIE).Interior.color = RGB(198, 239, 206) Then
            ApplyBetragsZuordnung wsBK, r
        End If
        
        ' DropDown für ROT und GELB setzen
        Dim reEvalFarbe As Long
        reEvalFarbe = wsBK.Cells(r, BK_COL_KATEGORIE).Interior.color
        If reEvalFarbe = RGB(255, 199, 206) Or reEvalFarbe = RGB(255, 235, 156) Then
            SetzeKategorieDropDown wsBK, r
        End If
        
        anzahlNeu = anzahlNeu + 1
        
NextRowReEval:
    Next r
    
    wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    If anzahlNeu > 0 Then
        Debug.Print "Re-Evaluierung: " & anzahlNeu & " Zeilen für IBAN " & Left(ibanClean, 8) & "... neu bewertet."
    End If
    
End Sub


' ===============================================================
' Prüft ob der Nutzer manuell Beträge in Spalten M-Z eingetragen hat
' ===============================================================
Private Function HatManuelleBetragseingabe(ByVal wsBK As Worksheet, _
                                            ByVal rowBK As Long) As Boolean
    Dim c As Long
    HatManuelleBetragseingabe = False
    
    For c = BK_COL_EINNAHMEN_START To BK_COL_AUSGABEN_ENDE  ' M bis Z
        If wsBK.Cells(rowBK, c).value <> "" And wsBK.Cells(rowBK, c).value <> 0 Then
            HatManuelleBetragseingabe = True
            Exit Function
        End If
    Next c
End Function

