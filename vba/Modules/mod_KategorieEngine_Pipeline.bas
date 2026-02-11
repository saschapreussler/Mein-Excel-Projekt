Attribute VB_Name = "mod_KategorieEngine_Pipeline"
Option Explicit

' ===============================================================
' KATEGORIEENGINE PIPELINE
' VERSION: 5.1 - 08.02.2026
' MERGE: v4.0 Logik (funktionierend) +
'        v4.3 Infrastruktur (kein Named Range, Cache,
'        HatManuelleKategorie, ReEvaluiereAlleNichtManuellen)
' FIX: Evaluator-Signatur v9.0 (wsData + lastRuleRow)
' FIX: Einstellungen-Cache Lade/Entlade um Pipeline
' FIX: HatManuelleKategorie schützt manuelle Eingaben
' v5.1: Sammelzahlung-Filter in AktualisierKategorieListen
'       entfernt - alle Kategorien aus Spalte J stehen im
'       DropDown zur Verfügung (inkl. Sammelzahlung)
' ===============================================================

' ---------------------------------------------------------------
' Hauptpipeline: Evaluiert NUR Zeilen ohne manuelle Eingabe
' Wird nach CSV-Import aufgerufen.
' ---------------------------------------------------------------
Public Sub KategorieEngine_Pipeline(Optional ByVal wsBK As Worksheet)

    ' WICHTIG: Vererbtes "On Error Resume Next" vom Aufrufer
    ' (mod_Banking_Data) SOFORT deaktivieren!
    On Error GoTo 0

    Dim wsData As Worksheet
    Dim lastRuleRow As Long
    Dim lastRowBK As Long
    Dim r As Long

    If wsBK Is Nothing Then Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsData = ThisWorkbook.Worksheets(WS_DATEN)
    
    ' Zeilengrenzen dynamisch ermitteln statt Named Range
    lastRuleRow = wsData.Cells(wsData.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRuleRow < DATA_START_ROW Then
        Debug.Print "Pipeline ABBRUCH: Keine Kategorien in Daten!J ab Zeile " & DATA_START_ROW & " gefunden!"
        Exit Sub
    End If

    lastRowBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRowBK < BK_START_ROW Then Exit Sub

    ' Kategorie-Listen auf Daten! aktualisieren (für DropDowns)
    AktualisierKategorieListen
    
    ' Einstellungen-Cache laden (Performance)
    LadeEinstellungenCache

    ' Blattschutz SELBST aufheben - nicht vom Aufrufer abhängig!
    On Error Resume Next
    wsBK.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For r = BK_START_ROW To lastRowBK

        Dim normText As String
        normText = NormalizeBankkontoZeile(wsBK, r)
        If normText = "" Then GoTo nextRow
        
        ' Manuelle Betragseingabe? NICHT anfassen
        If HatManuelleBetragseingabe(wsBK, r) Then GoTo nextRow
        
        ' Manuelle Kategorie? NICHT anfassen
        If HatManuelleKategorie(wsBK, r) Then GoTo nextRow

        ' Kategorie ermitteln (v9.0: wsData + lastRuleRow)
        On Error Resume Next
        EvaluateKategorieEngineRow wsBK, r, wsData, lastRuleRow
        If Err.Number <> 0 Then
            Debug.Print "Evaluator Fehler Zeile " & r & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0

        ' Betrag nur zuordnen wenn Kategorie GRÜN ist
        On Error Resume Next
        If wsBK.Cells(r, BK_COL_KATEGORIE).Interior.color = RGB(198, 239, 206) Then
            ApplyBetragsZuordnung wsBK, r
        End If
        If Err.Number <> 0 Then
            Debug.Print "Betragszuordnung Fehler Zeile " & r & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0

        ' DropDown für ROT und GELB setzen
        Dim katFarbe As Long
        katFarbe = wsBK.Cells(r, BK_COL_KATEGORIE).Interior.color
        If katFarbe = RGB(255, 199, 206) Or katFarbe = RGB(255, 235, 156) Then
            SetzeKategorieDropDown wsBK, r
        End If

nextRow:
    Next r

    ' Einstellungen-Cache freigeben
    EntladeEinstellungenCache

    ' Blattschutz wieder aktivieren
    On Error Resume Next
    wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub


' ---------------------------------------------------------------
' Re-Evaluierung ALLER nicht-manuellen Zeilen
' Wird aufgerufen wenn Kategorie-Regeln oder Einstellungen
' geändert werden. Überspringt manuelle Eingaben.
' ---------------------------------------------------------------
Public Sub ReEvaluiereAlleNichtManuellen()

    On Error GoTo 0

    Dim wsBK As Worksheet
    Dim wsData As Worksheet
    Dim lastRuleRow As Long
    Dim lastRowBK As Long
    Dim r As Long
    Dim anzahlNeu As Long

    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsData = ThisWorkbook.Worksheets(WS_DATEN)

    lastRuleRow = wsData.Cells(wsData.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRuleRow < DATA_START_ROW Then Exit Sub

    lastRowBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRowBK < BK_START_ROW Then Exit Sub

    ' Listen aktualisieren
    AktualisierKategorieListen

    ' Einstellungen-Cache laden
    LadeEinstellungenCache

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error Resume Next
    wsBK.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0

    anzahlNeu = 0

    For r = BK_START_ROW To lastRowBK

        ' Manuelle Betragseingabe? NICHT anfassen
        If HatManuelleBetragseingabe(wsBK, r) Then GoTo NextRowReEvalAll
        
        ' Manuelle Kategorie? NICHT anfassen
        If HatManuelleKategorie(wsBK, r) Then GoTo NextRowReEvalAll
        
        ' Alte Kategorie, Bemerkung und Validierung löschen
        wsBK.Cells(r, BK_COL_KATEGORIE).value = ""
        wsBK.Cells(r, BK_COL_KATEGORIE).Interior.ColorIndex = xlNone
        wsBK.Cells(r, BK_COL_KATEGORIE).Font.color = vbBlack
        wsBK.Cells(r, BK_COL_BEMERKUNG).value = ""
        On Error Resume Next
        wsBK.Cells(r, BK_COL_KATEGORIE).Validation.Delete
        On Error GoTo 0

        ' Neu evaluieren
        On Error Resume Next
        EvaluateKategorieEngineRow wsBK, r, wsData, lastRuleRow
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0

        ' Betrag nur zuordnen wenn GRÜN
        On Error Resume Next
        If wsBK.Cells(r, BK_COL_KATEGORIE).Interior.color = RGB(198, 239, 206) Then
            ApplyBetragsZuordnung wsBK, r
        End If
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0

        ' DropDown für ROT und GELB setzen
        Dim reEvalFarbeAll As Long
        reEvalFarbeAll = wsBK.Cells(r, BK_COL_KATEGORIE).Interior.color
        If reEvalFarbeAll = RGB(255, 199, 206) Or reEvalFarbeAll = RGB(255, 235, 156) Then
            SetzeKategorieDropDown wsBK, r
        End If

        anzahlNeu = anzahlNeu + 1

NextRowReEvalAll:
    Next r

    ' Einstellungen-Cache freigeben
    EntladeEinstellungenCache

    On Error Resume Next
    wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If anzahlNeu > 0 Then
        Debug.Print "Re-Evaluierung ALLE: " & anzahlNeu & " Zeilen neu bewertet."
    End If

End Sub


' ===============================================================
' Kategorie-Listen auf Daten! AF + AG befüllen
' (Eindeutige Kategorienamen, getrennt nach E und A)
' v5.1: Sammelzahlung wird NICHT mehr ausgefiltert!
'       Alle Kategorien aus Spalte J stehen im DropDown
'       zur Verfügung - auch Sammelzahlung.
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
        
        If ea = "E" Then
            If Not dictE.Exists(kat) Then dictE.Add kat, True
        ElseIf ea = "A" Then
            If Not dictA.Exists(kat) Then dictA.Add kat, True
        End If
NextListRow:
    Next r
    
    ' Alte Listen löschen - Daten-Blatt kurz entsperren
    On Error Resume Next
    wsData.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
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
    
    On Error Resume Next
    wsData.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
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
    
    ' Zelle muss entsperrt sein für Validation
    wsBK.Cells(rowBK, BK_COL_KATEGORIE).Locked = False
    
    ' Alte Validierung sicher löschen
    On Error Resume Next
    wsBK.Cells(rowBK, BK_COL_KATEGORIE).Validation.Delete
    On Error GoTo 0
    
    ' Validierungs-Formel als Bereichsreferenz
    Dim listRange As String
    listRange = "='" & wsData.Name & "'!" & _
                wsData.Cells(DATA_START_ROW, listCol).Address(True, True) & ":" & _
                wsData.Cells(lastListRow, listCol).Address(True, True)
    
    ' Neue DropDown-Validierung setzen
    On Error Resume Next
    With wsBK.Cells(rowBK, BK_COL_KATEGORIE).Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:=listRange
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With
    If Err.Number <> 0 Then
        Debug.Print "DropDown Fehler Zeile " & rowBK & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
End Sub


' ===============================================================
' Re-Evaluierung nach EntityRole-Änderung (für eine IBAN)
' ===============================================================
Public Sub ReEvaluiereNachEntityRoleAenderung(ByVal geaenderteIBAN As String)

    On Error GoTo 0
    
    Dim wsBK As Worksheet
    Dim wsData As Worksheet
    Dim lastRuleRow As Long
    Dim lastRowBK As Long
    Dim r As Long
    Dim ibanClean As String
    Dim zeilenIBAN As String
    Dim anzahlNeu As Long
    
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsData = ThisWorkbook.Worksheets(WS_DATEN)
    
    ' Zeilengrenzen dynamisch
    lastRuleRow = wsData.Cells(wsData.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRuleRow < DATA_START_ROW Then Exit Sub
    
    ibanClean = UCase(Replace(geaenderteIBAN, " ", ""))
    If ibanClean = "" Then Exit Sub
    
    lastRowBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRowBK < BK_START_ROW Then Exit Sub
    
    ' Listen aktualisieren
    AktualisierKategorieListen
    
    ' Einstellungen-Cache laden
    LadeEinstellungenCache
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error Resume Next
    wsBK.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    anzahlNeu = 0
    
    For r = BK_START_ROW To lastRowBK
        zeilenIBAN = UCase(Replace(Trim(CStr(wsBK.Cells(r, BK_COL_IBAN).value)), " ", ""))
        If zeilenIBAN <> ibanClean Then GoTo NextRowReEval
        
        ' Manuelle Betragseingabe? NICHT anfassen
        If HatManuelleBetragseingabe(wsBK, r) Then GoTo NextRowReEval
        
        ' Manuelle Kategorie? NICHT anfassen
        If HatManuelleKategorie(wsBK, r) Then GoTo NextRowReEval
        
        ' Alte Kategorie, Bemerkung und Validierung löschen
        wsBK.Cells(r, BK_COL_KATEGORIE).value = ""
        wsBK.Cells(r, BK_COL_KATEGORIE).Interior.ColorIndex = xlNone
        wsBK.Cells(r, BK_COL_KATEGORIE).Font.color = vbBlack
        wsBK.Cells(r, BK_COL_BEMERKUNG).value = ""
        On Error Resume Next
        wsBK.Cells(r, BK_COL_KATEGORIE).Validation.Delete
        On Error GoTo 0
        
        ' Neu evaluieren (v9.0: wsData + lastRuleRow)
        On Error Resume Next
        EvaluateKategorieEngineRow wsBK, r, wsData, lastRuleRow
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        
        ' Betrag nur zuordnen wenn GRÜN
        On Error Resume Next
        If wsBK.Cells(r, BK_COL_KATEGORIE).Interior.color = RGB(198, 239, 206) Then
            ApplyBetragsZuordnung wsBK, r
        End If
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        
        ' DropDown für ROT und GELB setzen
        Dim reEvalFarbe As Long
        reEvalFarbe = wsBK.Cells(r, BK_COL_KATEGORIE).Interior.color
        If reEvalFarbe = RGB(255, 199, 206) Or reEvalFarbe = RGB(255, 235, 156) Then
            SetzeKategorieDropDown wsBK, r
        End If
        
        anzahlNeu = anzahlNeu + 1
        
NextRowReEval:
    Next r
    
    ' Einstellungen-Cache freigeben
    EntladeEinstellungenCache
    
    On Error Resume Next
    wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
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
    
    For c = BK_COL_EINNAHMEN_START To BK_COL_AUSGABEN_ENDE
        If wsBK.Cells(rowBK, c).value <> "" And wsBK.Cells(rowBK, c).value <> 0 Then
            HatManuelleBetragseingabe = True
            Exit Function
        End If
    Next c
End Function


' ===============================================================
' Prüft ob der Nutzer manuell eine Kategorie gewählt/geändert hat.
' Manuelle Kategorie = Kategorie-Zelle hat einen Wert UND
' die Zelle hat KEINE der Engine-Farben (GRÜN/GELB/ROT).
' GRÜN-Zeilen werden ebenfalls übersprungen, da sie bereits
' erfolgreich automatisch zugeordnet wurden.
' ===============================================================
Private Function HatManuelleKategorie(ByVal wsBK As Worksheet, _
                                       ByVal rowBK As Long) As Boolean
    HatManuelleKategorie = False
    
    Dim katWert As String
    katWert = Trim(CStr(wsBK.Cells(rowBK, BK_COL_KATEGORIE).value))
    
    ' Keine Kategorie -> nicht manuell, Engine soll evaluieren
    If katWert = "" Then Exit Function
    
    Dim katFarbe As Long
    katFarbe = wsBK.Cells(rowBK, BK_COL_KATEGORIE).Interior.color
    
    ' GRÜN = erfolgreich automatisch zugeordnet -> nicht überschreiben
    If katFarbe = RGB(198, 239, 206) Then
        HatManuelleKategorie = True
        Exit Function
    End If
    
    ' ROT oder GELB = Engine hat zugeordnet aber unsicher
    ' -> darf von der Engine bei Re-Evaluierung überschrieben werden
    If katFarbe = RGB(255, 199, 206) Then Exit Function  ' ROT
    If katFarbe = RGB(255, 235, 156) Then Exit Function  ' GELB
    
    ' Jede andere Farbe oder keine Farbe mit Wert = manuell
    HatManuelleKategorie = True
    
End Function

