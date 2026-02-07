Attribute VB_Name = "mod_KategorieEngine_Pipeline"
Option Explicit

' ===============================================================
' KATEGORIEENGINE PIPELINE
' VERSION: 2.0 - 08.02.2026
' FIX: ApplyBetragsZuordnung nur bei GRÜN aufrufen,
'      nicht bei GELB (Sammelzahlung) oder ROT.
'      Neue Sub für Re-Evaluierung nach EntityRole-Änderung.
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

nextRow:
    Next r

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

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
        ' Manuell geändert = Spalte K (Status) enthält "Manuell" oder
        ' die Kategorie wurde vom Nutzer überschrieben.
        ' Kriterium: Wenn die Zeile GRÜN ist und einen Betrag in M-Z hat,
        ' wurde sie korrekt zugeordnet ? nur neu evaluieren wenn ROT oder GELB
        
        kategorieFarbe = wsBK.Cells(r, BK_COL_KATEGORIE).Interior.color
        
        ' GRÜN = erfolgreich zugeordnet ? NICHT anfassen
        If kategorieFarbe = RGB(198, 239, 206) Then GoTo NextRowReEval
        
        ' Prüfen ob der Nutzer manuell etwas in den Betragsspalten eingetragen hat
        If HatManuelleBetragseingabe(wsBK, r) Then GoTo NextRowReEval
        
        ' Diese Zeile darf neu evaluiert werden!
        ' Alte Kategorie und Bemerkung löschen
        wsBK.Cells(r, BK_COL_KATEGORIE).value = ""
        wsBK.Cells(r, BK_COL_KATEGORIE).Interior.ColorIndex = xlNone
        wsBK.Cells(r, BK_COL_KATEGORIE).Font.color = vbBlack
        wsBK.Cells(r, BK_COL_BEMERKUNG).value = ""
        
        ' Neu evaluieren
        EvaluateKategorieEngineRow wsBK, r, rngRules
        
        ' Betrag nur zuordnen wenn GRÜN
        If wsBK.Cells(r, BK_COL_KATEGORIE).Interior.color = RGB(198, 239, 206) Then
            ApplyBetragsZuordnung wsBK, r
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

