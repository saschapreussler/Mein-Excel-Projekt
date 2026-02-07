Attribute VB_Name = "mod_BetragsZuordnung"
Option Explicit

' ***************************************************************
' MODUL: mod_BetragsZuordnung
' VERSION: 2.0 - 08.02.2026
' FIX: GELB (Sammelzahlung) wird übersprungen - keine
'      Betragszuordnung und kein Überschreiben auf ROT!
'      Bessere Bemerkungen bei fehlender Zielspalte.
' ***************************************************************

Public Sub ApplyBetragsZuordnung(ByVal wsBK As Worksheet, _
                                 ByVal rowBK As Long)

    Dim betrag As Double
    betrag = wsBK.Cells(rowBK, BK_COL_BETRAG).value
    If betrag = 0 Then Exit Sub

    Dim category As String
    category = Trim(wsBK.Cells(rowBK, BK_COL_KATEGORIE).value)
    If category = "" Then Exit Sub

    ' ROT = manuelle Nacharbeit ? keine Automatik
    If wsBK.Cells(rowBK, BK_COL_KATEGORIE).Interior.color = RGB(255, 199, 206) Then Exit Sub
    
    ' GELB = Sammelzahlung/Mehrdeutigkeit ? NICHT anfassen!
    ' Der Nutzer muss die Beträge manuell aufteilen.
    If wsBK.Cells(rowBK, BK_COL_KATEGORIE).Interior.color = RGB(255, 235, 156) Then Exit Sub

    ' Zielüberschrift aus Kategorietabelle
    Dim targetHeader As String
    targetHeader = GetTargetHeaderByCategory(category)

    If targetHeader = "" Then
        ' Nur Bemerkung setzen wenn noch keine vorhanden
        If Trim(wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value) = "" Then
            wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
                "Kategorie '" & category & "' hat keine Zielspalte " & _
                "(Spalte N in der Kategorie-Tabelle auf Daten! ist leer)"
        End If
        ' NICHT auf ROT setzen - die Kategorie selbst kann korrekt sein!
        Exit Sub
    End If

    ' Zielspalte abhängig vom Vorzeichen suchen
    Dim targetCol As Long
    If betrag >= 0 Then
        targetCol = FindBankkontoColumnByHeader(wsBK, targetHeader, 13, 19)
    Else
        targetCol = FindBankkontoColumnByHeader(wsBK, targetHeader, 20, 26)
    End If

    If targetCol = 0 Then
        If Trim(wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value) = "" Then
            wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
                "Zielspalte '" & targetHeader & "' nicht im " & _
                IIf(betrag >= 0, "Einnahmen", "Ausgaben") & "-Bereich gefunden"
        End If
        Exit Sub
    End If

    ' Betrag eintragen
    wsBK.Cells(rowBK, targetCol).value = betrag

End Sub

' ---------------------------------------------------------------
' Zielspaltenname aus Kategorietabelle (Spalte N)
' ---------------------------------------------------------------
Private Function GetTargetHeaderByCategory(ByVal category As String) As String

    Dim wsRules As Worksheet
    Set wsRules = ThisWorkbook.Worksheets(WS_DATEN)

    Dim lastRow As Long
    lastRow = wsRules.Cells(wsRules.Rows.count, 10).End(xlUp).Row

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        If Trim(wsRules.Cells(r, 10).value) = category Then
            GetTargetHeaderByCategory = Trim(wsRules.Cells(r, 14).value)
            Exit Function
        End If
    Next r

    GetTargetHeaderByCategory = ""
End Function

' ---------------------------------------------------------------
' Zielspalte anhand Überschrift UND Bereich finden
' ---------------------------------------------------------------
Private Function FindBankkontoColumnByHeader(ByVal wsBK As Worksheet, _
                                             ByVal headerText As String, _
                                             ByVal firstCol As Long, _
                                             ByVal lastCol As Long) As Long
    Dim c As Long
    For c = firstCol To lastCol
        If Trim(wsBK.Cells(BK_HEADER_ROW, c).value) = headerText Then
            FindBankkontoColumnByHeader = c
            Exit Function
        End If
    Next c

    FindBankkontoColumnByHeader = 0
End Function

