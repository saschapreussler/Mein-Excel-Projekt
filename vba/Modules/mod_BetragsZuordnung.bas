Attribute VB_Name = "mod_BetragsZuordnung"
Option Explicit

' ***************************************************************
' MODUL: mod_BetragsZuordnung
' ZWECK: Richtige Betragszuordnung je nach Einnahme / Ausgabe
'        unter Berücksichtigung identischer Zielspaltennamen
'        + Konsistenzprüfung Kategorie
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

    ' Zielüberschrift aus Kategorietabelle
    Dim targetHeader As String
    targetHeader = GetTargetHeaderByCategory(category)

    If targetHeader = "" Then
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            "Keine Zielspalte für Kategorie definiert"
        MarkKategorieRed wsBK, rowBK
        Exit Sub
    End If

    ' Zielspalte abhängig vom Vorzeichen suchen
    Dim targetCol As Long
    If betrag >= 0 Then
        ' Einnahmen M:S
        targetCol = FindBankkontoColumnByHeader(wsBK, targetHeader, 13, 19)
    Else
        ' Ausgaben T:Z
        targetCol = FindBankkontoColumnByHeader(wsBK, targetHeader, 20, 26)
    End If

    If targetCol = 0 Then
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            "Zielspalte '" & targetHeader & "' nicht im passenden Bereich gefunden"
        MarkKategorieRed wsBK, rowBK
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
            GetTargetHeaderByCategory = Trim(wsRules.Cells(r, 14).value) ' Spalte N
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
        If Trim(wsBK.Cells(27, c).value) = headerText Then
            FindBankkontoColumnByHeader = c
            Exit Function
        End If
    Next c

    FindBankkontoColumnByHeader = 0
End Function

' ---------------------------------------------------------------
' Kategorie als manuell zu prüfen markieren
' ---------------------------------------------------------------
Private Sub MarkKategorieRed(ByVal wsBK As Worksheet, ByVal rowBK As Long)
    With wsBK.Cells(rowBK, BK_COL_KATEGORIE)
        .Interior.color = RGB(255, 199, 206)
        .Font.color = vbRed
    End With
End Sub




