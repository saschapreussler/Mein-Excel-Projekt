Attribute VB_Name = "mod_KategorieEngine_Evaluator"
Option Explicit

' =====================================================
' KATEGORIE-ENGINE – KONTEXT + EVALUATOR (FINAL)
' =====================================================

' -----------------------------
' Kontext erstellen
' -----------------------------
Public Function BuildKategorieContext(ByVal wsBK As Worksheet, _
                                      ByVal rowBK As Long) As Object
    Dim ctx As Object
    Set ctx = CreateObject("Scripting.Dictionary")

    Dim amount As Double
    amount = wsBK.Cells(rowBK, BK_COL_BETRAG).value

    Dim normText As String
    normText = NormalizeBankkontoZeile(wsBK, rowBK)

    Dim iban As String
    iban = Trim(wsBK.Cells(rowBK, BK_COL_IBAN).value)

    Dim entityRole As String
    entityRole = GetEntityRoleByIBAN(iban)

    ctx("Amount") = amount
    ctx("NormText") = normText

    ctx("IsEinnahme") = (amount > 0)
    ctx("IsAusgabe") = (amount < 0)

    ctx("EntityRole") = entityRole
    ctx("EntityType") = entityRole   ' bewusst identisch geführt

    ctx("IsEntgeltabschluss") = _
        (InStr(normText, "entgeltabschluss") > 0) Or _
        (InStr(normText, "abschluss") > 0)

    ' Rollenbasierte Rückzahlung
    ctx("IsRueckzahlungVersorger") = _
        (entityRole = "VERSORGER" And ctx("IsEinnahme"))

    ctx("IsRueckzahlungMitglied") = _
        (entityRole = "MITGLIED" And ctx("IsAusgabe"))

    Set BuildKategorieContext = ctx
End Function

' -----------------------------
' EntityRole über IBAN bestimmen
' -----------------------------
Private Function GetEntityRoleByIBAN(ByVal strIBAN As String) As String
    Dim wsD As Worksheet
    Set wsD = Worksheets(WS_DATEN)

    Dim lastRow As Long
    lastRow = wsD.Cells(wsD.Rows.Count, DATA_MAP_COL_IBAN).End(xlUp).Row

    Dim ibanClean As String
    ibanClean = UCase(Replace(strIBAN, " ", ""))

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        If UCase(Replace(wsD.Cells(r, DATA_MAP_COL_IBAN).value, " ", "")) = ibanClean Then
            GetEntityRoleByIBAN = Trim(wsD.Cells(r, DATA_MAP_COL_ENTITYROLE).value)
            Exit Function
        End If
    Next r

    GetEntityRoleByIBAN = ""
End Function

' -----------------------------
' Evaluator
' -----------------------------
Public Sub EvaluateKategorieEngineRow(ByVal wsBK As Worksheet, _
                                      ByVal rowBK As Long, _
                                      ByVal rngRules As Range)

    If Trim(wsBK.Cells(rowBK, BK_COL_KATEGORIE).value) <> "" Then Exit Sub

    Dim ctx As Object
    Set ctx = BuildKategorieContext(wsBK, rowBK)

    ' --------------------------------
    ' HARTE SICHERHEITSREGEL
    ' --------------------------------
    If ctx("IsEntgeltabschluss") Then
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       "Entgeltabschluss (Kontoführung)", "GRUEN"
        ' Direkt Betragszuordnung
        ApplyBetragsZuordnung wsBK, rowBK
        Exit Sub
    End If

    Dim bestCategory As String
    Dim bestScore As Long
    bestScore = -999

    Dim hitCategories As Object
    Set hitCategories = CreateObject("Scripting.Dictionary")

    Dim ruleRow As Range
    For Each ruleRow In rngRules.Rows

        Dim category As String
        Dim keyword As String
        Dim prio As Long
        Dim einAus As String

        category = Trim(ruleRow.Cells(1, 1).value)
        einAus = Trim(ruleRow.Cells(1, 2).value)
        keyword = Trim(ruleRow.Cells(1, 3).value)
        prio = Val(ruleRow.Cells(1, 4).value)

        If category = "" Or keyword = "" Then GoTo NextRule

        ' --------------------------------
        ' ROLLENTRENNUNG – ABSOLUT
        ' --------------------------------
        If ctx("EntityRole") = "VERSORGER" Then
            If LCase(category) Like "*mitglied*" Then GoTo NextRule
        End If

        If ctx("EntityRole") = "MITGLIED" Then
            If LCase(category) Like "*versorger*" Then GoTo NextRule
        End If

        keyword = NormalizeText(keyword)

        If InStr(ctx("NormText"), keyword) > 0 Then

            If (einAus = "E" And ctx("IsAusgabe")) Or _
               (einAus = "A" And ctx("IsEinnahme")) Then GoTo NextRule

            If Not hitCategories.Exists(category) Then
                hitCategories.Add category, True
            End If

            Dim score As Long
            score = 10 - prio

            If einAus = "E" And ctx("IsEinnahme") Then score = score + 2
            If einAus = "A" And ctx("IsAusgabe") Then score = score + 2

            If ctx("IsRueckzahlungVersorger") Then score = score + 3
            If ctx("IsRueckzahlungMitglied") Then score = score - 1

            If score > bestScore Then
                bestScore = score
                bestCategory = category
            End If
        End If

NextRule:
    Next ruleRow

    ' --------------------------------
    ' SAMMELZAHLUNG MITGLIED
    ' --------------------------------
    If hitCategories.Count > 1 _
       And ctx("EntityRole") = "MITGLIED" _
       And ctx("IsEinnahme") Then

        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            "Mehrere Positionen: " & Join(hitCategories.Keys, " | ")

        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       "Sammelzahlung Mitglied (mehrere Positionen)", "GELB"

        ' ROT für manuelle Nacharbeit, keine automatische Betragszuordnung
        wsBK.Cells(rowBK, BK_COL_KATEGORIE).Interior.color = RGB(255, 199, 206)
        Exit Sub
    End If

    ' --------------------------------
    ' ERGEBNIS
    ' --------------------------------
    If bestCategory <> "" Then
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), bestCategory, "GRUEN"
        ' Direkt Betragszuordnung auf Zielspalte
        ApplyBetragsZuordnung wsBK, rowBK
    Else
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), "", "ROT"
    End If
End Sub

' -----------------------------
' ApplyKategorie
' -----------------------------
Public Sub ApplyKategorie(ByVal targetCell As Range, _
                          ByVal category As String, _
                          ByVal confidence As String)
    With targetCell
        .value = category
        .Font.color = vbBlack
        .Interior.Pattern = xlSolid

        Select Case confidence
            Case "GRUEN": .Interior.color = RGB(198, 239, 206)
            Case "GELB":  .Interior.color = RGB(255, 235, 156)
            Case "ROT"
                .Interior.color = RGB(255, 199, 206)
                .Font.color = vbRed
        End Select
    End With
End Sub




