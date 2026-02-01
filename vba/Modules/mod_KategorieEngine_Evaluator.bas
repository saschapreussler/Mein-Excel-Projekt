Attribute VB_Name = "mod_KategorieEngine_Evaluator"
Option Explicit

' =====================================================
' KATEGORIE-ENGINE - EVALUATOR (VOLLSTAENDIG UEBERARBEITET)
' VERSION: 2.0 - 01.02.2026
' AENDERUNG: Sonderregel fuer 0-Euro-Betraege bei ABSCHLUSS
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

    Dim kontoName As String
    kontoName = LCase(Trim(wsBK.Cells(rowBK, BK_COL_NAME).value))
    
    Dim buchungsText As String
    buchungsText = LCase(Trim(wsBK.Cells(rowBK, BK_COL_BUCHUNGSTEXT).value))

    ctx("Amount") = amount
    ctx("NormText") = normText
    ctx("KontoName") = kontoName
    ctx("IBAN") = iban
    ctx("BuchungsText") = buchungsText

    ctx("IsEinnahme") = (amount > 0)
    ctx("IsAusgabe") = (amount < 0)
    ctx("IsNullBetrag") = (amount = 0)  ' NEU: 0-Euro-Betraege

    ctx("EntityRole") = entityRole

    ' Entgeltabschluss-Erkennung (Bankgebuehren)
    ' ERWEITERT: Auch bei ABSCHLUSS im Buchungstext
    ctx("IsEntgeltabschluss") = _
        (InStr(normText, "entgeltabschluss") > 0) Or _
        (InStr(normText, "kontoabschluss") > 0) Or _
        (InStr(normText, "abschluss") > 0 And InStr(normText, "entgelt") > 0) Or _
        (buchungsText = "abschluss") Or _
        (buchungsText = "entgeltabschluss")

    ' Bargeldauszahlung-Erkennung
    ctx("IsBargeldauszahlung") = _
        (InStr(normText, "bargeld") > 0) Or _
        (InStr(normText, "auszahlung") > 0 And InStr(normText, "geldautomat") > 0) Or _
        (InStr(normText, "abhebung") > 0)

    Set BuildKategorieContext = ctx
End Function

' -----------------------------
' EntityRole ueber IBAN bestimmen
' -----------------------------
Private Function GetEntityRoleByIBAN(ByVal strIBAN As String) As String
    Dim wsD As Worksheet
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)

    Dim lastRow As Long
    lastRow = wsD.Cells(wsD.Rows.Count, DATA_MAP_COL_IBAN).End(xlUp).Row

    Dim ibanClean As String
    ibanClean = UCase(Replace(strIBAN, " ", ""))
    
    If ibanClean = "" Then
        GetEntityRoleByIBAN = ""
        Exit Function
    End If

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        If UCase(Replace(wsD.Cells(r, DATA_MAP_COL_IBAN).value, " ", "")) = ibanClean Then
            GetEntityRoleByIBAN = UCase(Trim(wsD.Cells(r, DATA_MAP_COL_ENTITYROLE).value))
            Exit Function
        End If
    Next r

    GetEntityRoleByIBAN = ""
End Function

' -----------------------------
' Hauptfunktion: Kategorie evaluieren
' -----------------------------
Public Sub EvaluateKategorieEngineRow(ByVal wsBK As Worksheet, _
                                      ByVal rowBK As Long, _
                                      ByVal rngRules As Range)

    ' Bereits kategorisiert? Ueberspringen
    If Trim(wsBK.Cells(rowBK, BK_COL_KATEGORIE).value) <> "" Then Exit Sub

    Dim ctx As Object
    Set ctx = BuildKategorieContext(wsBK, rowBK)

    ' ================================
    ' PHASE 0: SONDERREGEL FUER 0-EURO-BETRAEGE
    ' ================================
    ' Bei 0,00 Euro und Buchungstext "ABSCHLUSS" -> Entgeltabschluss (Kontofuehrung)
    If ctx("IsNullBetrag") And ctx("IsEntgeltabschluss") Then
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       "Entgeltabschluss (Kontofuehrung)", "GRUEN"
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = "0-Euro-Abschluss automatisch zugeordnet"
        ' Bei 0 Euro keine Betragszuordnung noetig
        Exit Sub
    End If

    ' ================================
    ' PHASE 1: HARTE SONDERREGELN
    ' ================================
    
    ' 1a) Entgeltabschluss (Bankgebuehren)
    If ctx("IsEntgeltabschluss") And ctx("IsAusgabe") Then
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       "Entgeltabschluss (Kontofuehrung)", "GRUEN"
        ApplyBetragsZuordnung wsBK, rowBK
        Exit Sub
    End If
    
    ' 1b) Bargeldauszahlung
    If ctx("IsBargeldauszahlung") And ctx("IsAusgabe") Then
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       "Bargeldauszahlung", "GRUEN"
        ApplyBetragsZuordnung wsBK, rowBK
        Exit Sub
    End If

    ' ================================
    ' PHASE 2: KEYWORD-MATCHING MIT SCORING
    ' ================================
    
    Dim bestCategory As String
    Dim bestScore As Long
    Dim bestPriority As Long
    bestScore = -999
    bestPriority = 999
    bestCategory = ""

    Dim hitCategories As Object
    Set hitCategories = CreateObject("Scripting.Dictionary")

    Dim ruleRow As Range
    For Each ruleRow In rngRules.Rows

        Dim category As String
        Dim einAus As String
        Dim keyword As String
        Dim prio As Long

        ' Spaltenreihenfolge: J=Kategorie, K=E/A, L=Keyword, M=Prioritaet
        category = Trim(ruleRow.Cells(1, 1).value)      ' Spalte J - Kategorie
        einAus = UCase(Trim(ruleRow.Cells(1, 2).value)) ' Spalte K - E/A
        keyword = Trim(ruleRow.Cells(1, 3).value)       ' Spalte L - Keyword
        prio = Val(ruleRow.Cells(1, 4).value)           ' Spalte M - Prioritaet
        If prio = 0 Then prio = 5

        If category = "" Or keyword = "" Then GoTo NextRule

        ' ================================
        ' FILTER 1: Einnahme/Ausgabe MUSS passen!
        ' Bei 0-Euro-Betraegen: beide erlauben
        ' ================================
        If Not ctx("IsNullBetrag") Then
            If einAus = "E" And ctx("IsAusgabe") Then GoTo NextRule
            If einAus = "A" And ctx("IsEinnahme") Then GoTo NextRule
        End If

        ' ================================
        ' FILTER 2: EntityRole-Trennung (wenn bekannt)
        ' ================================
        If ctx("EntityRole") <> "" Then
            ' VERSORGER darf keine Mitglieder-Kategorien bekommen
            If ctx("EntityRole") = "VERSORGER" Then
                If LCase(category) Like "*mitglied*" Then GoTo NextRule
                If LCase(category) Like "*pacht*" Then GoTo NextRule
                If LCase(category) Like "*endabrechnung*" Then GoTo NextRule
                If LCase(category) Like "*vorauszahlung*" Then GoTo NextRule
            End If
            
            ' MITGLIED darf keine Versorger-Kategorien bekommen
            If ctx("EntityRole") = "MITGLIED" Or _
               ctx("EntityRole") = "MITGLIED_MIT_PACHT" Or _
               ctx("EntityRole") = "MITGLIED_OHNE_PACHT" Then
                If LCase(category) Like "*versorger*" Then GoTo NextRule
                If LCase(category) Like "*stadtwerke*" Then GoTo NextRule
                If LCase(category) Like "*rueckzahlung versorger*" Then GoTo NextRule
            End If
            
            ' BANK darf nur Bank-Kategorien bekommen
            If ctx("EntityRole") = "BANK" Then
                If Not (LCase(category) Like "*bank*" Or _
                        LCase(category) Like "*entgelt*" Or _
                        LCase(category) Like "*gebuehr*" Or _
                        LCase(category) Like "*kontofuehrung*") Then GoTo NextRule
            End If
        End If

        ' ================================
        ' KEYWORD-MATCHING
        ' ================================
        Dim normKeyword As String
        normKeyword = NormalizeText(keyword)
        
        If InStr(ctx("NormText"), normKeyword) > 0 Then

            Dim score As Long
            score = 100
            
            score = score + (10 - prio) * 5
            
            If ctx("EntityRole") <> "" Then
                score = score + 20
            End If
            
            If (einAus = "E" And ctx("IsEinnahme")) Or _
               (einAus = "A" And ctx("IsAusgabe")) Then
                score = score + 15
            End If

            If Not hitCategories.Exists(category) Then
                hitCategories.Add category, prio
            End If

            If score > bestScore Or (score = bestScore And prio < bestPriority) Then
                bestScore = score
                bestPriority = prio
                bestCategory = category
            End If
        End If

NextRule:
    Next ruleRow

    ' ================================
    ' PHASE 3: ERGEBNIS AUSWERTEN
    ' ================================
    
    If hitCategories.Count > 1 Then
        If ctx("EntityRole") = "MITGLIED" And ctx("IsEinnahme") Then
            wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
                "Mehrere Positionen erkannt: " & Join(hitCategories.Keys, " | ")
            ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                           "Sammelzahlung Mitglied", "GELB"
            Exit Sub
        Else
            wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
                "Mehrere moegliche Kategorien: " & Join(hitCategories.Keys, " | ") & _
                " - Bitte pruefen!"
            ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), bestCategory, "GELB"
            Exit Sub
        End If
    End If

    If bestCategory <> "" Then
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), bestCategory, "GRUEN"
        ApplyBetragsZuordnung wsBK, rowBK
        Exit Sub
    End If

    If ctx("EntityRole") = "" Then
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            "Keine Kategorie gefunden. IBAN nicht zugeordnet - bitte Entity-Mapping pruefen!"
    Else
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            "Keine passende Kategorie gefunden (EntityRole: " & ctx("EntityRole") & ")"
    End If
    ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), "", "ROT"

End Sub

' -----------------------------
' Kategorie anwenden mit Ampelfarbe
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

