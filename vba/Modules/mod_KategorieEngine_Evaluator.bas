Attribute VB_Name = "mod_KategorieEngine_Evaluator"
Option Explicit

' =====================================================
' KATEGORIE-ENGINE - EVALUATOR (Kern-Modul)
' VERSION: 9.5 - 13.02.2026
' Refactoring: Aufgeteilt in 3 Module:
'   - mod_KategorieEngine_Evaluator (dieses Modul)
'     Kontext-Aufbau, Hauptschleife, Ergebnis-Ausgabe
'   - mod_KategorieEngine_Scoring
'     MatchKeyword, ExactMatchBonus, WordCountBonus,
'     PasstEntityRoleZuKategorie
'   - mod_KategorieEngine_Zeitraum
'     Einstellungen-Cache, Betrags-/Zeitvalidierung,
'     ErmittleMonatPeriode, IstMonatInListe
' =====================================================

' Mindest-Score-Differenz f�r sichere Zuordnung
Private Const SCORE_DOMINANZ_SCHWELLE As Long = 20

' Kategorie f�r echte Mehrdeutigkeit (nur programmatisch!)
Private Const KAT_SAMMELZAHLUNG As String = "Sammelzahlung (mehrere Positionen) Mitglied"

' Farbe f�r "Folgemonat manuell best�tigt" (hell-gr�n)
Private Const FARBE_HELLGRUEN As Long = 12968900  ' RGB(196, 225, 196) -> &HC4E1C4 -> Long


' -----------------------------
' EntityInfo �ber IBAN bestimmen (kombiniert: Role + Parzelle)
' -----------------------------
Private Sub GetEntityInfoByIBAN(ByVal strIBAN As String, _
                                 ByRef outRole As String, _
                                 ByRef outParzelle As String)
    outRole = ""
    outParzelle = ""
    
    Dim ibanClean As String
    ibanClean = UCase(Replace(strIBAN, " ", ""))
    If ibanClean = "" Then Exit Sub
    
    Dim wsD As Worksheet
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    
    Dim lastRow As Long
    lastRow = wsD.Cells(wsD.Rows.count, DATA_MAP_COL_IBAN).End(xlUp).Row
    
    Dim r As Long
    For r = DATA_START_ROW To lastRow
        If UCase(Replace(wsD.Cells(r, DATA_MAP_COL_IBAN).value, " ", "")) = ibanClean Then
            outRole = UCase(Trim(CStr(wsD.Cells(r, DATA_MAP_COL_ENTITYROLE).value)))
            outParzelle = Trim(CStr(wsD.Cells(r, DATA_MAP_COL_PARZELLE).value))
            Exit Sub
        End If
    Next r
End Sub

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

    ' Kombinierter Aufruf statt 2 separate Loops
    Dim entityRole As String
    Dim entityParzelle As String
    GetEntityInfoByIBAN iban, entityRole, entityParzelle

    Dim kontoname As String
    kontoname = LCase(Trim(wsBK.Cells(rowBK, BK_COL_NAME).value))
    
    Dim buchungstext As String
    buchungstext = LCase(Trim(wsBK.Cells(rowBK, BK_COL_BUCHUNGSTEXT).value))
    
    Dim buchungsDatum As Variant
    buchungsDatum = wsBK.Cells(rowBK, BK_COL_DATUM).value

    ctx("Amount") = amount
    ctx("AbsAmount") = Abs(amount)
    ctx("NormText") = normText
    ctx("KontoName") = kontoname
    ctx("IBAN") = iban
    ctx("BuchungsText") = buchungstext
    ctx("Datum") = buchungsDatum

    ctx("IsEinnahme") = (amount > 0)
    ctx("IsAusgabe") = (amount < 0)
    ctx("IsNullBetrag") = (amount = 0)

    ctx("EntityRole") = entityRole
    ctx("EntityParzelle") = entityParzelle
    
    ' FIX: Spaces statt Underscores! EntityKey Manager speichert
    ' "MITGLIED MIT PACHT" (Spaces), nicht "MITGLIED_MIT_PACHT"
    ctx("IsMitglied") = (entityRole = "MITGLIED" Or _
                          entityRole = "MITGLIED MIT PACHT" Or _
                          entityRole = "MITGLIED OHNE PACHT")
    
    ctx("IsEhemaligesMitglied") = (entityRole = "EHEMALIGES MITGLIED")
    
    ctx("IsVersorger") = (entityRole = "VERSORGER")
    ctx("IsBank") = (entityRole = "BANK")

    ctx("IsEntgeltabschluss") = _
        (InStr(normText, "entgeltabschluss") > 0) Or _
        (InStr(normText, "kontoabschluss") > 0) Or _
        (InStr(normText, "abschluss") > 0 And InStr(normText, "entgelt") > 0) Or _
        (buchungstext = "abschluss") Or _
        (buchungstext = "entgeltabschluss")

    ctx("IsBargeldauszahlung") = _
        (InStr(normText, "bargeld") > 0) Or _
        (InStr(normText, "auszahlung") > 0 And InStr(normText, "geldautomat") > 0) Or _
        (InStr(normText, "abhebung") > 0)

    Set BuildKategorieContext = ctx
End Function

' =====================================================
' Hauptfunktion: Kategorie evaluieren (v9.3)
' Braucht KEINEN Named Range! Liest Regeln direkt vom
' Daten-Blatt �ber DATA_CAT_COL_* Konstanten.
' Scoring-Logik aus v7.0 wiederhergestellt.
' v9.3: WordCountBonus + erh�hter Prio-Bonus
' =====================================================
Public Sub EvaluateKategorieEngineRow(ByVal wsBK As Worksheet, _
                                      ByVal rowBK As Long, _
                                      ByVal wsData As Worksheet, _
                                      ByVal lastRuleRow As Long)

    ' Bereits kategorisiert? �berspringen
    If Trim(wsBK.Cells(rowBK, BK_COL_KATEGORIE).value) <> "" Then Exit Sub

    Dim ctx As Object
    Set ctx = BuildKategorieContext(wsBK, rowBK)

    ' ================================
    ' PHASE 0: SONDERREGEL F�R 0-EURO-BETR�GE
    ' ================================
    If ctx("IsNullBetrag") And ctx("IsEntgeltabschluss") Then
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       "Entgeltabschluss (Kontof" & ChrW(252) & "hrung)", "GRUEN"
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = "0-Euro-Abschluss automatisch zugeordnet"
        Exit Sub
    End If

    ' 0-Euro ohne Sonderregel -> �berspringen
    If ctx("IsNullBetrag") Then Exit Sub

    Dim normText As String
    normText = ctx("NormText")
    If normText = "" Then Exit Sub

    ' ================================
    ' PHASE 1: HARTE SONDERREGELN
    ' ================================
    
    ' 1a) Entgeltabschluss (Bankgeb�hren)
    If ctx("IsEntgeltabschluss") And ctx("IsAusgabe") Then
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       "Entgeltabschluss (Kontof" & ChrW(252) & "hrung)", "GRUEN"
        Exit Sub
    End If
    
    ' 1b) Bargeldauszahlung
    If ctx("IsBargeldauszahlung") And ctx("IsAusgabe") Then
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       "Bargeldauszahlung", "GRUEN"
        Exit Sub
    End If

    ' ================================
    ' PHASE 2: KEYWORD-MATCHING MIT ERWEITERTEM SCORING
    ' ================================
    
    Dim bestCategory As String
    Dim bestScore As Long
    Dim bestPriority As Long
    bestScore = -999
    bestPriority = 999
    bestCategory = ""

    ' Dictionary: Kategorie -> Score (h�chster Score je Kategorie)
    Dim hitCategories As Object
    Set hitCategories = CreateObject("Scripting.Dictionary")

    Dim dataRow As Long
    For dataRow = DATA_START_ROW To lastRuleRow

        Dim category As String
        Dim einAus As String
        Dim keyword As String
        Dim prio As Long
        Dim faelligkeit As String

        ' Spalten �ber Konstanten lesen
        category = Trim(CStr(wsData.Cells(dataRow, DATA_CAT_COL_KATEGORIE).value))    ' J
        einAus = UCase(Trim(CStr(wsData.Cells(dataRow, DATA_CAT_COL_EINAUS).value)))   ' K
        keyword = Trim(CStr(wsData.Cells(dataRow, DATA_CAT_COL_KEYWORD).value))        ' L
        prio = Val(wsData.Cells(dataRow, DATA_CAT_COL_PRIORITAET).value)               ' M
        faelligkeit = LCase(Trim(CStr(wsData.Cells(dataRow, DATA_CAT_COL_FAELLIGKEIT).value))) ' O
        If prio = 0 Then prio = 5

        If category = "" Or keyword = "" Then GoTo NextRule

        ' ================================
        ' FILTER 0: Sammelzahlung-Kategorie NIEMALS per Keyword!
        ' ================================
        If LCase(category) Like "*sammelzahlung*" Then GoTo NextRule

        ' ================================
        ' FILTER 1: Einnahme/Ausgabe MUSS passen
        ' ================================
        If Not ctx("IsNullBetrag") Then
            If einAus = "E" And ctx("IsAusgabe") Then GoTo NextRule
            If einAus = "A" And ctx("IsEinnahme") Then GoTo NextRule
        End If

        ' ================================
        ' FILTER 2: Strenge EntityRole-Trennung
        ' ================================
        If Not PasstEntityRoleZuKategorie(ctx, category, einAus) Then GoTo NextRule

        ' ================================
        ' KEYWORD-MATCHING (Multi-Word v7.0)
        ' ================================
        Dim normKeyword As String
        normKeyword = NormalizeText(keyword)
        
        If MatchKeyword(normText, normKeyword) Then

            Dim score As Long
            score = 100
            
            ' Priorit�tsbonus (niedrigere Prio = h�herer Bonus)
            ' v9.3: Faktor 8 statt 5 f�r st�rkere Differenzierung
            score = score + (10 - prio) * 8
            
            ' EntityRole bekannt = h�here Konfidenz (+20 wie in v7.0)
            If ctx("EntityRole") <> "" Then
                score = score + 20
            End If
            
            ' Einnahme/Ausgabe stimmt exakt �berein
            If (einAus = "E" And ctx("IsEinnahme")) Or _
               (einAus = "A" And ctx("IsAusgabe")) Then
                score = score + 15
            End If
            
            ' Keyword-L�nge als Qualit�tsfaktor
            Dim kwLen As Long
            kwLen = Len(normKeyword)
            If kwLen >= 12 Then
                score = score + 20
            ElseIf kwLen >= 8 Then
                score = score + 12
            ElseIf kwLen >= 5 Then
                score = score + 5
            End If
            
            ' ExactMatchBonus (v8.0: +10 wenn Keyword zusammenh�ngend im Text)
            score = score + ExactMatchBonus(normText, normKeyword)
            
            ' WordCountBonus (v9.3: Anzahl W�rter im Keyword * 5)
            score = score + WordCountBonus(normKeyword)
            
             ' Betragsvalidierung �ber Einstellungen
            Dim betragBonus As Long
            betragBonus = PruefeBetragGegenEinstellungen(category, ctx("AbsAmount"))
            score = score + betragBonus
            
             ' Zeitfenstervalidierung �ber Einstellungen
            If IsDate(ctx("Datum")) Then
                Dim zeitBonus As Long
                zeitBonus = PruefeZeitfenster(category, CDate(ctx("Datum")), faelligkeit)
                score = score + zeitBonus
            End If

            If Not hitCategories.Exists(category) Then
                hitCategories.Add category, score
            Else
                If score > CLng(hitCategories(category)) Then
                    hitCategories(category) = score
                End If
            End If

            If score > bestScore Or (score = bestScore And prio < bestPriority) Then
                bestScore = score
                bestPriority = prio
                bestCategory = category
            End If
        End If

NextRule:
    Next dataRow

    ' ================================
    ' PHASE 3: ERGEBNIS AUSWERTEN MIT SCORE-DOMINANZ
    ' ================================
    
    If hitCategories.count > 1 Then
        ' Zweitbesten Score ermitteln
        Dim zweitBesterScore As Long
        zweitBesterScore = -999
        Dim katKey As Variant
        For Each katKey In hitCategories.keys
            If CStr(katKey) <> bestCategory Then
                If CLng(hitCategories(katKey)) > zweitBesterScore Then
                    zweitBesterScore = CLng(hitCategories(katKey))
                End If
            End If
        Next katKey
        
        Dim scoreDifferenz As Long
        scoreDifferenz = bestScore - zweitBesterScore
        
        If scoreDifferenz >= SCORE_DOMINANZ_SCHWELLE Then
            ' SICHERER TREFFER trotz mehrerer Matches
            ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), bestCategory, "GRUEN"
            Exit Sub
        End If
        
        ' ECHTE MEHRDEUTIGKEIT: Detaillierte Bemerkung
        ' v9.1: Kein Instruktions-Satz, nur Kategorieliste
        Dim bemerkung As String
        bemerkung = hitCategories.count & " Kategorien passen:" & vbLf
        
        Dim katNr As Long
        katNr = 0
        For Each katKey In hitCategories.keys
            katNr = katNr + 1
            If katNr < hitCategories.count Then
                bemerkung = bemerkung & katNr & ") " & CStr(katKey) & vbLf
            Else
                ' Letzte Kategorie: KEIN abschlie�endes vbLf
                bemerkung = bemerkung & katNr & ") " & CStr(katKey)
            End If
        Next katKey
        
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = bemerkung
        
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       KAT_SAMMELZAHLUNG, "GELB"
        
        Call EntsperreBetragsspalten(wsBK, rowBK, ctx("IsEinnahme"))
        Exit Sub
    End If

    ' Genau 1 Treffer = sicher GR�N
    If bestCategory <> "" Then
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), bestCategory, "GRUEN"
        Exit Sub
    End If

    ' Kein Treffer = ROT
    If ctx("EntityRole") = "" Then
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
        "Keine Kategorie gefunden. IBAN nicht zugeordnet - bitte Entity-Mapping pr�fen!"
    Else
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            "Keine passende Kategorie gefunden (EntityRole: " & ctx("EntityRole") & ")"
    End If
    ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), "Bitte Auswahl treffen!", "ROT"

End Sub


' =====================================================
' Betragsspalten entsperren f�r manuelle Eingabe
' =====================================================
Private Sub EntsperreBetragsspalten(ByVal wsBK As Worksheet, _
                                    ByVal rowBK As Long, _
                                    ByVal istEinnahme As Boolean)
    Dim startCol As Long
    Dim endCol As Long
    Dim c As Long
    
    If istEinnahme Then
        startCol = BK_COL_EINNAHMEN_START
        endCol = BK_COL_EINNAHMEN_ENDE
    Else
        startCol = BK_COL_AUSGABEN_START
        endCol = BK_COL_AUSGABEN_ENDE
    End If
    
    On Error Resume Next
    For c = startCol To endCol
        wsBK.Cells(rowBK, c).Locked = False
    Next c
    On Error GoTo 0
End Sub


' -----------------------------
' Kategorie anwenden mit Ampelfarbe (originale Signatur v7.0)
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

