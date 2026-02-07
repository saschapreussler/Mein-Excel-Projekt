Attribute VB_Name = "mod_KategorieEngine_Evaluator"
Option Explicit

' =====================================================
' KATEGORIE-ENGINE - EVALUATOR
' VERSION: 4.0 - 07.02.2026
' ÄNDERUNG: Score-Dominanz-Logik: Wenn bester Score
'           deutlich höher als zweitbester ? GRÜN (sicher).
'           Sammelzahlung-Kategorie bei echter Mehrdeutigkeit.
'           Editierbare Betragsspalten bei GELB.
'           Strengere EntityRole-Trennung.
' =====================================================

' Mindest-Score-Differenz für sichere Zuordnung
Private Const SCORE_DOMINANZ_SCHWELLE As Long = 20

' Kategorie für echte Mehrdeutigkeit bei Mitgliedern
Private Const KAT_SAMMELZAHLUNG As String = "Sammelzahlung (mehrere Positionen) Mitglied"

' -----------------------------
' Kontext erstellen (erweitert)
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
    
    Dim entityParzelle As String
    entityParzelle = GetEntityParzelleByIBAN(iban)

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
    
    ' Ist es ein Mitglied (alle Varianten)?
    ctx("IsMitglied") = (entityRole = "MITGLIED" Or _
                          entityRole = "MITGLIED_MIT_PACHT" Or _
                          entityRole = "MITGLIED_OHNE_PACHT")
    
    ctx("IsVersorger") = (entityRole = "VERSORGER")
    ctx("IsBank") = (entityRole = "BANK")

    ' Entgeltabschluss-Erkennung (Bankgebühren)
    ctx("IsEntgeltabschluss") = _
        (InStr(normText, "entgeltabschluss") > 0) Or _
        (InStr(normText, "kontoabschluss") > 0) Or _
        (InStr(normText, "abschluss") > 0 And InStr(normText, "entgelt") > 0) Or _
        (buchungstext = "abschluss") Or _
        (buchungstext = "entgeltabschluss")

    ' Bargeldauszahlung-Erkennung
    ctx("IsBargeldauszahlung") = _
        (InStr(normText, "bargeld") > 0) Or _
        (InStr(normText, "auszahlung") > 0 And InStr(normText, "geldautomat") > 0) Or _
        (InStr(normText, "abhebung") > 0)

    Set BuildKategorieContext = ctx
End Function

' -----------------------------
' EntityRole über IBAN bestimmen
' -----------------------------
Private Function GetEntityRoleByIBAN(ByVal strIBAN As String) As String
    Dim wsD As Worksheet
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)

    Dim lastRow As Long
    lastRow = wsD.Cells(wsD.Rows.count, DATA_MAP_COL_IBAN).End(xlUp).Row

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
' Parzelle über IBAN bestimmen
' -----------------------------
Private Function GetEntityParzelleByIBAN(ByVal strIBAN As String) As String
    Dim wsD As Worksheet
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)

    Dim lastRow As Long
    lastRow = wsD.Cells(wsD.Rows.count, DATA_MAP_COL_IBAN).End(xlUp).Row

    Dim ibanClean As String
    ibanClean = UCase(Replace(strIBAN, " ", ""))
    
    If ibanClean = "" Then
        GetEntityParzelleByIBAN = ""
        Exit Function
    End If

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        If UCase(Replace(wsD.Cells(r, DATA_MAP_COL_IBAN).value, " ", "")) = ibanClean Then
            GetEntityParzelleByIBAN = Trim(wsD.Cells(r, DATA_MAP_COL_PARZELLE).value)
            Exit Function
        End If
    Next r

    GetEntityParzelleByIBAN = ""
End Function

' -----------------------------
' Hauptfunktion: Kategorie evaluieren
' -----------------------------
Public Sub EvaluateKategorieEngineRow(ByVal wsBK As Worksheet, _
                                      ByVal rowBK As Long, _
                                      ByVal rngRules As Range)

    ' Bereits kategorisiert? Überspringen
    If Trim(wsBK.Cells(rowBK, BK_COL_KATEGORIE).value) <> "" Then Exit Sub

    Dim ctx As Object
    Set ctx = BuildKategorieContext(wsBK, rowBK)

    ' ================================
    ' PHASE 0: SONDERREGEL FÜR 0-EURO-BETRÄGE
    ' ================================
    If ctx("IsNullBetrag") And ctx("IsEntgeltabschluss") Then
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       "Entgeltabschluss (Kontoführung)", "GRUEN"
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = "0-Euro-Abschluss automatisch zugeordnet"
        Exit Sub
    End If

    ' ================================
    ' PHASE 1: HARTE SONDERREGELN
    ' ================================
    
    ' 1a) Entgeltabschluss (Bankgebühren)
    If ctx("IsEntgeltabschluss") And ctx("IsAusgabe") Then
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       "Entgeltabschluss (Kontoführung)", "GRUEN"
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
    ' PHASE 2: KEYWORD-MATCHING MIT ERWEITERTEM SCORING
    ' ================================
    
    Dim bestCategory As String
    Dim bestScore As Long
    Dim bestPriority As Long
    bestScore = -999
    bestPriority = 999
    bestCategory = ""

    ' Dictionary: Kategorie -> Score (höchster Score je Kategorie)
    Dim hitCategories As Object
    Set hitCategories = CreateObject("Scripting.Dictionary")

    Dim ruleRow As Range
    For Each ruleRow In rngRules.Rows

        Dim category As String
        Dim einAus As String
        Dim keyword As String
        Dim prio As Long
        Dim faelligkeit As String

        ' Spalten: J=Kategorie, K=E/A, L=Keyword, M=Priorität, N=Zielspalte, O=Fälligkeit
        category = Trim(ruleRow.Cells(1, 1).value)      ' Spalte J
        einAus = UCase(Trim(ruleRow.Cells(1, 2).value))  ' Spalte K
        keyword = Trim(ruleRow.Cells(1, 3).value)        ' Spalte L
        prio = Val(ruleRow.Cells(1, 4).value)            ' Spalte M
        faelligkeit = LCase(Trim(ruleRow.Cells(1, 6).value)) ' Spalte O
        If prio = 0 Then prio = 5

        If category = "" Or keyword = "" Then GoTo NextRule

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
        ' KEYWORD-MATCHING
        ' ================================
        Dim normKeyword As String
        normKeyword = NormalizeText(keyword)
        
        If InStr(ctx("NormText"), normKeyword) > 0 Then

            Dim score As Long
            score = 100
            
            ' Prioritätsbonus
            score = score + (10 - prio) * 5
            
            ' EntityRole bekannt = höhere Konfidenz
            If ctx("EntityRole") <> "" Then
                score = score + 20
            End If
            
            ' Einnahme/Ausgabe stimmt exakt überein
            If (einAus = "E" And ctx("IsEinnahme")) Or _
               (einAus = "A" And ctx("IsAusgabe")) Then
                score = score + 15
            End If
            
            ' Keyword-Länge als Qualitätsfaktor:
            ' Längere Keywords sind spezifischer und verdienen mehr Score
            Dim kwLen As Long
            kwLen = Len(normKeyword)
            If kwLen >= 10 Then
                score = score + 15        ' Sehr spezifisches Keyword
            ElseIf kwLen >= 6 Then
                score = score + 8         ' Mittleres Keyword
            Else
                score = score + 0         ' Kurzes, generisches Keyword (z.B. "wasser")
            End If
            
            ' Betragsvalidierung über Einstellungen
            Dim betragBonus As Long
            betragBonus = PruefeBetragGegenEinstellungen(category, ctx("AbsAmount"))
            score = score + betragBonus
            
            ' Zeitfenstervalidierung über Einstellungen
            If IsDate(ctx("Datum")) Then
                Dim zeitBonus As Long
                zeitBonus = PruefeZeitfenster(category, CDate(ctx("Datum")), faelligkeit)
                score = score + zeitBonus
            End If

            If Not hitCategories.Exists(category) Then
                hitCategories.Add category, score
            Else
                ' Höheren Score behalten
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
    Next ruleRow

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
        
        ' Score-Dominanz prüfen:
        ' Wenn der beste Score DEUTLICH höher ist als der zweitbeste,
        ' ist die Zuordnung eindeutig trotz mehrerer Keyword-Treffer
        Dim scoreDifferenz As Long
        scoreDifferenz = bestScore - zweitBesterScore
        
        If scoreDifferenz >= SCORE_DOMINANZ_SCHWELLE Then
            ' ========================================
            ' SICHERER TREFFER trotz mehrerer Matches
            ' Der beste Kandidat dominiert klar
            ' ========================================
            ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), bestCategory, "GRUEN"
            ApplyBetragsZuordnung wsBK, rowBK
            Exit Sub
        End If
        
        ' ========================================
        ' ECHTE MEHRDEUTIGKEIT: Scores liegen nah beieinander
        ' ========================================
        
        ' Liste der konkurrierenden Kategorien erstellen
        Dim konkurrenten As String
        konkurrenten = ""
        For Each katKey In hitCategories.keys
            If konkurrenten <> "" Then konkurrenten = konkurrenten & " | "
            konkurrenten = konkurrenten & CStr(katKey) & " (" & hitCategories(katKey) & ")"
        Next katKey
        
        If ctx("IsMitglied") Then
            ' Mitglied ? Sammelzahlung-Kategorie
            wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
                "Mehrere Positionen möglich: " & konkurrenten & _
                " - Bitte manuell in den Betragsspalten aufteilen!"
            ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                           KAT_SAMMELZAHLUNG, "GELB"
            
            ' Betragsspalten editierbar machen (Zellschutz aufheben)
            Call EntsperreBetragsspalten(wsBK, rowBK, ctx("IsEinnahme"))
            Exit Sub
        Else
            ' Kein Mitglied ? normale Mehrdeutigkeitsmeldung
            wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
                "Mehrere mögliche Kategorien: " & konkurrenten & _
                " - Bitte prüfen!"
            ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                           KAT_SAMMELZAHLUNG, "GELB"
            
            Call EntsperreBetragsspalten(wsBK, rowBK, ctx("IsEinnahme"))
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
            "Keine Kategorie gefunden. IBAN nicht zugeordnet - bitte Entity-Mapping prüfen!"
    Else
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            "Keine passende Kategorie gefunden (EntityRole: " & ctx("EntityRole") & ")"
    End If
    ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), "", "ROT"

End Sub


' =====================================================
' NEU: Betragsspalten entsperren für manuelle Eingabe
' Bei GELB/Sammelzahlung soll der Nutzer die Beträge
' manuell auf die richtigen Spalten aufteilen können.
' Einnahme (Betrag > 0) ? Spalten M-S editierbar
' Ausgabe  (Betrag < 0) ? Spalten T-Z editierbar
' =====================================================
Private Sub EntsperreBetragsspalten(ByVal wsBK As Worksheet, _
                                    ByVal rowBK As Long, _
                                    ByVal istEinnahme As Boolean)
    Dim startCol As Long
    Dim endCol As Long
    Dim c As Long
    
    If istEinnahme Then
        startCol = BK_COL_EINNAHMEN_START   ' M = 13
        endCol = BK_COL_EINNAHMEN_ENDE      ' S = 19
    Else
        startCol = BK_COL_AUSGABEN_START    ' T = 20
        endCol = BK_COL_AUSGABEN_ENDE       ' Z = 26
    End If
    
    For c = startCol To endCol
        wsBK.Cells(rowBK, c).Locked = False
    Next c
End Sub


' =====================================================
' FILTER: Strenge EntityRole-Kategorie-Trennung
' Verhindert unlogische Zuordnungen wie
' "Strom Rückzahlung Versorger" bei MITGLIED
' =====================================================
Private Function PasstEntityRoleZuKategorie(ByVal ctx As Object, _
                                             ByVal category As String, _
                                             ByVal einAus As String) As Boolean
    
    Dim catLower As String
    catLower = LCase(category)
    Dim role As String
    role = ctx("EntityRole")
    
    PasstEntityRoleZuKategorie = True
    
    If role = "" Then Exit Function
    
    ' --- VERSORGER: Nur Versorger-typische Kategorien ---
    If ctx("IsVersorger") Then
        ' Versorger darf KEINE Mitglieder-Kategorien
        If catLower Like "*mitglied*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*pacht*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*endabrechnung*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*vorauszahlung*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*spende*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*beitrag*" And Not catLower Like "*verband*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*sammelzahlung*" Then PasstEntityRoleZuKategorie = False: Exit Function
        
        ' Versorger bei Einnahme (Rückzahlung VOM Versorger) = OK
        ' Versorger bei Ausgabe (Zahlung AN Versorger) = OK
    End If
    
    ' --- MITGLIED: Nur Mitglieder-typische Kategorien ---
    If ctx("IsMitglied") Then
        ' Mitglied darf KEINE Versorger-Kategorien
        If catLower Like "*versorger*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*stadtwerke*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*energieversorger*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*wasserwerk*" Then PasstEntityRoleZuKategorie = False: Exit Function
        
        ' Mitglied darf KEINE Rückzahlung-Versorger-Kombinationen
        ' (z.B. "Strom Rückzahlung Versorger" ist NICHT für Mitglieder)
        If catLower Like "*rueckzahlung*versorger*" Or _
           catLower Like "*rückzahlung*versorger*" Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
        
        ' Mitglied darf KEINE Bank-Kategorien
        If catLower Like "*entgeltabschluss*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*kontoführung*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*kontofuehrung*" Then PasstEntityRoleZuKategorie = False: Exit Function
        
        ' Mitglied bei Ausgabe = Rückerstattung AN Mitglied (selten, aber möglich)
        If ctx("IsAusgabe") Then
            ' Ausgaben an Mitglieder sind ungewöhnlich - nur Rückerstattung erlauben
            If Not (catLower Like "*rück*" Or catLower Like "*rueck*" Or _
                    catLower Like "*erstattung*" Or catLower Like "*gutschrift*") Then
                PasstEntityRoleZuKategorie = False: Exit Function
            End If
        End If
    End If
    
    ' --- BANK: Nur Bank-typische Kategorien ---
    If ctx("IsBank") Then
        If Not (catLower Like "*bank*" Or _
                catLower Like "*entgelt*" Or _
                catLower Like "*gebühr*" Or catLower Like "*gebuehr*" Or _
                catLower Like "*kontoführung*" Or catLower Like "*kontofuehrung*" Or _
                catLower Like "*zins*") Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
    End If
    
    ' --- EHEMALIGES MITGLIED: Wie Mitglied, aber eingeschränkter ---
    If role = "EHEMALIGES MITGLIED" Then
        If catLower Like "*versorger*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*stadtwerke*" Then PasstEntityRoleZuKategorie = False: Exit Function
        ' Ehemalige können noch Endabrechnungen/Rückzahlungen haben
    End If
    
End Function


' =====================================================
' Betragsvalidierung über Einstellungen!
' Prüft ob der Betrag zum Soll-Betrag der Kategorie passt
' Rückgabe: Score-Bonus (0 = kein Match, 25 = exakter Match)
' =====================================================
Private Function PruefeBetragGegenEinstellungen(ByVal category As String, _
                                                 ByVal absBetrag As Double) As Long
    Dim wsES As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim sollBetrag As Double
    
    On Error Resume Next
    Set wsES = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    If wsES Is Nothing Then
        PruefeBetragGegenEinstellungen = 0
        Exit Function
    End If
    
    lastRow = wsES.Cells(wsES.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    If lastRow < ES_START_ROW Then
        PruefeBetragGegenEinstellungen = 0
        Exit Function
    End If
    
    For r = ES_START_ROW To lastRow
        If Trim(wsES.Cells(r, ES_COL_KATEGORIE).value) = category Then
            sollBetrag = Val(wsES.Cells(r, ES_COL_SOLL_BETRAG).value)
            If sollBetrag > 0 Then
                ' Exakter Betragsvergleich (±1 Cent Toleranz)
                If Abs(absBetrag - sollBetrag) <= 0.01 Then
                    PruefeBetragGegenEinstellungen = 25
                    Exit Function
                End If
                ' Vielfaches des Soll-Betrags (Sammelzahlung)?
                If sollBetrag > 0 And absBetrag > sollBetrag Then
                    Dim rest As Double
                    rest = absBetrag - (Int(absBetrag / sollBetrag) * sollBetrag)
                    If Abs(rest) <= 0.01 Then
                        PruefeBetragGegenEinstellungen = 15
                        Exit Function
                    End If
                End If
            End If
            ' Kategorie gefunden aber Betrag passt nicht
            PruefeBetragGegenEinstellungen = 0
            Exit Function
        End If
    Next r
    
    ' Kategorie nicht in Einstellungen vorhanden (kein Malus)
    PruefeBetragGegenEinstellungen = 0
End Function


' =====================================================
' Zeitfensterprüfung über Einstellungen!
' Prüft ob das Buchungsdatum im Toleranzfenster liegt
' Rückgabe: Score-Bonus (0 = außerhalb, 20 = innerhalb)
' =====================================================
Private Function PruefeZeitfenster(ByVal category As String, _
                                    ByVal buchungsDatum As Date, _
                                    ByVal faelligkeit As String) As Long
    Dim wsES As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    On Error Resume Next
    Set wsES = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    If wsES Is Nothing Then
        PruefeZeitfenster = 0
        Exit Function
    End If
    
    lastRow = wsES.Cells(wsES.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    If lastRow < ES_START_ROW Then
        PruefeZeitfenster = 0
        Exit Function
    End If
    
    For r = ES_START_ROW To lastRow
        If Trim(wsES.Cells(r, ES_COL_KATEGORIE).value) = category Then
            
            Dim sollTag As Long
            Dim vorlauf As Long
            Dim nachlauf As Long
            Dim stichtagFix As String
            
            sollTag = Val(wsES.Cells(r, ES_COL_SOLL_TAG).value)
            vorlauf = Val(wsES.Cells(r, ES_COL_VORLAUF).value)
            nachlauf = Val(wsES.Cells(r, ES_COL_NACHLAUF).value)
            stichtagFix = Trim(wsES.Cells(r, ES_COL_STICHTAG_FIX).value)
            
            ' Variante A: Fester Stichtag (z.B. "15.03" = 15. März)
            If stichtagFix <> "" Then
                Dim stichDatum As Date
                On Error Resume Next
                stichDatum = CDate(stichtagFix & "." & Year(buchungsDatum))
                If Err.Number <> 0 Then
                    Err.Clear
                    On Error GoTo 0
                    GoTo WeiterNaechsteZeile
                End If
                On Error GoTo 0
                
                If buchungsDatum >= (stichDatum - vorlauf) And _
                   buchungsDatum <= (stichDatum + nachlauf) Then
                    PruefeZeitfenster = 20
                    Exit Function
                End If
            End If
            
            ' Variante B: Monatlich wiederkehrend (Soll-Tag)
            If sollTag >= 1 And sollTag <= 31 Then
                Dim sollDatum As Date
                On Error Resume Next
                sollDatum = DateSerial(Year(buchungsDatum), Month(buchungsDatum), sollTag)
                If Err.Number <> 0 Then
                    ' Tag existiert nicht in diesem Monat (z.B. 31. Februar)
                    Err.Clear
                    sollDatum = DateSerial(Year(buchungsDatum), Month(buchungsDatum) + 1, 0)
                End If
                On Error GoTo 0
                
                If buchungsDatum >= (sollDatum - vorlauf) And _
                   buchungsDatum <= (sollDatum + nachlauf) Then
                    PruefeZeitfenster = 20
                    Exit Function
                End If
                
                ' Auch Vormonat prüfen (Zahlung am Ende des Vormonats für nächsten Monat)
                Dim sollDatumVormonat As Date
                On Error Resume Next
                sollDatumVormonat = DateSerial(Year(buchungsDatum), Month(buchungsDatum) - 1, sollTag)
                If Err.Number <> 0 Then
                    Err.Clear
                    sollDatumVormonat = DateSerial(Year(buchungsDatum), Month(buchungsDatum), 0)
                End If
                On Error GoTo 0
                
                If buchungsDatum >= (sollDatumVormonat - vorlauf) And _
                   buchungsDatum <= (sollDatumVormonat + nachlauf) Then
                    PruefeZeitfenster = 15  ' Etwas weniger Score für Vormonat-Fenster
                    Exit Function
                End If
            End If
            
WeiterNaechsteZeile:
        End If
    Next r
    
    PruefeZeitfenster = 0
End Function


' =====================================================
' Monat/Periode intelligent ermitteln
' Berücksichtigt Fälligkeit und Zahlungstermine
' =====================================================
Public Function ErmittleMonatPeriode(ByVal category As String, _
                                     ByVal buchungsDatum As Date, _
                                     ByVal faelligkeit As String) As String
    
    Dim wsES As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    ' Fallback: Buchungsmonat
    Dim monatBuchung As Long
    monatBuchung = Month(buchungsDatum)
    
    ' Fälligkeit prüfen
    If faelligkeit = "" Then faelligkeit = "monatlich"
    
    Select Case LCase(faelligkeit)
        Case "jährlich", "jaehrlich"
            ErmittleMonatPeriode = "Jahresbeitrag " & Year(buchungsDatum)
            Exit Function
        Case "einmalig"
            ErmittleMonatPeriode = MonthName(monatBuchung) & " (einmalig)"
            Exit Function
        Case "quartalsweise", "quartal"
            Dim quartal As Long
            quartal = Int((monatBuchung - 1) / 3) + 1
            ErmittleMonatPeriode = "Q" & quartal & " " & Year(buchungsDatum)
            Exit Function
    End Select
    
    ' Monatliche Fälligkeit: Prüfen ob Zahlung für Folgemonat
    On Error Resume Next
    Set wsES = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    
    If wsES Is Nothing Then
        ErmittleMonatPeriode = MonthName(monatBuchung)
        Exit Function
    End If
    
    lastRow = wsES.Cells(wsES.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    
    For r = ES_START_ROW To lastRow
        If Trim(wsES.Cells(r, ES_COL_KATEGORIE).value) = category Then
            Dim sollTag As Long
            Dim vorlauf As Long
            
            sollTag = Val(wsES.Cells(r, ES_COL_SOLL_TAG).value)
            vorlauf = Val(wsES.Cells(r, ES_COL_VORLAUF).value)
            
            If sollTag >= 1 And sollTag <= 31 Then
                Dim tagBuchung As Long
                tagBuchung = Day(buchungsDatum)
                
                ' Wenn Buchungstag > Soll-Tag (z.B. 26. für Soll-Tag 5.)
                ' UND Vorlauf erlaubt es, dann ist es für den Folgemonat
                If tagBuchung > sollTag And vorlauf > 0 Then
                    Dim differenzTage As Long
                    ' Tage bis zum Soll-Tag im Folgemonat berechnen
                    Dim sollDatumFolge As Date
                    On Error Resume Next
                    sollDatumFolge = DateSerial(Year(buchungsDatum), Month(buchungsDatum) + 1, sollTag)
                    If Err.Number <> 0 Then
                        Err.Clear
                        On Error GoTo 0
                        GoTo FallbackMonat
                    End If
                    On Error GoTo 0
                    
                    differenzTage = CLng(sollDatumFolge - buchungsDatum)
                    
                    If differenzTage >= 0 And differenzTage <= vorlauf Then
                        ' Zahlung ist Vorlauf für Folgemonat
                        ErmittleMonatPeriode = MonthName(Month(sollDatumFolge))
                        Exit Function
                    End If
                End If
            End If
            
            ' Kategorie gefunden, aber keine Folgemonat-Zuordnung nötig
            GoTo FallbackMonat
        End If
    Next r
    
FallbackMonat:
    ErmittleMonatPeriode = MonthName(monatBuchung)
End Function


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


