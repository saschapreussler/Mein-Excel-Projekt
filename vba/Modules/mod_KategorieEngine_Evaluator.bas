Attribute VB_Name = "mod_KategorieEngine_Evaluator"
Option Explicit

' =====================================================
' KATEGORIE-ENGINE - EVALUATOR
' VERSION: 7.1 DEBUG - 08.02.2026
' NEU: Multi-Word-Matching (MatchKeyword)
' NEU: DEBUG-Protokoll im Direktbereich (Strg+G)
' HINWEIS: Debug.Print Zeilen nach Analyse entfernen!
' =====================================================

' Mindest-Score-Differenz fuer sichere Zuordnung
Private Const SCORE_DOMINANZ_SCHWELLE As Long = 20

' Kategorie fuer echte Mehrdeutigkeit (nur programmatisch!)
Private Const KAT_SAMMELZAHLUNG As String = "Sammelzahlung (mehrere Positionen) Mitglied"

' DEBUG: Zaehler fuer Zeilennummern im Protokoll
Private mDebugRowBK As Long

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
    
    ctx("IsMitglied") = (entityRole = "MITGLIED" Or _
                          entityRole = "MITGLIED_MIT_PACHT" Or _
                          entityRole = "MITGLIED_OHNE_PACHT")
    
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

' -----------------------------
' EntityRole ueber IBAN bestimmen
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
' Parzelle ueber IBAN bestimmen
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

' =====================================================
' MULTI-WORD-MATCHING (v7.0)
' Prueft ob ALLE Woerter des Keywords im Text vorkommen.
' Reihenfolge ist egal. Zusammengeschriebene Woerter
' werden ebenfalls erkannt (Substring-Matching je Wort).
' =====================================================
Private Function MatchKeyword(ByVal normText As String, _
                               ByVal normKeyword As String) As Boolean
    
    ' Schnelltest: Wenn Keyword keine Leerzeichen hat -> einfaches InStr
    If InStr(normKeyword, " ") = 0 Then
        MatchKeyword = (InStr(normText, normKeyword) > 0)
        Exit Function
    End If
    
    ' Multi-Word: Alle Woerter muessen als Substring vorkommen
    Dim woerter() As String
    woerter = Split(normKeyword, " ")
    
    Dim w As Long
    For w = LBound(woerter) To UBound(woerter)
        If Len(woerter(w)) > 0 Then
            If InStr(normText, woerter(w)) = 0 Then
                MatchKeyword = False
                Exit Function
            End If
        End If
    Next w
    
    MatchKeyword = True
End Function

' =====================================================
' DEBUG: Detailliertes Match-Protokoll fuer eine Zeile
' Wird im Direktbereich ausgegeben (Strg+G im VBA-Editor)
' =====================================================
Private Sub DebugLogMatch(ByVal keyword As String, ByVal normKeyword As String, _
                           ByVal category As String, ByVal einAus As String, _
                           ByVal prio As Long, ByVal matched As Boolean, _
                           ByVal filterGrund As String, ByVal score As Long)
    
    If matched Then
        Debug.Print "    [MATCH] Kat=""" & category & """ KW=""" & keyword & """ normKW=""" & normKeyword & """ Prio=" & prio & " Score=" & score
    ElseIf filterGrund <> "" Then
        Debug.Print "    [SKIP ] Kat=""" & category & """ KW=""" & keyword & """ -> " & filterGrund
    End If
End Sub

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
    
    ' === DEBUG START ===
    mDebugRowBK = rowBK
    Debug.Print ""
    Debug.Print "========== ZEILE " & rowBK & " =========="
    Debug.Print "  Name:    " & ctx("KontoName")
    Debug.Print "  Betrag:  " & ctx("Amount")
    Debug.Print "  E/A:     " & IIf(ctx("IsEinnahme"), "EINNAHME", IIf(ctx("IsAusgabe"), "AUSGABE", "NULL"))
    Debug.Print "  Role:    " & ctx("EntityRole")
    Debug.Print "  IsMitgl: " & ctx("IsMitglied")
    Debug.Print "  NormText:" & Left(ctx("NormText"), 120)
    Debug.Print "  BuchTxt: " & ctx("BuchungsText")
    Debug.Print "  --- Regel-Pruefung ---"
    ' === DEBUG END ===

    ' ================================
    ' PHASE 0: SONDERREGEL FUER 0-EURO-BETRAEGE
    ' ================================
    If ctx("IsNullBetrag") And ctx("IsEntgeltabschluss") Then
        Debug.Print "  -> PHASE 0: 0-Euro Entgeltabschluss"
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       "Entgeltabschluss (Kontof" & ChrW(252) & "hrung)", "GRUEN"
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = "0-Euro-Abschluss automatisch zugeordnet"
        Exit Sub
    End If

    ' ================================
    ' PHASE 1: HARTE SONDERREGELN
    ' ================================
    
    If ctx("IsEntgeltabschluss") And ctx("IsAusgabe") Then
        Debug.Print "  -> PHASE 1a: Entgeltabschluss"
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       "Entgeltabschluss (Kontof" & ChrW(252) & "hrung)", "GRUEN"
        Exit Sub
    End If
    
    If ctx("IsBargeldauszahlung") And ctx("IsAusgabe") Then
        Debug.Print "  -> PHASE 1b: Bargeldauszahlung"
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

    Dim hitCategories As Object
    Set hitCategories = CreateObject("Scripting.Dictionary")

    Dim ruleRow As Range
    For Each ruleRow In rngRules.Rows

        Dim category As String
        Dim einAus As String
        Dim keyword As String
        Dim prio As Long
        Dim faelligkeit As String

        category = Trim(ruleRow.Cells(1, 1).value)
        einAus = UCase(Trim(ruleRow.Cells(1, 2).value))
        keyword = Trim(ruleRow.Cells(1, 3).value)
        prio = Val(ruleRow.Cells(1, 4).value)
        faelligkeit = LCase(Trim(ruleRow.Cells(1, 6).value))
        If prio = 0 Then prio = 5

        If category = "" Or keyword = "" Then GoTo NextRule

        ' FILTER 0: Sammelzahlung NIEMALS per Keyword
        If LCase(category) Like "*sammelzahlung*" Then
            ' Kein Debug-Log fuer Sammelzahlung (wuerde nur fluten)
            GoTo NextRule
        End If

        ' FILTER 1: Einnahme/Ausgabe MUSS passen
        If Not ctx("IsNullBetrag") Then
            If einAus = "E" And ctx("IsAusgabe") Then
                DebugLogMatch keyword, "", category, einAus, prio, False, "E/A-Filter (E vs Ausgabe)", 0
                GoTo NextRule
            End If
            If einAus = "A" And ctx("IsEinnahme") Then
                DebugLogMatch keyword, "", category, einAus, prio, False, "E/A-Filter (A vs Einnahme)", 0
                GoTo NextRule
            End If
        End If

        ' FILTER 2: Strenge EntityRole-Trennung
        If Not PasstEntityRoleZuKategorie(ctx, category, einAus) Then
            DebugLogMatch keyword, "", category, einAus, prio, False, "EntityRole-Filter (Role=" & ctx("EntityRole") & ")", 0
            GoTo NextRule
        End If

        ' KEYWORD-MATCHING (v7.0 Multi-Word)
        Dim normKeyword As String
        normKeyword = NormalizeText(keyword)
        
        Dim matched As Boolean
        matched = MatchKeyword(ctx("NormText"), normKeyword)
        
        If matched Then

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
            
            Dim kwLen As Long
            kwLen = Len(normKeyword)
            If kwLen >= 12 Then
                score = score + 20
            ElseIf kwLen >= 8 Then
                score = score + 12
            ElseIf kwLen >= 5 Then
                score = score + 5
            End If
            
            Dim betragBonus As Long
            betragBonus = PruefeBetragGegenEinstellungen(category, ctx("AbsAmount"))
            score = score + betragBonus
            
            If IsDate(ctx("Datum")) Then
                Dim zeitBonus As Long
                zeitBonus = PruefeZeitfenster(category, CDate(ctx("Datum")), faelligkeit)
                score = score + zeitBonus
            End If

            ' === DEBUG ===
            DebugLogMatch keyword, normKeyword, category, einAus, prio, True, "", score
            
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
        Else
            ' DEBUG: Nur fuer Strom-relevante Keywords loggen (sonst zu viel)
            If LCase(keyword) Like "*strom*" Or LCase(keyword) Like "*abschlag*" Or _
               LCase(keyword) Like "*vorauszahlung*" Or LCase(keyword) Like "*stvom*" Then
                Debug.Print "    [MISS ] Kat=""" & category & """ KW=""" & keyword & """ normKW=""" & normKeyword & """"
                ' DEBUG: Einzelwort-Analyse
                If InStr(normKeyword, " ") > 0 Then
                    Dim dbgWords() As String
                    dbgWords = Split(normKeyword, " ")
                    Dim dw As Long
                    For dw = LBound(dbgWords) To UBound(dbgWords)
                        If Len(dbgWords(dw)) > 0 Then
                            If InStr(ctx("NormText"), dbgWords(dw)) > 0 Then
                                Debug.Print "             Wort """ & dbgWords(dw) & """ -> GEFUNDEN"
                            Else
                                Debug.Print "             Wort """ & dbgWords(dw) & """ -> NICHT GEFUNDEN ***"
                            End If
                        End If
                    Next dw
                End If
            End If
        End If

NextRule:
    Next ruleRow

    ' ================================
    ' PHASE 3: ERGEBNIS AUSWERTEN
    ' ================================
    
    ' === DEBUG ===
    Debug.Print "  --- Ergebnis ---"
    Debug.Print "  Treffer-Kategorien: " & hitCategories.count
    Dim dbgKey As Variant
    For Each dbgKey In hitCategories.keys
        Debug.Print "    " & CStr(dbgKey) & " = " & CLng(hitCategories(dbgKey))
    Next dbgKey
    Debug.Print "  Best: """ & bestCategory & """ Score=" & bestScore
    
    If hitCategories.count > 1 Then
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
        
        Debug.Print "  Differenz: " & scoreDifferenz & " (Schwelle: " & SCORE_DOMINANZ_SCHWELLE & ")"
        
        If scoreDifferenz >= SCORE_DOMINANZ_SCHWELLE Then
            Debug.Print "  -> GRUEN (dominanter Treffer)"
            ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), bestCategory, "GRUEN"
            Exit Sub
        End If
        
        Debug.Print "  -> GELB (Sammelzahlung - Mehrdeutigkeit)"
        
        Dim bemerkung As String
        bemerkung = hitCategories.count & " Kategorien passen:" & vbLf
        
        Dim katNr As Long
        katNr = 0
        For Each katKey In hitCategories.keys
            katNr = katNr + 1
            bemerkung = bemerkung & katNr & ") " & CStr(katKey) & vbLf
        Next katKey
        
        bemerkung = bemerkung & vbLf & _
                    "Bitte Kategorie manuell w" & ChrW(228) & "hlen und Betr" & ChrW(228) & "ge in Spalten M-Z aufteilen!"
        
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = bemerkung
        
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       KAT_SAMMELZAHLUNG, "GELB"
        
        Call EntsperreBetragsspalten(wsBK, rowBK, ctx("IsEinnahme"))
        Exit Sub
    End If

    If bestCategory <> "" Then
        Debug.Print "  -> GRUEN (einziger Treffer)"
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), bestCategory, "GRUEN"
        Exit Sub
    End If

    ' Kein Treffer = ROT
    Debug.Print "  -> ROT (kein Treffer)"
    If ctx("EntityRole") = "" Then
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            "Keine Kategorie gefunden. IBAN nicht zugeordnet - bitte Entity-Mapping pr" & ChrW(252) & "fen!"
    Else
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            "Keine passende Kategorie gefunden (EntityRole: " & ctx("EntityRole") & ")"
    End If
    ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), "Bitte Auswahl treffen!", "ROT"

End Sub


' =====================================================
' Betragsspalten entsperren fuer manuelle Eingabe
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



' =====================================================
' FILTER: Strenge EntityRole-Kategorie-Trennung
' v6.0: EHEMALIGES MITGLIED darf Auszahlung/Guthaben
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
        If catLower Like "*mitglied*" Then PasstEntityRoleZuKategorie = False: Exit Function
        
        If catLower Like "*pacht*" And catLower Like "*mitglied*" Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
        
        If catLower Like "*endabrechnung*" And catLower Like "*mitglied*" Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
        If catLower Like "*vorauszahlung*" And catLower Like "*mitglied*" Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
        If catLower Like "*spende*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*beitrag*" And Not catLower Like "*verband*" Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
        If catLower Like "*sammelzahlung*" Then PasstEntityRoleZuKategorie = False: Exit Function
    End If
    
    ' --- MITGLIED: Nur Mitglieder-typische Kategorien ---
    If ctx("IsMitglied") Then
        If catLower Like "*versorger*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*stadtwerke*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*energieversorger*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*wasserwerk*" Then PasstEntityRoleZuKategorie = False: Exit Function
        
        If catLower Like "*rueckzahlung*versorger*" Or _
           catLower Like "*r" & ChrW(252) & "ckzahlung*versorger*" Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
        
        If catLower Like "*miete*" And (catLower Like "*grundst" & ChrW(252) & "ck*" Or catLower Like "*grundstueck*") Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
        
        If catLower Like "*entgeltabschluss*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*kontof" & ChrW(252) & "hrung*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*kontofuehrung*" Then PasstEntityRoleZuKategorie = False: Exit Function
        
        If ctx("IsAusgabe") Then
            If Not (catLower Like "*r" & ChrW(252) & "ck*" Or catLower Like "*rueck*" Or _
                    catLower Like "*erstattung*" Or catLower Like "*gutschrift*" Or _
                    catLower Like "*auszahlung*" Or catLower Like "*guthaben*") Then
                PasstEntityRoleZuKategorie = False: Exit Function
            End If
        End If
    End If
    
    ' --- BANK: Nur Bank-typische Kategorien ---
    If ctx("IsBank") Then
        If Not (catLower Like "*bank*" Or _
                catLower Like "*entgelt*" Or _
                catLower Like "*geb" & ChrW(252) & "hr*" Or catLower Like "*gebuehr*" Or _
                catLower Like "*kontof" & ChrW(252) & "hrung*" Or catLower Like "*kontofuehrung*" Or _
                catLower Like "*zins*") Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
    End If
    
    ' --- EHEMALIGES MITGLIED ---
    If ctx("IsEhemaligesMitglied") Then
        If catLower Like "*versorger*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*stadtwerke*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*entgeltabschluss*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*kontof" & ChrW(252) & "hrung*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*kontofuehrung*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*miete*" And (catLower Like "*grundst" & ChrW(252) & "ck*" Or catLower Like "*grundstueck*") Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
        
        If ctx("IsAusgabe") Then
            If Not (catLower Like "*r" & ChrW(252) & "ck*" Or catLower Like "*rueck*" Or _
                    catLower Like "*erstattung*" Or catLower Like "*gutschrift*" Or _
                    catLower Like "*auszahlung*" Or catLower Like "*guthaben*" Or _
                    catLower Like "*endabrechnung*") Then
                PasstEntityRoleZuKategorie = False: Exit Function
            End If
        End If
    End If
    
End Function


' =====================================================
' Betragsvalidierung ueber Einstellungen!
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
                If Abs(absBetrag - sollBetrag) <= 0.01 Then
                    PruefeBetragGegenEinstellungen = 25
                    Exit Function
                End If
                If absBetrag > sollBetrag Then
                    Dim rest As Double
                    rest = absBetrag - (Int(absBetrag / sollBetrag) * sollBetrag)
                    If Abs(rest) <= 0.01 Then
                        PruefeBetragGegenEinstellungen = 15
                        Exit Function
                    End If
                End If
            End If
            PruefeBetragGegenEinstellungen = 0
            Exit Function
        End If
    Next r
    
    PruefeBetragGegenEinstellungen = 0
End Function


' =====================================================
' Zeitfensterpruefung ueber Einstellungen!
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
            
            If sollTag >= 1 And sollTag <= 31 Then
                Dim sollDatum As Date
                On Error Resume Next
                sollDatum = DateSerial(Year(buchungsDatum), Month(buchungsDatum), sollTag)
                If Err.Number <> 0 Then
                    Err.Clear
                    sollDatum = DateSerial(Year(buchungsDatum), Month(buchungsDatum) + 1, 0)
                End If
                On Error GoTo 0
                
                If buchungsDatum >= (sollDatum - vorlauf) And _
                   buchungsDatum <= (sollDatum + nachlauf) Then
                    PruefeZeitfenster = 20
                    Exit Function
                End If
                
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
                    PruefeZeitfenster = 15
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
' =====================================================
Public Function ErmittleMonatPeriode(ByVal category As String, _
                                     ByVal buchungsDatum As Date, _
                                     ByVal faelligkeit As String) As String
    
    Dim wsES As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim monatBuchung As Long
    monatBuchung = Month(buchungsDatum)
    
    If faelligkeit = "" Then faelligkeit = "monatlich"
    
    Select Case LCase(faelligkeit)
        Case "j" & ChrW(228) & "hrlich", "jaehrlich"
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
                
                If tagBuchung > sollTag And vorlauf > 0 Then
                    Dim differenzTage As Long
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
                        ErmittleMonatPeriode = MonthName(Month(sollDatumFolge))
                        Exit Function
                    End If
                End If
            End If
            
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

