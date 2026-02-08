Attribute VB_Name = "mod_KategorieEngine_Evaluator"
Option Explicit

' =====================================================
' KATEGORIE-ENGINE - EVALUATOR
' VERSION: 9.1 - 09.02.2026
' MERGE: v7.0 Scoring-Logik (funktionierend) +
'        v8.2 Infrastruktur (kein Named Range, Cache,
'        kombiniertes GetEntityInfo, ExactMatchBonus)
' FIX: Sonderregel fuer 0-Euro wieder aktiv
' FIX: PasstEntityRoleZuKategorie mit detaillierter Logik
' FIX: Sammelzahlung-Filter vor Scoring wieder aktiv
' FIX: ApplyKategorie mit originaler Signatur
' FIX: EntsperreBetragsspalten bei GELB wieder aktiv
' FIX: Detaillierte Bemerkungen bei GELB und ROT
' FIX: Betrags-Vielfaches-Check wiederhergestellt
' FIX: Zeitfenster mit Vormonat + faelligkeit-Parameter
' FIX: EntityRole-Bonus wieder auf +20
' v9.1: GELB-Bemerkung bereinigt (kein Instruktions-Satz)
' v9.1: ErmittleMonatPeriode mit Folgemonat-Erkennung
'       via SollTag + Vorlauf aus Einstellungen-Cache
' =====================================================

' Mindest-Score-Differenz fuer sichere Zuordnung
Private Const SCORE_DOMINANZ_SCHWELLE As Long = 20

' Kategorie fuer echte Mehrdeutigkeit (nur programmatisch!)
Private Const KAT_SAMMELZAHLUNG As String = "Sammelzahlung (mehrere Positionen) Mitglied"

' =====================================================
' EINSTELLUNGEN-CACHE (Performance)
' Wird einmal geladen, dann fuer alle Zeilen verwendet
' Spalten: B=Kategorie, C=Soll-Betrag, D=Soll-Tag,
'          E=Stichtag, F=Vorlauf, G=Nachlauf
' =====================================================
Private mCacheGeladen As Boolean
Private mCacheKat() As String
Private mCacheSoll() As Double
Private mCacheSollTag() As Long
Private mCacheStichtag() As Variant
Private mCacheVorlauf() As Long
Private mCacheNachlauf() As Long
Private mCacheAnzahl As Long

Public Sub LadeEinstellungenCache()
    Dim wsES As Worksheet
    Set wsES = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    
    Dim lastRow As Long
    lastRow = wsES.Cells(wsES.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    
    If lastRow < ES_START_ROW Then
        mCacheAnzahl = 0
        mCacheGeladen = True
        Exit Sub
    End If
    
    mCacheAnzahl = lastRow - ES_START_ROW + 1
    ReDim mCacheKat(1 To mCacheAnzahl)
    ReDim mCacheSoll(1 To mCacheAnzahl)
    ReDim mCacheSollTag(1 To mCacheAnzahl)
    ReDim mCacheStichtag(1 To mCacheAnzahl)
    ReDim mCacheVorlauf(1 To mCacheAnzahl)
    ReDim mCacheNachlauf(1 To mCacheAnzahl)
    
    Dim i As Long
    Dim r As Long
    For i = 1 To mCacheAnzahl
        r = ES_START_ROW + i - 1
        mCacheKat(i) = Trim(CStr(wsES.Cells(r, ES_COL_KATEGORIE).value))
        mCacheSoll(i) = wsES.Cells(r, ES_COL_SOLL_BETRAG).value
        
        On Error Resume Next
        mCacheSollTag(i) = CLng(wsES.Cells(r, ES_COL_SOLL_TAG).value)
        If Err.Number <> 0 Then mCacheSollTag(i) = 0: Err.Clear
        On Error GoTo 0
        
        mCacheStichtag(i) = wsES.Cells(r, ES_COL_STICHTAG_FIX).value
        
        On Error Resume Next
        mCacheVorlauf(i) = CLng(wsES.Cells(r, ES_COL_VORLAUF).value)
        If Err.Number <> 0 Then mCacheVorlauf(i) = 0: Err.Clear
        mCacheNachlauf(i) = CLng(wsES.Cells(r, ES_COL_NACHLAUF).value)
        If Err.Number <> 0 Then mCacheNachlauf(i) = 0: Err.Clear
        On Error GoTo 0
    Next i
    
    mCacheGeladen = True
End Sub

Public Sub EntladeEinstellungenCache()
    mCacheGeladen = False
    mCacheAnzahl = 0
    Erase mCacheKat
    Erase mCacheSoll
    Erase mCacheSollTag
    Erase mCacheStichtag
    Erase mCacheVorlauf
    Erase mCacheNachlauf
End Sub

' -----------------------------
' EntityInfo ueber IBAN bestimmen (kombiniert: Role + Parzelle)
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
' MULTI-WORD-MATCHING (v7.0)
' Prueft ob ALLE Woerter des Keywords im Text vorkommen.
' Reihenfolge ist egal. Zusammengeschriebene Woerter
' werden ebenfalls erkannt (Substring-Matching je Wort).
' =====================================================
Private Function MatchKeyword(ByVal normText As String, _
                               ByVal normKeyword As String) As Boolean
    
    If InStr(normKeyword, " ") = 0 Then
        MatchKeyword = (InStr(normText, normKeyword) > 0)
        Exit Function
    End If
    
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
' ExactMatchBonus (v8.0)
' Gibt Bonuspunkte wenn das normalisierte Keyword als
' zusammenhaengender Substring im Text vorkommt.
' =====================================================
Private Function ExactMatchBonus(ByVal normText As String, _
                                  ByVal normKeyword As String) As Long
    If InStr(normText, normKeyword) > 0 Then
        ExactMatchBonus = 10
    Else
        ExactMatchBonus = 0
    End If
End Function

' =====================================================
' Hauptfunktion: Kategorie evaluieren (v9.0)
' Braucht KEINEN Named Range! Liest Regeln direkt vom
' Daten-Blatt ueber DATA_CAT_COL_* Konstanten.
' Scoring-Logik aus v7.0 wiederhergestellt.
' =====================================================
Public Sub EvaluateKategorieEngineRow(ByVal wsBK As Worksheet, _
                                      ByVal rowBK As Long, _
                                      ByVal wsData As Worksheet, _
                                      ByVal lastRuleRow As Long)

    ' Bereits kategorisiert? Ueberspringen
    If Trim(wsBK.Cells(rowBK, BK_COL_KATEGORIE).value) <> "" Then Exit Sub

    Dim ctx As Object
    Set ctx = BuildKategorieContext(wsBK, rowBK)

    ' ================================
    ' PHASE 0: SONDERREGEL FUER 0-EURO-BETRAEGE
    ' ================================
    If ctx("IsNullBetrag") And ctx("IsEntgeltabschluss") Then
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       "Entgeltabschluss (Kontof" & ChrW(252) & "hrung)", "GRUEN"
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = "0-Euro-Abschluss automatisch zugeordnet"
        Exit Sub
    End If

    ' 0-Euro ohne Sonderregel -> ueberspringen
    If ctx("IsNullBetrag") Then Exit Sub

    Dim normText As String
    normText = ctx("NormText")
    If normText = "" Then Exit Sub

    ' ================================
    ' PHASE 1: HARTE SONDERREGELN
    ' ================================
    
    ' 1a) Entgeltabschluss (Bankgebuehren)
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

    ' Dictionary: Kategorie -> Score (hoechster Score je Kategorie)
    Dim hitCategories As Object
    Set hitCategories = CreateObject("Scripting.Dictionary")

    Dim dataRow As Long
    For dataRow = DATA_START_ROW To lastRuleRow

        Dim category As String
        Dim einAus As String
        Dim keyword As String
        Dim prio As Long
        Dim faelligkeit As String

        ' Spalten ueber Konstanten lesen
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
            
            ' Prioritaetsbonus (niedrigere Prio = hoeherer Bonus)
            score = score + (10 - prio) * 5
            
            ' EntityRole bekannt = hoehere Konfidenz (+20 wie in v7.0)
            If ctx("EntityRole") <> "" Then
                score = score + 20
            End If
            
            ' Einnahme/Ausgabe stimmt exakt ueberein
            If (einAus = "E" And ctx("IsEinnahme")) Or _
               (einAus = "A" And ctx("IsAusgabe")) Then
                score = score + 15
            End If
            
            ' Keyword-Laenge als Qualitaetsfaktor
            Dim kwLen As Long
            kwLen = Len(normKeyword)
            If kwLen >= 12 Then
                score = score + 20
            ElseIf kwLen >= 8 Then
                score = score + 12
            ElseIf kwLen >= 5 Then
                score = score + 5
            End If
            
            ' ExactMatchBonus (v8.0: +10 wenn Keyword zusammenhaengend im Text)
            score = score + ExactMatchBonus(normText, normKeyword)
            
            ' Betragsvalidierung ueber Einstellungen
            Dim betragBonus As Long
            betragBonus = PruefeBetragGegenEinstellungen(category, ctx("AbsAmount"))
            score = score + betragBonus
            
            ' Zeitfenstervalidierung ueber Einstellungen
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
                ' Letzte Kategorie: KEIN abschliessendes vbLf
                bemerkung = bemerkung & katNr & ") " & CStr(katKey)
            End If
        Next katKey
        
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = bemerkung
        
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), _
                       KAT_SAMMELZAHLUNG, "GELB"
        
        Call EntsperreBetragsspalten(wsBK, rowBK, ctx("IsEinnahme"))
        Exit Sub
    End If

    ' Genau 1 Treffer = sicher GRUEN
    If bestCategory <> "" Then
        ApplyKategorie wsBK.Cells(rowBK, BK_COL_KATEGORIE), bestCategory, "GRUEN"
        Exit Sub
    End If

    ' Kein Treffer = ROT
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
' FILTER: Strenge EntityRole-Kategorie-Trennung (v7.0)
' Detaillierte Logik wiederhergestellt!
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
        
        ' Mitglied bei Ausgabe = nur Rueckerstattung/Auszahlung/Guthaben
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
        
        ' Ehemalige bei Ausgabe: Auszahlung/Guthaben/Rueckzahlung erlaubt
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
' Betragsvalidierung ueber Einstellungen (Cache-Version)
' mit Vielfaches-Check aus v7.0
' =====================================================
Private Function PruefeBetragGegenEinstellungen(ByVal category As String, _
                                                 ByVal absBetrag As Double) As Long
    PruefeBetragGegenEinstellungen = 0
    If Not mCacheGeladen Then Exit Function
    If mCacheAnzahl = 0 Then Exit Function
    
    Dim i As Long
    For i = 1 To mCacheAnzahl
        If StrComp(mCacheKat(i), category, vbTextCompare) = 0 Then
            Dim sollBetrag As Double
            sollBetrag = Abs(mCacheSoll(i))
            If sollBetrag = 0 Then Exit Function
            
            ' Exakter Treffer
            If Abs(absBetrag - sollBetrag) <= 0.01 Then
                PruefeBetragGegenEinstellungen = 25
                Exit Function
            End If
            
            ' Vielfaches-Check (z.B. 3x Monatsbeitrag)
            If absBetrag > sollBetrag Then
                Dim rest As Double
                rest = absBetrag - (Int(absBetrag / sollBetrag) * sollBetrag)
                If Abs(rest) <= 0.01 Then
                    PruefeBetragGegenEinstellungen = 15
                    Exit Function
                End If
            End If
            
            Exit Function
        End If
    Next i
End Function


' =====================================================
' Zeitfensterpruefung (Cache-Version + faelligkeit)
' mit Vormonat-Check aus v7.0
' =====================================================
Private Function PruefeZeitfenster(ByVal category As String, _
                                    ByVal buchungsDatum As Date, _
                                    ByVal faelligkeit As String) As Long
    PruefeZeitfenster = 0
    If Not mCacheGeladen Then Exit Function
    If mCacheAnzahl = 0 Then Exit Function
    
    Dim i As Long
    For i = 1 To mCacheAnzahl
        If StrComp(mCacheKat(i), category, vbTextCompare) = 0 Then
            
            Dim sollTag As Long
            Dim vorlauf As Long
            Dim nachlauf As Long
            
            sollTag = mCacheSollTag(i)
            vorlauf = mCacheVorlauf(i)
            nachlauf = mCacheNachlauf(i)
            
            ' Fester Stichtag vorhanden?
            If IsDate(mCacheStichtag(i)) Then
                Dim stichDatum As Date
                On Error Resume Next
                stichDatum = CDate(CStr(mCacheStichtag(i)))
                If Err.Number <> 0 Then
                    Err.Clear
                    On Error GoTo 0
                    GoTo WeiterNaechsteZeile
                End If
                On Error GoTo 0
                
                ' Stichtag auf gleiches Jahr wie Buchung setzen
                Dim stichtagAktuell As Date
                On Error Resume Next
                stichtagAktuell = DateSerial(Year(buchungsDatum), Month(stichDatum), Day(stichDatum))
                If Err.Number <> 0 Then
                    Err.Clear
                    On Error GoTo 0
                    GoTo WeiterNaechsteZeile
                End If
                On Error GoTo 0
                
                If buchungsDatum >= (stichtagAktuell - vorlauf) And _
                   buchungsDatum <= (stichtagAktuell + nachlauf) Then
                    PruefeZeitfenster = 20
                    Exit Function
                End If
            End If
            
            ' Soll-Tag (Tag im Monat, z.B. 15 = jeweils am 15.)
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
                
                ' Vormonat-Check (aus v7.0)
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
    Next i
End Function


' =====================================================
' Monat/Periode intelligent ermitteln (v9.1)
' Nutzt Einstellungen-Cache fuer Folgemonat-Erkennung:
' Wenn SollTag und Vorlauf gesetzt sind, wird geprueft
' ob die Zahlung bereits fuer den Folgemonat gilt.
' Beispiel: SollTag=5, Vorlauf=10
'   -> Folgemonat-Faelligkeit = 05.02.
'   -> Fruehester Zahlungstag = 05.02. - 10 = 26.01.
'   -> Zahlung am 27.01. >= 26.01. -> gilt fuer Februar
' =====================================================
Public Function ErmittleMonatPeriode(ByVal category As String, _
                                     ByVal buchungsDatum As Date, _
                                     ByVal faelligkeit As String) As String
    
    Dim monatBuchung As Long
    monatBuchung = Month(buchungsDatum)
    
    If faelligkeit = "" Then faelligkeit = "monatlich"
    
    ' Nicht-monatliche Perioden: direkt zuordnen
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
        Case "halbjaehrlich", "halbj" & ChrW(228) & "hrlich"
            Dim halbjahr As Long
            halbjahr = IIf(monatBuchung <= 6, 1, 2)
            ErmittleMonatPeriode = "H" & halbjahr & " " & Year(buchungsDatum)
            Exit Function
    End Select
    
    ' ==============================================
    ' Monatlich: Folgemonat-Erkennung via Cache
    ' ==============================================
    If Not mCacheGeladen Then
        ErmittleMonatPeriode = MonthName(monatBuchung)
        Exit Function
    End If
    
    Dim idx As Long
    For idx = 1 To mCacheAnzahl
        If StrComp(mCacheKat(idx), category, vbTextCompare) = 0 Then
            
            Dim sollTag As Long
            Dim vorlauf As Long
            
            sollTag = mCacheSollTag(idx)
            vorlauf = mCacheVorlauf(idx)
            
            ' Pruefe zuerst festen Stichtag (Spalte E)
            If IsDate(mCacheStichtag(idx)) Then
                Dim stichDatum As Date
                On Error Resume Next
                stichDatum = CDate(CStr(mCacheStichtag(idx)))
                If Err.Number = 0 Then
                    On Error GoTo 0
                    ' Stichtag im aktuellen Buchungsmonat
                    Dim stichAktuell As Date
                    stichAktuell = DateSerial(Year(buchungsDatum), Month(buchungsDatum), Day(stichDatum))
                    
                    ' Stichtag im Folgemonat
                    Dim stichFolge As Date
                    stichFolge = DateSerial(Year(buchungsDatum), Month(buchungsDatum) + 1, Day(stichDatum))
                    
                    ' Liegt Buchung im Vorlauf-Fenster des Folgemonats?
                    If vorlauf > 0 And buchungsDatum >= (stichFolge - vorlauf) Then
                        ErmittleMonatPeriode = MonthName(Month(stichFolge))
                        Exit Function
                    End If
                    
                    ' Liegt Buchung im Vorlauf-Fenster des aktuellen Monats?
                    If vorlauf > 0 And buchungsDatum >= (stichAktuell - vorlauf) Then
                        ErmittleMonatPeriode = MonthName(Month(stichAktuell))
                        Exit Function
                    End If
                Else
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
            
            ' Pruefe SollTag (Spalte D) - Tag im Monat
            If sollTag >= 1 And sollTag <= 31 Then
                Dim tagBuchung As Long
                tagBuchung = Day(buchungsDatum)
                
                ' Folgemonat-Pruefung: Zahlung NACH dem SollTag des
                ' aktuellen Monats UND innerhalb des Vorlauf-Fensters
                ' fuer den SollTag des Folgemonats
                If vorlauf > 0 And tagBuchung > sollTag Then
                    Dim sollDatumFolge As Date
                    On Error Resume Next
                    sollDatumFolge = DateSerial(Year(buchungsDatum), Month(buchungsDatum) + 1, sollTag)
                    If Err.Number <> 0 Then
                        Err.Clear
                        On Error GoTo 0
                        GoTo FallbackMonat
                    End If
                    On Error GoTo 0
                    
                    Dim differenzTage As Long
                    differenzTage = CLng(sollDatumFolge - buchungsDatum)
                    
                    If differenzTage >= 0 And differenzTage <= vorlauf Then
                        ErmittleMonatPeriode = MonthName(Month(sollDatumFolge))
                        Exit Function
                    End If
                End If
            End If
            
            GoTo FallbackMonat
        End If
    Next idx
    
FallbackMonat:
    ErmittleMonatPeriode = MonthName(monatBuchung)
End Function


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


