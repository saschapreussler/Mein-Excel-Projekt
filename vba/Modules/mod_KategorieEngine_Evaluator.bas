Attribute VB_Name = "mod_KategorieEngine_Evaluator"
Option Explicit

' =====================================================
' KATEGORIE-ENGINE - EVALUATOR
' VERSION: 8.0 - 08.02.2026
' FIX: IsMitglied-Bug (Spaces statt Underscores)
' FIX: Kombiniertes GetEntityInfoByIBAN (1 Loop statt 2)
' NEU: ExactMatchBonus gegen Keyword-Kollisionen
' NEU: Einstellungen-Cache (Arrays statt Blattzugriff)
' ENTFERNT: Alle Debug.Print Zeilen
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
' Kontext erstellen (v8.0)
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
    
    ' FIX v8.0: Spaces statt Underscores! EntityKey Manager speichert
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
' ExactMatchBonus (v8.0)
' Gibt Bonuspunkte wenn das normalisierte Keyword als
' zusammenhaengender Substring im Text vorkommt.
' Verhindert Kollisionen bei Multi-Word-Matching.
' z.B. "stvom wasser" exakt in Text -> +10 Punkte
' =====================================================
Private Function ExactMatchBonus(ByVal normText As String, _
                                  ByVal normKeyword As String) As Long
    If InStr(normText, normKeyword) > 0 Then
        ExactMatchBonus = 10
    Else
        ExactMatchBonus = 0
    End If
End Function

' -----------------------------
' Hauptfunktion: Kategorie evaluieren (v8.0)
' -----------------------------
Public Sub EvaluateKategorieEngineRow(ByVal wsBK As Worksheet, _
                                      ByVal rowBK As Long, _
                                      ByVal rngRules As Range)

    ' --- Phase 1: Kontext ---
    Dim ctx As Object
    Set ctx = BuildKategorieContext(wsBK, rowBK)

    If ctx("IsNullBetrag") Then Exit Sub

    Dim normText As String
    normText = ctx("NormText")
    If normText = "" Then Exit Sub

    ' --- Phase 2: Regeln durchlaufen ---
    Dim bestKat As String
    Dim bestScore As Long
    Dim secondScore As Long
    Dim matchCount As Long
    bestKat = ""
    bestScore = 0
    secondScore = 0
    matchCount = 0

    Dim rRow As Long
    For rRow = 1 To rngRules.Rows.count
        Dim rawKeyword As String
        rawKeyword = Trim(CStr(rngRules.Cells(rRow, 1).value))
        If rawKeyword = "" Then GoTo NextRule

        Dim category As String
        category = Trim(CStr(rngRules.Cells(rRow, 2).value))
        If category = "" Then GoTo NextRule

        ' Keyword normalisieren
        Dim normKeyword As String
        normKeyword = NormalizeText(rawKeyword)
        If normKeyword = "" Then GoTo NextRule

        ' Keyword im Text suchen (Multi-Word-Matching)
        If Not MatchKeyword(normText, normKeyword) Then GoTo NextRule

        ' E/A aus Regeltabelle
        Dim einAus As String
        einAus = UCase(Trim(CStr(rngRules.Cells(rRow, 3).value)))

        ' E/A-Filter pruefen
        If einAus = "E" And Not ctx("IsEinnahme") Then GoTo NextRule
        If einAus = "A" And Not ctx("IsAusgabe") Then GoTo NextRule

        ' Prioritaet
        Dim prio As Long
        prio = 5
        On Error Resume Next
        prio = CLng(rngRules.Cells(rRow, 4).value)
        On Error GoTo 0
        If prio < 1 Then prio = 1
        If prio > 10 Then prio = 10

        ' EntityRole-Filter (Spalte 5)
        Dim roleFilter As String
        roleFilter = UCase(Trim(CStr(rngRules.Cells(rRow, 5).value)))
        If roleFilter <> "" Then
            If Not PasstEntityRoleZuKategorie(ctx, roleFilter) Then GoTo NextRule
        End If

        ' === SCORING ===
        Dim score As Long
        score = 100

        ' Prioritaetsbonus (Prio 1 = +45, Prio 5 = +25, Prio 10 = 0)
        score = score + (10 - prio) * 5

        ' EntityRole-Bonus (+20 wenn Rolle passt)
        If roleFilter <> "" Then score = score + 20

        ' E/A-Match-Bonus (+15)
        If einAus <> "" Then score = score + 15

        ' Keyword-Laengen-Bonus (laengere Keywords = spezifischer)
        If Len(normKeyword) > 20 Then
            score = score + 20
        ElseIf Len(normKeyword) > 10 Then
            score = score + 12
        ElseIf Len(normKeyword) > 5 Then
            score = score + 5
        End If

        ' ExactMatchBonus (v8.0: +10 wenn Keyword zusammenhaengend im Text)
        score = score + ExactMatchBonus(normText, normKeyword)

        ' Betrags-Bonus aus Einstellungen
        score = score + PruefeBetragGegenEinstellungen(category, ctx("AbsAmount"))

        ' Zeitfenster-Bonus
        score = score + PruefeZeitfenster(category, ctx("Datum"))

        ' Ergebnis vergleichen
        matchCount = matchCount + 1
        If score > bestScore Then
            secondScore = bestScore
            bestScore = score
            bestKat = category
        ElseIf score > secondScore Then
            secondScore = score
        End If

NextRule:
    Next rRow

    ' --- Phase 3: Ergebnis anwenden ---
    If matchCount = 0 Then
        ' ROT: Kein Match
        ApplyKategorie wsBK, rowBK, "", RGB(255, 199, 206), "Keine passende Kategorie gefunden"
        Exit Sub
    End If

    If matchCount = 1 Then
        ' GRUEN: Eindeutig
        ApplyKategorie wsBK, rowBK, bestKat, RGB(198, 239, 206), ""
        Exit Sub
    End If

    ' Mehrere Matches: Dominanz pruefen
    If (bestScore - secondScore) >= SCORE_DOMINANZ_SCHWELLE Then
        ' GRUEN: Klarer Sieger
        ApplyKategorie wsBK, rowBK, bestKat, RGB(198, 239, 206), ""
    Else
        ' GELB: Mehrdeutigkeit
        ApplyKategorie wsBK, rowBK, KAT_SAMMELZAHLUNG, RGB(255, 235, 156), _
            "Mehrere Kategorien moeglich (Diff=" & (bestScore - secondScore) & ")"
    End If

End Sub


' =====================================================
' EntityRole-Filter pruefen
' =====================================================
Private Function PasstEntityRoleZuKategorie(ByVal ctx As Object, _
                                             ByVal roleFilter As String) As Boolean
    PasstEntityRoleZuKategorie = False

    Select Case roleFilter
        Case "MITGLIED"
            PasstEntityRoleZuKategorie = ctx("IsMitglied")
        Case "VERSORGER"
            PasstEntityRoleZuKategorie = ctx("IsVersorger")
        Case "BANK"
            PasstEntityRoleZuKategorie = ctx("IsBank")
        Case "EHEMALIGES MITGLIED"
            PasstEntityRoleZuKategorie = ctx("IsEhemaligesMitglied")
        Case "ALLE"
            PasstEntityRoleZuKategorie = True
        Case Else
            ' Direktvergleich
            PasstEntityRoleZuKategorie = (ctx("EntityRole") = roleFilter)
    End Select
End Function


' =====================================================
' Betrags-Pruefung gegen Einstellungen (Cache-Version)
' =====================================================
Private Function PruefeBetragGegenEinstellungen(ByVal category As String, _
                                                 ByVal absAmount As Double) As Long
    PruefeBetragGegenEinstellungen = 0
    If Not mCacheGeladen Then Exit Function
    If mCacheAnzahl = 0 Then Exit Function
    
    Dim i As Long
    For i = 1 To mCacheAnzahl
        If StrComp(mCacheKat(i), category, vbTextCompare) = 0 Then
            Dim sollBetrag As Double
            sollBetrag = mCacheSoll(i)
            If sollBetrag = 0 Then Exit Function
            
            Dim diff As Double
            diff = Abs(absAmount - Abs(sollBetrag))
            
            If diff < 0.01 Then
                PruefeBetragGegenEinstellungen = 25
            ElseIf diff <= Abs(sollBetrag) * 0.15 Then
                PruefeBetragGegenEinstellungen = 15
            End If
            Exit Function
        End If
    Next i
End Function


' =====================================================
' Zeitfenster-Pruefung (Cache-Version)
' Nutzt Stichtag + Vorlauf/Nachlauf aus Einstellungen
' =====================================================
Private Function PruefeZeitfenster(ByVal category As String, _
                                    ByVal buchungsDatum As Variant) As Long
    PruefeZeitfenster = 0
    If Not mCacheGeladen Then Exit Function
    If mCacheAnzahl = 0 Then Exit Function
    If Not IsDate(buchungsDatum) Then Exit Function
    
    Dim buchDat As Date
    buchDat = CDate(buchungsDatum)
    
    Dim i As Long
    For i = 1 To mCacheAnzahl
        If StrComp(mCacheKat(i), category, vbTextCompare) = 0 Then
            
            ' Stichtag-basierte Pruefung
            Dim sollTag As Long
            sollTag = mCacheSollTag(i)
            
            Dim vorlauf As Long
            Dim nachlauf As Long
            vorlauf = mCacheVorlauf(i)
            nachlauf = mCacheNachlauf(i)
            
            ' Fester Stichtag vorhanden?
            If IsDate(mCacheStichtag(i)) Then
                Dim stichtag As Date
                stichtag = CDate(mCacheStichtag(i))
                
                ' Stichtag auf gleiches Jahr wie Buchung setzen
                Dim stichtagAktuell As Date
                On Error Resume Next
                stichtagAktuell = DateSerial(Year(buchDat), Month(stichtag), Day(stichtag))
                If Err.Number <> 0 Then
                    Err.Clear
                    On Error GoTo 0
                    Exit Function
                End If
                On Error GoTo 0
                
                Dim diffTage As Long
                diffTage = Abs(CLng(buchDat - stichtagAktuell))
                
                If diffTage = 0 Then
                    PruefeZeitfenster = 20   ' Exakt am Stichtag
                    Exit Function
                End If
                
                ' Vorlauf/Nachlauf-Toleranz
                If vorlauf > 0 Or nachlauf > 0 Then
                    Dim fruehestens As Date
                    Dim spaetestens As Date
                    fruehestens = stichtagAktuell - vorlauf
                    spaetestens = stichtagAktuell + nachlauf
                    
                    If buchDat >= fruehestens And buchDat <= spaetestens Then
                        PruefeZeitfenster = 15   ' Im Toleranzfenster
                    End If
                End If
                
                Exit Function
            End If
            
            ' Soll-Tag (Tag im Monat, z.B. 15 = jeweils am 15.)
            If sollTag > 0 And sollTag <= 31 Then
                Dim buchTag As Long
                buchTag = Day(buchDat)
                
                If buchTag = sollTag Then
                    PruefeZeitfenster = 20   ' Exakt am Soll-Tag
                    Exit Function
                End If
                
                ' Toleranz um den Soll-Tag
                If vorlauf > 0 Or nachlauf > 0 Then
                    If buchTag >= (sollTag - vorlauf) And buchTag <= (sollTag + nachlauf) Then
                        PruefeZeitfenster = 15
                    End If
                End If
                
                Exit Function
            End If
            
            ' Kein Zeitfenster definiert
            Exit Function
        End If
    Next i
End Function


' =====================================================
' Kategorie + Farbe + Bemerkung anwenden
' =====================================================
Private Sub ApplyKategorie(ByVal wsBK As Worksheet, ByVal rowBK As Long, _
                            ByVal kat As String, ByVal farbe As Long, _
                            ByVal bemerkung As String)

    With wsBK.Cells(rowBK, BK_COL_KATEGORIE)
        .value = kat
        .Interior.color = farbe
        
        ' Schriftfarbe
        If farbe = RGB(198, 239, 206) Then
            .Font.color = RGB(0, 97, 0)         ' Dunkelgruen
        ElseIf farbe = RGB(255, 235, 156) Then
            .Font.color = RGB(156, 101, 0)      ' Dunkelgelb
        ElseIf farbe = RGB(255, 199, 206) Then
            .Font.color = RGB(156, 0, 6)        ' Dunkelrot
        Else
            .Font.color = vbBlack
        End If
    End With

    If bemerkung <> "" Then
        wsBK.Cells(rowBK, BK_COL_BEMERKUNG).value = bemerkung
    End If
End Sub

