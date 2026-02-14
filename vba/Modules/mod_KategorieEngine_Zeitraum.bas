Attribute VB_Name = "mod_KategorieEngine_Zeitraum"
Option Explicit

' =====================================================
' KATEGORIE-ENGINE - ZEITRAUM & EINSTELLUNGEN-CACHE
' Ausgelagert aus mod_KategorieEngine_Evaluator
' Enth�lt: Einstellungen-Cache, Betragsvalidierung,
'          Zeitfensterpr�fung, Periodenermittlung
' =====================================================


' =====================================================
' EINSTELLUNGEN-CACHE (Performance)
' Wird einmal geladen, dann f�r alle Zeilen verwendet
' Spalten: B=Kategorie, C=Soll-Betrag, D=Soll-Tag,
'          E=Soll-Monate, F=Stichtag, G=Vorlauf, H=Nachlauf
' =====================================================
Private mCacheGeladen As Boolean
Private mCacheKat() As String
Private mCacheSoll() As Double
Private mCacheSollTag() As Long
Private mCacheSollMonate() As String
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
    ReDim mCacheSollMonate(1 To mCacheAnzahl)
    ReDim mCacheStichtag(1 To mCacheAnzahl)
    ReDim mCacheVorlauf(1 To mCacheAnzahl)
    ReDim mCacheNachlauf(1 To mCacheAnzahl)
    
    Dim i As Long
    Dim r As Long
    For i = 1 To mCacheAnzahl
        r = ES_START_ROW + i - 1
        mCacheKat(i) = Trim(CStr(wsES.Cells(r, ES_COL_KATEGORIE).value))
        mCacheSoll(i) = wsES.Cells(r, ES_COL_SOLL_BETRAG).value
        
        ' Soll-Tag: Zahl (1-31)
        Dim tagWert As String
        tagWert = Trim(CStr(wsES.Cells(r, ES_COL_SOLL_TAG).value))
        On Error Resume Next
        mCacheSollTag(i) = CLng(tagWert)
        If Err.Number <> 0 Then mCacheSollTag(i) = 0: Err.Clear
        On Error GoTo 0
        
        ' Soll-Monate: "03, 06, 09" oder leer (= alle Monate)
        mCacheSollMonate(i) = Trim(CStr(wsES.Cells(r, ES_COL_SOLL_MONATE).value))
        
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
    Erase mCacheSollMonate
    Erase mCacheStichtag
    Erase mCacheVorlauf
    Erase mCacheNachlauf
End Sub


' =====================================================
' Betragsvalidierung �ber Einstellungen (Cache-Version)
' mit Vielfaches-Check aus v7.0
' =====================================================
Public Function PruefeBetragGegenEinstellungen(ByVal category As String, _
                                                ByVal absBetrag As Double) As Long
    PruefeBetragGegenEinstellungen = 0
    If Not mCacheGeladen Then Exit Function
    If mCacheAnzahl = 0 Then Exit Function
    
    Dim i As Long
    For i = 1 To mCacheAnzahl
        If StrComp(mCacheKat(i), category, vbTextCompare) = 0 Then
            Dim SollBetrag As Double
            SollBetrag = Abs(mCacheSoll(i))
            If SollBetrag = 0 Then Exit Function
            
            ' Exakter Treffer
            If Abs(absBetrag - SollBetrag) <= 0.01 Then
                PruefeBetragGegenEinstellungen = 25
                Exit Function
            End If
            
            ' Vielfaches-Check (z.B. 3x Monatsbeitrag)
            If absBetrag > SollBetrag Then
                Dim rest As Double
                rest = absBetrag - (Int(absBetrag / SollBetrag) * SollBetrag)
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
' Zeitfensterpr�fung (Cache-Version + F�lligkeit + Soll-Monate)
' Priorit�ten:
' 1. Spalte F (Stichtag Fix) -> exaktes Datum
' 2. Spalte D + E (Tag + Monate) -> kombiniert
' 3. Spalte D allein -> monatlich
' 4. Tag 31 + Monate -> letzter Tag im jeweiligen Monat
' =====================================================
Public Function PruefeZeitfenster(ByVal category As String, _
                                   ByVal buchungsDatum As Date, _
                                   ByVal faelligkeit As String) As Long
    PruefeZeitfenster = 0
    If Not mCacheGeladen Then Exit Function
    If mCacheAnzahl = 0 Then Exit Function
    
    Dim i As Long
    For i = 1 To mCacheAnzahl
        If StrComp(mCacheKat(i), category, vbTextCompare) = 0 Then
            
            Dim SollTag As Long
            Dim vorlauf As Long
            Dim nachlauf As Long
            Dim SollMonate As String
            
            SollTag = mCacheSollTag(i)
            vorlauf = mCacheVorlauf(i)
            nachlauf = mCacheNachlauf(i)
            SollMonate = mCacheSollMonate(i)
            
            ' ===== PRIO 1: Fester Stichtag (Spalte F) =====
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
                ' Stichtag hat Vorrang -> nicht weiter pr�fen
                GoTo WeiterNaechsteZeile
            End If
            
            ' ===== PRIO 2/3/4: Tag und/oder Monate =====
            Dim buchungsMonat As Long
            buchungsMonat = Month(buchungsDatum)
            
            ' Pr�fe ob der Buchungsmonat in den Soll-Monaten liegt
            Dim monatPasst As Boolean
            monatPasst = True  ' Default: alle Monate (Spalte E leer)
            
            If SollMonate <> "" Then
                monatPasst = IstMonatInListe(buchungsMonat, SollMonate)
            End If
            
            ' PRIO 4: Tag 31 + Monate = letzter Tag im jeweiligen Monat
            ' (z.B. 28. Feb, 30. Apr, 31. Jan usw.)
            If SollTag = 31 And SollMonate <> "" Then
                If monatPasst Then
                    Dim letzterTag As Date
                    letzterTag = DateSerial(Year(buchungsDatum), buchungsMonat + 1, 0)
                    
                    If buchungsDatum >= (letzterTag - vorlauf) And _
                       buchungsDatum <= (letzterTag + nachlauf) Then
                        PruefeZeitfenster = 20
                        Exit Function
                    End If
                End If
                GoTo WeiterNaechsteZeile
            End If
            
            ' PRIO 2/3: Soll-Tag vorhanden (1-31)
            If SollTag >= 1 And SollTag <= 31 Then
                If monatPasst Then
                    Dim sollDatum As Date
                    On Error Resume Next
                    sollDatum = DateSerial(Year(buchungsDatum), buchungsMonat, SollTag)
                    If Err.Number <> 0 Then
                        Err.Clear
                        sollDatum = DateSerial(Year(buchungsDatum), buchungsMonat + 1, 0)
                    End If
                    On Error GoTo 0
                    
                    If buchungsDatum >= (sollDatum - vorlauf) And _
                       buchungsDatum <= (sollDatum + nachlauf) Then
                        PruefeZeitfenster = 20
                        Exit Function
                    End If
                End If
                
                ' Vormonat-Check: Pr�fe auch ob Buchung im Vorlauf des n�chsten passenden Monats liegt
                If SollMonate <> "" Then
                    ' Pr�fe ob der Folgemonat in der Liste ist
                    Dim folgeMonat As Long
                    folgeMonat = buchungsMonat + 1
                    If folgeMonat > 12 Then folgeMonat = 1
                    
                    If IstMonatInListe(folgeMonat, SollMonate) Then
                        Dim sollDatumFolge As Date
                        On Error Resume Next
                        sollDatumFolge = DateSerial(Year(buchungsDatum), buchungsMonat + 1, SollTag)
                        If Err.Number <> 0 Then
                            Err.Clear
                            sollDatumFolge = DateSerial(Year(buchungsDatum), buchungsMonat + 2, 0)
                        End If
                        On Error GoTo 0
                        
                        If buchungsDatum >= (sollDatumFolge - vorlauf) And _
                           buchungsDatum <= (sollDatumFolge + nachlauf) Then
                            PruefeZeitfenster = 15
                            Exit Function
                        End If
                    End If
                Else
                    ' Kein Monatsfilter -> Vormonat-Check wie bisher
                    Dim sollDatumVormonat As Date
                    On Error Resume Next
                    sollDatumVormonat = DateSerial(Year(buchungsDatum), buchungsMonat - 1, SollTag)
                    If Err.Number <> 0 Then
                        Err.Clear
                        sollDatumVormonat = DateSerial(Year(buchungsDatum), buchungsMonat, 0)
                    End If
                    On Error GoTo 0
                    
                    If buchungsDatum >= (sollDatumVormonat - vorlauf) And _
                       buchungsDatum <= (sollDatumVormonat + nachlauf) Then
                        PruefeZeitfenster = 15
                        Exit Function
                    End If
                End If
            End If
            
WeiterNaechsteZeile:
        End If
    Next i
End Function


' =====================================================
' Hilfsfunktion: Pr�ft ob ein Monat (1-12) in einer
' kommaseparierten Monatsliste enthalten ist.
' z.B. IstMonatInListe(3, "03, 06, 09, 12") -> True
' PUBLIC (wird auch in mod_Zahlungspruefung ben�tigt)
' =====================================================
Public Function IstMonatInListe(ByVal monat As Long, ByVal monatListe As String) As Boolean
    Dim teile() As String
    Dim t As Long
    
    teile = Split(Replace(monatListe, " ", ""), ",")
    
    For t = LBound(teile) To UBound(teile)
        If IsNumeric(teile(t)) Then
            If CLng(teile(t)) = monat Then
                IstMonatInListe = True
                Exit Function
            End If
        End If
    Next t
    
    IstMonatInListe = False
End Function



' =====================================================
' Monat/Periode intelligent ermitteln (v10.0)
' v10.0 NEU:
'   - "j�hrlich (jahr)":       -> "[Kategoriename] [Jahr]"
'   - "j�hrlich (jahr/folgejahr)": -> "[Kategoriename] [Jahr]/[Folgejahr]"
'   - "j�hrlich" Fallback:     -> "j�hrlich"
'   - Sammelzahlung wird NICHT mit "Jahresbeitrag" bef�llt
'   - Ultimo-5 Bemerkung ohne "Ultimo-5:" Pr�fix
'   - Dynamischer Kategoriename aus Blatt "Daten" Spalte J
' =====================================================
Public Function ErmittleMonatPeriode(ByVal category As String, _
                                     ByVal buchungsDatum As Date, _
                                     ByVal faelligkeit As String, _
                                     Optional ByVal wsBK As Worksheet = Nothing, _
                                     Optional ByVal aktuelleZeile As Long = 0) As String
    
    Dim monatBuchung As Long
    monatBuchung = Month(buchungsDatum)
    
    Dim jahrBuchung As Long
    jahrBuchung = Year(buchungsDatum)
    
    If faelligkeit = "" Then faelligkeit = "monatlich"
    
    ' =============================================
    ' v10.0: Sammelzahlung NIEMALS automatisch zuordnen!
    ' Spalte I wird von VerarbeiteSammelzahlung gesetzt.
    ' =============================================
    If LCase(category) Like "*sammelzahlung*" Then
        ErmittleMonatPeriode = ""
        Exit Function
    End If
    
    ' =============================================
    ' Nicht-monatliche Perioden: direkt zuordnen
    ' v10.0: Neue F�lligkeitstypen mit Jahr/Folgejahr
    ' =============================================
    Dim faelligkeitLC As String
    faelligkeitLC = LCase(faelligkeit)
    
    ' --- "j�hrlich (jahr/folgejahr)" ---
    ' z.B. Versicherung -> "Versicherung 2025/2026"
    If faelligkeitLC Like "*hrlich (jahr/folgejahr)*" Or _
       faelligkeitLC Like "*jaehrlich (jahr/folgejahr)*" Or _
       faelligkeitLC = "j" & ChrW(228) & "hrlich (jahr/folgejahr)" Then
        ErmittleMonatPeriode = category & " " & jahrBuchung & "/" & (jahrBuchung + 1)
        Exit Function
    End If
    
    ' --- "j�hrlich (jahr)" ---
    ' z.B. Endabrechnung -> "Endabrechnung 2025"
    If faelligkeitLC Like "*hrlich (jahr)*" Or _
       faelligkeitLC Like "*jaehrlich (jahr)*" Or _
       faelligkeitLC = "j" & ChrW(228) & "hrlich (jahr)" Then
        ErmittleMonatPeriode = category & " " & jahrBuchung
        Exit Function
    End If
    
    ' --- "j�hrlich" (Fallback) ---
    If faelligkeitLC = "j" & ChrW(228) & "hrlich" Or _
       faelligkeitLC = "jaehrlich" Then
        ErmittleMonatPeriode = "j" & ChrW(228) & "hrlich"
        Exit Function
    End If
    
    ' --- Einmalig ---
    If faelligkeitLC = "einmalig" Then
        ErmittleMonatPeriode = MonthName(monatBuchung) & " (einmalig)"
        Exit Function
    End If
    
    ' --- Quartal ---
    If faelligkeitLC = "quartalsweise" Or faelligkeitLC = "quartal" Then
        Dim quartal As Long
        quartal = Int((monatBuchung - 1) / 3) + 1
        ErmittleMonatPeriode = "Q" & quartal & " " & jahrBuchung
        Exit Function
    End If
    
    ' --- Halbj�hrlich ---
    If faelligkeitLC = "halbjaehrlich" Or _
       faelligkeitLC = "halbj" & ChrW(228) & "hrlich" Then
        Dim halbjahr As Long
        halbjahr = IIf(monatBuchung <= 6, 1, 2)
        ErmittleMonatPeriode = "H" & halbjahr & " " & jahrBuchung
        Exit Function
    End If
    
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
            
            Dim SollTag As Long
            Dim vorlauf As Long
            Dim SollMonate As String
            
            SollTag = mCacheSollTag(idx)
            vorlauf = mCacheVorlauf(idx)
            SollMonate = mCacheSollMonate(idx)
            
            ' Pr�fe zuerst festen Stichtag (Spalte F)
            If IsDate(mCacheStichtag(idx)) Then
                Dim stichDatum As Date
                On Error Resume Next
                stichDatum = CDate(CStr(mCacheStichtag(idx)))
                If Err.Number = 0 Then
                    On Error GoTo 0
                    Dim stichAktuell As Date
                    stichAktuell = DateSerial(Year(buchungsDatum), Month(buchungsDatum), Day(stichDatum))
                    
                    Dim stichFolge As Date
                    stichFolge = DateSerial(Year(buchungsDatum), Month(buchungsDatum) + 1, Day(stichDatum))
                    
                    If vorlauf > 0 And buchungsDatum >= (stichFolge - vorlauf) Then
                        ErmittleMonatPeriode = MonthName(Month(stichFolge))
                        Exit Function
                    End If
                    
                    If vorlauf > 0 And buchungsDatum >= (stichAktuell - vorlauf) Then
                        ErmittleMonatPeriode = MonthName(Month(stichAktuell))
                        Exit Function
                    End If
                Else
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
            
            ' =============================================
            ' Ultimo-5-Logik (v9.5, Bemerkung v10.0 angepasst)
            ' =============================================
            Dim letzterTagMonat As Long
            letzterTagMonat = Day(DateSerial(Year(buchungsDatum), monatBuchung + 1, 0))
            
            Dim effektiverTag As Long
            If SollTag = 31 Or SollTag = 0 Then
                effektiverTag = letzterTagMonat
            Else
                effektiverTag = SollTag
            End If
            
            Dim istUltimoBereich As Boolean
            istUltimoBereich = (effektiverTag >= (letzterTagMonat - 5))
            
            Dim tagBuchung As Long
            tagBuchung = Day(buchungsDatum)
            
            If istUltimoBereich And tagBuchung >= (letzterTagMonat - 5) And tagBuchung < letzterTagMonat Then
                
                Dim folgeMonatNr As Long
                folgeMonatNr = monatBuchung + 1
                If folgeMonatNr > 12 Then folgeMonatNr = 1
                
                If SollMonate <> "" And Not IstMonatInListe(folgeMonatNr, SollMonate) Then
                    GoTo FallbackMonat
                End If
                
                ' Lern-Check
                If Not wsBK Is Nothing And aktuelleZeile > 0 Then
                    Dim ibanAktuell As String
                    ibanAktuell = UCase(Replace(Trim(wsBK.Cells(aktuelleZeile, BK_COL_IBAN).value), " ", ""))
                    
                    If ibanAktuell <> "" Then
                        Dim suchZeile As Long
                        Dim musterGefunden As Boolean
                        musterGefunden = False
                        
                        For suchZeile = BK_START_ROW To aktuelleZeile - 1
                            Dim ibanZeile As String
                            ibanZeile = UCase(Replace(Trim(wsBK.Cells(suchZeile, BK_COL_IBAN).value), " ", ""))
                            If ibanZeile <> ibanAktuell Then GoTo NaechsteLernZeile
                            
                            If StrComp(Trim(wsBK.Cells(suchZeile, BK_COL_KATEGORIE).value), category, vbTextCompare) <> 0 Then GoTo NaechsteLernZeile
                            
                            If wsBK.Cells(suchZeile, BK_COL_MONAT_PERIODE).Interior.color = RGB(198, 239, 206) Then
                                If InStr(LCase(CStr(wsBK.Cells(suchZeile, BK_COL_BEMERKUNG).value)), "folgemonat") > 0 Then
                                    musterGefunden = True
                                    Exit For
                                End If
                            End If
NaechsteLernZeile:
                        Next suchZeile
                        
                        If musterGefunden Then
                            ErmittleMonatPeriode = MonthName(folgeMonatNr)
                            Exit Function
                        End If
                    End If
                End If
                
                ' v10.0: GELB-R�ckgabe OHNE "Ultimo-5:" Pr�fix
                ErmittleMonatPeriode = "GELB|" & MonthName(monatBuchung)
                Exit Function
                
            End If
            
            ' Bisherige Logik: SollTag + Vorlauf (nicht Ultimo-Bereich)
            If effektiverTag >= 1 And effektiverTag <= 31 Then
                If vorlauf > 0 And tagBuchung > effektiverTag Then
                    Dim sollDatumFolge As Date
                    On Error Resume Next
                    sollDatumFolge = DateSerial(Year(buchungsDatum), Month(buchungsDatum) + 1, effektiverTag)
                    If Err.Number <> 0 Then
                        Err.Clear
                        On Error GoTo 0
                        GoTo FallbackMonat
                    End If
                    On Error GoTo 0
                    
                    Dim differenzTage As Long
                    differenzTage = CLng(sollDatumFolge - buchungsDatum)
                    
                    If differenzTage >= 0 And differenzTage <= vorlauf Then
                        Dim folgeMon As Long
                        folgeMon = Month(sollDatumFolge)
                        If SollMonate = "" Or IstMonatInListe(folgeMon, SollMonate) Then
                            ErmittleMonatPeriode = MonthName(folgeMon)
                            Exit Function
                        End If
                    End If
                End If
            End If
            
            GoTo FallbackMonat
        End If
    Next idx
    
FallbackMonat:
    ErmittleMonatPeriode = MonthName(monatBuchung)
End Function
