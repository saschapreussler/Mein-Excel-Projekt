Attribute VB_Name = "mod_Zahlungspruefung"
Option Explicit

' ***************************************************************
' MODUL: mod_Zahlungspruefung (Orchestrator)
' VERSION: 3.2 - 01.03.2026
' ZWECK: Zahlungspruefung fuer Mitgliederliste + Einstellungen
'        - Prueft Zahlungseingaenge gegen Soll-Werte
'        - Behandelt Dezember-Vorauszahlungen
'        - Cache-Verwaltung (Einstellungen, IBAN, Dezember)
'        - SetzeMonatPeriode
'        - HoleFaelligkeitFuerKategorie
' AUSGELAGERT:
'   - mod_ZP_DropDowns: SetzeBankkontoDropDowns, Kategorie-/Monat-
'     DropDowns, Hilfsspalten AF/AG, Spaltenentsperrung
'   - mod_ZP_Sammelzuordnung: Sammelueberweisungen, manuelle
'     Monatszuordnung
' FIX v3.1: PruefeZahlungen nutzt jetzt Spalte I (Monat/Periode)
'           statt Month(Buchungsdatum) f?r Monats-Zuordnung
' NEU v3.2: Frist-/Toleranzpr?fung mit Vorlauf/Nachlauf aus
'           Einstellungen (Spalte G/H). S?umnishinweis in Bemerkung.
' ***************************************************************

' ===============================================================
' CACHE FUER EINSTELLUNGEN (Performance-Optimierung)
' ===============================================================
Private Type EinstellungsRegelZP
    kategorie As String
    SollBetrag As Double
    SollTag As Long
    SollMonate As String           ' z.B. "03, 06, 09"
    StichtagFix As String          ' z.B. "15.03"
    VorlaufTage As Long
    NachlaufTage As Long
    saeumnisGebuehr As Double
End Type

Private m_EinstellungenCacheZP() As EinstellungsRegelZP
Private m_EinstellungenGeladenZP As Boolean

' ===============================================================
' IBAN-CACHE: EntityKey -> IBAN (aus Daten!R+S)
' ===============================================================
Private m_EntityIBANCacheZP As Object   ' Dictionary: EntityKey -> IBAN
Private m_EntityIBANCacheGeladenZP As Boolean

' ===============================================================
' DEZEMBER-CACHE (fuer Vorauszahlungen)
' Struktur: Schluessel = IBAN|Kategorie, Wert = Collection von Betraegen
' ===============================================================
Private m_DezemberCacheZP As Object

' ===============================================================
' AMPELFARBEN (Konsistenz mit KategorieEngine)
' ===============================================================
Private Const AMPEL_GRUEN As Long = 12968900
Private Const AMPEL_GELB As Long = 10086143
Private Const AMPEL_ROT As Long = 9871103


' ===============================================================
' HAUPTFUNKTION: Prueft ALLE Zahlungen eines Mitglieds/einer Kategorie
' Wird von mod_Uebersicht_Generator aufgerufen
'
' Rueckgabe: "STATUS|Soll:XX.XX|Ist:XX.XX" oder
'            "STATUS|Soll:XX.XX|Ist:XX.XX|Bemerkungstext"
'           Dezimaltrenner im Rueckgabewert ist IMMER Punkt (.)
'
' v3.2: Frist-/Toleranzpruefung:
'   - Vorlauf (Spalte G) und Nachlauf (Spalte H) aus Einstellungen
'   - F?lligkeitsdatum wird berechnet (BerechneSollDatumZP)
'   - Zahlung innerhalb [F?lligkeit - Vorlauf, F?lligkeit + Nachlauf] = p?nktlich
'   - Zahlung eingegangen aber NACH F?lligkeit + Nachlauf = GELB + S?umnis
'   - Keine Zahlung = ROT
' ===============================================================
Public Function PruefeZahlungen(ByVal entityKey As String, _
                                 ByVal kategorie As String, _
                                 ByVal monat As Long, _
                                 ByVal jahr As Long) As String
    
    On Error GoTo ErrorHandler
    
    Dim wsBK As Worksheet
    Dim soll As Double
    Dim ist As Double
    Dim status As String
    Dim r As Long
    Dim lastRow As Long
    Dim zahlDatum As Date
    Dim zahlBetrag As Double
    Dim zahlKat As String
    Dim ibanZeile As String
    Dim entityIBAN As String
    Dim bemerkung As String
    
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    ' Einstellungen-Cache laden (falls noch nicht geschehen)
    If Not m_EinstellungenGeladenZP Then Call LadeEinstellungenCacheZP
    
    ' IBAN-Cache laden (falls noch nicht geschehen)
    If Not m_EntityIBANCacheGeladenZP Then Call LadeEntityIBANCacheZP
    
    ' 1. IBAN zum EntityKey aufloesen (ueber Daten!R+S)
    entityIBAN = ""
    If Not m_EntityIBANCacheZP Is Nothing Then
        If m_EntityIBANCacheZP.Exists(entityKey) Then
            entityIBAN = m_EntityIBANCacheZP(entityKey)
        End If
    End If
    
    If entityIBAN = "" Then
        PruefeZahlungen = "GELB|Soll:0.00|Ist:0.00|Keine IBAN zum EntityKey"
        Exit Function
    End If
    
    ' 2. Soll-Wert aus Einstellungen holen
    soll = HoleSollBetragZP(kategorie)
    
    ' v2.0 FIX: NICHT mehr abbrechen bei soll=0!
    ' Bei variablem Betrag (soll=0) wird trotzdem geprueft ob Zahlung da ist.
    
    ' 3. Ist-Wert aus Bankkonto ermitteln
    '    v3.1: Monat-Matching jetzt ueber Spalte I (Monat/Periode)
    '          statt ueber Month(Buchungsdatum), da Spalte I bereits
    '          die korrekte Monats-Zuordnung enthaelt (inkl. Vorlauf/Nachlauf)
    ist = 0
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    ' Erwarteter Monatname (z.B. "Januar", "Februar", ...)
    Dim erwarteterMonat As String
    erwarteterMonat = MonthName(monat)
    
    ' v3.2: Fruehestes Zahlungsdatum merken (fuer Fristpruefung)
    Dim fruehestesZahlDatum As Date
    Dim hatZahlung As Boolean
    hatZahlung = False
    
    For r = BK_START_ROW To lastRow
        ' Datum pruefen (muss vorhanden sein fuer Jahr-Check)
        If Not IsDate(wsBK.Cells(r, BK_COL_DATUM).value) Then GoTo NextZahlRow
        zahlDatum = wsBK.Cells(r, BK_COL_DATUM).value
        
        ' Jahr pruefen ueber Buchungsdatum
        If Year(zahlDatum) <> jahr Then
            ' Dezember-Sonderfall: Vorauszahlung Dezember Vorjahr fuer Januar
            If monat = 1 And Month(zahlDatum) = 12 And Year(zahlDatum) = jahr - 1 Then
                ' Vorauszahlung aus Dezember des Vorjahres -> zulaessig
            Else
                GoTo NextZahlRow
            End If
        End If
        
        ' Monat pruefen ueber Spalte I (Monat/Periode)
        Dim monatPeriode As String
        monatPeriode = Trim(CStr(wsBK.Cells(r, BK_COL_MONAT_PERIODE).value))
        If StrComp(monatPeriode, erwarteterMonat, vbTextCompare) <> 0 Then GoTo NextZahlRow
        
        ' IBAN pruefen (Spalte D = BK_COL_IBAN)
        ibanZeile = Replace(Trim(CStr(wsBK.Cells(r, BK_COL_IBAN).value)), " ", "")
        If StrComp(ibanZeile, entityIBAN, vbTextCompare) <> 0 Then GoTo NextZahlRow
        
        ' Kategorie pruefen (Spalte H = BK_COL_KATEGORIE)
        zahlKat = Trim(CStr(wsBK.Cells(r, BK_COL_KATEGORIE).value))
        If StrComp(zahlKat, kategorie, vbTextCompare) <> 0 Then GoTo NextZahlRow
        
        ' Betrag addieren (Spalte B = BK_COL_BETRAG)
        zahlBetrag = wsBK.Cells(r, BK_COL_BETRAG).value
        ist = ist + Abs(zahlBetrag)
        
        ' v3.2: Fruehestes Zahlungsdatum merken
        If Not hatZahlung Then
            fruehestesZahlDatum = zahlDatum
            hatZahlung = True
        ElseIf zahlDatum < fruehestesZahlDatum Then
            fruehestesZahlDatum = zahlDatum
        End If
        
NextZahlRow:
    Next r
    
    ' 4. Status ermitteln (GRUEN/GELB/ROT)
    '    v3.2: Mit Frist-/Toleranzpruefung
    bemerkung = ""
    
    ' Faelligkeitsdatum und Toleranzen aus Einstellungen holen
    Dim sollDatum As Date
    Dim vorlauf As Long
    Dim nachlauf As Long
    Dim saeumnisGebuehr As Double
    sollDatum = BerechneSollDatumZP(kategorie, monat, jahr)
    Call HoleToleranzZP(kategorie, vorlauf, nachlauf, saeumnisGebuehr)
    
    If soll > 0 Then
        ' Fester Soll-Betrag vorhanden: Betrags-Vergleich + Fristpruefung
        If ist >= soll Then
            ' Betrag ausreichend -> Fristpruefung
            If hatZahlung And (vorlauf > 0 Or nachlauf > 0) Then
                Dim fristEnde As Date
                fristEnde = sollDatum + nachlauf
                
                If fruehestesZahlDatum > fristEnde Then
                    ' Zahlung NACH Frist -> GELB + Saeumnis
                    status = "GELB"
                    bemerkung = "Versp" & ChrW(228) & "tet (" & Format(fruehestesZahlDatum, "dd.mm.yyyy") & _
                                ", Frist: " & Format(fristEnde, "dd.mm.yyyy") & ")"
                    If saeumnisGebuehr > 0 Then
                        bemerkung = bemerkung & " | S" & ChrW(228) & "umnis: " & _
                                    Format(saeumnisGebuehr, "#,##0.00") & " " & ChrW(8364)
                    End If
                Else
                    ' Zahlung fristgerecht
                    status = "GR" & ChrW(220) & "N"
                End If
            Else
                ' Keine Toleranz definiert oder keine Zahlung mit Datum
                status = "GR" & ChrW(220) & "N"
            End If
        ElseIf ist > 0 Then
            status = "GELB"
            bemerkung = "Teilzahlung (Soll: " & Format(soll, "#,##0.00") & _
                        ", Ist: " & Format(ist, "#,##0.00") & ")"
        Else
            ' Keine Zahlung -> ROT, aber nur wenn Faelligkeit schon erreicht
            If Date >= sollDatum Then
                status = "ROT"
                If nachlauf > 0 And Date <= sollDatum + nachlauf Then
                    status = "GELB"
                    bemerkung = "Noch offen (Frist bis " & Format(sollDatum + nachlauf, "dd.mm.yyyy") & ")"
                Else
                    If saeumnisGebuehr > 0 And Date > sollDatum + nachlauf Then
                        bemerkung = "S" & ChrW(228) & "umnis: " & _
                                    Format(saeumnisGebuehr, "#,##0.00") & " " & ChrW(8364)
                    End If
                End If
            Else
                ' Faelligkeit noch nicht erreicht -> GELB (ausstehend)
                status = "GELB"
                bemerkung = "F" & ChrW(228) & "llig am " & Format(sollDatum, "dd.mm.yyyy")
            End If
        End If
    Else
        ' Kein fester Soll-Betrag (variabel): nur Eingangs-Pruefung
        If ist > 0 Then
            status = "GR" & ChrW(220) & "N"
        Else
            If Date >= sollDatum Then
                status = "ROT"
            Else
                status = "GELB"
                bemerkung = "F" & ChrW(228) & "llig am " & Format(sollDatum, "dd.mm.yyyy")
            End If
        End If
    End If
    
    ' 5. Ergebnis formatieren (IMMER Punkt als Dezimaltrenner!)
    PruefeZahlungen = status & "|Soll:" & FormatDezimalPunkt(soll) & "|Ist:" & FormatDezimalPunkt(ist)
    If bemerkung <> "" Then
        PruefeZahlungen = PruefeZahlungen & "|" & bemerkung
    End If
    Exit Function
    
ErrorHandler:
    PruefeZahlungen = "ROT|Fehler:" & Err.Description
    
End Function


' ===============================================================
' HILFSFUNKTION: Double -> String mit Punkt als Dezimaltrenner
' ===============================================================
Private Function FormatDezimalPunkt(ByVal wert As Double) As String
    Dim s As String
    s = Format(wert, "0.00")
    s = Replace(s, ",", ".")
    FormatDezimalPunkt = s
End Function


' ===============================================================
' Holt Vorlauf/Nachlauf/S?umnis-Geb?hr aus dem Einstellungen-Cache
' fuer eine bestimmte Kategorie
' ===============================================================
Public Sub HoleToleranzZP(ByVal kategorie As String, _
                            ByRef vorlauf As Long, _
                            ByRef nachlauf As Long, _
                            ByRef saeumnisGebuehr As Double)
    
    Dim i As Long
    vorlauf = 0
    nachlauf = 0
    saeumnisGebuehr = 0
    
    If Not m_EinstellungenGeladenZP Then Exit Sub
    
    On Error Resume Next
    For i = LBound(m_EinstellungenCacheZP) To UBound(m_EinstellungenCacheZP)
        If StrComp(m_EinstellungenCacheZP(i).kategorie, kategorie, vbTextCompare) = 0 Then
            vorlauf = m_EinstellungenCacheZP(i).VorlaufTage
            nachlauf = m_EinstellungenCacheZP(i).NachlaufTage
            saeumnisGebuehr = m_EinstellungenCacheZP(i).saeumnisGebuehr
            Exit Sub
        End If
    Next i
    On Error GoTo 0
    
End Sub


' ===============================================================
' IBAN-CACHE: Laedt EntityKey -> IBAN Zuordnung aus Daten!R+S
' ===============================================================
Private Sub LadeEntityIBANCacheZP()
    
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim ek As String
    Dim iban As String
    
    Set m_EntityIBANCacheZP = CreateObject("Scripting.Dictionary")
    m_EntityIBANCacheGeladenZP = False
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    
    lastRow = wsDaten.Cells(wsDaten.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    If lastRow < EK_START_ROW Then Exit Sub
    
    For r = EK_START_ROW To lastRow
        ek = Trim(CStr(wsDaten.Cells(r, EK_COL_ENTITYKEY).value))
        iban = Replace(Trim(CStr(wsDaten.Cells(r, EK_COL_IBAN).value)), " ", "")
        
        If ek <> "" And iban <> "" Then
            If Not m_EntityIBANCacheZP.Exists(ek) Then
                m_EntityIBANCacheZP.Add ek, iban
            End If
        End If
    Next r
    
    m_EntityIBANCacheGeladenZP = True
    
End Sub


' ===============================================================
' IBAN-CACHE: Freigeben
' ===============================================================
Private Sub EntladeEntityIBANCacheZP()
    
    Set m_EntityIBANCacheZP = Nothing
    m_EntityIBANCacheGeladenZP = False
    
End Sub


' ===============================================================
' Soll-Betrag aus Einstellungen holen (mit Cache)
' ===============================================================
Private Function HoleSollBetragZP(ByVal kategorie As String) As Double
    
    Dim i As Long
    
    If Not m_EinstellungenGeladenZP Then Call LadeEinstellungenCacheZP
    
    On Error Resume Next
    For i = LBound(m_EinstellungenCacheZP) To UBound(m_EinstellungenCacheZP)
        If StrComp(m_EinstellungenCacheZP(i).kategorie, kategorie, vbTextCompare) = 0 Then
            HoleSollBetragZP = m_EinstellungenCacheZP(i).SollBetrag
            Exit Function
        End If
    Next i
    On Error GoTo 0
    
    HoleSollBetragZP = 0
    
End Function


' ===============================================================
' Soll-Datum berechnen (mit Spalte D/E vs F Logik)
' ===============================================================
Public Function BerechneSollDatumZP(ByVal kategorie As String, _
                                      ByVal monat As Long, _
                                      ByVal jahr As Long) As Date
    
    Dim i As Long
    Dim regel As EinstellungsRegelZP
    Dim tag As Long
    Dim istMonatGueltig As Boolean
    
    If Not m_EinstellungenGeladenZP Then Call LadeEinstellungenCacheZP
    
    ' 1. Regel finden
    For i = LBound(m_EinstellungenCacheZP) To UBound(m_EinstellungenCacheZP)
        If StrComp(m_EinstellungenCacheZP(i).kategorie, kategorie, vbTextCompare) = 0 Then
            regel = m_EinstellungenCacheZP(i)
            Exit For
        End If
    Next i
    
    If regel.kategorie = "" Then
        BerechneSollDatumZP = DateSerial(jahr, monat, 1)
        Exit Function
    End If
    
    ' 2. Pruefen: Spalte F (Stichtag Fix) hat Vorrang
    If regel.StichtagFix <> "" Then
        Dim teile() As String
        teile = Split(regel.StichtagFix, ".")
        If UBound(teile) >= 1 Then
            tag = CLng(teile(0))
            Dim fixMonat As Long
            fixMonat = CLng(teile(1))
            If fixMonat = monat Then
                BerechneSollDatumZP = DateSerial(jahr, monat, tag)
            Else
                BerechneSollDatumZP = DateSerial(jahr, monat, 1)
            End If
            Exit Function
        End If
    End If
    
    ' 3. Spalte D/E verwenden (SollTag + SollMonate)
    istMonatGueltig = False
    If regel.SollMonate <> "" Then
        Dim monate() As String
        monate = Split(regel.SollMonate, ",")
        Dim m As Long
        For m = LBound(monate) To UBound(monate)
            If IsNumeric(Trim(monate(m))) Then
                If CLng(Trim(monate(m))) = monat Then
                    istMonatGueltig = True
                    Exit For
                End If
            End If
        Next m
    Else
        istMonatGueltig = True
    End If
    
    If Not istMonatGueltig Then
        BerechneSollDatumZP = DateSerial(jahr, monat, 1)
        Exit Function
    End If
    
    tag = regel.SollTag
    If tag = 0 Then tag = 1
    If tag > 28 Then
        BerechneSollDatumZP = DateSerial(jahr, monat + 1, 0)
    Else
        BerechneSollDatumZP = DateSerial(jahr, monat, tag)
    End If
    
End Function


' ===============================================================
' Einstellungen-Cache laden (Performance-Optimierung)
' ===============================================================
Public Sub LadeEinstellungenCacheZP()
    
    Dim wsEinst As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim idx As Long
    
    On Error Resume Next
    Set wsEinst = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    
    If wsEinst Is Nothing Then
        m_EinstellungenGeladenZP = False
        Exit Sub
    End If
    
    lastRow = wsEinst.Cells(wsEinst.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    If lastRow < ES_START_ROW Then
        m_EinstellungenGeladenZP = False
        Exit Sub
    End If
    
    ReDim m_EinstellungenCacheZP(0 To lastRow - ES_START_ROW)
    idx = 0
    
    For r = ES_START_ROW To lastRow
        Dim kat As String
        kat = Trim(CStr(wsEinst.Cells(r, ES_COL_KATEGORIE).value))
        If kat = "" Then GoTo NextEinstRow
        
        With m_EinstellungenCacheZP(idx)
            .kategorie = kat
            
            If IsNumeric(wsEinst.Cells(r, ES_COL_SOLL_BETRAG).value) Then
                .SollBetrag = CDbl(wsEinst.Cells(r, ES_COL_SOLL_BETRAG).value)
            Else
                .SollBetrag = 0
            End If
            
            If IsNumeric(wsEinst.Cells(r, ES_COL_SOLL_TAG).value) Then
                .SollTag = CLng(wsEinst.Cells(r, ES_COL_SOLL_TAG).value)
            Else
                .SollTag = 0
            End If
            
            .SollMonate = Trim(CStr(wsEinst.Cells(r, ES_COL_SOLL_MONATE).value))
            .StichtagFix = Trim(CStr(wsEinst.Cells(r, ES_COL_STICHTAG_FIX).value))
            
            If IsNumeric(wsEinst.Cells(r, ES_COL_VORLAUF).value) Then
                .VorlaufTage = CLng(wsEinst.Cells(r, ES_COL_VORLAUF).value)
            Else
                .VorlaufTage = 0
            End If
            
            If IsNumeric(wsEinst.Cells(r, ES_COL_NACHLAUF).value) Then
                .NachlaufTage = CLng(wsEinst.Cells(r, ES_COL_NACHLAUF).value)
            Else
                .NachlaufTage = 0
            End If
            
            If IsNumeric(wsEinst.Cells(r, ES_COL_SAEUMNIS).value) Then
                .saeumnisGebuehr = CDbl(wsEinst.Cells(r, ES_COL_SAEUMNIS).value)
            Else
                .saeumnisGebuehr = 0
            End If
        End With
        
        idx = idx + 1
        
NextEinstRow:
    Next r
    
    If idx > 0 Then
        ReDim Preserve m_EinstellungenCacheZP(0 To idx - 1)
        m_EinstellungenGeladenZP = True
    Else
        m_EinstellungenGeladenZP = False
    End If
    
End Sub


' ===============================================================
' Einstellungen-Cache freigeben (Speicher sparen)
' ===============================================================
Public Sub EntladeEinstellungenCacheZP()
    
    Erase m_EinstellungenCacheZP
    m_EinstellungenGeladenZP = False
    
    Call EntladeEntityIBANCacheZP
    
End Sub


' ===============================================================
' DEZEMBER-VORAUSZAHLUNGEN: Cache initialisieren
' ===============================================================
Public Sub InitialisiereNachDezemberCacheZP(ByVal jahr As Long)
    
    Dim wsBK As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim zahlDatum As Date
    Dim zahlBetrag As Double
    Dim ibanWert As String
    Dim kategorie As String
    Dim col As Collection
    
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set m_DezemberCacheZP = CreateObject("Scripting.Dictionary")
    
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    For r = BK_START_ROW To lastRow
        If Not IsDate(wsBK.Cells(r, BK_COL_DATUM).value) Then GoTo NextDezRow
        zahlDatum = wsBK.Cells(r, BK_COL_DATUM).value
        
        If Year(zahlDatum) <> jahr - 1 Then GoTo NextDezRow
        If Month(zahlDatum) <> 12 Then GoTo NextDezRow
        
        ibanWert = Replace(Trim(CStr(wsBK.Cells(r, BK_COL_IBAN).value)), " ", "")
        If ibanWert = "" Then GoTo NextDezRow
        
        kategorie = Trim(CStr(wsBK.Cells(r, BK_COL_KATEGORIE).value))
        If kategorie = "" Then GoTo NextDezRow
        
        zahlBetrag = Abs(wsBK.Cells(r, BK_COL_BETRAG).value)
        
        Dim cacheKey As String
        cacheKey = ibanWert & "|" & kategorie
        
        If Not m_DezemberCacheZP.Exists(cacheKey) Then
            Set col = New Collection
            m_DezemberCacheZP.Add cacheKey, col
        Else
            Set col = m_DezemberCacheZP(cacheKey)
        End If
        
        col.Add zahlBetrag
        
NextDezRow:
    Next r
    
End Sub


' ===============================================================
' DEZEMBER-VORAUSZAHLUNGEN: Betrag aus Cache holen
' ===============================================================
Public Function HoleDezemberVorauszahlungZP(ByVal entityKey As String, _
                                             ByVal kategorie As String) As Double
    
    Dim cacheKey As String
    Dim col As Collection
    Dim summe As Double
    Dim v As Variant
    Dim entityIBAN As String
    
    entityIBAN = ""
    If Not m_EntityIBANCacheZP Is Nothing Then
        If m_EntityIBANCacheZP.Exists(entityKey) Then
            entityIBAN = m_EntityIBANCacheZP(entityKey)
        End If
    End If
    
    If entityIBAN = "" Then
        HoleDezemberVorauszahlungZP = 0
        Exit Function
    End If
    
    cacheKey = entityIBAN & "|" & kategorie
    
    If m_DezemberCacheZP Is Nothing Then
        HoleDezemberVorauszahlungZP = 0
        Exit Function
    End If
    
    If Not m_DezemberCacheZP.Exists(cacheKey) Then
        HoleDezemberVorauszahlungZP = 0
        Exit Function
    End If
    
    Set col = m_DezemberCacheZP(cacheKey)
    summe = 0
    
    For Each v In col
        summe = summe + CDbl(v)
    Next v
    
    HoleDezemberVorauszahlungZP = summe
    
End Function


' ===============================================================
' MONAT/PERIODE SETZEN (ueberarbeitet)
' FIX v1.5: Application.EnableEvents = False VOR dem Beschreiben
'           von Spalte I, damit Worksheet_Change NICHT getriggert wird.
' v2.0: Am Ende wird SetzeBankkontoDropDowns aufgerufen (fuer H + I)
' ===============================================================
Public Sub SetzeMonatPeriode(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim r As Long
    Dim monatWert As Variant
    Dim datumWert As Variant
    Dim kategorie As String
    Dim faelligkeit As String
    Dim ergebnis As String
    
    Dim eventsWaren As Boolean
    eventsWaren = Application.EnableEvents
    
    On Error GoTo SetzeMonatPeriodeError
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    ' Events ABSCHALTEN bevor Spalte I beschrieben wird
    Application.EnableEvents = False
    
    Dim wsDaten As Worksheet
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    
    ' Einstellungen-Cache laden
    Call LadeEinstellungenCache
    
    For r = BK_START_ROW To lastRow
        datumWert = ws.Cells(r, BK_COL_DATUM).value
        monatWert = ws.Cells(r, BK_COL_MONAT_PERIODE).value
        
        If IsDate(datumWert) And (isEmpty(monatWert) Or CStr(monatWert) = "") Then
            kategorie = Trim(CStr(ws.Cells(r, BK_COL_KATEGORIE).value))
            
            If kategorie <> "" Then
                faelligkeit = HoleFaelligkeitFuerKategorie(wsDaten, kategorie)
                
                ergebnis = mod_KategorieEngine_Zeitraum.ErmittleMonatPeriode( _
                    kategorie, CDate(datumWert), faelligkeit, ws, r)
                
                If Left(ergebnis, 5) = "GELB|" Then
                    Dim monatName As String
                    monatName = Mid(ergebnis, 6)
                    
                    ws.Cells(r, BK_COL_MONAT_PERIODE).value = monatName
                    ws.Cells(r, BK_COL_MONAT_PERIODE).Interior.color = RGB(255, 235, 156)
                    
                    Dim bestehendeBemerkung As String
                    bestehendeBemerkung = Trim(CStr(ws.Cells(r, BK_COL_BEMERKUNG).value))
                    
                    Dim gelbHinweis As String
                    gelbHinweis = "Bitte pr" & ChrW(252) & "fen ob Zahlung f" & ChrW(252) & "r " & _
                                  monatName & " oder Folgemonat gilt"
                    
                    If bestehendeBemerkung = "" Then
                        ws.Cells(r, BK_COL_BEMERKUNG).value = gelbHinweis
                    Else
                        ws.Cells(r, BK_COL_BEMERKUNG).value = bestehendeBemerkung & vbLf & gelbHinweis
                    End If
                    
                    ' Hell-gelber Hintergrund fuer die Bemerkung (gleiche Farbe wie Spalte I)
                    ws.Cells(r, BK_COL_BEMERKUNG).Interior.color = RGB(255, 235, 156)
                Else
                    ws.Cells(r, BK_COL_MONAT_PERIODE).value = ergebnis
                    ' Ampelfarbe Gruen = Monat eindeutig bestimmt
                    ws.Cells(r, BK_COL_MONAT_PERIODE).Interior.color = RGB(198, 239, 206)
                End If
            Else
                ws.Cells(r, BK_COL_MONAT_PERIODE).value = MonthName(Month(datumWert))
            End If
        End If
    Next r
    
    ' Einstellungen-Cache wieder freigeben
    Call mod_KategorieEngine_Zeitraum.EntladeEinstellungenCache
    
    ' v1.5 FIX: Events wieder einschalten
    Application.EnableEvents = eventsWaren
    
    ' v2.0: ALLE DropDowns setzen (H + I) und Spalten entsperren
    ' (ausgelagert nach mod_ZP_DropDowns)
    Call mod_ZP_DropDowns.SetzeBankkontoDropDowns(ws)
    
    Exit Sub

SetzeMonatPeriodeError:
    Application.EnableEvents = eventsWaren
    Debug.Print "Fehler in SetzeMonatPeriode: " & Err.Number & " - " & Err.Description
    
End Sub


' ===============================================================
' FAELLIGKEIT AUS KATEGORIE-TABELLE (Spalte O) HOLEN
' ===============================================================
Public Function HoleFaelligkeitFuerKategorie(ByVal wsDaten As Worksheet, _
                                              ByVal kategorie As String) As String
    Dim lastRow As Long
    Dim r As Long
    
    ' PRIO 1: Einstellungen-Blatt pruefen (Spalte B = Kategorie)
    Dim wsEinst As Worksheet
    On Error Resume Next
    Set wsEinst = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    
    If Not wsEinst Is Nothing Then
        lastRow = wsEinst.Cells(wsEinst.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
        For r = ES_START_ROW To lastRow
            If StrComp(Trim(CStr(wsEinst.Cells(r, ES_COL_KATEGORIE).value)), kategorie, vbTextCompare) = 0 Then
                Dim SollMonate As String
                SollMonate = Trim(CStr(wsEinst.Cells(r, ES_COL_SOLL_MONATE).value))
                If SollMonate = "" Then
                    HoleFaelligkeitFuerKategorie = "monatlich"
                Else
                    Dim anzMonate As Long
                    anzMonate = UBound(Split(SollMonate, ",")) + 1
                    Select Case anzMonate
                        Case 1: HoleFaelligkeitFuerKategorie = "j" & ChrW(228) & "hrlich"
                        Case 2: HoleFaelligkeitFuerKategorie = "halbj" & ChrW(228) & "hrlich"
                        Case 4: HoleFaelligkeitFuerKategorie = "viertelj" & ChrW(228) & "hrlich"
                        Case Else: HoleFaelligkeitFuerKategorie = "monatlich"
                    End Select
                End If
                Exit Function
            End If
        Next r
    End If
    
    ' PRIO 2: Fallback auf Daten-Blatt (Spalte O = Faelligkeit)
    lastRow = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    
    For r = DATA_START_ROW To lastRow
        If Trim(wsDaten.Cells(r, DATA_CAT_COL_KATEGORIE).value) = kategorie Then
            HoleFaelligkeitFuerKategorie = LCase(Trim(wsDaten.Cells(r, DATA_CAT_COL_FAELLIGKEIT).value))
            Exit Function
        End If
    Next r
    
    HoleFaelligkeitFuerKategorie = "monatlich"
End Function









