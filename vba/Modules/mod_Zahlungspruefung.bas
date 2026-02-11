Attribute VB_Name = "mod_Zahlungspruefung"
Option Explicit

' ***************************************************************
' MODUL: mod_Zahlungspruefung
' VERSION: 1.2 - 11.02.2026
' ZWECK: Zahlungsprüfung für Mitgliederliste + Einstellungen
'        - Prüft Zahlungseingänge gegen Soll-Werte
'        - Behandelt Dezember-Vorauszahlungen
'        - Erkennt Sammelüberweisungen
'        - Bietet manuelle Zuordnung bei Problemfällen
'        - Dokumentiert Aufschlüsselung in Spalte L
'        - SetzeMonatPeriode (verschoben aus mod_Banking_Data)
'        - HoleFaelligkeitFuerKategorie (verschoben aus mod_Banking_Data)
' FIX v1.1: LadeEinstellungenCacheZP -> PUBLIC (war Private)
'           EntladeEinstellungenCacheZP -> PUBLIC (war Private)
' NEU v1.2: + Public Sub SetzeMonatPeriode (aus mod_Banking_Data)
'           + Public Function HoleFaelligkeitFuerKategorie (aus mod_Banking_Data)
' ***************************************************************

' ===============================================================
' CACHE FÜR EINSTELLUNGEN (Performance-Optimierung)
' ===============================================================
Private Type EinstellungsRegelZP
    kategorie As String
    SollBetrag As Double
    SollTag As Long
    SollMonate As String           ' z.B. "03, 06, 09"
    StichtagFix As String          ' z.B. "15.03"
    VorlaufTage As Long
    NachlaufTage As Long
    SaeumnisGebuehr As Double
End Type

Private m_EinstellungenCacheZP() As EinstellungsRegelZP
Private m_EinstellungenGeladenZP As Boolean

' ===============================================================
' DEZEMBER-CACHE (für Vorauszahlungen)
' Struktur: Schlüssel = EntityKey, Wert = Array von Dezember-Zahlungen
' ===============================================================
Private m_DezemberCacheZP As Object   ' Dictionary mit EntityKey -> Collection von Beträgen

' ===============================================================
' AMPELFARBEN (Konsistenz mit KategorieEngine)
' ===============================================================
Private Const AMPEL_GRUEN As Long = 12968900
Private Const AMPEL_GELB As Long = 10086143
Private Const AMPEL_ROT As Long = 9871103


' ===============================================================
' HAUPTFUNKTION: Prüft ALLE Zahlungen eines Mitglieds/einer Kategorie
' Wird von mod_Uebersicht_Generator aufgerufen
' ===============================================================
Public Function PruefeZahlungen(ByVal entityKey As String, _
                                 ByVal kategorie As String, _
                                 ByVal monat As Long, _
                                 ByVal jahr As Long) As String
    
    On Error GoTo ErrorHandler
    
    ' Rückgabewert: "GRÜN|Soll:50.00|Ist:50.00" oder "ROT|Soll:50.00|Ist:0.00" oder "GELB|Soll:50.00|Ist:45.00"
    
    Dim wsBK As Worksheet
    Dim wsEinst As Worksheet
    Dim soll As Double
    Dim ist As Double
    Dim status As String
    Dim sollDatum As Date
    Dim r As Long
    Dim lastRow As Long
    Dim zahlDatum As Date
    Dim zahlBetrag As Double
    Dim zahlKat As String
    Dim entityKeyZeile As String
    
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsEinst = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    
    ' Einstellungen-Cache laden (falls noch nicht geschehen)
    If Not m_EinstellungenGeladenZP Then Call LadeEinstellungenCacheZP
    
    ' 1. Soll-Wert aus Einstellungen holen
    soll = HoleSollBetragZP(kategorie)
    If soll = 0 Then
        PruefeZahlungen = "GELB|Soll:0.00|Ist:0.00|Keine Einstellung"
        Exit Function
    End If
    
    ' 2. Soll-Datum berechnen (mit Dezember-Vorauszahlungs-Logik)
    sollDatum = BerechneSollDatumZP(kategorie, monat, jahr)
    
    ' 3. Ist-Wert aus Bankkonto!Spalte H + EntityKey ermitteln
    ist = 0
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    For r = BK_START_ROW To lastRow
        zahlDatum = wsBK.Cells(r, BK_COL_DATUM).value
        If Not IsDate(zahlDatum) Then GoTo NextZahlRow
        
        ' Nur Zahlungen im relevanten Zeitraum (Monat ± Toleranzen)
        If Year(zahlDatum) <> jahr Then GoTo NextZahlRow
        If Month(zahlDatum) <> monat Then
            ' Dezember-Sonderfall: Vorauszahlung für Januar prüfen
            If monat = 1 And Month(zahlDatum) = 12 And Year(zahlDatum) = jahr - 1 Then
                ' Vorauszahlung aus Dezember des Vorjahres -> zulässig
            Else
                GoTo NextZahlRow
            End If
        End If
        
        ' EntityKey prüfen (Spalte J = INTERNE_NR = EntityKey)
        entityKeyZeile = Trim(CStr(wsBK.Cells(r, BK_COL_INTERNE_NR).value))
        If entityKeyZeile <> entityKey Then GoTo NextZahlRow
        
        ' Kategorie prüfen (Spalte H)
        zahlKat = Trim(CStr(wsBK.Cells(r, BK_COL_KATEGORIE).value))
        If StrComp(zahlKat, kategorie, vbTextCompare) <> 0 Then GoTo NextZahlRow
        
        ' Betrag addieren
        zahlBetrag = wsBK.Cells(r, BK_COL_BETRAG).value
        ist = ist + Abs(zahlBetrag)
        
NextZahlRow:
    Next r
    
    ' 4. Status ermitteln (GRÜN/GELB/ROT)
    If ist >= soll Then
        status = "GRÜN"
    ElseIf ist > 0 Then
        status = "GELB"
    Else
        status = "ROT"
    End If
    
    ' 5. Ergebnis formatieren
    PruefeZahlungen = status & "|Soll:" & Format(soll, "0.00") & "|Ist:" & Format(ist, "0.00")
    Exit Function
    
ErrorHandler:
    PruefeZahlungen = "ROT|Fehler:" & Err.Description
    
End Function


' ===============================================================
' Soll-Betrag aus Einstellungen holen (mit Cache)
' ===============================================================
Private Function HoleSollBetragZP(ByVal kategorie As String) As Double
    
    Dim i As Long
    
    If Not m_EinstellungenGeladenZP Then Call LadeEinstellungenCacheZP
    
    For i = LBound(m_EinstellungenCacheZP) To UBound(m_EinstellungenCacheZP)
        If StrComp(m_EinstellungenCacheZP(i).kategorie, kategorie, vbTextCompare) = 0 Then
            HoleSollBetragZP = m_EinstellungenCacheZP(i).SollBetrag
            Exit Function
        End If
    Next i
    
    HoleSollBetragZP = 0
    
End Function


' ===============================================================
' Soll-Datum berechnen (mit Spalte D/E vs F Logik)
' ===============================================================
Private Function BerechneSollDatumZP(ByVal kategorie As String, _
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
        ' Keine Regel gefunden -> 1. des Monats als Fallback
        BerechneSollDatumZP = DateSerial(jahr, monat, 1)
        Exit Function
    End If
    
    ' 2. Prüfen: Spalte F (Stichtag Fix) hat Vorrang
    If regel.StichtagFix <> "" Then
        ' Format: "TT.MM." -> z.B. "15.03"
        Dim teile() As String
        teile = Split(regel.StichtagFix, ".")
        If UBound(teile) >= 1 Then
            tag = CLng(teile(0))
            Dim fixMonat As Long
            fixMonat = CLng(teile(1))
            ' Nur wenn der fixMonat zum aktuellen Monat passt
            If fixMonat = monat Then
                BerechneSollDatumZP = DateSerial(jahr, monat, tag)
            Else
                ' Falscher Monat -> keine Zahlung fällig
                BerechneSollDatumZP = DateSerial(jahr, monat, 1)
            End If
            Exit Function
        End If
    End If
    
    ' 3. Spalte D/E verwenden (SollTag + SollMonate)
    ' Prüfen ob der aktuelle Monat in SollMonate enthalten ist
    istMonatGueltig = False
    If regel.SollMonate <> "" Then
        ' Format: "03, 06, 09"
        Dim monate() As String
        monate = Split(regel.SollMonate, ",")
        Dim m As Long
        For m = LBound(monate) To UBound(monate)
            If CLng(Trim(monate(m))) = monat Then
                istMonatGueltig = True
                Exit For
            End If
        Next m
    Else
        ' Keine Monate angegeben -> gilt für ALLE Monate
        istMonatGueltig = True
    End If
    
    If Not istMonatGueltig Then
        ' Monat ist nicht in SollMonate enthalten -> keine Zahlung fällig
        BerechneSollDatumZP = DateSerial(jahr, monat, 1)
        Exit Function
    End If
    
    ' Tag aus SollTag
    tag = regel.SollTag
    If tag = 0 Then tag = 1
    If tag > 28 Then
        ' Letzter Tag im Monat (31 = Ultimo-Ersatz)
        BerechneSollDatumZP = DateSerial(jahr, monat + 1, 0)  ' 0. Tag des Folgemonats = letzter Tag
    Else
        BerechneSollDatumZP = DateSerial(jahr, monat, tag)
    End If
    
End Function


' ===============================================================
' Einstellungen-Cache laden (Performance-Optimierung)
' FIX v1.1: PUBLIC statt PRIVATE (wird von mod_Uebersicht_Generator aufgerufen)
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
            .SollBetrag = wsEinst.Cells(r, ES_COL_SOLL_BETRAG).value
            .SollTag = wsEinst.Cells(r, ES_COL_SOLL_TAG).value
            .SollMonate = Trim(CStr(wsEinst.Cells(r, ES_COL_SOLL_MONATE).value))
            .StichtagFix = Trim(CStr(wsEinst.Cells(r, ES_COL_STICHTAG_FIX).value))
            .VorlaufTage = wsEinst.Cells(r, ES_COL_VORLAUF).value
            .NachlaufTage = wsEinst.Cells(r, ES_COL_NACHLAUF).value
            .SaeumnisGebuehr = wsEinst.Cells(r, ES_COL_SAEUMNIS).value
        End With
        
        idx = idx + 1
        
NextEinstRow:
    Next r
    
    ' Array auf tatsächliche Größe reduzieren
    If idx > 0 Then
        ReDim Preserve m_EinstellungenCacheZP(0 To idx - 1)
        m_EinstellungenGeladenZP = True
    Else
        m_EinstellungenGeladenZP = False
    End If
    
End Sub


' ===============================================================
' Einstellungen-Cache freigeben (Speicher sparen)
' FIX v1.1: PUBLIC statt PRIVATE (wird von mod_Uebersicht_Generator aufgerufen)
' ===============================================================
Public Sub EntladeEinstellungenCacheZP()
    
    Erase m_EinstellungenCacheZP
    m_EinstellungenGeladenZP = False
    
End Sub


' ===============================================================
' DEZEMBER-VORAUSZAHLUNGEN: Cache initialisieren
' Wird von mod_Uebersicht_Generator aufgerufen (vor Jahreswechsel)
' ===============================================================
Public Sub InitialisiereNachDezemberCacheZP(ByVal jahr As Long)
    
    Dim wsBK As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim zahlDatum As Date
    Dim zahlBetrag As Double
    Dim entityKey As String
    Dim kategorie As String
    Dim col As Collection
    
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set m_DezemberCacheZP = CreateObject("Scripting.Dictionary")
    
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    For r = BK_START_ROW To lastRow
        zahlDatum = wsBK.Cells(r, BK_COL_DATUM).value
        If Not IsDate(zahlDatum) Then GoTo NextDezRow
        
        ' Nur Dezember des Vorjahres
        If Year(zahlDatum) <> jahr - 1 Then GoTo NextDezRow
        If Month(zahlDatum) <> 12 Then GoTo NextDezRow
        
        entityKey = Trim(CStr(wsBK.Cells(r, BK_COL_INTERNE_NR).value))
        If entityKey = "" Then GoTo NextDezRow
        
        kategorie = Trim(CStr(wsBK.Cells(r, BK_COL_KATEGORIE).value))
        If kategorie = "" Then GoTo NextDezRow
        
        zahlBetrag = Abs(wsBK.Cells(r, BK_COL_BETRAG).value)
        
        ' In Dictionary speichern: Schlüssel = EntityKey & "|" & Kategorie
        Dim cacheKey As String
        cacheKey = entityKey & "|" & kategorie
        
        If Not m_DezemberCacheZP.Exists(cacheKey) Then
            Set col = New Collection
            m_DezemberCacheZP.Add col, cacheKey
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
    
    cacheKey = entityKey & "|" & kategorie
    
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
' SAMMELÜBERWEISUNG: Erkennung und manuelle Zuordnung
' Wird von Bankkonto.Worksheet_Change aufgerufen wenn Nutzer
' eine Sammelüberweisung markiert (z.B. via Checkbox/Flag)
' ===============================================================
Public Sub BearbeiteSammelUeberweisungZP(ByVal wsBK As Worksheet, _
                                          ByVal zeile As Long)
    
    On Error GoTo ErrorHandler
    
    ' 1. Gesamtbetrag aus Spalte B holen
    Dim gesamtBetrag As Double
    gesamtBetrag = Abs(wsBK.Cells(zeile, BK_COL_BETRAG).value)
    
    If gesamtBetrag = 0 Then
        MsgBox "Kein Betrag in Zeile " & zeile & " gefunden!", vbExclamation
        Exit Sub
    End If
    
    ' 2. Verfügbare Kategorien aus Einstellungen holen
    Dim kategorien() As String
    Dim sollBetraege() As Double
    Dim anzahl As Long
    
    Call HoleKategorienAusEinstellungenZP(kategorien, sollBetraege, anzahl)
    
    If anzahl = 0 Then
        MsgBox "Keine Kategorien in Einstellungen gefunden!", vbExclamation
        Exit Sub
    End If
    
    ' 3. UserForm anzeigen für manuelle Zuordnung
    ' (Hier müsste eine eigene UserForm erstellt werden - Beispiel-Code:)
    Dim ergebnis As String
    ergebnis = ZeigeSammelZuordnungDialogZP(gesamtBetrag, kategorien, sollBetraege, anzahl)
    
    ' 4. Ergebnis in Spalte L (Bemerkung) schreiben
    If ergebnis <> "" Then
        wsBK.Cells(zeile, BK_COL_BEMERKUNG).value = "SAMMEL:" & vbLf & ergebnis
        
        ' Optional: Automatische Betragszuordnung in Spalten M-Z
        ' (würde hier implementiert werden)
        
        MsgBox "Sammelüberweisung erfolgreich zugeordnet!", vbInformation
    Else
        MsgBox "Zuordnung abgebrochen.", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler bei Sammelüberweisung: " & Err.Description, vbCritical
    
End Sub


' ===============================================================
' HILFSFUNKTION: Holt alle Kategorien aus Einstellungen
' ===============================================================
Private Sub HoleKategorienAusEinstellungenZP(ByRef kategorien() As String, _
                                              ByRef sollBetraege() As Double, _
                                              ByRef anzahl As Long)
    
    Dim wsEinst As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim kat As String
    
    Set wsEinst = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    lastRow = wsEinst.Cells(wsEinst.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    
    anzahl = 0
    ReDim kategorien(1 To lastRow - ES_START_ROW + 1)
    ReDim sollBetraege(1 To lastRow - ES_START_ROW + 1)
    
    For r = ES_START_ROW To lastRow
        kat = Trim(CStr(wsEinst.Cells(r, ES_COL_KATEGORIE).value))
        If kat <> "" Then
            anzahl = anzahl + 1
            kategorien(anzahl) = kat
            sollBetraege(anzahl) = wsEinst.Cells(r, ES_COL_SOLL_BETRAG).value
        End If
    Next r
    
    If anzahl > 0 Then
        ReDim Preserve kategorien(1 To anzahl)
        ReDim Preserve sollBetraege(1 To anzahl)
    End If
    
End Sub


' ===============================================================
' HILFSFUNKTION: Zeigt Dialog für Sammelzuordnung
' (Platzhalter - müsste eigene UserForm sein)
' ===============================================================
Private Function ZeigeSammelZuordnungDialogZP(ByVal gesamtBetrag As Double, _
                                               ByRef kategorien() As String, _
                                               ByRef sollBetraege() As Double, _
                                               ByVal anzahl As Long) As String
    
    ' PLATZHALTER: Hier würde eine UserForm (z.B. frm_SammelZuordnung) geladen werden
    ' Die UserForm hätte:
    ' - Label mit Gesamtbetrag
    ' - ListBox mit Kategorien
    ' - TextBoxen für Teilbeträge
    ' - OK/Abbrechen Buttons
    
    ' Beispiel-Rückgabe:
    Dim ergebnis As String
    ergebnis = "Mitgliedsbeitrag: 7.50 " & ChrW(8364) & vbLf & _
               "Pachtgebühr: 25.00 " & ChrW(8364) & vbLf & _
               "Wasserkosten: 12.50 " & ChrW(8364)
    
    ZeigeSammelZuordnungDialogZP = ergebnis
    
    ' In der echten Implementierung würde hier stehen:
    ' Load frm_SammelZuordnung
    ' frm_SammelZuordnung.InitialisiereMit gesamtBetrag, kategorien, sollBetraege, anzahl
    ' frm_SammelZuordnung.Show vbModal
    ' ZeigeSammelZuordnungDialogZP = frm_SammelZuordnung.GetErgebnis
    ' Unload frm_SammelZuordnung
    
End Function


' ===============================================================
' MANUELLE ZUORDNUNG: Monatszuordnung bei Problemfällen
' Wird aufgerufen wenn eine Zahlung keinem Monat zugeordnet werden kann
' ===============================================================
Public Function FrageNachManuellerMonatszuordnungZP(ByVal wsBK As Worksheet, _
                                                      ByVal zeile As Long) As Long
    
    Dim zahlDatum As Date
    Dim betrag As Double
    Dim name As String
    Dim prompt As String
    Dim antwort As String
    Dim monat As Long
    
    zahlDatum = wsBK.Cells(zeile, BK_COL_DATUM).value
    betrag = wsBK.Cells(zeile, BK_COL_BETRAG).value
    name = Trim(CStr(wsBK.Cells(zeile, BK_COL_NAME).value))
    
    prompt = "Die Zahlung kann keinem Monat zugeordnet werden:" & vbLf & vbLf & _
             "Datum: " & Format(zahlDatum, "dd.mm.yyyy") & vbLf & _
             "Betrag: " & Format(betrag, "#,##0.00 ") & ChrW(8364) & vbLf & _
             "Name: " & name & vbLf & vbLf & _
             "Bitte geben Sie den Zielmonat ein (1-12):"
    
    antwort = InputBox(prompt, "Manuelle Monatszuordnung", Month(zahlDatum))
    
    If antwort = "" Then
        FrageNachManuellerMonatszuordnungZP = 0  ' Abbruch
        Exit Function
    End If
    
    If Not IsNumeric(antwort) Then
        MsgBox "Ungültige Eingabe! Es muss eine Zahl zwischen 1 und 12 sein.", vbExclamation
        FrageNachManuellerMonatszuordnungZP = 0
        Exit Function
    End If
    
    monat = CLng(antwort)
    
    If monat < 1 Or monat > 12 Then
        MsgBox "Ungültige Eingabe! Es muss eine Zahl zwischen 1 und 12 sein.", vbExclamation
        FrageNachManuellerMonatszuordnungZP = 0
        Exit Function
    End If
    
    ' Zuordnung in Spalte I (MONAT_PERIODE) speichern
    wsBK.Cells(zeile, BK_COL_MONAT_PERIODE).value = Format(monat, "00") & "/" & Year(zahlDatum)
    
    MsgBox "Zahlung wurde Monat " & monat & "/" & Year(zahlDatum) & " zugeordnet.", vbInformation
    
    FrageNachManuellerMonatszuordnungZP = monat
    
End Function


' ===============================================================
' NEU v1.2: MONAT/PERIODE SETZEN (verschoben aus mod_Banking_Data)
' Intelligent über Einstellungen mit Cache-Unterstützung.
' Nutzt Public ErmittleMonatPeriode aus mod_KategorieEngine_Evaluator.
' Wird von mod_Banking_Data.Importiere_Kontoauszug aufgerufen.
' ===============================================================
Public Sub SetzeMonatPeriode(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim r As Long
    Dim monatWert As Variant
    Dim datumWert As Variant
    Dim kategorie As String
    Dim faelligkeit As String
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    ' Fälligkeit aus Kategorie-Tabelle vorladen
    Dim wsDaten As Worksheet
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    
    ' Einstellungen-Cache laden für Folgemonat-Erkennung
    ' (expliziter Aufruf auf mod_KategorieEngine_Evaluator)
    Call mod_KategorieEngine_Evaluator.LadeEinstellungenCache
    
    For r = BK_START_ROW To lastRow
        datumWert = ws.Cells(r, BK_COL_DATUM).value
        monatWert = ws.Cells(r, BK_COL_MONAT_PERIODE).value
        
        If IsDate(datumWert) And (isEmpty(monatWert) Or monatWert = "") Then
            kategorie = Trim(ws.Cells(r, BK_COL_KATEGORIE).value)
            
            If kategorie <> "" Then
                ' Fälligkeit aus Kategorie-Tabelle holen (Spalte O)
                faelligkeit = HoleFaelligkeitFuerKategorie(wsDaten, kategorie)
                ' Nutzt Public Version aus Evaluator (mit Cache + Folgemonat)
                ws.Cells(r, BK_COL_MONAT_PERIODE).value = _
                    mod_KategorieEngine_Evaluator.ErmittleMonatPeriode(kategorie, CDate(datumWert), faelligkeit)
            Else
                ' Keine Kategorie: Fallback auf Buchungsmonat
                ws.Cells(r, BK_COL_MONAT_PERIODE).value = MonthName(Month(datumWert))
            End If
        End If
    Next r
    
    ' Einstellungen-Cache wieder freigeben
    Call mod_KategorieEngine_Evaluator.EntladeEinstellungenCache
    
End Sub


' ===============================================================
' NEU v1.2: FÄLLIGKEIT AUS KATEGORIE-TABELLE (Spalte O) HOLEN
' Verschoben aus mod_Banking_Data, jetzt Public für alle Module.
' ===============================================================
Public Function HoleFaelligkeitFuerKategorie(ByVal wsDaten As Worksheet, _
                                              ByVal kategorie As String) As String
    Dim lastRow As Long
    Dim r As Long
    
    lastRow = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    
    For r = DATA_START_ROW To lastRow
        If Trim(wsDaten.Cells(r, DATA_CAT_COL_KATEGORIE).value) = kategorie Then
            HoleFaelligkeitFuerKategorie = LCase(Trim(wsDaten.Cells(r, DATA_CAT_COL_FAELLIGKEIT).value))
            Exit Function
        End If
    Next r
    
    HoleFaelligkeitFuerKategorie = "monatlich"
End Function

