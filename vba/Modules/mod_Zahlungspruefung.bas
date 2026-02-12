Attribute VB_Name = "mod_Zahlungspruefung"
Option Explicit

' ***************************************************************
' MODUL: mod_Zahlungspruefung
' VERSION: 1.5 - 12.02.2026
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
' FIX v1.3: PruefeZahlungen komplett überarbeitet:
'           - EntityKey wird über Daten!R+S (IBAN) aufgelöst
'           - Bankkonto-Suche läuft über BK_COL_IBAN (Spalte D)
'           - Dezimalformat: Punkt als Trenner (systemunabhängig)
'           - IBAN-Cache (EntityKey -> IBAN) hinzugefügt
'           - Dezember-Cache: ebenfalls über IBAN statt INTERNE_NR
' NEU v1.4: SetzeMonatPeriode überarbeitet:
'           - Verarbeitet "GELB|Monatsname" Rückgabe aus
'             ErmittleMonatPeriode (Ultimo-5-Logik)
'           - GELB-Hintergrund in Spalte I bei unklaren Fällen
'           - DropDown-Liste (Januar-Dezember) auf ALLE Zellen
'             in Spalte I (ab BK_START_ROW)
'           - Übergibt wsBK + aktuelleZeile an ErmittleMonatPeriode
'             für den Lern-Mechanismus
'           - Entsperrt Spalten H, I, J, L für Nutzereingaben
'           - NEU: SetzeMonatDropDowns (Hilfsprozedur)
'           - NEU: EntsperreSpaltenFuerNutzer (Hilfsprozedur)
' FIX v1.5: SetzeMonatPeriode:
'           - Application.EnableEvents = False VOR dem Beschreiben
'             von Spalte I, damit Worksheet_Change NICHT getriggert
'             wird und der Lern-Vermerk NICHT in alle Zeilen
'             geschrieben wird ("Typen-Unverträglichkeit" behoben).
'           - Application.EnableEvents wird am Ende wieder auf True
'             gesetzt (mit sicherem Cleanup bei Fehler).
' ***************************************************************

' ===============================================================
' CACHE FÜR EINSTELLUNGEN (Performance-Optimierung)
' ===============================================================
Private Type EinstellungsRegelZP
    kategorie As String
    SollBetrag As Double
    SollTag As Long
    sollMonate As String           ' z.B. "03, 06, 09"
    StichtagFix As String          ' z.B. "15.03"
    VorlaufTage As Long
    NachlaufTage As Long
    SaeumnisGebuehr As Double
End Type

Private m_EinstellungenCacheZP() As EinstellungsRegelZP
Private m_EinstellungenGeladenZP As Boolean

' ===============================================================
' IBAN-CACHE: EntityKey -> IBAN (aus Daten!R+S)
' ===============================================================
Private m_EntityIBANCacheZP As Object   ' Dictionary: EntityKey -> IBAN
Private m_EntityIBANCacheGeladenZP As Boolean

' ===============================================================
' DEZEMBER-CACHE (für Vorauszahlungen)
' Struktur: Schlüssel = IBAN|Kategorie, Wert = Collection von Beträgen
' ===============================================================
Private m_DezemberCacheZP As Object   ' Dictionary mit IBAN|Kategorie -> Collection von Beträgen

' ===============================================================
' AMPELFARBEN (Konsistenz mit KategorieEngine)
' ===============================================================
Private Const AMPEL_GRUEN As Long = 12968900   ' RGB(196, 225, 196) -> hell-grün (Lern-Marker)
Private Const AMPEL_GELB As Long = 10086143    ' RGB(255, 235, 156) -> gelb (unklar)
Private Const AMPEL_ROT As Long = 9871103      ' RGB(255, 199, 206) -> rot

' Hell-grün für manuell bestätigte Monatszuordnung (Spalte I)
Private Const FARBE_HELLGRUEN_MANUELL As Long = 13565382  ' RGB(198, 239, 206)


' ===============================================================
' HAUPTFUNKTION: Prüft ALLE Zahlungen eines Mitglieds/einer Kategorie
' Wird von mod_Uebersicht_Generator aufgerufen
'
' Rückgabe: "STATUS|Soll:XX.XX|Ist:XX.XX"
'           Dezimaltrenner im Rückgabewert ist IMMER Punkt (.)
'           damit das Parsen systemunabhängig funktioniert.
'
' LOGIK v1.3:
'   1. EntityKey -> IBAN über Daten-Blatt (Spalte R+S) auflösen
'   2. Bankkonto nach IBAN (Spalte D) + Kategorie (Spalte H)
'      + Monat/Jahr (Spalte A) durchsuchen
'   3. Beträge (Spalte B) summieren -> Ist-Wert
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
    
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    ' Einstellungen-Cache laden (falls noch nicht geschehen)
    If Not m_EinstellungenGeladenZP Then Call LadeEinstellungenCacheZP
    
    ' IBAN-Cache laden (falls noch nicht geschehen)
    If Not m_EntityIBANCacheGeladenZP Then Call LadeEntityIBANCacheZP
    
    ' 1. IBAN zum EntityKey auflösen (über Daten!R+S)
    entityIBAN = ""
    If Not m_EntityIBANCacheZP Is Nothing Then
        If m_EntityIBANCacheZP.Exists(entityKey) Then
            entityIBAN = m_EntityIBANCacheZP(entityKey)
        End If
    End If
    
    If entityIBAN = "" Then
        ' Kein IBAN zum EntityKey gefunden -> keine Prüfung möglich
        PruefeZahlungen = "GELB|Soll:0.00|Ist:0.00|Keine IBAN zum EntityKey"
        Exit Function
    End If
    
    ' 2. Soll-Wert aus Einstellungen holen
    soll = HoleSollBetragZP(kategorie)
    If soll = 0 Then
        PruefeZahlungen = "GELB|Soll:0.00|Ist:0.00|Keine Einstellung"
        Exit Function
    End If
    
    ' 3. Ist-Wert aus Bankkonto ermitteln
    '    Suche über: IBAN (Spalte D) + Kategorie (Spalte H) + Monat/Jahr (Spalte A)
    ist = 0
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    For r = BK_START_ROW To lastRow
        ' Datum prüfen
        If Not IsDate(wsBK.Cells(r, BK_COL_DATUM).value) Then GoTo NextZahlRow
        zahlDatum = wsBK.Cells(r, BK_COL_DATUM).value
        
        ' Jahr prüfen
        If Year(zahlDatum) <> jahr Then
            ' Dezember-Sonderfall: Vorauszahlung aus Dezember des Vorjahres für Januar
            If monat = 1 And Month(zahlDatum) = 12 And Year(zahlDatum) = jahr - 1 Then
                ' Vorauszahlung aus Dezember des Vorjahres -> zulässig
            Else
                GoTo NextZahlRow
            End If
        End If
        
        ' Monat prüfen (nur wenn Jahr passt)
        If Year(zahlDatum) = jahr Then
            If Month(zahlDatum) <> monat Then GoTo NextZahlRow
        End If
        
        ' IBAN prüfen (Spalte D = BK_COL_IBAN)
        ibanZeile = Replace(Trim(CStr(wsBK.Cells(r, BK_COL_IBAN).value)), " ", "")
        If StrComp(ibanZeile, entityIBAN, vbTextCompare) <> 0 Then GoTo NextZahlRow
        
        ' Kategorie prüfen (Spalte H = BK_COL_KATEGORIE)
        zahlKat = Trim(CStr(wsBK.Cells(r, BK_COL_KATEGORIE).value))
        If StrComp(zahlKat, kategorie, vbTextCompare) <> 0 Then GoTo NextZahlRow
        
        ' Betrag addieren (Spalte B = BK_COL_BETRAG)
        zahlBetrag = wsBK.Cells(r, BK_COL_BETRAG).value
        ist = ist + Abs(zahlBetrag)
        
NextZahlRow:
    Next r
    
    ' 4. Status ermitteln (GRÜN/GELB/ROT)
    If ist >= soll Then
        status = "GR" & ChrW(220) & "N"
    ElseIf ist > 0 Then
        status = "GELB"
    Else
        status = "ROT"
    End If
    
    ' 5. Ergebnis formatieren (IMMER Punkt als Dezimaltrenner!)
    PruefeZahlungen = status & "|Soll:" & FormatDezimalPunkt(soll) & "|Ist:" & FormatDezimalPunkt(ist)
    Exit Function
    
ErrorHandler:
    PruefeZahlungen = "ROT|Fehler:" & Err.Description
    
End Function


' ===============================================================
' HILFSFUNKTION: Double -> String mit Punkt als Dezimaltrenner
' Wird intern verwendet, damit das Parsen systemunabhängig
' funktioniert (deutsch: Komma -> Punkt).
' ===============================================================
Private Function FormatDezimalPunkt(ByVal wert As Double) As String
    Dim s As String
    s = Format(wert, "0.00")
    ' Lokales Dezimalkomma durch Punkt ersetzen
    s = Replace(s, ",", ".")
    FormatDezimalPunkt = s
End Function


' ===============================================================
' IBAN-CACHE: Lädt EntityKey -> IBAN Zuordnung aus Daten!R+S
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
            ' 1 EntityKey = genau 1 IBAN
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
    If regel.sollMonate <> "" Then
        ' Format: "03, 06, 09"
        Dim monate() As String
        monate = Split(regel.sollMonate, ",")
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
            .sollMonate = Trim(CStr(wsEinst.Cells(r, ES_COL_SOLL_MONATE).value))
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
' v1.3: Gibt auch IBAN-Cache frei
' ===============================================================
Public Sub EntladeEinstellungenCacheZP()
    
    Erase m_EinstellungenCacheZP
    m_EinstellungenGeladenZP = False
    
    ' IBAN-Cache ebenfalls freigeben
    Call EntladeEntityIBANCacheZP
    
End Sub


' ===============================================================
' DEZEMBER-VORAUSZAHLUNGEN: Cache initialisieren
' Wird von mod_Uebersicht_Generator aufgerufen (vor Jahreswechsel)
' v1.3: Suche über IBAN (Spalte D) statt INTERNE_NR (Spalte J)
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
        
        ' Nur Dezember des Vorjahres
        If Year(zahlDatum) <> jahr - 1 Then GoTo NextDezRow
        If Month(zahlDatum) <> 12 Then GoTo NextDezRow
        
        ' IBAN aus Spalte D (statt EntityKey aus Spalte J)
        ibanWert = Replace(Trim(CStr(wsBK.Cells(r, BK_COL_IBAN).value)), " ", "")
        If ibanWert = "" Then GoTo NextDezRow
        
        kategorie = Trim(CStr(wsBK.Cells(r, BK_COL_KATEGORIE).value))
        If kategorie = "" Then GoTo NextDezRow
        
        zahlBetrag = Abs(wsBK.Cells(r, BK_COL_BETRAG).value)
        
        ' In Dictionary speichern: Schlüssel = IBAN & "|" & Kategorie
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
' v1.3: Parameter geändert: entityKey -> wird intern zu IBAN aufgelöst
' ===============================================================
Public Function HoleDezemberVorauszahlungZP(ByVal entityKey As String, _
                                             ByVal kategorie As String) As Double
    
    Dim cacheKey As String
    Dim col As Collection
    Dim summe As Double
    Dim v As Variant
    Dim entityIBAN As String
    
    ' EntityKey -> IBAN auflösen
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
    Dim Name As String
    Dim prompt As String
    Dim antwort As String
    Dim monat As Long
    
    zahlDatum = wsBK.Cells(zeile, BK_COL_DATUM).value
    betrag = wsBK.Cells(zeile, BK_COL_BETRAG).value
    Name = Trim(CStr(wsBK.Cells(zeile, BK_COL_NAME).value))
    
    prompt = "Die Zahlung kann keinem Monat zugeordnet werden:" & vbLf & vbLf & _
             "Datum: " & Format(zahlDatum, "dd.mm.yyyy") & vbLf & _
             "Betrag: " & Format(betrag, "#,##0.00 ") & ChrW(8364) & vbLf & _
             "Name: " & Name & vbLf & vbLf & _
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
' v1.5: MONAT/PERIODE SETZEN (überarbeitet)
' Intelligent über Einstellungen mit Cache-Unterstützung.
' Nutzt Public ErmittleMonatPeriode aus mod_KategorieEngine_Evaluator.
' Wird von mod_Banking_Data.Importiere_Kontoauszug aufgerufen.
'
' FIX v1.5: Application.EnableEvents = False VOR dem Beschreiben
'           von Spalte I (und anderen Spalten), damit
'           Worksheet_Change NICHT getriggert wird.
'           Dadurch werden folgende Fehler behoben:
'           1. "Typen-Unverträglichkeit" bei der Übersichts-Erstellung
'           2. "Folgemonat manuell bestätigt" steht NICHT mehr
'              in jeder Zeile, sondern NUR wo der Nutzer manuell
'              bestätigt hat (über das DropDown in Spalte I bei
'              GELB markierten Zellen).
'
' NEU v1.4:
'   - Verarbeitet "GELB|Monatsname" Rückgabe (Ultimo-5-Logik)
'   - Setzt GELB-Hintergrund in Spalte I bei unklaren Fällen
'   - Übergibt wsBK + aktuelleZeile für Lern-Mechanismus
'   - Setzt DropDown-Listen (Januar-Dezember) auf ALLE Zellen
'     in Spalte I (ab BK_START_ROW)
'   - Entsperrt Spalten H, I, J, L für Nutzereingaben
' ===============================================================
Public Sub SetzeMonatPeriode(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim r As Long
    Dim monatWert As Variant
    Dim datumWert As Variant
    Dim kategorie As String
    Dim faelligkeit As String
    Dim ergebnis As String
    
    ' v1.5 FIX: Vorherigen Zustand von EnableEvents merken und sicher abschalten
    Dim eventsWaren As Boolean
    eventsWaren = Application.EnableEvents
    
    On Error GoTo SetzeMonatPeriodeError
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    ' =============================================
    ' v1.5 FIX: Events ABSCHALTEN bevor Spalte I
    ' beschrieben wird! Sonst triggert jedes .value =
    ' den Worksheet_Change -> VerarbeiteMonatAenderung
    ' -> Lern-Vermerk in JEDER Zeile + Typ-Fehler!
    ' =============================================
    Application.EnableEvents = False
    
    ' Fälligkeit aus Kategorie-Tabelle vorladen
    Dim wsDaten As Worksheet
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    
    ' Einstellungen-Cache laden für Folgemonat-Erkennung
    ' (expliziter Aufruf auf mod_KategorieEngine_Evaluator)
    Call mod_KategorieEngine_Evaluator.LadeEinstellungenCache
    
    For r = BK_START_ROW To lastRow
        datumWert = ws.Cells(r, BK_COL_DATUM).value
        monatWert = ws.Cells(r, BK_COL_MONAT_PERIODE).value
        
        If IsDate(datumWert) And (isEmpty(monatWert) Or CStr(monatWert) = "") Then
            kategorie = Trim(CStr(ws.Cells(r, BK_COL_KATEGORIE).value))
            
            If kategorie <> "" Then
                ' Fälligkeit aus Kategorie-Tabelle holen (Spalte O)
                faelligkeit = HoleFaelligkeitFuerKategorie(wsDaten, kategorie)
                
                ' Nutzt Public Version aus Evaluator (mit Cache + Folgemonat + Lern-Check)
                ' v1.4: Übergibt jetzt wsBK + aktuelleZeile für den Lern-Mechanismus
                ergebnis = mod_KategorieEngine_Evaluator.ErmittleMonatPeriode( _
                    kategorie, CDate(datumWert), faelligkeit, ws, r)
                
                ' =============================================
                ' v1.4: "GELB|Monatsname" Rückgabe verarbeiten
                ' Wenn ErmittleMonatPeriode unsicher ist (Ultimo-5-Bereich,
                ' kein Lernmuster), kommt "GELB|Januar" zurück.
                ' -> Spalte I: Monatsname setzen + GELB hinterlegen
                ' =============================================
                If Left(ergebnis, 5) = "GELB|" Then
                    ' GELB-Fall: Monat extrahieren und setzen
                    Dim monatName As String
                    monatName = Mid(ergebnis, 6) ' alles nach "GELB|"
                    
                    ws.Cells(r, BK_COL_MONAT_PERIODE).value = monatName
                    
                    ' GELB-Hintergrund setzen (= Nutzer soll prüfen)
                    ws.Cells(r, BK_COL_MONAT_PERIODE).Interior.color = RGB(255, 235, 156)
                    
                    ' Bemerkung: Hinweis anhängen
                    Dim bestehendeBemerkung As String
                    bestehendeBemerkung = Trim(CStr(ws.Cells(r, BK_COL_BEMERKUNG).value))
                    
                    Dim gelbHinweis As String
                    gelbHinweis = "Ultimo-5: Bitte pr" & ChrW(252) & "fen ob Zahlung f" & ChrW(252) & "r " & _
                                  monatName & " oder Folgemonat gilt"
                    
                    If bestehendeBemerkung = "" Then
                        ws.Cells(r, BK_COL_BEMERKUNG).value = gelbHinweis
                    Else
                        ws.Cells(r, BK_COL_BEMERKUNG).value = bestehendeBemerkung & vbLf & gelbHinweis
                    End If
                Else
                    ' Normaler Fall: Monat direkt setzen (kein GELB)
                    ws.Cells(r, BK_COL_MONAT_PERIODE).value = ergebnis
                End If
            Else
                ' Keine Kategorie: Fallback auf Buchungsmonat
                ws.Cells(r, BK_COL_MONAT_PERIODE).value = MonthName(Month(datumWert))
            End If
        End If
    Next r
    
    ' Einstellungen-Cache wieder freigeben
    Call mod_KategorieEngine_Evaluator.EntladeEinstellungenCache
    
    ' =============================================
    ' v1.4: DropDown-Listen auf ALLE Zellen in Spalte I setzen
    ' (Januar bis Dezember) - auch auf bereits befüllte Zellen
    ' =============================================
    Call SetzeMonatDropDowns(ws, lastRow)
    
    ' =============================================
    ' v1.4: Spalten H, I, J, L entsperren für Nutzereingaben
    ' (Locked = False, damit der Nutzer trotz Blattschutz
    '  Kategorien, Monate, Spalte J und Bemerkungen ändern kann)
    ' =============================================
    Call EntsperreSpaltenFuerNutzer(ws, lastRow)
    
    ' v1.5 FIX: Events wieder einschalten
    Application.EnableEvents = eventsWaren
    Exit Sub

SetzeMonatPeriodeError:
    ' v1.5 FIX: Bei Fehler Events SICHER wieder einschalten
    Application.EnableEvents = eventsWaren
    ' Fehler nicht verschlucken - Debug-Info ausgeben
    Debug.Print "Fehler in SetzeMonatPeriode: " & Err.Number & " - " & Err.Description
    
End Sub


' ===============================================================
' v1.4: DropDown-Listen (Januar-Dezember) auf Spalte I setzen
' Setzt Data Validation auf alle Zellen in Spalte I
' von BK_START_ROW bis lastRow.
' Bestehende DropDowns werden überschrieben (kein Schaden,
' da die Liste immer gleich ist).
' Bereits manuell bestätigte Werte (hell-grün) bleiben erhalten,
' der Nutzer kann sie aber jederzeit über das DropDown ändern.
' ===============================================================
Private Sub SetzeMonatDropDowns(ByVal ws As Worksheet, ByVal lastRow As Long)
    
    If lastRow < BK_START_ROW Then Exit Sub
    
    ' Monatsliste als kommaseparierter String für Data Validation
    Dim monatsListe As String
    monatsListe = "Januar,Februar,M" & ChrW(228) & "rz,April,Mai,Juni," & _
                  "Juli,August,September,Oktober,November,Dezember"
    
    Dim rngMonat As Range
    Set rngMonat = ws.Range(ws.Cells(BK_START_ROW, BK_COL_MONAT_PERIODE), _
                            ws.Cells(lastRow, BK_COL_MONAT_PERIODE))
    
    On Error Resume Next
    With rngMonat.Validation
        .Delete   ' Bestehende Validation entfernen
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertInformation, _
             Operator:=xlBetween, _
             Formula1:=monatsListe
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = False  ' Keine Fehlermeldung bei freier Eingabe
        '                     (z.B. "Q1 2026", "Jahresbeitrag 2026" etc.)
    End With
    On Error GoTo 0
    
End Sub


' ===============================================================
' v1.4: Spalten H, I, J, L entsperren für Nutzereingaben
' Setzt Locked = False auf die Zellen in den angegebenen Spalten,
' damit der Nutzer trotz Blattschutz (UserInterfaceOnly:=True)
' folgende Spalten bearbeiten kann:
'   H = Kategorie (BK_COL_KATEGORIE)
'   I = Monat/Periode (BK_COL_MONAT_PERIODE)
'   J = Interne Nr (BK_COL_INTERNE_NR) - Spalte 10
'   L = Bemerkung (BK_COL_BEMERKUNG)
' ===============================================================
Private Sub EntsperreSpaltenFuerNutzer(ByVal ws As Worksheet, ByVal lastRow As Long)
    
    If lastRow < BK_START_ROW Then Exit Sub
    
    On Error Resume Next
    
    ' Spalte H (Kategorie)
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_KATEGORIE), _
             ws.Cells(lastRow, BK_COL_KATEGORIE)).Locked = False
    
    ' Spalte I (Monat/Periode)
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_MONAT_PERIODE), _
             ws.Cells(lastRow, BK_COL_MONAT_PERIODE)).Locked = False
    
    ' Spalte J (Interne Nr) = Spalte 10
    ws.Range(ws.Cells(BK_START_ROW, 10), _
             ws.Cells(lastRow, 10)).Locked = False
    
    ' Spalte L (Bemerkung)
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_BEMERKUNG), _
             ws.Cells(lastRow, BK_COL_BEMERKUNG)).Locked = False
    
    On Error GoTo 0
    
End Sub


' ===============================================================
' v1.4: FÄLLIGKEIT AUS KATEGORIE-TABELLE (Spalte O) HOLEN
' Verschoben aus mod_Banking_Data, jetzt Public für alle Module.
' v1.3: Prüft zuerst Einstellungen-Blatt (Spalte B = Kategorie),
'       dann erst Daten-Blatt (Spalte O = Fälligkeit) als Fallback.
' ===============================================================
Public Function HoleFaelligkeitFuerKategorie(ByVal wsDaten As Worksheet, _
                                              ByVal kategorie As String) As String
    Dim lastRow As Long
    Dim r As Long
    
    ' PRIO 1: Einstellungen-Blatt prüfen (Spalte B = Kategorie)
    '         Wenn Kategorie dort existiert, ist sie als "monatlich" zu werten
    '         (Einstellungen definieren die Zahlungstermine)
    Dim wsEinst As Worksheet
    On Error Resume Next
    Set wsEinst = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    
    If Not wsEinst Is Nothing Then
        lastRow = wsEinst.Cells(wsEinst.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
        For r = ES_START_ROW To lastRow
            If StrComp(Trim(CStr(wsEinst.Cells(r, ES_COL_KATEGORIE).value)), kategorie, vbTextCompare) = 0 Then
                ' Kategorie in Einstellungen gefunden
                ' Fälligkeit aus SollMonate ableiten
                Dim sollMonate As String
                sollMonate = Trim(CStr(wsEinst.Cells(r, ES_COL_SOLL_MONATE).value))
                If sollMonate = "" Then
                    ' Keine Monate angegeben -> gilt für ALLE Monate = monatlich
                    HoleFaelligkeitFuerKategorie = "monatlich"
                Else
                    ' Spezifische Monate angegeben
                    Dim anzMonate As Long
                    anzMonate = UBound(Split(sollMonate, ",")) + 1
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
    
    ' PRIO 2: Fallback auf Daten-Blatt (Spalte O = Fälligkeit)
    lastRow = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    
    For r = DATA_START_ROW To lastRow
        If Trim(wsDaten.Cells(r, DATA_CAT_COL_KATEGORIE).value) = kategorie Then
            HoleFaelligkeitFuerKategorie = LCase(Trim(wsDaten.Cells(r, DATA_CAT_COL_FAELLIGKEIT).value))
            Exit Function
        End If
    Next r
    
    HoleFaelligkeitFuerKategorie = "monatlich"
End Function

