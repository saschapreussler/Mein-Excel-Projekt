Public Function HoleFaelligkeitFuerKategorie(ByVal wsEinst As Worksheet, ByVal kategorie As String) As String
    ' v2.1: Zuerst Fälligkeitsspalte O auf Blatt "Daten" prüfen
    Dim faellDaten As String
    faellDaten = ""
    Dim wsDatenZP As Worksheet
    On Error Resume Next
    Set wsDatenZP = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    If Not wsDatenZP Is Nothing Then
        Dim lastRuleRowZP As Long
        lastRuleRowZP = wsDatenZP.Cells(wsDatenZP.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
        Dim rZP As Long
        For rZP = DATA_START_ROW To lastRuleRowZP
            If StrComp(Trim(CStr(wsDatenZP.Cells(rZP, DATA_CAT_COL_KATEGORIE).value)), kategorie, vbTextCompare) = 0 Then
                faellDaten = LCase(Trim(CStr(wsDatenZP.Cells(rZP, DATA_CAT_COL_FAELLIGKEIT).value)))
                Exit For
            End If
        Next rZP
    End If
    ' Wenn in Spalte O ein spezieller Typ steht, diesen zurückgeben
    If faellDaten Like "*hrlich (jahr/folgejahr)*" Or _
       faellDaten = "j" & ChrW(228) & "hrlich (jahr/folgejahr)" Then
        HoleFaelligkeitFuerKategorie = "j" & ChrW(228) & "hrlich (jahr/folgejahr)"
        Exit Function
    ElseIf faellDaten Like "*hrlich (jahr)*" Or _
           faellDaten = "j" & ChrW(228) & "hrlich (jahr)" Then
        HoleFaelligkeitFuerKategorie = "j" & ChrW(228) & "hrlich (jahr)"
        Exit Function
    End If
    ' Fallback: Bisherige Logik über SollMonate
    Dim SollMonate As String
    SollMonate = Trim(CStr(wsEinst.Cells(2, ES_COL_SOLL_MONATE).value)) ' Annahme: Zeile 2 als Beispiel, ggf. anpassen
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
End Function
Attribute VB_Name = "mod_Zahlungspruefung"
Option Explicit

' ***************************************************************
' MODUL: mod_Zahlungspruefung
' VERSION: 2.0 - 12.02.2026
' ZWECK: Zahlungspruefung fuer Mitgliederliste + Einstellungen
'        - Prueft Zahlungseingaenge gegen Soll-Werte
'        - Behandelt Dezember-Vorauszahlungen
'        - Erkennt Sammelueberweisungen
'        - Bietet manuelle Zuordnung bei Problemfaellen
'        - Dokumentiert Aufschluesselung in Spalte L
'        - SetzeMonatPeriode (verschoben aus mod_Banking_Data)
'        - HoleFaelligkeitFuerKategorie (verschoben aus mod_Banking_Data)
' NEU v2.0:
'   - SetzeBankkontoDropDowns: Oeffentliche Prozedur die ALLE
'     DropDowns (H + I) setzt. Wird von Worksheet_Activate aufgerufen.
'   - SetzeKategorieDropDowns: DropDown in Spalte H (Bankkonto)
'     dynamisch aus Hilfsspalten AF/AG auf Blatt "Daten"
'   - AktualisiereKategorieHilfsspalten: Befuellt Spalte AF + AG
'     auf Blatt "Daten" mit eindeutigen Kategorienamen (E/A getrennt)
'   - PruefeZahlungen: Bei Soll=0 (variabler Betrag) wird trotzdem
'     geprueft ob eine Zahlung eingegangen ist (nicht mehr Abbruch)
'   - SetzeMonatDropDowns + EntsperreSpaltenFuerNutzer: Blattschutz
'     wird korrekt aufgehoben und danach wieder gesetzt
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

' Hell-gruen fuer manuell bestaetigte Monatszuordnung (Spalte I)
Private Const FARBE_HELLGRUEN_MANUELL As Long = 13565382


' ===============================================================
' v2.0 NEU: OEFFENTLICH: Setzt ALLE DropDowns auf dem Bankkonto-Blatt
' Wird von Tabelle3.Worksheet_Activate UND nach CSV-Import aufgerufen.
' Setzt:
'   - Spalte H (Kategorie): E- oder A-Kategorien je nach Betrag
'   - Spalte I (Monat/Periode): Januar bis Dezember
'   - Entsperrt editierbare Spalten (H, I, J, L)
' ===============================================================
Public Sub SetzeBankkontoDropDowns(ByVal wsBK As Worksheet)
    
    Dim lastRow As Long
    
    If wsBK Is Nothing Then Exit Sub
    
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    ' Hilfsspalten auf Daten-Blatt aktualisieren (AF + AG)
    Call AktualisiereKategorieHilfsspalten
    
    ' Blattschutz aufheben (noetig fuer Data Validation)
    On Error Resume Next
    wsBK.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' DropDowns setzen
    Call SetzeKategorieDropDowns(wsBK, lastRow)
    Call SetzeMonatDropDowns(wsBK, lastRow)
    
    ' Spalten entsperren fuer Nutzereingaben
    Call EntsperreSpaltenFuerNutzer(wsBK, lastRow)
    
    ' Blattschutz wieder aktivieren
    On Error Resume Next
    wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
End Sub


' ===============================================================
' v2.0 NEU: Befuellt Hilfsspalten AF (32) + AG (33) auf Blatt "Daten"
' mit eindeutigen Kategorienamen, getrennt nach E und A.
' AF = Einnahmen-Kategorien (K = "E")
' AG = Ausgaben-Kategorien (K = "A")
' Quelle: Spalte J (DATA_CAT_COL_KATEGORIE = 10)
'         Spalte K (DATA_CAT_COL_EINAUS = 11)
' ===============================================================
Public Sub AktualisiereKategorieHilfsspalten()
    
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim katName As String
    Dim einAus As String
    
    Dim dictE As Object
    Dim dictA As Object
    
    Set dictE = CreateObject("Scripting.Dictionary")
    Set dictA = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    
    On Error Resume Next
    wsDaten.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    lastRow = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then GoTo ProtectAndExit
    
    ' Eindeutige Kategorien sammeln
    For r = DATA_START_ROW To lastRow
        katName = Trim(CStr(wsDaten.Cells(r, DATA_CAT_COL_KATEGORIE).value))
        If katName = "" Then GoTo NextHilfsRow
        
        einAus = UCase(Trim(CStr(wsDaten.Cells(r, DATA_CAT_COL_EINAUS).value)))
        
        If einAus = "E" Then
            If Not dictE.Exists(katName) Then dictE.Add katName, katName
        ElseIf einAus = "A" Then
            If Not dictA.Exists(katName) Then dictA.Add katName, katName
        End If
        
NextHilfsRow:
    Next r
    
    ' Hilfsspalten leeren (ab Zeile 4, max 200 Zeilen sicherheitshalber)
    Dim maxClear As Long
    maxClear = wsDaten.Cells(wsDaten.Rows.count, DATA_COL_KAT_EINNAHMEN).End(xlUp).Row
    If maxClear < DATA_START_ROW + 200 Then maxClear = DATA_START_ROW + 200
    
    wsDaten.Range(wsDaten.Cells(DATA_START_ROW, DATA_COL_KAT_EINNAHMEN), _
                  wsDaten.Cells(maxClear, DATA_COL_KAT_EINNAHMEN)).ClearContents
    wsDaten.Range(wsDaten.Cells(DATA_START_ROW, DATA_COL_KAT_AUSGABEN), _
                  wsDaten.Cells(maxClear, DATA_COL_KAT_AUSGABEN)).ClearContents
    
    ' Einnahmen in Spalte AF (DATA_COL_KAT_EINNAHMEN = 32) schreiben
    Dim idx As Long
    idx = DATA_START_ROW
    Dim key As Variant
    For Each key In dictE.keys
        wsDaten.Cells(idx, DATA_COL_KAT_EINNAHMEN).value = CStr(key)
        idx = idx + 1
    Next key
    
    ' Ausgaben in Spalte AG (DATA_COL_KAT_AUSGABEN = 33) schreiben
    idx = DATA_START_ROW
    For Each key In dictA.keys
        wsDaten.Cells(idx, DATA_COL_KAT_AUSGABEN).value = CStr(key)
        idx = idx + 1
    Next key
    
ProtectAndExit:
    On Error Resume Next
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
    Set dictE = Nothing
    Set dictA = Nothing
    
End Sub


' ===============================================================
' v2.0 NEU: Setzt DropDown-Listen in Spalte H (Kategorie)
' Fuer jede Zeile: Betrag > 0 -> Einnahmen (AF), Betrag < 0 -> Ausgaben (AG)
' Referenziert dynamisch auf den befuellten Bereich in AF bzw. AG
' ===============================================================
Private Sub SetzeKategorieDropDowns(ByVal ws As Worksheet, ByVal lastRow As Long)
    
    If lastRow < BK_START_ROW Then Exit Sub
    
    Dim wsDaten As Worksheet
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    
    ' Letzter befuellter Eintrag in AF und AG ermitteln
    Dim lastE As Long
    lastE = wsDaten.Cells(wsDaten.Rows.count, DATA_COL_KAT_EINNAHMEN).End(xlUp).Row
    If lastE < DATA_START_ROW Then lastE = DATA_START_ROW
    
    Dim lastA As Long
    lastA = wsDaten.Cells(wsDaten.Rows.count, DATA_COL_KAT_AUSGABEN).End(xlUp).Row
    If lastA < DATA_START_ROW Then lastA = DATA_START_ROW
    
    ' Spaltenbuchstaben fuer Validation-Formeln berechnen
    Dim spalteBuchstabeE As String
    spalteBuchstabeE = SpalteNrZuBuchstabe(DATA_COL_KAT_EINNAHMEN)
    
    Dim spalteBuchstabeA As String
    spalteBuchstabeA = SpalteNrZuBuchstabe(DATA_COL_KAT_AUSGABEN)
    
    ' Daten-Blattname fuer Formel
    Dim datenName As String
    datenName = wsDaten.Name
    
    ' Validation-Formeln: =Daten!$AF$4:$AF$xx
    Dim formelEinnahmen As String
    formelEinnahmen = "=" & datenName & "!$" & spalteBuchstabeE & "$" & DATA_START_ROW & _
                      ":$" & spalteBuchstabeE & "$" & lastE
    
    Dim formelAusgaben As String
    formelAusgaben = "=" & datenName & "!$" & spalteBuchstabeA & "$" & DATA_START_ROW & _
                     ":$" & spalteBuchstabeA & "$" & lastA
    
    ' Pro Zeile die passende Validation setzen
    Dim r As Long
    Dim betrag As Double
    Dim formel As String
    
    On Error Resume Next
    
    For r = BK_START_ROW To lastRow
        betrag = 0
        If IsNumeric(ws.Cells(r, BK_COL_BETRAG).value) Then
            betrag = CDbl(ws.Cells(r, BK_COL_BETRAG).value)
        End If
        
        If betrag > 0 Then
            formel = formelEinnahmen
        ElseIf betrag < 0 Then
            formel = formelAusgaben
        Else
            ' Betrag = 0 oder leer: Einnahmen als Default
            formel = formelEinnahmen
        End If
        
        With ws.Cells(r, BK_COL_KATEGORIE).Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertInformation, _
                 Operator:=xlBetween, _
                 Formula1:=formel
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = False
        End With
    Next r
    
    On Error GoTo 0
    
End Sub


' ===============================================================
' v2.0 NEU: Hilfsfunktion: Spaltennummer -> Spaltenbuchstabe
' (1="A", 26="Z", 27="AA", 28="AB", 32="AF", 33="AG" etc.)
' ===============================================================
Private Function SpalteNrZuBuchstabe(ByVal spalte As Long) As String
    Dim temp As String
    temp = ""
    Do While spalte > 0
        Dim rest As Long
        rest = (spalte - 1) Mod 26
        temp = Chr(65 + rest) & temp
        spalte = (spalte - 1) \ 26
    Loop
    SpalteNrZuBuchstabe = temp
End Function


' ===============================================================
' HAUPTFUNKTION: Prueft ALLE Zahlungen eines Mitglieds/einer Kategorie
' Wird von mod_Uebersicht_Generator aufgerufen
'
' Rueckgabe: "STATUS|Soll:XX.XX|Ist:XX.XX"
'           Dezimaltrenner im Rueckgabewert ist IMMER Punkt (.)
'
' v2.0 FIX: Bei Soll=0 (variabler Betrag, z.B. "Strom/Wasser
'           Abschlagszahlung") wird NICHT mehr abgebrochen, sondern
'           es wird geprueft ob eine Zahlung eingegangen ist.
'           Status: GRUEN wenn Ist>0, ROT wenn Ist=0
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
    ' Alte Logik war:
    '   If soll = 0 Then PruefeZahlungen = "GELB|..." : Exit Function
    ' Das ist ENTFERNT, damit auch Kategorien ohne festen Soll-Betrag
    ' (z.B. "Strom/Wasser Abschlagszahlung") geprueft werden.
    
    ' 3. Ist-Wert aus Bankkonto ermitteln
    ist = 0
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    For r = BK_START_ROW To lastRow
        ' Datum pruefen
        If Not IsDate(wsBK.Cells(r, BK_COL_DATUM).value) Then GoTo NextZahlRow
        zahlDatum = wsBK.Cells(r, BK_COL_DATUM).value
        
        ' Jahr pruefen
        If Year(zahlDatum) <> jahr Then
            ' Dezember-Sonderfall: Vorauszahlung Dezember Vorjahr fuer Januar
            If monat = 1 And Month(zahlDatum) = 12 And Year(zahlDatum) = jahr - 1 Then
                ' Vorauszahlung aus Dezember des Vorjahres -> zulaessig
            Else
                GoTo NextZahlRow
            End If
        End If
        
        ' Monat pruefen (nur wenn Jahr passt)
        If Year(zahlDatum) = jahr Then
            If Month(zahlDatum) <> monat Then GoTo NextZahlRow
        End If
        
        ' IBAN pruefen (Spalte D = BK_COL_IBAN)
        ibanZeile = Replace(Trim(CStr(wsBK.Cells(r, BK_COL_IBAN).value)), " ", "")
        If StrComp(ibanZeile, entityIBAN, vbTextCompare) <> 0 Then GoTo NextZahlRow
        
        ' Kategorie pruefen (Spalte H = BK_COL_KATEGORIE)
        zahlKat = Trim(CStr(wsBK.Cells(r, BK_COL_KATEGORIE).value))
        If StrComp(zahlKat, kategorie, vbTextCompare) <> 0 Then GoTo NextZahlRow
        
        ' Betrag addieren (Spalte B = BK_COL_BETRAG)
        zahlBetrag = wsBK.Cells(r, BK_COL_BETRAG).value
        ist = ist + Abs(zahlBetrag)
        
NextZahlRow:
    Next r
    
    ' 4. Status ermitteln (GRUEN/GELB/ROT)
    '    v2.0: Unterscheidung fester Soll vs. variabler Soll
    If soll > 0 Then
        ' Fester Soll-Betrag vorhanden: Betrags-Vergleich
        If ist >= soll Then
            status = "GR" & ChrW(220) & "N"
        ElseIf ist > 0 Then
            status = "GELB"
        Else
            status = "ROT"
        End If
    Else
        ' Kein fester Soll-Betrag (variabel): nur Eingangs-Pruefung
        ' GRUEN wenn Zahlung eingegangen, ROT wenn nicht
        If ist > 0 Then
            status = "GR" & ChrW(220) & "N"
        Else
            status = "ROT"
        End If
    End If
    
    ' 5. Ergebnis formatieren (IMMER Punkt als Dezimaltrenner!)
    PruefeZahlungen = status & "|Soll:" & FormatDezimalPunkt(soll) & "|Ist:" & FormatDezimalPunkt(ist)
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
                .SaeumnisGebuehr = CDbl(wsEinst.Cells(r, ES_COL_SAEUMNIS).value)
            Else
                .SaeumnisGebuehr = 0
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
' SAMMELUEBERWEISUNGEN: Erkennung und manuelle Zuordnung
' ===============================================================
Public Sub BearbeiteSammelUeberweisungZP(ByVal wsBK As Worksheet, _
                                          ByVal zeile As Long)
    
    On Error GoTo ErrorHandler
    
    Dim gesamtBetrag As Double
    gesamtBetrag = Abs(wsBK.Cells(zeile, BK_COL_BETRAG).value)
    
    If gesamtBetrag = 0 Then
        MsgBox "Kein Betrag in Zeile " & zeile & " gefunden!", vbExclamation
        Exit Sub
    End If
    
    Dim kategorien() As String
    Dim sollBetraege() As Double
    Dim anzahl As Long
    
    Call HoleKategorienAusEinstellungenZP(kategorien, sollBetraege, anzahl)
    
    If anzahl = 0 Then
        MsgBox "Keine Kategorien in Einstellungen gefunden!", vbExclamation
        Exit Sub
    End If
    
    Dim ergebnis As String
    ergebnis = ZeigeSammelZuordnungDialogZP(gesamtBetrag, kategorien, sollBetraege, anzahl)
    
    If ergebnis <> "" Then
        wsBK.Cells(zeile, BK_COL_BEMERKUNG).value = "SAMMEL:" & vbLf & ergebnis
        MsgBox "Sammel" & ChrW(252) & "berweisung erfolgreich zugeordnet!", vbInformation
    Else
        MsgBox "Zuordnung abgebrochen.", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler bei Sammel" & ChrW(252) & "berweisung: " & Err.Description, vbCritical
    
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
' HILFSFUNKTION: Zeigt Dialog fuer Sammelzuordnung (Platzhalter)
' ===============================================================
Private Function ZeigeSammelZuordnungDialogZP(ByVal gesamtBetrag As Double, _
                                               ByRef kategorien() As String, _
                                               ByRef sollBetraege() As Double, _
                                               ByVal anzahl As Long) As String
    
    Dim ergebnis As String
    ergebnis = "Mitgliedsbeitrag: 7,50 " & ChrW(8364) & vbLf & _
               "Pachtgeb" & ChrW(252) & "hr: 25,00 " & ChrW(8364) & vbLf & _
               "Wasserkosten: 12,50 " & ChrW(8364)
    
    ZeigeSammelZuordnungDialogZP = ergebnis
    
End Function


' ===============================================================
' MANUELLE ZUORDNUNG: Monatszuordnung bei Problemfaellen
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
        FrageNachManuellerMonatszuordnungZP = 0
        Exit Function
    End If
    
    If Not IsNumeric(antwort) Then
        MsgBox "Ung" & ChrW(252) & "ltige Eingabe! Es muss eine Zahl zwischen 1 und 12 sein.", vbExclamation
        FrageNachManuellerMonatszuordnungZP = 0
        Exit Function
    End If
    
    monat = CLng(antwort)
    
    If monat < 1 Or monat > 12 Then
        MsgBox "Ung" & ChrW(252) & "ltige Eingabe! Es muss eine Zahl zwischen 1 und 12 sein.", vbExclamation
        FrageNachManuellerMonatszuordnungZP = 0
        Exit Function
    End If
    
    wsBK.Cells(zeile, BK_COL_MONAT_PERIODE).value = Format(monat, "00") & "/" & Year(zahlDatum)
    
    MsgBox "Zahlung wurde Monat " & monat & "/" & Year(zahlDatum) & " zugeordnet.", vbInformation
    
    FrageNachManuellerMonatszuordnungZP = monat
    
End Function


' ===============================================================
' v1.5: MONAT/PERIODE SETZEN (ueberarbeitet)
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
    Call mod_KategorieEngine_Zeitraum.LadeEinstellungenCache
    
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
                Else
                    ws.Cells(r, BK_COL_MONAT_PERIODE).value = ergebnis
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
    Call SetzeBankkontoDropDowns(ws)
    
    Exit Sub

SetzeMonatPeriodeError:
    Application.EnableEvents = eventsWaren
    Debug.Print "Fehler in SetzeMonatPeriode: " & Err.Number & " - " & Err.Description
    
End Sub


' ===============================================================
' v2.1: DropDown-Listen auf Spalte I setzen
' Enth�lt jetzt: Januar-Dezember + dynamische j�hrliche Eintr�ge
' aus der Kategorie-Tabelle (F�lligkeit "j�hrlich (jahr)" und
' "j�hrlich (jahr/folgejahr)")
' ===============================================================
Private Sub SetzeMonatDropDowns(ByVal ws As Worksheet, ByVal lastRow As Long)
    
    If lastRow < BK_START_ROW Then Exit Sub
    
    ' Abrechnungsjahr aus Startmen�!F1 lesen
    Dim abrJahr As Long
    On Error Resume Next
    abrJahr = CLng(ThisWorkbook.Worksheets("Startmen" & ChrW(252)).Range("F1").value)
    On Error GoTo 0
    If abrJahr = 0 Then abrJahr = Year(Date)
    
    ' Basis: Januar bis Dezember
    Dim monatsListe As String
    monatsListe = "Januar,Februar,M" & ChrW(228) & "rz,April,Mai,Juni," & _
                  "Juli,August,September,Oktober,November,Dezember"
    
    ' Dynamische Eintr�ge aus Kategorie-Tabelle (Daten!J+O)
    Dim wsDaten As Worksheet
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If Not wsDaten Is Nothing Then
        Dim lastRuleRow As Long
        lastRuleRow = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
        
        ' Dictionary f�r Eindeutigkeit
        Dim dictExtra As Object
        Set dictExtra = CreateObject("Scripting.Dictionary")
        
        Dim r As Long
        Dim katName As String
        Dim katFaell As String
        Dim extraEintrag As String
        
        For r = DATA_START_ROW To lastRuleRow
            katName = Trim(CStr(wsDaten.Cells(r, DATA_CAT_COL_KATEGORIE).value))
            katFaell = LCase(Trim(CStr(wsDaten.Cells(r, DATA_CAT_COL_FAELLIGKEIT).value)))
            
            If katName = "" Then GoTo NextDDRow
            
            extraEintrag = ""
            
            ' "j�hrlich (jahr/folgejahr)" -> "Kategoriename Jahr/Folgejahr"
            If katFaell Like "*hrlich (jahr/folgejahr)*" Or _
               katFaell = "j" & ChrW(228) & "hrlich (jahr/folgejahr)" Then
                extraEintrag = katName & " " & abrJahr & "/" & (abrJahr + 1)
            
            ' "j�hrlich (jahr)" -> "Kategoriename Jahr"
            ElseIf katFaell Like "*hrlich (jahr)*" Or _
                   katFaell = "j" & ChrW(228) & "hrlich (jahr)" Then
                extraEintrag = katName & " " & abrJahr
            End If
            
            If extraEintrag <> "" Then
                If Not dictExtra.Exists(extraEintrag) Then
                    dictExtra.Add extraEintrag, True
                End If
            End If
NextDDRow:
        Next r
        
        ' Zus�tzliche Eintr�ge anh�ngen
        Dim k As Variant
        For Each k In dictExtra.keys
            monatsListe = monatsListe & "," & CStr(k)
        Next k
        
        ' "Sammelzahlung" als festen Eintrag hinzuf�gen
        monatsListe = monatsListe & ",Sammelzahlung"
        
        ' "j�hrlich" als Fallback-Eintrag hinzuf�gen
        monatsListe = monatsListe & ",j" & ChrW(228) & "hrlich"
        
        Set dictExtra = Nothing
    End If
    
    Dim rngMonat As Range
    Set rngMonat = ws.Range(ws.Cells(BK_START_ROW, BK_COL_MONAT_PERIODE), _
                            ws.Cells(lastRow, BK_COL_MONAT_PERIODE))
    
    On Error Resume Next
    With rngMonat.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertInformation, _
             Operator:=xlBetween, _
             Formula1:=monatsListe
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = False
    End With
    On Error GoTo 0
    
End Sub

' ===============================================================
' v1.4: Spalten H, I, J, L entsperren fuer Nutzereingaben
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
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_INTERNE_NR), _
             ws.Cells(lastRow, BK_COL_INTERNE_NR)).Locked = False
    
    ' Spalte L (Bemerkung)
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_BEMERKUNG), _
             ws.Cells(lastRow, BK_COL_BEMERKUNG)).Locked = False
    
    On Error GoTo 0
    
End Sub


                ' v2.1: Zuerst F�lligkeitsspalte O auf Blatt "Daten" pr�fen
                '       (dort stehen die neuen Typen "j�hrlich (jahr)" etc.)
                Dim faellDaten As String
                faellDaten = ""
                
                Dim wsDatenZP As Worksheet
                On Error Resume Next
                Set wsDatenZP = ThisWorkbook.Worksheets(WS_DATEN)
                On Error GoTo 0
                
                If Not wsDatenZP Is Nothing Then
                    Dim lastRuleRowZP As Long
                    lastRuleRowZP = wsDatenZP.Cells(wsDatenZP.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
                    Dim rZP As Long
                    For rZP = DATA_START_ROW To lastRuleRowZP
                        If StrComp(Trim(CStr(wsDatenZP.Cells(rZP, DATA_CAT_COL_KATEGORIE).value)), kategorie, vbTextCompare) = 0 Then
                            faellDaten = LCase(Trim(CStr(wsDatenZP.Cells(rZP, DATA_CAT_COL_FAELLIGKEIT).value)))
                            Exit For
                        End If
                    Next rZP
                End If
                
                ' Wenn in Spalte O ein spezieller Typ steht, diesen zur�ckgeben
                If faellDaten Like "*hrlich (jahr/folgejahr)*" Or _
                   faellDaten = "j" & ChrW(228) & "hrlich (jahr/folgejahr)" Then
                    HoleFaelligkeitFuerKategorie = "j" & ChrW(228) & "hrlich (jahr/folgejahr)"
                    Exit Function
                ElseIf faellDaten Like "*hrlich (jahr)*" Or _
                       faellDaten = "j" & ChrW(228) & "hrlich (jahr)" Then
                    HoleFaelligkeitFuerKategorie = "j" & ChrW(228) & "hrlich (jahr)"
                    Exit Function
                End If
                
                ' Fallback: Bisherige Logik �ber SollMonate
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

