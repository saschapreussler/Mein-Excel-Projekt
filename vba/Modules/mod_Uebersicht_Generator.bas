Attribute VB_Name = "mod_Uebersicht_Generator"
Option Explicit

' ***************************************************************
' MODUL: mod_Uebersicht_Generator
' VERSION: 4.6 - 15.03.2026
' ZWECK: Generiert ?bersichtsblatt (Variante 2: Lange Tabelle)
'        - 14 Mitglieder (Parzellen 1-14)
'        - Kategorien DYNAMISCH aus Einstellungen-Blatt (Spalte B)
'        - Zeigt Soll/Ist/Status f?r jede Kombination
'        - Behandelt SHARE-Keys (Gemeinschaftskonten) korrekt
'        - Bei Kategorien OHNE festen Soll-Betrag:
'          Soll-Zelle bleibt leer + hell-gelb + editierbar
'          Nur Zahlungstermin-Pr?fung (p?nktlich / S?umnis)
'        - S?umnis-Geb?hren werden in Bemerkung angezeigt
' FIX v1.1: InitialisiereNachDezemberCache -> InitialisiereNachDezemberCacheZP
' FIX v1.2: Val() statt CDbl() f?r systemunabh?ngiges Parsen
' FIX v1.3: "Typen unvertr?glich" behoben (Variant, StrComp, etc.)
' FIX v1.4: ChrW() in Const nicht erlaubt -> Private Variablen
' NEU v2.0: Kategorien DYNAMISCH aus Einstellungen-Blatt
'           - Keine hart kodierten Kategorienamen mehr
'           - Soll-Betrag 0 -> Zelle leer + hell-gelb + editierbar
'           - Zahlungstermin-Pr?fung auch ohne Soll-Betrag
'           - S?umnis-Geb?hren in Bemerkung
' NEU v3.0: HoleAktiveMitglieder liest jetzt aus Daten-Blatt
'           (EntityKey-Tabelle R-W) statt aus Mitgliederliste
'           - SHARE-Keys: Parzelle "2, 5" wird aufgeteilt
'           - stummModus f?r automatische Aufrufe (keine MsgBox)
'           - Trigger: Bankkonto H/I + Einstellungen -> auto-Update
' NEU v4.0: Monatsweise Bef?llung der ?bersicht
'           - Nur Monate mit importierten CSV-Daten werden angezeigt
'           - ErmittleImportierteMonate() scannt Bankkonto Spalte A
'           - Eintrag erscheint nur wenn:
'             a) Zahlung vorhanden (Ist > 0) -> GR?N/GELB
'             b) Frist abgelaufen + keine Zahlung -> ROT
'           - Einheitliches Datumsformat: "Januar 2026"
' NEU v4.1: F?lligkeit-basierte Kategoriefilterung
'           - Kategorien erscheinen nur im F?lligkeitsmonat
'             (nicht mehr in allen 12 Monaten)
'           - F?lligkeit aus Daten Spalte O (Kategorie-Tabelle)
'           - Vorjahr-Speicher (Daten Spalten CA-CF):
'             Okt-Dez Zahlungen des Vorjahres f?r Jan-M?rz Zuordnung
'           - Spalte C linksb?ndig, Format "M?rz 2025"
'           - PruefeZahlungen: flexibler Perioden-Vergleich
' SPLIT v4.2: Datenquellen + Vorjahr-Speicher ausgelagert nach
'             mod_Uebersicht_Daten (LadeKategorienAusEinstellungen,
'             HoleAktiveMitglieder, Ermittle*, Vorjahr-Speicher)
' NEU v4.3: - Spalte C NumberFormat "@" (Text) verhindert
'             Excel-Datumserkennung ("Jan 25" -> "Januar 2025")
'           - Variabler Soll: Folgemonat-Uebernahme aus Vormonat
'             HoleManuellSollAusVormonat() nutzt gesicherte Werte
'             Gelb->Gruen wenn Ist = manueller Soll
'           - Zebra basiert auf sichtbaren Zeilen (nach Filter)
' NEU v4.4: - Ehrenmitglied: kein Mitgliedsbeitrag noetig
'           - Partner-Mitgliedsbeitrag: Doppelzahlung erkennen
'             (15 EUR statt 7.50 EUR -> deckt beide Parzellen-Mitgl.)
'           - Zahlungsdatum in Bemerkung (Spalte H)
'           - AutoFilter-Dropdowns immer auf Zeile 3
'           - Gruppenblock entfernt (Filter-kompatibel)
' NEU v4.5: - Parzelle-basierte Kategorien: Pacht, Fixkosten,
'             Abschlagszahlungen, Endabrechnung, Betriebskosten
'             erscheinen nur 1x pro Parzelle (nicht pro Mitglied)
'           - Nur Mitgliedsbeitrag wird pro Mitglied angezeigt
'           - Worksheet_Change Event: Gelb->Gruen bei manueller
'             Soll-Eingabe + MsgBox fuer Folgemonat-Uebernahme
' FIX v4.5b: - Reentrancy-Schutz (m_IsGenerating)
'            - AutoFilter VOR ClearContents entfernen
'            - Robustere Fehlerbehandlung (Blattschutz im ErrorHandler)
' NEU v4.6: - Ehrenmitglied-Rolle: ErmittleEntityRoleVonFunktion
'             erkennt jetzt "Ehrenmitglied" aus Mitgliederliste Spalte O
'           - Januar-Schutz: Wenn keine Vorjahr-Daten vorhanden,
'             wird ROT auf GELB herabgestuft statt falsche Saeumnis
'             (Dezember-Zahlung des Vorjahres koennte fehlen)
' ***************************************************************

' ===============================================================
' KONSTANTEN
' ===============================================================
Private Const UEBERSICHT_START_ROW As Long = 4
Private Const UEBERSICHT_HEADER_ROW As Long = 3

' Spalten im ?bersichtsblatt
Private Const UEB_COL_PARZELLE As Long = 1      ' A - Parzelle
Private Const UEB_COL_MITGLIED As Long = 2      ' B - Mitglied
Private Const UEB_COL_MONAT As Long = 3         ' C - Monat
Private Const UEB_COL_KATEGORIE As Long = 4     ' D - Kategorie
Private Const UEB_COL_SOLL As Long = 5          ' E - Soll
Private Const UEB_COL_IST As Long = 6           ' F - Ist
Private Const UEB_COL_STATUS As Long = 7        ' G - Status (GR?N/GELB/ROT)
Private Const UEB_COL_BEMERKUNG As Long = 8     ' H - Bemerkung
Private Const UEB_COL_SUMME_IST As Long = 9     ' I - Summe Ist (kumuliert)

' Farbe fuer Summen-Spalte (dezentes Hellblau)
Private Const FARBE_SUMME As Long = 16247773    ' RGB(189, 215, 248)
Private Const FARBE_SUMME_ZEBRA As Long = 15790320 ' RGB(208, 228, 241)

' Ampelfarben
Private Const AMPEL_GRUEN As Long = 12968900    ' RGB(196, 225, 196)
Private Const AMPEL_GELB As Long = 10086143     ' RGB(255, 235, 156)
Private Const AMPEL_ROT As Long = 9871103       ' RGB(255, 199, 206)

' Hell-gelb f?r "bitte manuell bef?llen" (Soll-Betrag variabel)
Private Const FARBE_HELLGELB_MANUELL As Long = 10092543  ' RGB(255, 255, 153)

' Zebra-Farbe (identisch mit Bankkonto / EntityKey-Tabelle)
Private Const ZEBRA_COLOR As Long = &HDEE5E3

' Status-String f?r GR?N (Encoding-sicher, wird in Init gesetzt)
Private m_STATUS_GRUEN As String
Private m_StatusInitialisiert As Boolean

' v4.5b: Reentrancy-Schutz (verhindert doppelten Aufruf)
Private m_IsGenerating As Boolean


' ===============================================================
' Type f?r eine dynamische Kategorie aus Einstellungen
' ===============================================================
Public Type UebKategorie
    Name As String
    SollBetrag As Double
    HatFestenSoll As Boolean      ' True wenn Spalte C > 0
    saeumnisGebuehr As Double     ' Spalte I auf Einstellungen
    SollMonate As String          ' Spalte E: "03, 06, 09" oder leer = alle
    faelligkeit As String         ' Spalte O auf Daten: "monatlich", "jaehrlich" etc.
End Type


' ===============================================================
' Initialisiert Status-String (Encoding-sicher)
' ===============================================================
Private Sub InitStatus()
    
    If m_StatusInitialisiert Then Exit Sub
    
    m_STATUS_GRUEN = "GR" & ChrW(220) & "N"
    m_StatusInitialisiert = True
    
End Sub


' ===============================================================
' v4.5b: Gibt zurueck ob GeneriereUebersicht gerade laeuft.
' Wird von Workbook_SheetChange geprueft um Events waehrend
' der Generierung zu ignorieren.
' ===============================================================
Public Function IsGenerating() As Boolean
    IsGenerating = m_IsGenerating
End Function


' ===============================================================
' HAUPTFUNKTION: Generiert komplettes ?bersichtsblatt
' v2.0: Kategorien DYNAMISCH aus Einstellungen-Blatt
' v3.0: stummModus f?r automatische Aufrufe (ohne MsgBox)
' ===============================================================
Public Sub GeneriereUebersicht(Optional ByVal jahr As Long = 0, _
                                Optional ByVal stummModus As Boolean = False)
    
    On Error GoTo ErrorHandler
    
    ' v4.5b: Reentrancy-Schutz - Doppelten Aufruf verhindern
    If m_IsGenerating Then
        Debug.Print "[" & ChrW(220) & "bersicht] WARNUNG: GeneriereUebersicht bereits aktiv - " & _
                    "Aufruf ignoriert (Reentrancy-Schutz)"
        Exit Sub
    End If
    m_IsGenerating = True
    
    ' Status initialisieren (Encoding-sicher)
    Call InitStatus
    
    Dim wsUeb As Worksheet
    Dim wsDaten As Worksheet
    Dim startTime As Double
    Dim monat As Long
    Dim kategorie As String
    Dim mitglieder As Collection
    Dim mitglied As Object
    Dim entityKey As String
    Dim ergebnis As String
    Dim teile() As String
    Dim soll As Double
    Dim ist As Double
    Dim status As String
    Dim rowIdx As Long
    
    startTime = Timer
    
    ' Jahr-Parameter validieren
    ' v5.1: Prim?r aus Startmen?!F1 lesen, Abgleich mit Bankkonto-Daten
    If jahr = 0 Then
        ' 1. Abrechnungsjahr aus Startmen?!F1 lesen
        Dim wsStart As Worksheet
        On Error Resume Next
        Set wsStart = ThisWorkbook.Worksheets("Startmen" & ChrW(252))
        On Error GoTo ErrorHandler
        
        Dim jahrF1 As Long
        jahrF1 = 0
        If Not wsStart Is Nothing Then
            If IsNumeric(wsStart.Range("F1").value) Then
                jahrF1 = CLng(wsStart.Range("F1").value)
            End If
        End If
        
        ' 2. Jahr aus Bankkonto-Daten ermitteln (h?ufigstes Jahr)
        Dim jahrBK As Long
        jahrBK = mod_Uebersicht_Daten.ErmittleJahrAusBankkonto()
        
        ' 3. Entscheidungslogik
        If jahrF1 > 0 And jahrBK > 0 Then
            If jahrF1 = jahrBK Then
                ' Gleich -> verwenden
                jahr = jahrF1
            Else
                If Not stummModus Then
                    ' Abweichung -> Nutzer fragen
                    Dim antwortJahr As VbMsgBoxResult
                    antwortJahr = MsgBox("Das Abrechnungsjahr in Startmen" & ChrW(252) & "!F1 ist " & _
                                         jahrF1 & "," & vbLf & _
                                         "aber die meisten Kontoausz" & ChrW(252) & "ge stammen aus " & _
                                         jahrBK & "." & vbLf & vbLf & _
                                         "Soll die " & ChrW(220) & "bersicht f" & ChrW(252) & "r " & _
                                         jahrF1 & " (Startmen" & ChrW(252) & ") erstellt werden?" & vbLf & _
                                         "  Ja = " & jahrF1 & " (Startmen" & ChrW(252) & "!F1)" & vbLf & _
                                         "  Nein = " & jahrBK & " (Kontoausz" & ChrW(252) & "ge)", _
                                         vbQuestion + vbYesNo, "Abrechnungsjahr")
                    If antwortJahr = vbYes Then
                        jahr = jahrF1
                    Else
                        jahr = jahrBK
                    End If
                Else
                    ' stummModus: Startmen?!F1 hat Vorrang
                    jahr = jahrF1
                    Debug.Print "[" & ChrW(220) & "bersicht] HINWEIS: F1=" & jahrF1 & _
                                " vs Bankkonto=" & jahrBK & " -> verwende F1"
                End If
            End If
        ElseIf jahrF1 > 0 Then
            jahr = jahrF1
        ElseIf jahrBK > 0 Then
            jahr = jahrBK
        Else
            jahr = Year(Date)
        End If
    End If
    Debug.Print "[" & ChrW(220) & "bersicht] Verwende Jahr: " & jahr
    
    ' =============================================
    ' v2.0: Kategorien DYNAMISCH aus Einstellungen laden
    ' =============================================
    Dim kategorien() As UebKategorie
    Dim anzahlKat As Long
    Call mod_Uebersicht_Daten.LadeKategorienAusEinstellungen(kategorien, anzahlKat)
    
    If anzahlKat = 0 Then
        If Not stummModus Then
            MsgBox "Keine Kategorien im Einstellungen-Blatt (Spalte B) gefunden!" & vbLf & _
                   "Bitte mindestens eine Kategorie mit Zahlungstermin anlegen.", _
                   vbCritical, "Fehler"
        End If
        m_IsGenerating = False
        Exit Sub
    End If
    
    ' Worksheets holen
    On Error Resume Next
    Set wsUeb = ThisWorkbook.Worksheets(WS_UEBERSICHT)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo ErrorHandler
    
    If wsUeb Is Nothing Or wsDaten Is Nothing Then
        If Not stummModus Then
            MsgBox "Blatt '" & ChrW(220) & "bersicht' oder 'Daten' nicht gefunden!", vbCritical
        End If
        m_IsGenerating = False
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Blatt entsperren
    On Error Resume Next
    wsUeb.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    ' v4.3: Manuell eingetragene Soll-Werte SICHERN bevor Inhalt geloescht wird
    Dim gespeicherteSoll As Object
    Set gespeicherteSoll = SammleManuelleSollWerte(wsUeb)
    
    ' v4.5b: AutoFilter VORHER entfernen (verhindert Probleme mit gefilterten Zeilen)
    If wsUeb.AutoFilterMode Then wsUeb.AutoFilterMode = False
    
    ' Alten Inhalt l?schen (ab Zeile 4, inkl. Spalte I)
    wsUeb.Range(wsUeb.Cells(UEBERSICHT_START_ROW, 1), _
                wsUeb.Cells(wsUeb.Rows.count, UEB_COL_SUMME_IST)).ClearContents
    wsUeb.Range(wsUeb.Cells(UEBERSICHT_START_ROW, 1), _
                wsUeb.Cells(wsUeb.Rows.count, UEB_COL_SUMME_IST)).Interior.ColorIndex = xlNone
    
    ' Header setzen
    Call SetzeUebersichtHeader(wsUeb)
    
    ' Einstellungen-Cache laden (Performance)
    Call mod_Zahlungspruefung.LadeEinstellungenCacheZP
    
    ' Dezember-Cache initialisieren (f?r Vorauszahlungen)
    Call mod_Zahlungspruefung.InitialisiereNachDezemberCacheZP(jahr)
    
    ' v4.0: Vorjahr-Speicher bef?llen (Okt-Dez Vorjahr)
    Call mod_Uebersicht_Daten.BefuelleVorjahrSpeicher(jahr - 1)
    
    ' v4.0: Vorjahr-Speicher automatisch loeschen (ab August)
    Call mod_Uebersicht_Daten.PruefeVorjahrSpeicherAblauf
    
    ' Aktive Mitglieder aus Daten-Blatt EntityKey-Tabelle laden
    Set mitglieder = mod_Uebersicht_Daten.HoleAktiveMitglieder(wsDaten)
    
    ' Debug-Diagnose: Mitglieder und Kategorien protokollieren
    Debug.Print "[" & ChrW(220) & "bersicht] Kategorien: " & anzahlKat & _
                " | Mitglieder: " & mitglieder.count
    
    If mitglieder.count = 0 Then
        Debug.Print "[" & ChrW(220) & "bersicht] WARNUNG: Keine aktiven Mitglieder gefunden!"
        Debug.Print "[" & ChrW(220) & "bersicht] Pr" & ChrW(252) & "fe Daten-Blatt: " & _
                    "EntityKey (R), Parzelle (V), Role (W)"
        
        ' v4.5b: Blatt trotzdem schuetzen und aufraumen
        On Error Resume Next
        wsUeb.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        On Error GoTo 0
        
        m_IsGenerating = False
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        If Not stummModus Then
            MsgBox "Keine aktiven Mitglieder im Daten-Blatt gefunden!" & vbLf & vbLf & _
                   "Bitte sicherstellen dass:" & vbLf & _
                   "- Spalte R (EntityKey) bef" & ChrW(252) & "llt ist" & vbLf & _
                   "- Spalte V (Parzelle) eine Nummer 1-14 enth" & ChrW(228) & "lt" & vbLf & _
                   "- Spalte W (Role) 'MITGLIED MIT PACHT' oder 'MITGLIED OHNE PACHT' enth" & ChrW(228) & "lt", _
                   vbExclamation, ChrW(220) & "bersicht"
        End If
        Exit Sub
    End If
    
    ' =============================================
    ' v4.0: DATEN GENERIEREN - Nur relevante Eintraege!
    '
    ' Ein Eintrag erscheint NUR wenn:
    '   a) Zahlung vorhanden (Ist > 0) -> GRUEN/GELB
    '   b) Monat hat importierte CSV-Daten UND
    '      Frist abgelaufen UND keine Zahlung -> ROT
    '
    ' KEIN Eintrag wenn:
    '   - Monat hat keine CSV-Daten UND Frist nicht abgelaufen
    ' =============================================
    
    ' 1. Ermittle welche Monate Bankkonto-Buchungen haben
    Dim importierteMonate() As Boolean
    importierteMonate = mod_Uebersicht_Daten.ErmittleImportierteMonate(jahr)
    
    ' Debug: Importierte Monate anzeigen
    Dim dbgMonate As String
    dbgMonate = ""
    For monat = 1 To 12
        If importierteMonate(monat) Then
            If dbgMonate <> "" Then dbgMonate = dbgMonate & ", "
            dbgMonate = dbgMonate & MonthName(monat, True)
        End If
    Next monat
    Debug.Print "[" & ChrW(220) & "bersicht] Importierte Monate: " & _
                IIf(dbgMonate = "", "(keine)", dbgMonate)
    
    ' v4.5: Dictionary um zu tracken welche Parzelle+Monat+Kategorie-Kombinationen
    ' bereits geschrieben wurden (fuer Parzelle-basierte Kategorien)
    Dim geschriebeneParzKat As Object
    Set geschriebeneParzKat = CreateObject("Scripting.Dictionary")
    geschriebeneParzKat.CompareMode = vbTextCompare
    
    ' v5.0: Dictionary fuer kumulierte Ist-Summen pro Parzelle+Kategorie
    ' Key = "Parzelle|Kategorie", Value = kumulierter Ist-Betrag
    Dim summeIstDict As Object
    Set summeIstDict = CreateObject("Scripting.Dictionary")
    summeIstDict.CompareMode = vbTextCompare
    
    ' Sammelt Zeilennummern pro Parzelle+Kategorie fuer nachtraegliches Befuellen
    ' Key = "Parzelle|Kategorie", Value = Collection von Zeilennummern
    Dim summeIstZeilen As Object
    Set summeIstZeilen = CreateObject("Scripting.Dictionary")
    summeIstZeilen.CompareMode = vbTextCompare
    
    ' 2. Daten generieren
    rowIdx = UEBERSICHT_START_ROW
    
    For Each mitglied In mitglieder
        Dim parzelleWert As Variant
        parzelleWert = mitglied("Parzelle")
        entityKey = mitglied("EntityKey")
        Dim mitgliedName As String
        mitgliedName = mitglied("Name")
        Dim mitgliedRole As String
        mitgliedRole = mitglied("Role")
        
        For monat = 1 To 12
            Dim k As Long
            For k = 0 To anzahlKat - 1
                
                ' Pruefen ob diese Kategorie in diesem Monat faellig ist
                If Not IstKategorieImMonatFaellig(kategorien(k), monat) Then
                    GoTo NextKat
                End If
                
                kategorie = kategorien(k).Name
                
                ' v4.0: MITGLIED OHNE PACHT zahlt NUR Mitgliedsbeitrag
                ' Alle anderen Kategorien (Pacht, Betriebskosten, Fixkosten,
                ' Abschlagszahlungen Strom/Wasser etc.) ueberspringen
                If InStr(UCase(mitgliedRole), "OHNE PACHT") > 0 Then
                    If StrComp(kategorie, "Mitgliedsbeitrag", vbTextCompare) <> 0 Then
                        GoTo NextKat
                    End If
                End If
                
                ' v4.4: Ehrenmitglied zahlt keinen Mitgliedsbeitrag
                If InStr(UCase(mitgliedRole), "EHREN") > 0 Then
                    If StrComp(kategorie, "Mitgliedsbeitrag", vbTextCompare) = 0 Then
                        GoTo NextKat
                    End If
                End If
                
                ' v4.5: Parzelle-basierte Kategorien nur 1x pro Parzelle anzeigen
                ' Nur Mitgliedsbeitrag wird pro Mitglied angezeigt,
                ' alle anderen Kategorien (Pacht, Fixkosten, Abschlagszahlungen,
                ' Endabrechnung, Betriebskosten) nur 1x pro Parzelle.
                If StrComp(kategorie, "Mitgliedsbeitrag", vbTextCompare) <> 0 Then
                    Dim parzKatKey As String
                    parzKatKey = CStr(parzelleWert) & "|" & monat & "|" & kategorie
                    If geschriebeneParzKat.Exists(parzKatKey) Then
                        GoTo NextKat
                    End If
                End If
                
                ' Zahlung pruefen (mod_Zahlungspruefung)
                ergebnis = mod_Zahlungspruefung.PruefeZahlungen(entityKey, kategorie, monat, jahr)
                
                ' Ergebnis parsen: "GRUEN|Soll:50.00|Ist:50.00"
                soll = 0
                ist = 0
                status = "ROT"
                
                teile = Split(ergebnis, "|")
                If UBound(teile) >= 2 Then
                    status = teile(0)
                    
                    Dim sollTeile() As String
                    sollTeile = Split(teile(1), ":")
                    If UBound(sollTeile) >= 1 Then
                        soll = val(sollTeile(1))
                    End If
                    
                    Dim istTeile() As String
                    istTeile = Split(teile(2), ":")
                    If UBound(istTeile) >= 1 Then
                        ist = val(istTeile(1))
                    End If
                ElseIf UBound(teile) >= 0 Then
                    status = teile(0)
                End If
                
                ' v4.0: Vorjahr-Zahlungen pruefen (Jan-Maerz)
                ' Dezember-Zahlung des Vorjahres die fuer diesen Monat gilt
                If monat <= 3 And ist = 0 Then
                    Dim vjBetrag As Double
                    vjBetrag = mod_Uebersicht_Daten.HoleVorjahrZahlung(entityKey, kategorie, monat)
                    If vjBetrag > 0 Then
                        ist = ist + vjBetrag
                        ' Status aktualisieren
                        If soll > 0 And ist >= soll Then
                            status = m_STATUS_GRUEN
                        ElseIf ist > 0 Then
                            status = m_STATUS_GRUEN
                        End If
                    End If
                End If
                
                ' v4.6: Januar-Schutz: Wenn keine Vorjahr-Daten vorhanden sind,
                ' koennte eine Dezember-Zahlung des Vorjahres fuer Januar gelten.
                ' In diesem Fall ROT -> GELB herabstufen mit Hinweis.
                If monat = 1 And ist = 0 And Not mod_Uebersicht_Daten.HatVorjahrDaten() Then
                    If StrComp(status, "ROT", vbTextCompare) = 0 Then
                        status = "GELB"
                    End If
                End If
                
                ' v4.4: Partner-Zahlung pruefen bei Mitgliedsbeitrag
                ' Wenn ein Mitglied auf der gleichen Parzelle >= 2x Soll
                ' bezahlt hat, gilt der Beitrag als mitbezahlt
                Dim partnerInfo As String
                partnerInfo = ""
                If ist = 0 And soll > 0 And StrComp(kategorie, "Mitgliedsbeitrag", vbTextCompare) = 0 Then
                    partnerInfo = PruefePartnerMitgliedsbeitrag( _
                        mitglieder, CLng(parzelleWert), entityKey, monat, jahr, soll)
                    If partnerInfo <> "" Then
                        ist = soll
                        status = m_STATUS_GRUEN
                    End If
                End If
                
                ' =============================================
                ' v4.0: FILTER - Nur relevante Eintraege anzeigen
                ' ROT-Eintraege erscheinen NUR fuer Monate mit
                ' importierten Kontoauszuegen auf dem Bankkonto-Blatt.
                ' =============================================
                Dim zeigeEintrag As Boolean
                zeigeEintrag = False
                
                If ist > 0 Then
                    ' Fall a) Zahlung vorhanden -> IMMER anzeigen
                    zeigeEintrag = True
                Else
                    ' Fall b) Keine Zahlung -> nur anzeigen wenn:
                    '   - Monat hat CSV-Daten (importiert) UND
                    '   - Frist (SollDatum + Nachlauf) ist abgelaufen
                    If importierteMonate(monat) Then
                        Dim sollDatumUeb As Date
                        Dim vorlaufUeb As Long
                        Dim nachlaufUeb As Long
                        Dim saeumnisUeb As Double
                        
                        sollDatumUeb = mod_Zahlungspruefung.BerechneSollDatumZP(kategorie, monat, jahr)
                        Call mod_Zahlungspruefung.HoleToleranzZP(kategorie, vorlaufUeb, nachlaufUeb, saeumnisUeb)
                        
                        ' Frist abgelaufen = Heute >= SollDatum + Nachlauf
                        If Date >= DateAdd("d", nachlaufUeb, sollDatumUeb) Then
                            zeigeEintrag = True
                        End If
                    End If
                End If
                
                ' Wenn nicht relevant -> naechste Kategorie
                If Not zeigeEintrag Then GoTo NextKat
                
                ' Zeile schreiben
                wsUeb.Cells(rowIdx, UEB_COL_PARZELLE).value = parzelleWert
                wsUeb.Cells(rowIdx, UEB_COL_MITGLIED).value = mitgliedName
                ' v4.3: Monat als Text schreiben (verhindert Excel-Datumserkennung)
                wsUeb.Cells(rowIdx, UEB_COL_MONAT).NumberFormat = "@"
                wsUeb.Cells(rowIdx, UEB_COL_MONAT).value = MonthName(monat) & " " & jahr
                wsUeb.Cells(rowIdx, UEB_COL_KATEGORIE).value = kategorie
                
                ' =============================================
                ' v2.0: Soll-Betrag Logik
                ' =============================================
                If kategorien(k).HatFestenSoll Then
                    ' Fester Soll-Betrag aus Einstellungen
                    wsUeb.Cells(rowIdx, UEB_COL_SOLL).value = soll
                Else
                    ' KEIN fester Soll-Betrag -> Zelle hell-gelb (editierbar)
                    ' v4.3: Zuerst pruefen ob Nutzer bereits einen Betrag fuer
                    ' diese Kategorie+Parzelle in einem frueheren Monat gesetzt hat
                    Dim manuellSoll As Double
                    manuellSoll = HoleManuellSollAusVormonat(gespeicherteSoll, CStr(parzelleWert), kategorie)
                    
                    If manuellSoll > 0 Then
                        ' Folgemonat: Betrag aus Vormonat uebernehmen
                        wsUeb.Cells(rowIdx, UEB_COL_SOLL).value = manuellSoll
                        soll = manuellSoll
                        wsUeb.Cells(rowIdx, UEB_COL_SOLL).Locked = False
                        
                        ' Pruefen ob Ist den manuellen Soll erreicht
                        If ist > 0 And Abs(ist - manuellSoll) < 0.01 Then
                            status = m_STATUS_GRUEN
                            wsUeb.Cells(rowIdx, UEB_COL_SOLL).Interior.color = AMPEL_GRUEN
                        ElseIf ist > 0 Then
                            status = m_STATUS_GRUEN
                            wsUeb.Cells(rowIdx, UEB_COL_SOLL).Interior.color = FARBE_HELLGELB_MANUELL
                        Else
                            wsUeb.Cells(rowIdx, UEB_COL_SOLL).Interior.color = FARBE_HELLGELB_MANUELL
                        End If
                    Else
                        ' Erster Monat ohne Vorgabe -> leer + hell-gelb
                        wsUeb.Cells(rowIdx, UEB_COL_SOLL).value = ""
                        wsUeb.Cells(rowIdx, UEB_COL_SOLL).Interior.color = FARBE_HELLGELB_MANUELL
                        wsUeb.Cells(rowIdx, UEB_COL_SOLL).Locked = False
                    End If
                    
                    ' Status bei variablem Betrag: Termin-Pruefung
                    If ist > 0 Then
                        status = m_STATUS_GRUEN
                    End If
                End If
                
                wsUeb.Cells(rowIdx, UEB_COL_IST).value = ist
                wsUeb.Cells(rowIdx, UEB_COL_STATUS).value = status
                
                ' Farbe setzen
                If StrComp(status, m_STATUS_GRUEN, vbTextCompare) = 0 Then
                    wsUeb.Cells(rowIdx, UEB_COL_STATUS).Interior.color = AMPEL_GRUEN
                ElseIf StrComp(status, "GELB", vbTextCompare) = 0 Then
                    wsUeb.Cells(rowIdx, UEB_COL_STATUS).Interior.color = AMPEL_GELB
                Else
                    wsUeb.Cells(rowIdx, UEB_COL_STATUS).Interior.color = AMPEL_ROT
                End If
                
                ' =============================================
                ' v2.0: Bemerkung mit S?umnis-Info
                ' =============================================
                Dim bemerkung As String
                bemerkung = ""
                
                ' Zusatzinfo aus Ergebnis (4. Teil)
                If UBound(teile) >= 3 Then
                    bemerkung = teile(3)
                End If
                
                ' S?umnis-Geb?hr anh?ngen wenn Status ROT und Geb?hr definiert
                If StrComp(status, "ROT", vbTextCompare) = 0 Then
                    If kategorien(k).saeumnisGebuehr > 0 Then
                        Dim saeumnisText As String
                        saeumnisText = "S" & ChrW(228) & "umnis-Geb" & ChrW(252) & "hr: " & _
                                       Format(kategorien(k).saeumnisGebuehr, "#,##0.00") & _
                                       " " & ChrW(8364)
                        If bemerkung = "" Then
                            bemerkung = saeumnisText
                        Else
                            bemerkung = bemerkung & " | " & saeumnisText
                        End If
                    End If
                End If
                
                ' Kein fester Soll -> Hinweis
                If Not kategorien(k).HatFestenSoll Then
                    Dim variabelHinweis As String
                    ' v4.3: Unterschiedliche Bemerkung je nach Quelle
                    If manuellSoll > 0 Then
                        variabelHinweis = "Soll aus Vormonat " & ChrW(252) & "bernommen (" & _
                                          Format(manuellSoll, "#,##0.00") & " " & ChrW(8364) & ")"
                    Else
                        variabelHinweis = "Soll-Betrag variabel (bitte manuell eintragen)"
                    End If
                    If bemerkung = "" Then
                        bemerkung = variabelHinweis
                    Else
                        bemerkung = bemerkung & " | " & variabelHinweis
                    End If
                End If
                
                ' v4.4: Partner-Info anhaengen wenn Mitgliedsbeitrag mitbezahlt
                If partnerInfo <> "" Then
                    If bemerkung = "" Then
                        bemerkung = partnerInfo
                    Else
                        bemerkung = bemerkung & " | " & partnerInfo
                    End If
                End If
                
                ' v4.6: Hinweis wenn Januar ohne Vorjahr-Daten auf GELB herabgestuft wurde
                If monat = 1 And ist = 0 And Not mod_Uebersicht_Daten.HatVorjahrDaten() Then
                    Dim vjHinweis As String
                    vjHinweis = "Keine Vorjahr-Daten: Zahlung evtl. im Dezember erfolgt"
                    If bemerkung = "" Then
                        bemerkung = vjHinweis
                    Else
                        bemerkung = bemerkung & " | " & vjHinweis
                    End If
                End If
                
                wsUeb.Cells(rowIdx, UEB_COL_BEMERKUNG).value = bemerkung
                
                ' v4.5: Parzelle-basierte Kategorie als geschrieben markieren
                If StrComp(kategorie, "Mitgliedsbeitrag", vbTextCompare) <> 0 Then
                    geschriebeneParzKat(CStr(parzelleWert) & "|" & monat & "|" & kategorie) = entityKey
                End If
                
                ' v5.0: Kumulierten Ist-Betrag tracken
                Dim sumKey As String
                sumKey = CStr(parzelleWert) & "|" & kategorie
                If summeIstDict.Exists(sumKey) Then
                    summeIstDict(sumKey) = summeIstDict(sumKey) + ist
                Else
                    summeIstDict(sumKey) = ist
                End If
                ' Zeilennummer merken
                If Not summeIstZeilen.Exists(sumKey) Then
                    Set summeIstZeilen(sumKey) = New Collection
                End If
                summeIstZeilen(sumKey).Add rowIdx
                
                rowIdx = rowIdx + 1
                
NextKat:
            Next k
        Next monat
    Next mitglied
    
    ' v5.0: Kumulierte Ist-Summen in Spalte I schreiben
    Dim sk As Variant
    For Each sk In summeIstZeilen.keys
        Dim zeilen As Collection
        Set zeilen = summeIstZeilen(sk)
        Dim sumWert As Double
        sumWert = summeIstDict(sk)
        Dim zi As Long
        For zi = 1 To zeilen.count
            wsUeb.Cells(zeilen(zi), UEB_COL_SUMME_IST).value = sumWert
        Next zi
    Next sk
    
    ' Formatierung anwenden
    Call FormatiereUebersicht(wsUeb, UEBERSICHT_START_ROW, rowIdx - 1)
    
    ' v4.4: AutoFilter auf Header-Zeile aktivieren (Dropdown-Pfeile immer sichtbar)
    If wsUeb.AutoFilterMode Then wsUeb.AutoFilterMode = False
    wsUeb.Range(wsUeb.Cells(UEBERSICHT_HEADER_ROW, UEB_COL_PARZELLE), _
                wsUeb.Cells(rowIdx - 1, UEB_COL_SUMME_IST)).AutoFilter
    
    ' Monats-Register (Shape-Tabs) erstellen/aktualisieren
    Call mod_Uebersicht_Filter.ErstelleMonatsRegister
    
    ' Einstellungen-Cache freigeben
    Call mod_Zahlungspruefung.EntladeEinstellungenCacheZP
    
    ' Blatt sch?tzen (Soll-Zellen ohne festen Betrag bleiben editierbar)
    ' AllowFiltering: Nutzer kann AutoFilter ohne Blattschutz-Aufhebung verwenden
    On Error Resume Next
    wsUeb.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo ErrorHandler
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Dim endTime As Double
    endTime = Timer
    
    ' Erfolgsmeldung nur bei manuellem Aufruf (nicht im stummModus)
    Debug.Print "[" & ChrW(220) & "bersicht] Erfolgreich: " & _
                (rowIdx - UEBERSICHT_START_ROW) & " Zeilen in " & _
                Format(endTime - startTime, "0.00") & "s"
    
    If Not stummModus Then
        MsgBox ChrW(220) & "bersicht erfolgreich generiert!" & vbLf & vbLf & _
               "Zeilen: " & (rowIdx - UEBERSICHT_START_ROW) & vbLf & _
               "Kategorien: " & anzahlKat & " (dynamisch aus Einstellungen)" & vbLf & _
               "Dauer: " & Format(endTime - startTime, "0.00") & " Sekunden", _
               vbInformation, "Fertig"
    End If
    
    m_IsGenerating = False
    Exit Sub
    
ErrorHandler:
    m_IsGenerating = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' v4.5b: Blatt im Fehlerfall trotzdem schuetzen
    On Error Resume Next
    If Not wsUeb Is Nothing Then
        wsUeb.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    On Error GoTo 0
    
    ' IMMER Debug.Print bei Fehler (auch im stummModus)
    Debug.Print "[" & ChrW(220) & "bersicht] FEHLER: " & Err.Number & " - " & Err.Description
    
    If Not stummModus Then
        MsgBox "Fehler beim Generieren der " & ChrW(220) & "bersicht:" & vbLf & vbLf & _
               Err.Description, vbCritical, "Fehler"
    End If
    
End Sub


' ===============================================================
' v4.0: Pr?ft ob eine Kategorie in einem bestimmten Monat f?llig ist
' Kombiniert SollMonate (Einstellungen Spalte E) mit Faelligkeit
' (Daten Spalte O).
' - monatlich: SollMonate oder alle Monate
' - jaehrlich: nur SollMonate (1 Monat)
' - halbjaehrlich: nur SollMonate (2 Monate)
' - quartalsweise/vierteljaehrlich: nur SollMonate (3-4 Monate)
' - benutzerdefiniert: nur SollMonate
' Wenn SollMonate leer UND Faelligkeit nicht monatlich ->
'   Kategorie ist NICHT in allen Monaten faellig!
'   Dann Fallback: Kategorie nie anzeigen (muss in Einstellungen
'   Spalte E konfiguriert werden)
' ===============================================================
Private Function IstKategorieImMonatFaellig(ByRef kat As UebKategorie, _
                                             ByVal monat As Long) As Boolean
    
    ' Wenn SollMonate definiert -> nur diese pruefen
    If kat.SollMonate <> "" Then
        IstKategorieImMonatFaellig = mod_KategorieEngine_Zeitraum.IstMonatInListe(monat, kat.SollMonate)
        Exit Function
    End If
    
    ' SollMonate leer -> Faelligkeit entscheidet
    Dim fl As String
    fl = kat.faelligkeit
    
    ' Monatlich oder leer -> in ALLEN Monaten faellig
    If fl = "" Or fl = "monatlich" Then
        IstKategorieImMonatFaellig = True
        Exit Function
    End If
    
    ' Nicht-monatlich OHNE SollMonate -> NICHT anzeigen
    ' (Kategorie muss in Einstellungen Spalte E konfiguriert werden)
    Debug.Print "[" & ChrW(220) & "bersicht] WARNUNG: Kategorie '" & kat.Name & _
                "' ist '" & fl & "' aber Spalte E (Soll-Monate) ist leer! " & _
                "Bitte in Einstellungen die Soll-Monate eintragen."
    IstKategorieImMonatFaellig = False
    
End Function


' ===============================================================
' Header im ?bersichtsblatt setzen
' ===============================================================
Private Sub SetzeUebersichtHeader(ByVal wsUeb As Worksheet)
    
    With wsUeb
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_PARZELLE).value = "Parzelle"
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_MITGLIED).value = "Mitglied"
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_MONAT).value = "Monat"
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_KATEGORIE).value = "Kategorie"
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_SOLL).value = "Soll"
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_IST).value = "Ist"
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_STATUS).value = "Status"
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_BEMERKUNG).value = "Bemerkung"
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_SUMME_IST).value = ChrW(931) & " Ist"
        
        ' Header formatieren
        Dim rngHeader As Range
        Set rngHeader = .Range(.Cells(UEBERSICHT_HEADER_ROW, UEB_COL_PARZELLE), _
                                .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_SUMME_IST))
        
        With rngHeader
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.color = RGB(217, 217, 217)  ' Hellgrau
            .Borders.LineStyle = xlContinuous
        End With
    End With
    
End Sub


' ===============================================================
' Formatierung des ?bersichtsblatts
' ===============================================================
Private Sub FormatiereUebersicht(ByVal wsUeb As Worksheet, _
                                   ByVal startRow As Long, _
                                   ByVal endRow As Long)
    
    Dim r As Long
    Dim rngTable As Range
    
    If endRow < startRow Then Exit Sub
    
    Set rngTable = wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_PARZELLE), _
                                wsUeb.Cells(endRow, UEB_COL_SUMME_IST))
    
    ' Zebramuster (identisch mit Bankkonto/EntityKey-Tabelle)
    ' Ungerade Zeilen (1., 3., 5. Datenzeile) = weiss
    ' Gerade Zeilen (2., 4., 6. Datenzeile) = ZEBRA_COLOR
    ' ACHTUNG: Nicht ueberschreiben bei Soll-Spalte (hell-gelb) und Status-Spalte (Ampel)
    For r = startRow To endRow
        Dim c As Long
        If (r - startRow) Mod 2 = 1 Then
            ' Gerade Datenzeile -> Zebra-Farbe
            For c = UEB_COL_PARZELLE To UEB_COL_SUMME_IST
                If c = UEB_COL_STATUS Then
                    ' Status-Spalte (G) behaelt IMMER ihre Ampelfarbe
                ElseIf c = UEB_COL_SOLL Then
                    If wsUeb.Cells(r, c).Interior.color <> FARBE_HELLGELB_MANUELL And _
                       wsUeb.Cells(r, c).Interior.color <> AMPEL_GRUEN Then
                        wsUeb.Cells(r, c).Interior.color = ZEBRA_COLOR
                    End If
                ElseIf c = UEB_COL_SUMME_IST Then
                    ' Summen-Spalte: eigene Zebra-Farbe (hellblau)
                    wsUeb.Cells(r, c).Interior.color = FARBE_SUMME_ZEBRA
                Else
                    wsUeb.Cells(r, c).Interior.color = ZEBRA_COLOR
                End If
            Next c
        Else
            ' Ungerade Datenzeile -> weiss (aber Soll/Status/Summe auslassen)
            For c = UEB_COL_PARZELLE To UEB_COL_SUMME_IST
                If c = UEB_COL_STATUS Then
                    ' Status-Spalte behaelt Ampelfarbe
                ElseIf c = UEB_COL_SOLL Then
                    If wsUeb.Cells(r, c).Interior.color <> FARBE_HELLGELB_MANUELL And _
                       wsUeb.Cells(r, c).Interior.color <> AMPEL_GRUEN Then
                        wsUeb.Cells(r, c).Interior.ColorIndex = xlNone
                    End If
                ElseIf c = UEB_COL_SUMME_IST Then
                    ' Summen-Spalte: dezentes Hellblau
                    wsUeb.Cells(r, c).Interior.color = FARBE_SUMME
                Else
                    wsUeb.Cells(r, c).Interior.ColorIndex = xlNone
                End If
            Next c
        End If
    Next r
    
    ' Rahmen
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' Trennlinie links von Summen-Spalte
    wsUeb.Range(wsUeb.Cells(UEBERSICHT_HEADER_ROW, UEB_COL_SUMME_IST), _
                wsUeb.Cells(endRow, UEB_COL_SUMME_IST)).Borders(xlEdgeLeft).Weight = xlMedium
    
    ' Spaltenbreiten: AutoFit basierend auf Inhalt
    Dim colAutoFit As Long
    For colAutoFit = UEB_COL_PARZELLE To UEB_COL_SUMME_IST
        wsUeb.Columns(colAutoFit).AutoFit
        ' Mindestbreite sicherstellen (Header nicht abschneiden)
        If wsUeb.Columns(colAutoFit).ColumnWidth < 10 Then
            wsUeb.Columns(colAutoFit).ColumnWidth = 10
        End If
        ' Etwas Puffer fuer bessere Lesbarkeit
        wsUeb.Columns(colAutoFit).ColumnWidth = wsUeb.Columns(colAutoFit).ColumnWidth + 2
    Next colAutoFit
    
    ' Deutsches Zahlenformat mit Euro-Zeichen (Spalte E + F)
    wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_SOLL), _
                wsUeb.Cells(endRow, UEB_COL_SOLL)).NumberFormat = "#,##0.00 " & ChrW(8364)
    wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_IST), _
                wsUeb.Cells(endRow, UEB_COL_IST)).NumberFormat = "#,##0.00 " & ChrW(8364)
    wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_SUMME_IST), _
                wsUeb.Cells(endRow, UEB_COL_SUMME_IST)).NumberFormat = "#,##0.00 " & ChrW(8364)
    
    ' Ausrichtung
    wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_PARZELLE), _
                wsUeb.Cells(endRow, UEB_COL_PARZELLE)).HorizontalAlignment = xlCenter
    ' Spalte C (Monat) linksbuendig
    wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_MONAT), _
                wsUeb.Cells(endRow, UEB_COL_MONAT)).HorizontalAlignment = xlLeft
    wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_STATUS), _
                wsUeb.Cells(endRow, UEB_COL_STATUS)).HorizontalAlignment = xlCenter
    
    ' Vertikale Zentrierung
    rngTable.VerticalAlignment = xlCenter
    
End Sub


' ===============================================================
' v4.3: Sammelt manuell eingetragene Soll-Werte aus der
' bestehenden Uebersicht BEVOR diese geloescht wird.
' Gibt ein Dictionary zurueck: Key = "Parzelle|Kategorie"
'                               Value = Soll-Betrag (Double)
' Nur Zeilen mit hell-gelber oder gruener Soll-Zelle werden
' beruecksichtigt (= variable Soll-Betraege).
' ===============================================================
Private Function SammleManuelleSollWerte(ByVal wsUeb As Worksheet) As Object
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    Dim lastRow As Long
    lastRow = wsUeb.Cells(wsUeb.Rows.count, UEB_COL_PARZELLE).End(xlUp).Row
    
    If lastRow < UEBERSICHT_START_ROW Then
        Set SammleManuelleSollWerte = dict
        Exit Function
    End If
    
    Dim r As Long
    For r = UEBERSICHT_START_ROW To lastRow
        Dim parzelle As String
        parzelle = CStr(wsUeb.Cells(r, UEB_COL_PARZELLE).value)
        
        ' v4.4: Keine Gruppenblock-Logik mehr noetig, jede Zeile hat Parzelle
        If parzelle = "" Then GoTo NextSollRow
        
        ' Nur variable Soll-Zellen (hell-gelb oder gruen) beruecksichtigen
        Dim sollFarbe As Long
        sollFarbe = wsUeb.Cells(r, UEB_COL_SOLL).Interior.color
        
        If sollFarbe = FARBE_HELLGELB_MANUELL Or sollFarbe = AMPEL_GRUEN Then
            Dim sollWert As Double
            sollWert = val(CStr(wsUeb.Cells(r, UEB_COL_SOLL).value))
            
            If sollWert > 0 Then
                Dim kat As String
                kat = CStr(wsUeb.Cells(r, UEB_COL_KATEGORIE).value)
                Dim dictKey As String
                dictKey = parzelle & "|" & kat
                
                ' Letzten Wert speichern (neuester Monat gewinnt)
                dict(dictKey) = sollWert
            End If
        End If
NextSollRow:
    Next r
    
    Set SammleManuelleSollWerte = dict
    
End Function


' ===============================================================
' v4.3: Sucht im gesicherten Dictionary nach einem manuell
' eingetragenen Soll-Betrag fuer die gleiche Parzelle+Kategorie.
' Gibt den Betrag zurueck (0 wenn nicht gefunden).
' ===============================================================
Private Function HoleManuellSollAusVormonat(ByVal gespeicherteSoll As Object, _
                                             ByVal parzelle As String, _
                                             ByVal kategorie As String) As Double
    
    HoleManuellSollAusVormonat = 0
    
    If gespeicherteSoll Is Nothing Then Exit Function
    
    Dim dictKey As String
    dictKey = parzelle & "|" & kategorie
    
    If gespeicherteSoll.Exists(dictKey) Then
        HoleManuellSollAusVormonat = gespeicherteSoll(dictKey)
    End If
    
End Function


' ===============================================================
' v4.4: Pruefen ob ein Partner auf der gleichen Parzelle den
' Mitgliedsbeitrag fuer beide bezahlt hat (z.B. 15 EUR statt 7.50).
' Gibt Partner-Info-String zurueck wenn mitbezahlt, sonst "".
' ===============================================================
Private Function PruefePartnerMitgliedsbeitrag( _
                    ByVal mitglieder As Collection, _
                    ByVal parzelle As Long, _
                    ByVal meineEntityKey As String, _
                    ByVal monat As Long, _
                    ByVal jahr As Long, _
                    ByVal sollProPerson As Double) As String
    
    PruefePartnerMitgliedsbeitrag = ""
    
    Dim partner As Object
    For Each partner In mitglieder
        ' Gleiche Parzelle, anderer EntityKey?
        If CLng(partner("Parzelle")) = parzelle Then
            If StrComp(CStr(partner("EntityKey")), meineEntityKey, vbTextCompare) <> 0 Then
                ' Partner gefunden - Zahlung pruefen
                Dim partnerErgebnis As String
                partnerErgebnis = mod_Zahlungspruefung.PruefeZahlungen( _
                    CStr(partner("EntityKey")), "Mitgliedsbeitrag", monat, jahr)
                
                ' IST aus Ergebnis parsen
                Dim partnerTeile() As String
                partnerTeile = Split(partnerErgebnis, "|")
                
                Dim partnerIst As Double
                partnerIst = 0
                If UBound(partnerTeile) >= 2 Then
                    Dim pIT() As String
                    pIT = Split(partnerTeile(2), ":")
                    If UBound(pIT) >= 1 Then partnerIst = val(pIT(1))
                End If
                
                ' Partner hat >= 2x Soll bezahlt -> deckt beide ab
                If partnerIst >= sollProPerson * 2 - 0.01 Then
                    PruefePartnerMitgliedsbeitrag = _
                        "Mitbezahlt durch " & CStr(partner("Name"))
                    Exit Function
                End If
            End If
        End If
    Next partner
    
End Function



















































