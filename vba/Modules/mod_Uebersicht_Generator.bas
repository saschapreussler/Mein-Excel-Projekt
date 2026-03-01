Attribute VB_Name = "mod_Uebersicht_Generator"
Option Explicit

' ***************************************************************
' MODUL: mod_Uebersicht_Generator
' VERSION: 4.1 - 01.06.2026
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


' ===============================================================
' Type f?r eine dynamische Kategorie aus Einstellungen
' ===============================================================
Private Type UebKategorie
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
' HAUPTFUNKTION: Generiert komplettes ?bersichtsblatt
' v2.0: Kategorien DYNAMISCH aus Einstellungen-Blatt
' v3.0: stummModus f?r automatische Aufrufe (ohne MsgBox)
' ===============================================================
Public Sub GeneriereUebersicht(Optional ByVal jahr As Long = 0, _
                                Optional ByVal stummModus As Boolean = False)
    
    On Error GoTo ErrorHandler
    
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
    ' v4.0: Wenn kein Jahr angegeben -> aus Bankkonto-Daten ermitteln
    If jahr = 0 Then
        jahr = ErmittleJahrAusBankkonto()
        If jahr = 0 Then jahr = Year(Date)
    End If
    Debug.Print "[" & ChrW(220) & "bersicht] Verwende Jahr: " & jahr
    
    ' =============================================
    ' v2.0: Kategorien DYNAMISCH aus Einstellungen laden
    ' =============================================
    Dim kategorien() As UebKategorie
    Dim anzahlKat As Long
    Call LadeKategorienAusEinstellungen(kategorien, anzahlKat)
    
    If anzahlKat = 0 Then
        If Not stummModus Then
            MsgBox "Keine Kategorien im Einstellungen-Blatt (Spalte B) gefunden!" & vbLf & _
                   "Bitte mindestens eine Kategorie mit Zahlungstermin anlegen.", _
                   vbCritical, "Fehler"
        End If
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
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Blatt entsperren
    On Error Resume Next
    wsUeb.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    ' Alten Inhalt l?schen (ab Zeile 4)
    wsUeb.Range(wsUeb.Cells(UEBERSICHT_START_ROW, 1), _
                wsUeb.Cells(wsUeb.Rows.count, UEB_COL_BEMERKUNG)).ClearContents
    wsUeb.Range(wsUeb.Cells(UEBERSICHT_START_ROW, 1), _
                wsUeb.Cells(wsUeb.Rows.count, UEB_COL_BEMERKUNG)).Interior.ColorIndex = xlNone
    
    ' Header setzen
    Call SetzeUebersichtHeader(wsUeb)
    
    ' Einstellungen-Cache laden (Performance)
    Call mod_Zahlungspruefung.LadeEinstellungenCacheZP
    
    ' Dezember-Cache initialisieren (f?r Vorauszahlungen)
    Call mod_Zahlungspruefung.InitialisiereNachDezemberCacheZP(jahr)
    
    ' v4.0: Vorjahr-Speicher bef?llen (Okt-Dez Vorjahr)
    Call BefuelleVorjahrSpeicher(jahr - 1)
    
    ' v4.0: Vorjahr-Speicher automatisch loeschen (ab August)
    Call PruefeVorjahrSpeicherAblauf
    
    ' Aktive Mitglieder aus Daten-Blatt EntityKey-Tabelle laden
    Set mitglieder = HoleAktiveMitglieder(wsDaten)
    
    ' Debug-Diagnose: Mitglieder und Kategorien protokollieren
    Debug.Print "[" & ChrW(220) & "bersicht] Kategorien: " & anzahlKat & _
                " | Mitglieder: " & mitglieder.count
    
    If mitglieder.count = 0 Then
        Debug.Print "[" & ChrW(220) & "bersicht] WARNUNG: Keine aktiven Mitglieder gefunden!"
        Debug.Print "[" & ChrW(220) & "bersicht] Pr" & ChrW(252) & "fe Daten-Blatt: " & _
                    "EntityKey (R), Parzelle (V), Role (W)"
        
        ' Auch im stummModus eine Warnung ausgeben, da Mitglieder-Daten fehlen
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
    importierteMonate = ErmittleImportierteMonate(jahr)
    
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
                    vjBetrag = HoleVorjahrZahlung(entityKey, kategorie, monat)
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
                
                ' =============================================
                ' v4.0: FILTER - Nur relevante Eintraege anzeigen
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
                ' v4.0: Einheitliches Datumsformat "Januar 2026"
                wsUeb.Cells(rowIdx, UEB_COL_MONAT).value = MonthName(monat) & " " & jahr
                wsUeb.Cells(rowIdx, UEB_COL_KATEGORIE).value = kategorie
                
                ' =============================================
                ' v2.0: Soll-Betrag Logik
                ' =============================================
                If kategorien(k).HatFestenSoll Then
                    ' Fester Soll-Betrag aus Einstellungen
                    wsUeb.Cells(rowIdx, UEB_COL_SOLL).value = soll
                Else
                    ' KEIN fester Soll-Betrag -> Zelle leer + hell-gelb
                    ' Nutzer kann hier pro Parzelle den individuellen Betrag eintragen
                    wsUeb.Cells(rowIdx, UEB_COL_SOLL).value = ""
                    wsUeb.Cells(rowIdx, UEB_COL_SOLL).Interior.color = FARBE_HELLGELB_MANUELL
                    wsUeb.Cells(rowIdx, UEB_COL_SOLL).Locked = False
                    
                    ' Status bei variablem Betrag: nur Termin-Pr?fung
                    ' Wenn Ist > 0 -> Zahlung eingegangen -> GR?N
                    ' Wenn Ist = 0 -> Keine Zahlung -> ROT oder GELB
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
                    variabelHinweis = "Soll-Betrag variabel (bitte manuell eintragen)"
                    If bemerkung = "" Then
                        bemerkung = variabelHinweis
                    Else
                        bemerkung = bemerkung & " | " & variabelHinweis
                    End If
                End If
                
                wsUeb.Cells(rowIdx, UEB_COL_BEMERKUNG).value = bemerkung
                
                rowIdx = rowIdx + 1
                
NextKat:
            Next k
        Next monat
    Next mitglied
    
    ' Formatierung anwenden
    Call FormatiereUebersicht(wsUeb, UEBERSICHT_START_ROW, rowIdx - 1)
    
    ' Einstellungen-Cache freigeben
    Call mod_Zahlungspruefung.EntladeEinstellungenCacheZP
    
    ' Blatt sch?tzen (Soll-Zellen ohne festen Betrag bleiben editierbar)
    On Error Resume Next
    wsUeb.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
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
    
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' IMMER Debug.Print bei Fehler (auch im stummModus)
    Debug.Print "[" & ChrW(220) & "bersicht] FEHLER: " & Err.Number & " - " & Err.Description
    
    If Not stummModus Then
        MsgBox "Fehler beim Generieren der " & ChrW(220) & "bersicht:" & vbLf & vbLf & _
               Err.Description, vbCritical, "Fehler"
    End If
    
End Sub


' ===============================================================
' v2.0: L?dt Kategorien DYNAMISCH aus Einstellungen-Blatt
' Liest Spalte B (Kategorie), C (Soll-Betrag), E (Soll-Monate),
' I (S?umnis-Geb?hr)
' Gibt eindeutige Kategorien zur?ck (keine Duplikate)
' ===============================================================
Private Sub LadeKategorienAusEinstellungen(ByRef kategorien() As UebKategorie, _
                                            ByRef anzahl As Long)
    
    Dim wsEinst As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim katName As String
    Dim dict As Object
    
    anzahl = 0
    
    On Error Resume Next
    Set wsEinst = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    
    If wsEinst Is Nothing Then Exit Sub
    
    lastRow = wsEinst.Cells(wsEinst.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    If lastRow < ES_START_ROW Then Exit Sub
    
    ' Dictionary f?r Eindeutigkeit
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Zuerst z?hlen f?r ReDim
    For r = ES_START_ROW To lastRow
        katName = Trim(CStr(wsEinst.Cells(r, ES_COL_KATEGORIE).value))
        If katName <> "" Then
            If Not dict.Exists(katName) Then
                dict.Add katName, r  ' Merke Zeilennummer f?r sp?teres Lesen
            End If
        End If
    Next r
    
    anzahl = dict.count
    If anzahl = 0 Then Exit Sub
    
    ReDim kategorien(0 To anzahl - 1)
    
    Dim idx As Long
    idx = 0
    Dim key As Variant
    
    For Each key In dict.keys
        r = dict(key)  ' Zeilennummer aus Dictionary
        
        With kategorien(idx)
            .Name = CStr(key)
            
            ' Soll-Betrag aus Spalte C
            Dim sollWert As Variant
            sollWert = wsEinst.Cells(r, ES_COL_SOLL_BETRAG).value
            If IsNumeric(sollWert) Then
                .SollBetrag = CDbl(sollWert)
            Else
                .SollBetrag = 0
            End If
            .HatFestenSoll = (.SollBetrag > 0)
            
            ' S?umnis-Geb?hr aus Spalte I
            Dim saeumnisWert As Variant
            saeumnisWert = wsEinst.Cells(r, ES_COL_SAEUMNIS).value
            If IsNumeric(saeumnisWert) Then
                .saeumnisGebuehr = CDbl(saeumnisWert)
            Else
                .saeumnisGebuehr = 0
            End If
            
            ' Soll-Monate aus Spalte E (z.B. "03, 06, 09" oder leer = alle)
            .SollMonate = Trim(CStr(wsEinst.Cells(r, ES_COL_SOLL_MONATE).value))
            
            ' Faelligkeit aus Daten-Blatt Spalte O (Kategorie-Tabelle)
            .faelligkeit = ""
        End With
        
        idx = idx + 1
    Next key
    
    ' Faelligkeit aus Daten-Blatt nachladen (Spalte O)
    Dim wsDatenKat As Worksheet
    On Error Resume Next
    Set wsDatenKat = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If Not wsDatenKat Is Nothing Then
        Dim lastRowDaten As Long
        lastRowDaten = wsDatenKat.Cells(wsDatenKat.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
        
        Dim ki As Long
        For ki = 0 To anzahl - 1
            Dim rD As Long
            For rD = DATA_START_ROW To lastRowDaten
                If StrComp(Trim(CStr(wsDatenKat.Cells(rD, DATA_CAT_COL_KATEGORIE).value)), _
                           kategorien(ki).Name, vbTextCompare) = 0 Then
                    kategorien(ki).faelligkeit = LCase(Trim(CStr( _
                        wsDatenKat.Cells(rD, DATA_CAT_COL_FAELLIGKEIT).value)))
                    Exit For
                End If
            Next rD
        Next ki
    End If
    
    Set dict = Nothing
    
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
        
        ' Header formatieren
        Dim rngHeader As Range
        Set rngHeader = .Range(.Cells(UEBERSICHT_HEADER_ROW, UEB_COL_PARZELLE), _
                                .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_BEMERKUNG))
        
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
' Holt alle aktiven Mitglieder aus Daten-Blatt (EntityKey-Tabelle)
' Spalten: R=EntityKey, S=IBAN, T=Kontoname, U=Zuordnung, V=Parzelle, W=Role
' Bei SHARE-Keys k?nnen mehrere Parzellen in V stehen (z.B. "2, 5")
' v4.0: Mehrere Mitglieder pro Parzelle erlaubt (z.B. MIT + OHNE PACHT)
'       Dedup ueber EntityKey+Parzelle (nicht nur Parzelle)
'       Name aus Spalte T (Kontoname), Fallback auf Spalte U (Zuordnung)
' ===============================================================
Private Function HoleAktiveMitglieder(ByVal wsDaten As Worksheet) As Collection
    
    Dim col As Collection
    Set col = New Collection
    
    Dim lastRow As Long
    lastRow = wsDaten.Cells(wsDaten.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    
    If lastRow < EK_START_ROW Then
        Set HoleAktiveMitglieder = col
        Exit Function
    End If
    
    ' Dictionary f?r bereits verarbeitete EntityKey+Parzelle-Kombinationen
    Dim verarbeiteteKombis As Object
    Set verarbeiteteKombis = CreateObject("Scripting.Dictionary")
    
    Dim r As Long
    Dim entityKey As String
    Dim zuordnung As String
    Dim parzelleWert As String
    Dim roleWert As String
    Dim dict As Object
    
    For r = EK_START_ROW To lastRow
        entityKey = Trim(CStr(wsDaten.Cells(r, EK_COL_ENTITYKEY).value))
        If entityKey = "" Then GoTo NextDatenRow
        
        ' Role pr?fen: nur aktive Mitglieder
        ' "MITGLIED MIT PACHT" und "MITGLIED OHNE PACHT" -> ja
        ' "EHEMALIGES MITGLIED" -> nein (ausschlie?en)
        roleWert = UCase(Trim(CStr(wsDaten.Cells(r, EK_COL_ROLE).value)))
        If InStr(roleWert, "MITGLIED") = 0 Then GoTo NextDatenRow
        If InStr(roleWert, "EHEMALIGES") > 0 Then GoTo NextDatenRow
        
        ' Parzelle(n) lesen (kann "2" oder "2, 5" sein bei SHARE-Keys)
        parzelleWert = Trim(CStr(wsDaten.Cells(r, EK_COL_PARZELLE).value))
        If parzelleWert = "" Then GoTo NextDatenRow
        
        ' Zuordnung (Kontoname) aus Spalte T - der echte Kontoinhaber
        zuordnung = Trim(CStr(wsDaten.Cells(r, EK_COL_KONTONAME).value))
        ' Falls Kontoname leer -> Fallback auf Zuordnung (Spalte U)
        If zuordnung = "" Then
            zuordnung = Trim(CStr(wsDaten.Cells(r, EK_COL_ZUORDNUNG).value))
        End If
        
        ' Parzelle(n) aufteilen (bei SHARE-Keys: "2, 5" -> 2 Eintr?ge)
        Dim parzellen() As String
        parzellen = Split(parzelleWert, ",")
        
        Dim p As Long
        For p = LBound(parzellen) To UBound(parzellen)
            Dim einzelParzelle As String
            einzelParzelle = Trim(parzellen(p))
            
            If IsNumeric(einzelParzelle) Then
                Dim parzelleNr As Long
                parzelleNr = CLng(einzelParzelle)
                
                ' Nur Parzellen 1-14
                If parzelleNr >= 1 And parzelleNr <= 14 Then
                    ' Duplikat-Pr?fung: EntityKey+Parzelle nur einmal
                    Dim kombiKey As String
                    kombiKey = entityKey & "_" & parzelleNr
                    
                    If Not verarbeiteteKombis.Exists(kombiKey) Then
                        verarbeiteteKombis.Add kombiKey, True
                        
                        Set dict = CreateObject("Scripting.Dictionary")
                        dict.Add "Parzelle", parzelleNr
                        dict.Add "EntityKey", entityKey
                        dict.Add "Name", zuordnung
                        dict.Add "Role", roleWert
                        
                        col.Add dict
                    End If
                End If
            End If
        Next p
        
NextDatenRow:
    Next r
    
    Set verarbeiteteKombis = Nothing
    Set HoleAktiveMitglieder = col
    
End Function


' ===============================================================
' v4.0: Ermittelt das haeufigste Jahr aus Bankkonto-Daten
' Scannt Spalte A (Datum) und zaehlt welches Jahr am meisten
' vorkommt. Gibt 0 zurueck wenn keine Daten vorhanden.
' ===============================================================
Private Function ErmittleJahrAusBankkonto() As Long
    
    Dim wsBK As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim zellWert As Variant
    Dim buchDatum As Date
    Dim jahrZaehler As Object
    Dim jahrKey As String
    
    ErmittleJahrAusBankkonto = 0
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo 0
    
    If wsBK Is Nothing Then Exit Function
    
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Function
    
    Set jahrZaehler = CreateObject("Scripting.Dictionary")
    
    For r = BK_START_ROW To lastRow
        zellWert = wsBK.Cells(r, BK_COL_DATUM).value
        
        If IsDate(zellWert) Then
            buchDatum = CDate(zellWert)
            jahrKey = CStr(Year(buchDatum))
            
            If jahrZaehler.Exists(jahrKey) Then
                jahrZaehler(jahrKey) = jahrZaehler(jahrKey) + 1
            Else
                jahrZaehler.Add jahrKey, 1
            End If
        End If
    Next r
    
    ' Haeufigtes Jahr finden
    If jahrZaehler.count = 0 Then
        Set jahrZaehler = Nothing
        Exit Function
    End If
    
    Dim maxAnzahl As Long
    Dim maxJahr As String
    Dim key As Variant
    maxAnzahl = 0
    
    For Each key In jahrZaehler.keys
        If jahrZaehler(key) > maxAnzahl Then
            maxAnzahl = jahrZaehler(key)
            maxJahr = CStr(key)
        End If
    Next key
    
    ErmittleJahrAusBankkonto = CLng(maxJahr)
    
    Debug.Print "[" & ChrW(220) & "bersicht] Jahr aus Bankkonto erkannt: " & maxJahr & _
                " (" & maxAnzahl & " Buchungen)"
    
    Set jahrZaehler = Nothing
    
End Function


' ===============================================================
' v4.0: Ermittelt welche Monate im Bankkonto CSV-Daten haben
' Scannt Spalte A (Datum) ab BK_START_ROW und setzt True
' fuer jeden Monat der mindestens eine Buchung enthaelt
' Gibt Boolean-Array(1 To 12) zurueck
' ===============================================================
Private Function ErmittleImportierteMonate(ByVal jahr As Long) As Boolean()
    
    Dim result() As Boolean
    ReDim result(1 To 12)
    Dim wsBK As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim zellWert As Variant
    Dim buchDatum As Date
    Dim m As Long
    
    ' Array initialisieren (alles False - ReDim setzt bereits auf False)
    For m = 1 To 12
        result(m) = False
    Next m
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo 0
    
    If wsBK Is Nothing Then
        Debug.Print "[" & ChrW(220) & "bersicht] WARNUNG: Blatt 'Bankkonto' nicht gefunden!"
        ErmittleImportierteMonate = result
        Exit Function
    End If
    
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    Debug.Print "[" & ChrW(220) & "bersicht] Bankkonto lastRow=" & lastRow & _
                " (BK_START_ROW=" & BK_START_ROW & ")"
    
    If lastRow < BK_START_ROW Then
        Debug.Print "[" & ChrW(220) & "bersicht] Keine Buchungen im Bankkonto gefunden."
        ErmittleImportierteMonate = result
        Exit Function
    End If
    
    For r = BK_START_ROW To lastRow
        zellWert = wsBK.Cells(r, BK_COL_DATUM).value
        
        If IsDate(zellWert) Then
            buchDatum = CDate(zellWert)
            
            If Year(buchDatum) = jahr Then
                result(Month(buchDatum)) = True
            End If
        End If
    Next r
    
    ErmittleImportierteMonate = result
    
End Function


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
                                wsUeb.Cells(endRow, UEB_COL_BEMERKUNG))
    
    ' Zebramuster (identisch mit Bankkonto/EntityKey-Tabelle)
    ' Ungerade Zeilen (1., 3., 5. Datenzeile) = weiss
    ' Gerade Zeilen (2., 4., 6. Datenzeile) = ZEBRA_COLOR
    ' ACHTUNG: Nicht ueberschreiben bei Soll-Spalte (hell-gelb) und Status-Spalte (Ampel)
    For r = startRow To endRow
        Dim c As Long
        If (r - startRow) Mod 2 = 1 Then
            ' Gerade Datenzeile -> Zebra-Farbe
            For c = UEB_COL_PARZELLE To UEB_COL_BEMERKUNG
                If c = UEB_COL_STATUS Then
                    ' Status-Spalte (G) behaelt IMMER ihre Ampelfarbe
                ElseIf c = UEB_COL_SOLL Then
                    ' Soll-Spalte: Nur Zebra wenn NICHT hell-gelb (variabel)
                    If wsUeb.Cells(r, c).Interior.color <> FARBE_HELLGELB_MANUELL Then
                        wsUeb.Cells(r, c).Interior.color = ZEBRA_COLOR
                    End If
                Else
                    wsUeb.Cells(r, c).Interior.color = ZEBRA_COLOR
                End If
            Next c
        Else
            ' Ungerade Datenzeile -> weiss (aber Soll/Status auslassen)
            For c = UEB_COL_PARZELLE To UEB_COL_BEMERKUNG
                If c = UEB_COL_STATUS Then
                    ' Status-Spalte behaelt Ampelfarbe
                ElseIf c = UEB_COL_SOLL Then
                    If wsUeb.Cells(r, c).Interior.color <> FARBE_HELLGELB_MANUELL Then
                        wsUeb.Cells(r, c).Interior.ColorIndex = xlNone
                    End If
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
    
    ' Spaltenbreiten
    wsUeb.Columns(UEB_COL_PARZELLE).ColumnWidth = 10
    wsUeb.Columns(UEB_COL_MITGLIED).ColumnWidth = 25
    wsUeb.Columns(UEB_COL_MONAT).ColumnWidth = 18
    wsUeb.Columns(UEB_COL_KATEGORIE).ColumnWidth = 22
    wsUeb.Columns(UEB_COL_SOLL).ColumnWidth = 14
    wsUeb.Columns(UEB_COL_IST).ColumnWidth = 14
    wsUeb.Columns(UEB_COL_STATUS).ColumnWidth = 10
    wsUeb.Columns(UEB_COL_BEMERKUNG).ColumnWidth = 45
    
    ' Deutsches Zahlenformat mit Euro-Zeichen (Spalte E + F)
    wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_SOLL), _
                wsUeb.Cells(endRow, UEB_COL_SOLL)).NumberFormat = "#,##0.00 " & ChrW(8364)
    wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_IST), _
                wsUeb.Cells(endRow, UEB_COL_IST)).NumberFormat = "#,##0.00 " & ChrW(8364)
    
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
' v4.0: VORJAHR-SPEICHER - Okt-Dez des Vorjahres cachen
' Kopiert relevante Bankkonto-Buchungen (Okt-Dez Vorjahr) in den
' Hilfsspeicher auf Blatt Daten ab Spalte CA.
' Zweck: Dezember-Zahlungen die fuer Januar gelten erkennen
' ===============================================================
Public Sub BefuelleVorjahrSpeicher(ByVal vorjahr As Long)
    
    Dim wsBK As Worksheet
    Dim wsDaten As Worksheet
    Dim lastRowBK As Long
    Dim r As Long
    Dim vjRow As Long
    Dim zahlDatum As Date
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsBK Is Nothing Or wsDaten Is Nothing Then Exit Sub
    
    ' Zuerst alten Speicher loeschen
    Call LoescheVorjahrSpeicher
    
    ' Header setzen
    wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_DATUM).value = "VJ Datum"
    wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_BETRAG).value = "VJ Betrag"
    wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_IBAN).value = "VJ IBAN"
    wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_KATEGORIE).value = "VJ Kategorie"
    wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_MONAT_PERIODE).value = "VJ Monat/Periode"
    wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_ENTITYKEY).value = "VJ EntityKey"
    
    ' Header formatieren
    Dim rngVJHeader As Range
    Set rngVJHeader = wsDaten.Range(wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_DATUM), _
                                     wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_ENTITYKEY))
    rngVJHeader.Font.Bold = True
    rngVJHeader.Interior.color = RGB(217, 217, 217)
    
    lastRowBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    vjRow = VJ_START_ROW
    
    For r = BK_START_ROW To lastRowBK
        If Not IsDate(wsBK.Cells(r, BK_COL_DATUM).value) Then GoTo NextVJRow
        zahlDatum = CDate(wsBK.Cells(r, BK_COL_DATUM).value)
        
        ' Nur Okt-Dez des Vorjahres
        If Year(zahlDatum) <> vorjahr Then GoTo NextVJRow
        If Month(zahlDatum) < 10 Then GoTo NextVJRow
        
        ' Nur wenn Kategorie und IBAN vorhanden
        Dim vjKat As String
        vjKat = Trim(CStr(wsBK.Cells(r, BK_COL_KATEGORIE).value))
        If vjKat = "" Then GoTo NextVJRow
        
        Dim vjIBAN As String
        vjIBAN = Replace(Trim(CStr(wsBK.Cells(r, BK_COL_IBAN).value)), " ", "")
        If vjIBAN = "" Then GoTo NextVJRow
        
        ' In Speicher schreiben
        wsDaten.Cells(vjRow, VJ_COL_DATUM).value = zahlDatum
        wsDaten.Cells(vjRow, VJ_COL_DATUM).NumberFormat = "DD.MM.YYYY"
        wsDaten.Cells(vjRow, VJ_COL_BETRAG).value = wsBK.Cells(r, BK_COL_BETRAG).value
        wsDaten.Cells(vjRow, VJ_COL_IBAN).value = vjIBAN
        wsDaten.Cells(vjRow, VJ_COL_KATEGORIE).value = vjKat
        wsDaten.Cells(vjRow, VJ_COL_MONAT_PERIODE).value = _
            Trim(CStr(wsBK.Cells(r, BK_COL_MONAT_PERIODE).value))
        
        ' EntityKey via IBAN aufloesen (ueber EntityKey-Tabelle)
        Dim vjEK As String
        vjEK = ""
        Dim ek As Long
        Dim ekLastRow As Long
        ekLastRow = wsDaten.Cells(wsDaten.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
        For ek = EK_START_ROW To ekLastRow
            Dim ekIBAN As String
            ekIBAN = Replace(Trim(CStr(wsDaten.Cells(ek, EK_COL_IBAN).value)), " ", "")
            If StrComp(ekIBAN, vjIBAN, vbTextCompare) = 0 Then
                vjEK = Trim(CStr(wsDaten.Cells(ek, EK_COL_ENTITYKEY).value))
                Exit For
            End If
        Next ek
        wsDaten.Cells(vjRow, VJ_COL_ENTITYKEY).value = vjEK
        
        vjRow = vjRow + 1
        
NextVJRow:
    Next r
    
    Debug.Print "[" & ChrW(220) & "bersicht] Vorjahr-Speicher: " & _
                (vjRow - VJ_START_ROW) & " Buchungen aus Okt-Dez " & vorjahr & " gecached"
    
End Sub


' ===============================================================
' v4.0: Loescht den Vorjahr-Speicher auf Blatt Daten (ab CA)
' ===============================================================
Public Sub LoescheVorjahrSpeicher()
    
    Dim wsDaten As Worksheet
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = wsDaten.Cells(wsDaten.Rows.count, VJ_COL_DATUM).End(xlUp).Row
    
    If lastRow >= VJ_HEADER_ROW Then
        wsDaten.Range(wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_DATUM), _
                       wsDaten.Cells(lastRow, VJ_COL_ENTITYKEY)).ClearContents
        wsDaten.Range(wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_DATUM), _
                       wsDaten.Cells(lastRow, VJ_COL_ENTITYKEY)).Interior.ColorIndex = xlNone
    End If
    
    Debug.Print "[" & ChrW(220) & "bersicht] Vorjahr-Speicher gel" & ChrW(246) & "scht"
    
End Sub


' ===============================================================
' v4.0: Prueft automatisch ob Vorjahr-Speicher geloescht werden soll
' Ab August des Folgejahres wird der Speicher automatisch geleert
' ===============================================================
Public Sub PruefeVorjahrSpeicherAblauf()
    
    If Month(Date) >= 8 Then
        Dim wsDaten As Worksheet
        On Error Resume Next
        Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
        On Error GoTo 0
        
        If wsDaten Is Nothing Then Exit Sub
        
        ' Pruefen ob noch Daten im Speicher sind
        Dim ersteDatum As Variant
        ersteDatum = wsDaten.Cells(VJ_START_ROW, VJ_COL_DATUM).value
        
        If IsDate(ersteDatum) Then
            If Year(CDate(ersteDatum)) < Year(Date) - 1 Then
                ' Daten sind aelter als Vorjahr -> loeschen
                Call LoescheVorjahrSpeicher
            ElseIf Year(CDate(ersteDatum)) = Year(Date) - 1 Then
                ' Vorjahr-Daten und wir sind >= August -> loeschen
                Call LoescheVorjahrSpeicher
            End If
        End If
    End If
    
End Sub


' ===============================================================
' v4.0: Holt Vorjahr-Zahlungsbetrag aus dem Speicher
' Prueft ob fuer den EntityKey + Kategorie eine Dezember-Zahlung
' vorliegt, die fuer Januar des Folgejahres gelten koennte
' (basierend auf Monat/Periode in Spalte CE)
' ===============================================================
Public Function HoleVorjahrZahlung(ByVal entityKey As String, _
                                    ByVal kategorie As String, _
                                    ByVal monat As Long) As Double
    HoleVorjahrZahlung = 0
    
    ' Nur fuer fruehe Monate relevant (Jan-Maerz)
    If monat > 3 Then Exit Function
    
    Dim wsDaten As Worksheet
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Function
    
    Dim lastRow As Long
    lastRow = wsDaten.Cells(wsDaten.Rows.count, VJ_COL_DATUM).End(xlUp).Row
    If lastRow < VJ_START_ROW Then Exit Function
    
    Dim r As Long
    Dim vjMonatPeriode As String
    Dim erwarteterMonat As String
    erwarteterMonat = MonthName(monat)
    
    For r = VJ_START_ROW To lastRow
        ' EntityKey pruefen
        If StrComp(Trim(CStr(wsDaten.Cells(r, VJ_COL_ENTITYKEY).value)), _
                   entityKey, vbTextCompare) <> 0 Then GoTo NextVJPruefRow
        
        ' Kategorie pruefen
        If StrComp(Trim(CStr(wsDaten.Cells(r, VJ_COL_KATEGORIE).value)), _
                   kategorie, vbTextCompare) <> 0 Then GoTo NextVJPruefRow
        
        ' Monat/Periode pruefen
        vjMonatPeriode = Trim(CStr(wsDaten.Cells(r, VJ_COL_MONAT_PERIODE).value))
        
        If StrComp(vjMonatPeriode, erwarteterMonat, vbTextCompare) = 0 Then
            ' Direkt-Match: Monat/Periode = "Januar"
            HoleVorjahrZahlung = HoleVorjahrZahlung + Abs(wsDaten.Cells(r, VJ_COL_BETRAG).value)
        End If
        
NextVJPruefRow:
    Next r
    
End Function















