Attribute VB_Name = "mod_ZP_Periode"
Option Explicit

' ***************************************************************
' MODUL: mod_ZP_Periode
' VERSION: 1.0 - 15.03.2026
' ZWECK: Monat/Periode-Logik für Zahlungsprüfung
'        - SetzeMonatPeriode: Spalte I (Monat/Periode) befüllen
'        - HoleFaelligkeitFuerKategorie: Fälligkeit ermitteln
' QUELLE: Extrahiert aus mod_Zahlungspruefung v3.2
' ***************************************************************


' ===============================================================
' MONAT/PERIODE SETZEN (überarbeitet)
' FIX v1.5: Application.EnableEvents = False VOR dem Beschreiben
'           von Spalte I, damit Worksheet_Change NICHT getriggert wird.
' v2.0: Am Ende wird SetzeBankkontoDropDowns aufgerufen (für H + I)
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
                
                On Error GoTo Zeilenfehler
                ergebnis = mod_KategorieEngine_Zeitraum.ErmittleMonatPeriode( _
                    kategorie, CDate(datumWert), faelligkeit, ws, r)
                On Error GoTo SetzeMonatPeriodeError
                
                If Left(ergebnis, 5) = "GELB|" Then
                    Dim monatName As String
                    monatName = mid(ergebnis, 6)
                    
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
                    
                    ' Hell-gelber Hintergrund für die Bemerkung (gleiche Farbe wie Spalte I)
                    ws.Cells(r, BK_COL_BEMERKUNG).Interior.color = RGB(255, 235, 156)
                Else
                    ws.Cells(r, BK_COL_MONAT_PERIODE).value = ergebnis
                    ' Ampelfarbe Grün = Monat eindeutig bestimmt
                    ws.Cells(r, BK_COL_MONAT_PERIODE).Interior.color = RGB(198, 239, 206)
                    If IsDate(datumWert) And _
                       StrComp(ergebnis, MonthName(Month(CDate(datumWert))), vbTextCompare) <> 0 Then
                        Dim automatischeBemerkung As String
                        automatischeBemerkung = "Folgemonat automatisch zugeordnet: " & ergebnis
                        If InStr(1, CStr(ws.Cells(r, BK_COL_BEMERKUNG).value), automatischeBemerkung, vbTextCompare) = 0 Then
                            If Trim$(CStr(ws.Cells(r, BK_COL_BEMERKUNG).value)) = "" Then
                                ws.Cells(r, BK_COL_BEMERKUNG).value = automatischeBemerkung
                            Else
                                ws.Cells(r, BK_COL_BEMERKUNG).value = _
                                    CStr(ws.Cells(r, BK_COL_BEMERKUNG).value) & vbLf & automatischeBemerkung
                            End If
                        End If
                    End If
                End If
            Else
                ws.Cells(r, BK_COL_MONAT_PERIODE).value = MonthName(Month(datumWert))
            End If
        End If
WeiterMitNaechsterZeile:
    Next r
    
    ' Einstellungen-Cache wieder freigeben
    Call mod_KategorieEngine_Zeitraum.EntladeEinstellungenCache
    
    ' v1.5 FIX: Events wieder einschalten
    Application.EnableEvents = eventsWaren
    
    ' v2.0: ALLE DropDowns setzen (H + I) und Spalten entsperren
    ' (ausgelagert nach mod_ZP_DropDowns)
    Call mod_ZP_DropDowns.SetzeBankkontoDropDowns(ws)
    
    Exit Sub

Zeilenfehler:
    Debug.Print "[Periodenautomatik] Fehler " & Err.Number & " - " & Err.Description & _
                " | Zeile=" & r & " | Datum=" & CStr(datumWert) & _
                " | Kategorie=" & kategorie
    Err.Clear
    On Error Resume Next
    ws.Cells(r, BK_COL_MONAT_PERIODE).value = MonthName(Month(CDate(datumWert)))
    ws.Cells(r, BK_COL_MONAT_PERIODE).Interior.color = RGB(255, 235, 156)
    ws.Cells(r, BK_COL_BEMERKUNG).value = Trim$(CStr(ws.Cells(r, BK_COL_BEMERKUNG).value)) & _
        IIf(Trim$(CStr(ws.Cells(r, BK_COL_BEMERKUNG).value)) = "", "", vbLf) & _
        "Periodenautomatik Fehler - bitte Monat/Periode prüfen"
    Application.EnableEvents = eventsWaren
    On Error GoTo 0
    GoTo WeiterMitNaechsterZeile

SetzeMonatPeriodeError:
    Application.EnableEvents = eventsWaren
    Debug.Print "Fehler in SetzeMonatPeriode: " & Err.Number & " - " & Err.Description
    
End Sub


' ===============================================================
' FÄLLIGKEIT AUS KATEGORIE-TABELLE (Spalte O) HOLEN
' ===============================================================
Public Function HoleFaelligkeitFuerKategorie(ByVal wsDaten As Worksheet, _
                                              ByVal kategorie As String) As String
    Dim lastRow As Long
    Dim r As Long
    
    ' PRIO 1: Einstellungen-Blatt prüfen (Spalte B = Kategorie)
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
                    ' SollMonate leer -> Daten-Blatt Spalte O als Fallback prüfen
                    GoTo PruefeDatenBlatt
                Else
                    Dim anzMonate As Long
                    anzMonate = UBound(Split(SollMonate, ",")) + 1
                    Select Case anzMonate
                        Case 1: HoleFaelligkeitFuerKategorie = "j" & ChrW(228) & "hrlich"
                        Case 2: HoleFaelligkeitFuerKategorie = "halbj" & ChrW(228) & "hrlich"
                        Case 4: HoleFaelligkeitFuerKategorie = "quartalsweise"
                        Case Else: HoleFaelligkeitFuerKategorie = "monatlich"
                    End Select
                End If
                Exit Function
            End If
        Next r
    End If
    
PruefeDatenBlatt:
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

Public Function IstPeriodeFuerMonat(ByVal periode As String, ByVal kategorie As String, _
                                    ByVal monat As Long, ByVal jahr As Long, _
                                    ByVal istMonatlich As Boolean) As Boolean
    Dim text As String
    Dim erwarteterMonat As String
    Dim q As Long
    Dim h As Long
    Dim bereich As String

    text = LCase$(Trim$(periode))
    erwarteterMonat = LCase$(MonthName(monat))
    IstPeriodeFuerMonat = False
    If text = erwarteterMonat Then IstPeriodeFuerMonat = True: Exit Function
    If InStr(1, text, LCase$(kategorie), vbTextCompare) > 0 And InStr(text, CStr(jahr)) > 0 Then
        IstPeriodeFuerMonat = True
        Exit Function
    End If
    If text Like "q# " & CStr(jahr) Then
        q = CLng(mid$(text, 2, 1))
        IstPeriodeFuerMonat = (q = Int((monat - 1) / 3) + 1)
        Exit Function
    End If
    If text Like "h# " & CStr(jahr) Then
        h = CLng(mid$(text, 2, 1))
        IstPeriodeFuerMonat = ((h = 1 And monat <= 6) Or (h = 2 And monat >= 7))
        Exit Function
    End If
    If InStr(text, CStr(jahr)) > 0 Then
        If InStr(text, "januar bis m") > 0 Then IstPeriodeFuerMonat = (monat >= 1 And monat <= 3)
        If InStr(text, "april bis juni") > 0 Then IstPeriodeFuerMonat = (monat >= 4 And monat <= 6)
        If InStr(text, "juli bis september") > 0 Then IstPeriodeFuerMonat = (monat >= 7 And monat <= 9)
        If InStr(text, "oktober bis dezember") > 0 Then IstPeriodeFuerMonat = (monat >= 10 And monat <= 12)
    End If
End Function


' ===============================================================
' ÜBER-PÜNKTLICHE DAUERAUFTRÄGE EINMALIG KLÄREN
' ===============================================================
' Fachlicher Hintergrund:
' Eine Kategorie, die erst zum Monatsletzten fällig ist, hat ein
' Erkennungsproblem. Eine Zahlung am Monatsende sieht wie die
' pünktliche Zahlung des laufenden Monats aus. Zahlt ein Mitglied
' per Dauerauftrag aber schon am Ende des Vormonats, dann wurde die
' Januar-Zahlung bereits im Dezember des Vorjahres geleistet, und
' jede weitere Zahlung gehört einen Monat später. Ohne Klärung zieht
' sich diese Verschiebung durch das ganze Jahr: der Februar gilt als
' offen, obwohl er Ende Januar bezahlt wurde.
'
' Liegen Vorjahresdaten aus Oktober bis Dezember vor, belegen diese
' den Sachverhalt und es wird nichts gefragt. Nur beim ersten Lauf
' eines neuen Programms fehlt dieser Beleg. Genau dann wird einmal
' je Bankverbindung und Kategorie nachgefragt, nicht je Buchung.
'
' Eine bestätigte Antwort schreibt den Lernvermerk "Folgemonat
' manuell bestätigt". Ab dann greift die vorhandene Automatik in
' mod_KategorieEngine_Zeitraum von selbst und es wird nicht erneut
' gefragt.
' ===============================================================
Public Sub PruefeUeberpuenktlicheZahler(ByVal ws As Worksheet)

    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim schluessel As String
    Dim iban As String
    Dim kategorie As String
    Dim datumWert As Variant
    Dim ersteZeile As Object
    Dim erstesDatum As Object
    Dim bereitsGeklaert As Object
    Dim k As Variant
    Dim eventsWaren As Boolean
    Dim antwort As VbMsgBoxResult
    Dim buchDatum As Date
    Dim folgeMonatNr As Long
    Dim letzterTag As Long
    Dim betrag As Double
    Dim anzahlBestaetigt As Long

    On Error GoTo ZahlerFehler

    If ws Is Nothing Then Exit Sub

    ' Sind Vorjahresdaten vorhanden, ist der Fall belegt.
    If mod_Uebersicht_Daten.HatVorjahrDaten() Then Exit Sub

    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub

    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    Set ersteZeile = CreateObject("Scripting.Dictionary")
    Set erstesDatum = CreateObject("Scripting.Dictionary")
    Set bereitsGeklaert = CreateObject("Scripting.Dictionary")

    ' --- 1. Früheste Buchung je Bankverbindung und Kategorie suchen ---
    For r = BK_START_ROW To lastRow
        datumWert = ws.Cells(r, BK_COL_DATUM).value
        If Not IsDate(datumWert) Then GoTo NaechsteZahlerZeile

        kategorie = Trim$(CStr(ws.Cells(r, BK_COL_KATEGORIE).value))
        If kategorie = "" Then GoTo NaechsteZahlerZeile

        iban = UCase$(Replace(Trim$(CStr(ws.Cells(r, BK_COL_IBAN).value)), " ", ""))
        If iban = "" Then GoTo NaechsteZahlerZeile

        schluessel = iban & "|" & UCase$(kategorie)

        If IstZahlerFrageGeklaert(ws, r) Then
            If Not bereitsGeklaert.exists(schluessel) Then bereitsGeklaert.Add schluessel, True
        End If

        If Not erstesDatum.exists(schluessel) Then
            erstesDatum.Add schluessel, CDate(datumWert)
            ersteZeile.Add schluessel, r
        ElseIf CDate(datumWert) < CDate(erstesDatum(schluessel)) Then
            erstesDatum(schluessel) = CDate(datumWert)
            ersteZeile(schluessel) = r
        End If
NaechsteZahlerZeile:
    Next r

    ' --- 2. Je offenem Fall genau einmal fragen ---
    eventsWaren = Application.EnableEvents
    Application.EnableEvents = False

    For Each k In erstesDatum.keys
        schluessel = CStr(k)
        If bereitsGeklaert.exists(schluessel) Then GoTo NaechsterZahlerFall

        r = CLng(ersteZeile(schluessel))
        buchDatum = CDate(erstesDatum(schluessel))

        ' Nur Zahlungen kurz vor Monatsende sind überhaupt mehrdeutig.
        If Day(buchDatum) < 20 Then GoTo NaechsterZahlerFall

        kategorie = Trim$(CStr(ws.Cells(r, BK_COL_KATEGORIE).value))

        ' Ohne monatliche Fälligkeit gibt es keinen Folgemonat.
        If InStr(1, HoleFaelligkeitFuerKategorie(wsDaten, kategorie), _
                 "monatlich", vbTextCompare) = 0 Then GoTo NaechsterZahlerFall

        folgeMonatNr = Month(buchDatum) + 1
        If folgeMonatNr > 12 Then folgeMonatNr = 1
        letzterTag = Day(DateSerial(Year(buchDatum), Month(buchDatum) + 1, 0))

        betrag = 0
        If IsNumeric(ws.Cells(r, BK_COL_BETRAG).value) Then
            betrag = Abs(CDbl(ws.Cells(r, BK_COL_BETRAG).value))
        End If

        antwort = MsgBox( _
            "Zahlung am Monatsende - für welchen Monat gilt sie?" & vbCrLf & vbCrLf & _
            "Kontoinhaber: " & CStr(ws.Cells(r, BK_COL_NAME).value) & vbCrLf & _
            "Kategorie: " & kategorie & vbCrLf & _
            "Betrag: " & Format(betrag, "#,##0.00") & " " & ChrW(8364) & vbCrLf & _
            "Eingang: " & Format(buchDatum, "dd.mm.yyyy") & _
            "  (Tag " & Day(buchDatum) & " von " & letzterTag & ")" & vbCrLf & vbCrLf & _
            "Diese Kategorie ist erst zum Monatsletzten fällig. Die Zahlung" & vbCrLf & _
            "kann deshalb für " & MonthName(Month(buchDatum)) & " gelten - oder das Mitglied zahlt" & vbCrLf & _
            "per Dauerauftrag über-pünktlich bereits für " & MonthName(folgeMonatNr) & "." & vbCrLf & vbCrLf & _
            "Es liegen noch keine Vorjahresdaten aus Oktober bis Dezember" & vbCrLf & _
            "vor, die das belegen könnten. Daher diese einmalige Rückfrage." & vbCrLf & vbCrLf & _
            "Zahlt dieses Mitglied über-pünktlich für den Folgemonat?" & vbCrLf & vbCrLf & _
            "  Ja = alle Monatsend-Zahlungen einen Monat weiterschieben" & vbCrLf & _
            "  Nein = Zahlung gilt für den laufenden Monat" & vbCrLf & _
            "  Abbrechen = restliche Fälle überspringen", _
            vbYesNoCancel + vbQuestion, _
            "Über-pünktlicher Dauerauftrag?")

        If antwort = vbCancel Then Exit For

        If antwort = vbYes Then
            Call VerschiebeMonatsendzahlungen(ws, schluessel, lastRow)
            anzahlBestaetigt = anzahlBestaetigt + 1
        Else
            Call MerkeZahlerFrageGeklaert(ws, r)
        End If

NaechsterZahlerFall:
    Next k

    Application.EnableEvents = eventsWaren

    If anzahlBestaetigt > 0 Then
        Debug.Print "[Periodenautomatik] " & anzahlBestaetigt & _
                    " über-pünktliche(r) Dauerauftrag/Daueraufträge bestätigt."
    End If

    Exit Sub

ZahlerFehler:
    Application.EnableEvents = eventsWaren
    Debug.Print "Fehler in PruefeUeberpuenktlicheZahler: " & Err.Number & " - " & Err.Description

End Sub


' ===============================================================
' Wurde für diese Zeile bereits entschieden?
' Erkennbar an einem Lernvermerk, an einer manuellen Änderung oder
' an der ausdrücklichen Verneinung aus dieser Rückfrage.
' ===============================================================
Private Function IstZahlerFrageGeklaert(ByVal ws As Worksheet, _
                                        ByVal r As Long) As Boolean

    Dim bem As String
    bem = LCase$(CStr(ws.Cells(r, BK_COL_BEMERKUNG).value))

    IstZahlerFrageGeklaert = (InStr(bem, "folgemonat") > 0) Or _
                             (InStr(bem, "manuell ge") > 0) Or _
                             (InStr(bem, "monatsendzahlung gepr") > 0)

End Function


' ===============================================================
' Verneinung festhalten, damit nicht erneut gefragt wird.
' ===============================================================
Private Sub MerkeZahlerFrageGeklaert(ByVal ws As Worksheet, ByVal r As Long)

    Dim bem As String
    Dim hinweis As String

    hinweis = "Monatsendzahlung geprüft: gilt für den laufenden Monat"
    bem = Trim$(CStr(ws.Cells(r, BK_COL_BEMERKUNG).value))

    If bem = "" Then
        ws.Cells(r, BK_COL_BEMERKUNG).value = hinweis
    Else
        ws.Cells(r, BK_COL_BEMERKUNG).value = bem & vbLf & hinweis
    End If

End Sub


' ===============================================================
' Alle Monatsend-Zahlungen dieser Bankverbindung und Kategorie um
' einen Monat weiterschieben.
'
' Von Hand geänderte Zeilen bleiben unangetastet: was der Nutzer
' selbst gesetzt hat, darf die Automatik nicht überschreiben.
' Der Lernvermerk sorgt dafür, dass künftige Buchungen ohne weitere
' Rückfrage richtig zugeordnet werden.
' ===============================================================
Private Sub VerschiebeMonatsendzahlungen(ByVal ws As Worksheet, _
                                         ByVal schluessel As String, _
                                         ByVal lastRow As Long)

    Dim r As Long
    Dim iban As String
    Dim kategorie As String
    Dim datumWert As Variant
    Dim buchDatum As Date
    Dim folgeMonatNr As Long
    Dim neuerMonat As String
    Dim bem As String
    Dim vermerk As String

    For r = BK_START_ROW To lastRow
        datumWert = ws.Cells(r, BK_COL_DATUM).value
        If Not IsDate(datumWert) Then GoTo NaechsteVerschiebeZeile

        kategorie = Trim$(CStr(ws.Cells(r, BK_COL_KATEGORIE).value))
        iban = UCase$(Replace(Trim$(CStr(ws.Cells(r, BK_COL_IBAN).value)), " ", ""))
        If iban & "|" & UCase$(kategorie) <> schluessel Then GoTo NaechsteVerschiebeZeile

        buchDatum = CDate(datumWert)
        If Day(buchDatum) < 20 Then GoTo NaechsteVerschiebeZeile

        bem = LCase$(CStr(ws.Cells(r, BK_COL_BEMERKUNG).value))
        If InStr(bem, "manuell ge") > 0 Then GoTo NaechsteVerschiebeZeile

        folgeMonatNr = Month(buchDatum) + 1
        If folgeMonatNr > 12 Then folgeMonatNr = 1
        neuerMonat = MonthName(folgeMonatNr)

        ws.Cells(r, BK_COL_MONAT_PERIODE).value = neuerMonat
        ws.Cells(r, BK_COL_MONAT_PERIODE).Interior.color = RGB(198, 239, 206)

        ' Den alten Rückfragehinweis entfernen, er ist beantwortet.
        Call EntferneGelbHinweisPeriode(ws, r)

        vermerk = "Folgemonat manuell best" & ChrW(228) & "tigt: " & neuerMonat
        bem = Trim$(CStr(ws.Cells(r, BK_COL_BEMERKUNG).value))
        If InStr(1, bem, vermerk, vbTextCompare) = 0 Then
            If bem = "" Then
                ws.Cells(r, BK_COL_BEMERKUNG).value = vermerk
            Else
                ws.Cells(r, BK_COL_BEMERKUNG).value = bem & vbLf & vermerk
            End If
        End If
        ws.Cells(r, BK_COL_BEMERKUNG).Interior.ColorIndex = xlNone

NaechsteVerschiebeZeile:
    Next r

End Sub


' ===============================================================
' Den gelben Hinweis "Bitte prüfen ob Zahlung für ... gilt"
' entfernen, sobald die Frage beantwortet ist.
' ===============================================================
Private Sub EntferneGelbHinweisPeriode(ByVal ws As Worksheet, ByVal r As Long)

    Dim zeilen() As String
    Dim neu As String
    Dim i As Long

    zeilen = Split(CStr(ws.Cells(r, BK_COL_BEMERKUNG).value), vbLf)

    For i = LBound(zeilen) To UBound(zeilen)
        If InStr(1, zeilen(i), "Bitte pr", vbTextCompare) = 0 Or _
           InStr(1, zeilen(i), "Folgemonat gilt", vbTextCompare) = 0 Then
            If Trim$(zeilen(i)) <> "" Then
                If neu = "" Then
                    neu = zeilen(i)
                Else
                    neu = neu & vbLf & zeilen(i)
                End If
            End If
        End If
    Next i

    ws.Cells(r, BK_COL_BEMERKUNG).value = neu

End Sub








































































































































