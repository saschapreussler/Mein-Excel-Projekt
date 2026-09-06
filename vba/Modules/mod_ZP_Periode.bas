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
' MONAT/PERIODE SETZEN (ueberarbeitet)
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




































































































































