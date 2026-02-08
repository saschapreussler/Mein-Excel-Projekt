Attribute VB_Name = "mod_Banking_Data"
Option Explicit

' ***************************************************************
' MODUL: mod_Banking_Data
' VERSION: 3.8 - 09.02.2026
' Zentrales Modul fuer den CSV-Import von Bankdaten,
' Formatierung, Sortierung und ListBox-Verwaltung.
'
' FIXES/CHANGES:
' v3.7 -> v3.8: Setze_Monat_Periode nutzt Public
'   ErmittleMonatPeriode aus mod_KategorieEngine_Evaluator
'   mit Cache-Unterstuetzung (Folgemonat-Erkennung).
'   Private ErmittleMonatPeriode ENTFERNT (doppelt/veraltet).
'
' Abschnitte:
' 1. IMPORT-HAUPTFUNKTION (Importiere_Kontoauszug)
' 2. ENTITY-KEY-PRUEFUNG
' 3. FORMATIERUNG (Zebra, Border, Zahlenformate)
' 4. SORTIERUNG nach Datum
' 5. IBAN-Import aus Buchungen
' 6. MONAT/PERIODE ZUORDNUNG (Setze_Monat_Periode)
' 7. IMPORT REPORT LISTBOX (ActiveX)
' 8. HILFSFUNKTIONEN
' 9. SORTIERE TABELLEN DATEN
' ***************************************************************


' ===============================================================
' LISTBOX-KONFIGURATION (Farben, Speicher, Limits)
' ===============================================================
Private Const LB_COLOR_GRUEN As Long = 13561798    ' RGB(198, 239, 206)
Private Const LB_COLOR_GELB As Long = 10283775     ' RGB(255, 235, 156)
Private Const LB_COLOR_ROT As Long = 13485311      ' RGB(255, 199, 206)
Private Const LB_COLOR_WEISS As Long = 16777215    ' RGB(255, 255, 255)

Private Const FORM_LISTBOX_NAME As String = "lst_ImportReport"
Private Const PROTO_ZEILE As Long = 500
Private Const PROTO_SPALTE As Long = 25             ' Spalte Y
Private Const PROTO_SEP As String = "||"
Private Const MAX_ZEILEN As Long = 500              ' 100 Bloecke x 5 Zeilen


' ===============================================================
' 1. IMPORT-HAUPTFUNKTION
' ===============================================================
Public Sub Importiere_Kontoauszug()

    Dim ws As Worksheet
    Dim wsDaten As Worksheet
    Dim dateiPfad As String
    Dim ff As Integer
    Dim zeile As String
    Dim felder() As String
    Dim importZeile As Long
    Dim totalRows As Long
    Dim imported As Long
    Dim dupes As Long
    Dim failed As Long
    Dim errorRows As String
    Dim zeilenNr As Long
    Dim headerGefunden As Boolean
    Dim i As Long
    
    ' Datei auswaehlen
    dateiPfad = Application.GetOpenFilename( _
        FileFilter:="CSV-Dateien (*.csv),*.csv", _
        Title:="Kontoauszug CSV-Datei ausw" & ChrW(228) & "hlen")
    
    If dateiPfad = "Falsch" Or dateiPfad = "False" Or dateiPfad = "" Then Exit Sub
    
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' Naechste freie Zeile bestimmen
    importZeile = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row + 1
    If importZeile < BK_START_ROW Then importZeile = BK_START_ROW
    
    totalRows = 0
    imported = 0
    dupes = 0
    failed = 0
    errorRows = ""
    
    ' CSV oeffnen und verarbeiten
    ff = FreeFile
    Open dateiPfad For Input As #ff
    
    headerGefunden = False
    zeilenNr = 0
    
    Do While Not EOF(ff)
        Line Input #ff, zeile
        zeilenNr = zeilenNr + 1
        
        ' Header-Zeile erkennen und ueberspringen
        If Not headerGefunden Then
            If InStr(LCase(zeile), "buchungstag") > 0 Or _
               InStr(LCase(zeile), "buchungsdatum") > 0 Or _
               InStr(LCase(zeile), "valuta") > 0 Then
                headerGefunden = True
            End If
            GoTo NaechsteZeile
        End If
        
        ' Leere Zeilen ueberspringen
        If Len(Trim(zeile)) = 0 Then GoTo NaechsteZeile
        
        totalRows = totalRows + 1
        
        ' CSV parsen (Semikolon-getrennt)
        felder = SplitCSV(zeile, ";")
        
        ' Mindestens 7 Felder erwartet
        If UBound(felder) < 6 Then
            failed = failed + 1
            errorRows = errorRows & zeilenNr & ", "
            GoTo NaechsteZeile
        End If
        
        ' Datum parsen
        Dim buchungsDatum As Variant
        buchungsDatum = ParseDatum(felder(0))
        
        If isEmpty(buchungsDatum) Then
            failed = failed + 1
            errorRows = errorRows & zeilenNr & ", "
            GoTo NaechsteZeile
        End If
        
        ' Betrag parsen
        Dim betragStr As String
        betragStr = felder(4)
        If UBound(felder) >= 5 Then
            If InStr(felder(4), ",") = 0 And InStr(felder(5), ",") > 0 Then
                betragStr = felder(4) & "," & felder(5)
            End If
        End If
        
        Dim betrag As Double
        betrag = ParseBetrag(betragStr)
        
        ' Duplikatpruefung
        Dim isDuplicate As Boolean
        isDuplicate = False
        
        Dim checkRow As Long
        For checkRow = BK_START_ROW To importZeile - 1
            If IsDate(ws.Cells(checkRow, BK_COL_DATUM).value) Then
                If CDate(ws.Cells(checkRow, BK_COL_DATUM).value) = CDate(buchungsDatum) And _
                   ws.Cells(checkRow, BK_COL_BETRAG).value = betrag And _
                   Trim(ws.Cells(checkRow, BK_COL_VERWENDUNGSZWECK).value) = CleanField(felder(3)) Then
                    isDuplicate = True
                    Exit For
                End If
            End If
        Next checkRow
        
        If isDuplicate Then
            dupes = dupes + 1
            GoTo NaechsteZeile
        End If
        
        ' Daten eintragen
        ws.Cells(importZeile, BK_COL_DATUM).value = CDate(buchungsDatum)
        ws.Cells(importZeile, BK_COL_BETRAG).value = betrag
        ws.Cells(importZeile, BK_COL_NAME).value = CleanField(felder(2))
        
        ' IBAN aus Feld 6 oder 7 (je nach Format)
        Dim ibanText As String
        ibanText = ""
        For i = 5 To UBound(felder)
            If LikeIBAN(felder(i)) Then
                ibanText = CleanField(felder(i))
                Exit For
            End If
        Next i
        ws.Cells(importZeile, BK_COL_IBAN).value = ibanText
        
        ws.Cells(importZeile, BK_COL_VERWENDUNGSZWECK).value = CleanField(felder(3))
        ws.Cells(importZeile, BK_COL_BUCHUNGSTEXT).value = CleanField(felder(1))
        
        ' Sichtbarkeitsfilter (Spalte G) auf TRUE setzen
        ws.Cells(importZeile, 7).value = True
        
        imported = imported + 1
        importZeile = importZeile + 1
        
NaechsteZeile:
    Loop
    
    Close #ff
    
    ' -------------------------------------------------------
    ' NACHVERARBEITUNG (nur wenn Daten importiert wurden)
    ' -------------------------------------------------------
    If imported > 0 Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        
        On Error GoTo ImportCleanup
        
        ' 1. IBANs in Entity-Tabelle uebernehmen
        Call ImportiereIBANs(ws)
        
        ' 2. Entity-Keys aktualisieren
        Call AktualisiereEntityKeys(ws)
        
        ' 3. Sortierung nach Datum
        Call SortiereBankkontoNachDatum(ws)
        
        ' 4. Formatierung
        Call Anwende_Zebra_Bankkonto(ws)
        Call Anwende_Border_Bankkonto(ws)
        Call Anwende_Formatierung_Bankkonto(ws)
        
        ' 5. Kategorie-Engine
        Call KategorieEngine_Pipeline(ws)
        
        ' 6. Monat/Periode zuordnen
        Call Setze_Monat_Periode(ws)
        
        ' 7. Formeln wiederherstellen
        Call StelleFormelnWiederHer(ws)
        
ImportCleanup:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' ListBox aktualisieren
    Call Update_ImportReport_ListBox(totalRows, imported, dupes, failed)
    
    ' Ergebnis anzeigen
    Dim msg As String
    msg = imported & " von " & totalRows & " Datens" & ChrW(228) & "tzen importiert." & vbCrLf
    If dupes > 0 Then msg = msg & dupes & " Duplikate erkannt." & vbCrLf
    If failed > 0 Then msg = msg & failed & " Fehler (Zeilen: " & Left(errorRows, Len(errorRows) - 2) & ")"
    
    MsgBox msg, IIf(failed > 0, vbExclamation, vbInformation), "Import abgeschlossen"
    
End Sub


' ===============================================================
' 2. ENTITY-KEY-PRUEFUNG
' ===============================================================
Private Sub PruefeUnvollstaendigeEntityKeys(ByVal ws As Worksheet)
    ' Wird bei Bedarf aufgerufen - aktuell leer (Platzhalter)
End Sub


' ===============================================================
' 3. FORMATIERUNG
' ===============================================================

' ---------------------------------------------------------------
' 3a. Zebra-Streifen (abwechselnde Zeilenfarben)
' ---------------------------------------------------------------
Private Sub Anwende_Zebra_Bankkonto(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    For r = BK_START_ROW To lastRow
        If (r - BK_START_ROW) Mod 2 = 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.color = RGB(242, 242, 242)
        Else
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.color = RGB(255, 255, 255)
        End If
    Next r
End Sub

' ---------------------------------------------------------------
' 3b. Rahmenlinien
' ---------------------------------------------------------------
Private Sub Anwende_Border_Bankkonto(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim rng As Range
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    Set rng = ws.Range(ws.Cells(BK_START_ROW, 1), ws.Cells(lastRow, 26))
    
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .color = RGB(200, 200, 200)
    End With
End Sub

' ---------------------------------------------------------------
' 3c. Zahlenformate und Spaltenbreiten
' ---------------------------------------------------------------
Private Sub Anwende_Formatierung_Bankkonto(ByVal ws As Worksheet)
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    ' Datum formatieren
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_DATUM), ws.Cells(lastRow, BK_COL_DATUM)).NumberFormat = "DD.MM.YYYY"
    
    ' Betrag formatieren
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_BETRAG), ws.Cells(lastRow, BK_COL_BETRAG)).NumberFormat = "#,##0.00"
    
    ' Betragsspalten M-Z formatieren
    ws.Range(ws.Cells(BK_START_ROW, 13), ws.Cells(lastRow, 26)).NumberFormat = "#,##0.00"
End Sub


' ===============================================================
' 4. SORTIERUNG NACH DATUM
' ===============================================================
Private Sub SortiereBankkontoNachDatum(ByVal ws As Worksheet)
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Cells(BK_START_ROW, BK_COL_DATUM), _
                             Order:=xlAscending
        .SetRange ws.Range(ws.Cells(BK_START_ROW, 1), ws.Cells(lastRow, 26))
        .Header = xlNo
        .Apply
    End With
End Sub


' ===============================================================
' 5. IBAN-IMPORT AUS BUCHUNGEN
' ===============================================================
Private Sub ImportiereIBANs(ByVal ws As Worksheet)
    Dim wsDaten As Worksheet
    Dim lastRowBK As Long
    Dim lastRowMap As Long
    Dim r As Long
    Dim iban As String
    Dim name As String
    Dim found As Boolean
    Dim mapRow As Long
    
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    
    On Error Resume Next
    wsDaten.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    lastRowBK = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    For r = BK_START_ROW To lastRowBK
        iban = Trim(ws.Cells(r, BK_COL_IBAN).value)
        If iban = "" Then GoTo NaechsteIBAN
        
        name = Trim(ws.Cells(r, BK_COL_NAME).value)
        
        ' Pruefen ob IBAN schon existiert
        lastRowMap = wsDaten.Cells(wsDaten.Rows.count, DATA_MAP_COL_IBAN).End(xlUp).Row
        found = False
        
        For mapRow = DATA_START_ROW To lastRowMap
            If UCase(Replace(wsDaten.Cells(mapRow, DATA_MAP_COL_IBAN).value, " ", "")) = _
               UCase(Replace(iban, " ", "")) Then
                found = True
                Exit For
            End If
        Next mapRow
        
        If Not found Then
            ' Neue IBAN eintragen
            lastRowMap = lastRowMap + 1
            wsDaten.Cells(lastRowMap, DATA_MAP_COL_IBAN).value = iban
            wsDaten.Cells(lastRowMap, DATA_MAP_COL_NAME).value = name
        End If
        
NaechsteIBAN:
    Next r
    
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
End Sub


' ===============================================================
' 5b. ENTITY-KEYS AKTUALISIEREN
' ===============================================================
Private Sub AktualisiereEntityKeys(ByVal ws As Worksheet)
    ' Delegiert an mod_EntityKey_Manager
    On Error Resume Next
    Call AktualisiereAlleEntityKeys
    On Error GoTo 0
End Sub


' ===============================================================
' 6. MONAT/PERIODE ZUORDNUNG
'    v3.8: Nutzt Public ErmittleMonatPeriode aus
'    mod_KategorieEngine_Evaluator mit Cache-Unterstuetzung.
' ===============================================================
Private Sub Setze_Monat_Periode(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim r As Long
    Dim monatWert As Variant
    Dim datumWert As Variant
    Dim kategorie As String
    Dim faelligkeit As String
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    ' Faelligkeit aus Kategorie-Tabelle vorladen
    Dim wsDaten As Worksheet
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    
    ' v3.8: Einstellungen-Cache laden fuer Folgemonat-Erkennung
    Call LadeEinstellungenCache
    
    For r = BK_START_ROW To lastRow
        datumWert = ws.Cells(r, BK_COL_DATUM).value
        monatWert = ws.Cells(r, BK_COL_MONAT_PERIODE).value
        
        If IsDate(datumWert) And (isEmpty(monatWert) Or monatWert = "") Then
            kategorie = Trim(ws.Cells(r, BK_COL_KATEGORIE).value)
            
            If kategorie <> "" Then
                ' Faelligkeit aus Kategorie-Tabelle holen (Spalte O)
                faelligkeit = HoleFaelligkeitFuerKategorie(wsDaten, kategorie)
                ' v3.8: Nutzt Public Version aus Evaluator (mit Cache + Folgemonat)
                ws.Cells(r, BK_COL_MONAT_PERIODE).value = _
                    ErmittleMonatPeriode(kategorie, CDate(datumWert), faelligkeit)
            Else
                ' Keine Kategorie: Fallback auf Buchungsmonat
                ws.Cells(r, BK_COL_MONAT_PERIODE).value = MonthName(Month(datumWert))
            End If
        End If
    Next r
    
    ' v3.8: Einstellungen-Cache wieder freigeben
    Call EntladeEinstellungenCache
    
End Sub

' ---------------------------------------------------------------
' 6b. Faelligkeit aus Kategorie-Tabelle (Spalte O) holen
' ---------------------------------------------------------------
Private Function HoleFaelligkeitFuerKategorie(ByVal wsDaten As Worksheet, _
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

' ===============================================================
' 7. IMPORT REPORT LISTBOX (ACTIVEX STEUERELEMENT)
'    -----------------------------------------------
'    Architektur:
'    - ActiveX ListBox "lst_ImportReport" auf Bankkonto-Blatt
'    - Speicher: Daten!Y500 (eine einzige Zelle, serialisiert
'      mit "||" als Trennzeichen zwischen Zeilen)
'    - Befuellung: .Clear / .AddItem (ActiveX-Methoden)
'    - Hintergrundfarbe: .BackColor direkt auf der ListBox
'    - Pro Import-Vorgang: 5 Zeilen (Datum, X/Y, Dupes, Fehler, ----)
'    - Max 100 Bloecke = 500 Zeilen Historie
'    - WICHTIG: EnableEvents=False beim Schreiben in Daten!Y500
'      um Worksheet_Change-Kaskade zu verhindern
'    - WICHTIG: Position/Groesse werden VOR .Clear gesichert
'      und NACH .AddItem wiederhergestellt, da ActiveX-ListBox
'      .AddItem die OLE-Container-Groesse veraendern kann.
'      Der Designer bestimmt die Ausgangsgroesse.
' ===============================================================

' ---------------------------------------------------------------
' 7a. Initialize: Liest Y500, befuellt ActiveX ListBox,
'     setzt Hintergrundfarbe.
'     Aufruf: Workbook_Open, Worksheet_Activate, nach Loeschen
' ---------------------------------------------------------------
Public Sub Initialize_ImportReport_ListBox()
    
    Dim wsBK As Worksheet
    Dim wsDaten As Worksheet
    Dim lb As MSForms.ListBox
    Dim oleObj As OLEObject
    Dim gespeichert As String
    Dim zeilen() As String
    Dim anzahl As Long
    Dim i As Long
    Dim savLeft As Double, savTop As Double
    Dim savWidth As Double, savHeight As Double
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsBK Is Nothing Or wsDaten Is Nothing Then Exit Sub
    
    ' OLEObject holen und Position/Groesse VORHER sichern
    On Error Resume Next
    Set oleObj = wsBK.OLEObjects(FORM_LISTBOX_NAME)
    On Error GoTo 0
    If oleObj Is Nothing Then Exit Sub
    
    savLeft = oleObj.Left
    savTop = oleObj.Top
    savWidth = oleObj.Width
    savHeight = oleObj.Height
    
    ' Placement auf freifliegend setzen
    On Error Resume Next
    oleObj.Placement = xlFreeFloating
    On Error GoTo 0
    
    ' ActiveX ListBox holen
    On Error Resume Next
    Set lb = oleObj.Object
    On Error GoTo 0
    If lb Is Nothing Then Exit Sub
    
    ' ListBox leeren
    lb.Clear
    
    ' Gespeichertes Protokoll aus Y500 lesen
    gespeichert = CStr(wsDaten.Cells(PROTO_ZEILE, PROTO_SPALTE).value)
    
    If gespeichert = "" Or gespeichert = "0" Then
        ' Kein Protokoll vorhanden - Standardtext
        lb.AddItem "Kein Status Report"
        lb.AddItem "vorhanden."
        lb.BackColor = LB_COLOR_WEISS
    Else
        ' Protokoll-Zeilen aus Y500 deserialisieren und einfuegen
        zeilen = Split(gespeichert, PROTO_SEP)
        anzahl = UBound(zeilen) + 1
        If anzahl > MAX_ZEILEN Then anzahl = MAX_ZEILEN
        
        For i = 0 To anzahl - 1
            lb.AddItem zeilen(i)
        Next i
        
        ' Farbe aus juengstem Block bestimmen
        Call FaerbeListBoxAusProtokoll(lb, zeilen)
    End If
    
    ' Position und Groesse WIEDERHERSTELLEN (AddItem kann sie aendern)
    On Error Resume Next
    oleObj.Left = savLeft
    oleObj.Top = savTop
    oleObj.Width = savWidth
    oleObj.Height = savHeight
    On Error GoTo 0
    
End Sub

' ---------------------------------------------------------------
' 7b. Update: Neuen 5-Zeilen-Block OBEN einfuegen,
'     in Y500 serialisiert speichern, ListBox aktualisieren.
' ---------------------------------------------------------------
Private Sub Update_ImportReport_ListBox(ByVal totalRows As Long, ByVal imported As Long, _
                                         ByVal dupes As Long, ByVal failed As Long)
    
    Dim wsBK As Worksheet
    Dim wsDaten As Worksheet
    Dim lb As MSForms.ListBox
    Dim oleObj As OLEObject
    Dim altGespeichert As String
    Dim neuerBlock As String
    Dim gesamt As String
    Dim zeilen() As String
    Dim anzahl As Long
    Dim i As Long
    Dim eventsWaren As Boolean
    Dim savLeft As Double, savTop As Double
    Dim savWidth As Double, savHeight As Double
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsBK Is Nothing Or wsDaten Is Nothing Then Exit Sub
    
    ' OLEObject holen und Position/Groesse VORHER sichern
    On Error Resume Next
    Set oleObj = wsBK.OLEObjects(FORM_LISTBOX_NAME)
    On Error GoTo 0
    If oleObj Is Nothing Then Exit Sub
    
    savLeft = oleObj.Left
    savTop = oleObj.Top
    savWidth = oleObj.Width
    savHeight = oleObj.Height
    
    ' Placement auf freifliegend setzen
    On Error Resume Next
    oleObj.Placement = xlFreeFloating
    On Error GoTo 0
    
    ' --- 5-Zeilen-Block zusammenbauen ---
    neuerBlock = "Import: " & Format(Now, "DD.MM.YYYY  HH:MM:SS") & _
                 PROTO_SEP & _
                 imported & " / " & totalRows & " Datens" & ChrW(228) & "tze importiert" & _
                 PROTO_SEP & _
                 dupes & " Duplikate erkannt" & _
                 PROTO_SEP & _
                 failed & " Fehler" & _
                 PROTO_SEP & _
                 "--------------------------------------"
    
    ' --- WICHTIG: Events deaktivieren BEVOR in Daten geschrieben wird ---
    eventsWaren = Application.EnableEvents
    Application.EnableEvents = False
    
    ' --- Daten-Blatt entsperren ---
    On Error Resume Next
    wsDaten.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' --- Alten Inhalt aus Y500 laden ---
    altGespeichert = CStr(wsDaten.Cells(PROTO_ZEILE, PROTO_SPALTE).value)
    
    If altGespeichert = "" Or altGespeichert = "0" Then
        gesamt = neuerBlock
    Else
        gesamt = neuerBlock & PROTO_SEP & altGespeichert
    End If
    
    ' --- Auf MAX_ZEILEN begrenzen ---
    zeilen = Split(gesamt, PROTO_SEP)
    anzahl = UBound(zeilen) + 1
    If anzahl > MAX_ZEILEN Then
        gesamt = zeilen(0)
        For i = 1 To MAX_ZEILEN - 1
            gesamt = gesamt & PROTO_SEP & zeilen(i)
        Next i
        anzahl = MAX_ZEILEN
    End If
    
    ' --- In Y500 speichern (eine einzige Zelle!) ---
    wsDaten.Cells(PROTO_ZEILE, PROTO_SPALTE).value = gesamt
    
    ' --- Daten-Blatt schuetzen ---
    On Error Resume Next
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
    ' --- Events wieder herstellen ---
    Application.EnableEvents = eventsWaren
    
    ' --- ActiveX ListBox aktualisieren ---
    On Error Resume Next
    Set lb = oleObj.Object
    On Error GoTo 0
    
    If Not lb Is Nothing Then
        lb.Clear
        zeilen = Split(gesamt, PROTO_SEP)
        For i = 0 To anzahl - 1
            lb.AddItem zeilen(i)
        Next i
        
        ' Farbcodierung
        Call FaerbeListBoxNachImport(lb, imported, dupes, failed)
    End If
    
    ' Position und Groesse WIEDERHERSTELLEN (AddItem kann sie aendern)
    On Error Resume Next
    oleObj.Left = savLeft
    oleObj.Top = savTop
    oleObj.Width = savWidth
    oleObj.Height = savHeight
    On Error GoTo 0
    
End Sub

' ---------------------------------------------------------------
' 7c. Farbcodierung nach Import-Ergebnis (direkt auf ListBox)
'     GRUEN  = Alles OK (dupes = 0, failed = 0)
'     GELB   = Duplikate vorhanden (dupes > 0, failed = 0)
'     ROT    = Fehler vorhanden (failed > 0)
' ---------------------------------------------------------------
Private Sub FaerbeListBoxNachImport(ByVal lb As MSForms.ListBox, _
                                     ByVal imported As Long, _
                                     ByVal dupes As Long, _
                                     ByVal failed As Long)
    
    If failed > 0 Then
        lb.BackColor = LB_COLOR_ROT
    ElseIf dupes > 0 Then
        lb.BackColor = LB_COLOR_GELB
    Else
        lb.BackColor = LB_COLOR_GRUEN
    End If
    
End Sub

' ---------------------------------------------------------------
' 7d. Farbcodierung aus gespeichertem Protokoll bestimmen
'     Liest Index 2: "X Duplikate erkannt"
'     Liest Index 3: "X Fehler"
' ---------------------------------------------------------------
Private Sub FaerbeListBoxAusProtokoll(ByVal lb As MSForms.ListBox, ByRef zeilen() As String)
    
    Dim dupes As Long
    Dim failed As Long
    
    If UBound(zeilen) < 3 Then
        lb.BackColor = LB_COLOR_WEISS
        Exit Sub
    End If
    
    dupes = ExtrahiereZahl(CStr(zeilen(2)))
    failed = ExtrahiereZahl(CStr(zeilen(3)))
    
    If failed > 0 Then
        lb.BackColor = LB_COLOR_ROT
    ElseIf dupes > 0 Then
        lb.BackColor = LB_COLOR_GELB
    Else
        lb.BackColor = LB_COLOR_GRUEN
    End If
    
End Sub

' ---------------------------------------------------------------
' 7e. Zahl am Anfang eines Strings extrahieren
'     "123 Duplikate erkannt" -> 123
' ---------------------------------------------------------------
Private Function ExtrahiereZahl(ByVal text As String) As Long
    
    Dim i As Long
    Dim zahlStr As String
    
    zahlStr = ""
    For i = 1 To Len(text)
        If Mid(text, i, 1) >= "0" And Mid(text, i, 1) <= "9" Then
            zahlStr = zahlStr & Mid(text, i, 1)
        Else
            If zahlStr <> "" Then Exit For
        End If
    Next i
    
    If zahlStr <> "" Then
        ExtrahiereZahl = CLng(zahlStr)
    Else
        ExtrahiereZahl = 0
    End If
    
End Function


' ===============================================================
' 8. HILFSFUNKTIONEN
' ===============================================================

' ===============================================================
' Stellt die Formeln auf dem Bankkonto-Blatt wieder her,
' die durch ClearContents oder Import verloren gehen koennen.
' Betrifft: C3, E8-E14, E16-E21, E23
' WICHTIG: Formeln werden 1:1 als FormulaLocal gesetzt!
' ===============================================================
Private Sub StelleFormelnWiederHer(ByVal ws As Worksheet)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    On Error Resume Next
    
    ' C3: Kontostand-Anzeige mit Monatsfilter
    ws.Range("C3").FormulaLocal = _
        "=WENN(Daten!$AE$4=0;WENN(ANZAHL(Bankkonto!$A$28:$A$3433)=0;"""";" & _
        """Kontostand nach der letzten Buchung im Monat am: "" & TEXT(MAX(Bankkonto!$A$28:$A$5000);""TT.MM.JJJJ""));" & _
        "WENN(Z" & ChrW(196) & "HLENWENNS(Bankkonto!$A$28:$A$5000;"">="" & DATUM(Startmen" & ChrW(252) & "!$F$1;Daten!$AE$4;1);" & _
        "Bankkonto!$A$28:$A$5000;""<="" & DATUM(Startmen" & ChrW(252) & "!$F$1;Daten!$AE$4+1;0))=0;"""";" & _
        """Kontostand nach der letzten Buchung im Monat am: "" & TEXT(MAXWENNS(Bankkonto!$A$28:$A$5000;" & _
        "Bankkonto!$A$28:$A$5000;"">="" & DATUM(Startmen" & ChrW(252) & "!$F$1;Daten!$AE$4;1);" & _
        "Bankkonto!$A$28:$A$5000;""<="" & DATUM(Startmen" & ChrW(252) & "!$F$1;Daten!$AE$4+1;0));""TT.MM.JJJJ"")))"
    
    ' E8-E14: Einnahmen (Spalten M-S) mit SUMMEWENNS + WENN=0 leer
    ws.Range("E8").FormulaLocal = _
        "=WENN(SUMMEWENNS(M28:M5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(M28:M5000;G28:G5000;WAHR))"
    ws.Range("E9").FormulaLocal = _
        "=WENN(SUMMEWENNS(N28:N5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(N28:N5000;G28:G5000;WAHR))"
    ws.Range("E10").FormulaLocal = _
        "=WENN(SUMMEWENNS(O28:O5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(O28:O5000;G28:G5000;WAHR))"
    ws.Range("E11").FormulaLocal = _
        "=WENN(SUMMEWENNS(P28:P5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(P28:P5000;G28:G5000;WAHR))"
    ws.Range("E12").FormulaLocal = _
        "=WENN(SUMMEWENNS(Q28:Q5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(Q28:Q5000;G28:G5000;WAHR))"
    ws.Range("E13").FormulaLocal = _
        "=WENN(SUMMEWENNS(R28:R5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(R28:R5000;G28:G5000;WAHR))"
    ws.Range("E14").FormulaLocal = _
        "=WENN(SUMMEWENNS(S28:S5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(S28:S5000;G28:G5000;WAHR))"
    
    ' E16-E21: Ausgaben (Spalten T-Y) mit SUMMEWENNS + WENN=0 leer
    ws.Range("E16").FormulaLocal = _
        "=WENN(SUMMEWENNS(T28:T5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(T28:T5000;G28:G5000;WAHR))"
    ws.Range("E17").FormulaLocal = _
        "=WENN(SUMMEWENNS(U28:U5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(U28:U5000;G28:G5000;WAHR))"
    ws.Range("E18").FormulaLocal = _
        "=WENN(SUMMEWENNS(V28:V5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(V28:V5000;G28:G5000;WAHR))"
    ws.Range("E19").FormulaLocal = _
        "=WENN(SUMMEWENNS(W28:W5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(W28:W5000;G28:G5000;WAHR))"
    ws.Range("E20").FormulaLocal = _
        "=WENN(SUMMEWENNS(X28:X5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(X28:X5000;G28:G5000;WAHR))"
    ws.Range("E21").FormulaLocal = _
        "=WENN(SUMMEWENNS(Y28:Y5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(Y28:Y5000;G28:G5000;WAHR))"
    
    ' E23: Auszahlung Kasse (Spalte Z)
    ws.Range("E23").FormulaLocal = _
        "=WENN(SUMMEWENNS(Z28:Z5000;G28:G5000;WAHR)=0;"""";SUMMEWENNS(Z28:Z5000;G28:G5000;WAHR))"
    
    On Error GoTo 0
    
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
End Sub

' ---------------------------------------------------------------
' 8b. Alle Bankkontozeilen loeschen
' ---------------------------------------------------------------
Public Sub LoescheAlleBankkontoZeilen()
    
    Dim ws As Worksheet
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim antwort As VbMsgBoxResult
    Dim eventsWaren As Boolean
    
    antwort = MsgBox("ACHTUNG: Alle Daten auf dem Bankkonto-Blatt werden geloescht!" & vbCrLf & vbCrLf & _
                     "Fortfahren?", vbYesNo + vbCritical, "Alle Daten loeschen?")
    
    If antwort <> vbYes Then Exit Sub
    
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    If lastRow >= BK_START_ROW Then
        ws.Range(ws.Cells(BK_START_ROW, 1), ws.Cells(lastRow, 26)).ClearContents
        ws.Range(ws.Cells(BK_START_ROW, 1), ws.Cells(lastRow, 26)).Interior.ColorIndex = xlNone
    End If
    
    ' Formeln wiederherstellen (wurden durch ClearContents geloescht)
    Call StelleFormelnWiederHer(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' Protokoll-Speicher leeren (Events aus!)
    eventsWaren = Application.EnableEvents
    Application.EnableEvents = False
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    If Not wsDaten Is Nothing Then
        wsDaten.Unprotect PASSWORD:=PASSWORD
        wsDaten.Cells(PROTO_ZEILE, PROTO_SPALTE).ClearContents
        wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    On Error GoTo 0
    
    Application.EnableEvents = eventsWaren
    
    Call Initialize_ImportReport_ListBox
    
    MsgBox "Alle Daten wurden geloescht.", vbInformation
    
End Sub

' ---------------------------------------------------------------
' 8c. Formatierung Bankkonto aktualisieren
' ---------------------------------------------------------------
Public Sub AktualisiereFormatierungBankkonto()
    
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    Call Anwende_Zebra_Bankkonto(ws)
    Call Anwende_Border_Bankkonto(ws)
    Call Anwende_Formatierung_Bankkonto(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    MsgBox "Formatierung aktualisiert!", vbInformation
    
End Sub

' ===============================================================
' 9. SORTIERE TABELLEN DATEN
' ===============================================================
Public Sub Sortiere_Tabellen_Daten()

    Dim ws As Worksheet
    Dim lr As Long
    
    Application.EnableEvents = False
    On Error GoTo ExitClean

    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ExitClean

    lr = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lr >= DATA_START_ROW Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add key:=ws.Cells(DATA_START_ROW, DATA_CAT_COL_KATEGORIE), _
                                 Order:=xlAscending
            .SetRange ws.Range(ws.Cells(DATA_START_ROW, DATA_CAT_COL_START), ws.Cells(lr, DATA_CAT_COL_END))
            .Header = xlNo
            .Apply
        End With
    End If

    lr = ws.Cells(ws.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    If lr >= EK_START_ROW Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add key:=ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                                 Order:=xlAscending
            .SetRange ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), ws.Cells(lr, EK_COL_DEBUG))
            .Header = xlNo
            .Apply
        End With
    End If
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True

ExitClean:
    Application.EnableEvents = True
End Sub


