VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ===============================================================
' Modulweite Variablen ffuer den Vergleich (alter Wert)
' Werden in Worksheet_SelectionChange gespeichert und in
' Worksheet_Change ausgewertet, um unnoetige Verarbeitung
' bei "Doppelklick + Enter ohne Aenderung" zu verhindern.
' ===============================================================
Private m_AlterWertH As String
Private m_AlterWertI As String

Private Sub btn_CSVimport_Click()
    ' Ruft die Hauptprozedur im Standardmodul mod_Datenpflege auf
    Call Importiere_Kontoauszug
End Sub

Private Sub Worksheet_Activate()

    ' Events ABSCHALTEN bevor die ComboBox beFuellt wird!
    ' Sonst loest .ListIndex = 0 sofort cbo_RechnungsMonat_Change aus,
    ' was bei fehlenden Daten die ComboBox wieder leert (Endlosschleife).
    Application.EnableEvents = False

    Dim arrMonate(1 To 13) As String
    
    arrMonate(1) = "ganzes Jahr"
    arrMonate(2) = "Januar"
    arrMonate(3) = "Februar"
    arrMonate(4) = "M" & ChrW(228) & "rz"
    arrMonate(5) = "April"
    arrMonate(6) = "Mai"
    arrMonate(7) = "Juni"
    arrMonate(8) = "Juli"
    arrMonate(9) = "August"
    arrMonate(10) = "September"
    arrMonate(11) = "Oktober"
    arrMonate(12) = "November"
    arrMonate(13) = "Dezember"
    
    With Me.cbo_RechnungsMonat
        .Clear
        .List = Application.Transpose(arrMonate)
        .ListIndex = 0
        
        If TypeName(Me.cbo_RechnungsMonat) = "ComboBox" Then
            .Style = 2 ' fmStyleDropDownList
        End If
    End With

    Application.EnableEvents = True

    Call mod_Banking_Report.Initialize_ImportReport_ListBox

    ' Hinweispruefung beim Oeffnen: Januar-Fehlbetraege ohne Vorjahrdaten.
    On Error Resume Next
    Call mod_Uebersicht_Generator.PruefeVorjahrHinweisBeimOeffnen
    On Error GoTo 0

    ' Fenster fixieren: Zeilen 1-3 bleiben stehen, ab Zeile 4 scrollbar.
    On Error Resume Next
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.SplitColumn = 0
        ActiveWindow.SplitRow = 3
        ActiveWindow.FreezePanes = True
    End If
    On Error GoTo 0
End Sub

Private Sub cbo_RechnungsMonat_Change()

    Dim wsBank As Worksheet
    Dim wsDaten As Worksheet
    
    Dim jahr As Long
    Dim monatsIndex As Long
    Dim letzteDatenZeile As Long
    
    Const ersteDatenZeile As Long = 30
    Const ersteZeileFilterbereich As Long = 29
    Const letzteSpalteFilterbereich As Long = 26
    
    Dim rngFilterBereich As Range
    Dim visibleCellsCount As Long
    
    Set wsBank = ThisWorkbook.Sheets("Bankkonto")
    Set wsDaten = ThisWorkbook.Sheets("Daten")

    On Error GoTo SafeExit
    Application.EnableEvents = False

    With wsBank
    
        ' Blattschutz aufheben (AutoFilter erfordert ungeschuetztes Blatt)
        On Error Resume Next
        .Unprotect PASSWORD:=PASSWORD
        On Error GoTo SafeExit
    
        ' Vorherige Kriterien aufheben, aber die Filter-Pfeile behalten.
        On Error Resume Next
        If .FilterMode Then .ShowAllData
        On Error GoTo SafeExit
        
        If Me.cbo_RechnungsMonat.ListIndex = -1 Then GoTo SafeExit
        
        ' v6.0: Abrechnungsjahr aus Einstellungen lesen
        jahr = HoleAbrechnungsjahr()
        If jahr = 0 Then
            MsgBox "Fehler: Kein Abrechnungsjahr auf dem Blatt Einstellungen hinterlegt.", vbCritical
            GoTo SafeExit
        End If
        
        ' Monatsindex bestimmen
        Select Case Me.cbo_RechnungsMonat.value
            Case "ganzes Jahr"
                monatsIndex = 0
            Case "Januar": monatsIndex = 1
            Case "Februar": monatsIndex = 2
            Case "M" & ChrW(228) & "rz": monatsIndex = 3
            Case "April": monatsIndex = 4
            Case "Mai": monatsIndex = 5
            Case "Juni": monatsIndex = 6
            Case "Juli": monatsIndex = 7
            Case "August": monatsIndex = 8
            Case "September": monatsIndex = 9
            Case "Oktober": monatsIndex = 10
            Case "November": monatsIndex = 11
            Case "Dezember": monatsIndex = 12
            Case Else
                GoTo SafeExit
        End Select
        
        ' Monat in Hilfszelle AE4 schreiben (DATA_COL_HILFSZELLE_FILTER = 31 = Spalte AE)
        ' 0 = ganzes Jahr, 1-12 = jeweiliger Monat
        wsDaten.Range("AE4").value = monatsIndex
        
        ' Letzte Datenzeile ermitteln
        letzteDatenZeile = .Cells(.Rows.count, "A").End(xlUp).Row
        
        If letzteDatenZeile < ersteDatenZeile Then
            ' Keine Daten vorhanden - still zuruecksetzen
            Me.cbo_RechnungsMonat.ListIndex = 0
            wsDaten.Range("AE4").value = 0
            GoTo SafeExit
        End If
        
        Set rngFilterBereich = .Range(.Cells(ersteZeileFilterbereich, 1), _
                                      .Cells(letzteDatenZeile, letzteSpalteFilterbereich))
        
        ' Basis-Filter auf dem kompletten Kopfbereich immer aktiv halten.
        rngFilterBereich.AutoFilter
        
        .Range("C27").value = "Auszug: " & Me.cbo_RechnungsMonat.value & " " & jahr
        
        ' Datumsgrenzen nur wenn nicht "ganzes Jahr"
        If monatsIndex > 0 Then
            Dim erstesDesMonats As Date
            Dim letztesDesMonats As Date
            
            erstesDesMonats = DateSerial(jahr, monatsIndex, 1)
            letztesDesMonats = DateSerial(jahr, monatsIndex + 1, 0)
            
            rngFilterBereich.AutoFilter Field:=1, _
                Criteria1:=">=" & CLng(erstesDesMonats), _
                Operator:=xlAnd, _
                Criteria2:="<=" & CLng(letztesDesMonats)
        End If
        
        ' Sichtbare Daten Pruefen
        On Error Resume Next
        visibleCellsCount = rngFilterBereich.SpecialCells(xlCellTypeVisible).count
        On Error GoTo 0
        
        If visibleCellsCount <= 1 And monatsIndex > 0 Then
            ' Keine sichtbaren Daten - still zuruecksetzen
            .ShowAllData
            Me.cbo_RechnungsMonat.ListIndex = 0
            wsDaten.Range("AE4").value = 0
        End If
        
    End With

SafeExit:
    ' Blattschutz wiederherstellen (inkl. Sortier-/Filterrechte)
    On Error Resume Next
    mod_Banking_Format.Schuetze_BankkontoBlatt wsBank
    On Error GoTo 0
    
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "Fehler in cbo_RechnungsMonat_Change: " & Err.Description, vbCritical
        Err.Clear
    End If

End Sub

' ===============================================================
' NEU v9.7: Worksheet_SelectionChange
' Speichert den aktuellen Wert der ausgewaehlten Zelle in
' m_AlterWertH (Spalte H) bzw. m_AlterWertI (Spalte I),
' damit Worksheet_Change Pruefen kann ob sich der Wert
' tatsaechlich geaendert hat.
' ===============================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Nur einzelne Zelle merken
    If Target.Cells.CountLarge <> 1 Then Exit Sub
    If Target.Row < BK_START_ROW Then Exit Sub
    
    Select Case Target.Column
        Case BK_COL_KATEGORIE
            m_AlterWertH = Trim(CStr(Target.value))
        Case BK_COL_MONAT_PERIODE
            m_AlterWertI = Trim(CStr(Target.value))
    End Select
End Sub

' ===============================================================
' Worksheet_Change: Manuelle Auswahl in Spalte H oder Spalte I
' ---------------------------------------------------------------
' v9.7: Reagiert auf ZWEI Spalten:
'   Spalte H (BK_COL_KATEGORIE = 8): Kategorie-Auswahl
'   Spalte I (BK_COL_MONAT_PERIODE = 9): Monat-Auswahl (DropDown)
'
' NEU v9.7: prueft ob sich der Wert tatsaechlich geaendert hat.
'   Verhindert unnoetige Verarbeitung bei Doppelklick+Enter
'   ohne echte WertAenderung. Alter Wert wird in
'   Worksheet_SelectionChange gespeichert.
'
' Spalte H - Kategorie:
'   - Sammelzahlung: prueft ob bereits zugeordnet (gruen), fragt
'     per MsgBox ob neu zugeordnet werden soll.
'   - Normale Kategorie: Spalten M-Z leeren, gruen setzen,
'     "manuell zugeordnet" als Bemerkung, ApplyBetragsZuordnung
'   - ErmittleMonatPeriode mit 5 Parametern (inkl. Lern-Check)
'   - "GELB|Monatsname" Rueckgabe wird erkannt und verarbeitet
'
' Spalte I - Monat/Periode:
'   - v9.6 FIX: Lern-Vermerk NUR wenn Zelle vorher GELB war
'     (= Automatik war unsicher). Programmatische Aenderungen
'     durch SetzeMonatPeriode triggern KEINEN Vermerk mehr.
'   - Spalte I: hell-gruener Hintergrund (manuell bestaetigt)
'   - Spalte L: Lern-Vermerk anhaengen (mit vbLf)
' ===============================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    
    ' Nur einzelne Zelle, ab Datenzeile 28
    If Target.Cells.CountLarge <> 1 Then Exit Sub
    If Target.Row < BK_START_ROW Then Exit Sub
    
    ' =============================================
    ' WEICHE: Spalte H oder Spalte I?
    ' =============================================
    Select Case Target.Column
    
        Case BK_COL_KATEGORIE
            ' --- Spalte H: Kategorie-Auswahl ---
            ' v9.7: Pruefen ob sich der Wert tatsaechlich geaendert hat
            Dim neuerWertH As String
            neuerWertH = Trim(CStr(Target.value))
            
            If neuerWertH = m_AlterWertH Then
                ' Sammelzahlung: Trotzdem verarbeiten (InputBox muss erscheinen)
                If Not (LCase(neuerWertH) Like "*sammelzahlung*") Then
                    Exit Sub
                End If
            End If
            
            Call VerarbeiteKategorieAenderung(Target)
        
        Case BK_COL_MONAT_PERIODE
            ' --- Spalte I: Monat/Periode manuell gewaehlt ---
            ' v9.7: Pruefen ob sich der Wert tatsaechlich geaendert hat
            Dim neuerWertI As String
            neuerWertI = Trim(CStr(Target.value))
            
            If neuerWertI = m_AlterWertI Then
                ' Keine Aenderung -> nichts tun
                Exit Sub
            End If
            
            Call VerarbeiteMonatAenderung(Target)
        
        Case Else
            ' Andere Spalten: nichts tun
            Exit Sub
    
    End Select
    
    ' v3.0: Uebersicht nach manueller Aenderung in H/I automatisch aktualisieren
    ' v5.1: Kein hardcoded Year(Date) mehr - Jahr wird aus Startmenue!F1 / Bankkonto ermittelt
    On Error Resume Next
    Call mod_Uebersicht_Generator.GeneriereUebersicht(stummModus:=True)
    Call mod_Uebersicht_Dashboard.GeneriereUebersichtNeu(stummModus:=True)
    On Error GoTo 0
    
    ' Punkt 11: BK-/KA-Nummern bei Kategorie-Aenderung neu vergeben
    If Target.Column = BK_COL_KATEGORIE Then
        On Error Resume Next
        Call mod_BK_KA_Nummern.NeuberechneAlleBKNummern(Me)
        On Error GoTo 0
    End If
    
End Sub


' ===============================================================
' Spalte H: Kategorie-Aenderung verarbeiten
' (Bisherige Logik aus Worksheet_Change, jetzt ausgelagert)
' v9.5: ErmittleMonatPeriode mit 5 Parametern + GELB-Parsing
' ===============================================================
Private Sub VerarbeiteKategorieAenderung(ByVal Target As Range)
    
    Dim kategorie As String
    kategorie = Trim(Target.value)
    If kategorie = "" Then Exit Sub
    
    On Error GoTo ChangeError
    Application.EnableEvents = False
    
    Dim ws As Worksheet
    Set ws = Me
    
    Dim rowBK As Long
    rowBK = Target.Row
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ChangeError
    
    ' Sammelzahlung? -> Spezialbehandlung
    If LCase(kategorie) Like "*sammelzahlung*" Then
        
        ' Pruefen ob die Zeile bereits gruen ist (= bereits vollstaendig zugeordnet)
        Dim aktFarbe As Long
        aktFarbe = Target.Interior.color
        
        If aktFarbe = RGB(198, 239, 206) Then
            ' Bereits vollstaendig zugeordnet -> Nutzer fragen
            Dim antwort As VbMsgBoxResult
            antwort = MsgBox("Diese Sammelzahlung wurde bereits vollst" & ChrW(228) & "ndig zugeordnet." & vbLf & vbLf & _
                             "M" & ChrW(246) & "chten Sie die Zuordnung verwerfen und neu aufteilen?", _
                             vbQuestion + vbYesNo, "Sammelzahlung - Zeile " & rowBK)
            
            If antwort = vbNo Then
                ' Nutzer will NICHT neu zuordnen -> nichts tun, gruen bleibt
                GoTo ChangeCleanup
            End If
            ' Nutzer will NEU zuordnen -> weiter zu VerarbeiteSammelzahlung
        End If
        
        Call VerarbeiteSammelzahlung(ws, rowBK)
        GoTo ChangeCleanup
    End If
    
    ' --- NORMALE KATEGORIE (auch bei Aenderung von gruen) ---
    
    ' 1. Alte Betraege in Spalten M-Z Loeschen?
    Dim c As Long
    For c = BK_COL_EINNAHMEN_START To BK_COL_AUSGABEN_ENDE
        ws.Cells(rowBK, c).ClearContents
    Next c
    
    ' 2. Kategorie gruen faerben
    ApplyKategorie ws.Cells(rowBK, BK_COL_KATEGORIE), kategorie, "GRUEN"
    
    ' 3. Bemerkung setzen
    ws.Cells(rowBK, BK_COL_BEMERKUNG).value = "manuell zugeordnet"
    
    ' 4. Monat/Periode setzen (falls leer)
    If Trim(ws.Cells(rowBK, BK_COL_MONAT_PERIODE).value) = "" Then
        If IsDate(ws.Cells(rowBK, BK_COL_DATUM).value) Then
            Dim wsDaten As Worksheet
            Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
            Dim faelligkeit As String
            Dim lastRuleDaten As Long
            lastRuleDaten = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
            Dim rD As Long
            faelligkeit = "monatlich"
            For rD = DATA_START_ROW To lastRuleDaten
                If Trim(wsDaten.Cells(rD, DATA_CAT_COL_KATEGORIE).value) = kategorie Then
                    faelligkeit = LCase(Trim(wsDaten.Cells(rD, DATA_CAT_COL_FAELLIGKEIT).value))
                    If faelligkeit = "" Then faelligkeit = "monatlich"
                    Exit For
                End If
            Next rD
            
            Call LadeEinstellungenCache
            
            ' v9.5: ErmittleMonatPeriode mit 5 Parametern
            ' (inkl. wsBK + aktuelleZeile ffuer Lern-Mechanismus)
            Dim ergebnis As String
            ergebnis = ErmittleMonatPeriode(kategorie, CDate(ws.Cells(rowBK, BK_COL_DATUM).value), _
                                            faelligkeit, ws, rowBK)
            
            ' v9.5: "GELB|Monatsname" Rueckgabe verarbeiten
            If Left(ergebnis, 5) = "GELB|" Then
                ' GELB-Fall: Monat extrahieren und setzen
                Dim monatName As String
                monatName = Mid(ergebnis, 6) ' alles nach "GELB|"
                
                ws.Cells(rowBK, BK_COL_MONAT_PERIODE).value = monatName
                
                ' GELB-Hintergrund setzen (= Nutzer soll Pruefen)
                ws.Cells(rowBK, BK_COL_MONAT_PERIODE).Interior.color = RGB(255, 235, 156)
                
                ' Bemerkung: Hinweis anhaengen
                Dim bestehend As String
                bestehend = Trim(CStr(ws.Cells(rowBK, BK_COL_BEMERKUNG).value))
                
                Dim gelbHinweis As String
                gelbHinweis = "Bitte pr" & ChrW(252) & "fen ob Zahlung f" & ChrW(252) & "r " & _
                              monatName & " oder Folgemonat gilt"
                
                If bestehend = "" Then
                    ws.Cells(rowBK, BK_COL_BEMERKUNG).value = gelbHinweis
                Else
                    ws.Cells(rowBK, BK_COL_BEMERKUNG).value = bestehend & vbLf & gelbHinweis
                End If
                
                ' Hell-gelber Hintergrund fuer die Bemerkung (gleiche Farbe wie Spalte I)
                ws.Cells(rowBK, BK_COL_BEMERKUNG).Interior.color = RGB(255, 235, 156)
            Else
                ' Normaler Fall: Monat direkt setzen (kein GELB)
                ws.Cells(rowBK, BK_COL_MONAT_PERIODE).value = ergebnis
                ' Ampelfarbe Gruen = Monat eindeutig bestimmt
                ws.Cells(rowBK, BK_COL_MONAT_PERIODE).Interior.color = RGB(198, 239, 206)
            End If
            
            Call EntladeEinstellungenCache
        End If
    End If
    
    ' 5. Betragszuordnung ausfuehren
    ApplyBetragsZuordnung ws, rowBK

ChangeCleanup:
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    Application.EnableEvents = True
    Exit Sub

ChangeError:
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "Fehler in Worksheet_Change (Kategorie): " & Err.Description, vbExclamation
        Err.Clear
    End If
End Sub


' ===============================================================
' Spalte I: Monat/Periode manuell geaendert (DropDown)
' v9.8: JEDE manuelle Aenderung wird in Spalte L notiert.
' ---------------------------------------------------------------
' - Bei GELBER Zelle: Folgemonat-Lernvermerk (Marker ffuer Lern-Check)
' - Bei gruenER Zelle mit bestehendem Folgemonat-Vermerk: Aktualisierung
' - Bei ALLEN anderen Fuellen: Allgemeiner Vermerk "manuell geaendert"
' ===============================================================
Private Sub VerarbeiteMonatAenderung(ByVal Target As Range)
    
    Dim monatWert As String
    monatWert = Trim(Target.value)
    If monatWert = "" Then Exit Sub
    
    Dim vorherigeFarbe As Long
    vorherigeFarbe = Target.Interior.color
    
    ' GELB = RGB(255, 235, 156) = Automatik war unsicher
    Dim istVorherGelb As Boolean
    istVorherGelb = (vorherigeFarbe = RGB(255, 235, 156))
    
    On Error GoTo MonatError
    Application.EnableEvents = False
    
    Dim ws As Worksheet
    Set ws = Me
    
    Dim rowBK As Long
    rowBK = Target.Row
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo MonatError
    
    ' 1. Hell-gruenen Hintergrund setzen (= manuell bestaetigt)
    ws.Cells(rowBK, BK_COL_MONAT_PERIODE).Interior.color = RGB(198, 239, 206)
    
    ' 2. Bemerkung in Spalte L aktualisieren
    Dim bestehendeBemerkung As String
    bestehendeBemerkung = Trim(CStr(ws.Cells(rowBK, BK_COL_BEMERKUNG).value))
    
    If istVorherGelb Then
        ' =============================================
        ' GELB-Fall: Folgemonat-Lernvermerk schreiben
        ' Dieser Marker wird vom Lern-Check in
        ' ErmittleMonatPeriode ausgewertet.
        ' =============================================
        Dim lernVermerk As String
        lernVermerk = "Folgemonat manuell best" & ChrW(228) & "tigt: " & monatWert
        
        ' Alten Folgemonat-Vermerk ersetzen (falls vorhanden)
        bestehendeBemerkung = EntferneBemerkungsZeile(bestehendeBemerkung, "folgemonat manuell best")
        ' Alten "manuell geaendert"-Vermerk ersetzen (falls vorhanden)
        bestehendeBemerkung = EntferneBemerkungsZeile(bestehendeBemerkung, "monat/periode manuell ge")
        
        If bestehendeBemerkung = "" Then
            ws.Cells(rowBK, BK_COL_BEMERKUNG).value = lernVermerk
        Else
            ws.Cells(rowBK, BK_COL_BEMERKUNG).value = bestehendeBemerkung & vbLf & lernVermerk
        End If
    Else
        ' =============================================
        ' Alle anderen Fuelle: Allgemeiner Vermerk
        ' Folgemonat-Vermerk NUR aktualisieren wenn vorhanden
        ' =============================================
        Dim hatFolgemonatVermerk As Boolean
        hatFolgemonatVermerk = (InStr(LCase(bestehendeBemerkung), "folgemonat manuell best") > 0)
        
        If hatFolgemonatVermerk Then
            ' Bestehenden Folgemonat-Vermerk aktualisieren
            Dim neuerLern As String
            neuerLern = "Folgemonat manuell best" & ChrW(228) & "tigt: " & monatWert
            bestehendeBemerkung = EntferneBemerkungsZeile(bestehendeBemerkung, "folgemonat manuell best")
            bestehendeBemerkung = EntferneBemerkungsZeile(bestehendeBemerkung, "monat/periode manuell ge")
            
            If bestehendeBemerkung = "" Then
                ws.Cells(rowBK, BK_COL_BEMERKUNG).value = neuerLern
            Else
                ws.Cells(rowBK, BK_COL_BEMERKUNG).value = bestehendeBemerkung & vbLf & neuerLern
            End If
        Else
            ' Allgemeinen Vermerk schreiben
            Dim manuellVermerk As String
            manuellVermerk = "Monat/Periode manuell ge" & ChrW(228) & "ndert auf: " & monatWert
            
            ' Alten "manuell geaendert"-Vermerk ersetzen
            bestehendeBemerkung = EntferneBemerkungsZeile(bestehendeBemerkung, "monat/periode manuell ge")
            
            If bestehendeBemerkung = "" Then
                ws.Cells(rowBK, BK_COL_BEMERKUNG).value = manuellVermerk
            Else
                ws.Cells(rowBK, BK_COL_BEMERKUNG).value = bestehendeBemerkung & vbLf & manuellVermerk
            End If
        End If
    End If
    
    ' 3. Auch den evtl. vorhandenen GELB-Hinweis entfernen
    '    (der ist jetzt nicht mehr relevant, da der Nutzer manuell bestaetigt hat)
    Dim aktBemerkung As String
    aktBemerkung = CStr(ws.Cells(rowBK, BK_COL_BEMERKUNG).value)
    
    If InStr(LCase(aktBemerkung), "bitte pr" & ChrW(252) & "fen ob zahlung") > 0 Then
        ws.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            EntferneBemerkungsZeile(aktBemerkung, "bitte pr" & ChrW(252) & "fen ob zahlung")
        ' Hell-gelben Hintergrund der Bemerkung zuruecksetzen
        ws.Cells(rowBK, BK_COL_BEMERKUNG).Interior.ColorIndex = xlNone
    End If

MonatCleanup:
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    Application.EnableEvents = True
    Exit Sub

MonatError:
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "Fehler in Worksheet_Change (Monat): " & Err.Description, vbExclamation
        Err.Clear
    End If
End Sub


' ===============================================================
' Sammelzahlung verarbeiten (v9.5)
' ---------------------------------------------------------------
' Komfortable InputBox mit:
' - Name des Zahlers und Verwendungszweck als Kontext
' - Buchstaben-Kuerzel (A, B, C...) ffuer KATEGORIEN (nicht Spalten!)
'   Die Kategorien werden dynamisch aus der Kategorie-Tabelle
'   (Spalte J auf Daten!) gelesen, nach E/A getrennt, nicht redundant.
' - Eingabe: Kuerzel;Betrag (z.B. "A;25,00")
'   Semikolon trennt Kuerzel vom Betrag, Komma = Dezimalzeichen
' - Jede Kategorie hat ueber Spalte N (Zielspalte) eine Zuordnung
'   wohin der Betrag auf dem Bankkonto gebucht wird.
' - Bei mehrfacher Buchung in die gleiche Spalte werden Betraege addiert.
' - gruen nur wenn ALLES zugeordnet, GELB bei teilweise/nichts.
' - Die Schleife endet sofort wenn der Betrag vollstaendig zugeordnet
'   ist - es erscheint KEINE weitere InputBox mit "0,00 offen".
' ===============================================================
Private Sub VerarbeiteSammelzahlung(ByVal ws As Worksheet, ByVal rowBK As Long)
    
    Dim betrag As Double
    betrag = ws.Cells(rowBK, BK_COL_BETRAG).value
    Dim absBetrag As Double
    absBetrag = Abs(betrag)
    
    ' Rundung auf 2 Nachkommastellen (verhindert Fliesskomma-Probleme)
    absBetrag = Round(absBetrag, 2)
    
    Dim istEinnahme As Boolean
    istEinnahme = (betrag > 0)
    
    ' Kontext-Info aus der Buchungszeile lesen
    Dim zahlerName As String
    zahlerName = Trim(ws.Cells(rowBK, BK_COL_NAME).value)
    If zahlerName = "" Then zahlerName = "(unbekannt)"
    
    Dim verwendungszweck As String
    verwendungszweck = Trim(ws.Cells(rowBK, BK_COL_VERWENDUNGSZWECK).value)
    If verwendungszweck = "" Then verwendungszweck = "(kein Verwendungszweck)"
    
    ' Spalten M-Z komplett leeren
    Dim c As Long
    For c = BK_COL_EINNAHMEN_START To BK_COL_AUSGABEN_ENDE
        ws.Cells(rowBK, c).ClearContents
    Next c
    
    ' ---------------------------------------------------------------
    ' Kategorien dynamisch aus der Kategorie-Tabelle laden
    ' (Spalte J auf Daten!, nach E/A getrennt, nicht redundant)
    ' ---------------------------------------------------------------
    Dim wsDaten As Worksheet
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    
    Dim lastRuleRow As Long
    lastRuleRow = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    
    ' Eindeutige Kategorien sammeln + deren Zielspalte (Spalte N)
    Dim dictKat As Object
    Set dictKat = CreateObject("Scripting.Dictionary")
    
    Dim filterEA As String
    If istEinnahme Then filterEA = "E" Else filterEA = "A"
    
    Dim r As Long
    For r = DATA_START_ROW To lastRuleRow
        Dim katName As String
        Dim katEA As String
        Dim katZiel As String
        
        katName = Trim(CStr(wsDaten.Cells(r, DATA_CAT_COL_KATEGORIE).value))
        katEA = UCase(Trim(CStr(wsDaten.Cells(r, DATA_CAT_COL_EINAUS).value)))
        katZiel = Trim(CStr(wsDaten.Cells(r, DATA_CAT_COL_ZIELSPALTE).value))
        
        If katName = "" Then GoTo NextKatRow
        If katEA <> filterEA Then GoTo NextKatRow
        
        ' Sammelzahlung selbst nicht als Unterkategorie anbieten
        If LCase(katName) Like "*sammelzahlung*" Then GoTo NextKatRow
        
        ' Nur aufnehmen wenn noch nicht vorhanden
        If Not dictKat.exists(katName) Then
            dictKat.Add katName, katZiel
        End If
NextKatRow:
    Next r
    
    Dim anzahlKat As Long
    anzahlKat = dictKat.count
    
    If anzahlKat = 0 Then
        MsgBox "Keine Kategorien f" & ChrW(252) & "r " & IIf(istEinnahme, "Einnahmen", "Ausgaben") & _
               " in der Kategorie-Tabelle gefunden!", vbCritical
        Exit Sub
    End If
    
    ' Kuerzel-Arrays aufbauen: A=erste Kategorie, B=zweite, usw.
    Dim kuerzelBuchstabe() As String
    Dim kuerzelKatName() As String
    Dim kuerzelZielHeader() As String
    ReDim kuerzelBuchstabe(1 To anzahlKat)
    ReDim kuerzelKatName(1 To anzahlKat)
    ReDim kuerzelZielHeader(1 To anzahlKat)
    
    Dim idx As Long
    idx = 0
    Dim k As Variant
    For Each k In dictKat.keys
        idx = idx + 1
        kuerzelBuchstabe(idx) = Chr(64 + idx)   ' A, B, C, ...
        kuerzelKatName(idx) = CStr(k)
        kuerzelZielHeader(idx) = CStr(dictKat(k))
    Next k
    
    ' Kuerzel-Liste ffuer Anzeige zusammenbauen
    Dim spaltenInfo As String
    spaltenInfo = IIf(istEinnahme, "EINNAHMEN", "AUSGABEN") & "-Kategorien:" & vbLf
    Dim i As Long
    For i = 1 To anzahlKat
        spaltenInfo = spaltenInfo & "  " & kuerzelBuchstabe(i) & ": " & _
                      kuerzelKatName(i) & vbLf
    Next i
    
    ' Eingabeschleife
    Dim restBetrag As Double
    restBetrag = absBetrag
    Dim zugeordnet As Double
    zugeordnet = 0
    
    ' Spaltenbereich ffuer Einnahmen/Ausgaben
    Dim startCol As Long
    Dim endCol As Long
    If istEinnahme Then
        startCol = BK_COL_EINNAHMEN_START
        endCol = BK_COL_EINNAHMEN_ENDE
    Else
        startCol = BK_COL_AUSGABEN_START
        endCol = BK_COL_AUSGABEN_ENDE
    End If
    
    ' Sammelzahlungs-Dokumentation ffuer Spalte L aufbauen
    Dim sammelDoku As String
    sammelDoku = ""
    
    Do
        ' *** Pruefung VOR der InputBox: Ist der Betrag bereits vollstaendig zugeordnet? ***
        restBetrag = Round(absBetrag - zugeordnet, 2)
        
        If restBetrag <= 0.01 Then
            ' Betrag ist vollstaendig zugeordnet -> Schleife sofort beenden
            ' KEINE weitere InputBox anzeigen!
            Exit Do
        End If
        
        Dim prompt As String
        prompt = "=== BUCHUNGSDETAILS ===" & vbLf & _
                 "Name: " & zahlerName & vbLf & _
                 "Verwendungszweck: " & verwendungszweck & vbLf & _
                 "Gesamtbetrag: " & Format(absBetrag, "#,##0.00") & " " & ChrW(8364) & vbLf & vbLf & _
                 spaltenInfo & vbLf & _
                 "Bereits zugeordnet: " & Format(zugeordnet, "#,##0.00") & " " & ChrW(8364) & vbLf & _
                 "Noch offen: " & Format(restBetrag, "#,##0.00") & " " & ChrW(8364) & vbLf & vbLf & _
                 "Eingabe: K" & ChrW(252) & "rzel;Betrag  (z.B. A;25,00)" & vbLf & _
                 "Abbrechen = Rest nicht zuordnen"
        
        Dim eingabe As String
        eingabe = InputBox(prompt, "Sammelzahlung - Zeile " & rowBK)
        
        ' Abbrechen oder leer
        If eingabe = "" Then Exit Do
        
        ' Leerzeichen entfernen
        eingabe = Replace(eingabe, " ", "")
        
        ' Eingabe parsen: "Kuerzel;Betrag" - Semikolon ist IMMER der Trenner
        ' Komma ist Dezimalzeichen (Euro,Cent)
        Dim trennPos As Long
        trennPos = InStr(eingabe, ";")
        
        If trennPos = 0 Or trennPos >= Len(eingabe) Then
            MsgBox "Ung" & ChrW(252) & "ltiges Format!" & vbLf & _
                   "Bitte im Format 'A;25,00' eingeben." & vbLf & _
                   "Semikolon trennt K" & ChrW(252) & "rzel vom Betrag.", vbExclamation
            GoTo WeiterEingabe
        End If
        
        Dim eingabeKuerzel As String
        eingabeKuerzel = UCase(Left(eingabe, trennPos - 1))
        
        Dim teilBetragStr As String
        teilBetragStr = Mid(eingabe, trennPos + 1)
        ' Komma -> Punkt ffuer Val() (Komma ist Dezimalzeichen)
        teilBetragStr = Replace(teilBetragStr, ",", ".")
        
        Dim teilBetrag As Double
        teilBetrag = val(teilBetragStr)
        teilBetrag = Round(teilBetrag, 2)
        
        If teilBetrag <= 0 Then
            MsgBox "Ung" & ChrW(252) & "ltiger Betrag: " & Mid(eingabe, trennPos + 1), vbExclamation
            GoTo WeiterEingabe
        End If
        
        If teilBetrag > restBetrag + 0.01 Then
            MsgBox "Betrag " & Format(teilBetrag, "#,##0.00") & " " & ChrW(8364) & _
                   " ist gr" & ChrW(246) & ChrW(223) & "er als der Restbetrag " & _
                   Format(restBetrag, "#,##0.00") & " " & ChrW(8364) & "!", vbExclamation
            GoTo WeiterEingabe
        End If
        
        ' Kuerzel in Kategorie uebersetzen
        Dim gefundenIdx As Long
        gefundenIdx = 0
        For i = 1 To anzahlKat
            If kuerzelBuchstabe(i) = eingabeKuerzel Then
                gefundenIdx = i
                Exit For
            End If
        Next i
        
        If gefundenIdx = 0 Then
            MsgBox "K" & ChrW(252) & "rzel '" & eingabeKuerzel & "' nicht gefunden!" & vbLf & _
                   "Erlaubte K" & ChrW(252) & "rzel: A bis " & kuerzelBuchstabe(anzahlKat), vbExclamation
            GoTo WeiterEingabe
        End If
        
        ' Zielspalte ueber die Kategorietabelle bestimmen (Spalte N = Zielspalte)
        Dim zielHeader As String
        zielHeader = kuerzelZielHeader(gefundenIdx)
        
        If zielHeader = "" Then
            MsgBox "Kategorie '" & kuerzelKatName(gefundenIdx) & "' hat keine Zielspalte " & _
                   "(Spalte N in der Kategorie-Tabelle ist leer)!", vbExclamation
            GoTo WeiterEingabe
        End If
        
        ' Zielspalte im richtigen Bereich (M-S oder T-Z) finden
        Dim zielCol As Long
        zielCol = 0
        For c = startCol To endCol
            If StrComp(Trim(ws.Cells(BK_HEADER_ROW, c).value), zielHeader, vbTextCompare) = 0 Then
                zielCol = c
                Exit For
            End If
        Next c
        
        If zielCol = 0 Then
            MsgBox "Zielspalte '" & zielHeader & "' nicht im " & _
                   IIf(istEinnahme, "Einnahmen", "Ausgaben") & "-Bereich gefunden!", vbExclamation
            GoTo WeiterEingabe
        End If
        
        ' Betrag eintragen (ADDIEREN wenn schon ein Wert in der Zelle steht)
        Dim vorhandenerBetrag As Double
        vorhandenerBetrag = 0
        If ws.Cells(rowBK, zielCol).value <> "" Then
            vorhandenerBetrag = CDbl(ws.Cells(rowBK, zielCol).value)
        End If
        
        If istEinnahme Then
            ws.Cells(rowBK, zielCol).value = Round(vorhandenerBetrag + teilBetrag, 2)
        Else
            ws.Cells(rowBK, zielCol).value = Round(vorhandenerBetrag + (-teilBetrag), 2)
        End If
        
        zugeordnet = Round(zugeordnet + teilBetrag, 2)
        
        ' Dokumentation ffuer Spalte L aufbauen
        If sammelDoku <> "" Then sammelDoku = sammelDoku & vbLf
        sammelDoku = sammelDoku & kuerzelKatName(gefundenIdx) & ": " & _
                     Format(teilBetrag, "#,##0.00") & " " & ChrW(8364)
        
WeiterEingabe:
    Loop
    
    ' ===============================================================
    ' Farbe und Bemerkung abhaengig vom Zuordnungsstatus
    ' ===============================================================
    
    ' Endgueltigen Restbetrag nochmal sauber berechnen
    restBetrag = Round(absBetrag - zugeordnet, 2)
    
    If zugeordnet < 0.01 Then
        ' NICHTS zugeordnet (User hat sofort Abbrechen geklickt)
        ' -> GELB bleibt, Bemerkung bleibt wie von der Engine gesetzt
        Exit Sub
    End If
    
    If restBetrag <= 0.01 Then
        ' ALLES zugeordnet -> gruen (sicher)
        ApplyKategorie ws.Cells(rowBK, BK_COL_KATEGORIE), _
                       ws.Cells(rowBK, BK_COL_KATEGORIE).value, "GRUEN"
        ws.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            "manuell zugeordnet (Sammelzahlung vollst" & ChrW(228) & "ndig aufgeteilt)" & _
            IIf(sammelDoku <> "", vbLf & "Aufschl" & ChrW(252) & "sselung:" & vbLf & sammelDoku, "")
    Else
        ' TEILWEISE zugeordnet -> GELB bleibt
        ApplyKategorie ws.Cells(rowBK, BK_COL_KATEGORIE), _
                       ws.Cells(rowBK, BK_COL_KATEGORIE).value, "GELB"
        ws.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            "Sammelzahlung teilweise aufgeteilt, " & _
            Format(restBetrag, "#,##0.00") & " " & ChrW(8364) & " nicht zugeordnet" & _
            IIf(sammelDoku <> "", vbLf & "Aufschl" & ChrW(252) & "sselung:" & vbLf & sammelDoku, "")
    End If
    
End Sub


' ===============================================================
' Hilfsfunktion: Entfernt alle Zeilen aus einer vbLf-getrennten
' Bemerkung, die den Suchtext (lowercase) enthalten.
' Gibt den bereinigten Text zurueck.
' ===============================================================
Private Function EntferneBemerkungsZeile(ByVal bemerkung As String, _
                                          ByVal suchTextLower As String) As String
    If bemerkung = "" Then
        EntferneBemerkungsZeile = ""
        Exit Function
    End If
    
    If InStr(LCase(bemerkung), suchTextLower) = 0 Then
        EntferneBemerkungsZeile = bemerkung
        Exit Function
    End If
    
    Dim zeilen() As String
    zeilen = Split(bemerkung, vbLf)
    
    Dim ergebnis As String
    ergebnis = ""
    
    Dim zz As Long
    For zz = LBound(zeilen) To UBound(zeilen)
        If InStr(LCase(zeilen(zz)), suchTextLower) = 0 Then
            If ergebnis = "" Then
                ergebnis = zeilen(zz)
            Else
                ergebnis = ergebnis & vbLf & zeilen(zz)
            End If
        End If
    Next zz
    
    EntferneBemerkungsZeile = ergebnis
End Function











































































































































