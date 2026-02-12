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

Private Sub btn_CSVimport_Click()
    ' Ruft die Hauptprozedur im Standardmodul mod_Datenpflege auf
    Call Importiere_Kontoauszug
End Sub

Private Sub Worksheet_Activate()

    ' Events ABSCHALTEN bevor die ComboBox befüllt wird!
    ' Sonst löst .ListIndex = 0 sofort cbo_RechnungsMonat_Change aus,
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

    Call Initialize_ImportReport_ListBox
End Sub

Private Sub cbo_RechnungsMonat_Change()

    Dim wsStart As Worksheet
    Dim wsBank As Worksheet
    Dim wsDaten As Worksheet
    
    Dim jahr As Long
    Dim monatsIndex As Long
    Dim letzteDatenZeile As Long
    
    Const ersteDatenZeile As Long = 28
    Const ersteZeileFilterbereich As Long = 27
    
    Dim rngFilterBereich As Range
    Dim visibleCellsCount As Long
    
    Set wsStart = ThisWorkbook.Sheets("Startmen" & ChrW(252))
    Set wsBank = ThisWorkbook.Sheets("Bankkonto")
    Set wsDaten = ThisWorkbook.Sheets("Daten")

    On Error GoTo SafeExit
    Application.EnableEvents = False

    With wsBank
    
        ' Vorherige Filter immer aufheben
        On Error Resume Next
        .ShowAllData
        On Error GoTo SafeExit
        
        If Me.cbo_RechnungsMonat.ListIndex = -1 Then GoTo SafeExit
        
        ' Abrechnungsjahr lesen
        If IsNumeric(wsStart.Range("F1").value) Then
            jahr = CLng(wsStart.Range("F1").value)
        Else
            MsgBox "Fehler: Ung" & ChrW(252) & "ltiges Abrechnungsjahr in Startmen" & ChrW(252) & "!F1", vbCritical
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
            ' Keine Daten vorhanden - still zur" & ChrW(252) & "cksetzen
            Me.cbo_RechnungsMonat.ListIndex = 0
            wsDaten.Range("AE4").value = 0
            GoTo SafeExit
        End If
        
        Set rngFilterBereich = .Range("A" & ersteZeileFilterbereich & ":A" & letzteDatenZeile)
        
        .Range("C25").value = "Auszug: " & Me.cbo_RechnungsMonat.value & " " & jahr
        
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
        
        ' Sichtbare Daten pr" & ChrW(252) & "fen
        On Error Resume Next
        visibleCellsCount = rngFilterBereich.SpecialCells(xlCellTypeVisible).count
        On Error GoTo 0
        
        If visibleCellsCount <= 1 And monatsIndex > 0 Then
            ' Keine sichtbaren Daten - still zur" & ChrW(252) & "cksetzen
            .ShowAllData
            Me.cbo_RechnungsMonat.ListIndex = 0
            wsDaten.Range("AE4").value = 0
        End If
        
    End With

SafeExit:
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "Fehler in cbo_RechnungsMonat_Change: " & Err.Description, vbCritical
        Err.Clear
    End If

End Sub

' ===============================================================
' Worksheet_Change: Manuelle Auswahl in Spalte H oder Spalte I
' ---------------------------------------------------------------
' v9.6: Reagiert auf ZWEI Spalten:
'   Spalte H (BK_COL_KATEGORIE = 8): Kategorie-Auswahl
'   Spalte I (BK_COL_MONAT_PERIODE = 9): Monat-Auswahl (DropDown)
'
' Spalte H - Kategorie:
'   - Sammelzahlung: Pr" & ChrW(252) & "ft ob bereits zugeordnet (GR" & ChrW(220) & "N), fragt
'     per MsgBox ob neu zugeordnet werden soll.
'   - Normale Kategorie: Spalten M-Z leeren, GR" & ChrW(220) & "N setzen,
'     "manuell zugeordnet" als Bemerkung, ApplyBetragsZuordnung
'   - ErmittleMonatPeriode mit 5 Parametern (inkl. Lern-Check)
'   - "GELB|Monatsname" R" & ChrW(252) & "ckgabe wird erkannt und verarbeitet
'
' Spalte I - Monat/Periode:
'   - v9.6 FIX: Lern-Vermerk NUR wenn Zelle vorher GELB war
'     (= Automatik war unsicher). Programmatische Änderungen
'     durch SetzeMonatPeriode triggern KEINEN Vermerk mehr.
'   - Spalte I: hell-gr" & ChrW(252) & "ner Hintergrund (manuell best" & ChrW(228) & "tigt)
'   - Spalte L: Lern-Vermerk anh" & ChrW(228) & "ngen (mit vbLf)
' ===============================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    
    ' Nur einzelne Zelle, ab Datenzeile 28
    If Target.Cells.count <> 1 Then Exit Sub
    If Target.Row < BK_START_ROW Then Exit Sub
    
    ' =============================================
    ' WEICHE: Spalte H oder Spalte I?
    ' =============================================
    Select Case Target.Column
    
        Case BK_COL_KATEGORIE
            ' --- Spalte H: Kategorie-Auswahl ---
            Call VerarbeiteKategorieAenderung(Target)
        
        Case BK_COL_MONAT_PERIODE
            ' --- Spalte I: Monat/Periode manuell gew" & ChrW(228) & "hlt ---
            Call VerarbeiteMonatAenderung(Target)
        
        Case Else
            ' Andere Spalten: nichts tun
            Exit Sub
    
    End Select
    
End Sub


' ===============================================================
' Spalte H: Kategorie-" & ChrW(196) & "nderung verarbeiten
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
        
        ' Pr" & ChrW(252) & "fen ob die Zeile bereits GR" & ChrW(220) & "N ist (= bereits vollst" & ChrW(228) & "ndig zugeordnet)
        Dim aktFarbe As Long
        aktFarbe = Target.Interior.color
        
        If aktFarbe = RGB(198, 239, 206) Then
            ' Bereits vollst" & ChrW(228) & "ndig zugeordnet -> Nutzer fragen
            Dim antwort As VbMsgBoxResult
            antwort = MsgBox("Diese Sammelzahlung wurde bereits vollst" & ChrW(228) & "ndig zugeordnet." & vbLf & vbLf & _
                             "M" & ChrW(246) & "chten Sie die Zuordnung verwerfen und neu aufteilen?", _
                             vbQuestion + vbYesNo, "Sammelzahlung - Zeile " & rowBK)
            
            If antwort = vbNo Then
                ' Nutzer will NICHT neu zuordnen -> nichts tun, GR" & ChrW(220) & "N bleibt
                GoTo ChangeCleanup
            End If
            ' Nutzer will NEU zuordnen -> weiter zu VerarbeiteSammelzahlung
        End If
        
        Call VerarbeiteSammelzahlung(ws, rowBK)
        GoTo ChangeCleanup
    End If
    
    ' --- NORMALE KATEGORIE (auch bei " & ChrW(196) & "nderung von GR" & ChrW(220) & "N) ---
    
    ' 1. Alte Betr" & ChrW(228) & "ge in Spalten M-Z l" & ChrW(246) & "schen
    Dim c As Long
    For c = BK_COL_EINNAHMEN_START To BK_COL_AUSGABEN_ENDE
        ws.Cells(rowBK, c).ClearContents
    Next c
    
    ' 2. Kategorie GR" & ChrW(220) & "N f" & ChrW(228) & "rben
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
            ' (inkl. wsBK + aktuelleZeile f" & ChrW(252) & "r Lern-Mechanismus)
            Dim ergebnis As String
            ergebnis = ErmittleMonatPeriode(kategorie, CDate(ws.Cells(rowBK, BK_COL_DATUM).value), _
                                            faelligkeit, ws, rowBK)
            
            ' v9.5: "GELB|Monatsname" R" & ChrW(252) & "ckgabe verarbeiten
            If Left(ergebnis, 5) = "GELB|" Then
                ' GELB-Fall: Monat extrahieren und setzen
                Dim monatName As String
                monatName = Mid(ergebnis, 6) ' alles nach "GELB|"
                
                ws.Cells(rowBK, BK_COL_MONAT_PERIODE).value = monatName
                
                ' GELB-Hintergrund setzen (= Nutzer soll pr" & ChrW(252) & "fen)
                ws.Cells(rowBK, BK_COL_MONAT_PERIODE).Interior.color = RGB(255, 235, 156)
                
                ' Bemerkung: Hinweis anh" & ChrW(228) & "ngen
                Dim bestehend As String
                bestehend = Trim(CStr(ws.Cells(rowBK, BK_COL_BEMERKUNG).value))
                
                Dim gelbHinweis As String
                gelbHinweis = "Ultimo-5: Bitte pr" & ChrW(252) & "fen ob Zahlung f" & ChrW(252) & "r " & _
                              monatName & " oder Folgemonat gilt"
                
                If bestehend = "" Then
                    ws.Cells(rowBK, BK_COL_BEMERKUNG).value = gelbHinweis
                Else
                    ws.Cells(rowBK, BK_COL_BEMERKUNG).value = bestehend & vbLf & gelbHinweis
                End If
            Else
                ' Normaler Fall: Monat direkt setzen (kein GELB)
                ws.Cells(rowBK, BK_COL_MONAT_PERIODE).value = ergebnis
            End If
            
            Call EntladeEinstellungenCache
        End If
    End If
    
    ' 5. Betragszuordnung ausf" & ChrW(252) & "hren
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
' Spalte I: Monat/Periode manuell ge" & ChrW(228) & "ndert (DropDown)
' v9.6: FIX - Lern-Vermerk NUR wenn Zelle vorher GELB war
' ---------------------------------------------------------------
' WICHTIG: Der Lern-Vermerk wird NUR geschrieben wenn der
' Hintergrund der Zelle vorher GELB war (RGB 255,235,156).
' Das bedeutet: Die Automatik (ErmittleMonatPeriode) war
' unsicher und hat die Zelle GELB markiert. Wenn der Nutzer
' dann manuell einen Monat w" & ChrW(228) & "hlt, ist DAS eine echte
' manuelle Best" & ChrW(228) & "tigung die dokumentiert werden muss.
'
' Wenn die Zelle NICHT gelb war (z.B. programmatisch gesetzt
' durch SetzeMonatPeriode oder bereits gr" & ChrW(252) & "n), wird KEIN
' Lern-Vermerk geschrieben.
' ===============================================================
Private Sub VerarbeiteMonatAenderung(ByVal Target As Range)
    
    Dim monatWert As String
    monatWert = Trim(Target.value)
    If monatWert = "" Then Exit Sub
    
    ' =============================================
    ' v9.6 FIX: Pr" & ChrW(252) & "fe ob die Zelle vorher GELB war
    ' Nur bei GELB handelt es sich um eine echte manuelle
    ' Best" & ChrW(228) & "tigung durch den Nutzer. Bei allen anderen
    ' Hintergrundfarben (kein Hintergrund, gr" & ChrW(252) & "n, etc.)
    ' wurde der Wert programmatisch gesetzt -> kein Vermerk.
    ' =============================================
    Dim vorherigeFarbe As Long
    vorherigeFarbe = Target.Interior.color
    
    ' GELB = RGB(255, 235, 156) = Automatik war unsicher
    Dim istVorherGelb As Boolean
    istVorherGelb = (vorherigeFarbe = RGB(255, 235, 156))
    
    ' Wenn NICHT gelb -> keine manuelle Best" & ChrW(228) & "tigung, nichts tun
    If Not istVorherGelb Then Exit Sub
    
    On Error GoTo MonatError
    Application.EnableEvents = False
    
    Dim ws As Worksheet
    Set ws = Me
    
    Dim rowBK As Long
    rowBK = Target.Row
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo MonatError
    
    ' 1. Hell-gr" & ChrW(252) & "nen Hintergrund setzen (= manuell best" & ChrW(228) & "tigt)
    ws.Cells(rowBK, BK_COL_MONAT_PERIODE).Interior.color = RGB(198, 239, 206)
    
    ' 2. Lern-Vermerk in Spalte L anh" & ChrW(228) & "ngen
    '    Der Vermerk enth" & ChrW(228) & "lt "Folgemonat" als Marker f" & ChrW(252) & "r den Lern-Check
    '    in ErmittleMonatPeriode. Zus" & ChrW(228) & "tzlich wird der gew" & ChrW(228) & "hlte Monat
    '    dokumentiert, damit nachvollziehbar ist was der Nutzer gew" & ChrW(228) & "hlt hat.
    Dim bestehendeBemerkung As String
    bestehendeBemerkung = Trim(CStr(ws.Cells(rowBK, BK_COL_BEMERKUNG).value))
    
    Dim lernVermerk As String
    lernVermerk = "Folgemonat manuell best" & ChrW(228) & "tigt: " & monatWert
    
    If bestehendeBemerkung = "" Then
        ws.Cells(rowBK, BK_COL_BEMERKUNG).value = lernVermerk
    Else
        ' Pr" & ChrW(252) & "fen ob der Lern-Vermerk schon vorhanden ist (Vermeidung von Duplikaten)
        If InStr(LCase(bestehendeBemerkung), "folgemonat manuell best") = 0 Then
            ' Noch kein Lern-Vermerk -> anh" & ChrW(228) & "ngen
            ws.Cells(rowBK, BK_COL_BEMERKUNG).value = bestehendeBemerkung & vbLf & lernVermerk
        Else
            ' Lern-Vermerk bereits vorhanden -> aktualisieren
            ' (alten Vermerk durch neuen ersetzen)
            Dim zeilen() As String
            zeilen = Split(bestehendeBemerkung, vbLf)
            
            Dim neueBemerkung As String
            neueBemerkung = ""
            
            Dim z As Long
            For z = LBound(zeilen) To UBound(zeilen)
                If InStr(LCase(zeilen(z)), "folgemonat manuell best") = 0 Then
                    ' Zeile behalten (kein alter Lern-Vermerk)
                    If neueBemerkung = "" Then
                        neueBemerkung = zeilen(z)
                    Else
                        neueBemerkung = neueBemerkung & vbLf & zeilen(z)
                    End If
                End If
            Next z
            
            ' Neuen Lern-Vermerk anh" & ChrW(228) & "ngen
            If neueBemerkung = "" Then
                ws.Cells(rowBK, BK_COL_BEMERKUNG).value = lernVermerk
            Else
                ws.Cells(rowBK, BK_COL_BEMERKUNG).value = neueBemerkung & vbLf & lernVermerk
            End If
        End If
    End If
    
    ' 3. Auch den evtl. vorhandenen GELB-Hinweis aus der Ultimo-5-Logik entfernen
    '    (der ist jetzt nicht mehr relevant, da der Nutzer manuell best" & ChrW(228) & "tigt hat)
    Dim aktBemerkung As String
    aktBemerkung = CStr(ws.Cells(rowBK, BK_COL_BEMERKUNG).value)
    
    If InStr(LCase(aktBemerkung), "ultimo-5: bitte pr") > 0 Then
        Dim zeilen2() As String
        zeilen2 = Split(aktBemerkung, vbLf)
        
        Dim bereinigt As String
        bereinigt = ""
        
        Dim z2 As Long
        For z2 = LBound(zeilen2) To UBound(zeilen2)
            If InStr(LCase(zeilen2(z2)), "ultimo-5: bitte pr") = 0 Then
                If bereinigt = "" Then
                    bereinigt = zeilen2(z2)
                Else
                    bereinigt = bereinigt & vbLf & zeilen2(z2)
                End If
            End If
        Next z2
        
        ws.Cells(rowBK, BK_COL_BEMERKUNG).value = bereinigt
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
' - Buchstaben-K" & ChrW(252) & "rzel (A, B, C...) f" & ChrW(252) & "r KATEGORIEN (nicht Spalten!)
'   Die Kategorien werden dynamisch aus der Kategorie-Tabelle
'   (Spalte J auf Daten!) gelesen, nach E/A getrennt, nicht redundant.
' - Eingabe: K" & ChrW(252) & "rzel;Betrag (z.B. "A;25,00")
'   Semikolon trennt K" & ChrW(252) & "rzel vom Betrag, Komma = Dezimalzeichen
' - Jede Kategorie hat " & ChrW(252) & "ber Spalte N (Zielspalte) eine Zuordnung
'   wohin der Betrag auf dem Bankkonto gebucht wird.
' - Bei mehrfacher Buchung in die gleiche Spalte werden Betr" & ChrW(228) & "ge addiert.
' - GR" & ChrW(220) & "N nur wenn ALLES zugeordnet, GELB bei teilweise/nichts.
' - Die Schleife endet sofort wenn der Betrag vollst" & ChrW(228) & "ndig zugeordnet
'   ist - es erscheint KEINE weitere InputBox mit "0,00 offen".
' ===============================================================
Private Sub VerarbeiteSammelzahlung(ByVal ws As Worksheet, ByVal rowBK As Long)
    
    Dim betrag As Double
    betrag = ws.Cells(rowBK, BK_COL_BETRAG).value
    Dim absBetrag As Double
    absBetrag = Abs(betrag)
    
    ' Rundung auf 2 Nachkommastellen (verhindert Flie" & ChrW(223) & "komma-Probleme)
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
        If Not dictKat.Exists(katName) Then
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
    
    ' K" & ChrW(252) & "rzel-Arrays aufbauen: A=erste Kategorie, B=zweite, usw.
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
    
    ' K" & ChrW(252) & "rzel-Liste f" & ChrW(252) & "r Anzeige zusammenbauen
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
    
    ' Spaltenbereich f" & ChrW(252) & "r Einnahmen/Ausgaben
    Dim startCol As Long
    Dim endCol As Long
    If istEinnahme Then
        startCol = BK_COL_EINNAHMEN_START
        endCol = BK_COL_EINNAHMEN_ENDE
    Else
        startCol = BK_COL_AUSGABEN_START
        endCol = BK_COL_AUSGABEN_ENDE
    End If
    
    ' Sammelzahlungs-Dokumentation f" & ChrW(252) & "r Spalte L aufbauen
    Dim sammelDoku As String
    sammelDoku = ""
    
    Do
        ' *** PR" & ChrW(220) & "FUNG VOR der InputBox: Ist der Betrag bereits vollst" & ChrW(228) & "ndig zugeordnet? ***
        restBetrag = Round(absBetrag - zugeordnet, 2)
        
        If restBetrag <= 0.01 Then
            ' Betrag ist vollst" & ChrW(228) & "ndig zugeordnet -> Schleife sofort beenden
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
        
        ' Eingabe parsen: "K" & ChrW(252) & "rzel;Betrag" - Semikolon ist IMMER der Trenner
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
        ' Komma -> Punkt f" & ChrW(252) & "r Val() (Komma ist Dezimalzeichen)
        teilBetragStr = Replace(teilBetragStr, ",", ".")
        
        Dim teilBetrag As Double
        teilBetrag = Val(teilBetragStr)
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
        
        ' K" & ChrW(252) & "rzel in Kategorie " & ChrW(252) & "bersetzen
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
        
        ' Zielspalte " & ChrW(252) & "ber die Kategorietabelle bestimmen (Spalte N = Zielspalte)
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
        
        ' Dokumentation f" & ChrW(252) & "r Spalte L aufbauen
        If sammelDoku <> "" Then sammelDoku = sammelDoku & vbLf
        sammelDoku = sammelDoku & kuerzelKatName(gefundenIdx) & ": " & _
                     Format(teilBetrag, "#,##0.00") & " " & ChrW(8364)
        
WeiterEingabe:
    Loop
    
    ' ===============================================================
    ' Farbe und Bemerkung abh" & ChrW(228) & "ngig vom Zuordnungsstatus
    ' ===============================================================
    
    ' Endg" & ChrW(252) & "ltigen Restbetrag nochmal sauber berechnen
    restBetrag = Round(absBetrag - zugeordnet, 2)
    
    If zugeordnet < 0.01 Then
        ' NICHTS zugeordnet (User hat sofort Abbrechen geklickt)
        ' -> GELB bleibt, Bemerkung bleibt wie von der Engine gesetzt
        Exit Sub
    End If
    
    If restBetrag <= 0.01 Then
        ' ALLES zugeordnet -> GR" & ChrW(220) & "N (sicher)
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

