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
' WORKSHEET: Tabelle3 (Bankkonto)
' VERSION: 10.0 - 13.02.2026
' ZWECK: Event-Handler für Blatt "Bankkonto"
' ÄNDERUNGEN v10.0:
'   - Punkt 2: GELB-Bestätigung bei gleichem Wert ermöglichen
'   - Punkt 3: Sammelzahlung Spalte I "Sammelzahlung" + GELB?GRÜN
'   - Punkt 4: Neue Fälligkeitstypen "jährlich (Jahr)" und
'              "jährlich (Jahr/Folgejahr)" für Spalte I
'   - Punkt 5: Bei Sammelzahlung steht "Sammelzahlung" in Spalte I
'   - Punkt 6: "Ultimo-5:" entfernt, nur noch
'              "Bitte prüfen ob Zahlung für ... oder Folgemonat gilt"
'   - Punkt 7: Versicherung dynamisch [Jahr/Folgejahr]
' ===============================================================

' ===============================================================
' Modulweite Variablen für den Vergleich (alter Wert)
' Werden in Worksheet_SelectionChange gespeichert und in
' Worksheet_Change ausgewertet, um unnötige Verarbeitung
' bei "Doppelklick + Enter ohne Änderung" zu verhindern.
' ===============================================================
Private m_AlterWertH As String
Private m_AlterWertI As String
' NEU v10.0: Merke auch ob die Zelle GELB war (für Bestätigung bei gleichem Wert)
Private m_AlterHintergrundH As Long
Private m_AlterHintergrundI As Long

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
            ' Keine Daten vorhanden - still zurücksetzen
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
        
        ' Sichtbare Daten prüfen
        On Error Resume Next
        visibleCellsCount = rngFilterBereich.SpecialCells(xlCellTypeVisible).count
        On Error GoTo 0
        
        If visibleCellsCount <= 1 And monatsIndex > 0 Then
            ' Keine sichtbaren Daten - still zurücksetzen
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
' NEU v10.0: Worksheet_SelectionChange
' Speichert den aktuellen Wert UND die Hintergrundfarbe der
' ausgewählten Zelle, damit Worksheet_Change prüfen kann ob
' sich der Wert tatsächlich geändert hat UND ob die Zelle
' GELB war (= Bestätigung durch gleichen Wert möglich).
' ===============================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Nur einzelne Zelle merken
    If Target.Cells.count <> 1 Then Exit Sub
    If Target.Row < BK_START_ROW Then Exit Sub
    
    Select Case Target.Column
        Case BK_COL_KATEGORIE
            m_AlterWertH = Trim(CStr(Target.value))
            m_AlterHintergrundH = Target.Interior.color
        Case BK_COL_MONAT_PERIODE
            m_AlterWertI = Trim(CStr(Target.value))
            m_AlterHintergrundI = Target.Interior.color
    End Select
End Sub

' ===============================================================
' Worksheet_Change: Manuelle Auswahl in Spalte H oder Spalte I
' ---------------------------------------------------------------
' v10.0: Bei GELB hinterlegten Zellen wird die Verarbeitung
'        IMMER ausgelöst (auch bei gleichem Wert), damit der
'        Nutzer den automatisch zugeordneten Wert bestätigen kann.
'        Bei nicht-GELBEN Zellen bleibt die bisherige Logik
'        "gleicher Wert = nichts tun" bestehen, aber der Nutzer
'        kann trotzdem via DropDown den Wert ändern.
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
            Dim neuerWertH As String
            neuerWertH = Trim(CStr(Target.value))
            
            ' v10.0: Bei GELB hinterlegter Zelle IMMER verarbeiten
            ' (auch bei gleichem Wert = Bestätigung)
            Dim warGelbH As Boolean
            warGelbH = (m_AlterHintergrundH = RGB(255, 235, 156))
            
            If neuerWertH = m_AlterWertH And Not warGelbH Then
                ' Keine Änderung und nicht GELB -> nichts tun
                Exit Sub
            End If
            
            Call VerarbeiteKategorieAenderung(Target, warGelbH)
        
        Case BK_COL_MONAT_PERIODE
            ' --- Spalte I: Monat/Periode manuell gewählt ---
            Dim neuerWertI As String
            neuerWertI = Trim(CStr(Target.value))
            
            ' v10.0: Bei GELB hinterlegter Zelle IMMER verarbeiten
            Dim warGelbI As Boolean
            warGelbI = (m_AlterHintergrundI = RGB(255, 235, 156))
            
            If neuerWertI = m_AlterWertI And Not warGelbI Then
                ' Keine Änderung und nicht GELB -> nichts tun
                Exit Sub
            End If
            
            Call VerarbeiteMonatAenderung(Target)
        
        Case Else
            ' Andere Spalten: nichts tun
            Exit Sub
    
    End Select
    
End Sub


' ===============================================================
' Spalte H: Kategorie-Änderung verarbeiten
' v10.0: Parameter warGelb hinzugefügt für Bestätigungslogik
'        Neue Fälligkeitstypen + Sammelzahlung Spalte I
' ===============================================================
Private Sub VerarbeiteKategorieAenderung(ByVal Target As Range, ByVal warGelb As Boolean)
    
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
        
        ' Prüfen ob die Zeile bereits GRÜN ist (= bereits vollständig zugeordnet)
        Dim aktFarbe As Long
        aktFarbe = Target.Interior.color
        
        If aktFarbe = RGB(198, 239, 206) Then
            ' Bereits vollständig zugeordnet -> Nutzer fragen
            Dim antwort As VbMsgBoxResult
            antwort = MsgBox("Diese Sammelzahlung wurde bereits vollst" & ChrW(228) & "ndig zugeordnet." & vbLf & vbLf & _
                             "M" & ChrW(246) & "chten Sie die Zuordnung verwerfen und neu aufteilen?", _
                             vbQuestion + vbYesNo, "Sammelzahlung - Zeile " & rowBK)
            
            If antwort = vbNo Then
                ' Nutzer will NICHT neu zuordnen -> nichts tun, GRÜN bleibt
                GoTo ChangeCleanup
            End If
            ' Nutzer will NEU zuordnen -> weiter zu VerarbeiteSammelzahlung
        End If
        
        ' v10.0 Punkt 5: "Sammelzahlung" in Spalte I setzen (GELB)
        ws.Cells(rowBK, BK_COL_MONAT_PERIODE).value = "Sammelzahlung"
        ws.Cells(rowBK, BK_COL_MONAT_PERIODE).Interior.color = RGB(255, 235, 156)
        
        Call VerarbeiteSammelzahlung(ws, rowBK)
        GoTo ChangeCleanup
    End If
    
    ' =============================================
    ' v10.0: Bei GELB-Bestätigung (gleicher Wert wie zuvor)
    ' -> GRÜN setzen, Bemerkung "manuell bestätigt"
    ' =============================================
    If warGelb Then
        ' GELB -> GRÜN (Bestätigung)
        ApplyKategorie ws.Cells(rowBK, BK_COL_KATEGORIE), kategorie, "GRUEN"
        
        ' Bemerkung aktualisieren
        Dim bestBemerkung As String
        bestBemerkung = Trim(CStr(ws.Cells(rowBK, BK_COL_BEMERKUNG).value))
        
        Dim bestaetigungsVermerk As String
        bestaetigungsVermerk = "Kategorie manuell best" & ChrW(228) & "tigt: " & kategorie
        
        If bestBemerkung = "" Then
            ws.Cells(rowBK, BK_COL_BEMERKUNG).value = bestaetigungsVermerk
        Else
            ws.Cells(rowBK, BK_COL_BEMERKUNG).value = bestBemerkung & vbLf & bestaetigungsVermerk
        End If
        
        ' Betragszuordnung trotzdem ausführen (falls noch nicht geschehen)
        ApplyBetragsZuordnung ws, rowBK
        GoTo ChangeCleanup
    End If
    
    ' --- NORMALE KATEGORIE (auch bei Änderung von GRÜN) ---
    
    ' 1. Alte Beträge in Spalten M-Z löschen
    Dim c As Long
    For c = BK_COL_EINNAHMEN_START To BK_COL_AUSGABEN_ENDE
        ws.Cells(rowBK, c).ClearContents
    Next c
    
    ' 2. Kategorie GRÜN färben
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
            
            ' v10.0 Punkt 4: ErmittleMonatPeriode mit neuen Fälligkeitstypen
            Dim ergebnis As String
            ergebnis = ErmittleMonatPeriode(kategorie, CDate(ws.Cells(rowBK, BK_COL_DATUM).value), _
                                            faelligkeit, ws, rowBK)
            
            ' v10.0 Punkt 4: Neue Fälligkeitstypen verarbeiten
            ' Ergebnis kann sein:
            '   "GELB|Monatsname" -> Ultimo-Prüfung
            '   "JAEHRLICH_JAHR|Kategoriename" -> z.B. "Endabrechnung 2025"
            '   "JAEHRLICH_JAHRFJ|Kategoriename" -> z.B. "Versicherung 2025/2026"
            '   Normaler Monatsname
            
            If Left(ergebnis, 5) = "GELB|" Then
                ' GELB-Fall: Monat extrahieren und setzen
                Dim monatName As String
                monatName = Mid(ergebnis, 6) ' alles nach "GELB|"
                
                ws.Cells(rowBK, BK_COL_MONAT_PERIODE).value = monatName
                
                ' GELB-Hintergrund setzen (= Nutzer soll prüfen)
                ws.Cells(rowBK, BK_COL_MONAT_PERIODE).Interior.color = RGB(255, 235, 156)
                
                ' v10.0 Punkt 6: Bemerkung OHNE "Ultimo-5:"
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
            Else
                ' Normaler Fall: Monat/Periode direkt setzen (kein GELB)
                ws.Cells(rowBK, BK_COL_MONAT_PERIODE).value = ergebnis
            End If
            
            Call EntladeEinstellungenCache
        End If
    End If
    
    ' 5. Betragszuordnung ausführen
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
' Spalte I: Monat/Periode manuell geändert (DropDown)
' v10.0: Auch bei gleichem Wert verarbeiten wenn Zelle GELB war
'        (= Bestätigung der automatischen Zuordnung).
'        Bereinigung der Bemerkung ohne "Ultimo-5:" Prefix.
' ===============================================================
Private Sub VerarbeiteMonatAenderung(ByVal Target As Range)
    
    Dim monatWert As String
    monatWert = Trim(Target.value)
    If monatWert = "" Then Exit Sub
    
    ' =============================================
    ' v10.0: Prüfe ob die Zelle vorher GELB war
    ' Bei GELB = echte manuelle Bestätigung oder Änderung
    ' Bei nicht-GELB = normale DropDown-Änderung
    ' =============================================
    Dim vorherigeFarbe As Long
    vorherigeFarbe = m_AlterHintergrundI
    
    ' GELB = RGB(255, 235, 156) = Automatik war unsicher
    Dim istVorherGelb As Boolean
    istVorherGelb = (vorherigeFarbe = RGB(255, 235, 156))
    
    ' Wenn NICHT gelb -> keine manuelle Bestätigung, nichts tun
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
    
    ' 1. Hell-grünen Hintergrund setzen (= manuell bestätigt)
    ws.Cells(rowBK, BK_COL_MONAT_PERIODE).Interior.color = RGB(198, 239, 206)
    
    ' 2. Lern-Vermerk in Spalte L anhängen
    Dim bestehendeBemerkung As String
    bestehendeBemerkung = Trim(CStr(ws.Cells(rowBK, BK_COL_BEMERKUNG).value))
    
    Dim lernVermerk As String
    lernVermerk = "Folgemonat manuell best" & ChrW(228) & "tigt: " & monatWert
    
    If bestehendeBemerkung = "" Then
        ws.Cells(rowBK, BK_COL_BEMERKUNG).value = lernVermerk
    Else
        ' Prüfen ob der Lern-Vermerk schon vorhanden ist (Vermeidung von Duplikaten)
        If InStr(LCase(bestehendeBemerkung), "folgemonat manuell best") = 0 Then
            ' Noch kein Lern-Vermerk -> anhängen
            ws.Cells(rowBK, BK_COL_BEMERKUNG).value = bestehendeBemerkung & vbLf & lernVermerk
        Else
            ' Lern-Vermerk bereits vorhanden -> aktualisieren
            Dim zeilen() As String
            zeilen = Split(bestehendeBemerkung, vbLf)
            
            Dim neueBemerkung As String
            neueBemerkung = ""
            
            Dim z As Long
            For z = LBound(zeilen) To UBound(zeilen)
                If InStr(LCase(zeilen(z)), "folgemonat manuell best") = 0 Then
                    If neueBemerkung = "" Then
                        neueBemerkung = zeilen(z)
                    Else
                        neueBemerkung = neueBemerkung & vbLf & zeilen(z)
                    End If
                End If
            Next z
            
            If neueBemerkung = "" Then
                ws.Cells(rowBK, BK_COL_BEMERKUNG).value = lernVermerk
            Else
                ws.Cells(rowBK, BK_COL_BEMERKUNG).value = neueBemerkung & vbLf & lernVermerk
            End If
        End If
    End If
    
    ' 3. v10.0 Punkt 6: GELB-Hinweis bereinigen (OHNE "Ultimo-5:" Prefix)
    '    Suche nach "bitte pr" statt "ultimo-5: bitte pr"
    Dim aktBemerkung As String
    aktBemerkung = CStr(ws.Cells(rowBK, BK_COL_BEMERKUNG).value)
    
    ' Alte "Ultimo-5:" Einträge UND neue "Bitte prüfen ob..." entfernen
    If InStr(LCase(aktBemerkung), "bitte pr") > 0 Then
        Dim zeilen2() As String
        zeilen2 = Split(aktBemerkung, vbLf)
        
        Dim bereinigt As String
        bereinigt = ""
        
        Dim z2 As Long
        For z2 = LBound(zeilen2) To UBound(zeilen2)
            ' Entferne sowohl alte "Ultimo-5:" als auch neue "Bitte prüfen" Zeilen
            If InStr(LCase(zeilen2(z2)), "ultimo-5: bitte pr") = 0 And _
               InStr(LCase(zeilen2(z2)), "bitte pr" & ChrW(252) & "fen ob zahlung f") = 0 Then
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
' Sammelzahlung verarbeiten (v10.0)
' ---------------------------------------------------------------
' v10.0: Spalte I wird auf "Sammelzahlung" gesetzt (GELB).
'        Nach vollständiger Zuordnung wechselt Spalte I auf GRÜN.
'        Dokumentation der Aufschlüsselung in Spalte L.
' ===============================================================
Private Sub VerarbeiteSammelzahlung(ByVal ws As Worksheet, ByVal rowBK As Long)
    
    Dim betrag As Double
    betrag = ws.Cells(rowBK, BK_COL_BETRAG).value
    Dim absBetrag As Double
    absBetrag = Abs(betrag)
    
    ' Rundung auf 2 Nachkommastellen (verhindert Fließkomma-Probleme)
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
    
    ' Kürzel-Arrays aufbauen: A=erste Kategorie, B=zweite, usw.
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
    
    ' Kürzel-Liste für Anzeige zusammenbauen
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
    
    ' Spaltenbereich für Einnahmen/Ausgaben
    Dim startCol As Long
    Dim endCol As Long
    If istEinnahme Then
        startCol = BK_COL_EINNAHMEN_START
        endCol = BK_COL_EINNAHMEN_ENDE
    Else
        startCol = BK_COL_AUSGABEN_START
        endCol = BK_COL_AUSGABEN_ENDE
    End If
    
    ' Sammelzahlungs-Dokumentation für Spalte L aufbauen
    Dim sammelDoku As String
    sammelDoku = ""
    
    Do
        ' *** PRÜFUNG VOR der InputBox: Ist der Betrag bereits vollständig zugeordnet? ***
        restBetrag = Round(absBetrag - zugeordnet, 2)
        
        If restBetrag <= 0.01 Then
            ' Betrag ist vollständig zugeordnet -> Schleife sofort beenden
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
        
        ' Eingabe parsen: "Kürzel;Betrag" - Semikolon ist IMMER der Trenner
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
        ' Komma -> Punkt für Val() (Komma ist Dezimalzeichen)
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
        
        ' Kürzel in Kategorie übersetzen
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
        
        ' Zielspalte über die Kategorietabelle bestimmen (Spalte N = Zielspalte)
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
        
        ' Dokumentation für Spalte L aufbauen
        If sammelDoku <> "" Then sammelDoku = sammelDoku & vbLf
        sammelDoku = sammelDoku & kuerzelKatName(gefundenIdx) & ": " & _
                     Format(teilBetrag, "#,##0.00") & " " & ChrW(8364)
        
WeiterEingabe:
    Loop
    
    ' ===============================================================
    ' Farbe und Bemerkung abhängig vom Zuordnungsstatus
    ' ===============================================================
    
    ' Endgültigen Restbetrag nochmal sauber berechnen
    restBetrag = Round(absBetrag - zugeordnet, 2)
    
    If zugeordnet < 0.01 Then
        ' NICHTS zugeordnet (User hat sofort Abbrechen geklickt)
        ' -> GELB bleibt, Bemerkung bleibt wie von der Engine gesetzt
        Exit Sub
    End If
    
    If restBetrag <= 0.01 Then
        ' ALLES zugeordnet -> GRÜN (sicher)
        ApplyKategorie ws.Cells(rowBK, BK_COL_KATEGORIE), _
                       ws.Cells(rowBK, BK_COL_KATEGORIE).value, "GRUEN"
        
        ' v10.0 Punkt 5: Auch Spalte I auf GRÜN setzen
        ws.Cells(rowBK, BK_COL_MONAT_PERIODE).value = "Sammelzahlung"
        ws.Cells(rowBK, BK_COL_MONAT_PERIODE).Interior.color = RGB(198, 239, 206)
        
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

