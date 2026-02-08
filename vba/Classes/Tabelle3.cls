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

Private Sub lst_ImportReport_Click()
End Sub

Private Sub Worksheet_Activate()
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
        On Error GoTo 0
        
        If Me.cbo_RechnungsMonat.ListIndex = -1 Then GoTo SafeExit
        
        ' Abrechnungsjahr lesen
        If IsNumeric(wsStart.Range("F1").value) Then
            jahr = CLng(wsStart.Range("F1").value)
        Else
            MsgBox "Fehler: Ungültiges Abrechnungsjahr in Startmenü!F1", vbCritical
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
        
        ' Monat in Hilfszelle AG4 schreiben
        ' 0 = ganzes Jahr, 1-12 = jeweiliger Monat
        wsDaten.Range("AG4").value = monatsIndex
        
        ' Letzte Datenzeile ermitteln
        letzteDatenZeile = .Cells(.Rows.count, "A").End(xlUp).Row
        
        If letzteDatenZeile < ersteDatenZeile Then
            ' Keine Daten vorhanden - still zurücksetzen
            Me.cbo_RechnungsMonat.ListIndex = 0
            wsDaten.Range("AG4").value = 0
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
            wsDaten.Range("AG4").value = 0
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
' Worksheet_Change: Manuelle Kategorie-Auswahl in Spalte H
' ---------------------------------------------------------------
' Erkennt wenn der Nutzer per DropDown eine Kategorie in Spalte H
' (BK_COL_KATEGORIE = 8) manuell auswählt und:
' - Normale Kategorie: Spalten M-Z leeren, GRÜN setzen,
'   "manuell zugeordnet" als Bemerkung, ApplyBetragsZuordnung
' - Sammelzahlung: VerarbeiteSammelzahlung aufrufen
' ===============================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    
    ' Nur Spalte H (Kategorie), ab Datenzeile 28, einzelne Zelle
    If Target.Cells.count <> 1 Then Exit Sub
    If Target.Column <> BK_COL_KATEGORIE Then Exit Sub
    If Target.Row < BK_START_ROW Then Exit Sub
    
    Dim kategorie As String
    kategorie = Trim(Target.value)
    If kategorie = "" Then Exit Sub
    
    ' Bereits GRÜN? Dann nicht erneut verarbeiten
    If Target.Interior.color = RGB(198, 239, 206) Then Exit Sub
    
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
        Call VerarbeiteSammelzahlung(ws, rowBK)
        GoTo ChangeCleanup
    End If
    
    ' --- NORMALE KATEGORIE ---
    
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
            ws.Cells(rowBK, BK_COL_MONAT_PERIODE).value = _
                ErmittleMonatPeriode(kategorie, CDate(ws.Cells(rowBK, BK_COL_DATUM).value), faelligkeit)
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
        MsgBox "Fehler in Worksheet_Change: " & Err.Description, vbExclamation
        Err.Clear
    End If
End Sub


' ===============================================================
' Sammelzahlung verarbeiten (v9.3)
' ---------------------------------------------------------------
' Komfortable InputBox mit:
' - Name des Zahlers und Verwendungszweck als Kontext
' - Buchstaben-Kürzel (A, B, C...) statt langer Spaltennamen
' - Eingabe: Kürzel;Betrag (z.B. "A;25,00" oder "A;25")
'   Semikolon trennt Kürzel vom Betrag, Komma = Dezimalzeichen
' - Laufende Summe: Gesamt / Zugeordnet / Noch offen
' - GRÜN nur wenn ALLES zugeordnet
' - GELB bleibt wenn teilweise oder gar nichts zugeordnet
' Funktioniert identisch für Einnahmen und Ausgaben.
' ===============================================================
Private Sub VerarbeiteSammelzahlung(ByVal ws As Worksheet, ByVal rowBK As Long)
    
    Dim betrag As Double
    betrag = ws.Cells(rowBK, BK_COL_BETRAG).value
    Dim absBetrag As Double
    absBetrag = Abs(betrag)
    
    Dim istEinnahme As Boolean
    istEinnahme = (betrag > 0)
    
    ' Kontext-Info aus der Buchungszeile lesen
    Dim zahlerName As String
    zahlerName = Trim(ws.Cells(rowBK, BK_COL_NAME).value)
    If zahlerName = "" Then zahlerName = "(unbekannt)"
    
    Dim verwendungszweck As String
    verwendungszweck = Trim(ws.Cells(rowBK, BK_COL_VERWENDUNGSZWECK).value)
    If verwendungszweck = "" Then verwendungszweck = "(kein Verwendungszweck)"
    
    ' Bereich bestimmen (Einnahmen oder Ausgaben)
    Dim startCol As Long
    Dim endCol As Long
    If istEinnahme Then
        startCol = BK_COL_EINNAHMEN_START
        endCol = BK_COL_EINNAHMEN_ENDE
    Else
        startCol = BK_COL_AUSGABEN_START
        endCol = BK_COL_AUSGABEN_ENDE
    End If
    
    ' Spalten M-Z komplett leeren
    Dim c As Long
    For c = BK_COL_EINNAHMEN_START To BK_COL_AUSGABEN_ENDE
        ws.Cells(rowBK, c).ClearContents
    Next c
    
    ' ---------------------------------------------------------------
    ' Kürzel-Zuordnung aufbauen: A=erste Spalte, B=zweite, usw.
    ' ---------------------------------------------------------------
    Dim kuerzelBuchstabe() As String
    Dim kuerzelSpalte() As Long
    Dim kuerzelName() As String
    Dim anzahlSpalten As Long
    anzahlSpalten = 0
    
    ' Erst zählen
    For c = startCol To endCol
        Dim hdrTmp As String
        hdrTmp = Trim(ws.Cells(BK_HEADER_ROW, c).value)
        If hdrTmp <> "" Then anzahlSpalten = anzahlSpalten + 1
    Next c
    
    If anzahlSpalten = 0 Then
        MsgBox "Keine Spaltenüberschriften in Zeile " & BK_HEADER_ROW & " gefunden!", vbCritical
        Exit Sub
    End If
    
    ReDim kuerzelBuchstabe(1 To anzahlSpalten)
    ReDim kuerzelSpalte(1 To anzahlSpalten)
    ReDim kuerzelName(1 To anzahlSpalten)
    
    Dim idx As Long
    idx = 0
    For c = startCol To endCol
        Dim headerText As String
        headerText = Trim(ws.Cells(BK_HEADER_ROW, c).value)
        If headerText <> "" Then
            idx = idx + 1
            kuerzelBuchstabe(idx) = Chr(64 + idx)   ' A, B, C, D, E, F, G
            kuerzelSpalte(idx) = c
            kuerzelName(idx) = headerText
        End If
    Next c
    
    ' Kürzel-Liste für Anzeige zusammenbauen
    Dim spaltenInfo As String
    spaltenInfo = IIf(istEinnahme, "EINNAHMEN", "AUSGABEN") & "-Spalten:" & vbLf
    Dim i As Long
    For i = 1 To anzahlSpalten
        spaltenInfo = spaltenInfo & "  " & kuerzelBuchstabe(i) & ": " & _
                      kuerzelName(i) & vbLf
    Next i
    
    ' Eingabeschleife
    Dim restBetrag As Double
    restBetrag = absBetrag
    Dim zugeordnet As Double
    zugeordnet = 0
    
    Do While restBetrag > 0.01
        Dim prompt As String
        prompt = "=== BUCHUNGSDETAILS ===" & vbLf & _
                 "Name: " & zahlerName & vbLf & _
                 "Verwendungszweck: " & verwendungszweck & vbLf & _
                 "Gesamtbetrag: " & Format(absBetrag, "#,##0.00") & " " & ChrW(8364) & vbLf & vbLf & _
                 spaltenInfo & vbLf & _
                 "Bereits zugeordnet: " & Format(zugeordnet, "#,##0.00") & " " & ChrW(8364) & vbLf & _
                 "Noch offen: " & Format(restBetrag, "#,##0.00") & " " & ChrW(8364) & vbLf & vbLf & _
                 "Eingabe: Kürzel;Betrag  (z.B. A;25,00)" & vbLf & _
                 "Abbrechen = Rest nicht zuordnen"
        
        Dim eingabe As String
        eingabe = InputBox(prompt, "Sammelzahlung - Zeile " & rowBK)
        
        ' Abbrechen oder leer
        If eingabe = "" Then Exit Do
        
        ' Leerzeichen entfernen
        eingabe = Replace(eingabe, " ", "")
        
        ' Eingabe parsen: "Kürzel;Betrag" - Semikolon ist IMMER der Trenner
        ' Komma ist Dezimalzeichen (Euro,Cent)
        Dim trennPos As Long
        trennPos = InStr(eingabe, ";")
        
        If trennPos = 0 Or trennPos >= Len(eingabe) Then
            MsgBox "Ungültiges Format!" & vbLf & _
                   "Bitte im Format 'A;25,00' eingeben." & vbLf & _
                   "Semikolon trennt Kürzel vom Betrag.", vbExclamation
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
        
        If teilBetrag <= 0 Then
            MsgBox "Ungültiger Betrag: " & Mid(eingabe, trennPos + 1), vbExclamation
            GoTo WeiterEingabe
        End If
        
        If teilBetrag > restBetrag + 0.01 Then
            MsgBox "Betrag " & Format(teilBetrag, "#,##0.00") & " " & ChrW(8364) & _
                   " ist größer als der Restbetrag " & _
                   Format(restBetrag, "#,##0.00") & " " & ChrW(8364) & "!", vbExclamation
            GoTo WeiterEingabe
        End If
        
        ' Kürzel in Spalte übersetzen
        Dim zielCol As Long
        zielCol = 0
        For i = 1 To anzahlSpalten
            If kuerzelBuchstabe(i) = eingabeKuerzel Then
                zielCol = kuerzelSpalte(i)
                Exit For
            End If
        Next i
        
        If zielCol = 0 Then
            MsgBox "Kürzel '" & eingabeKuerzel & "' nicht gefunden!" & vbLf & _
                   "Erlaubte Kürzel: " & kuerzelBuchstabe(1) & _
                   " bis " & kuerzelBuchstabe(anzahlSpalten), vbExclamation
            GoTo WeiterEingabe
        End If
        
        ' Betrag eintragen (mit korrektem Vorzeichen)
        If istEinnahme Then
            ws.Cells(rowBK, zielCol).value = teilBetrag
        Else
            ws.Cells(rowBK, zielCol).value = -teilBetrag
        End If
        
        zugeordnet = zugeordnet + teilBetrag
        restBetrag = absBetrag - zugeordnet
        
WeiterEingabe:
    Loop
    
    ' ===============================================================
    ' Farbe und Bemerkung abhängig vom Zuordnungsstatus
    ' ===============================================================
    
    If zugeordnet < 0.01 Then
        ' NICHTS zugeordnet (User hat sofort Abbrechen geklickt)
        ' -> GELB bleibt, Bemerkung bleibt wie von der Engine gesetzt
        Exit Sub
    End If
    
    If restBetrag <= 0.01 Then
        ' ALLES zugeordnet -> GRÜN (sicher)
        ApplyKategorie ws.Cells(rowBK, BK_COL_KATEGORIE), _
                       ws.Cells(rowBK, BK_COL_KATEGORIE).value, "GRUEN"
        ws.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            "manuell zugeordnet (Sammelzahlung vollständig aufgeteilt)"
    Else
        ' TEILWEISE zugeordnet -> GELB bleibt
        ApplyKategorie ws.Cells(rowBK, BK_COL_KATEGORIE), _
                       ws.Cells(rowBK, BK_COL_KATEGORIE).value, "GELB"
        ws.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            "Sammelzahlung teilweise aufgeteilt, " & _
            Format(restBetrag, "#,##0.00") & " " & ChrW(8364) & " nicht zugeordnet"
    End If
    
End Sub

