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
' Worksheet_Change: Manuelle Kategorie-Auswahl in Spalte H
' ---------------------------------------------------------------
' Erkennt wenn der Nutzer per DropDown eine Kategorie in Spalte H
' (BK_COL_KATEGORIE = 8) manuell auswählt oder ändert und:
' - Sammelzahlung: Prüft ob bereits zugeordnet (GRÜN), fragt
'   per MsgBox ob neu zugeordnet werden soll.
' - Normale Kategorie: Spalten M-Z leeren, GRÜN setzen,
'   "manuell zugeordnet" als Bemerkung, ApplyBetragsZuordnung
' - Bei Kategorie-Änderung (auch von GRÜN): alte Beträge
'   werden gelöscht und neu zugeordnet.
' ===============================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    
    ' Nur Spalte H (Kategorie), ab Datenzeile 28, einzelne Zelle
    If Target.Cells.count <> 1 Then Exit Sub
    If Target.Column <> BK_COL_KATEGORIE Then Exit Sub
    If Target.Row < BK_START_ROW Then Exit Sub
    
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
            antwort = MsgBox("Diese Sammelzahlung wurde bereits vollständig zugeordnet." & vbLf & vbLf & _
                             "Möchten Sie die Zuordnung verwerfen und neu aufteilen?", _
                             vbQuestion + vbYesNo, "Sammelzahlung - Zeile " & rowBK)
            
            If antwort = vbNo Then
                ' Nutzer will NICHT neu zuordnen -> nichts tun, GRÜN bleibt
                GoTo ChangeCleanup
            End If
            ' Nutzer will NEU zuordnen -> weiter zu VerarbeiteSammelzahlung
        End If
        
        Call VerarbeiteSammelzahlung(ws, rowBK)
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
' Sammelzahlung verarbeiten (v9.5)
' ---------------------------------------------------------------
' Komfortable InputBox mit:
' - Name des Zahlers und Verwendungszweck als Kontext
' - Buchstaben-Kürzel (A, B, C...) für KATEGORIEN (nicht Spalten!)
'   Die Kategorien werden dynamisch aus der Kategorie-Tabelle
'   (Spalte J auf Daten!) gelesen, nach E/A getrennt, nicht redundant.
' - Eingabe: Kürzel;Betrag (z.B. "A;25,00")
'   Semikolon trennt Kürzel vom Betrag, Komma = Dezimalzeichen
' - Jede Kategorie hat über Spalte N (Zielspalte) eine Zuordnung
'   wohin der Betrag auf dem Bankkonto gebucht wird.
' - Bei mehrfacher Buchung in die gleiche Spalte werden Beträge addiert.
' - GRÜN nur wenn ALLES zugeordnet, GELB bei teilweise/nichts.
' - Die Schleife endet sofort wenn der Betrag vollständig zugeordnet
'   ist - es erscheint KEINE weitere InputBox mit "0,00 offen".
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
        MsgBox "Keine Kategorien für " & IIf(istEinnahme, "Einnahmen", "Ausgaben") & _
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
    
    Do
        ' *** PRÜFUNG VOR der InputBox: Ist der Betrag bereits vollständig zugeordnet? ***
        restBetrag = Round(absBetrag - zugeordnet, 2)
        
        If restBetrag <= 0.01 Then
            ' Betrag ist vollständig zugeordnet -> Schleife sofort beenden
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
        teilBetrag = Round(teilBetrag, 2)
        
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
            MsgBox "Kürzel '" & eingabeKuerzel & "' nicht gefunden!" & vbLf & _
                   "Erlaubte Kürzel: A bis " & kuerzelBuchstabe(anzahlKat), vbExclamation
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

