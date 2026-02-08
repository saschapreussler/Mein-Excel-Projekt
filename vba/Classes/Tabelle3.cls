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
        
        ' Monat in Hilfszelle AG4 schreiben
        ' 0 = ganzes Jahr, 1-12 = jeweiliger Monat
        wsDaten.Range("AG4").value = monatsIndex
        
        ' Letzte Datenzeile ermitteln
        letzteDatenZeile = .Cells(.Rows.count, "A").End(xlUp).Row
        
        If letzteDatenZeile < ersteDatenZeile Then
            ' Keine Daten vorhanden - still zuruecksetzen
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
        
        ' Sichtbare Daten pruefen
        On Error Resume Next
        visibleCellsCount = rngFilterBereich.SpecialCells(xlCellTypeVisible).count
        On Error GoTo 0
        
        If visibleCellsCount <= 1 And monatsIndex > 0 Then
            ' Keine sichtbaren Daten - still zuruecksetzen
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
' (BK_COL_KATEGORIE = 8) manuell auswaehlt und:
' - Normale Kategorie: Spalten M-Z leeren, GRUEN setzen,
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
    
    ' Bereits GRUEN? Dann nicht erneut verarbeiten
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
    
    ' 1. Alte Betraege in Spalten M-Z loeschen
    Dim c As Long
    For c = BK_COL_EINNAHMEN_START To BK_COL_AUSGABEN_ENDE
        ws.Cells(rowBK, c).ClearContents
    Next c
    
    ' 2. Kategorie GRUEN faerben
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
        MsgBox "Fehler in Worksheet_Change: " & Err.Description, vbExclamation
        Err.Clear
    End If
End Sub


' ===============================================================
' Sammelzahlung verarbeiten: InputBox mit Spaltenkoepfen,
' schrittweise Betragseingabe mit laufender Summe.
' ===============================================================
Private Sub VerarbeiteSammelzahlung(ByVal ws As Worksheet, ByVal rowBK As Long)
    
    Dim betrag As Double
    betrag = ws.Cells(rowBK, BK_COL_BETRAG).value
    Dim absBetrag As Double
    absBetrag = Abs(betrag)
    
    Dim istEinnahme As Boolean
    istEinnahme = (betrag > 0)
    
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
    
    ' Spaltenkoepfe aus Zeile 27 sammeln
    Dim spaltenListe As String
    spaltenListe = "Verf" & ChrW(252) & "gbare Spalten (" & _
                   IIf(istEinnahme, "Einnahmen", "Ausgaben") & "):" & vbLf & vbLf
    
    Dim colCount As Long
    colCount = 0
    For c = startCol To endCol
        Dim headerText As String
        headerText = Trim(ws.Cells(BK_HEADER_ROW, c).value)
        If headerText <> "" Then
            colCount = colCount + 1
            spaltenListe = spaltenListe & colCount & ") " & headerText & _
                           "  [Spalte " & Chr(64 + c) & "]" & vbLf
        End If
    Next c
    
    ' Eingabeschleife
    Dim restBetrag As Double
    restBetrag = absBetrag
    Dim zugeordnet As Double
    zugeordnet = 0
    
    Do While restBetrag > 0.01
        Dim prompt As String
        prompt = spaltenListe & vbLf & _
                 "Gesamtbetrag: " & Format(absBetrag, "#,##0.00") & " " & ChrW(8364) & vbLf & _
                 "Bereits zugeordnet: " & Format(zugeordnet, "#,##0.00") & " " & ChrW(8364) & vbLf & _
                 "Noch offen: " & Format(restBetrag, "#,##0.00") & " " & ChrW(8364) & vbLf & vbLf & _
                 "Geben Sie den Spaltennamen (z.B. '" & _
                 Trim(ws.Cells(BK_HEADER_ROW, startCol).value) & _
                 "') und den Betrag ein:" & vbLf & _
                 "Format: Spaltenname;Betrag" & vbLf & _
                 "(z.B. Beitrag;25,00)" & vbLf & vbLf & _
                 "Abbrechen = restlichen Betrag nicht zuordnen"
        
        Dim eingabe As String
        eingabe = InputBox(prompt, "Sammelzahlung aufteilen - Zeile " & rowBK)
        
        ' Abbrechen oder leer
        If eingabe = "" Then Exit Do
        
        ' Eingabe parsen: "Spaltenname;Betrag"
        Dim teile() As String
        teile = Split(eingabe, ";")
        
        If UBound(teile) < 1 Then
            MsgBox "Ung" & ChrW(252) & "ltiges Format!" & vbLf & _
                   "Bitte im Format 'Spaltenname;Betrag' eingeben.", vbExclamation
            GoTo WeiterEingabe
        End If
        
        Dim zielSpalte As String
        zielSpalte = Trim(teile(0))
        
        Dim teilBetragStr As String
        teilBetragStr = Trim(teile(1))
        ' Komma -> Punkt fuer Val()
        teilBetragStr = Replace(teilBetragStr, ",", ".")
        
        Dim teilBetrag As Double
        teilBetrag = Val(teilBetragStr)
        
        If teilBetrag <= 0 Then
            MsgBox "Ung" & ChrW(252) & "ltiger Betrag: " & teile(1), vbExclamation
            GoTo WeiterEingabe
        End If
        
        If teilBetrag > restBetrag + 0.01 Then
            MsgBox "Betrag " & Format(teilBetrag, "#,##0.00") & " " & ChrW(8364) & _
                   " ist gr" & ChrW(246) & ChrW(223) & "er als der Restbetrag " & _
                   Format(restBetrag, "#,##0.00") & " " & ChrW(8364) & "!", vbExclamation
            GoTo WeiterEingabe
        End If
        
        ' Zielspalte finden (nach Ueberschrift in Zeile 27)
        Dim zielCol As Long
        zielCol = 0
        For c = startCol To endCol
            If StrComp(Trim(ws.Cells(BK_HEADER_ROW, c).value), zielSpalte, vbTextCompare) = 0 Then
                zielCol = c
                Exit For
            End If
        Next c
        
        If zielCol = 0 Then
            MsgBox "Spalte '" & zielSpalte & "' nicht gefunden!" & vbLf & _
                   "Bitte exakten Spaltennamen aus der Liste verwenden.", vbExclamation
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
    
    ' Kategorie GRUEN faerben und Bemerkung setzen
    ApplyKategorie ws.Cells(rowBK, BK_COL_KATEGORIE), _
                   ws.Cells(rowBK, BK_COL_KATEGORIE).value, "GRUEN"
    
    If restBetrag > 0.01 Then
        ws.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            "manuell zugeordnet (Sammelzahlung aufgeteilt, " & _
            Format(restBetrag, "#,##0.00") & " " & ChrW(8364) & " nicht zugeordnet)"
    Else
        ws.Cells(rowBK, BK_COL_BEMERKUNG).value = _
            "manuell zugeordnet (Sammelzahlung aufgeteilt)"
    End If
    
End Sub

