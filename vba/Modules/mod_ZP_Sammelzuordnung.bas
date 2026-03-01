Attribute VB_Name = "mod_ZP_Sammelzuordnung"
Option Explicit

' ===============================================================
' MODUL: mod_ZP_Sammelzuordnung
' Ausgelagert aus mod_Zahlungspruefung
' Enth?lt: Sammel?berweisungen erkennen + manuelle Monatszuordnung
' ===============================================================


' ===============================================================
' SAMMELUEBERWEISUNGEN: Erkennung und manuelle Zuordnung
' ===============================================================
Public Sub BearbeiteSammelUeberweisungZP(ByVal wsBK As Worksheet, _
                                          ByVal zeile As Long)
    
    On Error GoTo ErrorHandler
    
    Dim gesamtBetrag As Double
    gesamtBetrag = Abs(wsBK.Cells(zeile, BK_COL_BETRAG).value)
    
    If gesamtBetrag = 0 Then
        MsgBox "Kein Betrag in Zeile " & zeile & " gefunden!", vbExclamation
        Exit Sub
    End If
    
    Dim kategorien() As String
    Dim sollBetraege() As Double
    Dim anzahl As Long
    
    Call HoleKategorienAusEinstellungenZP(kategorien, sollBetraege, anzahl)
    
    If anzahl = 0 Then
        MsgBox "Keine Kategorien in Einstellungen gefunden!", vbExclamation
        Exit Sub
    End If
    
    Dim ergebnis As String
    ergebnis = ZeigeSammelZuordnungDialogZP(gesamtBetrag, kategorien, sollBetraege, anzahl)
    
    If ergebnis <> "" Then
        wsBK.Cells(zeile, BK_COL_BEMERKUNG).value = "SAMMEL:" & vbLf & ergebnis
        MsgBox "Sammel" & ChrW(252) & "berweisung erfolgreich zugeordnet!", vbInformation
    Else
        MsgBox "Zuordnung abgebrochen.", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler bei Sammel" & ChrW(252) & "berweisung: " & Err.Description, vbCritical
    
End Sub


' ===============================================================
' HILFSFUNKTION: Holt alle Kategorien aus Einstellungen
' ===============================================================
Private Sub HoleKategorienAusEinstellungenZP(ByRef kategorien() As String, _
                                              ByRef sollBetraege() As Double, _
                                              ByRef anzahl As Long)
    
    Dim wsEinst As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim kat As String
    
    Set wsEinst = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    lastRow = wsEinst.Cells(wsEinst.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    
    anzahl = 0
    ReDim kategorien(1 To lastRow - ES_START_ROW + 1)
    ReDim sollBetraege(1 To lastRow - ES_START_ROW + 1)
    
    For r = ES_START_ROW To lastRow
        kat = Trim(CStr(wsEinst.Cells(r, ES_COL_KATEGORIE).value))
        If kat <> "" Then
            anzahl = anzahl + 1
            kategorien(anzahl) = kat
            sollBetraege(anzahl) = wsEinst.Cells(r, ES_COL_SOLL_BETRAG).value
        End If
    Next r
    
    If anzahl > 0 Then
        ReDim Preserve kategorien(1 To anzahl)
        ReDim Preserve sollBetraege(1 To anzahl)
    End If
    
End Sub


' ===============================================================
' HILFSFUNKTION: Zeigt Dialog fuer Sammelzuordnung (Platzhalter)
' ===============================================================
Private Function ZeigeSammelZuordnungDialogZP(ByVal gesamtBetrag As Double, _
                                               ByRef kategorien() As String, _
                                               ByRef sollBetraege() As Double, _
                                               ByVal anzahl As Long) As String
    
    Dim ergebnis As String
    ergebnis = "Mitgliedsbeitrag: 7,50 " & ChrW(8364) & vbLf & _
               "Pachtgeb" & ChrW(252) & "hr: 25,00 " & ChrW(8364) & vbLf & _
               "Wasserkosten: 12,50 " & ChrW(8364)
    
    ZeigeSammelZuordnungDialogZP = ergebnis
    
End Function


' ===============================================================
' MANUELLE ZUORDNUNG: Monatszuordnung bei Problemfaellen
' ===============================================================
Public Function FrageNachManuellerMonatszuordnungZP(ByVal wsBK As Worksheet, _
                                                      ByVal zeile As Long) As Long
    
    Dim zahlDatum As Date
    Dim betrag As Double
    Dim Name As String
    Dim prompt As String
    Dim antwort As String
    Dim monat As Long
    
    zahlDatum = wsBK.Cells(zeile, BK_COL_DATUM).value
    betrag = wsBK.Cells(zeile, BK_COL_BETRAG).value
    Name = Trim(CStr(wsBK.Cells(zeile, BK_COL_NAME).value))
    
    prompt = "Die Zahlung kann keinem Monat zugeordnet werden:" & vbLf & vbLf & _
             "Datum: " & Format(zahlDatum, "dd.mm.yyyy") & vbLf & _
             "Betrag: " & Format(betrag, "#,##0.00 ") & ChrW(8364) & vbLf & _
             "Name: " & Name & vbLf & vbLf & _
             "Bitte geben Sie den Zielmonat ein (1-12):"
    
    antwort = InputBox(prompt, "Manuelle Monatszuordnung", Month(zahlDatum))
    
    If antwort = "" Then
        FrageNachManuellerMonatszuordnungZP = 0
        Exit Function
    End If
    
    If Not IsNumeric(antwort) Then
        MsgBox "Ung" & ChrW(252) & "ltige Eingabe! Es muss eine Zahl zwischen 1 und 12 sein.", vbExclamation
        FrageNachManuellerMonatszuordnungZP = 0
        Exit Function
    End If
    
    monat = CLng(antwort)
    
    If monat < 1 Or monat > 12 Then
        MsgBox "Ung" & ChrW(252) & "ltige Eingabe! Es muss eine Zahl zwischen 1 und 12 sein.", vbExclamation
        FrageNachManuellerMonatszuordnungZP = 0
        Exit Function
    End If
    
    wsBK.Cells(zeile, BK_COL_MONAT_PERIODE).value = Format(monat, "00") & "/" & Year(zahlDatum)
    
    MsgBox "Zahlung wurde Monat " & monat & "/" & Year(zahlDatum) & " zugeordnet.", vbInformation
    
    FrageNachManuellerMonatszuordnungZP = monat
    
End Function












