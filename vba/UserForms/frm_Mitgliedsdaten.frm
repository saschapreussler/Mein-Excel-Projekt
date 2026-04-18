VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Mitgliedsdaten 
   Caption         =   "Mitgliedsdaten"
   ClientHeight    =   8580.001
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   7800
   OleObjectBlob   =   "frm_Mitgliedsdaten.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frm_Mitgliedsdaten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
































Option Explicit

Private Const WS_NAME_MITGLIEDER As String = "Mitgliederliste"
Private m_AlreadyInitialized As Boolean  ' Flag um doppelte Initialisierung zu vermeiden


' ***************************************************************
' HILFSPROZEDUR: Extrahiert lRow aus Tag (unterst�tzt auch "lRow|Grund|..." Format)
' ***************************************************************
Private Function GetLRowFromTag() As Long
    Dim tagStr As String
    Dim tagParts() As String
    
    tagStr = CStr(Me.tag)
    
    ' Pr�fe ob Tag das Format "lRow|..." hat
    If InStr(tagStr, "|") > 0 Then
        tagParts = Split(tagStr, "|")
        ' Pr�fe ob erstes Element numerisch ist
        If mod_Mitglieder_Logik.IsNumericTag(tagParts(0)) Then
            GetLRowFromTag = CLng(tagParts(0))
        Else
            ' F�r "NACHPAECHTER_NEU|..." Format
            GetLRowFromTag = 0
        End If
    Else
        ' Normales Format: nur lRow oder "NEU"
        If mod_Mitglieder_Logik.IsNumericTag(tagStr) Then
            GetLRowFromTag = CLng(tagStr)
        Else
            GetLRowFromTag = 0
        End If
    End If
End Function


' ***************************************************************
' HILFSPROZEDUR: Aktualisiert Labels basierend auf Funktion
' ***************************************************************
Private Sub AktualisiereLabelsFuerFunktion()
    Dim istMitgliedOhnePacht As Boolean
    
    ' Pr�fe ob cbo_Funktion einen Wert hat
    If Me.cbo_Funktion.value = "" Then
        ' Default setzen
        Me.lbl_PachtbeginnBezeichner.Caption = "Pachtbeginn"
        Me.lbl_PachtendeBezeichner.Caption = "Pachtende"
        Exit Sub
    End If
    
    istMitgliedOhnePacht = (Me.cbo_Funktion.value = "Mitglied ohne Pacht")
    
    If istMitgliedOhnePacht Then
        Me.lbl_PachtbeginnBezeichner.Caption = "Mitgliedsbeginn"
        Me.lbl_PachtendeBezeichner.Caption = "Mitgliedsende"
    Else
        Me.lbl_PachtbeginnBezeichner.Caption = "Pachtbeginn"
        Me.lbl_PachtendeBezeichner.Caption = "Pachtende"
    End If
End Sub

' ***************************************************************
' HILFSPROZEDUR: Setzt den Anzeigemodus der Form
' FIX v2: IsRemovalMode wird jetzt korrekt ausgewertet!
'      Im RemovalMode:
'      - Alle Daten-Labels (lbl_Anrede, lbl_Vorname, ...) SICHTBAR
'        ? zeigen die Daten des austretenden Mitglieds
'      - Alle Bezeichner-Labels (lbl_PachtbeginnBezeichner,
'        lbl_PachtendeBezeichner) SICHTBAR
'      - Alle TextBoxen UNSICHTBAR, AUSSER txt_Pachtende
'      - Alle ComboBoxen UNSICHTBAR
'      - Nur Buttons �bernehmen + Abbrechen sichtbar
' ***************************************************************
Public Sub SetMode(ByVal EditMode As Boolean, Optional ByVal IsNewEntry As Boolean = False, Optional ByVal IsRemovalMode As Boolean = False)
    
    Dim ctl As MSForms.Control
    
    If IsRemovalMode Then
        ' ===================================================
        ' AUSTRITTS-MODUS: Daten-Labels sichtbar (read-only),
        '                  nur txt_Pachtende editierbar
        ' ===================================================
        For Each ctl In Me.Controls
            If TypeOf ctl Is MSForms.label And Left(ctl.Name, 4) = "lbl_" Then
                ' ALLE Labels sichtbar: sowohl Bezeichner-Labels
                ' (lbl_PachtbeginnBezeichner, lbl_PachtendeBezeichner)
                ' als auch Daten-Labels (lbl_Anrede, lbl_Vorname, ...)
                ctl.Visible = True
            ElseIf TypeOf ctl Is MSForms.TextBox Then
                ' Nur txt_Pachtende sichtbar (editierbar)
                If ctl.Name = "txt_Pachtende" Then
                    ctl.Visible = True
                Else
                    ctl.Visible = False
                End If
            ElseIf TypeOf ctl Is MSForms.ComboBox Then
                ' Alle ComboBoxen ausblenden
                ctl.Visible = False
            End If
        Next ctl
        
        ' Buttons: nur �bernehmen + Abbrechen
        Me.cmd_Uebernehmen.Visible = True
        Me.cmd_Abbrechen.Visible = True
        Me.cmd_Bearbeiten.Visible = False
        Me.cmd_Entfernen.Visible = False
        Me.cmd_Anlegen.Visible = False
        
        ' Label-Text anpassen (Pachtende / Mitgliedsende)
        Call AktualisiereLabelsFuerFunktion
        
        Exit Sub
    End If
    
    ' ===================================================
    ' NORMALER MODUS (wie bisher, unver�ndert)
    ' ===================================================
    For Each ctl In Me.Controls
        If TypeOf ctl Is MSForms.label And Left(ctl.Name, 4) = "lbl_" Then
            ' Bezeichner-Labels sollen IMMER sichtbar sein
            If ctl.Name = "lbl_PachtbeginnBezeichner" Or ctl.Name = "lbl_PachtendeBezeichner" Then
                ctl.Visible = True
            Else
                ' Alle anderen Labels (Datenlabels): SICHTBAR im ViewMode, UNSICHTBAR im EditMode
                ctl.Visible = Not EditMode
            End If
        ElseIf TypeOf ctl Is MSForms.TextBox Or TypeOf ctl Is MSForms.ComboBox Then
            ' TextBoxen/ComboBoxen: UNSICHTBAR im ViewMode, SICHTBAR im EditMode
            ctl.Visible = EditMode
        End If
    Next ctl
    
    If CStr(Me.tag) = "NEU" Or InStr(CStr(Me.tag), "NACHPAECHTER_NEU") > 0 Then
        Me.cmd_Bearbeiten.Visible = False
        Me.cmd_Entfernen.Visible = False
        Me.cmd_Uebernehmen.Visible = False
        Me.cmd_Anlegen.Visible = True
        Me.cmd_Abbrechen.Visible = True
        
    ElseIf EditMode = True Then
        Me.cmd_Bearbeiten.Visible = False
        Me.cmd_Entfernen.Visible = False
        Me.cmd_Anlegen.Visible = False
        Me.cmd_Uebernehmen.Visible = True
        Me.cmd_Abbrechen.Visible = True
        
    Else
        ' ViewMode (Vorschau)
        Me.cmd_Bearbeiten.Visible = True
        Me.cmd_Entfernen.Visible = True
        Me.cmd_Uebernehmen.Visible = False
        Me.cmd_Anlegen.Visible = False
        Me.cmd_Abbrechen.Visible = False
    End If
    
    If EditMode = False Then Exit Sub
    
    ' Aktualisiere Labels nach Funktion
    Call AktualisiereLabelsFuerFunktion
    
    If CStr(Me.tag) <> "NEU" And InStr(CStr(Me.tag), "NACHPAECHTER_NEU") = 0 Then
        Dim lRow As Long
        lRow = GetLRowFromTag()
        
        If lRow > 0 Then
            Me.cbo_Parzelle.value = Me.lbl_Parzelle.Caption
            Me.cbo_Anrede.value = Me.lbl_Anrede.Caption
            Me.txt_Vorname.value = Me.lbl_Vorname.Caption
            Me.txt_Nachname.value = Me.lbl_Nachname.Caption
            Me.txt_Strasse.value = Me.lbl_Strasse.Caption
            Me.txt_Nummer.value = Me.lbl_Nummer.Caption
            Me.txt_PLZ.value = Me.lbl_PLZ.Caption
            Me.txt_Wohnort.value = Me.lbl_Wohnort.Caption
            Me.txt_Telefon.value = Me.lbl_Telefon.Caption
            Me.txt_Mobil.value = Me.lbl_Mobil.Caption
            Me.txt_Geburtstag.value = Me.lbl_Geburtstag.Caption
            Me.txt_Email.value = Me.lbl_Email.Caption
            Me.cbo_Funktion.value = Me.lbl_Funktion.Caption
            Me.txt_Pachtbeginn.value = Me.lbl_Pachtbeginn.Caption
            Me.txt_Pachtende.value = Me.lbl_Pachtende.Caption
        End If
    End If
    
End Sub

Private Sub cbo_Parzelle_Change()
' ***************************************************************
' EVENT: ComboBox Parzelle-�nderung
' Pr�ft ob Parzelle belegt ist und bietet Adress�bernahme an
' ***************************************************************
    Dim parzelle As String
    Dim tagStr As String
    
    ' Nur im NEU-Modus aktiv (nicht beim Bearbeiten)
    tagStr = CStr(Me.tag)
    If tagStr <> "NEU" And InStr(tagStr, "NACHPAECHTER_NEU") = 0 Then
        Exit Sub
    End If
    
    parzelle = Trim(Me.cbo_Parzelle.value)
    If parzelle = "" Then Exit Sub
    
    ' Pr�fe ob Parzelle belegt ist
    Call PruefeUndUebernehmeAdresse(parzelle)
    
    ' Setze Fokus auf cbo_Anrede
    On Error Resume Next
    Me.cbo_Anrede.SetFocus
    On Error GoTo 0
End Sub

' ***************************************************************
' HILFSPROZEDUR: Pr�ft Parzellenbelegung und bietet Adress�bernahme an
' ***************************************************************
Private Sub PruefeUndUebernehmeAdresse(ByVal parzelle As String)
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim mitgliederAufParzelle As Collection
    Dim mitgliedInfo As Variant
    Dim antwort As VbMsgBoxResult
    Dim auswahlIndex As Long
    Dim auswahlText As String
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    Set mitgliederAufParzelle = New Collection
    
    lastRow = ws.Cells(ws.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    ' Sammle alle aktiven Mitglieder auf dieser Parzelle
    For r = M_START_ROW To lastRow
        If StrComp(Trim(ws.Cells(r, M_COL_PARZELLE).value), parzelle, vbTextCompare) = 0 Then
            ' Nur aktive Mitglieder (ohne Pachtende)
            If Trim(ws.Cells(r, M_COL_PACHTENDE).value) = "" Then
                ' Speichere: Zeile|Nachname|Vorname|Strasse|Nummer|PLZ|Wohnort
                mitgliedInfo = Array( _
                    r, _
                    ws.Cells(r, M_COL_NACHNAME).value, _
                    ws.Cells(r, M_COL_VORNAME).value, _
                    ws.Cells(r, M_COL_STRASSE).value, _
                    ws.Cells(r, M_COL_NUMMER).value, _
                    ws.Cells(r, M_COL_PLZ).value, _
                    ws.Cells(r, M_COL_WOHNORT).value _
                )
                mitgliederAufParzelle.Add mitgliedInfo
            End If
        End If
    Next r
    
    ' Keine Mitglieder auf Parzelle gefunden
    If mitgliederAufParzelle.count = 0 Then
        Exit Sub
    End If
    
    ' EIN Mitglied auf Parzelle
    If mitgliederAufParzelle.count = 1 Then
        mitgliedInfo = mitgliederAufParzelle(1)
        
        antwort = MsgBox("Auf Parzelle " & parzelle & " ist bereits gemeldet:" & vbCrLf & _
                        mitgliedInfo(1) & ", " & mitgliedInfo(2) & vbCrLf & vbCrLf & _
                        "Adresse: " & mitgliedInfo(3) & " " & mitgliedInfo(4) & ", " & _
                        mitgliedInfo(5) & " " & mitgliedInfo(6) & vbCrLf & vbCrLf & _
                        "M�chten Sie diese Adresse �bernehmen?", _
                        vbYesNo + vbQuestion, "Adresse �bernehmen?")
        
        If antwort = vbYes Then
            Me.txt_Strasse.value = mitgliedInfo(3)
            Me.txt_Nummer.value = mitgliedInfo(4)
            Me.txt_PLZ.value = mitgliedInfo(5)
            Me.txt_Wohnort.value = mitgliedInfo(6)
        End If
        
    Else
        ' MEHRERE Mitglieder auf Parzelle - Auswahl anbieten
        auswahlText = "Auf Parzelle " & parzelle & " sind mehrere Personen gemeldet:" & vbCrLf & vbCrLf
        
        For i = 1 To mitgliederAufParzelle.count
            mitgliedInfo = mitgliederAufParzelle(i)
            auswahlText = auswahlText & i & ") " & mitgliedInfo(1) & ", " & mitgliedInfo(2) & vbCrLf & _
                         "    " & mitgliedInfo(3) & " " & mitgliedInfo(4) & ", " & _
                         mitgliedInfo(5) & " " & mitgliedInfo(6) & vbCrLf & vbCrLf
        Next i
        
        auswahlText = auswahlText & "M�chten Sie eine Adresse �bernehmen?"
        
        antwort = MsgBox(auswahlText, vbYesNo + vbQuestion, "Adresse �bernehmen?")
        
        If antwort = vbYes Then
            ' Zeige Auswahl-Dialog
            auswahlIndex = mod_Mitglieder_Logik.ZeigeAdressAuswahl(mitgliederAufParzelle)
            
            If auswahlIndex > 0 And auswahlIndex <= mitgliederAufParzelle.count Then
                mitgliedInfo = mitgliederAufParzelle(auswahlIndex)
                Me.txt_Strasse.value = mitgliedInfo(3)
                Me.txt_Nummer.value = mitgliedInfo(4)
                Me.txt_PLZ.value = mitgliedInfo(5)
                Me.txt_Wohnort.value = mitgliedInfo(6)
            End If
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Fehler bei Adress�bernahme: " & Err.Description
End Sub


Private Sub cmd_Bearbeiten_Click()
    Call SetMode(True, False, False)
End Sub

Private Sub cmd_Abbrechen_Click()
    Dim tagStr As String
    
    tagStr = CStr(Me.tag)
    
    If tagStr = "NEU" Or InStr(tagStr, "NACHPAECHTER_NEU") > 0 Then
        Unload Me
        Exit Sub
    End If
    
    ' Wenn Tag im Format "lRow|Grund|..." ist (nach Abbruch eines Austritts), stelle urspr�nglichen Tag wieder her
    If InStr(tagStr, "|") > 0 Then
        Dim tagParts() As String
        tagParts = Split(tagStr, "|")
        If mod_Mitglieder_Logik.IsNumericTag(tagParts(0)) Then
            Me.tag = tagParts(0)  ' Nur lRow behalten
        End If
    End If
    
    Call SetMode(False)
End Sub

' ***************************************************************
' EVENT: ComboBox Funktion-�nderung
' ***************************************************************
Private Sub cbo_Funktion_Change()
    Call AktualisiereLabelsFuerFunktion
End Sub

Private Sub cmd_Entfernen_Click()
    
    Dim lRow As Long
    Dim nachname As String
    Dim vorname As String
    Dim OldParzelle As String
    Dim OldMemberID As String
    Dim austrittsDatum As Date
    Dim ChangeReason As String
    Dim pachtEndeVal As String
    Dim auswahlOption As Integer
    Dim nachpaechterID As String
    Dim nachpaechterName As String
    Dim tagStr As String
    
    ' Sichere Tag-Extraktion mit Fehlerbehandlung
    On Error GoTo TagError
    tagStr = CStr(Me.tag)
    
    ' Extrahiere lRow aus Tag (unterst�tzt auch "lRow|Grund|..." Format)
    lRow = GetLRowFromTag()
    
    If lRow < M_START_ROW Then
        MsgBox "Interner Fehler: Keine g�ltige Zeilennummer f�r das Entfernen gefunden.", vbCritical
        Exit Sub
    End If
    
    On Error GoTo 0
    
    OldParzelle = Me.lbl_Parzelle.Caption
    
    ' === SICHERHEITSCHECK: Verein-Parzelle darf NIEMALS gel�scht werden ===
    If UCase(Trim(OldParzelle)) = "VEREIN" Then
        MsgBox "FEHLER: Die Verein-Parzelle darf nicht gel�scht oder entfernt werden!", vbCritical, "Operation nicht erlaubt"
        Exit Sub
    End If
    
    nachname = Me.lbl_Nachname.Caption
    vorname = Me.lbl_Vorname.Caption
    OldMemberID = ThisWorkbook.Worksheets(WS_MITGLIEDER).Cells(lRow, M_COL_MEMBER_ID).value
    
    ' Pr�fe ob Pachtende bereits gef�llt ist
    pachtEndeVal = Trim(Me.lbl_Pachtende.Caption)
    
    ' Zeige Austrittsauswahl-Dialog
    With frm_Austrittsauswahl
        .Show vbModal
        auswahlOption = .SelectedOption
        ChangeReason = .CustomReason
        nachpaechterID = .nachpaechterID
        nachpaechterName = .nachpaechterName
        Unload frm_Austrittsauswahl
    End With
    
    If auswahlOption = 0 Then
        ' Benutzer hat abgebrochen - stelle urspr�nglichen Tag wieder her
        Me.tag = lRow
        Exit Sub
    End If
    
    Select Case auswahlOption
        Case 1 ' Nachp�chter
            If ChangeReason = "" Then ChangeReason = "�bergabe an Nachp�chter"
            
            ' Pr�fe ob neuer Nachp�chter angelegt werden muss
            If nachpaechterID = "NACHPAECHTER_NEU" Then
                ' Speichere aktuellen Zustand im Tag
                Me.tag = lRow & "|" & ChangeReason & "|NACHPAECHTER_NEU|" & OldParzelle
                
                ' Verstecke aktuelles Formular
                Me.Hide
                
                ' Lade NEUES Formular f�r Nachp�chter
                Dim frmNachpaechter As frm_Mitgliedsdaten
                Set frmNachpaechter = New frm_Mitgliedsdaten
                
                With frmNachpaechter
                    .tag = "NACHPAECHTER_NEU|" & OldParzelle & "|" & Format(Date, "dd.mm.yyyy")
                    
                    ' Leere alle Felder
                    .cbo_Anrede.value = ""
                    .txt_Vorname.value = ""
                    .txt_Nachname.value = ""
                    .txt_Strasse.value = ""
                    .txt_Nummer.value = ""
                    .txt_PLZ.value = ""
                    .txt_Wohnort.value = ""
                    .txt_Telefon.value = ""
                    .txt_Mobil.value = ""
                    .txt_Geburtstag.value = ""
                    .txt_Email.value = ""
                    .txt_Pachtende.value = ""
                    
                    ' Vorbef�llen: Parzelle, Funktion, Pachtbeginn
                    .cbo_Parzelle.value = OldParzelle
                    .cbo_Funktion.value = "Mitglied mit Pacht"
                    .txt_Pachtbeginn.value = Format(Date, "dd.mm.yyyy")
                    
                    ' Setze Modus auf Bearbeiten
                    Call .SetMode(True, True, False)
                    
                    .Show vbModal
                End With
                
                ' Aufr�umen
                Set frmNachpaechter = Nothing
                
                ' Zeige aktuelles Formular wieder
                Me.Show
                
                ' Nach R�ckkehr: Verarbeite Austritt mit neuem Nachp�chter
                Call VerarbeiteAustrittNachNachpaechterErfassung(lRow, OldParzelle, OldMemberID, nachname, vorname, Date, ChangeReason)
                Exit Sub
            Else
                ' Bestehender Nachp�chter wurde ausgew�hlt
                ' Pr�fe ob Nachp�chter bereits eine Parzelle hat
                Call BearbeiteNachpaechterUebernahme(nachpaechterID, nachpaechterName, OldParzelle, lRow, OldMemberID, nachname, vorname, Date, ChangeReason)
                Exit Sub
            End If
            
        Case 2 ' Tod
            If ChangeReason = "" Then ChangeReason = "Tod des Mitglieds"
            nachpaechterID = ""
            nachpaechterName = ""
            GoTo PruefeMehrfachParzellen
            
        Case 3 ' K�ndigung
            If ChangeReason = "" Then ChangeReason = "K�ndigung"
            nachpaechterID = ""
            nachpaechterName = ""
            GoTo PruefeMehrfachParzellen
            
        ' ENTFERNT: Case 4 ' Parzellenwechsel
            
        Case 5 ' Sonstiges
            If ChangeReason = "" Then ChangeReason = "Sonstiges"
            nachpaechterID = ""
            nachpaechterName = ""
            GoTo PruefeMehrfachParzellen
    End Select
    
    Exit Sub  ' Sicherheits-Exit

' ==========================================================
' NEU v2.8: Pr�fung ob Mitglied mehrere Parzellen hat
' ==========================================================
PruefeMehrfachParzellen:
    Dim alleParzellen As String
    Dim parzellenArray() As String
    Dim anzahlParzellen As Long
    Dim mehrfachAntwort As VbMsgBoxResult
    Dim parzellenOhneVerein As String
    Dim tmpArray() As String
    Dim p As Long
    
    alleParzellen = mod_Mitglieder_Logik.GetParzellenVonMitglied(OldMemberID)
    
    ' Z�hle Parzellen (ohne "Verein")
    parzellenOhneVerein = ""
    If alleParzellen <> "" Then
        tmpArray = Split(alleParzellen, ", ")
        For p = LBound(tmpArray) To UBound(tmpArray)
            If UCase(Trim(tmpArray(p))) <> "VEREIN" Then
                If parzellenOhneVerein = "" Then
                    parzellenOhneVerein = Trim(tmpArray(p))
                Else
                    parzellenOhneVerein = parzellenOhneVerein & ", " & Trim(tmpArray(p))
                End If
            End If
        Next p
    End If
    
    If parzellenOhneVerein <> "" Then
        parzellenArray = Split(parzellenOhneVerein, ", ")
        anzahlParzellen = UBound(parzellenArray) - LBound(parzellenArray) + 1
    Else
        anzahlParzellen = 0
    End If
    
    If anzahlParzellen > 1 Then
        ' *** MITGLIED HAT MEHRERE PARZELLEN ***
        mehrfachAntwort = MsgBox( _
            "HINWEIS: " & vorname & " " & nachname & " hat " & anzahlParzellen & " Parzellen:" & vbCrLf & _
            parzellenOhneVerein & vbCrLf & vbCrLf & _
            "Aktuell wird der Austritt f" & ChrW(252) & "r Parzelle " & OldParzelle & " bearbeitet." & vbCrLf & _
            "(Grund: " & ChangeReason & ")" & vbCrLf & vbCrLf & _
            "M" & ChrW(246) & "chten Sie:" & vbCrLf & _
            "JA = Komplett austreten (ALLE " & anzahlParzellen & " Parzellen abgeben)" & vbCrLf & _
            "NEIN = Nur Parzelle " & OldParzelle & " abgeben (auf den anderen Parzellen bleiben)" & vbCrLf & _
            "ABBRECHEN = Vorgang abbrechen", _
            vbYesNoCancel + vbQuestion, "Mehrere Parzellen erkannt")
        
        If mehrfachAntwort = vbCancel Then
            ' Abbruch
            Me.tag = lRow
            Exit Sub
            
        ElseIf mehrfachAntwort = vbYes Then
            ' *** KOMPLETT-AUSTRITT: Alle Parzellen abgeben ***
            GoTo AustrittBearbeitenKomplett
            
        Else ' vbNo
            ' *** NUR DIESE PARZELLE: Nur die aktuelle Parzelle abgeben ***
            GoTo AustrittBearbeiten
        End If
    Else
        ' *** NUR EINE PARZELLE: Normaler Austritt ***
        GoTo AustrittBearbeiten
    End If

' ==========================================================
' NEU v2.8: Komplett-Austritt bei mehreren Parzellen
' ==========================================================
AustrittBearbeitenKomplett:
    If pachtEndeVal = "" Then
        ' Pachtende ist noch leer - Benutzer kann es eintragen
        Call SetMode(True, False, True)
        
        ' Tag-Format: lRow|Grund|NachpaechterID|NachpaechterName|KOMPLETT
        Me.tag = lRow & "|" & ChangeReason & "|" & nachpaechterID & "|" & nachpaechterName & "|KOMPLETT"
        
        ' F�lle Pachtende mit heutigem Datum und MARKIERE ES komplett
        Me.txt_Pachtende.value = Format(Date, "dd.mm.yyyy")
        Me.txt_Pachtende.SetFocus
        Me.txt_Pachtende.SelStart = 0
        Me.txt_Pachtende.SelLength = Len(Me.txt_Pachtende.value)
        
        MsgBox "KOMPLETT-AUSTRITT: Das Austrittsdatum wurde auf heute gesetzt." & vbCrLf & _
               "Grund: " & ChangeReason & vbCrLf & vbCrLf & _
               "ALLE Parzellen (" & parzellenOhneVerein & ") werden abgegeben!" & vbCrLf & vbCrLf & _
               "Bitte best" & ChrW(228) & "tigen Sie das Datum und klicken Sie dann '" & ChrW(220) & "bernehmen'.", vbInformation, "Komplett-Austritt"
        Exit Sub
    Else
        ' Pachtende ist bereits gesetzt
        austrittsDatum = CDate(pachtEndeVal)
    End If
    
    ' F�hre Komplett-Austritt sofort durch
    Call mod_Mitglieder_Logik.VerschiebeAlleParzellenInHistorie(OldMemberID, nachname, vorname, austrittsDatum, ChangeReason)
    
    ' Formatierung neu anwenden
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
    If mod_Mitglieder_Logik.IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    Unload Me
    Exit Sub
    
AustrittBearbeiten:
    If pachtEndeVal = "" Then
        ' Pachtende ist noch leer - Benutzer kann es eintragen
        ' FIX: IsRemovalMode = True -> nur txt_Pachtende sichtbar!
        Call SetMode(True, False, True)
        
        ' Speichere Grund tempor�r im Tag des Formulars (ohne KOMPLETT-Flag)
        Me.tag = lRow & "|" & ChangeReason & "|" & nachpaechterID & "|" & nachpaechterName
        
        ' F�lle Pachtende mit heutigem Datum und MARKIERE ES komplett
        Me.txt_Pachtende.value = Format(Date, "dd.mm.yyyy")
        Me.txt_Pachtende.SetFocus
        Me.txt_Pachtende.SelStart = 0
        Me.txt_Pachtende.SelLength = Len(Me.txt_Pachtende.value)
        
        MsgBox "Das Austrittsdatum wurde auf heute gesetzt." & vbCrLf & _
               "Grund: " & ChangeReason & vbCrLf & vbCrLf & _
               "Nur Parzelle " & OldParzelle & " wird abgegeben." & vbCrLf & vbCrLf & _
               "Bitte best" & ChrW(228) & "tigen Sie es (oder " & ChrW(228) & "ndern Sie es) und klicken Sie dann '" & ChrW(220) & "bernehmen'.", vbInformation, "Austrittsdatum"
        Exit Sub
    Else
        ' Pachtende ist bereits gesetzt - Mitglied in Historie verschieben
        austrittsDatum = CDate(pachtEndeVal)
    End If
    
    ' Verschiebe NUR DIESE PARZELLE in Mitgliederhistorie
    Call mod_Mitglieder_Logik.VerschiebeInHistorie(lRow, OldParzelle, OldMemberID, nachname, vorname, austrittsDatum, ChangeReason, nachpaechterName, nachpaechterID)
    
    ' Formatierung neu anwenden
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
    If mod_Mitglieder_Logik.IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    Unload Me
    Exit Sub
    
TagError:
    MsgBox "Fehler beim Lesen der Zeilennummer: " & Err.Description, vbCritical
    Exit Sub
End Sub


' ***************************************************************
' HILFSPROZEDUR: BearbeiteNachpaechterUebernahme
' Behandelt die �bernahme einer Parzelle durch einen registrierten Nachp�chter
' ***************************************************************
Private Sub BearbeiteNachpaechterUebernahme(ByVal nachpaechterID As String, ByVal nachpaechterName As String, _
                                             ByVal neueParzelle As String, ByVal alteLRow As Long, _
                                             ByVal alteMemberID As String, ByVal alteNachname As String, _
                                             ByVal alteVorname As String, ByVal austrittsDatum As Date, _
                                             ByVal grund As String)
    
    Dim wsM As Worksheet
    Dim alteParzellen As String
    Dim antwort As VbMsgBoxResult
    Dim r As Long
    Dim lastRow As Long
    Dim nachpaechterParzelle As String
    Dim nachpaechterRow As Long
    
    Set wsM = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    
    ' Finde alle Parzellen des Nachp�chters
    alteParzellen = mod_Mitglieder_Logik.GetParzellenVonMitglied(nachpaechterID)
    
    If alteParzellen = "" Then
        ' Nachp�chter hat keine Parzelle - einfach neue Parzelle zuweisen
        Call UebernehmeParzelleOhneWechsel(nachpaechterID, nachpaechterName, neueParzelle, alteLRow, alteMemberID, alteNachname, alteVorname, austrittsDatum, grund)
    Else
        ' Nachp�chter hat bereits Parzelle(n) - Benutzer fragen
        antwort = MsgBox("Der Nachp�chter " & nachpaechterName & " ist bereits auf Parzelle " & alteParzellen & " gemeldet." & vbCrLf & vbCrLf & _
                        "M�chten Sie:" & vbCrLf & _
                        "JA = Parzelle " & alteParzellen & " verlassen und zu Parzelle " & neueParzelle & " wechseln" & vbCrLf & _
                        "NEIN = Beide Parzellen (" & alteParzellen & " und " & neueParzelle & ") behalten" & vbCrLf & _
                        "ABBRECHEN = Vorgang abbrechen", _
                        vbYesNoCancel + vbQuestion, "Nachp�chter bereits registriert")
        
        If antwort = vbYes Then
            ' Parzelle wechseln - pr�fe ob alte Parzelle noch zahlende Mitglieder hat
            ' Bei mehreren Parzellen: Pr�fe jede einzeln
            Dim parzellenArray() As String
            parzellenArray = Split(alteParzellen, ", ")
            
            Dim kannWechseln As Boolean
            kannWechseln = True
            Dim problematischeParzelle As String
            
            Dim i As Integer
            For i = LBound(parzellenArray) To UBound(parzellenArray)
                If Not mod_Mitglieder_Logik.HatParzelleNochZahlendesMitglied(parzellenArray(i), nachpaechterID) Then
                    kannWechseln = False
                    problematischeParzelle = parzellenArray(i)
                    Exit For
                End If
            Next i
            
            If Not kannWechseln Then
                MsgBox "Der Wechsel ist nicht m�glich!" & vbCrLf & vbCrLf & _
                       "Sie sind das einzige zahlende Mitglied auf Parzelle " & problematischeParzelle & "." & vbCrLf & _
                       "Ein Wechsel w�rde die Parzelle ohne zahlendes Mitglied zur�cklassen.", vbCritical, "Wechsel nicht m�glich"
                Exit Sub
            End If
            
            ' Wechsel durchf�hren - alle alten Eintr�ge in Historie verschieben
            Call NachpaechterParzellenWechsel(nachpaechterID, nachpaechterName, neueParzelle, austrittsDatum, alteLRow, alteMemberID, alteNachname, alteVorname, grund)
            
        ElseIf antwort = vbNo Then
            ' Pr�fe ob Nachp�chter bereits auf der NEUEN Parzelle ist (Doppel-Check!)
            If mod_Mitglieder_Logik.ExistiertBereitsAufParzelle(nachpaechterID, neueParzelle) Then
                MsgBox "FEHLER: " & nachpaechterName & " ist bereits auf Parzelle " & neueParzelle & " registriert!" & vbCrLf & _
                       "Doppelte Eintr�ge sind nicht erlaubt.", vbCritical, "Doppelter Eintrag verhindert"
                Exit Sub
            End If
            
            ' Beide Parzellen behalten - neue Zeile hinzuf�gen
            Call NachpaechterZusaetzlicheParzelle(nachpaechterID, nachpaechterName, neueParzelle, austrittsDatum, alteLRow, alteMemberID, alteNachname, alteVorname, grund)
            
        Else
            ' Abbrechen
            Exit Sub
        End If
    End If
    
End Sub

' ***************************************************************
' HILFSPROZEDUR: UebernehmeParzelleOhneWechsel
' Nachp�chter ohne bestehende Parzelle �bernimmt neue Parzelle
' ***************************************************************
Private Sub UebernehmeParzelleOhneWechsel(ByVal nachpaechterID As String, ByVal nachpaechterName As String, _
                                           ByVal neueParzelle As String, ByVal alteLRow As Long, _
                                           ByVal alteMemberID As String, ByVal alteNachname As String, _
                                           ByVal alteVorname As String, ByVal austrittsDatum As Date, _
                                           ByVal grund As String)
    
    Dim wsM As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim nachpaechterRow As Long
    Dim nachpaechterPachtbeginn As String
    
    Set wsM = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    wsM.Unprotect PASSWORD:=PASSWORD
    
    ' Finde Zeile des Nachp�chters
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    nachpaechterRow = 0
    
    For r = M_START_ROW To lastRow
        If wsM.Cells(r, M_COL_MEMBER_ID).value = nachpaechterID Then
            nachpaechterRow = r
            nachpaechterPachtbeginn = wsM.Cells(r, M_COL_PACHTANFANG).value
            Exit For
        End If
    Next r
    
    If nachpaechterRow > 0 Then
        ' Aktualisiere Parzelle des Nachp�chters
        wsM.Cells(nachpaechterRow, M_COL_PARZELLE).value = neueParzelle
        wsM.Cells(nachpaechterRow, M_COL_SEITE).value = mod_Mitglieder_Logik.GetSeiteFromParzelle(neueParzelle)
    End If
    
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' Verschiebe altes Mitglied in Historie
    Call mod_Mitglieder_Logik.VerschiebeInHistorie(alteLRow, neueParzelle, alteMemberID, alteNachname, alteVorname, austrittsDatum, grund, nachpaechterName, nachpaechterID)
    
    ' Formatierung neu anwenden
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
    If mod_Mitglieder_Logik.IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    MsgBox "Parzelle " & neueParzelle & " wurde an " & nachpaechterName & " �bergeben.", vbInformation
    
    Unload Me
End Sub

' ***************************************************************
' HILFSPROZEDUR: NachpaechterParzellenWechsel
' Nachp�chter verl�sst alte Parzelle(n) komplett und wechselt zur neuen
' ***************************************************************
Private Sub NachpaechterParzellenWechsel(ByVal nachpaechterID As String, ByVal nachpaechterName As String, _
                                          ByVal neueParzelle As String, ByVal austrittsDatum As Date, _
                                          ByVal alteLRow As Long, ByVal alteMemberID As String, _
                                          ByVal alteNachname As String, ByVal alteVorname As String, _
                                          ByVal grund As String)
    
    Dim wsM As Worksheet
    Dim wsH As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim nachpaechterPachtbeginn As String
    Dim alteParzelle As String
    Dim nachpaechterNachname As String
    Dim nachpaechterVorname As String
    Dim nachpaechterAnrede As String
    Dim nachpaechterStrasse As String
    Dim nachpaechterNummer As String
    Dim nachpaechterPLZ As String
    Dim nachpaechterWohnort As String
    Dim nachpaechterTelefon As String
    Dim nachpaechterMobil As String
    Dim nachpaechterGeburtstag As String
    Dim nachpaechterEmail As String
    Dim nachpaechterFunktion As String
    
    Set wsM = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    
    wsM.Unprotect PASSWORD:=PASSWORD
    wsH.Unprotect PASSWORD:=PASSWORD
    
    ' WICHTIG: Sammle ALLE Daten des Nachp�chters VOR dem L�schen!
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        If wsM.Cells(r, M_COL_MEMBER_ID).value = nachpaechterID Then
            ' Speichere alle Daten beim ERSTEN Fund
            If nachpaechterPachtbeginn = "" Then
                nachpaechterPachtbeginn = wsM.Cells(r, M_COL_PACHTANFANG).value
                nachpaechterNachname = wsM.Cells(r, M_COL_NACHNAME).value
                nachpaechterVorname = wsM.Cells(r, M_COL_VORNAME).value
                nachpaechterAnrede = wsM.Cells(r, M_COL_ANREDE).value
                nachpaechterStrasse = wsM.Cells(r, M_COL_STRASSE).value
                nachpaechterNummer = wsM.Cells(r, M_COL_NUMMER).value
                nachpaechterPLZ = wsM.Cells(r, M_COL_PLZ).value
                nachpaechterWohnort = wsM.Cells(r, M_COL_WOHNORT).value
                nachpaechterTelefon = wsM.Cells(r, M_COL_TELEFON).value
                nachpaechterMobil = wsM.Cells(r, M_COL_MOBIL).value
                nachpaechterGeburtstag = wsM.Cells(r, M_COL_GEBURTSTAG).value
                nachpaechterEmail = wsM.Cells(r, M_COL_EMAIL).value
                nachpaechterFunktion = wsM.Cells(r, M_COL_FUNKTION).value
            End If
        End If
    Next r
    
    ' Jetzt l�sche alle Zeilen des Nachp�chters und schreibe in Historie (r�ckw�rts!)
    For r = lastRow To M_START_ROW Step -1
        If wsM.Cells(r, M_COL_MEMBER_ID).value = nachpaechterID Then
            ' Speichere alte Parzelle
            alteParzelle = wsM.Cells(r, M_COL_PARZELLE).value
            
            ' === SICHERHEITSCHECK: NIEMALS Verein-Zeile l�schen ===
            If UCase(Trim(alteParzelle)) = "VEREIN" Then
                ' �berspringe diese Zeile - NICHT L�SCHEN!
                Debug.Print "WARNUNG: Verein-Zeile �bersprungen (Zeile " & r & ")"
                GoTo nextRow
            End If
            
            ' Schreibe in Mitgliederhistorie
            Dim nextHistRow As Long
            nextHistRow = wsH.Cells(wsH.Rows.count, H_COL_NAME_EHEM_PAECHTER).End(xlUp).Row + 1
            If nextHistRow < H_START_ROW Then nextHistRow = H_START_ROW
            
            wsH.Cells(nextHistRow, H_COL_PARZELLE).value = alteParzelle
            wsH.Cells(nextHistRow, H_COL_MEMBER_ID_ALT).value = nachpaechterID
            wsH.Cells(nextHistRow, H_COL_NAME_EHEM_PAECHTER).value = nachpaechterNachname & ", " & nachpaechterVorname
            
            On Error Resume Next
            wsH.Cells(nextHistRow, H_COL_AUST_DATUM).value = austrittsDatum
            If Err.Number = 0 Then
                wsH.Cells(nextHistRow, H_COL_AUST_DATUM).NumberFormat = "dd.mm.yyyy"
            End If
            On Error GoTo 0
            
            wsH.Cells(nextHistRow, H_COL_GRUND).value = "Parzellenwechsel"
            wsH.Cells(nextHistRow, H_COL_NACHPAECHTER_NAME).value = ""
            wsH.Cells(nextHistRow, H_COL_NACHPAECHTER_ID).value = ""
            wsH.Cells(nextHistRow, H_COL_KOMMENTAR).value = ""
            wsH.Cells(nextHistRow, H_COL_ENDABRECHNUNG).value = ""
            
            On Error Resume Next
            wsH.Cells(nextHistRow, H_COL_SYSTEMZEIT).value = Now
            If Err.Number = 0 Then
                wsH.Cells(nextHistRow, H_COL_SYSTEMZEIT).NumberFormat = "dd.mm.yyyy hh:mm:ss"
            End If
            On Error GoTo 0
            
            ' L�sche Zeile
            wsM.Rows(r).Delete Shift:=xlUp
        End If
nextRow:
    Next r
    
    ' Erstelle neue Zeile f�r Nachp�chter auf neuer Parzelle
    Dim newRow As Long
    newRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row + 1
    
    ' Schreibe alle gespeicherten Daten in neue Zeile
    wsM.Cells(newRow, M_COL_MEMBER_ID).value = nachpaechterID
    wsM.Cells(newRow, M_COL_PARZELLE).value = neueParzelle
    wsM.Cells(newRow, M_COL_SEITE).value = mod_Mitglieder_Logik.GetSeiteFromParzelle(neueParzelle)
    wsM.Cells(newRow, M_COL_ANREDE).value = nachpaechterAnrede
    wsM.Cells(newRow, M_COL_NACHNAME).value = nachpaechterNachname
    wsM.Cells(newRow, M_COL_VORNAME).value = nachpaechterVorname
    wsM.Cells(newRow, M_COL_STRASSE).value = nachpaechterStrasse
    wsM.Cells(newRow, M_COL_NUMMER).value = nachpaechterNummer
    wsM.Cells(newRow, M_COL_PLZ).value = nachpaechterPLZ
    wsM.Cells(newRow, M_COL_WOHNORT).value = nachpaechterWohnort
    wsM.Cells(newRow, M_COL_TELEFON).value = nachpaechterTelefon
    wsM.Cells(newRow, M_COL_MOBIL).value = nachpaechterMobil
    wsM.Cells(newRow, M_COL_GEBURTSTAG).value = nachpaechterGeburtstag
    wsM.Cells(newRow, M_COL_EMAIL).value = nachpaechterEmail
    wsM.Cells(newRow, M_COL_FUNKTION).value = nachpaechterFunktion
    
    ' Pachtbeginn nur setzen wenn vorhanden - MIT FEHLERBEHANDLUNG
    If nachpaechterPachtbeginn <> "" Then
        On Error Resume Next
        wsM.Cells(newRow, M_COL_PACHTANFANG).value = CDate(nachpaechterPachtbeginn)
        If Err.Number = 0 Then
            wsM.Cells(newRow, M_COL_PACHTANFANG).NumberFormat = "dd.mm.yyyy"
        End If
        On Error GoTo 0
    End If
    
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    wsH.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' Verschiebe altes Mitglied in Historie (muss neu gesucht werden, da Zeilen verschoben wurden)
    ' Finde die neue lRow des austretenden Mitglieds
    Dim neueAlteLRow As Long
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        If wsM.Cells(r, M_COL_MEMBER_ID).value = alteMemberID And _
           wsM.Cells(r, M_COL_PARZELLE).value = neueParzelle Then
            neueAlteLRow = r
            Exit For
        End If
    Next r
    
    If neueAlteLRow > 0 Then
        Call mod_Mitglieder_Logik.VerschiebeInHistorie(neueAlteLRow, neueParzelle, alteMemberID, alteNachname, alteVorname, austrittsDatum, grund, nachpaechterName, nachpaechterID)
    End If
    
    ' Formatierung neu anwenden
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
    If mod_Mitglieder_Logik.IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    MsgBox "Nachp�chter " & nachpaechterName & " ist von allen bisherigen Parzellen zu Parzelle " & neueParzelle & " gewechselt.", vbInformation
    
    Unload Me
End Sub

' ***************************************************************
' HILFSPROZEDUR: NachpaechterZusaetzlicheParzelle
' Nachp�chter beh�lt alte Parzelle und bekommt zus�tzlich neue
' ***************************************************************
Private Sub NachpaechterZusaetzlicheParzelle(ByVal nachpaechterID As String, ByVal nachpaechterName As String, _
                                              ByVal neueParzelle As String, ByVal austrittsDatum As Date, _
                                              ByVal alteLRow As Long, ByVal alteMemberID As String, _
                                              ByVal alteNachname As String, ByVal alteVorname As String, _
                                              ByVal grund As String)
    
    Dim wsM As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim newRow As Long
    Dim vorlagenRow As Long
    
    Set wsM = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    
    ' === SICHERHEITSCHECK: Pr�fe ob bereits auf dieser Parzelle ===
    If mod_Mitglieder_Logik.ExistiertBereitsAufParzelle(nachpaechterID, neueParzelle) Then
        MsgBox "FEHLER: " & nachpaechterName & " ist bereits auf Parzelle " & neueParzelle & " registriert!" & vbCrLf & _
               "Doppelte Eintr�ge sind nicht erlaubt.", vbCritical, "Doppelter Eintrag verhindert"
        Exit Sub
    End If
    
    wsM.Unprotect PASSWORD:=PASSWORD
    
    ' Finde eine Zeile des Nachp�chters als Vorlage
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    vorlagenRow = 0
    
    For r = M_START_ROW To lastRow
        If wsM.Cells(r, M_COL_MEMBER_ID).value = nachpaechterID Then
            vorlagenRow = r
            Exit For
        End If
    Next r
    
    If vorlagenRow = 0 Then
        MsgBox "Fehler: Nachp�chter nicht gefunden.", vbCritical
        wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        Exit Sub
    End If
    
    ' Erstelle neue Zeile f�r zus�tzliche Parzelle
    newRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row + 1
    
    ' Kopiere alle Daten von Vorlagenzeile
    wsM.Cells(newRow, M_COL_MEMBER_ID).value = wsM.Cells(vorlagenRow, M_COL_MEMBER_ID).value
    wsM.Cells(newRow, M_COL_PARZELLE).value = neueParzelle
    wsM.Cells(newRow, M_COL_SEITE).value = mod_Mitglieder_Logik.GetSeiteFromParzelle(neueParzelle)
    wsM.Cells(newRow, M_COL_ANREDE).value = wsM.Cells(vorlagenRow, M_COL_ANREDE).value
    wsM.Cells(newRow, M_COL_NACHNAME).value = wsM.Cells(vorlagenRow, M_COL_NACHNAME).value
    wsM.Cells(newRow, M_COL_VORNAME).value = wsM.Cells(vorlagenRow, M_COL_VORNAME).value
    wsM.Cells(newRow, M_COL_STRASSE).value = wsM.Cells(vorlagenRow, M_COL_STRASSE).value
    wsM.Cells(newRow, M_COL_NUMMER).value = wsM.Cells(vorlagenRow, M_COL_NUMMER).value
    wsM.Cells(newRow, M_COL_PLZ).value = wsM.Cells(vorlagenRow, M_COL_PLZ).value
    wsM.Cells(newRow, M_COL_WOHNORT).value = wsM.Cells(vorlagenRow, M_COL_WOHNORT).value
    wsM.Cells(newRow, M_COL_TELEFON).value = wsM.Cells(vorlagenRow, M_COL_TELEFON).value
    wsM.Cells(newRow, M_COL_MOBIL).value = wsM.Cells(vorlagenRow, M_COL_MOBIL).value
    wsM.Cells(newRow, M_COL_GEBURTSTAG).value = wsM.Cells(vorlagenRow, M_COL_GEBURTSTAG).value
    wsM.Cells(newRow, M_COL_EMAIL).value = wsM.Cells(vorlagenRow, M_COL_EMAIL).value
    wsM.Cells(newRow, M_COL_FUNKTION).value = wsM.Cells(vorlagenRow, M_COL_FUNKTION).value
    
    ' Pachtbeginn = �bernahmedatum (AustrittsDatum) - MIT FEHLERBEHANDLUNG
    On Error Resume Next
    wsM.Cells(newRow, M_COL_PACHTANFANG).value = austrittsDatum
    If Err.Number = 0 Then
        wsM.Cells(newRow, M_COL_PACHTANFANG).NumberFormat = "dd.mm.yyyy"
    End If
    On Error GoTo 0
    
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' Verschiebe altes Mitglied in Historie
    Call mod_Mitglieder_Logik.VerschiebeInHistorie(alteLRow, neueParzelle, alteMemberID, alteNachname, alteVorname, austrittsDatum, grund, nachpaechterName, nachpaechterID)
    
    ' Formatierung neu anwenden
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
    If mod_Mitglieder_Logik.IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    MsgBox "Nachp�chter " & nachpaechterName & " hat zus�tzlich Parzelle " & neueParzelle & " �bernommen.", vbInformation
    
    Unload Me
End Sub

' ***************************************************************
' HILFSPROZEDUR: VerarbeiteAustrittNachNachpaechterErfassung
' Wird aufgerufen nachdem ein neuer Nachp�chter erfasst wurde
' ***************************************************************
Private Sub VerarbeiteAustrittNachNachpaechterErfassung(ByVal lRow As Long, ByVal parzelle As String, _
                                                          ByVal memberID As String, ByVal nachname As String, _
                                                          ByVal vorname As String, ByVal austrittsDatum As Date, _
                                                          ByVal grund As String)
    
    Dim wsM As Worksheet
    Dim newMemberID As String
    Dim newMemberName As String
    Dim r As Long
    Dim lastRow As Long
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    ' Finde den neu angelegten Nachp�chter (letzte Zeile mit gleicher Parzelle)
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    For r = lastRow To M_START_ROW Step -1
        If StrComp(Trim(wsM.Cells(r, M_COL_PARZELLE).value), parzelle, vbTextCompare) = 0 Then
            ' Pr�fe ob es nicht das alte Mitglied ist
            If r <> lRow Then
                newMemberID = wsM.Cells(r, M_COL_MEMBER_ID).value
                newMemberName = wsM.Cells(r, M_COL_NACHNAME).value & ", " & wsM.Cells(r, M_COL_VORNAME).value
                Exit For
            End If
        End If
    Next r
    
    ' Verschiebe altes Mitglied in Historie mit Nachp�chter-Daten
    Call mod_Mitglieder_Logik.VerschiebeInHistorie(lRow, parzelle, memberID, nachname, vorname, austrittsDatum, grund, newMemberName, newMemberID)
    
    ' Formatierung neu anwenden
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
    If mod_Mitglieder_Logik.IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    Unload Me
End Sub


' ***************************************************************
' NEUE VERSION: cmd_Uebernehmen_Click mit Parzellenwechsel-Logik
' ***************************************************************
Private Sub cmd_Uebernehmen_Click()
    
    Dim tagParts() As String
    Dim lRow As Long
    Dim grund As String
    Dim nachpaechterID As String
    Dim nachpaechterName As String
    
    ' Pr�fe ob Tag im Format "lRow|Grund|NachpaechterID|NachpaechterName[|KOMPLETT]" vorliegt
    If InStr(Me.tag, "|") > 0 Then
        tagParts = Split(Me.tag, "|")
        
        ' Pr�fe ob erstes Element numerisch ist
        If mod_Mitglieder_Logik.IsNumericTag(tagParts(0)) And UBound(tagParts) >= 1 Then
            ' Austritt-Modus mit Grund
            lRow = CLng(tagParts(0))
            grund = tagParts(1)
            If UBound(tagParts) >= 2 Then nachpaechterID = tagParts(2)
            If UBound(tagParts) >= 3 Then nachpaechterName = tagParts(3)
            
            ' NEU v2.8: Pr�fe auf KOMPLETT-Flag
            Dim istKomplettAustritt As Boolean
            istKomplettAustritt = False
            If UBound(tagParts) >= 4 Then
                If UCase(tagParts(4)) = "KOMPLETT" Then
                    istKomplettAustritt = True
                End If
            End If
            
            If istKomplettAustritt Then
                ' KOMPLETT-AUSTRITT: Alle Parzellen des Mitglieds verschieben
                Call cmd_Uebernehmen_MitKomplettAustritt(lRow, grund)
            Else
                ' Normaler Einzel-Parzellen-Austritt
                Call cmd_Uebernehmen_MitAustritt(lRow, grund, nachpaechterName, nachpaechterID)
            End If
            Exit Sub
        End If
    End If
    
    On Error GoTo 0
    
    Dim wsM As Worksheet
    Dim autoSeite As String
    Dim funktion As String
    Dim istMitgliedOhnePacht As Boolean
    Dim OldParzelle As String
    Dim NewParzelle As String
    Dim nachname As String
    Dim vorname As String
    Dim currentMemberID As String
    Dim antwort As VbMsgBoxResult
    Dim zielParzelleHatMitglied As Boolean
    Dim istWechsel As Boolean
    Dim mitgliedNameAufZiel As String
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    
    ' === PFLICHTFELDER VALIDIERUNG ===
    If Trim(Me.txt_Nachname.value) = "" Or Trim(Me.txt_Vorname.value) = "" Then
        MsgBox "Nachname und Vorname d�rfen nicht leer sein.", vbCritical
        Exit Sub
    End If
    
    ' === DATUMSVALIDIERUNG ===
    If Not mod_Mitglieder_Logik.IstGueltigesDatum(Me.txt_Geburtstag.value) Then
        MsgBox "Geburtstag: Bitte ein g�ltiges Datum eingeben (Format: TT.MM.JJJJ).", vbExclamation
        Exit Sub
    End If
    
    If Not mod_Mitglieder_Logik.IstGueltigesDatum(Me.txt_Pachtbeginn.value) Then
        MsgBox "Pachtbeginn: Bitte ein g�ltiges Datum eingeben (Format: TT.MM.JJJJ).", vbExclamation
        Exit Sub
    End If
    
    If Not mod_Mitglieder_Logik.IstGueltigesDatum(Me.txt_Pachtende.value) Then
        MsgBox "Pachtende: Bitte ein g�ltiges Datum eingeben (Format: TT.MM.JJJJ).", vbExclamation
        Exit Sub
    End If
    
    funktion = Me.cbo_Funktion.value
    istMitgliedOhnePacht = (funktion = "Mitglied ohne Pacht")
    
    OldParzelle = Me.lbl_Parzelle.Caption
    NewParzelle = Me.cbo_Parzelle.value
    nachname = Me.txt_Nachname.value
    vorname = Me.txt_Vorname.value
    currentMemberID = wsM.Cells(lRow, M_COL_MEMBER_ID).value
    
    ' === SICHERHEITSCHECK: Verein-Parzelle darf nicht bearbeitet werden ===
    If UCase(Trim(OldParzelle)) = "VEREIN" Then
        MsgBox "FEHLER: Die Verein-Parzelle darf nicht bearbeitet werden!", vbCritical
        Exit Sub
    End If
    
    ' === VALIDIERUNG: "Mitglied ohne Pacht" darf keine leere Parzelle beziehen ===
    If istMitgliedOhnePacht Then
        If NewParzelle <> "" And mod_Mitglieder_Logik.IstParzelleLeer(NewParzelle) Then
            ' Pr�fe ob es ein Wechsel von "Mitglied mit Pacht" zu "Mitglied ohne Pacht" ist
            Dim alteFunktion As String
            alteFunktion = wsM.Cells(lRow, M_COL_FUNKTION).value
            
            If alteFunktion <> "Mitglied ohne Pacht" Then
                ' Wechsel von zahlendem Mitglied zu "ohne Pacht"
                ' Pr�fe ob andere zahlende Mitglieder auf der Parzelle sind
                If Not mod_Mitglieder_Logik.HatParzelleNochZahlendesMitglied(NewParzelle, currentMemberID) Then
                    MsgBox "FEHLER: Ein Mitglied ohne Pacht kann nicht das einzige Mitglied auf einer Parzelle sein!" & vbCrLf & vbCrLf & _
                           "Es muss immer ein zahlendes Mitglied (Mitglied mit Pacht oder Vorstandsmitglied) auf der Parzelle sein.", _
                           vbCritical, "Validierungsfehler"
                    Exit Sub
                End If
            Else
                ' War schon "Mitglied ohne Pacht" und will auf leere Parzelle
                MsgBox "FEHLER: Ein Mitglied ohne Pacht kann keine leere Parzelle beziehen!" & vbCrLf & vbCrLf & _
                       "Die Parzelle " & NewParzelle & " hat kein zahlendes Mitglied.", _
                       vbCritical, "Validierungsfehler"
                Exit Sub
            End If
        End If
    End If
    
    ' === VALIDIERUNG: Duplikate (gleicher Vor- und Nachname auf Parzelle) ===
    If mod_Mitglieder_Logik.ExistiertPersonAufParzelle(vorname, nachname, NewParzelle, lRow) Then
        MsgBox "FEHLER: Eine Person mit dem Namen " & nachname & ", " & vorname & _
               " ist bereits auf Parzelle " & NewParzelle & " registriert!" & vbCrLf & vbCrLf & _
               "Doppelte Eintr�ge sind nicht erlaubt.", vbCritical, "Doppelter Eintrag verhindert"
        Exit Sub
    End If
    
    ' --- VALIDIERUNG: Pachtbeginn je nach Funktion ---
    If Not istMitgliedOhnePacht Then
        ' Mit Pacht: Pachtbeginn ist mandatory
        If Me.txt_Pachtbeginn.value = "" Then
            MsgBox "F�r diese Funktion ist ein Pachtbeginn erforderlich.", vbCritical
            Exit Sub
        End If
    End If
    
    ' === PARZELLENWECHSEL-LOGIK ===
    If OldParzelle <> NewParzelle And OldParzelle <> "" And NewParzelle <> "" Then
        ' Parzellenwechsel erkannt!
        zielParzelleHatMitglied = Not mod_Mitglieder_Logik.IstParzelleLeer(NewParzelle)
        
        If zielParzelleHatMitglied Then
            ' Zielparzelle hat bereits Mitglieder
            mitgliedNameAufZiel = mod_Mitglieder_Logik.GetMitgliedNameAufParzelle(NewParzelle)
            
            antwort = MsgBox("Die Parzelle " & NewParzelle & " hat bereits ein Mitglied (" & mitgliedNameAufZiel & ")." & vbCrLf & vbCrLf & _
                           "M�chten Sie:" & vbCrLf & _
                           "JA = Parzelle " & NewParzelle & " zus�tzlich pachten (beide Parzellen behalten)" & vbCrLf & _
                           "NEIN = Parzelle " & OldParzelle & " verlassen und zu " & NewParzelle & " wechseln (Umzug)" & vbCrLf & _
                           "ABBRECHEN = Vorgang abbrechen", _
                           vbYesNoCancel + vbQuestion, "Parzellenwechsel")
        Else
            ' Zielparzelle ist leer
            antwort = MsgBox("Die Parzelle " & NewParzelle & " ist leer." & vbCrLf & vbCrLf & _
                           "M�chten Sie:" & vbCrLf & _
                           "JA = Parzelle " & NewParzelle & " zus�tzlich pachten (beide Parzellen behalten)" & vbCrLf & _
                           "NEIN = Parzelle " & OldParzelle & " verlassen und zu " & NewParzelle & " wechseln (Umzug)" & vbCrLf & _
                           "ABBRECHEN = Vorgang abbrechen", _
                           vbYesNoCancel + vbQuestion, "Parzellenwechsel")
        End If
        
        If antwort = vbCancel Then
            Exit Sub
        End If
        
        ' GE�NDERT: JA = Zus�tzliche Parzelle, NEIN = Wechsel
        istWechsel = (antwort = vbNo)
        
        If istWechsel Then
            ' === UMZUG: Alte Parzelle verlassen ===
            
            ' PR�FUNG 1: Ist die neue Parzelle leer UND ist das Mitglied KEIN zahlendes Mitglied?
            If mod_Mitglieder_Logik.IstParzelleLeer(NewParzelle) Then
                If Not (funktion = "Mitglied mit Pacht" Or _
                        funktion = "1. Vorsitzende(r)" Or _
                        funktion = "2. Vorsitzende(r)" Or _
                        funktion = "Kassierer(in)" Or _
                        funktion = "Schriftf�hrer(in)") Then
                    MsgBox "FEHLER: Ein 'Mitglied ohne Pacht' kann nicht alleine auf eine leere Parzelle wechseln!" & vbCrLf & vbCrLf & _
                           "Die Parzelle " & NewParzelle & " ist leer und ben�tigt ein zahlendes Mitglied " & _
                           "(Mitglied mit Pacht oder Vorstandsmitglied).", vbCritical, "Wechsel nicht m�glich"
                    Exit Sub
                End If
            End If
            
            ' PR�FUNG 2: Pr�fe ob auf alter Parzelle noch zahlende Mitglieder bleiben
            If Not mod_Mitglieder_Logik.HatParzelleNochZahlendesMitglied(OldParzelle, currentMemberID) Then
                Dim warnAntwort As VbMsgBoxResult
                warnAntwort = MsgBox("WARNUNG: Sie sind das einzige zahlende Mitglied auf Parzelle " & OldParzelle & "!" & vbCrLf & vbCrLf & _
                               "Nach Ihrem Wechsel wird die Parzelle ohne zahlendes Mitglied sein." & vbCrLf & vbCrLf & _
                               "M�chten Sie trotzdem wechseln?", vbYesNo + vbExclamation, "Warnung")
                If warnAntwort = vbNo Then
                    Exit Sub
                End If
            End If
            
            ' Speichere �nderungen in Mitgliederliste (neue Parzelle)
            Call SpeichereMitgliedsdaten(wsM, lRow, NewParzelle)
            
            ' Speichere Parzellenwechsel in Historie (Member ID bleibt erhalten!)
            Call mod_Mitglieder_Logik.SpeichereParzellenwechselInHistorie(OldParzelle, NewParzelle, currentMemberID, nachname, vorname, "Parzellenwechsel (Umzug)")
            
        Else
            ' === ZUS�TZLICHE PARZELLE: Neue Zeile anlegen (JA wurde gedr�ckt) ===
            ' WICHTIG: Die bestehende Zeile (OldParzelle) wird NICHT ge�ndert!
            
            ' Pr�fe ob Mitglied bereits auf der neuen Parzelle existiert (Duplikat-Check)
            If mod_Mitglieder_Logik.ExistiertBereitsAufParzelle(currentMemberID, NewParzelle, 0) Then
                MsgBox "FEHLER: Sie sind bereits auf Parzelle " & NewParzelle & " registriert!" & vbCrLf & _
                       "Doppelte Eintr�ge sind nicht erlaubt.", vbCritical, "Doppelter Eintrag verhindert"
                Exit Sub
            End If
            
            ' Erstelle nur die neue Zeile f�r die zus�tzliche Parzelle
            Call ErstelleZusaetzlicheParzelleZeile(wsM, lRow, NewParzelle, currentMemberID)
            
            ' Speichere in Historie
            Call mod_Mitglieder_Logik.SpeichereParzellenwechselInHistorie(OldParzelle, NewParzelle, currentMemberID, nachname, vorname, "Zus�tzliche Parzelle gepachtet")
        End If
        
    Else
        ' === NORMALE �NDERUNG (kein Parzellenwechsel) ===
        Call SpeichereMitgliedsdaten(wsM, lRow, NewParzelle)
        
        ' Normale �nderung - nur Sortierung und Formatierung
        Call mod_Mitglieder_UI.Sortiere_Mitgliederliste_Nach_Parzelle
        Call mod_Mitglieder_UI.Fuelle_MemberIDs_Wenn_Fehlend
    End If
    
    ' Formatierung neu anwenden
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
    If mod_Mitglieder_Logik.IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    MsgBox "�nderungen f�r Mitglied " & nachname & " erfolgreich gespeichert.", vbInformation
    
    Unload Me
    Exit Sub
    
TagError:
    MsgBox "Fehler beim Lesen der Zeilennummer: " & Err.Description, vbCritical
    Exit Sub
    
ErrorHandler:
    On Error GoTo 0
    If Not wsM Is Nothing Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler beim Speichern der �nderungen: " & Err.Description, vbCritical
End Sub

' ***************************************************************
' HILFSPROZEDUR: Speichert Mitgliedsdaten in Worksheet
' ***************************************************************
Private Sub SpeichereMitgliedsdaten(ByRef wsM As Worksheet, ByVal lRow As Long, ByVal parzelle As String)
    Dim autoSeite As String
    
    wsM.Unprotect PASSWORD:=PASSWORD
    
    On Error Resume Next
    
    autoSeite = mod_Mitglieder_Logik.GetSeiteFromParzelle(parzelle)
    
    wsM.Cells(lRow, M_COL_PARZELLE).value = parzelle
    wsM.Cells(lRow, M_COL_SEITE).value = autoSeite
    wsM.Cells(lRow, M_COL_ANREDE).value = Me.cbo_Anrede.value
    wsM.Cells(lRow, M_COL_NACHNAME).value = Me.txt_Nachname.value
    wsM.Cells(lRow, M_COL_VORNAME).value = Me.txt_Vorname.value
    wsM.Cells(lRow, M_COL_STRASSE).value = Me.txt_Strasse.value
    wsM.Cells(lRow, M_COL_NUMMER).value = Me.txt_Nummer.value
    wsM.Cells(lRow, M_COL_PLZ).value = Me.txt_PLZ.value
    wsM.Cells(lRow, M_COL_WOHNORT).value = Me.txt_Wohnort.value
    wsM.Cells(lRow, M_COL_TELEFON).value = Me.txt_Telefon.value
    wsM.Cells(lRow, M_COL_MOBIL).value = Me.txt_Mobil.value
    wsM.Cells(lRow, M_COL_GEBURTSTAG).value = Me.txt_Geburtstag.value
    wsM.Cells(lRow, M_COL_EMAIL).value = Me.txt_Email.value
    wsM.Cells(lRow, M_COL_FUNKTION).value = Me.cbo_Funktion.value
    
    ' Pachtbeginn mit Fehlerbehandlung
    If Me.txt_Pachtbeginn.value <> "" Then
        wsM.Cells(lRow, M_COL_PACHTANFANG).value = CDate(Me.txt_Pachtbeginn.value)
        If Err.Number = 0 Then
            wsM.Cells(lRow, M_COL_PACHTANFANG).NumberFormat = "dd.mm.yyyy"
        End If
        Err.Clear
    End If
    
    ' Pachtende mit Fehlerbehandlung
    If Me.txt_Pachtende.value <> "" Then
        wsM.Cells(lRow, M_COL_PACHTENDE).value = CDate(Me.txt_Pachtende.value)
        If Err.Number = 0 Then
            wsM.Cells(lRow, M_COL_PACHTENDE).NumberFormat = "dd.mm.yyyy"
        End If
        Err.Clear
    End If
    
    On Error GoTo 0
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
End Sub

' ***************************************************************
' HILFSPROZEDUR: Erstellt neue Zeile f�r zus�tzliche Parzelle
' Member ID wird beibehalten!
' ***************************************************************
Private Sub ErstelleZusaetzlicheParzelleZeile(ByRef wsM As Worksheet, ByVal vorlagenRow As Long, _
                                               ByVal neueParzelle As String, ByVal memberID As String)
    Dim newRow As Long
    Dim autoSeite As String
    
    wsM.Unprotect PASSWORD:=PASSWORD
    
    newRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row + 1
    autoSeite = mod_Mitglieder_Logik.GetSeiteFromParzelle(neueParzelle)
    
    ' Kopiere alle Daten von Vorlagenzeile mit GLEICHER Member ID
    wsM.Cells(newRow, M_COL_MEMBER_ID).value = memberID  ' WICHTIG: Gleiche Member ID!
    wsM.Cells(newRow, M_COL_PARZELLE).value = neueParzelle
    wsM.Cells(newRow, M_COL_SEITE).value = autoSeite
    wsM.Cells(newRow, M_COL_ANREDE).value = Me.cbo_Anrede.value
    wsM.Cells(newRow, M_COL_NACHNAME).value = Me.txt_Nachname.value
    wsM.Cells(newRow, M_COL_VORNAME).value = Me.txt_Vorname.value
    wsM.Cells(newRow, M_COL_STRASSE).value = Me.txt_Strasse.value
    wsM.Cells(newRow, M_COL_NUMMER).value = Me.txt_Nummer.value
    wsM.Cells(newRow, M_COL_PLZ).value = Me.txt_PLZ.value
    wsM.Cells(newRow, M_COL_WOHNORT).value = Me.txt_Wohnort.value
    wsM.Cells(newRow, M_COL_TELEFON).value = Me.txt_Telefon.value
    wsM.Cells(newRow, M_COL_MOBIL).value = Me.txt_Mobil.value
    wsM.Cells(newRow, M_COL_GEBURTSTAG).value = Me.txt_Geburtstag.value
    wsM.Cells(newRow, M_COL_EMAIL).value = Me.txt_Email.value
    wsM.Cells(newRow, M_COL_FUNKTION).value = Me.cbo_Funktion.value
    
    ' Pachtbeginn = heute (�bernahmedatum)
    On Error Resume Next
    wsM.Cells(newRow, M_COL_PACHTANFANG).value = Date
    wsM.Cells(newRow, M_COL_PACHTANFANG).NumberFormat = "dd.mm.yyyy"
    On Error GoTo 0
    
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' Sortiere und formatiere
    Call mod_Mitglieder_UI.Sortiere_Mitgliederliste_Nach_Parzelle
End Sub


' ***************************************************************
' HILFSPROZEDUR: cmd_Uebernehmen_MitAustritt
' Wird aufgerufen wenn Austritt mit Grund durchgef�hrt wird
' ***************************************************************
Private Sub cmd_Uebernehmen_MitAustritt(ByVal lRow As Long, ByVal grund As String, _
                                         Optional ByVal nachpaechterName As String = "", _
                                         Optional ByVal nachpaechterID As String = "")
    
    Dim wsM As Worksheet
    Dim nachname As String
    Dim vorname As String
    Dim OldParzelle As String
    Dim OldMemberID As String
    Dim austrittsDatum As Date
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    
    If Me.txt_Pachtende.value = "" Then
        MsgBox "Austrittsdatum darf nicht leer sein.", vbCritical
        Exit Sub
    End If
    
    If Not mod_Mitglieder_Logik.IstGueltigesDatum(Me.txt_Pachtende.value) Then
        MsgBox "Austrittsdatum: Bitte ein g�ltiges Datum eingeben (Format: TT.MM.JJJJ).", vbExclamation
        Exit Sub
    End If
    
    austrittsDatum = CDate(Me.txt_Pachtende.value)
    nachname = wsM.Cells(lRow, M_COL_NACHNAME).value
    vorname = wsM.Cells(lRow, M_COL_VORNAME).value
    OldParzelle = wsM.Cells(lRow, M_COL_PARZELLE).value
    OldMemberID = wsM.Cells(lRow, M_COL_MEMBER_ID).value
    
    ' === SICHERHEITSCHECK: Verein-Parzelle darf NIEMALS gel�scht werden ===
    If UCase(Trim(OldParzelle)) = "VEREIN" Then
        MsgBox "FEHLER: Die Verein-Parzelle darf nicht gel�scht werden!", vbCritical, "Operation nicht erlaubt"
        Exit Sub
    End If
    
    ' Verschiebe Mitglied in Mitgliederhistorie
    Call mod_Mitglieder_Logik.VerschiebeInHistorie(lRow, OldParzelle, OldMemberID, nachname, vorname, austrittsDatum, grund, nachpaechterName, nachpaechterID)
    
    ' Formatierung neu anwenden
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
    If mod_Mitglieder_Logik.IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    Unload Me
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler beim Austritt: " & Err.Description, vbCritical
End Sub


' ***************************************************************
' NEU v2.8: cmd_Uebernehmen_MitKomplettAustritt
' Wird aufgerufen wenn Komplett-Austritt (alle Parzellen) best�tigt wird
' ***************************************************************
Private Sub cmd_Uebernehmen_MitKomplettAustritt(ByVal lRow As Long, ByVal grund As String)
    
    Dim wsM As Worksheet
    Dim nachname As String
    Dim vorname As String
    Dim OldMemberID As String
    Dim austrittsDatum As Date
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    
    If Me.txt_Pachtende.value = "" Then
        MsgBox "Austrittsdatum darf nicht leer sein.", vbCritical
        Exit Sub
    End If
    
    If Not mod_Mitglieder_Logik.IstGueltigesDatum(Me.txt_Pachtende.value) Then
        MsgBox "Austrittsdatum: Bitte ein g" & ChrW(252) & "ltiges Datum eingeben (Format: TT.MM.JJJJ).", vbExclamation
        Exit Sub
    End If
    
    austrittsDatum = CDate(Me.txt_Pachtende.value)
    nachname = wsM.Cells(lRow, M_COL_NACHNAME).value
    vorname = wsM.Cells(lRow, M_COL_VORNAME).value
    OldMemberID = wsM.Cells(lRow, M_COL_MEMBER_ID).value
    
    ' Verschiebe ALLE Parzellen des Mitglieds in die Historie
    Call mod_Mitglieder_Logik.VerschiebeAlleParzellenInHistorie(OldMemberID, nachname, vorname, austrittsDatum, grund)
    
    ' Formatierung neu anwenden
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
    If mod_Mitglieder_Logik.IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    Unload Me
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler beim Komplett-Austritt: " & Err.Description, vbCritical
End Sub


Private Sub cmd_Anlegen_Click()
    Dim wsM As Worksheet
    Dim lRow As Long
    Dim autoSeite As String
    Dim funktion As String
    Dim istMitgliedOhnePacht As Boolean
    Dim antwort As VbMsgBoxResult
    Dim parzelle As String
    Dim parzelleHatMitgliedMitPacht As Boolean
    Dim r As Long
    Dim lastRow As Long
    Dim funktion_in_zeile As String
    Dim newMemberID As String
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    
    ' === PFLICHTFELDER VALIDIERUNG ===
    If Trim(Me.txt_Nachname.value) = "" Or Trim(Me.txt_Vorname.value) = "" Then
        MsgBox "Nachname und Vorname d�rfen nicht leer sein.", vbCritical
        Exit Sub
    End If
    
    If Trim(Me.cbo_Parzelle.value) = "" Then
        MsgBox "Parzelle muss ausgef�llt werden.", vbCritical
        Exit Sub
    End If
    
    If Trim(Me.cbo_Funktion.value) = "" Then
        MsgBox "Funktion muss ausgew�hlt werden.", vbCritical
        Exit Sub
    End If
    
    ' === DATUMSVALIDIERUNG ===
    If Not mod_Mitglieder_Logik.IstGueltigesDatum(Me.txt_Geburtstag.value) Then
        MsgBox "Geburtstag: Bitte ein g�ltiges Datum eingeben (Format: TT.MM.JJJJ).", vbExclamation
        Exit Sub
    End If
    
    If Not mod_Mitglieder_Logik.IstGueltigesDatum(Me.txt_Pachtbeginn.value) Then
        MsgBox "Pachtbeginn: Bitte ein g�ltiges Datum eingeben (Format: TT.MM.JJJJ).", vbExclamation
        Exit Sub
    End If
    
    If Not mod_Mitglieder_Logik.IstGueltigesDatum(Me.txt_Pachtende.value) Then
        MsgBox "Pachtende: Bitte ein g�ltiges Datum eingeben (Format: TT.MM.JJJJ).", vbExclamation
        Exit Sub
    End If
    
    funktion = Me.cbo_Funktion.value
    parzelle = Me.cbo_Parzelle.value
    istMitgliedOhnePacht = (funktion = "Mitglied ohne Pacht")
    
    ' --- VALIDIERUNG: Pachtbeginn je nach Funktion ---
    If Not istMitgliedOhnePacht Then
        ' Mit Pacht: Pachtbeginn MANDATORY
        If Me.txt_Pachtbeginn.value = "" Then
            MsgBox "F�r diese Funktion ist ein Pachtbeginn erforderlich.", vbCritical
            Exit Sub
        End If
    End If
    
    ' --- VALIDIERUNG: Parzelle je nach Funktion ---
    If Not istMitgliedOhnePacht Then
        ' Mit Pacht: Muss eine Parzelle haben
        If parzelle = "" Then
            MsgBox "F�r diese Funktion muss eine Parzelle ausgew�hlt sein.", vbCritical
            Exit Sub
        End If
    Else
        ' Mitglied ohne Pacht: SPEZIELLE REGEL
        ' - Parzelle kann leer sein, ODER
        ' - Parzelle muss bereits ein "Pacht-Mitglied" haben (mit Pacht oder Vorstandsmitglied)
        
        If parzelle <> "" Then
            ' Pr�fe ob auf dieser Parzelle ein Mitglied mit Pacht existiert
            parzelleHatMitgliedMitPacht = False
            lastRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row
            
            For r = M_START_ROW To lastRow
                ' Suche Mitglieder auf dieser Parzelle
                If StrComp(Trim(wsM.Cells(r, M_COL_PARZELLE).value), parzelle, vbTextCompare) = 0 Then
                    ' Diese Parzelle hat ein Mitglied - hat es Pacht?
                    funktion_in_zeile = wsM.Cells(r, M_COL_FUNKTION).value
                    
                    ' REGEL: Folgende Funktionen sind IMMER mit Pacht:
                    If funktion_in_zeile = "Mitglied mit Pacht" Or _
                       funktion_in_zeile = "1. Vorsitzende(r)" Or _
                       funktion_in_zeile = "2. Vorsitzende(r)" Or _
                       funktion_in_zeile = "Kassierer(in)" Or _
                       funktion_in_zeile = "Schriftf�hrer(in)" Then
                        parzelleHatMitgliedMitPacht = True
                        Exit For
                    End If
                End If
            Next r
            
            If Not parzelleHatMitgliedMitPacht Then
                MsgBox "Ein Mitglied ohne Pacht darf nur auf eine Parzelle mit einem Mitglied mit Pacht oder einem Vorstandsmitglied angemeldet werden.", vbCritical
                Exit Sub
            End If
        End If
    End If
    
    ' --- VALIDIERUNG: Pr�fe Duplikate bei Vorsitzende ---
    If funktion = "1. Vorsitzende(r)" Or funktion = "2. Vorsitzende(r)" Then
        If mod_Mitglieder_Logik.FunktionExistiertBereits(funktion, "") Then
            antwort = MsgBox("Es gibt bereits einen/eine " & funktion & "!" & vbCrLf & vbCrLf & _
                           "Soll wirklich ein(e) weitere(r) " & funktion & " angelegt werden?", vbYesNo + vbExclamation, "Warnung")
            If antwort = vbNo Then Exit Sub
        End If
    End If

    wsM.Unprotect PASSWORD:=PASSWORD
    
    lRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row + 1
    
    newMemberID = mod_Mitglieder_UI.CreateGUID_Public()
    
    ' === SICHERHEITSCHECK: Doppelte Eintr�ge verhindern ===
    If mod_Mitglieder_Logik.ExistiertBereitsAufParzelle(newMemberID, parzelle) Then
        wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        MsgBox "FEHLER: Diese Person existiert bereits auf Parzelle " & parzelle & "!" & vbCrLf & _
               "Doppelte Eintr�ge sind nicht erlaubt.", vbCritical, "Doppelter Eintrag verhindert"
        Exit Sub
    End If
    
    wsM.Cells(lRow, M_COL_MEMBER_ID).value = newMemberID
    
    autoSeite = mod_Mitglieder_Logik.GetSeiteFromParzelle(Me.cbo_Parzelle.value)
    
    On Error Resume Next
    
    wsM.Cells(lRow, M_COL_PARZELLE).value = Me.cbo_Parzelle.value
    wsM.Cells(lRow, M_COL_SEITE).value = autoSeite
    wsM.Cells(lRow, M_COL_ANREDE).value = Me.cbo_Anrede.value
    wsM.Cells(lRow, M_COL_NACHNAME).value = Me.txt_Nachname.value
    wsM.Cells(lRow, M_COL_VORNAME).value = Me.txt_Vorname.value
    wsM.Cells(lRow, M_COL_STRASSE).value = Me.txt_Strasse.value
    wsM.Cells(lRow, M_COL_NUMMER).value = Me.txt_Nummer.value
    wsM.Cells(lRow, M_COL_PLZ).value = Me.txt_PLZ.value
    wsM.Cells(lRow, M_COL_WOHNORT).value = Me.txt_Wohnort.value
    wsM.Cells(lRow, M_COL_TELEFON).value = Me.txt_Telefon.value
    wsM.Cells(lRow, M_COL_MOBIL).value = Me.txt_Mobil.value
    wsM.Cells(lRow, M_COL_GEBURTSTAG).value = Me.txt_Geburtstag.value
    wsM.Cells(lRow, M_COL_EMAIL).value = Me.txt_Email.value
    wsM.Cells(lRow, M_COL_FUNKTION).value = Me.cbo_Funktion.value
    
    ' Pachtbeginn mit Fehlerbehandlung
    If Me.txt_Pachtbeginn.value <> "" Then
        wsM.Cells(lRow, M_COL_PACHTANFANG).value = CDate(Me.txt_Pachtbeginn.value)
        If Err.Number = 0 Then
            wsM.Cells(lRow, M_COL_PACHTANFANG).NumberFormat = "dd.mm.yyyy"
        End If
        Err.Clear
    End If
    
    ' Pachtende mit Fehlerbehandlung
    If Me.txt_Pachtende.value <> "" Then
        wsM.Cells(lRow, M_COL_PACHTENDE).value = CDate(Me.txt_Pachtende.value)
        If Err.Number = 0 Then
            wsM.Cells(lRow, M_COL_PACHTENDE).NumberFormat = "dd.mm.yyyy"
        End If
        Err.Clear
    End If
    
    On Error GoTo 0
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Call mod_Mitglieder_UI.Sortiere_Mitgliederliste_Nach_Parzelle
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
    If mod_Mitglieder_Logik.IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    MsgBox "Neues Mitglied " & Me.txt_Nachname.value & " erfolgreich angelegt.", vbInformation
    
    Unload Me
    Exit Sub
ErrorHandler:
    On Error GoTo 0
    If Not wsM Is Nothing Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler beim Anlegen des neuen Mitglieds: " & Err.Description, vbCritical
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 1
    m_AlreadyInitialized = False  ' Flag zur�cksetzen
    
    On Error GoTo ErrorHandler
    
    Me.cbo_Anrede.RowSource = "Daten!D4:D9"
    
    ' Funktion dynamisch f�llen
    Call FuelleFunktionComboDB
    
    ' Fuelle cbo_Parzelle OHNE "Verein"
    Call FuelleParzelleComboDB
    
    ' Setze default Captions f�r die Label-Bezeichner IMMER
    Me.lbl_PachtbeginnBezeichner.Caption = "Pachtbeginn"
    Me.lbl_PachtendeBezeichner.Caption = "Pachtende"
    
    Exit Sub
ErrorHandler:
    MsgBox "Fehler beim Initialisieren der Form: " & Err.Description, vbCritical
End Sub

' ***************************************************************
' EVENT: UserForm_Activate - wird NACH Tag-Setzen ausgef�hrt!
' ***************************************************************
Private Sub UserForm_Activate()
    ' Verhindere doppelte Ausf�hrung
    If m_AlreadyInitialized Then Exit Sub
    m_AlreadyInitialized = True
    
    On Error GoTo ErrorHandler
    
    Dim tagStr As String
    tagStr = CStr(Me.tag)
    
    ' DEBUG: Zeige Tag-Wert (optional auskommentieren nach Test)
    Debug.Print "DEBUG UserForm_Activate - Tag = '" & tagStr & "'"
    
    ' Pr�fe ob es ein Nachp�chter-NEU Modus ist
    If InStr(tagStr, "NACHPAECHTER_NEU") > 0 Then
        Debug.Print "DEBUG: NACHPAECHTER_NEU erkannt - setze EditMode"
        Call SetMode(True, True, False)
        Exit Sub
    End If
    
    ' Pr�fe ob "NEU" f�r neues Mitglied
    If tagStr = "NEU" Then
        Debug.Print "DEBUG: NEU erkannt - leere Felder und setze EditMode"
        
        ' Leere alle Felder
        Me.cbo_Parzelle.value = ""
        Me.cbo_Anrede.value = ""
        Me.txt_Vorname.value = ""
        Me.txt_Nachname.value = ""
        Me.txt_Strasse.value = ""
        Me.txt_Nummer.value = ""
        Me.txt_PLZ.value = ""
        Me.txt_Wohnort.value = ""
        Me.txt_Telefon.value = ""
        Me.txt_Mobil.value = ""
        Me.txt_Geburtstag.value = ""
        Me.txt_Email.value = ""
        Me.cbo_Funktion.value = ""
        Me.txt_Pachtende.value = ""
        
        ' F�lle txt_Pachtbeginn mit aktuellem Datum
        Me.txt_Pachtbeginn.value = Format(Date, "dd.mm.yyyy")
        
        Call SetMode(True, True, False)
        Exit Sub
    End If
    
    ' F�r bestehende Mitglieder: ViewMode (nur Labels sichtbar)
    Debug.Print "DEBUG: Bestehendes Mitglied - setze ViewMode"
    Call SetMode(False, False, False)
    
    Exit Sub
ErrorHandler:
    MsgBox "Fehler beim Aktivieren der Form: " & Err.Description, vbCritical
End Sub

' ***************************************************************
' HILFSPROZEDUR: FuelleParzelleComboDB
' F�llt die Parzelle ComboBox mit allen Werten AUSSER "Verein"
' ***************************************************************
Private Sub FuelleParzelleComboDB()
    Dim ws As Worksheet
    Dim lRow As Long
    Dim parzelleValue As String
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    If ws Is Nothing Then
        ' Fallback: Nutze F4:F17 (OHNE F18 = "Verein")
        Me.cbo_Parzelle.RowSource = "Daten!F4:F17"
        Exit Sub
    End If
    
    ' Leere die ComboBox zuerst
    Me.cbo_Parzelle.Clear
    
    ' Lese alle Werte von F4:F17 und f�ge sie hinzu, AUSSER "Verein"
    For lRow = 4 To 17
        parzelleValue = Trim(ws.Cells(lRow, 6).value)
        
        ' �berspringe leere Zellen und "Verein"
        If parzelleValue <> "" And UCase(parzelleValue) <> "VEREIN" Then
            Me.cbo_Parzelle.AddItem parzelleValue
        End If
    Next lRow
    
    Exit Sub
ErrorHandler:
    ' Fallback bei Fehler: Nutze F4:F17
    Me.cbo_Parzelle.RowSource = "Daten!F4:F17"
End Sub

' ***************************************************************
' HILFSPROZEDUR: FuelleFunktionComboDB
' F�llt die Funktion ComboBox dynamisch aus Daten!B4:B?
' ***************************************************************
Private Sub FuelleFunktionComboDB()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lRow As Long
    Dim funktionValue As String
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    If ws Is Nothing Then
        ' Fallback
        Me.cbo_Funktion.RowSource = "Daten!B4:B12"
        Exit Sub
    End If
    
    ' Leere die ComboBox zuerst
    Me.cbo_Funktion.Clear
    
    ' Finde letzte gef�llte Zeile in Spalte B
    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    
    ' Lese alle Werte von B4 bis lastRow
    For lRow = 4 To lastRow
        funktionValue = Trim(ws.Cells(lRow, 2).value)
        
        ' �berspringe leere Zellen
        If funktionValue <> "" Then
            Me.cbo_Funktion.AddItem funktionValue
        End If
    Next lRow
    
    Exit Sub
ErrorHandler:
    ' Fallback bei Fehler
    Me.cbo_Funktion.RowSource = "Daten!B4:B12"
End Sub
