VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Mitgliedsdaten 
   Caption         =   "Mitgliedsdaten"
   ClientHeight    =   8580.001
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   7790
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

' --- Hilfsfunktion für Parzelle -> Seite ---
Private Function GetSeiteFromParzelle(ByVal parzelle As String) As String
    Dim parzelleNum As Long
    
    If UCase(Trim(parzelle)) = "VEREIN" Then
        GetSeiteFromParzelle = "zentral"
        Exit Function
    End If
    
    On Error Resume Next
    parzelleNum = CLng(Left(parzelle, InStr(parzelle & " ", " ") - 1))
    On Error GoTo 0
    
    If parzelleNum = 0 Then
        GetSeiteFromParzelle = ""
        Exit Function
    End If
    
    If parzelleNum >= 1 And parzelleNum <= 9 Then
        GetSeiteFromParzelle = "rechts"
    ElseIf parzelleNum >= 10 And parzelleNum <= 14 Then
        GetSeiteFromParzelle = "links"
    Else
        GetSeiteFromParzelle = ""
    End If
    
End Function

' --- Prüfe ob Funktion bereits existiert ---
Private Function FunktionExistiertBereits(ByVal funktion As String, ByVal ausschlussParzelle As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    lastRow = ws.Cells(ws.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        If ws.Cells(r, M_COL_FUNKTION).value = funktion And _
           ws.Cells(r, M_COL_PARZELLE).value <> ausschlussParzelle And _
           ws.Cells(r, M_COL_PARZELLE).value <> "" Then
            FunktionExistiertBereits = True
            Exit Function
        End If
    Next r
    
    FunktionExistiertBereits = False
End Function

' --- Hilfsfunktion: Prüfe ob String eine Zahl ist ---
Private Function IsNumeric(ByVal value As String) As Boolean
    On Error Resume Next
    IsNumeric = Not IsError(CLng(value))
    On Error GoTo 0
End Function

' ***************************************************************
' HILFSPROZEDUR: Aktualisiert Labels basierend auf Funktion
' ***************************************************************
Private Sub AktualisiereLabelsFuerFunktion()
    Dim istMitgliedOhnePacht As Boolean
    
    ' Prüfe ob cbo_Funktion einen Wert hat
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
' ***************************************************************
Public Sub SetMode(ByVal EditMode As Boolean, Optional ByVal IsNewEntry As Boolean = False, Optional ByVal IsRemovalMode As Boolean = False)
    
    Dim ctl As MSForms.Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is MSForms.Label And Left(ctl.Name, 4) = "lbl_" Then
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
    
    If CStr(Me.Tag) = "NEU" Then
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
    
    If CStr(Me.Tag) <> "NEU" Then
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
    ElseIf IsNewEntry Then
        ' Für neue Mitglieder: Felder leer lassen
        Me.txt_Pachtbeginn.value = ""
        Me.txt_Pachtende.value = ""
    End If
    
End Sub

Private Sub cmd_Bearbeiten_Click()
    Call SetMode(True, False, False)
End Sub

Private Sub cmd_Abbrechen_Click()
    If CStr(Me.Tag) = "NEU" Then
        Unload Me
        Exit Sub
    End If
    Call SetMode(False)
End Sub

' ***************************************************************
' EVENT: ComboBox Funktion-Änderung
' ***************************************************************
Private Sub cbo_Funktion_Change()
    Call AktualisiereLabelsFuerFunktion
End Sub

Private Sub cmd_Entfernen_Click()
    
    Dim lRow As Long
    Dim Nachname As String
    Dim Vorname As String
    Dim OldParzelle As String
    Dim OldMemberID As String
    Dim AustrittsDatum As Date
    Dim ChangeReason As String
    Dim pachtEndeVal As String
    Dim auswahlOption As Integer
    
    If Not IsNumeric(Me.Tag) Or CLng(Me.Tag) < M_START_ROW Then
        MsgBox "Interner Fehler: Keine gültige Zeilennummer für das Entfernen gefunden.", vbCritical
        Exit Sub
    End If
    
    lRow = CLng(Me.Tag)
    Nachname = Me.lbl_Nachname.Caption
    Vorname = Me.lbl_Vorname.Caption
    OldParzelle = Me.lbl_Parzelle.Caption
    OldMemberID = ThisWorkbook.Worksheets(WS_MITGLIEDER).Cells(lRow, M_COL_MEMBER_ID).value
    
    ' Prüfe ob Pachtende bereits gefüllt ist
    pachtEndeVal = Trim(Me.lbl_Pachtende.Caption)
    
    ' Zeige Austrittsauswahl-Dialog
    With frm_Austrittsauswahl
        .Show vbModal
        auswahlOption = .SelectedOption
        ChangeReason = .CustomReason
        Unload frm_Austrittsauswahl
    End With
    
    If auswahlOption = 0 Then
        ' Benutzer hat abgebrochen
        Exit Sub
    End If
    
    Select Case auswahlOption
        Case 1 ' Nachpächter
            If ChangeReason = "" Then ChangeReason = "Übergabe an Nachpächter"
            GoTo AustrittBearbeiten
            
        Case 2 ' Tod
            If ChangeReason = "" Then ChangeReason = "Tod des Mitglieds"
            GoTo AustrittBearbeiten
            
        Case 3 ' Kündigung
            If ChangeReason = "" Then ChangeReason = "Kündigung"
            GoTo AustrittBearbeiten
            
        Case 4 ' Parzellenwechsel
            ChangeReason = "Parzellenwechsel"
            Call SetMode(True, False, False)
            MsgBox "Bitte geben Sie die neue Parzellennummer in das Feld 'Parzelle' ein und klicken Sie dann 'Übernehmen'.", vbInformation, "Parzellenwechsel"
            Exit Sub
            
        Case 5 ' Sonstiges
            If ChangeReason = "" Then ChangeReason = "Sonstiges"
            GoTo AustrittBearbeiten
    End Select
    
AustrittBearbeiten:
    If pachtEndeVal = "" Then
        ' Pachtende ist noch leer - Benutzer kann es eintragen
        Call SetMode(True, False, False)
        
        ' Speichere Grund temporär im Tag des Formulars
        Me.Tag = lRow & "|" & ChangeReason
        
        ' Fülle Pachtende mit heutigem Datum und MARKIERE ES komplett
        Me.txt_Pachtende.value = Format(Date, "dd.mm.yyyy")
        Me.txt_Pachtende.SetFocus
        Me.txt_Pachtende.SelStart = 0
        Me.txt_Pachtende.SelLength = Len(Me.txt_Pachtende.value)
        
        MsgBox "Das Austrittsdatum wurde auf heute gesetzt." & vbCrLf & _
               "Grund: " & ChangeReason & vbCrLf & vbCrLf & _
               "Bitte bestätigen Sie es (oder ändern Sie es) und klicken Sie dann 'Übernehmen'.", vbInformation, "Austrittsdatum"
        Exit Sub
    Else
        ' Pachtende ist bereits gesetzt - Mitglied in Historie verschieben
        AustrittsDatum = CDate(pachtEndeVal)
    End If
    
    ' Verschiebe Mitglied in Mitgliederhistorie
    Call VerschiebeInHistorie(lRow, OldParzelle, OldMemberID, Nachname, Vorname, AustrittsDatum, ChangeReason)
    
    ' Formatierung neu anwenden
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
    If IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    Unload Me
    
End Sub

' ***************************************************************
' HILFSPROZEDUR: VerschiebeInHistorie
' Verschiebt ein Mitglied von Mitgliederliste in Mitgliederhistorie
' ***************************************************************
Private Sub VerschiebeInHistorie(ByVal lRow As Long, ByVal parzelle As String, ByVal memberID As String, _
                                   ByVal Nachname As String, ByVal Vorname As String, _
                                   ByVal AustrittsDatum As Date, ByVal grund As String)
    
    Dim wsM As Worksheet
    Dim wsH As Worksheet
    Dim nextHistRow As Long
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    
    ' Entsperre beide Blätter
    wsM.Unprotect PASSWORD:=PASSWORD
    wsH.Unprotect PASSWORD:=PASSWORD
    
    ' Finde nächste freie Zeile in Mitgliederhistorie (ab Zeile 4)
    nextHistRow = wsH.Cells(wsH.Rows.Count, 3).End(xlUp).Row + 1
    If nextHistRow < 4 Then nextHistRow = 4
    
    ' Schreibe Daten in Mitgliederhistorie
    wsH.Cells(nextHistRow, 1).value = parzelle              ' Spalte A: Parzelle
    wsH.Cells(nextHistRow, 2).value = memberID              ' Spalte B: Member ID (alt)
    wsH.Cells(nextHistRow, 3).value = Nachname              ' Spalte C: Nachname
    wsH.Cells(nextHistRow, 4).value = Vorname               ' Spalte D: Vorname
    wsH.Cells(nextHistRow, 5).value = AustrittsDatum        ' Spalte E: Austrittsdatum
    wsH.Cells(nextHistRow, 5).NumberFormat = "dd.mm.yyyy"
    wsH.Cells(nextHistRow, 6).value = grund                 ' Spalte F: Grund
    wsH.Cells(nextHistRow, 7).value = ""                    ' Spalte G: Endabrechnung (leer)
    
    ' Lösche Zeile aus Mitgliederliste
    wsM.Rows(lRow).Delete Shift:=xlUp
    
    ' Schütze Blätter wieder
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    wsH.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    MsgBox "Mitglied " & Nachname & " wurde in die Mitgliederhistorie verschoben." & vbCrLf & _
           "Grund: " & grund, vbInformation
    
    Exit Sub
ErrorHandler:
    On Error GoTo 0
    If Not wsM Is Nothing Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    If Not wsH Is Nothing Then wsH.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler beim Verschieben in Historie: " & Err.Description, vbCritical
End Sub

Private Sub cmd_Uebernehmen_Click()
    
    Dim tagParts() As String
    Dim lRow As Long
    Dim grund As String
    
    ' Prüfe ob Tag im Format "lRow|Grund" vorliegt (bei Austritt)
    If InStr(Me.Tag, "|") > 0 Then
        tagParts = Split(Me.Tag, "|")
        If UBound(tagParts) >= 1 Then
            ' Austritt-Modus mit Grund
            Call cmd_Uebernehmen_MitAustritt(CLng(tagParts(0)), tagParts(1))
            Exit Sub
        End If
    End If
    
    ' Normale Validierung für Bearbeiten-Modus
    If Not IsNumeric(Me.Tag) Or CLng(Me.Tag) < M_START_ROW Then
        MsgBox "Interner Fehler: Keine gültige Zeilennummer für das Speichern gefunden.", vbCritical
        Exit Sub
    End If
    
    lRow = CLng(Me.Tag)
    
    Dim wsM As Worksheet
    Dim autoSeite As String
    Dim funktion As String
    Dim istMitgliedOhnePacht As Boolean
    Dim OldParzelle As String
    Dim NewParzelle As String
    Dim Nachname As String
    Dim ChangeReason As String
    Dim AustrittsDatum As Date
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    
    If Me.txt_Nachname.value = "" Or Me.txt_Vorname.value = "" Then
        MsgBox "Nachname und Vorname dürfen nicht leer sein.", vbCritical
        Exit Sub
    End If
    
    funktion = Me.cbo_Funktion.value
    istMitgliedOhnePacht = (funktion = "Mitglied ohne Pacht")
    
    ' === SICHERHEITSCHECK: Verein-Parzelle darf nicht bearbeitet werden ===
    If Trim(Me.lbl_Parzelle.Caption) = PARZELLE_VEREIN Then
        MsgBox "FEHLER: Die Verein-Parzelle darf nicht bearbeitet werden!", vbCritical
        Exit Sub
    End If
    
    ' --- VALIDIERUNG: Pachtbeginn je nach Funktion ---
    If Not istMitgliedOhnePacht Then
        ' Mit Pacht: Pachtbeginn ist mandatory
        If Me.txt_Pachtbeginn.value = "" Then
            MsgBox "Für diese Funktion ist ein Pachtbeginn erforderlich.", vbCritical
            Exit Sub
        End If
        If Not IsDate(Me.txt_Pachtbeginn.value) Then
            MsgBox "Pachtbeginn: Bitte ein gültiges Datum eingeben.", vbExclamation
            Exit Sub
        End If
    Else
        ' Ohne Pacht: Pachtbeginn ist optional
        If Me.txt_Pachtbeginn.value <> "" Then
            If Not IsDate(Me.txt_Pachtbeginn.value) Then
                MsgBox "Mitgliedsbeginn: Bitte ein gültiges Datum eingeben.", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    If Me.txt_Pachtende.value <> "" Then
        If Not IsDate(Me.txt_Pachtende.value) Then
            MsgBox "Pachtende: Bitte ein gültiges Datum eingeben.", vbExclamation
            Exit Sub
        End If
    End If

    ' Prüfe auf Parzellenwechsel
    OldParzelle = Me.lbl_Parzelle.Caption
    NewParzelle = Me.cbo_Parzelle.value
    Nachname = Me.txt_Nachname.value

    wsM.Unprotect PASSWORD:=PASSWORD
    
    On Error Resume Next
    
    autoSeite = GetSeiteFromParzelle(Me.cbo_Parzelle.value)
    
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
    
    If Me.txt_Pachtbeginn.value <> "" Then
        wsM.Cells(lRow, M_COL_PACHTANFANG).value = CDate(Me.txt_Pachtbeginn.value)
        wsM.Cells(lRow, M_COL_PACHTANFANG).NumberFormat = "dd.mm.yyyy"
    End If
    
    If Me.txt_Pachtende.value <> "" Then
        wsM.Cells(lRow, M_COL_PACHTENDE).value = CDate(Me.txt_Pachtende.value)
        wsM.Cells(lRow, M_COL_PACHTENDE).NumberFormat = "dd.mm.yyyy"
    End If
    
    On Error GoTo 0
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' Prüfe auf Parzellenwechsel und speichere ggf. in Historie
    If OldParzelle <> "" And NewParzelle <> OldParzelle Then
        ' Parzellenwechsel erkannt
        Call mod_Mitglieder_UI.Speichere_Historie_und_Aktualisiere_Mitgliederliste( _
             lRow, OldParzelle, "", Nachname, Date, NewParzelle, "", "Parzellenwechsel")
    Else
        ' Normale Änderung - nur Sortierung und Formatierung
        Call mod_Mitglieder_UI.Sortiere_Mitgliederliste_Nach_Parzelle
        Call mod_Mitglieder_UI.Fuelle_MemberIDs_Wenn_Fehlend
    End If
    
    ' Formatierung neu anwenden
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
    If IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    MsgBox "Änderungen für Mitglied " & Me.txt_Nachname.value & " erfolgreich gespeichert.", vbInformation
    
    Unload Me
    Exit Sub
ErrorHandler:
    On Error GoTo 0
    If Not wsM Is Nothing Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler beim Speichern der Änderungen: " & Err.Description, vbCritical
End Sub

' ***************************************************************
' HILFSPROZEDUR: cmd_Uebernehmen_MitAustritt
' Wird aufgerufen wenn Austritt mit Grund durchgeführt wird
' ***************************************************************
Private Sub cmd_Uebernehmen_MitAustritt(ByVal lRow As Long, ByVal grund As String)
    
    Dim wsM As Worksheet
    Dim Nachname As String
    Dim Vorname As String
    Dim OldParzelle As String
    Dim OldMemberID As String
    Dim AustrittsDatum As Date
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    
    If Me.txt_Pachtende.value = "" Then
        MsgBox "Austrittsdatum darf nicht leer sein.", vbCritical
        Exit Sub
    End If
    
    If Not IsDate(Me.txt_Pachtende.value) Then
        MsgBox "Austrittsdatum: Bitte ein gültiges Datum eingeben.", vbExclamation
        Exit Sub
    End If
    
    AustrittsDatum = CDate(Me.txt_Pachtende.value)
    Nachname = wsM.Cells(lRow, M_COL_NACHNAME).value
    Vorname = wsM.Cells(lRow, M_COL_VORNAME).value
    OldParzelle = wsM.Cells(lRow, M_COL_PARZELLE).value
    OldMemberID = wsM.Cells(lRow, M_COL_MEMBER_ID).value
    
    ' Verschiebe Mitglied in Mitgliederhistorie
    Call VerschiebeInHistorie(lRow, OldParzelle, OldMemberID, Nachname, Vorname, AustrittsDatum, grund)
    
    ' Formatierung neu anwenden
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
    If IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    Unload Me
    Exit Sub
ErrorHandler:
    MsgBox "Fehler beim Austritt: " & Err.Description, vbCritical
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
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    
    If Me.txt_Nachname.value = "" Or Me.txt_Vorname.value = "" Then
        MsgBox "Nachname und Vorname dürfen nicht leer sein.", vbCritical
        Exit Sub
    End If
    
    funktion = Me.cbo_Funktion.value
    parzelle = Me.cbo_Parzelle.value
    istMitgliedOhnePacht = (funktion = "Mitglied ohne Pacht")
    
    ' --- VALIDIERUNG 1: Pachtbeginn je nach Funktion ---
    If Not istMitgliedOhnePacht Then
        ' Mit Pacht: Pachtbeginn MANDATORY - Auto-Ausfüllung wenn leer
        If Me.txt_Pachtbeginn.value = "" Then
            Me.txt_Pachtbeginn.value = Format(Date, "dd.mm.yyyy")
        End If
        If Not IsDate(Me.txt_Pachtbeginn.value) Then
            MsgBox "Pachtbeginn: Bitte ein gültiges Datum eingeben.", vbExclamation
            Exit Sub
        End If
    Else
        ' Mitglied ohne Pacht: Pachtbeginn optional
        If Me.txt_Pachtbeginn.value <> "" Then
            If Not IsDate(Me.txt_Pachtbeginn.value) Then
                MsgBox "Mitgliedsbeginn: Bitte ein gültiges Datum eingeben.", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    ' --- VALIDIERUNG 2: Pachtende ---
    If Me.txt_Pachtende.value <> "" Then
        If Not IsDate(Me.txt_Pachtende.value) Then
            MsgBox "Pachtende: Bitte ein gültiges Datum eingeben.", vbExclamation
            Exit Sub
        End If
    End If
    
    ' --- VALIDIERUNG 3: Parzelle je nach Funktion ---
    If Not istMitgliedOhnePacht Then
        ' Mit Pacht: Muss eine Parzelle haben
        If parzelle = "" Then
            MsgBox "Für diese Funktion muss eine Parzelle ausgewählt sein.", vbCritical
            Exit Sub
        End If
    Else
        ' Mitglied ohne Pacht: SPEZIELLE REGEL
        ' - Parzelle kann leer sein, ODER
        ' - Parzelle muss bereits ein "Pacht-Mitglied" haben (mit Pacht oder Vorstandsmitglied)
        
        If parzelle <> "" Then
            ' Prüfe ob auf dieser Parzelle ein Mitglied mit Pacht existiert
            parzelleHatMitgliedMitPacht = False
            lastRow = wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
            
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
                       funktion_in_zeile = "Schriftführer(in)" Then
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
    
    ' --- VALIDIERUNG 4: Prüfe Duplikate bei Vorsitzende ---
    If funktion = "1. Vorsitzende(r)" Or funktion = "2. Vorsitzende(r)" Then
        If FunktionExistiertBereits(funktion, "") Then
            antwort = MsgBox("Es gibt bereits einen/eine " & funktion & "!" & vbCrLf & vbCrLf & _
                           "Soll wirklich ein(e) weitere(r) " & funktion & " angelegt werden?", vbYesNo + vbExclamation, "Warnung")
            If antwort = vbNo Then Exit Sub
        End If
    End If

    wsM.Unprotect PASSWORD:=PASSWORD
    
    lRow = wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).Row + 1
    
    wsM.Cells(lRow, M_COL_MEMBER_ID).value = mod_Mitglieder_UI.CreateGUID_Public()
    
    autoSeite = GetSeiteFromParzelle(Me.cbo_Parzelle.value)
    
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
    
    If Me.txt_Pachtbeginn.value <> "" Then
        wsM.Cells(lRow, M_COL_PACHTANFANG).value = CDate(Me.txt_Pachtbeginn.value)
        wsM.Cells(lRow, M_COL_PACHTANFANG).NumberFormat = "dd.mm.yyyy"
    End If
    
    If Me.txt_Pachtende.value <> "" Then
        wsM.Cells(lRow, M_COL_PACHTENDE).value = CDate(Me.txt_Pachtende.value)
        wsM.Cells(lRow, M_COL_PACHTENDE).NumberFormat = "dd.mm.yyyy"
    End If
    
    On Error GoTo 0
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Call mod_Mitglieder_UI.Sortiere_Mitgliederliste_Nach_Parzelle
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
    If IsFormLoaded("frm_Mitgliederverwaltung") Then
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
    
    On Error GoTo ErrorHandler
    
    Me.cbo_Anrede.RowSource = "Daten!D4:D9"
    
    ' Funktion dynamisch füllen
    Call FuelleFunktionComboDB
    
    ' Fuelle cbo_Parzelle OHNE "Verein"
    Call FuelleParzelleComboDB
    
    ' Setze default Captions für die Label-Bezeichner IMMER
    Me.lbl_PachtbeginnBezeichner.Caption = "Pachtbeginn"
    Me.lbl_PachtendeBezeichner.Caption = "Pachtende"
    
    ' Rufe SetMode NUR für NEUE Mitglieder auf
    If CStr(Me.Tag) = "NEU" Then
        Call SetMode(True, True, False)
    Else
        ' Für bestehende Mitglieder: Explizit alle TextBoxen und ComboBoxen ausblenden
        Dim ctl As MSForms.Control
        For Each ctl In Me.Controls
            If TypeOf ctl Is MSForms.TextBox Or TypeOf ctl Is MSForms.ComboBox Then
                ctl.Visible = False
            End If
        Next ctl
        
        ' Stelle sicher, dass die Bezeichner-Labels sichtbar sind
        Me.lbl_PachtbeginnBezeichner.Visible = True
        Me.lbl_PachtendeBezeichner.Visible = True
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Fehler beim Initialisieren der Form: " & Err.Description, vbCritical
End Sub

' ***************************************************************
' HILFSPROZEDUR: FuelleParzelleComboDB
' Füllt die Parzelle ComboBox mit allen Werten AUSSER "Verein"
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
    
    ' Lese alle Werte von F4:F17 und füge sie hinzu, AUSSER "Verein"
    For lRow = 4 To 17
        parzelleValue = Trim(ws.Cells(lRow, 6).value)
        
        ' Überspringe leere Zellen und "Verein"
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
' Füllt die Funktion ComboBox dynamisch aus Daten!B4:B?
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
    
    ' Finde letzte gefüllte Zeile in Spalte B
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' Lese alle Werte von B4 bis lastRow
    For lRow = 4 To lastRow
        funktionValue = Trim(ws.Cells(lRow, 2).value)
        
        ' Überspringe leere Zellen
        If funktionValue <> "" Then
            Me.cbo_Funktion.AddItem funktionValue
        End If
    Next lRow
    
    Exit Sub
ErrorHandler:
    ' Fallback bei Fehler
    Me.cbo_Funktion.RowSource = "Daten!B4:B12"
End Sub

Private Function IsFormLoaded(ByVal FormName As String) As Boolean
    Dim i As Long
    
    For i = 0 To VBA.UserForms.Count - 1
        If StrComp(VBA.UserForms.Item(i).Name, FormName, vbTextCompare) = 0 Then
            IsFormLoaded = True
            Exit Function
        End If
    Next i
    
    IsFormLoaded = False
End Function

