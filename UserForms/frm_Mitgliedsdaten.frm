VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Mitgliedsdaten 
   Caption         =   "Mitgliedsdaten"
   ClientHeight    =   8590.001
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

' ***************************************************************
' INTERNE VARIABLEN FÜR DIE ID-BASIERTE VERARBEITUNG
' ***************************************************************
Private sMemberID As String
Private sMode As String ' Speichert den Modus ("NEU", "EDIT", oder "ANSICHT")
Private sOriginalParzelle As String ' Speichert die ursprüngliche Parzelle
Private sOriginalFunktion As String ' Speichert die ursprüngliche Funktion
Private lMemberRow As Long ' Speichert die Zeilennummer NUR zum Initialisieren (wird später über ID gesucht!)

' ***************************************************************
' ÖFFENTLICHE INITIALISIERUNG
' ***************************************************************
Public Sub Init_MemberData(ByVal RowIndex As Long)

    Dim wsM As Worksheet
    Dim sParzelle As String

    On Error GoTo InitErrorHandler ' Fehlerbehandlung hinzugefügt

    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    lMemberRow = RowIndex

    If lMemberRow < M_START_ROW Then
        ' --- NEUES MITGLIED ANLEGEN ---
        sMode = "NEU"
        sMemberID = ""
        Me.Caption = "Neues Mitglied anlegen"
        
        ' Setze Voreinstellungen und mache Pachtbeginn sichtbar
        Me.txt_Geburtstag.Value = "" ' Geburtstag leer lassen
        Me.txt_Pachtbeginn.Value = Format(Date, "dd.mm.yyyy") ' Pachtbeginn vorbelegen
        Me.txt_Pachtende.Value = "" ' Pachtende leer lassen
        
        ' Controls für NEU-Modus sichtbar/unsichtbar machen
        Me.lbl_Pachtbeginn.Visible = True
        Me.txt_Pachtbeginn.Visible = True
        Me.lbl_Pachtende.Visible = False
        Me.txt_Pachtende.Visible = False
        
        Call SetMode(True) ' Sofort in den Eingabemodus
        GoTo InitExit
    Else
        ' --- BESTEHENDES MITGLIED LADEN (Start im ANSICHT-Modus) ---
        sMode = "ANSICHT"
        
        ' Lade IDs und Originalwerte
        sMemberID = CStr(wsM.Cells(lMemberRow, M_COL_MEMBER_ID).Value)
        sParzelle = CStr(wsM.Cells(lMemberRow, M_COL_PARZELLE).Value)
        sOriginalParzelle = sParzelle
        sOriginalFunktion = CStr(wsM.Cells(lMemberRow, M_COL_FUNKTION).Value)
        
        Me.Caption = "Mitgliedsdaten: " & sParzelle & " - " & CStr(wsM.Cells(lMemberRow, M_COL_NACHNAME).Value)

        ' Daten in die Labels laden
        Me.lbl_Parzelle.Caption = sParzelle
        Me.lbl_Anrede.Caption = CStr(wsM.Cells(lMemberRow, M_COL_ANREDE).Value)
        Me.lbl_Nachname.Caption = CStr(wsM.Cells(lMemberRow, M_COL_NACHNAME).Value)
        Me.lbl_Vorname.Caption = CStr(wsM.Cells(lMemberRow, M_COL_VORNAME).Value)
        Me.lbl_Strasse.Caption = CStr(wsM.Cells(lMemberRow, M_COL_STRASSE).Value)
        Me.lbl_Nummer.Caption = CStr(wsM.Cells(lMemberRow, M_COL_NUMMER).Value)
        Me.lbl_PLZ.Caption = CStr(wsM.Cells(lMemberRow, M_COL_PLZ).Value)
        Me.lbl_Wohnort.Caption = CStr(wsM.Cells(lMemberRow, M_COL_WOHNORT).Value)
        Me.lbl_Telefon.Caption = CStr(wsM.Cells(lMemberRow, M_COL_TELEFON).Value)
        Me.lbl_Mobil.Caption = CStr(wsM.Cells(lMemberRow, M_COL_MOBIL).Value)
        Me.lbl_Geburtstag.Caption = CStr(wsM.Cells(lMemberRow, M_COL_GEBURTSTAG).Value)
        Me.lbl_Email.Caption = CStr(wsM.Cells(lMemberRow, M_COL_EMAIL).Value)
        Me.lbl_Funktion.Caption = CStr(wsM.Cells(lMemberRow, M_COL_FUNKTION).Value)

        ' Pacht-Labels und Textboxen im ANSICHT-Modus unsichtbar
        Me.lbl_Pachtbeginn.Visible = False
        Me.txt_Pachtbeginn.Visible = False
        Me.lbl_Pachtende.Visible = False
        Me.txt_Pachtende.Visible = False
        
        Call SetMode(False) ' Start im ANSICHT-Modus
    End If

InitExit:
    Exit Sub

InitErrorHandler:
    MsgBox "Interner Fehler in Init_MemberData: Das Arbeitsblatt '" & WS_MITGLIEDER & "' ist nicht verfügbar oder Konstanten fehlen. " & Err.Description, vbCritical
    Exit Sub
End Sub

' ***************************************************************
' HILFSPROZEDUR: Setzt den Anzeigemodus der Form
' ***************************************************************
Public Sub SetMode(ByVal EditMode As Boolean) ' EditMode = True -> Bearbeiten/Neu; EditMode = False -> Ansicht
    
    Dim ctl As MSForms.Control
    Dim lCurrentRow As Long
    
    ' 1. Umschalten der Labels vs. Textboxen/Comboboxen
    For Each ctl In Me.Controls
        ' Schließe Pachtfelder von der automatischen Umschaltung aus
        If ctl.Name <> Me.lbl_Pachtbeginn.Name And ctl.Name <> Me.txt_Pachtbeginn.Name And _
            ctl.Name <> Me.lbl_Pachtende.Name And ctl.Name <> Me.txt_Pachtende.Name Then
            
            If TypeOf ctl Is MSForms.Label And Left(ctl.Name, 4) = "lbl_" Then
                ctl.Visible = Not EditMode ' Labels anzeigen in Ansicht
            ElseIf TypeOf ctl Is MSForms.TextBox Or TypeOf ctl Is MSForms.ComboBox Then
                ctl.Visible = EditMode ' Eingabefelder anzeigen in Edit/Neu
            End If
        End If
    Next ctl
    
    ' 2. Button-Sichtbarkeit
    If sMode = "NEU" Then
        ' NEU-MODUS
        Me.cmd_Bearbeiten.Visible = False
        Me.cmd_Austritt.Visible = False
        Me.cmd_Uebernehmen.Visible = False
        Me.cmd_Anlegen.Visible = True ' NEU-Button
        Me.cmd_Abbrechen.Visible = True
        
    ElseIf EditMode = True Then
        ' EDIT-MODUS
        sMode = "EDIT" ' Modus aktualisieren
        Me.cmd_Bearbeiten.Visible = False
        Me.cmd_Anlegen.Visible = False
        Me.cmd_Austritt.Visible = True ' Austritt sichtbar in Edit
        Me.cmd_Uebernehmen.Visible = True ' Übernehmen sichtbar in Edit
        Me.cmd_Abbrechen.Visible = True
        
    Else
        ' ANSICHT-MODUS (EditMode = False, sMode = ANSICHT)
        sMode = "ANSICHT" ' Modus aktualisieren
        Me.cmd_Bearbeiten.Visible = True ' Ändern sichtbar in Ansicht
        Me.cmd_Anlegen.Visible = True ' Anlegen sichtbar in Ansicht (optional, falls es ein eigener Button ist)
        Me.cmd_Austritt.Visible = False
        Me.cmd_Uebernehmen.Visible = False
        Me.cmd_Abbrechen.Visible = True
    End If
    
    ' 3. Vorbefüllen der Felder im EDIT-Modus (ANSICHT hat bereits Labels)
    If EditMode = True And sMode <> "NEU" Then
        ' Füllen der Textboxen/Comboboxen mit Werten aus den Labels/Daten
        lCurrentRow = Finde_Zeile_durch_MemberID(sMemberID)
        Dim wsM As Worksheet
        If lCurrentRow > 0 Then Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
        
        Me.cbo_Parzelle.Value = Me.lbl_Parzelle.Caption
        Me.cbo_Anrede.Value = Me.lbl_Anrede.Caption
        Me.txt_Vorname.Value = Me.lbl_Vorname.Caption
        Me.txt_Nachname.Value = Me.lbl_Nachname.Caption
        Me.txt_Strasse.Value = Me.lbl_Strasse.Caption
        Me.txt_Nummer.Value = Me.lbl_Nummer.Caption
        Me.txt_PLZ.Value = Me.lbl_PLZ.Caption
        Me.txt_Wohnort.Value = Me.lbl_Wohnort.Caption
        Me.txt_Telefon.Value = Me.lbl_Telefon.Caption
        Me.txt_Mobil.Value = Me.lbl_Mobil.Caption
        Me.txt_Geburtstag.Value = Me.lbl_Geburtstag.Caption
        Me.txt_Email.Value = Me.lbl_Email.Caption
        Me.cbo_Funktion.Value = Me.lbl_Funktion.Caption
        
        ' Laden der unsichtbaren Pachtdaten für den späteren Speichervorgang
        If Not wsM Is Nothing Then
            Me.txt_Pachtbeginn.Value = wsM.Cells(lCurrentRow, M_COL_PACHTBEGINN).Value
            Me.txt_Pachtende.Value = wsM.Cells(lCurrentRow, M_COL_PACHTENDE).Value
        End If
    End If
    
End Sub

' ***************************************************************
' Prozedur: cmd_Bearbeiten_Click ("Ändern" Button)
' ***************************************************************
Private Sub cmd_Bearbeiten_Click()
    ' Schaltet die Form in den Bearbeitungsmodus
    Call SetMode(True)
End Sub

' ***************************************************************
' Prozedur: cmd_Abbrechen_Click
' Schließt die Userform IMMER, unabhängig vom Modus
' ***************************************************************
Private Sub cmd_Abbrechen_Click()
    Unload Me
End Sub

' ***************************************************************
' Prozedur: cmd_Austritt_Click (Startet die komplexe Austrittslogik)
' ***************************************************************
Private Sub cmd_Austritt_Click()
    
    If sMemberID = "" Or sMode = "NEU" Then
        MsgBox "Interner Fehler: Keine eindeutige MemberID für den Austritt gefunden.", vbCritical
        Exit Sub
    End If
    
    Dim OldParzelle As String
    OldParzelle = Me.lbl_Parzelle.Caption ' Oder besser: Lese aus der unsichtbaren Textbox/Liste, falls geändert
    If sMode = "EDIT" Then OldParzelle = Me.cbo_Parzelle.Value

    If MsgBox("Wollen Sie den Austritt von " & Me.txt_Nachname.Value & " aus der Parzelle " & OldParzelle & " erfassen? Dies löst die Nachfolgelogik aus.", _
              vbYesNo + vbExclamation, "Austritt bestätigen") = vbNo Then
        Exit Sub
    End If

    ' AUFRUF DER LOGIK ZUR BEHANDLUNG DES AUSTRITTS (Wird in einem separaten Formular/Prozess behandelt)
    ' Simulierter Aufruf der Logik (wird meist durch ein Folge-Formular ausgelöst)
    MsgBox "Austrittslogik muss hier über ein Formular zur Erfassung von Austrittsdatum und Nachfolger gestartet werden.", vbExclamation
    ' Call frm_Austritt.Init_Austritt_Daten(sMemberID, OldParzelle, Me.txt_Nachname.Value)
    
    ' Nach simulierter Bearbeitung:
    ' Unload Me
    
End Sub


' ***************************************************************
' Prozedur: cmd_Uebernehmen_Click (Speichert die Änderungen)
' ***************************************************************
Private Sub cmd_Uebernehmen_Click()
    
    If sMemberID = "" Or sMode = "NEU" Then
        MsgBox "Fehler: MemberID fehlt oder es ist der falsche Speichermodus. Speichern abgebrochen.", vbCritical
        Exit Sub
    End If
    
    Dim wsM As Worksheet
    Dim lCurrentRow As Long
    
    On Error GoTo ErrorHandler
    
    ' 1. Aktuelle Zeile anhand der MemberID suchen (sicher gegen Sortierung)
    lCurrentRow = Finde_Zeile_durch_MemberID(sMemberID)
    
    If lCurrentRow = 0 Then
        MsgBox "Fehler: Die MemberID " & sMemberID & " konnte in der Mitgliederliste nicht gefunden werden. Speichern abgebrochen.", vbCritical
        Exit Sub
    End If

    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    ' --- 2. VALIDIERUNG ---
    If Trim(Me.txt_Nachname.Value) = "" Or Trim(Me.txt_Vorname.Value) = "" Then
        MsgBox "Nachname und Vorname dürfen nicht leer sein.", vbCritical
        Me.txt_Nachname.SetFocus
        Exit Sub
    End If
    
    ' Datumsvalidierung für Pachtbeginn (ist in Edit unsichtbar, muss aber gültig sein, wenn es befüllt ist)
    If Trim(Me.txt_Pachtbeginn.Value) <> "" And Not IsDate(Me.txt_Pachtbeginn.Value) Then
        MsgBox "Pachtbeginn muss ein gültiges Datum sein.", vbCritical
        Exit Sub
    End If
    
    ' NEU: Vorstand Validierung (KORRIGIERT: Übergibt nur MemberID und prüft den Boolean-Rückgabewert)
    If UCase(Me.cbo_Funktion.Value) = UCase(VORSTAND_STATUS) Then
        If mod_Mitglieder_UI.Check_Vorstand_Eindeutigkeit(sMemberID) = False Then
            ' Wenn False, gibt es bereits einen anderen Vorstand.
            MsgBox "Die Funktion '" & VORSTAND_STATUS & "' ist bereits einem anderen aktiven Mitglied zugewiesen." & vbCrLf & _
                   "Bitte wählen Sie eine andere Funktion.", vbCritical
            Me.cbo_Funktion.SetFocus
            Exit Sub
        End If
    End If
    
    ' --- 3. DATENSPEICHERUNG (STAMMDATEN) ---
    Call mod_Mitglieder_UI.UnprotectSheet(wsM)
    
    ' Daten in die gefundene Zeile lCurrentRow schreiben
    wsM.Cells(lCurrentRow, M_COL_PARZELLE).Value = Me.cbo_Parzelle.Value
    wsM.Cells(lCurrentRow, M_COL_ANREDE).Value = Me.cbo_Anrede.Value
    wsM.Cells(lCurrentRow, M_COL_NACHNAME).Value = Me.txt_Nachname.Value
    wsM.Cells(lCurrentRow, M_COL_VORNAME).Value = Me.txt_Vorname.Value
    wsM.Cells(lCurrentRow, M_COL_STRASSE).Value = Me.txt_Strasse.Value
    wsM.Cells(lCurrentRow, M_COL_NUMMER).Value = Me.txt_Nummer.Value
    wsM.Cells(lCurrentRow, M_COL_PLZ).Value = Me.txt_PLZ.Value
    wsM.Cells(lCurrentRow, M_COL_WOHNORT).Value = Me.txt_Wohnort.Value
    wsM.Cells(lCurrentRow, M_COL_TELEFON).Value = Me.txt_Telefon.Value
    wsM.Cells(lCurrentRow, M_COL_MOBIL).Value = Me.txt_Mobil.Value
    wsM.Cells(lCurrentRow, M_COL_GEBURTSTAG).Value = Me.txt_Geburtstag.Value
    wsM.Cells(lCurrentRow, M_COL_EMAIL).Value = Me.txt_Email.Value
    wsM.Cells(lCurrentRow, M_COL_FUNKTION).Value = Me.cbo_Funktion.Value
    
    ' Pachtdaten werden nur gespeichert, wenn sie gefüllt wurden (aus SetMode(True) geladen)
    wsM.Cells(lCurrentRow, M_COL_PACHTBEGINN).Value = Me.txt_Pachtbeginn.Value
    wsM.Cells(lCurrentRow, M_COL_PACHTENDE).Value = Me.txt_Pachtende.Value
    
    Call mod_Mitglieder_UI.ProtectSheet(wsM)
    
    ' --- 4. AUFRÄUMEN & AKTUALISIEREN ---
    Call mod_Mitglieder_UI.RefreshAllLists
    
    MsgBox "Änderungen für Mitglied " & Me.txt_Nachname.Value & " erfolgreich gespeichert.", vbInformation
    
    Unload Me
    Exit Sub
ErrorHandler:
    If Not wsM Is Nothing Then Call mod_Mitglieder_UI.ProtectSheet(wsM)
    MsgBox "Fehler beim Speichern der Änderungen: " & Err.Description, vbCritical
End Sub

' ***************************************************************
' Prozedur: cmd_Anlegen_Click (Fügt ein neues Mitglied hinzu)
' ***************************************************************
Private Sub cmd_Anlegen_Click()
    
    If sMode <> "NEU" Then
        ' Wenn der Button "Anlegen" im ANSICHT-Modus gedrückt wird, soll dies eine NEU-Anlage starten.
        Unload Me
        MsgBox "Um ein neues Mitglied anzulegen, starten Sie die Funktion bitte über das Hauptmenü.", vbExclamation
        Exit Sub
    End If
    
    ' Speichervorgang für NEU
    Dim wsM As Worksheet
    Dim lRow As Long
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    ' --- 1. VALIDIERUNG ---
    If Trim(Me.txt_Nachname.Value) = "" Or Trim(Me.txt_Vorname.Value) = "" Then
        MsgBox "Nachname und Vorname dürfen nicht leer sein.", vbCritical
        Me.txt_Nachname.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(Me.txt_Pachtbeginn.Value) Then
        MsgBox "Pachtbeginn muss ein gültiges Datum sein.", vbCritical
        Me.txt_Pachtbeginn.SetFocus
        Exit Sub
    End If
    
    ' NEU: Vorstand Validierung (KORRIGIERT: Übergibt nur MemberID ("") und prüft den Boolean-Rückgabewert)
    If UCase(Me.cbo_Funktion.Value) = UCase(VORSTAND_STATUS) Then
        ' sMemberID ist bei Neuanlage "" oder leer, was die Funktion korrekt behandelt.
        If mod_Mitglieder_UI.Check_Vorstand_Eindeutigkeit(sMemberID) = False Then
            ' Wenn False, gibt es bereits einen anderen Vorstand.
            MsgBox "Die Funktion '" & VORSTAND_STATUS & "' ist bereits einem anderen aktiven Mitglied zugewiesen." & vbCrLf & _
                   "Bitte wählen Sie eine andere Funktion.", vbCritical
            Me.cbo_Funktion.SetFocus
            Exit Sub
        End If
    End If

    ' --- 2. DATENSPEICHERUNG ---
    Call mod_Mitglieder_UI.UnprotectSheet(wsM)
    
    lRow = Application.Max(M_START_ROW, wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).Row + 1)
    
    Dim sNewID As String
    ' GUID aus mod_Mitglieder_UI holen
    sNewID = mod_Mitglieder_UI.CreateGUID()

    wsM.Cells(lRow, M_COL_MEMBER_ID).Value = sNewID
    
    ' Daten in die Spalten schreiben
    wsM.Cells(lRow, M_COL_PARZELLE).Value = Me.cbo_Parzelle.Value
    wsM.Cells(lRow, M_COL_ANREDE).Value = Me.cbo_Anrede.Value
    wsM.Cells(lRow, M_COL_NACHNAME).Value = Me.txt_Nachname.Value
    wsM.Cells(lRow, M_COL_VORNAME).Value = Me.txt_Vorname.Value
    wsM.Cells(lRow, M_COL_STRASSE).Value = Me.txt_Strasse.Value
    wsM.Cells(lRow, M_COL_NUMMER).Value = Me.txt_Nummer.Value
    wsM.Cells(lRow, M_COL_PLZ).Value = Me.txt_PLZ.Value
    wsM.Cells(lRow, M_COL_WOHNORT).Value = Me.txt_Wohnort.Value
    wsM.Cells(lRow, M_COL_TELEFON).Value = Me.txt_Telefon.Value
    wsM.Cells(lRow, M_COL_MOBIL).Value = Me.txt_Mobil.Value
    wsM.Cells(lRow, M_COL_GEBURTSTAG).Value = Me.txt_Geburtstag.Value
    wsM.Cells(lRow, M_COL_EMAIL).Value = Me.txt_Email.Value
    wsM.Cells(lRow, M_COL_FUNKTION).Value = Me.cbo_Funktion.Value
    wsM.Cells(lRow, M_COL_PACHTBEGINN).Value = Me.txt_Pachtbeginn.Value
    wsM.Cells(lRow, M_COL_PACHTENDE).Value = Me.txt_Pachtende.Value ' Ist leer
    
    Call mod_Mitglieder_UI.ProtectSheet(wsM)
    
    ' --- 3. HISTORIE & AUFRÄUMEN ---
    ' HIER MUSS EINE NEUE PROZEDUR FÜR NEUANLAGE EINGEFÜGT WERDEN (Bsp. in mod_Mitglieder_UI)
    ' Simuliert den Aufruf der Historien-Schreibfunktion:
    ' Call mod_Mitglieder_UI.SchreibeHistorie(sNewID, Me.cbo_Parzelle.Value, Me.txt_Nachname.Value, Me.txt_Pachtbeginn.Value, "", "", "Neuanlage / Pachtbeginn")
    
    Call mod_Mitglieder_UI.RefreshAllLists
    
    MsgBox "Neues Mitglied " & Me.txt_Nachname.Value & " erfolgreich angelegt.", vbInformation
    
    Unload Me
    Exit Sub
ErrorHandler:
    If Not wsM Is Nothing Then Call mod_Mitglieder_UI.ProtectSheet(wsM)
    MsgBox "Fehler beim Anlegen des neuen Mitglieds: " & Err.Description, vbCritical
End Sub

' ***************************************************************
' Prozedur: UserForm_Initialize (REDUZIERT!)
' ***************************************************************
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 1
End Sub

' ***************************************************************
' NEUE PROZEDUR: UserForm_Activate (Datenbindung hierher verschoben)
' ***************************************************************
Private Sub UserForm_Activate()
    Static bInitialRun As Boolean
    
    If Not bInitialRun Then
        Call LoadComboboxes
        bInitialRun = True
    End If
End Sub

' ***************************************************************
' NEUE PROZEDUR: Lädt ComboBox-Daten (zur Entlastung von Initialize)
' ***************************************************************
Private Sub LoadComboboxes()

    Dim wsDaten As Worksheet
    Dim i As Long
    Dim sErrorMsg As String
    
    On Error GoTo ErrorHandler ' Fängt Fehler beim Laden der Daten ab

    ' 1. Zugriff auf das Datenblatt über Konstante WS_DATEN
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN) ' Nutzt die Konstante aus mod_Const

    ' 2. cbo_Anrede befüllen (Daten!D4:D9)
    Me.cbo_Anrede.Clear
    For i = 4 To 9
        If Trim(wsDaten.Cells(i, "D").Value) <> "" Then
            Me.cbo_Anrede.AddItem wsDaten.Cells(i, "D").Value
        End If
    Next i
    
    ' 3. cbo_Funktion befüllen (Daten!B4:B12) - KORRIGIERT auf Zeile 12
    Me.cbo_Funktion.Clear
    For i = 4 To 12 ' Korrigiert von 11 auf 12
        If Trim(wsDaten.Cells(i, "B").Value) <> "" Then
            Me.cbo_Funktion.AddItem wsDaten.Cells(i, "B").Value
        End If
    Next i
    
    ' 4. cbo_Parzelle befüllen (Daten!F4:F18)
    Me.cbo_Parzelle.Clear
    Me.cbo_Parzelle.AddItem "" ' Erster Eintrag ist leer (Mitglied ohne Pacht)
    For i = 4 To 18 ' Beispielbereich
        If Trim(wsDaten.Cells(i, "F").Value) <> "" Then
            Me.cbo_Parzelle.AddItem wsDaten.Cells(i, "F").Value
        End If
    Next i
    
    Exit Sub ' Erfolgreicher Ausgang

ErrorHandler:
    ' Spezifische Fehlermeldung zur Diagnose des Problems
    sErrorMsg = "Fehler beim Laden der ComboBox-Daten." & vbCrLf & vbCrLf & _
                "Ursache:" & vbCrLf
    
    If Err.Number = 9 Then ' Subscript Out of Range (häufig bei falschem Blattnamen)
        sErrorMsg = sErrorMsg & "Das Arbeitsblatt mit dem Namen '" & WS_DATEN & "' existiert nicht, ist falsch benannt oder die Konstante WS_DATEN ist falsch."
    Else
        sErrorMsg = sErrorMsg & "Laufzeitfehler " & Err.Number & ": " & Err.Description & vbCrLf & _
                    "Dies deutet auf einen ungültigen Zellbereich (z.B. falsche Spalte oder Zeilennummer) auf dem Blatt '" & WS_DATEN & "' hin."
    End If
    
    MsgBox sErrorMsg, vbCritical
    Unload Me ' Formular entladen, um weiteren Absturz zu verhindern
End Sub

' ***************************************************************
' Hilfsfunktionen (unverändert)
' ***************************************************************
Private Function Finde_Zeile_durch_MemberID(ByVal MemberID As String) As Long
    Dim wsM As Worksheet
    Dim rngFind As Range
    Dim lastRow As Long
    
    Finde_Zeile_durch_MemberID = 0
    
    If MemberID = "" Then Exit Function
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_MEMBER_ID).End(xlUp).Row
    
    ' *** WICHTIGE ÄNDERUNG: LookAt:=xlWhole und SearchFormat:=False hinzufügen ***
    ' Die zusätzliche Angabe von SearchFormat:=False stellt sicher, dass die Suche
    ' nicht durch sichtbare Zellen eingeschränkt wird und daher auch in ausgeblendeter
    ' Spalte A die ID findet, was das Bearbeiten erst ermöglicht.
    Set rngFind = wsM.Range(wsM.Cells(M_START_ROW, M_COL_MEMBER_ID), wsM.Cells(lastRow, M_COL_MEMBER_ID)).Find( _
        What:=MemberID, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)
        
    If Not rngFind Is Nothing Then
        Finde_Zeile_durch_MemberID = rngFind.Row
    End If
End Function

Private Function IsFormLoaded(ByVal FormName As String) As Boolean
    
    Dim i As Long
    
    On Error Resume Next
    For i = 0 To VBA.UserForms.Count - 1
        If StrComp(VBA.UserForms.Item(i).Name, FormName, vbTextCompare) = 0 Then
            IsFormLoaded = True
            Exit Function
        End If
    Next i
    On Error GoTo 0
    
    IsFormLoaded = False
End Function




