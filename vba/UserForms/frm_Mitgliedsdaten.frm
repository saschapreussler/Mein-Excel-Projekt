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
        If ws.Cells(r, M_COL_FUNKTION).Value = funktion And _
           ws.Cells(r, M_COL_PARZELLE).Value <> ausschlussParzelle And _
           ws.Cells(r, M_COL_PARZELLE).Value <> "" Then
            FunktionExistiertBereits = True
            Exit Function
        End If
    Next r
    
    FunktionExistiertBereits = False
End Function

' ***************************************************************
' HILFSPROZEDUR: Setzt den Anzeigemodus der Form
' ***************************************************************
Public Sub SetMode(ByVal EditMode As Boolean, Optional ByVal IsNewEntry As Boolean = False, Optional ByVal IsRemovalMode As Boolean = False)
    
    Dim ctl As MSForms.Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is MSForms.Label And Left(ctl.Name, 4) = "lbl_" Then
            ctl.Visible = Not EditMode
        ElseIf TypeOf ctl Is MSForms.TextBox Or TypeOf ctl Is MSForms.ComboBox Then
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
        Me.cmd_Bearbeiten.Visible = True
        Me.cmd_Entfernen.Visible = True
        Me.cmd_Uebernehmen.Visible = False
        Me.cmd_Anlegen.Visible = False
        Me.cmd_Abbrechen.Visible = False
    End If
    
    If EditMode = False Then Exit Sub
    
    If CStr(Me.Tag) <> "NEU" Then
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
        Me.txt_Pachtanfang.Value = Me.lbl_Pachtanfang.Caption
        Me.txt_Pachtende.Value = Me.lbl_Pachtende.Caption
    ElseIf IsNewEntry Then
        Me.txt_Pachtanfang.Value = Format(Date, "dd.mm.yyyy")
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

Private Sub cmd_Entfernen_Click()
    
    Dim lRow As Long
    Dim Nachname As String
    Dim OldParzelle As String
    Dim AustrittsDatum As Date
    Dim NewParzelleNr As String
    Dim NewMemberID As String
    Dim ChangeReason As String
    Dim AustrittAustritt As VbMsgBoxResult
    Dim NewParzelleVal As String
    Dim AustrittsDatumStr As String
    
    If Not IsNumeric(Me.Tag) Or CLng(Me.Tag) < M_START_ROW Then
        MsgBox "Interner Fehler: Keine gültige Zeilennummer für das Entfernen gefunden.", vbCritical
        Exit Sub
    End If
    
    lRow = CLng(Me.Tag)
    Nachname = Me.lbl_Nachname.Caption
    OldParzelle = Me.lbl_Parzelle.Caption
    
    AustrittAustritt = MsgBox("Wählen Sie den Grund für die Änderung:" & vbCrLf & vbCrLf & _
                                "Ja = Parzellenwechsel (Mitglied behält Mitgliedschaft, bekommt neue Parzelle)" & vbCrLf & _
                                "Nein = Austritt (Mitglied gibt Parzelle ab und tritt aus)", vbYesNo + vbQuestion, "Parzellenwechsel oder Austritt?")
    
    If AustrittAustritt = vbYes Then
        NewParzelleVal = InputBox("Bitte geben Sie die neue Parzellennummer ein:", "Neue Parzelle")
        NewParzelleVal = Trim(NewParzelleVal)
        
        If NewParzelleVal = "" Then
            Exit Sub
        End If
        
        If NewParzelleVal = OldParzelle Then
            MsgBox "Die neue Parzelle darf nicht identisch mit der alten sein.", vbExclamation
            Exit Sub
        End If
        
        AustrittsDatum = Date
        NewParzelleNr = NewParzelleVal
        NewMemberID = ""
        ChangeReason = "Parzellenwechsel"
        
    Else
        AustrittsDatumStr = InputBox("Bitte geben Sie das Austrittsdatum ein (z.B. 31.12.2025):", "Austrittsdatum")
        AustrittsDatumStr = Trim(AustrittsDatumStr)
        
        If AustrittsDatumStr = "" Then
            Exit Sub
        End If
        
        If Not IsDate(AustrittsDatumStr) Then
            MsgBox "Bitte ein gültiges Datum eingeben (z.B. 31.12.2025).", vbExclamation
            Exit Sub
        End If
        
        AustrittsDatum = CDate(AustrittsDatumStr)
        NewParzelleNr = ""
        NewMemberID = ""
        ChangeReason = "Austritt aus Parzelle"
    End If
    
    Call mod_Mitglieder_UI.Speichere_Historie_und_Aktualisiere_Mitgliederliste( _
         lRow, OldParzelle, "", Nachname, AustrittsDatum, NewParzelleNr, NewMemberID, ChangeReason)
    
    Unload Me
    
End Sub

Private Sub cmd_Uebernehmen_Click()
    
    If Not IsNumeric(Me.Tag) Or CLng(Me.Tag) < M_START_ROW Then
        MsgBox "Interner Fehler: Keine gültige Zeilennummer für das Speichern gefunden.", vbCritical
        Exit Sub
    End If
    
    Dim lRow As Long
    Dim wsM As Worksheet
    Dim autoSeite As String
    
    On Error GoTo ErrorHandler
    
    lRow = CLng(Me.Tag)
    Set wsM = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    
    If Me.txt_Nachname.Value = "" Or Me.txt_Vorname.Value = "" Then
        MsgBox "Nachname und Vorname dürfen nicht leer sein.", vbCritical
        Exit Sub
    End If
    
    If Me.txt_Pachtanfang.Value <> "" Then
        If Not IsDate(Me.txt_Pachtanfang.Value) Then
            MsgBox "Pachtanfang: Bitte ein gültiges Datum eingeben.", vbExclamation
            Exit Sub
        End If
    End If
    
    If Me.txt_Pachtende.Value <> "" Then
        If Not IsDate(Me.txt_Pachtende.Value) Then
            MsgBox "Pachtende: Bitte ein gültiges Datum eingeben.", vbExclamation
            Exit Sub
        End If
    End If

    wsM.Unprotect PASSWORD:=PASSWORD
    
    autoSeite = GetSeiteFromParzelle(Me.cbo_Parzelle.Value)
    
    wsM.Cells(lRow, M_COL_PARZELLE).Value = Me.cbo_Parzelle.Value
    wsM.Cells(lRow, M_COL_SEITE).Value = autoSeite
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
    
    If Me.txt_Pachtanfang.Value <> "" Then
        wsM.Cells(lRow, M_COL_PACHTANFANG).Value = CDate(Me.txt_Pachtanfang.Value)
        wsM.Cells(lRow, M_COL_PACHTANFANG).NumberFormat = "dd.mm.yyyy"
    End If
    
    If Me.txt_Pachtende.Value <> "" Then
        wsM.Cells(lRow, M_COL_PACHTENDE).Value = CDate(Me.txt_Pachtende.Value)
        wsM.Cells(lRow, M_COL_PACHTENDE).NumberFormat = "dd.mm.yyyy"
    End If
    
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Call mod_Mitglieder_UI.Sortiere_Mitgliederliste_Nach_Parzelle
    Call mod_Mitglieder_UI.Fuelle_MemberIDs_Wenn_Fehlend
    
    If IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    MsgBox "Änderungen für Mitglied " & Me.txt_Nachname.Value & " erfolgreich gespeichert.", vbInformation
    
    Unload Me
    Exit Sub
ErrorHandler:
    If Not wsM Is Nothing Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler beim Speichern der Änderungen: " & Err.Description, vbCritical
End Sub

Private Sub cmd_Anlegen_Click()
    Dim wsM As Worksheet
    Dim lRow As Long
    Dim autoSeite As String
    Dim funktion As String
    Dim antwort As VbMsgBoxResult
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    
    If Me.txt_Nachname.Value = "" Or Me.txt_Vorname.Value = "" Then
        MsgBox "Nachname und Vorname dürfen nicht leer sein.", vbCritical
        Exit Sub
    End If
    If Me.cbo_Parzelle.Value = "" Then
        MsgBox "Die Parzelle muss gesetzt sein.", vbCritical
        Exit Sub
    End If
    
    If Me.txt_Pachtanfang.Value = "" Then
        MsgBox "Pachtanfang: Das Datum muss festgelegt werden.", vbExclamation
        Exit Sub
    End If
    
    If Not IsDate(Me.txt_Pachtanfang.Value) Then
        MsgBox "Pachtanfang: Bitte ein gültiges Datum eingeben.", vbExclamation
        Exit Sub
    End If
    
    If Me.txt_Pachtende.Value <> "" Then
        If Not IsDate(Me.txt_Pachtende.Value) Then
            MsgBox "Pachtende: Bitte ein gültiges Datum eingeben.", vbExclamation
            Exit Sub
        End If
    End If
    
    funktion = Me.cbo_Funktion.Value
    If funktion = "1. Vorsitzende(r)" Or funktion = "2. Vorsitzende(r)" Then
        If FunktionExistiertBereits(funktion, "") Then
            antwort = MsgBox("Es gibt bereits einen/eine " & funktion & "!" & vbCrLf & vbCrLf & _
                           "Soll wirklich ein(e) weitere(r) " & funktion & " angelegt werden?", vbYesNo + vbExclamation, "Warnung")
            If antwort = vbNo Then Exit Sub
        End If
    End If

    wsM.Unprotect PASSWORD:=PASSWORD
    
    lRow = wsM.Cells(wsM.Rows.Count, M_COL_PARZELLE).End(xlUp).Row + 1
    
    wsM.Cells(lRow, M_COL_MEMBER_ID).Value = mod_Mitglieder_UI.CreateGUID_Public()
    
    autoSeite = GetSeiteFromParzelle(Me.cbo_Parzelle.Value)
    
    wsM.Cells(lRow, M_COL_PARZELLE).Value = Me.cbo_Parzelle.Value
    wsM.Cells(lRow, M_COL_SEITE).Value = autoSeite
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
    
    wsM.Cells(lRow, M_COL_PACHTANFANG).Value = CDate(Me.txt_Pachtanfang.Value)
    wsM.Cells(lRow, M_COL_PACHTANFANG).NumberFormat = "dd.mm.yyyy"
    
    If Me.txt_Pachtende.Value <> "" Then
        wsM.Cells(lRow, M_COL_PACHTENDE).Value = CDate(Me.txt_Pachtende.Value)
        wsM.Cells(lRow, M_COL_PACHTENDE).NumberFormat = "dd.mm.yyyy"
    End If
    
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Call mod_Mitglieder_UI.Sortiere_Mitgliederliste_Nach_Parzelle
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
    If IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    MsgBox "Neues Mitglied " & Me.txt_Nachname.Value & " erfolgreich angelegt.", vbInformation
    
    Unload Me
    Exit Sub
ErrorHandler:
    If Not wsM Is Nothing Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler beim Anlegen des neuen Mitglieds: " & Err.Description, vbCritical
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 1
    
    Me.cbo_Anrede.RowSource = "Daten!D4:D9"
    Me.cbo_Funktion.RowSource = "Daten!B4:B11"
    Me.cbo_Parzelle.RowSource = "Daten!F4:F18"
    
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
