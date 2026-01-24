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
            ' Bezeichner-Labels sollen IMMER sichtbar sein (auch im EditMode)
            If ctl.Name = "lbl_PachtbeginnBezeichner" Or ctl.Name = "lbl_PachtendeBezeichner" Then
                ctl.Visible = True
            Else
                ' Alle anderen Labels: unsichtbar im EditMode, sichtbar im ViewMode
                ctl.Visible = Not EditMode
            End If
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
    Dim OldParzelle As String
    Dim AustrittsDatum As Date
    Dim NewParzelleNr As String
    Dim NewMemberID As String
    Dim ChangeReason As String
    Dim AustrittAustritt As VbMsgBoxResult
    Dim pachtEndeVal As String
    
    If Not IsNumeric(Me.Tag) Or CLng(Me.Tag) < M_START_ROW Then
        MsgBox "Interner Fehler: Keine gültige Zeilennummer für das Entfernen gefunden.", vbCritical
        Exit Sub
    End If
    
    lRow = CLng(Me.Tag)
    Nachname = Me.lbl_Nachname.Caption
    OldParzelle = Me.lbl_Parzelle.Caption
    
    ' Prüfe ob Pachtende bereits gefüllt ist
    pachtEndeVal = Trim(Me.lbl_Pachtende.Caption)
    
    AustrittAustritt = MsgBox("Wählen Sie den Grund für die Änderung:" & vbCrLf & vbCrLf & _
                                "Ja = Parzellenwechsel (Mitglied behält Mitgliedschaft, bekommt neue Parzelle)" & vbCrLf & _
                                "Nein = Austritt (Mitglied gibt Parzelle ab und tritt aus)", vbYesNo + vbQuestion, "Parzellenwechsel oder Austritt?")
    
    If AustrittAustritt = vbYes Then
        ' PARZELLENWECHSEL - Im Edit-Modus die neue Parzelle eingeben
        Call SetMode(True, False, False)
        MsgBox "Bitte geben Sie die neue Parzellennummer in das Feld 'Parzelle' ein und klicken Sie dann 'Übernehmen'.", vbInformation, "Parzellenwechsel"
        Exit Sub
        
    Else
        ' AUSTRITT - Im Edit-Modus das Austrittsdatum eingeben
        If pachtEndeVal = "" Then
            ' Pachtende ist noch leer - Benutzer kann es eintragen
            Call SetMode(True, False, False)
            Me.txt_Pachtende.value = Format(Date, "dd.mm.yyyy")
            MsgBox "Bitte bestätigen Sie das Austrittsdatum (oder ändern Sie es) und klicken Sie dann 'Übernehmen'.", vbInformation, "Austrittsdatum"
            Exit Sub
        Else
            ' Pachtende ist bereits gesetzt - einfach als Austritt eintragen
            AustrittsDatum = CDate(pachtEndeVal)
        End If
        
        NewParzelleNr = ""
        NewMemberID = ""
        ChangeReason = "Austritt aus Parzelle"
        
        Call mod_Mitglieder_UI.Speichere_Historie_und_Aktualisiere_Mitgliederliste( _
             lRow, OldParzelle, "", Nachname, AustrittsDatum, NewParzelleNr, NewMemberID, ChangeReason)
        
        ' Formatierung neu anwenden
        Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
        
        Unload Me
    End If
    
End Sub

Private Sub cmd_Uebernehmen_Click()
    
    If Not IsNumeric(Me.Tag) Or CLng(Me.Tag) < M_START_ROW Then
        MsgBox "Interner Fehler: Keine gültige Zeilennummer für das Speichern gefunden.", vbCritical
        Exit Sub
    End If
    
    Dim lRow As Long
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
    
    lRow = CLng(Me.Tag)
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
        If Me.txt_Pachtbeginn.value = "" Then
            MsgBox "Für diese Funktion ist ein Pachtbeginn erforderlich.", vbCritical
            Exit Sub
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
                    ' - "Mitglied mit Pacht" (explizit)
                    ' - "1. Vorsitzende(r)" (Vorstand = immer mit Pacht)
                    ' - "2. Vorsitzende(r)" (Vorstand = immer mit Pacht)
                    ' - "Kassierer(in)" (Vorstand = immer mit Pacht)
                    ' - "Schriftführer(in)" (Vorstand = immer mit Pacht)
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
    Me.cbo_Funktion.RowSource = "Daten!B4:B11"
    
    ' Fuelle cbo_Parzelle OHNE "Verein"
    Call FuelleParzelleComboDB
    
    ' Setze default Captions für die Label-Bezeichner IMMER
    Me.lbl_PachtbeginnBezeichner.Caption = "Pachtbeginn"
    Me.lbl_PachtendeBezeichner.Caption = "Pachtende"
    
    ' Rufe SetMode auf, um die Form zu initialisieren
    ' Für neue Mitglieder (Tag = "NEU") wird EditMode direkt gesetzt
    If CStr(Me.Tag) = "NEU" Then
        Call SetMode(True, True, False)
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Fehler beim Initialisieren der Form: " & Err.Description, vbCritical
End Sub

' ***************************************************************
' NEUE HILFSPROZEDUR: FuelleParzelleComboDB
' Füllt die Parzelle ComboBox mit allen Werten AUßER "Verein"
' ***************************************************************
Private Sub FuelleParzelleComboDB()
    Dim ws As Worksheet
    Dim lRow As Long
    Dim parzelleValue As String
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    If ws Is Nothing Then
        ' Fallback: Nutze die Original-RowSource minus "Verein"
        Me.cbo_Parzelle.RowSource = "Daten!F4:F18"
        Exit Sub
    End If
    
    ' Leere die ComboBox zuerst
    Me.cbo_Parzelle.Clear
    
    ' Lese alle Werte von F4:F18 und füge sie hinzu, AUSSER "Verein"
    For lRow = 4 To 18
        parzelleValue = Trim(ws.Cells(lRow, 6).value)
        
        ' Überspringe leere Zellen und "Verein"
        If parzelleValue <> "" And UCase(parzelleValue) <> "VEREIN" Then
            Me.cbo_Parzelle.AddItem parzelleValue
        End If
    Next lRow
    
    Exit Sub
ErrorHandler:
    ' Fallback bei Fehler: Nutze RowSource
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

