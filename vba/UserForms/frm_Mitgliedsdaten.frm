VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Mitgliedsdaten 
   Caption         =   "Mitgliedsdaten"
   ClientHeight    =   8960.001
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
' KONSTANTEN: Für die Formulareingabe (müssen mit den tatsächlichen Steuerelementen übereinstimmen)
' ***************************************************************
Private Const WS_NAME_MITGLIEDER As String = "Mitgliederliste"

' ***************************************************************
' HILFSPROZEDUR: Setzt den Anzeigemodus der Form
' ***************************************************************
Public Sub SetMode(ByVal EditMode As Boolean, Optional ByVal PreFillAddress As Boolean = False)
    
    ' Labels (Anzeigen) vs. Eingabefelder (Editieren/Anlegen) umschalten
    Dim ctl As MSForms.Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is MSForms.Label And Left(ctl.Name, 4) = "lbl_" Then
            ' Umschalten von Label (Anzeige)
            ctl.Visible = Not EditMode
        ElseIf TypeOf ctl Is MSForms.TextBox Or TypeOf ctl Is MSForms.ComboBox Then
            ' Umschalten von Eingabefeldern
            ctl.Visible = EditMode
        End If
    Next ctl
    
    ' Die Buttons je nach Modus sichtbar machen
    If CStr(Me.Tag) = "NEU" Then
        ' ANLEGE-MODUS
        Me.cmd_Bearbeiten.Visible = False
        Me.cmd_Entfernen.Visible = False
        Me.cmd_Uebernehmen.Visible = False
        Me.cmd_Anlegen.Visible = True
        Me.cmd_Abbrechen.Visible = True
        
    ElseIf EditMode = True Then
        ' BEARBEITEN-MODUS
        Me.cmd_Bearbeiten.Visible = False
        Me.cmd_Entfernen.Visible = False
        Me.cmd_Anlegen.Visible = False
        Me.cmd_Uebernehmen.Visible = True
        Me.cmd_Abbrechen.Visible = True
        
    Else
        ' ANZEIGE-MODUS (Standard beim DblClick)
        Me.cmd_Bearbeiten.Visible = True
        Me.cmd_Entfernen.Visible = True ' Entspricht dem Button "Parzellenwechsel und Mitgliedsaustritt"
        Me.cmd_Uebernehmen.Visible = False
        Me.cmd_Anlegen.Visible = False
        Me.cmd_Abbrechen.Visible = False
    End If
    
    If EditMode = False Then Exit Sub ' Alle weiteren Schritte nur im Eingabemodus
    
    ' Im Edit-Modus alle Eingabefelder mit aktuellen Werten füllen
    If CStr(Me.Tag) <> "NEU" Then
        Me.cbo_Parzelle.Value = Me.lbl_Parzelle.Caption
        Me.cbo_Seite.Value = Me.lbl_Seite.Caption
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
' ***************************************************************
Private Sub cmd_Abbrechen_Click()
    ' Bei NEU: Einfach schließen
    If CStr(Me.Tag) = "NEU" Then
        Unload Me
        Exit Sub
    End If
    
    ' Bei Bearbeiten: Zurück in den Anzeigen-Modus (Daten bleiben unverändert, wenn nicht gespeichert)
    Call SetMode(False)
End Sub

' ***************************************************************
' Prozedur: cmd_Entfernen_Click ("Entfernen" Button)
' Startet die Historien- und Parzellenwechsel-Logik
' ***************************************************************
Private Sub cmd_Entfernen_Click()
    
    Dim lRow As Long
    Dim Nachname As String
    Dim OldParzelle As String
    
    ' 1. Prüfen, ob eine Zeile zum Entfernen gespeichert ist (Tag muss die Zeilennummer enthalten)
    If Not IsNumeric(Me.Tag) Or CLng(Me.Tag) < M_START_ROW Then
        MsgBox "Interner Fehler: Keine gültige Zeilennummer für das Entfernen gefunden.", vbCritical
        Exit Sub
    End If
    
    lRow = CLng(Me.Tag)
    Nachname = Me.lbl_Nachname.Caption
    OldParzelle = Me.lbl_Parzelle.Caption
    
    ' 2. UserForm für Parzellenwechsel/Austritt initialisieren und anzeigen
    Unload Me ' Detail-Form schließen, bevor die nächste Form geöffnet wird
    
    On Error Resume Next
    ' Wir nehmen an, dass die UserForm frm_Parzellenwechsel existiert
    With frm_Parzellenwechsel
        ' Initialisiere die Form mit den Daten des betroffenen Mitglieds
        .Init_Wechsel_Daten lRow, OldParzelle, Nachname
        .Show
    End With
    On Error GoTo 0
    
    ' Hinweis: Die Aktualisierung der Mitgliederliste erfolgt am Ende der Logik in frm_Parzellenwechsel!
    
End Sub

' ***************************************************************
' Prozedur: cmd_Uebernehmen_Click (Speichert die Änderungen)
' ***************************************************************
Private Sub cmd_Uebernehmen_Click()
    
    If Not IsNumeric(Me.Tag) Or CLng(Me.Tag) < M_START_ROW Then
        MsgBox "Interner Fehler: Keine gültige Zeilennummer für das Speichern gefunden.", vbCritical
        Exit Sub
    End If
    
    Dim lRow As Long
    Dim wsM As Worksheet
    
    On Error GoTo ErrorHandler
    
    lRow = CLng(Me.Tag)
    Set wsM = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    
    ' --- 1. VALIDIERUNG ---
    If Me.txt_Nachname.Value = "" Or Me.txt_Vorname.Value = "" Then
        MsgBox "Nachname und Vorname dürfen nicht leer sein.", vbCritical
        Exit Sub
    End If
    ' Weitere notwendige Validierungen (z.B. Parzelle, Datumsformat Geburtstag) hier einfügen...

    ' --- 2. DATENSPEICHERUNG ---
    wsM.Unprotect PASSWORD:=PASSWORD ' Muss bekannt sein
    
    ' Daten in die Spalten 2 (M_COL_PARZELLE) bis 15 (M_COL_FUNKTION) schreiben
    wsM.Cells(lRow, M_COL_PARZELLE).Value = Me.cbo_Parzelle.Value
    wsM.Cells(lRow, M_COL_SEITE).Value = Me.cbo_Seite.Value
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
    
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' --- 3. AUFRÄUMEN & AKTUALISIEREN ---
    Call mod_Mitglieder_UI.Sortiere_Mitgliederliste_Nach_Parzelle ' Sortiert und formatiert neu
    
    ' Hauptformular aktualisieren (Falls es noch offen ist)
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

' ***************************************************************
' Prozedur: cmd_Anlegen_Click (Fügt ein neues Mitglied hinzu)
' ***************************************************************
Private Sub cmd_Anlegen_Click()
    Dim wsM As Worksheet
    Dim lRow As Long
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    
    ' --- 1. VALIDIERUNG ---
    If Me.txt_Nachname.Value = "" Or Me.txt_Vorname.Value = "" Then
        MsgBox "Nachname und Vorname dürfen nicht leer sein.", vbCritical
        Exit Sub
    End If
    If Me.cbo_Parzelle.Value = "" Then
        MsgBox "Die Parzelle muss gesetzt sein.", vbCritical
        Exit Sub
    End If

    ' --- 2. DATENSPEICHERUNG ---
    wsM.Unprotect PASSWORD:=PASSWORD
    
    lRow = wsM.Cells(wsM.Rows.Count, M_COL_PARZELLE).End(xlUp).Row + 1 ' Neue freie Zeile finden
    
    ' Daten in die Spalten 2 (M_COL_PARZELLE) bis 15 (M_COL_FUNKTION) schreiben
    wsM.Cells(lRow, M_COL_PARZELLE).Value = Me.cbo_Parzelle.Value
    wsM.Cells(lRow, M_COL_SEITE).Value = Me.cbo_Seite.Value
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
    
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' --- 3. AUFRÄUMEN & AKTUALISIEREN ---
    Call mod_Mitglieder_UI.Sortiere_Mitgliederliste_Nach_Parzelle
    
    ' Hauptformular aktualisieren
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

' ***************************************************************
' Prozedur: UserForm_Initialize
' ***************************************************************
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 1
    
    ' WICHTIG: RowSource-Zuweisungen für die ComboBoxen
    ' Stellen Sie sicher, dass die Konstanten für die Dropdown-Bereiche bekannt sind.
    Me.cbo_Anrede.RowSource = "Daten!D4:D9"
    Me.cbo_Funktion.RowSource = "Daten!B4:B11"
    Me.cbo_Seite.RowSource = "Daten!H4:H6"
    Me.cbo_Parzelle.RowSource = "Daten!F4:F18" ' Alle Parzellen
    
End Sub

' ***************************************************************
' HILFSFUNKTION: Prüfen, ob eine UserForm geladen ist (KORRIGIERT FÜR EXCEL)
' ***************************************************************
Private Function IsFormLoaded(ByVal FormName As String) As Boolean
    
    Dim i As Long
    
    ' Gehe alle geladenen UserForms in der VBA-Collection durch
    ' ACHTUNG: Wir nutzen VBA.UserForms, da die Forms-Collection in Excel nicht existiert.
    For i = 0 To VBA.UserForms.Count - 1
        ' Vergleiche den Namen des Formulars (Klassenname) mit dem gesuchten Namen
        If StrComp(VBA.UserForms.Item(i).Name, FormName, vbTextCompare) = 0 Then
            IsFormLoaded = True ' Formular gefunden und ist geladen
            Exit Function
        End If
    Next i
    
    IsFormLoaded = False ' Formular nicht gefunden
    
End Function


