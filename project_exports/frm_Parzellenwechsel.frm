VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Parzellenwechsel 
   Caption         =   "UserForm1"
   ClientHeight    =   4360
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5800
   OleObjectBlob   =   "frm_Parzellenwechsel.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frm_Parzellenwechsel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ***************************************************************
' Globale Variablen für die Übergabe der Daten
' ***************************************************************
Private lMemberRow As Long      ' Zeile des Mitglieds in der Mitgliederliste (Muss gespeichert werden)
Private sOldParzelle As String  ' Alte Parzellennummer
Private sMemberName As String   ' Name des Mitglieds (Zur Anzeige)
Private sEntityKeyOld As String ' EntityKey (ID) des ausscheidenden Mitglieds

' ***************************************************************
' INIT-PROZEDUR: Wird von frm_Mitgliedsdaten oder mod_Mitglieder_UI aufgerufen
' ***************************************************************
Public Sub Init_Wechsel_Daten(ByVal selectedRow As Long, ByVal OldParzelle As String, ByVal Nachname As String)
    
    Dim wsD As Worksheet
    
    lMemberRow = selectedRow
    sOldParzelle = OldParzelle
    sMemberName = Nachname
    
    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    ' EntityKey des ausscheidenden Mitglieds über die Parzelle aus dem Datenblatt holen
    ' Wir nehmen an, GetEntityKeyByParzelle ist in mod_Mitglieder_UI (oder einem anderen Public Modul)
    If Not wsD Is Nothing Then
        sEntityKeyOld = mod_Mitglieder_UI.GetEntityKeyByParzelle(sOldParzelle)
    End If
    
    ' Formular-Titel und Info setzen
    Me.Caption = "Parzellenwechsel / Mitgliedsaustritt"
    Me.lbl_MitgliedInfo.Caption = "Mitglied: " & sMemberName & " gibt Parzelle " & sOldParzelle & " ab."
    
    ' Initialwerte setzen
    Me.txt_Austrittsdatum.Value = Format(Date, "dd.mm.yyyy") ' Heutiges Datum vorschlagen
    Me.opt_Austritt.Value = True ' Standard: Austritt
    Me.cbo_Nachpaechter.Value = ""
    Me.txt_Bemerkung.Value = ""
    
    ' UI anpassen
    Call UpdateUIState
    
    Me.Show
End Sub

' ***************************************************************
' PROZEDUR: Passt die Sichtbarkeit der Felder an (Austritt vs. Wechsel) (KORRIGIERT)
' ***************************************************************
Private Sub UpdateUIState()
    If Me.opt_Austritt.Value = True Then
        ' AUSTRITT/FREIGABE
        Me.cbo_NeueParzelle.Visible = False
        Me.cbo_Nachpaechter.Visible = True ' Optional: Nachpächter/Käufer protokollieren
        ' Fehler: Me.lbl_Fehler.Caption = "" <--- ENTFERNT
        Me.cmd_Uebernehmen.Caption = "Austritt protokollieren & Mitgliederliste aktualisieren"
    Else
        ' PARZELLENWECHSEL
        Me.cbo_NeueParzelle.Visible = True
        Me.cbo_Nachpaechter.Visible = False
        Me.cmd_Uebernehmen.Caption = "Wechsel protokollieren & Mitgliederliste aktualisieren"
    End If
End Sub



' ***************************************************************
' EVENTS: Wechsel zwischen Optionsbuttons
' ***************************************************************
Private Sub opt_Austritt_Click()
    Call UpdateUIState
End Sub

Private Sub opt_Wechsel_Click()
    Call UpdateUIState
End Sub

' ***************************************************************
' EVENT: Klick auf "Speichern/Übernehmen" (KORRIGIERT: Alle lbl_Fehler entfernt)
' ***************************************************************
Private Sub cmd_Uebernehmen_Click()
    
    Dim AustrittsDatum As Date
    Dim NewParzelleNr As String
    Dim ChangeReason As String
    Dim NewMemberID As String ' ID des neuen Pächters (optional)
    
    ' --- 1. VALIDIERUNG ---
    If lMemberRow < M_START_ROW Then
        ' Fehler: Me.lbl_Fehler.Caption = "Interner Fehler: Zeilennummer fehlt." <--- ENTFERNT
        MsgBox "Interner Fehler: Die Zeilennummer des Mitglieds fehlt. Vorgang abgebrochen.", vbCritical
        Exit Sub
    End If
    
    If Not IsDate(Me.txt_Austrittsdatum.Value) Then
        ' Fehler: Me.lbl_Fehler.Caption = "Bitte ein gültiges Datum eingeben (z.B. 31.12.2025)." <--- ENTFERNT
        MsgBox "Bitte ein gültiges Datum eingeben (z.B. 31.12.2025).", vbExclamation
        Exit Sub
    End If
    
    AustrittsDatum = CDate(Me.txt_Austrittsdatum.Value)
    
    If Me.opt_Wechsel.Value = True Then
        ' Validierung im WECHSEL-Modus
        If Me.cbo_NeueParzelle.Value = "" Then
            ' Fehler: Me.lbl_Fehler.Caption = "Bitte die neue Parzelle für den Wechsel auswählen." <--- ENTFERNT
            MsgBox "Bitte die neue Parzelle für den Wechsel auswählen.", vbExclamation
            Exit Sub
        End If
        If Me.cbo_NeueParzelle.Value = sOldParzelle Then
            ' Fehler: Me.lbl_Fehler.Caption = "Alte und neue Parzelle dürfen nicht identisch sein." <--- ENTFERNT
            MsgBox "Alte und neue Parzelle dürfen nicht identisch sein.", vbExclamation
            Exit Sub
        End If
        ChangeReason = "Parzellenwechsel"
        NewParzelleNr = Me.cbo_NeueParzelle.Value
        ' ACHTUNG: Der MemberID des neuen Pächters kann im Rahmen des Wechsels nicht einfach ermittelt werden.
        ' Im Normalfall bleibt der EntityKey des aktuellen Mitglieds unverändert, nur die Parzelle wird umgeschrieben.
        NewMemberID = ""
        
    Else ' opt_Austritt = True
        ' Validierung im AUSTRITT-Modus
        ChangeReason = "Austritt aus Parzelle"
        NewParzelleNr = "" ' Parzelle wird in Mitgliederliste geleert
        
        ' Hier müsste ggf. die ID des Nachpächters (cbo_Nachpaechter) ermittelt werden, falls dieser sofort einzieht.
        ' Da diese Logik komplex ist (Mapping von Name zu ID), setzen wir sie hier vorerst auf ""
        NewMemberID = ""
    End If
    
    ' Me.lbl_Fehler.Caption = "" <--- ENTFERNT (Wird durch MsgBox ersetzt)
    
    ' --- 2. DATENSPEICHERUNG ÜBER DAS MODUL ---
    ' Ruft die zentrale Logik in mod_Mitglieder_UI auf
    Call mod_Mitglieder_UI.Speichere_Historie_und_Aktualisiere_Mitgliederliste( _
         lMemberRow, _
         sOldParzelle, _
         sEntityKeyOld, _
         sMemberName, _
         AustrittsDatum, _
         NewParzelleNr, _
         NewMemberID, _
         ChangeReason)
    
    Unload Me
    
End Sub

' ***************************************************************
' EVENT: Abbruch
' ***************************************************************
Private Sub cmd_Abbrechen_Click()
    Unload Me
End Sub

' ***************************************************************
' EVENT: Formular Initialisierung
' ***************************************************************
Private Sub UserForm_Initialize()
    
    Me.StartUpPosition = 1 ' Zentriert anzeigen
    
    ' Dropdown-Zuweisungen (Müssen mit mod_Const übereinstimmen)
    
    ' Neue Parzelle für den Wechsel (alle Parzellen)
    Me.cbo_NeueParzelle.RowSource = "Daten!F4:F18"
    
    ' Nachpächter/Käufer (Hier wird eine Liste aller existierenden Mitglieder benötigt)
    ' Dies ist ein komplexes Mapping von Name -> ID, daher wird hier vorerst nur die Namensliste befüllt.
    ' Nehmen Sie an, es gibt einen benannten Bereich 'rng_MitgliederNamen'
    ' oder Sie verwenden eine Liste von Nachnamen:
    On Error Resume Next
    Me.cbo_Nachpaechter.RowSource = "rng_MitgliederNamen"
    On Error GoTo 0
    
    ' Ggf. Nachpächter-Dropdown manuell befüllen, falls RowSource fehlt
    If Me.cbo_Nachpaechter.ListCount = 0 Then
        ' Alternativ: Nur die aktuell in der Mitgliederliste geführten Nachnamen anbieten
        ' Das ist komplexer, daher die Empfehlung, einen benannten Bereich zu nutzen.
    End If
    
End Sub

