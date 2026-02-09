VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Austrittsauswahl 
   Caption         =   "Austrittsgrund"
   ClientHeight    =   2910
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4320
   OleObjectBlob   =   "frm_Austrittsauswahl.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frm_Austrittsauswahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_SelectedOption As Integer
Private m_CustomReason As String
Private m_NachpaechterID As String
Private m_NachpaechterName As String
Private m_AlteParzelleNr As String
Private m_AustrittsDatum As String

Public Property Get SelectedOption() As Integer
    SelectedOption = m_SelectedOption
End Property

Public Property Get CustomReason() As String
    CustomReason = m_CustomReason
End Property

Public Property Get nachpaechterID() As String
    nachpaechterID = m_NachpaechterID
End Property

Public Property Get nachpaechterName() As String
    nachpaechterName = m_NachpaechterName
End Property

Public Property Let AlteParzelleNr(ByVal value As String)
    m_AlteParzelleNr = value
End Property

Public Property Let austrittsDatum(ByVal value As String)
    m_AustrittsDatum = value
End Property

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 1
    Me.Caption = "Austrittsgrund wählen"
    
    Me.opt_Nachpaechter.value = False
    Me.opt_Tod.value = False
    Me.opt_Kuendigung.value = False
    ' ENTFERNT: Me.opt_Parzellenwechsel.value = False
    Me.opt_Sonstiges.value = False
    
    Me.txt_Sonstiges.Visible = False
    Me.txt_Sonstiges.value = ""
    
    Me.cbo_Nachpaechter.Visible = False
    Me.cbo_Nachpaechter.Clear
    
    m_SelectedOption = 0
    m_CustomReason = ""
    m_NachpaechterID = ""
    m_NachpaechterName = ""
    m_AlteParzelleNr = ""
    m_AustrittsDatum = ""
End Sub

Private Sub opt_Nachpaechter_Click()
    Me.txt_Sonstiges.Visible = False
    Me.cbo_Nachpaechter.Visible = Me.opt_Nachpaechter.value
    
    If Me.opt_Nachpaechter.value Then
        Call FuelleNachpaechterComboBox
    End If
End Sub

Private Sub opt_Tod_Click()
    Me.txt_Sonstiges.Visible = False
    Me.cbo_Nachpaechter.Visible = False
End Sub

Private Sub opt_Kuendigung_Click()
    Me.txt_Sonstiges.Visible = False
    Me.cbo_Nachpaechter.Visible = False
End Sub

' ENTFERNT: Private Sub opt_Parzellenwechsel_Click()

Private Sub opt_Sonstiges_Click()
    Me.txt_Sonstiges.Visible = Me.opt_Sonstiges.value
    Me.cbo_Nachpaechter.Visible = False
    
    If Me.opt_Sonstiges.value Then
        Me.txt_Sonstiges.SetFocus
    End If
End Sub

Private Sub cmd_OK_Click()
    If Me.opt_Nachpaechter.value Then
        ' Prüfe ob Nachpächter ausgewählt wurde
        If Trim(Me.cbo_Nachpaechter.value) = "" Then
            ' Frage ob Nachpächter bereits im System ist
            Dim antwort As VbMsgBoxResult
            antwort = MsgBox("Ist der Nachpächter bereits im System registriert?", vbYesNoCancel + vbQuestion, "Nachpächter registriert?")
            
            If antwort = vbYes Then
                ' Ja - Nachpächter muss ausgewählt werden
                MsgBox "Bitte wählen Sie den Nachpächter aus der Liste aus.", vbExclamation
                Me.cbo_Nachpaechter.SetFocus
                Exit Sub
            ElseIf antwort = vbNo Then
                ' Nein - Neuer Nachpächter muss angelegt werden
                MsgBox "Es muss ein neuer Nachpächter erfasst werden.", vbInformation, "Nachpächter erfassen"
                
                m_SelectedOption = 1
                m_CustomReason = "Übergabe an Nachpächter"
                m_NachpaechterID = "NACHPAECHTER_NEU"
                m_NachpaechterName = ""
                Me.Hide
                Exit Sub
            Else
                ' Abbrechen
                Exit Sub
            End If
        Else
            ' Nachpächter wurde ausgewählt
            m_SelectedOption = 1
            m_CustomReason = "Übergabe an Nachpächter"
            
            ' Hole Member ID aus ComboBox (versteckte Spalte)
            Dim ws As Worksheet
            Dim lRow As Long
            Dim lastRow As Long
            Dim selectedName As String
            
            Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
            selectedName = Me.cbo_Nachpaechter.value
            
            ' Suche Member ID anhand des Namens
            lastRow = ws.Cells(ws.Rows.count, M_COL_NACHNAME).End(xlUp).Row
            For lRow = M_START_ROW To lastRow
                Dim fullName As String
                fullName = ws.Cells(lRow, M_COL_NACHNAME).value & ", " & ws.Cells(lRow, M_COL_VORNAME).value
                
                If StrComp(Trim(fullName), Trim(selectedName), vbTextCompare) = 0 Then
                    m_NachpaechterID = ws.Cells(lRow, M_COL_MEMBER_ID).value
                    m_NachpaechterName = selectedName
                    Exit For
                End If
            Next lRow
        End If
        
    ElseIf Me.opt_Tod.value Then
        m_SelectedOption = 2
        m_CustomReason = "Tod des Mitglieds"
        
    ElseIf Me.opt_Kuendigung.value Then
        m_SelectedOption = 3
        m_CustomReason = "Kündigung"
        
    ' ENTFERNT: ElseIf Me.opt_Parzellenwechsel.value Then
    '     m_SelectedOption = 4
    '     m_CustomReason = "Parzellenwechsel"
        
    ElseIf Me.opt_Sonstiges.value Then
        If Trim(Me.txt_Sonstiges.value) = "" Then
            MsgBox "Bitte geben Sie einen Grund ein.", vbExclamation
            Me.txt_Sonstiges.SetFocus
            Exit Sub
        End If
        m_SelectedOption = 5
        m_CustomReason = Trim(Me.txt_Sonstiges.value)
    Else
        MsgBox "Bitte wählen Sie eine Option aus.", vbExclamation
        Exit Sub
    End If
    
    Me.Hide
End Sub

Private Sub cmd_Abbrechen_Click()
    m_SelectedOption = 0
    m_CustomReason = ""
    m_NachpaechterID = ""
    m_NachpaechterName = ""
    Me.Hide
End Sub

Private Sub FuelleNachpaechterComboBox()
    Dim ws As Worksheet
    Dim lRow As Long
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    Me.cbo_Nachpaechter.Clear
    
    lastRow = ws.Cells(ws.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    For lRow = M_START_ROW To lastRow
        ' Nur aktive Mitglieder (Pachtende leer) und nicht "Verein"
        If Trim(ws.Cells(lRow, M_COL_NACHNAME).value) <> "" And _
           Trim(ws.Cells(lRow, M_COL_PACHTENDE).value) = "" And _
           StrComp(Trim(ws.Cells(lRow, M_COL_PARZELLE).value), "Verein", vbTextCompare) <> 0 Then
            
            ' Format: "Nachname, Vorname"
            Me.cbo_Nachpaechter.AddItem ws.Cells(lRow, M_COL_NACHNAME).value & ", " & ws.Cells(lRow, M_COL_VORNAME).value
        End If
    Next lRow
End Sub

