VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Austrittsauswahl 
   Caption         =   "Austrittsgrund"
   ClientHeight    =   2380
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4540
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

Public Property Get SelectedOption() As Integer
    SelectedOption = m_SelectedOption
End Property

Public Property Get CustomReason() As String
    CustomReason = m_CustomReason
End Property

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 1
    Me.Caption = "Austrittsgrund wählen"
    
    Me.opt_Nachpaechter.value = False
    Me.opt_Tod.value = False
    Me.opt_Kuendigung.value = False
    Me.opt_Parzellenwechsel.value = False
    Me.opt_Sonstiges.value = False
    
    Me.txt_Sonstiges.Visible = False
    Me.txt_Sonstiges.value = ""
    
    m_SelectedOption = 0
    m_CustomReason = ""
End Sub

Private Sub opt_Sonstiges_Click()
    Me.txt_Sonstiges.Visible = Me.opt_Sonstiges.value
    If Me.opt_Sonstiges.value Then
        Me.txt_Sonstiges.SetFocus
    End If
End Sub

Private Sub opt_Nachpaechter_Click()
    Me.txt_Sonstiges.Visible = False
End Sub

Private Sub opt_Tod_Click()
    Me.txt_Sonstiges.Visible = False
End Sub

Private Sub opt_Kuendigung_Click()
    Me.txt_Sonstiges.Visible = False
End Sub

Private Sub opt_Parzellenwechsel_Click()
    Me.txt_Sonstiges.Visible = False
End Sub

Private Sub cmd_OK_Click()
    If Me.opt_Nachpaechter.value Then
        m_SelectedOption = 1
        m_CustomReason = "Übergabe an Nachpächter"
    ElseIf Me.opt_Tod.value Then
        m_SelectedOption = 2
        m_CustomReason = "Tod des Mitglieds"
    ElseIf Me.opt_Kuendigung.value Then
        m_SelectedOption = 3
        m_CustomReason = "Kündigung"
    ElseIf Me.opt_Parzellenwechsel.value Then
        m_SelectedOption = 4
        m_CustomReason = "Parzellenwechsel"
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
    Me.Hide
End Sub
