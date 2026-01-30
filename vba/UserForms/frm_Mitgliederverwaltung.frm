VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Mitgliederverwaltung 
   Caption         =   "Mitgliederverwaltung"
   ClientHeight    =   9720.001
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   18740
   OleObjectBlob   =   "frm_Mitgliederverwaltung.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frm_Mitgliederverwaltung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_MitgliedEdit_Click()
    If Me.lst_Mitgliederliste.ListIndex >= 0 Then
        Call OeffneMitgliedsDetails
    Else
        MsgBox "Bitte wählen Sie zuerst ein Mitglied aus der Liste aus.", vbExclamation
    End If
End Sub

Private Sub cmd_NeuesMitglied_Click()
    
    With frm_Mitgliedsdaten
        .Tag = "NEU"
        .Show
    End With

End Sub

Private Sub lst_Mitgliederliste_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call OeffneMitgliedsDetails
End Sub

Public Sub OeffneMitgliedsDetails()

    If Me.lst_Mitgliederliste.ListIndex < 0 Then Exit Sub
    
    Dim ws As Worksheet
    Dim lRow As Long
    
    Set ws = ThisWorkbook.Worksheets("Mitgliederliste")
    
    Dim ParzelleToFind As String
    Dim NachnameToFind As String
    Dim VornameToFind As String
    
    ParzelleToFind = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 0)
    NachnameToFind = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 4)
    VornameToFind = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 3)
    lRow = 0
    
    Dim r As Long
    For r = 6 To ws.Cells(ws.Rows.Count, 2).End(xlUp).row
        If StrComp(Trim(ws.Cells(r, 2).value), ParzelleToFind, vbTextCompare) = 0 And _
           StrComp(Trim(ws.Cells(r, 5).value), NachnameToFind, vbTextCompare) = 0 And _
           StrComp(Trim(ws.Cells(r, 6).value), VornameToFind, vbTextCompare) = 0 Then
             lRow = r
             Exit For
        End If
    Next r

    If lRow < 6 Then
        MsgBox "Fehler: Datenzeile in der Tabelle nicht gefunden.", vbCritical
        Exit Sub
    End If

    With frm_Mitgliedsdaten
        .Tag = lRow
        
        .lbl_Parzelle.Caption = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 0)
        .lbl_Anrede.Caption = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 2)
        .lbl_Vorname.Caption = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 3)
        .lbl_Nachname.Caption = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 4)
        .lbl_Strasse.Caption = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 5)
        .lbl_Nummer.Caption = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 6)
        .lbl_PLZ.Caption = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 7)
        .lbl_Wohnort.Caption = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 8)
        .lbl_Telefon.Caption = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 9)
        .lbl_Mobil.Caption = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 10)
        .lbl_Geburtstag.Caption = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 11)
        .lbl_Email.Caption = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 12)
        .lbl_Funktion.Caption = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 13)
        
        ' Fülle auch die Pachtbeginn und Pachtende Labels
        If Me.lst_Mitgliederliste.ColumnCount > 14 Then
            .lbl_Pachtbeginn.Caption = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 14)
        Else
            .lbl_Pachtbeginn.Caption = ""
        End If
        
        If Me.lst_Mitgliederliste.ColumnCount > 15 Then
            .lbl_Pachtende.Caption = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 15)
        Else
            .lbl_Pachtende.Caption = ""
        End If
        
        .cbo_Parzelle.value = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 0)
        .cbo_Anrede.value = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 2)
        .txt_Vorname.value = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 3)
        .txt_Nachname.value = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 4)
        .txt_Strasse.value = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 5)
        .txt_Nummer.value = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 6)
        .txt_PLZ.value = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 7)
        .txt_Wohnort.value = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 8)
        .txt_Telefon.value = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 9)
        .txt_Mobil.value = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 10)
        .txt_Geburtstag.value = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 11)
        .txt_Email.value = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 12)
        .cbo_Funktion.value = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 13)
        
        ' Fülle auch die Pachtbeginn und Pachtende TextBoxen
        If Me.lst_Mitgliederliste.ColumnCount > 14 Then
            .txt_Pachtbeginn.value = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 14)
        Else
            .txt_Pachtbeginn.value = ""
        End If
        
        If Me.lst_Mitgliederliste.ColumnCount > 15 Then
            .txt_Pachtende.value = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 15)
        Else
            .txt_Pachtende.value = ""
        End If
        
    End With
    
    frm_Mitgliedsdaten.Show
    
End Sub

Private Sub LoadListBoxData()
    Dim iZeile As Long
    Dim AnzArr As Long
    Dim arr() As Variant

    lst_Mitgliederliste.ColumnCount = 16
    lst_Mitgliederliste.ColumnHeads = False
    
    AnzArr = 0
    
    With Worksheets("Mitgliederliste")
        For iZeile = 6 To .Cells(.Rows.Count, 2).End(xlUp).row
            If Trim(.Cells(iZeile, 2).value) <> "" And _
               StrComp(Trim(.Cells(iZeile, 2).value), "Verein", vbTextCompare) <> 0 And _
               Trim(.Cells(iZeile, M_COL_PACHTENDE).value) = "" Then
                AnzArr = AnzArr + 1
            End If
        Next iZeile
        
        If AnzArr > 0 Then
             ReDim arr(0 To AnzArr - 1, 0 To 15)
        Else
             lst_Mitgliederliste.Clear
             Exit Sub
        End If
        
        AnzArr = 0
        
        For iZeile = 6 To .Cells(.Rows.Count, 2).End(xlUp).row
            If Trim(.Cells(iZeile, 2).value) <> "" And _
               StrComp(Trim(.Cells(iZeile, 2).value), "Verein", vbTextCompare) <> 0 And _
               Trim(.Cells(iZeile, M_COL_PACHTENDE).value) = "" Then
                
                arr(AnzArr, 0) = .Cells(iZeile, 2).value    ' B: Parzelle
                arr(AnzArr, 1) = .Cells(iZeile, 3).value    ' C: Seite
                arr(AnzArr, 2) = .Cells(iZeile, 4).value    ' D: Anrede
                arr(AnzArr, 3) = .Cells(iZeile, 6).value    ' F: Vorname
                arr(AnzArr, 4) = .Cells(iZeile, 5).value    ' E: Nachname
                arr(AnzArr, 5) = .Cells(iZeile, 7).value    ' G: Strasse
                arr(AnzArr, 6) = .Cells(iZeile, 8).value    ' H: Nummer
                arr(AnzArr, 7) = .Cells(iZeile, 9).value    ' I: PLZ
                arr(AnzArr, 8) = .Cells(iZeile, 10).value   ' J: Wohnort
                arr(AnzArr, 9) = .Cells(iZeile, 11).value   ' K: Telefon
                arr(AnzArr, 10) = .Cells(iZeile, 12).value  ' L: Mobil
                arr(AnzArr, 11) = .Cells(iZeile, 13).value  ' M: Geburtstag
                arr(AnzArr, 12) = .Cells(iZeile, 14).value  ' N: Email
                arr(AnzArr, 13) = .Cells(iZeile, 15).value  ' O: Funktion
                arr(AnzArr, 14) = .Cells(iZeile, 16).value  ' P: Pachtbeginn
                arr(AnzArr, 15) = .Cells(iZeile, 17).value  ' Q: Pachtende
                
                AnzArr = AnzArr + 1
            End If
        Next iZeile
        
        lst_Mitgliederliste.List = arr
    End With
End Sub

Public Sub RefreshMitgliederListe()
    Call LoadListBoxData
    Call AktualisiereDatumLabel
End Sub

Private Sub UserForm_Initialize()
    Call LoadListBoxData
    Call AktualisiereDatumLabel
End Sub

' ***************************************************************
' HILFSPROZEDUR: AktualisiereDatumLabel
' Setzt das Label lbl_ListDatum mit dem Stand-Datum aus D2
' ***************************************************************
Private Sub AktualisiereDatumLabel()
    Dim ws As Worksheet
    Dim standDatum As Variant
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Mitgliederliste")
    
    If Not ws Is Nothing Then
        standDatum = ws.Cells(2, 4).value
        
        If IsDate(standDatum) Then
            Me.lbl_ListDatum.Caption = "Stand: " & Format(standDatum, "dd.mm.yyyy")
        Else
            Me.lbl_ListDatum.Caption = "Stand: " & CStr(standDatum)
        End If
    Else
        Me.lbl_ListDatum.Caption = "Stand: (unbekannt)"
    End If
    
    On Error GoTo 0
End Sub

