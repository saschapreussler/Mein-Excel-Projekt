VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub btn_CSVimport_Click()
    ' Ruft die Hauptprozedur im Standardmodul mod_Datenpflege auf
    Call Importiere_Kontoauszug
End Sub

Private Sub lst_ImportReport_Click()
End Sub

Private Sub Worksheet_Activate()
    Dim arrMonate(1 To 13) As String
    
    arrMonate(1) = "ganzes Jahr"
    arrMonate(2) = "Januar"
    arrMonate(3) = "Februar"
    arrMonate(4) = "M?rz"
    arrMonate(5) = "April"
    arrMonate(6) = "Mai"
    arrMonate(7) = "Juni"
    arrMonate(8) = "Juli"
    arrMonate(9) = "August"
    arrMonate(10) = "September"
    arrMonate(11) = "Oktober"
    arrMonate(12) = "November"
    arrMonate(13) = "Dezember"
    
    With Me.cbo_RechnungsMonat
        .Clear
        .List = Application.Transpose(arrMonate)
        .ListIndex = 0
        
        If TypeName(Me.cbo_RechnungsMonat) = "ComboBox" Then
            .Style = 2 ' fmStyleDropDownList
        End If
    End With

    Call Initialize_ImportReport_ListBox
End Sub

Private Sub cbo_RechnungsMonat_Change()

    Dim wsStart As Worksheet
    Dim wsBank As Worksheet
    Dim wsDaten As Worksheet
    
    Dim jahr As Long
    Dim monatsIndex As Long
    Dim letzteDatenZeile As Long
    
    Const ersteDatenZeile As Long = 28
    Const ersteZeileFilterbereich As Long = 27
    
    Dim rngFilterBereich As Range
    Dim visibleCellsCount As Long
    
    Set wsStart = ThisWorkbook.Sheets("Startmenü")
    Set wsBank = ThisWorkbook.Sheets("Bankkonto")
    Set wsDaten = ThisWorkbook.Sheets("Daten")

    On Error GoTo SafeExit
    Application.EnableEvents = False

    With wsBank
    
        ' Vorherige Filter immer aufheben
        On Error Resume Next
        .ShowAllData
        On Error GoTo 0
        
        If Me.cbo_RechnungsMonat.ListIndex = -1 Then GoTo SafeExit
        
        ' Abrechnungsjahr lesen
        If IsNumeric(wsStart.Range("F1").value) Then
            jahr = CLng(wsStart.Range("F1").value)
        Else
            MsgBox "Fehler: Ung?ltiges Abrechnungsjahr in Startmen?!F1", vbCritical
            GoTo SafeExit
        End If
        
        ' Monatsindex bestimmen
        Select Case Me.cbo_RechnungsMonat.value
            Case "ganzes Jahr"
                monatsIndex = 0
            Case "Januar": monatsIndex = 1
            Case "Februar": monatsIndex = 2
            Case "M?rz": monatsIndex = 3
            Case "April": monatsIndex = 4
            Case "Mai": monatsIndex = 5
            Case "Juni": monatsIndex = 6
            Case "Juli": monatsIndex = 7
            Case "August": monatsIndex = 8
            Case "September": monatsIndex = 9
            Case "Oktober": monatsIndex = 10
            Case "November": monatsIndex = 11
            Case "Dezember": monatsIndex = 12
            Case Else
                GoTo SafeExit
        End Select
        
        ' Monat in Hilfszelle AG4 schreiben
        ' 0 = ganzes Jahr, 1-12 = jeweiliger Monat
        wsDaten.Range("AG4").value = monatsIndex
        
        ' Letzte Datenzeile ermitteln
        letzteDatenZeile = .Cells(.Rows.count, "A").End(xlUp).Row
        
        If letzteDatenZeile < ersteDatenZeile Then
            ' Keine Daten vorhanden - still zuruecksetzen
            Me.cbo_RechnungsMonat.ListIndex = 0
            wsDaten.Range("AG4").value = 0
            GoTo SafeExit
        End If
        
        Set rngFilterBereich = .Range("A" & ersteZeileFilterbereich & ":A" & letzteDatenZeile)
        
        .Range("C25").value = "Auszug: " & Me.cbo_RechnungsMonat.value & " " & jahr
        
        ' Datumsgrenzen nur wenn nicht "ganzes Jahr"
        If monatsIndex > 0 Then
            Dim erstesDesMonats As Date
            Dim letztesDesMonats As Date
            
            erstesDesMonats = DateSerial(jahr, monatsIndex, 1)
            letztesDesMonats = DateSerial(jahr, monatsIndex + 1, 0)
            
            rngFilterBereich.AutoFilter Field:=1, _
                Criteria1:=">=" & CLng(erstesDesMonats), _
                Operator:=xlAnd, _
                Criteria2:="<=" & CLng(letztesDesMonats)
        End If
        
        ' Sichtbare Daten pruefen
        On Error Resume Next
        visibleCellsCount = rngFilterBereich.SpecialCells(xlCellTypeVisible).count
        On Error GoTo 0
        
        If visibleCellsCount <= 1 And monatsIndex > 0 Then
            ' Keine sichtbaren Daten - still zuruecksetzen
            .ShowAllData
            Me.cbo_RechnungsMonat.ListIndex = 0
            wsDaten.Range("AG4").value = 0
        End If
        
    End With

SafeExit:
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "Fehler in cbo_RechnungsMonat_Change: " & Err.Description, vbCritical
        Err.Clear
    End If

End Sub

