Attribute VB_Name = "mod_Vereinskasse_Filter"
Option Explicit

' ===============================================================
' MODUL: mod_Vereinskasse_Filter
' VERSION: 1.0 - 18.04.2026
' ZWECK: Monatsfilter fuer Vereinskasse (analog Bankkonto)
'        - ComboBox erstellen (cbo_MonatFilter_VK)
'        - Filterlogik fuer Spalte A (Datum), Daten ab Zeile 27
'        - Kontostand-Formel in C24
' ===============================================================


' ===============================================================
' 1. COMBOBOX ERSTELLEN (falls noch nicht vorhanden)
'    Wird bei Workbook_Open aufgerufen
' ===============================================================
Public Sub InitialisiereVereinskasseComboBox()
    Dim wsVK As Worksheet
    
    On Error Resume Next
    Set wsVK = ThisWorkbook.Worksheets(WS_VEREINSKASSE)
    On Error GoTo 0
    
    If wsVK Is Nothing Then Exit Sub
    
    ' Pruefen ob ComboBox bereits existiert
    Dim oleObj As OLEObject
    Dim cbExists As Boolean
    cbExists = False
    
    On Error Resume Next
    Set oleObj = wsVK.OLEObjects("cbo_MonatFilter_VK")
    If Not oleObj Is Nothing Then cbExists = True
    Err.Clear
    On Error GoTo 0
    
    If cbExists Then Exit Sub
    
    ' ComboBox erstellen
    On Error Resume Next
    wsVK.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    Set oleObj = wsVK.OLEObjects.Add( _
        ClassType:="Forms.ComboBox.1", _
        Left:=wsVK.Range("A24").Left, _
        Top:=wsVK.Range("A24").Top + 2, _
        Width:=130, _
        Height:=22)
    
    With oleObj
        .Name = "cbo_MonatFilter_VK"
        .PrintObject = False
    End With
    
    ' ComboBox-Style auf DropDownList setzen
    On Error Resume Next
    oleObj.Object.Style = 2  ' fmStyleDropDownList
    Err.Clear
    On Error GoTo 0
    
    On Error Resume Next
    wsVK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub


' ===============================================================
' 2. COMBOBOX BEFUELLEN (bei Worksheet_Activate)
' ===============================================================
Public Sub BefuelleVereinskasseComboBox(ByVal wsVK As Worksheet)
    Dim oleObj As OLEObject
    
    On Error Resume Next
    Set oleObj = wsVK.OLEObjects("cbo_MonatFilter_VK")
    On Error GoTo 0
    
    If oleObj Is Nothing Then Exit Sub
    
    Application.EnableEvents = False
    
    Dim arrMonate(1 To 13) As String
    arrMonate(1) = "ganzes Jahr"
    arrMonate(2) = "Januar"
    arrMonate(3) = "Februar"
    arrMonate(4) = "M" & ChrW(228) & "rz"
    arrMonate(5) = "April"
    arrMonate(6) = "Mai"
    arrMonate(7) = "Juni"
    arrMonate(8) = "Juli"
    arrMonate(9) = "August"
    arrMonate(10) = "September"
    arrMonate(11) = "Oktober"
    arrMonate(12) = "November"
    arrMonate(13) = "Dezember"
    
    With oleObj.Object
        .Clear
        .List = Application.Transpose(arrMonate)
        .ListIndex = 0
    End With
    
    Application.EnableEvents = True
End Sub


' ===============================================================
' 3. FILTER ANWENDEN (bei ComboBox-Change)
' ===============================================================
Public Sub WendeVereinskasseFilterAn(ByVal wsVK As Worksheet, ByVal monatsWert As String)
    Dim wsDaten As Worksheet
    Dim jahr As Long
    Dim monatsIndex As Long
    Dim letzteDatenZeile As Long
    Dim rngFilterBereich As Range
    Dim visibleCellsCount As Long
    
    On Error GoTo FilterExit
    Application.EnableEvents = False
    
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    
    On Error Resume Next
    wsVK.Unprotect PASSWORD:=PASSWORD
    On Error GoTo FilterExit
    
    ' Vorherige Filter aufheben
    On Error Resume Next
    If wsVK.AutoFilterMode Then wsVK.AutoFilterMode = False
    On Error GoTo FilterExit
    
    ' Abrechnungsjahr aus Einstellungen
    jahr = HoleAbrechnungsjahr()
    If jahr = 0 Then
        MsgBox "Fehler: Kein Abrechnungsjahr auf dem Blatt Einstellungen hinterlegt.", vbCritical
        GoTo FilterExit
    End If
    
    ' Monatsindex bestimmen
    Select Case monatsWert
        Case "ganzes Jahr": monatsIndex = 0
        Case "Januar": monatsIndex = 1
        Case "Februar": monatsIndex = 2
        Case "M" & ChrW(228) & "rz": monatsIndex = 3
        Case "April": monatsIndex = 4
        Case "Mai": monatsIndex = 5
        Case "Juni": monatsIndex = 6
        Case "Juli": monatsIndex = 7
        Case "August": monatsIndex = 8
        Case "September": monatsIndex = 9
        Case "Oktober": monatsIndex = 10
        Case "November": monatsIndex = 11
        Case "Dezember": monatsIndex = 12
        Case Else: GoTo FilterExit
    End Select
    
    ' Letzte Datenzeile ermitteln (Spalte A = Datum)
    letzteDatenZeile = wsVK.Cells(wsVK.Rows.count, VK_COL_DATUM).End(xlUp).Row
    
    If letzteDatenZeile < VK_START_ROW Then
        ' Keine Daten vorhanden
        wsVK.Range("C24").value = "Auszug: ganzes Jahr " & jahr
        GoTo FilterExit
    End If
    
    ' Filterbereich: ab Header-Zeile bis letzte Datenzeile
    Set rngFilterBereich = wsVK.Range("A" & VK_HEADER_ROW & ":A" & letzteDatenZeile)
    
    ' Anzeige aktualisieren
    wsVK.Range("C24").value = "Auszug: " & monatsWert & " " & jahr
    
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
        ' Keine sichtbaren Daten - Filter zuruecksetzen
        wsVK.ShowAllData
        wsVK.Range("C24").value = "Auszug: ganzes Jahr " & jahr & " (keine Daten f" & ChrW(252) & "r " & monatsWert & ")"
    End If

FilterExit:
    On Error Resume Next
    wsVK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
    Application.EnableEvents = True
    If Err.Number <> 0 Then Err.Clear
End Sub


' ===============================================================
' 4. KONTOSTAND-FORMEL IN C24 SETZEN
'    Analog zu Bankkonto E2
' ===============================================================
Public Sub SetzeVereinskasseFormeln(ByVal wsVK As Worksheet)
    On Error Resume Next
    wsVK.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    On Error Resume Next
    
    ' C24: Kontostand laufend mit Filter
    ' Wenn kein Monatsfilter -> Kontostand Vorjahr (Einstellungen C5)
    ' Sonst: Kontostand Vorjahr + Summe aller Vereinskasse-Buchungen bis Filtermonat
    ' Hinweis: Vereinskasse nutzt eigene Hilfszelle fuer Monatsfilter
    ' Wir verwenden dieselbe Daten!AE4 Hilfszelle wie Bankkonto
    wsVK.Range("C24").FormulaLocal = _
        "=WENN(Daten!$AE$4<=1;Einstellungen!$C$5;" & _
        "Einstellungen!$C$5+SUMMEWENNS(Vereinskasse!$B$" & VK_START_ROW & ":$B$5000;" & _
        "Vereinskasse!$A$" & VK_START_ROW & ":$A$5000;"">=""&DATUM(Einstellungen!$C$4;1;1);" & _
        "Vereinskasse!$A$" & VK_START_ROW & ":$A$5000;""<""&DATUM(Einstellungen!$C$4;Daten!$AE$4;1)))"
    
    wsVK.Range("C24").NumberFormat = "#,##0.00 " & ChrW(8364)
    
    On Error GoTo 0
    
    On Error Resume Next
    wsVK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub


