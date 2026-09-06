Attribute VB_Name = "mod_Format_Dropdowns"
Option Explicit

' ***************************************************************
' MODUL: mod_Format_Dropdowns
' ZWECK: DropDown-Listen-Verwaltung für Kategorien (AF, AG, AH)
' ABGELEITET AUS: mod_Formatierung (Modularisierung)
' VERSION: 1.0 - 01.03.2026
' FUNKTIONEN:
'   - AktualisiereKategorieDropdownListen: Dropdown-Quellen aktualisieren
'   - ErstelleKategorieNamedRanges: Named Ranges erstellen/aktualisieren
' ***************************************************************

' ===============================================================
' DROPDOWN-LISTEN FÜR KATEGORIEN AKTUALISIEREN (AF + AG + AH)
' ===============================================================
Public Sub AktualisiereKategorieDropdownListen(Optional ByRef ws As Worksheet = Nothing)
    
    Dim lastRow As Long
    Dim r As Long
    Dim kategorie As String
    Dim einAus As String
    Dim dictEinnahmen As Object
    Dim dictAusgaben As Object
    Dim key As Variant
    Dim nextRowE As Long
    Dim nextRowA As Long
    Dim eigenePerioden As Object
    Dim periodenZeile As Long
    Dim abrechnungsjahr As Long
    Dim periodenJahr As Long
    Dim standardPeriode As Variant
    
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    Set dictEinnahmen = CreateObject("Scripting.Dictionary")
    Set dictAusgaben = CreateObject("Scripting.Dictionary")
    Set eigenePerioden = CreateObject("Scripting.Dictionary")
    eigenePerioden.CompareMode = vbTextCompare

    For r = 4 To 1000
        If Trim$(CStr(ws.Cells(r, DATA_COL_MONAT_PERIODE).value)) <> "" Then
            eigenePerioden(Trim$(CStr(ws.Cells(r, DATA_COL_MONAT_PERIODE).value))) = True
        End If
    Next r
    
    lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub
    
    For r = DATA_START_ROW To lastRow
        kategorie = Trim(ws.Cells(r, DATA_CAT_COL_KATEGORIE).value)
        einAus = UCase(Trim(ws.Cells(r, DATA_CAT_COL_EINAUS).value))
        
        If kategorie <> "" Then
            If einAus = "E" Then
                If Not dictEinnahmen.exists(kategorie) Then
                    dictEinnahmen.Add kategorie, kategorie
                End If
            ElseIf einAus = "A" Then
                If Not dictAusgaben.exists(kategorie) Then
                    dictAusgaben.Add kategorie, kategorie
                End If
            End If
        End If
    Next r
    
    On Error Resume Next
    ws.Range("AF4:AF1000").ClearContents
    ws.Range("AG4:AG1000").ClearContents
    ws.Range("AH4:AH1000").ClearContents
    On Error GoTo 0
    
    nextRowE = 4
    For Each key In dictEinnahmen.keys
        ws.Cells(nextRowE, DATA_COL_KAT_EINNAHMEN).value = key
        nextRowE = nextRowE + 1
    Next key
    
    nextRowA = 4
    For Each key In dictAusgaben.keys
        ws.Cells(nextRowA, DATA_COL_KAT_AUSGABEN).value = key
        nextRowA = nextRowA + 1
    Next key
    
    ws.Cells(3, DATA_COL_MONAT_PERIODE).value = "Monat/Periode"
    ws.Cells(4, DATA_COL_MONAT_PERIODE).value = "Januar"
    ws.Cells(5, DATA_COL_MONAT_PERIODE).value = "Februar"
    ws.Cells(6, DATA_COL_MONAT_PERIODE).value = "M" & ChrW(228) & "rz"
    ws.Cells(7, DATA_COL_MONAT_PERIODE).value = "April"
    ws.Cells(8, DATA_COL_MONAT_PERIODE).value = "Mai"
    ws.Cells(9, DATA_COL_MONAT_PERIODE).value = "Juni"
    ws.Cells(10, DATA_COL_MONAT_PERIODE).value = "Juli"
    ws.Cells(11, DATA_COL_MONAT_PERIODE).value = "August"
    ws.Cells(12, DATA_COL_MONAT_PERIODE).value = "September"
    ws.Cells(13, DATA_COL_MONAT_PERIODE).value = "Oktober"
    ws.Cells(14, DATA_COL_MONAT_PERIODE).value = "November"
    ws.Cells(15, DATA_COL_MONAT_PERIODE).value = "Dezember"

    periodenZeile = 16
    ws.Cells(periodenZeile, DATA_COL_MONAT_PERIODE).value = "jährlich"
    periodenZeile = periodenZeile + 1
    abrechnungsjahr = HoleAbrechnungsjahr()
    If abrechnungsjahr <= 0 Then abrechnungsjahr = Year(Date)
    For periodenJahr = abrechnungsjahr - 1 To abrechnungsjahr + 1
        For Each standardPeriode In Array("Endabrechnung " & periodenJahr, "Pacht " & periodenJahr, _
                                         "Fixkosten " & periodenJahr, "Q1 " & periodenJahr, "Q2 " & periodenJahr, _
                                         "Q3 " & periodenJahr, "Q4 " & periodenJahr, "H1 " & periodenJahr, "H2 " & periodenJahr)
            If not eigenePerioden.exists(CStr(standardPeriode)) Then
                ws.Cells(periodenZeile, DATA_COL_MONAT_PERIODE).value = CStr(standardPeriode)
                periodenZeile = periodenZeile + 1
            End If
        Next standardPeriode
    Next periodenJahr

    For Each key In eigenePerioden.keys
        If periodenZeile <= 1000 Then
            If InStr(1, "|Januar|Februar|M" & ChrW(228) & "rz|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember|", _
                     "|" & CStr(key) & "|", vbTextCompare) = 0 Then
                ws.Cells(periodenZeile, DATA_COL_MONAT_PERIODE).value = CStr(key)
                periodenZeile = periodenZeile + 1
            End If
        End If
    Next key
    
    Call ErstelleKategorieNamedRanges(ws, nextRowE - 1, nextRowA - 1, periodenZeile)
    
    Call mod_Format_Spalten.FormatiereSingleSpalte(ws, 32, True)  ' AF
    Call mod_Format_Spalten.FormatiereSingleSpalte(ws, 33, True)  ' AG
    Call mod_Format_Spalten.FormatiereSingleSpalte(ws, 34, True)  ' AH
    
End Sub

' ===============================================================
' NAMED RANGES FÜR KATEGORIEN ERSTELLEN
' ===============================================================
Private Sub ErstelleKategorieNamedRanges(ByRef ws As Worksheet, _
                                         ByVal lastRowE As Long, _
                                         ByVal lastRowA As Long, _
                                         ByVal periodenZeile As Long)
    
    On Error Resume Next
    
    ThisWorkbook.names("lst_KategorienEinnahmen").Delete
    ThisWorkbook.names("lst_KategorienAusgaben").Delete
    ThisWorkbook.names("lst_MonatPeriode").Delete
    
    If lastRowE >= 4 Then
        ThisWorkbook.names.Add Name:="lst_KategorienEinnahmen", _
            RefersTo:="=" & ws.Name & "!$AF$4:$AF$" & lastRowE
    Else
        ThisWorkbook.names.Add Name:="lst_KategorienEinnahmen", _
            RefersTo:="=" & ws.Name & "!$AF$4"
    End If
    
    If lastRowA >= 4 Then
        ThisWorkbook.names.Add Name:="lst_KategorienAusgaben", _
            RefersTo:="=" & ws.Name & "!$AG$4:$AG$" & lastRowA
    Else
        ThisWorkbook.names.Add Name:="lst_KategorienAusgaben", _
            RefersTo:="=" & ws.Name & "!$AG$4"
    End If
    
    ThisWorkbook.names.Add Name:="lst_MonatPeriode", _
        RefersTo:="=" & ws.Name & "!$AH$4:$AH$" & Application.Max(15, periodenZeile - 1)
    
    On Error GoTo 0
    
End Sub


















































































































































