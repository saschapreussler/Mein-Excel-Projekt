Attribute VB_Name = "mod_Einstellungen_DropDowns"
Option Explicit

' ===============================================================
' MODUL: mod_Einstellungen_DropDowns
' Ausgelagert aus mod_Einstellungen
' Enth?lt: SetzeDropDowns, HoleAlleKategorien
' FIX v2.1: Hilfsspalte f?r Fallback-DropDown = Daten!BA
' ===============================================================


' ===============================================================
' DROPDOWN-LISTEN SETZEN
'    Spalte B: Kategorie-DropDown (nur nicht-verwendete)
'    FIX v2.1: Hilfsspalte ist jetzt Daten!BA (DATA_COL_ES_HILF)
' ===============================================================
Public Sub SetzeDropDowns(ByVal ws As Worksheet)
    
    Dim lastRow As Long
    Dim nextRow As Long
    Dim r As Long
    Dim tagListe As String
    Dim toleranzListe As String
    Dim eigeneKat As String
    Dim zeilenListe As String
    
    lastRow = LetzteZeile(ws)
    nextRow = lastRow + 1
    If nextRow < ES_START_ROW Then nextRow = ES_START_ROW
    
    ' ===================================================================
    ' SPALTE B: Kategorie-DropDown (pro Zeile individuell berechnet)
    ' ===================================================================
    
    ' 1. Alle Kategorien aus Daten!J holen (dedupliziert, case-insensitive)
    Dim alleKategorien As Object
    Set alleKategorien = HoleAlleKategorien()
    
    ' 2. Alle bereits in Einstellungen!B verwendeten Kategorien sammeln
    Dim verwendete As Object
    Set verwendete = CreateObject("Scripting.Dictionary")
    verwendete.CompareMode = vbTextCompare
    
    Dim tmpKat As String
    For r = ES_START_ROW To lastRow
        tmpKat = Trim(CStr(ws.Cells(r, ES_COL_KATEGORIE).value))
        If tmpKat <> "" Then
            If Not verwendete.Exists(tmpKat) Then
                verwendete.Add tmpKat, r
            End If
        End If
    Next r
    
    ' 3. Verf?gbare Kategorien = Alle aus Daten!J MINUS bereits in Einstellungen!B verwendete
    Dim verfuegbar As Object
    Set verfuegbar = CreateObject("Scripting.Dictionary")
    verfuegbar.CompareMode = vbTextCompare
    
    Dim k As Variant
    For Each k In alleKategorien.keys
        If Not verwendete.Exists(CStr(k)) Then
            verfuegbar.Add CStr(k), True
        End If
    Next k
    
    ' 4. Basisliste als String (f?r leere Zeilen / n?chste freie Zeile)
    Dim basisListe As String
    basisListe = ""
    If verfuegbar.count > 0 Then
        basisListe = Join(verfuegbar.keys, ",")
    End If
    
    ' 5. Hilfsspalte BA auf Daten vorbereiten
    Dim wsDaten As Worksheet
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If Not wsDaten Is Nothing Then
        On Error Resume Next
        wsDaten.Unprotect PASSWORD:=PASSWORD
        On Error GoTo 0
        
        wsDaten.Range(wsDaten.Cells(1, DATA_COL_ES_HILF), _
                      wsDaten.Cells(wsDaten.Rows.count, DATA_COL_ES_HILF)).ClearContents
        
        wsDaten.Cells(DATA_HEADER_ROW, DATA_COL_ES_HILF).value = "ES-Hilf"
        
        Dim hilfZeile As Long
        hilfZeile = DATA_START_ROW
        For Each k In verfuegbar.keys
            wsDaten.Cells(hilfZeile, DATA_COL_ES_HILF).value = CStr(k)
            hilfZeile = hilfZeile + 1
        Next k
        
        wsDaten.Columns(DATA_COL_ES_HILF).Hidden = True
        wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    
    Dim hilfLetzte As Long
    hilfLetzte = hilfZeile - 1
    If hilfLetzte < DATA_START_ROW Then hilfLetzte = DATA_START_ROW
    
    ' 6. Pro Zeile das DropDown setzen
    For r = ES_START_ROW To nextRow
    
        eigeneKat = ""
        zeilenListe = ""
        
        On Error Resume Next
        ws.Cells(r, ES_COL_KATEGORIE).Validation.Delete
        On Error GoTo 0
        
        eigeneKat = Trim(CStr(ws.Cells(r, ES_COL_KATEGORIE).value))
        
        If eigeneKat = "" Then
            zeilenListe = basisListe
        Else
            If basisListe <> "" Then
                zeilenListe = eigeneKat & "," & basisListe
            Else
                zeilenListe = eigeneKat
            End If
        End If
        
        If zeilenListe <> "" Then
            If Len(zeilenListe) <= 255 Then
                On Error Resume Next
                With ws.Cells(r, ES_COL_KATEGORIE).Validation
                    .Add Type:=xlValidateList, _
                         AlertStyle:=xlValidAlertStop, _
                         Formula1:=zeilenListe
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = False
                    .ShowError = True
                End With
                If Err.Number <> 0 Then
                    Debug.Print "FEHLER Validation.Add Zeile " & r & ": " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
            Else
                If eigeneKat <> "" And Not wsDaten Is Nothing Then
                    On Error Resume Next
                    wsDaten.Unprotect PASSWORD:=PASSWORD
                    On Error GoTo 0
                    wsDaten.Cells(hilfLetzte + 1, DATA_COL_ES_HILF).value = eigeneKat
                    On Error Resume Next
                    With ws.Cells(r, ES_COL_KATEGORIE).Validation
                        .Add Type:=xlValidateList, _
                             AlertStyle:=xlValidAlertStop, _
                             Formula1:="='" & WS_DATEN & "'!$BA$" & DATA_START_ROW & ":$BA$" & (hilfLetzte + 1)
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .ShowInput = False
                        .ShowError = True
                    End With
                    If Err.Number <> 0 Then
                        Debug.Print "FEHLER Validation.Add Fallback+ Zeile " & r & ": " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo 0
                    wsDaten.Cells(hilfLetzte + 1, DATA_COL_ES_HILF).ClearContents
                    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
                ElseIf Not wsDaten Is Nothing Then
                    On Error Resume Next
                    With ws.Cells(r, ES_COL_KATEGORIE).Validation
                        .Add Type:=xlValidateList, _
                             AlertStyle:=xlValidAlertStop, _
                             Formula1:="='" & WS_DATEN & "'!$BA$" & DATA_START_ROW & ":$BA$" & hilfLetzte
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .ShowInput = False
                        .ShowError = True
                    End With
                    If Err.Number <> 0 Then
                        Debug.Print "FEHLER Validation.Add Fallback Zeile " & r & ": " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
    Next r
    
    ' ===================================================================
    ' SPALTE D: Tag 1-31
    ' ===================================================================
    tagListe = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31"
    
    For r = ES_START_ROW To nextRow
        On Error Resume Next
        ws.Cells(r, ES_COL_SOLL_TAG).Validation.Delete
        On Error GoTo 0
        With ws.Cells(r, ES_COL_SOLL_TAG).Validation
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Formula1:=tagListe
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = True
        End With
    Next r
    
    ' ===================================================================
    ' SPALTE E: KEIN DropDown
    ' ===================================================================
    For r = ES_START_ROW To nextRow
        On Error Resume Next
        ws.Cells(r, ES_COL_SOLL_MONATE).Validation.Delete
        On Error GoTo 0
    Next r
    
    ' ===================================================================
    ' SPALTE F: KEIN DropDown
    ' ===================================================================
    For r = ES_START_ROW To nextRow
        On Error Resume Next
        ws.Cells(r, ES_COL_STICHTAG_FIX).Validation.Delete
        On Error GoTo 0
    Next r
    
    ' ===================================================================
    ' SPALTE G: Vorlauf 0-31
    ' ===================================================================
    toleranzListe = "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31"
    
    For r = ES_START_ROW To nextRow
        On Error Resume Next
        ws.Cells(r, ES_COL_VORLAUF).Validation.Delete
        On Error GoTo 0
        With ws.Cells(r, ES_COL_VORLAUF).Validation
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Formula1:=toleranzListe
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = True
        End With
    Next r
    
    ' ===================================================================
    ' SPALTE H: Nachlauf 0-31
    ' ===================================================================
    For r = ES_START_ROW To nextRow
        On Error Resume Next
        ws.Cells(r, ES_COL_NACHLAUF).Validation.Delete
        On Error GoTo 0
        With ws.Cells(r, ES_COL_NACHLAUF).Validation
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Formula1:=toleranzListe
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = True
        End With
    Next r
    
End Sub


' ===============================================================
' HILFSFUNKTION: Alle Kategorien aus Daten!J holen
' ===============================================================
Public Function HoleAlleKategorien() As Object
    
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim kat As String
    Dim dict As Object
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then
        Set HoleAlleKategorien = dict
        Exit Function
    End If
    
    lastRow = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    
    For r = DATA_START_ROW To lastRow
        kat = Trim(CStr(wsDaten.Cells(r, DATA_CAT_COL_KATEGORIE).value))
        If kat <> "" Then
            If Not dict.Exists(kat) Then
                dict.Add kat, True
            End If
        End If
    Next r
    
    Set HoleAlleKategorien = dict
    
End Function


' ===============================================================
' HILFSFUNKTION: Letzte belegte Zeile in Spalte B
' ===============================================================
Private Function LetzteZeile(ByVal ws As Worksheet) As Long
    Dim lr As Long
    lr = ws.Cells(ws.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    If lr < ES_START_ROW Then lr = ES_START_ROW - 1
    LetzteZeile = lr
End Function






































