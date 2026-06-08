Attribute VB_Name = "mod_Jahreswechsel"
'==================================================================
'  PUNKT 13: Neues Kalenderjahr starten
'  - Erstellt eine Archiv-Kopie der aktuellen Mappe
'  - Loescht jahresbezogene Daten in der Original-Mappe
'  - Behaelt: Mitgliederliste, EntityKeys, Kategorien, Zaehlerstaende
'  - Setzt Abrechnungsjahr in Einstellungen +1
'==================================================================
Option Explicit

' Lokale Konstante - UEBERSICHT_START_ROW ist in mod_Uebersicht_Generator
' als Private Const deklariert und daher hier nicht sichtbar.
' Wert muss mit mod_Uebersicht_Generator / mod_Uebersicht_Filter uebereinstimmen.
Private Const UEBERSICHT_START_ROW As Long = 4


Public Sub StarteNeuesJahr()

    On Error GoTo Fehler

    Dim altJahr As Long, neuJahr As Long
    altJahr = HoleAbrechnungsjahr
    If altJahr <= 0 Then altJahr = Year(Date)
    neuJahr = altJahr + 1

    Dim antwort As VbMsgBoxResult
    antwort = MsgBox( _
        "Neues Kalenderjahr starten?" & vbCrLf & vbCrLf & _
        "Aktuelles Abrechnungsjahr: " & altJahr & vbCrLf & _
        "Neues Abrechnungsjahr:     " & neuJahr & vbCrLf & vbCrLf & _
        "Ablauf:" & vbCrLf & _
        "  1) Archiv-Kopie speichern (Original mit Jahr " & altJahr & ")" & vbCrLf & _
        "  2) Bankkonto / Vereinskasse leeren (au" & ChrW(223) & "er Okt-Dez " & altJahr & ")" & vbCrLf & _
        "  3) " & ChrW(220) & "bersicht / Dashboard / Finanz-" & ChrW(220) & "bersicht leeren" & vbCrLf & _
        "  4) Abrechnungsjahr auf " & neuJahr & " setzen" & vbCrLf & vbCrLf & _
        "Mitgliederliste, EntityKeys, Kategorien und Z" & ChrW(228) & "hlerst" & ChrW(228) & "nde bleiben erhalten." & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbYesNo + vbExclamation, "Neues Kalenderjahr")

    If antwort <> vbYes Then Exit Sub

    ' --- 1) Archiv-Kopie speichern ---
    Dim vorschlag As String
    vorschlag = "Kassenbuch_" & altJahr & "_archiv.xlsm"

    Dim ziel As Variant
    ziel = Application.GetSaveAsFilename( _
        InitialFileName:=vorschlag, _
        FileFilter:="Excel-Mappe mit Makros (*.xlsm), *.xlsm", _
        Title:="Archiv-Kopie speichern unter ...")

    If VarType(ziel) = vbBoolean Then
        MsgBox "Vorgang abgebrochen. Es wurden keine " & ChrW(196) & "nderungen vorgenommen.", vbInformation, "Abgebrochen"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    ThisWorkbook.SaveCopyAs CStr(ziel)
    Dim saveErr As Long: saveErr = Err.Number
    On Error GoTo Fehler
    If saveErr <> 0 Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        MsgBox "Archiv-Kopie konnte nicht gespeichert werden." & vbCrLf & _
               "Es wurden keine " & ChrW(196) & "nderungen vorgenommen.", vbCritical, "Fehler"
        Exit Sub
    End If

    ' --- 2) Bankkonto leeren (ausser Okt-Dez altJahr) ---
    Call LeereBankkonto(altJahr)

    ' --- 2b) Vereinskasse leeren (ausser Okt-Dez altJahr) ---
    Call LeereVereinskasse(altJahr)

    ' --- 3) Uebersicht / Dashboard / Finanz-Uebersicht leeren ---
    Call LeereBlatt(WS_UEBERSICHT(), UEBERSICHT_START_ROW)
    Call LeereBlatt("Dashboard Mitgliederzahlungen", DASH_MATRIX_START_ROW)
    Call LeereFinanzUebersicht

    ' --- 4) Abrechnungsjahr +1 ---
    Call SetzeAbrechnungsjahr(neuJahr)

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Neues Kalenderjahr erfolgreich initialisiert." & vbCrLf & vbCrLf & _
           "Archiv-Kopie:   " & CStr(ziel) & vbCrLf & _
           "Neues Jahr:     " & neuJahr & vbCrLf & vbCrLf & _
           "Bitte die Mappe jetzt speichern (STRG+S).", _
           vbInformation, "Fertig"
    Exit Sub

Fehler:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Jahreswechsel:" & vbCrLf & Err.Description, vbCritical, "Fehler"
End Sub


' ==================================================================
'  Bankkonto leeren - behaelt Datensaetze ab Okt des altJahr
' ==================================================================
Private Sub LeereBankkonto(ByVal altJahr As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim wasGeschuetzt As Boolean
    wasGeschuetzt = ws.ProtectContents
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row

    Dim r As Long
    For r = lastRow To BK_START_ROW Step -1
        Dim dat As Variant
        dat = ws.Cells(r, BK_COL_DATUM).value
        Dim behalten As Boolean: behalten = False
        If IsDate(dat) Then
            If Year(dat) = altJahr And Month(dat) >= 10 Then behalten = True
            If Year(dat) > altJahr Then behalten = True
        End If
        If Not behalten Then
            ws.Rows(r).Delete
        End If
    Next r

    If wasGeschuetzt Then
        On Error Resume Next
        ws.Protect PASSWORD:=PASSWORD, AllowFiltering:=True, _
                   DrawingObjects:=True, Contents:=True, Scenarios:=True
        On Error GoTo 0
    End If
End Sub


' ==================================================================
'  Vereinskasse leeren - behaelt Datensaetze ab Okt des altJahr
' ==================================================================
Private Sub LeereVereinskasse(ByVal altJahr As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_VEREINSKASSE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim wasGeschuetzt As Boolean
    wasGeschuetzt = ws.ProtectContents
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, VK_COL_DATUM).End(xlUp).Row

    Dim r As Long
    For r = lastRow To VK_START_ROW Step -1
        Dim dat As Variant
        dat = ws.Cells(r, VK_COL_DATUM).value
        Dim behalten As Boolean: behalten = False
        If IsDate(dat) Then
            If Year(dat) = altJahr And Month(dat) >= 10 Then behalten = True
            If Year(dat) > altJahr Then behalten = True
        End If
        If Not behalten Then
            ws.Rows(r).Delete
        End If
    Next r

    If wasGeschuetzt Then
        On Error Resume Next
        ws.Protect PASSWORD:=PASSWORD, AllowFiltering:=True, _
                   DrawingObjects:=True, Contents:=True, Scenarios:=True
        On Error GoTo 0
    End If
End Sub


' ==================================================================
'  Generisches Leeren ab einer Startzeile (Werte loeschen)
' ==================================================================
Private Sub LeereBlatt(ByVal sheetName As String, ByVal startRow As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim wasGeschuetzt As Boolean
    wasGeschuetzt = ws.ProtectContents
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0

    Dim lastRow As Long
    lastRow = ws.UsedRange.Rows.count + ws.UsedRange.Row - 1
    If lastRow >= startRow Then
        On Error Resume Next
        ws.Range(ws.Rows(startRow), ws.Rows(lastRow)).ClearContents
        On Error GoTo 0
    End If

    If wasGeschuetzt Then
        On Error Resume Next
        ws.Protect PASSWORD:=PASSWORD, AllowFiltering:=True, _
                   DrawingObjects:=True, Contents:=True, Scenarios:=True
        On Error GoTo 0
    End If
End Sub


' ==================================================================
'  Finanz-Uebersicht leeren (KPIs/Auswertungen)
' ==================================================================
Private Sub LeereFinanzUebersicht()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_FINANZ_UEBERSICHT())
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim wasGeschuetzt As Boolean
    wasGeschuetzt = ws.ProtectContents
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0

    On Error Resume Next
    ws.UsedRange.ClearContents
    On Error GoTo 0

    If wasGeschuetzt Then
        On Error Resume Next
        ws.Protect PASSWORD:=PASSWORD, AllowFiltering:=True, _
                   DrawingObjects:=True, Contents:=True, Scenarios:=True
        On Error GoTo 0
    End If
End Sub


' ==================================================================
'  Abrechnungsjahr in Einstellungen setzen
' ==================================================================
Private Sub SetzeAbrechnungsjahr(ByVal jahr As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim wasGeschuetzt As Boolean
    wasGeschuetzt = ws.ProtectContents
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0

    ws.Cells(ES_CFG_ABRECHNUNGSJAHR_ROW, ES_CFG_VALUE_COL).value = jahr

    If wasGeschuetzt Then
        On Error Resume Next
        ws.Protect PASSWORD:=PASSWORD, AllowFiltering:=True, _
                   DrawingObjects:=True, Contents:=True, Scenarios:=True
        On Error GoTo 0
    End If
End Sub




























