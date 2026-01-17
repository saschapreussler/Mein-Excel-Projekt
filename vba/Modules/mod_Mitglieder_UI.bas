Attribute VB_Name = "mod_Mitglieder_UI"
Option Explicit

' ***************************************************************
' PROZEDUR: AktualisiereDatenstand
' ***************************************************************
Public Sub AktualisiereDatenstand()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = Worksheets(WS_MITGLIEDER)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ws.Unprotect PASSWORD:=PASSWORD
        With ws.Cells(M_STAND_ROW, M_STAND_COL)
            .Value = Now
        End With
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Else
        Debug.Print "Fehler: Tabellenblatt '" & WS_MITGLIEDER & "' nicht gefunden."
    End If
End Sub

' ***************************************************************
' PROZEDUR: Fuelle_MemberIDs_Wenn_Fehlend
' ***************************************************************
Public Sub Fuelle_MemberIDs_Wenn_Fehlend()

    Dim wsM As Worksheet
    Dim lastRow As Long
    Dim lRow As Long
    Dim wasProtected As Boolean
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If wsM Is Nothing Then Exit Sub
    
    wasProtected = wsM.ProtectContents
    If wasProtected Then wsM.Unprotect PASSWORD:=PASSWORD
    
    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
    
    If lastRow < M_START_ROW Then GoTo Cleanup
    
    Application.ScreenUpdating = False
    
    ' Header setzen
    wsM.Cells(M_HEADER_ROW, M_COL_MEMBER_ID).Value = "Member ID"
    
    ' Schleife durch alle Zeilen ab M_START_ROW
    For lRow = M_START_ROW To lastRow
        ' Prüfen, ob eine MemberID fehlt und ob der Datensatz nicht leer ist
        If wsM.Cells(lRow, M_COL_MEMBER_ID).Value = "" And _
           wsM.Cells(lRow, M_COL_NACHNAME).Value <> "" Then
            
            ' GUID generieren und eintragen
            wsM.Cells(lRow, M_COL_MEMBER_ID).Value = CreateGUID_Public()
        End If
    Next lRow
    
    ' Spalte A sperren
    With wsM.Range(wsM.Cells(M_START_ROW, M_COL_MEMBER_ID), wsM.Cells(lastRow + 1000, M_COL_MEMBER_ID))
        .Locked = True
        .FormulaHidden = True
    End With
    
Cleanup:
    Application.ScreenUpdating = True
    If wasProtected Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler beim Füllen der MemberIDs: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ***************************************************************
' HILFSFUNKTION: GUID erstellen (PUBLIC - für frm_Mitgliedsdaten zugänglich)
' ***************************************************************
Public Function CreateGUID_Public() As String
    
    On Error Resume Next
    Dim TypeLib As Object
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    CreateGUID_Public = Mid(TypeLib.GUID, 2, 36)
    
    If CreateGUID_Public = "" Then
        Randomize
        CreateGUID_Public = Format(Now, "yyyymmddhhmmss") & "-" & Int((99999 - 10000 + 1) * Rnd + 10000)
    End If
    
    Set TypeLib = Nothing
End Function

' ***************************************************************
' PROZEDUR: ApplyMitgliederDropdowns
' ***************************************************************
Public Sub ApplyMitgliederDropdowns()
    Dim ws As Worksheet
    On Error GoTo ErrorHandler
    Set ws = Worksheets(WS_MITGLIEDER)
    ws.Unprotect PASSWORD:=PASSWORD
    
    ws.Range(ws.Cells(M_START_ROW, M_COL_PARZELLE), ws.Cells(1000, M_COL_PARZELLE)).Locked = False
    ws.Range(ws.Cells(M_START_ROW, M_COL_ANREDE), ws.Cells(1000, M_COL_ANREDE)).Locked = False
    ws.Range(ws.Cells(M_START_ROW, M_COL_FUNKTION), ws.Cells(1000, M_COL_FUNKTION)).Locked = False

    Call ApplyDropdown(ws.Range(ws.Cells(M_START_ROW, M_COL_PARZELLE), ws.Cells(1000, M_COL_PARZELLE)), "=Daten!$F$4:$F$18", True)
    Call ApplyDropdown(ws.Range(ws.Cells(M_START_ROW, M_COL_SEITE), ws.Cells(1000, M_COL_SEITE)), "=Daten!$H$4:$H$6", True)
    Call ApplyDropdown(ws.Range(ws.Cells(M_START_ROW, M_COL_ANREDE), ws.Cells(1000, M_COL_ANREDE)), "=Daten!$D$4:$D$9", True)
    Call ApplyDropdown(ws.Range(ws.Cells(M_START_ROW, M_COL_FUNKTION), ws.Cells(1000, M_COL_FUNKTION)), "=Daten!$B$4:$B$11", True)

    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Exit Sub
ErrorHandler:
    If Not ws Is Nothing Then ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler beim Setzen der Dropdown-Listen: " & Err.Description, vbCritical
End Sub

Public Sub Reapply_Data_Validation()
    Call ApplyMitgliederDropdowns
End Sub

Private Sub ApplyDropdown(ByVal targetRange As Range, ByVal sourceFormula As String, ByVal allowBlanks As Boolean)
    With targetRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=sourceFormula
        .IgnoreBlank = allowBlanks
        .InCellDropdown = True
        .ErrorTitle = "Ungültiger Wert"
        .ErrorMessage = "Bitte wählen Sie einen Wert aus der Liste."
    End With
End Sub

' ***************************************************************
' PROZEDUR: Sortiere_Mitgliederliste_Nach_Parzelle
' ***************************************************************
Public Sub Sortiere_Mitgliederliste_Nach_Parzelle()

    Dim ws As Worksheet
    Dim rngSort As Range
    Dim lastRow As Long
    Dim wasProtected As Boolean
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If ws Is Nothing Then Exit Sub
    
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect PASSWORD:=PASSWORD
    
    lastRow = ws.Cells(ws.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
    
    If lastRow < M_START_ROW Then GoTo Cleanup
    
    Set rngSort = ws.Range(ws.Cells(M_START_ROW, 1), ws.Cells(lastRow, M_COL_PACHTENDE))
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Columns(M_COL_PACHTENDE), SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=ws.Columns(M_COL_PARZELLE), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .SortFields.Add Key:=ws.Columns(M_COL_ANREDE), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange rngSort
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Call Reapply_Data_Validation
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
Cleanup:
    If Not ws Is Nothing Then
        If wasProtected Then ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    Exit Sub

ErrorHandler:
    If Not ws Is Nothing Then
        If wasProtected Then ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    MsgBox "FEHLER BEIM SORTIEREN: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ***************************************************************
' FUNKTION: GetEntityKeyByParzelle
' ***************************************************************
Public Function GetEntityKeyByParzelle(ByVal ParzelleNr As String) As String
    Dim wsD As Worksheet
    Dim lastRow As Long
    Dim rngFind As Range
    
    If ParzelleNr = "" Then
        GetEntityKeyByParzelle = ""
        Exit Function
    End If
    
    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    If wsD Is Nothing Then GoTo ErrorHandler
    
    lastRow = wsD.Cells(wsD.Rows.Count, DATA_MAP_COL_PARZELLE).End(xlUp).Row
    Set rngFind = wsD.Range(wsD.Cells(DATA_START_ROW, DATA_MAP_COL_PARZELLE), wsD.Cells(lastRow, DATA_MAP_COL_PARZELLE))
    
    Set rngFind = rngFind.Find(What:=ParzelleNr, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rngFind Is Nothing Then
        GetEntityKeyByParzelle = wsD.Cells(rngFind.Row, DATA_MAP_COL_ENTITYKEY).Value
    Else
        GetEntityKeyByParzelle = ""
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Fehler in GetEntityKeyByParzelle: " & Err.Description
    GetEntityKeyByParzelle = ""
End Function

' ***************************************************************
' PROZEDUR: Speichere_Historie_und_Aktualisiere_Mitgliederliste
' ***************************************************************
Public Sub Speichere_Historie_und_Aktualisiere_Mitgliederliste( _
    ByVal selectedRow As Long, _
    ByVal OldParzelle As String, _
    ByVal OldMemberID As String, _
    ByVal Nachname As String, _
    ByVal AustrittsDatum As Date, _
    ByVal NewParzelleNr As String, _
    ByVal NewMemberID As String, _
    ByVal ChangeReason As String)

    Dim wsM As Worksheet
    Dim wsH As Worksheet
    Dim NextRow As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    
    ' --- 1. HISTORIE SPEICHERN ---
    wsH.Unprotect PASSWORD:=PASSWORD
    NextRow = wsH.Cells(wsH.Rows.Count, H_COL_PARZELLE).End(xlUp).Row + 1
    If NextRow < H_START_ROW Then NextRow = H_START_ROW
    
    wsH.Cells(NextRow, H_COL_PARZELLE).Value = OldParzelle
    wsH.Cells(NextRow, H_COL_MITGL_ID).Value = OldMemberID
    wsH.Cells(NextRow, H_COL_NACHNAME).Value = Nachname
    wsH.Cells(NextRow, H_COL_AUST_DATUM).Value = AustrittsDatum
    wsH.Cells(NextRow, H_COL_NEUER_PAECHTER_ID).Value = NewMemberID
    wsH.Cells(NextRow, H_COL_GRUND).Value = ChangeReason
    wsH.Cells(NextRow, H_COL_SYSTEMZEIT).Value = Now
    
    wsH.Cells(NextRow, H_COL_AUST_DATUM).NumberFormat = "dd.mm.yyyy"
    wsH.Cells(NextRow, H_COL_SYSTEMZEIT).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    wsH.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True

    ' --- 2. MITGLIEDERLISTE AKTUALISIEREN ---
    wsM.Unprotect PASSWORD:=PASSWORD
    
    If ChangeReason = "Parzellenwechsel" And NewParzelleNr <> "" Then
        wsM.Cells(selectedRow, M_COL_PARZELLE).Value = NewParzelleNr
    ElseIf ChangeReason = "Austritt aus Parzelle" Then
        wsM.Cells(selectedRow, M_COL_PARZELLE).Value = ""
        wsM.Cells(selectedRow, M_COL_PACHTENDE).Value = AustrittsDatum
        wsM.Cells(selectedRow, M_COL_PACHTENDE).NumberFormat = "dd.mm.yyyy"
        wsM.Cells(selectedRow, M_COL_FUNKTION).Value = AUSTRITT_STATUS
    End If
    
    Call AktualisiereDatenstand
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' --- 3. AUFRÄUMEN & AKTUALISIERUNG ---
    On Error Resume Next
    Call mod_Hilfsfunktionen.AktualisiereNamedRange_MitgliederNamen
    Call Sortiere_Mitgliederliste_Nach_Parzelle
    Call mod_Banking_Data.Aktualisiere_Parzellen_Mapping_Final
    Call mod_Banking_Data.Sortiere_Tabellen_Daten
    Call mod_ZaehlerLogik.Ermittle_Kennzahlen_Mitgliederliste
    Call mod_ZaehlerLogik.ErzeugeParzellenUebersicht
    Call mod_ZaehlerLogik.AktualisiereZaehlerTabellenSpalteA

    If IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    On Error GoTo 0
    
    MsgBox "Historien-Eintrag erfolgreich gespeichert und Mitgliederliste aktualisiert.", vbInformation
    
    Exit Sub
ErrorHandler:
    If Not wsM Is Nothing Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    If Not wsH Is Nothing Then wsH.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "FEHLER BEI DER DATENVERARBEITUNG: " & Err.Description, vbCritical
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' ***************************************************************
' HILFSFUNKTION: Prüfen, ob eine UserForm geladen ist
' ***************************************************************
Private Function IsFormLoaded(ByVal FormName As String) As Boolean
    
    Dim f As Object
    
    For Each f In VBA.UserForms
        If StrComp(f.Name, FormName, vbTextCompare) = 0 Then
            IsFormLoaded = True
            Exit Function
        End If
    Next f
    
    IsFormLoaded = False
    
End Function

