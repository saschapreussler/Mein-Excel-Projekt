Attribute VB_Name = "mod_Mitglieder_UI"
Option Explicit

' -------------------------------------------------------------------------
' Modul: mod_Mitglieder_UI
' Zweck : UI- und Mitgliedslisten-Hilfsfunktionen (Sortieren, Formatieren,
'         Dropdowns, Suche, Historie-Handling)
' Änderung: Robustere Passwort-Quelle + automatisches, stilles Entsperren/
'           Sperren (versucht mehrere Quellen für das Passwort).
' -------------------------------------------------------------------------

' Modulweite Flags
Private g_ProtectionLocked As Boolean                ' True wenn mindestens ein Blatt mit unbekanntem Passwort existiert

' ----------------------
' Passwort-Quelle: versucht mehrere Orte (const PASSWORD, definierter Name, DocProps, Named Range auf WS)
' ----------------------
Private Function GetEffectivePassword() As String
    Dim pwd As String
    Dim nm As Name
    Dim tmp As String
    On Error Resume Next

    ' 1) Konstanten-Variable PASSWORD (sofern vorhanden)
    pwd = ""
    On Error Resume Next
    pwd = PASSWORD
    On Error GoTo 0
    If Len(Trim$(pwd)) > 0 Then
        Debug.Print "GetEffectivePassword: Passwort aus Konstanten (mod_Const) gefunden, Länge=" & Len(pwd)
        GetEffectivePassword = pwd
        Exit Function
    End If

    ' 2) Namensbereich ThisWorkbook.Name "PASSWORD" (oder "Config_PASSWORD")
    On Error Resume Next
    Set nm = ThisWorkbook.Names("PASSWORD")
    If Not nm Is Nothing Then
        tmp = ""
        On Error Resume Next
        If Len(nm.RefersTo) > 0 Then
            ' Wenn Name auf Range zeigt, hole Wert; sonst entferne '=' aus RefersTo
            If Left$(nm.RefersTo, 1) = "=" Then
                tmp = ""
                On Error Resume Next
                tmp = ThisWorkbook.Names("PASSWORD").RefersToRange.Value
                On Error GoTo 0
                If Len(Trim$(tmp)) = 0 Then
                    ' fallback to text in RefersTo
                    tmp = Replace(nm.RefersTo, "=", "")
                End If
            Else
                tmp = nm.RefersTo
            End If
        End If
        If Len(Trim$(tmp)) > 0 Then
            Debug.Print "GetEffectivePassword: Passwort aus ThisWorkbook.Names('PASSWORD') gefunden, Länge=" & Len(tmp)
            GetEffectivePassword = CStr(tmp)
            Exit Function
        End If
    End If
    On Error GoTo 0

    ' 3) Suche nach alternativen Namen (z.B. Config_Password, PW, PASS)
    On Error Resume Next
    For Each nm In ThisWorkbook.Names
        If UCase(nm.Name) Like "*PASS*" Or UCase(nm.Name) Like "*PW*" Then
            tmp = ""
            On Error Resume Next
            tmp = nm.RefersToRange.Value
            On Error GoTo 0
            If Len(Trim$(tmp)) > 0 Then
                Debug.Print "GetEffectivePassword: Passwort aus Namen '" & nm.Name & "' gefunden, Länge=" & Len(tmp)
                GetEffectivePassword = CStr(tmp)
                Exit Function
            End If
        End If
    Next nm
    On Error GoTo 0

    ' 4) CustomDocumentProperties (falls vorhanden)
    On Error Resume Next
    tmp = ""
    tmp = ThisWorkbook.CustomDocumentProperties("PASSWORD").Value
    If Len(Trim$(tmp)) > 0 Then
        Debug.Print "GetEffectivePassword: Passwort aus CustomDocumentProperties gefunden, Länge=" & Len(tmp)
        GetEffectivePassword = CStr(tmp)
        Exit Function
    End If
    On Error GoTo 0

    ' 5) Konfigurationsblätter: prüfe ein Blatt namens "Konfiguration" oder "Daten" oder "Config"
    On Error Resume Next
    If SheetExists("Konfiguration") Then
        tmp = Trim(CStr(ThisWorkbook.Worksheets("Konfiguration").Range("A1").Value))
        If Len(tmp) > 0 Then
            Debug.Print "GetEffectivePassword: Passwort aus Blatt 'Konfiguration' Zelle A1 gefunden, Länge=" & Len(tmp)
            GetEffectivePassword = tmp
            Exit Function
        End If
    End If
    If SheetExists("Config") Then
        tmp = Trim(CStr(ThisWorkbook.Worksheets("Config").Range("A1").Value))
        If Len(tmp) > 0 Then
            Debug.Print "GetEffectivePassword: Passwort aus Blatt 'Config' Zelle A1 gefunden, Länge=" & Len(tmp)
            GetEffectivePassword = tmp
            Exit Function
        End If
    End If
    If SheetExists(WS_DATEN) Then
        tmp = Trim(CStr(ThisWorkbook.Worksheets(WS_DATEN).Range("A1").Value))
        If Len(tmp) > 0 Then
            Debug.Print "GetEffectivePassword: Passwort aus Blatt WS_DATEN Zelle A1 gefunden, Länge=" & Len(tmp)
            GetEffectivePassword = tmp
            Exit Function
        End If
    End If
    On Error GoTo 0

    ' keine Quelle gefunden -> leer zurückgeben
    Debug.Print "GetEffectivePassword: Kein Passwort gefunden (alle Quellen leer)"
    GetEffectivePassword = ""
End Function

Private Function SheetExists(ByVal sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sName)
    SheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

' ----------------------
' Unprotect / Protect (robust, still, keine Dialoge)
' ----------------------
Public Function UnprotectSheet(ByRef ws As Worksheet) As Boolean
    ' Versucht still, ein Worksheet zu entsperren.
    ' Liefert True bei Erfolg, False bei Fehlschlag.
    Dim pwd As String
    Dim oldAlerts As Boolean

    If ws Is Nothing Then
        UnprotectSheet = False
        Exit Function
    End If

    ' Wenn globaler Lock gesetzt ist, vermeide weitere Versuche
    If g_ProtectionLocked Then
        Debug.Print "UnprotectSheet: Abbruch für '" & ws.Name & "' wegen bekanntem Passwort-Lock."
        UnprotectSheet = False
        Exit Function
    End If

    ' Ermittle effektives Passwort (kann leer sein)
    pwd = GetEffectivePassword()

    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next

    If Len(Trim$(pwd)) > 0 Then
        ws.Unprotect PASSWORD:=pwd
    Else
        ' Benutzer hat gesagt, dass er manuell ohne Passwort auf-/abschalten kann.
        ' Wir versuchen direktes Unprotect ohne Passwort (wie manuell).
        ws.Unprotect
    End If

    On Error GoTo 0
    Application.DisplayAlerts = oldAlerts

    If ws.ProtectContents Then
        ' Konnte nicht entsperrt werden -> markiere Lock, damit wir nicht wiederholt Dialogs provozieren
        g_ProtectionLocked = True
        Debug.Print "UnprotectSheet: Blatt '" & ws.Name & "' konnte nicht entsperrt werden (Passwort fehlt/falsch). Automatische Versuche werden unterbunden."
        UnprotectSheet = False
    Else
        UnprotectSheet = True
    End If
End Function

Public Function ProtectSheet(ByRef ws As Worksheet) As Boolean
    Dim pwd As String
    Dim oldAlerts As Boolean

    If ws Is Nothing Then
        ProtectSheet = False
        Exit Function
    End If

    ' Wenn globaler Lock gesetzt ist, vermeiden wir weiteren Schutzversuch
    If g_ProtectionLocked Then
        Debug.Print "ProtectSheet: Schutzoperation für '" & ws.Name & "' übersprungen (geschütztes Blatt/Lock aktiv)."
        ProtectSheet = True
        Exit Function
    End If

    pwd = GetEffectivePassword() ' kann leer sein

    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next

    If Len(Trim$(pwd)) > 0 Then
        ws.Protect PASSWORD:=pwd, UserInterfaceOnly:=True
    Else
        ws.Protect UserInterfaceOnly:=True
    End If

    If Err.Number <> 0 Then
        Debug.Print "ProtectSheet: Konnte Blatt '" & ws.Name & "' nicht schützen: " & Err.Description
        Err.Clear
        ProtectSheet = False
    Else
        ProtectSheet = True
    End If

    On Error GoTo 0
    Application.DisplayAlerts = oldAlerts
End Function

' ----------------------
' Debug-Helfer
' ----------------------
Public Sub ListProtectedSheets()
    Dim ws As Worksheet
    Debug.Print "---- Geschützte Blätter ----"
    For Each ws In ThisWorkbook.Worksheets
        If ws.ProtectContents Then Debug.Print ws.Name & "  (ProtectContents=True)"
    Next ws
    Debug.Print "----------------------------"
End Sub

' ----------------------
' Zentrale Aktualisierer
' ----------------------
Public Sub RefreshAllLists()
    On Error Resume Next
    Sortiere_Mitgliederliste_Nach_Parzelle
    AktualisiereNamedRange_MitgliederNamen
    On Error GoTo 0
End Sub

Public Sub AktualisiereDatenstand()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(WS_MITGLIEDER)
    On Error GoTo 0
    If ws Is Nothing Then
        Debug.Print "AktualisiereDatenstand: Blatt '" & WS_MITGLIEDER & "' nicht gefunden."
        Exit Sub
    End If

    If Not UnprotectSheet(ws) Then
        Debug.Print "AktualisiereDatenstand: Blatt '" & ws.Name & "' konnte nicht entsperrt werden. Abbruch."
        Exit Sub
    End If

    On Error GoTo Cleanup
    ws.Cells(M_STAND_ROW, M_STAND_COL).Value = Now

Cleanup:
    Call ProtectSheet(ws)
End Sub

' ----------------------
' Member ID Handling
' ----------------------
Public Sub Fuelle_MemberIDs_Wenn_Fehlend()
    Dim wsM As Worksheet
    Dim lastRow As Long, lRow As Long
    Dim wasProtected As Boolean

    On Error GoTo ErrorHandler
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If wsM Is Nothing Then Exit Sub

    wasProtected = wsM.ProtectContents
    If wasProtected Then
        If Not UnprotectSheet(wsM) Then
            Debug.Print "Fuelle_MemberIDs_Wenn_Fehlend: Blatt konnte nicht entsperrt werden."
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False

    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
    If lastRow < M_START_ROW Then GoTo Cleanup

    wsM.Cells(M_HEADER_ROW, M_COL_MEMBER_ID).Value = "Member ID"

    For lRow = M_START_ROW To lastRow
        If Trim(wsM.Cells(lRow, M_COL_MEMBER_ID).Value & "") = "" And _
           Trim(wsM.Cells(lRow, M_COL_NACHNAME).Value & "") <> "" Then
            wsM.Cells(lRow, M_COL_MEMBER_ID).Value = CreateGUID()
        End If
    Next lRow

    With wsM.Range(wsM.Cells(M_START_ROW, M_COL_MEMBER_ID), wsM.Cells(lastRow + 1000, M_COL_MEMBER_ID))
        .Locked = True
        .FormulaHidden = True
    End With

Cleanup:
    Application.ScreenUpdating = True
    If wasProtected Then ProtectSheet wsM
    Exit Sub

ErrorHandler:
    Debug.Print "Fehler beim Füllen der MemberIDs: " & Err.Description
    Resume Cleanup
End Sub

Public Function CreateGUID() As String
    Dim TypeLib As Object
    On Error Resume Next
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    On Error GoTo 0

    If Not TypeLib Is Nothing Then
        CreateGUID = Mid(TypeLib.GUID, 2, 36)
    End If

    If CreateGUID = "" Then
        Randomize
        CreateGUID = Format(Now, "yyyymmddhhmmss") & "-" & CStr(Int((99999 - 10000 + 1) * Rnd + 10000))
    End If

    Set TypeLib = Nothing
End Function

' ----------------------
' Dropdown / Validation
' ----------------------
Public Sub ApplyMitgliederDropdowns()
    Dim ws As Worksheet
    On Error GoTo ErrorHandler
    Set ws = Worksheets(WS_MITGLIEDER)
    If ws Is Nothing Then Exit Sub

    ' Versuche, Blatt still und automatisch zu entsperren
    If Not UnprotectSheet(ws) Then
        Debug.Print "ApplyMitgliederDropdowns: Blatt '" & WS_MITGLIEDER & "' konnte nicht entsperrt werden. Dropdowns werden nicht gesetzt."
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Entsperren der editierbaren Spalten (falls erforderlich)
    On Error Resume Next
    ws.Range(ws.Cells(M_START_ROW, M_COL_PARZELLE), ws.Cells(1000, M_COL_PARZELLE)).Locked = False
    ws.Range(ws.Cells(M_START_ROW, M_COL_ANREDE), ws.Cells(1000, M_COL_ANREDE)).Locked = False
    ws.Range(ws.Cells(M_START_ROW, M_COL_FUNKTION), ws.Cells(1000, M_COL_FUNKTION)).Locked = False
    On Error GoTo ErrorHandler

    ' Dropdowns (Calls)
    Call ApplyDropdown(ws.Range(ws.Cells(M_START_ROW, M_COL_PARZELLE), ws.Cells(1000, M_COL_PARZELLE)), "=Daten!$F$4:$F$18", True)
    Call ApplyDropdown(ws.Range(ws.Cells(M_START_ROW, M_COL_SEITE), ws.Cells(1000, M_COL_SEITE)), "=Daten!$H$4:$H$6", True)
    Call ApplyDropdown(ws.Range(ws.Cells(M_START_ROW, M_COL_ANREDE), ws.Cells(1000, M_COL_ANREDE)), "=Daten!$D$4:$D$9", True)
    Call ApplyDropdown(ws.Range(ws.Cells(M_START_ROW, M_COL_FUNKTION), ws.Cells(1000, M_COL_FUNKTION)), "=Daten!$B$4:$B$11", True)

    Application.ScreenUpdating = True
    Call ProtectSheet(ws)
    Exit Sub

ErrorHandler:
    On Error Resume Next
    ProtectSheet ws
    Debug.Print "Fehler beim Setzen der Dropdown-Listen: " & Err.Description
    On Error GoTo 0
End Sub

Private Sub ApplyDropdown(ByVal targetRange As Range, ByVal sourceFormula As String, ByVal allowBlanks As Boolean)
    If targetRange Is Nothing Then Exit Sub

    On Error Resume Next
    targetRange.Validation.Delete
    On Error GoTo 0

    On Error Resume Next
    targetRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=sourceFormula
    targetRange.Validation.IgnoreBlank = allowBlanks
    targetRange.Validation.InCellDropdown = True
    targetRange.Validation.ErrorTitle = "Ungültiger Wert"
    targetRange.Validation.ErrorMessage = "Bitte wählen Sie einen Wert aus der Liste."
    If Err.Number <> 0 Then
        Debug.Print "ApplyDropdown: Validation konnte nicht gesetzt werden auf Range " & targetRange.Address & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Sub Reapply_Data_Validation()
    ApplyMitgliederDropdowns
End Sub

' ----------------------
' Zebra-Formatierung (bedingte Formatierung)
' ----------------------
Public Sub Anwende_Zebra_Formatierung(ByVal ws As Worksheet, ByVal startCol As Long, ByVal endCol As Long, ByVal startRow As Long, ByVal dataCheckCol As Long)
    Const ZEBRA_COLOR As Long = &HDEE5E3

    If ws Is Nothing Then Exit Sub
    Dim rngFullData As Range
    Dim sFormula As String
    Dim checkColLetter As String

    Set rngFullData = ws.Range(ws.Cells(startRow, startCol), ws.Cells(1000, endCol))

    On Error Resume Next
    rngFullData.FormatConditions.Delete
    On Error GoTo 0

    rngFullData.Interior.ColorIndex = xlNone

    checkColLetter = Split(ws.Columns(dataCheckCol).Address(False, False), ":")(0)
    sFormula = "=UND(NICHT(ISTLEER($" & checkColLetter & startRow & ")); REST(ZEILE();2)=0)"

    With rngFullData.FormatConditions.Add(Type:=xlExpression, Formula1:=sFormula)
        .Interior.color = ZEBRA_COLOR
        .StopIfTrue = True
    End With
End Sub

' ----------------------
' Gesamtformatierung der Tabellen
' ----------------------
Public Sub Formatiere_Alle_Tabellen_Neu()
    Dim wsM As Worksheet, wsD As Worksheet
    Dim wasProtectedM As Boolean, wasProtectedD As Boolean

    Const DATA_START_ROW As Long = 4
    Const M_START_COL As Long = 1
    Const M_CHECK_COL As Long = 5 ' Nachname
    Const D_ENTITYKEY_END_COL As Long = 21

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If Not wsM Is Nothing Then
        wasProtectedM = wsM.ProtectContents
        If wasProtectedM Then
            If Not UnprotectSheet(wsM) Then
                Debug.Print "Formatiere_Alle_Tabellen_Neu: Blatt '" & wsM.Name & "' konnte nicht entsperrt werden. Überspringe Formatierung."
                GoTo SkipM
            End If
        End If

        Anwende_Zebra_Formatierung wsM, M_START_COL, M_COL_PACHTENDE, M_START_ROW, M_CHECK_COL

        Call ProtectSheet(wsM)
SkipM:
    End If

    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    If Not wsD Is Nothing Then
        wasProtectedD = wsD.ProtectContents
        If wasProtectedD Then
            If Not UnprotectSheet(wsD) Then
                Debug.Print "Formatiere_Alle_Tabellen_Neu: Blatt '" & wsD.Name & "' konnte nicht entsperrt werden. Überspringe Formatierung."
                GoTo SkipD
            End If
        End If

        Anwende_Zebra_Formatierung wsD, DATA_CAT_COL_START, DATA_CAT_COL_END, DATA_START_ROW, DATA_CAT_COL_START
        Anwende_Zebra_Formatierung wsD, DATA_MAP_COL_ENTITYKEY, D_ENTITYKEY_END_COL, DATA_START_ROW, DATA_MAP_COL_ENTITYKEY

        Call ProtectSheet(wsD)
SkipD:
    End If

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Debug.Print "FEHLER beim Formatieren der Tabellen: " & Err.Description
    Resume Cleanup
End Sub

' ----------------------
' Sortierung Mitgliederliste nach Parzelle
' ----------------------
Public Sub Sortiere_Mitgliederliste_Nach_Parzelle()
    Dim ws As Worksheet
    Dim rngSort As Range
    Dim lastRow As Long
    Dim wasProtected As Boolean

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If ws Is Nothing Then Exit Sub

    wasProtected = ws.ProtectContents
    If wasProtected Then
        If Not UnprotectSheet(ws) Then
            Debug.Print "Sortiere_Mitgliederliste_Nach_Parzelle: Blatt '" & ws.Name & "' konnte nicht entsperrt werden. Abbruch."
            Exit Sub
        End If
    End If

    lastRow = ws.Cells(ws.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
    If lastRow < M_START_ROW Then GoTo Cleanup

    Set rngSort = ws.Range(ws.Cells(M_START_ROW, 1), ws.Cells(lastRow, M_COL_PACHTENDE))

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(M_START_ROW, M_COL_PACHTENDE), ws.Cells(lastRow, M_COL_PACHTENDE)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=ws.Range(ws.Cells(M_START_ROW, M_COL_PARZELLE), ws.Cells(lastRow, M_COL_PARZELLE)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .SortFields.Add Key:=ws.Range(ws.Cells(M_START_ROW, M_COL_ANREDE), ws.Cells(lastRow, M_COL_ANREDE)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        .SetRange rngSort
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Reapply_Data_Validation
    Formatiere_Alle_Tabellen_Neu

Cleanup:
    If Not ws Is Nothing Then
        If wasProtected Then ProtectSheet ws
    End If
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    If Not ws Is Nothing Then
        If wasProtected Then ProtectSheet ws
    End If
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Debug.Print "FEHLER BEIM SORTIEREN (mod_Mitglieder_UI):" & Err.Number & " - " & Err.Description
    Resume Cleanup
End Sub

' ----------------------
' Named Range Aktualisierung
' ----------------------
Public Sub AktualisiereNamedRange_MitgliederNamen()
    Dim wsM As Worksheet
    Dim lastRow As Long
    Dim rngTarget As Range

    On Error GoTo ErrorHandler
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If wsM Is Nothing Then Exit Sub

    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
    If lastRow < M_START_ROW Then
        On Error Resume Next
        ThisWorkbook.Names("MitgliederNamen").Delete
        On Error GoTo 0
        Exit Sub
    End If

    Set rngTarget = wsM.Range(wsM.Cells(M_START_ROW, M_COL_NACHNAME), wsM.Cells(lastRow, M_COL_NACHNAME))

    On Error Resume Next
    ThisWorkbook.Names("MitgliederNamen").Delete
    On Error GoTo ErrorHandler

    ThisWorkbook.Names.Add Name:="MitgliederNamen", RefersTo:=rngTarget
    Exit Sub

ErrorHandler:
    Debug.Print "Fehler beim Aktualisieren des Named Range 'MitgliederNamen': " & Err.Description
End Sub

' ----------------------
' Suche & Hilfsfunktionen
' ----------------------
Public Function GetEntityKeyByParzelle(ByVal ParzelleNr As String) As String
    Dim wsD As Worksheet
    Dim lastRow As Long
    Dim rngFind As Range

    If Trim(ParzelleNr) = "" Then
        GetEntityKeyByParzelle = ""
        Exit Function
    End If

    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    If wsD Is Nothing Then Exit Function

    lastRow = wsD.Cells(wsD.Rows.Count, DATA_MAP_COL_PARZELLE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Function

    Set rngFind = wsD.Range(wsD.Cells(DATA_START_ROW, DATA_MAP_COL_PARZELLE), wsD.Cells(lastRow, DATA_MAP_COL_PARZELLE)) _
                    .Find(What:=ParzelleNr, LookIn:=xlValues, LookAt:=xlWhole)

    If Not rngFind Is Nothing Then
        GetEntityKeyByParzelle = wsD.Cells(rngFind.Row, DATA_MAP_COL_ENTITYKEY).Value & ""
    Else
        GetEntityKeyByParzelle = ""
    End If
End Function

Private Function FindeRowByMemberID(ByVal MemberID As String) As Long
    Dim wsM As Worksheet
    Dim rngSearch As Range
    Dim rngFind As Range
    Dim lastRow As Long
    Dim bWasProtected As Boolean

    FindeRowByMemberID = 0
    If Trim(MemberID & "") = "" Then Exit Function

    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If wsM Is Nothing Then Exit Function

    bWasProtected = wsM.ProtectContents
    If bWasProtected Then UnprotectSheet wsM

    If wsM.AutoFilterMode Then
        If wsM.FilterMode Then
            On Error Resume Next
            wsM.ShowAllData
            On Error GoTo 0
        End If
    End If

    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_MEMBER_ID).End(xlUp).Row
    If lastRow < M_START_ROW Then GoTo CleanExit

    Set rngSearch = wsM.Range(wsM.Cells(M_START_ROW, M_COL_MEMBER_ID), wsM.Cells(lastRow, M_COL_MEMBER_ID))

    Set rngFind = rngSearch.Find(What:=MemberID, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If Not rngFind Is Nothing Then FindeRowByMemberID = rngFind.Row

CleanExit:
    If bWasProtected Then ProtectSheet wsM
End Function

' ----------------------
' Historie & Mitglieder-Update (wird von UserForm ausgelöst)
' ----------------------
Public Sub Speichere_Historie_und_Aktualisiere_Mitgliederliste( _
    ByVal selectedRow As Long, _
    ByVal OldParzelle As String, _
    ByVal OldMemberID As String, _
    ByVal Nachname As String, _
    ByVal AustrittsDatum As Date, _
    ByVal NewParzelleNr As String, _
    ByVal NewMemberID As String, _
    ByVal ChangeReason As String)

    Dim wsM As Worksheet, wsH As Worksheet
    Dim NextRow As Long, UebernehmerRow As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)

    ' 1) Historie speichern
    If Not UnprotectSheet(wsH) Then
        Debug.Print "Speichere_Historie: Historie-Blatt konnte nicht entsperrt werden. Abbruch."
        GoTo CleanExit
    End If

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
    ProtectSheet wsH

    ' 2) Mitgliederliste aktualisieren
    If Not UnprotectSheet(wsM) Then
        Debug.Print "Speichere_Historie: Mitglieder-Blatt konnte nicht entsperrt werden. Abbruch."
        GoTo CleanExit
    End If

    If ChangeReason = "Parzellenwechsel" And Trim(NewParzelleNr & "") <> "" Then
        wsM.Cells(selectedRow, M_COL_PARZELLE).Value = NewParzelleNr
    ElseIf ChangeReason = "Austritt aus Parzelle" Or ChangeReason = "Austritt mit Pachtübernahme" Then
        wsM.Cells(selectedRow, M_COL_PARZELLE).Value = ""
        wsM.Cells(selectedRow, M_COL_PACHTENDE).Value = AustrittsDatum
        wsM.Cells(selectedRow, M_COL_PACHTENDE).NumberFormat = "dd.mm.yyyy"
        wsM.Cells(selectedRow, M_COL_FUNKTION).Value = AUSTRITT_STATUS_DISPLAY
    End If

    If ChangeReason = "Austritt mit Pachtübernahme" And Trim(NewMemberID & "") <> "" Then
        UebernehmerRow = FindeRowByMemberID(NewMemberID)
        If UebernehmerRow > 0 Then
            wsM.Cells(UebernehmerRow, M_COL_FUNKTION).Value = PAECHTER_STATUS
            MsgBox "Pachtvertrag für Parzelle " & OldParzelle & " erfolgreich auf " & wsM.Cells(UebernehmerRow, M_COL_NACHNAME).Value & " übertragen.", vbInformation
        Else
            MsgBox "FEHLER: MemberID des Übernehmers '" & NewMemberID & "' konnte nicht gefunden werden.", vbCritical
        End If
    End If

    AktualisiereDatenstand
    ProtectSheet wsM

    ' 3) Nachfolgende Aktualisierungen
    AktualisiereNamedRange_MitgliederNamen
    Sortiere_Mitgliederliste_Nach_Parzelle

    On Error Resume Next
    Call mod_Banking_Data.Aktualisiere_Parzellen_Mapping_Final
    Call mod_Banking_Data.Sortiere_Tabellen_Daten
    Call mod_ZaehlerLogik.Ermittle_Kennzahlen_Mitgliederliste
    Call mod_ZaehlerLogik.ErzeugeParzellenUebersicht
    Call mod_ZaehlerLogik.AktualisiereZaehlerTabellenSpalteA
    On Error GoTo 0

    If IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If

    If ChangeReason <> "Austritt mit Pachtübernahme" Then
        MsgBox "Historien-Eintrag erfolgreich gespeichert und Mitgliederliste aktualisiert.", vbInformation
    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    If Not wsM Is Nothing Then ProtectSheet wsM
    If Not wsH Is Nothing Then ProtectSheet wsH
    MsgBox "FEHLER BEI DER DATENVERARBEITUNG NACH FORMULARABSCHLUSS: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' ----------------------
' UI Helfer & Prüfungen
' ----------------------
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

Public Function CheckIfLastPaechter(ByVal PaeffelParzelle As String, ByVal MemberIDToExclude As String) As Boolean
    Dim wsM As Worksheet
    Dim lastRowM As Long, lRow As Long
    Dim PachterCount As Long
    Dim currentParzelle As String, currentMemberID As String, currentFunktion As String

    CheckIfLastPaechter = False
    On Error GoTo ErrorHandler

    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If wsM Is Nothing Then Exit Function

    lastRowM = wsM.Cells(wsM.Rows.Count, M_COL_PARZELLE).End(xlUp).Row
    If lastRowM < M_START_ROW Then
        CheckIfLastPaechter = True
        Exit Function
    End If

    For lRow = M_START_ROW To lastRowM
        currentParzelle = Trim(CStr(wsM.Cells(lRow, M_COL_PARZELLE).Value))
        currentMemberID = Trim(CStr(wsM.Cells(lRow, M_COL_MEMBER_ID).Value))
        currentFunktion = Trim(CStr(wsM.Cells(lRow, M_COL_FUNKTION).Value))
        If UCase(currentParzelle) = UCase(PaeffelParzelle) Then
            If UCase(currentFunktion) = UCase(PAECHTER_STATUS) Then
                If UCase(currentMemberID) <> UCase(MemberIDToExclude) Then
                    CheckIfLastPaechter = False
                    Exit Function
                End If
            End If
        End If
    Next lRow

    CheckIfLastPaechter = True
    Exit Function

ErrorHandler:
    MsgBox "Fehler in CheckIfLastPaechter: " & Err.Description, vbCritical
    CheckIfLastPaechter = True
End Function

Public Function GetSekundaerMitgliederAufParzelle(ByVal ParzelleNr As String) As Variant
    Dim wsM As Worksheet
    Dim lastRowM As Long, lRow As Long
    Dim SekundaerList() As String
    Dim i As Long: i = -1

    On Error GoTo ErrorHandler
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If wsM Is Nothing Then GetSekundaerMitgliederAufParzelle = Array(): Exit Function

    lastRowM = wsM.Cells(wsM.Rows.Count, M_COL_PARZELLE).End(xlUp).Row
    If lastRowM < M_START_ROW Then GetSekundaerMitgliederAufParzelle = Array(): Exit Function

    For lRow = M_START_ROW To lastRowM
        If UCase(Trim(CStr(wsM.Cells(lRow, M_COL_PARZELLE).Value))) = UCase(ParzelleNr) And _
           UCase(Trim(CStr(wsM.Cells(lRow, M_COL_FUNKTION).Value))) = UCase(SEKUNDAER_STATUS) And _
           (IsDate(wsM.Cells(lRow, M_COL_PACHTENDE).Value) = False) Then

            i = i + 1
            ReDim Preserve SekundaerList(0 To i)
            SekundaerList(i) = Trim(CStr(wsM.Cells(lRow, M_COL_NACHNAME).Value)) & ", " & _
                               Trim(CStr(wsM.Cells(lRow, M_COL_VORNAME).Value)) & "|" & _
                               Trim(CStr(wsM.Cells(lRow, M_COL_MEMBER_ID).Value))
        End If
    Next lRow

    If i >= 0 Then
        GetSekundaerMitgliederAufParzelle = SekundaerList
    Else
        GetSekundaerMitgliederAufParzelle = Array()
    End If
    Exit Function

ErrorHandler:
    MsgBox "Fehler in GetSekundaerMitgliederAufParzelle: " & Err.Description, vbCritical
    GetSekundaerMitgliederAufParzelle = Array()
End Function

Public Function Check_Vorstand_Eindeutigkeit(ByVal CheckMemberID As String) As Boolean
    Dim wsM As Worksheet
    Dim lastRowM As Long, lRow As Long
    Dim currentMemberID As String, currentFunktion As String

    Check_Vorstand_Eindeutigkeit = True
    On Error GoTo ErrorHandler

    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If wsM Is Nothing Then Exit Function

    lastRowM = wsM.Cells(wsM.Rows.Count, M_COL_FUNKTION).End(xlUp).Row
    If lastRowM < M_START_ROW Then Exit Function

    For lRow = M_START_ROW To lastRowM
        currentFunktion = Trim(CStr(wsM.Cells(lRow, M_COL_FUNKTION).Value))
        currentMemberID = Trim(CStr(wsM.Cells(lRow, M_COL_MEMBER_ID).Value))
        If UCase(currentFunktion) = UCase(VORSTAND_STATUS) Then
            If UCase(currentMemberID) <> UCase(CheckMemberID) Then
                Check_Vorstand_Eindeutigkeit = False
                Exit Function
            End If
        End If
    Next lRow
    Exit Function

ErrorHandler:
    MsgBox "Fehler in Check_Vorstand_Eindeutigkeit: " & Err.Description, vbCritical
    Check_Vorstand_Eindeutigkeit = False
End Function

' -------------------------------------------------------------------------
' Ende Modul
' -------------------------------------------------------------------------


