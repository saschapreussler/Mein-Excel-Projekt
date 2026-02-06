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
            .value = Now
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
    
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    If lastRow < M_START_ROW Then GoTo CleanUp
    
    Application.ScreenUpdating = False
    
    ' Header setzen
    wsM.Cells(M_HEADER_ROW, M_COL_MEMBER_ID).value = "Member ID"
    
    ' Schleife durch alle Zeilen ab M_START_ROW
    For lRow = M_START_ROW To lastRow
        ' Pruefen, ob eine MemberID fehlt und ob der Datensatz nicht leer ist
        If wsM.Cells(lRow, M_COL_MEMBER_ID).value = "" And _
           wsM.Cells(lRow, M_COL_NACHNAME).value <> "" Then
            
            ' GUID generieren und eintragen
            wsM.Cells(lRow, M_COL_MEMBER_ID).value = CreateGUID_Public()
        End If
    Next lRow
    
    ' Spalte A sperren
    With wsM.Range(wsM.Cells(M_START_ROW, M_COL_MEMBER_ID), wsM.Cells(lastRow + 1000, M_COL_MEMBER_ID))
        .Locked = True
        .FormulaHidden = True
    End With
    
CleanUp:
    Application.ScreenUpdating = True
    If wasProtected Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler beim Fuellen der MemberIDs: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

' ***************************************************************
' HILFSFUNKTION: GUID erstellen (PUBLIC - fuer frm_Mitgliedsdaten zugaenglich)
' ***************************************************************
Public Function CreateGUID_Public() As String
    
    On Error Resume Next
    Dim TypeLib As Object
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    CreateGUID_Public = Mid(TypeLib.guid, 2, 36)
    
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
    Call ApplyDropdown(ws.Range(ws.Cells(M_START_ROW, M_COL_FUNKTION), ws.Cells(1000, M_COL_FUNKTION)), "=Daten!$B$4:$B$12", True)

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
        .ErrorTitle = "Ungueltiger Wert"
        .ErrorMessage = "Bitte waehlen Sie einen Wert aus der Liste."
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
    
    lastRow = ws.Cells(ws.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    If lastRow < M_START_ROW Then GoTo CleanUp
    
    Set rngSort = ws.Range(ws.Cells(M_START_ROW, 1), ws.Cells(lastRow, M_COL_PACHTENDE))
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Columns(M_COL_PACHTENDE), SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add key:=ws.Columns(M_COL_PARZELLE), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .SortFields.Add key:=ws.Columns(M_COL_ANREDE), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange rngSort
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Call Reapply_Data_Validation
    Call mod_Formatierung.Formatiere_Alle_Tabellen_Neu
    
CleanUp:
    If Not ws Is Nothing Then
        If wasProtected Then ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    Exit Sub

ErrorHandler:
    If Not ws Is Nothing Then
        If wasProtected Then ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    MsgBox "FEHLER BEIM SORTIEREN: " & Err.Description, vbCritical
    Resume CleanUp
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
    
    lastRow = wsD.Cells(wsD.Rows.count, DATA_MAP_COL_PARZELLE).End(xlUp).Row
    Set rngFind = wsD.Range(wsD.Cells(DATA_START_ROW, DATA_MAP_COL_PARZELLE), wsD.Cells(lastRow, DATA_MAP_COL_PARZELLE))
    
    Set rngFind = rngFind.Find(What:=ParzelleNr, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rngFind Is Nothing Then
        GetEntityKeyByParzelle = wsD.Cells(rngFind.Row, DATA_MAP_COL_ENTITYKEY).value
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
' SICHERHEITSKRITISCH: Schuetzt die Verein-Parzelle vor Datenueberschreibung
' ***************************************************************
Public Sub Speichere_Historie_und_Aktualisiere_Mitgliederliste( _
    ByVal selectedRow As Long, _
    ByVal OldParzelle As String, _
    ByVal OldMemberID As String, _
    ByVal nachname As String, _
    ByVal austrittsDatum As Date, _
    ByVal NewParzelleNr As String, _
    ByVal newMemberID As String, _
    ByVal ChangeReason As String)

    Dim wsM As Worksheet
    Dim wsH As Worksheet
    Dim nextRow As Long
    Dim lastRow As Long
    Dim iRow As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    
    ' === SICHERHEITSCHECK 1: selectedRow darf nicht die Verein-Parzelle sein ===
    If selectedRow >= M_START_ROW Then
        If Trim(wsM.Cells(selectedRow, M_COL_PARZELLE).value) = PARZELLE_VEREIN Then
            MsgBox "FEHLER: Die Verein-Parzelle darf nicht geaendert werden!", vbCritical
            GoTo CleanUp
        End If
    End If
    
    ' --- 1. HISTORIE SPEICHERN ---
    wsH.Unprotect PASSWORD:=PASSWORD
    nextRow = wsH.Cells(wsH.Rows.count, H_COL_PARZELLE).End(xlUp).Row + 1
    If nextRow < H_START_ROW Then nextRow = H_START_ROW
    
    wsH.Cells(nextRow, H_COL_PARZELLE).value = OldParzelle
    wsH.Cells(nextRow, H_COL_MITGL_ID).value = OldMemberID
    wsH.Cells(nextRow, H_COL_NACHNAME).value = nachname
    wsH.Cells(nextRow, H_COL_AUST_DATUM).value = austrittsDatum
    wsH.Cells(nextRow, H_COL_NEUER_PAECHTER_ID).value = newMemberID
    wsH.Cells(nextRow, H_COL_GRUND).value = ChangeReason
    wsH.Cells(nextRow, H_COL_SYSTEMZEIT).value = Now
    
    wsH.Cells(nextRow, H_COL_AUST_DATUM).NumberFormat = "dd.mm.yyyy"
    wsH.Cells(nextRow, H_COL_SYSTEMZEIT).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    wsH.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True

    ' --- 2. MITGLIEDERLISTE AKTUALISIEREN ---
    wsM.Unprotect PASSWORD:=PASSWORD
    
    If ChangeReason = "Parzellenwechsel" And NewParzelleNr <> "" Then
        ' === SICHERHEITSCHECK 2: NewParzelleNr darf nicht "Verein" sein ===
        If Trim(NewParzelleNr) = PARZELLE_VEREIN Then
            MsgBox "FEHLER: Austretende Mitglieder duerfen nicht zur Verein-Parzelle wechseln!", vbCritical
            wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
            GoTo CleanUp
        End If
        
        wsM.Cells(selectedRow, M_COL_PARZELLE).value = NewParzelleNr
        
    ElseIf ChangeReason = "Austritt aus Parzelle" Then
        ' === SICHERHEITSCHECK 3: Stelle sicher, dass wir nicht die Verein-Parzelle antasten ===
        If Trim(wsM.Cells(selectedRow, M_COL_PARZELLE).value) = PARZELLE_VEREIN Then
            MsgBox "FEHLER: Die Verein-Parzelle kann nicht aufgeloest werden!", vbCritical
            wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
            GoTo CleanUp
        End If
        
        ' Setze Parzelle auf leer (Member ist ausgetreten)
        wsM.Cells(selectedRow, M_COL_PARZELLE).value = ""
        wsM.Cells(selectedRow, M_COL_PACHTENDE).value = austrittsDatum
        wsM.Cells(selectedRow, M_COL_PACHTENDE).NumberFormat = "dd.mm.yyyy"
        wsM.Cells(selectedRow, M_COL_FUNKTION).value = AUSTRITT_STATUS
    End If
    
    Call AktualisiereDatenstand
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' === SICHERHEITSCHECK 4: Verifikation vor Sortierung ===
    Call VerifikationVereinsParzelleIntakt
    
    ' --- 3. AUFRAEUMEN & AKTUALISIERUNG ---
    On Error Resume Next
    Call mod_Hilfsfunktionen.AktualisiereNamedRange_MitgliederNamen
    Call Sortiere_Mitgliederliste_Nach_Parzelle
    Call mod_EntityKey_Manager.ImportiereIBANsAusBankkonto
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
    
CleanUp:
    If Not wsM Is Nothing Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    If Not wsH Is Nothing Then wsH.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    If Not wsM Is Nothing Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    If Not wsH Is Nothing Then wsH.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "FEHLER BEI DER DATENVERARBEITUNG: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' ***************************************************************
' HILFSFUNKTION: Verifikation dass Verein-Parzelle intakt ist
' ***************************************************************
Private Sub VerifikationVereinsParzelleIntakt()
    Dim ws As Worksheet
    Dim lRow As Long
    Dim vereinParzelleRow As Long
    Dim vereinRowNachname As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    If ws Is Nothing Then Exit Sub
    
    vereinParzelleRow = 0
    
    ' Finde die Zeile mit "Verein"-Parzelle
    For lRow = M_START_ROW To ws.Cells(ws.Rows.count, M_COL_PARZELLE).End(xlUp).Row
        If Trim(ws.Cells(lRow, M_COL_PARZELLE).value) = PARZELLE_VEREIN Then
            vereinParzelleRow = lRow
            vereinRowNachname = Trim(ws.Cells(lRow, M_COL_NACHNAME).value)
            
            ' SICHERHEITSPRUEFUNG: Die Verein-Zeile sollte leer sein oder nur spezielle Marker enthalten
            ' Falls Nachname leer: OK
            ' Falls Nachname nicht leer: Warnung (zeigt manuell uebernommene Daten an)
            If vereinRowNachname <> "" Then
                ' Die Zeile hat Mitgliederdaten - dies sollte nicht passieren!
                Debug.Print "WARNUNG: Verein-Parzelle-Zeile (" & vereinParzelleRow & ") enthaelt Mitgliederdaten: " & vereinRowNachname
            End If
            Exit For
        End If
    Next lRow
End Sub

' ***************************************************************
' HILFSFUNKTION: Pruefen, ob eine UserForm geladen ist
' ***************************************************************
Private Function IsFormLoaded(ByVal FormName As String) As Boolean
    
    Dim f As Object
    
    For Each f In VBA.UserForms
        If StrComp(f.name, FormName, vbTextCompare) = 0 Then
            IsFormLoaded = True
            Exit Function
        End If
    Next f
    
    IsFormLoaded = False
    
End Function

