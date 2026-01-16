Attribute VB_Name = "mod_Hilfsfunktionen"
Option Explicit

' **********************************************************
' MODUL: mod_Hilfsfunktionen (Nur Hilfsroutinen - Konstanten kommen aus mod_Const)
' **********************************************************

' Der temporäre Blattname ist PRIVATE und kann hier bleiben, falls er nur lokal verwendet wird
Private Const TEMP_WS_NAME As String = "TEMP_LISTEN"

' **********************************************************
' TEIL 1: GENERISCHE HILFSPROZEDUREN (Listen & Benannte Bereiche)
' **********************************************************

' PROZEDUR: AktualisiereNamedRange_MitgliederNamen
Public Sub AktualisiereNamedRange_MitgliederNamen()
    
    ' HINWEIS: Dieser Code verwendet nun Konstanten (z.B. WS_MITGLIEDER, PASSWORD, M_COL_...)
    ' die im Modul mod_Const definiert sein müssen, um zu kompilieren.
    
    Dim wsM As Worksheet
    Dim wsTemp As Worksheet
    Dim lastRow As Long
    Dim tempRow As Long
    Dim rngTarget As Range
    Dim wasProtected As Boolean
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    ' 1. Temporäres Arbeitsblatt erstellen/finden
    On Error Resume Next
    Set wsTemp = ThisWorkbook.Worksheets(TEMP_WS_NAME)
    On Error GoTo 0
    
    If wsTemp Is Nothing Then
        Set wsTemp = ThisWorkbook.Worksheets.Add(After:=wsM)
        wsTemp.Name = TEMP_WS_NAME
    Else
        wsTemp.Cells.Clear
    End If
    
    ' 2. Daten kopieren und filtern (Nur aktive Mitglieder)
    wasProtected = wsM.ProtectContents
    If wasProtected Then Call UnprotectSheet(wsM) ' Nutzt die Hilfsprozedur
    
    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
    
    If lastRow >= M_START_ROW Then
        
        wsM.Range(wsM.Cells(M_HEADER_ROW, 1), wsM.Cells(lastRow, M_COL_FUNKTION)).AutoFilter
        
        wsM.Range(wsM.Cells(M_HEADER_ROW, 1), wsM.Cells(lastRow, M_COL_PACHTENDE)).AutoFilter _
             Field:=M_COL_PACHTENDE, Criteria1:="="
        
        tempRow = 1
        Dim copyCols As Variant
        copyCols = Array(M_COL_NACHNAME, M_COL_VORNAME, M_COL_PARZELLE)
        Dim i As Long
        
        For i = LBound(copyCols) To UBound(copyCols)
            wsM.Columns(copyCols(i)).SpecialCells(xlCellTypeVisible).Copy
            wsTemp.Cells(tempRow, i + 1).PasteSpecial xlPasteValues
        Next i
        
        Application.CutCopyMode = False
        wsM.AutoFilterMode = False
        
        ' 3. Kombinierte Namen-Liste erstellen
        tempRow = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
        
        If tempRow > 1 Then
            For i = 2 To tempRow
                wsTemp.Cells(i, 4).Value = wsTemp.Cells(i, 1).Value & ", " & wsTemp.Cells(i, 2).Value
            Next i
            
            ' 4. Benannten Bereich erstellen/aktualisieren
            Set rngTarget = wsTemp.Range(wsTemp.Cells(2, 4), wsTemp.Cells(tempRow, 4))
            
            On Error Resume Next
            ThisWorkbook.Names("rng_MitgliederNamen").Delete
            On Error GoTo 0
            
            ThisWorkbook.Names.Add Name:="rng_MitgliederNamen", RefersTo:=rngTarget
        End If
    End If
    
Cleanup:
    Application.ScreenUpdating = True
    If Not wsM Is Nothing Then
        If wsM.AutoFilterMode Then wsM.AutoFilterMode = False
        If wasProtected Then Call ProtectSheet(wsM) ' Nutzt die Hilfsprozedur
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler in AktualisiereNamedRange_MitgliederNamen: " & Err.Description, vbCritical
    Resume Cleanup

End Sub

' HILFSFUNKTION: Prüfen, ob eine UserForm geladen ist
Private Function IsFormLoaded(ByVal FormName As String) As Boolean
    
    Dim i As Long
    
    On Error Resume Next
    For i = 0 To VBA.UserForms.Count - 1
        If StrComp(VBA.UserForms.Item(i).Name, FormName, vbTextCompare) = 0 Then
            IsFormLoaded = True
            Exit Function
        End If
    Next i
    On Error GoTo 0
    
    IsFormLoaded = False
    
End Function


' **********************************************************
' TEIL 2: UI- und Listen-Hilfen (Basisfunktionen)
' **********************************************************

Public Sub RefreshAllLists()
    ' Aktualisiert die ListBox im Hauptformular und die Named Ranges
    If IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    Call AktualisiereNamedRange_MitgliederNamen
End Sub

Public Sub UnprotectSheet(ByRef ws As Worksheet)
    If ws.ProtectContents Then
        ws.Unprotect PASSWORD:=PASSWORD
    End If
End Sub

Public Sub ProtectSheet(ByRef ws As Worksheet)
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
End Sub


' **********************************************************
' TEIL 3: KERN-GESCHÄFTSLOGIK (Muss hier bleiben, falls kein mod_Mitglieder_Logik erstellt wird)
' **********************************************************

' FUNKTION: Generiert eine eindeutige ID (GUID)
Public Function CreateGUID() As String
    On Error Resume Next
    CreateGUID = CreateObject("Scriptlet.TypeLib").GUID
    If Err.Number <> 0 Then
        CreateGUID = Format(Now, "yyyymmddhhmmss") & "-" & Int(Rnd() * 100000)
    Else
        CreateGUID = Mid(CreateGUID, 2, Len(CreateGUID) - 2)
    End If
    On Error GoTo 0
End Function

' PROZEDUR: Schreibt einen Eintrag in das Historien-Protokoll
Public Sub SchreibeHistorie(ByVal MemberID As String, ByVal Parzelle As String, ByVal Nachname As String, _
                            ByVal Datum As Variant, ByVal AlterWert As String, ByVal NeuerWert As String, _
                            ByVal Aktion As String)
    
    Dim wsH As Worksheet
    Dim lRow As Long
    
    On Error GoTo ErrorHandler
    
    ' KORREKTUR: Die Konstante muss WS_MITGLIEDER_HISTORIE lauten
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    Call UnprotectSheet(wsH)
    
    lRow = wsH.Cells(wsH.Rows.Count, 1).End(xlUp).Row + 1
    
    wsH.Cells(lRow, 1).Value = Now            ' Datum/Zeit der Eintragung
    wsH.Cells(lRow, 2).Value = MemberID        ' Mitglieds-ID
    wsH.Cells(lRow, 3).Value = Parzelle        ' Parzelle
    wsH.Cells(lRow, 4).Value = Nachname        ' Nachname
    wsH.Cells(lRow, 5).Value = Aktion          ' Aktion (z.B. "Neuanlage", "Austritt", "Adresse geändert")
    wsH.Cells(lRow, 6).Value = Datum           ' Datum der Änderung (z.B. Pachtbeginn/ende)
    wsH.Cells(lRow, 7).Value = AlterWert       ' Alter Wert
    wsH.Cells(lRow, 8).Value = NeuerWert       ' Neuer Wert
    
    Call ProtectSheet(wsH)
    Exit Sub

ErrorHandler:
    If Not wsH Is Nothing Then Call ProtectSheet(wsH)
    MsgBox "Fehler beim Schreiben der Historie: " & Err.Description, vbCritical
End Sub

' PROZEDUR: HandleAustritt
Public Sub HandleAustritt(ByVal MemberID As String, ByVal OldParzelle As String, ByVal Nachname As String)
    
    MsgBox "Austrittslogik für " & Nachname & " (Parzelle: " & OldParzelle & ") wird gestartet." & vbCrLf & _
           "Hier müsste die dedizierte Austritts-UserForm erscheinen, um Pachtende zu setzen.", vbInformation
            
End Sub






' **********************************************************
' TEIL 4: MEMBER LOOKUP HELPERS
' **********************************************************

' FUNCTION: FindMemberRowByID
' Finds a member row by MemberID in column A using Find method
' Returns the row number, or 0 if not found
' This function is not affected by hidden columns
Public Function FindMemberRowByID(ws As Worksheet, memberID As Variant) As Long
    
    Dim foundCell As Range
    
    On Error Resume Next
    
    ' Use Find method on column A to locate the MemberID
    ' Find is not affected by hidden columns or rows
    Set foundCell = ws.Columns("A").Find( _
        What:=memberID, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False _
    )
    
    If Not foundCell Is Nothing Then
        FindMemberRowByID = foundCell.Row
    Else
        FindMemberRowByID = 0
    End If
    
    On Error GoTo 0
    
End Function
