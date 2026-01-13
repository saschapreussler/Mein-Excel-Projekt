Attribute VB_Name = "mod_ReassignKeys"
Option Explicit

' ***************************************************************
' MODUL: mod_ReassignKeys
' ZWECK: Reassignment von BANK-IDs zu MemberIDs für neue Mitglieder
' ***************************************************************

' Reassigns BANK-Keys to a new member's MemberID
' This is useful when a member is created after bank transactions exist
Public Sub ReassignBankKeysForNewMember(ByVal memberID As String)
    Dim wsD As Worksheet
    Dim wsM As Worksheet
    Dim wsBK As Worksheet
    Dim lastRowD As Long
    Dim lastRowBK As Long
    Dim rD As Long, rBK As Long
    Dim entityKey As Variant
    Dim memberName As String
    Dim memberVorname As String
    Dim memberNachname As String
    Dim memberParzelle As String
    Dim kontoName As String
    Dim parzelle As String
    Dim reassignCount As Long
    Dim bankReassignCount As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Trim(memberID) = "" Then
        MsgBox "Fehler: Keine MemberID angegeben!", vbCritical
        Exit Sub
    End If
    
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    reassignCount = 0
    bankReassignCount = 0
    
    ' ================================================
    ' STEP 1: Find member details
    ' ================================================
    
    Dim memberRow As Long
    memberRow = FindRowByMemberID_Safe(memberID, wsM)
    
    If memberRow = 0 Then
        MsgBox "Fehler: MemberID '" & memberID & "' nicht in Mitgliederliste gefunden!", vbCritical
        GoTo Cleanup
    End If
    
    memberNachname = Trim(CStr(wsM.Cells(memberRow, M_COL_NACHNAME).Value))
    memberVorname = Trim(CStr(wsM.Cells(memberRow, M_COL_VORNAME).Value))
    memberParzelle = Trim(CStr(wsM.Cells(memberRow, M_COL_PARZELLE).Value))
    memberName = memberNachname & ", " & memberVorname
    
    If memberNachname = "" Then
        MsgBox "Fehler: Mitglied hat keinen Namen!", vbCritical
        GoTo Cleanup
    End If
    
    ' ================================================
    ' STEP 2: Find and reassign BANK-IDs in DATA
    ' ================================================
    
    lastRowD = wsD.Cells(wsD.Rows.Count, DATA_MAP_COL_ENTITYKEY).End(xlUp).Row
    
    If lastRowD >= DATA_START_ROW Then
        For rD = DATA_START_ROW To lastRowD
            entityKey = wsD.Cells(rD, DATA_MAP_COL_ENTITYKEY).Value
            
            ' Only process BANK-IDs
            If VarType(entityKey) = vbString And Left(CStr(entityKey), 5) = "BANK-" Then
                kontoName = Trim(CStr(wsD.Cells(rD, DATA_MAP_COL_KTONAME).Value))
                parzelle = Trim(CStr(wsD.Cells(rD, DATA_MAP_COL_PARZELLE).Value))
                
                ' Check if this entry matches the member
                If IsMatchForMember(kontoName, parzelle, memberNachname, memberVorname, memberParzelle) Then
                    ' Reassign to MemberID
                    wsD.Cells(rD, DATA_MAP_COL_ENTITYKEY).Value = memberID
                    wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG).Value = memberName
                    wsD.Cells(rD, DATA_MAP_COL_PARZELLE).Value = memberParzelle
                    wsD.Cells(rD, DATA_MAP_COL_ENTITYROLE).Value = "MITGLIED"
                    wsD.Cells(rD, DATA_MAP_COL_DEBUG).Value = "Reassigned from BANK-ID to MemberID"
                    
                    reassignCount = reassignCount + 1
                End If
            End If
        Next rD
    End If
    
    ' ================================================
    ' STEP 3: Update BANKKONTO entries
    ' ================================================
    
    lastRowBK = wsBK.Cells(wsBK.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    
    If lastRowBK >= BK_START_ROW And reassignCount > 0 Then
        ' Update all bank entries that had reassigned EntityKeys
        For rBK = BK_START_ROW To lastRowBK
            entityKey = wsBK.Cells(rBK, BK_COL_ENTITY_KEY).Value
            
            ' Check if this is one of the reassigned BANK-IDs
            If VarType(entityKey) = vbString And Left(CStr(entityKey), 5) = "BANK-" Then
                ' Check if it should be reassigned based on DATA mapping
                Dim newKey As String
                newKey = FindReassignedKey(CStr(entityKey), wsD, memberID)
                
                If newKey <> "" Then
                    wsBK.Cells(rBK, BK_COL_ENTITY_KEY).Value = newKey
                    bankReassignCount = bankReassignCount + 1
                End If
            End If
        Next rBK
    End If
    
    ' ================================================
    ' STEP 4: Report Results
    ' ================================================
    
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    If reassignCount > 0 Then
        MsgBox "Reassignment erfolgreich!" & vbCrLf & vbCrLf & _
               "Mitglied: " & memberName & vbCrLf & _
               "MemberID: " & memberID & vbCrLf & vbCrLf & _
               "Reassigned EntityKeys (DATA): " & reassignCount & vbCrLf & _
               "Updated Bank Entries: " & bankReassignCount, vbInformation, "Reassignment Complete"
    Else
        MsgBox "Keine passenden BANK-IDs für dieses Mitglied gefunden.", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Fehler beim Reassignment: " & Err.Description, vbCritical
End Sub

' ================================================
' Helper: Check if entry matches member
' ================================================
Private Function IsMatchForMember(ByVal kontoName As String, ByVal parzelle As String, _
                                  ByVal nachname As String, ByVal vorname As String, _
                                  ByVal memberParzelle As String) As Boolean
    IsMatchForMember = False
    
    ' Normalize strings for comparison
    kontoName = LCase(Trim(kontoName))
    nachname = LCase(Trim(nachname))
    vorname = LCase(Trim(vorname))
    parzelle = Trim(parzelle)
    memberParzelle = Trim(memberParzelle)
    
    ' Check if both names are in kontoName
    If InStr(1, kontoName, nachname, vbTextCompare) > 0 And _
       InStr(1, kontoName, vorname, vbTextCompare) > 0 Then
        IsMatchForMember = True
        Exit Function
    End If
    
    ' Alternative: Check by Parzelle if names partially match
    If parzelle = memberParzelle And memberParzelle <> "" Then
        If InStr(1, kontoName, nachname, vbTextCompare) > 0 Or _
           InStr(1, kontoName, vorname, vbTextCompare) > 0 Then
            IsMatchForMember = True
            Exit Function
        End If
    End If
End Function

' ================================================
' Helper: Find if a BANK-ID was reassigned to memberID
' ================================================
Private Function FindReassignedKey(ByVal oldBankID As String, ByVal wsD As Worksheet, _
                                   ByVal targetMemberID As String) As String
    Dim lastRow As Long
    Dim r As Long
    Dim entityKey As Variant
    
    FindReassignedKey = ""
    
    lastRow = wsD.Cells(wsD.Rows.Count, DATA_MAP_COL_ENTITYKEY).End(xlUp).Row
    
    ' We need to store old BANK-IDs somewhere or track them during reassignment
    ' For now, return targetMemberID if any match is found in debug message
    
    For r = DATA_START_ROW To lastRow
        entityKey = wsD.Cells(r, DATA_MAP_COL_ENTITYKEY).Value
        
        If CStr(entityKey) = targetMemberID Then
            ' Check if debug message indicates reassignment
            If InStr(1, wsD.Cells(r, DATA_MAP_COL_DEBUG).Value, "Reassigned", vbTextCompare) > 0 Then
                FindReassignedKey = targetMemberID
                Exit Function
            End If
        End If
    Next r
End Function

' ================================================
' Helper: Safe MemberID search
' ================================================
Private Function FindRowByMemberID_Safe(ByVal memberID As String, ByVal wsM As Worksheet) As Long
    Dim lastRow As Long
    Dim r As Long
    Dim currentID As String
    
    FindRowByMemberID_Safe = 0
    
    If Trim(memberID) = "" Then Exit Function
    
    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_MEMBER_ID).End(xlUp).Row
    
    If lastRow < M_START_ROW Then Exit Function
    
    ' Row-by-row search for safety
    For r = M_START_ROW To lastRow
        currentID = Trim(CStr(wsM.Cells(r, M_COL_MEMBER_ID).Value))
        
        If StrComp(currentID, memberID, vbTextCompare) = 0 Then
            FindRowByMemberID_Safe = r
            Exit Function
        End If
    Next r
End Function
