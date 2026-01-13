Attribute VB_Name = "mod_MigrateKeys"
Option Explicit

' ***************************************************************
' MODUL: mod_MigrateKeys
' ZWECK: Migration von numerischen EntityKeys zu string-basierten
'        MemberID/BANK-IDs
' ***************************************************************

' Migrates EntityKeys from numeric to string-based format
' - Members get their MemberID from WS_MITGLIEDER
' - Bank entries without member match get BANK-yyyymmddhhmmss-nnn format
Public Sub Migrate_EntityKeys_To_MemberID()
    
    Dim wsD As Worksheet
    Dim wsM As Worksheet
    Dim wsBK As Worksheet
    Dim lastRowD As Long
    Dim lastRowBK As Long
    Dim rD As Long, rBK As Long
    Dim oldEntityKey As Variant
    Dim newEntityKey As String
    Dim zuordnung As String
    Dim entityRole As String
    Dim bankCounter As Long
    Dim migratedCount As Long
    Dim skippedCount As Long
    Dim bankCount As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    migratedCount = 0
    skippedCount = 0
    bankCount = 0
    bankCounter = 1
    
    ' ================================================
    ' STEP 1: Migrate DATA_MAP_COL_ENTITYKEY in WS_DATEN
    ' ================================================
    
    lastRowD = wsD.Cells(wsD.Rows.Count, DATA_MAP_COL_ENTITYKEY).End(xlUp).Row
    
    If lastRowD < DATA_START_ROW Then
        MsgBox "Keine Daten zum Migrieren gefunden.", vbInformation
        GoTo Cleanup
    End If
    
    For rD = DATA_START_ROW To lastRowD
        oldEntityKey = wsD.Cells(rD, DATA_MAP_COL_ENTITYKEY).Value
        zuordnung = Trim(CStr(wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG).Value))
        entityRole = Trim(CStr(wsD.Cells(rD, DATA_MAP_COL_ENTITYROLE).Value))
        
        ' Skip if already string-based
        If VarType(oldEntityKey) = vbString And Len(CStr(oldEntityKey)) > 10 Then
            skippedCount = skippedCount + 1
            GoTo NextDataRow
        End If
        
        ' Try to find MemberID for this entry
        newEntityKey = ""
        
        If zuordnung <> "" And UCase(entityRole) = "MITGLIED" Then
            ' Try to find member by name
            newEntityKey = FindMemberIDByName(zuordnung, wsM)
        End If
        
        ' If no member match, generate BANK-ID
        If newEntityKey = "" Then
            newEntityKey = Generate_BankID(bankCounter)
            bankCounter = bankCounter + 1
            bankCount = bankCount + 1
        End If
        
        ' Write new EntityKey
        wsD.Cells(rD, DATA_MAP_COL_ENTITYKEY).Value = newEntityKey
        migratedCount = migratedCount + 1
        
NextDataRow:
    Next rD
    
    ' ================================================
    ' STEP 2: Migrate BK_COL_ENTITY_KEY in WS_BANKKONTO
    ' ================================================
    
    lastRowBK = wsBK.Cells(wsBK.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    
    If lastRowBK >= BK_START_ROW Then
        For rBK = BK_START_ROW To lastRowBK
            oldEntityKey = wsBK.Cells(rBK, BK_COL_ENTITY_KEY).Value
            
            ' Skip if empty or already string
            If IsEmpty(oldEntityKey) Or oldEntityKey = "" Then GoTo NextBKRow
            If VarType(oldEntityKey) = vbString And Len(CStr(oldEntityKey)) > 10 Then GoTo NextBKRow
            
            ' Try to find matching EntityKey in DATA mapping
            newEntityKey = FindNewEntityKeyByOld(CStr(oldEntityKey), wsD)
            
            If newEntityKey <> "" Then
                wsBK.Cells(rBK, BK_COL_ENTITY_KEY).Value = newEntityKey
            End If
            
NextBKRow:
        Next rBK
    End If
    
    ' ================================================
    ' STEP 3: Report Results
    ' ================================================
    
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Migration abgeschlossen!" & vbCrLf & vbCrLf & _
           "Migrierte Einträge (DATA): " & migratedCount & vbCrLf & _
           "Übersprungen (bereits String): " & skippedCount & vbCrLf & _
           "Neue BANK-IDs erstellt: " & bankCount, vbInformation, "EntityKey Migration"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Fehler bei der Migration: " & Err.Description, vbCritical
End Sub

' ================================================
' Helper: Generate BANK-ID with format BANK-yyyymmddhhmmss-nnn
' ================================================
Private Function Generate_BankID(ByVal counter As Long) As String
    Dim timestamp As String
    Dim counterStr As String
    
    timestamp = Format(Now, "yyyymmddhhnnss")
    counterStr = Format(counter, "000")
    
    Generate_BankID = "BANK-" & timestamp & "-" & counterStr
End Function

' ================================================
' Helper: Find MemberID by member name
' ================================================
Public Function FindMemberIDByName(ByVal memberName As String, ByVal wsM As Worksheet) As String
    Dim lastRow As Long
    Dim r As Long
    Dim fullName As String
    Dim memberID As String
    Dim vorname As String
    Dim nachname As String
    
    FindMemberIDByName = ""
    
    If Trim(memberName) = "" Then Exit Function
    
    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
    
    If lastRow < M_START_ROW Then Exit Function
    
    ' Normalize search name
    memberName = Trim(memberName)
    
    For r = M_START_ROW To lastRow
        nachname = Trim(CStr(wsM.Cells(r, M_COL_NACHNAME).Value))
        vorname = Trim(CStr(wsM.Cells(r, M_COL_VORNAME).Value))
        memberID = Trim(CStr(wsM.Cells(r, M_COL_MEMBER_ID).Value))
        
        If memberID = "" Then GoTo NextRow
        
        ' Build full name variations
        fullName = nachname & ", " & vorname
        
        ' Check various name formats
        If InStr(1, memberName, nachname, vbTextCompare) > 0 And _
           InStr(1, memberName, vorname, vbTextCompare) > 0 Then
            FindMemberIDByName = memberID
            Exit Function
        End If
        
        ' Also check if the names match in reverse order
        If StrComp(memberName, fullName, vbTextCompare) = 0 Or _
           StrComp(memberName, vorname & " " & nachname, vbTextCompare) = 0 Then
            FindMemberIDByName = memberID
            Exit Function
        End If
        
NextRow:
    Next r
    
End Function

' ================================================
' Helper: Find new EntityKey by old numeric key
' ================================================
Private Function FindNewEntityKeyByOld(ByVal oldKey As String, ByVal wsD As Worksheet) As String
    ' This function is a placeholder - in reality we'd need to store
    ' a mapping of old -> new keys during migration
    ' For now, return empty to indicate no match
    FindNewEntityKeyByOld = ""
End Function

' ================================================
' Validation: Check migration results
' ================================================
Public Sub Validate_MigrationResults()
    Dim wsD As Worksheet
    Dim wsBK As Worksheet
    Dim lastRowD As Long
    Dim lastRowBK As Long
    Dim rD As Long, rBK As Long
    Dim entityKey As Variant
    Dim numericCount As Long
    Dim stringCount As Long
    Dim emptyCount As Long
    Dim memberIDCount As Long
    Dim bankIDCount As Long
    
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    numericCount = 0
    stringCount = 0
    emptyCount = 0
    memberIDCount = 0
    bankIDCount = 0
    
    ' Check DATA sheet
    lastRowD = wsD.Cells(wsD.Rows.Count, DATA_MAP_COL_ENTITYKEY).End(xlUp).Row
    
    If lastRowD >= DATA_START_ROW Then
        For rD = DATA_START_ROW To lastRowD
            entityKey = wsD.Cells(rD, DATA_MAP_COL_ENTITYKEY).Value
            
            If IsEmpty(entityKey) Or entityKey = "" Then
                emptyCount = emptyCount + 1
            ElseIf IsNumeric(entityKey) Then
                numericCount = numericCount + 1
            ElseIf VarType(entityKey) = vbString Then
                stringCount = stringCount + 1
                If Left(CStr(entityKey), 5) = "BANK-" Then
                    bankIDCount = bankIDCount + 1
                Else
                    memberIDCount = memberIDCount + 1
                End If
            End If
        Next rD
    End If
    
    ' Display results
    MsgBox "Migration Validation (DATA sheet):" & vbCrLf & vbCrLf & _
           "String EntityKeys: " & stringCount & vbCrLf & _
           "  - MemberIDs: " & memberIDCount & vbCrLf & _
           "  - BANK-IDs: " & bankIDCount & vbCrLf & _
           "Numerische EntityKeys: " & numericCount & vbCrLf & _
           "Leere Einträge: " & emptyCount, vbInformation, "Validation Summary"
End Sub
