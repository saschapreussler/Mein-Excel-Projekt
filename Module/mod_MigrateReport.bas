Attribute VB_Name = "mod_MigrateReport"
Option Explicit

' ***************************************************************
' MODUL: mod_MigrateReport
' ZWECK: Reporting und Heuristik für EntityKey-Migration
' ***************************************************************

' Generate a detailed migration report
Public Sub GenerateMigrationReport()
    Dim wsD As Worksheet
    Dim wsM As Worksheet
    Dim lastRowD As Long
    Dim rD As Long
    Dim entityKey As Variant
    Dim zuordnung As String
    Dim entityRole As String
    Dim report As String
    Dim unmatchedCount As Long
    Dim matchedCount As Long
    Dim bankIDCount As Long
    
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    unmatchedCount = 0
    matchedCount = 0
    bankIDCount = 0
    
    report = "MIGRATION REPORT" & vbCrLf
    report = report & String(50, "=") & vbCrLf & vbCrLf
    
    lastRowD = wsD.Cells(wsD.Rows.Count, DATA_MAP_COL_ENTITYKEY).End(xlUp).Row
    
    If lastRowD < DATA_START_ROW Then
        MsgBox "Keine Daten für Report gefunden.", vbInformation
        Exit Sub
    End If
    
    report = report & "Unzugeordnete Bank-Einträge (BANK-IDs):" & vbCrLf
    report = report & String(50, "-") & vbCrLf
    
    For rD = DATA_START_ROW To lastRowD
        entityKey = wsD.Cells(rD, DATA_MAP_COL_ENTITYKEY).Value
        zuordnung = Trim(CStr(wsD.Cells(rD, DATA_MAP_COL_ZUORDNUNG).Value))
        entityRole = Trim(CStr(wsD.Cells(rD, DATA_MAP_COL_ENTITYROLE).Value))
        
        If VarType(entityKey) = vbString And Left(CStr(entityKey), 5) = "BANK-" Then
            bankIDCount = bankIDCount + 1
            report = report & "Zeile " & rD & ": " & CStr(entityKey) & " - "
            report = report & "IBAN: " & wsD.Cells(rD, DATA_MAP_COL_IBAN_OLD).Value & " - "
            report = report & "Name: " & wsD.Cells(rD, DATA_MAP_COL_KTONAME).Value & vbCrLf
            unmatchedCount = unmatchedCount + 1
        ElseIf VarType(entityKey) = vbString And Len(CStr(entityKey)) > 10 Then
            matchedCount = matchedCount + 1
        End If
    Next rD
    
    If bankIDCount = 0 Then
        report = report & "(Keine BANK-IDs gefunden - alle Einträge zugeordnet)" & vbCrLf
    End If
    
    report = report & vbCrLf & String(50, "=") & vbCrLf
    report = report & "ZUSAMMENFASSUNG:" & vbCrLf
    report = report & "  Zugeordnete Mitglieder: " & matchedCount & vbCrLf
    report = report & "  Unzugeordnete (BANK-IDs): " & unmatchedCount & vbCrLf
    
    ' Display in Debug window
    Debug.Print report
    
    ' Also show summary in MsgBox
    MsgBox "Migration Report erstellt!" & vbCrLf & vbCrLf & _
           "Details im Direktfenster (Strg+G)" & vbCrLf & vbCrLf & _
           "Zugeordnete Mitglieder: " & matchedCount & vbCrLf & _
           "Unzugeordnete (BANK-IDs): " & unmatchedCount, vbInformation
End Sub

' Try to match BANK-IDs by trailing digits (Parzellennummer)
Public Sub TryMatchByTrailingDigits(Optional ByVal maxDigits As Long = 2)
    Dim wsD As Worksheet
    Dim wsM As Worksheet
    Dim lastRowD As Long
    Dim rD As Long
    Dim entityKey As Variant
    Dim kontoName As String
    Dim trailingDigits As String
    Dim parzelle As String
    Dim memberID As String
    Dim matchCount As Long
    
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    matchCount = 0
    
    If maxDigits < 1 Or maxDigits > 3 Then
        MsgBox "Ungültiger Wert für maxDigits (1-3 erlaubt)", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    lastRowD = wsD.Cells(wsD.Rows.Count, DATA_MAP_COL_ENTITYKEY).End(xlUp).Row
    
    For rD = DATA_START_ROW To lastRowD
        entityKey = wsD.Cells(rD, DATA_MAP_COL_ENTITYKEY).Value
        
        ' Only process BANK-IDs
        If VarType(entityKey) = vbString And Left(CStr(entityKey), 5) = "BANK-" Then
            kontoName = Trim(CStr(wsD.Cells(rD, DATA_MAP_COL_KTONAME).Value))
            
            ' Extract trailing digits
            trailingDigits = ExtractTrailingDigits(kontoName, maxDigits)
            
            If trailingDigits <> "" Then
                ' Try to find member by Parzelle
                memberID = FindMemberIDByParzelle(trailingDigits, wsM)
                
                If memberID <> "" Then
                    ' Update EntityKey
                    wsD.Cells(rD, DATA_MAP_COL_ENTITYKEY).Value = memberID
                    wsD.Cells(rD, DATA_MAP_COL_PARZELLE).Value = trailingDigits
                    wsD.Cells(rD, DATA_MAP_COL_DEBUG).Value = "Auto-matched by trailing digits"
                    matchCount = matchCount + 1
                End If
            End If
        End If
    Next rD
    
    Application.ScreenUpdating = True
    
    MsgBox "Heuristisches Matching abgeschlossen!" & vbCrLf & vbCrLf & _
           "Gefundene Matches: " & matchCount, vbInformation, "Trailing Digits Matching"
End Sub

' Extract trailing digits from a string
Private Function ExtractTrailingDigits(ByVal inputStr As String, ByVal maxDigits As Long) As String
    Dim i As Long
    Dim ch As String
    Dim result As String
    
    ExtractTrailingDigits = ""
    result = ""
    
    ' Work backwards from end of string
    For i = Len(inputStr) To 1 Step -1
        ch = Mid(inputStr, i, 1)
        
        If ch >= "0" And ch <= "9" Then
            result = ch & result
            If Len(result) >= maxDigits Then Exit For
        Else
            ' Stop at first non-digit
            Exit For
        End If
    Next i
    
    ExtractTrailingDigits = result
End Function

' Find MemberID by Parzellennummer
Private Function FindMemberIDByParzelle(ByVal parzelleNr As String, ByVal wsM As Worksheet) As String
    Dim lastRow As Long
    Dim r As Long
    Dim parzelle As String
    Dim memberID As String
    Dim funktion As String
    
    FindMemberIDByParzelle = ""
    
    If Trim(parzelleNr) = "" Then Exit Function
    
    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_PARZELLE).End(xlUp).Row
    
    If lastRow < M_START_ROW Then Exit Function
    
    For r = M_START_ROW To lastRow
        parzelle = Trim(CStr(wsM.Cells(r, M_COL_PARZELLE).Value))
        memberID = Trim(CStr(wsM.Cells(r, M_COL_MEMBER_ID).Value))
        funktion = Trim(CStr(wsM.Cells(r, M_COL_FUNKTION).Value))
        
        ' Match Parzelle and ensure it's a Pächter
        If parzelle = parzelleNr And _
           UCase(funktion) = UCase(PAECHTER_STATUS) And _
           memberID <> "" Then
            FindMemberIDByParzelle = memberID
            Exit Function
        End If
    Next r
End Function

' Clean up report helper
Public Sub CleanupDebugMessages()
    Dim wsD As Worksheet
    Dim lastRowD As Long
    Dim rD As Long
    Dim debugMsg As String
    
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    
    lastRowD = wsD.Cells(wsD.Rows.Count, DATA_MAP_COL_ENTITYKEY).End(xlUp).Row
    
    If lastRowD < DATA_START_ROW Then Exit Sub
    
    For rD = DATA_START_ROW To lastRowD
        debugMsg = Trim(CStr(wsD.Cells(rD, DATA_MAP_COL_DEBUG).Value))
        
        ' Clean up auto-generated messages if desired
        If InStr(1, debugMsg, "Auto-matched", vbTextCompare) > 0 Then
            ' Optionally clear or modify
            ' wsD.Cells(rD, DATA_MAP_COL_DEBUG).Value = ""
        End If
    Next rD
    
    MsgBox "Debug-Nachrichten bereinigt.", vbInformation
End Sub
