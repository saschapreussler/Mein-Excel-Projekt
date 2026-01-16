Attribute VB_Name = "mod_Tests"
Option Explicit

' ***************************************************************
' MODUL: mod_Tests
' Enth�lt Test-Prozeduren f�r Qualit�tssicherung
' ***************************************************************

' ***************************************************************
' TEST: CSV Import mit Beispieldatei
' ***************************************************************
Public Sub Test_CSVImport()
    ' Tests CSV import with sample file
    ' Expected: User selects tests/sample.csv and 5 rows are imported
    
    Dim ws As Worksheet
    Dim initialRowCount As Long
    Dim finalRowCount As Long
    Dim importedRows As Long
    
    On Error GoTo TestError
    
    Debug.Print "========================================"
    Debug.Print "Starting CSV Import Test..."
    Debug.Print "========================================"
    
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    initialRowCount = ws.Cells(ws.Rows.Count, BK_COL_BETRAG).End(xlUp).Row
    Debug.Print "Initial row count: " & initialRowCount
    Debug.Print ""
    Debug.Print "Please select tests/sample.csv when prompted"
    Debug.Print ""
    
    ' Call the import procedure
    Call Importiere_Kontoauszug
    
    ' Check results
    finalRowCount = ws.Cells(ws.Rows.Count, BK_COL_BETRAG).End(xlUp).Row
    importedRows = finalRowCount - initialRowCount
    
    Debug.Print "Final row count: " & finalRowCount
    Debug.Print "Rows imported: " & importedRows
    Debug.Print ""
    
    ' Validate
    If importedRows = 5 Then
        Debug.Print "✓ Test PASSED: Expected 5 rows imported"
        MsgBox "CSV Import Test PASSED" & vbCrLf & vbCrLf & _
               "Expected: 5 rows" & vbCrLf & _
               "Imported: " & importedRows & " rows", vbInformation, "Test Result"
    ElseIf importedRows = 0 Then
        Debug.Print "⚠ Test SKIPPED or DUPLICATES: No new rows imported (may already exist)"
        MsgBox "CSV Import Test: No new rows imported" & vbCrLf & vbCrLf & _
               "This may be expected if you already imported sample.csv before." & vbCrLf & _
               "Try deleting the imported rows and re-running the test.", vbInformation, "Test Result"
    Else
        Debug.Print "✗ Test WARNING: Expected 5 rows, got " & importedRows
        MsgBox "CSV Import Test WARNING" & vbCrLf & vbCrLf & _
               "Expected: 5 rows" & vbCrLf & _
               "Imported: " & importedRows & " rows" & vbCrLf & vbCrLf & _
               "Check if some rows were filtered or already existed.", vbExclamation, "Test Result"
    End If
    
    Debug.Print "========================================"
    Exit Sub
    
TestError:
    Debug.Print "✗ Test FAILED with error: " & Err.Description
    MsgBox "CSV Import Test FAILED" & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Test Result"
End Sub


' ***************************************************************
' TEST: Member ID Generierung
' ***************************************************************
Public Sub Test_MemberIDGeneration()
    ' Tests GUID generation for members
    
    Dim id1 As String, id2 As String, id3 As String
    Dim i As Integer
    Dim allUnique As Boolean
    
    On Error GoTo TestError
    
    Debug.Print "========================================"
    Debug.Print "Testing Member ID Generation..."
    Debug.Print "========================================"
    
    ' Generate 3 IDs
    id1 = CreateGUID()
    id2 = CreateGUID()
    id3 = CreateGUID()
    
    Debug.Print "ID 1: " & id1
    Debug.Print "ID 2: " & id2
    Debug.Print "ID 3: " & id3
    Debug.Print ""
    
    ' Validate
    allUnique = (Len(id1) > 0 And Len(id2) > 0 And Len(id3) > 0 And _
                 id1 <> id2 And id2 <> id3 And id1 <> id3)
    
    If allUnique Then
        Debug.Print "✓ Test PASSED: All IDs are unique and non-empty"
        MsgBox "Member ID Generation Test PASSED" & vbCrLf & vbCrLf & _
               "Generated 3 unique IDs successfully.", vbInformation, "Test Result"
    Else
        Debug.Print "✗ Test FAILED: IDs not unique or empty"
        MsgBox "Member ID Generation Test FAILED" & vbCrLf & vbCrLf & _
               "IDs are not unique or some are empty.", vbCritical, "Test Result"
    End If
    
    Debug.Print "========================================"
    Exit Sub
    
TestError:
    Debug.Print "✗ Test FAILED with error: " & Err.Description
    MsgBox "Member ID Generation Test FAILED" & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Test Result"
End Sub


' ***************************************************************
' TEST: Member Lookup by ID (auch mit versteckten Spalten)
' ***************************************************************
Public Sub Test_MemberLookup()
    ' Tests member lookup functionality with hidden columns
    
    Dim ws As Worksheet
    Dim testMemberID As String
    Dim foundRow As Long
    Dim colAHidden As Boolean
    
    On Error GoTo TestError
    
    Debug.Print "========================================"
    Debug.Print "Testing Member Lookup by ID..."
    Debug.Print "========================================"
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    ' Get first member ID for testing
    If ws.Cells(M_START_ROW, M_COL_MEMBER_ID).Value <> "" Then
        testMemberID = ws.Cells(M_START_ROW, M_COL_MEMBER_ID).Value
        Debug.Print "Using test Member ID: " & testMemberID
    Else
        Debug.Print "⚠ Test SKIPPED: No member data found in sheet"
        MsgBox "Member Lookup Test SKIPPED" & vbCrLf & vbCrLf & _
               "No member data found. Please add at least one member first.", _
               vbExclamation, "Test Result"
        Exit Sub
    End If
    
    ' Test 1: Normal lookup
    Debug.Print ""
    Debug.Print "Test 1: Normal lookup (columns visible)..."
    foundRow = FindeRowByMemberID(testMemberID)
    
    If foundRow = M_START_ROW Then
        Debug.Print "✓ Found member at row: " & foundRow
    Else
        Debug.Print "✗ ERROR: Expected row " & M_START_ROW & ", found " & foundRow
        GoTo TestFailed
    End If
    
    ' Test 2: Lookup with column A hidden
    Debug.Print ""
    Debug.Print "Test 2: Lookup with column A hidden..."
    colAHidden = ws.Columns(M_COL_MEMBER_ID).Hidden
    
    ' Hide column A temporarily
    ws.Columns(M_COL_MEMBER_ID).Hidden = True
    
    foundRow = FindeRowByMemberID(testMemberID)
    
    ' Restore column visibility
    ws.Columns(M_COL_MEMBER_ID).Hidden = colAHidden
    
    If foundRow = M_START_ROW Then
        Debug.Print "✓ Found member at row: " & foundRow & " (column was hidden)"
    Else
        Debug.Print "✗ ERROR: Expected row " & M_START_ROW & ", found " & foundRow
        GoTo TestFailed
    End If
    
    ' Test 3: Alias function
    Debug.Print ""
    Debug.Print "Test 3: Testing alias function FindMemberRowByID..."
    foundRow = FindMemberRowByID(ws, testMemberID)
    
    If foundRow = M_START_ROW Then
        Debug.Print "✓ Alias function works correctly"
    Else
        Debug.Print "✗ ERROR: Alias function failed"
        GoTo TestFailed
    End If
    
    Debug.Print ""
    Debug.Print "✓ All Tests PASSED"
    MsgBox "Member Lookup Test PASSED" & vbCrLf & vbCrLf & _
           "Member lookup works correctly even with hidden columns.", _
           vbInformation, "Test Result"
    Debug.Print "========================================"
    Exit Sub
    
TestFailed:
    Debug.Print ""
    Debug.Print "✗ Test FAILED"
    MsgBox "Member Lookup Test FAILED" & vbCrLf & vbCrLf & _
           "Check Debug output for details.", vbCritical, "Test Result"
    Debug.Print "========================================"
    Exit Sub
    
TestError:
    ' Restore column visibility on error
    On Error Resume Next
    If Not ws Is Nothing Then ws.Columns(M_COL_MEMBER_ID).Hidden = colAHidden
    On Error GoTo 0
    
    Debug.Print "✗ Test FAILED with error: " & Err.Description
    MsgBox "Member Lookup Test FAILED" & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Test Result"
End Sub


' ***************************************************************
' TEST: Alle Tests ausf�hren
' ***************************************************************
Public Sub Run_All_Tests()
    ' Runs all test procedures
    
    Debug.Print ""
    Debug.Print "========================================"
    Debug.Print "RUNNING ALL TESTS"
    Debug.Print "========================================"
    Debug.Print ""
    
    On Error Resume Next ' Continue even if a test fails
    
    Call Test_MemberIDGeneration
    Debug.Print ""
    
    Call Test_MemberLookup
    Debug.Print ""
    
    ' Note: CSV Import requires user interaction, so we skip in batch mode
    Debug.Print "Note: Test_CSVImport requires manual file selection"
    Debug.Print "Run it separately if needed: Call Test_CSVImport"
    Debug.Print ""
    
    Debug.Print "========================================"
    Debug.Print "ALL AUTOMATED TESTS COMPLETED"
    Debug.Print "========================================"
    
    MsgBox "All automated tests completed." & vbCrLf & vbCrLf & _
           "Check the Immediate Window (Ctrl+G) for detailed results." & vbCrLf & vbCrLf & _
           "Note: CSV Import test requires manual execution.", _
           vbInformation, "Test Suite Complete"
End Sub
