Attribute VB_Name = "mod_Tests"
Option Explicit

' ***************************************************************
' MODUL: mod_Tests
' ZWECK: Test macros for validating functionality
' ***************************************************************

' Test macro for CSV import functionality
Public Sub Test_CSVImport()
    
    Dim testFile As String
    Dim wsBank As Worksheet
    Dim initialRowCount As Long
    Dim finalRowCount As Long
    Dim expectedRows As Long
    Dim testResult As String
    
    On Error GoTo TestError
    
    ' Setup
    testFile = ThisWorkbook.Path & "\tests\sample.csv"
    Set wsBank = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    ' Check if test file exists
    If Dir(testFile) = "" Then
        MsgBox "Test file not found: " & testFile & vbCrLf & _
               "Please ensure tests/sample.csv exists in the project directory.", vbCritical, "Test Failed"
        Exit Sub
    End If
    
    ' Count initial rows
    initialRowCount = wsBank.Cells(wsBank.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    If initialRowCount < BK_START_ROW Then initialRowCount = BK_START_ROW - 1
    
    ' Expected number of rows in sample.csv (4 data rows + 1 header = 4 data rows)
    expectedRows = 4
    
    MsgBox "Starting CSV Import Test..." & vbCrLf & _
           "Test file: " & testFile & vbCrLf & _
           "Initial row count: " & (initialRowCount - BK_START_ROW + 1), vbInformation
    
    ' Run the import (user will need to manually select the test file)
    ' TODO: Automate file selection for testing
    ' For now, manually select tests/sample.csv when prompted
    Call Importiere_Kontoauszug
    
    ' Count final rows
    finalRowCount = wsBank.Cells(wsBank.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    If finalRowCount < BK_START_ROW Then finalRowCount = BK_START_ROW - 1
    
    ' Validate
    Dim importedRows As Long
    importedRows = finalRowCount - initialRowCount
    
    If importedRows > 0 Then
        testResult = "CSV Import Test PASSED" & vbCrLf & _
                    "Rows imported: " & importedRows & vbCrLf & _
                    "Note: If running multiple times, duplicates will be skipped."
        MsgBox testResult, vbInformation, "Test Result"
    ElseIf importedRows = 0 Then
        testResult = "CSV Import Test - No new rows imported" & vbCrLf & _
                    "This is expected if test was run multiple times (duplicate detection)."
        MsgBox testResult, vbExclamation, "Test Result"
    Else
        testResult = "CSV Import Test FAILED" & vbCrLf & _
                    "Row count decreased unexpectedly."
        MsgBox testResult, vbCritical, "Test Result"
    End If
    
    Exit Sub

TestError:
    MsgBox "Error in Test_CSVImport: " & Err.Description, vbCritical, "Test Error"
    
End Sub
