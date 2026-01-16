# Test Files

This directory contains test data and test procedures for the VBA project.

## Sample Data Files

### sample.csv

A sample bank transaction CSV file for testing the import functionality.

**Format:** Semicolon-delimited, UTF-8 encoded
**Columns:**
- Buchungstag (Booking Date)
- Wertstellung (Value Date)
- Buchungstext (Transaction Text)
- Verwendungszweck (Purpose)
- Beguenstigter/Zahlungspflichtiger (Beneficiary/Payer)
- Kontonummer (IBAN)
- BLZ (Bank Code) - usually empty for IBAN
- Betrag (Amount)
- Waehrung (Currency)

## Test Procedures

### Testing CSV Import

1. Open the workbook
2. Go to the "Bankkonto" sheet
3. Click the CSV import button or run:
   ```vba
   Call Importiere_Kontoauszug
   ```
4. Select `tests/sample.csv`
5. Verify:
   - 5 transactions imported
   - Amounts correctly formatted
   - No duplicates on re-import

### Testing Member ID Lookup

1. Ensure "Mitglieder" sheet has data
2. In VBA Immediate Window (Ctrl+G):
   ```vba
   Dim testRow As Long
   testRow = FindMemberRowByID(ActiveSheet, "test-guid-here")
   Debug.Print "Found at row: " & testRow
   ```

### Testing Form Workflows

**Member Management Form:**
1. Run: `frm_Mitgliederverwaltung.Show`
2. Test adding a new member
3. Test editing existing member
4. Test member ID generation

**Member Data Form:**
1. Select a member row
2. Run the form
3. Test field validation
4. Test save operation

## Automated Test Macros

The following test macros are available in the VBA project:

### Test_CSVImport

```vba
Public Sub Test_CSVImport()
    ' Tests CSV import with sample file
    ' Expected: 5 rows imported successfully
    Debug.Print "Starting CSV Import Test..."
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    Dim initialRowCount As Long
    initialRowCount = ws.Cells(ws.Rows.Count, BK_COL_BETRAG).End(xlUp).Row
    
    Debug.Print "Initial row count: " & initialRowCount
    Debug.Print "Please select tests/sample.csv when prompted"
    
    Call Importiere_Kontoauszug
    
    Dim finalRowCount As Long
    finalRowCount = ws.Cells(ws.Rows.Count, BK_COL_BETRAG).End(xlUp).Row
    
    Debug.Print "Final row count: " & finalRowCount
    Debug.Print "Rows imported: " & (finalRowCount - initialRowCount)
    
    If finalRowCount - initialRowCount = 5 Then
        Debug.Print "✓ Test PASSED: Expected 5 rows imported"
    Else
        Debug.Print "✗ Test FAILED: Expected 5 rows, got " & (finalRowCount - initialRowCount)
    End If
End Sub
```

### Test_MemberIDGeneration

```vba
Public Sub Test_MemberIDGeneration()
    ' Tests GUID generation for members
    Debug.Print "Testing Member ID Generation..."
    
    Dim id1 As String, id2 As String
    id1 = CreateGUID()
    id2 = CreateGUID()
    
    Debug.Print "ID 1: " & id1
    Debug.Print "ID 2: " & id2
    
    If Len(id1) > 0 And Len(id2) > 0 And id1 <> id2 Then
        Debug.Print "✓ Test PASSED: Unique IDs generated"
    Else
        Debug.Print "✗ Test FAILED: IDs not unique or empty"
    End If
End Sub
```

## Creating New Test Files

When adding new test files:

1. Use realistic but anonymized data
2. Document the file format
3. Keep files small (< 100 rows for CSV)
4. Use UTF-8 encoding
5. Add test instructions above

## Test Checklist

Before submitting a PR, verify:

- [ ] `Debug → Compile VBAProject` succeeds
- [ ] All test macros pass
- [ ] Forms open without errors
- [ ] CSV import works with `sample.csv`
- [ ] No data corruption in existing sheets
- [ ] Member ID lookup works with hidden columns
- [ ] Sort functions maintain data integrity
