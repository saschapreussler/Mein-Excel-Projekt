# Contributing to Mein Excel Projekt

Thank you for contributing to this VBA project! This guide will help you understand the development workflow.

## Project Structure

```
Mein-Excel-Projekt/
├── project_exports/     # VBA source code (version controlled)
│   ├── *.bas           # Standard modules
│   ├── *.cls           # Class modules
│   ├── *.frm           # UserForm code
│   └── *.frx           # UserForm resources (not version controlled)
├── tools/              # Development utilities
│   ├── Modul1.bas     # Export helper
│   └── README.md      # Tools documentation
├── tests/              # Test files and sample data
├── .gitignore         # Git ignore patterns
└── CONTRIBUTING.md    # This file
```

## Development Workflow

### 1. Setting Up Your Environment

1. Clone the repository:
   ```bash
   git clone https://github.com/saschapreussler/Mein-Excel-Projekt.git
   cd Mein-Excel-Projekt
   ```

2. Create a new workbook or copy an existing `.xlsm` file (not tracked in git)

3. Open the workbook in Excel

### 2. Importing Modules

To load the VBA code into your workbook:

1. Open Excel workbook
2. Press `Alt+F11` to open VBA Editor
3. For each module in `project_exports/`:
   - File → Import File
   - Select the file (`.bas`, `.cls`, `.frm`)
   - Click Open

**Tip:** Import in this order for dependency resolution:
- `mod_Const.bas` first (constants)
- Other `mod_*.bas` files (standard modules)
- `frm_*.frm` files (forms)
- Worksheet modules (`Tabelle*.bas`)
- `DieseArbeitsmappe.bas` last

### 3. Making Changes

1. Create a feature branch:
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. Make your code changes in the VBA Editor

3. **Important:** Test your changes:
   - Debug → Compile VBAProject (Ctrl+F5)
   - Run affected macros manually
   - Verify no runtime errors

### 4. Exporting Changes

After making changes, export the modules:

1. Import `tools/Modul1.bas` into the workbook temporarily
2. Run the macro `ExportAllVBA`
3. All modules are exported to `project_exports/`
4. Remove `Modul1` from the VBA project
5. Save and close the workbook

### 5. Running Static Checks

Before committing, verify module balance:

```bash
python3 tools/check_vba_balance.py
```

This ensures every Sub/Function has a matching End Sub/End Function.

### 6. Committing Changes

```bash
git add project_exports/
git commit -m "Description of your changes"
git push origin feature/your-feature-name
```

### 7. Creating a Pull Request

1. Go to GitHub repository
2. Click "New Pull Request"
3. Select your feature branch
4. Fill in PR description:
   - What changed
   - Why it changed
   - Testing performed
   - Compile status

## Code Standards

### Naming Conventions

- **Modules:** `mod_ModuleName` (PascalCase)
- **Forms:** `frm_FormName` (PascalCase)
- **Public Procedures:** `ProcedureName` (PascalCase)
- **Private Procedures:** `ProcedureName` (PascalCase)
- **Variables:** `variableName` (camelCase) or `strVariableName` (Hungarian notation)
- **Constants:** `CONSTANT_NAME` (UPPER_SNAKE_CASE)

### Code Style

```vba
' Use Option Explicit in all modules
Option Explicit

' Use meaningful names
Public Sub CalculateTotalAmount()
    Dim totalAmount As Double
    Dim rowIndex As Long
    
    ' Comment complex logic
    For rowIndex = 2 To lastRow
        totalAmount = totalAmount + ws.Cells(rowIndex, 5).Value
    Next rowIndex
End Sub

' Include error handling for public procedures
Public Sub ImportData()
    On Error GoTo ErrorHandler
    
    ' Main logic here
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
```

### Best Practices

1. **Always use `Option Explicit`** at the top of each module
2. **Error handling:** Add error handlers to public procedures
3. **Screen updating:** Disable for batch operations
   ```vba
   Application.ScreenUpdating = False
   ' ... operations ...
   Application.ScreenUpdating = True
   ```
4. **Constants:** Define magic numbers and strings as constants in `mod_Const`
5. **Comments:** Document complex logic, not obvious statements
6. **Worksheet protection:** Use `UnprotectSheet` / `ProtectSheet` helpers

## Testing

### Manual Testing

Before committing:

1. **Compile Check:**
   - Open VBA Editor
   - Debug → Compile VBAProject
   - Fix any compilation errors

2. **Run Key Macros:**
   - `Initialize_ImportReport_ListBox`
   - `Importiere_Kontoauszug` (with sample CSV in `tests/`)
   - Form workflows (open each form and test inputs)

3. **Check Data Integrity:**
   - Verify no data loss
   - Check formulas still work
   - Validate dropdown lists

### Sample Data

Use files in `tests/` directory for testing:
- `tests/sample.csv` - Sample bank transactions

### Test Macros

Run these test procedures:

```vba
' In VBA Immediate Window (Ctrl+G):
Call Test_CSVImport
Call Fuelle_MemberIDs_Wenn_Fehlend
Call Sortiere_Mitgliederliste_Nach_Parzelle
```

## Common Issues

### "Method or Data Member Not Found"

- Module not imported
- Typo in procedure/variable name
- Module reference missing

**Solution:** Re-import all modules in correct order

### "Type Mismatch"

- Data type mismatch
- Range contains unexpected values

**Solution:** Add validation and error handling

### "Object Required"

- Variable not set with `Set`
- Object is Nothing

**Solution:** Use `If Not obj Is Nothing Then`

## Getting Help

- Check `tools/README.md` for development tools
- Review existing code in `project_exports/` for examples
- Ask questions in GitHub Issues

## Rollback Instructions

If something goes wrong:

```bash
# Discard uncommitted changes
git checkout -- project_exports/

# Revert last commit
git revert HEAD

# Reset to specific commit
git reset --hard <commit-hash>
git push --force-with-lease origin feature/your-feature
```

**Warning:** Use `--force-with-lease` carefully, it rewrites history!

## Additional Resources

- [Excel VBA Documentation](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
- [VBA Style Guide](https://github.com/Rubberduck-VBA/Rubberduck)
- [Git Best Practices](https://git-scm.com/book/en/v2)
