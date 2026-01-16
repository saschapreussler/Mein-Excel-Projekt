# Contributing to Mein-Excel-Projekt

Thank you for contributing to this Excel VBA project! This document provides guidelines for working with the VBA codebase.

## Prerequisites

- Microsoft Excel 2016 or later (with VBA support)
- Basic understanding of VBA programming
- Git for version control

## Project Structure

```
Mein-Excel-Projekt/
├── project_exports/       # Exported VBA modules (.bas, .cls, .frm files)
├── tools/                 # Development utilities (not part of runtime)
├── tests/                 # Test files and sample data
├── .gitignore            # Git ignore rules
└── CONTRIBUTING.md       # This file
```

## Workflow

### 1. Export VBA Modules

To export VBA code from the workbook to version control:

1. Open the workbook: `Programm Kassenbuch 2018_v2.6.0.xlsm`
2. Enable macros if prompted
3. Press `Alt+F11` to open the VBA Editor
4. In the VBA Editor, open the Immediate Window (`Ctrl+G`)
5. Import the export helper from `tools/Modul1.bas`
6. Run: `Call ExportAllVBA`
7. All modules will be exported to `project_exports/`

**Note:** The workbook binary (*.xlsm) should NOT be committed to version control.

### 2. Import VBA Modules

To import VBA code from version control into a workbook:

1. Open the target workbook
2. Press `Alt+F11` to open the VBA Editor
3. For each module in `project_exports/`:
   - **Standard Modules (.bas)**: File → Import File → Select the .bas file
   - **Class Modules (.cls)**: File → Import File → Select the .cls file
   - **User Forms (.frm)**: File → Import File → Select the .frm file (the .frx will be imported automatically)
4. **Document Modules** (Tabelle*.bas, DieseArbeitsmappe.bas): 
   - These cannot be imported directly
   - Copy/paste the code manually into the existing document modules

**Important:** Do NOT import modules from the `tools/` directory into production workbooks.

### 3. Run Static Parser Locally

Before committing changes, verify that all Sub/Function/Property declarations are properly closed:

```bash
python3 /path/to/vba_static_checker.py
```

This script checks that:
- Every `Sub` has a matching `End Sub`
- Every `Function` has a matching `End Function`
- Every `Property` has a matching `End Property`

### 4. Compile VBA Project

After making changes, always compile the VBA project to check for syntax errors:

1. Open the VBA Editor (`Alt+F11`)
2. Go to **Debug → Compile VBAProject**
3. Fix any compilation errors that appear
4. Test critical functionality:
   - `Call Initialize_ImportReport_ListBox`
   - `Call Test_CSVImport`

### 5. Testing

Run the test macros to validate functionality:

- **CSV Import Test**: `Call Test_CSVImport`
  - Uses sample data from `tests/sample.csv`
  - Verifies import functionality

Add new test macros to `mod_Tests.bas` for new features.

## Creating Pull Requests

1. **Create a new branch** from `main`:
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make your changes** in the VBA Editor

3. **Export modules** using `ExportAllVBA`

4. **Run static checker** to verify code structure

5. **Compile and test** your changes

6. **Commit your changes**:
   ```bash
   git add project_exports/
   git commit -m "Description of your changes"
   ```

7. **Push your branch**:
   ```bash
   git push origin feature/your-feature-name
   ```

8. **Create a Pull Request** on GitHub:
   - Provide a clear title and description
   - List changed files
   - Describe how to test the changes
   - Include any compilation output or test results

## Code Style Guidelines

- Use **Option Explicit** in all modules
- Use meaningful variable names
- Comment complex logic
- Follow existing naming conventions:
  - Module names: `mod_FeatureName`
  - Public constants: `UPPER_SNAKE_CASE`
  - Variables: `camelCase` or `PascalCase`
- Keep functions and subs focused and single-purpose
- Avoid deeply nested logic

## Common Issues

### Hidden Columns
Use `Find` method instead of iterating through cells to avoid issues with hidden columns:
```vba
Set foundCell = ws.Columns("A").Find(What:=value, LookIn:=xlValues)
```

### UTF-8 and BOM
CSV imports handle UTF-8 with BOM automatically. The import routine supports multiple delimiters (comma, semicolon, tab, pipe).

### Module Parsing Errors
If you see parsing errors, check:
- All Sub/Function/Property statements have matching End statements
- No orphaned End statements
- Proper line continuations (use `_`)

## Getting Help

If you encounter issues:
1. Check existing issues on GitHub
2. Review the code comments in the modules
3. Run the static checker for structural problems
4. Create a new issue with a detailed description

## License

This project is maintained by the KGA team. Please respect the project's license and usage terms.
