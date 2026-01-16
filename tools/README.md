# VBA Development Tools

This directory contains helper tools and utilities for VBA development and maintenance. These files are not imported into the runtime workbook.

## Export Helper: Modul1.bas

**Purpose:** Exports all VBA modules, forms, and class modules from the workbook to the `project_exports/` directory.

**Usage:**
1. Open the workbook in Excel
2. Press `Alt+F11` to open the VBA Editor
3. Import `Modul1.bas` into the workbook (File → Import File)
4. Run the macro `ExportAllVBA`
5. All modules will be exported to `project_exports/`
6. Remove `Modul1` from the workbook after export (right-click → Remove Modul1)

**Important:** This module should NOT be part of the production workbook. It's only used during development for version control purposes.

## Workflow

### Exporting Modules for Version Control

```vba
' After importing Modul1.bas temporarily:
Call ExportAllVBA
' Then remove Modul1 from the project
```

### Importing Modules into Workbook

1. Open the workbook
2. Open VBA Editor (`Alt+F11`)
3. For each file in `project_exports/`:
   - File → Import File
   - Select the `.bas`, `.cls`, or `.frm` file
4. Save the workbook

## Best Practices

- **Never commit the `.xlsm` file** - it's binary and causes merge conflicts
- **Always export modules** after making changes
- **Test imports** in a copy of the workbook first
- **Remove temporary/debug modules** before exporting for production

## Future Enhancements

Additional tools that could be added here:
- Static code analyzer for VBA syntax validation
- Duplicate function detector
- Sub/Function balance checker
- Code formatter
- Import automation script
