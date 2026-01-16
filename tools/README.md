# Tools Directory

This directory contains utility macros and helper modules for development and debugging.
These modules are **not** part of the runtime application.

## Contents

### Export/Import Helpers
- **Modul1.bas** - VBA module export helper (`ExportAllVBA`)
  - Exports all VBA components to `project_exports/` directory
  - Usage: Run `ExportAllVBA` from within Excel VBA Editor

## Purpose

These tools support:
- VBA code export and version control
- Debugging and diagnostics
- Development-time utilities

**Note:** These modules should not be imported into production workbooks.
