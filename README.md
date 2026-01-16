# Mein Excel Projekt - VBA Kassenbuch 2018

Ein Excel-VBA basiertes Kassenbuch-System für Vereinsverwaltung mit Mitgliederverwaltung, Parzellenverwaltung, Banking-Import und Kategorisierung.

## Überblick

Dieses Projekt verwaltet:
- **Mitgliederdaten** mit GUID-basierten IDs
- **Parzellenzuordnung** und Pächter-Historie
- **Bankkonto-Transaktionen** mit CSV-Import
- **Automatische Kategorisierung** von Transaktionen
- **Zählerverwaltung** (Wasser, Strom)
- **Formular-basierte UI** für Dateneingabe

## Projektstruktur

```
Mein-Excel-Projekt/
├── project_exports/         # VBA Quellcode (versioniert)
│   ├── mod_*.bas           # Standard-Module (Geschäftslogik)
│   ├── frm_*.frm           # UserForms (Benutzeroberfläche)
│   ├── Tabelle*.bas        # Worksheet-Module
│   └── DieseArbeitsmappe.bas  # Workbook-Module
├── tools/                   # Entwicklungstools
│   ├── Modul1.bas          # Export-Helper
│   ├── check_vba_balance.py # Statischer Code-Checker
│   └── README.md
├── tests/                   # Testdaten und Test-Makros
│   ├── sample.csv          # Beispiel-Transaktionen
│   └── README.md
├── .gitignore              # Git Ignore-Patterns
├── CONTRIBUTING.md         # Entwickler-Workflow
└── README.md               # Diese Datei
```

## Schnellstart

### Voraussetzungen

- Microsoft Excel 2016 oder neuer
- VBA-Makros müssen aktiviert sein
- Python 3.x für statische Code-Checks (optional)

### Workbook erstellen/laden

1. Repository klonen:
   ```bash
   git clone https://github.com/saschapreussler/Mein-Excel-Projekt.git
   cd Mein-Excel-Projekt
   ```

2. Neue Excel-Arbeitsmappe erstellen (`.xlsm` Format)

3. VBA-Module importieren:
   - Excel öffnen
   - `Alt+F11` für VBA-Editor
   - Für jede Datei in `project_exports/`:
     - Datei → Datei importieren
     - Datei auswählen und importieren

4. Makros aktivieren und testen:
   - `Debug → VBAProject kompilieren` (sollte ohne Fehler kompilieren)

### CSV Import testen

1. In Excel zum Blatt "Bankkonto" wechseln
2. Makro ausführen oder Button klicken
3. `tests/sample.csv` auswählen
4. 5 Transaktionen sollten importiert werden

## Hauptmodule

### Geschäftslogik

- **mod_Const.bas** - Zentrale Konstanten (Spaltennummern, Blattnamen, Status)
- **mod_Mitglieder_UI.bas** - Mitgliederverwaltung, Sortierung, Dropdowns
- **mod_Banking_Data.bas** - CSV-Import, IBAN-Mapping, Transaktionsverwaltung
- **mod_ZaehlerLogik.bas** - Zählerstandsverwaltung, Übersichten
- **mod_KategorieEngine_*.bas** - Regelbasierte Kategorisierung von Transaktionen

### Benutzeroberfläche

- **frm_Mitgliederverwaltung** - Hauptformular für Mitgliederverwaltung
- **frm_Mitgliedsdaten** - Detail-Formular für einzelne Mitglieder
- **frm_Zaehlerwechsel** - Formular für Zählerstandswechsel

### Tests

- **mod_Tests.bas** - Test-Prozeduren für Qualitätssicherung
  - `Test_CSVImport` - CSV-Import testen
  - `Test_MemberIDGeneration` - GUID-Generierung testen
  - `Test_MemberLookup` - Mitglieder-Suche testen (auch mit versteckten Spalten)
  - `Run_All_Tests` - Alle automatisierten Tests ausführen

## Wichtige Funktionen

### Mitglieder-ID Lookup

```vba
' Findet Mitglied anhand der Member ID
Dim row As Long
row = FindeRowByMemberID("guid-here")

' Oder mit Worksheet-Parameter (Alias)
row = FindMemberRowByID(ws, "guid-here")
```

### CSV Import

```vba
' Importiert Banktransaktionen aus CSV
Call Importiere_Kontoauszug
' Unterstützt: UTF-8, Semikolon-Delimiter, Duplikatserkennung
```

### Automatische Kategorisierung

```vba
' Wendet Kategorisierungsregeln an
Call mod_KategorieEngine_Pipeline.ApplyCategorizationPipeline()
```

## Entwicklung

### Module exportieren

```vba
' Modul1.bas temporär importieren, dann:
Call ExportAllVBA
' Exportiert alle Module nach project_exports/
```

### Statische Checks

```bash
# Prüft Sub/Function Balance
python3 tools/check_vba_balance.py
```

### Tests ausführen

```vba
' Im VBA Immediate Window (Strg+G):
Call Run_All_Tests

' Oder einzeln:
Call Test_CSVImport
Call Test_MemberLookup
```

Siehe [CONTRIBUTING.md](CONTRIBUTING.md) für detaillierte Entwicklungs-Workflows.

## Code-Standards

- **Option Explicit** in allen Modulen
- PascalCase für Public Procedures
- Ungarische Notation für Variablen (optional)
- Fehlerbehandlung in Public Procedures
- Kommentare für komplexe Logik

Beispiel:

```vba
Option Explicit

Public Sub ProcessMemberData()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    lastRow = ws.Cells(ws.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
    
    ' Verarbeitung hier...
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler: " & Err.Description, vbCritical
End Sub
```

## Häufige Probleme

### Kompilierungsfehler "Methode oder Datenelement nicht gefunden"

- Module in der richtigen Reihenfolge importieren
- Zuerst `mod_Const.bas`, dann andere Module
- Alle Referenzen überprüfen

### CSV-Import funktioniert nicht

- Dateiformat: Semikolon-separiert, UTF-8
- Spalten müssen Header in Zeile 1 haben
- Siehe `tests/sample.csv` als Beispiel

### Formular zeigt keine Daten

- Blattnamen in `mod_Const.bas` überprüfen
- Spalten-Konstanten validieren
- `Debug → VBAProject kompilieren` ausführen

## Testing

### Manuelle Tests

1. **Kompilierung:** `Debug → VBAProject kompilieren` (Strg+F5)
2. **CSV Import:** Mit `tests/sample.csv` testen
3. **Formulare:** Alle Formulare öffnen und testen
4. **Sortierung:** Mitgliederliste sortieren
5. **Dropdowns:** Validierung testen

### Automatisierte Tests

```vba
Call Run_All_Tests
```

Siehe [tests/README.md](tests/README.md) für Details.

## Mitwirken

Beiträge sind willkommen! Bitte:

1. Issue erstellen oder zuweisen lassen
2. Feature-Branch erstellen: `git checkout -b feature/mein-feature`
3. Änderungen machen und testen
4. Module exportieren (mit `tools/Modul1.bas`)
5. Pull Request erstellen

Siehe [CONTRIBUTING.md](CONTRIBUTING.md) für detaillierte Anleitung.

## Lizenz

[Lizenzinformationen hier einfügen]

## Kontakt

[Kontaktinformationen hier einfügen]

## Changelog

### v2.6.0 (2024-01-16)
- ✅ Projekt-Cleanup: Temp-Module in `tools/` verschoben
- ✅ `.gitignore` hinzugefügt, `.xlsm` Binaries entfernt
- ✅ Test-Framework hinzugefügt (`mod_Tests.bas`)
- ✅ Member-Lookup-Funktion öffentlich gemacht
- ✅ Dokumentation hinzugefügt (README, CONTRIBUTING, tools/README, tests/README)
- ✅ Statischer Code-Checker hinzugefügt (`tools/check_vba_balance.py`)
- ✅ Beispiel-CSV für Tests hinzugefügt
