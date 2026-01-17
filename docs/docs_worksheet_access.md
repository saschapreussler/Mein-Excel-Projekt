# Zentrale Worksheet-Zugriffsschicht

## Ziel

Dieses Dokument definiert eine **verbindliche Zugriffsschicht** für alle Excel-Tabellenblätter.

Ziel ist es:
- harte Abhängigkeiten (`Worksheets("…")`, Indexzugriffe) zu vermeiden
- Refactoring und Erweiterungen zu erleichtern
- GitHub Copilot eine klare API für Tabellenzugriffe zu geben

Ab sofort gilt:
> **Direkte Worksheet-Zugriffe außerhalb dieses Moduls sind nicht mehr erlaubt.**

---

## Prinzip

- Jedes fachlich relevante Tabellenblatt erhält **eine Zugriffsfunktion**
- Rückgabewert ist immer ein `Worksheet`
- Der Name des Blattes ist **zentral an einer Stelle definiert**

---

## Empfohlenes Modul

**Modulname:** `mod_WorksheetAccess`

Dieses Modul enthält ausschließlich Funktionen vom Typ `Function … As Worksheet`.

---

## Zugriffsfunktionen (verbindlich)

### Stammdaten

- `WsMitgliederliste()` → Tabellenblatt "Mitgliederliste"
- `WsMitgliederhistorie()` → Tabellenblatt "Mitgliederhistorie"

### Buchhaltung

- `WsBankkonto()` → Tabellenblatt "Bankkonto"
- `WsVereinskasse()` → Tabellenblatt "Vereinskasse"

### Verbrauch / Zähler

- `WsStrom()` → Tabellenblatt "Strom"
- `WsWasser()` → Tabellenblatt "Wasser"

### Steuerung / Regeln

- `WsDaten()` → Tabellenblatt "Daten"
- `WsEinstellungen()` → Tabellenblatt "Einstellungen"

### UI / Navigation

- `WsStartmenue()` → Tabellenblatt "Startmenü"
- `WsUebersicht()` → Tabellenblatt "Übersicht"

---

## Beispielhafte Implementierung (VBA)

```vba
Option Explicit

Public Function WsMitgliederliste() As Worksheet
    Set WsMitgliederliste = ThisWorkbook.Worksheets("Mitgliederliste")
End Function

Public Function WsBankkonto() As Worksheet
    Set WsBankkonto = ThisWorkbook.Worksheets("Bankkonto")
End Function

Public Function WsStrom() As Worksheet
    Set WsStrom = ThisWorkbook.Worksheets("Strom")
End Function

Public Function WsWasser() As Worksheet
    Set WsWasser = ThisWorkbook.Worksheets("Wasser")
End Function
```

---

## Vorteile

- Tabellenblattnamen müssen nur einmal angepasst werden
- Fehler durch Umbenennungen werden sofort sichtbar
- Copilot kann gezielt angewiesen werden:
  > „Verwende ausschließlich WsBankkonto() für Zugriffe“

---

## Migrationsstrategie

1. Modul `mod_WorksheetAccess` anlegen
2. Funktionen für alle relevanten Blätter erstellen
3. Bestehenden Code **schrittweise** umstellen:
   - von `Worksheets("Bankkonto")`
   - zu `WsBankkonto()`
4. Funktionale Logik dabei **nicht ändern**

---

## Verbotene Muster (ab jetzt)

- `Worksheets(3)`
- `ActiveWorkbook.Worksheets("…")`
- Direktzugriffe aus UserForms

---

*Diese Zugriffsschicht ist verpflichtend für alle zukünftigen Änderungen.*

