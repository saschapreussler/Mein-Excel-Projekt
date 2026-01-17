# VBA-Architektur – Vereinsverwaltung

## Ziel dieses Dokuments

Dieses Dokument beschreibt die **logische und technische Architektur** der VBA-Lösung.
Es dient als verbindliche Referenz für:
- Weiterentwicklung
- Refactoring
- Einsatz von GitHub Copilot
- saubere Trennung von Verantwortlichkeiten

Grundsatz: **Nicht jedes Modul darf alles.**

---

## Architekturprinzipien

1. **Single Source of Truth**
   - Fachliche Daten liegen ausschließlich in Excel-Tabellenblättern
   - VBA liest und schreibt definierte Bereiche

2. **Trennung von Verantwortlichkeiten**
   - UI (UserForms) enthalten keine Geschäftslogik
   - Module enthalten Fachlogik
   - Klassen kapseln Objekt- und Ereignislogik

3. **Dualer Bedienansatz**
   - Tabellenblatt-Eingaben und UserForms sind gleichwertig
   - Beide Wege nutzen dieselben Fachfunktionen

4. **Keine Logik in Worksheet_Change ohne Delegation**
   - Events dürfen nur triggern, nicht rechnen

---

## Schichtenmodell

### 1. UI-Schicht (Benutzerinteraktion)

**UserForms**
- `frm_Mitgliederverwaltung`
- `frm_Mitgliedsdaten`
- `frm_Zaehlerwechsel`

Verantwortung:
- Eingabevalidierung (formal)
- Benutzerführung
- Aufruf fachlicher Funktionen

Dürfen:
- Daten anzeigen
- Fachfunktionen aufrufen

Dürfen nicht:
- Berechnungen durchführen
- direkt komplexe Tabellenlogik implementieren

---

### 2. Orchestrierungs- / Service-Schicht

**Module (steuernd)**

- `mod_Mitglieder_UI`
- `mod_Banking_Data`
- `mod_ZaehlerLogik`

Verantwortung:
- Koordination von Abläufen
- Übergabe von Daten an Fachmodule
- Fehlerbehandlung

---

### 3. Fachlogik-Schicht

**Kategorie-Engine**

- `mod_KategorieEngine_Pipeline`
- `mod_KategorieEngine_Normalize`
- `mod_KategorieEngine_Evaluator`
- `mod_KategorieEngine_Apply`
- `mod_KategorieEngine_Utils`
- `mod_KategorieRegeln`
- `mod_KategorieZiel`

Verantwortung:
- Analyse von Bankbuchungen
- Scoring / Prioritäten
- Entscheidung über Kategorie
- Ermittlung der Zielspalte im Blatt „Bankkonto“

---

**Zähler- und Verbrauchslogik**

- `mod_ZaehlerLogik`

Verantwortung:
- Verarbeitung von Strom- und Wasserzählern
- Zählerwechsel
- Verbrauchsberechnung
- Kostenberechnung

Gemeinsamkeiten Strom / Wasser:
- identische Logik
- unterschiedliche Parameter und Blätter

---

### 4. Infrastruktur- / Hilfsschicht

- `mod_Hilfsfunktionen`
- `mod_Mapping_Tools`
- `mod_Const`

Verantwortung:
- Wiederverwendbare Funktionen
- Konstanten
- Mapping und Suchlogik

Keine Fachentscheidungen.

---

## Klassenmodule

### Workbook- und Worksheet-Klassen

- `DieseArbeitsmappe.cls`
- `Tabelle*.cls`

Verantwortung:
- Ereignisse (Open, Activate, Change)
- Weiterleitung an Module

Regel:
- Keine Fachlogik im Klassenmodul
- Nur Delegation

---

## Tabellenblatt-Zugriffe (verbindlich)

Direkte Zugriffe wie `Worksheets("Bankkonto")` sollen **nicht verstreut** auftreten.

Empfohlenes Muster:
- Zentrale Zugriffsfunktionen (werden später eingeführt)
- Einheitliche Benennung

Beispiele:
- `GetWsBankkonto()`
- `GetWsMitgliederliste()`
- `GetWsStrom()`
- `GetWsWasser()`

---

## Datenfluss (vereinfacht)

CSV-Import → Bankkonto → Kategorie-Engine → Zuordnung → Übersicht

Zählerstände → Verbrauch → Kosten → Übersicht → Endabrechnung

Mitgliederliste → Zuordnung → Bank / Zähler / Abrechnung

---

## Regeln für zukünftige Erweiterungen

- Neue Logik immer zuerst fachlich dokumentieren
- Kein direkter Tabellenzugriff aus UserForms
- Neue Module klar einer Schicht zuordnen
- Copilot nur mit Bezug auf dieses Dokument nutzen

---

*Dieses Dokument ist verbindlich für die weitere Entwicklung.*

