# Mein-Excel-Projekt

## Überblick
Dieses Repository enthält ein professionell strukturiertes **Excel/VBA-Projekt (.xlsm)** zur Verwaltung eines Vereins. Der Schwerpunkt liegt auf:

- Mitgliederverwaltung (inkl. Pachtlogik)
- Import und Verarbeitung von Bankkonto-CSV-Dateien
- Automatischer Kategorisierung von Buchungen über eine regelbasierte Engine
- Nachvollziehbarer, wartbarer VBA-Architektur
- Versionierung und Zusammenarbeit mit **GitHub** und **GitHub Copilot**

Das Projekt ist bewusst so aufgebaut, dass **Copilot das gesamte fachliche und technische Modell verstehen kann**.

---

## Zielsetzung

- Zentrale Excel-Arbeitsmappe als "Single Source of Truth"
- Saubere Trennung von:
  - UI (Tabellen, Buttons, UserForms)
  - Logik (Engine, Validierung, Zuordnung)
  - Daten (Tabellenblätter)
- Wiederverwendbarer, testbarer VBA-Code
- Grundlage für spätere Erweiterungen (z. B. Word-Serienbriefe)

---

## Technologiestack

- Microsoft Excel (.xlsm)
- VBA (Module, Klassen, UserForms)
- GitHub (Versionsverwaltung)
- GitHub Copilot (Code-Assistenz)

Aktuell **kein PowerQuery / PowerAutomate**, aber die Architektur schließt eine spätere Erweiterung nicht aus.

---

## Tabellenblätter (High-Level)

| Blatt | Zweck |
|------|------|
| Startmenü | Zentrale Navigation (Buttons zu allen Funktionen) |
| Mitgliederliste | Stammdaten der Mitglieder inkl. Pacht- und Funktionslogik |
| Bankkonto | Importierte Bankbuchungen + Kategorie-Engine |
| Daten | Lookup-Tabellen, Kategorien, Konfiguration |

Details siehe: `docs/Tabellenstruktur.md`

---

## VBA-Architektur (Zielbild)

Der Code ist modular aufgebaut. Bestehender Code wird **schrittweise migriert**, nicht verworfen.

### Geplante Modulstruktur

| Modul | Verantwortung |
|------|--------------|
| modGlobals | Globale Konstanten, Enums, zentrale Einstellungen |
| modNavigation | Navigation zwischen Tabellen (Home, Startmenü) |
| modCSVImport | Import und Vorverarbeitung von Bank-CSV-Dateien |
| modKategorienEngine | Regel- & Scoring-Logik zur Kategoriezuordnung |
| modMitglieder | CRUD-Logik für Mitglieder |
| modValidierung | Plausibilitäts- und Datenprüfungen |
| modLogging | Debug- und Laufzeitprotokolle |
| modUtilities | Allgemeine Hilfsfunktionen |

### Klassen (Beispiele)

- `clsMitglied`
- `clsBuchung`
- `clsKategorie`
- `clsEngineResult`

---

## UserForms

- **Mitgliederverwaltung**
  - Anlegen / Bearbeiten / Löschen von Mitgliedern
  - Nutzung der Mitgliederliste als Backend
  - Keine direkte Bearbeitung kritischer Felder (z. B. GUID)

---

## GUID-Strategie (vorläufig)

- Jede Person erhält eine **eindeutige GUID**
- Speicherung in `Mitgliederliste!A`
- Spalte ist technisch notwendig, aber für Nutzer ausgeblendet
- Erzeugung ausschließlich über VBA

(Detaillierte Festlegung folgt in separatem Dokument)

---

## Repository-Struktur

```text
Mein-Excel-Projekt/
├─ excel/
│  └─ MeinProjekt.xlsm
├─ src/
│  ├─ modules/
│  ├─ classes/
│  └─ forms/
├─ docs/
│  ├─ projektuebersicht.md
│  └─ tabellenstruktur.md
├─ scripts/
│  └─ export_import.bas
└─ README.md
```

---

## Arbeitsweise mit Copilot (wichtig)

- Fachliche Anforderungen stehen **immer zuerst in `/docs`**
- VBA-Code wird **nicht blind generiert**, sondern:
  - erklärt
  - kommentiert
  - schrittweise integriert
- Jede neue Logik bekommt eine dokumentierte Verantwortung

---

## Status

- [x] Tabellenstruktur dokumentiert
- [x] Ordnerstruktur vorbereitet
- [ ] VBA-Code-Migration
- [ ] Kategorie-Engine Refactoring
- [ ] CSV-Import vereinheitlichen
- [ ] Logging & Debugging standardisieren

---

## Hinweise

Dieses Projekt wird **inkrementell verbessert**. Stabilität und Nachvollziehbarkeit haben Vorrang vor Geschwindigkeit.


---

## Tabellenblatt "Strom" (Zähler – Elektrizität)

- Zweck: Erfassung, Historisierung und Abrechnung von Stromverbräuchen je Parzelle und Mitglied.
- Grundprinzip: Zähler sind zeitlich Parzellen und Mitgliedern zugeordnet; es können Zählerwechsel stattfinden, ohne historische Daten zu verlieren.

Empfohlene fachliche Spalten (Ist-Zustand kann abweichen, wird später abgeglichen):
- Zähler-ID (eindeutig, perspektivisch GUID)
- Parzelle
- Mitglieds-GUID
- Zählernummer (physischer Zähler)
- Startstand
- Startdatum
- Endstand
- Enddatum
- Verbrauch (berechnet)
- Preis pro kWh (ggf. aus "Einstellungen")
- Kosten (berechnet)
- Zählerstatus (aktiv / gewechselt / stillgelegt)
- Kommentar / Historie

Besonderheiten:
- Zählerwechsel erzeugt einen neuen Datensatz, der alte bleibt unverändert erhalten.
- Verbrauchsberechnung erfolgt ausschließlich aus Differenzen (Endstand – Startstand).
- Übergabe der Kosten an Endabrechnung und Übersicht.

---

## Tabellenblatt "Wasser" (Zähler – Wasseruhren)

- Zweck: Analog zum Tabellenblatt "Strom", jedoch für Wasserverbräuche.
- Struktur und Logik sind bewusst identisch gehalten, um VBA-Code wiederzuverwenden.

Empfohlene fachliche Spalten:
- Zähler-ID (GUID)
- Parzelle
- Mitglieds-GUID
- Zählernummer
- Startstand
- Startdatum
- Endstand
- Enddatum
- Verbrauch (m³, berechnet)
- Preis pro Einheit (aus "Einstellungen")
- Kosten (berechnet)
- Zählerstatus
- Kommentar / Historie

---

## Tabellenblatt "Übersicht"

- Zweck: Zentrale Kontroll- und Auswertungssicht für Vorstand / Kasse.
- Keine Buchungen, ausschließlich Auswertung und Statusanzeige.

Inhaltliche Dimensionen:
- Mitglied
- Parzelle
- Abrechnungszeitraum

Kennzahlen:
- Soll-Beiträge (Mitglied, Pacht, Verbrauch)
- Ist-Zahlungen (aus "Bankkonto")
- Differenz

Statuslogik:
- fristgerecht bezahlt
- verspätet bezahlt
- offen
- Strafgebühr fällig

Visuelle Logik:
- Ampelfarben (grün / gelb / rot)
- optional Symbole oder Hinweise

Abhängigkeiten:
- Stammdaten aus "Mitgliederliste"
- Buchungen aus "Bankkonto"
- Verbrauchskosten aus "Strom" und "Wasser"
- Fristen und Gebühren aus "Einstellungen"

---

## Tabellenblatt "Einstellungen"

- Zweck: Regel- und Parameterzentrale ohne operative Logik.

Typische Inhalte:
- Zahlungsfristen (Mitgliedsbeitrag, Pacht, Verbrauch)
- Strafgebühren (fix / pro Zeitraum / Deckelung)
- Preisparameter (Strom, Wasser)
- Gültigkeitszeiträume von Regeln

Grundsatz:
- Keine Berechnungen mit Geschäftslogik
- Wird ausschließlich von VBA-Modulen gelesen

---

Hinweis:
Die Trennung von Stammdaten, Bewegungsdaten, Regeln und Auswertung ist bewusst gewählt, um Copilot eine klare mentale Modellbildung des Systems zu ermöglichen und spätere Erweiterungen (z. B. Serienbriefe, Word, PowerQuery) vorzubereiten.

