# Vereinsverwaltung – Excel / VBA

## Überblick

Dieses Projekt ist eine umfangreiche **Excel-basierte Vereinsverwaltungs- und Abrechnungslösung** für einen Kleingarten- bzw. Parzellenverein.  
Primäre Zielgruppe ist der **Kassierer / Kassenwart**, perspektivisch auch Vorstand und weitere Funktionsträger.

Die Arbeitsmappe dient als **zentrale Daten- und Logikbasis** für:

- Mitglieder- und Parzellenverwaltung  
- Bankkonto- und Barkassenbuchführung  
- Regelbasierte Kategorisierung von Buchungen  
- Strom- und Wasserverbrauchserfassung inkl. Zählerwechsel  
- Fristen-, Status- und Strafgebührenlogik  
- Übersichten, Auswertungen und Endabrechnungen je Mitglied / Parzelle  

Die Lösung ist vollständig in **Microsoft Excel (Windows)** mit **VBA** umgesetzt und als **Single-User-System** konzipiert.

---

## Technischer Rahmen

- Plattform: Microsoft Excel 2019 oder neuer (Windows)
- Dateityp: `.xlsm`
- Programmiersprache: VBA
- Nutzung: Single-User
- Versionsverwaltung: GitHub  
  - VBA-Code wird als Module, Klassen und Forms exportiert
  - Excel-Datei selbst ist nicht Bestandteil der Versionshistorie

Perspektivisch denkbar (nicht Bestandteil des aktuellen Designs):
- Nutzung der Excel-Daten als Datenquelle für **Word-Serienbriefe** (z. B. Endabrechnungen)
- Einsatz von PowerQuery / PowerAutomate
- Weitere Automatisierung der Abrechnungsläufe

---

## Fachliche Zielsetzung

Das System soll in der Lage sein:

- CSV-Bankkontoauszüge zu importieren
- Buchungen automatisiert zu analysieren und zu kategorisieren
- Einnahmen und Ausgaben korrekt auf Zielspalten zu verteilen
- Mitglieder, Parzellen und externe Entitäten logisch zuzuordnen
- Strom- und Wasserverbräuche inklusive Zählerwechsel zu berechnen
- Zahlungsfristen zu überwachen
- Strafgebühren bei verspäteten Zahlungen zu berücksichtigen
- Eine konsistente Datenbasis für Jahres- und Endabrechnungen bereitzustellen

---

## Zentrale Tabellenblätter (Excel)

### Startmenü
- Einstiegspunkt der Arbeitsmappe
- Navigation per Buttons zu allen Funktionsbereichen
- Zentrale Bedienoberfläche für den Nutzer

---

### Mitgliederliste
- Zentrale Stammdatentabelle
- GUID-basierte Member-ID (technischer Schlüssel, ausgeblendet)
- Parzellen-, Seiten- und Funktionszuordnung
- Grundlage für:
  - Bankbuchungen
  - Verbrauchszähler
  - Abrechnungen
  - Übersichten

---

### Bankkonto
- Import von CSV-Kontoauszügen
- Speicherung aller relevanten Buchungsdaten
- Automatische Kategorisierung über die Kategorie-Engine
- Aufteilung der Beträge in Einnahmen- und Ausgabenspalten
- Grundlage für:
  - Beitragsstatus
  - Zahlungsüberwachung
  - Kassen- und Abrechnungslogik

---

### Vereinskasse
- Manuelle Barkassenführung
- Verknüpfung mit Bankkonto über interne BK / KA-Nummern
- Abbildung von Auszahlungen an die Kasse und Rückführungen

---

### Strom
- Erfassung von Stromzählerständen je Parzelle / Mitglied
- Unterstützung von Zählerwechseln
- Verbrauchsberechnung
- Kostenberechnung auf Basis der **im Blatt gepflegten Preisparameter**
- Übergabe der Ergebnisse an Übersicht und Abrechnung

---

### Wasser
- Analog zum Strom-Blatt aufgebaut
- Separate Zähler- und Preislogik
- Unabhängige Pflege der Parameter
- Fachlich bewusst getrennt von Strom

---

### Übersicht
- Zentrales Kontroll- und Auswertungsblatt
- Anzeige:
  - gezahlte Beiträge
  - offene Beträge
  - Fristgerechte / verspätete Zahlungen
- Visualisierung von Zahlungsstatus je Mitglied / Parzelle
- Grundlage für Strafgebührenlogik und Nachverfolgung

---

### Einstellungen
- Zentrale Parametertabelle
- Definition von:
  - Zahlungsfristen
  - Strafgebühren
  - allgemeinen Abrechnungsregeln
- Wird von Logik- und Übersichtsblättern ausgewertet

---

### Daten
- Zentrale Regel- und Zuordnungstabelle
- Beinhaltet u. a.:
  - Kategorien
  - Einnahme / Ausgabe-Kennzeichen
  - Keywords für Verwendungszwecke
  - Prioritäten (Scoring)
  - Zielspalten im Blatt „Bankkonto“
  - Entity-Zuordnungen (Mitglied, Bank, Versorger etc.)
- Technisches Herzstück der Kategorie-Engine

---

## Kategorie-Engine (fachliches Kernmodul)

Die Kategorie-Engine ordnet jede importierte Bankbuchung automatisch einer Kategorie zu.

Dabei werden u. a. berücksichtigt:
- Einnahme / Ausgabe
- Schlüsselwörter im Verwendungszweck und Buchungstext
- Prioritäten (Scoring-Logik)
- Plausibilität von Beträgen
- Zuordnung zu Mitgliedern oder externen Entitäten

Das Ergebnis wird:
- in der Kategorie-Spalte visualisiert
- farblich gekennzeichnet
- betragsmäßig korrekt auf Zielspalten verteilt

---

## Strom- und Wasser-Zählerlogik

- Separate Tabellenblätter für Strom und Wasser
- Unterstützung von:
  - laufenden Zählerständen
  - Zählerwechseln
  - Mehrjahresbetrachtungen
- Preisparameter werden **direkt im jeweiligen Fachblatt gepflegt**
- Ergebnisse fließen in:
  - Übersicht
  - Endabrechnungen
  - Beitragsbewertung

---

## Bedienkonzept

Das Projekt verfolgt einen **dualen Bedienansatz**:

1. **Direkte Bearbeitung über Tabellenblätter**
   - Für erfahrene Nutzer
   - Maximale Transparenz der Daten

2. **Menügeführte Bearbeitung über UserForms**
   - Für strukturierte Eingabe
   - Validierung, Führung und Komfort

Beide Wege:
- greifen auf dieselben Tabellen zu
- schreiben ausschließlich in definierte Zellbereiche
- sind fachlich gleichwertig

---

## VBA-Architektur (Überblick)

Die VBA-Logik ist modular aufgebaut:

- `frm_*` → UserForms (UI / Benutzerinteraktion)
- `mod_*` → Fach- und Logikmodule
- Klassenmodule → Objekt- und Ereignislogik (Workbook / Worksheets)

Wichtige Modulgruppen:
- Kategorie-Engine
- CSV-Import / Banking
- Mitglieder- und UI-Logik
- Strom- und Wasser-Zählerlogik
- Übersichts- und Statusberechnung
- Hilfs- und Mappingfunktionen

Bestehender VBA-Code ist vorhanden und wird:
- strukturiert nach GitHub exportiert
- schrittweise migriert, refaktoriert und erweitert

---

## Designprinzipien

- Klare Trennung von:
  - Stammdaten
  - Bewegungsdaten
  - Regeln / Parameter
  - Auswertung
- Keine Geschäftslogik in Excel-Formeln
- VBA als führende Logikschicht
- Erweiterbarkeit vor Optimierung
- Nachvollziehbarkeit vor Komplexität

---

## Zielgruppe & Wartung

Das Projekt richtet sich an technisch versierte Vereinsmitglieder mit soliden Excel-Kenntnissen.

Die Struktur ist bewusst so ausgelegt, dass:
- GitHub Copilot das Projektkonzept vollständig erfassen kann
- Weiterentwicklung kontrolliert möglich bleibt
- spätere Erweiterungen (z. B. Serienbriefe) integrierbar sind

---

*Dieses Projekt befindet sich in aktiver Weiterentwicklung.*
