# Projektübersicht – Vereinsverwaltung

## Zweck dieses Dokuments
Dieses Dokument beschreibt die **fachliche Gesamtlogik** des Excel-VBA-Projekts und dient als zentrale Orientierung für:
- Entwickler
- zukünftige Wartung
- **GitHub Copilot** (explizit)

Excel selbst ist für Copilot nicht lesbar – diese Übersicht ersetzt sozusagen das fehlende „Modell“ der Arbeitsmappe.

---

## Benutzerführung (UI-Konzept)

### Startmenü (Tabelle1 – "Startmenü")
- Wird **automatisch beim Öffnen der Arbeitsmappe** angezeigt
- Zentrale Navigationsoberfläche für den Kassierer
- Enthält:
  - Überblicksinformationen (Anzahl Mitglieder, Parzellen etc.)
  - Navigations-Buttons zu allen relevanten Tabellenblättern
  - Button zum Öffnen der Mitgliederverwaltung (UserForm)

### Navigation
- Jedes fachliche Tabellenblatt enthält einen Button:
  - "Home" → Rückkehr zum Startmenü
- Navigation erfolgt ausschließlich über Buttons (kein manuelles Wechseln vorgesehen)

### UserForms
- `frm_Mitgliederverwaltung`
  - Zentrale grafische Oberfläche zur:
    - Anlage neuer Mitglieder
    - Bearbeitung bestehender Mitglieder
    - Abwicklung von Austritten / Kündigungen

Weitere UserForms unterstützen Spezialfälle (z. B. Zählerwechsel).

---

## Fachliche Hauptprozesse

### 1. Bankdaten-Import & Kategorisierung

**Ablauf (vereinfacht):**

1. CSV-Datei der Bank wird importiert
2. Rohdaten werden normalisiert
3. Jede Buchung wird einzeln verarbeitet
4. Kategorie-Engine ordnet der Buchung genau **eine Kategorie** zu
5. Ergebnis wird im Blatt "Bankkonto" dargestellt
6. Betrag wird automatisch in die zur Kategorie gehörenden Spalten geschrieben

---

### 2. Kategorie-Engine (Konzept)

Die Kategorie-Zuordnung basiert auf einer **datengetriebenen Logik**.

#### Zentrale Datenquelle
- Tabelle auf Blatt `Daten`
- Enthält pro Kategorie u. a.:
  - Einnahme / Ausgabe
  - Priorität
  - Schlüsselwörter (Verwendungszweck, Name, IBAN etc.)
  - Zielspalten im Bankkonto
  - Fälligkeiten

#### Logische Kriterien (kombiniert)
- Textanalyse des Verwendungszwecks
- Einnahme vs. Ausgabe
- Logische Plausibilität des Betrags
- Mitgliedsbezug (falls vorhanden)
- Priorisierung konkurrierender Kategorien

> Die Kategorie-Engine ist bewusst modular aufgebaut und besteht aus mehreren Modulen (Normalize, Evaluator, Apply, Pipeline).

---

### 3. Barkasse (Vereinskasse)

- Manuelle Pflege durch den Kassierer
- Automatisierte Übernahme relevanter Buchungen vom Bankkonto
- Verknüpfung über eindeutige Referenzen (BK / KA-Nummern)
- Sicherstellung der Nachvollziehbarkeit Bank ↔ Barkasse

---

### 4. Verbrauchserfassung (Strom / Wasser)

- Erfassung aktueller Zählerstände
- Berechnung des Verbrauchs
- Kostenumlage auf Mitglieder / Parzellen
- Unterstützung von Zählerwechseln
- Historisierung aller relevanten Daten

---

### 5. Endabrechnung

- Zusammenführung aller Daten:
  - Mitgliedsdaten
  - Pacht
  - Bankbuchungen
  - Barkasse
  - Strom- und Wasserkosten
- Ziel:
  - vollständige Jahresabrechnung pro Mitglied
- Perspektivisch:
  - Übergabe an Word über Serienbrieffunktion

---

## Technische Leitlinien

- Excel-Datei = UI & Laufzeit
- VBA-Code = Logik & Regeln
- Tabellen = Datenhaltung
- Keine Geschäftslogik direkt in Worksheets
- Jede Automatik ist über Buttons oder klar definierte Entry-Points erreichbar

---

## Offene Punkte / Weiterentwicklung

- Konsolidierung der Kategorie-Engine
- Einführung klarer Zugriffsfunktionen für Tabellen
- Refactoring der bestehenden Worksheet-Klassen
- Einführung automatischer VBA Import-/Export-Routinen
