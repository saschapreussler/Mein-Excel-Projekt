# Test-Szenario: Kassenbuch v6.1 — Vollständiger Systemtest

## Voraussetzungen
- Excel-Datei: `Programm Kassenbuch 2018_v2.7.3.xlsm`
- CSV-Dateien: Sparkasse-Export 2024, 2025 oder 2026
- VBA-Editor erreichbar (Alt+F11)

---

## Phase 0: Sync prüfen & Diagnose

### 0.1 – Neues Modul importieren (KRITISCH!)

> **Wenn dein Sync-Tool nur bestehende Module aktualisiert**, muss `mod_Vereinskasse_Filter.bas`
> und `mod_Diagnose.bas` MANUELL importiert werden:

1. Alt+F11 → VBA-Editor öffnen
2. **Datei → Datei importieren** (Strg+M)
3. Navigiere zu `vba\Modules\mod_Vereinskasse_Filter.bas` → Importieren
4. Dasselbe für `vba\Modules\mod_Diagnose.bas`
5. Prüfe ob beide Module im Projektbaum unter "Module" erscheinen

### 0.2 – Kompilierung prüfen

1. VBA-Editor → **Debuggen → Kompilieren von VBAProject**
2. Wenn ein Fehler erscheint → das fehlende Modul importieren
3. Erst wenn "Kompilieren" ohne Fehler durchläuft, funktioniert `Workbook_Open`

### 0.3 – Diagnose ausführen

1. Direktfenster öffnen: **Ansicht → Direktfenster** (Strg+G)
2. Eingeben: `mod_Diagnose.DiagnoseAlles` → Enter
3. Eine MsgBox zeigt alle Testergebnisse:
   - ✅ = OK
   - ❌ = Problem gefunden
4. **WICHTIG**: Wenn Test 03 meldet, dass Vereinskasse NICHT "Tabelle4" ist,
   muss der Event-Code verschoben werden (siehe Phase 5)

---

## Phase 1: Workbook_Open testen

### 1.1 – Simulation (ohne Datei-Neustart)

1. Direktfenster: `mod_Diagnose.SimuliereWorkbookOpen` → Enter
2. MsgBox zeigt für jeden Schritt ✅ oder ❌
3. Alle 9 Schritte sollten ✅ zeigen

### 1.2 – InputBox Abrechnungsjahr testen

1. Direktfenster: `mod_Diagnose.Test_InputBox_Abrechnungsjahr` → Enter
2. Bestätige mit "Ja" → C4 wird geleert
3. InputBox erscheint → gib z.B. `2025` ein
4. Bestätige → Wert wird in Einstellungen C4 geschrieben
5. **Erwartung**: Zelle C4 zeigt "2025"

### 1.3 – InputBox Kontostand testen

1. Direktfenster: `mod_Diagnose.Test_InputBox_Kontostand` → Enter
2. Bestätige mit "Ja"
3. InputBox erscheint → gib z.B. `1.234,56` ein
4. Bestätige → Wert wird in Einstellungen C5 geschrieben
5. **Erwartung**: Zelle C5 zeigt "1.234,56 €"

### 1.4 – InputBox Vereinsdaten testen

1. Direktfenster: `mod_Diagnose.Test_InputBox_Vereinsdaten` → Enter
2. Bestätige mit "Ja"
3. 4 InputBoxen erscheinen nacheinander:
   - Vereinsname → z.B. `KGA "Elisabeth Scholle" e.V.`
   - Straße → z.B. `Musterweg 12`
   - PLZ → z.B. `12345`
   - Ort → z.B. `Berlin`
4. **Erwartung**: Einstellungen C16, C17, C18, E18 befüllt

### 1.5 – Echter Neustart

1. Datei speichern & schließen
2. Datei erneut öffnen
3. **Erwartung**: Keine InputBoxen (Werte bereits vorhanden)
4. Startmenü zeigt das neue Design mit Kacheln

---

## Phase 2: Startseite (WOW-Effekt)

### 2.1 – Design prüfen

1. Zum Startmenü wechseln
2. **Erwartung**:
   - Dunkles Hero-Banner oben mit "★ K A S S E N B U C H ★"
   - Vereinsname in Türkis
   - Adresszeile mit Straße • PLZ Ort | Abrechnungsjahr
   - 4 KPI-Karten: Abrechnungsjahr, Mitglieder, Parzellen, Kontostand
   - 9 Navigations-Kacheln in 3 Spalten
   - Serienbrief-Bereich
   - Footer mit Version

### 2.2 – Manueller Neuaufbau

1. Direktfenster: `mod_Diagnose.Startseite_Neu_Aufbauen` → Enter
2. **Erwartung**: ✅-MsgBox, dann Startmenü wird angezeigt mit neuem Design

### 2.3 – KPI-Werte prüfen

| KPI | Quelle | Erwartung |
|-----|--------|-----------|
| Abrechnungsjahr | Einstellungen C4 | Zeigt die Jahreszahl (z.B. 2025) |
| Mitglieder | Mitgliederliste (aktive) | Zahl > 0 wenn Mitglieder existieren |
| Parzellen | Mitgliederliste (belegte) | Zahl > 0 wenn Parzellen vergeben |
| Kontostand | Einstellungen C5 | Zeigt den Betrag in Euro |

### 2.4 – Navigation testen

Klicke nacheinander auf jede Kachel:

| Kachel | Ziel-Blatt |
|--------|-----------|
| 📊 Zahlungsübersicht | Übersicht |
| 🏦 Bankkonto | Bankkonto |
| 💰 Vereinskasse | Vereinskasse |
| 📈 Dashboard | Dashboard / Übersicht |
| ⚡ Strom | Strom |
| 💧 Wasser | Wasser |
| ⚙ Einstellungen | Einstellungen |
| 🗃 Daten | Daten |
| 👥 Mitgliederverwaltung | Mitgliederliste (UserForm) |

Auf jedem Blatt sollte oben rechts ein **"🏠 Startmenü"**-Button erscheinen.

---

## Phase 3: Bankkonto E2-Formel

### 3.1 – Formel prüfen

1. Wechsle zu Bankkonto
2. Klicke auf Zelle E2
3. **Erwartung**: Formel enthält `Einstellungen!$C$5` und `Einstellungen!$C$4`
4. **NICHT** `Startmenü!$F$4`

### 3.2 – Formel manuell reparieren

1. Direktfenster: `mod_Diagnose.BankkontoE2_Reparieren` → Enter
2. MsgBox zeigt VORHER und NACHHER Formel
3. **Erwartung**: NACHHER enthält "Einstellungen"

### 3.3 – Monatsfilter testen (wenn CSV importiert)

1. Bankkonto → ComboBox oben wähle "März"
2. **Erwartung**: Nur März-Buchungen sichtbar
3. E2 zeigt Kontostand bis Ende Februar
4. Wähle "ganzes Jahr" → alle Buchungen sichtbar

---

## Phase 4: Vereinskasse Filter

### 4.1 – ComboBox erstellen

1. Direktfenster: `mod_Diagnose.VereinskasseComboBox_Erstellen` → Enter
2. **Erwartung**: ✅-MsgBox
3. Wechsle zu Vereinskasse → ComboBox in Zeile 24 sichtbar

### 4.2 – Filter testen (nur wenn Daten vorhanden)

1. ComboBox → "ganzes Jahr" auswählen → alle Daten sichtbar
2. ComboBox → "Januar" → nur Januar-Buchungen
3. C24 zeigt "Auszug: Januar 2025"

---

## Phase 5: Vereinskasse Event-Code prüfen

> Dieser Schritt ist nur nötig, wenn Test 03 gemeldet hat,
> dass der CodeName NICHT "Tabelle4" ist.

### 5.1 – Richtiges Sheet-Modul finden

1. Führe Diagnose aus: `mod_Diagnose.Test_03_TabellenCodenames`
2. Notiere den tatsächlichen CodeName (z.B. "Tabelle27")

### 5.2 – Event-Code verschieben

1. VBA-Editor → Tabelle4 öffnen → Code kopieren
2. Das richtige Sheet-Modul (z.B. Tabelle27) öffnen
3. Code einfügen
4. Tabelle4 leeren (Option Explicit bleibt)

---

## Phase 6: Einstellungen → Startseite Auto-Refresh

### 6.1 – Abrechnungsjahr ändern

1. Gehe zu Einstellungen → Zelle C4
2. Ändere z.B. von 2025 auf 2026
3. Wechsle zum Startmenü
4. **Erwartung**: Hero-Banner zeigt "Abrechnungsjahr 2026"
5. KPI-Karte zeigt "2026"

### 6.2 – Vereinsname ändern

1. Einstellungen → C16 → neuen Namen eingeben
2. Wechsle zum Startmenü
3. **Erwartung**: Vereinsname in Türkis aktualisiert

### 6.3 – Kontostand ändern

1. Einstellungen → C5 → neuen Betrag eingeben
2. Wechsle zum Startmenü
3. **Erwartung**: KPI "Kontostand Vorjahr" zeigt neuen Betrag

---

## Phase 7: CSV-Import + Komplett-Durchlauf

### 7.1 – Vorbereitung

1. Setze Abrechnungsjahr auf das Jahr der CSV-Datei (z.B. 2025)
2. Kontostand Vorjahr z.B. auf 5.000,00 setzen

### 7.2 – CSV importieren

1. Über den normalen Import-Weg die CSV laden
2. **Erwartung**: Buchungen erscheinen auf Bankkonto

### 7.3 – Nach Import prüfen

| Prüfpunkt | Erwartung |
|-----------|-----------|
| Bankkonto E2 | Zeigt Kontostand (Vorjahr + Buchungen) |
| Bankkonto Monatsfilter | Funktioniert mit importierten Daten |
| Startseite KPIs | Mitglieder/Parzellen-Zahl korrekt |
| Zahlungsübersicht | Kategorisierte Buchungen |

---

## Fehlersuche

| Symptom | Mögliche Ursache | Lösung |
|---------|------------------|--------|
| Gar nichts passiert | VBA-Projekt kompiliert nicht | Debuggen → Kompilieren, fehlendes Modul importieren |
| Startseite sieht alt aus | InitialisiereStartseite lief nicht | `mod_Diagnose.Startseite_Neu_Aufbauen` |
| E2 zeigt #BEZUG! | Alte Formel mit Startmenü!$F$4 | `mod_Diagnose.BankkontoE2_Reparieren` |
| Keine InputBox bei Start | Werte bereits in Einstellungen vorhanden | Normal! Werte löschen → erneut öffnen |
| ComboBox Vereinskasse fehlt | mod_Vereinskasse_Filter nicht importiert | Modul importieren (Strg+M) |
| ComboBox reagiert nicht | Event-Code im falschen Sheet-Modul | Siehe Phase 5 |
| Startseite aktualisiert nicht | Tabelle9.cls nicht synchronisiert | Tabelle9.cls prüfen (VerarbeiteKonfigAenderung) |
