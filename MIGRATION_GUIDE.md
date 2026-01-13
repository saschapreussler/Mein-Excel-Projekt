# EntityKey Migration Guide

## Übersicht

Diese Anleitung beschreibt die Migration von numerischen EntityKeys zu string-basierten MemberID/BANK-ID Keys im Kassenbuch-Projekt.

## ⚠️ WICHTIG: Backup erstellen!

**Vor der Migration MUSS ein Backup erstellt werden!**

1. Schließen Sie die Excel-Datei
2. Erstellen Sie eine Kopie der Datei mit aktuellem Datum (z.B. `Kassenbuch_BACKUP_2026-01-13.xlsm`)
3. Öffnen Sie die Backup-Kopie für die Migration

## Migrations-Schritte

### Schritt 1: VBA-Code aktualisieren

Die folgenden Module wurden hinzugefügt/aktualisiert:

**Neue Module:**
- `mod_MigrateKeys.bas` - Hauptmigrations-Modul
- `mod_MigrateReport.bas` - Berichtserstellung und Heuristik
- `mod_ReassignKeys.bas` - Nachträgliche Zuordnung für neue Mitglieder

**Aktualisierte Module:**
- `mod_Banking_Data.bas` - String-basierte BANK-IDs statt numerischer IDs
- `mod_Mitglieder_UI.bas` - Robuste MemberID-Suche, sichere Spalten-Operationen
- `mod_Mapping_Tools.bas` - NormalizeString ist jetzt öffentlich

### Schritt 2: Code kompilieren

1. Öffnen Sie den VBA-Editor (Alt+F11)
2. Menü: **Debug** → **Compile VBAProject**
3. Beheben Sie eventuelle Fehler (sollten keine auftreten)

### Schritt 3: Migration ausführen

**Im VBA-Editor:**

```vba
' 1. Migration starten
Sub RunMigration()
    Call mod_MigrateKeys.Migrate_EntityKeys_To_MemberID
End Sub
```

**Was passiert:**
- Alle EntityKeys im Blatt "Daten" werden von numerisch auf String umgestellt
- Einträge mit Mitgliedszuordnung erhalten die MemberID aus der Mitgliederliste
- Nicht zugeordnete Einträge erhalten eine BANK-ID im Format: `BANK-yyyymmddhhmmss-nnn`
- EntityKeys im Blatt "Bankkonto" werden entsprechend aktualisiert

### Schritt 4: Migrationsbericht erstellen

```vba
' 2. Report generieren
Sub CreateReport()
    Call mod_MigrateReport.GenerateMigrationReport
End Sub
```

Der Report zeigt:
- Anzahl erfolgreich zugeordneter Mitglieder
- Anzahl nicht zugeordneter Einträge (BANK-IDs)
- Details zu allen BANK-IDs

### Schritt 5: Migration validieren

```vba
' 3. Validation durchführen
Sub ValidateMigration()
    Call mod_MigrateKeys.Validate_MigrationResults
End Sub
```

**Erwartetes Ergebnis:**
- Alle EntityKeys sollten Strings sein (keine numerischen Werte mehr)
- MemberIDs: GUIDs im Format `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
- BANK-IDs: Format `BANK-yyyymmddhhmmss-nnn`

### Schritt 6: Heuristisches Matching (Optional)

Falls viele BANK-IDs übrig sind, kann versucht werden, diese anhand der Parzellennummer zuzuordnen:

```vba
' 4. Heuristisches Matching (max. 2 Ziffern)
Sub TryHeuristicMatch()
    Call mod_MigrateReport.TryMatchByTrailingDigits(2)
End Sub
```

**Hinweis:** Dies versucht, BANK-IDs anhand der letzten Ziffern im Kontonamen (z.B. "Parzelle 15") mit Mitgliedern zu verbinden.

## Nach der Migration

### Neue Funktionalität

**1. Reassignment für nachträglich angelegte Mitglieder:**

Wenn ein Mitglied erst nach dem Import von Banktransaktionen angelegt wird:

```vba
' Beispiel: Reassign für neues Mitglied
Sub ReassignForNewMember()
    Dim memberID As String
    memberID = "12345678-1234-1234-1234-123456789012" ' MemberID aus Mitgliederliste
    
    Call mod_ReassignKeys.ReassignBankKeysForNewMember(memberID)
End Sub
```

**2. Neue Bank-Einträge:**

Bei zukünftigen CSV-Importen:
- Neue, nicht zuordenbare IBANs erhalten automatisch BANK-IDs
- Das Fuzzy-Matching versucht weiterhin, Namen zuzuordnen
- Manuelle Zuordnung bleibt möglich

### Validierung der Funktionalität

**UserForms testen:**
1. Öffnen Sie "Mitgliederverwaltung"
2. Erstellen Sie ein neues Mitglied
3. Bearbeiten Sie ein bestehendes Mitglied
4. ✓ Keine Laufzeitfehler sollten auftreten

**Banking-Import testen:**
1. Importieren Sie einen CSV-Kontoauszug
2. Prüfen Sie die Mapping-Tabelle im Blatt "Daten"
3. ✓ Neue Einträge sollten BANK-IDs erhalten

## Fehlerbehebung

### "Fehler beim Migrieren"

**Ursache:** Blattschutz oder fehlende Berechtigungen

**Lösung:**
1. Entsperren Sie alle Blätter manuell (Überprüfen → Blattschutz aufheben)
2. Führen Sie die Migration erneut aus

### "MemberID nicht gefunden"

**Ursache:** Mitglied hat keine MemberID in Spalte A

**Lösung:**
1. Führen Sie `Fuelle_MemberIDs_Wenn_Fehlend` aus
2. Wiederholen Sie die Migration

### "Spalten-Hidden Fehler"

**Ursache:** Arbeitsmappe ist strukturgeschützt

**Lösung:**
- Die neue Funktion `SafeSetColumnHidden` sollte dies automatisch behandeln
- Falls Fehler auftreten: Strukturschutz aufheben (Überprüfen → Arbeitsmappe schützen)

## Rückgängig machen

Falls die Migration nicht wie erwartet funktioniert:

1. Schließen Sie die Datei OHNE zu speichern
2. Öffnen Sie die Backup-Datei
3. Wiederholen Sie die Schritte mit angepassten Parametern

## Unterstützung

Bei Fragen oder Problemen:
- Prüfen Sie den Debug-Output (Strg+G im VBA-Editor)
- Erstellen Sie einen Screenshot der Fehlermeldung
- Kontaktieren Sie den Support mit dem Backup und Screenshots

## Änderungshistorie

- **2026-01-13**: Initiale Migration erstellt
  - mod_MigrateKeys.bas hinzugefügt
  - mod_MigrateReport.bas hinzugefügt
  - mod_ReassignKeys.bas hinzugefügt
  - mod_Banking_Data.bas aktualisiert (String-basierte IDs)
  - mod_Mitglieder_UI.bas aktualisiert (robuste Suche)
