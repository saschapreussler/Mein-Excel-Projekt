# Pull Request: Migration zu string-basierten EntityKeys (MemberID / BANK-IDs)

## Übersicht

Dieser PR implementiert die Migration von numerischen EntityKeys zu string-basierten MemberID/BANK-ID Keys und behebt kritische Runtime-Fehler im Zusammenhang mit geschützten Blättern und der Mitgliedersuche.

## Problemstellung

Die ursprünglichen Probleme waren:

1. **Runtime-Fehler beim Spalten-Verstecken**: Direkter Zugriff auf `Columns().Hidden` verursacht Fehler bei geschützten Blättern
2. **Fragile MemberID-Suche**: `.Find`-basierte Suche ist nicht Variant-robust und kann bei geschützten Blättern fehlschlagen
3. **Numerische EntityKeys**: Verwendung von `Application.Max()` auf gemischte String/Numerik-Werte führt zu Fehlern
4. **Fehlende Migration**: Keine Strategie für Übergang von numerischen zu string-basierten Keys

## Implementierte Lösung

### 1. Neue Migration-Module

**mod_MigrateKeys.bas**
- Hauptmigrations-Modul für EntityKey-Konvertierung
- `Migrate_EntityKeys_To_MemberID()` - Migriert alle EntityKeys
- `Validate_MigrationResults()` - Validiert Migrationsergebnisse
- `FindMemberIDByName()` - Findet MemberID anhand von Name
- `Generate_BankID()` - Erzeugt BANK-IDs im Format BANK-yyyymmddhhmmss-nnn

**mod_MigrateReport.bas**
- `GenerateMigrationReport()` - Detaillierter Bericht über Migration
- `TryMatchByTrailingDigits()` - Heuristisches Matching anhand Parzellennummer
- `CleanupDebugMessages()` - Debug-Nachrichten bereinigen

**mod_ReassignKeys.bas**
- `ReassignBankKeysForNewMember(memberID)` - Nachträgliche Zuordnung von BANK-IDs zu neuen Mitgliedern

### 2. Kernänderungen

**mod_Banking_Data.bas:**
- ❌ Entfernt: `Application.Max(wsD.Columns(DATA_MAP_COL_ENTITYKEY))` (numerische Annahme)
- ✅ Hinzugefügt: `Generate_BankID_String()` für BANK-yyyymmddhhmmss-nnn Format
- ✓ Neue Einträge erhalten string-basierte BANK-IDs statt numerischer IDs

**mod_Mitglieder_UI.bas:**
- ✓ `FindeRowByMemberID` komplett überarbeitet: zeilenweise, variant-robust
- ✓ `VarToSafeString` und `CleanMemberID` Helper hinzugefügt
- ✓ `SafeSetColumnHidden` für geschützte Blätter hinzugefügt

**mod_Mapping_Tools.bas:**
- ✓ `NormalizeString` ist jetzt Public für Wiederverwendung

### 3. String-basierte EntityKeys

**Für Mitglieder:**
- Format: MemberID (GUID) z.B. `12345678-1234-1234-1234-123456789012`
- Quelle: Spalte A der Mitgliederliste

**Für Bank-Einträge ohne Mitglied:**
- Format: `BANK-yyyymmddhhmmss-nnn`
- Beispiel: `BANK-20260113150225-001`
- Eindeutig durch Timestamp + Counter

### 4. Robuste MemberID-Suche

**Vorher (fragil):**
```vba
Set rngFind = rngSearch.Find(What:=MemberID, LookIn:=xlValues, LookAt:=xlWhole)
```

**Nachher (robust):**
```vba
For r = M_START_ROW To lastRow
    cellValue = wsM.Cells(r, M_COL_MEMBER_ID).Value
    cleanCellValue = CleanMemberID(VarToSafeString(cellValue))
    If StrComp(cleanCellValue, cleanSearchID, vbTextCompare) = 0 Then
        FindeRowByMemberID = r
        Exit For
    End If
Next r
```

### 5. Sichere Spalten-Operationen

**Neue Funktion:**
```vba
Public Function SafeSetColumnHidden(ByRef ws As Worksheet, ByVal colIndex As Long, ByVal isHidden As Boolean) As Boolean
    ' Prüft ProtectStructure und ProtectContents
    ' Unprotect/Protect mit Fehlerbehandlung
End Function
```

## Gelieferte Dateien

**Neue Module:**
- `Module/mod_MigrateKeys.bas` (9KB) - Hauptmigrations-Modul
- `Module/mod_MigrateReport.bas` (7KB) - Reporting und Heuristik
- `Module/mod_ReassignKeys.bas` (9KB) - Nachträgliche Zuordnung

**Aktualisierte Module:**
- `Module/mod_Banking_Data.bas` - String-basierte BANK-IDs
- `Module/mod_Mitglieder_UI.bas` - Robuste Suche, sichere Spaltenoperationen
- `Module/mod_Mapping_Tools.bas` - Öffentliche NormalizeString-Funktion

**Dokumentation:**
- `MIGRATION_GUIDE.md` - Schritt-für-Schritt Migrationsanleitung
- `TESTING_GUIDE.md` - Umfassende Test-Checkliste mit 10 Tests
- `PR_DESCRIPTION.md` - Diese Datei

## Testing-Anleitung

⚠️ **WICHTIG: Backup erstellen vor dem Testen!**

Bitte folgen Sie der Datei `TESTING_GUIDE.md` für die vollständige Test-Checkliste.

**Schnell-Test:**
1. Backup erstellen
2. Code kompilieren (Debug → Compile VBAProject)
3. Migration ausführen: `Call mod_MigrateKeys.Migrate_EntityKeys_To_MemberID`
4. Validation: `Call mod_MigrateKeys.Validate_MigrationResults`
5. UserForms testen (Neues Mitglied, Bearbeiten)

## Wichtige Hinweise

**Minimal-invasive Änderungen:**
- ✓ Keine UI-Umgestaltungen
- ✓ Datenmigration erfolgt NUR durch expliziten Aufruf
- ✓ Bestehende Funktionalität bleibt unverändert
- ✓ Rückgängig machen: Backup wiederherstellen

**Nicht-invasive Natur:**
- Migration ist optional und muss manuell ausgeführt werden
- Keine automatischen Änderungen beim Workbook_Open
- Backup-Strategie dokumentiert

**Rückwärtskompatibilität:**
- Neue Funktionen sind zusätzlich, nicht ersetzend
- Bestehende Prozeduren bleiben funktional
- Migration ist reversibel durch Backup

## Bekannte Einschränkungen

1. **FindNewEntityKeyByOld** in mod_MigrateKeys ist ein Platzhalter - Eine vollständige Implementierung würde ein Mapping speichern
2. **Heuristisches Matching** kann Fehlzuordnungen bei mehrdeutigen Parzellennummern verursachen
3. **Migration** kann bei sehr großen Datensätzen (>10000 Zeilen) mehrere Minuten dauern

## Review-Checkliste

- [ ] Code kompiliert ohne Fehler
- [ ] TESTING_GUIDE.md befolgt und alle Tests bestanden
- [ ] Screenshots der Validation-Summary erstellt
- [ ] Migration-Report generiert und geprüft
- [ ] UserForms funktionieren ohne Runtime-Fehler
- [ ] Banking-Import funktioniert mit neuen BANK-IDs
- [ ] Dokumentation ist vollständig und verständlich

## Nach erfolgreicher Review

**Nicht automatisch mergen!** Dies ist ein Review-PR.

Nach erfolgreichen lokalen Tests:
1. Genehmigen Sie den PR
2. Erstellen Sie eine Anleitung für die Produktiv-Migration
3. Planen Sie ein Backup-Fenster für die Produktiv-Umgebung
4. Führen Sie die Migration in der Produktiv-Datei durch

## Kontextbilder

Die im Problem-Statement referenzierten Screenshots zeigen:
- Fehler beim Hidden-Property setzen (behoben durch SafeSetColumnHidden)
- Migration abgeschlossen Dialog
- Validation Summary
- Ursprüngliche UserForm-Dialoge (unverändert)

## Fragen oder Probleme?

Bei Fragen oder Problemen während des Testens:
1. Prüfen Sie den Debug-Output (Strg+G im VBA-Editor)
2. Konsultieren Sie die Troubleshooting-Sektion in MIGRATION_GUIDE.md
3. Erstellen Sie ein Issue mit Screenshot und Beschreibung

## Änderungshistorie

**2026-01-13 - Initiale Implementation**
- Migration-Module erstellt
- Banking_Data und Mitglieder_UI aktualisiert
- Umfassende Dokumentation hinzugefügt
- Test-Checkliste erstellt
