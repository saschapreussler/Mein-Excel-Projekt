# Testing Guide - EntityKey Migration

## Test-Checkliste für Reviewer

### Vorbereitung

- [ ] Backup der Original-Datei erstellt
- [ ] Kopie für Tests geöffnet
- [ ] VBA-Editor geöffnet (Alt+F11)
- [ ] Code kompiliert ohne Fehler (Debug → Compile VBAProject)

### Test 1: Migration ausführen

**Ziel:** Numerische EntityKeys werden zu String-basierten MemberID/BANK-IDs

```vba
Sub Test1_Migration()
    Call mod_MigrateKeys.Migrate_EntityKeys_To_MemberID
End Sub
```

**Erwartetes Ergebnis:**
- ✓ MsgBox zeigt "Migration abgeschlossen!"
- ✓ Anzahl migrierter Einträge wird angezeigt
- ✓ Anzahl neuer BANK-IDs wird angezeigt

**Validierung:**
1. Öffnen Sie das Blatt "Daten"
2. Prüfen Sie Spalte S (EntityKey):
   - [ ] Alle Werte sind jetzt Strings (keine Zahlen)
   - [ ] MemberIDs haben GUID-Format: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
   - [ ] BANK-IDs haben Format: `BANK-20260113150225-001`

### Test 2: Validation Report

```vba
Sub Test2_Validation()
    Call mod_MigrateKeys.Validate_MigrationResults
End Sub
```

**Erwartetes Ergebnis:**
- ✓ MsgBox zeigt Statistik:
  - String EntityKeys: [Anzahl]
  - MemberIDs: [Anzahl]
  - BANK-IDs: [Anzahl]
  - Numerische EntityKeys: 0
  - Leere Einträge: 0 oder gering

### Test 3: Migration Report

```vba
Sub Test3_Report()
    Call mod_MigrateReport.GenerateMigrationReport
End Sub
```

**Erwartetes Ergebnis:**
- ✓ MsgBox zeigt Zusammenfassung
- ✓ Details im Direktfenster (Strg+G)
- ✓ Liste aller nicht zugeordneten BANK-IDs mit Details

### Test 4: Heuristisches Matching (Optional)

**Nur wenn viele BANK-IDs übrig sind:**

```vba
Sub Test4_HeuristicMatch()
    Call mod_MigrateReport.TryMatchByTrailingDigits(2)
End Sub
```

**Erwartetes Ergebnis:**
- ✓ MsgBox zeigt Anzahl gefundener Matches
- ✓ BANK-IDs mit erkannter Parzellennummer werden zu MemberIDs

**Nachprüfung:**
- Blatt "Daten", Spalte Y (Debug): "Auto-matched by trailing digits"

### Test 5: UserForm - Neues Mitglied

**Ziel:** UserForm funktioniert ohne Laufzeitfehler

1. Öffnen Sie das Blatt "Mitgliederliste"
2. Doppelklick auf eine leere Zeile (oder Button "Neues Mitglied")
3. Formular "Mitgliedsdaten" öffnet sich
4. Füllen Sie alle Pflichtfelder aus:
   - [ ] Nachname: "Testmann"
   - [ ] Vorname: "Test"
   - [ ] Parzelle: "99"
   - [ ] Funktion: "PÄCHTER"
5. Klicken Sie "Übernehmen"

**Erwartetes Ergebnis:**
- ✓ Keine Fehlermeldung
- ✓ Mitglied wird in Liste eingefügt
- ✓ MemberID wird automatisch generiert (Spalte A)
- ✓ Liste wird neu sortiert

### Test 6: UserForm - Mitglied bearbeiten

1. Doppelklick auf ein bestehendes Mitglied
2. Ändern Sie ein Feld (z.B. Telefonnummer)
3. Klicken Sie "Übernehmen"

**Erwartetes Ergebnis:**
- ✓ Keine Fehlermeldung
- ✓ Änderung wird gespeichert
- ✓ MemberID bleibt unverändert

### Test 7: Banking-Import

**Ziel:** CSV-Import funktioniert mit neuen BANK-IDs

1. Menü/Button zum CSV-Import aufrufen
2. Wählen Sie eine Test-CSV-Datei
3. Import durchführen

**Erwartetes Ergebnis:**
- ✓ Import erfolgreich
- ✓ Neue Einträge im Blatt "Bankkonto"
- ✓ Mapping-Tabelle in "Daten" wird aktualisiert
- ✓ Neue IBANs erhalten BANK-IDs (Format: BANK-...)

### Test 8: Reassignment für neues Mitglied

**Szenario:** Ein Mitglied wird nachträglich angelegt, es existieren aber bereits Banktransaktionen

```vba
Sub Test8_Reassign()
    Dim memberID As String
    
    ' MemberID des gerade erstellten Test-Mitglieds aus Schritt 5
    ' (aus Spalte A der Mitgliederliste kopieren)
    memberID = "[HIER MEMBERID EINFÜGEN]"
    
    Call mod_ReassignKeys.ReassignBankKeysForNewMember(memberID)
End Sub
```

**Erwartetes Ergebnis:**
- ✓ MsgBox zeigt Anzahl reassigned EntityKeys
- ✓ BANK-IDs, die zum Mitglied passen, werden zu dessen MemberID

### Test 9: Geschützte Blätter

**Ziel:** SafeSetColumnHidden funktioniert mit geschützten Blättern

1. Schützen Sie das Blatt "Mitgliederliste" (Überprüfen → Blatt schützen)
2. Führen Sie eine Operation aus, die Spalten ein-/ausblendet
   (z.B. Sortierung, Formatierung)

**Erwartetes Ergebnis:**
- ✓ Keine Fehlermeldung "Geschütztes Blatt"
- ✓ Operation wird korrekt ausgeführt
- ✓ Blattschutz bleibt bestehen

### Test 10: Robuste MemberID-Suche

**Ziel:** FindeRowByMemberID arbeitet korrekt mit Variants

```vba
Sub Test10_SearchRobustness()
    Dim wsM As Worksheet
    Dim testRow As Long
    Dim memberID As String
    
    Set wsM = ThisWorkbook.Worksheets("Mitgliederliste")
    
    ' Erste MemberID aus der Liste holen
    memberID = wsM.Cells(6, 1).Value ' M_START_ROW, M_COL_MEMBER_ID
    
    ' Test: Suche sollte Zeile 6 zurückgeben
    testRow = mod_Mitglieder_UI.FindeRowByMemberID(memberID)
    
    If testRow = 6 Then
        MsgBox "✓ Suche erfolgreich! Zeile gefunden: " & testRow
    Else
        MsgBox "✗ Fehler! Erwartete Zeile 6, erhalten: " & testRow
    End If
End Sub
```

**Hinweis:** Da FindeRowByMemberID privat ist, muss dieser Test über eine öffentliche Funktion erfolgen, die die Suche aufruft.

## Fehlerprotokoll

Falls Fehler auftreten, dokumentieren Sie bitte:

| Test | Fehler | Screenshot | Lösung |
|------|--------|------------|--------|
| Test X | Beschreibung | Link | Wie behoben |

## Erfolgs-Kriterien

Alle Tests müssen bestanden werden:

- [ ] Test 1: Migration ✓
- [ ] Test 2: Validation ✓
- [ ] Test 3: Report ✓
- [ ] Test 4: Heuristic (optional) ✓
- [ ] Test 5: Neues Mitglied ✓
- [ ] Test 6: Mitglied bearbeiten ✓
- [ ] Test 7: Banking-Import ✓
- [ ] Test 8: Reassignment ✓
- [ ] Test 9: Geschützte Blätter ✓
- [ ] Test 10: Robuste Suche ✓

## Nach erfolgreichen Tests

1. Speichern Sie die Test-Datei
2. Führen Sie eine finale Kompilierung durch
3. Erstellen Sie einen Screenshot der Validation-Summary
4. Dokumentieren Sie eventuelle Besonderheiten
5. Genehmigen Sie den PR

## Rollback-Plan

Falls kritische Fehler auftreten:

1. Schließen Sie die Datei ohne zu speichern
2. Dokumentieren Sie den Fehler
3. Erstellen Sie ein GitHub Issue mit:
   - Beschreibung des Fehlers
   - Screenshot
   - Schritte zur Reproduktion
4. Öffnen Sie das Backup für weitere Tests
