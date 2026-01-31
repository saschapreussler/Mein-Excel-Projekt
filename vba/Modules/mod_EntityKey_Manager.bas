Attribute VB_Name = "mod_EntityKey_Manager"
Option Explicit

' ***************************************************************
' MODUL: mod_EntityKey_Manager
' ZWECK: Verwaltung und Zuordnung von EntityKeys für Bankverkehr
' VERSION: 1.1 - 31.01.2026
' WICHTIG: Bestehende Daten werden NIEMALS überschrieben!
' ***************************************************************

' ===============================================================
' KONSTANTEN FÜR ENTITYKEY-TABELLE (Spalten S-Y auf Daten-Blatt)
' ===============================================================
Private Const EK_COL_ENTITYKEY As Long = 19      ' S - EntityKey (GUID)
Private Const EK_COL_IBAN As Long = 20           ' T - IBAN
Private Const EK_COL_KONTONAME As Long = 21      ' U - Zahler/Empfänger (Bank)
Private Const EK_COL_ZUORDNUNG As Long = 22      ' V - Mitglied(er)/Zuordnung
Private Const EK_COL_PARZELLE As Long = 23       ' W - Parzelle(n)
Private Const EK_COL_ROLE As Long = 24           ' X - EntityRole
Private Const EK_COL_DEBUG As Long = 25          ' Y - Debug Zuordnung

Private Const EK_START_ROW As Long = 4           ' Daten beginnen ab Zeile 4
Private Const EK_HEADER_ROW As Long = 3          ' Überschriften in Zeile 3

Private Const EK_ROLE_DROPDOWN_COL As Long = 32  ' AF - Dropdown-Quelle für EntityRole

' EntityRole-Präfixe
Private Const PREFIX_SHARE As String = "SHARE-"
Private Const PREFIX_VERSORGER As String = "VERS-"
Private Const PREFIX_BANK As String = "BANK-"
Private Const PREFIX_SHOP As String = "SHOP-"
Private Const PREFIX_EHEMALIG As String = "EX-"

' EntityRole-Werte
Private Const ROLE_MITGLIED_MIT_PACHT As String = "MITGLIED_MIT_PACHT"
Private Const ROLE_MITGLIED_OHNE_PACHT As String = "MITGLIED_OHNE_PACHT"
Private Const ROLE_EHEMALIGES_MITGLIED As String = "EHEMALIGES_MITGLIED"
Private Const ROLE_VERSORGER As String = "VERSORGER"
Private Const ROLE_BANK As String = "BANK"
Private Const ROLE_SHOP As String = "SHOP"

' ===============================================================
' ÖFFENTLICHE PROZEDUR: Importiert IBANs aus Bankkonto und
' erstellt EntityKey-Einträge
' ===============================================================
Public Sub ImportiereIBANsAusBankkonto()
    
    Dim wsBK As Worksheet
    Dim wsD As Worksheet
    Dim dictIBANs As Object
    Dim r As Long
    Dim lastRowBK As Long
    Dim lastRowD As Long
    Dim nextRowD As Long
    Dim currentIBAN As String
    Dim currentKontoName As String
    Dim ibanKey As Variant
    Dim anzahlNeu As Long
    Dim anzahlBereitsVorhanden As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set dictIBANs = CreateObject("Scripting.Dictionary")
    
    ' Schutz entfernen
    On Error Resume Next
    wsD.Unprotect PASSWORD:=PASSWORD
    wsBK.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    anzahlNeu = 0
    anzahlBereitsVorhanden = 0
    
    ' ============================================================
    ' SCHRITT 1: Bereits vorhandene IBANs in Daten-Tabelle merken
    ' ============================================================
    lastRowD = wsD.Cells(wsD.Rows.Count, EK_COL_IBAN).End(xlUp).Row
    
    If lastRowD >= EK_START_ROW Then
        For r = EK_START_ROW To lastRowD
            currentIBAN = NormalisiereIBAN(wsD.Cells(r, EK_COL_IBAN).value)
            If currentIBAN <> "" Then
                If Not dictIBANs.Exists(currentIBAN) Then
                    dictIBANs.Add currentIBAN, True
                    anzahlBereitsVorhanden = anzahlBereitsVorhanden + 1
                End If
            End If
        Next r
    End If
    
    Debug.Print "Bereits vorhandene IBANs: " & anzahlBereitsVorhanden
    
    ' ============================================================
    ' SCHRITT 2: IBANs aus Bankkonto sammeln (eindeutig)
    ' ============================================================
    Dim dictNeueIBANs As Object
    Set dictNeueIBANs = CreateObject("Scripting.Dictionary")
    
    lastRowBK = wsBK.Cells(wsBK.Rows.Count, BK_COL_IBAN).End(xlUp).Row
    
    Debug.Print "Bankkonto: Zeilen " & BK_START_ROW & " bis " & lastRowBK
    
    For r = BK_START_ROW To lastRowBK
        currentIBAN = NormalisiereIBAN(wsBK.Cells(r, BK_COL_IBAN).value)
        currentKontoName = Trim(wsBK.Cells(r, BK_COL_NAME).value)
        
        ' Überspringe leere oder ungültige IBANs
        If currentIBAN <> "" And currentIBAN <> "N.A." And Len(currentIBAN) >= 15 Then
            ' Ist diese IBAN bereits in der Daten-Tabelle?
            If Not dictIBANs.Exists(currentIBAN) Then
                ' Neue IBAN gefunden!
                If Not dictNeueIBANs.Exists(currentIBAN) Then
                    ' Speichere IBAN mit Kontoname(n)
                    dictNeueIBANs.Add currentIBAN, currentKontoName
                Else
                    ' IBAN bereits gesammelt - Kontoname ergänzen wenn anders
                    If InStr(dictNeueIBANs(currentIBAN), currentKontoName) = 0 Then
                        dictNeueIBANs(currentIBAN) = dictNeueIBANs(currentIBAN) & vbLf & currentKontoName
                    End If
                End If
            End If
        End If
    Next r
    
    Debug.Print "Neue eindeutige IBANs gefunden: " & dictNeueIBANs.Count
    
    ' ============================================================
    ' SCHRITT 3: Neue IBANs in Daten-Tabelle eintragen
    ' ============================================================
    If lastRowD < EK_START_ROW Then
        nextRowD = EK_START_ROW
    Else
        nextRowD = lastRowD + 1
    End If
    
    For Each ibanKey In dictNeueIBANs.Keys
        ' Nur IBAN und Kontoname eintragen - Rest macht die Automatik
        wsD.Cells(nextRowD, EK_COL_IBAN).value = ibanKey
        wsD.Cells(nextRowD, EK_COL_KONTONAME).value = dictNeueIBANs(ibanKey)
        
        anzahlNeu = anzahlNeu + 1
        nextRowD = nextRowD + 1
    Next ibanKey
    
    ' Schutz wieder aktivieren
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' ============================================================
    ' SCHRITT 4: Jetzt EntityKey-Zuordnung durchführen
    ' ============================================================
    If anzahlNeu > 0 Then
        Dim antwort As VbMsgBoxResult
        antwort = MsgBox("Import abgeschlossen!" & vbCrLf & vbCrLf & _
                        "Neue IBANs importiert: " & anzahlNeu & vbCrLf & _
                        "Bereits vorhanden (übersprungen): " & anzahlBereitsVorhanden & vbCrLf & vbCrLf & _
                        "Möchten Sie jetzt die automatische Mitglieder-Zuordnung starten?", _
                        vbYesNo + vbQuestion, "IBAN-Import erfolgreich")
        
        If antwort = vbYes Then
            Call AktualisiereAlleEntityKeys
        End If
    Else
        MsgBox "Keine neuen IBANs gefunden!" & vbCrLf & vbCrLf & _
               "Alle " & anzahlBereitsVorhanden & " IBANs aus dem Bankkonto sind bereits in der EntityKey-Tabelle.", _
               vbInformation, "Import abgeschlossen"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler beim IBAN-Import: " & Err.Description, vbCritical
End Sub

' ===============================================================
' HILFSFUNKTION: Normalisiert IBAN (entfernt Leerzeichen, Großbuchstaben)
' ===============================================================
Private Function NormalisiereIBAN(ByVal iban As Variant) As String
    Dim result As String
    
    If IsNull(iban) Or IsEmpty(iban) Then
        NormalisiereIBAN = ""
        Exit Function
    End If
    
    result = UCase(Trim(CStr(iban)))
    result = Replace(result, " ", "")
    result = Replace(result, "-", "")
    
    NormalisiereIBAN = result
End Function

' ===============================================================
' HAUPTPROZEDUR: Aktualisiert alle EntityKeys in der Tabelle
' WICHTIG: Bestehende Daten werden NIEMALS überschrieben!
' ===============================================================
Public Sub AktualisiereAlleEntityKeys()
    
    Dim wsD As Worksheet
    Dim wsM As Worksheet
    Dim wsH As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim iban As String
    Dim kontoName As String
    Dim currentEntityKey As String
    Dim currentZuordnung As String
    Dim currentParzelle As String
    Dim currentRole As String
    Dim currentDebug As String
    Dim newEntityKey As String
    Dim zuordnung As String
    Dim parzellen As String
    Dim entityRole As String
    Dim debugInfo As String
    Dim ampelStatus As Long  ' 1=Grün, 2=Gelb, 3=Rot
    Dim mitgliederGefunden As Collection
    Dim zeilenMitEingriff As Collection
    Dim zeilenNeu As Long
    Dim zeilenUnveraendert As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    Set zeilenMitEingriff = New Collection
    
    zeilenNeu = 0
    zeilenUnveraendert = 0
    
    ' Schutz entfernen
    On Error Resume Next
    wsD.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    lastRow = wsD.Cells(wsD.Rows.Count, EK_COL_IBAN).End(xlUp).Row
    If lastRow < EK_START_ROW Then lastRow = EK_START_ROW
    
    ' Dropdown für EntityRole einrichten (einmalig für alle Zeilen)
    Call SetupEntityRoleDropdown(wsD, lastRow)
    
    ' Jede Zeile durchgehen
    For r = EK_START_ROW To lastRow
        iban = Trim(wsD.Cells(r, EK_COL_IBAN).value)
        kontoName = Trim(wsD.Cells(r, EK_COL_KONTONAME).value)
        
        ' Bestehende Werte lesen
        currentEntityKey = Trim(wsD.Cells(r, EK_COL_ENTITYKEY).value)
        currentZuordnung = Trim(wsD.Cells(r, EK_COL_ZUORDNUNG).value)
        currentParzelle = Trim(wsD.Cells(r, EK_COL_PARZELLE).value)
        currentRole = Trim(wsD.Cells(r, EK_COL_ROLE).value)
        currentDebug = Trim(wsD.Cells(r, EK_COL_DEBUG).value)
        
        ' Überspringe leere Zeilen
        If iban = "" And kontoName = "" Then GoTo NextRow
        
        ' ============================================================
        ' WICHTIG: Bestehende Daten NIEMALS überschreiben!
        ' ============================================================
        If HatBereitsGueltigeDaten(currentEntityKey, currentZuordnung, currentRole) Then
            ' Zeile hat bereits gültige Daten - NICHT ÄNDERN
            zeilenUnveraendert = zeilenUnveraendert + 1
            
            ' Nur Ampelfarbe aktualisieren basierend auf vorhandenen Daten
            If currentRole <> "" Then
                Call SetzeAmpelFarbe(wsD, r, 1)  ' Grün - bereits zugeordnet
            End If
            GoTo NextRow
        End If
        
        ' ============================================================
        ' Neue Zeile ohne vollständige Daten - Automatik durchführen
        ' ============================================================
        zeilenNeu = zeilenNeu + 1
        
        ' Suche Mitglieder zum Kontonamen (aktive UND ehemalige!)
        Set mitgliederGefunden = SucheMitgliederZuKontoname(kontoName, wsM, wsH)
        
        ' Generiere EntityKey und Zuordnung basierend auf Suchergebnis
        Call GeneriereEntityKeyUndZuordnung(mitgliederGefunden, kontoName, _
                                             newEntityKey, zuordnung, parzellen, entityRole, debugInfo, ampelStatus)
        
        ' Schreibe nur LEERE Felder (niemals überschreiben!)
        If currentEntityKey = "" Then wsD.Cells(r, EK_COL_ENTITYKEY).value = newEntityKey
        If currentZuordnung = "" Then wsD.Cells(r, EK_COL_ZUORDNUNG).value = zuordnung
        If currentParzelle = "" Then wsD.Cells(r, EK_COL_PARZELLE).value = parzellen
        If currentRole = "" Then wsD.Cells(r, EK_COL_ROLE).value = entityRole
        If currentDebug = "" Then wsD.Cells(r, EK_COL_DEBUG).value = debugInfo
        
        ' Ampelfarbe setzen
        Call SetzeAmpelFarbe(wsD, r, ampelStatus)
        
        ' Zeilen mit Eingriff merken
        If ampelStatus = 3 Then
            zeilenMitEingriff.Add r
        End If
        
NextRow:
    Next r
    
    ' Formatierung anwenden
    Call FormatiereEntityKeyTabelle(wsD, lastRow)
    
    ' Schutz wieder aktivieren
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' Hinweis bei Zeilen mit erforderlichem Eingriff
    If zeilenMitEingriff.Count > 0 Then
        Call ZeigeEingriffsHinweis(wsD, zeilenMitEingriff, zeilenNeu, zeilenUnveraendert)
    Else
        MsgBox "EntityKey-Aktualisierung abgeschlossen!" & vbCrLf & vbCrLf & _
               "Neue Zeilen verarbeitet: " & zeilenNeu & vbCrLf & _
               "Bestehende Zeilen unverändert: " & zeilenUnveraendert & vbCrLf & vbCrLf & _
               "Alle Zuordnungen sind vollständig.", vbInformation, "Aktualisierung abgeschlossen"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler bei EntityKey-Aktualisierung: " & Err.Description, vbCritical
End Sub

' ===============================================================
' HILFSFUNKTION: Prüft ob Zeile bereits gültige Daten hat
' ===============================================================
Private Function HatBereitsGueltigeDaten(ByVal entityKey As String, _
                                          ByVal zuordnung As String, _
                                          ByVal role As String) As Boolean
    
    ' Eine Zeile gilt als "bereits zugeordnet" wenn:
    ' - EntityKey vorhanden UND kein reiner Zahlenwert (alte Nummerierung)
    ' - ODER Zuordnung UND Role vorhanden
    
    HatBereitsGueltigeDaten = False
    
    ' EntityKey vorhanden und kein reiner Zahlenwert
    If entityKey <> "" Then
        If Not IsNumeric(entityKey) Then
            HatBereitsGueltigeDaten = True
            Exit Function
        End If
    End If
    
    ' Zuordnung UND Role vorhanden
    If zuordnung <> "" And role <> "" Then
        HatBereitsGueltigeDaten = True
        Exit Function
    End If
    
End Function

' ===============================================================
' HILFSPROZEDUR: Zeigt Hinweis für Zeilen mit erforderlichem Eingriff
' ===============================================================
Private Sub ZeigeEingriffsHinweis(ByRef ws As Worksheet, ByRef zeilen As Collection, _
                                   ByVal zeilenNeu As Long, ByVal zeilenUnveraendert As Long)
    
    Dim msg As String
    Dim antwort As VbMsgBoxResult
    Dim ersteZeile As Long
    
    ersteZeile = zeilen(1)
    
    msg = "EntityKey-Aktualisierung abgeschlossen!" & vbCrLf & vbCrLf
    msg = msg & "Neue Zeilen verarbeitet: " & zeilenNeu & vbCrLf
    msg = msg & "Bestehende Zeilen unverändert: " & zeilenUnveraendert & vbCrLf & vbCrLf
    msg = msg & "ACHTUNG: " & zeilen.Count & " Zeile(n) erfordern Ihre manuelle Zuordnung!" & vbCrLf & vbCrLf
    msg = msg & "Diese Zeilen sind ROT markiert." & vbCrLf
    msg = msg & "Bitte ordnen Sie zu, ob es sich um:" & vbCrLf
    msg = msg & "  • MITGLIED (mit/ohne Pacht)" & vbCrLf
    msg = msg & "  • EHEMALIGES MITGLIED" & vbCrLf
    msg = msg & "  • VERSORGER (Strom, Gas, etc.)" & vbCrLf
    msg = msg & "  • BANK" & vbCrLf
    msg = msg & "  • SHOP (Online-Händler)" & vbCrLf
    msg = msg & "handelt." & vbCrLf & vbCrLf
    msg = msg & "Möchten Sie jetzt zur ersten betroffenen Zeile springen?"
    
    antwort = MsgBox(msg, vbYesNo + vbExclamation, "Manuelle Zuordnung erforderlich")
    
    If antwort = vbYes Then
        ws.Activate
        ws.Cells(ersteZeile, EK_COL_ROLE).Select
        
        MsgBox "Sie befinden sich nun in Zeile " & ersteZeile & "." & vbCrLf & vbCrLf & _
               "Nutzen Sie den Menüpunkt:" & vbCrLf & _
               "Alt+F8 ? 'EntityKeyDialogFuerAktuelleZeile'" & vbCrLf & vbCrLf & _
               "oder wählen Sie direkt aus der Dropdown-Liste in Spalte X.", _
               vbInformation, "Hinweis"
    End If
    
End Sub

' ===============================================================
' HILFSFUNKTION: Sucht Mitglieder anhand des Kontonamens
' Durchsucht BEIDE: Aktive Mitglieder UND Mitgliederhistorie!
' ===============================================================
Private Function SucheMitgliederZuKontoname(ByVal kontoName As String, _
                                             ByRef wsM As Worksheet, _
                                             ByRef wsH As Worksheet) As Collection
    
    Dim gefunden As New Collection
    Dim r As Long
    Dim lastRow As Long
    Dim nachname As String
    Dim vorname As String
    Dim memberID As String
    Dim parzelle As String
    Dim funktion As String
    Dim kontoNameNorm As String
    Dim mitgliedInfo(0 To 7) As Variant
    Dim zeilen As Variant
    Dim zeile As Variant
    Dim nameKombiniert As String
    Dim nameParts() As String
    Dim austrittsDatum As Date
    
    Set SucheMitgliederZuKontoname = gefunden
    
    If kontoName = "" Then Exit Function
    
    ' Kontoname kann mehrere Zeilen enthalten (vbLf getrennt)
    zeilen = Split(kontoName, vbLf)
    
    For Each zeile In zeilen
        kontoNameNorm = NormalisiereString(CStr(zeile))
        If kontoNameNorm = "" Then GoTo NextZeile
        
        ' ============================================================
        ' SCHRITT 1: Durchsuche AKTIVE Mitgliederliste
        ' ============================================================
        lastRow = wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
        
        For r = M_START_ROW To lastRow
            ' Nur aktive Mitglieder (ohne Pachtende)
            If Trim(wsM.Cells(r, M_COL_PACHTENDE).value) = "" Then
                nachname = Trim(wsM.Cells(r, M_COL_NACHNAME).value)
                vorname = Trim(wsM.Cells(r, M_COL_VORNAME).value)
                memberID = Trim(wsM.Cells(r, M_COL_MEMBER_ID).value)
                parzelle = Trim(wsM.Cells(r, M_COL_PARZELLE).value)
                funktion = Trim(wsM.Cells(r, M_COL_FUNKTION).value)
                
                ' Prüfe verschiedene Matching-Varianten
                If IstNameImKontoname(nachname, vorname, kontoNameNorm) Then
                    ' Prüfe ob dieses Mitglied bereits in Collection ist
                    If Not IstMitgliedBereitsGefunden(gefunden, memberID, False) Then
                        mitgliedInfo(0) = memberID
                        mitgliedInfo(1) = nachname
                        mitgliedInfo(2) = vorname
                        mitgliedInfo(3) = parzelle
                        mitgliedInfo(4) = funktion
                        mitgliedInfo(5) = r
                        mitgliedInfo(6) = False  ' Nicht ehemalig
                        mitgliedInfo(7) = CDate("01.01.1900")  ' Kein Austrittsdatum
                        gefunden.Add mitgliedInfo
                    End If
                End If
            End If
        Next r
        
        ' ============================================================
        ' SCHRITT 2: Durchsuche MITGLIEDERHISTORIE (ehemalige Mitglieder)
        ' ============================================================
        lastRow = wsH.Cells(wsH.Rows.Count, H_COL_NAME_EHEM_PAECHTER).End(xlUp).Row
        
        For r = H_START_ROW To lastRow
            ' Name ist in Historie als "Nachname, Vorname" gespeichert
            nameKombiniert = Trim(wsH.Cells(r, H_COL_NAME_EHEM_PAECHTER).value)
            memberID = Trim(wsH.Cells(r, H_COL_MEMBER_ID_ALT).value)
            parzelle = Trim(wsH.Cells(r, H_COL_PARZELLE).value)
            
            ' Austrittsdatum lesen
            On Error Resume Next
            austrittsDatum = wsH.Cells(r, H_COL_AUST_DATUM).value
            If Err.Number <> 0 Then austrittsDatum = CDate("01.01.1900")
            On Error GoTo 0
            
            ' Name aufteilen
            If InStr(nameKombiniert, ",") > 0 Then
                nameParts = Split(nameKombiniert, ",")
                nachname = Trim(nameParts(0))
                If UBound(nameParts) >= 1 Then
                    vorname = Trim(nameParts(1))
                Else
                    vorname = ""
                End If
            Else
                nachname = nameKombiniert
                vorname = ""
            End If
            
            ' Prüfe verschiedene Matching-Varianten
            If IstNameImKontoname(nachname, vorname, kontoNameNorm) Then
                ' Prüfe ob dieses Mitglied bereits in Collection ist
                If Not IstMitgliedBereitsGefunden(gefunden, memberID, True) Then
                    mitgliedInfo(0) = memberID
                    mitgliedInfo(1) = nachname
                    mitgliedInfo(2) = vorname
                    mitgliedInfo(3) = parzelle
                    mitgliedInfo(4) = "Ehemaliges Mitglied"
                    mitgliedInfo(5) = r
                    mitgliedInfo(6) = True  ' Ehemalig!
                    mitgliedInfo(7) = austrittsDatum
                    gefunden.Add mitgliedInfo
                End If
            End If
        Next r
        
NextZeile:
    Next zeile
    
    Set SucheMitgliederZuKontoname = gefunden
    
End Function

' ===============================================================
' HILFSFUNKTION: Prüft ob Name im Kontonamen enthalten ist
' ===============================================================
Private Function IstNameImKontoname(ByVal nachname As String, ByVal vorname As String, _
                                     ByVal kontoNameNorm As String) As Boolean
    
    Dim nachnameNorm As String
    Dim vornameNorm As String
    
    nachnameNorm = NormalisiereString(nachname)
    vornameNorm = NormalisiereString(vorname)
    
    IstNameImKontoname = False
    
    If nachnameNorm = "" Then Exit Function
    
    ' Variante 1: "Nachname Vorname" oder "Vorname Nachname"
    If InStr(kontoNameNorm, nachnameNorm & " " & vornameNorm) > 0 Then
        IstNameImKontoname = True
        Exit Function
    End If
    
    If vornameNorm <> "" Then
        If InStr(kontoNameNorm, vornameNorm & " " & nachnameNorm) > 0 Then
            IstNameImKontoname = True
            Exit Function
        End If
    End If
    
    ' Variante 2: Beide Namen separat enthalten (für "Müller Hans und Maria")
    If Len(nachnameNorm) >= 3 And Len(vornameNorm) >= 3 Then
        If InStr(kontoNameNorm, nachnameNorm) > 0 And InStr(kontoNameNorm, vornameNorm) > 0 Then
            IstNameImKontoname = True
            Exit Function
        End If
    End If
    
    ' Variante 3: Nur Nachname bei langem, eindeutigem Namen
    If Len(nachnameNorm) >= 6 Then
        If InStr(kontoNameNorm, nachnameNorm) > 0 Then
            IstNameImKontoname = True
            Exit Function
        End If
    End If
    
End Function

' ===============================================================
' HILFSFUNKTION: Normalisiert String für Vergleich
' ===============================================================
Private Function NormalisiereString(ByVal s As String) As String
    Dim result As String
    
    result = LCase(Trim(s))
    result = Replace(result, ",", " ")
    result = Replace(result, ".", " ")
    result = Replace(result, "-", " ")
    result = Replace(result, "  ", " ")
    result = Application.WorksheetFunction.Trim(result)
    
    ' Umlaute normalisieren
    result = Replace(result, "ä", "ae")
    result = Replace(result, "ö", "oe")
    result = Replace(result, "ü", "ue")
    result = Replace(result, "ß", "ss")
    
    NormalisiereString = result
End Function

' ===============================================================
' HILFSFUNKTION: Prüft ob MemberID bereits in Collection ist
' ===============================================================
Private Function IstMitgliedBereitsGefunden(ByRef col As Collection, _
                                             ByVal memberID As String, _
                                             ByVal istEhemalig As Boolean) As Boolean
    Dim item As Variant
    
    IstMitgliedBereitsGefunden = False
    
    For Each item In col
        ' Prüfe MemberID UND ob es der gleiche Status (aktiv/ehemalig) ist
        If item(0) = memberID And item(6) = istEhemalig Then
            IstMitgliedBereitsGefunden = True
            Exit Function
        End If
    Next item
End Function

' ===============================================================
' HILFSPROZEDUR: Generiert EntityKey und Zuordnung
' Berücksichtigt aktive UND ehemalige Mitglieder
' ===============================================================
Private Sub GeneriereEntityKeyUndZuordnung(ByRef mitglieder As Collection, _
                                            ByVal kontoName As String, _
                                            ByRef outEntityKey As String, _
                                            ByRef outZuordnung As String, _
                                            ByRef outParzellen As String, _
                                            ByRef outEntityRole As String, _
                                            ByRef outDebugInfo As String, _
                                            ByRef outAmpelStatus As Long)
    
    Dim mitgliedInfo As Variant
    Dim i As Long
    Dim memberIDs As String
    Dim uniqueMemberIDs As Object
    Dim hatAktiveMitglieder As Boolean
    Dim hatEhemaligeMitglieder As Boolean
    Dim anzahlAktive As Long
    Dim anzahlEhemalige As Long
    
    Set uniqueMemberIDs = CreateObject("Scripting.Dictionary")
    
    outEntityKey = ""
    outZuordnung = ""
    outParzellen = ""
    outEntityRole = ""
    outDebugInfo = ""
    outAmpelStatus = 1  ' Default: Grün
    
    ' Zähle aktive und ehemalige Mitglieder
    hatAktiveMitglieder = False
    hatEhemaligeMitglieder = False
    anzahlAktive = 0
    anzahlEhemalige = 0
    
    For i = 1 To mitglieder.Count
        mitgliedInfo = mitglieder(i)
        If mitgliedInfo(6) = True Then  ' Ehemalig
            hatEhemaligeMitglieder = True
            anzahlEhemalige = anzahlEhemalige + 1
        Else
            hatAktiveMitglieder = True
            anzahlAktive = anzahlAktive + 1
        End If
    Next i
    
    ' ============================================================
    ' Fall 1: Kein Mitglied gefunden
    ' ============================================================
    If mitglieder.Count = 0 Then
        If IstVersorger(kontoName) Then
            outEntityKey = PREFIX_VERSORGER & CreateGUID()
            outEntityRole = ROLE_VERSORGER
            outDebugInfo = "Automatisch als VERSORGER erkannt"
            outAmpelStatus = 1  ' Grün
        ElseIf IstBank(kontoName) Then
            outEntityKey = PREFIX_BANK & CreateGUID()
            outEntityRole = ROLE_BANK
            outDebugInfo = "Automatisch als BANK erkannt"
            outAmpelStatus = 1  ' Grün
        ElseIf IstShop(kontoName) Then
            outEntityKey = PREFIX_SHOP & CreateGUID()
            outEntityRole = ROLE_SHOP
            outDebugInfo = "Automatisch als SHOP erkannt"
            outAmpelStatus = 1  ' Grün
        Else
            ' MANUELLE ZUORDNUNG ERFORDERLICH
            outEntityKey = ""
            outEntityRole = ""
            outDebugInfo = "MANUELLE ZUORDNUNG ERFORDERLICH - Kein Mitglied gefunden"
            outAmpelStatus = 3  ' Rot
        End If
        Exit Sub
    End If
    
    ' ============================================================
    ' Fall 2: NUR ehemalige Mitglieder gefunden
    ' ============================================================
    If hatEhemaligeMitglieder And Not hatAktiveMitglieder Then
        ' Sammle alle eindeutigen Member-IDs
        For i = 1 To mitglieder.Count
            mitgliedInfo = mitglieder(i)
            If Not uniqueMemberIDs.Exists(CStr(mitgliedInfo(0))) Then
                uniqueMemberIDs.Add CStr(mitgliedInfo(0)), CStr(mitgliedInfo(0))
            End If
        Next i
        
        If uniqueMemberIDs.Count = 1 Then
            ' Ein ehemaliges Mitglied
            mitgliedInfo = mitglieder(1)
            outEntityKey = PREFIX_EHEMALIG & mitgliedInfo(0)
            outZuordnung = mitgliedInfo(1) & ", " & mitgliedInfo(2)
            outParzellen = mitgliedInfo(3) & " (bis " & Format(mitgliedInfo(7), "dd.mm.yyyy") & ")"
            outEntityRole = ROLE_EHEMALIGES_MITGLIED
            outDebugInfo = "Ehemaliges Mitglied gefunden (Austritt: " & Format(mitgliedInfo(7), "dd.mm.yyyy") & ")"
            outAmpelStatus = 1  ' Grün
        Else
            ' Mehrere ehemalige Mitglieder (Gemeinschaftskonto)
            memberIDs = ""
            Dim key As Variant
            For Each key In uniqueMemberIDs.Keys
                If memberIDs <> "" Then memberIDs = memberIDs & "_"
                memberIDs = memberIDs & key
            Next key
            
            outEntityKey = PREFIX_SHARE & PREFIX_EHEMALIG & memberIDs
            outEntityRole = ROLE_EHEMALIGES_MITGLIED
            outDebugInfo = "Ehem. Gemeinschaftskonto - " & uniqueMemberIDs.Count & " Personen"
            outAmpelStatus = 1  ' Grün
            
            ' Zuordnung und Parzellen zusammensetzen
            For i = 1 To mitglieder.Count
                mitgliedInfo = mitglieder(i)
                If outZuordnung <> "" Then outZuordnung = outZuordnung & vbLf
                outZuordnung = outZuordnung & mitgliedInfo(1) & ", " & mitgliedInfo(2) & " (ehem.)"
                
                Dim parzelleInfo As String
                parzelleInfo = CStr(mitgliedInfo(3))
                If InStr(outParzellen, parzelleInfo) = 0 Then
                    If outParzellen <> "" Then outParzellen = outParzellen & vbLf
                    outParzellen = outParzellen & parzelleInfo
                End If
            Next i
        End If
        Exit Sub
    End If
    
    ' ============================================================
    ' Fall 3: Aktive Mitglieder gefunden (ggf. auch ehemalige)
    ' ============================================================
    
    ' Sammle nur AKTIVE Member-IDs für EntityKey
    For i = 1 To mitglieder.Count
        mitgliedInfo = mitglieder(i)
        If mitgliedInfo(6) = False Then  ' Nur aktive
            If Not uniqueMemberIDs.Exists(CStr(mitgliedInfo(0))) Then
                uniqueMemberIDs.Add CStr(mitgliedInfo(0)), CStr(mitgliedInfo(0))
            End If
        End If
    Next i
    
    ' Genau 1 aktives Mitglied
    If uniqueMemberIDs.Count = 1 Then
        ' Finde das erste aktive Mitglied
        For i = 1 To mitglieder.Count
            mitgliedInfo = mitglieder(i)
            If mitgliedInfo(6) = False Then
                outEntityKey = CStr(mitgliedInfo(0))
                outZuordnung = mitgliedInfo(1) & ", " & mitgliedInfo(2)
                outEntityRole = ErmittleEntityRoleVonFunktion(CStr(mitgliedInfo(4)))
                outDebugInfo = "Eindeutiger Treffer"
                outAmpelStatus = 1  ' Grün
                Exit For
            End If
        Next i
        
        ' Parzellen sammeln (von allen aktiven Einträgen mit gleicher ID)
        For i = 1 To mitglieder.Count
            mitgliedInfo = mitglieder(i)
            If mitgliedInfo(6) = False Then
                If InStr(outParzellen, CStr(mitgliedInfo(3))) = 0 Then
                    If outParzellen <> "" Then outParzellen = outParzellen & vbLf
                    outParzellen = outParzellen & CStr(mitgliedInfo(3))
                End If
            End If
        Next i
        
        ' Hinweis wenn auch ehemalige gefunden wurden
        If hatEhemaligeMitglieder Then
            outDebugInfo = outDebugInfo & " (+ " & anzahlEhemalige & " ehem. Einträge in Historie)"
        End If
        
        Exit Sub
    End If
    
    ' Mehrere aktive Mitglieder (Gemeinschaftskonto)
    If uniqueMemberIDs.Count > 1 Then
        memberIDs = ""
        For Each key In uniqueMemberIDs.Keys
            If memberIDs <> "" Then memberIDs = memberIDs & "_"
            memberIDs = memberIDs & key
        Next key
        
        outEntityKey = PREFIX_SHARE & memberIDs
        outEntityRole = ROLE_MITGLIED_MIT_PACHT
        outDebugInfo = "Gemeinschaftskonto - " & uniqueMemberIDs.Count & " Personen"
        outAmpelStatus = 1  ' Grün
        
        ' Zuordnung und Parzellen zusammensetzen (nur aktive)
        For i = 1 To mitglieder.Count
            mitgliedInfo = mitglieder(i)
            If mitgliedInfo(6) = False Then  ' Nur aktive
                If outZuordnung <> "" Then outZuordnung = outZuordnung & vbLf
                outZuordnung = outZuordnung & mitgliedInfo(1) & ", " & mitgliedInfo(2)
                
                If InStr(outParzellen, CStr(mitgliedInfo(3))) = 0 Then
                    If outParzellen <> "" Then outParzellen = outParzellen & vbLf
                    outParzellen = outParzellen & CStr(mitgliedInfo(3))
                End If
            End If
        Next i
        
        ' Hinweis wenn auch ehemalige gefunden wurden
        If hatEhemaligeMitglieder Then
            outDebugInfo = outDebugInfo & " (+ " & anzahlEhemalige & " ehem. Einträge)"
        End If
    End If
    
End Sub

' ===============================================================
' HILFSFUNKTION: Ermittelt EntityRole aus Funktion
' ===============================================================
Private Function ErmittleEntityRoleVonFunktion(ByVal funktion As String) As String
    
    Select Case funktion
        Case "Mitglied ohne Pacht"
            ErmittleEntityRoleVonFunktion = ROLE_MITGLIED_OHNE_PACHT
        Case "Ehemaliges Mitglied"
            ErmittleEntityRoleVonFunktion = ROLE_EHEMALIGES_MITGLIED
        Case Else
            ErmittleEntityRoleVonFunktion = ROLE_MITGLIED_MIT_PACHT
    End Select
    
End Function

' ===============================================================
' HILFSFUNKTIONEN: Erkennung von Versorger/Bank/Shop
' ERWEITERT mit mehr Schlüsselwörtern
' ===============================================================
Private Function IstVersorger(ByVal name As String) As Boolean
    Dim keywords As Variant
    Dim kw As Variant
    
    ' Erweiterte Liste für Versorger
    keywords = Array( _
        "stadtwerke", "energie", "strom", "gas", "wasser", _
        "telekom", "vodafone", "o2", "1&1", "versicherung", _
        "allianz", "huk", "devk", "axa", "ergo", "enviam", _
        "enso", "ewe", "eon", "e.on", "rwe", "vattenfall", _
        "gvv", "signal iduna", "debeka", "lvm", "abfall", _
        "müll", "entsorgung", "abwasser", "kanal", _
        "wazv", "zweckverband", "wasserverband", "abwasserverband", _
        "grundstücksgesellschaft", "wohnungsbau", "wohnungsgesellschaft", _
        "hausverwaltung", "immobilien", "grundstück", _
        "finanzamt", "rundfunk", "gez", "beitragsservice", _
        "kfz", "haftpflicht", "hausrat", "rechtsschutz", _
        "krankenkasse", "aok", "barmer", "dak", "tk", "ikk", _
        "berufsgenossenschaft", "rentenversicherung", _
        "stadt ", "gemeinde ", "kommune", "landkreis", _
        "werder", "havel", "potsdam", "brandenburg")
    
    name = LCase(name)
    
    For Each kw In keywords
        If InStr(name, kw) > 0 Then
            IstVersorger = True
            Exit Function
        End If
    Next kw
    
    IstVersorger = False
End Function

Private Function IstBank(ByVal name As String) As Boolean
    Dim keywords As Variant
    Dim kw As Variant
    
    ' Erweiterte Liste für Banken
    keywords = Array( _
        "sparkasse", "volksbank", "raiffeisenbank", "commerzbank", _
        "deutsche bank", "postbank", "ing", "dkb", "targobank", _
        "sparda", "psd bank", "santander", "hypovereinsbank", _
        "unicredit", "n26", "comdirect", "consorsbank", _
        "mittelbrandenburgische", "mbs", "brandenburger bank", _
        "kreditbank", "landesbank", "girozentrale", _
        "bausparkasse", "schwäbisch hall", "lbs", "wüstenrot")
    
    name = LCase(name)
    
    For Each kw In keywords
        If InStr(name, kw) > 0 Then
            IstBank = True
            Exit Function
        End If
    Next kw
    
    IstBank = False
End Function

Private Function IstShop(ByVal name As String) As Boolean
    Dim keywords As Variant
    Dim kw As Variant
    
    ' Erweiterte Liste für Shops
    keywords = Array( _
        "amazon", "ebay", "paypal", "otto", "zalando", _
        "mediamarkt", "saturn", "lidl", "aldi", "rewe", _
        "edeka", "penny", "netto", "kaufland", "hornbach", _
        "obi", "bauhaus", "toom", "hagebau", "dehner", _
        "rossmann", "dm-drogerie", "müller drogerie", _
        "ikea", "poco", "roller", "mömax", "xxxlutz", _
        "h&m", "c&a", "kik", "takko", "ernsting", _
        "decathlon", "intersport", "karstadt", "galeria", _
        "thalia", "hugendubel", "weltbild", _
        "notebooksbilliger", "cyberport", "alternate", _
        "thomann", "musicstore", "conrad", "reichelt", _
        "fressnapf", "zooplus", "futterhaus", _
        "apotheke", "docmorris", "shop-apotheke")
    
    name = LCase(name)
    
    For Each kw In keywords
        If InStr(name, kw) > 0 Then
            IstShop = True
            Exit Function
        End If
    Next kw
    
    IstShop = False
End Function

' ===============================================================
' HILFSFUNKTION: Generiert neue GUID
' ===============================================================
Private Function CreateGUID() As String
    CreateGUID = mod_Mitglieder_UI.CreateGUID_Public()
End Function

' ===============================================================
' HILFSPROZEDUR: Setzt Dropdown für EntityRole
' Dropdown nur bis lastRow + 50 Pufferzeilen
' ===============================================================
Private Sub SetupEntityRoleDropdown(ByRef ws As Worksheet, ByVal lastRow As Long)
    
    Dim rngDropdown As Range
    Dim dropdownSource As String
    Dim lastRoleRow As Long
    Dim dropdownEndRow As Long
    
    lastRoleRow = ws.Cells(ws.Rows.Count, EK_ROLE_DROPDOWN_COL).End(xlUp).Row
    If lastRoleRow < 4 Then lastRoleRow = 10
    
    dropdownSource = "=$AF$4:$AF$" & lastRoleRow
    
    ' Dropdown nur bis lastRow + 50 Pufferzeilen für neue Einträge
    dropdownEndRow = lastRow + 50
    
    Set rngDropdown = ws.Range(ws.Cells(EK_START_ROW, EK_COL_ROLE), _
                               ws.Cells(dropdownEndRow, EK_COL_ROLE))
    
    On Error Resume Next
    With rngDropdown.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=dropdownSource
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
        .ErrorTitle = "Ungültige Eingabe"
        .ErrorMessage = "Bitte wählen Sie einen Wert aus der Liste."
    End With
    On Error GoTo 0
    
End Sub

' ===============================================================
' HILFSPROZEDUR: Setzt Ampelfarbe (nur Spalten V-Y)
' Grün = OK, Gelb = Prüfen, Rot = Eingriff nötig
' ===============================================================
Private Sub SetzeAmpelFarbe(ByRef ws As Worksheet, ByVal zeile As Long, ByVal ampelStatus As Long)
    
    Dim rng As Range
    Dim farbe As Long
    
    Set rng = ws.Range(ws.Cells(zeile, EK_COL_ZUORDNUNG), ws.Cells(zeile, EK_COL_DEBUG))
    
    Select Case ampelStatus
        Case 1
            farbe = RGB(198, 224, 180)  ' Grün - OK
        Case 2
            farbe = RGB(255, 230, 153)  ' Gelb - Prüfen
        Case 3
            farbe = RGB(255, 150, 150)  ' Rot - Eingriff nötig
        Case Else
            farbe = RGB(198, 224, 180)  ' Default: Grün
    End Select
    
    rng.Interior.color = farbe
    
End Sub

' ===============================================================
' HILFSPROZEDUR: Formatiert die EntityKey-Tabelle
' NUR bis zur letzten Datenzeile - NICHT darüber hinaus!
' Spalte S: Feste Breite 12, KEIN Textumbruch
' Spalten T-Y: AutoFit Breite, dann Zeilenhöhe AutoFit
' ===============================================================
Private Sub FormatiereEntityKeyTabelle(ByRef ws As Worksheet, ByVal lastRow As Long)
    
    Dim rngTable As Range
    Dim rngOhneEntityKey As Range
    
    If lastRow < EK_START_ROW Then Exit Sub
    
    ' NUR den Datenbereich formatieren (nicht darüber hinaus!)
    Set rngTable = ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                            ws.Cells(lastRow, EK_COL_DEBUG))
    
    ' Bereich OHNE EntityKey-Spalte (für Textumbruch)
    Set rngOhneEntityKey = ws.Range(ws.Cells(EK_START_ROW, EK_COL_IBAN), _
                                     ws.Cells(lastRow, EK_COL_DEBUG))
    
    ' Rahmen NUR für Datenbereich
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' Vertikale Ausrichtung für alle
    rngTable.VerticalAlignment = xlCenter
    
    ' Textumbruch NUR für Spalten T-Y (NICHT für S!)
    rngOhneEntityKey.WrapText = True
    
    ' Spalte S (EntityKey): KEIN Textumbruch, linksbündig, FESTE BREITE 12
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                  ws.Cells(lastRow, EK_COL_ENTITYKEY))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ENTITYKEY).ColumnWidth = 12
    
    ' Spalte W (Parzelle): zentriert
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_PARZELLE), _
             ws.Cells(lastRow, EK_COL_PARZELLE)).HorizontalAlignment = xlCenter
    
    ' Spalte X (Role): zentriert
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_ROLE), _
             ws.Cells(lastRow, EK_COL_ROLE)).HorizontalAlignment = xlCenter
    
    ' SCHRITT 1: AutoFit Spaltenbreite NUR für Spalten T-Y
    Dim col As Long
    For col = EK_COL_IBAN To EK_COL_DEBUG
        ws.Columns(col).AutoFit
    Next col
    
    ' SCHRITT 2: Zeilenhöhe AutoFit für gesamte Tabelle (ab Zeile 4)
    ws.Rows(EK_START_ROW & ":" & lastRow).AutoFit
    
End Sub

' ===============================================================
' DIALOG: Manuelle EntityKey-Zuordnung für aktuelle Zeile
' ===============================================================
Public Sub EntityKeyDialogFuerAktuelleZeile()
    
    Dim aktuelleZeile As Long
    Dim wsD As Worksheet
    Dim wsM As Worksheet
    Dim wsH As Worksheet
    Dim iban As String
    Dim kontoName As String
    Dim mitglieder As Collection
    Dim eingabe As String
    Dim auswahlText As String
    Dim i As Long
    Dim mitgliedInfo As Variant
    Dim neuerEntityKey As String
    Dim neueZuordnung As String
    Dim neueParzellen As String
    Dim neueRole As String
    Dim memberIDs As String
    Dim uniqueIDs As Object
    Dim key As Variant
    
    aktuelleZeile = ActiveCell.Row
    
    If aktuelleZeile < EK_START_ROW Then
        MsgBox "Bitte wählen Sie eine Datenzeile (ab Zeile " & EK_START_ROW & ").", vbExclamation
        Exit Sub
    End If
    
    If ActiveSheet.name <> WS_DATEN Then
        MsgBox "Bitte wechseln Sie zum Tabellenblatt 'Daten'.", vbExclamation
        Exit Sub
    End If
    
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    
    iban = Trim(wsD.Cells(aktuelleZeile, EK_COL_IBAN).value)
    kontoName = Trim(wsD.Cells(aktuelleZeile, EK_COL_KONTONAME).value)
    
    Set mitglieder = SucheMitgliederZuKontoname(kontoName, wsM, wsH)
    
    auswahlText = "=== EntityKey-Zuordnung (Zeile " & aktuelleZeile & ") ===" & vbCrLf & vbCrLf
    auswahlText = auswahlText & "IBAN: " & iban & vbCrLf
    auswahlText = auswahlText & "Kontoname: " & Replace(kontoName, vbLf, " / ") & vbCrLf & vbCrLf
    
    If mitglieder.Count > 0 Then
        auswahlText = auswahlText & "Gefundene Mitglieder:" & vbCrLf
        For i = 1 To mitglieder.Count
            mitgliedInfo = mitglieder(i)
            auswahlText = auswahlText & "  " & i & ") " & mitgliedInfo(1) & ", " & mitgliedInfo(2)
            If mitgliedInfo(6) = True Then
                auswahlText = auswahlText & " [EHEMALIG - Austritt: " & Format(mitgliedInfo(7), "dd.mm.yyyy") & "]"
            End If
            auswahlText = auswahlText & " (Parzelle " & mitgliedInfo(3) & ")" & vbCrLf
        Next i
        auswahlText = auswahlText & vbCrLf
    Else
        auswahlText = auswahlText & "Keine Mitglieder gefunden." & vbCrLf & vbCrLf
    End If
    
    auswahlText = auswahlText & "Bitte wählen Sie:" & vbCrLf
    auswahlText = auswahlText & "  M = MITGLIED (aktiv)" & vbCrLf
    auswahlText = auswahlText & "  E = EHEMALIGES MITGLIED" & vbCrLf
    auswahlText = auswahlText & "  G = GEMEINSCHAFTSKONTO" & vbCrLf
    auswahlText = auswahlText & "  V = VERSORGER" & vbCrLf
    auswahlText = auswahlText & "  B = BANK" & vbCrLf
    auswahlText = auswahlText & "  S = SHOP" & vbCrLf
    auswahlText = auswahlText & "  X = Abbrechen"
    
    eingabe = UCase(Trim(InputBox(auswahlText, "EntityKey-Zuordnung", "M")))
    
    If eingabe = "" Or eingabe = "X" Then Exit Sub
    
    wsD.Unprotect PASSWORD:=PASSWORD
    
    Set uniqueIDs = CreateObject("Scripting.Dictionary")
    
    Select Case eingabe
        Case "M"
            ' Erstes AKTIVES Mitglied
            For i = 1 To mitglieder.Count
                mitgliedInfo = mitglieder(i)
                If mitgliedInfo(6) = False Then
                    neuerEntityKey = CStr(mitgliedInfo(0))
                    neueZuordnung = mitgliedInfo(1) & ", " & mitgliedInfo(2)
                    neueParzellen = CStr(mitgliedInfo(3))
                    neueRole = ROLE_MITGLIED_MIT_PACHT
                    Exit For
                End If
            Next i
            If neuerEntityKey = "" Then
                neuerEntityKey = CreateGUID()
                neueZuordnung = kontoName
                neueRole = ROLE_MITGLIED_MIT_PACHT
            End If
            
        Case "E"
            ' Erstes EHEMALIGES Mitglied
            For i = 1 To mitglieder.Count
                mitgliedInfo = mitglieder(i)
                If mitgliedInfo(6) = True Then
                    neuerEntityKey = PREFIX_EHEMALIG & CStr(mitgliedInfo(0))
                    neueZuordnung = mitgliedInfo(1) & ", " & mitgliedInfo(2) & " (ehem.)"
                    neueParzellen = CStr(mitgliedInfo(3))
                    neueRole = ROLE_EHEMALIGES_MITGLIED
                    Exit For
                End If
            Next i
            If neuerEntityKey = "" Then
                neuerEntityKey = PREFIX_EHEMALIG & CreateGUID()
                neueZuordnung = kontoName
                neueRole = ROLE_EHEMALIGES_MITGLIED
            End If
            
        Case "G"
            memberIDs = ""
            For i = 1 To mitglieder.Count
                mitgliedInfo = mitglieder(i)
                If Not uniqueIDs.Exists(CStr(mitgliedInfo(0))) Then
                    uniqueIDs.Add CStr(mitgliedInfo(0)), CStr(mitgliedInfo(0))
                    If memberIDs <> "" Then memberIDs = memberIDs & "_"
                    memberIDs = memberIDs & CStr(mitgliedInfo(0))
                End If
                
                If neueZuordnung <> "" Then neueZuordnung = neueZuordnung & vbLf
                neueZuordnung = neueZuordnung & mitgliedInfo(1) & ", " & mitgliedInfo(2)
                If mitgliedInfo(6) = True Then neueZuordnung = neueZuordnung & " (ehem.)"
                
                If InStr(neueParzellen, CStr(mitgliedInfo(3))) = 0 Then
                    If neueParzellen <> "" Then neueParzellen = neueParzellen & vbLf
                    neueParzellen = neueParzellen & CStr(mitgliedInfo(3))
                End If
            Next i
            
            If memberIDs = "" Then memberIDs = CreateGUID()
            neuerEntityKey = PREFIX_SHARE & memberIDs
            neueRole = ROLE_MITGLIED_MIT_PACHT
            
        Case "V"
            neuerEntityKey = PREFIX_VERSORGER & CreateGUID()
            neueRole = ROLE_VERSORGER
            neueZuordnung = kontoName
            
        Case "B"
            neuerEntityKey = PREFIX_BANK & CreateGUID()
            neueRole = ROLE_BANK
            neueZuordnung = kontoName
            
        Case "S"
            neuerEntityKey = PREFIX_SHOP & CreateGUID()
            neueRole = ROLE_SHOP
            neueZuordnung = kontoName
            
        Case Else
            MsgBox "Ungültige Eingabe.", vbExclamation
            wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
            Exit Sub
    End Select
    
    ' NUR LEERE Felder füllen!
    If Trim(wsD.Cells(aktuelleZeile, EK_COL_ENTITYKEY).value) = "" Then
        wsD.Cells(aktuelleZeile, EK_COL_ENTITYKEY).value = neuerEntityKey
    End If
    If Trim(wsD.Cells(aktuelleZeile, EK_COL_ZUORDNUNG).value) = "" Then
        wsD.Cells(aktuelleZeile, EK_COL_ZUORDNUNG).value = neueZuordnung
    End If
    If Trim(wsD.Cells(aktuelleZeile, EK_COL_PARZELLE).value) = "" And neueParzellen <> "" Then
        wsD.Cells(aktuelleZeile, EK_COL_PARZELLE).value = neueParzellen
    End If
    If Trim(wsD.Cells(aktuelleZeile, EK_COL_ROLE).value) = "" Then
        wsD.Cells(aktuelleZeile, EK_COL_ROLE).value = neueRole
    End If
    
    wsD.Cells(aktuelleZeile, EK_COL_DEBUG).value = "Manuell zugeordnet am " & Format(Now, "dd.mm.yyyy hh:mm")
    
    Call SetzeAmpelFarbe(wsD, aktuelleZeile, 1)
    
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    MsgBox "EntityKey erfolgreich zugeordnet!" & vbCrLf & vbCrLf & _
           "EntityKey: " & neuerEntityKey & vbCrLf & _
           "Rolle: " & neueRole, vbInformation, "Zuordnung erfolgreich"
    
End Sub

' ===============================================================
' ÖFFENTLICHE PROZEDUR: Wird nach CSV-Import aufgerufen
' ===============================================================
Public Sub NachCSVImport_EntityKeysAktualisieren()
    Call AktualisiereAlleEntityKeys
End Sub

' ===============================================================
' HILFSPROZEDUR: Entfernt überflüssige Rahmenlinien
' (Einmalig ausführen wenn nötig)
' ===============================================================
Public Sub EntferneUeberfluesstigeRahmen()
    
    Dim ws As Worksheet
    Dim lastDataRow As Long
    Dim rngZuLoeschen As Range
    
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    ' Finde letzte echte Datenzeile (basierend auf IBAN in Spalte T)
    lastDataRow = ws.Cells(ws.Rows.Count, EK_COL_IBAN).End(xlUp).Row
    If lastDataRow < EK_START_ROW Then lastDataRow = EK_START_ROW
    
    ' Lösche Rahmen und Farben ab lastDataRow+1 bis Zeile 1000 in Spalten S-Y
    If lastDataRow < 1000 Then
        Set rngZuLoeschen = ws.Range(ws.Cells(lastDataRow + 1, EK_COL_ENTITYKEY), ws.Cells(1000, EK_COL_DEBUG))
        rngZuLoeschen.Borders.LineStyle = xlNone
        rngZuLoeschen.Interior.ColorIndex = xlNone
    End If
    
    MsgBox "Überflüssige Rahmenlinien entfernt!" & vbCrLf & _
           "Letzte Datenzeile: " & lastDataRow, vbInformation
    
End Sub

