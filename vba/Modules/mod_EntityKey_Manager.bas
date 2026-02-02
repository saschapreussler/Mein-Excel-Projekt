Attribute VB_Name = "mod_EntityKey_Manager"
' ***************************************************************
' MODUL: mod_EntityKey_Manager
' ZWECK: Verwaltung und Zuordnung von EntityKeys fuer Bankverkehr
' VERSION: 2.2 - 02.02.2026
' AENDERUNG: Alle Spalten korrigiert, EntityRole-DropDown dynamisch auf AD
'            GUID-Generierung fuer Mitglieder, VERSORGER, BANK etc.
' ***************************************************************

' ===============================================================
' KEINE SPALTEN-KONSTANTEN HIER - ALLE AUS mod_Const.bas!
' EK_COL_ENTITYKEY=18(R), EK_COL_IBAN=19(S), EK_COL_KONTONAME=20(T)
' EK_COL_ZUORDNUNG=21(U), EK_COL_PARZELLE=22(V), EK_COL_ROLE=23(W)
' EK_COL_DEBUG=24(X), DATA_COL_DD_ENTITYROLE=30(AD)
' ===============================================================

' EntityRole-Praefixe fuer GUID-Generierung
Public Const PREFIX_SHARE As String = "SHARE-"
Public Const PREFIX_VERSORGER As String = "VERS-"
Public Const PREFIX_BANK As String = "BANK-"
Public Const PREFIX_SHOP As String = "SHOP-"
Public Const PREFIX_EHEMALIG As String = "EX-"
Public Const PREFIX_SONSTIGE As String = "SONSTIGE-"

' EntityRole-Werte
Public Const ROLE_MITGLIED_MIT_PACHT As String = "MITGLIED_MIT_PACHT"
Public Const ROLE_MITGLIED_OHNE_PACHT As String = "MITGLIED_OHNE_PACHT"
Public Const ROLE_EHEMALIGES_MITGLIED As String = "EHEMALIGES_MITGLIED"
Public Const ROLE_VERSORGER As String = "VERSORGER"
Public Const ROLE_BANK As String = "BANK"
Public Const ROLE_SHOP As String = "SHOP"
Public Const ROLE_SONSTIGE As String = "SONSTIGE"

' Zebra-Farbe
Private Const ZEBRA_COLOR As Long = &HDEE5E3

' ===============================================================
' HILFSFUNKTION: Prueft ob Role eine Parzelle haben darf
' Erlaubt fuer: Mitglieder (alle Arten) und SONSTIGE
' Nicht erlaubt fuer: VERSORGER, BANK, SHOP
' ===============================================================
Private Function DarfParzelleHaben(ByVal role As String) As Boolean
    Dim normRole As String
    normRole = NormalisiereRoleString(role)
    
    DarfParzelleHaben = True
    
    If normRole = "VERSORGER" Or normRole = "BANK" Or normRole = "SHOP" Then
        DarfParzelleHaben = False
    End If
End Function

' ===============================================================
' OEFFENTLICHE PROZEDUR: Importiert IBANs aus Bankkonto
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
    Dim currentDatum As Variant
    Dim ibanKey As Variant
    Dim anzahlNeu As Long
    Dim anzahlBereitsVorhanden As Long
    Dim anzahlZeilenGeprueft As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set dictIBANs = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    wsD.Unprotect PASSWORD:=PASSWORD
    wsBK.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    anzahlNeu = 0
    anzahlBereitsVorhanden = 0
    anzahlZeilenGeprueft = 0
    
    ' Sammle bereits vorhandene IBANs aus Daten-Blatt (Spalte S = 19)
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
    
    Dim dictNeueIBANs As Object
    Set dictNeueIBANs = CreateObject("Scripting.Dictionary")
    
    lastRowBK = wsBK.Cells(wsBK.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    
    For r = BK_START_ROW To lastRowBK
        currentDatum = wsBK.Cells(r, BK_COL_DATUM).value
        
        If Not IsEmpty(currentDatum) And currentDatum <> "" Then
            anzahlZeilenGeprueft = anzahlZeilenGeprueft + 1
            
            currentIBAN = NormalisiereIBAN(wsBK.Cells(r, BK_COL_IBAN).value)
            currentKontoName = Trim(wsBK.Cells(r, BK_COL_NAME).value)
            
            If currentIBAN <> "" And currentIBAN <> "N.A." And Len(currentIBAN) >= 15 Then
                If Not dictIBANs.Exists(currentIBAN) Then
                    If Not dictNeueIBANs.Exists(currentIBAN) Then
                        dictNeueIBANs.Add currentIBAN, currentKontoName
                    Else
                        If InStr(dictNeueIBANs(currentIBAN), currentKontoName) = 0 Then
                            dictNeueIBANs(currentIBAN) = dictNeueIBANs(currentIBAN) & vbLf & currentKontoName
                        End If
                    End If
                End If
            End If
        End If
    Next r
    
    If lastRowD < EK_START_ROW Then
        nextRowD = EK_START_ROW
    Else
        nextRowD = lastRowD + 1
    End If
    
    For Each ibanKey In dictNeueIBANs.Keys
        wsD.Cells(nextRowD, EK_COL_IBAN).value = ibanKey
        wsD.Cells(nextRowD, EK_COL_KONTONAME).value = dictNeueIBANs(ibanKey)
        
        anzahlNeu = anzahlNeu + 1
        nextRowD = nextRowD + 1
    Next ibanKey
    
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    If anzahlNeu > 0 Then
        Dim antwort As VbMsgBoxResult
        antwort = MsgBox("Import abgeschlossen!" & vbCrLf & vbCrLf & _
                        "Bankzeilen geprueft: " & anzahlZeilenGeprueft & vbCrLf & _
                        "Neue IBANs importiert: " & anzahlNeu & vbCrLf & _
                        "Bereits vorhanden (uebersprungen): " & anzahlBereitsVorhanden & vbCrLf & vbCrLf & _
                        "Moechten Sie jetzt die automatische Mitglieder-Zuordnung starten?", _
                        vbYesNo + vbQuestion, "IBAN-Import erfolgreich")
        
        If antwort = vbYes Then
            Call AktualisiereAlleEntityKeys
        End If
    Else
        MsgBox "Keine neuen IBANs gefunden!" & vbCrLf & vbCrLf & _
               "Bankzeilen geprueft: " & anzahlZeilenGeprueft & vbCrLf & _
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
' HILFSFUNKTION: Normalisiert IBAN
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
' ===============================================================
Public Sub AktualisiereAlleEntityKeys()
    
    Dim wsD As Worksheet
    Dim wsM As Worksheet
    Dim wsH As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim iban As String
    Dim kontoname As String
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
    Dim ampelStatus As Long
    Dim mitgliederGefunden As Collection
    Dim zeilenRot As Collection
    Dim zeilenGelb As Collection
    Dim zeilenNeu As Long
    Dim zeilenUnveraendert As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    Set zeilenRot = New Collection
    Set zeilenGelb = New Collection
    
    zeilenNeu = 0
    zeilenUnveraendert = 0
    
    On Error Resume Next
    wsD.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    ' Letzte Zeile basierend auf IBAN-Spalte (S = 19)
    lastRow = wsD.Cells(wsD.Rows.Count, EK_COL_IBAN).End(xlUp).Row
    If lastRow < EK_START_ROW Then lastRow = EK_START_ROW
    
    ' EntityRole-DropDown einrichten (dynamisch aus AD)
    Call SetupEntityRoleDropdown(wsD, lastRow)
    
    ' Parzellen-DropDown einrichten (dynamisch aus F)
    Call SetupParzellenDropdown(wsD, lastRow)
    
    For r = EK_START_ROW To lastRow
        iban = Trim(wsD.Cells(r, EK_COL_IBAN).value)
        kontoname = Trim(wsD.Cells(r, EK_COL_KONTONAME).value)
        
        currentEntityKey = Trim(wsD.Cells(r, EK_COL_ENTITYKEY).value)
        currentZuordnung = Trim(wsD.Cells(r, EK_COL_ZUORDNUNG).value)
        currentParzelle = Trim(wsD.Cells(r, EK_COL_PARZELLE).value)
        currentRole = Trim(wsD.Cells(r, EK_COL_ROLE).value)
        currentDebug = Trim(wsD.Cells(r, EK_COL_DEBUG).value)
        
        If iban = "" And kontoname = "" Then GoTo NextRow
        
        ' Pruefen ob bereits gueltige Daten vorhanden
        If HatBereitsGueltigeDaten(currentEntityKey, currentZuordnung, currentRole) Then
            zeilenUnveraendert = zeilenUnveraendert + 1
            If currentRole <> "" Then
                Call SetzeAmpelFarbe(wsD, r, 1)
            End If
            GoTo NextRow
        End If
        
        zeilenNeu = zeilenNeu + 1
        
        ' Mitglieder suchen
        Set mitgliederGefunden = SucheMitgliederZuKontoname(kontoname, wsM, wsH)
        
        ' EntityKey und Zuordnung generieren
        Call GeneriereEntityKeyUndZuordnung(mitgliederGefunden, kontoname, wsM, _
                                             newEntityKey, zuordnung, parzellen, entityRole, debugInfo, ampelStatus)
        
        ' Werte eintragen (nur wenn leer)
        If currentEntityKey = "" And newEntityKey <> "" Then wsD.Cells(r, EK_COL_ENTITYKEY).value = newEntityKey
        If currentZuordnung = "" And zuordnung <> "" Then wsD.Cells(r, EK_COL_ZUORDNUNG).value = zuordnung
        
        ' Parzelle nur setzen wenn erlaubt
        If currentParzelle = "" And parzellen <> "" And DarfParzelleHaben(entityRole) Then
            wsD.Cells(r, EK_COL_PARZELLE).value = parzellen
        End If
        
        If currentRole = "" And entityRole <> "" Then wsD.Cells(r, EK_COL_ROLE).value = entityRole
        If currentDebug = "" Then wsD.Cells(r, EK_COL_DEBUG).value = debugInfo
        
        ' Formatierung
        Call SetzeAmpelFarbe(wsD, r, ampelStatus)
        Call SetzeZellschutzFuerZeile(wsD, r, entityRole)
        
        If ampelStatus = 3 Then
            zeilenRot.Add r
        ElseIf ampelStatus = 2 Then
            zeilenGelb.Add r
        End If
        
NextRow:
    Next r
    
    ' Gesamttabelle formatieren
    Call FormatiereEntityKeyTabelle(wsD, lastRow)
    
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    If zeilenRot.Count > 0 Or zeilenGelb.Count > 0 Then
        Call ZeigeEingriffsHinweis(wsD, zeilenRot, zeilenGelb, zeilenNeu, zeilenUnveraendert)
    Else
        MsgBox "EntityKey-Aktualisierung abgeschlossen!" & vbCrLf & vbCrLf & _
               "Neue Zeilen verarbeitet: " & zeilenNeu & vbCrLf & _
               "Bestehende Zeilen unveraendert: " & zeilenUnveraendert & vbCrLf & vbCrLf & _
               "Alle Zuordnungen sind vollstaendig (GRUEN).", vbInformation, "Aktualisierung abgeschlossen"
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
' HILFSPROZEDUR: Setzt Zellschutz basierend auf Role-Typ
' Parzelle (V) nur bearbeitbar fuer Mitglieder und SONSTIGE
' ===============================================================
Private Sub SetzeZellschutzFuerZeile(ByRef ws As Worksheet, ByVal zeile As Long, ByVal role As String)
    
    ' Spalten U, W, X immer bearbeitbar
    ws.Cells(zeile, EK_COL_ZUORDNUNG).Locked = False
    ws.Cells(zeile, EK_COL_ROLE).Locked = False
    ws.Cells(zeile, EK_COL_DEBUG).Locked = False
    
    ' Spalte V (Parzelle) nur bearbeitbar wenn erlaubt
    If DarfParzelleHaben(role) Then
        ws.Cells(zeile, EK_COL_PARZELLE).Locked = False
    Else
        ws.Cells(zeile, EK_COL_PARZELLE).Locked = True
        ws.Cells(zeile, EK_COL_PARZELLE).value = ""
    End If
    
End Sub

' ===============================================================
' HILFSFUNKTION: Prueft ob Zeile bereits gueltige Daten hat
' ===============================================================
Private Function HatBereitsGueltigeDaten(ByVal entityKey As String, _
                                          ByVal zuordnung As String, _
                                          ByVal role As String) As Boolean
    
    HatBereitsGueltigeDaten = False
    
    ' Wenn EntityKey vorhanden und KEINE einfache Zahl ist
    If entityKey <> "" Then
        If Not IsNumeric(entityKey) Then
            HatBereitsGueltigeDaten = True
            Exit Function
        End If
    End If
    
    ' Wenn Zuordnung UND Role vorhanden
    If zuordnung <> "" And role <> "" Then
        HatBereitsGueltigeDaten = True
        Exit Function
    End If
    
End Function

' ===============================================================
' HILFSPROZEDUR: Zeigt Hinweis fuer Zeilen mit erforderlichem Eingriff
' ===============================================================
Private Sub ZeigeEingriffsHinweis(ByRef ws As Worksheet, ByRef zeilenRot As Collection, _
                                   ByRef zeilenGelb As Collection, _
                                   ByVal zeilenNeu As Long, ByVal zeilenUnveraendert As Long)
    
    Dim msg As String
    Dim antwort As VbMsgBoxResult
    Dim ersteZeile As Long
    
    msg = "EntityKey-Aktualisierung abgeschlossen!" & vbCrLf & vbCrLf
    msg = msg & "Neue Zeilen verarbeitet: " & zeilenNeu & vbCrLf
    msg = msg & "Bestehende Zeilen unveraendert: " & zeilenUnveraendert & vbCrLf & vbCrLf
    
    If zeilenRot.Count > 0 Then
        msg = msg & "ROT: " & zeilenRot.Count & " Zeile(n) - Manuelle Zuordnung erforderlich!" & vbCrLf
    End If
    
    If zeilenGelb.Count > 0 Then
        msg = msg & "GELB: " & zeilenGelb.Count & " Zeile(n) - Nur Nachname gefunden, bitte pruefen!" & vbCrLf
    End If
    
    msg = msg & vbCrLf & "Moechten Sie jetzt zur ersten betroffenen Zeile springen?"
    
    antwort = MsgBox(msg, vbYesNo + vbExclamation, "Zuordnung pruefen")
    
    If antwort = vbYes Then
        If zeilenRot.Count > 0 Then
            ersteZeile = zeilenRot(1)
        ElseIf zeilenGelb.Count > 0 Then
            ersteZeile = zeilenGelb(1)
        Else
            Exit Sub
        End If
        
        ws.Activate
        ws.Cells(ersteZeile, EK_COL_ZUORDNUNG).Select
    End If
    
End Sub

' ===============================================================
' HILFSFUNKTION: Sucht Mitglieder anhand des Kontonamens
' ===============================================================
Private Function SucheMitgliederZuKontoname(ByVal kontoname As String, _
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
    Dim mitgliedInfo(0 To 8) As Variant
    Dim zeilen As Variant
    Dim zeile As Variant
    Dim nameKombiniert As String
    Dim nameParts() As String
    Dim austrittsDatum As Date
    Dim matchResult As Long
    
    Set SucheMitgliederZuKontoname = gefunden
    
    If kontoname = "" Then Exit Function
    
    zeilen = Split(kontoname, vbLf)
    
    For Each zeile In zeilen
        kontoNameNorm = NormalisiereStringFuerVergleich(CStr(zeile))
        If kontoNameNorm = "" Then GoTo NextZeile
        
        ' Aktive Mitglieder suchen
        lastRow = wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
        
        For r = M_START_ROW To lastRow
            ' Nur aktive Mitglieder (ohne Pachtende-Datum)
            If Trim(wsM.Cells(r, M_COL_PACHTENDE).value) = "" Then
                nachname = Trim(wsM.Cells(r, M_COL_NACHNAME).value)
                vorname = Trim(wsM.Cells(r, M_COL_VORNAME).value)
                memberID = Trim(wsM.Cells(r, M_COL_MEMBER_ID).value)
                parzelle = Trim(wsM.Cells(r, M_COL_PARZELLE).value)
                funktion = Trim(wsM.Cells(r, M_COL_FUNKTION).value)
                
                matchResult = PruefeNamensMatch(nachname, vorname, kontoNameNorm)
                
                If matchResult > 0 Then
                    If Not IstMitgliedBereitsGefunden(gefunden, memberID, False) Then
                        mitgliedInfo(0) = memberID
                        mitgliedInfo(1) = nachname
                        mitgliedInfo(2) = vorname
                        mitgliedInfo(3) = parzelle
                        mitgliedInfo(4) = funktion
                        mitgliedInfo(5) = r
                        mitgliedInfo(6) = False  ' Nicht ehemalig
                        mitgliedInfo(7) = CDate("01.01.1900")
                        mitgliedInfo(8) = matchResult
                        gefunden.Add mitgliedInfo
                    End If
                End If
            End If
        Next r
        
        ' Ehemalige Mitglieder suchen
        lastRow = wsH.Cells(wsH.Rows.Count, H_COL_NAME_EHEM_PAECHTER).End(xlUp).Row
        
        For r = H_START_ROW To lastRow
            nameKombiniert = Trim(wsH.Cells(r, H_COL_NAME_EHEM_PAECHTER).value)
            memberID = Trim(wsH.Cells(r, H_COL_MEMBER_ID_ALT).value)
            parzelle = Trim(wsH.Cells(r, H_COL_PARZELLE).value)
            
            On Error Resume Next
            austrittsDatum = wsH.Cells(r, H_COL_AUST_DATUM).value
            If Err.Number <> 0 Then austrittsDatum = CDate("01.01.1900")
            On Error GoTo 0
            
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
            
            matchResult = PruefeNamensMatch(nachname, vorname, kontoNameNorm)
            
            If matchResult > 0 Then
                If Not IstMitgliedBereitsGefunden(gefunden, memberID, True) Then
                    mitgliedInfo(0) = memberID
                    mitgliedInfo(1) = nachname
                    mitgliedInfo(2) = vorname
                    mitgliedInfo(3) = parzelle
                    mitgliedInfo(4) = "Ehemaliges Mitglied"
                    mitgliedInfo(5) = r
                    mitgliedInfo(6) = True  ' Ehemalig
                    mitgliedInfo(7) = austrittsDatum
                    mitgliedInfo(8) = matchResult
                    gefunden.Add mitgliedInfo
                End If
            End If
        Next r
        
NextZeile:
    Next zeile
    
    Set SucheMitgliederZuKontoname = gefunden
    
End Function

' ===============================================================
' HILFSFUNKTION: Prueft ob Name im Kontonamen enthalten ist
' Rueckgabe: 0=kein Match, 1=nur Nachname, 2=Vor- und Nachname
' ===============================================================
Private Function PruefeNamensMatch(ByVal nachname As String, ByVal vorname As String, _
                                    ByVal kontoNameNorm As String) As Long
    
    Dim nachnameNorm As String
    Dim vornameNorm As String
    Dim nachnameGefunden As Boolean
    Dim vornameGefunden As Boolean
    
    PruefeNamensMatch = 0
    
    nachnameNorm = NormalisiereStringFuerVergleich(nachname)
    vornameNorm = NormalisiereStringFuerVergleich(vorname)
    
    If nachnameNorm = "" Then Exit Function
    If Len(nachnameNorm) < 3 Then Exit Function
    
    nachnameGefunden = IstNachnameImKontoname(nachnameNorm, kontoNameNorm)
    
    If Not nachnameGefunden Then
        PruefeNamensMatch = 0
        Exit Function
    End If
    
    vornameGefunden = False
    
    If vornameNorm <> "" And Len(vornameNorm) >= 2 Then
        vornameGefunden = IstVornameImKontoname(vornameNorm, kontoNameNorm)
    End If
    
    If nachnameGefunden And vornameGefunden Then
        PruefeNamensMatch = 2  ' Exakter Treffer
    ElseIf nachnameGefunden Then
        PruefeNamensMatch = 1  ' Nur Nachname
    End If
    
End Function

' ===============================================================
' HILFSFUNKTION: Prueft ob Nachname im Kontonamen enthalten ist
' ===============================================================
Private Function IstNachnameImKontoname(ByVal nachnameNorm As String, _
                                         ByVal kontoNameNorm As String) As Boolean
    
    Dim teile() As String
    Dim teil As Variant
    
    IstNachnameImKontoname = False
    
    If InStr(kontoNameNorm, nachnameNorm) > 0 Then
        IstNachnameImKontoname = True
        Exit Function
    End If
    
    ' Mehrteilige Nachnamen pruefen
    If InStr(nachnameNorm, " ") > 0 Then
        teile = Split(nachnameNorm, " ")
        Dim alleGefunden As Boolean
        alleGefunden = True
        For Each teil In teile
            If Len(teil) >= 3 Then
                If InStr(kontoNameNorm, CStr(teil)) = 0 Then
                    alleGefunden = False
                    Exit For
                End If
            End If
        Next teil
        IstNachnameImKontoname = alleGefunden
    End If
    
End Function

' ===============================================================
' HILFSFUNKTION: Prueft ob Vorname im Kontonamen enthalten ist
' ===============================================================
Private Function IstVornameImKontoname(ByVal vornameNorm As String, _
                                        ByVal kontoNameNorm As String) As Boolean
    
    Dim pos As Long
    Dim vorZeichen As String
    Dim nachZeichen As String
    
    IstVornameImKontoname = False
    
    If vornameNorm = "" Then Exit Function
    
    pos = InStr(kontoNameNorm, vornameNorm)
    
    If pos > 0 Then
        If pos > 1 Then
            vorZeichen = Mid(kontoNameNorm, pos - 1, 1)
        Else
            vorZeichen = " "
        End If
        
        If pos + Len(vornameNorm) <= Len(kontoNameNorm) Then
            nachZeichen = Mid(kontoNameNorm, pos + Len(vornameNorm), 1)
        Else
            nachZeichen = " "
        End If
        
        If (vorZeichen = " " Or pos = 1) And (nachZeichen = " " Or pos + Len(vornameNorm) > Len(kontoNameNorm)) Then
            IstVornameImKontoname = True
        End If
    End If
    
End Function

' ===============================================================
' HILFSFUNKTION: Normalisiert String fuer Vergleich
' ===============================================================
Private Function NormalisiereStringFuerVergleich(ByVal s As String) As String
    Dim result As String
    
    result = LCase(Trim(s))
    result = Replace(result, ",", " ")
    result = Replace(result, ".", " ")
    result = Replace(result, "-", " ")
    
    ' Umlaute normalisieren
    result = Replace(result, "ae", "a")
    result = Replace(result, "oe", "o")
    result = Replace(result, "ue", "u")
    result = Replace(result, "ss", "s")
    
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    result = Trim(result)
    
    NormalisiereStringFuerVergleich = result
End Function

' ===============================================================
' HILFSFUNKTION: Prueft ob MemberID bereits in Collection ist
' ===============================================================
Private Function IstMitgliedBereitsGefunden(ByRef col As Collection, _
                                             ByVal memberID As String, _
                                             ByVal istEhemalig As Boolean) As Boolean
    Dim item As Variant
    
    IstMitgliedBereitsGefunden = False
    
    For Each item In col
        If item(0) = memberID And item(6) = istEhemalig Then
            IstMitgliedBereitsGefunden = True
            Exit Function
        End If
    Next item
End Function

' ===============================================================
' HILFSPROZEDUR: Generiert EntityKey und Zuordnung
' WICHTIG: Nutzt Member-ID (GUID) fuer Mitglieder!
' ===============================================================
Private Sub GeneriereEntityKeyUndZuordnung(ByRef mitglieder As Collection, _
                                            ByVal kontoname As String, _
                                            ByRef wsM As Worksheet, _
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
    Dim key As Variant
    
    Dim mitgliederExakt As New Collection
    Dim mitgliederNurNachname As New Collection
    
    Set uniqueMemberIDs = CreateObject("Scripting.Dictionary")
    
    ' Initialisierung
    outEntityKey = ""
    outZuordnung = ""
    outParzellen = ""
    outEntityRole = ""
    outDebugInfo = ""
    outAmpelStatus = 1  ' Gruen = OK
    
    ' Treffer nach Match-Qualitaet sortieren
    For i = 1 To mitglieder.Count
        mitgliedInfo = mitglieder(i)
        
        If mitgliedInfo(8) = 2 Then
            mitgliederExakt.Add mitgliedInfo
        ElseIf mitgliedInfo(8) = 1 Then
            mitgliederNurNachname.Add mitgliedInfo
        End If
    Next i
    
    ' ============================================================
    ' Fall 1: Keine exakten Treffer gefunden
    ' ============================================================
    If mitgliederExakt.Count = 0 Then
        ' Pruefen ob VERSORGER, BANK oder SHOP
        If IstVersorger(kontoname) Then
            outEntityKey = PREFIX_VERSORGER & CreateGUID()
            outEntityRole = ROLE_VERSORGER
            outZuordnung = ExtrahiereAnzeigeName(kontoname)
            outParzellen = ""
            outDebugInfo = "Automatisch als VERSORGER erkannt"
            outAmpelStatus = 1
            Exit Sub
        ElseIf IstBank(kontoname) Then
            outEntityKey = PREFIX_BANK & CreateGUID()
            outEntityRole = ROLE_BANK
            outZuordnung = ExtrahiereAnzeigeName(kontoname)
            outParzellen = ""
            outDebugInfo = "Automatisch als BANK erkannt"
            outAmpelStatus = 1
            Exit Sub
        ElseIf IstShop(kontoname) Then
            outEntityKey = PREFIX_SHOP & CreateGUID()
            outEntityRole = ROLE_SHOP
            outZuordnung = ExtrahiereAnzeigeName(kontoname)
            outParzellen = ""
            outDebugInfo = "Automatisch als SHOP erkannt"
            outAmpelStatus = 1
            Exit Sub
        End If
        
        ' Nur Nachname gefunden?
        If mitgliederNurNachname.Count > 0 Then
            outEntityKey = ""
            outZuordnung = ""
            outParzellen = ""
            outEntityRole = ""
            outDebugInfo = "NUR NACHNAME GEFUNDEN - Bitte pruefen! Moegliche Treffer:"
            outAmpelStatus = 2  ' Gelb
            
            For i = 1 To mitgliederNurNachname.Count
                mitgliedInfo = mitgliederNurNachname(i)
                outDebugInfo = outDebugInfo & vbLf & "  ? " & mitgliedInfo(1) & ", " & mitgliedInfo(2) & " (Parz. " & mitgliedInfo(3) & ")"
            Next i
            Exit Sub
        Else
            ' Gar nichts gefunden
            outEntityKey = ""
            outZuordnung = ""
            outParzellen = ""
            outEntityRole = ""
            outDebugInfo = "KEIN MITGLIED GEFUNDEN - Manuelle Zuordnung erforderlich"
            outAmpelStatus = 3  ' Rot
            Exit Sub
        End If
    End If
    
    ' ============================================================
    ' Fall 2: Exakte Treffer vorhanden
    ' ============================================================
    hatAktiveMitglieder = False
    hatEhemaligeMitglieder = False
    
    For i = 1 To mitgliederExakt.Count
        mitgliedInfo = mitgliederExakt(i)
        
        If mitgliedInfo(6) = False Then
            hatAktiveMitglieder = True
            If Not uniqueMemberIDs.Exists(CStr(mitgliedInfo(0))) Then
                uniqueMemberIDs.Add CStr(mitgliedInfo(0)), CStr(mitgliedInfo(0))
            End If
        Else
            hatEhemaligeMitglieder = True
        End If
    Next i
    
    ' Fall 2a: NUR ehemalige Mitglieder
    If hatEhemaligeMitglieder And Not hatAktiveMitglieder Then
        Set uniqueMemberIDs = CreateObject("Scripting.Dictionary")
        For i = 1 To mitgliederExakt.Count
            mitgliedInfo = mitgliederExakt(i)
            If Not uniqueMemberIDs.Exists(CStr(mitgliedInfo(0))) Then
                uniqueMemberIDs.Add CStr(mitgliedInfo(0)), CStr(mitgliedInfo(0))
            End If
        Next i
        
        If uniqueMemberIDs.Count = 1 Then
            ' Ein ehemaliges Mitglied
            mitgliedInfo = mitgliederExakt(1)
            outEntityKey = PREFIX_EHEMALIG & CStr(mitgliedInfo(0))
            outZuordnung = mitgliedInfo(1) & ", " & mitgliedInfo(2)
            outParzellen = mitgliedInfo(3) & " (bis " & Format(mitgliedInfo(7), "dd.mm.yyyy") & ")"
            outEntityRole = ROLE_EHEMALIGES_MITGLIED
            outDebugInfo = "Ehemaliges Mitglied - exakter Treffer"
            outAmpelStatus = 1
        Else
            ' Mehrere ehemalige = Gemeinschaftskonto
            memberIDs = ""
            For Each key In uniqueMemberIDs.Keys
                If memberIDs <> "" Then memberIDs = memberIDs & "_"
                memberIDs = memberIDs & key
            Next key
            
            outEntityKey = PREFIX_SHARE & PREFIX_EHEMALIG & memberIDs
            outEntityRole = ROLE_EHEMALIGES_MITGLIED
            outDebugInfo = "Ehem. Gemeinschaftskonto - " & uniqueMemberIDs.Count & " Personen"
            outAmpelStatus = 1
            
            Dim bereitsHinzu As Object
            Set bereitsHinzu = CreateObject("Scripting.Dictionary")
            
            For i = 1 To mitgliederExakt.Count
                mitgliedInfo = mitgliederExakt(i)
                If Not bereitsHinzu.Exists(CStr(mitgliedInfo(0))) Then
                    bereitsHinzu.Add CStr(mitgliedInfo(0)), True
                    If outZuordnung <> "" Then outZuordnung = outZuordnung & vbLf
                    outZuordnung = outZuordnung & mitgliedInfo(1) & ", " & mitgliedInfo(2) & " (ehem.)"
                    
                    If InStr(outParzellen, CStr(mitgliedInfo(3))) = 0 Then
                        If outParzellen <> "" Then outParzellen = outParzellen & vbLf
                        outParzellen = outParzellen & CStr(mitgliedInfo(3))
                    End If
                End If
            Next i
        End If
        Exit Sub
    End If
    
    ' Fall 2b: Aktive Mitglieder - genau 1
    If uniqueMemberIDs.Count = 1 Then
        For i = 1 To mitgliederExakt.Count
            mitgliedInfo = mitgliederExakt(i)
            If mitgliedInfo(6) = False Then
                ' WICHTIG: Member-ID (GUID) als EntityKey verwenden!
                outEntityKey = CStr(mitgliedInfo(0))
                outZuordnung = mitgliedInfo(1) & ", " & mitgliedInfo(2)
                outEntityRole = ErmittleEntityRoleVonFunktion(CStr(mitgliedInfo(4)))
                outDebugInfo = "Eindeutiger Treffer (Vor- und Nachname)"
                outAmpelStatus = 1
                Exit For
            End If
        Next i
        
        outParzellen = HoleAlleParzellen(CStr(mitgliedInfo(0)), wsM)
        
        If hatEhemaligeMitglieder Then
            outDebugInfo = outDebugInfo & " (+ ehem. Eintraege in Historie)"
        End If
        
        Exit Sub
    End If
    
    ' Fall 2c: Mehrere aktive Mitglieder = Gemeinschaftskonto
    If uniqueMemberIDs.Count > 1 Then
        memberIDs = ""
        For Each key In uniqueMemberIDs.Keys
            If memberIDs <> "" Then memberIDs = memberIDs & "_"
            memberIDs = memberIDs & key
        Next key
        
        outEntityKey = PREFIX_SHARE & memberIDs
        outEntityRole = ROLE_MITGLIED_MIT_PACHT
        outDebugInfo = "Gemeinschaftskonto - " & uniqueMemberIDs.Count & " Personen"
        outAmpelStatus = 1
        
        Dim bereitsHinzugefuegteMitglieder As Object
        Set bereitsHinzugefuegteMitglieder = CreateObject("Scripting.Dictionary")
        
        For i = 1 To mitgliederExakt.Count
            mitgliedInfo = mitgliederExakt(i)
            If mitgliedInfo(6) = False Then
                If Not bereitsHinzugefuegteMitglieder.Exists(CStr(mitgliedInfo(0))) Then
                    bereitsHinzugefuegteMitglieder.Add CStr(mitgliedInfo(0)), True
                    
                    If outZuordnung <> "" Then outZuordnung = outZuordnung & vbLf
                    outZuordnung = outZuordnung & mitgliedInfo(1) & ", " & mitgliedInfo(2)
                    
                    Dim dieseParzellen As String
                    dieseParzellen = HoleAlleParzellen(CStr(mitgliedInfo(0)), wsM)
                    
                    Dim parzArr() As String
                    Dim p As Long
                    parzArr = Split(dieseParzellen, vbLf)
                    For p = LBound(parzArr) To UBound(parzArr)
                        If Trim(parzArr(p)) <> "" Then
                            If InStr(outParzellen, Trim(parzArr(p))) = 0 Then
                                If outParzellen <> "" Then outParzellen = outParzellen & vbLf
                                outParzellen = outParzellen & Trim(parzArr(p))
                            End If
                        End If
                    Next p
                End If
            End If
        Next i
    End If
    
End Sub


' ---------------------------------------------------------
' Hilfsfunktion: Extrahiert Anzeigename aus Zuordnungstext
' ---------------------------------------------------------
Private Function ExtrahiereAnzeigeName(ByVal zuordnung As String) As String
    Dim pos As Long
    
    If Len(zuordnung) = 0 Then
        ExtrahiereAnzeigeName = ""
        Exit Function
    End If
    
    ' Format: "Nachname, Vorname (Parzelle X)" oder "Name (Info)"
    pos = InStr(zuordnung, " (")
    If pos > 0 Then
        ExtrahiereAnzeigeName = Left(zuordnung, pos - 1)
    Else
        ExtrahiereAnzeigeName = zuordnung
    End If
End Function

' ---------------------------------------------------------
' Hilfsfunktion: Holt alle Parzellen eines Mitglieds
' ---------------------------------------------------------
Private Function HoleAlleParzellen(ByVal mitgliedID As String) As String
    Dim wsMitglieder As Worksheet
    Dim lastRow As Long, i As Long
    Dim parzellen As String
    Dim currentID As String, currentParzelle As String
    
    On Error Resume Next
    Set wsMitglieder = ThisWorkbook.Worksheets(WS_NAME_MITGLIEDER)
    On Error GoTo 0
    
    If wsMitglieder Is Nothing Then
        HoleAlleParzellen = ""
        Exit Function
    End If
    
    lastRow = wsMitglieder.Cells(wsMitglieder.Rows.Count, MIT_COL_MEMBER_ID).End(xlUp).Row
    parzellen = ""
    
    For i = MIT_FIRST_DATA_ROW To lastRow
        currentID = Trim(CStr(wsMitglieder.Cells(i, MIT_COL_MEMBER_ID).value))
        If currentID = mitgliedID Then
            currentParzelle = Trim(CStr(wsMitglieder.Cells(i, MIT_COL_PARZELLE).value))
            If Len(currentParzelle) > 0 Then
                If Len(parzellen) > 0 Then
                    ' Prüfen ob Parzelle schon enthalten
                    If InStr(parzellen, currentParzelle) = 0 Then
                        parzellen = parzellen & ", " & currentParzelle
                    End If
                Else
                    parzellen = currentParzelle
                End If
            End If
        End If
    Next i
    
    HoleAlleParzellen = parzellen
End Function

' ---------------------------------------------------------
' Aktualisiert Parzellen für ein Mitglied in EntityKey-Tabelle
' ---------------------------------------------------------
Public Sub AktualisiereParzellenFuerMitglied(ByVal mitgliedID As String)
    Dim wsDaten As Worksheet
    Dim lastRow As Long, i As Long
    Dim currentKey As String
    Dim neueParzellen As String
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_NAME_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    
    lastRow = wsDaten.Cells(wsDaten.Rows.Count, EK_COL_ENTITYKEY).End(xlUp).Row
    neueParzellen = HoleAlleParzellen(mitgliedID)
    
    For i = EK_FIRST_DATA_ROW To lastRow
        currentKey = Trim(CStr(wsDaten.Cells(i, EK_COL_ENTITYKEY).value))
        If currentKey = mitgliedID Then
            wsDaten.Cells(i, EK_COL_PARZELLE).value = neueParzellen
        End If
    Next i
End Sub

' ---------------------------------------------------------
' Ermittelt EntityRole basierend auf Vereinsfunktion
' ---------------------------------------------------------
Private Function ErmittleEntityRoleVonFunktion(ByVal funktion As String) As String
    Dim func As String
    func = UCase(Trim(funktion))
    
    Select Case func
        Case "1. VORSITZENDER", "2. VORSITZENDER", "KASSENWART", _
             "SCHRIFTFÜHRER", "BEISITZER", "FACHBERATER"
            ErmittleEntityRoleVonFunktion = ROLE_MITGLIED_MIT_PACHT
        Case "EHRENMITGLIED"
            ErmittleEntityRoleVonFunktion = ROLE_MITGLIED_OHNE_PACHT
        Case Else
            ' Standard: Mitglied mit Pacht
            ErmittleEntityRoleVonFunktion = ROLE_MITGLIED_MIT_PACHT
    End Select
End Function

' ---------------------------------------------------------
' Prüft ob Kontoname auf Versorger hindeutet
' ---------------------------------------------------------
Private Function IstVersorger(ByVal kontoname As String) As Boolean
    Dim name As String
    name = UCase(kontoname)
    
    ' Strom/Energie
    If InStr(name, "ENERGIE") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "STROM") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "STADTWERKE") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "ENVIAM") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "ENVIA") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "MITNETZ") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "VATTENFALL") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "E.ON") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "EON") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "RWE") > 0 Then IstVersorger = True: Exit Function
    
    ' Wasser
    If InStr(name, "WASSER") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "WASSERWERK") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "ZWA") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "ABWASSER") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "ZWECKVERBAND") > 0 Then IstVersorger = True: Exit Function
    
    ' Gas
    If InStr(name, "GASAG") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "GASVERSORGUNG") > 0 Then IstVersorger = True: Exit Function
    
    ' Versicherungen
    If InStr(name, "VERSICHERUNG") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "ALLIANZ") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "ERGO") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "HDI") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "AXA") > 0 Then IstVersorger = True: Exit Function
    
    ' Kommunale Dienste
    If InStr(name, "STADTVERWALTUNG") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "FINANZAMT") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "GEMEINDE") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "LANDKREIS") > 0 Then IstVersorger = True: Exit Function
    
    ' Verbände
    If InStr(name, "VERBAND") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "LANDESVERBAND") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "BUNDESVERBAND") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "KREISVERBAND") > 0 Then IstVersorger = True: Exit Function
    
    ' Telekommunikation
    If InStr(name, "TELEKOM") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "VODAFONE") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "O2") > 0 Then IstVersorger = True: Exit Function
    
    ' Müll/Entsorgung
    If InStr(name, "ENTSORGUNG") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "ABFALL") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "MÜLL") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "WERTSTOFF") > 0 Then IstVersorger = True: Exit Function
    
    IstVersorger = False
End Function

' ---------------------------------------------------------
' Prüft ob Kontoname auf Bank hindeutet
' ---------------------------------------------------------
Private Function IstBank(ByVal kontoname As String) As Boolean
    Dim name As String
    name = UCase(kontoname)
    
    If InStr(name, "SPARKASSE") > 0 Then IstBank = True: Exit Function
    If InStr(name, "SPARDA") > 0 Then IstBank = True: Exit Function
    If InStr(name, "VOLKSBANK") > 0 Then IstBank = True: Exit Function
    If InStr(name, "RAIFFEISEN") > 0 Then IstBank = True: Exit Function
    If InStr(name, "COMMERZBANK") > 0 Then IstBank = True: Exit Function
    If InStr(name, "DEUTSCHE BANK") > 0 Then IstBank = True: Exit Function
    If InStr(name, "POSTBANK") > 0 Then IstBank = True: Exit Function
    If InStr(name, "HYPOVEREINSBANK") > 0 Then IstBank = True: Exit Function
    If InStr(name, "TARGOBANK") > 0 Then IstBank = True: Exit Function
    If InStr(name, "ING") > 0 Then IstBank = True: Exit Function
    If InStr(name, "COMDIRECT") > 0 Then IstBank = True: Exit Function
    If InStr(name, "DKB") > 0 Then IstBank = True: Exit Function
    If InStr(name, "KREDITINSTITUT") > 0 Then IstBank = True: Exit Function
    If InStr(name, "GENOSSENSCHAFTSBANK") > 0 Then IstBank = True: Exit Function
    If InStr(name, "LANDESBANK") > 0 Then IstBank = True: Exit Function
    If InStr(name, "N26") > 0 Then IstBank = True: Exit Function
    If InStr(name, "SANTANDER") > 0 Then IstBank = True: Exit Function
    If InStr(name, "KREDITKARTE") > 0 Then IstBank = True: Exit Function
    If InStr(name, "ZINSEN") > 0 Then IstBank = True: Exit Function
    If InStr(name, "KONTOFÜHRUNG") > 0 Then IstBank = True: Exit Function
    If InStr(name, "BANKGEBÜHR") > 0 Then IstBank = True: Exit Function
    
    IstBank = False
End Function

' ---------------------------------------------------------
' Prüft ob Kontoname auf Shop/Händler hindeutet
' ---------------------------------------------------------
Private Function IstShop(ByVal kontoname As String) As Boolean
    Dim name As String
    name = UCase(kontoname)
    
    ' Baumärkte
    If InStr(name, "BAUHAUS") > 0 Then IstShop = True: Exit Function
    If InStr(name, "OBI") > 0 Then IstShop = True: Exit Function
    If InStr(name, "HORNBACH") > 0 Then IstShop = True: Exit Function
    If InStr(name, "HAGEBAU") > 0 Then IstShop = True: Exit Function
    If InStr(name, "TOOM") > 0 Then IstShop = True: Exit Function
    If InStr(name, "BAUMARKT") > 0 Then IstShop = True: Exit Function
    If InStr(name, "HELLWEG") > 0 Then IstShop = True: Exit Function
    If InStr(name, "GLOBUS") > 0 Then IstShop = True: Exit Function
    
    ' Gartencenter
    If InStr(name, "GARTENCENTER") > 0 Then IstShop = True: Exit Function
    If InStr(name, "DEHNER") > 0 Then IstShop = True: Exit Function
    If InStr(name, "KÖLLE") > 0 Then IstShop = True: Exit Function
    If InStr(name, "PFLANZENCENTER") > 0 Then IstShop = True: Exit Function
    
    ' Elektro
    If InStr(name, "MEDIAMARKT") > 0 Then IstShop = True: Exit Function
    If InStr(name, "MEDIA MARKT") > 0 Then IstShop = True: Exit Function
    If InStr(name, "SATURN") > 0 Then IstShop = True: Exit Function
    If InStr(name, "CONRAD") > 0 Then IstShop = True: Exit Function
    If InStr(name, "EURONICS") > 0 Then IstShop = True: Exit Function
    If InStr(name, "EXPERT") > 0 Then IstShop = True: Exit Function
    
    ' Möbel
    If InStr(name, "IKEA") > 0 Then IstShop = True: Exit Function
    If InStr(name, "POCO") > 0 Then IstShop = True: Exit Function
    If InStr(name, "ROLLER") > 0 Then IstShop = True: Exit Function
    If InStr(name, "MÖBEL") > 0 Then IstShop = True: Exit Function
    
    ' Online-Shops
    If InStr(name, "AMAZON") > 0 Then IstShop = True: Exit Function
    If InStr(name, "EBAY") > 0 Then IstShop = True: Exit Function
    If InStr(name, "OTTO") > 0 Then IstShop = True: Exit Function
    If InStr(name, "PAYPAL") > 0 Then IstShop = True: Exit Function
    
    ' Bürobedarf
    If InStr(name, "STAPLES") > 0 Then IstShop = True: Exit Function
    If InStr(name, "BÜRO") > 0 Then IstShop = True: Exit Function
    If InStr(name, "OFFICE") > 0 Then IstShop = True: Exit Function
    
    ' Supermärkte (für Vereinsbedarf)
    If InStr(name, "REWE") > 0 Then IstShop = True: Exit Function
    If InStr(name, "EDEKA") > 0 Then IstShop = True: Exit Function
    If InStr(name, "LIDL") > 0 Then IstShop = True: Exit Function
    If InStr(name, "ALDI") > 0 Then IstShop = True: Exit Function
    If InStr(name, "KAUFLAND") > 0 Then IstShop = True: Exit Function
    If InStr(name, "NETTO") > 0 Then IstShop = True: Exit Function
    If InStr(name, "PENNY") > 0 Then IstShop = True: Exit Function
    
    ' Getränke
    If InStr(name, "GETRÄNKE") > 0 Then IstShop = True: Exit Function
    If InStr(name, "GETRAENKE") > 0 Then IstShop = True: Exit Function
    
    IstShop = False
End Function

' ---------------------------------------------------------
' Erzeugt eine neue GUID - nutzt mod_Mitglieder_UI.CreateGUID_Public
' ---------------------------------------------------------
Private Function CreateGUID() As String
    On Error Resume Next
    CreateGUID = mod_Mitglieder_UI.CreateGUID_Public()
    If Err.Number <> 0 Or Len(CreateGUID) = 0 Then
        ' Fallback: Eigene GUID generieren
        Randomize
        CreateGUID = Format(Now, "YYYYMMDDHHMMSS") & "-" & _
                     Format(Int(Rnd * 10000), "0000") & "-" & _
                     Format(Int(Rnd * 10000), "0000")
    End If
    On Error GoTo 0
End Function

' ---------------------------------------------------------
' Setzt Ampelfarbe in der Debug-Spalte
' ---------------------------------------------------------
Private Sub SetzeAmpelFarbe(ByVal rng As Range, ByVal statusCode As String)
    With rng
        .value = statusCode
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        
        Select Case UCase(statusCode)
            Case "OK", "GRÜN", "GREEN"
                .Interior.color = RGB(146, 208, 80)  ' Grün
                .Font.color = RGB(0, 0, 0)
            Case "WARN", "GELB", "YELLOW"
                .Interior.color = RGB(255, 255, 0)   ' Gelb
                .Font.color = RGB(0, 0, 0)
            Case "ERR", "ROT", "RED"
                .Interior.color = RGB(255, 0, 0)     ' Rot
                .Font.color = RGB(255, 255, 255)
            Case "INFO", "BLAU", "BLUE"
                .Interior.color = RGB(91, 155, 213)  ' Blau
                .Font.color = RGB(255, 255, 255)
            Case Else
                .Interior.ColorIndex = xlNone
                .Font.color = RGB(0, 0, 0)
        End Select
    End With
End Sub

' ---------------------------------------------------------
' Richtet das EntityRole-DropDown ein - DYNAMISCH auf Spalte AD!
' Parameter: ws = Worksheet, lastRow = letzte Datenzeile
' ---------------------------------------------------------
Private Sub SetupEntityRoleDropdown(ByRef ws As Worksheet, ByVal lastRow As Long)
    Dim wsDaten As Worksheet
    Dim lastRoleRow As Long
    Dim sourceRange As String
    Dim i As Long
    Dim zielZelle As Range
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_NAME_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    
    ' Dynamische Ermittlung der letzten Zeile in der EntityRole-Füllbereich-Spalte (AD)
    lastRoleRow = wsDaten.Cells(wsDaten.Rows.Count, DATA_COL_DD_ENTITYROLE).End(xlUp).Row
    If lastRoleRow < 4 Then lastRoleRow = 10  ' Fallback
    
    ' Quelle: Daten!$AD$4:$AD$[lastRow] - DYNAMISCH!
    sourceRange = "=" & WS_NAME_DATEN & "!$" & ColLetter(DATA_COL_DD_ENTITYROLE) & "$4:$" & _
                  ColLetter(DATA_COL_DD_ENTITYROLE) & "$" & lastRoleRow
    
    ' Für jede Datenzeile das DropDown einrichten
    For i = EK_FIRST_DATA_ROW To lastRow
        Set zielZelle = ws.Cells(i, EK_COL_ROLE)
        
        ' Bestehende Validierung entfernen
        On Error Resume Next
        zielZelle.Validation.Delete
        On Error GoTo 0
        
        ' Neue DropDown-Liste erstellen
        With zielZelle.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:=sourceRange
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = True
            .ErrorTitle = "Ungültige Eingabe"
            .ErrorMessage = "Bitte wählen Sie eine gültige EntityRole aus der Liste."
        End With
    Next i
End Sub
' ---------------------------------------------------------
' Hilfsfunktion: Spaltenummer zu Buchstabe
' ---------------------------------------------------------
Private Function ColLetter(ByVal colNum As Long) As String
    Dim temp As String
    
    If colNum < 1 Then
        ColLetter = ""
        Exit Function
    End If
    
    temp = ""
    Do While colNum > 0
        temp = Chr(((colNum - 1) Mod 26) + 65) & temp
        colNum = (colNum - 1) \ 26
    Loop
    
    ColLetter = temp
End Function

' ---------------------------------------------------------
' Richtet das Parzellen-DropDown ein (nur für Mitglieder/Sonstige)
' ---------------------------------------------------------
Public Sub SetupParzellenDropdown(ByVal zielZelle As Range)
    Dim wsDaten As Worksheet
    Dim lastParzelleRow As Long
    Dim sourceRange As String
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_NAME_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    If zielZelle Is Nothing Then Exit Sub
    
    ' Dynamische Ermittlung der letzten Zeile in der Parzellen-Füllbereich-Spalte (F)
    lastParzelleRow = wsDaten.Cells(wsDaten.Rows.Count, DATA_COL_DD_PARZELLE).End(xlUp).Row
    If lastParzelleRow < 4 Then lastParzelleRow = 100  ' Fallback
    
    ' Quelle: Daten!$F$4:$F$[lastRow] - DYNAMISCH!
    sourceRange = "=" & WS_NAME_DATEN & "!$" & ColLetter(DATA_COL_DD_PARZELLE) & "$4:$" & _
                  ColLetter(DATA_COL_DD_PARZELLE) & "$" & lastParzelleRow
    
    ' Bestehende Validierung entfernen
    On Error Resume Next
    zielZelle.Validation.Delete
    On Error GoTo 0
    
    ' Neue DropDown-Liste erstellen
    With zielZelle.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, _
             Operator:=xlBetween, Formula1:=sourceRange
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
        .ErrorTitle = "Hinweis"
        .ErrorMessage = "Sie können eine Parzelle aus der Liste wählen oder manuell eingeben."
    End With
End Sub


' ---------------------------------------------------------
' Sperrt/Entsperrt die Parzellen-Zelle basierend auf EntityRole
' ---------------------------------------------------------
Public Sub AktualisiereParzellenschutz(ByVal zeile As Long)
    Dim wsDaten As Worksheet
    Dim roleValue As String
    Dim parzelleZelle As Range
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_NAME_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    
    Set parzelleZelle = wsDaten.Cells(zeile, EK_COL_PARZELLE)
    roleValue = Trim(CStr(wsDaten.Cells(zeile, EK_COL_ROLE).value))
    
    If DarfParzelleHaben(roleValue) Then
        ' Parzelle erlaubt - DropDown einrichten
        SetupParzellenDropdown parzelleZelle
        parzelleZelle.Interior.ColorIndex = xlNone
    Else
        ' Parzelle nicht erlaubt - Zelle leeren und grau hinterlegen
        parzelleZelle.value = ""
        On Error Resume Next
        parzelleZelle.Validation.Delete
        On Error GoTo 0
        parzelleZelle.Interior.color = RGB(217, 217, 217)  ' Hellgrau
    End If
End Sub

' ---------------------------------------------------------
' Formatiert die gesamte EntityKey-Tabelle
' ---------------------------------------------------------
Public Sub FormatiereEntityKeyTabelle()
    Dim wsDaten As Worksheet
    Dim lastRow As Long, i As Long
    Dim headerRange As Range
    Dim dataRange As Range
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_NAME_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then
        Debug.Print "FormatiereEntityKeyTabelle: Worksheet '" & WS_NAME_DATEN & "' nicht gefunden!"
        Exit Sub
    End If
    
    lastRow = wsDaten.Cells(wsDaten.Rows.Count, EK_COL_ENTITYKEY).End(xlUp).Row
    If lastRow < EK_FIRST_DATA_ROW Then lastRow = EK_FIRST_DATA_ROW
    
    Application.ScreenUpdating = False
    
    ' Header formatieren (Zeile 3)
    Set headerRange = wsDaten.Range(wsDaten.Cells(3, EK_COL_ENTITYKEY), _
                                     wsDaten.Cells(3, EK_COL_DEBUG))
    With headerRange
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.color = RGB(68, 114, 196)  ' Blau
        .Font.color = RGB(255, 255, 255)     ' Weiß
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    ' Datenbereich formatieren
    Set dataRange = wsDaten.Range(wsDaten.Cells(EK_FIRST_DATA_ROW, EK_COL_ENTITYKEY), _
                                   wsDaten.Cells(lastRow, EK_COL_DEBUG))
    
    With dataRange
        ' Vertikale Ausrichtung: OBEN (nicht Mitte!)
        .VerticalAlignment = xlTop
        
        ' Rahmen
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideVertical).color = RGB(191, 191, 191)
        
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideHorizontal).color = RGB(217, 217, 217)
        
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    
    ' Spaltenbreiten und Ausrichtung
    ' EntityKey (R) - links, AutoFit
    wsDaten.Columns(EK_COL_ENTITYKEY).HorizontalAlignment = xlLeft
    wsDaten.Columns(EK_COL_ENTITYKEY).EntireColumn.AutoFit
    
    ' IBAN (S) - links, feste Breite
    wsDaten.Columns(EK_COL_IBAN).HorizontalAlignment = xlLeft
    wsDaten.Columns(EK_COL_IBAN).ColumnWidth = 28
    
    ' Zahler/Empfänger (T) - links, AutoFit
    wsDaten.Columns(EK_COL_ZAHLER).HorizontalAlignment = xlLeft
    wsDaten.Columns(EK_COL_ZAHLER).EntireColumn.AutoFit
    If wsDaten.Columns(EK_COL_ZAHLER).ColumnWidth > 35 Then
        wsDaten.Columns(EK_COL_ZAHLER).ColumnWidth = 35
    End If
    
    ' Zuordnung (U) - links, AutoFit
    wsDaten.Columns(EK_COL_ZUORDNUNG).HorizontalAlignment = xlLeft
    wsDaten.Columns(EK_COL_ZUORDNUNG).EntireColumn.AutoFit
    If wsDaten.Columns(EK_COL_ZUORDNUNG).ColumnWidth > 40 Then
        wsDaten.Columns(EK_COL_ZUORDNUNG).ColumnWidth = 40
    End If
    
    ' Parzelle (V) - zentriert
    wsDaten.Columns(EK_COL_PARZELLE).HorizontalAlignment = xlCenter
    wsDaten.Columns(EK_COL_PARZELLE).ColumnWidth = 10
    
    ' EntityRole (W) - LINKS, AutoFit
    wsDaten.Columns(EK_COL_ROLE).HorizontalAlignment = xlLeft
    wsDaten.Columns(EK_COL_ROLE).EntireColumn.AutoFit
    If wsDaten.Columns(EK_COL_ROLE).ColumnWidth < 20 Then
        wsDaten.Columns(EK_COL_ROLE).ColumnWidth = 20
    End If
    
    ' Debug (X) - zentriert
    wsDaten.Columns(EK_COL_DEBUG).HorizontalAlignment = xlCenter
    wsDaten.Columns(EK_COL_DEBUG).ColumnWidth = 8
    
    ' Einzelne Zeilen formatieren (DropDowns, Parzellenschutz)
    For i = EK_FIRST_DATA_ROW To lastRow
        FormatiereEntityKeyZeile i, False  ' Ohne Rahmen (bereits gemacht)
    Next i
    
    Application.ScreenUpdating = True
    
    Debug.Print "FormatiereEntityKeyTabelle: " & (lastRow - EK_FIRST_DATA_ROW + 1) & " Zeilen formatiert"
End Sub

' ---------------------------------------------------------
' Formatiert eine einzelne EntityKey-Zeile
' ---------------------------------------------------------
Public Sub FormatiereEntityKeyZeile(ByVal zeile As Long, Optional ByVal mitRahmen As Boolean = True)
    Dim wsDaten As Worksheet
    Dim roleValue As String
    Dim zeilenRange As Range
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_NAME_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    
    Set zeilenRange = wsDaten.Range(wsDaten.Cells(zeile, EK_COL_ENTITYKEY), _
                                     wsDaten.Cells(zeile, EK_COL_DEBUG))
    
    ' Vertikale Ausrichtung: OBEN
    zeilenRange.VerticalAlignment = xlTop
    
    ' Rahmen wenn gewünscht
    If mitRahmen Then
        With zeilenRange
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideVertical).Weight = xlThin
            .Borders(xlInsideVertical).color = RGB(191, 191, 191)
        End With
    End If
    
    ' EntityRole DropDown einrichten
    SetupEntityRoleDropdown wsDaten.Cells(zeile, EK_COL_ROLE)
    
    ' Parzellenschutz aktualisieren
    AktualisiereParzellenschutz zeile
    
    ' Spalte W (EntityRole) linksbündig
    wsDaten.Cells(zeile, EK_COL_ROLE).HorizontalAlignment = xlLeft
End Sub

' ---------------------------------------------------------
' Normalisiert Role-String für Vergleiche
' ---------------------------------------------------------
Private Function NormalisiereRoleString(ByVal roleString As String) As String
    Dim result As String
    result = UCase(Trim(roleString))
    result = Replace(result, " ", "_")
    result = Replace(result, "-", "_")
    NormalisiereRoleString = result
End Function

' ---------------------------------------------------------
' Prüft ob Role auf ehemaliges Mitglied hindeutet
' ---------------------------------------------------------
Private Function IstRoleEhemaligesMitglied(ByVal roleString As String) As Boolean
    Dim role As String
    role = NormalisiereRoleString(roleString)
    
    IstRoleEhemaligesMitglied = (role = NormalisiereRoleString(ROLE_EHEMALIGES_MITGLIED)) Or _
                                 (InStr(role, "EHEMAL") > 0) Or _
                                 (InStr(role, "AUSTRITT") > 0) Or _
                                 (InStr(role, "EX_") > 0)
End Function

' ---------------------------------------------------------
' Prüft ob EntityKey nicht zur Role passt
' ---------------------------------------------------------
Private Function EntityKeyPasstNichtZuRole(ByVal entityKey As String, ByVal roleString As String) As Boolean
    Dim role As String
    Dim keyPrefix As String
    
    role = NormalisiereRoleString(roleString)
    entityKey = Trim(entityKey)
    
    If Len(entityKey) = 0 Then
        EntityKeyPasstNichtZuRole = True
        Exit Function
    End If
    
    ' Prefix ermitteln
    If InStr(entityKey, "-") > 0 Then
        keyPrefix = UCase(Left(entityKey, InStr(entityKey, "-") - 1))
    Else
        ' Kein Prefix = vermutlich Member-ID (Zahl)
        If IsNumeric(entityKey) Then
            ' Sollte Mitglied sein
            EntityKeyPasstNichtZuRole = Not (InStr(role, "MITGLIED") > 0)
            Exit Function
        End If
        keyPrefix = ""
    End If
    
    ' Prefix vs. Role prüfen
    Select Case keyPrefix
        Case "VERS"
            EntityKeyPasstNichtZuRole = Not (role = NormalisiereRoleString(ROLE_VERSORGER))
        Case "BANK"
            EntityKeyPasstNichtZuRole = Not (role = NormalisiereRoleString(ROLE_BANK))
        Case "SHOP"
            EntityKeyPasstNichtZuRole = Not (role = NormalisiereRoleString(ROLE_SHOP))
        Case "EX"
            EntityKeyPasstNichtZuRole = Not IstRoleEhemaligesMitglied(roleString)
        Case "SONSTIGE"
            EntityKeyPasstNichtZuRole = Not (role = NormalisiereRoleString(ROLE_SONSTIGE))
        Case "SHARE"
            EntityKeyPasstNichtZuRole = False  ' SHARE passt zu mehreren Roles
        Case Else
            EntityKeyPasstNichtZuRole = False
    End Select
End Function

' ---------------------------------------------------------
' Verarbeitet manuelle Änderung der EntityRole
' ---------------------------------------------------------
Public Sub VerarbeiteManuelleRoleAenderung(ByVal zeile As Long)
    Dim wsDaten As Worksheet
    Dim neueRole As String
    Dim alterKey As String
    Dim neuerKey As String
    Dim kontoname As String
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_NAME_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    
    neueRole = Trim(CStr(wsDaten.Cells(zeile, EK_COL_ROLE).value))
    alterKey = Trim(CStr(wsDaten.Cells(zeile, EK_COL_ENTITYKEY).value))
    kontoname = Trim(CStr(wsDaten.Cells(zeile, EK_COL_ZAHLER).value))
    
    ' Prüfen ob EntityKey zur neuen Role passt
    If EntityKeyPasstNichtZuRole(alterKey, neueRole) Then
        ' Neuen EntityKey generieren
        Select Case NormalisiereRoleString(neueRole)
            Case NormalisiereRoleString(ROLE_VERSORGER)
                neuerKey = PREFIX_VERSORGER & CreateGUID()
            Case NormalisiereRoleString(ROLE_BANK)
                neuerKey = PREFIX_BANK & CreateGUID()
            Case NormalisiereRoleString(ROLE_SHOP)
                neuerKey = PREFIX_SHOP & CreateGUID()
            Case NormalisiereRoleString(ROLE_EHEMALIGES_MITGLIED)
                neuerKey = PREFIX_EHEMALIG & CreateGUID()
            Case NormalisiereRoleString(ROLE_SONSTIGE)
                neuerKey = PREFIX_SONSTIGE & CreateGUID()
            Case Else
                ' Bei Mitglied: Key beibehalten oder auf SONSTIGE setzen
                neuerKey = PREFIX_SONSTIGE & CreateGUID()
        End Select
        
        wsDaten.Cells(zeile, EK_COL_ENTITYKEY).value = neuerKey
        
        ' Zuordnung aktualisieren
        wsDaten.Cells(zeile, EK_COL_ZUORDNUNG).value = kontoname
    End If
    
    ' Parzellenschutz aktualisieren
    AktualisiereParzellenschutz zeile
    
    ' Debug-Status setzen
    SetzeAmpelFarbe wsDaten.Cells(zeile, EK_COL_DEBUG), "OK"
End Sub

' ---------------------------------------------------------
' Öffnet EntityKey-Dialog für aktuelle Zeile (falls vorhanden)
' ---------------------------------------------------------
Public Sub EntityKeyDialogFuerAktuelleZeile()
    Dim wsDaten As Worksheet
    Dim aktuelleZeile As Long
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_NAME_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then
        MsgBox "Worksheet '" & WS_NAME_DATEN & "' nicht gefunden!", vbExclamation
        Exit Sub
    End If
    
    ' Prüfen ob wir auf dem richtigen Sheet sind
    If Not ActiveSheet Is wsDaten Then
        MsgBox "Bitte wechseln Sie zum Worksheet '" & WS_NAME_DATEN & "'!", vbInformation
        Exit Sub
    End If
    
    aktuelleZeile = ActiveCell.Row
    
    ' Prüfen ob in EntityKey-Tabelle
    If aktuelleZeile < EK_FIRST_DATA_ROW Then
        MsgBox "Bitte wählen Sie eine Datenzeile in der EntityKey-Tabelle!", vbInformation
        Exit Sub
    End If
    
    If ActiveCell.Column < EK_COL_ENTITYKEY Or ActiveCell.Column > EK_COL_DEBUG Then
        MsgBox "Bitte wählen Sie eine Zelle in der EntityKey-Tabelle (Spalten R-X)!", vbInformation
        Exit Sub
    End If
    
    ' Hier könnte ein Dialog aufgerufen werden
    ' Für jetzt: Zeile neu formatieren und DropDowns einrichten
    FormatiereEntityKeyZeile aktuelleZeile, True
    
    MsgBox "EntityKey-Zeile " & aktuelleZeile & " wurde aktualisiert.", vbInformation
End Sub

' ---------------------------------------------------------
' Wird nach CSV-Import aufgerufen
' ---------------------------------------------------------
Public Sub NachCSVImport_EntityKeysAktualisieren()
    Dim startTime As Double
    startTime = Timer
    
    Debug.Print "=== NachCSVImport_EntityKeysAktualisieren START ==="
    
    ' EntityKeys aktualisieren
    AktualisiereAlleEntityKeys
    
    ' Tabelle formatieren
    FormatiereEntityKeyTabelle
    
    Debug.Print "=== NachCSVImport_EntityKeysAktualisieren ENDE (" & Format(Timer - startTime, "0.00") & "s) ==="
End Sub

' ---------------------------------------------------------
' Entfernt überflüssige Rahmen unter der Tabelle
' ---------------------------------------------------------
Public Sub EntferneUeberfluesstigeRahmen()
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim cleanRange As Range
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_NAME_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    
    lastRow = wsDaten.Cells(wsDaten.Rows.Count, EK_COL_ENTITYKEY).End(xlUp).Row
    
    ' 100 Zeilen unter der Tabelle bereinigen
    Set cleanRange = wsDaten.Range(wsDaten.Cells(lastRow + 1, EK_COL_ENTITYKEY), _
                                    wsDaten.Cells(lastRow + 100, EK_COL_DEBUG))
    
    With cleanRange
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Interior.ColorIndex = xlNone
    End With
End Sub

' ---------------------------------------------------------
' Debugging: Zeigt alle Konstanten im Immediate Window
' ---------------------------------------------------------
Public Sub DebugKonstanten()
    Debug.Print "=== EntityKey Manager Konstanten ==="
    Debug.Print "EK_COL_ENTITYKEY: " & EK_COL_ENTITYKEY & " (" & ColLetter(EK_COL_ENTITYKEY) & ")"
    Debug.Print "EK_COL_IBAN: " & EK_COL_IBAN & " (" & ColLetter(EK_COL_IBAN) & ")"
    Debug.Print "EK_COL_ZAHLER: " & EK_COL_ZAHLER & " (" & ColLetter(EK_COL_ZAHLER) & ")"
    Debug.Print "EK_COL_ZUORDNUNG: " & EK_COL_ZUORDNUNG & " (" & ColLetter(EK_COL_ZUORDNUNG) & ")"
    Debug.Print "EK_COL_PARZELLE: " & EK_COL_PARZELLE & " (" & ColLetter(EK_COL_PARZELLE) & ")"
    Debug.Print "EK_COL_ROLE: " & EK_COL_ROLE & " (" & ColLetter(EK_COL_ROLE) & ")"
    Debug.Print "EK_COL_DEBUG: " & EK_COL_DEBUG & " (" & ColLetter(EK_COL_DEBUG) & ")"
    Debug.Print "DATA_COL_DD_ENTITYROLE: " & DATA_COL_DD_ENTITYROLE & " (" & ColLetter(DATA_COL_DD_ENTITYROLE) & ")"
    Debug.Print "DATA_COL_DD_PARZELLE: " & DATA_COL_DD_PARZELLE & " (" & ColLetter(DATA_COL_DD_PARZELLE) & ")"
    Debug.Print "=== Ende Konstanten ==="
End Sub

' ============================================================
' ENDE VON mod_EntityKey_Manager.bas
' ============================================================

