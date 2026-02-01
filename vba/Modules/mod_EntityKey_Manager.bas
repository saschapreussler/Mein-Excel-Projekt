Attribute VB_Name = "mod_EntityKey_Manager"
' ***************************************************************
' MODUL: mod_EntityKey_Manager
' ZWECK: Verwaltung und Zuordnung von EntityKeys fuer Bankverkehr
' VERSION: 2.1 - 01.02.2026
' AENDERUNG: Spalten korrigiert nach Loeschung Spalte O
'            EntityRole-DropDown zeigt auf AD (Spalte 30)
'            Nutzt zentrale Konstanten aus mod_Const.bas
' ***************************************************************

' ===============================================================
' KEINE SPALTEN-KONSTANTEN MEHR HIER!
' Alle EK_COL_* Konstanten kommen aus mod_Const.bas:
' EK_COL_ENTITYKEY = 18 (R), EK_COL_IBAN = 19 (S), etc.
' ===============================================================

' EntityRole-DropDown Quelle: Spalte AD = 30
Private Const EK_ROLE_DROPDOWN_COL As Long = 30  ' AD - Dropdown-Quelle fuer EntityRole

' EntityRole-Praefixe
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

' Zebra-Farbe - IDENTISCH mit mod_Formatierung!
Private Const ZEBRA_COLOR As Long = &HDEE5E3

' ===============================================================
' HILFSFUNKTION: Prueft ob Role eine Parzelle haben darf
' ===============================================================
Private Function DarfParzelleHaben(ByVal role As String) As Boolean
    Dim normRole As String
    normRole = NormalisiereRoleString(role)
    
    DarfParzelleHaben = True
    
    If normRole = "VERSORGER" Then
        DarfParzelleHaben = False
    ElseIf normRole = "BANK" Then
        DarfParzelleHaben = False
    ElseIf normRole = "SHOP" Then
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
    
    lastRow = wsD.Cells(wsD.Rows.Count, EK_COL_IBAN).End(xlUp).Row
    If lastRow < EK_START_ROW Then lastRow = EK_START_ROW
    
    Call SetupEntityRoleDropdown(wsD, lastRow)
    
    For r = EK_START_ROW To lastRow
        iban = Trim(wsD.Cells(r, EK_COL_IBAN).value)
        kontoName = Trim(wsD.Cells(r, EK_COL_KONTONAME).value)
        
        currentEntityKey = Trim(wsD.Cells(r, EK_COL_ENTITYKEY).value)
        currentZuordnung = Trim(wsD.Cells(r, EK_COL_ZUORDNUNG).value)
        currentParzelle = Trim(wsD.Cells(r, EK_COL_PARZELLE).value)
        currentRole = Trim(wsD.Cells(r, EK_COL_ROLE).value)
        currentDebug = Trim(wsD.Cells(r, EK_COL_DEBUG).value)
        
        If iban = "" And kontoName = "" Then GoTo NextRow
        
        If HatBereitsGueltigeDaten(currentEntityKey, currentZuordnung, currentRole) Then
            zeilenUnveraendert = zeilenUnveraendert + 1
            If currentRole <> "" Then
                Call SetzeAmpelFarbe(wsD, r, 1)
            End If
            GoTo NextRow
        End If
        
        zeilenNeu = zeilenNeu + 1
        
        Set mitgliederGefunden = SucheMitgliederZuKontoname(kontoName, wsM, wsH)
        
        Call GeneriereEntityKeyUndZuordnung(mitgliederGefunden, kontoName, wsM, _
                                             newEntityKey, zuordnung, parzellen, entityRole, debugInfo, ampelStatus)
        
        If currentEntityKey = "" And newEntityKey <> "" Then wsD.Cells(r, EK_COL_ENTITYKEY).value = newEntityKey
        If currentZuordnung = "" And zuordnung <> "" Then wsD.Cells(r, EK_COL_ZUORDNUNG).value = zuordnung
        
        If currentParzelle = "" And parzellen <> "" And DarfParzelleHaben(entityRole) Then
            wsD.Cells(r, EK_COL_PARZELLE).value = parzellen
        End If
        
        If currentRole = "" And entityRole <> "" Then wsD.Cells(r, EK_COL_ROLE).value = entityRole
        If currentDebug = "" Then wsD.Cells(r, EK_COL_DEBUG).value = debugInfo
        
        Call SetzeAmpelFarbe(wsD, r, ampelStatus)
        Call SetzeZellschutzFuerZeile(wsD, r, entityRole)
        
        If ampelStatus = 3 Then
            zeilenRot.Add r
        ElseIf ampelStatus = 2 Then
            zeilenGelb.Add r
        End If
        
NextRow:
    Next r
    
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
' ===============================================================
Private Sub SetzeZellschutzFuerZeile(ByRef ws As Worksheet, ByVal zeile As Long, ByVal role As String)
    
    ws.Cells(zeile, EK_COL_ZUORDNUNG).Locked = False
    ws.Cells(zeile, EK_COL_ROLE).Locked = False
    ws.Cells(zeile, EK_COL_DEBUG).Locked = False
    
    If DarfParzelleHaben(role) Then
        ws.Cells(zeile, EK_COL_PARZELLE).Locked = False
    Else
        ws.Cells(zeile, EK_COL_PARZELLE).Locked = True
    End If
    
End Sub

' ===============================================================
' HILFSFUNKTION: Prueft ob Zeile bereits gueltige Daten hat
' ===============================================================
Private Function HatBereitsGueltigeDaten(ByVal entityKey As String, _
                                          ByVal zuordnung As String, _
                                          ByVal role As String) As Boolean
    
    HatBereitsGueltigeDaten = False
    
    If entityKey <> "" Then
        If Not IsNumeric(entityKey) Then
            HatBereitsGueltigeDaten = True
            Exit Function
        End If
    End If
    
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
    Dim mitgliedInfo(0 To 8) As Variant
    Dim zeilen As Variant
    Dim zeile As Variant
    Dim nameKombiniert As String
    Dim nameParts() As String
    Dim austrittsDatum As Date
    Dim matchResult As Long
    
    Set SucheMitgliederZuKontoname = gefunden
    
    If kontoName = "" Then Exit Function
    
    zeilen = Split(kontoName, vbLf)
    
    For Each zeile In zeilen
        kontoNameNorm = NormalisiereStringFuerVergleich(CStr(zeile))
        If kontoNameNorm = "" Then GoTo NextZeile
        
        lastRow = wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
        
        For r = M_START_ROW To lastRow
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
                        mitgliedInfo(6) = False
                        mitgliedInfo(7) = CDate("01.01.1900")
                        mitgliedInfo(8) = matchResult
                        gefunden.Add mitgliedInfo
                    End If
                End If
            End If
        Next r
        
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
                    mitgliedInfo(6) = True
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
        PruefeNamensMatch = 2
    ElseIf nachnameGefunden Then
        PruefeNamensMatch = 1
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
' ===============================================================
Private Sub GeneriereEntityKeyUndZuordnung(ByRef mitglieder As Collection, _
                                            ByVal kontoName As String, _
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
    
    outEntityKey = ""
    outZuordnung = ""
    outParzellen = ""
    outEntityRole = ""
    outDebugInfo = ""
    outAmpelStatus = 1
    
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
        If IstVersorger(kontoName) Then
            outEntityKey = PREFIX_VERSORGER & CreateGUID()
            outEntityRole = ROLE_VERSORGER
            outZuordnung = ExtrahiereAnzeigeName(kontoName)
            outParzellen = ""
            outDebugInfo = "Automatisch als VERSORGER erkannt"
            outAmpelStatus = 1
            Exit Sub
        ElseIf IstBank(kontoName) Then
            outEntityKey = PREFIX_BANK & CreateGUID()
            outEntityRole = ROLE_BANK
            outZuordnung = ExtrahiereAnzeigeName(kontoName)
            outParzellen = ""
            outDebugInfo = "Automatisch als BANK erkannt"
            outAmpelStatus = 1
            Exit Sub
        ElseIf IstShop(kontoName) Then
            outEntityKey = PREFIX_SHOP & CreateGUID()
            outEntityRole = ROLE_SHOP
            outZuordnung = ExtrahiereAnzeigeName(kontoName)
            outParzellen = ""
            outDebugInfo = "Automatisch als SHOP erkannt"
            outAmpelStatus = 1
            Exit Sub
        End If
        
        If mitgliederNurNachname.Count > 0 Then
            outEntityKey = ""
            outZuordnung = ""
            outParzellen = ""
            outEntityRole = ""
            outDebugInfo = "NUR NACHNAME GEFUNDEN - Bitte pruefen! Moegliche Treffer:"
            outAmpelStatus = 2
            
            For i = 1 To mitgliederNurNachname.Count
                mitgliedInfo = mitgliederNurNachname(i)
                outDebugInfo = outDebugInfo & vbLf & "  ? " & mitgliedInfo(1) & ", " & mitgliedInfo(2) & " (Parz. " & mitgliedInfo(3) & ")"
            Next i
            Exit Sub
        Else
            outEntityKey = ""
            outZuordnung = ""
            outParzellen = ""
            outEntityRole = ""
            outDebugInfo = "KEIN MITGLIED GEFUNDEN - Manuelle Zuordnung erforderlich"
            outAmpelStatus = 3
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
            mitgliedInfo = mitgliederExakt(1)
            outEntityKey = PREFIX_EHEMALIG & mitgliedInfo(0)
            outZuordnung = mitgliedInfo(1) & ", " & mitgliedInfo(2)
            outParzellen = mitgliedInfo(3) & " (bis " & Format(mitgliedInfo(7), "dd.mm.yyyy") & ")"
            outEntityRole = ROLE_EHEMALIGES_MITGLIED
            outDebugInfo = "Ehemaliges Mitglied - exakter Treffer"
            outAmpelStatus = 1
        Else
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

' ===============================================================
' HILFSFUNKTION: Extrahiert Anzeigename aus Kontoname
' ===============================================================
Private Function ExtrahiereAnzeigeName(ByVal kontoName As String) As String
    Dim zeilen() As String
    Dim erstesElement As String
    
    If kontoName = "" Then
        ExtrahiereAnzeigeName = ""
        Exit Function
    End If
    
    zeilen = Split(kontoName, vbLf)
    erstesElement = Trim(zeilen(0))
    
    If Len(erstesElement) > 50 Then
        erstesElement = Left(erstesElement, 50) & "..."
    End If
    
    ExtrahiereAnzeigeName = erstesElement
End Function

' ===============================================================
' HILFSFUNKTION: Holt ALLE Parzellen eines Mitglieds
' ===============================================================
Private Function HoleAlleParzellen(ByVal memberID As String, ByRef wsM As Worksheet) As String
    Dim r As Long
    Dim lastRow As Long
    Dim currentMemberID As String
    Dim parzelle As String
    Dim result As String
    
    result = ""
    
    If memberID = "" Then
        HoleAlleParzellen = ""
        Exit Function
    End If
    
    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_MEMBER_ID).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        currentMemberID = Trim(wsM.Cells(r, M_COL_MEMBER_ID).value)
        
        If currentMemberID = memberID Then
            parzelle = Trim(wsM.Cells(r, M_COL_PARZELLE).value)
            If parzelle <> "" Then
                If InStr(result, parzelle) = 0 Then
                    If result <> "" Then result = result & vbLf
                    result = result & parzelle
                End If
            End If
        End If
    Next r
    
    HoleAlleParzellen = result
End Function

' ===============================================================
' OEFFENTLICHE PROZEDUR: Aktualisiert Parzellen fuer ein Mitglied
' ===============================================================
Public Sub AktualisiereParzellenFuerMitglied(ByVal memberID As String)
    Dim wsD As Worksheet
    Dim wsM As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim entityKey As String
    Dim currentRole As String
    Dim neueParzellen As String
    
    If memberID = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    neueParzellen = HoleAlleParzellen(memberID, wsM)
    
    On Error Resume Next
    wsD.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    lastRow = wsD.Cells(wsD.Rows.Count, EK_COL_IBAN).End(xlUp).Row
    
    For r = EK_START_ROW To lastRow
        entityKey = Trim(wsD.Cells(r, EK_COL_ENTITYKEY).value)
        currentRole = Trim(wsD.Cells(r, EK_COL_ROLE).value)
        
        If entityKey = memberID Or _
           InStr(entityKey, memberID & "_") > 0 Or _
           InStr(entityKey, "_" & memberID) > 0 Then
            
            If DarfParzelleHaben(currentRole) Then
                wsD.Cells(r, EK_COL_PARZELLE).value = neueParzellen
                wsD.Cells(r, EK_COL_PARZELLE).Locked = False
            End If
        End If
    Next r
    
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
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
' ===============================================================
Public Function IstVersorger(ByVal Name As String) As Boolean
    Dim keywords As Variant
    Dim kw As Variant
    
    keywords = Array( _
        "stadtwerke", "energie", "strom", "gas", "wasser", _
        "telekom", "vodafone", "o2", "1&1", "versicherung", _
        "allianz", "huk", "devk", "axa", "ergo", "enviam", _
        "enso", "ewe", "eon", "e.on", "rwe", "vattenfall", _
        "gvv", "signal iduna", "debeka", "lvm", "abfall", _
        "muell", "entsorgung", "abwasser", "kanal", _
        "wazv", "zweckverband", "wasserverband", "abwasserverband", _
        "grundstuecksgesellschaft", "wohnungsbau", "wohnungsgesellschaft", _
        "hausverwaltung", "immobilien", "grundstueck", _
        "finanzamt", "rundfunk", "gez", "beitragsservice", _
        "kfz", "haftpflicht", "hausrat", "rechtsschutz", _
        "krankenkasse", "aok", "barmer", "dak", "tk", "ikk", _
        "berufsgenossenschaft", "rentenversicherung", _
        "stadt ", "gemeinde ", "kommune", "landkreis", _
        "werder", "havel", "potsdam", "brandenburg")
    
    Name = LCase(Name)
    
    For Each kw In keywords
        If InStr(Name, kw) > 0 Then
            IstVersorger = True
            Exit Function
        End If
    Next kw
    
    IstVersorger = False
End Function

Public Function IstBank(ByVal Name As String) As Boolean
    Dim keywords As Variant
    Dim kw As Variant
    
    keywords = Array( _
        "sparkasse", "volksbank", "raiffeisenbank", "commerzbank", _
        "deutsche bank", "postbank", "ing", "dkb", "targobank", _
        "sparda", "psd bank", "santander", "hypovereinsbank", _
        "unicredit", "n26", "comdirect", "consorsbank", _
        "mittelbrandenburgische", "mbs", "brandenburger bank", _
        "kreditbank", "landesbank", "girozentrale", _
        "bausparkasse", "schwaebisch hall", "lbs", "wuestenrot")
    
    Name = LCase(Name)
    
    For Each kw In keywords
        If InStr(Name, kw) > 0 Then
            IstBank = True
            Exit Function
        End If
    Next kw
    
    IstBank = False
End Function

Public Function IstShop(ByVal Name As String) As Boolean
    Dim keywords As Variant
    Dim kw As Variant
    
    keywords = Array( _
        "amazon", "ebay", "paypal", "otto", "zalando", _
        "mediamarkt", "saturn", "lidl", "aldi", "rewe", _
        "edeka", "penny", "netto", "kaufland", "hornbach", _
        "obi", "bauhaus", "toom", "hagebau", "dehner", _
        "rossmann", "dm-drogerie", "mueller drogerie", _
        "ikea", "poco", "roller", "moemax", "xxxlutz", _
        "h&m", "c&a", "kik", "takko", "ernsting", _
        "decathlon", "intersport", "karstadt", "galeria", _
        "thalia", "hugendubel", "weltbild", _
        "notebooksbilliger", "cyberport", "alternate", _
        "thomann", "musicstore", "conrad", "reichelt", _
        "fressnapf", "zooplus", "futterhaus", _
        "apotheke", "docmorris", "shop-apotheke")
    
    Name = LCase(Name)
    
    For Each kw In keywords
        If InStr(Name, kw) > 0 Then
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
' HILFSPROZEDUR: Setzt Dropdown fuer EntityRole
' KORRIGIERT: Zeigt auf Spalte AD (30) statt AF (32)
' ===============================================================
Private Sub SetupEntityRoleDropdown(ByRef ws As Worksheet, ByVal lastRow As Long)
    
    Dim rngDropdown As Range
    Dim dropdownSource As String
    Dim lastRoleRow As Long
    Dim dropdownEndRow As Long
    
    ' EntityRole-Liste aus Spalte AD (30) lesen
    lastRoleRow = ws.Cells(ws.Rows.Count, EK_ROLE_DROPDOWN_COL).End(xlUp).Row
    If lastRoleRow < 4 Then lastRoleRow = 10
    
    ' DropDown-Quelle: Spalte AD
    dropdownSource = "=$AD$4:$AD$" & lastRoleRow
    
    dropdownEndRow = lastRow + 50
    
    ' EK_COL_ROLE = 23 = Spalte W (nach Spaltenverschiebung)
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
        .ErrorTitle = "Ungueltige Eingabe"
        .ErrorMessage = "Bitte waehlen Sie einen Wert aus der Liste."
    End With
    On Error GoTo 0
    
End Sub

' ===============================================================
' HILFSPROZEDUR: Setzt Ampelfarbe (nur Spalten U-X)
' KORRIGIERT: Spalten nach Verschiebung
' ===============================================================
Private Sub SetzeAmpelFarbe(ByRef ws As Worksheet, ByVal zeile As Long, ByVal ampelStatus As Long)
    
    Dim rng As Range
    Dim farbe As Long
    
    ' EK_COL_ZUORDNUNG = 21 (U), EK_COL_DEBUG = 24 (X)
    Set rng = ws.Range(ws.Cells(zeile, EK_COL_ZUORDNUNG), ws.Cells(zeile, EK_COL_DEBUG))
    
    Select Case ampelStatus
        Case 1
            farbe = RGB(198, 224, 180)  ' Hellgruen
        Case 2
            farbe = RGB(255, 230, 153)  ' Hellgelb
        Case 3
            farbe = RGB(255, 150, 150)  ' Hellrot
        Case Else
            farbe = RGB(198, 224, 180)  ' Hellgruen
    End Select
    
    rng.Interior.color = farbe
    
End Sub

' ===============================================================
' HILFSPROZEDUR: Formatiert die EntityKey-Tabelle
' KORRIGIERT: Spalte W (EntityRole) LINKSBUENDIG + AutoFit
' ===============================================================
Private Sub FormatiereEntityKeyTabelle(ByRef ws As Worksheet, ByVal lastRow As Long)
    
    Dim rngTable As Range
    Dim rngOhneEntityKey As Range
    Dim rngZebra As Range
    Dim r As Long
    Dim col As Long
    Dim currentRole As String
    
    If lastRow < EK_START_ROW Then Exit Sub
    
    ' EK_COL_ENTITYKEY = 18 (R), EK_COL_DEBUG = 24 (X)
    Set rngTable = ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                            ws.Cells(lastRow, EK_COL_DEBUG))
    
    ' EK_COL_IBAN = 19 (S)
    Set rngOhneEntityKey = ws.Range(ws.Cells(EK_START_ROW, EK_COL_IBAN), _
                                     ws.Cells(lastRow, EK_COL_DEBUG))
    
    ' Rahmen
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' Vertikal zentriert
    rngTable.VerticalAlignment = xlCenter
    rngOhneEntityKey.WrapText = True
    
    ' Spalte R (EntityKey): Kein Textumbruch, feste Breite, linksbuendig
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                  ws.Cells(lastRow, EK_COL_ENTITYKEY))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ENTITYKEY).ColumnWidth = 14
    
    ' Spalte V (Parzelle): Zentriert
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_PARZELLE), _
             ws.Cells(lastRow, EK_COL_PARZELLE)).HorizontalAlignment = xlCenter
    
    ' Spalte W (EntityRole): LINKSBUENDIG + kein Textumbruch
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ROLE), _
                  ws.Cells(lastRow, EK_COL_ROLE))
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    
    ' Spalten R-T immer gesperrt
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
             ws.Cells(lastRow, EK_COL_KONTONAME)).Locked = True
    
    ' ============================================================
    ' ZELLSCHUTZ und ZEBRA pro Zeile
    ' ============================================================
    For r = EK_START_ROW To lastRow
        currentRole = Trim(ws.Cells(r, EK_COL_ROLE).value)
        
        Call SetzeZellschutzFuerZeile(ws, r, currentRole)
        
        ' Zebra fuer Spalten R-T
        Set rngZebra = ws.Range(ws.Cells(r, EK_COL_ENTITYKEY), ws.Cells(r, EK_COL_KONTONAME))
        
        If (r - EK_START_ROW) Mod 2 = 1 Then
            rngZebra.Interior.color = ZEBRA_COLOR
        Else
            rngZebra.Interior.ColorIndex = xlNone
        End If
    Next r
    
    ' AutoFit Spaltenbreiten fuer S-X
    For col = EK_COL_IBAN To EK_COL_DEBUG
        ws.Columns(col).AutoFit
    Next col
    
    ' Speziell: AutoFit fuer Spalte W (EntityRole)
    ws.Columns(EK_COL_ROLE).AutoFit
    
    ' Zeilenhoehe AutoFit
    ws.Rows(EK_START_ROW & ":" & lastRow).AutoFit
    
End Sub

' ===============================================================
' OEFFENTLICHE PROZEDUR: Formatiert eine einzelne Zeile
' ===============================================================
Public Sub FormatiereEntityKeyZeile(ByVal zeile As Long)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rngZebra As Range
    Dim col As Long
    Dim currentRole As String
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    If zeile < EK_START_ROW Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, EK_COL_IBAN).End(xlUp).Row
    If zeile > lastRow Then Exit Sub
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    currentRole = Trim(ws.Cells(zeile, EK_COL_ROLE).value)
    
    Call SetzeZellschutzFuerZeile(ws, zeile, currentRole)
    
    ' Zebra fuer Spalten R-T dieser Zeile
    Set rngZebra = ws.Range(ws.Cells(zeile, EK_COL_ENTITYKEY), ws.Cells(zeile, EK_COL_KONTONAME))
    
    If (zeile - EK_START_ROW) Mod 2 = 1 Then
        rngZebra.Interior.color = ZEBRA_COLOR
    Else
        rngZebra.Interior.ColorIndex = xlNone
    End If
    
    ' Vertikal zentriert fuer diese Zeile
    ws.Range(ws.Cells(zeile, EK_COL_ENTITYKEY), ws.Cells(zeile, EK_COL_DEBUG)).VerticalAlignment = xlCenter
    
    ' AutoFit Spaltenbreite S-X
    For col = EK_COL_IBAN To EK_COL_DEBUG
        ws.Columns(col).AutoFit
    Next col
    
    ' Zeilenhoehe AutoFit fuer diese Zeile
    ws.Rows(zeile).AutoFit
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
End Sub

' ===============================================================
' HILFSFUNKTION: Normalisiert Role-String fuer Vergleich
' ===============================================================
Private Function NormalisiereRoleString(ByVal role As String) As String
    Dim result As String
    result = UCase(Trim(role))
    result = Replace(result, " ", "")
    result = Replace(result, "_", "")
    NormalisiereRoleString = result
End Function

' ===============================================================
' HILFSFUNKTION: Prueft ob Role "EHEMALIGES MITGLIED" ist
' ===============================================================
Private Function IstRoleEhemaligesMitglied(ByVal role As String) As Boolean
    Dim normRole As String
    normRole = NormalisiereRoleString(role)
    
    IstRoleEhemaligesMitglied = (normRole = "EHEMALIGESVMITGLIED" Or _
                                 normRole = "EHEMALIGEMITGLIED" Or _
                                 normRole = "EHEMALIGES_MITGLIED" Or _
                                 normRole = "EHEMALIGESMITGLIED")
End Function

' ===============================================================
' HILFSFUNKTION: Prueft ob EntityKey zum Role-Typ passt
' ===============================================================
Private Function EntityKeyPasstNichtZuRole(ByVal entityKey As String, ByVal role As String) As Boolean
    Dim normRole As String
    
    EntityKeyPasstNichtZuRole = False
    
    If entityKey = "" Then
        EntityKeyPasstNichtZuRole = True
        Exit Function
    End If
    
    normRole = NormalisiereRoleString(role)
    
    If IstRoleEhemaligesMitglied(role) Then
        If Left(entityKey, Len(PREFIX_EHEMALIG)) <> PREFIX_EHEMALIG Then
            EntityKeyPasstNichtZuRole = True
        End If
        Exit Function
    End If
    
    If normRole = "VERSORGER" Then
        If Left(entityKey, Len(PREFIX_VERSORGER)) <> PREFIX_VERSORGER Then
            EntityKeyPasstNichtZuRole = True
        End If
        Exit Function
    End If
    
    If normRole = "BANK" Then
        If Left(entityKey, Len(PREFIX_BANK)) <> PREFIX_BANK Then
            EntityKeyPasstNichtZuRole = True
        End If
        Exit Function
    End If
    
    If normRole = "SHOP" Then
        If Left(entityKey, Len(PREFIX_SHOP)) <> PREFIX_SHOP Then
            EntityKeyPasstNichtZuRole = True
        End If
        Exit Function
    End If
    
    If normRole = "SONSTIGE" Then
        If Left(entityKey, Len(PREFIX_SONSTIGE)) <> PREFIX_SONSTIGE Then
            EntityKeyPasstNichtZuRole = True
        End If
        Exit Function
    End If
    
    If InStr(normRole, "MITGLIED") > 0 And InStr(normRole, "EHEMAL") = 0 Then
        If Left(entityKey, Len(PREFIX_EHEMALIG)) = PREFIX_EHEMALIG Or _
           Left(entityKey, Len(PREFIX_VERSORGER)) = PREFIX_VERSORGER Or _
           Left(entityKey, Len(PREFIX_BANK)) = PREFIX_BANK Or _
           Left(entityKey, Len(PREFIX_SHOP)) = PREFIX_SHOP Or _
           Left(entityKey, Len(PREFIX_SONSTIGE)) = PREFIX_SONSTIGE Then
            EntityKeyPasstNichtZuRole = True
        End If
        Exit Function
    End If
    
End Function

' ===============================================================
' OEFFENTLICHE PROZEDUR: Verarbeitet manuelle Aenderung in Spalte W
' ===============================================================
Public Sub VerarbeiteManuelleRoleAenderung(ByVal zeile As Long)
    Dim ws As Worksheet
    Dim wsM As Worksheet
    Dim wsH As Worksheet
    Dim kontoName As String
    Dim neueRole As String
    Dim entityKey As String
    Dim zuordnung As String
    Dim mitglieder As Collection
    Dim mitgliedInfo As Variant
    Dim i As Long
    Dim gefunden As Boolean
    Dim entityKeyMussAktualisiert As Boolean
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    
    If zeile < EK_START_ROW Then Exit Sub
    
    neueRole = Trim(ws.Cells(zeile, EK_COL_ROLE).value)
    kontoName = Trim(ws.Cells(zeile, EK_COL_KONTONAME).value)
    entityKey = Trim(ws.Cells(zeile, EK_COL_ENTITYKEY).value)
    zuordnung = Trim(ws.Cells(zeile, EK_COL_ZUORDNUNG).value)
    
    If neueRole = "" Then Exit Sub
    
    Application.EnableEvents = False
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    entityKeyMussAktualisiert = EntityKeyPasstNichtZuRole(entityKey, neueRole)
    
    If entityKeyMussAktualisiert Then
        
        If IstRoleEhemaligesMitglied(neueRole) Then
            Set mitglieder = SucheMitgliederZuKontoname(kontoName, wsM, wsH)
            gefunden = False
            
            For i = 1 To mitglieder.Count
                mitgliedInfo = mitglieder(i)
                If mitgliedInfo(6) = True Then
                    entityKey = PREFIX_EHEMALIG & CStr(mitgliedInfo(0))
                    If zuordnung = "" Then zuordnung = mitgliedInfo(1) & ", " & mitgliedInfo(2)
                    gefunden = True
                    Exit For
                End If
            Next i
            
            If Not gefunden Then
                entityKey = PREFIX_EHEMALIG & CreateGUID()
                If zuordnung = "" Then zuordnung = ExtrahiereNachnameVorname(kontoName)
            End If
            
        ElseIf NormalisiereRoleString(neueRole) = "VERSORGER" Then
            entityKey = PREFIX_VERSORGER & CreateGUID()
            If zuordnung = "" Then zuordnung = ExtrahiereAnzeigeName(kontoName)
            ws.Cells(zeile, EK_COL_PARZELLE).value = ""
            
        ElseIf NormalisiereRoleString(neueRole) = "BANK" Then
            entityKey = PREFIX_BANK & CreateGUID()
            If zuordnung = "" Then zuordnung = ExtrahiereAnzeigeName(kontoName)
            ws.Cells(zeile, EK_COL_PARZELLE).value = ""
            
        ElseIf NormalisiereRoleString(neueRole) = "SHOP" Then
            entityKey = PREFIX_SHOP & CreateGUID()
            If zuordnung = "" Then zuordnung = ExtrahiereAnzeigeName(kontoName)
            ws.Cells(zeile, EK_COL_PARZELLE).value = ""
            
        ElseIf NormalisiereRoleString(neueRole) = "SONSTIGE" Then
            entityKey = PREFIX_SONSTIGE & CreateGUID()
            If zuordnung = "" Then zuordnung = ExtrahiereAnzeigeName(kontoName)
            
        ElseIf InStr(NormalisiereRoleString(neueRole), "MITGLIED") > 0 Then
            Set mitglieder = SucheMitgliederZuKontoname(kontoName, wsM, wsH)
            gefunden = False
            
            For i = 1 To mitglieder.Count
                mitgliedInfo = mitglieder(i)
                If mitgliedInfo(6) = False Then
                    entityKey = CStr(mitgliedInfo(0))
                    If zuordnung = "" Then zuordnung = mitgliedInfo(1) & ", " & mitgliedInfo(2)
                    If Trim(ws.Cells(zeile, EK_COL_PARZELLE).value) = "" Then
                        ws.Cells(zeile, EK_COL_PARZELLE).value = HoleAlleParzellen(entityKey, wsM)
                    End If
                    gefunden = True
                    Exit For
                End If
            Next i
            
            If Not gefunden Then
                entityKey = CreateGUID()
                If zuordnung = "" Then zuordnung = ExtrahiereNachnameVorname(kontoName)
            End If
        End If
        
        ws.Cells(zeile, EK_COL_ENTITYKEY).value = entityKey
        If zuordnung <> "" And Trim(ws.Cells(zeile, EK_COL_ZUORDNUNG).value) = "" Then
            ws.Cells(zeile, EK_COL_ZUORDNUNG).value = zuordnung
        End If
        
        ws.Cells(zeile, EK_COL_DEBUG).value = "Manuell zugeordnet am " & Format(Now, "dd.mm.yyyy hh:mm")
    End If
    
    Call SetzeZellschutzFuerZeile(ws, zeile, neueRole)
    
    Call SetzeAmpelFarbe(ws, zeile, 1)
    Call FormatiereEntityKeyZeile(zeile)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
End Sub

' ===============================================================
' HILFSFUNKTION: Extrahiert "Nachname, Vorname" aus Kontoname
' ===============================================================
Private Function ExtrahiereNachnameVorname(ByVal kontoName As String) As String
    Dim teile() As String
    Dim erstesElement As String
    Dim worte() As String
    
    If kontoName = "" Then
        ExtrahiereNachnameVorname = ""
        Exit Function
    End If
    
    teile = Split(kontoName, vbLf)
    erstesElement = Trim(teile(0))
    
    worte = Split(erstesElement, " ")
    
    If UBound(worte) >= 1 Then
        ExtrahiereNachnameVorname = worte(UBound(worte)) & ", " & worte(0)
    Else
        ExtrahiereNachnameVorname = erstesElement
    End If
    
End Function

' ===============================================================
' DIALOG: Manuelle EntityKey-Zuordnung fuer aktuelle Zeile
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
        MsgBox "Bitte waehlen Sie eine Datenzeile (ab Zeile " & EK_START_ROW & ").", vbExclamation
        Exit Sub
    End If
    
    If ActiveSheet.Name <> WS_DATEN Then
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
                auswahlText = auswahlText & " [EHEMALIG]"
            End If
            auswahlText = auswahlText & " (Parzelle " & mitgliedInfo(3) & ")"
            If mitgliedInfo(8) = 2 Then
                auswahlText = auswahlText & " [EXAKT]"
            ElseIf mitgliedInfo(8) = 1 Then
                auswahlText = auswahlText & " [nur Nachname]"
            End If
            auswahlText = auswahlText & vbCrLf
        Next i
        auswahlText = auswahlText & vbCrLf
    Else
        auswahlText = auswahlText & "Keine Mitglieder gefunden." & vbCrLf & vbCrLf
    End If
    
    auswahlText = auswahlText & "Bitte waehlen Sie:" & vbCrLf
    auswahlText = auswahlText & "  M = MITGLIED (aktiv)" & vbCrLf
    auswahlText = auswahlText & "  E = EHEMALIGES MITGLIED" & vbCrLf
    auswahlText = auswahlText & "  G = GEMEINSCHAFTSKONTO" & vbCrLf
    auswahlText = auswahlText & "  V = VERSORGER" & vbCrLf
    auswahlText = auswahlText & "  B = BANK" & vbCrLf
    auswahlText = auswahlText & "  S = SHOP" & vbCrLf
    auswahlText = auswahlText & "  O = SONSTIGE" & vbCrLf
    auswahlText = auswahlText & "  X = Abbrechen"
    
    eingabe = UCase(Trim(InputBox(auswahlText, "EntityKey-Zuordnung", "M")))
    
    If eingabe = "" Or eingabe = "X" Then Exit Sub
    
    wsD.Unprotect PASSWORD:=PASSWORD
    
    Set uniqueIDs = CreateObject("Scripting.Dictionary")
    
    Select Case eingabe
        Case "M"
            For i = 1 To mitglieder.Count
                mitgliedInfo = mitglieder(i)
                If mitgliedInfo(6) = False Then
                    neuerEntityKey = CStr(mitgliedInfo(0))
                    neueZuordnung = mitgliedInfo(1) & ", " & mitgliedInfo(2)
                    neueParzellen = HoleAlleParzellen(CStr(mitgliedInfo(0)), wsM)
                    neueRole = ROLE_MITGLIED_MIT_PACHT
                    Exit For
                End If
            Next i
            If neuerEntityKey = "" Then
                neuerEntityKey = CreateGUID()
                neueZuordnung = ExtrahiereNachnameVorname(kontoName)
                neueRole = ROLE_MITGLIED_MIT_PACHT
            End If
            
        Case "E"
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
                neueZuordnung = ExtrahiereNachnameVorname(kontoName)
                neueRole = ROLE_EHEMALIGES_MITGLIED
            End If
            
        Case "G"
            memberIDs = ""
            For i = 1 To mitglieder.Count
                mitgliedInfo = mitglieder(i)
                If Not uniqueIDs.Exists(CStr(mitgliedInfo(0))) Then
                    uniqueIDs.Add CStr(mitgliedInfo(0)), True
                    If memberIDs <> "" Then memberIDs = memberIDs & "_"
                    memberIDs = memberIDs & CStr(mitgliedInfo(0))
                    
                    If neueZuordnung <> "" Then neueZuordnung = neueZuordnung & vbLf
                    neueZuordnung = neueZuordnung & mitgliedInfo(1) & ", " & mitgliedInfo(2)
                    If mitgliedInfo(6) = True Then neueZuordnung = neueZuordnung & " (ehem.)"
                    
                    Dim parz As String
                    parz = HoleAlleParzellen(CStr(mitgliedInfo(0)), wsM)
                    Dim parzArr() As String
                    Dim p As Long
                    parzArr = Split(parz, vbLf)
                    For p = LBound(parzArr) To UBound(parzArr)
                        If Trim(parzArr(p)) <> "" Then
                            If InStr(neueParzellen, Trim(parzArr(p))) = 0 Then
                                If neueParzellen <> "" Then neueParzellen = neueParzellen & vbLf
                                neueParzellen = neueParzellen & Trim(parzArr(p))
                            End If
                        End If
                    Next p
                End If
            Next i
            
            If memberIDs = "" Then memberIDs = CreateGUID()
            neuerEntityKey = PREFIX_SHARE & memberIDs
            neueRole = ROLE_MITGLIED_MIT_PACHT
            
        Case "V"
            neuerEntityKey = PREFIX_VERSORGER & CreateGUID()
            neueRole = ROLE_VERSORGER
            neueZuordnung = ExtrahiereAnzeigeName(kontoName)
            neueParzellen = ""
            
        Case "B"
            neuerEntityKey = PREFIX_BANK & CreateGUID()
            neueRole = ROLE_BANK
            neueZuordnung = ExtrahiereAnzeigeName(kontoName)
            neueParzellen = ""
            
        Case "S"
            neuerEntityKey = PREFIX_SHOP & CreateGUID()
            neueRole = ROLE_SHOP
            neueZuordnung = ExtrahiereAnzeigeName(kontoName)
            neueParzellen = ""
            
        Case "O"
            neuerEntityKey = PREFIX_SONSTIGE & CreateGUID()
            neueRole = ROLE_SONSTIGE
            neueZuordnung = ExtrahiereAnzeigeName(kontoName)
            
        Case Else
            MsgBox "Ungueltige Eingabe.", vbExclamation
            wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
            Exit Sub
    End Select
    
    wsD.Cells(aktuelleZeile, EK_COL_ENTITYKEY).value = neuerEntityKey
    
    If Trim(wsD.Cells(aktuelleZeile, EK_COL_ZUORDNUNG).value) = "" Then
        wsD.Cells(aktuelleZeile, EK_COL_ZUORDNUNG).value = neueZuordnung
    End If
    
    If DarfParzelleHaben(neueRole) Then
        If Trim(wsD.Cells(aktuelleZeile, EK_COL_PARZELLE).value) = "" And neueParzellen <> "" Then
            wsD.Cells(aktuelleZeile, EK_COL_PARZELLE).value = neueParzellen
        End If
    Else
        wsD.Cells(aktuelleZeile, EK_COL_PARZELLE).value = ""
    End If
    
    wsD.Cells(aktuelleZeile, EK_COL_ROLE).value = neueRole
    
    wsD.Cells(aktuelleZeile, EK_COL_DEBUG).value = "Manuell zugeordnet am " & Format(Now, "dd.mm.yyyy hh:mm")
    
    Call SetzeZellschutzFuerZeile(wsD, aktuelleZeile, neueRole)
    
    Call SetzeAmpelFarbe(wsD, aktuelleZeile, 1)
    Call FormatiereEntityKeyZeile(aktuelleZeile)
    
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    MsgBox "EntityKey erfolgreich zugeordnet!" & vbCrLf & vbCrLf & _
           "EntityKey: " & neuerEntityKey & vbCrLf & _
           "Rolle: " & neueRole, vbInformation, "Zuordnung erfolgreich"
    
End Sub

' ===============================================================
' OEFFENTLICHE PROZEDUR: Wird nach CSV-Import aufgerufen
' ===============================================================
Public Sub NachCSVImport_EntityKeysAktualisieren()
    Call AktualisiereAlleEntityKeys
End Sub

' ===============================================================
' HILFSPROZEDUR: Entfernt ueberfluessige Rahmenlinien
' ===============================================================
Public Sub EntferneUeberfluesstigeRahmen()
    
    Dim ws As Worksheet
    Dim lastDataRow As Long
    Dim rngZuLoeschen As Range
    
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    lastDataRow = ws.Cells(ws.Rows.Count, EK_COL_IBAN).End(xlUp).Row
    If lastDataRow < EK_START_ROW Then lastDataRow = EK_START_ROW
    
    If lastDataRow < 1000 Then
        Set rngZuLoeschen = ws.Range(ws.Cells(lastDataRow + 1, EK_COL_ENTITYKEY), ws.Cells(1000, EK_COL_DEBUG))
        rngZuLoeschen.Borders.LineStyle = xlNone
        rngZuLoeschen.Interior.ColorIndex = xlNone
    End If
    
    MsgBox "Ueberfluessige Rahmenlinien entfernt!" & vbCrLf & _
           "Letzte Datenzeile: " & lastDataRow, vbInformation
    
End Sub



