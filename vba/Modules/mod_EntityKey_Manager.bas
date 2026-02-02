Attribute VB_Name = "mod_EntityKey_Manager"
' ***************************************************************
' MODUL: mod_EntityKey_Manager
' ZWECK: Verwaltung und Zuordnung von EntityKeys fuer Bankverkehr
' VERSION: 2.3 - 02.02.2026 - KORRIGIERT
' ***************************************************************

' ===============================================================
' KEINE SPALTEN-KONSTANTEN HIER - ALLE AUS mod_Const.bas!
' ===============================================================

' EntityRole-Praefixe fuer GUID-Generierung
Private Const PREFIX_SHARE As String = "SHARE-"
Private Const PREFIX_VERSORGER As String = "VERS-"
Private Const PREFIX_BANK As String = "BANK-"
Private Const PREFIX_SHOP As String = "SHOP-"
Private Const PREFIX_EHEMALIG As String = "EX-"
Private Const PREFIX_SONSTIGE As String = "SONSTIGE-"

' EntityRole-Werte
Private Const ROLE_MITGLIED_MIT_PACHT As String = "MITGLIED_MIT_PACHT"
Private Const ROLE_MITGLIED_OHNE_PACHT As String = "MITGLIED_OHNE_PACHT"
Private Const ROLE_EHEMALIGES_MITGLIED As String = "EHEMALIGES_MITGLIED"
Private Const ROLE_VERSORGER As String = "VERSORGER"
Private Const ROLE_BANK As String = "BANK"
Private Const ROLE_SHOP As String = "SHOP"
Private Const ROLE_SONSTIGE As String = "SONSTIGE"

' ===============================================================
' HILFSFUNKTION: Prueft ob Role eine Parzelle haben darf
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
' HILFSFUNKTION: Normalisiert Role-String fuer Vergleiche
' ===============================================================
Private Function NormalisiereRoleString(ByVal roleString As String) As String
    Dim result As String
    result = UCase(Trim(roleString))
    result = Replace(result, " ", "_")
    result = Replace(result, "-", "_")
    NormalisiereRoleString = result
End Function

' ===============================================================
' HILFSFUNKTION: Spaltenummer zu Buchstabe
' ===============================================================
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

' ===============================================================
' HILFSFUNKTION: Erzeugt eine neue GUID
' ===============================================================
Private Function CreateGUID() As String
    On Error Resume Next
    CreateGUID = mod_Mitglieder_UI.CreateGUID_Public()
    If Err.Number <> 0 Or Len(CreateGUID) = 0 Then
        Randomize
        CreateGUID = Format(Now, "YYYYMMDDHHMMSS") & "-" & _
                     Format(Int(Rnd * 10000), "0000") & "-" & _
                     Format(Int(Rnd * 10000), "0000")
    End If
    On Error GoTo 0
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
                        "Bereits vorhanden: " & anzahlBereitsVorhanden & vbCrLf & vbCrLf & _
                        "Moechten Sie jetzt die automatische Zuordnung starten?", _
                        vbYesNo + vbQuestion, "IBAN-Import erfolgreich")
        
        If antwort = vbYes Then
            Call AktualisiereAlleEntityKeys
        End If
    Else
        MsgBox "Keine neuen IBANs gefunden!", vbInformation
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
    
    lastRow = wsD.Cells(wsD.Rows.Count, EK_COL_IBAN).End(xlUp).Row
    If lastRow < EK_START_ROW Then lastRow = EK_START_ROW
    
    ' EntityRole-DropDown einrichten (dynamisch aus AD)
    Call SetupEntityRoleDropdowns(wsD, lastRow)
    
    ' Parzellen-DropDown einrichten (dynamisch aus F)
    Call SetupParzellenDropdowns(wsD, lastRow)
    
    For r = EK_START_ROW To lastRow
        iban = Trim(wsD.Cells(r, EK_COL_IBAN).value)
        kontoname = Trim(wsD.Cells(r, EK_COL_KONTONAME).value)
        
        currentEntityKey = Trim(wsD.Cells(r, EK_COL_ENTITYKEY).value)
        currentZuordnung = Trim(wsD.Cells(r, EK_COL_ZUORDNUNG).value)
        currentParzelle = Trim(wsD.Cells(r, EK_COL_PARZELLE).value)
        currentRole = Trim(wsD.Cells(r, EK_COL_ROLE).value)
        currentDebug = Trim(wsD.Cells(r, EK_COL_DEBUG).value)
        
        If iban = "" And kontoname = "" Then GoTo NextRow
        
        If HatBereitsGueltigeDaten(currentEntityKey, currentZuordnung, currentRole) Then
            zeilenUnveraendert = zeilenUnveraendert + 1
            If currentRole <> "" Then
                Call SetzeAmpelFarbe(wsD, r, 1)
            End If
            GoTo NextRow
        End If
        
        zeilenNeu = zeilenNeu + 1
        
        Set mitgliederGefunden = SucheMitgliederZuKontoname(kontoname, wsM, wsH)
        
        Call GeneriereEntityKeyUndZuordnung(mitgliederGefunden, kontoname, wsM, _
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
               "Neue Zeilen: " & zeilenNeu & vbCrLf & _
               "Unveraendert: " & zeilenUnveraendert, vbInformation
    End If
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler: " & Err.Description, vbCritical
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
    
    If entityKey <> "" Then
        If Not IsNumeric(entityKey) Then
            HatBereitsGueltigeDaten = True
            Exit Function
        End If
    End If
    
    If zuordnung <> "" And role <> "" Then
        HatBereitsGueltigeDaten = True
    End If
End Function

' ===============================================================
' HILFSPROZEDUR: Zeigt Hinweis fuer Zeilen mit Eingriff
' ===============================================================
Private Sub ZeigeEingriffsHinweis(ByRef ws As Worksheet, ByRef zeilenRot As Collection, _
                                   ByRef zeilenGelb As Collection, _
                                   ByVal zeilenNeu As Long, ByVal zeilenUnveraendert As Long)
    Dim msg As String
    Dim antwort As VbMsgBoxResult
    Dim ersteZeile As Long
    
    msg = "EntityKey-Aktualisierung abgeschlossen!" & vbCrLf & vbCrLf
    msg = msg & "Neue Zeilen: " & zeilenNeu & vbCrLf
    msg = msg & "Unveraendert: " & zeilenUnveraendert & vbCrLf & vbCrLf
    
    If zeilenRot.Count > 0 Then
        msg = msg & "ROT: " & zeilenRot.Count & " - Manuelle Zuordnung!" & vbCrLf
    End If
    
    If zeilenGelb.Count > 0 Then
        msg = msg & "GELB: " & zeilenGelb.Count & " - Bitte pruefen!" & vbCrLf
    End If
    
    msg = msg & vbCrLf & "Zur ersten Zeile springen?"
    
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

'--- 2. Teil ---

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
    
    If mitgliederExakt.Count = 0 Then
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
        
        If mitgliederNurNachname.Count > 0 Then
            outDebugInfo = "NUR NACHNAME - Bitte pruefen!"
            outAmpelStatus = 2
            For i = 1 To mitgliederNurNachname.Count
                mitgliedInfo = mitgliederNurNachname(i)
                outDebugInfo = outDebugInfo & vbLf & "  ? " & mitgliedInfo(1) & ", " & mitgliedInfo(2)
            Next i
            Exit Sub
        Else
            outDebugInfo = "KEIN MITGLIED - Manuelle Zuordnung!"
            outAmpelStatus = 3
            Exit Sub
        End If
    End If
    
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
            outEntityKey = PREFIX_EHEMALIG & CStr(mitgliedInfo(0))
            outZuordnung = mitgliedInfo(1) & ", " & mitgliedInfo(2)
            outParzellen = mitgliedInfo(3)
            outEntityRole = ROLE_EHEMALIGES_MITGLIED
            outDebugInfo = "Ehemaliges Mitglied"
            outAmpelStatus = 1
        Else
            memberIDs = ""
            For Each key In uniqueMemberIDs.Keys
                If memberIDs <> "" Then memberIDs = memberIDs & "_"
                memberIDs = memberIDs & key
            Next key
            outEntityKey = PREFIX_SHARE & PREFIX_EHEMALIG & memberIDs
            outEntityRole = ROLE_EHEMALIGES_MITGLIED
            outDebugInfo = "Ehem. Gemeinschaftskonto"
            outAmpelStatus = 1
            
            Dim bereitsHinzu As Object
            Set bereitsHinzu = CreateObject("Scripting.Dictionary")
            For i = 1 To mitgliederExakt.Count
                mitgliedInfo = mitgliederExakt(i)
                If Not bereitsHinzu.Exists(CStr(mitgliedInfo(0))) Then
                    bereitsHinzu.Add CStr(mitgliedInfo(0)), True
                    If outZuordnung <> "" Then outZuordnung = outZuordnung & vbLf
                    outZuordnung = outZuordnung & mitgliedInfo(1) & ", " & mitgliedInfo(2)
                    If InStr(outParzellen, CStr(mitgliedInfo(3))) = 0 Then
                        If outParzellen <> "" Then outParzellen = outParzellen & vbLf
                        outParzellen = outParzellen & CStr(mitgliedInfo(3))
                    End If
                End If
            Next i
        End If
        Exit Sub
    End If
    
    If uniqueMemberIDs.Count = 1 Then
        For i = 1 To mitgliederExakt.Count
            mitgliedInfo = mitgliederExakt(i)
            If mitgliedInfo(6) = False Then
                outEntityKey = CStr(mitgliedInfo(0))
                outZuordnung = mitgliedInfo(1) & ", " & mitgliedInfo(2)
                outEntityRole = ErmittleEntityRoleVonFunktion(CStr(mitgliedInfo(4)))
                outDebugInfo = "Eindeutiger Treffer"
                outAmpelStatus = 1
                Exit For
            End If
        Next i
        outParzellen = HoleAlleParzellen(CStr(mitgliedInfo(0)), wsM)
        Exit Sub
    End If
    
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
        
        Dim bereitsHinzugefuegt As Object
        Set bereitsHinzugefuegt = CreateObject("Scripting.Dictionary")
        
        For i = 1 To mitgliederExakt.Count
            mitgliedInfo = mitgliederExakt(i)
            If mitgliedInfo(6) = False Then
                If Not bereitsHinzugefuegt.Exists(CStr(mitgliedInfo(0))) Then
                    bereitsHinzugefuegt.Add CStr(mitgliedInfo(0)), True
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
' HILFSFUNKTION: Extrahiert Anzeigename
' ===============================================================
Private Function ExtrahiereAnzeigeName(ByVal zuordnung As String) As String
    Dim pos As Long
    If Len(zuordnung) = 0 Then
        ExtrahiereAnzeigeName = ""
        Exit Function
    End If
    pos = InStr(zuordnung, " (")
    If pos > 0 Then
        ExtrahiereAnzeigeName = Left(zuordnung, pos - 1)
    Else
        ExtrahiereAnzeigeName = zuordnung
    End If
End Function

' ===============================================================
' HILFSFUNKTION: Holt alle Parzellen eines Mitglieds
' ===============================================================
Private Function HoleAlleParzellen(ByVal mitgliedID As String, ByRef wsM As Worksheet) As String
    Dim lastRow As Long, i As Long
    Dim parzellen As String
    Dim currentID As String, currentParzelle As String
    
    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_MEMBER_ID).End(xlUp).Row
    parzellen = ""
    
    For i = M_START_ROW To lastRow
        currentID = Trim(CStr(wsM.Cells(i, M_COL_MEMBER_ID).value))
        If currentID = mitgliedID Then
            currentParzelle = Trim(CStr(wsM.Cells(i, M_COL_PARZELLE).value))
            If Len(currentParzelle) > 0 Then
                If Len(parzellen) > 0 Then
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

' ===============================================================
' HILFSFUNKTION: Ermittelt EntityRole von Funktion
' ===============================================================
Private Function ErmittleEntityRoleVonFunktion(ByVal funktion As String) As String
    Dim func As String
    func = UCase(Trim(funktion))
    
    Select Case func
        Case "1. VORSITZENDER", "2. VORSITZENDER", "KASSENWART", _
             "SCHRIFTFUEHRER", "BEISITZER", "FACHBERATER"
            ErmittleEntityRoleVonFunktion = ROLE_MITGLIED_MIT_PACHT
        Case "EHRENMITGLIED"
            ErmittleEntityRoleVonFunktion = ROLE_MITGLIED_OHNE_PACHT
        Case Else
            ErmittleEntityRoleVonFunktion = ROLE_MITGLIED_MIT_PACHT
    End Select
End Function

'--- Teil 3 ---

' ===============================================================
' HILFSFUNKTION: Prueft ob Kontoname auf Versorger hindeutet
' ===============================================================
Private Function IstVersorger(ByVal kontoname As String) As Boolean
    Dim name As String
    name = UCase(kontoname)
    
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
    If InStr(name, "WASSER") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "ZWA") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "ABWASSER") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "ZWECKVERBAND") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "GASAG") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "GASVERSORGUNG") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "VERSICHERUNG") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "ALLIANZ") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "ERGO") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "HDI") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "AXA") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "STADTVERWALTUNG") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "FINANZAMT") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "GEMEINDE") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "LANDKREIS") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "VERBAND") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "TELEKOM") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "VODAFONE") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "ENTSORGUNG") > 0 Then IstVersorger = True: Exit Function
    If InStr(name, "ABFALL") > 0 Then IstVersorger = True: Exit Function
    
    IstVersorger = False
End Function

' ===============================================================
' HILFSFUNKTION: Prueft ob Kontoname auf Bank hindeutet
' ===============================================================
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
    If InStr(name, "TARGOBANK") > 0 Then IstBank = True: Exit Function
    If InStr(name, "ING") > 0 Then IstBank = True: Exit Function
    If InStr(name, "COMDIRECT") > 0 Then IstBank = True: Exit Function
    If InStr(name, "DKB") > 0 Then IstBank = True: Exit Function
    If InStr(name, "N26") > 0 Then IstBank = True: Exit Function
    If InStr(name, "SANTANDER") > 0 Then IstBank = True: Exit Function
    If InStr(name, "ZINSEN") > 0 Then IstBank = True: Exit Function
    If InStr(name, "KONTOFUEHRUNG") > 0 Then IstBank = True: Exit Function
    If InStr(name, "BANKGEBUEHR") > 0 Then IstBank = True: Exit Function
    
    IstBank = False
End Function

' ===============================================================
' HILFSFUNKTION: Prueft ob Kontoname auf Shop hindeutet
' ===============================================================
Private Function IstShop(ByVal kontoname As String) As Boolean
    Dim name As String
    name = UCase(kontoname)
    
    If InStr(name, "BAUHAUS") > 0 Then IstShop = True: Exit Function
    If InStr(name, "OBI") > 0 Then IstShop = True: Exit Function
    If InStr(name, "HORNBACH") > 0 Then IstShop = True: Exit Function
    If InStr(name, "HAGEBAU") > 0 Then IstShop = True: Exit Function
    If InStr(name, "TOOM") > 0 Then IstShop = True: Exit Function
    If InStr(name, "BAUMARKT") > 0 Then IstShop = True: Exit Function
    If InStr(name, "GARTENCENTER") > 0 Then IstShop = True: Exit Function
    If InStr(name, "DEHNER") > 0 Then IstShop = True: Exit Function
    If InStr(name, "MEDIAMARKT") > 0 Then IstShop = True: Exit Function
    If InStr(name, "SATURN") > 0 Then IstShop = True: Exit Function
    If InStr(name, "CONRAD") > 0 Then IstShop = True: Exit Function
    If InStr(name, "IKEA") > 0 Then IstShop = True: Exit Function
    If InStr(name, "POCO") > 0 Then IstShop = True: Exit Function
    If InStr(name, "AMAZON") > 0 Then IstShop = True: Exit Function
    If InStr(name, "EBAY") > 0 Then IstShop = True: Exit Function
    If InStr(name, "PAYPAL") > 0 Then IstShop = True: Exit Function
    If InStr(name, "REWE") > 0 Then IstShop = True: Exit Function
    If InStr(name, "EDEKA") > 0 Then IstShop = True: Exit Function
    If InStr(name, "LIDL") > 0 Then IstShop = True: Exit Function
    If InStr(name, "ALDI") > 0 Then IstShop = True: Exit Function
    If InStr(name, "KAUFLAND") > 0 Then IstShop = True: Exit Function
    
    IstShop = False
End Function

' ===============================================================
' HILFSPROZEDUR: Setzt Ampelfarbe - 3 Parameter Version
' ===============================================================
Private Sub SetzeAmpelFarbe(ByRef ws As Worksheet, ByVal zeile As Long, ByVal ampelStatus As Long)
    Dim zelle As Range
    Set zelle = ws.Cells(zeile, EK_COL_DEBUG)
    
    Select Case ampelStatus
        Case 1
            zelle.Interior.color = RGB(146, 208, 80)
        Case 2
            zelle.Interior.color = RGB(255, 255, 0)
        Case 3
            zelle.Interior.color = RGB(255, 0, 0)
        Case Else
            zelle.Interior.ColorIndex = xlNone
    End Select
End Sub

' ===============================================================
' HILFSPROZEDUR: Richtet EntityRole-DropDowns ein (mehrere Zeilen)
' ===============================================================
Private Sub SetupEntityRoleDropdowns(ByRef ws As Worksheet, ByVal lastRow As Long)
    Dim wsDaten As Worksheet
    Dim lastRoleRow As Long
    Dim sourceRange As String
    Dim i As Long
    Dim zielZelle As Range
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    
    lastRoleRow = wsDaten.Cells(wsDaten.Rows.Count, DATA_COL_DD_ENTITYROLE).End(xlUp).Row
    If lastRoleRow < 4 Then lastRoleRow = 10
    
    sourceRange = "=" & WS_DATEN & "!$" & ColLetter(DATA_COL_DD_ENTITYROLE) & "$4:$" & _
                  ColLetter(DATA_COL_DD_ENTITYROLE) & "$" & lastRoleRow
    
    For i = EK_START_ROW To lastRow
        Set zielZelle = ws.Cells(i, EK_COL_ROLE)
        
        On Error Resume Next
        zielZelle.Validation.Delete
        On Error GoTo 0
        
        With zielZelle.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:=sourceRange
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = True
        End With
    Next i
End Sub

' ===============================================================
' HILFSPROZEDUR: Richtet Parzellen-DropDowns ein (mehrere Zeilen)
' ===============================================================
Private Sub SetupParzellenDropdowns(ByRef ws As Worksheet, ByVal lastRow As Long)
    Dim wsDaten As Worksheet
    Dim lastParzelleRow As Long
    Dim sourceRange As String
    Dim i As Long
    Dim zielZelle As Range
    Dim roleValue As String
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    
    lastParzelleRow = wsDaten.Cells(wsDaten.Rows.Count, DATA_COL_DD_PARZELLE).End(xlUp).Row
    If lastParzelleRow < 4 Then lastParzelleRow = 100
    
    sourceRange = "=" & WS_DATEN & "!$" & ColLetter(DATA_COL_DD_PARZELLE) & "$4:$" & _
                  ColLetter(DATA_COL_DD_PARZELLE) & "$" & lastParzelleRow
    
    For i = EK_START_ROW To lastRow
        roleValue = Trim(ws.Cells(i, EK_COL_ROLE).value)
        Set zielZelle = ws.Cells(i, EK_COL_PARZELLE)
        
        On Error Resume Next
        zielZelle.Validation.Delete
        On Error GoTo 0
        
        If DarfParzelleHaben(roleValue) Then
            With zielZelle.Validation
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, _
                     Operator:=xlBetween, Formula1:=sourceRange
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = False
                .ShowError = True
            End With
            zielZelle.Interior.ColorIndex = xlNone
        Else
            zielZelle.Interior.color = RGB(217, 217, 217)
        End If
    Next i
End Sub

' ===============================================================
' HILFSPROZEDUR: Formatiert die EntityKey-Tabelle
' ===============================================================
Private Sub FormatiereEntityKeyTabelle(ByRef ws As Worksheet, ByVal lastRow As Long)
    Dim headerRange As Range
    Dim dataRange As Range
    
    If lastRow < EK_START_ROW Then lastRow = EK_START_ROW
    
    Set headerRange = ws.Range(ws.Cells(EK_START_ROW - 1, EK_COL_ENTITYKEY), _
                               ws.Cells(EK_START_ROW - 1, EK_COL_DEBUG))
    With headerRange
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.color = RGB(68, 114, 196)
        .Font.color = RGB(255, 255, 255)
    End With
    
    Set dataRange = ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                             ws.Cells(lastRow, EK_COL_DEBUG))
    
    With dataRange
        .VerticalAlignment = xlTop
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    ws.Columns(EK_COL_ENTITYKEY).HorizontalAlignment = xlLeft
    ws.Columns(EK_COL_ENTITYKEY).EntireColumn.AutoFit
    
    ws.Columns(EK_COL_IBAN).HorizontalAlignment = xlLeft
    ws.Columns(EK_COL_IBAN).ColumnWidth = 28
    
    ws.Columns(EK_COL_KONTONAME).HorizontalAlignment = xlLeft
    ws.Columns(EK_COL_KONTONAME).EntireColumn.AutoFit
    If ws.Columns(EK_COL_KONTONAME).ColumnWidth > 35 Then
        ws.Columns(EK_COL_KONTONAME).ColumnWidth = 35
    End If
    
    ws.Columns(EK_COL_ZUORDNUNG).HorizontalAlignment = xlLeft
    ws.Columns(EK_COL_ZUORDNUNG).EntireColumn.AutoFit
    If ws.Columns(EK_COL_ZUORDNUNG).ColumnWidth > 40 Then
        ws.Columns(EK_COL_ZUORDNUNG).ColumnWidth = 40
    End If
    
    ws.Columns(EK_COL_PARZELLE).HorizontalAlignment = xlCenter
    ws.Columns(EK_COL_PARZELLE).ColumnWidth = 10
    
    ws.Columns(EK_COL_ROLE).HorizontalAlignment = xlLeft
    ws.Columns(EK_COL_ROLE).EntireColumn.AutoFit
    If ws.Columns(EK_COL_ROLE).ColumnWidth < 20 Then
        ws.Columns(EK_COL_ROLE).ColumnWidth = 20
    End If
    
    ws.Columns(EK_COL_DEBUG).HorizontalAlignment = xlCenter
    ws.Columns(EK_COL_DEBUG).ColumnWidth = 8
End Sub

' ===============================================================
' OEFFENTLICHE PROZEDUR: Wird nach CSV-Import aufgerufen
' ===============================================================
Public Sub NachCSVImport_EntityKeysAktualisieren()
    Call AktualisiereAlleEntityKeys
End Sub

' ===============================================================
' OEFFENTLICHE PROZEDUR: Entfernt ueberfluessige Rahmen
' ===============================================================
Public Sub EntferneUeberfluesstigeRahmen()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cleanRange As Range
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, EK_COL_ENTITYKEY).End(xlUp).Row
    
    Set cleanRange = ws.Range(ws.Cells(lastRow + 1, EK_COL_ENTITYKEY), _
                              ws.Cells(lastRow + 100, EK_COL_DEBUG))
    
    With cleanRange
        .Borders.LineStyle = xlNone
        .Interior.ColorIndex = xlNone
    End With
End Sub


' ===============================================================
' OEFFENTLICHE PROZEDUR: Debug-Ausgabe der Konstanten
' ===============================================================
Public Sub DebugKonstanten()
    Debug.Print "=== EntityKey Manager Konstanten ==="
    Debug.Print "EK_COL_ENTITYKEY: " & EK_COL_ENTITYKEY & " (" & ColLetter(EK_COL_ENTITYKEY) & ")"
    Debug.Print "EK_COL_IBAN: " & EK_COL_IBAN & " (" & ColLetter(EK_COL_IBAN) & ")"
    Debug.Print "EK_COL_KONTONAME: " & EK_COL_KONTONAME & " (" & ColLetter(EK_COL_KONTONAME) & ")"
    Debug.Print "EK_COL_ZUORDNUNG: " & EK_COL_ZUORDNUNG & " (" & ColLetter(EK_COL_ZUORDNUNG) & ")"
    Debug.Print "EK_COL_PARZELLE: " & EK_COL_PARZELLE & " (" & ColLetter(EK_COL_PARZELLE) & ")"
    Debug.Print "EK_COL_ROLE: " & EK_COL_ROLE & " (" & ColLetter(EK_COL_ROLE) & ")"
    Debug.Print "EK_COL_DEBUG: " & EK_COL_DEBUG & " (" & ColLetter(EK_COL_DEBUG) & ")"
    Debug.Print "DATA_COL_DD_ENTITYROLE: " & DATA_COL_DD_ENTITYROLE & " (" & ColLetter(DATA_COL_DD_ENTITYROLE) & ")"
    Debug.Print "DATA_COL_DD_PARZELLE: " & DATA_COL_DD_PARZELLE & " (" & ColLetter(DATA_COL_DD_PARZELLE) & ")"
    Debug.Print "=== Ende Konstanten ==="

End Sub

' ===============================================================
' OEFFENTLICHE PROZEDUR: Verarbeitet manuelle Role-Aenderung
' ===============================================================
Public Sub VerarbeiteManuelleRoleAenderung(ByVal zeile As Long)
    Dim ws As Worksheet
    Dim neueRole As String
    Dim alterKey As String
    Dim neuerKey As String
    Dim kontoname As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    neueRole = Trim(CStr(ws.Cells(zeile, EK_COL_ROLE).value))
    alterKey = Trim(CStr(ws.Cells(zeile, EK_COL_ENTITYKEY).value))
    kontoname = Trim(CStr(ws.Cells(zeile, EK_COL_KONTONAME).value))
    
    Select Case NormalisiereRoleString(neueRole)
        Case "VERSORGER"
            If Left(alterKey, 5) <> "VERS-" Then
                neuerKey = PREFIX_VERSORGER & CreateGUID()
            End If
        Case "BANK"
            If Left(alterKey, 5) <> "BANK-" Then
                neuerKey = PREFIX_BANK & CreateGUID()
            End If
        Case "SHOP"
            If Left(alterKey, 5) <> "SHOP-" Then
                neuerKey = PREFIX_SHOP & CreateGUID()
            End If
        Case "EHEMALIGES_MITGLIED"
            If Left(alterKey, 3) <> "EX-" Then
                neuerKey = PREFIX_EHEMALIG & CreateGUID()
            End If
        Case "SONSTIGE"
            If Left(alterKey, 9) <> "SONSTIGE-" Then
                neuerKey = PREFIX_SONSTIGE & CreateGUID()
            End If
    End Select
    
    If neuerKey <> "" Then
        ws.Cells(zeile, EK_COL_ENTITYKEY).value = neuerKey
    End If
    
    If Trim(ws.Cells(zeile, EK_COL_ZUORDNUNG).value) = "" Then
        ws.Cells(zeile, EK_COL_ZUORDNUNG).value = ExtrahiereAnzeigeName(kontoname)
    End If
    
    If Not DarfParzelleHaben(neueRole) Then
        ws.Cells(zeile, EK_COL_PARZELLE).value = ""
        ws.Cells(zeile, EK_COL_PARZELLE).Interior.color = RGB(217, 217, 217)
    Else
        ws.Cells(zeile, EK_COL_PARZELLE).Interior.ColorIndex = xlNone
    End If
    
    Call SetzeAmpelFarbe(ws, zeile, 1)
End Sub

' ===============================================================
' OEFFENTLICHE PROZEDUR: Formatiert eine einzelne EntityKey-Zeile
' ===============================================================
Public Sub FormatiereEntityKeyZeile(ByVal zeile As Long)
    Dim ws As Worksheet
    Dim roleValue As String
    Dim zeilenRange As Range
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    Set zeilenRange = ws.Range(ws.Cells(zeile, EK_COL_ENTITYKEY), ws.Cells(zeile, EK_COL_DEBUG))
    
    zeilenRange.VerticalAlignment = xlTop
    
    With zeilenRange
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    
    ws.Cells(zeile, EK_COL_ROLE).HorizontalAlignment = xlLeft
End Sub
