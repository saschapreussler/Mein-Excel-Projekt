Attribute VB_Name = "mod_EntityKey_Manager"
Option Explicit

' ***************************************************************
' MODUL: mod_EntityKey_Manager
' ZWECK: Verwaltung und Zuordnung von EntityKeys für Bankverkehr
' VERSION: 3.1 - 04.02.2026
' SICHERHEIT:
'   - KEINE automatische Formatierung
'   - KEINE automatische Sortierung
'   - NUR Spalten R-X werden befüllt
'   - Spalten Y-AE werden NIEMALS berührt
'   - Bestehende Formatierung bleibt erhalten
' ***************************************************************

' ===============================================================
' KONSTANTEN (lokal - zusätzlich zu mod_Const)
' ===============================================================
Private Const EK_ROLE_DROPDOWN_COL As Long = 30  ' Spalte AD für Role-Dropdown-Quelle

' EntityKey Präfixe
Private Const PREFIX_SHARE As String = "SHARE-"
Private Const PREFIX_VERSORGER As String = "VERS-"
Private Const PREFIX_BANK As String = "BANK-"
Private Const PREFIX_SHOP As String = "SHOP-"
Private Const PREFIX_EHEMALIG As String = "EX-"
Private Const PREFIX_SONSTIGE As String = "SONST-"

' EntityRole Werte - OHNE UNTERSTRICHE (mit Leerzeichen)
Private Const ROLE_MITGLIED As String = "MITGLIED"
Private Const ROLE_MITGLIED_MIT_PACHT As String = "MITGLIED MIT PACHT"
Private Const ROLE_MITGLIED_OHNE_PACHT As String = "MITGLIED OHNE PACHT"
Private Const ROLE_EHEMALIGES_MITGLIED As String = "EHEMALIGES MITGLIED"
Private Const ROLE_VORSTAND As String = "VORSTAND"
Private Const ROLE_EHRENMITGLIED As String = "EHRENMITGLIED"
Private Const ROLE_VERSORGER As String = "VERSORGER"
Private Const ROLE_BANK As String = "BANK"
Private Const ROLE_SHOP As String = "SHOP"
Private Const ROLE_SONSTIGE As String = "SONSTIGE"

' ===============================================================
' HILFSFUNKTION: Prüft ob Role eine Parzelle haben darf
' ===============================================================
Private Function DarfParzelleHaben(ByVal role As String) As Boolean
    Dim normRole As String
    
    If Trim(role) = "" Then
        DarfParzelleHaben = False
        Exit Function
    End If
    
    normRole = UCase(Trim(role))
    
    If InStr(normRole, "MITGLIED") > 0 Then
        DarfParzelleHaben = True
    ElseIf InStr(normRole, "VORSTAND") > 0 Then
        DarfParzelleHaben = True
    ElseIf InStr(normRole, "EHRENMITGLIED") > 0 Then
        DarfParzelleHaben = True
    ElseIf normRole = "SONSTIGE" Then
        DarfParzelleHaben = True
    Else
        DarfParzelleHaben = False
    End If
End Function

' ===============================================================
' HILFSFUNKTION: Entfernt mehrfache Leerzeichen
' ===============================================================
Private Function EntferneMehrfacheLeerzeichen(ByVal s As String) As String
    Dim result As String
    result = s
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    EntferneMehrfacheLeerzeichen = Trim(result)
End Function

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
' ÖFFENTLICHE PROZEDUR: Importiert IBANs aus Bankkonto
' SICHER: Ändert NUR Spalten R-T, keine Formatierung
' ===============================================================
Public Sub ImportiereIBANsAusBankkonto()
    
    Dim wsBK As Worksheet
    Dim wsD As Worksheet
    Dim dictIBANs As Object
    Dim dictExisting As Object
    Dim r As Long
    Dim lastRowBK As Long
    Dim lastRowD As Long
    Dim nextRowD As Long
    Dim currentIBAN As String
    Dim currentKontoName As String
    Dim currentDatum As Variant
    Dim ibanKey As Variant
    Dim existingKontoName As String
    Dim anzahlNeu As Long
    Dim anzahlAktualisiert As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set dictIBANs = CreateObject("Scripting.Dictionary")
    Set dictExisting = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    wsD.Unprotect PASSWORD:=PASSWORD
    wsBK.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    anzahlNeu = 0
    anzahlAktualisiert = 0
    
    ' Bestehende IBANs sammeln
    lastRowD = wsD.Cells(wsD.Rows.count, EK_COL_IBAN).End(xlUp).Row
    
    If lastRowD >= EK_START_ROW Then
        For r = EK_START_ROW To lastRowD
            currentIBAN = NormalisiereIBAN(wsD.Cells(r, EK_COL_IBAN).value)
            If currentIBAN <> "" Then
                If Not dictExisting.Exists(currentIBAN) Then
                    dictExisting.Add currentIBAN, r
                End If
            End If
        Next r
    End If
    
    ' IBANs aus Bankkonto sammeln
    lastRowBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    For r = BK_START_ROW To lastRowBK
        currentDatum = wsBK.Cells(r, BK_COL_DATUM).value
        
        If Not IsEmpty(currentDatum) And currentDatum <> "" Then
            currentIBAN = NormalisiereIBAN(wsBK.Cells(r, BK_COL_IBAN).value)
            currentKontoName = EntferneMehrfacheLeerzeichen(Trim(wsBK.Cells(r, BK_COL_NAME).value))
            
            If currentIBAN <> "" And currentIBAN <> "N.A." And Len(currentIBAN) >= 15 Then
                If Not dictIBANs.Exists(currentIBAN) Then
                    dictIBANs.Add currentIBAN, currentKontoName
                End If
            End If
        End If
    Next r
    
    ' Neue IBANs einfügen
    If lastRowD < EK_START_ROW Then
        nextRowD = EK_START_ROW
    Else
        nextRowD = lastRowD + 1
    End If
    
    For Each ibanKey In dictIBANs.Keys
        currentIBAN = CStr(ibanKey)
        currentKontoName = EntferneMehrfacheLeerzeichen(dictIBANs(ibanKey))
        
        If Not dictExisting.Exists(currentIBAN) Then
            ' NUR Spalten S und T befüllen - KEINE Formatierung!
            wsD.Cells(nextRowD, EK_COL_IBAN).value = currentIBAN
            wsD.Cells(nextRowD, EK_COL_KONTONAME).value = currentKontoName
            anzahlNeu = anzahlNeu + 1
            nextRowD = nextRowD + 1
        End If
    Next ibanKey
    
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "IBAN-Import abgeschlossen!" & vbCrLf & vbCrLf & _
           "Neue IBANs: " & anzahlNeu, vbInformation, "Import"
    
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
' HAUPTPROZEDUR: Aktualisiert alle EntityKeys
' SICHER: NUR Spalten R-X, KEINE Formatierung, KEINE Sortierung
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
    Dim newEntityKey As String
    Dim zuordnung As String
    Dim parzellen As String
    Dim entityRole As String
    Dim debugInfo As String
    Dim ampelStatus As Long
    Dim mitgliederGefunden As Collection
    Dim zeilenNeu As Long
    Dim zeilenUnveraendert As Long
    Dim zeilenProbleme As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    
    zeilenNeu = 0
    zeilenUnveraendert = 0
    zeilenProbleme = 0
    
    On Error Resume Next
    wsD.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    lastRow = wsD.Cells(wsD.Rows.count, EK_COL_IBAN).End(xlUp).Row
    If lastRow < EK_START_ROW Then lastRow = EK_START_ROW
    
    For r = EK_START_ROW To lastRow
        iban = Trim(wsD.Cells(r, EK_COL_IBAN).value)
        kontoname = EntferneMehrfacheLeerzeichen(Trim(wsD.Cells(r, EK_COL_KONTONAME).value))
        
        ' Doppelte Leerzeichen in Spalte T bereinigen (nur Wert, keine Formatierung!)
        If wsD.Cells(r, EK_COL_KONTONAME).value <> kontoname Then
            wsD.Cells(r, EK_COL_KONTONAME).value = kontoname
        End If
        
        currentEntityKey = Trim(wsD.Cells(r, EK_COL_ENTITYKEY).value)
        currentZuordnung = Trim(wsD.Cells(r, EK_COL_ZUORDNUNG).value)
        currentParzelle = Trim(wsD.Cells(r, EK_COL_PARZELLE).value)
        currentRole = Trim(wsD.Cells(r, EK_COL_ROLE).value)
        
        If iban = "" And kontoname = "" Then GoTo NextRow
        
        ' WICHTIG: Bereits manuell zugeordnete Zeilen NICHT überschreiben
        If HatBereitsGueltigeDaten(currentEntityKey, currentZuordnung, currentRole) Then
            zeilenUnveraendert = zeilenUnveraendert + 1
            GoTo NextRow
        End If
        
        zeilenNeu = zeilenNeu + 1
        
        ' Suche Mitglieder im Kontonamen
        Set mitgliederGefunden = SucheMitgliederZuKontoname(kontoname, wsM, wsH)
        
        ' Generiere EntityKey und Zuordnung
        Call GeneriereEntityKeyUndZuordnung(mitgliederGefunden, kontoname, wsM, _
                                             newEntityKey, zuordnung, parzellen, entityRole, debugInfo, ampelStatus)
        
        ' NUR leere Zellen befüllen - NIEMALS überschreiben!
        If currentEntityKey = "" And newEntityKey <> "" Then
            wsD.Cells(r, EK_COL_ENTITYKEY).value = newEntityKey
        End If
        
        If currentZuordnung = "" And zuordnung <> "" Then
            wsD.Cells(r, EK_COL_ZUORDNUNG).value = zuordnung
        End If
        
        If currentParzelle = "" And parzellen <> "" And DarfParzelleHaben(entityRole) Then
            wsD.Cells(r, EK_COL_PARZELLE).value = parzellen
        End If
        
        If currentRole = "" And entityRole <> "" Then
            wsD.Cells(r, EK_COL_ROLE).value = entityRole
        End If
        
        ' Debug-Info nur wenn leer
        If Trim(wsD.Cells(r, EK_COL_DEBUG).value) = "" Then
            wsD.Cells(r, EK_COL_DEBUG).value = debugInfo
        End If
        
        If ampelStatus = 3 Then zeilenProbleme = zeilenProbleme + 1
        
NextRow:
    Next r
    
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "EntityKey-Aktualisierung abgeschlossen!" & vbCrLf & vbCrLf & _
           "Neue Zeilen verarbeitet: " & zeilenNeu & vbCrLf & _
           "Bestehende Zeilen unverändert: " & zeilenUnveraendert & vbCrLf & _
           "Zeilen mit Problemen (ROT): " & zeilenProbleme & vbCrLf & vbCrLf & _
           "HINWEIS: Formatierung wurde NICHT geändert.", vbInformation, "Aktualisierung"
    
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
' HILFSFUNKTION: Sucht Mitglieder im Kontonamen
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
    Dim matchResult As Long
    
    Set SucheMitgliederZuKontoname = gefunden
    
    If kontoname = "" Then Exit Function
    
    kontoNameNorm = NormalisiereStringFuerVergleich(kontoname)
    If kontoNameNorm = "" Then Exit Function
    
    ' === AKTIVE MITGLIEDER ===
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
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
    
    Set SucheMitgliederZuKontoname = gefunden
End Function

' ===============================================================
' HILFSFUNKTION: Prüft Namens-Match
' ===============================================================
Private Function PruefeNamensMatch(ByVal nachname As String, ByVal vorname As String, _
                                     ByVal kontoNameNorm As String) As Long
    
    Dim nachnameNorm As String
    Dim vornameNorm As String
    
    PruefeNamensMatch = 0
    
    nachnameNorm = NormalisiereStringFuerVergleich(nachname)
    vornameNorm = NormalisiereStringFuerVergleich(vorname)
    
    If nachnameNorm = "" Or Len(nachnameNorm) < 3 Then Exit Function
    
    If InStr(kontoNameNorm, nachnameNorm) = 0 Then Exit Function
    
    If vornameNorm <> "" And Len(vornameNorm) >= 2 Then
        If InStr(kontoNameNorm, vornameNorm) > 0 Then
            PruefeNamensMatch = 2  ' Exakt
        Else
            PruefeNamensMatch = 1  ' Nur Nachname
        End If
    Else
        PruefeNamensMatch = 1
    End If
End Function

' ===============================================================
' HILFSFUNKTION: Normalisiert String für Vergleich
' ===============================================================
Private Function NormalisiereStringFuerVergleich(ByVal s As String) As String
    Dim result As String
    
    result = LCase(Trim(s))
    result = Replace(result, ",", " ")
    result = Replace(result, ".", " ")
    result = Replace(result, "-", " ")
    result = Replace(result, "ä", "ae")
    result = Replace(result, "ö", "oe")
    result = Replace(result, "ü", "ue")
    result = Replace(result, "ß", "ss")
    result = Replace(result, "ae", "a")
    result = Replace(result, "oe", "o")
    result = Replace(result, "ue", "u")
    
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    NormalisiereStringFuerVergleich = Trim(result)
End Function

' ===============================================================
' HILFSFUNKTION: Prüft ob MemberID bereits gefunden
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



'--- Ende Teil 1 ---
'--- Anfang Teil 2 ---



' filepath: c:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\Modules\mod_EntityKey_Manager.bas

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
    
    ' Treffer sortieren nach Qualität
    For i = 1 To mitglieder.count
        mitgliedInfo = mitglieder(i)
        If mitgliedInfo(8) = 2 Then
            mitgliederExakt.Add mitgliedInfo
        ElseIf mitgliedInfo(8) = 1 Then
            mitgliederNurNachname.Add mitgliedInfo
        End If
    Next i
    
    ' ============================================================
    ' Fall 1: Keine exakten Treffer
    ' ============================================================
    If mitgliederExakt.count = 0 Then
        
        ' WICHTIG: Erst SHOP prüfen (hat Vorrang vor VERSORGER)
        If IstShop(kontoname) Then
            outEntityKey = PREFIX_SHOP & CreateGUID()
            outEntityRole = ROLE_SHOP
            outZuordnung = ExtrahiereAnzeigeName(kontoname)
            outDebugInfo = "Automatisch als SHOP erkannt"
            outAmpelStatus = 1
            Exit Sub
        End If
        
        If IstVersorger(kontoname) Then
            outEntityKey = PREFIX_VERSORGER & CreateGUID()
            outEntityRole = ROLE_VERSORGER
            outZuordnung = ExtrahiereAnzeigeName(kontoname)
            outDebugInfo = "Automatisch als VERSORGER erkannt"
            outAmpelStatus = 1
            Exit Sub
        End If
        
        If IstBank(kontoname) Then
            outEntityKey = PREFIX_BANK & CreateGUID()
            outEntityRole = ROLE_BANK
            outZuordnung = ExtrahiereAnzeigeName(kontoname)
            outDebugInfo = "Automatisch als BANK erkannt"
            outAmpelStatus = 1
            Exit Sub
        End If
        
        ' Nur Nachname gefunden?
        If mitgliederNurNachname.count > 0 Then
            outDebugInfo = "NUR NACHNAME - Bitte prüfen!"
            outAmpelStatus = 2
            Exit Sub
        End If
        
        ' Nichts gefunden
        outDebugInfo = "KEIN TREFFER - Manuelle Zuordnung"
        outAmpelStatus = 3
        Exit Sub
    End If
    
    ' ============================================================
    ' Fall 2: Exakte Treffer vorhanden
    ' ============================================================
    
    ' Sammle unique MemberIDs
    For i = 1 To mitgliederExakt.count
        mitgliedInfo = mitgliederExakt(i)
        If mitgliedInfo(6) = False Then  ' Nicht ehemalig
            If Not uniqueMemberIDs.Exists(CStr(mitgliedInfo(0))) Then
                uniqueMemberIDs.Add CStr(mitgliedInfo(0)), CStr(mitgliedInfo(0))
            End If
        End If
    Next i
    
    ' Fall 2a: Genau 1 aktives Mitglied
    If uniqueMemberIDs.count = 1 Then
        For i = 1 To mitgliederExakt.count
            mitgliedInfo = mitgliederExakt(i)
            If mitgliedInfo(6) = False Then
                outEntityKey = CStr(mitgliedInfo(0))
                ' Format: "Nachname, Vorname"
                outZuordnung = mitgliedInfo(1) & ", " & mitgliedInfo(2)
                outParzellen = CStr(mitgliedInfo(3))
                outEntityRole = ErmittleEntityRoleVonFunktion(CStr(mitgliedInfo(4)))
                outDebugInfo = "Eindeutiger Treffer"
                outAmpelStatus = 1
                Exit For
            End If
        Next i
        Exit Sub
    End If
    
    ' Fall 2b: Mehrere aktive Mitglieder = GEMEINSCHAFTSKONTO
    If uniqueMemberIDs.count > 1 Then
        memberIDs = ""
        For Each key In uniqueMemberIDs.Keys
            If memberIDs <> "" Then memberIDs = memberIDs & "_"
            memberIDs = memberIDs & key
        Next key
        
        outEntityKey = PREFIX_SHARE & memberIDs
        outEntityRole = ROLE_MITGLIED_MIT_PACHT
        outDebugInfo = "Gemeinschaftskonto - " & uniqueMemberIDs.count & " Personen"
        outAmpelStatus = 1
        
        ' Alle Namen mit "Nachname, Vorname" und vbLf
        Dim bereitsHinzu As Object
        Set bereitsHinzu = CreateObject("Scripting.Dictionary")
        
        For i = 1 To mitgliederExakt.count
            mitgliedInfo = mitgliederExakt(i)
            If mitgliedInfo(6) = False Then
                If Not bereitsHinzu.Exists(CStr(mitgliedInfo(0))) Then
                    bereitsHinzu.Add CStr(mitgliedInfo(0)), True
                    
                    If outZuordnung <> "" Then outZuordnung = outZuordnung & vbLf
                    outZuordnung = outZuordnung & mitgliedInfo(1) & ", " & mitgliedInfo(2)
                    
                    If outParzellen <> "" Then
                        If InStr(outParzellen, CStr(mitgliedInfo(3))) = 0 Then
                            outParzellen = outParzellen & ", " & CStr(mitgliedInfo(3))
                        End If
                    Else
                        outParzellen = CStr(mitgliedInfo(3))
                    End If
                End If
            End If
        Next i
    End If
End Sub

' ===============================================================
' HILFSFUNKTION: Extrahiert Anzeigename
' ===============================================================
Private Function ExtrahiereAnzeigeName(ByVal kontoname As String) As String
    Dim zeilen() As String
    Dim erstesElement As String
    
    If kontoname = "" Then
        ExtrahiereAnzeigeName = ""
        Exit Function
    End If
    
    zeilen = Split(kontoname, vbLf)
    erstesElement = EntferneMehrfacheLeerzeichen(Trim(zeilen(0)))
    
    If Len(erstesElement) > 50 Then
        erstesElement = Left(erstesElement, 50) & "..."
    End If
    
    ExtrahiereAnzeigeName = erstesElement
End Function

' ===============================================================
' HILFSFUNKTION: Ermittelt EntityRole aus Funktion
' ===============================================================
Private Function ErmittleEntityRoleVonFunktion(ByVal funktion As String) As String
    Dim funktionUpper As String
    funktionUpper = UCase(funktion)
    
    If InStr(funktionUpper, "VORSTAND") > 0 Or _
       InStr(funktionUpper, "VORSITZ") > 0 Or _
       InStr(funktionUpper, "KASSIERER") > 0 Or _
       InStr(funktionUpper, "SCHRIFTFÜHRER") > 0 Then
        ErmittleEntityRoleVonFunktion = ROLE_VORSTAND
    ElseIf InStr(funktionUpper, "EHRENMITGLIED") > 0 Then
        ErmittleEntityRoleVonFunktion = ROLE_EHRENMITGLIED
    ElseIf InStr(funktionUpper, "OHNE PACHT") > 0 Then
        ErmittleEntityRoleVonFunktion = ROLE_MITGLIED_OHNE_PACHT
    ElseIf InStr(funktionUpper, "EHEMALIG") > 0 Then
        ErmittleEntityRoleVonFunktion = ROLE_EHEMALIGES_MITGLIED
    Else
        ErmittleEntityRoleVonFunktion = ROLE_MITGLIED_MIT_PACHT
    End If
End Function

' ===============================================================
' IstShop - Prüft ob Kontoname ein Shop ist
' WICHTIG: Wird VOR IstVersorger geprüft!
' ===============================================================
Private Function IstShop(ByVal kontoname As String) As Boolean
    Dim n As String
    n = UCase(Trim(kontoname))
    
    IstShop = False
    If Len(n) = 0 Then Exit Function
    
    ' Supermärkte/Discounter
    If InStr(n, "LIDL") > 0 Then IstShop = True: Exit Function
    If InStr(n, "ALDI") > 0 Then IstShop = True: Exit Function
    If InStr(n, "REWE") > 0 Then IstShop = True: Exit Function
    If InStr(n, "EDEKA") > 0 Then IstShop = True: Exit Function
    If InStr(n, "PENNY") > 0 Then IstShop = True: Exit Function
    If InStr(n, "NETTO") > 0 Then IstShop = True: Exit Function
    If InStr(n, "KAUFLAND") > 0 Then IstShop = True: Exit Function
    
    ' Baumärkte
    If InStr(n, "BAUHAUS") > 0 Then IstShop = True: Exit Function
    If InStr(n, "HORNBACH") > 0 Then IstShop = True: Exit Function
    If InStr(n, "OBI") > 0 Then IstShop = True: Exit Function
    If InStr(n, "HAGEBAU") > 0 Then IstShop = True: Exit Function
    If InStr(n, "TOOM") > 0 Then IstShop = True: Exit Function
    If InStr(n, "HELLWEG") > 0 Then IstShop = True: Exit Function
    
    ' Online-Händler
    If InStr(n, "AMAZON") > 0 Then IstShop = True: Exit Function
    If InStr(n, "EBAY") > 0 Then IstShop = True: Exit Function
    If InStr(n, "ZALANDO") > 0 Then IstShop = True: Exit Function
    If InStr(n, "OTTO") > 0 Then IstShop = True: Exit Function
    If InStr(n, "MEDIAMARKT") > 0 Then IstShop = True: Exit Function
    If InStr(n, "SATURN") > 0 Then IstShop = True: Exit Function
    
    ' Drogerien
    If InStr(n, "ROSSMANN") > 0 Then IstShop = True: Exit Function
    If InStr(n, "MUELLER") > 0 Or InStr(n, "MÜLLER") > 0 Then IstShop = True: Exit Function
    
    ' Möbel/Garten
    If InStr(n, "IKEA") > 0 Then IstShop = True: Exit Function
    If InStr(n, "DEHNER") > 0 Then IstShop = True: Exit Function
    
    ' Tankstellen
    If InStr(n, "ARAL") > 0 Then IstShop = True: Exit Function
    If InStr(n, "SHELL") > 0 Then IstShop = True: Exit Function
    If InStr(n, "TANKSTELLE") > 0 Then IstShop = True: Exit Function
    
    ' Payment
    If InStr(n, "PAYPAL") > 0 Then IstShop = True: Exit Function
    If InStr(n, "KLARNA") > 0 Then IstShop = True: Exit Function
End Function

' ===============================================================
' IstVersorger - Prüft ob Kontoname ein Versorger ist
' ===============================================================
Private Function IstVersorger(ByVal kontoname As String) As Boolean
    Dim n As String
    n = UCase(Trim(kontoname))
    
    IstVersorger = False
    If Len(n) = 0 Then Exit Function
    
    ' Energie
    If InStr(n, "STADTWERK") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "ENERGIE") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "STROM") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "VATTENFALL") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "E.ON") > 0 Or InStr(n, "EON") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "RWE") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "ENVIA") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "GASAG") > 0 Then IstVersorger = True: Exit Function
    
    ' Wasser
    If InStr(n, "WASSER") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "ABWASSER") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "BWB") > 0 Then IstVersorger = True: Exit Function
    
    ' Versicherung
    If InStr(n, "VERSICHERUNG") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "ALLIANZ") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "DEVK") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "HUK") > 0 Then IstVersorger = True: Exit Function
    
    ' Telekommunikation
    If InStr(n, "TELEKOM") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "VODAFONE") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "1&1") > 0 Then IstVersorger = True: Exit Function
    
    ' Entsorgung
    If InStr(n, "BSR") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "ENTSORGUNG") > 0 Then IstVersorger = True: Exit Function
    
    ' Rundfunk
    If InStr(n, "RUNDFUNK") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "BEITRAGSSERVICE") > 0 Then IstVersorger = True: Exit Function
End Function

' ===============================================================
' IstBank - Prüft ob Kontoname eine Bank ist
' ===============================================================
Private Function IstBank(ByVal kontoname As String) As Boolean
    Dim n As String
    n = UCase(Trim(kontoname))
    
    IstBank = False
    If Len(n) = 0 Then Exit Function
    
    If InStr(n, "SPARKASSE") > 0 Then IstBank = True: Exit Function
    If InStr(n, "VOLKSBANK") > 0 Then IstBank = True: Exit Function
    If InStr(n, "RAIFFEISENBANK") > 0 Then IstBank = True: Exit Function
    If InStr(n, "COMMERZBANK") > 0 Then IstBank = True: Exit Function
    If InStr(n, "DEUTSCHE BANK") > 0 Then IstBank = True: Exit Function
    If InStr(n, "POSTBANK") > 0 Then IstBank = True: Exit Function
    If InStr(n, "ING") > 0 Then IstBank = True: Exit Function
    If InStr(n, "DKB") > 0 Then IstBank = True: Exit Function
    If InStr(n, "BANK") > 0 Then IstBank = True: Exit Function
End Function

' ===============================================================
' CreateGUID - Erzeugt eine GUID
' ===============================================================
Private Function CreateGUID() As String
    Dim guid As String
    Dim i As Integer
    
    Randomize Timer
    guid = ""
    
    For i = 1 To 8: guid = guid & Hex(Int(Rnd * 16)): Next i
    guid = guid & "-"
    For i = 1 To 4: guid = guid & Hex(Int(Rnd * 16)): Next i
    guid = guid & "-"
    For i = 1 To 4: guid = guid & Hex(Int(Rnd * 16)): Next i
    guid = guid & "-"
    For i = 1 To 4: guid = guid & Hex(Int(Rnd * 16)): Next i
    guid = guid & "-"
    For i = 1 To 12: guid = guid & Hex(Int(Rnd * 16)): Next i
    
    CreateGUID = LCase(guid)
End Function

' ===============================================================
' ÖFFENTLICH: Verarbeitet manuelle Role-Änderung
' SICHER: Ändert NUR die betroffene Zelle, keine Formatierung
' ===============================================================
Public Sub VerarbeiteManuelleRoleAenderung(ByVal Target As Range)
    Dim wsDaten As Worksheet
    Dim zeile As Long
    Dim neueRole As String
    
    On Error GoTo ErrorHandler
    
    If Target.Column <> EK_COL_ROLE Then Exit Sub
    If Target.Row < EK_START_ROW Then Exit Sub
    
    Set wsDaten = Target.Worksheet
    zeile = Target.Row
    neueRole = Trim(CStr(Target.value))
    
    ' Debug-Info aktualisieren (nur Wert, keine Formatierung!)
    wsDaten.Cells(zeile, EK_COL_DEBUG).value = "Manuell: " & neueRole & " (" & Format(Now, "dd.mm.yyyy hh:mm") & ")"
    
    ' Parzelle leeren wenn nicht erlaubt
    If Not DarfParzelleHaben(neueRole) Then
        wsDaten.Cells(zeile, EK_COL_PARZELLE).value = ""
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "FEHLER in VerarbeiteManuelleRoleAenderung: " & Err.Description
End Sub

' ===============================================================
' ÖFFENTLICH: Formatiert eine einzelne Zeile
' HINWEIS: Tut NICHTS mehr - Formatierung bleibt unverändert!
' ===============================================================
Public Sub FormatiereEntityKeyZeile(ByVal zeile As Long, Optional ByVal ws As Worksheet = Nothing)
    ' BEWUSST LEER - Formatierung wird NICHT mehr geändert!
    ' Diese Funktion existiert nur noch für Kompatibilität mit Tabelle8.cls
End Sub

