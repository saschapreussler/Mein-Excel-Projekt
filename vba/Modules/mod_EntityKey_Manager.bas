Attribute VB_Name = "mod_EntityKey_Manager"
Option Explicit

' ***************************************************************
' MODUL: mod_EntityKey_Manager
' ZWECK: Verwaltung und Zuordnung von EntityKeys fuer Bankverkehr
' VERSION: 3.2 - 06.02.2026
' FIX: IBAN-Import alle Spalte-D-Eintraege, mehrere Kontonamen
'      pro IBAN mit vbLf, MsgBoxen entfernt, Formatierung nach Import
' ***************************************************************

' ===============================================================
' KONSTANTEN (lokal)
' ===============================================================
Private Const EK_ROLE_DROPDOWN_COL As Long = 30

Private Const PREFIX_SHARE As String = "SHARE-"
Private Const PREFIX_VERSORGER As String = "VERS-"
Private Const PREFIX_BANK As String = "BANK-"
Private Const PREFIX_SHOP As String = "SHOP-"
Private Const PREFIX_EHEMALIG As String = "EX-"
Private Const PREFIX_SONSTIGE As String = "SONST-"

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
' HILFSFUNKTION: Prueft ob Role eine Parzelle haben darf
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
    
    If IsNull(iban) Or isEmpty(iban) Then
        NormalisiereIBAN = ""
        Exit Function
    End If
    
    result = UCase(Trim(CStr(iban)))
    result = Replace(result, " ", "")
    result = Replace(result, "-", "")
    
    NormalisiereIBAN = result
End Function

' ===============================================================
' OEFFENTLICHE PROZEDUR: Importiert IBANs aus Bankkonto
' NEU: Alle IBAN-Eintraege aus Spalte D wenn Spalte A Datum hat
'      Auch "0" in Spalte D wird beruecksichtigt
'      Mehrere Kontonamen pro IBAN werden mit vbLf gesammelt
'      MsgBox entfaellt, Formatierung wird nach Import aufgerufen
' ===============================================================
Public Sub ImportiereIBANsAusBankkonto()
    
    Dim wsBK As Worksheet
    Dim wsD As Worksheet
    Dim dictIBANs As Object
    Dim dictKontonamen As Object
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
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    Set dictIBANs = CreateObject("Scripting.Dictionary")
    Set dictKontonamen = CreateObject("Scripting.Dictionary")
    Set dictExisting = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    wsD.Unprotect PASSWORD:=PASSWORD
    wsBK.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    anzahlNeu = 0
    
    ' Bestehende IBANs in EntityKey-Tabelle sammeln
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
    
    ' IBANs aus Bankkonto sammeln - ALLE wo Spalte A ein Datum hat
    lastRowBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    For r = BK_START_ROW To lastRowBK
        currentDatum = wsBK.Cells(r, BK_COL_DATUM).value
        
        ' Pruefen ob Spalte A ein Datum enthaelt
        If Not isEmpty(currentDatum) And currentDatum <> "" Then
            ' IBAN aus Spalte D - auch "0" wird beruecksichtigt
            currentIBAN = NormalisiereIBAN(wsBK.Cells(r, BK_COL_IBAN).value)
            currentKontoName = EntferneMehrfacheLeerzeichen(Trim(CStr(wsBK.Cells(r, BK_COL_NAME).value)))
            
            ' Alle gueltigen IBAN-Eintraege uebernehmen (keine Laengenfilterung mehr)
            If currentIBAN <> "" Then
                ' IBAN merken
                If Not dictIBANs.Exists(currentIBAN) Then
                    dictIBANs.Add currentIBAN, currentKontoName
                    ' Kontonamen-Dictionary fuer diese IBAN starten
                    Dim dictNames As Object
                    Set dictNames = CreateObject("Scripting.Dictionary")
                    If currentKontoName <> "" Then
                        dictNames.Add currentKontoName, True
                    End If
                    Set dictKontonamen(currentIBAN) = dictNames
                Else
                    ' Weiteren Kontonamen sammeln (nicht-redundant)
                    If currentKontoName <> "" Then
                        If Not dictKontonamen(currentIBAN).Exists(currentKontoName) Then
                            dictKontonamen(currentIBAN).Add currentKontoName, True
                        End If
                    End If
                End If
            End If
        End If
    Next r
    
    ' Bestehende Kontonamen in Spalte T aktualisieren (mehrere Namen mit vbLf)
    If lastRowD >= EK_START_ROW Then
        For r = EK_START_ROW To lastRowD
            currentIBAN = NormalisiereIBAN(wsD.Cells(r, EK_COL_IBAN).value)
            If currentIBAN <> "" And dictKontonamen.Exists(currentIBAN) Then
                Dim allNames As String
                allNames = SammelKontonamen(dictKontonamen(currentIBAN))
                If allNames <> "" Then
                    wsD.Cells(r, EK_COL_KONTONAME).value = allNames
                End If
            End If
        Next r
    End If
    
    ' Neue IBANs einfuegen
    If lastRowD < EK_START_ROW Then
        nextRowD = EK_START_ROW
    Else
        nextRowD = lastRowD + 1
    End If
    
    For Each ibanKey In dictIBANs.Keys
        currentIBAN = CStr(ibanKey)
        
        If Not dictExisting.Exists(currentIBAN) Then
            ' Kontonamen zusammenbauen
            Dim kontoNamenGesamt As String
            kontoNamenGesamt = SammelKontonamen(dictKontonamen(currentIBAN))
            
            ' Spalten S und T befuellen
            wsD.Cells(nextRowD, EK_COL_IBAN).value = currentIBAN
            wsD.Cells(nextRowD, EK_COL_KONTONAME).value = kontoNamenGesamt
            anzahlNeu = anzahlNeu + 1
            nextRowD = nextRowD + 1
        End If
    Next ibanKey
    
    ' Formatierung der Daten-Tabelle nach Import aufrufen
    Call mod_Formatierung.FormatEntityKeyTableComplete(wsD)
    
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' Keine MsgBox mehr - Import laeuft still
    Debug.Print "IBAN-Import: " & anzahlNeu & " neue IBANs importiert."
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Debug.Print "Fehler beim IBAN-Import: " & Err.Description
End Sub

' ===============================================================
' HILFSFUNKTION: Sammelt alle Kontonamen aus Dictionary zu String
' Verbindet nicht-redundante Namen mit vbLf
' ===============================================================
Private Function SammelKontonamen(ByRef dictNames As Object) As String
    Dim key As Variant
    Dim result As String
    Dim cleanName As String
    
    result = ""
    
    For Each key In dictNames.Keys
        cleanName = EntferneMehrfacheLeerzeichen(Trim(CStr(key)))
        If cleanName <> "" Then
            If result <> "" Then
                result = result & vbLf & cleanName
            Else
                result = cleanName
            End If
        End If
    Next key
    
    SammelKontonamen = result
End Function

' ===============================================================
' HAUPTPROZEDUR: Aktualisiert alle EntityKeys
' NEU: MsgBox entfaellt
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
        kontoname = EntferneMehrfacheLeerzeichen(Trim(CStr(wsD.Cells(r, EK_COL_KONTONAME).value)))
        
        ' Doppelte Leerzeichen in Spalte T bereinigen
        If CStr(wsD.Cells(r, EK_COL_KONTONAME).value) <> kontoname Then
            wsD.Cells(r, EK_COL_KONTONAME).value = kontoname
        End If
        
        currentEntityKey = Trim(wsD.Cells(r, EK_COL_ENTITYKEY).value)
        currentZuordnung = Trim(wsD.Cells(r, EK_COL_ZUORDNUNG).value)
        currentParzelle = Trim(wsD.Cells(r, EK_COL_PARZELLE).value)
        currentRole = Trim(wsD.Cells(r, EK_COL_ROLE).value)
        
        If iban = "" And kontoname = "" Then GoTo nextRow
        
        If HatBereitsGueltigeDaten(currentEntityKey, currentZuordnung, currentRole) Then
            zeilenUnveraendert = zeilenUnveraendert + 1
            GoTo nextRow
        End If
        
        zeilenNeu = zeilenNeu + 1
        
        Set mitgliederGefunden = SucheMitgliederZuKontoname(kontoname, wsM, wsH)
        
        Call GeneriereEntityKeyUndZuordnung(mitgliederGefunden, kontoname, wsM, _
                                             newEntityKey, zuordnung, parzellen, entityRole, debugInfo, ampelStatus)
        
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
        
        If Trim(wsD.Cells(r, EK_COL_DEBUG).value) = "" Then
            wsD.Cells(r, EK_COL_DEBUG).value = debugInfo
        End If
        
        If ampelStatus = 3 Then zeilenProbleme = zeilenProbleme + 1
        
nextRow:
    Next r
    
    ' Formatierung nach EntityKey-Aktualisierung
    Call mod_Formatierung.FormatEntityKeyTableComplete(wsD)
    
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' Keine MsgBox mehr - laeuft still
    Debug.Print "EntityKey-Update: Neu=" & zeilenNeu & " Unv=" & zeilenUnveraendert & " Probleme=" & zeilenProbleme
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Debug.Print "Fehler bei EntityKey-Aktualisierung: " & Err.Description
End Sub




'--- Ende Teil 1 von 2 ---
'--- Anfang Teil 2 von 2 ---




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
' HILFSFUNKTION: Prueft Namens-Match
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
            PruefeNamensMatch = 2
        Else
            PruefeNamensMatch = 1
        End If
    Else
        PruefeNamensMatch = 1
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
    result = Replace(result, ChrW(228), "ae")
    result = Replace(result, ChrW(246), "oe")
    result = Replace(result, ChrW(252), "ue")
    result = Replace(result, ChrW(223), "ss")
    result = Replace(result, "ae", "a")
    result = Replace(result, "oe", "o")
    result = Replace(result, "ue", "u")
    
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    NormalisiereStringFuerVergleich = Trim(result)
End Function

' ===============================================================
' HILFSFUNKTION: Prueft ob MemberID bereits gefunden
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
    
    For i = 1 To mitglieder.count
        mitgliedInfo = mitglieder(i)
        If mitgliedInfo(8) = 2 Then
            mitgliederExakt.Add mitgliedInfo
        ElseIf mitgliedInfo(8) = 1 Then
            mitgliederNurNachname.Add mitgliedInfo
        End If
    Next i
    
    ' Fall 1: Keine exakten Treffer
    If mitgliederExakt.count = 0 Then
        
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
        
        If mitgliederNurNachname.count > 0 Then
            outDebugInfo = "NUR NACHNAME - Bitte pruefen!"
            outAmpelStatus = 2
            Exit Sub
        End If
        
        outDebugInfo = "KEIN TREFFER - Manuelle Zuordnung"
        outAmpelStatus = 3
        Exit Sub
    End If
    
    ' Fall 2: Exakte Treffer vorhanden
    For i = 1 To mitgliederExakt.count
        mitgliedInfo = mitgliederExakt(i)
        If mitgliedInfo(6) = False Then
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
       InStr(funktionUpper, "SCHRIFTF") > 0 Then
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
' IstShop
' ===============================================================
Private Function IstShop(ByVal kontoname As String) As Boolean
    Dim n As String
    n = UCase(Trim(kontoname))
    IstShop = False
    If Len(n) = 0 Then Exit Function
    
    If InStr(n, "LIDL") > 0 Then IstShop = True: Exit Function
    If InStr(n, "ALDI") > 0 Then IstShop = True: Exit Function
    If InStr(n, "REWE") > 0 Then IstShop = True: Exit Function
    If InStr(n, "EDEKA") > 0 Then IstShop = True: Exit Function
    If InStr(n, "PENNY") > 0 Then IstShop = True: Exit Function
    If InStr(n, "NETTO") > 0 Then IstShop = True: Exit Function
    If InStr(n, "KAUFLAND") > 0 Then IstShop = True: Exit Function
    If InStr(n, "BAUHAUS") > 0 Then IstShop = True: Exit Function
    If InStr(n, "HORNBACH") > 0 Then IstShop = True: Exit Function
    If InStr(n, "OBI") > 0 Then IstShop = True: Exit Function
    If InStr(n, "HAGEBAU") > 0 Then IstShop = True: Exit Function
    If InStr(n, "TOOM") > 0 Then IstShop = True: Exit Function
    If InStr(n, "HELLWEG") > 0 Then IstShop = True: Exit Function
    If InStr(n, "AMAZON") > 0 Then IstShop = True: Exit Function
    If InStr(n, "EBAY") > 0 Then IstShop = True: Exit Function
    If InStr(n, "ZALANDO") > 0 Then IstShop = True: Exit Function
    If InStr(n, "OTTO") > 0 Then IstShop = True: Exit Function
    If InStr(n, "MEDIAMARKT") > 0 Then IstShop = True: Exit Function
    If InStr(n, "SATURN") > 0 Then IstShop = True: Exit Function
    If InStr(n, "ROSSMANN") > 0 Then IstShop = True: Exit Function
    If InStr(n, "IKEA") > 0 Then IstShop = True: Exit Function
    If InStr(n, "DEHNER") > 0 Then IstShop = True: Exit Function
    If InStr(n, "ARAL") > 0 Then IstShop = True: Exit Function
    If InStr(n, "SHELL") > 0 Then IstShop = True: Exit Function
    If InStr(n, "TANKSTELLE") > 0 Then IstShop = True: Exit Function
    If InStr(n, "PAYPAL") > 0 Then IstShop = True: Exit Function
    If InStr(n, "KLARNA") > 0 Then IstShop = True: Exit Function
End Function

' ===============================================================
' IstVersorger
' ===============================================================
Private Function IstVersorger(ByVal kontoname As String) As Boolean
    Dim n As String
    n = UCase(Trim(kontoname))
    IstVersorger = False
    If Len(n) = 0 Then Exit Function
    
    If InStr(n, "STADTWERK") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "ENERGIE") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "STROM") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "VATTENFALL") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "E.ON") > 0 Or InStr(n, "EON") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "RWE") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "ENVIA") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "GASAG") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "WASSER") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "ABWASSER") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "BWB") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "VERSICHERUNG") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "ALLIANZ") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "DEVK") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "HUK") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "TELEKOM") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "VODAFONE") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "1&1") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "BSR") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "ENTSORGUNG") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "RUNDFUNK") > 0 Then IstVersorger = True: Exit Function
    If InStr(n, "BEITRAGSSERVICE") > 0 Then IstVersorger = True: Exit Function
End Function

' ===============================================================
' IstBank
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
' CreateGUID
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
' OEFFENTLICH: Verarbeitet manuelle Role-Aenderung
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
    
    wsDaten.Cells(zeile, EK_COL_DEBUG).value = "Manuell: " & neueRole & " (" & Format(Now, "dd.mm.yyyy hh:mm") & ")"
    
    If Not DarfParzelleHaben(neueRole) Then
        wsDaten.Cells(zeile, EK_COL_PARZELLE).value = ""
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "FEHLER in VerarbeiteManuelleRoleAenderung: " & Err.Description
End Sub

' ===============================================================
' OEFFENTLICH: Formatiert eine einzelne Zeile (Kompatibilitaet)
' ===============================================================
Public Sub FormatiereEntityKeyZeile(ByVal zeile As Long, Optional ByVal ws As Worksheet = Nothing)
    ' BEWUSST LEER - Formatierung wird durch mod_Formatierung gesteuert
End Sub

