Attribute VB_Name = "mod_EntityKey_Manager"
Option Explicit

' ***************************************************************
' MODUL: mod_EntityKey_Manager (ORCHESTRATOR)
' ZWECK: Verwaltung und Zuordnung von EntityKeys fuer Bankverkehr
' VERSION: 6.0 - 01.03.2026 (Modularisiert)
'
' SUB-MODULE:
'   mod_EntityKey_Normalize  - String-Normalisierung (IBAN, Namen)
'   mod_EntityKey_Kontoname  - Kontonamen-Deduplizierung
'   mod_EntityKey_Classifier - Klassifikation (Shop/Versorger/Bank)
'   mod_EntityKey_Matching   - Mitgliedersuche und Namensabgleich
'   mod_EntityKey_Ampel      - Ampelfarben-System (Gruen/Gelb/Rot)
'   mod_EntityKey_UI         - UI-Interaktion (manuelle Aenderungen)
'
' VERBLEIBENDE FUNKTIONEN:
'   - ImportiereIBANsAusBankkonto: IBAN-Import aus Bankkonto
'   - AktualisiereAlleEntityKeys: Haupt-Update aller EntityKeys
'   - GeneriereEntityKeyUndZuordnung: Zuordnungslogik
'   - AktualisiereEntityKeyBeiAustritt: EX-Prefix bei Austritt
'   - HatBereitsGueltigeDaten: Gueltigkeitspruefung
'   - CreateGUID: ID-Generator
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
Private Const ROLE_VERSORGER As String = "VERSORGER"
Private Const ROLE_BANK As String = "BANK"
Private Const ROLE_SHOP As String = "SHOP"
Private Const ROLE_SONSTIGE As String = "SONSTIGE"

' Ampelfarben
Private Const AMPEL_GRUEN As Long = 12968900
Private Const AMPEL_GELB As Long = 10086143
Private Const AMPEL_ROT As Long = 9871103

' ===============================================================
' OEFFENTLICHE PROZEDUR: Importiert IBANs aus Bankkonto
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
            currentIBAN = mod_EntityKey_Normalize.NormalisiereIBAN(wsD.Cells(r, EK_COL_IBAN).value)
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
        
        If Not isEmpty(currentDatum) And currentDatum <> "" Then
            currentIBAN = mod_EntityKey_Normalize.NormalisiereIBAN(wsBK.Cells(r, BK_COL_IBAN).value)
            currentKontoName = mod_EntityKey_Normalize.EntferneMehrfacheLeerzeichen(Trim(CStr(wsBK.Cells(r, BK_COL_NAME).value)))
            
            If currentIBAN <> "" Then
                If Not dictIBANs.Exists(currentIBAN) Then
                    dictIBANs.Add currentIBAN, currentKontoName
                    Dim dictNames As Object
                    Set dictNames = CreateObject("Scripting.Dictionary")
                    If currentKontoName <> "" Then
                        dictNames.Add UCase(Trim(currentKontoName)), currentKontoName
                    End If
                    Set dictKontonamen(currentIBAN) = dictNames
                Else
                    If currentKontoName <> "" Then
                        Dim nameKey As String
                        nameKey = UCase(Trim(currentKontoName))
                        If Not dictKontonamen(currentIBAN).Exists(nameKey) Then
                            If Not mod_EntityKey_Kontoname.IstKontonameRedundant(dictKontonamen(currentIBAN), currentKontoName) Then
                                dictKontonamen(currentIBAN).Add nameKey, currentKontoName
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next r
    
    ' Bestehende Kontonamen in Spalte T aktualisieren
    If lastRowD >= EK_START_ROW Then
        For r = EK_START_ROW To lastRowD
            currentIBAN = mod_EntityKey_Normalize.NormalisiereIBAN(wsD.Cells(r, EK_COL_IBAN).value)
            If currentIBAN <> "" And dictKontonamen.Exists(currentIBAN) Then
                Dim bereinigteNamen As Object
                Set bereinigteNamen = mod_EntityKey_Kontoname.BereinigeKontonamen(dictKontonamen(currentIBAN))
                Dim allNames As String
                allNames = mod_EntityKey_Kontoname.SammelKontonamen(bereinigteNamen)
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
    
    For Each ibanKey In dictIBANs.keys
        currentIBAN = CStr(ibanKey)
        
        If Not dictExisting.Exists(currentIBAN) Then
            Dim bereinigt As Object
            Set bereinigt = mod_EntityKey_Kontoname.BereinigeKontonamen(dictKontonamen(currentIBAN))
            Dim kontoNamenGesamt As String
            kontoNamenGesamt = mod_EntityKey_Kontoname.SammelKontonamen(bereinigt)
            
            wsD.Cells(nextRowD, EK_COL_IBAN).value = currentIBAN
            wsD.Cells(nextRowD, EK_COL_KONTONAME).value = kontoNamenGesamt
            anzahlNeu = anzahlNeu + 1
            nextRowD = nextRowD + 1
        End If
    Next ibanKey
    
    Call mod_Formatierung.FormatEntityKeyTableComplete(wsD)
    
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
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
' HAUPTPROZEDUR: Aktualisiert alle EntityKeys
' ===============================================================
Public Sub AktualisiereAlleEntityKeys()
    
    Dim wsD As Worksheet
    Dim wsM As Worksheet
    Dim wsH As Worksheet
    Dim wsBK As Worksheet
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
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
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
        kontoname = mod_EntityKey_Normalize.EntferneMehrfacheLeerzeichen(Trim(CStr(wsD.Cells(r, EK_COL_KONTONAME).value)))
        
        If CStr(wsD.Cells(r, EK_COL_KONTONAME).value) <> kontoname Then
            wsD.Cells(r, EK_COL_KONTONAME).value = kontoname
        End If
        
        currentEntityKey = Trim(wsD.Cells(r, EK_COL_ENTITYKEY).value)
        currentZuordnung = Trim(wsD.Cells(r, EK_COL_ZUORDNUNG).value)
        currentParzelle = Trim(wsD.Cells(r, EK_COL_PARZELLE).value)
        currentRole = Trim(wsD.Cells(r, EK_COL_ROLE).value)
        
        If iban = "" And kontoname = "" Then GoTo nextRow
        
        ' Pruefe Geldautomat-Abhebung VOR Bankabschluss
        If currentEntityKey = "" Then
            If mod_EntityKey_Classifier.IstGeldautomatAbhebung(iban, kontoname) Then
                wsD.Cells(r, EK_COL_ENTITYKEY).value = PREFIX_BANK & CreateGUID()
                wsD.Cells(r, EK_COL_ZUORDNUNG).value = "Bargeldabhebung Geldautomat (Vereinskasse)"
                wsD.Cells(r, EK_COL_ROLE).value = ROLE_BANK
                wsD.Cells(r, EK_COL_DEBUG).value = "Geldautomat erkannt (GA + BLZ)"
                Call mod_EntityKey_Ampel.SetzeAmpelFarbe(wsD, r, 1)
                GoTo nextRow
            End If
        End If
        
        ' Pruefe IBAN "0" oder "3529000972" + ABSCHLUSS
        If currentEntityKey = "" Then
            If mod_EntityKey_Classifier.IstBankAbschluss(iban, wsBK) Then
                wsD.Cells(r, EK_COL_ENTITYKEY).value = PREFIX_BANK & CreateGUID()
                wsD.Cells(r, EK_COL_ZUORDNUNG).value = "Bankabschluss / Kontogeb" & ChrW(252) & "hren"
                wsD.Cells(r, EK_COL_ROLE).value = ROLE_BANK
                wsD.Cells(r, EK_COL_DEBUG).value = "BANK erkannt (IBAN=" & iban & " + ABSCHLUSS)"
                GoTo nextRow
            End If
        End If
        
        If HatBereitsGueltigeDaten(currentEntityKey, currentZuordnung, currentRole) Then
            zeilenUnveraendert = zeilenUnveraendert + 1
            GoTo nextRow
        End If
        
        zeilenNeu = zeilenNeu + 1
        
        Set mitgliederGefunden = mod_EntityKey_Matching.SucheMitgliederZuKontoname(kontoname, wsM, wsH)
        
        Call GeneriereEntityKeyUndZuordnung(mitgliederGefunden, kontoname, wsM, _
                                             newEntityKey, zuordnung, parzellen, entityRole, debugInfo, ampelStatus)
        
        If currentEntityKey = "" And newEntityKey <> "" Then
            wsD.Cells(r, EK_COL_ENTITYKEY).value = newEntityKey
        End If
        
        If currentZuordnung = "" And zuordnung <> "" Then
            wsD.Cells(r, EK_COL_ZUORDNUNG).value = zuordnung
        End If
        
        If currentParzelle = "" And parzellen <> "" And mod_EntityKey_Classifier.DarfParzelleHaben(entityRole) Then
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
    
    ' Formatierung ZUERST (inkl. Sortierung)
    Call mod_Formatierung.FormatEntityKeyTableComplete(wsD)
    
    ' Ampelfarben DANACH (nach Sortierung!)
    Call mod_EntityKey_Ampel.SetzeAlleAmpelfarbenNachSortierung(wsD)
    
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Debug.Print "EntityKey-Update: Neu=" & zeilenNeu & " Unv=" & zeilenUnveraendert & " Probleme=" & zeilenProbleme
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Debug.Print "Fehler bei EntityKey-Aktualisierung: " & Err.Description
End Sub

' ===============================================================
' Prueft ob Zeile bereits gueltige Daten hat
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
' Generiert EntityKey und Zuordnung
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
    
    ' Fall 1: Keine exakten Treffer -> Automatik pruefen
    If mitgliederExakt.count = 0 Then
        
        If mod_EntityKey_Classifier.IstShop(kontoname) Then
            outEntityKey = PREFIX_SHOP & CreateGUID()
            outEntityRole = ROLE_SHOP
            outZuordnung = mod_EntityKey_Normalize.ExtrahiereAnzeigeName(kontoname)
            outDebugInfo = "Automatisch als SHOP erkannt"
            outAmpelStatus = 1
            Exit Sub
        End If
        
        ' VERSORGER mit Zweck-Erkennung
        Dim versorgerZweck As String
        versorgerZweck = mod_EntityKey_Classifier.ErmittleVersorgerZweck(kontoname)
        If versorgerZweck <> "" Then
            outEntityKey = PREFIX_VERSORGER & CreateGUID()
            outEntityRole = ROLE_VERSORGER
            outZuordnung = mod_EntityKey_Normalize.ExtrahiereAnzeigeName(kontoname)
            outDebugInfo = "Automatisch als VERSORGER erkannt (" & versorgerZweck & ")"
            outAmpelStatus = 1
            Exit Sub
        End If
        
        If mod_EntityKey_Classifier.IstBank(kontoname) Then
            outEntityKey = PREFIX_BANK & CreateGUID()
            outEntityRole = ROLE_BANK
            outZuordnung = mod_EntityKey_Normalize.ExtrahiereAnzeigeName(kontoname)
            outDebugInfo = "Automatisch als BANK erkannt"
            outAmpelStatus = 1
            Exit Sub
        End If
        
        If mitgliederNurNachname.count > 0 Then
            outDebugInfo = "NUR NACHNAME - Bitte pr" & ChrW(252) & "fen!"
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
    
    ' Fall 2a: Genau 1 aktives Mitglied -> GRUEN
    If uniqueMemberIDs.count = 1 Then
        For i = 1 To mitgliederExakt.count
            mitgliedInfo = mitgliederExakt(i)
            If mitgliedInfo(6) = False Then
                outEntityKey = CStr(mitgliedInfo(0))
                outZuordnung = mitgliedInfo(1) & ", " & mitgliedInfo(2)
                outParzellen = mod_EntityKey_Matching.HoleAlleParzellen(CStr(mitgliedInfo(0)), wsM)
                outEntityRole = mod_EntityKey_Classifier.ErmittleEntityRoleVonFunktion(CStr(mitgliedInfo(4)))
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
        For Each key In uniqueMemberIDs.keys
            If memberIDs <> "" Then memberIDs = memberIDs & "_"
            memberIDs = memberIDs & key
        Next key
        
        outEntityKey = PREFIX_SHARE & memberIDs
        outEntityRole = ROLE_MITGLIED_MIT_PACHT
        outDebugInfo = "Gemeinschaftskonto - " & uniqueMemberIDs.count & " Personen (automatisch erkannt)"
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
                    
                    Dim personParzellen As String
                    personParzellen = mod_EntityKey_Matching.HoleAlleParzellen(CStr(mitgliedInfo(0)), wsM)
                    
                    If personParzellen <> "" Then
                        If outParzellen <> "" Then
                            Dim arrP() As String
                            Dim p As Long
                            arrP = Split(personParzellen, ", ")
                            For p = LBound(arrP) To UBound(arrP)
                                If InStr(outParzellen, Trim(arrP(p))) = 0 Then
                                    outParzellen = outParzellen & ", " & Trim(arrP(p))
                                End If
                            Next p
                        Else
                            outParzellen = personParzellen
                        End If
                    End If
                End If
            End If
        Next i
    End If
End Sub

' ===============================================================
' CreateGUID - Generiert ID im Format "yyyymmddhhmmss-NNNNN"
' Public weil auch von mod_EntityKey_UI aufgerufen
' ===============================================================
Public Function CreateGUID() As String
    Randomize
    CreateGUID = Format(Now, "yyyymmddhhmmss") & "-" & Int((99999 - 10000 + 1) * Rnd + 10000)
End Function

' ===============================================================
' Aktualisiert EntityKey-Tabelle bei Mitglied-Austritt
' Sucht alle Zeilen mit der alten MemberID und setzt EX-Prefix
' Wird aufgerufen aus mod_Mitglieder_UI nach Historie-Eintrag
' ===============================================================
Public Sub AktualisiereEntityKeyBeiAustritt(ByVal alteMemberID As String)
    Dim wsD As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim currentEK As String
    Dim neuerEK As String
    Dim currentRole As String
    Dim anzahlAktualisiert As Long
    
    On Error GoTo ErrorHandler
    
    If alteMemberID = "" Then Exit Sub
    
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    
    On Error Resume Next
    wsD.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    lastRow = wsD.Cells(wsD.Rows.count, EK_COL_IBAN).End(xlUp).Row
    Dim lastRowR As Long
    lastRowR = wsD.Cells(wsD.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    If lastRowR > lastRow Then lastRow = lastRowR
    If lastRow < EK_START_ROW Then GoTo CleanUp
    
    anzahlAktualisiert = 0
    
    For r = EK_START_ROW To lastRow
        currentEK = Trim(wsD.Cells(r, EK_COL_ENTITYKEY).value)
        
        ' Pruefe ob EntityKey die alte MemberID ist (direkt oder als SHARE-Teil)
        If currentEK = alteMemberID Then
            ' Einzelne MemberID -> EX-Prefix + neue GUID
            neuerEK = PREFIX_EHEMALIG & CreateGUID()
            wsD.Cells(r, EK_COL_ENTITYKEY).value = neuerEK
            wsD.Cells(r, EK_COL_ROLE).value = ROLE_EHEMALIGES_MITGLIED
            wsD.Cells(r, EK_COL_DEBUG).value = "Austritt: MemberID ge" & ChrW(228) & "ndert zu EX- (" & Format(Now, "dd.mm.yyyy") & ")"
            Call mod_EntityKey_Ampel.SetzeAmpelFarbe(wsD, r, 1)
            anzahlAktualisiert = anzahlAktualisiert + 1
            
        ElseIf Left(currentEK, Len(PREFIX_SHARE)) = PREFIX_SHARE Then
            ' SHARE-Key: pruefe ob MemberID enthalten
            If InStr(currentEK, alteMemberID) > 0 Then
                ' Gemeinschaftskonto mit diesem Mitglied
                Dim sharePart As String
                sharePart = Mid(currentEK, Len(PREFIX_SHARE) + 1)
                
                Dim idParts() As String
                idParts = Split(sharePart, "_")
                
                Dim newShareParts As String
                newShareParts = ""
                Dim verbleibendeAnzahl As Long
                verbleibendeAnzahl = 0
                
                Dim p As Long
                For p = LBound(idParts) To UBound(idParts)
                    If idParts(p) <> alteMemberID Then
                        If newShareParts <> "" Then newShareParts = newShareParts & "_"
                        newShareParts = newShareParts & idParts(p)
                        verbleibendeAnzahl = verbleibendeAnzahl + 1
                    End If
                Next p
                
                If verbleibendeAnzahl = 1 Then
                    wsD.Cells(r, EK_COL_ENTITYKEY).value = newShareParts
                ElseIf verbleibendeAnzahl > 1 Then
                    wsD.Cells(r, EK_COL_ENTITYKEY).value = PREFIX_SHARE & newShareParts
                End If
                
                Dim altDebug As String
                altDebug = Trim(wsD.Cells(r, EK_COL_DEBUG).value)
                If altDebug <> "" Then
                    wsD.Cells(r, EK_COL_DEBUG).value = altDebug & " | Austritt eines Kontoinhabers (" & Format(Now, "dd.mm.yyyy") & ")"
                Else
                    wsD.Cells(r, EK_COL_DEBUG).value = "Austritt eines Kontoinhabers (" & Format(Now, "dd.mm.yyyy") & ")"
                End If
                
                anzahlAktualisiert = anzahlAktualisiert + 1
            End If
        End If
    Next r
    
CleanUp:
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    If anzahlAktualisiert > 0 Then
        Debug.Print "EntityKey-Austritt: " & anzahlAktualisiert & " Zeilen aktualisiert f" & ChrW(252) & "r MemberID " & alteMemberID
    End If
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Debug.Print "FEHLER in AktualisiereEntityKeyBeiAustritt: " & Err.Description
End Sub
























