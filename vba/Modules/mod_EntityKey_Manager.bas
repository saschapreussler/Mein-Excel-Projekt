Attribute VB_Name = "mod_EntityKey_Manager"
Option Explicit

' ***************************************************************
' MODUL: mod_EntityKey_Manager
' ZWECK: Verwaltung und Zuordnung von EntityKeys fuer Bankverkehr
' VERSION: 2.0 - 02.02.2026
' KORREKTUR: iban Parameter hinzugefuegt, SetzeZellschutzFuerZeile Public
' ***************************************************************

Public Const EK_COL_ENTITYKEY As Long = 18
Public Const EK_COL_IBAN As Long = 19
Public Const EK_COL_KONTONAME As Long = 20
Public Const EK_COL_ZUORDNUNG As Long = 21
Public Const EK_COL_PARZELLE As Long = 22
Public Const EK_COL_ROLE As Long = 23
Public Const EK_COL_DEBUG As Long = 24

Public Const EK_START_ROW As Long = 4
Public Const EK_HEADER_ROW As Long = 3

Private Const EK_ROLE_DROPDOWN_COL As Long = 30

Public Const PREFIX_SHARE As String = "SHARE-"
Public Const PREFIX_VERSORGER As String = "VERS-"
Public Const PREFIX_BANK As String = "BANK-"
Public Const PREFIX_SHOP As String = "SHOP-"
Public Const PREFIX_EHEMALIG As String = "EX-"
Public Const PREFIX_SONSTIGE As String = "SONSTIGE-"

Public Const ROLE_MITGLIED_MIT_PACHT As String = "MITGLIED_MIT_PACHT"
Public Const ROLE_MITGLIED_OHNE_PACHT As String = "MITGLIED_OHNE_PACHT"
Public Const ROLE_EHEMALIGES_MITGLIED As String = "EHEMALIGES_MITGLIED"
Public Const ROLE_VERSORGER As String = "VERSORGER"
Public Const ROLE_BANK As String = "BANK"
Public Const ROLE_SHOP As String = "SHOP"
Public Const ROLE_SONSTIGE As String = "SONSTIGE"

Private Const ZEBRA_COLOR As Long = &HDEE5E3

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
    
    On Error Resume Next
    wsD.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    Set dictIBANs = CreateObject("Scripting.Dictionary")
    
    lastRowBK = wsBK.Cells(wsBK.Rows.Count, BK_COL_DATUM).End(xlUp).Row
    If lastRowBK < BK_START_ROW Then lastRowBK = BK_START_ROW
    
    lastRowD = wsD.Cells(wsD.Rows.Count, EK_COL_IBAN).End(xlUp).Row
    If lastRowD < EK_START_ROW Then lastRowD = EK_START_ROW - 1
    
    For r = EK_START_ROW To lastRowD
        currentIBAN = Trim(wsD.Cells(r, EK_COL_IBAN).value)
        currentKontoName = Trim(wsD.Cells(r, EK_COL_KONTONAME).value)
        If currentIBAN <> "" Then
            dictIBANs(currentIBAN & "|" & currentKontoName) = True
        End If
    Next r
    
    anzahlBereitsVorhanden = dictIBANs.Count
    nextRowD = lastRowD + 1
    anzahlNeu = 0
    anzahlZeilenGeprueft = 0
    
    For r = BK_START_ROW To lastRowBK
        anzahlZeilenGeprueft = anzahlZeilenGeprueft + 1
        
        currentDatum = wsBK.Cells(r, BK_COL_DATUM).value
        If IsEmpty(currentDatum) Or currentDatum = "" Then GoTo NextRowImport
        
        currentIBAN = Trim(wsBK.Cells(r, BK_COL_IBAN).value)
        currentKontoName = Trim(wsBK.Cells(r, BK_COL_NAME).value)
        
        If currentIBAN = "" And currentKontoName = "" Then GoTo NextRowImport
        
        If Not dictIBANs.Exists(currentIBAN & "|" & currentKontoName) Then
            wsD.Cells(nextRowD, EK_COL_IBAN).value = currentIBAN
            wsD.Cells(nextRowD, EK_COL_KONTONAME).value = currentKontoName
            
            dictIBANs(currentIBAN & "|" & currentKontoName) = True
            nextRowD = nextRowD + 1
            anzahlNeu = anzahlNeu + 1
        End If
        
NextRowImport:
    Next r
    
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "IBAN-Import abgeschlossen!" & vbCrLf & vbCrLf & _
           "Zeilen auf Bankkonto geprueft: " & anzahlZeilenGeprueft & vbCrLf & _
           "Bereits vorhandene IBANs: " & anzahlBereitsVorhanden & vbCrLf & _
           "Neue IBANs importiert: " & anzahlNeu, vbInformation, "Import abgeschlossen"
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler beim IBAN-Import: " & Err.Description, vbCritical
End Sub


'--- Ende Teil 1 ---
'--- Anfang Teil 2 ---


Public Sub AktualisiereEntityKeys()
    
    Dim wsD As Worksheet
    Dim wsM As Worksheet
    Dim wsH As Worksheet
    Dim r As Long
    Dim lastRow As Long
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
    Dim zeilenNeu As Long
    Dim zeilenUnveraendert As Long
    Dim zeilenRot As Collection
    Dim zeilenGelb As Collection
    
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
        
        Call GeneriereEntityKeyUndZuordnung(mitgliederGefunden, iban, kontoName, wsM, _
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

Public Sub SetzeZellschutzFuerZeile(ByRef ws As Worksheet, ByVal zeile As Long, ByVal role As String)
    
    ws.Cells(zeile, EK_COL_ZUORDNUNG).Locked = False
    ws.Cells(zeile, EK_COL_ROLE).Locked = False
    ws.Cells(zeile, EK_COL_DEBUG).Locked = False
    
    If DarfParzelleHaben(role) Then
        ws.Cells(zeile, EK_COL_PARZELLE).Locked = False
    Else
        ws.Cells(zeile, EK_COL_PARZELLE).Locked = True
    End If
    
End Sub

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


'--- Ende Teil 2 ---
'--- Anfang Teil 3 ---


Private Sub ZeigeEingriffsHinweis(ByRef ws As Worksheet, ByRef zeilenRot As Collection, _
                                   ByRef zeilenGelb As Collection, _
                                   ByVal zeilenNeu As Long, ByVal zeilenUnveraendert As Long)
    
    Dim msg As String
    Dim antwort As VbMsgBoxResult
    Dim ersteZeile As Long
    
    msg = "EntityKey-Aktualisierung abgeschlossen!" & vbCrLf & vbCrLf & _
          "Neue Zeilen verarbeitet: " & zeilenNeu & vbCrLf & _
          "Bestehende Zeilen unveraendert: " & zeilenUnveraendert & vbCrLf & vbCrLf
    
    If zeilenRot.Count > 0 Then
        msg = msg & "ROT (manueller Eingriff erforderlich): " & zeilenRot.Count & " Zeilen" & vbCrLf
    End If
    
    If zeilenGelb.Count > 0 Then
        msg = msg & "GELB (Pruefung empfohlen): " & zeilenGelb.Count & " Zeilen" & vbCrLf
    End If
    
    msg = msg & vbCrLf & "Moechten Sie zur ersten markierten Zeile springen?"
    
    antwort = MsgBox(msg, vbYesNo + vbQuestion, "Aktualisierung mit Hinweisen")
    
    If antwort = vbYes Then
        If zeilenRot.Count > 0 Then
            ersteZeile = zeilenRot(1)
        ElseIf zeilenGelb.Count > 0 Then
            ersteZeile = zeilenGelb(1)
        End If
        
        If ersteZeile > 0 Then
            Application.GoTo ws.Cells(ersteZeile, EK_COL_ZUORDNUNG), True
        End If
    End If
    
End Sub

Private Sub SetzeAmpelFarbe(ByRef ws As Worksheet, ByVal zeile As Long, ByVal status As Long)
    
    Dim rngAmpel As Range
    Set rngAmpel = ws.Range(ws.Cells(zeile, EK_COL_ZUORDNUNG), ws.Cells(zeile, EK_COL_DEBUG))
    
    Select Case status
        Case 1
            rngAmpel.Interior.color = RGB(198, 239, 206)
        Case 2
            rngAmpel.Interior.color = RGB(255, 235, 156)
        Case 3
            rngAmpel.Interior.color = RGB(255, 199, 206)
        Case Else
            rngAmpel.Interior.ColorIndex = xlNone
    End Select
    
End Sub

Private Sub SetupEntityRoleDropdown(ByRef ws As Worksheet, ByVal lastRow As Long)
    
    Dim dropdownRange As Range
    Dim r As Long
    Dim roleList As String
    
    roleList = ROLE_MITGLIED_MIT_PACHT & "," & _
               ROLE_MITGLIED_OHNE_PACHT & "," & _
               ROLE_EHEMALIGES_MITGLIED & "," & _
               ROLE_VERSORGER & "," & _
               ROLE_BANK & "," & _
               ROLE_SHOP & "," & _
               ROLE_SONSTIGE
    
    For r = EK_START_ROW To lastRow
        On Error Resume Next
        ws.Cells(r, EK_COL_ROLE).Validation.Delete
        
        With ws.Cells(r, EK_COL_ROLE).Validation
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertWarning, _
                 Formula1:=roleList
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = True
            .ErrorTitle = "Ungueltige Eingabe"
            .ErrorMessage = "Bitte waehlen Sie einen Wert aus der Liste."
        End With
        On Error GoTo 0
    Next r
    
End Sub

Private Function SucheMitgliederZuKontoname(ByVal kontoName As String, _
                                             ByRef wsM As Worksheet, _
                                             ByRef wsH As Worksheet) As Collection
    
    Dim result As New Collection
    Dim r As Long
    Dim lastRowM As Long
    Dim lastRowH As Long
    Dim vorname As String
    Dim nachname As String
    Dim parzelle As String
    Dim memberID As String
    Dim kontoNameNorm As String
    Dim vornameNorm As String
    Dim nachnameNorm As String
    Dim matchScore As Long
    Dim mitgliedInfo(0 To 8) As Variant
    
    Set SucheMitgliederZuKontoname = result
    
    If Trim(kontoName) = "" Then Exit Function
    
    kontoNameNorm = NormalisiereStringFuerVergleich(kontoName)
    
    lastRowM = wsM.Cells(wsM.Rows.Count, M_COL_MEMBER_ID).End(xlUp).Row
    
    For r = M_START_ROW To lastRowM
        memberID = Trim(wsM.Cells(r, M_COL_MEMBER_ID).value)
        vorname = Trim(wsM.Cells(r, M_COL_VORNAME).value)
        nachname = Trim(wsM.Cells(r, M_COL_NACHNAME).value)
        parzelle = Trim(wsM.Cells(r, M_COL_PARZELLE).value)
        
        If nachname <> "" Then
            vornameNorm = NormalisiereStringFuerVergleich(vorname)
            nachnameNorm = NormalisiereStringFuerVergleich(nachname)
            
            matchScore = PruefeNamensMatch(kontoNameNorm, vornameNorm, nachnameNorm)
            
            If matchScore > 0 Then
                If Not IstMitgliedBereitsGefunden(result, memberID, False) Then
                    mitgliedInfo(0) = memberID
                    mitgliedInfo(1) = vorname
                    mitgliedInfo(2) = nachname
                    mitgliedInfo(3) = parzelle
                    mitgliedInfo(4) = r
                    mitgliedInfo(5) = "Mitglieder"
                    mitgliedInfo(6) = False
                    mitgliedInfo(7) = IIf(parzelle <> "", True, False)
                    mitgliedInfo(8) = matchScore
                    result.Add mitgliedInfo
                End If
            End If
        End If
    Next r
    
    On Error Resume Next
    lastRowH = wsH.Cells(wsH.Rows.Count, 1).End(xlUp).Row
    On Error GoTo 0
    
    If lastRowH > 0 Then
        For r = 2 To lastRowH
            memberID = Trim(wsH.Cells(r, 1).value)
            vorname = Trim(wsH.Cells(r, 2).value)
            nachname = Trim(wsH.Cells(r, 3).value)
            parzelle = Trim(wsH.Cells(r, 4).value)
            
            If nachname <> "" Then
                vornameNorm = NormalisiereStringFuerVergleich(vorname)
                nachnameNorm = NormalisiereStringFuerVergleich(nachname)
                
                matchScore = PruefeNamensMatch(kontoNameNorm, vornameNorm, nachnameNorm)
                
                If matchScore > 0 Then
                    If Not IstMitgliedBereitsGefunden(result, memberID, True) Then
                        mitgliedInfo(0) = memberID
                        mitgliedInfo(1) = vorname
                        mitgliedInfo(2) = nachname
                        mitgliedInfo(3) = parzelle
                        mitgliedInfo(4) = r
                        mitgliedInfo(5) = "Historie"
                        mitgliedInfo(6) = True
                        mitgliedInfo(7) = IIf(parzelle <> "", True, False)
                        mitgliedInfo(8) = matchScore
                        result.Add mitgliedInfo
                    End If
                End If
            End If
        Next r
    End If
    
    Set SucheMitgliederZuKontoname = result
    
End Function


'--- Ende Teil 3 ---
'--- Anfang Teil 4 ---


Private Function PruefeNamensMatch(ByVal kontoNameNorm As String, _
                                    ByVal vornameNorm As String, _
                                    ByVal nachnameNorm As String) As Long
    
    PruefeNamensMatch = 0
    
    If InStr(kontoNameNorm, nachnameNorm) > 0 And InStr(kontoNameNorm, vornameNorm) > 0 Then
        PruefeNamensMatch = 2
        Exit Function
    End If
    
    If InStr(kontoNameNorm, nachnameNorm) > 0 Then
        PruefeNamensMatch = 1
        Exit Function
    End If
    
End Function

Private Function NormalisiereStringFuerVergleich(ByVal inputStr As String) As String
    Dim result As String
    
    result = LCase(Trim(inputStr))
    
    result = Replace(result, ChrW(228), "a")
    result = Replace(result, ChrW(246), "o")
    result = Replace(result, ChrW(252), "u")
    result = Replace(result, ChrW(196), "a")
    result = Replace(result, ChrW(214), "o")
    result = Replace(result, ChrW(220), "u")
    result = Replace(result, ChrW(223), "ss")
    
    result = Replace(result, "ae", "a")
    result = Replace(result, "oe", "o")
    result = Replace(result, "ue", "u")
    
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    result = Trim(result)
    
    NormalisiereStringFuerVergleich = result
End Function

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

Private Sub GeneriereEntityKeyUndZuordnung(ByRef mitglieder As Collection, _
                                            ByVal iban As String, _
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
    
    If mitgliederExakt.Count = 0 Then
        If IstBankTransaktion(iban, kontoName) Or IstBank(kontoName) Then
            outEntityKey = PREFIX_BANK & CreateGUID()
            outEntityRole = ROLE_BANK
            outZuordnung = ExtrahiereAnzeigeName(kontoName)
            If outZuordnung = "" Then outZuordnung = "Bank-Transaktion"
            outParzellen = ""
            outDebugInfo = "Automatisch als BANK erkannt"
            outAmpelStatus = 1
            Exit Sub
        ElseIf IstVersorger(kontoName) Then
            outEntityKey = PREFIX_VERSORGER & CreateGUID()
            outEntityRole = ROLE_VERSORGER
            outZuordnung = ExtrahiereAnzeigeName(kontoName)
            outParzellen = ""
            outDebugInfo = "Automatisch als VERSORGER erkannt"
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
        
        If mitgliederNurNachname.Count = 1 Then
            mitgliedInfo = mitgliederNurNachname(1)
            
            If mitgliedInfo(6) = True Then
                outEntityKey = PREFIX_EHEMALIG & mitgliedInfo(0)
                outEntityRole = ROLE_EHEMALIGES_MITGLIED
                outZuordnung = mitgliedInfo(2) & ", " & mitgliedInfo(1) & " (Ehemalig)"
                outParzellen = mitgliedInfo(3)
                outDebugInfo = "Nur Nachname gefunden (Historie)"
                outAmpelStatus = 2
            Else
                outEntityKey = mitgliedInfo(0)
                If mitgliedInfo(7) Then
                    outEntityRole = ROLE_MITGLIED_MIT_PACHT
                Else
                    outEntityRole = ROLE_MITGLIED_OHNE_PACHT
                End If
                outZuordnung = mitgliedInfo(2) & ", " & mitgliedInfo(1)
                outParzellen = mitgliedInfo(3)
                outDebugInfo = "Nur Nachname gefunden (aktiv)"
                outAmpelStatus = 2
            End If
            Exit Sub
        ElseIf mitgliederNurNachname.Count > 1 Then
            outEntityKey = PREFIX_SONSTIGE & CreateGUID()
            outEntityRole = ROLE_SONSTIGE
            outZuordnung = ""
            outParzellen = ""
            outDebugInfo = "Mehrere Nachnamen-Treffer (" & mitgliederNurNachname.Count & ")"
            outAmpelStatus = 3
            Exit Sub
        End If
        
        outEntityKey = PREFIX_SONSTIGE & CreateGUID()
        outEntityRole = ROLE_SONSTIGE
        outZuordnung = ExtrahiereAnzeigeName(kontoName)
        outParzellen = ""
        outDebugInfo = "Keine Zuordnung gefunden"
        outAmpelStatus = 3
        Exit Sub
    End If
    
    
'--- Ende Teil 4 ---
'--- Anfang Teil 5 ---


    If mitgliederExakt.Count = 1 Then
        mitgliedInfo = mitgliederExakt(1)
        
        If mitgliedInfo(6) = True Then
            outEntityKey = PREFIX_EHEMALIG & mitgliedInfo(0)
            outEntityRole = ROLE_EHEMALIGES_MITGLIED
            outZuordnung = mitgliedInfo(2) & ", " & mitgliedInfo(1) & " (Ehemalig)"
            outParzellen = mitgliedInfo(3)
            outDebugInfo = "Exakter Treffer (Historie)"
            outAmpelStatus = 1
        Else
            outEntityKey = mitgliedInfo(0)
            If mitgliedInfo(7) Then
                outEntityRole = ROLE_MITGLIED_MIT_PACHT
            Else
                outEntityRole = ROLE_MITGLIED_OHNE_PACHT
            End If
            outZuordnung = mitgliedInfo(2) & ", " & mitgliedInfo(1)
            outParzellen = mitgliedInfo(3)
            outDebugInfo = "Exakter Treffer (aktiv)"
            outAmpelStatus = 1
        End If
        Exit Sub
    End If
    
    hatAktiveMitglieder = False
    hatEhemaligeMitglieder = False
    
    For i = 1 To mitgliederExakt.Count
        mitgliedInfo = mitgliederExakt(i)
        If mitgliedInfo(6) = False Then
            hatAktiveMitglieder = True
        Else
            hatEhemaligeMitglieder = True
        End If
        
        If Not uniqueMemberIDs.Exists(CStr(mitgliedInfo(0))) Then
            uniqueMemberIDs.Add CStr(mitgliedInfo(0)), mitgliedInfo
        End If
    Next i
    
    If uniqueMemberIDs.Count > 1 Then
        outEntityKey = PREFIX_SHARE & CreateGUID()
        
        If hatAktiveMitglieder And Not hatEhemaligeMitglieder Then
            outEntityRole = ROLE_MITGLIED_MIT_PACHT
        ElseIf Not hatAktiveMitglieder And hatEhemaligeMitglieder Then
            outEntityRole = ROLE_EHEMALIGES_MITGLIED
        Else
            outEntityRole = ROLE_MITGLIED_MIT_PACHT
        End If
        
        Dim namen As String
        Dim parzellenListe As String
        namen = ""
        parzellenListe = ""
        
        For Each key In uniqueMemberIDs.Keys
            mitgliedInfo = uniqueMemberIDs(key)
            If namen <> "" Then namen = namen & " / "
            namen = namen & mitgliedInfo(2) & ", " & mitgliedInfo(1)
            If mitgliedInfo(6) Then namen = namen & " (Eh.)"
            
            If mitgliedInfo(3) <> "" Then
                If InStr(parzellenListe, mitgliedInfo(3)) = 0 Then
                    If parzellenListe <> "" Then parzellenListe = parzellenListe & ", "
                    parzellenListe = parzellenListe & mitgliedInfo(3)
                End If
            End If
        Next key
        
        outZuordnung = namen
        outParzellen = parzellenListe
        outDebugInfo = "Gemeinsames Konto (" & uniqueMemberIDs.Count & " Mitglieder)"
        outAmpelStatus = 2
        Exit Sub
    Else
        mitgliedInfo = mitgliederExakt(1)
        
        If mitgliedInfo(6) = True Then
            outEntityKey = PREFIX_EHEMALIG & mitgliedInfo(0)
            outEntityRole = ROLE_EHEMALIGES_MITGLIED
            outZuordnung = mitgliedInfo(2) & ", " & mitgliedInfo(1) & " (Ehemalig)"
        Else
            outEntityKey = mitgliedInfo(0)
            If mitgliedInfo(7) Then
                outEntityRole = ROLE_MITGLIED_MIT_PACHT
            Else
                outEntityRole = ROLE_MITGLIED_OHNE_PACHT
            End If
            outZuordnung = mitgliedInfo(2) & ", " & mitgliedInfo(1)
        End If
        outParzellen = mitgliedInfo(3)
        outDebugInfo = "Eindeutige Zuordnung"
        outAmpelStatus = 1
    End If
    
End Sub

Private Function ExtrahiereAnzeigeName(ByVal kontoName As String) As String
    Dim result As String
    Dim pos As Long
    
    result = Trim(kontoName)
    
    pos = InStr(result, ",")
    If pos > 0 Then
        result = Trim(Left(result, pos - 1))
    End If
    
    If Len(result) > 50 Then
        result = Left(result, 47) & "..."
    End If
    
    ExtrahiereAnzeigeName = result
End Function

Public Function CreateGUID() As String
    Dim guid As String
    Dim i As Long
    Dim hexChars As String
    
    hexChars = "0123456789ABCDEF"
    guid = ""
    
    Randomize Timer
    
    For i = 1 To 8
        guid = guid & Mid(hexChars, Int(Rnd * 16) + 1, 1)
    Next i
    
    CreateGUID = guid
End Function


'--- Ende Teil 5 ---
'--- Anfang Teil 6 ---
    
    
Public Function IstVersorger(ByVal Name As String) As Boolean
    Dim keywords As Variant
    Dim kw As Variant
    Dim nameUpper As String
    
    keywords = Array( _
        "stadtwerke", "energieversorgung", "gasversorgung", _
        "stromversorgung", "wasserversorgung", "abwasser", _
        "entsorgung", "muellabfuhr", "abfallwirtschaft", _
        "telekom", "vodafone", "o2", "telefonica", _
        "1und1", "1&1", "unitymedia", "kabel deutschland", _
        "versicherung", "allianz", "axa", "ergo", "huk", _
        "debeka", "barmer", "aok", "tk", "dak", _
        "finanzamt", "gemeinde", "stadt ", "kreis", _
        "landratsamt", "bezirksamt", "behoerde", _
        "gez", "rundfunk", "beitragsservice", _
        "hausverwaltung", "wohnungsbau", "immobilien", _
        "enso", "drewag", "sachsenenergie", "ewe", "eon", _
        "rwe", "vattenfall", "enbw", "mainova", "entega")
    
    nameUpper = UCase(Trim(Name))
    
    If InStr(nameUpper, "LIDL") > 0 Then
        IstVersorger = False
        Exit Function
    End If
    
    IstVersorger = False
    
    For Each kw In keywords
        If InStr(nameUpper, UCase(kw)) > 0 Then
            IstVersorger = True
            Exit Function
        End If
    Next kw
    
End Function

Public Function IstBank(ByVal Name As String) As Boolean
    Dim keywords As Variant
    Dim kw As Variant
    Dim nameUpper As String
    
    keywords = Array( _
        "sparkasse", "volksbank", "raiffeisenbank", _
        "commerzbank", "deutsche bank", "postbank", _
        "ing-diba", "ing diba", "comdirect", "consorsbank", _
        "targobank", "santander", "sparda", "psd bank", _
        "dkb", "norisbank", "hypovereinsbank", "unicredit")
    
    nameUpper = UCase(Trim(Name))
    
    If nameUpper = "" Or nameUpper = "0" Then
        IstBank = False
        Exit Function
    End If
    
    For Each kw In keywords
        If InStr(nameUpper, UCase(kw)) > 0 Then
            IstBank = True
            Exit Function
        End If
    Next kw
    
    IstBank = False
End Function

Public Function IstBankTransaktion(ByVal iban As String, ByVal kontoName As String) As Boolean
    Dim wsBK As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim buchungsText As String
    Dim ibanNorm As String
    Dim gefundenIBAN As String
    
    IstBankTransaktion = False
    
    ibanNorm = UCase(Trim(Replace(iban, " ", "")))
    
    If ibanNorm = "0" Or ibanNorm = "3529000972" Then
        On Error Resume Next
        Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
        On Error GoTo 0
        
        If wsBK Is Nothing Then Exit Function
        
        lastRow = wsBK.Cells(wsBK.Rows.Count, BK_COL_DATUM).End(xlUp).Row
        
        For r = BK_START_ROW To lastRow
            gefundenIBAN = UCase(Trim(Replace(wsBK.Cells(r, BK_COL_IBAN).value, " ", "")))
            
            If gefundenIBAN = ibanNorm Then
                buchungsText = UCase(Trim(wsBK.Cells(r, BK_COL_BUCHUNGSTEXT).value))
                
                If InStr(buchungsText, "ENTGELTABSCHLUSS") > 0 Or _
                   InStr(buchungsText, "ABSCHLUSS") > 0 Or _
                   InStr(buchungsText, "RECHNUNG") > 0 Then
                    IstBankTransaktion = True
                    Exit Function
                End If
            End If
        Next r
    End If
    
    If Left(UCase(Trim(kontoName)), 5) = "GA NR" Then
        IstBankTransaktion = True
        Exit Function
    End If
    
    If Trim(kontoName) = "" And (ibanNorm = "0" Or ibanNorm = "3529000972") Then
        IstBankTransaktion = True
        Exit Function
    End If
    
End Function

Public Function IstShop(ByVal Name As String) As Boolean
    Dim keywords As Variant
    Dim kw As Variant
    Dim nameUpper As String
    
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
    
    nameUpper = UCase(Trim(Name))
    
    IstShop = False
    
    For Each kw In keywords
        If InStr(nameUpper, UCase(kw)) > 0 Then
            IstShop = True
            Exit Function
        End If
    Next kw
    
End Function


'--- Ende Teil 6 ---
'--- Anfang Teil 7 ---


Private Function NormalisiereRoleString(ByVal role As String) As String
    Dim result As String
    
    result = UCase(Trim(role))
    result = Replace(result, "_", "")
    result = Replace(result, "-", "")
    result = Replace(result, " ", "")
    
    If InStr(result, "VERSORGER") > 0 Then
        NormalisiereRoleString = "VERSORGER"
    ElseIf InStr(result, "BANK") > 0 Then
        NormalisiereRoleString = "BANK"
    ElseIf InStr(result, "SHOP") > 0 Then
        NormalisiereRoleString = "SHOP"
    ElseIf InStr(result, "SONSTIGE") > 0 Then
        NormalisiereRoleString = "SONSTIGE"
    ElseIf InStr(result, "EHEMALIG") > 0 Then
        NormalisiereRoleString = "EHEMALIG"
    ElseIf InStr(result, "MITGLIED") > 0 Then
        NormalisiereRoleString = "MITGLIED"
    Else
        NormalisiereRoleString = result
    End If
End Function

Private Sub FormatiereEntityKeyTabelle(ByRef ws As Worksheet, ByVal lastRow As Long)
    
    Dim rngTable As Range
    Dim rngZebra As Range
    Dim r As Long
    Dim currentRole As String
    
    If lastRow < EK_START_ROW Then Exit Sub
    
    Set rngTable = ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                            ws.Cells(lastRow, EK_COL_DEBUG))
    
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    rngTable.VerticalAlignment = xlCenter
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                  ws.Cells(lastRow, EK_COL_ENTITYKEY))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ENTITYKEY).ColumnWidth = 9
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_IBAN), _
                  ws.Cells(lastRow, EK_COL_IBAN))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_IBAN).ColumnWidth = 23
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_KONTONAME), _
                  ws.Cells(lastRow, EK_COL_KONTONAME))
        .WrapText = True
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_KONTONAME).ColumnWidth = 50
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ZUORDNUNG), _
                  ws.Cells(lastRow, EK_COL_ZUORDNUNG))
        .WrapText = True
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ZUORDNUNG).ColumnWidth = 30
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_PARZELLE), _
                  ws.Cells(lastRow, EK_COL_PARZELLE))
        .WrapText = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns(EK_COL_PARZELLE).ColumnWidth = 10
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_ROLE), _
                  ws.Cells(lastRow, EK_COL_ROLE))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_ROLE).AutoFit
    
    With ws.Range(ws.Cells(EK_START_ROW, EK_COL_DEBUG), _
                  ws.Cells(lastRow, EK_COL_DEBUG))
        .WrapText = False
        .HorizontalAlignment = xlLeft
    End With
    ws.Columns(EK_COL_DEBUG).AutoFit
    
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
             ws.Cells(lastRow, EK_COL_KONTONAME)).Locked = True
    
    For r = EK_START_ROW To lastRow
        currentRole = Trim(ws.Cells(r, EK_COL_ROLE).value)
        
        Call SetzeZellschutzFuerZeile(ws, r, currentRole)
        
        Set rngZebra = ws.Range(ws.Cells(r, EK_COL_ENTITYKEY), ws.Cells(r, EK_COL_KONTONAME))
        
        If (r - EK_START_ROW) Mod 2 = 1 Then
            rngZebra.Interior.color = ZEBRA_COLOR
        Else
            rngZebra.Interior.ColorIndex = xlNone
        End If
    Next r
    
    ws.Rows(EK_START_ROW & ":" & lastRow).AutoFit
    
End Sub

Public Sub FormatiereEntityKeyZeile(ByVal zeile As Long)
    
    Dim ws As Worksheet
    Dim currentRole As String
    Dim rngZebra As Range
    
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    currentRole = Trim(ws.Cells(zeile, EK_COL_ROLE).value)
    
    Call SetzeZellschutzFuerZeile(ws, zeile, currentRole)
    
    Set rngZebra = ws.Range(ws.Cells(zeile, EK_COL_ENTITYKEY), ws.Cells(zeile, EK_COL_KONTONAME))
    
    If (zeile - EK_START_ROW) Mod 2 = 1 Then
        rngZebra.Interior.color = ZEBRA_COLOR
    Else
        rngZebra.Interior.ColorIndex = xlNone
    End If
    
    ws.Rows(zeile).AutoFit
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
End Sub


'--- Ende Teil 7 ---
'--- Anfang Teil 8 ---


Public Sub OnEntityRoleChange(ByVal Target As Range)
    
    Dim ws As Worksheet
    Dim zeile As Long
    Dim neueRole As String
    Dim currentParzelle As String
    
    Set ws = Target.Worksheet
    zeile = Target.Row
    neueRole = Trim(Target.value)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    If Not DarfParzelleHaben(neueRole) Then
        currentParzelle = Trim(ws.Cells(zeile, EK_COL_PARZELLE).value)
        If currentParzelle <> "" Then
            Dim antwort As VbMsgBoxResult
            antwort = MsgBox("Der Role-Typ '" & neueRole & "' erlaubt keine Parzelle." & vbCrLf & _
                            "Die aktuelle Parzelle '" & currentParzelle & "' wird geloescht." & vbCrLf & vbCrLf & _
                            "Fortfahren?", vbYesNo + vbQuestion, "Parzelle loeschen?")
            If antwort = vbYes Then
                ws.Cells(zeile, EK_COL_PARZELLE).value = ""
            Else
                Application.EnableEvents = False
                Target.value = ""
                Application.EnableEvents = True
                ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
                Exit Sub
            End If
        End If
    End If
    
    Call SetzeZellschutzFuerZeile(ws, zeile, neueRole)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
End Sub

Public Sub LoescheAlleEntityKeys()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim antwort As VbMsgBoxResult
    
    antwort = MsgBox("ACHTUNG: Diese Aktion loescht ALLE EntityKey-Daten!" & vbCrLf & vbCrLf & _
                     "Dies umfasst:" & vbCrLf & _
                     "- Alle EntityKeys (Spalte R)" & vbCrLf & _
                     "- Alle IBANs (Spalte S)" & vbCrLf & _
                     "- Alle Kontonamen (Spalte T)" & vbCrLf & _
                     "- Alle Zuordnungen (Spalte U)" & vbCrLf & _
                     "- Alle Parzellen (Spalte V)" & vbCrLf & _
                     "- Alle Roles (Spalte W)" & vbCrLf & _
                     "- Alle Debug-Infos (Spalte X)" & vbCrLf & vbCrLf & _
                     "Sind Sie SICHER?", vbYesNo + vbCritical, "Alle EntityKeys loeschen?")
    
    If antwort <> vbYes Then Exit Sub
    
    antwort = MsgBox("LETZTE WARNUNG!" & vbCrLf & vbCrLf & _
                     "Diese Aktion kann NICHT rueckgaengig gemacht werden!" & vbCrLf & vbCrLf & _
                     "Wirklich ALLE EntityKey-Daten loeschen?", _
                     vbYesNo + vbCritical, "Endgueltige Bestaetigung")
    
    If antwort <> vbYes Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.Count, EK_COL_IBAN).End(xlUp).Row
    If lastRow < EK_START_ROW Then lastRow = EK_START_ROW
    
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
             ws.Cells(lastRow, EK_COL_DEBUG)).ClearContents
    
    ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
             ws.Cells(lastRow, EK_COL_DEBUG)).Interior.ColorIndex = xlNone
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Alle EntityKey-Daten wurden geloescht.", vbInformation
    
End Sub

Public Sub ExportiereEntityKeysAlsCSV()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim dateiPfad As String
    Dim fileNum As Integer
    Dim zeile As String
    
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    lastRow = ws.Cells(ws.Rows.Count, EK_COL_IBAN).End(xlUp).Row
    If lastRow < EK_START_ROW Then
        MsgBox "Keine EntityKey-Daten zum Exportieren vorhanden.", vbInformation
        Exit Sub
    End If
    
    dateiPfad = Application.GetSaveAsFilename( _
        InitialFileName:="EntityKeys_Export_" & Format(Date, "YYYYMMDD") & ".csv", _
        FileFilter:="CSV-Dateien (*.csv),*.csv", _
        Title:="EntityKeys exportieren")
    
    If dateiPfad = "Falsch" Or dateiPfad = "" Then Exit Sub
    
    fileNum = FreeFile
    Open dateiPfad For Output As #fileNum
    
    Print #fileNum, "EntityKey;IBAN;Kontoname;Zuordnung;Parzelle;Role;Debug"
    
    For r = EK_START_ROW To lastRow
        zeile = ws.Cells(r, EK_COL_ENTITYKEY).value & ";" & _
                ws.Cells(r, EK_COL_IBAN).value & ";" & _
                ws.Cells(r, EK_COL_KONTONAME).value & ";" & _
                ws.Cells(r, EK_COL_ZUORDNUNG).value & ";" & _
                ws.Cells(r, EK_COL_PARZELLE).value & ";" & _
                ws.Cells(r, EK_COL_ROLE).value & ";" & _
                ws.Cells(r, EK_COL_DEBUG).value
        Print #fileNum, zeile
    Next r
    
    Close #fileNum
    
    MsgBox "Export abgeschlossen!" & vbCrLf & vbCrLf & _
           "Datei: " & dateiPfad & vbCrLf & _
           "Zeilen exportiert: " & (lastRow - EK_START_ROW + 1), vbInformation
    
End Sub



