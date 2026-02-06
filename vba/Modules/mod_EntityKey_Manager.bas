Attribute VB_Name = "mod_EntityKey_Manager"
Option Explicit

' ***************************************************************
' MODUL: mod_EntityKey_Manager
' ZWECK: Verwaltung und Zuordnung von EntityKeys fuer Bankverkehr
' VERSION: 5.3.1 - 06.02.2026
' FIX: Manuelle SHOP/VERSORGER/BANK/SONSTIGE immer GRUEN
' NEU: EHEMALIGES MITGLIED -> InputBox Parzelle 1-14 wenn nicht in Historie
' NEU: EHEMALIGES MITGLIED in Historie -> GRUEN
' NEU: AktualisiereEntityKeyBeiAustritt (EX-Prefix bei Mitglied-Austritt)
' FIX: Debug-Spalte nur Datum (kein Uhrzeit) bei manuellen Zuordnungen
' FIX: Spalte X Breite 65 (via mod_Formatierung)
' FIX: Mehrere Parzellen pro Mitglied werden kommagetrennt in V angezeigt
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
' AMPELFARBEN: Setzt Farbe fuer Spalten U-X einer Zeile
' ampelStatus: 1=Gruen, 2=Gelb, 3=Rot
' ===============================================================
Private Sub SetzeAmpelFarbe(ByRef ws As Worksheet, ByVal zeile As Long, ByVal ampelStatus As Long)
    Dim rngAmpel As Range
    Dim farbe As Long
    
    Set rngAmpel = ws.Range(ws.Cells(zeile, EK_COL_ZUORDNUNG), _
                            ws.Cells(zeile, EK_COL_DEBUG))
    
    Select Case ampelStatus
        Case 1
            farbe = RGB(198, 224, 180)  ' Gruen
        Case 2
            farbe = RGB(255, 230, 153)  ' Gelb
        Case 3
            farbe = RGB(255, 150, 150)  ' Rot
        Case Else
            farbe = RGB(198, 224, 180)  ' Default Gruen
    End Select
    
    rngAmpel.Interior.color = farbe
End Sub

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
        
        If Not isEmpty(currentDatum) And currentDatum <> "" Then
            currentIBAN = NormalisiereIBAN(wsBK.Cells(r, BK_COL_IBAN).value)
            currentKontoName = EntferneMehrfacheLeerzeichen(Trim(CStr(wsBK.Cells(r, BK_COL_NAME).value)))
            
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
                            If Not IstKontonameRedundant(dictKontonamen(currentIBAN), currentKontoName) Then
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
            currentIBAN = NormalisiereIBAN(wsD.Cells(r, EK_COL_IBAN).value)
            If currentIBAN <> "" And dictKontonamen.Exists(currentIBAN) Then
                Dim bereinigteNamen As Object
                Set bereinigteNamen = BereinigeKontonamen(dictKontonamen(currentIBAN))
                Dim allNames As String
                allNames = SammelKontonamen(bereinigteNamen)
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
            Set bereinigt = BereinigeKontonamen(dictKontonamen(currentIBAN))
            Dim kontoNamenGesamt As String
            kontoNamenGesamt = SammelKontonamen(bereinigt)
            
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
' HILFSFUNKTION: Prueft ob ein Kontoname semantisch redundant ist
' ===============================================================
Private Function IstKontonameRedundant(ByRef dictNames As Object, ByVal neuerName As String) As Boolean
    Dim key As Variant
    Dim bestehenderName As String
    Dim neueWorte As Object
    Dim bestehendeWorte As Object
    
    IstKontonameRedundant = False
    
    Set neueWorte = ZerlegeInWorte(neuerName)
    
    For Each key In dictNames.keys
        bestehenderName = CStr(dictNames(key))
        Set bestehendeWorte = ZerlegeInWorte(bestehenderName)
        
        If SindWortmengenGleich(neueWorte, bestehendeWorte) Then
            IstKontonameRedundant = True
            Exit Function
        End If
        
        If IstTeilmenge(neueWorte, bestehendeWorte) Then
            IstKontonameRedundant = True
            Exit Function
        End If
    Next key
End Function

' ===============================================================
' HILFSFUNKTION: Zerlegt einen Namen in normalisierte Worte
' ===============================================================
Private Function ZerlegeInWorte(ByVal Name As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim parts() As String
    Dim wort As String
    Dim i As Long
    
    Name = UCase(Trim(Name))
    Name = Replace(Name, ",", " ")
    Name = Replace(Name, ".", " ")
    Name = Replace(Name, "-", " ")
    Name = EntferneMehrfacheLeerzeichen(Name)
    
    If Name = "" Then
        Set ZerlegeInWorte = dict
        Exit Function
    End If
    
    parts = Split(Name, " ")
    
    For i = LBound(parts) To UBound(parts)
        wort = Trim(parts(i))
        If wort <> "" And wort <> "UND" And wort <> "U" Then
            If Not dict.Exists(wort) Then
                dict.Add wort, True
            End If
        End If
    Next i
    
    Set ZerlegeInWorte = dict
End Function

' ===============================================================
' HILFSFUNKTION: Prueft ob zwei Wortmengen identisch sind
' ===============================================================
Private Function SindWortmengenGleich(ByRef dict1 As Object, ByRef dict2 As Object) As Boolean
    Dim key As Variant
    
    SindWortmengenGleich = False
    
    If dict1.count <> dict2.count Then Exit Function
    If dict1.count = 0 Then Exit Function
    
    For Each key In dict1.keys
        If Not dict2.Exists(key) Then Exit Function
    Next key
    
    SindWortmengenGleich = True
End Function

' ===============================================================
' HILFSFUNKTION: Prueft ob dict1 eine Teilmenge von dict2 ist
' ===============================================================
Private Function IstTeilmenge(ByRef dictKlein As Object, ByRef dictGross As Object) As Boolean
    Dim key As Variant
    
    IstTeilmenge = False
    
    If dictKlein.count = 0 Then Exit Function
    If dictKlein.count >= dictGross.count Then Exit Function
    
    For Each key In dictKlein.keys
        If Not dictGross.Exists(key) Then Exit Function
    Next key
    
    IstTeilmenge = True
End Function

' ===============================================================
' HILFSFUNKTION: Bereinigt Dictionary von redundanten Kontonamen
' ===============================================================
Private Function BereinigeKontonamen(ByRef dictNames As Object) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    Dim keys() As Variant
    Dim values() As Variant
    Dim istRedundant() As Boolean
    Dim i As Long, j As Long
    Dim cnt As Long
    Dim worteI As Object, worteJ As Object
    
    cnt = dictNames.count
    If cnt = 0 Then
        Set BereinigeKontonamen = result
        Exit Function
    End If
    
    If cnt = 1 Then
        Dim singleKey As Variant
        For Each singleKey In dictNames.keys
            result.Add singleKey, dictNames(singleKey)
        Next singleKey
        Set BereinigeKontonamen = result
        Exit Function
    End If
    
    ReDim keys(0 To cnt - 1)
    ReDim values(0 To cnt - 1)
    ReDim istRedundant(0 To cnt - 1)
    
    i = 0
    Dim k As Variant
    For Each k In dictNames.keys
        keys(i) = k
        values(i) = dictNames(k)
        istRedundant(i) = False
        i = i + 1
    Next k
    
    For i = 0 To cnt - 1
        If Not istRedundant(i) Then
            Set worteI = ZerlegeInWorte(CStr(values(i)))
            For j = i + 1 To cnt - 1
                If Not istRedundant(j) Then
                    Set worteJ = ZerlegeInWorte(CStr(values(j)))
                    
                    If SindWortmengenGleich(worteI, worteJ) Then
                        If Len(CStr(values(i))) >= Len(CStr(values(j))) Then
                            istRedundant(j) = True
                        Else
                            istRedundant(i) = True
                            Exit For
                        End If
                    ElseIf IstTeilmenge(worteI, worteJ) Then
                        istRedundant(i) = True
                        Exit For
                    ElseIf IstTeilmenge(worteJ, worteI) Then
                        istRedundant(j) = True
                    End If
                End If
            Next j
        End If
    Next i
    
    For i = 0 To cnt - 1
        If Not istRedundant(i) Then
            result.Add keys(i), values(i)
        End If
    Next i
    
    Set BereinigeKontonamen = result
End Function

' ===============================================================
' HILFSFUNKTION: Sammelt alle Kontonamen aus Dictionary zu String
' ===============================================================
Private Function SammelKontonamen(ByRef dictNames As Object) As String
    Dim key As Variant
    Dim result As String
    Dim cleanName As String
    
    result = ""
    
    For Each key In dictNames.keys
        cleanName = EntferneMehrfacheLeerzeichen(Trim(CStr(dictNames(key))))
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
' NEU v5.2: Prueft ob Kontoname eine Geldautomat-Abhebung ist
' Muster: IBAN="0", Name beginnt mit "GA " und enthaelt "BLZ"
' ===============================================================
Private Function IstGeldautomatAbhebung(ByVal iban As String, ByVal kontoname As String) As Boolean
    Dim normIBAN As String
    Dim nameUpper As String
    
    IstGeldautomatAbhebung = False
    normIBAN = NormalisiereIBAN(iban)
    
    If normIBAN <> "0" Then Exit Function
    
    nameUpper = UCase(Trim(kontoname))
    
    If Left(nameUpper, 3) = "GA " And InStr(nameUpper, "BLZ") > 0 Then
        IstGeldautomatAbhebung = True
    End If
End Function

' ===============================================================
' NEU v5.2: Prueft ob ehemaliges Mitglied in Mitgliederhistorie steht
' Gibt True zurueck wenn in Historie gefunden
' ===============================================================
Private Function PruefeObInHistorie(ByVal kontoname As String, ByRef wsH As Worksheet) As Boolean
    Dim r As Long
    Dim lastRow As Long
    Dim nachnameHist As String
    Dim kontoNameNorm As String
    Dim nachnameNorm As String
    
    PruefeObInHistorie = False
    
    If kontoname = "" Then Exit Function
    
    kontoNameNorm = NormalisiereStringFuerVergleich(kontoname)
    If kontoNameNorm = "" Then Exit Function
    
    lastRow = wsH.Cells(wsH.Rows.count, H_COL_NAME_EHEM_PAECHTER).End(xlUp).Row
    
    For r = H_START_ROW To lastRow
        nachnameHist = Trim(wsH.Cells(r, H_COL_NAME_EHEM_PAECHTER).value)
        If nachnameHist <> "" Then
            nachnameNorm = NormalisiereStringFuerVergleich(nachnameHist)
            If nachnameNorm <> "" And Len(nachnameNorm) >= 3 Then
                If InStr(kontoNameNorm, nachnameNorm) > 0 Then
                    PruefeObInHistorie = True
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

' ===============================================================
' NEU v5.3: Prueft ob ehemaliges Mitglied in Historie steht
' und gibt die Parzelle aus der Historie zurueck
' ===============================================================
Private Function HoleParzelleFuerEhemaligesAusHistorie(ByVal kontoname As String, ByRef wsH As Worksheet) As String
    Dim r As Long
    Dim lastRow As Long
    Dim nachnameHist As String
    Dim kontoNameNorm As String
    Dim nachnameNorm As String
    
    HoleParzelleFuerEhemaligesAusHistorie = ""
    
    If kontoname = "" Then Exit Function
    
    kontoNameNorm = NormalisiereStringFuerVergleich(kontoname)
    If kontoNameNorm = "" Then Exit Function
    
    lastRow = wsH.Cells(wsH.Rows.count, H_COL_NAME_EHEM_PAECHTER).End(xlUp).Row
    
    For r = H_START_ROW To lastRow
        nachnameHist = Trim(wsH.Cells(r, H_COL_NAME_EHEM_PAECHTER).value)
        If nachnameHist <> "" Then
            nachnameNorm = NormalisiereStringFuerVergleich(nachnameHist)
            If nachnameNorm <> "" And Len(nachnameNorm) >= 3 Then
                If InStr(kontoNameNorm, nachnameNorm) > 0 Then
                    HoleParzelleFuerEhemaligesAusHistorie = Trim(CStr(wsH.Cells(r, H_COL_PARZELLE).value))
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

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
        kontoname = EntferneMehrfacheLeerzeichen(Trim(CStr(wsD.Cells(r, EK_COL_KONTONAME).value)))
        
        If CStr(wsD.Cells(r, EK_COL_KONTONAME).value) <> kontoname Then
            wsD.Cells(r, EK_COL_KONTONAME).value = kontoname
        End If
        
        currentEntityKey = Trim(wsD.Cells(r, EK_COL_ENTITYKEY).value)
        currentZuordnung = Trim(wsD.Cells(r, EK_COL_ZUORDNUNG).value)
        currentParzelle = Trim(wsD.Cells(r, EK_COL_PARZELLE).value)
        currentRole = Trim(wsD.Cells(r, EK_COL_ROLE).value)
        
        If iban = "" And kontoname = "" Then GoTo nextRow
        
        ' NEU v5.2: Pruefe Geldautomat-Abhebung VOR Bankabschluss
        If currentEntityKey = "" Then
            If IstGeldautomatAbhebung(iban, kontoname) Then
                wsD.Cells(r, EK_COL_ENTITYKEY).value = PREFIX_BANK & CreateGUID()
                wsD.Cells(r, EK_COL_ZUORDNUNG).value = "Bargeldabhebung Geldautomat (Vereinskasse)"
                wsD.Cells(r, EK_COL_ROLE).value = ROLE_BANK
                wsD.Cells(r, EK_COL_DEBUG).value = "Geldautomat erkannt (GA + BLZ)"
                Call SetzeAmpelFarbe(wsD, r, 1)
                GoTo nextRow
            End If
        End If
        
        ' Pruefe IBAN "0" oder "3529000972" + ABSCHLUSS
        If currentEntityKey = "" Then
            If IstBankAbschluss(iban, wsBK) Then
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
    
    ' Formatierung ZUERST (inkl. Sortierung)
    Call mod_Formatierung.FormatEntityKeyTableComplete(wsD)
    
    ' Ampelfarben DANACH (nach Sortierung!)
    Call SetzeAlleAmpelfarbenNachSortierung(wsD)
    
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
' NEU v5.0: Setzt Ampelfarben fuer ALLE Zeilen NACH Sortierung
' FIX v5.3: EHEMALIGES MITGLIED in Historie -> GRUEN
' ===============================================================
Public Sub SetzeAlleAmpelfarbenNachSortierung(ByRef wsD As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    Dim entityKey As String
    Dim zuordnung As String
    Dim role As String
    Dim debugTxt As String
    Dim ampel As Long
    Dim kontoname As String
    Dim wsH As Worksheet
    
    On Error Resume Next
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    On Error GoTo 0
    
    lastRow = wsD.Cells(wsD.Rows.count, EK_COL_IBAN).End(xlUp).Row
    Dim lastRowR As Long
    lastRowR = wsD.Cells(wsD.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    If lastRowR > lastRow Then lastRow = lastRowR
    If lastRow < EK_START_ROW Then Exit Sub
    
    For r = EK_START_ROW To lastRow
        entityKey = Trim(wsD.Cells(r, EK_COL_ENTITYKEY).value)
        zuordnung = Trim(wsD.Cells(r, EK_COL_ZUORDNUNG).value)
        role = Trim(wsD.Cells(r, EK_COL_ROLE).value)
        debugTxt = Trim(wsD.Cells(r, EK_COL_DEBUG).value)
        
        ampel = BerechneAmpelStatus(entityKey, zuordnung, role, debugTxt)
        
        ' NEU v5.3: Bei EHEMALIGES MITGLIED pruefen ob in Historie
        If UCase(role) = "EHEMALIGES MITGLIED" Then
            If Not wsH Is Nothing Then
                kontoname = Trim(CStr(wsD.Cells(r, EK_COL_KONTONAME).value))
                If PruefeObInHistorie(kontoname, wsH) Then
                    ' In Historie gefunden -> GRUEN
                    ampel = 1
                Else
                    ' Nicht in Historie -> GELB
                    ampel = 2
                    ' Hinweis nur anfuegen wenn KEIN "nicht in Historie" bereits im Text
                    If InStr(debugTxt, "nicht in Historie") = 0 Then
                        If debugTxt <> "" Then
                            debugTxt = debugTxt & " | nicht in Historie"
                        Else
                            debugTxt = "nicht in Historie"
                        End If
                        wsD.Cells(r, EK_COL_DEBUG).value = debugTxt
                    End If
                End If
            End If
        End If
        
        Call SetzeAmpelFarbe(wsD, r, ampel)
    Next r
End Sub

' ===============================================================
' Berechnet den korrekten Ampelstatus einer Zeile
'
' GRUEN (1) = 100% sicher zugeordnet
' GELB (2) = Treffer unsicher / Teiluebereinstimmung
' ROT (3) = Kein Treffer, Nutzer MUSS manuell in W zuordnen
' FIX v5.3: EHEMALIGES MITGLIED default GELB (Historie-Check extern)
' ===============================================================
Private Function BerechneAmpelStatus(ByVal entityKey As String, _
                                      ByVal zuordnung As String, _
                                      ByVal role As String, _
                                      ByVal debugTxt As String) As Long
    Dim debugUpper As String
    debugUpper = UCase(debugTxt)
    
    ' ROT: Kein EntityKey UND kein Role
    If entityKey = "" And role = "" Then
        BerechneAmpelStatus = 3
        Exit Function
    End If
    
    ' ROT: Debug sagt KEIN TREFFER und keine Role gesetzt
    If InStr(debugUpper, "KEIN TREFFER") > 0 And role = "" Then
        BerechneAmpelStatus = 3
        Exit Function
    End If
    
    ' GELB: Nur Nachname gefunden, unsicher
    If InStr(debugUpper, "NUR NACHNAME") > 0 Then
        BerechneAmpelStatus = 2
        Exit Function
    End If
    
    ' GELB: EntityKey fehlt, obwohl Role vorhanden
    If entityKey = "" And role <> "" Then
        BerechneAmpelStatus = 2
        Exit Function
    End If
    
    ' GELB: Role fehlt, obwohl EntityKey vorhanden
    If role = "" And entityKey <> "" Then
        BerechneAmpelStatus = 2
        Exit Function
    End If
    
    ' GELB: Ehemaliges Mitglied (Historie-Check in SetzeAlleAmpelfarbenNachSortierung)
    If UCase(role) = "EHEMALIGES MITGLIED" Then
        BerechneAmpelStatus = 2
        Exit Function
    End If
    
    ' GRUEN: Alles vorhanden und sicher
    If entityKey <> "" And role <> "" Then
        If zuordnung <> "" Then
            BerechneAmpelStatus = 1
        Else
            BerechneAmpelStatus = 2
        End If
        Exit Function
    End If
    
    ' Default: Gelb
    BerechneAmpelStatus = 2
End Function

' ===============================================================
' Prueft ob IBAN eine Bank-Abschluss-IBAN ist
' FIX v5.2: Schliesst Geldautomat-Abhebung aus
' ===============================================================
Private Function IstBankAbschluss(ByVal iban As String, ByRef wsBK As Worksheet) As Boolean
    Dim normIBAN As String
    Dim r As Long
    Dim lastRow As Long
    Dim bkIBAN As String
    Dim buchungstext As String
    Dim bkKontoname As String
    
    IstBankAbschluss = False
    normIBAN = NormalisiereIBAN(iban)
    
    If normIBAN <> "0" And normIBAN <> "3529000972" Then Exit Function
    
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    For r = BK_START_ROW To lastRow
        bkIBAN = NormalisiereIBAN(wsBK.Cells(r, BK_COL_IBAN).value)
        If bkIBAN = normIBAN Then
            ' NEU v5.2: Geldautomat ausschliessen
            bkKontoname = Trim(CStr(wsBK.Cells(r, BK_COL_NAME).value))
            If IstGeldautomatAbhebung(CStr(wsBK.Cells(r, BK_COL_IBAN).value), bkKontoname) Then
                GoTo naechsteZeile
            End If
            
            buchungstext = UCase(Trim(CStr(wsBK.Cells(r, BK_COL_BUCHUNGSTEXT).value)))
            If InStr(buchungstext, "ABSCHLUSS") > 0 Or _
               InStr(buchungstext, "ENTGELTABSCHLUSS") > 0 Then
                IstBankAbschluss = True
                Exit Function
            End If
        End If
naechsteZeile:
    Next r
End Function

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
' Sucht Mitglieder im Kontonamen
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
' Prueft Namens-Match
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
' Normalisiert String fuer Vergleich
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
' Prueft ob MemberID bereits gefunden
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
' Generiert EntityKey und Zuordnung
' FIX v5.1: Debug-Spalte zeigt bei VERSORGER den Zweck
'           Gemeinschaftskonto zeigt "automatisch erkannt"
' FIX v5.2: Umlaute in sichtbaren Texten
' FIX v5.3.1: Alle Parzellen pro Mitglied via HoleAlleParzellen
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
        
        If IstShop(kontoname) Then
            outEntityKey = PREFIX_SHOP & CreateGUID()
            outEntityRole = ROLE_SHOP
            outZuordnung = ExtrahiereAnzeigeName(kontoname)
            outDebugInfo = "Automatisch als SHOP erkannt"
            outAmpelStatus = 1
            Exit Sub
        End If
        
        ' VERSORGER mit Zweck-Erkennung
        Dim versorgerZweck As String
        versorgerZweck = ErmittleVersorgerZweck(kontoname)
        If versorgerZweck <> "" Then
            outEntityKey = PREFIX_VERSORGER & CreateGUID()
            outEntityRole = ROLE_VERSORGER
            outZuordnung = ExtrahiereAnzeigeName(kontoname)
            outDebugInfo = "Automatisch als VERSORGER erkannt (" & versorgerZweck & ")"
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
                ' FIX v5.3.1: Alle Parzellen des Mitglieds sammeln (nicht nur erste)
                outParzellen = HoleAlleParzellen(CStr(mitgliedInfo(0)), wsM)
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
                    
                    ' FIX v5.3.1: Alle Parzellen pro Person sammeln (nicht nur erste)
                    Dim personParzellen As String
                    personParzellen = HoleAlleParzellen(CStr(mitgliedInfo(0)), wsM)
                    
                    If personParzellen <> "" Then
                        If outParzellen <> "" Then
                            ' Nur Parzellen anfuegen die noch nicht enthalten sind
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
' Ermittelt EntityRole aus Funktion
' ===============================================================
Private Function ErmittleEntityRoleVonFunktion(ByVal funktion As String) As String
    Dim funktionUpper As String
    funktionUpper = UCase(funktion)
    
    If InStr(funktionUpper, "OHNE PACHT") > 0 Then
        ErmittleEntityRoleVonFunktion = ROLE_MITGLIED_OHNE_PACHT
    ElseIf InStr(funktionUpper, "EHEMALIG") > 0 Then
        ErmittleEntityRoleVonFunktion = ROLE_EHEMALIGES_MITGLIED
    Else
        ErmittleEntityRoleVonFunktion = ROLE_MITGLIED_MIT_PACHT
    End If
End Function

' ===============================================================
' IstShop - Keyword-Liste
' ===============================================================
Private Function IstShop(ByVal kontoname As String) As Boolean
    Dim n As String
    Dim keywords As Variant
    Dim i As Long
    
    n = UCase(Trim(kontoname))
    IstShop = False
    If Len(n) = 0 Then Exit Function
    
    keywords = Array( _
        "LIDL", "ALDI", "REWE", "EDEKA", "PENNY", "NETTO", "KAUFLAND", _
        "NORMA", "REAL", "ROSSMANN", "DM-DROGERIE", "MUELLER DROGERIE", _
        "BAUHAUS", "HORNBACH", "OBI", "HAGEBAU", "TOOM", "HELLWEG", _
        "GLOBUS BAUMARKT", "BAYWA", "RAIFFEISEN MARKT", _
        "AMAZON", "EBAY", "ZALANDO", "OTTO", "MEDIAMARKT", "SATURN", _
        "CONRAD ELECTRONIC", "ALTERNATE", "NOTEBOOKSBILLIGER", _
        "IKEA", "POCO", "ROLLER", "XXX LUTZ", _
        "DEHNER", "PFLANZEN KOELLE", "OVERKAMP", _
        "ARAL", "SHELL", "TOTAL", "ESSO", "JET TANKSTELLE", "TANKSTELLE", _
        "PAYPAL", "KLARNA", "SUMUP", _
        "FRESSNAPF", "ZOOPLUS", "DAS FUTTERHAUS", _
        "APOTHEKE", "FIELMANN", "APOLLO OPTIK", _
        "ACTION", "TEDI", "WOOLWORTH", "KIK", _
        "DECATHLON", "INTERSPORT", _
        "H&M", "C&A", "PRIMARK", "DEICHMANN" _
    )
    
    For i = LBound(keywords) To UBound(keywords)
        If InStr(n, CStr(keywords(i))) > 0 Then
            IstShop = True
            Exit Function
        End If
    Next i
End Function

' ===============================================================
' NEU v5.1: ErmittleVersorgerZweck - Gibt den Zweck zurueck
' oder "" wenn kein Versorger erkannt.
' Ersetzt IstVersorger - prueft UND gibt den Grund zurueck.
' FIX v5.2: Umlaute in sichtbaren Texten
' ===============================================================
Private Function ErmittleVersorgerZweck(ByVal kontoname As String) As String
    Dim n As String
    
    n = UCase(Trim(kontoname))
    ErmittleVersorgerZweck = ""
    If Len(n) = 0 Then Exit Function
    
    ' --- Wasser / Abwasser ---
    If InStr(n, "WAZV") > 0 Then
        ErmittleVersorgerZweck = "Wasser/Abwasser Zweckverband"
        Exit Function
    End If
    If InStr(n, "BRAUCHWASSER") > 0 Or InStr(n, "EIGENBETRIEB") > 0 Then
        ErmittleVersorgerZweck = "Brauchwasserversorgung"
        Exit Function
    End If
    If InStr(n, "WASSER") > 0 Or InStr(n, "ABWASSER") > 0 Then
        ErmittleVersorgerZweck = "Wasser/Abwasser"
        Exit Function
    End If
    If InStr(n, "BWB") > 0 Or InStr(n, "BERLINER WASSERBETRIEBE") > 0 Then
        ErmittleVersorgerZweck = "Wasser/Abwasser"
        Exit Function
    End If
    If InStr(n, "ZWECKVERBAND") > 0 Then
        ErmittleVersorgerZweck = "Zweckverband"
        Exit Function
    End If
    
    ' --- Strom / Energie ---
    If InStr(n, "STADTWERK") > 0 Or InStr(n, "ENERGIE") > 0 Or InStr(n, "STROM") > 0 Then
        ErmittleVersorgerZweck = "Strom/Energie"
        Exit Function
    End If
    If InStr(n, "VATTENFALL") > 0 Or InStr(n, "E.ON") > 0 Or InStr(n, "EON") > 0 Then
        ErmittleVersorgerZweck = "Strom/Energie"
        Exit Function
    End If
    If InStr(n, "RWE") > 0 Or InStr(n, "ENVIA") > 0 Or InStr(n, "ENVIAM") > 0 Then
        ErmittleVersorgerZweck = "Strom/Energie"
        Exit Function
    End If
    If InStr(n, "ENBW") > 0 Or InStr(n, "MAINOVA") > 0 Or InStr(n, "ENTEGA") > 0 Then
        ErmittleVersorgerZweck = "Strom/Energie"
        Exit Function
    End If
    
    ' --- Gas / Heizung ---
    If InStr(n, "GASAG") > 0 Or InStr(n, "GAS") > 0 Then
        ErmittleVersorgerZweck = "Gas/Heizung"
        Exit Function
    End If
    If InStr(n, "FERNWAERME") > 0 Or InStr(n, "HEIZUNG") > 0 Then
        ErmittleVersorgerZweck = "Fernw" & ChrW(228) & "rme/Heizung"
        Exit Function
    End If
    
    ' --- Versicherung ---
    If InStr(n, "VERSICHERUNG") > 0 Or InStr(n, "ALLIANZ") > 0 Or InStr(n, "DEVK") > 0 Then
        ErmittleVersorgerZweck = "Versicherung"
        Exit Function
    End If
    If InStr(n, "HUK") > 0 Or InStr(n, "HDI") > 0 Or InStr(n, "ERGO") > 0 Then
        ErmittleVersorgerZweck = "Versicherung"
        Exit Function
    End If
    If InStr(n, "GENERALI") > 0 Or InStr(n, "AXA") > 0 Or InStr(n, "ZURICH") > 0 Then
        ErmittleVersorgerZweck = "Versicherung"
        Exit Function
    End If
    If InStr(n, "WUERTTEMBERGISCHE") > 0 Then
        ErmittleVersorgerZweck = "Versicherung"
        Exit Function
    End If
    
    ' --- Telekommunikation ---
    If InStr(n, "TELEKOM") > 0 Or InStr(n, "VODAFONE") > 0 Or InStr(n, "1&1") > 0 Then
        ErmittleVersorgerZweck = "Telekommunikation"
        Exit Function
    End If
    If InStr(n, "O2") > 0 Or InStr(n, "TELEFONICA") > 0 Then
        ErmittleVersorgerZweck = "Telekommunikation"
        Exit Function
    End If
    If InStr(n, "KABEL DEUTSCHLAND") > 0 Or InStr(n, "UNITYMEDIA") > 0 Then
        ErmittleVersorgerZweck = "Telekommunikation"
        Exit Function
    End If
    
    ' --- Abfall / Entsorgung ---
    If InStr(n, "BSR") > 0 Or InStr(n, "ENTSORGUNG") > 0 Or InStr(n, "STADTREINIGUNG") > 0 Then
        ErmittleVersorgerZweck = "Abfallwirtschaft/Entsorgung"
        Exit Function
    End If
    If InStr(n, "ABFALLWIRTSCHAFT") > 0 Or InStr(n, "ABFALL") > 0 Then
        ErmittleVersorgerZweck = "Abfallwirtschaft/Entsorgung"
        Exit Function
    End If
    If InStr(n, "LANDKREIS") > 0 Then
        ErmittleVersorgerZweck = "Abfallwirtschaft (Landkreis)"
        Exit Function
    End If
    
    ' --- Grundsteuer / Finanzamt ---
    If InStr(n, "GRUNDSTEUER") > 0 Or InStr(n, "FINANZAMT") > 0 Then
        ErmittleVersorgerZweck = "Grundsteuer/Steuern"
        Exit Function
    End If
    If InStr(n, "STADT WERDER") > 0 Or InStr(n, "STADT WERDER (HAVEL)") > 0 Then
        ErmittleVersorgerZweck = "Grundsteuer (Stadt)"
        Exit Function
    End If
    If InStr(n, "ABGABE") > 0 Then
        ErmittleVersorgerZweck = "Abgaben"
        Exit Function
    End If
    
    ' --- Rundfunk ---
    If InStr(n, "RUNDFUNK") > 0 Or InStr(n, "BEITRAGSSERVICE") > 0 Or InStr(n, "ARD ZDF") > 0 Then
        ErmittleVersorgerZweck = "Rundfunkbeitrag"
        Exit Function
    End If
    
    ' --- Verband ---
    If InStr(n, "VERBAND") > 0 Or InStr(n, "BEZIRKSVERBAND") > 0 Or InStr(n, "LANDESVERBAND") > 0 Then
        ErmittleVersorgerZweck = "Verband/Verb" & ChrW(228) & "nde"
        Exit Function
    End If
    If InStr(n, "VERPACHTUNG") > 0 Or InStr(n, "KLEINGARTENVERBAND") > 0 Then
        ErmittleVersorgerZweck = "Verpachtung/Kleingartenverband"
        Exit Function
    End If
    
    ' --- Miete / Grundstueck ---
    If InStr(n, "GRUNDSTUECKSGESELLSCHAFT") > 0 Or InStr(n, "GRUNDSTUCKSGESELLSCHAFT") > 0 Then
        ErmittleVersorgerZweck = "Grundst" & ChrW(252) & "cks-Miete"
        Exit Function
    End If
    If InStr(n, "HAUSVERWALTUNG") > 0 Then
        ErmittleVersorgerZweck = "Hausverwaltung/Miete"
        Exit Function
    End If
    If InStr(n, "HAUS- UND GRUNDSTUECK") > 0 Or InStr(n, "HAUS UND GRUNDSTUECK") > 0 Then
        ErmittleVersorgerZweck = "Grundst" & ChrW(252) & "cks-Miete"
        Exit Function
    End If
    If InStr(n, "HUG ") > 0 Or InStr(n, "H.U.G") > 0 Then
        ErmittleVersorgerZweck = "Grundst" & ChrW(252) & "cks-Miete"
        Exit Function
    End If
    If InStr(n, "MIETE") > 0 Or InStr(n, "MIETVERTRAG") > 0 Then
        ErmittleVersorgerZweck = "Miete"
        Exit Function
    End If
    If InStr(n, "PACHT") > 0 Then
        ErmittleVersorgerZweck = "Pacht"
        Exit Function
    End If
    
    ' Kein Versorger erkannt
    ErmittleVersorgerZweck = ""
End Function

' ===============================================================
' IstBank - Keyword-Liste
' ===============================================================
Private Function IstBank(ByVal kontoname As String) As Boolean
    Dim n As String
    Dim keywords As Variant
    Dim i As Long
    
    n = UCase(Trim(kontoname))
    IstBank = False
    If Len(n) = 0 Then Exit Function
    
    keywords = Array( _
        "SPARKASSE", "VOLKSBANK", "RAIFFEISENBANK", "RAIFFEISEN", _
        "COMMERZBANK", "DEUTSCHE BANK", "POSTBANK", _
        "ING DIBA", "ING-DIBA", "DKB", "DEUTSCHE KREDITBANK", _
        "COMDIRECT", "CONSORSBANK", "TARGOBANK", _
        "HYPOVEREINSBANK", "UNICREDIT", _
        "BERLINER BANK", "LANDESBANK", "SPARDA", _
        "PSD BANK", "NORISBANK", "N26", _
        "BANK" _
    )
    
    For i = LBound(keywords) To UBound(keywords)
        If InStr(n, CStr(keywords(i))) > 0 Then
            IstBank = True
            Exit Function
        End If
    Next i
End Function

' ===============================================================
' CreateGUID - v5.1: Generiert ID im Format "yyyymmddhhmmss-NNNNN"
' Genau wie der Fallback in CreateGUID_Public auf der Mitgliederliste
' ===============================================================
Private Function CreateGUID() As String
    Randomize
    CreateGUID = Format(Now, "yyyymmddhhmmss") & "-" & Int((99999 - 10000 + 1) * Rnd + 10000)
End Function

' ===============================================================
' Holt alle Parzellen fuer eine MemberID
' Durchsucht die gesamte Mitgliederliste nach allen Zeilen mit
' gleicher MemberID und sammelt die Parzellen kommagetrennt
' ===============================================================
Private Function HoleAlleParzellen(ByVal memberID As String, _
                                    ByRef wsM As Worksheet) As String
    Dim r As Long
    Dim lastRow As Long
    Dim parzelle As String
    Dim result As String
    Dim dictParzellen As Object
    
    Set dictParzellen = CreateObject("Scripting.Dictionary")
    result = ""
    
    If memberID = "" Then
        HoleAlleParzellen = ""
        Exit Function
    End If
    
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_MEMBER_ID).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        If Trim(wsM.Cells(r, M_COL_MEMBER_ID).value) = memberID Then
            parzelle = Trim(wsM.Cells(r, M_COL_PARZELLE).value)
            If parzelle <> "" Then
                If Not dictParzellen.Exists(parzelle) Then
                    dictParzellen.Add parzelle, True
                    If result <> "" Then result = result & ", "
                    result = result & parzelle
                End If
            End If
        End If
    Next r
    
    HoleAlleParzellen = result
End Function

' ===============================================================
' Verarbeitet manuelle Role-Aenderung in Spalte W
' FIX v5.3: Alle manuellen Zuordnungen GRUEN (ausser EHEMALIGES MITGLIED ohne Historie)
' NEU v5.3: EHEMALIGES MITGLIED -> InputBox Parzelle wenn nicht in Historie
' FIX v5.3: Debug-Spalte nur Datum (kein Uhrzeit)
' ===============================================================
Public Sub VerarbeiteManuelleRoleAenderung(ByVal Target As Range)
    Dim wsDaten As Worksheet
    Dim wsM As Worksheet
    Dim wsH As Worksheet
    Dim zeile As Long
    Dim neueRole As String
    Dim kontoname As String
    Dim currentEntityKey As String
    Dim neuerEntityKey As String
    Dim neueZuordnung As String
    Dim neueParzelle As String
    Dim neuerDebug As String
    Dim ampelStatus As Long
    Dim correctPrefix As String
    
    On Error GoTo ErrorHandler
    
    If Target.Column <> EK_COL_ROLE Then Exit Sub
    If Target.Row < EK_START_ROW Then Exit Sub
    
    Set wsDaten = Target.Worksheet
    zeile = Target.Row
    neueRole = UCase(Trim(CStr(Target.value)))
    kontoname = EntferneMehrfacheLeerzeichen(Trim(CStr(wsDaten.Cells(zeile, EK_COL_KONTONAME).value)))
    currentEntityKey = Trim(wsDaten.Cells(zeile, EK_COL_ENTITYKEY).value)
    
    Application.EnableEvents = False
    
    On Error Resume Next
    wsDaten.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    Select Case neueRole
        Case "MITGLIED MIT PACHT", "MITGLIED OHNE PACHT", "MITGLIED"
            Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
            Dim mitglieder As Collection
            Set mitglieder = SucheMitgliederZuKontoname(kontoname, wsM, _
                              ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE))
            
            If mitglieder.count > 0 Then
                Dim bestMatch As Variant
                bestMatch = FindeBestenTreffer(mitglieder)
                
                neuerEntityKey = CStr(bestMatch(0))
                neueZuordnung = bestMatch(1) & ", " & bestMatch(2)
                neueParzelle = HoleAlleParzellen(CStr(bestMatch(0)), wsM)
                neuerDebug = "Manuell: " & neueRole & " -> Mitglied gefunden (" & Format(Now, "dd.mm.yyyy") & ")"
                ampelStatus = 1
            Else
                neuerEntityKey = currentEntityKey
                neueZuordnung = ExtrahiereAnzeigeName(kontoname)
                neueParzelle = ""
                neuerDebug = "Manuell: " & neueRole & " -> KEIN Mitglied gefunden (" & Format(Now, "dd.mm.yyyy") & ")"
                ampelStatus = 2
            End If
            
        Case "EHEMALIGES MITGLIED"
            correctPrefix = PREFIX_EHEMALIG
            If Left(UCase(currentEntityKey), Len(correctPrefix)) <> correctPrefix Then
                neuerEntityKey = correctPrefix & CreateGUID()
            Else
                neuerEntityKey = currentEntityKey
            End If
            neueZuordnung = ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            ampelStatus = 2  ' Default GELB
            
            ' NEU v5.3: Pruefen ob in Mitgliederhistorie
            On Error Resume Next
            Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
            On Error GoTo ErrorHandler
            If Not wsH Is Nothing Then
                If PruefeObInHistorie(kontoname, wsH) Then
                    ' In Historie gefunden -> GRUEN + Parzelle aus Historie
                    ampelStatus = 1
                    Dim historieParzelle As String
                    historieParzelle = HoleParzelleFuerEhemaligesAusHistorie(kontoname, wsH)
                    If historieParzelle <> "" Then
                        neueParzelle = historieParzelle
                    End If
                    neuerDebug = "Manuell: EHEMALIGES MITGLIED - in Historie gefunden; " & Format(Now, "dd.mm.yyyy")
                Else
                    ' NICHT in Historie -> InputBox fuer Parzelle (1-14)
                    ampelStatus = 2
                    
                    Dim eingabe As String
                    Dim parzelleGueltig As Boolean
                    Dim parzelleNr As Long
                    
                    parzelleGueltig = False
                    Do
                        eingabe = InputBox("Welche Parzelle belegte das ehemalige Mitglied?" & vbCrLf & vbCrLf & _
                                           "Bitte eine Zahl von 1 bis 14 eingeben:" & vbCrLf & _
                                           "(Abbrechen = keine Parzelle zuweisen)", _
                                           "Parzelle f" & ChrW(252) & "r ehemaliges Mitglied", "")
                        
                        ' Abbrechen gedrueckt oder leer
                        If eingabe = "" Then
                            Exit Do
                        End If
                        
                        ' Pruefen ob gueltige Zahl 1-14
                        If IsNumeric(eingabe) Then
                            parzelleNr = CLng(eingabe)
                            If parzelleNr >= 1 And parzelleNr <= 14 Then
                                parzelleGueltig = True
                            Else
                                MsgBox "Ung" & ChrW(252) & "ltige Eingabe! Bitte eine Zahl zwischen 1 und 14 eingeben.", vbExclamation, "Ung" & ChrW(252) & "ltige Parzelle"
                            End If
                        Else
                            MsgBox "Ung" & ChrW(252) & "ltige Eingabe! Bitte nur eine Zahl eingeben.", vbExclamation, "Ung" & ChrW(252) & "ltige Eingabe"
                        End If
                    Loop Until parzelleGueltig
                    
                    If parzelleGueltig Then
                        neueParzelle = CStr(parzelleNr)
                        neuerDebug = "Manuell: EHEMALIGES MITGLIED - Parzelle " & neueParzelle & "; nicht in Historie; " & Format(Now, "dd.mm.yyyy")
                    Else
                        neuerDebug = "Manuell: EHEMALIGES MITGLIED; nicht in Historie; " & Format(Now, "dd.mm.yyyy")
                    End If
                End If
            End If
        Case "VERSORGER"
            correctPrefix = PREFIX_VERSORGER
            If Left(UCase(currentEntityKey), Len(correctPrefix)) <> UCase(correctPrefix) Then
                neuerEntityKey = correctPrefix & CreateGUID()
            Else
                neuerEntityKey = currentEntityKey
            End If
            neueZuordnung = ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            neuerDebug = "Manuell: VERSORGER (" & Format(Now, "dd.mm.yyyy") & ")"
            ampelStatus = 1
            
        Case "BANK"
            correctPrefix = PREFIX_BANK
            If Left(UCase(currentEntityKey), Len(correctPrefix)) <> UCase(correctPrefix) Then
                neuerEntityKey = correctPrefix & CreateGUID()
            Else
                neuerEntityKey = currentEntityKey
            End If
            neueZuordnung = ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            neuerDebug = "Manuell: BANK (" & Format(Now, "dd.mm.yyyy") & ")"
            ampelStatus = 1
            
        Case "SHOP"
            correctPrefix = PREFIX_SHOP
            If Left(UCase(currentEntityKey), Len(correctPrefix)) <> UCase(correctPrefix) Then
                neuerEntityKey = correctPrefix & CreateGUID()
            Else
                neuerEntityKey = currentEntityKey
            End If
            neueZuordnung = ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            neuerDebug = "Manuell: SHOP (" & Format(Now, "dd.mm.yyyy") & ")"
            ampelStatus = 1
            
        Case "SONSTIGE"
            correctPrefix = PREFIX_SONSTIGE
            If Left(UCase(currentEntityKey), Len(correctPrefix)) <> UCase(correctPrefix) Then
                neuerEntityKey = correctPrefix & CreateGUID()
            Else
                neuerEntityKey = currentEntityKey
            End If
            neueZuordnung = ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            neuerDebug = "Manuell: SONSTIGE (" & Format(Now, "dd.mm.yyyy") & ")"
            ampelStatus = 1
            
        Case ""
            neuerEntityKey = ""
            neueZuordnung = ""
            neueParzelle = ""
            neuerDebug = ""
            ampelStatus = 3
            
        Case Else
            correctPrefix = PREFIX_SONSTIGE
            If Left(UCase(currentEntityKey), Len(correctPrefix)) <> UCase(correctPrefix) Then
                neuerEntityKey = correctPrefix & CreateGUID()
            Else
                neuerEntityKey = currentEntityKey
            End If
            neueZuordnung = ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            neuerDebug = "Manuell: " & neueRole & " (" & Format(Now, "dd.mm.yyyy") & ")"
            ampelStatus = 2
    End Select
    
    ' EntityKey setzen
    If neuerEntityKey <> "" Then
        wsDaten.Cells(zeile, EK_COL_ENTITYKEY).value = neuerEntityKey
    ElseIf neueRole = "" Then
        wsDaten.Cells(zeile, EK_COL_ENTITYKEY).value = ""
    End If
    
    ' Zuordnung: aus Kontoname wenn leer, Nutzer kann spaeter aendern
    If neueZuordnung <> "" Then
        Dim aktuelleZuordnung As String
        aktuelleZuordnung = Trim(wsDaten.Cells(zeile, EK_COL_ZUORDNUNG).value)
        If aktuelleZuordnung = "" Then
            wsDaten.Cells(zeile, EK_COL_ZUORDNUNG).value = neueZuordnung
        End If
    ElseIf neueRole = "" Then
        wsDaten.Cells(zeile, EK_COL_ZUORDNUNG).value = ""
    End If
    
    ' Parzelle
    If DarfParzelleHaben(neueRole) Then
        If neueParzelle <> "" Then
            wsDaten.Cells(zeile, EK_COL_PARZELLE).value = neueParzelle
        End If
    Else
        ' NEU v5.3: EHEMALIGES MITGLIED darf Parzelle bekommen (aus Historie oder InputBox)
        If neueRole = "EHEMALIGES MITGLIED" And neueParzelle <> "" Then
            wsDaten.Cells(zeile, EK_COL_PARZELLE).value = neueParzelle
        Else
            wsDaten.Cells(zeile, EK_COL_PARZELLE).value = ""
        End If
    End If
    
    ' Debug-Spalte X: IMMER aktualisieren bei jeder manuellen Aenderung
    wsDaten.Cells(zeile, EK_COL_DEBUG).value = neuerDebug
    
    ' Ampelfarbe (vorlaeufig, wird nach Sortierung nochmal gesetzt)
    Call SetzeAmpelFarbe(wsDaten, zeile, ampelStatus)
    
    ' Dropdowns
    Call SetupEntityRoleDropdown(wsDaten, zeile)
    
    If neueRole = "EHEMALIGES MITGLIED" Or neueRole = "SONSTIGE" Then
        Call SetupParzelleDropdown(wsDaten, zeile)
    End If
    
    ' U, W, X immer editierbar
    wsDaten.Cells(zeile, EK_COL_ZUORDNUNG).Locked = False
    wsDaten.Cells(zeile, EK_COL_ROLE).Locked = False
    wsDaten.Cells(zeile, EK_COL_DEBUG).Locked = False
    
    ' NEU v5.3: Sortierung + Ampelfarben sofort nach manueller Aenderung
    Call mod_Formatierung.FormatEntityKeyTableComplete(wsDaten)
    Call SetzeAlleAmpelfarbenNachSortierung(wsDaten)
    
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Application.EnableEvents = True
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    On Error Resume Next
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Debug.Print "FEHLER in VerarbeiteManuelleRoleAenderung: " & Err.Description
End Sub


' ===============================================================
' NEU v5.3: Aktualisiert EntityKey-Tabelle bei Mitglied-Austritt
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
            Call SetzeAmpelFarbe(wsD, r, 1)  ' GRUEN - in Historie vorhanden
            anzahlAktualisiert = anzahlAktualisiert + 1
            
        ElseIf Left(currentEK, Len(PREFIX_SHARE)) = PREFIX_SHARE Then
            ' SHARE-Key: pruefe ob MemberID enthalten
            If InStr(currentEK, alteMemberID) > 0 Then
                ' Gemeinschaftskonto mit diesem Mitglied
                ' MemberID aus SHARE-Key entfernen
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
                    ' Nur noch 1 Person -> direkte MemberID (kein SHARE mehr)
                    wsD.Cells(r, EK_COL_ENTITYKEY).value = newShareParts
                ElseIf verbleibendeAnzahl > 1 Then
                    ' Noch mehrere -> SHARE beibehalten
                    wsD.Cells(r, EK_COL_ENTITYKEY).value = PREFIX_SHARE & newShareParts
                End If
                
                ' Zuordnung und Debug aktualisieren
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

' ===============================================================
' Findet den besten Treffer aus Collection
' ===============================================================
Private Function FindeBestenTreffer(ByRef mitglieder As Collection) As Variant
    Dim bestInfo As Variant
    Dim info As Variant
    Dim bestScore As Long
    Dim i As Long
    
    bestScore = 0
    
    For i = 1 To mitglieder.count
        info = mitglieder(i)
        If CLng(info(8)) > bestScore Then
            bestScore = CLng(info(8))
            bestInfo = info
        End If
    Next i
    
    If bestScore = 0 And mitglieder.count > 0 Then
        bestInfo = mitglieder(1)
    End If
    
    FindeBestenTreffer = bestInfo
End Function

' ===============================================================
' Setzt EntityRole-Dropdown fuer eine Zeile
' ===============================================================
Private Sub SetupEntityRoleDropdown(ByRef ws As Worksheet, ByVal zeile As Long)
    Dim lastRowDD As Long
    
    On Error Resume Next
    
    lastRowDD = ws.Cells(ws.Rows.count, DATA_COL_DD_ENTITYROLE).End(xlUp).Row
    If lastRowDD < DATA_START_ROW Then lastRowDD = DATA_START_ROW
    
    ws.Cells(zeile, EK_COL_ROLE).Validation.Delete
    With ws.Cells(zeile, EK_COL_ROLE).Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:="=" & WS_DATEN & "!$AD$" & DATA_START_ROW & ":$AD$" & lastRowDD
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    
    On Error GoTo 0
End Sub

' ===============================================================
' Setzt Parzellen-Dropdown fuer eine Zeile
' ===============================================================
Private Sub SetupParzelleDropdown(ByRef ws As Worksheet, ByVal zeile As Long)
    Dim lastRowParzelle As Long
    
    On Error Resume Next
    
    lastRowParzelle = ws.Cells(ws.Rows.count, DATA_COL_DD_PARZELLE).End(xlUp).Row
    If lastRowParzelle < DATA_START_ROW Then lastRowParzelle = DATA_START_ROW
    
    ws.Cells(zeile, EK_COL_PARZELLE).Validation.Delete
    With ws.Cells(zeile, EK_COL_PARZELLE).Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:="=" & WS_DATEN & "!$F$" & DATA_START_ROW & ":$F$" & lastRowParzelle
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    
    ws.Cells(zeile, EK_COL_PARZELLE).Locked = False
    
    On Error GoTo 0
End Sub

' ===============================================================
' Kompatibilitaet
' ===============================================================
Public Sub FormatiereEntityKeyZeile(ByVal zeile As Long, Optional ByVal ws As Worksheet = Nothing)
    ' BEWUSST LEER
End Sub



