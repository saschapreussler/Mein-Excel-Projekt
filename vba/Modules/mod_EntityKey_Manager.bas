Attribute VB_Name = "mod_EntityKey_Manager"
Option Explicit

' ***************************************************************
' MODUL: mod_EntityKey_Manager
' ZWECK: Verwaltung und Zuordnung von EntityKeys fuer Bankverkehr
' VERSION: 4.0 - 06.02.2026
' NEU: Ampelfarben, smarte Kontoname-Deduplizierung,
'      EntityRole vereinfacht (kein VORSTAND/EHRENMITGLIED),
'      VerarbeiteManuelleRoleAenderung mit EntityKey-Prefix,
'      bessere Keyword-Listen, HoleAlleParzellen
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
Private Const AMPEL_GRUEN As Long = 12968900   ' RGB(198,224,180) = &HC0E0B4 -> &H00B4E0C6
Private Const AMPEL_GELB As Long = 10086143     ' RGB(255,230,153) = &H99E6FF -> &H0099E6FF
Private Const AMPEL_ROT As Long = 9871103       ' RGB(255,150,150) = &H969696 -> &H009696FF

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
    
    ' IBANs aus Bankkonto sammeln - ALLE wo Spalte A ein Datum hat
    lastRowBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    For r = BK_START_ROW To lastRowBK
        currentDatum = wsBK.Cells(r, BK_COL_DATUM).value
        
        ' Pruefen ob Spalte A ein Datum enthaelt
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
                    ' Weiteren Kontonamen sammeln (nicht-redundant)
                    If currentKontoName <> "" Then
                        Dim nameKey As String
                        nameKey = UCase(Trim(currentKontoName))
                        If Not dictKontonamen(currentIBAN).Exists(nameKey) Then
                            ' Pruefen ob Name semantisch redundant ist
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
                ' Redundante Namen entfernen und besten behalten
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
    
    ' Formatierung nach Import
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
' Erkennt: gleiche Woerter in anderer Reihenfolge,
'          Teilmengen, "Vorname Nachname" vs "Nachname Vorname"
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
        
        ' Fall 1: Gleiche Wortmenge (nur Reihenfolge anders)
        If SindWortmengenGleich(neueWorte, bestehendeWorte) Then
            IstKontonameRedundant = True
            Exit Function
        End If
        
        ' Fall 2: Neuer Name ist Teilmenge des bestehenden
        If IstTeilmenge(neueWorte, bestehendeWorte) Then
            IstKontonameRedundant = True
            Exit Function
        End If
    Next key
End Function

' ===============================================================
' HILFSFUNKTION: Zerlegt einen Namen in normalisierte Worte
' Entfernt Fuellwoerter wie "UND", "U.", "U"
' ===============================================================
Private Function ZerlegeInWorte(ByVal name As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim parts() As String
    Dim wort As String
    Dim i As Long
    
    name = UCase(Trim(name))
    name = Replace(name, ",", " ")
    name = Replace(name, ".", " ")
    name = Replace(name, "-", " ")
    name = EntferneMehrfacheLeerzeichen(name)
    
    If name = "" Then
        Set ZerlegeInWorte = dict
        Exit Function
    End If
    
    parts = Split(name, " ")
    
    For i = LBound(parts) To UBound(parts)
        wort = Trim(parts(i))
        ' Fuellwoerter ignorieren
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
' Behaelt jeweils den laengsten/vollstaendigsten Namen
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
    
    ' Alle Keys/Values in Arrays kopieren
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
    
    ' Paarweise vergleichen
    For i = 0 To cnt - 1
        If Not istRedundant(i) Then
            Set worteI = ZerlegeInWorte(CStr(values(i)))
            For j = i + 1 To cnt - 1
                If Not istRedundant(j) Then
                    Set worteJ = ZerlegeInWorte(CStr(values(j)))
                    
                    ' Gleiche Wortmenge -> kuerzeren entfernen
                    If SindWortmengenGleich(worteI, worteJ) Then
                        If Len(CStr(values(i))) >= Len(CStr(values(j))) Then
                            istRedundant(j) = True
                        Else
                            istRedundant(i) = True
                            Exit For
                        End If
                    ' i ist Teilmenge von j -> i entfernen
                    ElseIf IstTeilmenge(worteI, worteJ) Then
                        istRedundant(i) = True
                        Exit For
                    ' j ist Teilmenge von i -> j entfernen
                    ElseIf IstTeilmenge(worteJ, worteI) Then
                        istRedundant(j) = True
                    End If
                End If
            Next j
        End If
    Next i
    
    ' Nicht-redundante Namen ins Ergebnis
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
' HAUPTPROZEDUR: Aktualisiert alle EntityKeys
' NEU: Setzt Ampelfarben nach Zuordnung
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
            ' Auch bei unveraenderten Zeilen Ampelfarbe setzen
            ampelStatus = ErmittleAmpelStatus(currentEntityKey, currentZuordnung, currentRole)
            Call SetzeAmpelFarbe(wsD, r, ampelStatus)
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
        
        ' Ampelfarbe setzen
        Call SetzeAmpelFarbe(wsD, r, ampelStatus)
        
        If ampelStatus = 3 Then zeilenProbleme = zeilenProbleme + 1
        
nextRow:
    Next r
    
    ' Formatierung nach EntityKey-Aktualisierung
    Call mod_Formatierung.FormatEntityKeyTableComplete(wsD)
    
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
' HILFSFUNKTION: Ermittelt Ampelstatus fuer bestehende Zeilen
' ===============================================================
Private Function ErmittleAmpelStatus(ByVal entityKey As String, _
                                      ByVal zuordnung As String, _
                                      ByVal role As String) As Long
    ' Wenn alles da ist -> Gruen
    If entityKey <> "" And zuordnung <> "" And role <> "" Then
        ErmittleAmpelStatus = 1
        Exit Function
    End If
    
    ' Wenn Role fehlt oder EntityKey fehlt -> Gelb
    If entityKey = "" Or role = "" Then
        ErmittleAmpelStatus = 2
        Exit Function
    End If
    
    ' Sonst Gruen
    ErmittleAmpelStatus = 1
End Function

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
        For Each key In uniqueMemberIDs.keys
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
' NEU: Kein VORSTAND/EHRENMITGLIED mehr - alles MITGLIED MIT PACHT
' ===============================================================
Private Function ErmittleEntityRoleVonFunktion(ByVal funktion As String) As String
    Dim funktionUpper As String
    funktionUpper = UCase(funktion)
    
    If InStr(funktionUpper, "OHNE PACHT") > 0 Then
        ErmittleEntityRoleVonFunktion = ROLE_MITGLIED_OHNE_PACHT
    ElseIf InStr(funktionUpper, "EHEMALIG") > 0 Then
        ErmittleEntityRoleVonFunktion = ROLE_EHEMALIGES_MITGLIED
    Else
        ' VORSTAND, EHRENMITGLIED, KASSIERER, SCHRIFTFUEHRER
        ' und alles andere -> MITGLIED MIT PACHT
        ErmittleEntityRoleVonFunktion = ROLE_MITGLIED_MIT_PACHT
    End If
End Function

' ===============================================================
' IstShop - erweiterte Keyword-Liste
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
        "ACTION", "TEDi", "WOOLWORTH", "KIK", _
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
' IstVersorger - erweiterte Keyword-Liste
' ===============================================================
Private Function IstVersorger(ByVal kontoname As String) As Boolean
    Dim n As String
    Dim keywords As Variant
    Dim i As Long
    
    n = UCase(Trim(kontoname))
    IstVersorger = False
    If Len(n) = 0 Then Exit Function
    
    keywords = Array( _
        "STADTWERK", "ENERGIE", "STROM", "VATTENFALL", "E.ON", "EON", _
        "RWE", "ENVIA", "ENVIAM", "ENBW", "MAINOVA", "ENTEGA", _
        "GASAG", "GAS", "FERNWAERME", "HEIZUNG", _
        "WASSER", "ABWASSER", "BWB", "BERLINER WASSERBETRIEBE", _
        "VERSICHERUNG", "ALLIANZ", "DEVK", "HUK", "HDI", "ERGO", _
        "GENERALI", "AXA", "ZURICH", "WUERTTEMBERGISCHE", _
        "TELEKOM", "VODAFONE", "1&1", "O2", "TELEFONICA", _
        "KABEL DEUTSCHLAND", "UNITYMEDIA", _
        "BSR", "ENTSORGUNG", "STADTREINIGUNG", "ABFALLWIRTSCHAFT", _
        "RUNDFUNK", "BEITRAGSSERVICE", "ARD ZDF", _
        "GRUNDSTEUER", "FINANZAMT", "ABGABE", _
        "VERBAND", "BEZIRKSVERBAND", "LANDESVERBAND", _
        "VERPACHTUNG", "KLEINGARTENVERBAND" _
    )
    
    For i = LBound(keywords) To UBound(keywords)
        If InStr(n, CStr(keywords(i))) > 0 Then
            IstVersorger = True
            Exit Function
        End If
    Next i
End Function

' ===============================================================
' IstBank - erweiterte Keyword-Liste
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
' HILFSFUNKTION: Holt alle Parzellen fuer eine MemberID
' aus Mitgliederliste (aktuelle) und Mitgliederhistorie (ehemalige)
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
    
    ' Aktive Mitglieder durchsuchen
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
' OEFFENTLICH: Verarbeitet manuelle Role-Aenderung in Spalte W
' NEU: Generiert passenden EntityKey mit Prefix bei Rolle-Wechsel
'      Sucht Mitglieder wenn Role=MITGLIED/EHEMALIGES MITGLIED
' ===============================================================
Public Sub VerarbeiteManuelleRoleAenderung(ByVal Target As Range)
    Dim wsDaten As Worksheet
    Dim wsM As Worksheet
    Dim zeile As Long
    Dim neueRole As String
    Dim kontoname As String
    Dim currentEntityKey As String
    Dim neuerEntityKey As String
    Dim neueZuordnung As String
    Dim neueParzelle As String
    Dim neuerDebug As String
    Dim ampelStatus As Long
    
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
            ' Mitglied suchen
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
                neuerDebug = "Manuell: " & neueRole & " -> Mitglied gefunden (" & Format(Now, "dd.mm.yyyy hh:mm") & ")"
                ampelStatus = 1
            Else
                neuerEntityKey = currentEntityKey
                neueZuordnung = ""
                neueParzelle = ""
                neuerDebug = "Manuell: " & neueRole & " -> KEIN Mitglied gefunden (" & Format(Now, "dd.mm.yyyy hh:mm") & ")"
                ampelStatus = 2
            End If
            
        Case "EHEMALIGES MITGLIED"
            neuerEntityKey = PREFIX_EHEMALIG & CreateGUID()
            neueZuordnung = ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""  ' Wird per Dropdown gesetzt
            neuerDebug = "Manuell: EHEMALIGES MITGLIED (" & Format(Now, "dd.mm.yyyy hh:mm") & ")"
            ampelStatus = 2
            
        Case "VERSORGER"
            neuerEntityKey = PREFIX_VERSORGER & CreateGUID()
            neueZuordnung = ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            neuerDebug = "Manuell: VERSORGER (" & Format(Now, "dd.mm.yyyy hh:mm") & ")"
            ampelStatus = 1
            
        Case "BANK"
            neuerEntityKey = PREFIX_BANK & CreateGUID()
            neueZuordnung = ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            neuerDebug = "Manuell: BANK (" & Format(Now, "dd.mm.yyyy hh:mm") & ")"
            ampelStatus = 1
            
        Case "SHOP"
            neuerEntityKey = PREFIX_SHOP & CreateGUID()
            neueZuordnung = ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            neuerDebug = "Manuell: SHOP (" & Format(Now, "dd.mm.yyyy hh:mm") & ")"
            ampelStatus = 1
            
        Case "SONSTIGE"
            neuerEntityKey = PREFIX_SONSTIGE & CreateGUID()
            neueZuordnung = ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            neuerDebug = "Manuell: SONSTIGE (" & Format(Now, "dd.mm.yyyy hh:mm") & ")"
            ampelStatus = 1
            
        Case ""
            ' Leere Role - nichts aendern
            neuerEntityKey = ""
            neueZuordnung = ""
            neueParzelle = ""
            neuerDebug = ""
            ampelStatus = 3
            
        Case Else
            neuerEntityKey = PREFIX_SONSTIGE & CreateGUID()
            neueZuordnung = ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            neuerDebug = "Manuell: " & neueRole & " (" & Format(Now, "dd.mm.yyyy hh:mm") & ")"
            ampelStatus = 2
    End Select
    
    ' Werte setzen (nur wenn noch leer oder Role geaendert)
    If neuerEntityKey <> "" Then
        wsDaten.Cells(zeile, EK_COL_ENTITYKEY).value = neuerEntityKey
    End If
    
    If neueZuordnung <> "" Then
        wsDaten.Cells(zeile, EK_COL_ZUORDNUNG).value = neueZuordnung
    End If
    
    ' Parzelle: setzen oder loeschen je nach Role
    If DarfParzelleHaben(neueRole) Then
        If neueParzelle <> "" Then
            wsDaten.Cells(zeile, EK_COL_PARZELLE).value = neueParzelle
        End If
    Else
        wsDaten.Cells(zeile, EK_COL_PARZELLE).value = ""
    End If
    
    If neuerDebug <> "" Then
        wsDaten.Cells(zeile, EK_COL_DEBUG).value = neuerDebug
    End If
    
    ' Ampelfarbe setzen
    Call SetzeAmpelFarbe(wsDaten, zeile, ampelStatus)
    
    ' EntityRole-Dropdown fuer diese Zeile setzen
    Call SetupEntityRoleDropdown(wsDaten, zeile)
    
    ' Parzellen-Dropdown fuer EHEMALIGES MITGLIED
    If neueRole = "EHEMALIGES MITGLIED" Then
        Call SetupParzelleDropdown(wsDaten, zeile)
    End If
    
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
' HILFSFUNKTION: Findet den besten Treffer aus Collection
' Priorisiert exakte Treffer (matchResult=2) vor Nachname-only (1)
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
    
    ' Falls kein Treffer, ersten nehmen
    If bestScore = 0 And mitglieder.count > 0 Then
        bestInfo = mitglieder(1)
    End If
    
    FindeBestenTreffer = bestInfo
End Function

' ===============================================================
' HILFSPROZEDUR: Setzt EntityRole-Dropdown fuer eine Zeile
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
' HILFSPROZEDUR: Setzt Parzellen-Dropdown fuer eine Zeile
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
' OEFFENTLICH: Formatiert eine einzelne Zeile (Kompatibilitaet)
' ===============================================================
Public Sub FormatiereEntityKeyZeile(ByVal zeile As Long, Optional ByVal ws As Worksheet = Nothing)
    ' BEWUSST LEER - Formatierung wird durch mod_Formatierung gesteuert
End Sub

