Attribute VB_Name = "mod_EntityKey_Manager"
Option Explicit

' ***************************************************************
' MODUL: mod_EntityKey_Manager
' ZWECK: Verwaltung und Zuordnung von EntityKeys für Bankverkehr
' VERSION: 3.0 - 04.02.2026
' ÄNDERUNGEN:
'   - Sortierung: Parzelle 1-14, dann VERSORGER, BANK, SHOP
'   - Gemeinschaftskonto: Alle Namen aus Kontoname suchen
'   - EntityRole ohne Unterstriche (Leerzeichen statt _)
'   - IstVersorger/IstShop komplett überarbeitet
'   - Doppelte Leerzeichen in Kontoname entfernen
'   - Dynamische Role-Dropdown
' ***************************************************************

' ===============================================================
' KONSTANTEN
' ===============================================================
Private Const EK_ROLE_DROPDOWN_COL As Long = 30
Private Const EK_PARZELLE_START_ROW As Long = 4
Private Const EK_PARZELLE_END_ROW As Long = 17
Private Const EK_PARZELLE_COL As Long = 6

Private Const ZEBRA_COLOR As Long = &HDEE5E3

' EntityKey Präfixe
Private Const PREFIX_SHARE As String = "SHARE-"
Private Const PREFIX_VERSORGER As String = "VERS-"
Private Const PREFIX_BANK As String = "BANK-"
Private Const PREFIX_SHOP As String = "SHOP-"
Private Const PREFIX_EHEMALIG As String = "EX-"
Private Const PREFIX_SONSTIGE As String = "SONST-"

' EntityRole Werte - OHNE UNTERSTRICHE (mit Leerzeichen)
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
' Erlaubt für: Mitglieder, Vorstand, Ehrenmitglieder, SONSTIGE
' Nicht erlaubt für: VERSORGER, BANK, SHOP oder LEER
' ===============================================================
Private Function DarfParzelleHaben(ByVal role As String) As Boolean
    Dim normRole As String
    
    If Trim(role) = "" Then
        DarfParzelleHaben = False
        Exit Function
    End If
    
    normRole = UCase(Trim(role))
    
    ' Erlaubt für Mitglieder-Typen und SONSTIGE
    If InStr(normRole, "MITGLIED") > 0 Then
        DarfParzelleHaben = True
    ElseIf InStr(normRole, "VORSTAND") > 0 Then
        DarfParzelleHaben = True
    ElseIf InStr(normRole, "EHRENMITGLIED") > 0 Then
        DarfParzelleHaben = True
    ElseIf normRole = "SONSTIGE" Then
        DarfParzelleHaben = True
    Else
        ' VERSORGER, BANK, SHOP = keine Parzelle
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
' ÖFFENTLICHE PROZEDUR: Importiert IBANs aus Bankkonto
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
    
    lastRowBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    For r = BK_START_ROW To lastRowBK
        currentDatum = wsBK.Cells(r, BK_COL_DATUM).value
        
        If Not IsEmpty(currentDatum) And currentDatum <> "" Then
            currentIBAN = NormalisiereIBAN(wsBK.Cells(r, BK_COL_IBAN).value)
            currentKontoName = EntferneMehrfacheLeerzeichen(Trim(wsBK.Cells(r, BK_COL_NAME).value))
            
            If currentIBAN <> "" And currentIBAN <> "N.A." And Len(currentIBAN) >= 15 Then
                If Not dictIBANs.Exists(currentIBAN) Then
                    dictIBANs.Add currentIBAN, currentKontoName
                Else
                    If currentKontoName <> "" Then
                        existingKontoName = dictIBANs(currentIBAN)
                        If InStr(existingKontoName, currentKontoName) = 0 Then
                            dictIBANs(currentIBAN) = existingKontoName & vbLf & currentKontoName
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
    
    For Each ibanKey In dictIBANs.Keys
        currentIBAN = CStr(ibanKey)
        currentKontoName = EntferneMehrfacheLeerzeichen(dictIBANs(ibanKey))
        
        If dictExisting.Exists(currentIBAN) Then
            r = dictExisting(currentIBAN)
            existingKontoName = Trim(wsD.Cells(r, EK_COL_KONTONAME).value)
            
            Dim neueNamen As String
            neueNamen = MergeKontonamen(existingKontoName, currentKontoName)
            
            If neueNamen <> existingKontoName Then
                wsD.Cells(r, EK_COL_KONTONAME).value = neueNamen
                anzahlAktualisiert = anzahlAktualisiert + 1
            End If
        Else
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
    
    Debug.Print "IBAN-Import: " & anzahlNeu & " neue, " & anzahlAktualisiert & " aktualisiert"
    
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
' HILFSFUNKTION: Führt Kontonamen zusammen (ohne Duplikate)
' ===============================================================
Private Function MergeKontonamen(ByVal existing As String, ByVal neu As String) As String
    Dim result As String
    Dim existingNames() As String
    Dim newNames() As String
    Dim dictNames As Object
    Dim i As Long
    Dim Name As String
    
    Set dictNames = CreateObject("Scripting.Dictionary")
    result = ""
    
    If existing <> "" Then
        existingNames = Split(existing, vbLf)
        For i = LBound(existingNames) To UBound(existingNames)
            Name = EntferneMehrfacheLeerzeichen(Trim(existingNames(i)))
            If Name <> "" And Not dictNames.Exists(Name) Then
                dictNames.Add Name, True
                If result <> "" Then result = result & vbLf
                result = result & Name
            End If
        Next i
    End If
    
    If neu <> "" Then
        newNames = Split(neu, vbLf)
        For i = LBound(newNames) To UBound(newNames)
            Name = EntferneMehrfacheLeerzeichen(Trim(newNames(i)))
            If Name <> "" And Not dictNames.Exists(Name) Then
                dictNames.Add Name, True
                If result <> "" Then result = result & vbLf
                result = result & Name
            End If
        Next i
    End If
    
    MergeKontonamen = result
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
' HAUPTPROZEDUR: Aktualisiert alle EntityKeys in der Tabelle
' WICHTIG: Bereits manuell zugeordnete Zeilen werden NICHT überschrieben
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
    
    lastRow = wsD.Cells(wsD.Rows.count, EK_COL_IBAN).End(xlUp).Row
    If lastRow < EK_START_ROW Then lastRow = EK_START_ROW
    
    Call SetupEntityRoleDropdown(wsD, lastRow)
    
    For r = EK_START_ROW To lastRow
        iban = Trim(wsD.Cells(r, EK_COL_IBAN).value)
        kontoname = EntferneMehrfacheLeerzeichen(Trim(wsD.Cells(r, EK_COL_KONTONAME).value))
        
        ' Doppelte Leerzeichen in Spalte T bereinigen
        wsD.Cells(r, EK_COL_KONTONAME).value = kontoname
        
        currentEntityKey = Trim(wsD.Cells(r, EK_COL_ENTITYKEY).value)
        currentZuordnung = Trim(wsD.Cells(r, EK_COL_ZUORDNUNG).value)
        currentParzelle = Trim(wsD.Cells(r, EK_COL_PARZELLE).value)
        currentRole = Trim(wsD.Cells(r, EK_COL_ROLE).value)
        currentDebug = Trim(wsD.Cells(r, EK_COL_DEBUG).value)
        
        If iban = "" And kontoname = "" Then GoTo NextRow
        
        ' WICHTIG: Bereits manuell zugeordnete Zeilen NICHT überschreiben
        If HatBereitsGueltigeDaten(currentEntityKey, currentZuordnung, currentRole) Then
            zeilenUnveraendert = zeilenUnveraendert + 1
            If currentRole <> "" Then
                Call SetzeAmpelFarbe(wsD, r, 1)
            End If
            Call SetupParzellenDropdown(wsD, r, currentRole)
            Call SetzeZellschutzFuerZeile(wsD, r, currentRole)
            GoTo NextRow
        End If
        
        zeilenNeu = zeilenNeu + 1
        
        ' Suche ALLE Mitglieder im Kontonamen (für Gemeinschaftskonten)
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
        Call SetupParzellenDropdown(wsD, r, entityRole)
        
        If ampelStatus = 3 Then
            zeilenRot.Add r
        ElseIf ampelStatus = 2 Then
            zeilenGelb.Add r
        End If
        
NextRow:
    Next r
    
    Call FormatiereEntityKeyTabelle(wsD, lastRow)
    Call SortiereEntityKeyTabelle(wsD)
    
    wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    If zeilenRot.count > 0 Or zeilenGelb.count > 0 Then
        Call ZeigeEingriffsHinweis(wsD, zeilenRot, zeilenGelb, zeilenNeu, zeilenUnveraendert)
    Else
        MsgBox "EntityKey-Aktualisierung abgeschlossen!" & vbCrLf & vbCrLf & _
               "Neue Zeilen verarbeitet: " & zeilenNeu & vbCrLf & _
               "Bestehende Zeilen unverändert: " & zeilenUnveraendert & vbCrLf & vbCrLf & _
               "Alle Zuordnungen sind vollständig (GRÜN).", vbInformation, "Aktualisierung abgeschlossen"
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
' SORTIERUNG: Parzelle 1-14, dann VERSORGER, dann BANK, dann SHOP
' ===============================================================
Public Sub SortiereEntityKeyTabelle(Optional ByRef ws As Worksheet = Nothing)
    
    Dim lastRow As Long
    Dim r As Long
    Dim sortKey As String
    Dim role As String
    Dim parzelle As String
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    End If
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.count, EK_COL_IBAN).End(xlUp).Row
    If lastRow < EK_START_ROW Then
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        Exit Sub
    End If
    
    ' Hilfsspalte für Sortierung erstellen (Spalte AE = 31)
    Const SORT_HELPER_COL As Long = 31
    
    For r = EK_START_ROW To lastRow
        parzelle = Trim(ws.Cells(r, EK_COL_PARZELLE).value)
        role = UCase(Trim(ws.Cells(r, EK_COL_ROLE).value))
        
        ' Sortierreihenfolge: 1=Parzellen, 2=VERSORGER, 3=BANK, 4=SHOP, 5=Rest
        If parzelle <> "" And IsNumeric(Left(parzelle, 2)) Then
            ' Parzelle 1-14: Sortierkey "1" + zweistellige Parzellennummer
            sortKey = "1" & Format(Val(parzelle), "00")
        ElseIf InStr(role, "VERSORGER") > 0 Then
            sortKey = "2000"
        ElseIf InStr(role, "BANK") > 0 Then
            sortKey = "3000"
        ElseIf InStr(role, "SHOP") > 0 Then
            sortKey = "4000"
        Else
            sortKey = "5000"
        End If
        
        ws.Cells(r, SORT_HELPER_COL).value = sortKey
    Next r
    
    ' Sortieren nach Hilfsspalte
    Dim sortRange As Range
    Set sortRange = ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), ws.Cells(lastRow, SORT_HELPER_COL))
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(EK_START_ROW, SORT_HELPER_COL), ws.Cells(lastRow, SORT_HELPER_COL)), _
                           SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    With ws.Sort
        .SetRange sortRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Hilfsspalte löschen
    ws.Columns(SORT_HELPER_COL).ClearContents
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
End Sub



'--- Ende Teil 1 von 3 ---
'--- Anfang Teil 2 von 3 ---



' ===============================================================
' HILFSPROZEDUR: Setzt Parzellen-Dropdown
' NUR wenn Role gefüllt UND Parzelle erlaubt ist
' ===============================================================
Private Sub SetupParzellenDropdown(ByRef ws As Worksheet, ByVal zeile As Long, ByVal role As String)
    
    Dim cell As Range
    Dim dropdownSource As String
    
    Set cell = ws.Cells(zeile, EK_COL_PARZELLE)
    
    On Error Resume Next
    cell.Validation.Delete
    On Error GoTo 0
    
    ' Dropdown NUR wenn Role gefüllt UND für diesen Typ erlaubt
    If DarfParzelleHaben(role) Then
        dropdownSource = "=$F$" & EK_PARZELLE_START_ROW & ":$F$" & EK_PARZELLE_END_ROW
        
        On Error Resume Next
        With cell.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:=dropdownSource
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = True
            .ErrorTitle = "Ungültige Eingabe"
            .ErrorMessage = "Bitte wählen Sie eine Parzelle aus der Liste."
        End With
        On Error GoTo 0
        
        cell.Locked = False
    Else
        ' Keine Dropdown, Zelle sperren, Inhalt löschen
        cell.Locked = True
        cell.value = ""
    End If
    
End Sub

' ===============================================================
' HILFSPROZEDUR: Setzt Dropdown für EntityRole (DYNAMISCH aus AD)
' Der Nutzer kann die Liste in Spalte AD jederzeit erweitern
' ===============================================================
Private Sub SetupEntityRoleDropdown(ByRef ws As Worksheet, ByVal lastRow As Long)
    
    Dim rngDropdown As Range
    Dim dropdownSource As String
    Dim lastRoleRow As Long
    Dim dropdownEndRow As Long
    
    ' Dynamisch: Letzte gefüllte Zeile in Spalte AD ermitteln
    lastRoleRow = ws.Cells(ws.Rows.count, EK_ROLE_DROPDOWN_COL).End(xlUp).Row
    If lastRoleRow < 4 Then lastRoleRow = 12 ' Mindestens 12 Zeilen für Standardrollen
    
    ' Dynamischer Bereich - Nutzer kann erweitern
    dropdownSource = "=$AD$4:$AD$" & lastRoleRow
    
    dropdownEndRow = lastRow + 100 ' Genug Reserve für neue Einträge
    
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
        .ErrorMessage = "Bitte wählen Sie einen Wert aus der Liste oder erweitern Sie die Liste in Spalte AD."
    End With
    On Error GoTo 0
    
End Sub

' ===============================================================
' HILFSPROZEDUR: Setzt Zellschutz basierend auf Role-Typ
' Spalten U, W, X sind IMMER bearbeitbar
' Spalte V nur wenn Role gefüllt UND erlaubt
' ===============================================================
Public Sub SetzeZellschutzFuerZeile(ByRef ws As Worksheet, ByVal zeile As Long, ByVal role As String)
    
    ' Spalten U (Zuordnung), W (Role), X (Debug) IMMER bearbeitbar
    ws.Cells(zeile, EK_COL_ZUORDNUNG).Locked = False
    ws.Cells(zeile, EK_COL_ROLE).Locked = False
    ws.Cells(zeile, EK_COL_DEBUG).Locked = False
    
    ' Spalte V (Parzelle) nur bearbeitbar wenn Role gefüllt UND erlaubt
    If DarfParzelleHaben(role) Then
        ws.Cells(zeile, EK_COL_PARZELLE).Locked = False
    Else
        ws.Cells(zeile, EK_COL_PARZELLE).Locked = True
    End If
    
End Sub

' ===============================================================
' HILFSFUNKTION: Prüft ob Zeile bereits gültige Daten hat
' Wenn ja, wird diese Zeile NICHT automatisch überschrieben
' ===============================================================
Private Function HatBereitsGueltigeDaten(ByVal entityKey As String, _
                                          ByVal zuordnung As String, _
                                          ByVal role As String) As Boolean
    
    HatBereitsGueltigeDaten = False
    
    ' Wenn EntityKey vorhanden und kein reiner Zahlenstring
    If entityKey <> "" Then
        If Not IsNumeric(entityKey) Then
            HatBereitsGueltigeDaten = True
            Exit Function
        End If
    End If
    
    ' Wenn sowohl Zuordnung als auch Role vorhanden
    If zuordnung <> "" And role <> "" Then
        HatBereitsGueltigeDaten = True
        Exit Function
    End If
    
End Function

' ===============================================================
' HILFSPROZEDUR: Zeigt Hinweis für Zeilen mit erforderlichem Eingriff
' ===============================================================
Private Sub ZeigeEingriffsHinweis(ByRef ws As Worksheet, ByRef zeilenRot As Collection, _
                                   ByRef zeilenGelb As Collection, _
                                   ByVal zeilenNeu As Long, ByVal zeilenUnveraendert As Long)
    
    Dim msg As String
    Dim antwort As VbMsgBoxResult
    Dim ersteZeile As Long
    
    msg = "EntityKey-Aktualisierung abgeschlossen!" & vbCrLf & vbCrLf
    msg = msg & "Neue Zeilen verarbeitet: " & zeilenNeu & vbCrLf
    msg = msg & "Bestehende Zeilen unverändert: " & zeilenUnveraendert & vbCrLf & vbCrLf
    
    If zeilenRot.count > 0 Then
        msg = msg & "ROT: " & zeilenRot.count & " Zeile(n) - Manuelle Zuordnung erforderlich!" & vbCrLf
    End If
    
    If zeilenGelb.count > 0 Then
        msg = msg & "GELB: " & zeilenGelb.count & " Zeile(n) - Nur Nachname gefunden, bitte prüfen!" & vbCrLf
    End If
    
    msg = msg & vbCrLf & "Möchten Sie jetzt zur ersten betroffenen Zeile springen?"
    
    antwort = MsgBox(msg, vbYesNo + vbExclamation, "Zuordnung prüfen")
    
    If antwort = vbYes Then
        If zeilenRot.count > 0 Then
            ersteZeile = zeilenRot(1)
        ElseIf zeilenGelb.count > 0 Then
            ersteZeile = zeilenGelb(1)
        Else
            Exit Sub
        End If
        
        ws.Activate
        ws.Cells(ersteZeile, EK_COL_ZUORDNUNG).Select
    End If
    
End Sub

' ===============================================================
' HILFSFUNKTION: Sucht ALLE Mitglieder im Kontonamen
' WICHTIG: Durchsucht den GESAMTEN String nach allen Mitgliedernamen
' Für Gemeinschaftskonten wie "Juliane Kelm, Julian Peter"
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
    Dim nameKombiniert As String
    Dim nameParts() As String
    Dim austrittsDatum As Date
    Dim matchResult As Long
    
    Set SucheMitgliederZuKontoname = gefunden
    
    If kontoname = "" Then Exit Function
    
    ' Gesamten Kontonamen normalisieren (NICHT in Zeilen splitten!)
    kontoNameNorm = NormalisiereStringFuerVergleich(kontoname)
    If kontoNameNorm = "" Then Exit Function
    
    ' === AKTIVE MITGLIEDER DURCHSUCHEN ===
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        If Trim(wsM.Cells(r, M_COL_PACHTENDE).value) = "" Then
            nachname = Trim(wsM.Cells(r, M_COL_NACHNAME).value)
            vorname = Trim(wsM.Cells(r, M_COL_VORNAME).value)
            memberID = Trim(wsM.Cells(r, M_COL_MEMBER_ID).value)
            parzelle = Trim(wsM.Cells(r, M_COL_PARZELLE).value)
            funktion = Trim(wsM.Cells(r, M_COL_FUNKTION).value)
            
            ' Prüfe ob dieser Name im Kontonamen vorkommt
            matchResult = PruefeNamensMatch(nachname, vorname, kontoNameNorm)
            
            If matchResult > 0 Then
                If Not IstMitgliedBereitsGefunden(gefunden, memberID, False) Then
                    mitgliedInfo(0) = memberID
                    mitgliedInfo(1) = nachname
                    mitgliedInfo(2) = vorname
                    mitgliedInfo(3) = parzelle
                    mitgliedInfo(4) = funktion
                    mitgliedInfo(5) = r
                    mitgliedInfo(6) = False ' nicht ehemalig
                    mitgliedInfo(7) = CDate("01.01.1900")
                    mitgliedInfo(8) = matchResult
                    gefunden.Add mitgliedInfo
                End If
            End If
        End If
    Next r
    
    ' === EHEMALIGE MITGLIEDER DURCHSUCHEN ===
    lastRow = wsH.Cells(wsH.Rows.count, H_COL_NAME_EHEM_PAECHTER).End(xlUp).Row
    
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
                mitgliedInfo(6) = True ' ehemalig
                mitgliedInfo(7) = austrittsDatum
                mitgliedInfo(8) = matchResult
                gefunden.Add mitgliedInfo
            End If
        End If
    Next r
    
    Set SucheMitgliederZuKontoname = gefunden
    
End Function

' ===============================================================
' HILFSFUNKTION: Prüft ob Name im Kontonamen enthalten ist
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
    
    nachnameGefunden = (InStr(kontoNameNorm, nachnameNorm) > 0)
    
    If Not nachnameGefunden Then
        PruefeNamensMatch = 0
        Exit Function
    End If
    
    vornameGefunden = False
    
    If vornameNorm <> "" And Len(vornameNorm) >= 2 Then
        vornameGefunden = (InStr(kontoNameNorm, vornameNorm) > 0)
    End If
    
    If nachnameGefunden And vornameGefunden Then
        PruefeNamensMatch = 2 ' Exakter Treffer
    ElseIf nachnameGefunden Then
        PruefeNamensMatch = 1 ' Nur Nachname
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
    result = Trim(result)
    
    NormalisiereStringFuerVergleich = result
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
        If item(0) = memberID And item(6) = istEhemalig Then
            IstMitgliedBereitsGefunden = True
            Exit Function
        End If
    Next item
End Function

' ===============================================================
' HILFSPROZEDUR: Generiert EntityKey und Zuordnung
' Bei Gemeinschaftskonten: Alle Namen mit "Nachname, Vorname" + vbLf
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
    
    For i = 1 To mitglieder.count
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
    If mitgliederExakt.count = 0 Then
        ' WICHTIG: Erst SHOP prüfen (hat Vorrang vor VERSORGER)
        If IstShop(kontoname) Then
            outEntityKey = PREFIX_SHOP & CreateGUID()
            outEntityRole = ROLE_SHOP
            outZuordnung = ExtrahiereAnzeigeName(kontoname)
            outParzellen = ""
            outDebugInfo = "Automatisch als SHOP erkannt"
            outAmpelStatus = 1
            Exit Sub
        ElseIf IstVersorger(kontoname) Then
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
        End If
        
        If mitgliederNurNachname.count > 0 Then
            outEntityKey = ""
            outZuordnung = ""
            outParzellen = ""
            outEntityRole = ""
            outDebugInfo = "NUR NACHNAME GEFUNDEN - Bitte prüfen! Mögliche Treffer:"
            outAmpelStatus = 2
            
            For i = 1 To mitgliederNurNachname.count
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
    
    For i = 1 To mitgliederExakt.count
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
        For i = 1 To mitgliederExakt.count
            mitgliedInfo = mitgliederExakt(i)
            If Not uniqueMemberIDs.Exists(CStr(mitgliedInfo(0))) Then
                uniqueMemberIDs.Add CStr(mitgliedInfo(0)), CStr(mitgliedInfo(0))
            End If
        Next i
        
        If uniqueMemberIDs.count = 1 Then
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
            outDebugInfo = "Ehem. Gemeinschaftskonto - " & uniqueMemberIDs.count & " Personen"
            outAmpelStatus = 1
            
            Dim bereitsHinzu As Object
            Set bereitsHinzu = CreateObject("Scripting.Dictionary")
            
            For i = 1 To mitgliederExakt.count
                mitgliedInfo = mitgliederExakt(i)
                If Not bereitsHinzu.Exists(CStr(mitgliedInfo(0))) Then
                    bereitsHinzu.Add CStr(mitgliedInfo(0)), True
                    ' Format: "Nachname, Vorname" mit vbLf
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
    
    ' Fall 2b: Aktive Mitglieder - genau 1
    If uniqueMemberIDs.count = 1 Then
        For i = 1 To mitgliederExakt.count
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
            outDebugInfo = outDebugInfo & " (+ ehem. Einträge in Historie)"
        End If
        
        Exit Sub
    End If
    
    ' Fall 2c: Mehrere aktive Mitglieder = GEMEINSCHAFTSKONTO
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
        
        Dim bereitsHinzugefuegteMitglieder As Object
        Set bereitsHinzugefuegteMitglieder = CreateObject("Scripting.Dictionary")
        
        ' WICHTIG: Alle gefundenen Mitglieder mit "Nachname, Vorname" + vbLf ausgeben
        For i = 1 To mitgliederExakt.count
            mitgliedInfo = mitgliederExakt(i)
            If mitgliedInfo(6) = False Then
                If Not bereitsHinzugefuegteMitglieder.Exists(CStr(mitgliedInfo(0))) Then
                    bereitsHinzugefuegteMitglieder.Add CStr(mitgliedInfo(0)), True
                    
                    ' Format: "Nachname, Vorname" mit vbLf zwischen den Namen
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
    
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_MEMBER_ID).End(xlUp).Row
    
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
' HILFSFUNKTION: Ermittelt EntityRole aus Funktion
' OHNE UNTERSTRICHE - mit Leerzeichen
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



'--- Ende Teil 2 von 3 ---
'--- Anfang Teil 3 von 3 ---



' =====================================================================
' TEIL 3 VON 3 - KLASSIFIZIERUNG UND FORMATIERUNG
' =====================================================================

' ---------------------------------------------------------------------
' IstVersorger - Prüft ob Kontoname ein ECHTER Versorger ist
' WICHTIG: Nur Strom, Gas, Wasser, Versicherungen - KEINE Shops!
' ---------------------------------------------------------------------
Private Function IstVersorger(ByVal kontoname As String) As Boolean
    Dim nameUpper As String
    nameUpper = UCase(Trim(kontoname))
    
    If Len(nameUpper) = 0 Then
        IstVersorger = False
        Exit Function
    End If
    
    ' Energieversorger (Strom, Gas)
    If InStr(1, nameUpper, "STADTWERK", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "ENERGIE", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "STROM", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "GAS", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "VATTENFALL", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "E.ON", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "EON", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "ENVIA", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "RWE", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "ENBW", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "GASAG", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "ENTEGA", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "MAINOVA", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "PFALZWERKE", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "SYNA", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "LICHTBLICK", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "YELLO", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "NATURSTROM", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "GREENPEACE ENERGY", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "POLARSTERN", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    
    ' Wasserversorger
    If InStr(1, nameUpper, "WASSER", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "ABWASSER", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "WASSERWERK", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "WASSERVERBAND", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "ZWECKVERBAND", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "BERLINWASSER", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "BERLINER WASSERBETRIEBE", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "BWB", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    
    ' Versicherungen
    If InStr(1, nameUpper, "VERSICHERUNG", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "ALLIANZ", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "ERGO", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "DEVK", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "HUK", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "AXA", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "GENERALI", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "ZURICH", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "GOTHAER", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "PROVINZIAL", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    
    ' Telekommunikation
    If InStr(1, nameUpper, "TELEKOM", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "VODAFONE", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "O2", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "1&1", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "UNITYMEDIA", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "KABEL DEUTSCHLAND", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "CONGSTAR", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    
    ' Entsorgung
    If InStr(1, nameUpper, "BSR", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "STADTREINIGUNG", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "ENTSORGUNG", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "ABFALLWIRTSCHAFT", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "AWB", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "REMONDIS", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "ALBA", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    
    ' GEZ/Rundfunk
    If InStr(1, nameUpper, "RUNDFUNK", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "BEITRAGSSERVICE", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "GEZ", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    If InStr(1, nameUpper, "ARD ZDF", vbTextCompare) > 0 Then IstVersorger = True: Exit Function
    
    IstVersorger = False
End Function

' ---------------------------------------------------------------------
' IstShop - Prüft ob Kontoname ein Shop/Händler ist
' WICHTIG: Wird VOR IstVersorger geprüft!
' ---------------------------------------------------------------------
Private Function IstShop(ByVal kontoname As String) As Boolean
    Dim nameUpper As String
    nameUpper = UCase(Trim(kontoname))
    
    If Len(nameUpper) = 0 Then
        IstShop = False
        Exit Function
    End If
    
    ' Supermärkte und Discounter
    If InStr(1, nameUpper, "LIDL", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "ALDI", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "REWE", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "EDEKA", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "PENNY", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "NETTO", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "KAUFLAND", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "REAL", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "GLOBUS", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "NORMA", vbTextCompare) > 0 Then IstShop = True: Exit Function
    
    ' Baumärkte
    If InStr(1, nameUpper, "BAUHAUS", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "HORNBACH", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "OBI", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "HAGEBAU", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "TOOM", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "HELLWEG", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "BAUMARKT", vbTextCompare) > 0 Then IstShop = True: Exit Function
    
    ' Online-Händler
    If InStr(1, nameUpper, "AMAZON", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "EBAY", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "ZALANDO", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "OTTO", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "MEDIAMARKT", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "MEDIA MARKT", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "SATURN", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "CONRAD", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "REICHELT", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "THOMANN", vbTextCompare) > 0 Then IstShop = True: Exit Function
    
    ' Drogerien
    If InStr(1, nameUpper, "DM", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "ROSSMANN", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "MUELLER", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "MÜLLER", vbTextCompare) > 0 Then IstShop = True: Exit Function
    
    ' Möbelhäuser
    If InStr(1, nameUpper, "IKEA", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "POCO", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "ROLLER", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "XXXLUTZ", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "HÖFFNER", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "HOEFFNER", vbTextCompare) > 0 Then IstShop = True: Exit Function
    
    ' Gartenmärkte
    If InStr(1, nameUpper, "DEHNER", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "GARTENCENTER", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "GARTENMARKT", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "PFLANZEN", vbTextCompare) > 0 Then IstShop = True: Exit Function
    
    ' Tankstellen
    If InStr(1, nameUpper, "ARAL", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "SHELL", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "ESSO", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "TOTAL", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "JET", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "TANKSTELLE", vbTextCompare) > 0 Then IstShop = True: Exit Function
    
    ' Bezahldienste / Payment
    If InStr(1, nameUpper, "PAYPAL", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "KLARNA", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "PAYDIREKT", vbTextCompare) > 0 Then IstShop = True: Exit Function
    
    ' Apotheken
    If InStr(1, nameUpper, "APOTHEKE", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "DOCMORRIS", vbTextCompare) > 0 Then IstShop = True: Exit Function
    If InStr(1, nameUpper, "SHOP APOTHEKE", vbTextCompare) > 0 Then IstShop = True: Exit Function
    
    IstShop = False
End Function

' ---------------------------------------------------------------------
' IstBank - Prüft ob Kontoname eine Bank ist
' ---------------------------------------------------------------------
Private Function IstBank(ByVal kontoname As String) As Boolean
    Dim nameUpper As String
    nameUpper = UCase(Trim(kontoname))
    
    If Len(nameUpper) = 0 Then
        IstBank = False
        Exit Function
    End If
    
    ' Deutsche Großbanken
    If InStr(1, nameUpper, "SPARKASSE", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "VOLKSBANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "RAIFFEISENBANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "COMMERZBANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "DEUTSCHE BANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "POSTBANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "HYPOVEREINSBANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "UNICREDIT", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "TARGOBANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "SANTANDER", vbTextCompare) > 0 Then IstBank = True: Exit Function
    
    ' Direktbanken
    If InStr(1, nameUpper, "ING", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "DKB", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "COMDIRECT", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "CONSORSBANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "N26", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "NORISBANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    
    ' Genossenschaftsbanken
    If InStr(1, nameUpper, "VR BANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "VR-BANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "PSD BANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "SPARDA", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "GLS BANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "ETHIKBANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    
    ' Landesbanken
    If InStr(1, nameUpper, "LANDESBANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "LBBW", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "HELABA", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "NORD/LB", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "BAYERNLB", vbTextCompare) > 0 Then IstBank = True: Exit Function
    
    ' Berliner Banken
    If InStr(1, nameUpper, "BERLINER SPARKASSE", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "BERLINER VOLKSBANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "WEBERBANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "INVESTITIONSBANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "IBB", vbTextCompare) > 0 Then IstBank = True: Exit Function
    
    ' Bausparkassen
    If InStr(1, nameUpper, "BAUSPARKASSE", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "LBS", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "SCHWÄBISCH HALL", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "WÜSTENROT", vbTextCompare) > 0 Then IstBank = True: Exit Function
    
    ' Allgemeine Begriffe
    If InStr(1, nameUpper, "BANK", vbTextCompare) > 0 Then IstBank = True: Exit Function
    If InStr(1, nameUpper, "KREDITINSTITUT", vbTextCompare) > 0 Then IstBank = True: Exit Function
    
    IstBank = False
End Function

' ---------------------------------------------------------------------
' CreateGUID - Erzeugt eine neue GUID für EntityKeys
' ---------------------------------------------------------------------
Private Function CreateGUID() As String
    ' Einfache GUID-Generierung basierend auf Zeit und Zufall
    Dim guid As String
    Dim i As Integer
    
    Randomize Timer
    
    guid = ""
    For i = 1 To 8
        guid = guid & Hex(Int(Rnd * 16))
    Next i
    guid = guid & "-"
    For i = 1 To 4
        guid = guid & Hex(Int(Rnd * 16))
    Next i
    guid = guid & "-"
    For i = 1 To 4
        guid = guid & Hex(Int(Rnd * 16))
    Next i
    guid = guid & "-"
    For i = 1 To 4
        guid = guid & Hex(Int(Rnd * 16))
    Next i
    guid = guid & "-"
    For i = 1 To 12
        guid = guid & Hex(Int(Rnd * 16))
    Next i
    
    CreateGUID = LCase(guid)
End Function

' ---------------------------------------------------------------------
' SetzeAmpelFarbe - Setzt die Hintergrundfarbe basierend auf Status
' ---------------------------------------------------------------------
Private Sub SetzeAmpelFarbe(ByVal zeile As Long, ByVal wsDaten As Worksheet, ByVal status As String)
    Dim rngZeile As Range
    Dim farbe As Long
    
    Set rngZeile = wsDaten.Range(wsDaten.Cells(zeile, EK_COL_ENTITYKEY), wsDaten.Cells(zeile, EK_COL_DEBUG))
    
    Select Case UCase(status)
        Case "GRUEN", "GRÜN", "GREEN", "OK"
            farbe = RGB(198, 239, 206)  ' Hellgrün
        Case "GELB", "YELLOW", "WARNUNG"
            farbe = RGB(255, 235, 156)  ' Hellgelb
        Case "ROT", "RED", "FEHLER"
            farbe = RGB(255, 199, 206)  ' Hellrot
        Case "WEISS", "WHITE", "NEUTRAL"
            farbe = RGB(255, 255, 255)  ' Weiß
        Case Else
            farbe = RGB(255, 255, 255)  ' Standard: Weiß
    End Select
    
    rngZeile.Interior.color = farbe
End Sub

' ---------------------------------------------------------------------
' FormatiereEntityKeyTabelle - Formatiert die gesamte EntityKey-Tabelle
' ---------------------------------------------------------------------
Public Sub FormatiereEntityKeyTabelle()
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim zeile As Long
    Dim rngHeader As Range
    Dim rngTabelle As Range
    
    On Error GoTo ErrorHandler
    
    Set wsDaten = ThisWorkbook.Worksheets("Daten")
    lastRow = wsDaten.Cells(wsDaten.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    
    If lastRow < EK_START_ROW Then
        Debug.Print "FormatiereEntityKeyTabelle: Keine Daten zum Formatieren"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Header formatieren (Zeile 3)
    Set rngHeader = wsDaten.Range(wsDaten.Cells(3, EK_COL_ENTITYKEY), wsDaten.Cells(3, EK_COL_DEBUG))
    With rngHeader
        .Font.Bold = True
        .Interior.color = RGB(68, 114, 196)  ' Blau
        .Font.color = RGB(255, 255, 255)     ' Weiß
        .HorizontalAlignment = xlCenter
    End With
    
    ' Datenbereich formatieren
    Set rngTabelle = wsDaten.Range(wsDaten.Cells(EK_START_ROW, EK_COL_ENTITYKEY), wsDaten.Cells(lastRow, EK_COL_DEBUG))
    
    With rngTabelle
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.color = RGB(180, 180, 180)
    End With
    
    ' Spaltenbreiten anpassen
    wsDaten.Columns(EK_COL_ENTITYKEY).ColumnWidth = 38     ' EntityKey
    wsDaten.Columns(EK_COL_IBAN).ColumnWidth = 24          ' IBAN
    wsDaten.Columns(EK_COL_KONTONAME).ColumnWidth = 30     ' Kontoname
    wsDaten.Columns(EK_COL_ZUORDNUNG).ColumnWidth = 25     ' Zuordnung
    wsDaten.Columns(EK_COL_PARZELLE).ColumnWidth = 10      ' Parzelle
    wsDaten.Columns(EK_COL_ROLE).ColumnWidth = 14          ' Role
    wsDaten.Columns(EK_COL_DEBUG).ColumnWidth = 40         ' Debug
    
    ' Jede Zeile einzeln formatieren
    For zeile = EK_START_ROW To lastRow
        FormatiereEntityKeyZeile zeile, wsDaten
    Next zeile
    
    Application.ScreenUpdating = True
    
    Debug.Print "FormatiereEntityKeyTabelle: " & (lastRow - EK_START_ROW + 1) & " Zeilen formatiert"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Debug.Print "FEHLER in FormatiereEntityKeyTabelle: " & Err.Description
End Sub

' ---------------------------------------------------------------------
' FormatiereEntityKeyZeile - Formatiert eine einzelne Zeile
' ---------------------------------------------------------------------
Public Sub FormatiereEntityKeyZeile(ByVal zeile As Long, ByVal wsDaten As Worksheet)
    Dim role As String
    Dim parzelle As Variant
    Dim status As String
    
    On Error GoTo ErrorHandler
    
    role = Trim(CStr(wsDaten.Cells(zeile, EK_COL_ROLE).value))
    parzelle = wsDaten.Cells(zeile, EK_COL_PARZELLE).value
    
    ' Status bestimmen basierend auf Role und Zuordnung
    If role = ROLE_MITGLIED Then
        If IsNumeric(parzelle) And parzelle >= 1 And parzelle <= 14 Then
            status = "GRUEN"
        Else
            status = "GELB"
        End If
    ElseIf role = ROLE_VERSORGER Or role = ROLE_BANK Or role = ROLE_SHOP Then
        status = "GRUEN"
    ElseIf role = ROLE_SONSTIGE Then
        status = "GELB"
    Else
        status = "ROT"
    End If
    
    SetzeAmpelFarbe zeile, wsDaten, status
    
    ' Zuordnung zentrieren wenn Parzelle
    If IsNumeric(parzelle) Then
        wsDaten.Cells(zeile, EK_COL_PARZELLE).HorizontalAlignment = xlCenter
    End If
    
    ' Role zentrieren
    wsDaten.Cells(zeile, EK_COL_ROLE).HorizontalAlignment = xlCenter
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "FEHLER in FormatiereEntityKeyZeile (Zeile " & zeile & "): " & Err.Description
End Sub

' ---------------------------------------------------------------------
' VerarbeiteManuelleRoleAenderung - Reagiert auf manuelle Änderungen
' Wird von Worksheet_Change aufgerufen
' ---------------------------------------------------------------------
Public Sub VerarbeiteManuelleRoleAenderung(ByVal Target As Range)
    Dim wsDaten As Worksheet
    Dim zeile As Long
    Dim neueRole As String
    Dim entityKey As String
    
    On Error GoTo ErrorHandler
    
    ' Nur wenn Änderung in Role-Spalte
    If Target.Column <> EK_COL_ROLE Then Exit Sub
    If Target.Row < EK_START_ROW Then Exit Sub
    
    Set wsDaten = Target.Worksheet
    zeile = Target.Row
    neueRole = Trim(CStr(Target.value))
    entityKey = Trim(CStr(wsDaten.Cells(zeile, EK_COL_ENTITYKEY).value))
    
    If Len(entityKey) = 0 Then Exit Sub
    
    ' Validierung der neuen Role
    Select Case neueRole
        Case ROLE_MITGLIED, ROLE_VERSORGER, ROLE_BANK, ROLE_SHOP, ROLE_SONSTIGE
            ' Gültige Role - nichts zu tun
        Case Else
            MsgBox "Ungültige EntityRole: " & neueRole & vbCrLf & vbCrLf & _
                   "Gültige Werte:" & vbCrLf & _
                   "- " & ROLE_MITGLIED & vbCrLf & _
                   "- " & ROLE_VERSORGER & vbCrLf & _
                   "- " & ROLE_BANK & vbCrLf & _
                   "- " & ROLE_SHOP & vbCrLf & _
                   "- " & ROLE_SONSTIGE, _
                   vbExclamation, "EntityRole-Fehler"
            Exit Sub
    End Select
    
    ' Debug-Info aktualisieren
    wsDaten.Cells(zeile, EK_COL_DEBUG).value = "Manuell: " & neueRole & " (" & Format(Now, "dd.mm.yyyy hh:mm") & ")"
    
    ' Zeile neu formatieren
    FormatiereEntityKeyZeile zeile, wsDaten
    
    ' Parzellen-Dropdown anpassen
    If neueRole = ROLE_MITGLIED Or neueRole = ROLE_SONSTIGE Then
        SetupParzellenDropdown wsDaten.Cells(zeile, EK_COL_PARZELLE)
    Else
        ' Dropdown entfernen und Zelle leeren
        With wsDaten.Cells(zeile, EK_COL_PARZELLE)
            .Validation.Delete
            .value = ""
        End With
    End If
    
    Debug.Print "VerarbeiteManuelleRoleAenderung: Zeile " & zeile & " -> " & neueRole
    Exit Sub
    
ErrorHandler:
    Debug.Print "FEHLER in VerarbeiteManuelleRoleAenderung: " & Err.Description
End Sub

' ---------------------------------------------------------------------
' GetEntityKeyStatistik - Gibt Statistik über EntityKeys zurück
' ---------------------------------------------------------------------
Public Function GetEntityKeyStatistik() As String
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim zeile As Long
    Dim role As String
    Dim cntMitglied As Long, cntVersorger As Long, cntBank As Long
    Dim cntShop As Long, cntSonstige As Long, cntGesamt As Long
    
    On Error GoTo ErrorHandler
    
    Set wsDaten = ThisWorkbook.Worksheets("Daten")
    lastRow = wsDaten.Cells(wsDaten.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    
    For zeile = EK_START_ROW To lastRow
        role = Trim(CStr(wsDaten.Cells(zeile, EK_COL_ROLE).value))
        cntGesamt = cntGesamt + 1
        
        Select Case role
            Case ROLE_MITGLIED: cntMitglied = cntMitglied + 1
            Case ROLE_VERSORGER: cntVersorger = cntVersorger + 1
            Case ROLE_BANK: cntBank = cntBank + 1
            Case ROLE_SHOP: cntShop = cntShop + 1
            Case ROLE_SONSTIGE: cntSonstige = cntSonstige + 1
        End Select
    Next zeile
    
    GetEntityKeyStatistik = "EntityKey-Statistik:" & vbCrLf & _
                            "===================" & vbCrLf & _
                            "Gesamt:    " & cntGesamt & vbCrLf & _
                            "Mitglied:  " & cntMitglied & vbCrLf & _
                            "Versorger: " & cntVersorger & vbCrLf & _
                            "Bank:      " & cntBank & vbCrLf & _
                            "Shop:      " & cntShop & vbCrLf & _
                            "Sonstige:  " & cntSonstige
    Exit Function
    
ErrorHandler:
    GetEntityKeyStatistik = "FEHLER: " & Err.Description
End Function

' ---------------------------------------------------------------------
' ZeigeEntityKeyStatistik - Zeigt Statistik in MsgBox
' ---------------------------------------------------------------------
Public Sub ZeigeEntityKeyStatistik()
    MsgBox GetEntityKeyStatistik(), vbInformation, "EntityKey-Statistik"
End Sub

' ---------------------------------------------------------------------
' ResetEntityKeyTabelle - Löscht alle EntityKey-Daten (mit Warnung!)
' ---------------------------------------------------------------------
Public Sub ResetEntityKeyTabelle()
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim antwort As VbMsgBoxResult
    
    antwort = MsgBox("ACHTUNG: Alle EntityKey-Daten werden gelöscht!" & vbCrLf & vbCrLf & _
                     "Dieser Vorgang kann NICHT rückgängig gemacht werden!" & vbCrLf & vbCrLf & _
                     "Wirklich fortfahren?", _
                     vbExclamation + vbYesNo + vbDefaultButton2, _
                     "EntityKey-Tabelle zurücksetzen")
    
    If antwort <> vbYes Then
        MsgBox "Abgebrochen.", vbInformation, "Reset abgebrochen"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Set wsDaten = ThisWorkbook.Worksheets("Daten")
    lastRow = wsDaten.Cells(wsDaten.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    
    If lastRow >= EK_START_ROW Then
        wsDaten.Range(wsDaten.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                      wsDaten.Cells(lastRow, EK_COL_DEBUG)).ClearContents
        wsDaten.Range(wsDaten.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                      wsDaten.Cells(lastRow, EK_COL_DEBUG)).Interior.ColorIndex = xlNone
    End If
    
    MsgBox "EntityKey-Tabelle wurde zurückgesetzt.", vbInformation, "Reset erfolgreich"
    Exit Sub
    
ErrorHandler:
    MsgBox "FEHLER beim Reset: " & Err.Description, vbCritical, "Fehler"
End Sub

' =====================================================================
' ENDE DES MODULS mod_EntityKey_Manager
' =====================================================================

