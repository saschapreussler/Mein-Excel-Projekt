Attribute VB_Name = "mod_Mitglieder_Logik"
' =============================================================================
' Modul:       mod_Mitglieder_Logik
' Beschreibung: Gesch?ftslogik f?r Mitgliederverwaltung
'               Extrahiert aus frm_Mitgliedsdaten.frm (SPLIT v1.0)
'               Enth?lt: Parzellen-Pr?fungen, Historie-Operationen,
'                        Validierungs- und Hilfsfunktionen
' Abh?ngigkeiten: mod_Const (alle Spalten-/Worksheet-Konstanten)
' Datum:         2025-06
' =============================================================================
Option Explicit

' =============================================================================
' HILFSFUNKTIONEN
' =============================================================================

' --- Hilfsfunktion f?r Parzelle -> Seite ---
Public Function GetSeiteFromParzelle(ByVal parzelle As String) As String
    Dim parzelleNum As Long
    
    If UCase(Trim(parzelle)) = "VEREIN" Then
        GetSeiteFromParzelle = "zentral"
        Exit Function
    End If
    
    On Error Resume Next
    parzelleNum = CLng(Left(parzelle, InStr(parzelle & " ", " ") - 1))
    On Error GoTo 0
    
    If parzelleNum = 0 Then
        GetSeiteFromParzelle = ""
        Exit Function
    End If
    
    If parzelleNum >= 1 And parzelleNum <= 9 Then
        GetSeiteFromParzelle = "rechts"
    ElseIf parzelleNum >= 10 And parzelleNum <= 14 Then
        GetSeiteFromParzelle = "links"
    Else
        GetSeiteFromParzelle = ""
    End If
    
End Function

' --- Pr?fe ob Funktion bereits existiert ---
Public Function FunktionExistiertBereits(ByVal funktion As String, ByVal ausschlussParzelle As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    lastRow = ws.Cells(ws.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        If ws.Cells(r, M_COL_FUNKTION).value = funktion And _
           ws.Cells(r, M_COL_PARZELLE).value <> ausschlussParzelle And _
           ws.Cells(r, M_COL_PARZELLE).value <> "" Then
            FunktionExistiertBereits = True
            Exit Function
        End If
    Next r
    
    FunktionExistiertBereits = False
End Function

' --- Hilfsfunktion: Pr?fe ob String eine Zahl ist ---
Public Function IsNumericTag(ByVal value As String) As Boolean
    Dim testVal As Long
    On Error Resume Next
    testVal = CLng(value)
    IsNumericTag = (Err.Number = 0)
    On Error GoTo 0
End Function

' --- Hilfsfunktion: Validiere Datumsformat ---
Public Function IstGueltigesDatum(ByVal datumStr As String) As Boolean
    If datumStr = "" Then
        IstGueltigesDatum = True  ' Leere Strings sind erlaubt
        Exit Function
    End If
    
    On Error Resume Next
    Dim testDatum As Date
    testDatum = CDate(datumStr)
    IstGueltigesDatum = (Err.Number = 0)
    On Error GoTo 0
End Function

' --- Pr?fe ob FormularForm geladen ist ---
Public Function IsFormLoaded(ByVal FormName As String) As Boolean
    Dim i As Long
    
    For i = 0 To VBA.UserForms.count - 1
        If StrComp(VBA.UserForms.item(i).Name, FormName, vbTextCompare) = 0 Then
            IsFormLoaded = True
            Exit Function
        End If
    Next i
    
    IsFormLoaded = False
End Function

' =============================================================================
' PARZELLEN-PR?FUNGEN
' =============================================================================

' ***************************************************************
' Pr?ft ob Person bereits auf dieser Parzelle existiert
' ***************************************************************
Public Function ExistiertBereitsAufParzelle(ByVal memberID As String, ByVal parzelle As String, Optional ByVal ausschlussZeile As Long = 0) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    lastRow = ws.Cells(ws.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        If r <> ausschlussZeile Then  ' Ignoriere die aktuelle Zeile bei Bearbeitung
            If ws.Cells(r, M_COL_MEMBER_ID).value = memberID And _
               StrComp(Trim(ws.Cells(r, M_COL_PARZELLE).value), Trim(parzelle), vbTextCompare) = 0 Then
                ExistiertBereitsAufParzelle = True
                Exit Function
            End If
        End If
    Next r
    
    ExistiertBereitsAufParzelle = False
End Function

' ***************************************************************
' Pr?ft ob auf einer Parzelle noch zahlende Mitglieder sind
' ***************************************************************
Public Function HatParzelleNochZahlendesMitglied(ByVal parzelle As String, ByVal ausschlussMemberID As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim funktion As String
    Dim memberID As String
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    lastRow = ws.Cells(ws.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        If StrComp(Trim(ws.Cells(r, M_COL_PARZELLE).value), parzelle, vbTextCompare) = 0 Then
            memberID = ws.Cells(r, M_COL_MEMBER_ID).value
            funktion = ws.Cells(r, M_COL_FUNKTION).value
            
            ' Ignoriere die auszuschlie?ende Member-ID
            If memberID <> ausschlussMemberID Then
                ' Pr?fe ob zahlendes Mitglied
                If funktion = "Mitglied mit Pacht" Or _
                   funktion = "1. Vorsitzende(r)" Or _
                   funktion = "2. Vorsitzende(r)" Or _
                   funktion = "Kassierer(in)" Or _
                   funktion = "Schriftf" & ChrW(252) & "hrer(in)" Then
                    HatParzelleNochZahlendesMitglied = True
                    Exit Function
                End If
            End If
        End If
    Next r
    
    HatParzelleNochZahlendesMitglied = False
End Function

' ***************************************************************
' Findet alle Parzellen eines Mitglieds anhand Member-ID
' ***************************************************************
Public Function GetParzellenVonMitglied(ByVal memberID As String) As String
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim parzellen As String
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    lastRow = ws.Cells(ws.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    parzellen = ""
    
    For r = M_START_ROW To lastRow
        If ws.Cells(r, M_COL_MEMBER_ID).value = memberID Then
            If parzellen = "" Then
                parzellen = ws.Cells(r, M_COL_PARZELLE).value
            Else
                parzellen = parzellen & ", " & ws.Cells(r, M_COL_PARZELLE).value
            End If
        End If
    Next r
    
    GetParzellenVonMitglied = parzellen
End Function

' ***************************************************************
' Pr?ft ob eine Parzelle zahlendes Mitglied hat
' ***************************************************************
Public Function ParzelleHatZahlendesMitglied(ByVal parzelle As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim funktion As String
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    lastRow = ws.Cells(ws.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        If StrComp(Trim(ws.Cells(r, M_COL_PARZELLE).value), Trim(parzelle), vbTextCompare) = 0 Then
            If Trim(ws.Cells(r, M_COL_PACHTENDE).value) = "" Then
                funktion = ws.Cells(r, M_COL_FUNKTION).value
                
                If funktion = "Mitglied mit Pacht" Or _
                   funktion = "1. Vorsitzende(r)" Or _
                   funktion = "2. Vorsitzende(r)" Or _
                   funktion = "Kassierer(in)" Or _
                   funktion = "Schriftf" & ChrW(252) & "hrer(in)" Then
                    ParzelleHatZahlendesMitglied = True
                    Exit Function
                End If
            End If
        End If
    Next r
    
    ParzelleHatZahlendesMitglied = False
End Function

' ***************************************************************
' Pr?ft ob Person auf Parzelle existiert
' ***************************************************************
Public Function ExistiertPersonAufParzelle(ByVal vorname As String, ByVal nachname As String, _
                                             ByVal parzelle As String, Optional ByVal ausschlussZeile As Long = 0) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    lastRow = ws.Cells(ws.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        If r <> ausschlussZeile Then
            If StrComp(Trim(ws.Cells(r, M_COL_PARZELLE).value), Trim(parzelle), vbTextCompare) = 0 And _
               StrComp(Trim(ws.Cells(r, M_COL_VORNAME).value), Trim(vorname), vbTextCompare) = 0 And _
               StrComp(Trim(ws.Cells(r, M_COL_NACHNAME).value), Trim(nachname), vbTextCompare) = 0 Then
                ExistiertPersonAufParzelle = True
                Exit Function
            End If
        End If
    Next r
    
    ExistiertPersonAufParzelle = False
End Function

' ***************************************************************
' Pr?ft ob Parzelle leer ist (keine aktiven Mitglieder)
' ***************************************************************
Public Function IstParzelleLeer(ByVal parzelle As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    lastRow = ws.Cells(ws.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        If StrComp(Trim(ws.Cells(r, M_COL_PARZELLE).value), Trim(parzelle), vbTextCompare) = 0 Then
            If Trim(ws.Cells(r, M_COL_PACHTENDE).value) = "" Then
                IstParzelleLeer = False
                Exit Function
            End If
        End If
    Next r
    
    IstParzelleLeer = True
End Function

' ***************************************************************
' Holt Namen des ersten aktiven Mitglieds auf Parzelle
' ***************************************************************
Public Function GetMitgliedNameAufParzelle(ByVal parzelle As String) As String
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    lastRow = ws.Cells(ws.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        If StrComp(Trim(ws.Cells(r, M_COL_PARZELLE).value), Trim(parzelle), vbTextCompare) = 0 Then
            If Trim(ws.Cells(r, M_COL_PACHTENDE).value) = "" Then
                GetMitgliedNameAufParzelle = ws.Cells(r, M_COL_NACHNAME).value & ", " & ws.Cells(r, M_COL_VORNAME).value
                Exit Function
            End If
        End If
    Next r
    
    GetMitgliedNameAufParzelle = ""
End Function

' =============================================================================
' AUSWAHL-DIALOG
' =============================================================================

' ***************************************************************
' Zeigt Auswahl-Dialog f?r mehrere Mitglieder auf einer Parzelle
' ***************************************************************
Public Function ZeigeAdressAuswahl(ByRef mitglieder As Collection) As Long
    Dim eingabe As String
    Dim auswahlText As String
    Dim i As Long
    Dim mitgliedInfo As Variant
    Dim auswahlNummer As Long
    
    auswahlText = "Geben Sie die Nummer des Mitglieds ein:" & vbCrLf & vbCrLf
    
    For i = 1 To mitglieder.count
        mitgliedInfo = mitglieder(i)
        auswahlText = auswahlText & i & " = " & mitgliedInfo(1) & ", " & mitgliedInfo(2) & vbCrLf
    Next i
    
    auswahlText = auswahlText & vbCrLf & "0 = Abbrechen"
    
    eingabe = InputBox(auswahlText, "Adresse ausw" & ChrW(228) & "hlen", "1")
    
    If eingabe = "" Then
        ZeigeAdressAuswahl = 0
        Exit Function
    End If
    
    On Error Resume Next
    auswahlNummer = CLng(eingabe)
    On Error GoTo 0
    
    If auswahlNummer < 0 Or auswahlNummer > mitglieder.count Then
        MsgBox "Ung" & ChrW(252) & "ltige Auswahl.", vbExclamation
        ZeigeAdressAuswahl = 0
    Else
        ZeigeAdressAuswahl = auswahlNummer
    End If
End Function

' =============================================================================
' HISTORIE-OPERATIONEN
' =============================================================================

' ***************************************************************
' Verschiebt ein Mitglied von Mitgliederliste in Mitgliederhistorie
' NEUE STRUKTUR: 10 Spalten (A-J)
' ***************************************************************
Public Sub VerschiebeInHistorie(ByVal lRow As Long, ByVal parzelle As String, ByVal memberID As String, _
                                   ByVal nachname As String, ByVal vorname As String, _
                                   ByVal austrittsDatum As Date, ByVal grund As String, _
                                   Optional ByVal nachpaechterName As String = "", _
                                   Optional ByVal nachpaechterID As String = "")
    
    Dim wsM As Worksheet
    Dim wsH As Worksheet
    Dim nextHistRow As Long
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    
    ' === SICHERHEITSCHECK: NIEMALS Verein-Parzelle l?schen ===
    If UCase(Trim(parzelle)) = "VEREIN" Then
        MsgBox "KRITISCHER FEHLER: Versuch, die Verein-Parzelle zu l" & ChrW(246) & "schen wurde verhindert!" & vbCrLf & _
               "Zeile " & lRow & ", Member-ID: " & memberID, vbCritical, "Sicherheitswarnung"
        Exit Sub
    End If
    
    ' Entsperre beide Bl?tter
    wsM.Unprotect PASSWORD:=PASSWORD
    wsH.Unprotect PASSWORD:=PASSWORD
    
    ' Finde n?chste freie Zeile in Mitgliederhistorie (ab Zeile 4)
    nextHistRow = wsH.Cells(wsH.Rows.count, H_COL_NAME_EHEM_PAECHTER).End(xlUp).Row + 1
    If nextHistRow < H_START_ROW Then nextHistRow = H_START_ROW
    
    ' Schreibe Daten in Mitgliederhistorie (10 Spalten A-J) - MIT FEHLERBEHANDLUNG
    wsH.Cells(nextHistRow, H_COL_PARZELLE).value = parzelle                          ' A: Parzelle
    wsH.Cells(nextHistRow, H_COL_MEMBER_ID_ALT).value = memberID                     ' B: Member ID (alt)
    wsH.Cells(nextHistRow, H_COL_NAME_EHEM_PAECHTER).value = nachname & ", " & vorname  ' C: Name ehem. P?chter (kombiniert)
    
    On Error Resume Next
    wsH.Cells(nextHistRow, H_COL_AUST_DATUM).value = austrittsDatum                  ' D: Austrittsdatum
    If Err.Number = 0 Then
        wsH.Cells(nextHistRow, H_COL_AUST_DATUM).NumberFormat = "dd.mm.yyyy"
    End If
    On Error GoTo 0
    
    wsH.Cells(nextHistRow, H_COL_GRUND).value = grund                                ' E: Grund
    wsH.Cells(nextHistRow, H_COL_NACHPAECHTER_NAME).value = nachpaechterName         ' F: Name neuer P?chter
    wsH.Cells(nextHistRow, H_COL_NACHPAECHTER_ID).value = nachpaechterID             ' G: ID neuer P?chter
    wsH.Cells(nextHistRow, H_COL_KOMMENTAR).value = ""                               ' H: Kommentar (leer)
    wsH.Cells(nextHistRow, H_COL_ENDABRECHNUNG).value = ""                           ' I: Endabrechnung (leer)
    
    On Error Resume Next
    wsH.Cells(nextHistRow, H_COL_SYSTEMZEIT).value = Now                             ' J: Systemzeit
    If Err.Number = 0 Then
        wsH.Cells(nextHistRow, H_COL_SYSTEMZEIT).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    End If
    On Error GoTo 0
    
    ' L?sche Zeile aus Mitgliederliste
    wsM.Rows(lRow).Delete Shift:=xlUp
    
    ' Sch?tze Bl?tter wieder
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    wsH.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Dim nachpaechterInfo As String
    If nachpaechterName <> "" Then
        nachpaechterInfo = vbCrLf & "Nachp" & ChrW(228) & "chter: " & nachpaechterName
    Else
        nachpaechterInfo = ""
    End If
    
    MsgBox "Mitglied " & nachname & " wurde in die Mitgliederhistorie verschoben." & vbCrLf & _
           "Grund: " & grund & nachpaechterInfo, vbInformation
    
    Exit Sub
ErrorHandler:
    On Error GoTo 0
    If Not wsM Is Nothing Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    If Not wsH Is Nothing Then wsH.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler beim Verschieben in Historie: " & Err.Description, vbCritical
End Sub

' ***************************************************************
' Speichert Parzellenwechsel in Mitgliederhistorie
' ***************************************************************
Public Sub SpeichereParzellenwechselInHistorie(ByVal alteParzelle As String, ByVal neueParzelle As String, _
                                                  ByVal memberID As String, ByVal nachname As String, _
                                                  ByVal vorname As String, ByVal grund As String)
    Dim wsH As Worksheet
    Dim nextHistRow As Long
    
    On Error GoTo ErrorHandler
    
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    
    wsH.Unprotect PASSWORD:=PASSWORD
    
    nextHistRow = wsH.Cells(wsH.Rows.count, H_COL_NAME_EHEM_PAECHTER).End(xlUp).Row + 1
    If nextHistRow < H_START_ROW Then nextHistRow = H_START_ROW
    
    wsH.Cells(nextHistRow, H_COL_PARZELLE).value = alteParzelle                     ' A: Alte Parzelle
    wsH.Cells(nextHistRow, H_COL_MEMBER_ID_ALT).value = memberID                    ' B: Member ID (bleibt gleich)
    wsH.Cells(nextHistRow, H_COL_NAME_EHEM_PAECHTER).value = nachname & ", " & vorname  ' C: Name
    
    On Error Resume Next
    wsH.Cells(nextHistRow, H_COL_AUST_DATUM).value = Date                           ' D: Wechseldatum
    wsH.Cells(nextHistRow, H_COL_AUST_DATUM).NumberFormat = "dd.mm.yyyy"
    On Error GoTo ErrorHandler
    
    wsH.Cells(nextHistRow, H_COL_GRUND).value = grund                               ' E: Grund
    wsH.Cells(nextHistRow, H_COL_NACHPAECHTER_NAME).value = ""                      ' F: kein Nachp?chter
    wsH.Cells(nextHistRow, H_COL_NACHPAECHTER_ID).value = ""                        ' G: kein Nachp?chter
    wsH.Cells(nextHistRow, H_COL_KOMMENTAR).value = "Neue Parzelle: " & neueParzelle ' H: Kommentar
    wsH.Cells(nextHistRow, H_COL_ENDABRECHNUNG).value = ""                          ' I: keine Endabrechnung
    
    On Error Resume Next
    wsH.Cells(nextHistRow, H_COL_SYSTEMZEIT).value = Now                            ' J: Systemzeit
    wsH.Cells(nextHistRow, H_COL_SYSTEMZEIT).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    On Error GoTo ErrorHandler
    
    wsH.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Exit Sub
    
ErrorHandler:
    If Not wsH Is Nothing Then wsH.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Debug.Print "Fehler beim Speichern in Historie: " & Err.Description
End Sub

' ***************************************************************
' Verschiebt ALLE Eintr?ge eines Mitglieds (alle Parzellen)
' von der Mitgliederliste in die Mitgliederhistorie.
' Wird beim Komplett-Austritt bei Mehrfach-Parzellen aufgerufen.
' ***************************************************************
Public Sub VerschiebeAlleParzellenInHistorie(ByVal memberID As String, _
                                               ByVal nachname As String, ByVal vorname As String, _
                                               ByVal austrittsDatum As Date, ByVal grund As String)
    
    Dim wsM As Worksheet
    Dim wsH As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim nextHistRow As Long
    Dim parzelle As String
    Dim deletedCount As Long
    Dim parzellenListe As String
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    
    ' Entsperre beide Bl?tter
    wsM.Unprotect PASSWORD:=PASSWORD
    wsH.Unprotect PASSWORD:=PASSWORD
    
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    deletedCount = 0
    parzellenListe = ""
    
    ' R?CKW?RTS durchlaufen wegen Zeilen-L?schung!
    For r = lastRow To M_START_ROW Step -1
        If wsM.Cells(r, M_COL_MEMBER_ID).value = memberID Then
            parzelle = wsM.Cells(r, M_COL_PARZELLE).value
            
            ' === SICHERHEITSCHECK: NIEMALS Verein-Zeile l?schen ===
            If UCase(Trim(parzelle)) = "VEREIN" Then
                Debug.Print "WARNUNG: Verein-Zeile " & ChrW(252) & "bersprungen (Zeile " & r & ") bei Komplett-Austritt"
                GoTo NextRowKomplett
            End If
            
            ' Sammle Parzellen f?r die MsgBox
            If parzellenListe = "" Then
                parzellenListe = parzelle
            Else
                parzellenListe = parzelle & ", " & parzellenListe
            End If
            
            ' Finde n?chste freie Zeile in Mitgliederhistorie
            nextHistRow = wsH.Cells(wsH.Rows.count, H_COL_NAME_EHEM_PAECHTER).End(xlUp).Row + 1
            If nextHistRow < H_START_ROW Then nextHistRow = H_START_ROW
            
            ' Schreibe Daten in Mitgliederhistorie (Spalten A-J)
            wsH.Cells(nextHistRow, H_COL_PARZELLE).value = parzelle                            ' A: Parzelle
            wsH.Cells(nextHistRow, H_COL_MEMBER_ID_ALT).value = memberID                       ' B: Member ID
            wsH.Cells(nextHistRow, H_COL_NAME_EHEM_PAECHTER).value = nachname & ", " & vorname  ' C: Name
            
            On Error Resume Next
            wsH.Cells(nextHistRow, H_COL_AUST_DATUM).value = austrittsDatum                    ' D: Austrittsdatum
            If Err.Number = 0 Then
                wsH.Cells(nextHistRow, H_COL_AUST_DATUM).NumberFormat = "dd.mm.yyyy"
            End If
            On Error GoTo ErrorHandler
            
            wsH.Cells(nextHistRow, H_COL_GRUND).value = grund                                  ' E: Grund
            wsH.Cells(nextHistRow, H_COL_NACHPAECHTER_NAME).value = ""                         ' F: kein Nachp?chter
            wsH.Cells(nextHistRow, H_COL_NACHPAECHTER_ID).value = ""                           ' G: kein Nachp?chter
            wsH.Cells(nextHistRow, H_COL_KOMMENTAR).value = "Komplett-Austritt (alle Parzellen)" ' H: Kommentar
            wsH.Cells(nextHistRow, H_COL_ENDABRECHNUNG).value = ""                             ' I: Endabrechnung
            
            On Error Resume Next
            wsH.Cells(nextHistRow, H_COL_SYSTEMZEIT).value = Now                               ' J: Systemzeit
            If Err.Number = 0 Then
                wsH.Cells(nextHistRow, H_COL_SYSTEMZEIT).NumberFormat = "dd.mm.yyyy hh:mm:ss"
            End If
            On Error GoTo ErrorHandler
            
            ' L?sche Zeile aus Mitgliederliste
            wsM.Rows(r).Delete Shift:=xlUp
            deletedCount = deletedCount + 1
        End If
NextRowKomplett:
    Next r
    
    ' Sch?tze Bl?tter wieder
    wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    wsH.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    MsgBox "KOMPLETT-AUSTRITT durchgef" & ChrW(252) & "hrt!" & vbCrLf & vbCrLf & _
           "Mitglied: " & nachname & ", " & vorname & vbCrLf & _
           "Parzellen: " & parzellenListe & vbCrLf & _
           "Anzahl verschobener Eintr" & ChrW(228) & "ge: " & deletedCount & vbCrLf & _
           "Grund: " & grund & vbCrLf & _
           "Datum: " & Format(austrittsDatum, "dd.mm.yyyy"), vbInformation, "Austritt abgeschlossen"
    
    Exit Sub
    
ErrorHandler:
    On Error GoTo 0
    If Not wsM Is Nothing Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    If Not wsH Is Nothing Then wsH.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler beim Komplett-Austritt: " & Err.Description, vbCritical
End Sub


































