Attribute VB_Name = "mod_Mitglieder_UI"
Option Explicit

' ***************************************************************
' HILFSFUNKTIONEN ZUM BLATTSCHUTZ & UI-AKTUALISIERUNG
' ***************************************************************
Public Sub UnprotectSheet(ByRef ws As Worksheet)
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Unprotect PASSWORD:=PASSWORD
        On Error GoTo 0
    End If
End Sub

Public Sub ProtectSheet(ByRef ws As Worksheet)
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        On Error GoTo 0
    End If
End Sub

Public Sub RefreshAllLists()
    ' Aktualisiert alle Dropdowns und sortiert neu (zentraler Aufruf)
    Call Sortiere_Mitgliederliste_Nach_Parzelle
    ' Füge hier weitere Aktualisierungsfunktionen hinzu, falls nötig
End Sub

' ***************************************************************
' PROZEDUR: AktualisiereDatenstand (KORRIGIERT: Nutzt Unprotect/ProtectSheet)
' ***************************************************************
Public Sub AktualisiereDatenstand()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = Worksheets(WS_MITGLIEDER)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        Call UnprotectSheet(ws)
        With ws.Cells(M_STAND_ROW, M_STAND_COL)
            .Value = Now
        End With
        Call ProtectSheet(ws)
    Else
        Debug.Print "Fehler: Tabellenblatt '" & WS_MITGLIEDER & "' nicht gefunden."
    End If
End Sub

' ***************************************************************
' PROZEDUR: Fuelle_MemberIDs_Wenn_Fehlend (NEU: Fügt die eindeutige MemberID in Spalte A hinzu)
' ***************************************************************
Public Sub Fuelle_MemberIDs_Wenn_Fehlend()

    Dim wsM As Worksheet
    Dim lastRow As Long
    Dim lRow As Long
    Dim wasProtected As Boolean
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If wsM Is Nothing Then Exit Sub
    
    wasProtected = wsM.ProtectContents
    If wasProtected Then Call UnprotectSheet(wsM)
    
    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
    
    If lastRow < M_START_ROW Then GoTo Cleanup
    
    Application.ScreenUpdating = False
    
    ' Header setzen
    wsM.Cells(M_HEADER_ROW, M_COL_MEMBER_ID).Value = "Member ID"
    
    ' Schleife durch alle Zeilen ab M_START_ROW
    For lRow = M_START_ROW To lastRow
        ' Prüfen, ob eine MemberID fehlt und ob der Datensatz nicht leer ist (Nachname gefüllt)
        If wsM.Cells(lRow, M_COL_MEMBER_ID).Value = "" And _
           wsM.Cells(lRow, M_COL_NACHNAME).Value <> "" Then
            
            ' GUID generieren und eintragen
            wsM.Cells(lRow, M_COL_MEMBER_ID).Value = CreateGUID()
        End If
    Next lRow
    
    ' *** ZELLSPERRUNG FÜR SPALTE A ***
    ' Spalte A sperren, damit die ID nicht manuell verändert wird.
    With wsM.Range(wsM.Cells(M_START_ROW, M_COL_MEMBER_ID), wsM.Cells(lastRow + 1000, M_COL_MEMBER_ID))
        .Locked = True
        .FormulaHidden = True
    End With
    
Cleanup:
    Application.ScreenUpdating = True
    If wasProtected Then Call ProtectSheet(wsM)
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler beim Füllen der MemberIDs: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ***************************************************************
' HILFSFUNKTION: GUID (Globally Unique Identifier) erstellen (NEU) (KORRIGIERT: On Error GoTo 0 hinzugefügt)
' ***************************************************************
Public Function CreateGUID() As String
    ' Benötigt KEINEN Verweis, nutzt das Scriptlet.TypeLib-Objekt zur Laufzeit.
    
    Dim TypeLib As Object
    ' Versuch 1: GUID per Scripting Runtime
    On Error Resume Next ' Nur für diesen Aufruf
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    On Error GoTo 0      ' Fehlerbehandlung zurücksetzen
    
    If Not TypeLib Is Nothing Then
        CreateGUID = Mid(TypeLib.GUID, 2, 36) ' Entfernt die Klammern und gibt die reine GUID zurück
    End If
    
    If CreateGUID = "" Then
        ' Versuch 2: Notfall-GUID (falls TypeLib blockiert/nicht verfügbar)
        Randomize
        CreateGUID = Format(Now, "yyyymmddhhmmss") & "-" & Int((99999 - 10000 + 1) * Rnd + 10000)
    End If
    
    Set TypeLib = Nothing
End Function

' ***************************************************************
' PROZEDUR: ApplyMitgliederDropdowns (FINAL KORRIGIERT: Spalte C/Seite gesperrt, Spalte B/Parzelle entsperrt)
' ***************************************************************
Public Sub ApplyMitgliederDropdowns()
    Dim ws As Worksheet
    On Error GoTo ErrorHandler
    Set ws = Worksheets(WS_MITGLIEDER)
    Call UnprotectSheet(ws)
    
    ' ***************************************************************
    ' KORREKTUR DER SPERRUNGEN (Locked = False)
    ' ***************************************************************
    
    ' Spalte B (Parzelle): Muss entsperrt werden, damit der Benutzer Parzelle wählen kann.
    ws.Range(ws.Cells(M_START_ROW, M_COL_PARZELLE), ws.Cells(1000, M_COL_PARZELLE)).Locked = False
    
    ' Spalte D (Anrede): Entsperren für Dropdown-Auswahl.
    ws.Range(ws.Cells(M_START_ROW, M_COL_ANREDE), ws.Cells(1000, M_COL_ANREDE)).Locked = False
    
    ' Spalte O (Funktion): Entsperren für Dropdown-Auswahl.
    ws.Range(ws.Cells(M_START_ROW, M_COL_FUNKTION), ws.Cells(1000, M_COL_FUNKTION)).Locked = False

    ' Dropdowns anwenden
    ' Dropdown für Parzelle (B) wieder aktivieren! (Entsperrte Zelle braucht die Liste)
    Call ApplyDropdown(ws.Range(ws.Cells(M_START_ROW, M_COL_PARZELLE), ws.Cells(1000, M_COL_PARZELLE)), "=Daten!$F$4:$F$18", True)
    
    ' Dropdown für Seite (C) aktivieren! (Gesperrte Zelle braucht die Liste für UserInterfaceOnly)
    Call ApplyDropdown(ws.Range(ws.Cells(M_START_ROW, M_COL_SEITE), ws.Cells(1000, M_COL_SEITE)), "=Daten!$H$4:$H$6", True)
    
    Call ApplyDropdown(ws.Range(ws.Cells(M_START_ROW, M_COL_ANREDE), ws.Cells(1000, M_COL_ANREDE)), "=Daten!$D$4:$D$9", True)
    Call ApplyDropdown(ws.Range(ws.Cells(M_START_ROW, M_COL_FUNKTION), ws.Cells(1000, M_COL_FUNKTION)), "=Daten!$B$4:$B$11", True)

    Call ProtectSheet(ws)
    Exit Sub
ErrorHandler:
    Call ProtectSheet(ws)
    MsgBox "Fehler beim Setzen der Dropdown-Listen: " & Err.Description, vbCritical
End Sub

Public Sub Reapply_Data_Validation()
    Call ApplyMitgliederDropdowns
End Sub

Private Sub ApplyDropdown(ByVal targetRange As Range, ByVal sourceFormula As String, ByVal allowBlanks As Boolean)
    With targetRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=sourceFormula
        .IgnoreBlank = allowBlanks
        .InCellDropdown = True
        .ErrorTitle = "Ungültiger Wert"
        .ErrorMessage = "Bitte wählen Sie einen Wert aus der Liste."
    End With
End Sub


' ***************************************************************
' PROZEDUR: Anwende_Zebra_Formatierung (Universelle BF mit Prüfspalte)
' ***************************************************************
Public Sub Anwende_Zebra_Formatierung(ByVal ws As Worksheet, ByVal startCol As Long, ByVal endCol As Long, ByVal startRow As Long, ByVal dataCheckCol As Long)
    
    Const ZEBRA_COLOR As Long = &HDEE5E3
    
    If ws Is Nothing Then Exit Sub

    Dim rngFullData As Range
    Dim sFormula As String
    
    ' 1. Zielbereich definieren
    Set rngFullData = ws.Range(ws.Cells(startRow, startCol), ws.Cells(1000, endCol))
    
    ' 2. Bestehende Regeln im BF-Bereich LÖSCHEN
    On Error Resume Next
    rngFullData.FormatConditions.Delete
    On Error GoTo 0
    
    ' 3. Explizites Entfernen aller manuellen Zellfüllungen im Bereich
    rngFullData.Interior.color = xlNone
    
    ' 4. Formel erstellen: =UND(NICHT(ISTLEER($[Prüfspalte][Startzeile])); REST(ZEILE();2)=0)
    Dim checkColLetter As String
    checkColLetter = Split(ws.Columns(dataCheckCol).Address(False, False), ":")(0)
    
    sFormula = "=UND(NICHT(ISTLEER($" & checkColLetter & startRow & ")); REST(ZEILE();2)=0)"
    
    With rngFullData.FormatConditions.Add(Type:=xlExpression, Formula1:=sFormula)
        .Interior.color = ZEBRA_COLOR
        .StopIfTrue = True
        .Priority = 1
    End With

End Sub


' ***************************************************************
' PROZEDUR: Formatiere_Alle_Tabellen_Neu (KORRIGIERT: Hardcoding Spalten B und U)
' ***************************************************************
Public Sub Formatiere_Alle_Tabellen_Neu()

    Dim wsM As Worksheet
    Dim wsD As Worksheet
    Dim wasProtectedM As Boolean
    Dim wasProtectedD As Boolean
    
    ' --- ZUSÄTZLICHE KONSTANTEN FÜR DIESE PROZEDUR (Spalten Hardcoded zur Stabilität) ---
    Const DATA_START_ROW As Long = 4
    Const M_START_COL As Long = 1      ' KORRIGIERT: Startet jetzt in Spalte A (MemberID)
    Const M_CHECK_COL As Long = 5      ' Spalte E (Nachname) - Prüfspalte Mitglieder
    Const D_ENTITYKEY_END_COL As Long = 21 ' Spalte U - Endspalte EntityKey
    ' --- ENDE ZUSÄTZLICHE KONSTANTEN ---

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' 1. Mitgliederliste (WS_MITGLIEDER)
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If Not wsM Is Nothing Then
        wasProtectedM = wsM.ProtectContents
        If wasProtectedM Then Call UnprotectSheet(wsM)
        
        ' BF: Start A (1), Ende P (M_COL_PACHTENDE), Startzeile 6, Prüfspalte E (5)
        Call Anwende_Zebra_Formatierung(wsM, M_START_COL, M_COL_PACHTENDE, M_START_ROW, M_CHECK_COL)
        
        If wasProtectedM Then Call ProtectSheet(wsM)
    End If
    
    ' 2. Datenblatt (WS_DATEN)
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    If Not wsD Is Nothing Then
        wasProtectedD = wsD.ProtectContents
        If wasProtectedD Then Call UnprotectSheet(wsD)
        
        ' BF 1: Kategorie-Regeln (J bis Q, Startzeile 4, Prüfspalte J)
        Call Anwende_Zebra_Formatierung(wsD, DATA_CAT_COL_START, DATA_CAT_COL_END, DATA_START_ROW, DATA_CAT_COL_START)
        
        ' BF 2: EntityKey/Mapping-Tabelle (S bis U (21), Startzeile 4, Prüfspalte S)
        Call Anwende_Zebra_Formatierung(wsD, DATA_MAP_COL_ENTITYKEY, D_ENTITYKEY_END_COL, DATA_START_ROW, DATA_MAP_COL_ENTITYKEY)
        
        If wasProtectedD Then Call ProtectSheet(wsD)
    End If
    
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "FEHLER beim Formatieren der Tabellen: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ***************************************************************
' PROZEDUR: Sortiere_Mitgliederliste_Nach_Parzelle (KORRIGIERT: Sortierbereich umfasst jetzt Spalte A)
' ***************************************************************
Public Sub Sortiere_Mitgliederliste_Nach_Parzelle()

    Dim ws As Worksheet
    Dim rngSort As Range
    Dim lastRow As Long
    Dim wasProtected As Boolean
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    If ws Is Nothing Then Exit Sub ' Exit, wenn Blatt nicht gefunden wird
    
    wasProtected = ws.ProtectContents
    If wasProtected Then Call UnprotectSheet(ws)
    
    ' Nachname (E) wird als robusteste Spalte zum Finden der letzten Zeile verwendet
    lastRow = ws.Cells(ws.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
    
    If lastRow < M_START_ROW Then GoTo Cleanup
    
    ' Sortierbereich: KORRIGIERT: Von Spalte A (1) bis P (M_COL_PACHTENDE)
    Set rngSort = ws.Range(ws.Cells(M_START_ROW, 1), ws.Cells(lastRow, M_COL_PACHTENDE))
    
    With ws.Sort
        .SortFields.Clear
        ' 1. Sortierkriterium: Pachtende (P) - um Ehemalige unten zu halten
        .SortFields.Add Key:=ws.Columns(M_COL_PACHTENDE), SortOn:=xlSortOnValues, Order:=xlAscending
        ' 2. Sortierkriterium: Parzelle (B) - Hauptsortierung
        .SortFields.Add Key:=ws.Columns(M_COL_PARZELLE), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        ' 3. Sortierkriterium: Anrede (D) - sekundäre Sortierung
        .SortFields.Add Key:=ws.Columns(M_COL_ANREDE), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange rngSort
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Nach Sortierung: Validation und Formatierung erneut anwenden
    Call Reapply_Data_Validation
    Call Formatiere_Alle_Tabellen_Neu
    
Cleanup:
    If Not ws Is Nothing Then
        If wasProtected Then Call ProtectSheet(ws)
    End If
    Exit Sub

ErrorHandler:
    If Not ws Is Nothing Then
        If wasProtected Then Call ProtectSheet(ws)
    End If
    MsgBox "FEHLER BEIM SORTIEREN (mod_Mitglieder_UI):" & vbCrLf & "Nr: " & Err.Number & vbCrLf & "Beschreibung: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ***************************************************************
' PROZEDUR: AktualisiereNamedRange_MitgliederNamen (NEU)
' ***************************************************************
Public Sub AktualisiereNamedRange_MitgliederNamen()
    ' Muss im Modul UI sein, da es das UI-Element ComboBox/Dropdown beeinflusst
    
    Dim wsM As Worksheet
    Dim lastRow As Long
    Dim rngTarget As Range
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_NACHNAME).End(xlUp).Row
    
    ' Named Range auf Spalte D (Nachname) von M_START_ROW bis zur letzten Zeile
    Set rngTarget = wsM.Range(wsM.Cells(M_START_ROW, M_COL_NACHNAME), wsM.Cells(lastRow, M_COL_NACHNAME))
    
    ' Definiere den Namen 'MitgliederNamen' neu oder aktualisiere ihn
    On Error Resume Next ' Fehler ignorieren, falls der Name nicht existiert
    ThisWorkbook.Names("MitgliederNamen").Delete
    On Error GoTo ErrorHandler ' Fehlerbehandlung wieder aktivieren
    
    ThisWorkbook.Names.Add Name:="MitgliederNamen", RefersTo:=rngTarget
    
    Exit Sub
ErrorHandler:
    Debug.Print "Fehler beim Aktualisieren des Named Range 'MitgliederNamen': " & Err.Description
End Sub


' ***************************************************************
' FUNKTION: GetEntityKeyByParzelle
' Sucht den EntityKey im Blatt "Daten" basierend auf der Parzellennummer.
' *********************************************************************************
Public Function GetEntityKeyByParzelle(ByVal ParzelleNr As String) As String
    ' WICHTIG: Sie wird beibehalten, da sie für das *Banking-Mapping* relevant sein kann,
    ' aber für die Mitglieder-UI ist sie unsicher bei Doppelbelegung!
    Dim wsD As Worksheet
    Dim lastRow As Long
    Dim rngFind As Range
    
    If ParzelleNr = "" Then
        GetEntityKeyByParzelle = ""
        Exit Function
    End If
    
    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    If wsD Is Nothing Then GoTo ErrorHandler
    
    ' Sucht in der Spalte DATA_MAP_COL_PARZELLE (W)
    lastRow = wsD.Cells(wsD.Rows.Count, DATA_MAP_COL_PARZELLE).End(xlUp).Row
    Set rngFind = wsD.Range(wsD.Cells(DATA_START_ROW, DATA_MAP_COL_PARZELLE), wsD.Cells(lastRow, DATA_MAP_COL_PARZELLE))
    
    ' Führt die Suche durch
    Set rngFind = rngFind.Find(What:=ParzelleNr, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rngFind Is Nothing Then
        ' Wenn gefunden, geben wir den Wert aus Spalte DATA_MAP_COL_ENTITYKEY (S) in der gleichen Zeile zurück
        GetEntityKeyByParzelle = wsD.Cells(rngFind.Row, DATA_MAP_COL_ENTITYKEY).Value
    Else
        GetEntityKeyByParzelle = ""
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Fehler in GetEntityKeyByParzelle: " & Err.Description
    GetEntityKeyByParzelle = ""
End Function


Private Function FindeRowByMemberID(ByVal MemberID As String) As Long

    Dim wsM As Worksheet
    Dim rngSearch As Range
    Dim rngFind As Range
    Dim lastRow As Long
    Dim bWasProtected As Boolean

    FindeRowByMemberID = 0
    If Trim(MemberID) = "" Then Exit Function

    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)

    ' ------------------------------------------------------------
    ' Blattschutz merken und ggf. aufheben
    ' ------------------------------------------------------------
    bWasProtected = wsM.ProtectContents
    If bWasProtected Then
        mod_Mitglieder_UI.UnprotectSheet wsM
    End If

    ' ------------------------------------------------------------
    ' Filter zuverlässig entfernen
    ' ------------------------------------------------------------
    If wsM.AutoFilterMode Then
        If wsM.FilterMode Then wsM.ShowAllData
    End If

    lastRow = wsM.Cells(wsM.Rows.Count, M_COL_MEMBER_ID).End(xlUp).Row
    If lastRow < M_START_ROW Then GoTo CleanExit

    Set rngSearch = wsM.Range( _
        wsM.Cells(M_START_ROW, M_COL_MEMBER_ID), _
        wsM.Cells(lastRow, M_COL_MEMBER_ID) _
    )

    Set rngFind = rngSearch.Find( _
        What:=MemberID, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False _
    )

    If Not rngFind Is Nothing Then
        FindeRowByMemberID = rngFind.Row
    End If

CleanExit:
    ' ------------------------------------------------------------
    ' Blattschutz wiederherstellen
    ' ------------------------------------------------------------
    If bWasProtected Then
        mod_Mitglieder_UI.ProtectSheet wsM
    End If

End Function



' ***************************************************************
' PROZEDUR: Speichere_Historie_und_Aktualisiere_Mitgliederliste (KORRIGIERT & ERWEITERT)
' DIESE PROZEDUR WIRD NACHHER AUS DER USERFORM AUSGELÖST!
' ***************************************************************
Public Sub Speichere_Historie_und_Aktualisiere_Mitgliederliste( _
    ByVal selectedRow As Long, _
    ByVal OldParzelle As String, _
    ByVal OldMemberID As String, _
    ByVal Nachname As String, _
    ByVal AustrittsDatum As Date, _
    ByVal NewParzelleNr As String, _
    ByVal NewMemberID As String, _
    ByVal ChangeReason As String)

    Dim wsM As Worksheet
    Dim wsH As Worksheet
    Dim NextRow As Long
    Dim UebernehmerRow As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    
    ' --- 1. HISTORIE SPEICHERN ---
    Call UnprotectSheet(wsH) ' Historie-Blatt entsperren
    NextRow = wsH.Cells(wsH.Rows.Count, H_COL_PARZELLE).End(xlUp).Row + 1
    If NextRow < H_START_ROW Then NextRow = H_START_ROW ' Sicherstellen, dass mindestens H_START_ROW verwendet wird
    
    ' Daten in das Historie-Blatt schreiben
    wsH.Cells(NextRow, H_COL_PARZELLE).Value = OldParzelle
    wsH.Cells(NextRow, H_COL_MITGL_ID).Value = OldMemberID
    wsH.Cells(NextRow, H_COL_NACHNAME).Value = Nachname
    wsH.Cells(NextRow, H_COL_AUST_DATUM).Value = AustrittsDatum
    wsH.Cells(NextRow, H_COL_NEUER_PAECHTER_ID).Value = NewMemberID ' ID des Nachpächters/Übernehmers
    wsH.Cells(NextRow, H_COL_GRUND).Value = ChangeReason
    wsH.Cells(NextRow, H_COL_SYSTEMZEIT).Value = Now
    
    ' Formatierung
    wsH.Cells(NextRow, H_COL_AUST_DATUM).NumberFormat = "dd.mm.yyyy"
    wsH.Cells(NextRow, H_COL_SYSTEMZEIT).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    Call ProtectSheet(wsH) ' Historie-Blatt wieder sperren

    
    ' --- 2. MITGLIEDERLISTE AKTUALISIEREN ---
    Call UnprotectSheet(wsM)
    
    ' *** 2a. AUSSCHEIDENDES MITGLIED AKTUALISIEREN (selectedRow) ***
    
    If ChangeReason = "Parzellenwechsel" And NewParzelleNr <> "" Then
        ' Parzellenwechsel: Parzellennummer des Mitglieds in der Mitgliederliste ändern
        wsM.Cells(selectedRow, M_COL_PARZELLE).Value = NewParzelleNr
        ' Das Pachtende muss hier *nicht* gesetzt werden, da das Mitglied aktiv bleibt
        
    ElseIf ChangeReason = "Austritt aus Parzelle" Or ChangeReason = "Austritt mit Pachtübernahme" Then
        
        ' Bei Austritt oder Übernahme muss das ausscheidende Mitglied immer auf "Ehemalig" gesetzt werden.
        
        ' Parzellennummer des Mitglieds in der Mitgliederliste leeren (für Austretenden)
        wsM.Cells(selectedRow, M_COL_PARZELLE).Value = ""
        
        ' WICHTIG: Pachtende setzen
        wsM.Cells(selectedRow, M_COL_PACHTENDE).Value = AustrittsDatum
        wsM.Cells(selectedRow, M_COL_PACHTENDE).NumberFormat = "dd.mm.yyyy"
        
        ' KORREKTUR: Austrittsstatus in Spalte Funtkion (M_COL_FUNKTION) setzen
        wsM.Cells(selectedRow, M_COL_FUNKTION).Value = AUSTRITT_STATUS_DISPLAY
    End If
    
    ' *** 2b. NEUES/ÜBERNEHMENDES MITGLIED AKTUALISIEREN (falls vorhanden) ***
    
    If ChangeReason = "Austritt mit Pachtübernahme" And NewMemberID <> "" Then
        
        ' Zeile des Übernehmers anhand der ID finden
        UebernehmerRow = FindeRowByMemberID(NewMemberID)
        
        If UebernehmerRow > 0 Then
            ' Funktion auf PÄCHTER_STATUS setzen
            wsM.Cells(UebernehmerRow, M_COL_FUNKTION).Value = PAECHTER_STATUS
            ' Hinweis: Die Parzelle des Übernehmers (die OldParzelle) wird NICHT geändert,
            ' da Sekundärmitglieder bereits die richtige Parzelle eingetragen haben.
            
            MsgBox "Pachtvertrag für Parzelle " & OldParzelle & " erfolgreich auf " & wsM.Cells(UebernehmerRow, M_COL_NACHNAME).Value & " übertragen.", vbInformation
            
        Else
            ' Fehlerfall
            MsgBox "FEHLER: MemberID des Übernehmers '" & NewMemberID & "' konnte in der Mitgliederliste nicht gefunden werden.", vbCritical
        End If
    End If
    
    ' Aktualisiere das Datum der letzten Änderung in D2
    Call AktualisiereDatenstand
    
    Call ProtectSheet(wsM)
    
    ' --- 3. ÜBERGREIFENDE AUFRÄUM- UND AKTUALISIERUNGS-LOGIK (UpdateAllDependencies) ---
    
    ' 3a) Named Range für Nachpächter-Dropdown aktualisieren (KORRIGIERT)
    Call AktualisiereNamedRange_MitgliederNamen
    
    ' 3b) Sortieren der Mitgliederliste (enthält Formatierung/Validation)
    Call Sortiere_Mitgliederliste_Nach_Parzelle
    
    ' 3c) Aktualisierung der nachgelagerten Abhängigkeiten (Banking, Zähler, etc.)
    On Error Resume Next
    
    ' Annahme: Diese Module existieren und die Prozeduren sind Public
    ' Wenn das Modul mod_Banking_Data nicht existiert, tritt hier kein Fehler auf (wegen On Error Resume Next)
    Call mod_Banking_Data.Aktualisiere_Parzellen_Mapping_Final
    Call mod_Banking_Data.Sortiere_Tabellen_Daten
    
    ' Wenn das Modul mod_ZaehlerLogik nicht existiert, tritt hier kein Fehler auf (wegen On Error Resume Next)
    Call mod_ZaehlerLogik.Ermittle_Kennzahlen_Mitgliederliste
    Call mod_ZaehlerLogik.ErzeugeParzellenUebersicht
    Call mod_ZaehlerLogik.AktualisiereZaehlerTabellenSpalteA

    ' 3d) Hauptformular aktualisieren
    If IsFormLoaded("frm_Mitgliederverwaltung") Then
        frm_Mitgliederverwaltung.RefreshMitgliederListe
    End If
    
    On Error GoTo 0
    
    ' Zeige nur die generische Meldung, wenn keine spezielle Übernahme-Meldung kam
    If ChangeReason <> "Austritt mit Pachtübernahme" Then
        MsgBox "Historien-Eintrag erfolgreich gespeichert und Mitgliederliste aktualisiert.", vbInformation
    End If
    
    Exit Sub
ErrorHandler:
    If Not wsM Is Nothing Then Call ProtectSheet(wsM)
    If Not wsH Is Nothing Then Call ProtectSheet(wsH)
    MsgBox "FEHLER BEI DER DATENVERARBEITUNG NACH FORMULARABSCHLUSS: " & Err.Description, vbCritical
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' ***************************************************************
' HILFSFUNKTION: Prüfen, ob eine UserForm geladen ist (KORRIGIERT)
' ***************************************************************
Private Function IsFormLoaded(ByVal FormName As String) As Boolean
    
    Dim f As Object
    
    ' Durchläuft die UserForms-Collection des VBA-Projekts (Korrekt für Excel)
    For Each f In VBA.UserForms
        ' Vergleicht den Namen (Case-insensitive)
        If StrComp(f.Name, FormName, vbTextCompare) = 0 Then
            IsFormLoaded = True
            Exit Function
        End If
    Next f
    
    IsFormLoaded = False
    
End Function


' ***************************************************************
' PRÜFFUNKTION: Ist das angegebene Mitglied der letzte aktive Pächter?
' ***************************************************************
Public Function CheckIfLastPaechter(ByVal PaeffelParzelle As String, ByVal MemberIDToExclude As String) As Boolean
    
    Dim wsM As Worksheet
    Dim lastRowM As Long
    Dim lRow As Long
    Dim PachterCount As Long
    Dim currentParzelle As String
    Dim currentMemberID As String
    Dim currentFunktion As String
    
    ' Standardmäßig annehmen, es gibt noch andere Pächter
    CheckIfLastPaechter = False
    PachterCount = 0
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    lastRowM = wsM.Cells(wsM.Rows.Count, M_COL_PARZELLE).End(xlUp).Row
    
    If lastRowM < M_START_ROW Then
        ' Wenn die Liste leer ist, kann es keinen letzten Pächter geben
        Exit Function
    End If
    
    For lRow = M_START_ROW To lastRowM
        currentParzelle = Trim(CStr(wsM.Cells(lRow, M_COL_PARZELLE).Value))
        currentMemberID = Trim(CStr(wsM.Cells(lRow, M_COL_MEMBER_ID).Value))
        currentFunktion = Trim(CStr(wsM.Cells(lRow, M_COL_FUNKTION).Value))
        
        ' 1. Prüfe, ob es die relevante Parzelle ist
        If UCase(currentParzelle) = UCase(PaeffelParzelle) Then
            
            ' 2. Prüfe, ob das Mitglied aktiv und ein Pächter ist
            If UCase(currentFunktion) = UCase(PAECHTER_STATUS) Then
            
                ' 3. Schließe das Mitglied aus, das gerade bearbeitet wird (es ist das potenziell austretende)
                If UCase(currentMemberID) <> UCase(MemberIDToExclude) Then
                    PachterCount = PachterCount + 1
                    ' Wenn wir einen weiteren Pächter gefunden haben, brechen wir die Schleife ab (Performance-Optimierung)
                    If PachterCount > 0 Then
                        CheckIfLastPaechter = False ' Es gibt noch einen anderen Pächter
                        Exit Function
                    End If
                End If
            End If
        End If
    Next lRow
    
    ' Wenn die Schleife durchgelaufen ist und PachterCount 0 ist:
    ' bedeutet Count 0, dass der MemberIDToExclude der einzige Pächter ist.
    CheckIfLastPaechter = True
    Exit Function
    
ErrorHandler:
    MsgBox "Fehler in CheckIfLastPaechter: " & Err.Description, vbCritical
    CheckIfLastPaechter = True ' Im Fehlerfall sicherheitshalber davon ausgehen, dass es der Letzte ist, um Abbruch zu erzwingen
End Function


' ***************************************************************
' HILFSFUNKTION: Sucht Sekundärmitglieder auf einer Parzelle
' ***************************************************************
' Gibt ein Array von Strings zurück (z.B. "Nachname, Vorname|MemberID")
Public Function GetSekundaerMitgliederAufParzelle(ByVal ParzelleNr As String) As Variant
    
    Dim wsM As Worksheet
    Dim lastRowM As Long
    Dim lRow As Long
    Dim SekundaerList() As String
    Dim i As Long: i = -1
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    lastRowM = wsM.Cells(wsM.Rows.Count, M_COL_PARZELLE).End(xlUp).Row
    
    If lastRowM < M_START_ROW Then
        GetSekundaerMitgliederAufParzelle = Array()
        Exit Function
    End If
    
    For lRow = M_START_ROW To lastRowM
        ' Prüfe Parzelle, Funktion und Pachtende
        If UCase(Trim(CStr(wsM.Cells(lRow, M_COL_PARZELLE).Value))) = UCase(ParzelleNr) And _
           UCase(Trim(CStr(wsM.Cells(lRow, M_COL_FUNKTION).Value))) = UCase(SEKUNDAER_STATUS) And _
           IsDate(wsM.Cells(lRow, M_COL_PACHTENDE).Value) = False Then ' Nur aktive Mitglieder
            
            i = i + 1
            ReDim Preserve SekundaerList(0 To i)
            
            ' Speichere Name und ID getrennt durch ein "|" (Pipe-Symbol)
            SekundaerList(i) = Trim(CStr(wsM.Cells(lRow, M_COL_NACHNAME).Value)) & ", " & _
                               Trim(CStr(wsM.Cells(lRow, M_COL_VORNAME).Value)) & "|" & _
                               Trim(CStr(wsM.Cells(lRow, M_COL_MEMBER_ID).Value))
        End If
    Next lRow
    
    If i >= 0 Then
        GetSekundaerMitgliederAufParzelle = SekundaerList
    Else
        GetSekundaerMitgliederAufParzelle = Array()
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Fehler in GetSekundaerMitgliederAufParzelle: " & Err.Description, vbCritical
    GetSekundaerMitgliederAufParzelle = Array()
End Function


' ***************************************************************
' NEUE PRÜFFUNKTION: Check_Vorstand_Eindeutigkeit
' Prüft, ob bereits ein Mitglied den Status "Vorstand" hat,
' unter Ausschluss des aktuell zu speichernden Mitglieds (anhand der MemberID).
' Wird von frm_Mitgliedsdaten zur Validierung vor dem Speichern aufgerufen.
' ***************************************************************
Public Function Check_Vorstand_Eindeutigkeit(ByVal CheckMemberID As String) As Boolean
    
    Dim wsM As Worksheet
    Dim lastRowM As Long
    Dim lRow As Long
    Dim currentMemberID As String
    Dim currentFunktion As String
    
    ' Standardmäßig annehmen, dass die Eindeutigkeit gegeben ist (True)
    Check_Vorstand_Eindeutigkeit = True
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    lastRowM = wsM.Cells(wsM.Rows.Count, M_COL_FUNKTION).End(xlUp).Row
    
    If lastRowM < M_START_ROW Then
        Exit Function ' Liste ist leer
    End If
    
    For lRow = M_START_ROW To lastRowM
        currentFunktion = Trim(CStr(wsM.Cells(lRow, M_COL_FUNKTION).Value))
        currentMemberID = Trim(CStr(wsM.Cells(lRow, M_COL_MEMBER_ID).Value))
        
        ' 1. Prüfen, ob die Funktion "Vorstand" ist (Gross-/Kleinschreibung ignorieren)
        If UCase(currentFunktion) = UCase(VORSTAND_STATUS) Then
            
            ' 2. Prüfen, ob dies NICHT die ID des aktuell zu speichernden/bearbeitenden Mitglieds ist
            If UCase(currentMemberID) <> UCase(CheckMemberID) Then
                ' Ein aktives Mitglied mit der Funktion "Vorstand" wurde gefunden, das nicht das aktuelle ist.
                Check_Vorstand_Eindeutigkeit = False
                Exit Function
            End If
        End If
    Next lRow
    
    Exit Function
    
ErrorHandler:
    MsgBox "Fehler in Check_Vorstand_Eindeutigkeit: " & Err.Description, vbCritical
    Check_Vorstand_Eindeutigkeit = False ' Im Fehlerfall lieber False zurückgeben (Eindeutigkeit nicht garantiert)
End Function




