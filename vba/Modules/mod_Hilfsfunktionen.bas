Attribute VB_Name = "mod_Hilfsfunktionen"
Option Explicit

' **********************************************************
' MODUL: mod_Hilfsfunktionen
' ZWECK: Generische Hilfsroutinen (Named Ranges, Listgenerierung)
' **********************************************************

Private Const TEMP_WS_NAME As String = "TEMP_LISTEN"

' **********************************************************
' PROZEDUR: AktualisiereNamedRange_MitgliederNamen
' Erstellt oder aktualisiert einen benannten Bereich
' mit den Namen aller aktiven Mitglieder.
' WICHTIG: Das temporäre Worksheet wird am Ende IMMER gelöscht!
' **********************************************************
Public Sub AktualisiereNamedRange_MitgliederNamen()
    
    Dim wsM As Worksheet
    Dim wsTemp As Worksheet
    Dim lastRow As Long
    Dim tempRow As Long
    Dim rngTarget As Range
    Dim wasProtected As Boolean
    Dim arrNames() As Variant
    Dim nameCount As Long
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    
    ' 1. Temporäres Arbeitsblatt erstellen/finden
    On Error Resume Next
    Set wsTemp = ThisWorkbook.Worksheets(TEMP_WS_NAME)
    On Error GoTo ErrorHandler
    
    If wsTemp Is Nothing Then
        ' Deaktiviere DisplayAlerts um Warnungen zu unterdrücken
        Application.DisplayAlerts = False
        Set wsTemp = ThisWorkbook.Worksheets.Add(After:=wsM)
        wsTemp.Name = TEMP_WS_NAME
        ' Verstecke das Worksheet (optional, für zusätzliche Sicherheit)
        wsTemp.Visible = xlSheetVeryHidden
        Application.DisplayAlerts = True
    Else
        ' Vorherige Daten löschen
        wsTemp.Cells.Clear
    End If
    
    ' 2. Daten kopieren und filtern (Nur aktive Mitglieder)
    wasProtected = wsM.ProtectContents
    If wasProtected Then wsM.Unprotect PASSWORD:=PASSWORD
    
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    If lastRow >= M_START_ROW Then
        
        ' Filterbereich definieren (Header bis letzte Zeile)
        wsM.Range(wsM.Cells(M_HEADER_ROW, 1), wsM.Cells(lastRow, M_COL_FUNKTION)).AutoFilter
        
        ' Filtern: Pachtende (M_COL_PACHTENDE) muss leer sein (Aktives Mitglied)
        wsM.Range(wsM.Cells(M_HEADER_ROW, 1), wsM.Cells(lastRow, M_COL_PACHTENDE)).AutoFilter _
             Field:=M_COL_PACHTENDE, Criteria1:="="
        
        tempRow = 1
        ' Kopiere die Spalten: Nachname (5), Vorname (6), Parzelle (2)
        Dim copyCols As Variant
        copyCols = Array(M_COL_NACHNAME, M_COL_VORNAME, M_COL_PARZELLE)
        
        For i = LBound(copyCols) To UBound(copyCols)
            wsM.Columns(copyCols(i)).SpecialCells(xlCellTypeVisible).Copy
            ' Fügen Sie in die temporäre Tabelle in Spalten A, B, C ein
            wsTemp.Cells(tempRow, i + 1).PasteSpecial xlPasteValues
        Next i
        
        Application.CutCopyMode = False
        wsM.AutoFilterMode = False ' Filter aufheben
        
        ' 3. Kombinierte Namen-Liste erstellen (Nachname, Vorname)
        tempRow = wsTemp.Cells(wsTemp.Rows.count, 1).End(xlUp).Row
        
        If tempRow > 1 Then ' Zeile 1 enthält die Header/Erste Zeile des kopierten Bereichs
            For i = 2 To tempRow
                ' Spalte D: Nachname, Vorname (wird im Dropdown angezeigt)
                wsTemp.Cells(i, 4).value = wsTemp.Cells(i, 1).value & ", " & wsTemp.Cells(i, 2).value
            Next i
            
            ' 4. Benannten Bereich erstellen/aktualisieren (Spalte D, ab Zeile 2)
            Set rngTarget = wsTemp.Range(wsTemp.Cells(2, 4), wsTemp.Cells(tempRow, 4))
            
            ' Bestehenden benannten Bereich löschen
            On Error Resume Next
            ThisWorkbook.Names("rng_MitgliederNamen").Delete
            On Error GoTo ErrorHandler
            
            ' Neuen benannten Bereich definieren
            ThisWorkbook.Names.Add Name:="rng_MitgliederNamen", RefersTo:=rngTarget
        End If
    End If
    
    ' *** WICHTIG: Temporäres Worksheet IMMER löschen! ***
    Call LoescheTempWorksheet
    
CleanUp:
    Application.ScreenUpdating = True
    If Not wsM Is Nothing Then
        If wsM.AutoFilterMode Then wsM.AutoFilterMode = False
        If wasProtected Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler in AktualisiereNamedRange_MitgliederNamen: " & Err.Description, vbCritical
    ' Versuche trotz Fehler das Temp-Worksheet zu löschen
    Call LoescheTempWorksheet
    Resume CleanUp

End Sub

' **********************************************************
' PROZEDUR: LoescheTempWorksheet
' Löscht das temporäre Worksheet sicher
' **********************************************************
Private Sub LoescheTempWorksheet()
    Dim wsTemp As Worksheet
    
    On Error Resume Next
    Set wsTemp = ThisWorkbook.Worksheets(TEMP_WS_NAME)
    
    If Not wsTemp Is Nothing Then
        Application.DisplayAlerts = False
        wsTemp.Delete
        Application.DisplayAlerts = True
    End If
    
    On Error GoTo 0
End Sub

' **********************************************************
' PROZEDUR: BereinigeTempWorksheets
' Öffentliche Prozedur zum Bereinigen aller temporären Worksheets
' Kann manuell oder beim Öffnen der Arbeitsmappe aufgerufen werden
' **********************************************************
Public Sub BereinigeTempWorksheets()
    Dim ws As Worksheet
    Dim wsToDelete As Collection
    Dim tempName As Variant
    
    Set wsToDelete = New Collection
    
    ' Sammle alle Worksheets die "TEMP" im Namen haben
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "TEMP", vbTextCompare) > 0 Then
            wsToDelete.Add ws.Name
        End If
    Next ws
    
    ' Lösche gesammelte Worksheets
    Application.DisplayAlerts = False
    For Each tempName In wsToDelete
        On Error Resume Next
        ThisWorkbook.Worksheets(CStr(tempName)).Delete
        On Error GoTo 0
    Next tempName
    Application.DisplayAlerts = True
End Sub

' ***************************************************************
' HILFSFUNKTION: Prüfen, ob eine UserForm geladen ist (KORRIGIERT FÜR EXCEL)
' ***************************************************************
Private Function IsFormLoaded(ByVal FormName As String) As Boolean
    
    Dim i As Long
    
    ' Gehe alle geladenen UserForms in der VBA-Collection durch
    For i = 0 To VBA.UserForms.count - 1
        ' Vergleiche den Klassennamen des geladenen Formulars mit dem gesuchten Namen
        If StrComp(VBA.UserForms.item(i).Name, FormName, vbTextCompare) = 0 Then
            IsFormLoaded = True ' Formular gefunden und ist geladen
            Exit Function
        End If
    Next i
    
    IsFormLoaded = False ' Formular nicht gefunden
    
End Function

