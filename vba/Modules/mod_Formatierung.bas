Attribute VB_Name = "mod_Formatierung"
Option Explicit

' ***************************************************************
' MODUL: mod_Formatierung (ORCHESTRATOR)
' ZWECK: Formatierung und DropDown-Listen-Verwaltung - Haupteinstieg
' VERSION: 6.0 - 01.03.2026 (Modularisiert)
'
' SUB-MODULE:
'   mod_Format_Spalten     - Einzelspalten-Zebra, Verdichtung
'   mod_Format_Kategorie   - Kategorie-Tabelle J-P
'   mod_Format_EntityKey   - EntityKey-Tabelle R-X
'   mod_Format_Bankkonto   - Bankkonto-Blatt Formatierung
'   mod_Format_Protection  - Blattschutz, Sperren/Entsperren
'   mod_Format_Dropdowns   - DropDown-Listen AF/AG/AH
'
' VERBLEIBENDE FUNKTIONEN:
'   - Formatiere_Alle_Tabellen_Neu: Haupt-Orchestrator
'   - FormatiereBlattDaten: Daten-Blatt Orchestrator
'   - FormatKategorieTableComplete: Public Wrapper Kategorie
'   - FormatEntityKeyTableComplete: Public Wrapper EntityKey
'   - OnDatenChange: Worksheet_Change Dispatcher
'   - OnKategorieChange: Kategorie-Aenderung Validierung
'   - ZentriereAlleZellenVertikal: Vertikale Zentrierung
'   - FormatiereNeuesBlatt: Neues Blatt formatieren
' ***************************************************************

' ===============================================================
' Zentriert ALLE Zellen auf ALLEN Blaettern vertikal
' ===============================================================
Public Sub ZentriereAlleZellenVertikal()
    
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error Resume Next
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect PASSWORD:=PASSWORD
        ws.Cells.VerticalAlignment = xlCenter
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    Next ws
    
    On Error GoTo 0
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub


' ===============================================================
' v5.1: Stellt sicher, dass AutoFilter-Dropdowns auf allen
' relevanten Blaettern vorhanden und nutzbar sind.
' Wird von Workbook_Open aufgerufen.
' ===============================================================
Public Sub StelleAutoFilterBereit()
    
    Dim cfg As Variant
    ' Array: Blattname, Header-Zeile
    cfg = Array( _
        Array(WS_BANKKONTO, 27), _
        Array(WS_VEREINSKASSE, 26), _
        Array("Strom", 7), _
        Array("Wasser", 9), _
        Array(WS_MITGLIEDER, 5), _
        Array(WS_MITGLIEDER_HISTORIE, 3), _
        Array(WS_EINSTELLUNGEN, 3), _
        Array(WS_DATEN, 3))
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    
    Dim i As Long
    For i = LBound(cfg) To UBound(cfg)
        Dim shName As String
        shName = cfg(i)(0)
        Dim hRow As Long
        hRow = cfg(i)(1)
        
        Dim ws As Worksheet
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(shName)
        If ws Is Nothing Then GoTo NextSheet
        
        ws.Unprotect PASSWORD:=PASSWORD
        
        ' Letzte Spalte mit Daten in Header-Zeile ermitteln
        Dim lastCol As Long
        lastCol = ws.Cells(hRow, ws.Columns.count).End(xlToLeft).Column
        If lastCol < 1 Then lastCol = 1
        
        ' Letzte Zeile mit Daten ermitteln
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
        If lastRow <= hRow Then lastRow = hRow + 1
        
        ' Bestehenden AutoFilter entfernen
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        
        ' AutoFilter auf Header-Bereich aktivieren
        ws.Range(ws.Cells(hRow, 1), ws.Cells(lastRow, lastCol)).AutoFilter
        
        ' Blatt mit AllowFiltering schuetzen
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
        
NextSheet:
    Next i
    
    On Error GoTo 0
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub


' ===============================================================
' Wird aufgerufen wenn ein neues Blatt erstellt wird
' ===============================================================
Public Sub FormatiereNeuesBlatt(ByVal ws As Worksheet)
    
    On Error Resume Next
    
    ws.Unprotect PASSWORD:=PASSWORD
    ws.Cells.VerticalAlignment = xlCenter
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    On Error GoTo 0
    
End Sub

' ===============================================================
' HAUPTPROZEDUR: Formatiert ALLE relevanten Tabellen neu
' ===============================================================
Public Sub Formatiere_Alle_Tabellen_Neu()
    
    Dim wsD As Worksheet
    Dim wsBK As Worksheet
    Dim wsM As Worksheet
    Dim ws As Worksheet
    Dim lastRowBK As Long
    Dim euroFormat As String
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    euroFormat = "#,##0.00 " & ChrW(8364)
    
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect PASSWORD:=PASSWORD
        ws.Cells.VerticalAlignment = xlCenter
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        On Error GoTo ErrorHandler
    Next ws
    
    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo ErrorHandler
    
    If Not wsD Is Nothing Then
        On Error Resume Next
        wsD.Unprotect PASSWORD:=PASSWORD
        On Error GoTo ErrorHandler
        
        Call mod_Format_Spalten.FormatiereAlleDatenSpalten(wsD)
        Call mod_Format_Kategorie.FormatiereKategorieTabelle(wsD)
        Call mod_Format_EntityKey.FormatiereEntityKeyTabelleKomplett(wsD)
        Call mod_Format_Dropdowns.AktualisiereKategorieDropdownListen(wsD)
        Call mod_Format_Kategorie.SortiereKategorieTabelle(wsD)
        Call mod_Format_EntityKey.SortiereEntityKeyTabelle(wsD)
        Call mod_Format_Protection.EntspeerreEditierbareSpalten(wsD)
        
        wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo ErrorHandler
    
    If Not wsBK Is Nothing Then
        On Error Resume Next
        wsBK.Unprotect PASSWORD:=PASSWORD
        On Error GoTo ErrorHandler
        
        lastRowBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
        If lastRowBK < BK_START_ROW Then lastRowBK = BK_START_ROW
        
        With wsBK.Range(wsBK.Cells(BK_START_ROW, BK_COL_BEMERKUNG), _
                        wsBK.Cells(lastRowBK, BK_COL_BEMERKUNG))
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        
        wsBK.Rows(BK_START_ROW & ":" & lastRowBK).AutoFit
        
        wsBK.Range(wsBK.Cells(BK_START_ROW, BK_COL_BETRAG), _
                   wsBK.Cells(lastRowBK, BK_COL_BETRAG)).NumberFormat = euroFormat
        
        wsBK.Range(wsBK.Cells(BK_START_ROW, BK_COL_MITGL_BEITR), _
                   wsBK.Cells(lastRowBK, BK_COL_AUSZAHL_KASSE)).NumberFormat = euroFormat
        
        wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    
    On Error Resume Next
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    On Error GoTo ErrorHandler
    
    If Not wsM Is Nothing Then
        On Error Resume Next
        wsM.Unprotect PASSWORD:=PASSWORD
        wsM.Cells.VerticalAlignment = xlCenter
        wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        On Error GoTo ErrorHandler
    End If
    
    Call mod_Format_Protection.BlendeDatenSpaltenAus
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not wsD Is Nothing Then wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    If Not wsBK Is Nothing Then wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    If Not wsM Is Nothing Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Debug.Print "Fehler in Formatiere_Alle_Tabellen_Neu: " & Err.Description
End Sub

' ===============================================================
' HAUPTPROZEDUR: Formatiert das gesamte Daten-Blatt
' ===============================================================
Public Sub FormatiereBlattDaten()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    ws.Cells.VerticalAlignment = xlCenter
    
    Call mod_Format_Spalten.FormatiereAlleDatenSpalten(ws)
    Call mod_Format_Kategorie.FormatiereKategorieTabelle(ws)
    Call mod_Format_EntityKey.FormatiereEntityKeyTabelleKomplett(ws)
    Call mod_Format_Dropdowns.AktualisiereKategorieDropdownListen(ws)
    Call mod_Format_Kategorie.SortiereKategorieTabelle(ws)
    Call mod_Format_EntityKey.SortiereEntityKeyTabelle(ws)
    Call mod_Format_Protection.EntspeerreEditierbareSpalten(ws)
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Call mod_Format_Protection.BlendeDatenSpaltenAus
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Formatierung des Daten-Blatts abgeschlossen!" & vbCrLf & vbCrLf & _
           "- Alle Zellen vertikal zentriert" & vbCrLf & _
           "- Alle Spalten mit Zebra-Formatierung" & vbCrLf & _
           "- Kategorie-Tabelle formatiert und sortiert" & vbCrLf & _
           "- EntityKey-Tabelle formatiert und sortiert" & vbCrLf & _
           "- DropDown-Listen aktualisiert" & vbCrLf & _
           "- Editierbare Spalten und Eingabezeile entsperrt", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler bei der Formatierung: " & Err.Description, vbCritical
End Sub

' ===============================================================
' PUBLIC WRAPPER: Formatiert Kategorie-Tabelle komplett
' ===============================================================
Public Sub FormatKategorieTableComplete(ByRef ws As Worksheet)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    If Err.Number <> 0 Then
        Debug.Print "FormatKatComplete: Unprotect fehlgeschlagen: " & Err.Description
        Err.Clear
    End If
    
    ' Jeder Schritt einzeln abgesichert - Fehler in einem Schritt
    ' duerfen die folgenden Schritte NICHT blockieren
    Call mod_Format_Spalten.VerdichteSpalteOhneLuecken(ws, DATA_CAT_COL_KATEGORIE, DATA_CAT_COL_START, DATA_CAT_COL_END)
    If Err.Number <> 0 Then
        Debug.Print "FormatKatComplete: Verdichte fehlgeschlagen: " & Err.Description
        Err.Clear
    End If
    
    Call mod_Format_Kategorie.FormatiereKategorieTabelle(ws)
    If Err.Number <> 0 Then
        Debug.Print "FormatKatComplete: Formatiere fehlgeschlagen: " & Err.Description
        Err.Clear
    End If
    
    Call mod_Format_Kategorie.SortiereKategorieTabelle(ws)
    If Err.Number <> 0 Then
        Debug.Print "FormatKatComplete: Sortiere fehlgeschlagen: " & Err.Description
        Err.Clear
    End If
    
    Call mod_Format_Protection.EntspeerreEditierbareSpalten(ws)
    If Err.Number <> 0 Then
        Debug.Print "FormatKatComplete: Entsperre fehlgeschlagen: " & Err.Description
        Err.Clear
    End If
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    If Err.Number <> 0 Then
        Debug.Print "FormatKatComplete: Protect fehlgeschlagen: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
End Sub

' ===============================================================
' PUBLIC WRAPPER: Formatiert EntityKey-Tabelle komplett
' ===============================================================
Public Sub FormatEntityKeyTableComplete(ByRef ws As Worksheet)
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    If Err.Number <> 0 Then
        Debug.Print "FormatEKComplete: Unprotect fehlgeschlagen: " & Err.Description
        Err.Clear
    End If
    
    Call mod_Format_Spalten.VerdichteSpalteOhneLuecken(ws, EK_COL_IBAN, EK_COL_ENTITYKEY, EK_COL_DEBUG)
    If Err.Number <> 0 Then
        Debug.Print "FormatEKComplete: Verdichte fehlgeschlagen: " & Err.Description
        Err.Clear
    End If
    
    Call mod_Format_EntityKey.FormatiereEntityKeyTabelleKomplett(ws)
    If Err.Number <> 0 Then
        Debug.Print "FormatEKComplete: Formatiere fehlgeschlagen: " & Err.Description
        Err.Clear
    End If
    
    Call mod_Format_EntityKey.SortiereEntityKeyTabelle(ws)
    If Err.Number <> 0 Then
        Debug.Print "FormatEKComplete: Sortiere fehlgeschlagen: " & Err.Description
        Err.Clear
    End If
    
    Call mod_Format_Protection.EntspeerreEditierbareSpalten(ws)
    If Err.Number <> 0 Then
        Debug.Print "FormatEKComplete: Entsperre fehlgeschlagen: " & Err.Description
        Err.Clear
    End If
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    If Err.Number <> 0 Then
        Debug.Print "FormatEKComplete: Protect fehlgeschlagen: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
End Sub

' ===============================================================
' Validierung und Reaktion bei Aenderung in Kategorie-Tabelle
' ===============================================================
Public Sub OnKategorieChange(ByVal Target As Range)
    
    Dim ws As Worksheet
    Dim zeile As Long
    Dim kategorie As String
    Dim r As Long
    Dim lastRow As Long
    
    Set ws = Target.Worksheet
    zeile = Target.Row
    
    If zeile < DATA_START_ROW Then Exit Sub
    
    ' Spalte J oder K geaendert: Dropdown-Listen aktualisieren
    If Target.Column = DATA_CAT_COL_KATEGORIE Or Target.Column = DATA_CAT_COL_EINAUS Then
        Call mod_Format_Dropdowns.AktualisiereKategorieDropdownListen(ws)
    End If
    
    ' Spalte K (E/A) geaendert: Konsistenzpruefung
    If Target.Column = DATA_CAT_COL_EINAUS Then
        Dim einAus As String
        einAus = UCase(Trim(Target.value))
        
        kategorie = Trim(ws.Cells(zeile, DATA_CAT_COL_KATEGORIE).value)
        
        If kategorie <> "" And einAus <> "" Then
            lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
            
            Dim bestehenderTyp As String
            bestehenderTyp = ""
            
            For r = DATA_START_ROW To lastRow
                If r <> zeile Then
                    If StrComp(Trim(ws.Cells(r, DATA_CAT_COL_KATEGORIE).value), kategorie, vbTextCompare) = 0 Then
                        Dim vorhandenerEA As String
                        vorhandenerEA = UCase(Trim(ws.Cells(r, DATA_CAT_COL_EINAUS).value))
                        If vorhandenerEA <> "" Then
                            bestehenderTyp = vorhandenerEA
                            Exit For
                        End If
                    End If
                End If
            Next r
            
            If bestehenderTyp <> "" And bestehenderTyp <> einAus Then
                Application.EnableEvents = False
                Target.value = bestehenderTyp
                Application.EnableEvents = True
                
                Dim typBeschreibung As String
                If bestehenderTyp = "E" Then
                    typBeschreibung = "Einnahme (E)"
                Else
                    typBeschreibung = "Ausgabe (A)"
                End If
                
                MsgBox "Die Kategorie """ & kategorie & """ ist bereits als " & typBeschreibung & " eingetragen." & vbCrLf & vbCrLf & _
                       "Bei gleicher Kategorie kann nur einheitlich zwischen Einnahme oder Ausgabe gew" & ChrW(228) & "hlt werden." & vbCrLf & _
                       "Gemischte Angaben sind nicht gestattet." & vbCrLf & vbCrLf & _
                       "Der Wert wurde automatisch auf """ & bestehenderTyp & """ korrigiert.", _
                       vbInformation, "Kategorie-Konsistenz"
            End If
        End If
        
        einAus = UCase(Trim(Target.value))
        Call mod_Format_Kategorie.SetzeZielspalteDropdown(ws, zeile, einAus)
    End If
    
    ' Spalte L (Keyword) geaendert: Duplikat-Pruefung
    If Target.Column = DATA_CAT_COL_KEYWORD Then
        Dim neuesKeyword As String
        neuesKeyword = Trim(Target.value)
        
        If neuesKeyword <> "" Then
            kategorie = Trim(ws.Cells(zeile, DATA_CAT_COL_KATEGORIE).value)
            
            If kategorie <> "" Then
                lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
                
                For r = DATA_START_ROW To lastRow
                    If r <> zeile Then
                        If StrComp(Trim(ws.Cells(r, DATA_CAT_COL_KATEGORIE).value), kategorie, vbTextCompare) = 0 Then
                            If StrComp(Trim(ws.Cells(r, DATA_CAT_COL_KEYWORD).value), neuesKeyword, vbTextCompare) = 0 Then
                                MsgBox "F" & ChrW(252) & "r die Kategorie """ & kategorie & """ gibt es bereits das gew" & ChrW(228) & "hlte Schl" & ChrW(252) & "sselwort """ & neuesKeyword & """.", _
                                       vbExclamation, "Doppeltes Schl" & ChrW(252) & "sselwort"
                                Exit For
                            End If
                        End If
                    End If
                Next r
            End If
        End If
    End If
    
End Sub

' ===============================================================
' WORKSHEET_CHANGE HANDLER fuer dynamische Formatierung
' ===============================================================
Public Sub OnDatenChange(ByVal Target As Range, ByVal ws As Worksheet)
    
    On Error GoTo ErrorHandler
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Call mod_Format_Spalten.FormatiereAlleDatenSpalten(ws)
    
    Call mod_Format_Kategorie.FormatiereKategorieTabelle(ws)
    Call mod_Format_Kategorie.SortiereKategorieTabelle(ws)
    
    Call mod_Format_EntityKey.FormatiereEntityKeyTabelleKomplett(ws)
    Call mod_Format_EntityKey.SortiereEntityKeyTabelle(ws)
    
    If Not Intersect(Target, ws.Range("J:P")) Is Nothing Then
        Call OnKategorieChange(Target)
    End If
    
    Call mod_Format_Protection.BlendeDatenSpaltenAus
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Debug.Print "Fehler in OnDatenChange: " & Err.Description
    End If
    
End Sub

















































































