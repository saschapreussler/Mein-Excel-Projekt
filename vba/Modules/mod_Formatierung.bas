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
'   - OnKategorieChange: Kategorie-änderung Validierung
'   - ZentriereAlleZellenVertikal: Vertikale Zentrierung
'   - FormatiereNeuesBlatt: Neues Blatt formatieren
' ***************************************************************

' ===============================================================
' Zentriert ALLE Zellen auf ALLEN Blättern vertikal
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
' relevanten Blättern vorhanden und nutzbar sind.
' Wird von Workbook_Open aufgerufen.
' ===============================================================
Public Sub StelleAutoFilterBereit()
    
    Dim cfg As Variant
    ' Array: Blattname, Header-Zeile
    ' HINWEIS: WS_BANKKONTO bewusst NICHT hier - das macht
    '          mod_Banking_Format.Aktiviere_BankkontoFilter dediziert
    '          (mit Header-Entsperren + Unmerge), weil dieses Blatt
    '          mit dem generischen Code Probleme hatte.
    cfg = Array( _
        Array(WS_VEREINSKASSE, 26), _
        Array("Strom", 9), _
        Array("Wasser", 11), _
        Array(WS_MITGLIEDER, 5), _
        Array(WS_MITGLIEDER_HISTORIE, 3), _
        Array(WS_EINSTELLUNGEN, 22), _
        Array(WS_DATEN, 3), _
        Array("Dashboard Mitgliederzahlungen", 10), _
        Array(WS_UEBERSICHT(), 3))
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim i As Long
    For i = LBound(cfg) To UBound(cfg)
        Dim shName As String
        shName = cfg(i)(0)
        Dim hRow As Long
        hRow = cfg(i)(1)
        
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(shName)
        Err.Clear
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextSheet
        
        On Error Resume Next
        ws.Unprotect PASSWORD:=PASSWORD
        Err.Clear
        On Error GoTo 0
        
        ' Header-Zeile validieren: muss mindestens 1 nicht-leere Zelle haben
        Dim headerCheck As String
        headerCheck = ""
        On Error Resume Next
        headerCheck = Trim(CStr(ws.Cells(hRow, 1).value))
        On Error GoTo 0
        
        ' Letzte Spalte mit NICHT-LEEREM Header ermitteln
        Dim lastCol As Long
        lastCol = 1
        Dim c As Long
        On Error Resume Next
        For c = 1 To ws.Cells(hRow, ws.Columns.count).End(xlToLeft).Column
            If Trim(CStr(ws.Cells(hRow, c).value)) <> "" Then lastCol = c
        Next c
        Err.Clear
        On Error GoTo 0
        If lastCol < 1 Then lastCol = 1
        
        ' Letzte Zeile mit Daten ermitteln
        Dim lastRow As Long
        lastRow = hRow + 1
        On Error Resume Next
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
        Err.Clear
        On Error GoTo 0
        If lastRow <= hRow Then lastRow = hRow + 1
        
        ' Bestehenden AutoFilter entfernen
        On Error Resume Next
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        Err.Clear
        On Error GoTo 0
        
        ' Header-Zellen entsperren damit Filter trotz Blattschutz nutzbar sind
        Dim hc As Long
        On Error Resume Next
        For hc = 1 To lastCol
            ws.Cells(hRow, hc).Locked = False
        Next hc
        Err.Clear
        On Error GoTo 0
        
        ' AutoFilter auf Header-Bereich aktivieren (nur belegte Spalten)
        On Error Resume Next
        ws.Range(ws.Cells(hRow, 1), ws.Cells(lastRow, lastCol)).AutoFilter
        If Err.Number <> 0 Then
            Debug.Print "[AutoFilter] FEHLER auf " & shName & " (Zeile " & hRow & "): " & Err.Description
            Err.Clear
        Else
            Debug.Print "[AutoFilter] OK: " & shName & " (Zeile " & hRow & ", Spalten 1-" & lastCol & ")"
        End If
        On Error GoTo 0
        
        ' Blatt mit AllowFiltering schützen
        On Error Resume Next
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
        Err.Clear
        On Error GoTo 0
        
NextSheet:
    Next i
    
    ' Zurück zur Startseite navigieren
    Dim wsStart As Worksheet
    Set wsStart = Nothing
    On Error Resume Next
    Set wsStart = ThisWorkbook.Worksheets(WS_STARTMENUE())
    On Error GoTo 0
    If Not wsStart Is Nothing Then
        wsStart.Activate
        wsStart.Range("A1").Select
    End If
    
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
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    
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
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
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
        
        wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
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
        
        wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    
    On Error Resume Next
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    On Error GoTo ErrorHandler
    
    If Not wsM Is Nothing Then
        On Error Resume Next
        wsM.Unprotect PASSWORD:=PASSWORD
        wsM.Cells.VerticalAlignment = xlCenter
        wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
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
    If Not wsD Is Nothing Then wsD.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    If Not wsBK Is Nothing Then wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    If Not wsM Is Nothing Then wsM.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
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
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    
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
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
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
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
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
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    If Err.Number <> 0 Then
        Debug.Print "FormatEKComplete: Protect fehlgeschlagen: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
End Sub

' ===============================================================
' Validierung und Reaktion bei änderung in Kategorie-Tabelle
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
    
    ' Spalte J oder K geändert: Dropdown-Listen aktualisieren
    If Target.Column = DATA_CAT_COL_KATEGORIE Or Target.Column = DATA_CAT_COL_EINAUS Then
        Call mod_Format_Dropdowns.AktualisiereKategorieDropdownListen(ws)
    End If
    
    ' Spalte K (E/A) geändert: Konsistenzprüfung
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
    
    ' Spalte L (Keyword) geändert: Duplikat-Prüfung
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
' WORKSHEET_CHANGE HANDLER für dynamische Formatierung
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


' ===============================================================
' FormatiereMitgliederlisteKomplett
' ---------------------------------------------------------------
' Sorgt für ein durchgehend konsistentes Layout der
' Mitgliederliste, immer wenn sich Daten Ändern. Wird gerufen von:
'   - Workbook_Open
'   - Worksheet_Change auf Mitgliederliste (Tabelle7)
'   - mod_Mitglieder_UI nach Speichern / Löschen / Sortieren
'
' Setzt:
'   - Vertikale + horizontale Zentrierung im Datenbereich
'   - Duenne durchgaengige Rahmen (innen + aussen)
'   - Zebra-Streifen (weiss / hellgrau alternierend)
'   - AutoFit auf Datenspalten + sinnvolle Mindestbreite
'   - Datumsformat für Geburtstag / Pacht-Spalten
'
' Wird intern stets unter ScreenUpdating=False und EnableEvents=False
' ausgeführt, damit es keine Reentrancy in Worksheet_Change gibt.
' ===============================================================
Public Sub FormatiereMitgliederlisteKomplett()

    Dim wsM As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim datenBereich As Range
    Dim headerBereich As Range
    Dim warEnableEvents As Boolean
    Dim warScreenUpdating As Boolean

    On Error GoTo CleanExit

    Set wsM = Nothing
    On Error Resume Next
    Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    On Error GoTo CleanExit
    If wsM Is Nothing Then Exit Sub

    warEnableEvents = Application.EnableEvents
    warScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    On Error Resume Next
    wsM.Unprotect PASSWORD:=PASSWORD
    On Error GoTo CleanExit

    ' Letzte belegte Datenzeile + Spalte ermitteln
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    If lastRow < M_START_ROW Then lastRow = M_START_ROW
    lastCol = M_COL_FUNKTION   ' bis Spalte O (15) - alles dahinter ist Verwaltungs-/Versteckt-Bereich

    Set headerBereich = wsM.Range(wsM.Cells(M_HEADER_ROW, 1), _
                                   wsM.Cells(M_HEADER_ROW, lastCol))
    Set datenBereich = wsM.Range(wsM.Cells(M_START_ROW, 1), _
                                  wsM.Cells(lastRow, lastCol))

    ' Ausrichtung: zentriert horizontal + vertikal
    datenBereich.HorizontalAlignment = xlCenter
    datenBereich.VerticalAlignment = xlCenter
    headerBereich.HorizontalAlignment = xlCenter
    headerBereich.VerticalAlignment = xlCenter

    ' Rahmen entfernen, dann neu setzen (innen + aussen, duenn, schwarz)
    Dim rngGesamt As Range
    Set rngGesamt = wsM.Range(wsM.Cells(M_HEADER_ROW, 1), _
                               wsM.Cells(lastRow, lastCol))
    rngGesamt.Borders.LineStyle = xlNone

    Dim borderIds As Variant
    borderIds = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
                      xlInsideVertical, xlInsideHorizontal)
    Dim i As Long
    For i = LBound(borderIds) To UBound(borderIds)
        With rngGesamt.Borders(borderIds(i))
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = RGB(150, 150, 150)
        End With
    Next i

    ' Zebra-Streifen: ungerade Datenzeile = weiss, gerade = hellgrau
    For r = M_START_ROW To lastRow
        Dim rowRange As Range
        Set rowRange = wsM.Range(wsM.Cells(r, 1), wsM.Cells(r, lastCol))
        If ((r - M_START_ROW) Mod 2) = 0 Then
            rowRange.Interior.color = RGB(255, 255, 255)
        Else
            rowRange.Interior.color = RGB(242, 242, 242)
        End If
    Next r

    ' Datumsformate
    On Error Resume Next
    wsM.Range(wsM.Cells(M_START_ROW, M_COL_GEBURTSTAG), _
              wsM.Cells(lastRow, M_COL_GEBURTSTAG)).NumberFormat = "DD.MM.YYYY"
    wsM.Range(wsM.Cells(M_START_ROW, M_COL_PACHTANFANG), _
              wsM.Cells(lastRow, M_COL_PACHTENDE)).NumberFormat = "DD.MM.YYYY"
    On Error GoTo CleanExit

    ' Spaltenbreiten: AutoFit, dann sinnvolle Mindestbreite
    Dim c As Long
    wsM.Range(wsM.Cells(M_HEADER_ROW, 1), wsM.Cells(lastRow, lastCol)) _
        .Columns.AutoFit
    For c = 1 To lastCol
        If wsM.Columns(c).ColumnWidth < 8 Then
            wsM.Columns(c).ColumnWidth = 8
        End If
    Next c

    ' Zeilenhöhe AutoFit
    On Error Resume Next
    wsM.Rows(M_HEADER_ROW & ":" & lastRow).AutoFit
    On Error GoTo CleanExit

CleanExit:
    On Error Resume Next
    wsM.Protect PASSWORD:=PASSWORD, _
                UserInterfaceOnly:=True, _
                AllowFiltering:=True, _
                AllowSorting:=True
    Application.EnableEvents = warEnableEvents
    Application.ScreenUpdating = warScreenUpdating
    On Error GoTo 0
End Sub













































































































































