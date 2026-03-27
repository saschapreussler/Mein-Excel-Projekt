Attribute VB_Name = "mod_Uebersicht_Dashboard"
Option Explicit

' ***************************************************************
' MODUL: mod_Uebersicht_Dashboard
' VERSION: 1.0 - 17.06.2025
' ZWECK: Generiert visuelles Dashboard auf neuem Blatt
'        - KPI-Karten: Mitglieder, Eingegangen, Offen, Saeumnis
'        - Zahlungsmatrix: Parzelle x Kategorie (farbcodiert)
'        - Verzugsdetail: Offene Zahlungen nach Schwere sortiert
' DATENQUELLEN: mod_Uebersicht_Daten, mod_Zahlungspruefung
' ***************************************************************

' ============================================================
'  FARBEN (initialisiert via InitFarben, da RGB() in Const
'  nicht erlaubt und manuelle Long-Berechnung fehleranfaellig)
' ============================================================
Private m_CLR_NAVY As Long
Private m_CLR_HEADER_BG As Long
Private m_CLR_WEISS As Long
Private m_CLR_KPI_BLAU As Long
Private m_CLR_KPI_GRUEN As Long
Private m_CLR_KPI_ROT As Long
Private m_CLR_KPI_ORANGE As Long
Private m_CLR_ZELLE_GRUEN As Long
Private m_CLR_ZELLE_GELB As Long
Private m_CLR_ZELLE_ROT As Long
Private m_CLR_ZELLE_GRAU As Long
Private m_CLR_TEXT_GRUEN As Long
Private m_CLR_TEXT_DUNKELROT As Long
Private m_FarbenInit As Boolean

' Layout
Private Const TITEL_ROW As Long = 2
Private Const KPI_LABEL_ROW As Long = 5
Private Const KPI_WERT_ROW As Long = 6
Private Const KPI_DETAIL_ROW As Long = 7
Private Const MATRIX_HEADER_ROW As Long = 10
Private Const MATRIX_START_ROW As Long = 11


' ============================================================
'  TYPES (muessen VOR allen Sub/Function stehen!)
' ============================================================
Public Type ParzelleInfo
    parzNr As Long
    mitgliedName As String
    entityKeys As String      ' komma-getrennt falls mehrere
    roles As String
End Type

Public Type VerzugEintrag
    parzNr As Long
    mitglied As String
    kategorie As String
    monatNr As Long
    monatText As String
    soll As Double
    ist As Double
    differenz As Double
    saeumnis As Double
    tageVerzug As Long
    bemerkung As String
End Type


' ============================================================
'  FARB-INITIALISIERUNG
' ============================================================
Private Sub InitFarben()
    If m_FarbenInit Then Exit Sub
    m_CLR_NAVY = RGB(23, 37, 84)
    m_CLR_HEADER_BG = RGB(68, 84, 106)
    m_CLR_WEISS = RGB(255, 255, 255)
    m_CLR_KPI_BLAU = RGB(41, 128, 185)
    m_CLR_KPI_GRUEN = RGB(39, 174, 96)
    m_CLR_KPI_ROT = RGB(231, 76, 60)
    m_CLR_KPI_ORANGE = RGB(243, 156, 18)
    m_CLR_ZELLE_GRUEN = RGB(198, 239, 206)
    m_CLR_ZELLE_GELB = RGB(255, 230, 153)
    m_CLR_ZELLE_ROT = RGB(255, 199, 206)
    m_CLR_ZELLE_GRAU = RGB(242, 242, 242)
    m_CLR_TEXT_GRUEN = RGB(0, 128, 80)
    m_CLR_TEXT_DUNKELROT = RGB(192, 0, 0)
    m_FarbenInit = True
End Sub
' ============================================================
Public Sub GeneriereUebersichtNeu(Optional ByVal stummModus As Boolean = False)
    
    On Error GoTo ErrorHandler
    
    ' Farben initialisieren
    Call InitFarben
    
    Dim startTime As Double
    startTime = Timer
    
    ' --- 1. Jahr ermitteln ---
    Dim jahr As Long
    jahr = ErmittleDashboardJahr(stummModus)
    If jahr = 0 Then Exit Sub
    
    ' --- 2. Kategorien und Mitglieder laden ---
    Dim kategorien() As UebKategorie
    Dim anzKat As Long
    Call mod_Uebersicht_Daten.LadeKategorienAusEinstellungen(kategorien, anzKat)
    
    If anzKat = 0 Then
        If Not stummModus Then
            MsgBox "Keine Kategorien im Einstellungen-Blatt gefunden!", vbCritical, "Dashboard"
        End If
        Exit Sub
    End If
    
    Dim wsDaten As Worksheet
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo ErrorHandler
    If wsDaten Is Nothing Then
        If Not stummModus Then
            MsgBox "Blatt 'Daten' nicht gefunden!", vbCritical, "Dashboard"
        End If
        Exit Sub
    End If
    
    Dim mitglieder As Collection
    Set mitglieder = mod_Uebersicht_Daten.HoleAktiveMitglieder(wsDaten)
    
    ' --- 3. Parzellen gruppieren ---
    Dim parzellen() As ParzelleInfo
    Dim anzParz As Long
    Call GruppiereParzellen(mitglieder, parzellen, anzParz)
    
    If anzParz = 0 Then
        If Not stummModus Then
            MsgBox "Keine aktiven Mitglieder gefunden!", vbExclamation, "Dashboard"
        End If
        Exit Sub
    End If
    
    ' --- 4. Sheet erstellen / leeren ---
    Dim wsDash As Worksheet
    Set wsDash = HoleOderErstelleSheet()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    wsDash.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    wsDash.Cells.Clear
    
    ' Hintergrund weiss
    wsDash.Cells.Interior.color = m_CLR_WEISS
    
    ' --- 5. Einstellungen-Cache laden ---
    Call mod_Zahlungspruefung.LadeEinstellungenCacheZP
    Call mod_Zahlungspruefung.InitialisiereNachDezemberCacheZP(jahr)
    Call mod_Uebersicht_Daten.BefuelleVorjahrSpeicher(jahr - 1)
    
    ' --- 6. Titel schreiben ---
    Call SchreibeTitel(wsDash, jahr, anzKat)
    
    ' --- 7. Daten sammeln und Matrix + Verzug schreiben ---
    Dim matrixEndRow As Long
    Dim kpiSummeIst As Double
    Dim kpiSummeSoll As Double
    Dim kpiSummeSaeumnis As Double
    Dim kpiAnzahlOffen As Long
    Dim kpiAnzahlBezahlt As Long
    
    Dim verzugListe() As VerzugEintrag
    Dim anzVerzug As Long
    anzVerzug = 0
    ReDim verzugListe(0 To 499)
    
    Call SchreibeMatrixMitDaten(wsDash, jahr, kategorien, anzKat, _
                                parzellen, anzParz, mitglieder, _
                                matrixEndRow, _
                                kpiSummeIst, kpiSummeSoll, kpiSummeSaeumnis, _
                                kpiAnzahlOffen, kpiAnzahlBezahlt, _
                                verzugListe, anzVerzug)
    
    ' --- 8. KPI-Karten schreiben ---
    Call SchreibeKPI(wsDash, anzParz, kpiSummeIst, kpiSummeSoll, _
                     kpiSummeSaeumnis, kpiAnzahlOffen, kpiAnzahlBezahlt)
    
    ' --- 9. Verzugsdetail schreiben ---
    Dim verzugEndRow As Long
    If anzVerzug > 0 Then
        ReDim Preserve verzugListe(0 To anzVerzug - 1)
        Call SortiereVerzug(verzugListe, anzVerzug)
        Call SchreibeVerzugsdetail(wsDash, matrixEndRow + 3, verzugListe, _
                                   anzVerzug, verzugEndRow)
    End If
    
    ' --- 10. Cache freigeben ---
    Call mod_Zahlungspruefung.EntladeEinstellungenCacheZP
    
    ' --- 11. Spaltenbreiten ---
    Call PasseSpaltenAn(wsDash, anzKat)
    
    ' --- 12. Blatt schuetzen ---
    On Error Resume Next
    wsDash.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo ErrorHandler
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    If Not stummModus Then
        wsDash.Activate
        wsDash.Range("A1").Select
    End If
    
    Dim endTime As Double
    endTime = Timer
    
    Debug.Print "[Dashboard] Erfolgreich: " & anzParz & " Parzellen, " & _
                anzKat & " Kategorien in " & Format(endTime - startTime, "0.00") & "s"
    
    If Not stummModus Then
        MsgBox "Dashboard erfolgreich generiert!" & vbLf & vbLf & _
               "Parzellen: " & anzParz & vbLf & _
               "Kategorien: " & anzKat & vbLf & _
               IIf(anzVerzug > 0, "Offene Posten: " & anzVerzug & vbLf, "") & _
               "Dauer: " & Format(endTime - startTime, "0.00") & " Sekunden", _
               vbInformation, "Dashboard"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    On Error Resume Next
    Call mod_Zahlungspruefung.EntladeEinstellungenCacheZP
    If Not wsDash Is Nothing Then
        wsDash.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End If
    On Error GoTo 0
    
    MsgBox "Fehler beim Erstellen des Dashboards:" & vbLf & vbLf & _
           Err.Description, vbCritical, "Dashboard-Fehler"
    Debug.Print "[Dashboard] FEHLER: " & Err.Number & " - " & Err.Description
    
End Sub


' ============================================================
'  JAHR ERMITTELN (gleiche Logik wie GeneriereUebersicht)
' ============================================================
Private Function ErmittleDashboardJahr(Optional ByVal stummModus As Boolean = False) As Long
    
    Dim jahrF1 As Long
    jahrF1 = 0
    
    Dim wsStart As Worksheet
    On Error Resume Next
    Set wsStart = ThisWorkbook.Worksheets("Startmen" & ChrW(252))
    On Error GoTo 0
    
    If Not wsStart Is Nothing Then
        If IsNumeric(wsStart.Range("F1").value) Then
            jahrF1 = CLng(wsStart.Range("F1").value)
        End If
    End If
    
    Dim jahrBK As Long
    jahrBK = mod_Uebersicht_Daten.ErmittleJahrAusBankkonto()
    
    If jahrF1 > 0 And jahrBK > 0 Then
        If jahrF1 = jahrBK Then
            ErmittleDashboardJahr = jahrF1
        Else
            If Not stummModus Then
                Dim antwort As VbMsgBoxResult
                antwort = MsgBox("Startmen" & ChrW(252) & "!F1 = " & jahrF1 & _
                                 ", Bankkonto = " & jahrBK & "." & vbLf & vbLf & _
                                 "Dashboard f" & ChrW(252) & "r " & jahrF1 & _
                                 " (Startmen" & ChrW(252) & ") erstellen?", _
                                 vbQuestion + vbYesNo, "Abrechnungsjahr")
                If antwort = vbYes Then
                    ErmittleDashboardJahr = jahrF1
                Else
                    ErmittleDashboardJahr = jahrBK
                End If
            Else
                ' stummModus: F1 hat Vorrang
                ErmittleDashboardJahr = jahrF1
            End If
        End If
    ElseIf jahrF1 > 0 Then
        ErmittleDashboardJahr = jahrF1
    ElseIf jahrBK > 0 Then
        ErmittleDashboardJahr = jahrBK
    Else
        ErmittleDashboardJahr = Year(Date)
    End If
    
End Function


' ============================================================
'  SHEET ERSTELLEN / LEEREN
' ============================================================
Private Function HoleOderErstelleSheet() As Worksheet
    
    Dim sheetName As String
    sheetName = ChrW(220) & "bersicht (neu)"
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = sheetName
    End If
    
    Set HoleOderErstelleSheet = ws
    
End Function


' ============================================================
'  PARZELLEN GRUPPIEREN (1 Zeile pro Parzelle im Dashboard)
' ============================================================
Public Sub GruppiereParzellen(ByVal mitglieder As Collection, _
                                ByRef parzellen() As ParzelleInfo, _
                                ByRef anzParz As Long)
    
    ' Dictionary speichert nur den Index ins tempArr (UDTs
    ' koennen in Standardmodulen nicht als Variant abgelegt werden)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim tempArr() As ParzelleInfo
    Dim tempCount As Long
    tempCount = 0
    ReDim tempArr(1 To mitglieder.count)
    
    Dim m As Object
    For Each m In mitglieder
        Dim pNr As Long
        pNr = CLng(m("Parzelle"))
        
        Dim pKey As String
        pKey = CStr(pNr)
        
        If Not dict.Exists(pKey) Then
            tempCount = tempCount + 1
            tempArr(tempCount).parzNr = pNr
            tempArr(tempCount).mitgliedName = m("Name")
            tempArr(tempCount).entityKeys = m("EntityKey")
            tempArr(tempCount).roles = m("Role")
            dict.Add pKey, tempCount   ' merke Index
        Else
            ' Weitere Mitglieder auf gleicher Parzelle
            Dim ix As Long
            ix = dict(pKey)
            If InStr(tempArr(ix).entityKeys, m("EntityKey")) = 0 Then
                tempArr(ix).entityKeys = tempArr(ix).entityKeys & "," & m("EntityKey")
                tempArr(ix).mitgliedName = tempArr(ix).mitgliedName & " / " & m("Name")
                tempArr(ix).roles = tempArr(ix).roles & "," & m("Role")
            End If
        End If
    Next m
    
    anzParz = tempCount
    If anzParz = 0 Then
        Set dict = Nothing
        Exit Sub
    End If
    
    ReDim parzellen(1 To anzParz)
    
    ' Sortiert nach Parzellennummer einfuegen
    Dim idx As Long
    idx = 1
    Dim p As Long
    For p = 1 To 14
        If dict.Exists(CStr(p)) Then
            parzellen(idx) = tempArr(dict(CStr(p)))
            idx = idx + 1
        End If
    Next p
    
    ' Falls Parzellen > 14 existieren
    Dim key As Variant
    For Each key In dict.keys
        If CLng(key) > 14 Then
            If idx <= anzParz Then
                parzellen(idx) = tempArr(dict(key))
                idx = idx + 1
            End If
        End If
    Next key
    
    Set dict = Nothing
    
End Sub


' ============================================================
'  TITEL SCHREIBEN
' ============================================================
Private Sub SchreibeTitel(ByVal ws As Worksheet, ByVal jahr As Long, _
                           ByVal anzKat As Long)
    
    ' Zeile 1: Abstand oben
    ws.Rows(1).RowHeight = 8
    
    ' Zeile 2: Haupttitel
    Dim letzteSpalte As Long
    letzteSpalte = 3 + anzKat + 2  ' A,B,C + Kategorien + Gesamt + Quote
    If letzteSpalte < 8 Then letzteSpalte = 8
    
    With ws.Range(ws.Cells(TITEL_ROW, 1), ws.Cells(TITEL_ROW, letzteSpalte))
        .Merge
        .value = "ZAHLUNGS-DASHBOARD " & jahr
        .Font.Name = "Calibri"
        .Font.Size = 22
        .Font.Bold = True
        .Font.color = m_CLR_WEISS
        .Interior.color = m_CLR_NAVY
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 45
    End With
    
    ' Zeile 3: Untertitel
    With ws.Range(ws.Cells(3, 1), ws.Cells(3, letzteSpalte))
        .Merge
        .value = "Erstellt am " & Format(Now, "dd.mm.yyyy") & _
                 " um " & Format(Now, "hh:mm") & " Uhr"
        .Font.Name = "Calibri"
        .Font.Size = 9
        .Font.color = RGB(128, 128, 128)
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
        .RowHeight = 20
    End With
    
    ' Zeile 4: Trennlinie
    ws.Rows(4).RowHeight = 6
    
End Sub


' ============================================================
'  KPI-KARTEN SCHREIBEN
' ============================================================
Private Sub SchreibeKPI(ByVal ws As Worksheet, _
                         ByVal anzMitglieder As Long, _
                         ByVal summeIst As Double, _
                         ByVal summeSoll As Double, _
                         ByVal summeSaeumnis As Double, _
                         ByVal anzOffen As Long, _
                         ByVal anzBezahlt As Long)
    
    ' 4 KPI-Karten in Zeilen 5-7
    ' Karte 1: Spalte A-B (Mitglieder)
    Call SchreibeKPIKarte(ws, KPI_LABEL_ROW, 1, 2, _
                          "MITGLIEDER", _
                          CStr(anzMitglieder), _
                          "Aktive Parzellen", _
                          m_CLR_KPI_BLAU)
    
    ' Karte 2: Spalte C-D (Eingegangen)
    Call SchreibeKPIKarte(ws, KPI_LABEL_ROW, 3, 4, _
                          "EINGEGANGEN", _
                          Format(summeIst, "#,##0.00") & " " & ChrW(8364), _
                          CStr(anzBezahlt) & " Posten bezahlt", _
                          m_CLR_KPI_GRUEN)
    
    ' Karte 3: Spalte E-F (Offen)
    Dim offenBetrag As Double
    offenBetrag = summeSoll - summeIst
    If offenBetrag < 0 Then offenBetrag = 0
    
    Call SchreibeKPIKarte(ws, KPI_LABEL_ROW, 5, 6, _
                          "OFFEN", _
                          Format(offenBetrag, "#,##0.00") & " " & ChrW(8364), _
                          CStr(anzOffen) & " Posten offen", _
                          m_CLR_KPI_ROT)
    
    ' Karte 4: Spalte G-H (Saeumnis)
    Call SchreibeKPIKarte(ws, KPI_LABEL_ROW, 7, 8, _
                          "S" & ChrW(196) & "UMNIS", _
                          Format(summeSaeumnis, "#,##0.00") & " " & ChrW(8364), _
                          "Anfallende Geb" & ChrW(252) & "hren", _
                          m_CLR_KPI_ORANGE)
    
    ' Zeile 8: Trennlinie
    ws.Rows(8).RowHeight = 4
    
    ' Zeile 9: Abschnitt-Titel "ZAHLUNGSMATRIX"
    Dim matrixTitelSpalte As Long
    matrixTitelSpalte = ws.Cells(TITEL_ROW, 1).MergeArea.Columns.count
    If matrixTitelSpalte < 4 Then matrixTitelSpalte = 4
    
    With ws.Range(ws.Cells(9, 1), ws.Cells(9, matrixTitelSpalte))
        .Merge
        .value = ChrW(9632) & " ZAHLUNGSMATRIX"
        .Font.Name = "Calibri"
        .Font.Size = 13
        .Font.Bold = True
        .Font.color = m_CLR_NAVY
        .VerticalAlignment = xlCenter
        .RowHeight = 25
    End With
    
End Sub


' ============================================================
'  EINZELNE KPI-KARTE
' ============================================================
Private Sub SchreibeKPIKarte(ByVal ws As Worksheet, _
                              ByVal startRow As Long, _
                              ByVal colStart As Long, _
                              ByVal colEnd As Long, _
                              ByVal labelText As String, _
                              ByVal wertText As String, _
                              ByVal detailText As String, _
                              ByVal kartenFarbe As Long)
    
    ' Label-Zeile (Zeile 5)
    With ws.Range(ws.Cells(startRow, colStart), ws.Cells(startRow, colEnd))
        .Merge
        .value = labelText
        .Font.Name = "Calibri"
        .Font.Size = 9
        .Font.Bold = True
        .Font.color = m_CLR_WEISS
        .Interior.color = kartenFarbe
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 20
    End With
    
    ' Wert-Zeile (Zeile 6)
    With ws.Range(ws.Cells(startRow + 1, colStart), ws.Cells(startRow + 1, colEnd))
        .Merge
        .value = wertText
        .Font.Name = "Calibri"
        .Font.Size = 16
        .Font.Bold = True
        .Font.color = kartenFarbe
        .Interior.color = m_CLR_WEISS
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 32
        .Borders(xlEdgeLeft).color = kartenFarbe
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).color = kartenFarbe
        .Borders(xlEdgeRight).Weight = xlMedium
    End With
    
    ' Detail-Zeile (Zeile 7)
    With ws.Range(ws.Cells(startRow + 2, colStart), ws.Cells(startRow + 2, colEnd))
        .Merge
        .value = detailText
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.color = RGB(100, 100, 100)
        .Interior.color = m_CLR_WEISS
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 18
        .Borders(xlEdgeBottom).color = kartenFarbe
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeLeft).color = kartenFarbe
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).color = kartenFarbe
        .Borders(xlEdgeRight).Weight = xlMedium
    End With
    
End Sub


' ============================================================
'  MATRIX MIT DATEN SCHREIBEN
'  Sammelt gleichzeitig KPI-Werte und Verzug-Eintraege
' ============================================================
Public Sub SchreibeMatrixMitDaten(ByVal ws As Worksheet, _
                                    ByVal jahr As Long, _
                                    ByRef kategorien() As UebKategorie, _
                                    ByVal anzKat As Long, _
                                    ByRef parzellen() As ParzelleInfo, _
                                    ByVal anzParz As Long, _
                                    ByVal mitglieder As Collection, _
                                    ByRef matrixEndRow As Long, _
                                    ByRef kpiSummeIst As Double, _
                                    ByRef kpiSummeSoll As Double, _
                                    ByRef kpiSummeSaeumnis As Double, _
                                    ByRef kpiAnzahlOffen As Long, _
                                    ByRef kpiAnzahlBezahlt As Long, _
                                    ByRef verzugListe() As VerzugEintrag, _
                                    ByRef anzVerzug As Long)
    
    Dim importierteMonate() As Boolean
    importierteMonate = mod_Uebersicht_Daten.ErmittleImportierteMonate(jahr)
    
    ' --- Matrix-Header schreiben ---
    Dim headerRow As Long
    headerRow = MATRIX_HEADER_ROW
    
    With ws.Cells(headerRow, 1)
        .value = "Nr"
    End With
    With ws.Cells(headerRow, 2)
        .value = "Parzelle"
    End With
    With ws.Cells(headerRow, 3)
        .value = "Mitglied"
    End With
    
    ' Kategorie-Spalten ab Spalte 4
    Dim k As Long
    For k = 0 To anzKat - 1
        ws.Cells(headerRow, 4 + k).value = kategorien(k).Name
    Next k
    
    ' Gesamt-Spalten nach Kategorien
    Dim colGesamt As Long
    colGesamt = 4 + anzKat
    ws.Cells(headerRow, colGesamt).value = "Gesamt"
    ws.Cells(headerRow, colGesamt + 1).value = "Quote"
    
    Dim letzteSpalte As Long
    letzteSpalte = colGesamt + 1
    
    ' Header formatieren
    Dim rngHeader As Range
    Set rngHeader = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, letzteSpalte))
    With rngHeader
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .Font.color = m_CLR_WEISS
        .Interior.color = m_CLR_HEADER_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 28
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.color = m_CLR_WEISS
        .Borders.Weight = xlThin
    End With
    
    ' --- Datenzeilen ---
    Dim rowIdx As Long
    rowIdx = MATRIX_START_ROW
    
    ' KPI initialisieren
    kpiSummeIst = 0
    kpiSummeSoll = 0
    kpiSummeSaeumnis = 0
    kpiAnzahlOffen = 0
    kpiAnzahlBezahlt = 0
    
    ' Dictionary fuer Parzelle-basierte Kat (identisch mit GeneriereUebersicht)
    Dim geschriebeneKat As Object
    Set geschriebeneKat = CreateObject("Scripting.Dictionary")
    geschriebeneKat.CompareMode = vbTextCompare
    
    Dim p As Long
    For p = 1 To anzParz
        ' Nr + Parzelle + Mitglied
        ws.Cells(rowIdx, 1).value = p
        ws.Cells(rowIdx, 2).value = parzellen(p).parzNr
        ws.Cells(rowIdx, 3).value = parzellen(p).mitgliedName
        
        ' Zellen-Formatierung Basis
        ws.Range(ws.Cells(rowIdx, 1), ws.Cells(rowIdx, 3)).Font.Name = "Calibri"
        ws.Range(ws.Cells(rowIdx, 1), ws.Cells(rowIdx, 3)).Font.Size = 10
        ws.Cells(rowIdx, 1).HorizontalAlignment = xlCenter
        ws.Cells(rowIdx, 2).HorizontalAlignment = xlCenter
        ws.Cells(rowIdx, 2).Font.Bold = True
        
        Dim zeileSoll As Double
        Dim zeileIst As Double
        Dim zeileSaeumnis As Double
        zeileSoll = 0
        zeileIst = 0
        zeileSaeumnis = 0
        
        ' EntityKeys dieser Parzelle
        Dim eKeys() As String
        eKeys = Split(parzellen(p).entityKeys, ",")
        
        Dim roles As String
        roles = UCase(parzellen(p).roles)
        
        For k = 0 To anzKat - 1
            Dim kategorie As String
            kategorie = kategorien(k).Name
            
            Dim katCol As Long
            katCol = 4 + k
            
            ' Kategorie-Filter (wie in GeneriereUebersicht)
            ' OHNE PACHT: nur Mitgliedsbeitrag (nur wenn KEIN MIT PACHT auf der Parzelle)
            Dim istNurMitgliedsbeitrag As Boolean
            istNurMitgliedsbeitrag = (InStr(roles, "MIT PACHT") = 0 And _
                                      InStr(roles, "OHNE PACHT") > 0 And _
                                      StrComp(kategorie, "Mitgliedsbeitrag", vbTextCompare) <> 0)
            
            ' Ehrenmitglied: kein Mitgliedsbeitrag
            Dim istEhrenMitglied As Boolean
            istEhrenMitglied = (InStr(roles, "EHREN") > 0 And _
                                StrComp(kategorie, "Mitgliedsbeitrag", vbTextCompare) = 0)
            
            If istNurMitgliedsbeitrag Or istEhrenMitglied Then
                ' Nicht anwendbar
                With ws.Cells(rowIdx, katCol)
                    .value = ChrW(8212)  ' Gedankenstrich
                    .Font.color = RGB(180, 180, 180)
                    .HorizontalAlignment = xlCenter
                    .Interior.color = m_CLR_ZELLE_GRAU
                End With
                GoTo NextKatDash
            End If
            
            ' Pro Kategorie aggregieren ueber alle faelligen Monate
            Dim bezahltMonate As Long
            Dim faelligMonate As Long
            Dim katSoll As Double
            Dim katIst As Double
            Dim katSaeumnis As Double
            Dim katHatRot As Boolean
            Dim katHatGelb As Boolean
            bezahltMonate = 0
            faelligMonate = 0
            katSoll = 0
            katIst = 0
            katSaeumnis = 0
            katHatRot = False
            katHatGelb = False
            
            Dim monat As Long
            For monat = 1 To 12
                ' Pruefen ob Kategorie in diesem Monat faellig ist
                If Not IstKatImMonatFaellig(kategorien(k), monat) Then
                    GoTo NextMonatDash
                End If
                
                ' Pruefen ob Monat relevant (Daten vorhanden oder Frist abgelaufen)
                Dim monatRelevant As Boolean
                monatRelevant = False
                
                If importierteMonate(monat) Then
                    monatRelevant = True
                Else
                    ' Noch pruefen ob Frist abgelaufen
                    Dim testSollDatum As Date
                    Dim testNachlauf As Long
                    Dim testVorlauf As Long
                    Dim testSaeumnis As Double
                    testSollDatum = mod_Zahlungspruefung.BerechneSollDatumZP(kategorie, monat, jahr)
                    Call mod_Zahlungspruefung.HoleToleranzZP(kategorie, testVorlauf, testNachlauf, testSaeumnis)
                    If Date >= DateAdd("d", testNachlauf, testSollDatum) Then
                        monatRelevant = True
                    End If
                End If
                
                If Not monatRelevant Then GoTo NextMonatDash
                
                faelligMonate = faelligMonate + 1
                
                ' PruefeZahlungen fuer jeden EntityKey dieser Parzelle
                Dim eIdx As Long
                Dim besteIst As Double
                Dim besteSoll As Double
                Dim besterStatus As String
                Dim besteBemerkung As String
                besteIst = 0
                besteSoll = 0
                besterStatus = "ROT"
                besteBemerkung = ""
                
                For eIdx = LBound(eKeys) To UBound(eKeys)
                    Dim ek As String
                    ek = Trim(eKeys(eIdx))
                    If ek = "" Then GoTo NextEKDash
                    
                    Dim ergebnis As String
                    ergebnis = mod_Zahlungspruefung.PruefeZahlungen(ek, kategorie, monat, jahr)
                    
                    ' Parsen: "STATUS|Soll:X.XX|Ist:Y.YY|Bemerkung"
                    Dim teile() As String
                    teile = Split(ergebnis, "|")
                    
                    Dim tmpStatus As String
                    Dim tmpSoll As Double
                    Dim tmpIst As Double
                    Dim tmpBem As String
                    tmpStatus = "ROT"
                    tmpSoll = 0
                    tmpIst = 0
                    tmpBem = ""
                    
                    If UBound(teile) >= 2 Then
                        tmpStatus = teile(0)
                        Dim sollT() As String
                        sollT = Split(teile(1), ":")
                        If UBound(sollT) >= 1 Then tmpSoll = val(sollT(1))
                        Dim istT() As String
                        istT = Split(teile(2), ":")
                        If UBound(istT) >= 1 Then tmpIst = val(istT(1))
                    ElseIf UBound(teile) >= 0 Then
                        tmpStatus = teile(0)
                    End If
                    If UBound(teile) >= 3 Then tmpBem = teile(3)
                    
                    ' Bestes Ergebnis waehlen (Mitgliedsbeitrag: pro Person)
                    If tmpIst > besteIst Then besteIst = tmpIst
                    If tmpSoll > besteSoll Then besteSoll = tmpSoll
                    
                    ' Status: GRUEN > GELB > ROT
                    Dim statusGruen As String
                    statusGruen = "GR" & ChrW(220) & "N"
                    If StrComp(tmpStatus, statusGruen, vbTextCompare) = 0 Then
                        besterStatus = statusGruen
                    ElseIf StrComp(tmpStatus, "GELB", vbTextCompare) = 0 Then
                        If StrComp(besterStatus, statusGruen, vbTextCompare) <> 0 Then
                            besterStatus = "GELB"
                        End If
                    End If
                    If tmpBem <> "" Then besteBemerkung = tmpBem
                    
NextEKDash:
                Next eIdx
                
                ' Ergebnis auswerten
                katSoll = katSoll + besteSoll
                katIst = katIst + besteIst
                
                If StrComp(besterStatus, statusGruen, vbTextCompare) = 0 Then
                    bezahltMonate = bezahltMonate + 1
                    kpiAnzahlBezahlt = kpiAnzahlBezahlt + 1
                ElseIf StrComp(besterStatus, "GELB", vbTextCompare) = 0 Then
                    katHatGelb = True
                    If besteIst > 0 Then
                        bezahltMonate = bezahltMonate + 1
                        kpiAnzahlBezahlt = kpiAnzahlBezahlt + 1
                    Else
                        kpiAnzahlOffen = kpiAnzahlOffen + 1
                    End If
                Else
                    katHatRot = True
                    kpiAnzahlOffen = kpiAnzahlOffen + 1
                    
                    ' Saeumnis berechnen
                    If kategorien(k).saeumnisGebuehr > 0 Then
                        katSaeumnis = katSaeumnis + kategorien(k).saeumnisGebuehr
                    End If
                    
                    ' Verzug-Eintrag sammeln
                    If anzVerzug <= UBound(verzugListe) Then
                        With verzugListe(anzVerzug)
                            .parzNr = parzellen(p).parzNr
                            .mitglied = parzellen(p).mitgliedName
                            .kategorie = kategorie
                            .monatNr = monat
                            .monatText = MonthName(monat) & " " & jahr
                            .soll = besteSoll
                            .ist = besteIst
                            .differenz = besteSoll - besteIst
                            .saeumnis = kategorien(k).saeumnisGebuehr
                            .bemerkung = besteBemerkung
                            
                            ' Tage Verzug berechnen
                            Dim sollDatum As Date
                            Dim vorlauf As Long
                            Dim nachlauf As Long
                            Dim sgb As Double
                            sollDatum = mod_Zahlungspruefung.BerechneSollDatumZP(kategorie, monat, jahr)
                            Call mod_Zahlungspruefung.HoleToleranzZP(kategorie, vorlauf, nachlauf, sgb)
                            Dim fristDatum As Date
                            fristDatum = DateAdd("d", nachlauf, sollDatum)
                            If Date > fristDatum Then
                                .tageVerzug = DateDiff("d", fristDatum, Date)
                            Else
                                .tageVerzug = 0
                            End If
                        End With
                        anzVerzug = anzVerzug + 1
                    End If
                End If
                
NextMonatDash:
            Next monat
            
            ' Matrix-Zelle schreiben
            kpiSummeSoll = kpiSummeSoll + katSoll
            kpiSummeIst = kpiSummeIst + katIst
            kpiSummeSaeumnis = kpiSummeSaeumnis + katSaeumnis
            zeileSoll = zeileSoll + katSoll
            zeileIst = zeileIst + katIst
            zeileSaeumnis = zeileSaeumnis + katSaeumnis
            
            Call SchreibeMatrixZelle(ws, rowIdx, katCol, _
                                     faelligMonate, bezahltMonate, _
                                     katSoll, katIst, katHatRot, katHatGelb)
            
NextKatDash:
        Next k
        
        ' Gesamt-Spalten
        ws.Cells(rowIdx, colGesamt).value = zeileIst
        ws.Cells(rowIdx, colGesamt).NumberFormat = "#,##0.00"
        ws.Cells(rowIdx, colGesamt).Font.Bold = True
        ws.Cells(rowIdx, colGesamt).Font.Name = "Calibri"
        ws.Cells(rowIdx, colGesamt).Font.Size = 10
        ws.Cells(rowIdx, colGesamt).HorizontalAlignment = xlRight
        
        ' Quote
        Dim quote As Double
        If zeileSoll > 0 Then
            quote = zeileIst / zeileSoll
        Else
            quote = 1
        End If
        
        With ws.Cells(rowIdx, colGesamt + 1)
            .value = quote
            .NumberFormat = "0%"
            .Font.Name = "Calibri"
            .Font.Size = 10
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            ' Farbe basierend auf Quote
            If quote >= 1 Then
                .Font.color = m_CLR_TEXT_GRUEN
                .Interior.color = m_CLR_ZELLE_GRUEN
            ElseIf quote >= 0.5 Then
                .Font.color = RGB(120, 100, 0)
                .Interior.color = m_CLR_ZELLE_GELB
            Else
                .Font.color = m_CLR_TEXT_DUNKELROT
                .Interior.color = m_CLR_ZELLE_ROT
            End If
        End With
        
        ' Zebra-Streifen fuer Nr, Parzelle, Mitglied
        If p Mod 2 = 0 Then
            ws.Range(ws.Cells(rowIdx, 1), ws.Cells(rowIdx, 3)).Interior.color = RGB(245, 245, 250)
        End If
        
        ' Rahmen
        ws.Range(ws.Cells(rowIdx, 1), ws.Cells(rowIdx, letzteSpalte)).Borders.LineStyle = xlContinuous
        ws.Range(ws.Cells(rowIdx, 1), ws.Cells(rowIdx, letzteSpalte)).Borders.color = RGB(220, 220, 220)
        ws.Range(ws.Cells(rowIdx, 1), ws.Cells(rowIdx, letzteSpalte)).Borders.Weight = xlThin
        ws.Rows(rowIdx).RowHeight = 26
        
        rowIdx = rowIdx + 1
    Next p
    
    matrixEndRow = rowIdx - 1
    
    ' Summenzeile unter der Matrix
    rowIdx = matrixEndRow + 1
    ws.Cells(rowIdx, 3).value = "SUMME"
    ws.Cells(rowIdx, 3).Font.Bold = True
    ws.Cells(rowIdx, 3).Font.Name = "Calibri"
    ws.Cells(rowIdx, 3).HorizontalAlignment = xlRight
    ws.Cells(rowIdx, colGesamt).value = kpiSummeIst
    ws.Cells(rowIdx, colGesamt).NumberFormat = "#,##0.00"
    ws.Cells(rowIdx, colGesamt).Font.Bold = True
    ws.Cells(rowIdx, colGesamt).Font.Name = "Calibri"
    
    Dim gesamtQuote As Double
    If kpiSummeSoll > 0 Then
        gesamtQuote = kpiSummeIst / kpiSummeSoll
    Else
        gesamtQuote = 1
    End If
    ws.Cells(rowIdx, colGesamt + 1).value = gesamtQuote
    ws.Cells(rowIdx, colGesamt + 1).NumberFormat = "0%"
    ws.Cells(rowIdx, colGesamt + 1).Font.Bold = True
    ws.Cells(rowIdx, colGesamt + 1).Font.Name = "Calibri"
    ws.Cells(rowIdx, colGesamt + 1).HorizontalAlignment = xlCenter
    
    ' Summenzeile Rahmen oben (doppelt)
    With ws.Range(ws.Cells(rowIdx, 1), ws.Cells(rowIdx, letzteSpalte))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeTop).color = m_CLR_HEADER_BG
        .RowHeight = 24
    End With
    
    matrixEndRow = rowIdx
    
    ' Datenbalken fuer Quote-Spalte (Conditional Formatting)
    On Error Resume Next
    Dim rngQuote As Range
    Set rngQuote = ws.Range(ws.Cells(MATRIX_START_ROW, colGesamt + 1), _
                             ws.Cells(matrixEndRow - 1, colGesamt + 1))
    rngQuote.FormatConditions.Delete
    
    Dim db As Object
    Set db = rngQuote.FormatConditions.AddDatabar
    If Not db Is Nothing Then
        db.BarColor.color = RGB(41, 128, 185)
        db.BarFillType = xlDataBarFillGradient
        db.MinPoint.Modify newtype:=xlConditionValueNumber, newValue:=0
        db.MaxPoint.Modify newtype:=xlConditionValueNumber, newValue:=1
        db.BarBorder.Type = xlDataBarBorderSolid
        db.BarBorder.color.color = RGB(41, 128, 185)
        db.ShowValue = True
    End If
    On Error GoTo 0
    
    Set geschriebeneKat = Nothing
    
End Sub


' ============================================================
'  MATRIX-ZELLE SCHREIBEN (eine Kategorie pro Parzelle)
' ============================================================
Private Sub SchreibeMatrixZelle(ByVal ws As Worksheet, _
                                 ByVal zeile As Long, _
                                 ByVal spalte As Long, _
                                 ByVal faellig As Long, _
                                 ByVal bezahlt As Long, _
                                 ByVal soll As Double, _
                                 ByVal ist As Double, _
                                 ByVal hatRot As Boolean, _
                                 ByVal hatGelb As Boolean)
    
    With ws.Cells(zeile, spalte)
        .Font.Name = "Calibri"
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        If faellig = 0 Then
            ' Noch nicht faellig / keine Daten
            .value = ChrW(8212)  ' Gedankenstrich
            .Font.color = RGB(180, 180, 180)
            .Interior.color = m_CLR_ZELLE_GRAU
            
        ElseIf bezahlt >= faellig And Not hatRot Then
            ' Alles bezahlt
            .value = ChrW(10004) & " " & Format(ist, "#,##0.00")
            .Font.color = m_CLR_TEXT_GRUEN
            .Font.Bold = True
            .Interior.color = m_CLR_ZELLE_GRUEN
            
        ElseIf hatRot And bezahlt = 0 Then
            ' Nichts bezahlt, ueberfaellig
            .value = ChrW(10008) & " " & Format(soll, "#,##0.00")
            .Font.color = m_CLR_TEXT_DUNKELROT
            .Font.Bold = True
            .Interior.color = m_CLR_ZELLE_ROT
            
        ElseIf hatRot Then
            ' Teilweise bezahlt mit offenen Posten
            .value = CStr(bezahlt) & "/" & CStr(faellig) & " " & ChrW(8226) & " " & _
                     Format(ist, "#,##0.00")
            .Font.color = m_CLR_TEXT_DUNKELROT
            .Interior.color = m_CLR_ZELLE_ROT
            
        ElseIf hatGelb Then
            ' Teilweise / verspaetet
            .value = CStr(bezahlt) & "/" & CStr(faellig) & " " & ChrW(8226) & " " & _
                     Format(ist, "#,##0.00")
            .Font.color = RGB(120, 100, 0)
            .Interior.color = m_CLR_ZELLE_GELB
            
        Else
            ' Teils bezahlt, kein Rot
            .value = Format(ist, "#,##0.00")
            .Font.color = m_CLR_TEXT_GRUEN
            .Interior.color = m_CLR_ZELLE_GRUEN
        End If
    End With
    
End Sub


' ============================================================
'  FAELLIGKEIT PRUEFEN (Kopie der Private-Logik aus Generator)
' ============================================================
Public Function IstKatImMonatFaellig(ByRef kat As UebKategorie, _
                                       ByVal monat As Long) As Boolean
    
    ' SollMonate definiert -> nur diese Monate
    If kat.SollMonate <> "" Then
        IstKatImMonatFaellig = mod_KategorieEngine_Zeitraum.IstMonatInListe(monat, kat.SollMonate)
        Exit Function
    End If
    
    ' Faelligkeit pruefen
    Dim fl As String
    fl = kat.faelligkeit
    
    ' Monatlich oder leer -> alle Monate
    If fl = "" Or fl = "monatlich" Then
        IstKatImMonatFaellig = True
        Exit Function
    End If
    
    ' Nicht-monatlich ohne SollMonate -> nicht anzeigen
    IstKatImMonatFaellig = False
    
End Function


' ============================================================
'  VERZUG SORTIEREN (Bubble Sort nach TageVerzug absteigend)
' ============================================================
Public Sub SortiereVerzug(ByRef liste() As VerzugEintrag, ByVal anzahl As Long)
    
    Dim i As Long, j As Long
    Dim temp As VerzugEintrag
    
    For i = 0 To anzahl - 2
        For j = 0 To anzahl - 2 - i
            If liste(j).tageVerzug < liste(j + 1).tageVerzug Then
                temp = liste(j)
                liste(j) = liste(j + 1)
                liste(j + 1) = temp
            End If
        Next j
    Next i
    
End Sub


' ============================================================
'  VERZUGSDETAIL SCHREIBEN
' ============================================================
Public Sub SchreibeVerzugsdetail(ByVal ws As Worksheet, _
                                    ByVal startRow As Long, _
                                    ByRef liste() As VerzugEintrag, _
                                    ByVal anzahl As Long, _
                                    ByRef endRow As Long)
    
    ' Abschnitt-Titel
    Dim titelCol As Long
    titelCol = ws.Cells(TITEL_ROW, 1).MergeArea.Columns.count
    If titelCol < 10 Then titelCol = 10
    
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, titelCol))
        .Merge
        .value = ChrW(9632) & " VERZUGSDETAIL " & ChrW(8212) & _
                 " OFFENE ZAHLUNGEN (" & anzahl & ")"
        .Font.Name = "Calibri"
        .Font.Size = 13
        .Font.Bold = True
        .Font.color = m_CLR_KPI_ROT
        .VerticalAlignment = xlCenter
        .RowHeight = 28
    End With
    
    ' Header
    Dim hRow As Long
    hRow = startRow + 1
    
    Dim headers As Variant
    headers = Array("Parzelle", "Mitglied", "Kategorie", "Monat", _
                    "Soll", "Ist", "Differenz", _
                    "S" & ChrW(228) & "umnis", "Tage Verzug", "Bemerkung")
    
    Dim c As Long
    For c = 0 To 9
        ws.Cells(hRow, c + 1).value = headers(c)
    Next c
    
    Dim rngVHeader As Range
    Set rngVHeader = ws.Range(ws.Cells(hRow, 1), ws.Cells(hRow, 10))
    With rngVHeader
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .Font.color = m_CLR_WEISS
        .Interior.color = m_CLR_KPI_ROT
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 24
        .Borders.LineStyle = xlContinuous
        .Borders.color = m_CLR_WEISS
        .Borders.Weight = xlThin
    End With
    
    ' Datenzeilen
    Dim dRow As Long
    dRow = hRow + 1
    
    Dim i As Long
    For i = 0 To anzahl - 1
        With liste(i)
            ws.Cells(dRow, 1).value = .parzNr
            ws.Cells(dRow, 2).value = .mitglied
            ws.Cells(dRow, 3).value = .kategorie
            ws.Cells(dRow, 4).value = .monatText
            ws.Cells(dRow, 5).value = .soll
            ws.Cells(dRow, 5).NumberFormat = "#,##0.00"
            ws.Cells(dRow, 6).value = .ist
            ws.Cells(dRow, 6).NumberFormat = "#,##0.00"
            ws.Cells(dRow, 7).value = .differenz
            ws.Cells(dRow, 7).NumberFormat = "#,##0.00"
            ws.Cells(dRow, 8).value = .saeumnis
            ws.Cells(dRow, 8).NumberFormat = "#,##0.00"
            ws.Cells(dRow, 9).value = .tageVerzug
            ws.Cells(dRow, 10).value = .bemerkung
        End With
        
        ' Formatierung
        Dim rngVRow As Range
        Set rngVRow = ws.Range(ws.Cells(dRow, 1), ws.Cells(dRow, 10))
        With rngVRow
            .Font.Name = "Calibri"
            .Font.Size = 9
            .VerticalAlignment = xlCenter
            .RowHeight = 22
            .Borders.LineStyle = xlContinuous
            .Borders.color = RGB(220, 220, 220)
            .Borders.Weight = xlThin
        End With
        
        ws.Cells(dRow, 1).HorizontalAlignment = xlCenter
        ws.Cells(dRow, 4).HorizontalAlignment = xlCenter
        ws.Cells(dRow, 9).HorizontalAlignment = xlCenter
        
        ' Farbe basierend auf Tage Verzug
        If liste(i).tageVerzug > 60 Then
            ' Stark ueberfaellig -> intensives Rot
            ws.Range(ws.Cells(dRow, 1), ws.Cells(dRow, 10)).Interior.color = RGB(255, 220, 220)
            ws.Cells(dRow, 9).Font.Bold = True
            ws.Cells(dRow, 9).Font.color = m_CLR_TEXT_DUNKELROT
        ElseIf liste(i).tageVerzug > 30 Then
            ' Mittel ueberfaellig -> helles Rot
            ws.Range(ws.Cells(dRow, 1), ws.Cells(dRow, 10)).Interior.color = m_CLR_ZELLE_ROT
        ElseIf liste(i).tageVerzug > 0 Then
            ' Leicht ueberfaellig -> Gelb
            ws.Range(ws.Cells(dRow, 1), ws.Cells(dRow, 10)).Interior.color = m_CLR_ZELLE_GELB
        End If
        
        ' Tage-Verzug Balken-Effekt in Spalte 9
        If liste(i).tageVerzug > 0 Then
            Dim balkenLaenge As Long
            balkenLaenge = liste(i).tageVerzug \ 10
            If balkenLaenge > 10 Then balkenLaenge = 10
            If balkenLaenge < 1 Then balkenLaenge = 1
            ws.Cells(dRow, 9).value = CStr(liste(i).tageVerzug) & " " & _
                                       Application.WorksheetFunction.Rept(ChrW(9608), balkenLaenge)
        End If
        
        ' Zebra
        If i Mod 2 = 1 And liste(i).tageVerzug = 0 Then
            rngVRow.Interior.color = RGB(250, 250, 252)
        End If
        
        dRow = dRow + 1
    Next i
    
    endRow = dRow - 1
    
    ' Summenzeile
    If anzahl > 0 Then
        ws.Cells(dRow, 4).value = "SUMME:"
        ws.Cells(dRow, 4).Font.Bold = True
        ws.Cells(dRow, 4).HorizontalAlignment = xlRight
        ws.Cells(dRow, 4).Font.Name = "Calibri"
        
        ' Summenformeln
        ws.Cells(dRow, 5).Formula = "=SUM(" & ws.Cells(hRow + 1, 5).Address & _
                                    ":" & ws.Cells(dRow - 1, 5).Address & ")"
        ws.Cells(dRow, 5).NumberFormat = "#,##0.00"
        ws.Cells(dRow, 5).Font.Bold = True
        
        ws.Cells(dRow, 7).Formula = "=SUM(" & ws.Cells(hRow + 1, 7).Address & _
                                    ":" & ws.Cells(dRow - 1, 7).Address & ")"
        ws.Cells(dRow, 7).NumberFormat = "#,##0.00"
        ws.Cells(dRow, 7).Font.Bold = True
        ws.Cells(dRow, 7).Font.color = m_CLR_TEXT_DUNKELROT
        
        ws.Cells(dRow, 8).Formula = "=SUM(" & ws.Cells(hRow + 1, 8).Address & _
                                    ":" & ws.Cells(dRow - 1, 8).Address & ")"
        ws.Cells(dRow, 8).NumberFormat = "#,##0.00"
        ws.Cells(dRow, 8).Font.Bold = True
        ws.Cells(dRow, 8).Font.color = m_CLR_KPI_ORANGE
        
        With ws.Range(ws.Cells(dRow, 1), ws.Cells(dRow, 10))
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeTop).color = m_CLR_KPI_ROT
            .RowHeight = 22
        End With
        
        endRow = dRow
    End If
    
End Sub


' ============================================================
'  SPALTENBREITEN ANPASSEN
' ============================================================
Private Sub PasseSpaltenAn(ByVal ws As Worksheet, ByVal anzKat As Long)
    
    ' Feste Breiten
    ws.Columns(1).ColumnWidth = 5   ' Nr
    ws.Columns(2).ColumnWidth = 10  ' Parzelle
    ws.Columns(3).ColumnWidth = 22  ' Mitglied
    
    ' Kategorie-Spalten
    Dim k As Long
    For k = 0 To anzKat - 1
        ws.Columns(4 + k).ColumnWidth = 18
    Next k
    
    ' Gesamt + Quote
    ws.Columns(4 + anzKat).ColumnWidth = 14    ' Gesamt
    ws.Columns(4 + anzKat + 1).ColumnWidth = 10 ' Quote
    
    ' Verzugsdetail Spalten (1-10): AutoFit soweit moeglich
    ' Spalte 10 (Bemerkung) breiter
    If ws.Cells(ws.Rows.count, 10).End(xlUp).Row > MATRIX_START_ROW + 20 Then
        ws.Columns(10).ColumnWidth = 35
    End If
    
End Sub












