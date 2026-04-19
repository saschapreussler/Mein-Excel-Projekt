Attribute VB_Name = "mod_Uebersicht_Dashboard"
Option Explicit

' ***************************************************************
' MODUL: mod_Uebersicht_Dashboard
' VERSION: 2.0 - 27.03.2026
' ZWECK: Visuelles Dashboard auf Blatt "Uebersicht (neu)"
'        Einstiegspunkt, KPI-Karten, Parzellen-Gruppierung
' ABHAENGIGKEITEN: mod_Dashboard_Matrix, mod_Uebersicht_Daten,
'                  mod_Zahlungspruefung
' ***************************************************************

' ============================================================
'  TYPES (muessen VOR allen Sub/Function stehen!)
' ============================================================
Public Type ParzelleInfo
    parzNr As Long
    mitgliedNamen As String    ' vbLf-getrennt
    entityKeys As String       ' komma-getrennt
    roles As String            ' komma-getrennt (parallel zu entityKeys)
    eintritte As String        ' komma-getrennt (parallel zu entityKeys, Format YYYYMMDD)
    anzMitglieder As Long
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
'  FARBEN (Public fuer mod_Dashboard_Matrix)
' ============================================================
Public m_CLR_NAVY As Long
Public m_CLR_HEADER_BG As Long
Public m_CLR_WEISS As Long
Public m_CLR_KPI_BLAU As Long
Public m_CLR_KPI_GRUEN As Long
Public m_CLR_KPI_ROT As Long
Public m_CLR_KPI_ORANGE As Long
Public m_CLR_ZELLE_GRUEN As Long
Public m_CLR_ZELLE_GELB As Long
Public m_CLR_ZELLE_ROT As Long
Public m_CLR_ZELLE_GRAU As Long
Public m_CLR_TEXT_GRUEN As Long
Public m_CLR_TEXT_DUNKELROT As Long
Private m_FarbenInit As Boolean

' Layout-Konstanten (Public fuer Matrix-Modul)
Public Const DASH_TITEL_ROW As Long = 2
Public Const DASH_KPI_LABEL_ROW As Long = 5
Public Const DASH_KPI_WERT_ROW As Long = 6
Public Const DASH_KPI_DETAIL_ROW As Long = 7
Public Const DASH_MATRIX_HEADER_ROW As Long = 10
Public Const DASH_MATRIX_START_ROW As Long = 11


Public Sub InitFarben()
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
'  HAUPTFUNKTION
' ============================================================
Public Sub GeneriereUebersichtNeu(Optional ByVal stummModus As Boolean = False)
    
    On Error GoTo ErrorHandler
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
    
    ' --- 3a. Mitglieder aus Mitgliederliste ergaenzen ---
    ' Auch Mitglieder OHNE eigenen EntityKey werden angezeigt
    Dim mitgliederML As Collection
    Set mitgliederML = mod_Uebersicht_Daten.HoleMitgliederAusMitgliederliste()
    
    Dim anzMitglieder As Long
    anzMitglieder = mitgliederML.count
    
    Call ErgaenzeParzellennamen(parzellen, anzParz, mitgliederML)
    
    ' --- 4. Soll-Werte aus Uebersicht laden ---
    Dim sollDict As Object
    Set sollDict = LadeSollAusUebersicht()
    
    ' --- 5. Sheet erstellen / leeren ---
    Dim wsDash As Worksheet
    Set wsDash = HoleOderErstelleSheet()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    wsDash.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    wsDash.Cells.Clear
    wsDash.Cells.Interior.color = m_CLR_WEISS
    
    ' --- 6. Einstellungen-Cache laden ---
    Call mod_Zahlungspruefung.LadeEinstellungenCacheZP
    Call mod_Zahlungspruefung.InitialisiereNachDezemberCacheZP(jahr)
    Call mod_Uebersicht_Daten.BefuelleVorjahrSpeicher(jahr - 1)
    
    ' --- 7. Titel schreiben ---
    Call SchreibeTitel(wsDash, jahr, anzKat)
    
    ' --- 8. Matrix + Verzug + KPI-Daten ---
    Dim matrixEndRow As Long
    Dim kpiSummeIst As Double
    Dim kpiSummeSoll As Double
    Dim kpiSummeSaeumnis As Double
    Dim kpiAnzahlOffen As Long
    Dim kpiAnzahlBezahlt As Long
    Dim kpiAnzahlSaeumnis As Long
    Dim kpiOffenOhneSoll As Long
    Dim kpiOffenBetrag As Double
    
    Dim verzugListe() As VerzugEintrag
    Dim anzVerzug As Long
    anzVerzug = 0
    ReDim verzugListe(0 To 999)
    
    Call mod_Dashboard_Matrix.SchreibeMatrixMitDaten( _
        wsDash, jahr, kategorien, anzKat, _
        parzellen, anzParz, mitglieder, sollDict, _
        matrixEndRow, _
        kpiSummeIst, kpiSummeSoll, kpiSummeSaeumnis, _
        kpiAnzahlOffen, kpiAnzahlBezahlt, _
        kpiAnzahlSaeumnis, kpiOffenOhneSoll, _
        kpiOffenBetrag, _
        verzugListe, anzVerzug)
    
    ' --- 9. KPI-Karten ---
    ' anzMitglieder wurde oben aus Mitgliederliste ermittelt (Schritt 3a)
    
    Call SchreibeKPI(wsDash, anzParz, anzMitglieder, _
                     kpiSummeIst, kpiSummeSoll, kpiSummeSaeumnis, _
                     kpiAnzahlOffen, kpiAnzahlBezahlt, _
                     kpiAnzahlSaeumnis, kpiOffenOhneSoll, _
                     kpiOffenBetrag)
    
    ' --- 10. Verzugsdetail ---
    Dim verzugEndRow As Long
    If anzVerzug > 0 Then
        ReDim Preserve verzugListe(0 To anzVerzug - 1)
        Call mod_Dashboard_Matrix.SortiereVerzug(verzugListe, anzVerzug)
        Call mod_Dashboard_Matrix.SchreibeVerzugsdetail( _
            wsDash, matrixEndRow + 3, verzugListe, anzVerzug, verzugEndRow)
    End If
    
    ' --- 11. Cache freigeben ---
    Call mod_Zahlungspruefung.EntladeEinstellungenCacheZP
    
    ' --- 12. Spaltenbreiten ---
    Call mod_Dashboard_Matrix.PasseSpaltenAn(wsDash, anzKat)
    
    ' --- 13. Blatt schuetzen ---
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
               "Mitglieder: " & anzMitglieder & vbLf & _
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
'  JAHR ERMITTELN
' ============================================================
Private Function ErmittleDashboardJahr(Optional ByVal stummModus As Boolean = False) As Long
    
    ' v6.0: Abrechnungsjahr aus Einstellungen statt Startmenue!F1
    Dim jahrF1 As Long
    jahrF1 = HoleAbrechnungsjahr()
    
    Dim jahrBK As Long
    jahrBK = mod_Uebersicht_Daten.ErmittleJahrAusBankkonto()
    
    If jahrF1 > 0 And jahrBK > 0 Then
        If jahrF1 = jahrBK Then
            ErmittleDashboardJahr = jahrF1
        Else
            If Not stummModus Then
                Dim antwort As VbMsgBoxResult
                antwort = MsgBox("Abrechnungsjahr = " & jahrF1 & _
                                 ", Bankkonto = " & jahrBK & "." & vbLf & vbLf & _
                                 "Dashboard f" & ChrW(252) & "r " & jahrF1 & _
                                 " (Einstellungen) erstellen?", _
                                 vbQuestion + vbYesNo, "Abrechnungsjahr")
                If antwort = vbYes Then
                    ErmittleDashboardJahr = jahrF1
                Else
                    ErmittleDashboardJahr = jahrBK
                End If
            Else
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
    sheetName = "Dashboard Mitgliederzahlungen"
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Direkt hinter dem Blatt "Zahlungsuebersicht" einfuegen
        Dim wsUeb As Worksheet
        On Error Resume Next
        Set wsUeb = ThisWorkbook.Worksheets(WS_UEBERSICHT())
        On Error GoTo 0
        
        If Not wsUeb Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=wsUeb)
        Else
            Set ws = ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        End If
        ws.Name = sheetName
    End If
    
    Set HoleOderErstelleSheet = ws
    
End Function


' ============================================================
'  PARZELLEN GRUPPIEREN (1 Zeile pro Parzelle)
' ============================================================
Public Sub GruppiereParzellen(ByVal mitglieder As Collection, _
                                ByRef parzellen() As ParzelleInfo, _
                                ByRef anzParz As Long)
    
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
            tempArr(tempCount).mitgliedNamen = m("Name")
            tempArr(tempCount).entityKeys = m("EntityKey")
            tempArr(tempCount).roles = m("Role")
            ' v5.2: Eintrittsdatum parallel zu entityKeys speichern
            If IsDate(m("Eintritt")) And CDate(m("Eintritt")) > 0 Then
                tempArr(tempCount).eintritte = Format(CDate(m("Eintritt")), "YYYYMMDD")
            Else
                tempArr(tempCount).eintritte = ""
            End If
            tempArr(tempCount).anzMitglieder = 1
            dict.Add pKey, tempCount
        Else
            Dim ix As Long
            ix = dict(pKey)
            If InStr(tempArr(ix).entityKeys, m("EntityKey")) = 0 Then
                ' Neuer EntityKey fuer diese Parzelle -> hinzufuegen
                tempArr(ix).entityKeys = tempArr(ix).entityKeys & "," & m("EntityKey")
                tempArr(ix).roles = tempArr(ix).roles & "," & m("Role")
                ' v5.2: Eintrittsdatum parallel speichern
                Dim eDat As String
                If IsDate(m("Eintritt")) And CDate(m("Eintritt")) > 0 Then
                    eDat = Format(CDate(m("Eintritt")), "YYYYMMDD")
                Else
                    eDat = ""
                End If
                tempArr(ix).eintritte = tempArr(ix).eintritte & "," & eDat
            End If
            
            ' v5.2: Name IMMER pruefen (auch wenn EntityKey bereits bekannt),
            '        damit alle Personen auf der Parzelle dargestellt werden
            If InStr(1, tempArr(ix).mitgliedNamen, m("Name"), vbTextCompare) = 0 Then
                tempArr(ix).mitgliedNamen = tempArr(ix).mitgliedNamen & vbLf & m("Name")
                tempArr(ix).anzMitglieder = tempArr(ix).anzMitglieder + 1
            End If
        End If
    Next m
    
    anzParz = tempCount
    If anzParz = 0 Then
        Set dict = Nothing
        Exit Sub
    End If
    
    ReDim parzellen(1 To anzParz)
    
    Dim idx As Long
    idx = 1
    Dim p As Long
    For p = 1 To 14
        If dict.Exists(CStr(p)) Then
            parzellen(idx) = tempArr(dict(CStr(p)))
            idx = idx + 1
        End If
    Next p
    
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
'  SOLL-WERTE AUS UEBERSICHT-BLATT LADEN
'  Liest manuell eingetragene Soll-Betraege (Parzelle|Kategorie)
' ============================================================
Public Function LadeSollAusUebersicht() As Object
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    Dim wsUeb As Worksheet
    On Error Resume Next
    Set wsUeb = ThisWorkbook.Worksheets(WS_UEBERSICHT())
    On Error GoTo 0
    
    If wsUeb Is Nothing Then
        Set LadeSollAusUebersicht = dict
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = wsUeb.Cells(wsUeb.Rows.count, 4).End(xlUp).Row
    If lastRow < 4 Then
        Set LadeSollAusUebersicht = dict
        Exit Function
    End If
    
    ' Spalte A=Parzelle(1), D=Kategorie(4), E=Soll(5)
    Dim r As Long
    For r = 4 To lastRow
        Dim parz As String
        parz = Trim(CStr(wsUeb.Cells(r, 1).value))
        If parz = "" Then GoTo NextSollRow
        
        Dim kat As String
        kat = Trim(CStr(wsUeb.Cells(r, 4).value))
        If kat = "" Then GoTo NextSollRow
        
        Dim sollWert As Double
        sollWert = 0
        If IsNumeric(wsUeb.Cells(r, 5).value) Then
            sollWert = CDbl(wsUeb.Cells(r, 5).value)
        End If
        
        If sollWert > 0 Then
            Dim dictKey As String
            dictKey = parz & "|" & kat
            dict(dictKey) = sollWert
        End If
        
NextSollRow:
    Next r
    
    Set LadeSollAusUebersicht = dict
    
End Function


' ============================================================
'  TITEL SCHREIBEN
' ============================================================
Private Sub SchreibeTitel(ByVal ws As Worksheet, ByVal jahr As Long, _
                           ByVal anzKat As Long)
    ws.Rows(1).RowHeight = 8
    
    Dim letzteSpalte As Long
    letzteSpalte = 2 + anzKat + 2
    If letzteSpalte < 8 Then letzteSpalte = 8
    
    With ws.Range(ws.Cells(DASH_TITEL_ROW, 1), ws.Cells(DASH_TITEL_ROW, letzteSpalte))
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
    
    ws.Rows(4).RowHeight = 6
End Sub


' ============================================================
'  KPI-KARTEN SCHREIBEN
' ============================================================
Private Sub SchreibeKPI(ByVal ws As Worksheet, _
                         ByVal anzParzellen As Long, _
                         ByVal anzMitglieder As Long, _
                         ByVal summeIst As Double, _
                         ByVal summeSoll As Double, _
                         ByVal summeSaeumnis As Double, _
                         ByVal anzOffen As Long, _
                         ByVal anzBezahlt As Long, _
                         ByVal anzSaeumnis As Long, _
                         ByVal offenOhneSoll As Long, _
                         ByVal offenBetragKPI As Double)
    
    ' Karte 1: Parzellen & Mitglieder
    Call SchreibeKPIKarte(ws, DASH_KPI_LABEL_ROW, 1, 2, _
                          "PARZELLEN & MITGLIEDER", _
                          CStr(anzParzellen) & " Parzellen", _
                          CStr(anzMitglieder) & " Mitglieder aktiv", _
                          m_CLR_KPI_BLAU)
    
    ' Karte 2: Eingegangen
    Call SchreibeKPIKarte(ws, DASH_KPI_LABEL_ROW, 3, 4, _
                          "EINGEGANGEN", _
                          Format(summeIst, "#,##0.00") & " " & ChrW(8364), _
                          CStr(anzBezahlt) & " Posten bezahlt", _
                          m_CLR_KPI_GRUEN)
    
    ' Karte 3: Offen (v5.3: direkt akkumulierter Betrag statt Global-Differenz)
    Dim offenBetrag As Double
    offenBetrag = offenBetragKPI
    
    Dim offenDetail As String
    offenDetail = CStr(anzOffen) & " Posten offen"
    If offenOhneSoll > 0 Then
        offenDetail = offenDetail & " (davon " & offenOhneSoll & " ohne Soll)"
    End If
    
    Call SchreibeKPIKarte(ws, DASH_KPI_LABEL_ROW, 5, 6, _
                          "OFFEN", _
                          Format(offenBetrag, "#,##0.00") & " " & ChrW(8364), _
                          offenDetail, _
                          m_CLR_KPI_ROT)
    
    ' Karte 4: Saeumnis
    Call SchreibeKPIKarte(ws, DASH_KPI_LABEL_ROW, 7, 8, _
                          "S" & ChrW(196) & "UMNIS", _
                          Format(summeSaeumnis, "#,##0.00") & " " & ChrW(8364), _
                          CStr(anzSaeumnis) & " Vorkommen", _
                          m_CLR_KPI_ORANGE)
    
    ws.Rows(8).RowHeight = 4
    
    Dim matrixTitelSpalte As Long
    matrixTitelSpalte = ws.Cells(DASH_TITEL_ROW, 1).MergeArea.Columns.count
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
'  v5.4: PARZELLEN MIT MITGLIEDERLISTE-NAMEN ERGAENZEN
'  Fuegt Mitglieder aus der Mitgliederliste hinzu, die keinen
'  eigenen EntityKey haben (z.B. Partner zahlt Kosten mit).
'  Vergleicht per InStr (case-insensitive) um Doppeleintraege
'  zu vermeiden - nutzt Nachname als Schluesselwort.
' ============================================================
Private Sub ErgaenzeParzellennamen(ByRef parzellen() As ParzelleInfo, _
                                    ByVal anzParz As Long, _
                                    ByVal mitgliederML As Collection)
    
    Dim mlM As Object
    For Each mlM In mitgliederML
        Dim pNr As Long
        pNr = CLng(mlM("Parzelle"))
        
        ' Passende Parzelle im Array finden
        Dim p As Long
        For p = 1 To anzParz
            If parzellen(p).parzNr = pNr Then
                Dim mlName As String
                mlName = CStr(mlM("Name"))
                
                ' Pruefen ob Name schon enthalten ist
                ' (auch Teilstring-Match, da EntityKey-Tabelle ggf.
                '  Kontoname hat und Mitgliederliste Vorname+Nachname)
                Dim bereitsVorhanden As Boolean
                bereitsVorhanden = False
                
                ' Nachname extrahieren (letztes Wort) fuer robusteren Vergleich
                Dim nameParts() As String
                nameParts = Split(mlName, " ")
                Dim nachname As String
                nachname = nameParts(UBound(nameParts))
                
                If InStr(1, parzellen(p).mitgliedNamen, nachname, vbTextCompare) > 0 Then
                    bereitsVorhanden = True
                End If
                
                If Not bereitsVorhanden Then
                    parzellen(p).mitgliedNamen = parzellen(p).mitgliedNamen & vbLf & mlName
                    parzellen(p).anzMitglieder = parzellen(p).anzMitglieder + 1
                    Debug.Print "[Dashboard] Mitglied ergaenzt: " & mlName & _
                                " auf Parzelle " & pNr & " (ohne eigenen EntityKey)"
                End If
                
                Exit For
            End If
        Next p
    Next mlM
    
End Sub





























