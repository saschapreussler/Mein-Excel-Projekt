Attribute VB_Name = "mod_Startseite"
Option Explicit

' ===============================================================
' MODUL: mod_Startseite
' VERSION: 1.0 - 18.04.2026
' ZWECK: Startseite (Startmenue) als Eye-Catcher gestalten
'        - KPIs: Abrechnungsjahr, Mitglieder, Parzellen, Kontostand
'        - Navigations-Buttons zu allen relevanten Blaettern
'        - Blau-Grau Farbschema (professionell)
' ===============================================================

' --- Farben (Blau-Grau Schema) ---
Private Const CLR_HEADER As Long = 2894892      ' RGB(44, 62, 80) - Dunkles Blau-Grau
Private Const CLR_SECTION As Long = 6182740      ' RGB(52, 73, 94) - Mittleres Blau-Grau
Private Const CLR_KPI_BG As Long = 15853804      ' RGB(236, 240, 241) - Helles Grau
Private Const CLR_BTN_NAV As Long = 12161833     ' RGB(41, 128, 185) - Blau Akzent
Private Const CLR_BTN_SERIENBR As Long = 6723942 ' RGB(230, 126, 102) - Gedaempftes Orange
Private Const CLR_BTN_MITGL As Long = 5408340    ' RGB(52, 152, 82) - Gruen
Private Const CLR_WHITE As Long = 16777215       ' RGB(255, 255, 255)
Private Const CLR_DARK_TEXT As Long = 2500134     ' RGB(38, 50, 56) - Fast Schwarz
Private Const CLR_BG As Long = 16448250          ' RGB(250, 250, 250) - Hintergrund

' --- Layout-Konstanten ---
Private Const START_COL_LEFT As Long = 2         ' Spalte B
Private Const START_COL_RIGHT As Long = 9        ' Spalte I (rechte Grenze)


' ===============================================================
' HAUPTPROZEDUR: Startseite initialisieren und gestalten
' Wird bei Workbook_Open aufgerufen
' ===============================================================
Public Sub InitialisiereStartseite()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_STARTMENUE())
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' Bestehende Shapes entfernen (ausser ActiveX cmdMitgliederverwaltung)
    Call EntferneAlteShapes(ws)
    
    ' Blatt vorbereiten
    Call VorbereiteBlatt(ws)
    
    ' Header-Bereich (Titel + Vereinsname)
    Call SchreibeHeader(ws)
    
    ' KPI-Karten
    Call SchreibeKPIs(ws)
    
    ' Navigations-Buttons
    Call ErstelleNavigationsButtons(ws)
    
    ' Blattschutz
    ws.Cells.Locked = True
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


' ===============================================================
' ALTE SHAPES ENTFERNEN (ausser ActiveX-Controls)
' ===============================================================
Private Sub EntferneAlteShapes(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim i As Long
    
    ' Rueckwaerts durchlaufen um Indexprobleme zu vermeiden
    For i = ws.Shapes.count To 1 Step -1
        Set shp = ws.Shapes(i)
        ' ActiveX-Controls (OLEObjects) nicht loeschen
        If shp.Type <> msoOLEControlObject Then
            shp.Delete
        End If
    Next i
    
    ' ActiveX cmdMitgliederverwaltung entfernen (wird als Shape neu erstellt)
    On Error Resume Next
    ws.OLEObjects("cmdMitgliederverwaltung").Delete
    Err.Clear
    On Error GoTo 0
End Sub


' ===============================================================
' BLATT VORBEREITEN: Zellen loeschen, Hintergrund, Spaltenbreiten
' ===============================================================
Private Sub VorbereiteBlatt(ByVal ws As Worksheet)
    ' Inhalte loeschen (Zeilen 1-40)
    ws.Range("A1:N40").ClearContents
    ws.Range("A1:N40").ClearFormats
    
    ' Hintergrundfarbe
    ws.Range("A1:N40").Interior.color = CLR_BG
    
    ' Gitternetzlinien ausblenden
    Dim wnd As Window
    For Each wnd In Application.Windows
        If wnd.Caption = ThisWorkbook.Name Then
            wnd.DisplayGridlines = False
        End If
    Next wnd
    
    ' Spaltenbreiten setzen
    ws.Columns("A").ColumnWidth = 3      ' Rand links
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 4      ' Abstand
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 3      ' Rand rechts
    
    ' Zeilenhoehen
    ws.Rows("1").RowHeight = 8           ' Rand oben
    ws.Rows("2").RowHeight = 40          ' Titel
    ws.Rows("3").RowHeight = 22          ' Untertitel
    ws.Rows("4").RowHeight = 8           ' Abstand
    ws.Rows("5").RowHeight = 18          ' KPI-Header
    ws.Rows("6").RowHeight = 50          ' KPI-Werte
    ws.Rows("7").RowHeight = 18          ' KPI-Labels
    ws.Rows("8").RowHeight = 12          ' Abstand
    
    Dim r As Long
    For r = 9 To 25
        ws.Rows(r).RowHeight = 38        ' Button-Zeilen
    Next r
End Sub


' ===============================================================
' HEADER: Titel und Vereinsname
' ===============================================================
Private Sub SchreibeHeader(ByVal ws As Worksheet)
    ' Titel-Zeile
    With ws.Range("B2:H2")
        .Merge
        .value = "Kassenbuch"
        .Font.Size = 22
        .Font.Bold = True
        .Font.color = CLR_WHITE
        .Interior.color = CLR_HEADER
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Untertitel mit Vereinsname und Abrechnungsjahr
    Dim untertitel As String
    Dim vereinsname As String
    Dim abrJahr As Long
    
    vereinsname = HoleVereinsname()
    abrJahr = HoleAbrechnungsjahr()
    
    If vereinsname <> "" Then
        untertitel = vereinsname
    Else
        untertitel = "Kleingartenverein"
    End If
    
    If abrJahr > 0 Then
        untertitel = untertitel & " - Abrechnungsjahr " & abrJahr
    End If
    
    With ws.Range("B3:H3")
        .Merge
        .value = untertitel
        .Font.Size = 12
        .Font.Italic = True
        .Font.color = CLR_WHITE
        .Interior.color = CLR_SECTION
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub


' ===============================================================
' KPI-KARTEN: Mitglieder, Parzellen, Kontostand
' ===============================================================
Private Sub SchreibeKPIs(ByVal ws As Worksheet)
    ' KPI-Header-Zeile
    With ws.Range("B5:H5")
        .Interior.color = CLR_BG
    End With
    
    ' --- KPI 1: Mitglieder ---
    Call SchreibeEinzelKPI(ws, "B", "C", ZaehleMitglieder(), "Mitglieder")
    
    ' --- KPI 2: Belegte Parzellen ---
    Call SchreibeEinzelKPI(ws, "D", "E", ZaehleBelegteParzellen(), "Parzellen belegt")
    
    ' --- KPI 3: Kontostand Vorjahr ---
    Dim kontostand As Double
    kontostand = HoleKontostandVorjahr()
    
    Dim kontoText As String
    If kontostand = 0 Then
        kontoText = "---"
    Else
        kontoText = Format$(kontostand, "#,##0.00") & " " & ChrW(8364)
    End If
    
    ' KPI 3 ueber Spalten F-H (breiter fuer Waehrung)
    With ws.Range("F6:H6")
        .Merge
        .value = kontoText
        .Font.Size = 18
        .Font.Bold = True
        .Font.color = CLR_DARK_TEXT
        .Interior.color = CLR_KPI_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.color = RGB(189, 195, 199)
        .Borders.Weight = xlThin
    End With
    
    With ws.Range("F7:H7")
        .Merge
        .value = "Kontostand Vorjahr"
        .Font.Size = 9
        .Font.color = RGB(127, 140, 141)
        .Interior.color = CLR_KPI_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With
End Sub


' ===============================================================
' EINZEL-KPI: Formatierte KPI-Karte schreiben
' ===============================================================
Private Sub SchreibeEinzelKPI(ByVal ws As Worksheet, _
                                ByVal spalte1 As String, _
                                ByVal spalte2 As String, _
                                ByVal wert As Long, _
                                ByVal beschreibung As String)
    
    ' Wert (grosse Zahl)
    With ws.Range(spalte1 & "6:" & spalte2 & "6")
        .Merge
        .value = wert
        .Font.Size = 24
        .Font.Bold = True
        .Font.color = CLR_DARK_TEXT
        .Interior.color = CLR_KPI_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "0"
        .Borders.color = RGB(189, 195, 199)
        .Borders.Weight = xlThin
    End With
    
    ' Beschreibung
    With ws.Range(spalte1 & "7:" & spalte2 & "7")
        .Merge
        .value = beschreibung
        .Font.Size = 9
        .Font.color = RGB(127, 140, 141)
        .Interior.color = CLR_KPI_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With
End Sub


' ===============================================================
' NAVIGATIONS-BUTTONS ERSTELLEN
' ===============================================================
Private Sub ErstelleNavigationsButtons(ByVal ws As Worksheet)
    ' Section-Header: Navigation
    With ws.Range("B9:H9")
        .Merge
        .value = "Navigation"
        .Font.Size = 11
        .Font.Bold = True
        .Font.color = CLR_WHITE
        .Interior.color = CLR_SECTION
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Buttons in 2 Spalten anordnen
    ' Linke Spalte: B-D (Left ~ 22, Width ~ 190)
    ' Rechte Spalte: F-H (Left ~ 280, Width ~ 190)
    
    Dim leftX As Double
    Dim rightX As Double
    Dim btnW As Double
    Dim btnH As Double
    Dim rowH As Double
    Dim startTop As Double
    
    ' Positionen aus Zellgeometrie berechnen
    leftX = ws.Range("B10").Left + 4
    rightX = ws.Range("F10").Left + 4
    btnW = ws.Range("B10:D10").Width - 8
    btnH = 30
    startTop = ws.Range("B10").Top + 4
    rowH = ws.Rows("10").RowHeight
    
    ' --- Linke Spalte ---
    Call ErstelleButton(ws, "btn_Uebersicht", ChrW(128202) & " Zahlungs" & ChrW(252) & "bersicht", _
                        leftX, startTop, btnW, btnH, CLR_BTN_NAV, _
                        "'mod_Navigation.NavigiereZu_Uebersicht'")
    
    Call ErstelleButton(ws, "btn_Bankkonto", ChrW(127974) & " Bankkonto", _
                        leftX, startTop + rowH, btnW, btnH, CLR_BTN_NAV, _
                        "'mod_Navigation.NavigiereZu_Bankkonto'")
    
    Call ErstelleButton(ws, "btn_Strom", ChrW(9889) & " Strom", _
                        leftX, startTop + rowH * 2, btnW, btnH, CLR_BTN_NAV, _
                        "'mod_Navigation.NavigiereZu_Strom'")
    
    Call ErstelleButton(ws, "btn_Einstellungen", ChrW(9881) & " Einstellungen", _
                        leftX, startTop + rowH * 3, btnW, btnH, CLR_BTN_NAV, _
                        "'mod_Navigation.NavigiereZu_Einstellungen'")
    
    Call ErstelleButton(ws, "btn_Mitgliederverwaltung", ChrW(128101) & " Mitgliederverwaltung", _
                        leftX, startTop + rowH * 4, btnW, btnH, CLR_BTN_MITGL, _
                        "'mod_Navigation.ZeigeMitgliederverwaltung'")
    
    ' --- Rechte Spalte ---
    Call ErstelleButton(ws, "btn_Dashboard", ChrW(128200) & " Dashboard", _
                        rightX, startTop, btnW, btnH, CLR_BTN_NAV, _
                        "'mod_Navigation.NavigiereZu_Dashboard'")
    
    Call ErstelleButton(ws, "btn_Vereinskasse", ChrW(128176) & " Vereinskasse", _
                        rightX, startTop + rowH, btnW, btnH, CLR_BTN_NAV, _
                        "'mod_Navigation.NavigiereZu_Vereinskasse'")
    
    Call ErstelleButton(ws, "btn_Wasser", ChrW(128167) & " Wasser", _
                        rightX, startTop + rowH * 2, btnW, btnH, CLR_BTN_NAV, _
                        "'mod_Navigation.NavigiereZu_Wasser'")
    
    Call ErstelleButton(ws, "btn_Daten", ChrW(128451) & " Daten", _
                        rightX, startTop + rowH * 3, btnW, btnH, CLR_BTN_NAV, _
                        "'mod_Navigation.NavigiereZu_Daten'")
    
    ' --- Serienbrief-Bereich ---
    With ws.Range("B16:H16")
        .Merge
        .value = "Serienbrief (Word-Dokumente)"
        .Font.Size = 11
        .Font.Bold = True
        .Font.color = CLR_WHITE
        .Interior.color = CLR_SECTION
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Dim sbTop As Double
    sbTop = ws.Range("B17").Top + 4
    
    Call ErstelleButton(ws, "btn_Betriebskosten", ChrW(128196) & " Betriebskostenabrechnung", _
                        leftX, sbTop, btnW, btnH, CLR_BTN_SERIENBR, _
                        "'mod_Navigation.ZeigeSerienbrief_Betriebskosten'")
    
    Call ErstelleButton(ws, "btn_Endabrechnung", ChrW(128196) & " Endabrechnung", _
                        rightX, sbTop, btnW, btnH, CLR_BTN_SERIENBR, _
                        "'mod_Navigation.ZeigeSerienbrief_Endabrechnung'")
End Sub


' ===============================================================
' BUTTON ERSTELLEN: Gerundetes Rechteck mit Text
' ===============================================================
Private Sub ErstelleButton(ByVal ws As Worksheet, _
                            ByVal btnName As String, _
                            ByVal btnText As String, _
                            ByVal x As Double, _
                            ByVal y As Double, _
                            ByVal w As Double, _
                            ByVal h As Double, _
                            ByVal farbe As Long, _
                            ByVal makroName As String)
    
    On Error GoTo BtnErr
    
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    
    With shp
        .Name = btnName
        .Fill.ForeColor.RGB = farbe
        .Line.Visible = msoFalse
        
        ' Abgerundete Ecken etwas staerker
        On Error Resume Next
        .Adjustments(1) = 0.25
        On Error GoTo BtnErr
        
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 8
            .MarginRight = 8
            .MarginTop = 2
            .MarginBottom = 2
            .WordWrap = msoFalse
            
            With .TextRange
                .text = btnText
                .Font.Fill.ForeColor.RGB = CLR_WHITE
                .Font.Size = 11
                .Font.Bold = msoTrue
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
        
        .OnAction = makroName
        .Placement = xlFreeFloating
    End With
    
    Exit Sub
    
BtnErr:
    Debug.Print "[Startseite] Button '" & btnName & "' Fehler: " & Err.Description
    Err.Clear
End Sub


' ===============================================================
' HILFSFUNKTIONEN: KPI-Daten ermitteln
' ===============================================================

' Zaehlt aktive Mitglieder (nicht redundant) aus Mitgliederliste
Public Function ZaehleMitglieder() As Long
    Dim wsMitgl As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim anzahl As Long
    Dim dictNamen As Object
    
    On Error Resume Next
    Set wsMitgl = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    On Error GoTo 0
    
    If wsMitgl Is Nothing Then
        ZaehleMitglieder = 0
        Exit Function
    End If
    
    Set dictNamen = CreateObject("Scripting.Dictionary")
    lastRow = wsMitgl.Cells(wsMitgl.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        Dim nachname As String
        Dim vorname As String
        Dim pachtAnfang As Variant
        Dim pachtEnde As Variant
        
        nachname = Trim(CStr(wsMitgl.Cells(r, M_COL_NACHNAME).value))
        vorname = Trim(CStr(wsMitgl.Cells(r, M_COL_VORNAME).value))
        pachtAnfang = wsMitgl.Cells(r, M_COL_PACHTANFANG).value
        pachtEnde = wsMitgl.Cells(r, M_COL_PACHTENDE).value
        
        If nachname = "" Then GoTo NextMitglied
        
        ' Nur aktive Mitglieder (Pachtanfang vorhanden, Pachtende leer oder in Zukunft)
        If Not IsDate(pachtAnfang) Then
            If Not IsNumeric(pachtAnfang) Then GoTo NextMitglied
        End If
        
        If IsDate(pachtEnde) Then
            If CDate(pachtEnde) < Date Then GoTo NextMitglied
        End If
        
        ' KGA-Eintraege ignorieren
        Dim anrede As String
        anrede = Trim(CStr(wsMitgl.Cells(r, M_COL_ANREDE).value))
        If anrede = ANREDE_KGA Then GoTo NextMitglied
        
        ' Eindeutigkeit per Name
        Dim schluessel As String
        schluessel = LCase(nachname & "|" & vorname)
        If Not dictNamen.Exists(schluessel) Then
            dictNamen.Add schluessel, True
        End If
        
NextMitglied:
    Next r
    
    ZaehleMitglieder = dictNamen.count
    Set dictNamen = Nothing
End Function


' Zaehlt belegte (verpachtete) Parzellen aus Mitgliederliste
Public Function ZaehleBelegteParzellen() As Long
    Dim wsMitgl As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim dictParz As Object
    
    On Error Resume Next
    Set wsMitgl = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    On Error GoTo 0
    
    If wsMitgl Is Nothing Then
        ZaehleBelegteParzellen = 0
        Exit Function
    End If
    
    Set dictParz = CreateObject("Scripting.Dictionary")
    lastRow = wsMitgl.Cells(wsMitgl.Rows.count, M_COL_PARZELLE).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        Dim parzelle As String
        Dim pAnfang As Variant
        Dim pEnde As Variant
        
        parzelle = Trim(CStr(wsMitgl.Cells(r, M_COL_PARZELLE).value))
        pAnfang = wsMitgl.Cells(r, M_COL_PACHTANFANG).value
        pEnde = wsMitgl.Cells(r, M_COL_PACHTENDE).value
        
        If parzelle = "" Then GoTo NextParzelle
        If parzelle = PARZELLE_VEREIN Then GoTo NextParzelle
        
        ' Nur aktive Pacht
        If Not IsDate(pAnfang) Then
            If Not IsNumeric(pAnfang) Then GoTo NextParzelle
        End If
        
        If IsDate(pEnde) Then
            If CDate(pEnde) < Date Then GoTo NextParzelle
        End If
        
        ' KGA ignorieren
        Dim anredeP As String
        anredeP = Trim(CStr(wsMitgl.Cells(r, M_COL_ANREDE).value))
        If anredeP = ANREDE_KGA Then GoTo NextParzelle
        
        ' Parzellennummer normalisieren
        Dim parzNorm As String
        parzNorm = LCase(Trim(parzelle))
        If Not dictParz.Exists(parzNorm) Then
            dictParz.Add parzNorm, True
        End If
        
NextParzelle:
    Next r
    
    ZaehleBelegteParzellen = dictParz.count
    Set dictParz = Nothing
End Function


' Vereinsname aus Einstellungen lesen
Private Function HoleVereinsname() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    
    If ws Is Nothing Then
        HoleVereinsname = ""
        Exit Function
    End If
    
    HoleVereinsname = Trim(CStr(ws.Cells(ES_CFG_VEREINSNAME_ROW, ES_CFG_VALUE_COL).value))
End Function


' Kontostand Vorjahr aus Einstellungen lesen
Private Function HoleKontostandVorjahr() As Double
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    
    If ws Is Nothing Then
        HoleKontostandVorjahr = 0
        Exit Function
    End If
    
    Dim wert As Variant
    wert = ws.Cells(ES_CFG_KONTOSTAND_ROW, ES_CFG_VALUE_COL).value
    If IsNumeric(wert) And wert <> "" Then
        HoleKontostandVorjahr = CDbl(wert)
    Else
        HoleKontostandVorjahr = 0
    End If
End Function
