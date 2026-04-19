Attribute VB_Name = "mod_Startseite"
Option Explicit

' ===============================================================
' MODUL: mod_Startseite
' VERSION: 2.0 - 18.04.2026
' ZWECK: Startseite (Startmenue) als professioneller Eye-Catcher
'        - Gradient-Look mit Farbverlauf-Effekt
'        - KPIs: Abrechnungsjahr, Mitglieder, Parzellen, Kontostand
'        - Navigations-Buttons in Kachel-Optik
'        - Serienbrief-Platzhalter
' ===============================================================

' --- Farben (Modernes Blau-Schema mit Akzenten) ---
Private Const CLR_HERO_DARK As Long = 2763306    ' RGB(26, 35, 42) - Hero-Banner dunkel
Private Const CLR_HERO_MED As Long = 4735033     ' RGB(41, 50, 72) - Hero-Banner mittel
Private Const CLR_ACCENT As Long = 14521384      ' RGB(40, 167, 221) - Akzent-Tuerkis
Private Const CLR_KPI_BG As Long = 16119285      ' RGB(245, 246, 250) - KPI Hintergrund
Private Const CLR_KPI_BORDER As Long = 14408667  ' RGB(219, 223, 219) - KPI Rahmen
Private Const CLR_BTN_FINANCE As Long = 11948081 ' RGB(41, 128, 182) - Finanzen-Blau
Private Const CLR_BTN_METER As Long = 7168108    ' RGB(108, 117, 109) - Zaehler Grau-Gruen
Private Const CLR_BTN_ADMIN As Long = 6260068    ' RGB(100, 120, 95) - Verwaltung
Private Const CLR_BTN_SERIENBR As Long = 5202271 ' RGB(95, 110, 79) - Gedaempftes Gruen
Private Const CLR_BTN_MITGL As Long = 5408340    ' RGB(52, 152, 82) - Mitglieder Gruen
Private Const CLR_WHITE As Long = 16777215
Private Const CLR_DARK_TEXT As Long = 2500134     ' RGB(38, 50, 56)
Private Const CLR_LIGHT_TEXT As Long = 12632256   ' RGB(192, 192, 192)
Private Const CLR_BG As Long = 16777215           ' Weiss
Private Const CLR_SECTION_BG As Long = 15921906   ' RGB(242, 242, 242) - Sections
Private Const CLR_DIVIDER As Long = 14408667      ' RGB(219, 223, 219)


' ===============================================================
' HAUPTPROZEDUR: Startseite initialisieren und gestalten
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
    
    Call EntferneAlteShapes(ws)
    Call VorbereiteBlatt(ws)
    Call SchreibeHeroBanner(ws)
    Call SchreibeKPIBereich(ws)
    Call ErstelleNavigationsKacheln(ws)
    Call SchreibeFooter(ws)
    
    ws.Cells.Locked = True
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


' ===============================================================
' ALTE SHAPES ENTFERNEN
' ===============================================================
Private Sub EntferneAlteShapes(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim i As Long
    
    For i = ws.Shapes.count To 1 Step -1
        Set shp = ws.Shapes(i)
        If shp.Type <> msoOLEControlObject Then
            shp.Delete
        End If
    Next i
    
    On Error Resume Next
    ws.OLEObjects("cmdMitgliederverwaltung").Delete
    Err.Clear
    On Error GoTo 0
End Sub


' ===============================================================
' BLATT VORBEREITEN
' ===============================================================
Private Sub VorbereiteBlatt(ByVal ws As Worksheet)
    ws.Range("A1:P55").ClearContents
    ws.Range("A1:P55").ClearFormats
    ws.Range("A1:P55").Interior.color = CLR_BG
    
    ' Gitternetzlinien aus
    Dim wnd As Window
    For Each wnd In Application.Windows
        If wnd.Caption = ThisWorkbook.Name Then
            wnd.DisplayGridlines = False
        End If
    Next wnd
    
    ' Spaltenbreiten (sauberes KPI-Layout)
    ws.Columns("A").ColumnWidth = 2      ' Rand
    ws.Columns("B").ColumnWidth = 4      ' Padding links
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 4      ' Trennspalte (leer)
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 14
    ws.Columns("J").ColumnWidth = 14
    ws.Columns("K").ColumnWidth = 4      ' Padding rechts
    ws.Columns("L").ColumnWidth = 2      ' Rand
    
    ' Zeilenhoehen
    ws.Rows("1").RowHeight = 6           ' Top-Rand
    ws.Rows("2").RowHeight = 50          ' Hero Titel
    ws.Rows("3").RowHeight = 24          ' Hero Untertitel
    ws.Rows("4").RowHeight = 30          ' Hero Vereinsname+Jahr
    ws.Rows("5").RowHeight = 4           ' Akzentlinie
    ws.Rows("6").RowHeight = 10          ' Abstand
    ws.Rows("7").RowHeight = 18          ' KPI-Header
    ws.Rows("8").RowHeight = 48          ' KPI-Werte Zeile 1
    ws.Rows("9").RowHeight = 18          ' KPI-Labels Zeile 1
    ws.Rows("10").RowHeight = 48         ' KPI-Werte Zeile 2 (Kontostand)
    ws.Rows("11").RowHeight = 18         ' KPI-Labels Zeile 2
    ws.Rows("12").RowHeight = 16         ' Abstand
    ws.Rows("13").RowHeight = 26         ' Sections-Header "Navigation"
    ws.Rows("14").RowHeight = 8          ' Abstand
    
    Dim r As Long
    For r = 15 To 18
        ws.Rows(r).RowHeight = 42        ' Button-Zeilen (4 Reihen Navigation)
    Next r
    
    ws.Rows("19").RowHeight = 26         ' Section-Header "Serienbrief"
    ws.Rows("20").RowHeight = 8          ' Abstand
    ws.Rows("21").RowHeight = 42         ' Serienbrief-Buttons
    ws.Rows("22").RowHeight = 16         ' Abstand
    ws.Rows("23").RowHeight = 20         ' Footer
End Sub


' ===============================================================
' HERO-BANNER: Titel, Untertitel, Akzentlinie
' ===============================================================
Private Sub SchreibeHeroBanner(ByVal ws As Worksheet)
    ' Hero-Block (Zeilen 2-4) - dunkler Hintergrund
    With ws.Range("A2:L4")
        .Interior.color = CLR_HERO_DARK
    End With
    
    ' Titel: "KASSENBUCH" gross und auffaellig
    With ws.Range("B2:K2")
        .Merge
        .value = ChrW(9733) & "  K A S S E N B U C H  " & ChrW(9733)
        .Font.Size = 24
        .Font.Bold = True
        .Font.color = CLR_WHITE
        .Interior.color = CLR_HERO_DARK
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Vereinsname prominent anzeigen
    Dim vereinsname As String
    Dim abrJahr As Long
    vereinsname = HoleVereinsname()
    abrJahr = HoleAbrechnungsjahr()
    
    Dim vereinsZeile As String
    If vereinsname <> "" Then
        vereinsZeile = vereinsname
    Else
        vereinsZeile = "Dein Kleingartenverein"
    End If
    
    With ws.Range("B3:K3")
        .Merge
        .value = vereinsZeile
        .Font.Size = 13
        .Font.Bold = True
        .Font.color = CLR_ACCENT
        .Interior.color = CLR_HERO_DARK
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Adresse + Abrechnungsjahr
    Dim adressZeile As String
    Dim strasse As String
    Dim plz As String
    Dim ort As String
    
    strasse = HoleVereinsStrasse()
    plz = HoleVereinsPLZ()
    ort = HoleVereinsOrt()
    
    adressZeile = ""
    If strasse <> "" Then adressZeile = strasse
    If plz <> "" Or ort <> "" Then
        If adressZeile <> "" Then adressZeile = adressZeile & " " & ChrW(8226) & " "
        If plz <> "" Then adressZeile = adressZeile & plz & " "
        If ort <> "" Then adressZeile = adressZeile & ort
    End If
    If abrJahr > 0 Then
        If adressZeile <> "" Then adressZeile = adressZeile & "  |  "
        adressZeile = adressZeile & "Abrechnungsjahr " & abrJahr
    End If
    If adressZeile = "" Then adressZeile = "Finanzverwaltung " & ChrW(8226) & " " & ChrW(220) & "bersicht"
    
    With ws.Range("B4:K4")
        .Merge
        .value = adressZeile
        .Font.Size = 10
        .Font.Bold = False
        .Font.color = CLR_LIGHT_TEXT
        .Interior.color = CLR_HERO_DARK
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Akzentlinie (schmaler Tuerkis-Streifen)
    With ws.Range("A5:L5")
        .Interior.color = CLR_ACCENT
    End With
End Sub


' ===============================================================
' KPI-BEREICH: 4 Kennzahlen-Karten
' ===============================================================
Private Sub SchreibeKPIBereich(ByVal ws As Worksheet)
    ' Hintergrund KPI-Bereich (erweitert fuer 2 KPI-Zeilen)
    ws.Range("A6:L12").Interior.color = CLR_SECTION_BG
    
    ' KPI-Header
    With ws.Range("B7:J7")
        .Merge
        .value = ChrW(9473) & ChrW(9473) & "  KENNZAHLEN  " & ChrW(9473) & ChrW(9473)
        .Font.Size = 9
        .Font.Bold = True
        .Font.color = RGB(140, 140, 140)
        .Interior.color = CLR_SECTION_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' --- KPI-Zeile 1: Abrechnungsjahr, Mitglieder, Parzellen ---
    Dim abrJahr As Long
    abrJahr = HoleAbrechnungsjahr()
    Dim jahrText As String
    If abrJahr > 0 Then jahrText = CStr(abrJahr) Else jahrText = "---"
    Call SchreibeKPIKarte(ws, "C", "E", 8, 9, jahrText, "Abrechnungsjahr", RGB(41, 128, 185))
    
    ' Spalte F bleibt leer (Trennung)
    ws.Range("F8:F9").Interior.color = CLR_SECTION_BG
    
    Call SchreibeKPIKarte(ws, "G", "H", 8, 9, CStr(ZaehleMitglieder()), "Mitglieder", RGB(39, 174, 96))
    
    Call SchreibeKPIKarte(ws, "I", "J", 8, 9, CStr(ZaehleBelegteParzellen()), "Parzellen", RGB(142, 68, 173))
    
    ' --- KPI-Zeile 2: Kontostand Vorjahr + Aktuell ---
    Dim kontoVorjahr As Double
    kontoVorjahr = HoleKontostandVorjahr()
    Dim vorjahrText As String
    vorjahrText = Format$(kontoVorjahr, "#,##0.00") & " " & ChrW(8364)
    
    Dim vorjahrFarbe As Long
    If kontoVorjahr >= 0 Then vorjahrFarbe = RGB(41, 128, 185) Else vorjahrFarbe = RGB(231, 76, 60)
    
    Call SchreibeKPIKarte(ws, "C", "E", 10, 11, vorjahrText, "Kontostand Vorjahr", vorjahrFarbe)
    
    ' Spalte F bleibt leer (Trennung)
    ws.Range("F10:F11").Interior.color = CLR_SECTION_BG
    
    Dim kontostand As Double
    kontostand = HoleAktuellerKontostand()
    Dim kontoText As String
    kontoText = Format$(kontostand, "#,##0.00") & " " & ChrW(8364)
    
    Dim kontoFarbe As Long
    If kontostand >= 0 Then kontoFarbe = RGB(39, 174, 96) Else kontoFarbe = RGB(231, 76, 60)
    
    Call SchreibeKPIKarte(ws, "G", "J", 10, 11, kontoText, "Kontostand Aktuell | " & HoleLetztesBuchungsdatum(), kontoFarbe)
End Sub


' ===============================================================
' EINZEL-KPI-KARTE
' ===============================================================
Private Sub SchreibeKPIKarte(ByVal ws As Worksheet, _
                              ByVal col1 As String, _
                              ByVal col2 As String, _
                              ByVal wertZeile As Long, _
                              ByVal labelZeile As Long, _
                              ByVal wertText As String, _
                              ByVal label As String, _
                              ByVal akzentFarbe As Long)
    
    ' Wert-Zelle
    With ws.Range(col1 & wertZeile & ":" & col2 & wertZeile)
        .Merge
        .value = wertText
        .Font.Size = 16
        .Font.Bold = True
        .Font.color = CLR_DARK_TEXT
        .Interior.color = CLR_WHITE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        ' Rahmen
        .Borders(xlEdgeBottom).color = akzentFarbe
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeTop).color = CLR_KPI_BORDER
        .Borders(xlEdgeTop).Weight = xlHairline
        .Borders(xlEdgeLeft).color = CLR_KPI_BORDER
        .Borders(xlEdgeLeft).Weight = xlHairline
        .Borders(xlEdgeRight).color = CLR_KPI_BORDER
        .Borders(xlEdgeRight).Weight = xlHairline
    End With
    
    ' Label
    With ws.Range(col1 & labelZeile & ":" & col2 & labelZeile)
        .Merge
        .value = label
        .Font.Size = 8
        .Font.Bold = True
        .Font.color = RGB(120, 120, 120)
        .Interior.color = CLR_SECTION_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With
End Sub


' ===============================================================
' NAVIGATIONS-KACHELN: Gruppiert in 3 Spalten
' ===============================================================
Private Sub ErstelleNavigationsKacheln(ByVal ws As Worksheet)
    ' Sections-Header
    With ws.Range("B13:J13")
        .Merge
        .value = ChrW(9654) & "  NAVIGATION"
        .Font.Size = 10
        .Font.Bold = True
        .Font.color = CLR_WHITE
        .Interior.color = CLR_HERO_DARK
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    
    ' --- Positionen berechnen ---
    Dim col1Left As Double, col2Left As Double, col3Left As Double
    Dim kachelW As Double, kachelH As Double
    Dim gapY As Double
    
    col1Left = ws.Range("C15").Left
    col2Left = ws.Range("F15").Left
    col3Left = ws.Range("I15").Left
    kachelW = ws.Range("C15:D15").Width
    kachelH = 34
    gapY = ws.Rows("15").RowHeight
    
    ' --- Spalte 1: Finanzen ---
    Call ErstelleKachel(ws, "kachel_Uebersicht", _
        ChrW(9654) & " Zahlungs" & ChrW(252) & "bersicht", _
        col1Left, ws.Range("C15").Top + 4, kachelW, kachelH, _
        CLR_BTN_FINANCE, "'mod_Navigation.NavigiereZu_Uebersicht'")
    
    Call ErstelleKachel(ws, "kachel_Bankkonto", _
        ChrW(9733) & " Bankkonto", _
        col1Left, ws.Range("C16").Top + 4, kachelW, kachelH, _
        CLR_BTN_FINANCE, "'mod_Navigation.NavigiereZu_Bankkonto'")
    
    Call ErstelleKachel(ws, "kachel_Vereinskasse", _
        ChrW(9830) & " Vereinskasse", _
        col1Left, ws.Range("C17").Top + 4, kachelW, kachelH, _
        CLR_BTN_FINANCE, "'mod_Navigation.NavigiereZu_Vereinskasse'")
    
    ' --- Spalte 2: Verbrauch & Verwaltung ---
    Call ErstelleKachel(ws, "kachel_Dashboard", _
        ChrW(9650) & " Dashboard", _
        col2Left, ws.Range("F15").Top + 4, kachelW, kachelH, _
        CLR_BTN_FINANCE, "'mod_Navigation.NavigiereZu_Dashboard'")
    
    Call ErstelleKachel(ws, "kachel_Strom", _
        ChrW(9889) & " Strom", _
        col2Left, ws.Range("F16").Top + 4, kachelW, kachelH, _
        CLR_BTN_METER, "'mod_Navigation.NavigiereZu_Strom'")
    
    Call ErstelleKachel(ws, "kachel_Wasser", _
        ChrW(8776) & " Wasser", _
        col2Left, ws.Range("F17").Top + 4, kachelW, kachelH, _
        CLR_BTN_METER, "'mod_Navigation.NavigiereZu_Wasser'")
    
    ' --- Spalte 3: Admin ---
    Call ErstelleKachel(ws, "kachel_Einstellungen", _
        ChrW(9881) & " Einstellungen", _
        col3Left, ws.Range("I15").Top + 4, kachelW, kachelH, _
        CLR_BTN_ADMIN, "'mod_Navigation.NavigiereZu_Einstellungen'")
    
    Call ErstelleKachel(ws, "kachel_Daten", _
        ChrW(9632) & " Daten", _
        col3Left, ws.Range("I16").Top + 4, kachelW, kachelH, _
        CLR_BTN_ADMIN, "'mod_Navigation.NavigiereZu_Daten'")
    
    Call ErstelleKachel(ws, "kachel_Mitglieder", _
        ChrW(9679) & " Mitgliederverwaltung", _
        col3Left, ws.Range("I17").Top + 4, kachelW, kachelH, _
        CLR_BTN_MITGL, "'mod_Navigation.ZeigeMitgliederverwaltung'")
    
    ' --- Zeile 4: Finanz-Uebersicht ---
    Call ErstelleKachel(ws, "kachel_FinanzUebersicht", _
        ChrW(9654) & " Finanz-" & ChrW(220) & "bersicht", _
        col1Left, ws.Range("C18").Top + 4, kachelW, kachelH, _
        CLR_BTN_FINANCE, "'mod_Navigation.NavigiereZu_FinanzUebersicht'")
    
    ' --- Serienbrief-Bereich ---
    With ws.Range("B19:J19")
        .Merge
        .value = ChrW(9654) & "  SERIENBRIEF (Word-Dokumente)"
        .Font.Size = 10
        .Font.Bold = True
        .Font.color = CLR_WHITE
        .Interior.color = CLR_HERO_DARK
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    
    Call ErstelleKachel(ws, "kachel_Betriebskosten", _
        ChrW(9633) & " Betriebskostenabrechnung", _
        col1Left, ws.Range("C21").Top + 4, kachelW, kachelH, _
        CLR_BTN_SERIENBR, "'mod_Navigation.ZeigeSerienbrief_Betriebskosten'")
    
    Call ErstelleKachel(ws, "kachel_Endabrechnung", _
        ChrW(9633) & " Endabrechnung", _
        col2Left, ws.Range("F21").Top + 4, kachelW, kachelH, _
        CLR_BTN_SERIENBR, "'mod_Navigation.ZeigeSerienbrief_Endabrechnung'")
End Sub


' ===============================================================
' FOOTER: Versionsinfo und Hinweis
' ===============================================================
Private Sub SchreibeFooter(ByVal ws As Worksheet)
    With ws.Range("B23:J23")
        .Merge
        .value = "Kassenbuch v2.7  |  " & ChrW(169) & " " & Year(Date)
        .Font.Size = 8
        .Font.color = RGB(160, 160, 160)
        .Interior.color = CLR_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub


' ===============================================================
' KACHEL-BUTTON ERSTELLEN (modernes Design mit Schatten-Effekt)
' ===============================================================
Private Sub ErstelleKachel(ByVal ws As Worksheet, _
                            ByVal btnName As String, _
                            ByVal btnText As String, _
                            ByVal x As Double, _
                            ByVal y As Double, _
                            ByVal w As Double, _
                            ByVal h As Double, _
                            ByVal farbe As Long, _
                            ByVal makroName As String)
    
    On Error GoTo KachelErr
    
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    
    With shp
        .Name = btnName
        .Fill.ForeColor.RGB = farbe
        .Line.Visible = msoFalse
        
        ' Leicht abgerundete Ecken
        On Error Resume Next
        .Adjustments(1) = 0.18
        On Error GoTo KachelErr
        
        ' Dezenter Schatten fuer 3D-Effekt
        With .Shadow
            .Visible = msoTrue
            .Type = msoShadow14
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0.75
            .OffsetX = 2
            .OffsetY = 2
            .Blur = 4
        End With
        
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 10
            .MarginRight = 10
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
    
KachelErr:
    Debug.Print "[Startseite] Kachel '" & btnName & "' Fehler: " & Err.Description
    Err.Clear
End Sub


' ===============================================================
' HILFSFUNKTIONEN: KPI-Daten ermitteln
' ===============================================================

Public Function ZaehleMitglieder() As Long
    Dim wsMitgl As Worksheet
    Dim r As Long
    Dim lastRow As Long
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
        
        If Not IsDate(pachtAnfang) Then
            If Not IsNumeric(pachtAnfang) Then GoTo NextMitglied
        End If
        
        If IsDate(pachtEnde) Then
            If CDate(pachtEnde) < Date Then GoTo NextMitglied
        End If
        
        Dim anrede As String
        anrede = Trim(CStr(wsMitgl.Cells(r, M_COL_ANREDE).value))
        If anrede = ANREDE_KGA Then GoTo NextMitglied
        
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
        
        If Not IsDate(pAnfang) Then
            If Not IsNumeric(pAnfang) Then GoTo NextParzelle
        End If
        
        If IsDate(pEnde) Then
            If CDate(pEnde) < Date Then GoTo NextParzelle
        End If
        
        Dim anredeP As String
        anredeP = Trim(CStr(wsMitgl.Cells(r, M_COL_ANREDE).value))
        If anredeP = ANREDE_KGA Then GoTo NextParzelle
        
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


Public Function HoleKontostandVorjahr() As Double
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


Private Function HoleAktuellerKontostand() As Double
    Dim vorjahr As Double
    vorjahr = HoleKontostandVorjahr()
    
    Dim wsBK As Worksheet
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo 0
    
    If wsBK Is Nothing Then
        HoleAktuellerKontostand = vorjahr
        Exit Function
    End If
    
    Dim summe As Double
    On Error Resume Next
    summe = Application.WorksheetFunction.Sum(wsBK.Range("B" & BK_START_ROW & ":B5000"))
    On Error GoTo 0
    
    HoleAktuellerKontostand = vorjahr + summe
End Function


Private Function HoleLetztesBuchungsdatum() As String
    Dim wsBK As Worksheet
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo 0
    
    If wsBK Is Nothing Then
        HoleLetztesBuchungsdatum = "---"
        Exit Function
    End If
    
    On Error Resume Next
    Dim maxDatum As Double
    maxDatum = Application.WorksheetFunction.Max(wsBK.Range("A" & BK_START_ROW & ":A5000"))
    On Error GoTo 0
    
    If maxDatum > 0 Then
        HoleLetztesBuchungsdatum = "Stand: " & Format$(CDate(maxDatum), "dd.mm.yyyy")
    Else
        HoleLetztesBuchungsdatum = "keine Buchungen"
    End If
End Function


Private Function HoleVereinsStrasse() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    If ws Is Nothing Then HoleVereinsStrasse = "": Exit Function
    HoleVereinsStrasse = Trim(CStr(ws.Cells(ES_CFG_STRASSE_ROW, ES_CFG_VALUE_COL).value))
End Function


Private Function HoleVereinsPLZ() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    If ws Is Nothing Then HoleVereinsPLZ = "": Exit Function
    HoleVereinsPLZ = Trim(CStr(ws.Cells(ES_CFG_PLZ_ORT_ROW, ES_CFG_VALUE_COL).value))
End Function


Private Function HoleVereinsOrt() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    If ws Is Nothing Then HoleVereinsOrt = "": Exit Function
    HoleVereinsOrt = Trim(CStr(ws.Cells(ES_CFG_PLZ_ORT_ROW, 5).value))
End Function




















