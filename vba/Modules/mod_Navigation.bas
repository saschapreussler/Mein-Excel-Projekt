Attribute VB_Name = "mod_Navigation"
Option Explicit

' ===============================================================
' MODUL: mod_Navigation
' VERSION: 1.0 - 18.04.2026
' ZWECK: Navigation zwischen Tabellenblaettern
'        - Startseite -> alle Blätter (Button-Handler)
'        - Alle Blätter -> Startseite (Home-Button)
'        - Home-Buttons auf allen Blättern erstellen/entfernen
' ===============================================================

Private Const HOME_BTN_NAME As String = "btn_Home"


' ===============================================================
' NAVIGATION: Startseite aktivieren
' ===============================================================
Public Sub NavigiereZuStartseite()
    Dim ws As Worksheet
    Set ws = FindeTabellenblattRobust(WS_STARTMENUE(), "Startseite")
    AktiviereZielblattStabil ws
End Sub


' ===============================================================
' NAVIGATION: Einzelne Blätter aktivieren (Button-Handler)
' ===============================================================
Public Sub NavigiereZu_Bankkonto()
    AktiviereTabellenblatt WS_BANKKONTO
End Sub

Public Sub NavigiereZu_Einstellungen()
    AktiviereTabellenblatt WS_EINSTELLUNGEN
End Sub

Public Sub NavigiereZu_Vereinskasse()
    AktiviereTabellenblatt WS_VEREINSKASSE
End Sub

Public Sub NavigiereZu_Strom()
    AktiviereTabellenblatt "Strom"
End Sub

Public Sub NavigiereZu_Wasser()
    AktiviereTabellenblatt "Wasser"
End Sub

Public Sub NavigiereZu_Daten()
    AktiviereTabellenblatt WS_DATEN
End Sub

Public Sub NavigiereZu_FinanzUebersicht()
    ' Blatt erstellen falls nicht vorhanden, dann aktivieren
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_FINANZ_UEBERSICHT())
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Blatt wird beim ersten Aufruf erstellt
        mod_FinanzUebersicht.ErstelleFinanzUebersicht
    Else
        AktiviereZielblattStabil ws
    End If
End Sub

Public Sub NavigiereZu_Uebersicht()
    Dim wsUeb As Worksheet
    Set wsUeb = FindeTabellenblattRobust(WS_UEBERSICHT(), "Uebersicht")

    If wsUeb Is Nothing Then
        On Error Resume Next
        Call mod_Uebersicht_Generator.GeneriereUebersicht
        On Error GoTo 0
        Set wsUeb = FindeTabellenblattRobust(WS_UEBERSICHT(), "Uebersicht")
    End If

    If wsUeb Is Nothing Then
        MsgBox "Zahlungs" & ChrW(252) & "bersicht konnte nicht gefunden oder erzeugt werden.", _
               vbExclamation, "Navigation"
        Exit Sub
    End If

    AktiviereZielblattStabil wsUeb
End Sub

Public Sub NavigiereZu_Dashboard()
    ' Dashboard wird dynamisch erzeugt - Name kann variieren
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = FindeTabellenblattRobust("Dashboard Mitgliederzahlungen", "Dashboard")
    On Error GoTo 0
    
    If ws Is Nothing Then
        On Error Resume Next
        Call mod_Uebersicht_Dashboard.GeneriereUebersichtNeu(True)
        On Error GoTo 0
        Set ws = FindeTabellenblattRobust("Dashboard Mitgliederzahlungen", "Dashboard")
        If ws Is Nothing Then
            MsgBox "Das Dashboard wurde noch nicht erstellt." & vbLf & vbLf & _
                   "Bitte zuerst die Zahlungs" & ChrW(252) & "bersicht " & _
                   "oder das Dashboard generieren.", _
                   vbInformation, "Dashboard nicht vorhanden"
            Exit Sub
        End If
    End If
    
    AktiviereZielblattStabil ws
End Sub

' ===============================================================
' Failsafe gegen vertauschte Startseiten-Button-Verknuepfungen.
' Wenn ein Navigationsmakro durch eine andere Startkachel ausgeloest
' wurde, wird auf das richtige Ziel umgeleitet.
' Rueckgabe: True = Umleitung ausgefuehrt, Aufrufer soll Exit Sub.
' ===============================================================
Private Function LeiteBeiFehlverdrahtungWeiter(ByVal erwarteteKachel As String) As Boolean
    LeiteBeiFehlverdrahtungWeiter = False

    Dim callerName As String
    callerName = ""
    On Error Resume Next
    callerName = LCase$(Trim$(CStr(Application.Caller)))
    On Error GoTo 0

    If callerName = "" Then Exit Function
    If Left$(callerName, 7) <> "kachel_" Then Exit Function
    If callerName = erwarteteKachel Then Exit Function

    If RouteStartkachelDirekt(callerName) Then
        LeiteBeiFehlverdrahtungWeiter = True
    End If
End Function

' ===============================================================
' Direkte Zielzuordnung nach Startkachel-Name.
' ===============================================================
Private Function RouteStartkachelDirekt(ByVal callerName As String) As Boolean
    RouteStartkachelDirekt = True

    Select Case callerName
        Case "kachel_bankkonto"
            AktiviereTabellenblatt WS_BANKKONTO
        Case "kachel_strom"
            AktiviereTabellenblatt "Strom"
        Case "kachel_wasser"
            AktiviereTabellenblatt "Wasser"
        Case "kachel_einstellungen"
            AktiviereTabellenblatt WS_EINSTELLUNGEN
        Case "kachel_daten"
            AktiviereTabellenblatt WS_DATEN
        Case "kachel_vereinskasse"
            AktiviereTabellenblatt WS_VEREINSKASSE
        Case "kachel_uebersicht"
            AktiviereTabellenblatt WS_UEBERSICHT()
        Case "kachel_dashboard"
            Dim wsDash As Worksheet
            Set wsDash = Nothing
            Set wsDash = FindeTabellenblattRobust("Dashboard Mitgliederzahlungen", "Dashboard")
            If Not wsDash Is Nothing Then
                AktiviereZielblattStabil wsDash
            Else
                MsgBox "Das Dashboard wurde noch nicht erstellt.", vbInformation, "Dashboard"
            End If
        Case "kachel_finanzuebersicht"
            NavigiereZu_FinanzUebersicht
        Case "kachel_mitglieder"
            frm_Mitgliederverwaltung.Show
        Case Else
            RouteStartkachelDirekt = False
    End Select
End Function

Public Sub ZeigeMitgliederverwaltung()
    frm_Mitgliederverwaltung.Show
End Sub

Public Sub ZeigeSerienbrief_Betriebskosten()
    MsgBox "Die Serienbrief-Funktion f" & ChrW(252) & "r die " & _
           "Betriebskostenabrechnung wird in einem sp" & ChrW(228) & _
           "teren Schritt implementiert.", _
           vbInformation, "Betriebskostenabrechnung"
End Sub

Public Sub ZeigeSerienbrief_Endabrechnung()
    MsgBox "Die Serienbrief-Funktion f" & ChrW(252) & "r die " & _
           "Endabrechnung wird in einem sp" & ChrW(228) & _
           "teren Schritt implementiert.", _
           vbInformation, "Endabrechnung"
End Sub


' ===============================================================
' HILFSFUNKTION: Tabellenblatt aktivieren (intern)
' ===============================================================
Private Sub AktiviereTabellenblatt(ByVal blattName As String)
    Dim ws As Worksheet
    Set ws = FindeTabellenblattRobust(blattName)

    If Not ws Is Nothing Then
        AktiviereZielblattStabil ws
    Else
        MsgBox "Tabellenblatt """ & blattName & """ nicht gefunden.", _
               vbExclamation, "Navigation"
    End If
End Sub

' ===============================================================
' STABILE AKTIVIERUNG (Sicherheitsnetz gegen falsche Zielblaetter):
' Aktiviert das Zielblatt und laesst dessen Worksheet_Activate-Logik
' einmal laufen. Falls ein Nebeneffekt (z.B. Cross-Sheet-DropDowns)
' danach ein ANDERES Blatt aktiv laesst, wird das gewuenschte Ziel
' OHNE erneute Events zwingend wiederhergestellt. So landet ein
' Button-Klick garantiert auf dem richtigen Tabellenblatt.
' ===============================================================
Public Sub AktiviereZielblattStabil(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Activate
    DoEvents

    If Not IstAktivesBlatt(ws) Then
        Application.EnableEvents = False
        ws.Activate
        Application.EnableEvents = True
    End If

    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Private Function IstAktivesBlatt(ByVal ws As Worksheet) As Boolean
    IstAktivesBlatt = False
    On Error Resume Next
    IstAktivesBlatt = (Not ActiveSheet Is Nothing) And (ActiveSheet Is ws)
    On Error GoTo 0
End Function

Public Function FindeTabellenblattRobust(ByVal blattName As String, Optional ByVal fallbackName As String = "") As Worksheet
    Dim ws As Worksheet
    Dim aliases As Variant
    Dim i As Long

    Set FindeTabellenblattRobust = Nothing

    Set ws = FindeTabellenblattExakt(blattName)
    If Not ws Is Nothing Then
        Set FindeTabellenblattRobust = ws
        Exit Function
    End If

    If Len(fallbackName) > 0 Then
        Set ws = FindeTabellenblattExakt(fallbackName)
        If Not ws Is Nothing Then
            Set FindeTabellenblattRobust = ws
            Exit Function
        End If
    End If

    aliases = HoleBlattAliase(blattName, fallbackName)
    For i = LBound(aliases) To UBound(aliases)
        Set ws = FindeTabellenblattExakt(CStr(aliases(i)))
        If Not ws Is Nothing Then
            Set FindeTabellenblattRobust = ws
            Exit Function
        End If
    Next i

    For i = LBound(aliases) To UBound(aliases)
        Set ws = FindeTabellenblattEnthaelt(CStr(aliases(i)))
        If Not ws Is Nothing Then
            Set FindeTabellenblattRobust = ws
            Exit Function
        End If
    Next i
End Function

Private Function FindeTabellenblattExakt(ByVal blattName As String) As Worksheet
    Dim ws As Worksheet

    Set FindeTabellenblattExakt = Nothing

    If Len(blattName) = 0 Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(blattName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set FindeTabellenblattExakt = ws
        Exit Function
    End If

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, blattName, vbTextCompare) = 0 Then
            Set FindeTabellenblattExakt = ws
            Exit Function
        End If
    Next ws
End Function

Private Function FindeTabellenblattEnthaelt(ByVal teil As String) As Worksheet
    Dim ws As Worksheet

    Set FindeTabellenblattEnthaelt = Nothing
    If Len(teil) = 0 Then Exit Function

    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, teil, vbTextCompare) > 0 Then
            Set FindeTabellenblattEnthaelt = ws
            Exit Function
        End If
    Next ws
End Function

Private Function HoleBlattAliase(ByVal blattName As String, ByVal fallbackName As String) As Variant
    Dim key As String

    key = LCase$(Trim$(blattName & "|" & fallbackName))

    Select Case True
        Case InStr(1, key, "zahlungs", vbTextCompare) > 0 Or InStr(1, key, "uebersicht", vbTextCompare) > 0
            HoleBlattAliase = Array("Zahlungs" & Chr$(252) & "bersicht", "Zahlungsuebersicht", _
                                    Chr$(220) & "bersicht", "Uebersicht")
        Case InStr(1, key, "startmen", vbTextCompare) > 0 Or InStr(1, key, "startseite", vbTextCompare) > 0
            HoleBlattAliase = Array("Startmen" & Chr$(252), "Startmenue", "Startseite")
        Case InStr(1, key, "finanz", vbTextCompare) > 0
            HoleBlattAliase = Array("Finanz-" & Chr$(220) & "bersicht", "Finanz-Uebersicht", "Finanz" & Chr$(252) & "bersicht")
        Case InStr(1, key, "dashboard", vbTextCompare) > 0
            HoleBlattAliase = Array("Dashboard Mitgliederzahlungen", "Dashboard")
        Case InStr(1, key, "strom", vbTextCompare) > 0
            HoleBlattAliase = Array("Strom")
        Case InStr(1, key, "wasser", vbTextCompare) > 0
            HoleBlattAliase = Array("Wasser")
        Case Else
            HoleBlattAliase = Array(blattName, fallbackName)
    End Select
End Function


' ===============================================================
' HOME-BUTTONS: Auf allen Blättern erstellen (ausser Startseite)
' Wird bei Workbook_Open aufgerufen
' ===============================================================
Public Sub SetzeHomeButtonsAufAllenBlaettern()
    Dim ws As Worksheet
    Dim startName As String
    startName = WS_STARTMENUE()
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> startName Then
            ' Navigationszeilen-Höhe korrigieren (falls bereits migriert)
            On Error Resume Next
            If Application.WorksheetFunction.CountA(ws.Rows(1)) = 0 Then
                ws.Unprotect PASSWORD:=PASSWORD
                ws.Rows(1).RowHeight = 30
                ' Auf der übersicht liegt in Zeile 2 die Monats-Register-Leiste,
                ' deshalb braucht sie dort mehr Platz als der schmale Spacer (3) auf den anderen Blättern.
                If ws.Name = WS_UEBERSICHT() Then
                    ws.Rows(2).RowHeight = 26
                Else
                    ws.Rows(2).RowHeight = 3
                End If
                ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
            End If
            On Error GoTo 0
            Call ErstelleHomeButton(ws)
        End If
    Next ws
End Sub


' ===============================================================
' HOME-BUTTON: Einzelnen Button auf Blatt erstellen
' ===============================================================
Public Sub ErstelleHomeButton(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' Bestehenden Button entfernen falls vorhanden
    Call EntferneHomeButton(ws)
    
    On Error GoTo BtnFehler
    
    ' Größe abhängig vom Blatt-Typ anpassen
    Dim btnW As Double, btnH As Double, fontSize As Double
    Select Case ws.Name
        Case WS_BANKKONTO
            ' Breites Blatt - etwas groesserer Button
            btnW = 100: btnH = 30: fontSize = 11
        Case WS_UEBERSICHT()
            ' Schmales Blatt mit Filter-Buttons darunter - kompakter
            btnW = 76: btnH = 24: fontSize = 9.5
        Case WS_FINANZ_UEBERSICHT()
            btnW = 80: btnH = 26: fontSize = 10
        Case Else
            ' Standard für alle anderen Blätter
            btnW = 88: btnH = 26: fontSize = 10
    End Select
    
    ' Button oben links positionieren (Zelle A1/B1)
    Dim btnLeft As Double
    Dim btnTop As Double
    btnLeft = ws.Range("A1").Left + 4
    btnTop = ws.Range("A1").Top + 3
    
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                  btnLeft, btnTop, _
                                  btnW, btnH)
    
    With shp
        .Name = HOME_BTN_NAME
        .Fill.ForeColor.RGB = RGB(44, 62, 80)
        .Line.Visible = msoFalse
        
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 4
            .MarginRight = 4
            .MarginTop = 2
            .MarginBottom = 2
            
            With .TextRange
                .text = ChrW(8962) & " Home"
                .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Font.Size = fontSize
                .Font.Bold = msoTrue
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
        
        .OnAction = "'" & ThisWorkbook.Name & "'!mod_Navigation.NavigiereZuStartseite"
        .Placement = xlFreeFloating
    End With
    
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0
    Exit Sub

BtnFehler:
    Debug.Print "[Navigation] Home-Button auf """ & ws.Name & """ fehlgeschlagen: " & Err.Description
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0
End Sub


' ===============================================================
' HOME-BUTTON: Bestehenden Button entfernen
' ===============================================================
Private Sub EntferneHomeButton(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Shapes(HOME_BTN_NAME).Delete
    Err.Clear
    On Error GoTo 0
End Sub


' ===============================================================
' MIGRATION: 2 Zeilen oben einfügen für Navigationsleiste
' Wird einmalig aufgerufen wenn Blatt noch keine Navigationszeilen hat.
' Prüft ob Zeile 1 leer ist (= bereits migriert) oder Daten enthält.
' Betroffene Blätter: Bankkonto, Strom, Wasser, Einstellungen,
'                      Zahlungsübersicht, Finanz-übersicht
' ===============================================================
Public Sub MigriereNavigationszeilen()
    Dim blaetter As Variant
    blaetter = Array(WS_BANKKONTO, "Strom", "Wasser", WS_EINSTELLUNGEN)
    
    Dim i As Long
    Dim ws As Worksheet
    
    For i = LBound(blaetter) To UBound(blaetter)
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(CStr(blaetter(i)))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            Call FuegeNavigationsZeilenEin(ws)
        End If
    Next i
    
    ' übersicht und Finanz-übersicht separat (Function-Konstanten)
    On Error Resume Next
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets(WS_UEBERSICHT())
    On Error GoTo 0
    If Not ws Is Nothing Then Call FuegeNavigationsZeilenEin(ws)
    
    On Error Resume Next
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets(WS_FINANZ_UEBERSICHT())
    On Error GoTo 0
    If Not ws Is Nothing Then Call FuegeNavigationsZeilenEin(ws)
End Sub

Private Sub FuegeNavigationsZeilenEin(ByVal ws As Worksheet)
    ' Prüfen ob bereits migriert: Zeile 1 muss leer sein UND
    ' Zeile 3 muss Daten enthalten (sonst ist das Blatt neu/leer)
    If Application.WorksheetFunction.CountA(ws.Rows(1)) = 0 And _
       Application.WorksheetFunction.CountA(ws.Rows(2)) = 0 Then
        ' Bereits migriert oder leer -> nichts tun
        Exit Sub
    End If
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' 2 leere Zeilen oben einfügen
    ws.Rows("1:2").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ' Eingefuegte Zeilen bereinigen
    ws.Range("A1:AZ2").Clear
    ws.Rows(1).RowHeight = 30
    ' Auf der übersicht liegt in Zeile 2 die Monats-Register-Leiste -> mehr Höhe
    If ws.Name = WS_UEBERSICHT() Then
        ws.Rows(2).RowHeight = 26
    Else
        ws.Rows(2).RowHeight = 3
    End If
    
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0
    
    Debug.Print "[Navigation] Navigationszeilen eingefügt auf: " & ws.Name
End Sub








































































