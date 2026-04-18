Attribute VB_Name = "mod_Navigation"
Option Explicit

' ===============================================================
' MODUL: mod_Navigation
' VERSION: 1.0 - 18.04.2026
' ZWECK: Navigation zwischen Tabellenblaettern
'        - Startseite -> alle Blaetter (Button-Handler)
'        - Alle Blaetter -> Startseite (Home-Button)
'        - Home-Buttons auf allen Blaettern erstellen/entfernen
' ===============================================================

Private Const HOME_BTN_NAME As String = "btn_Home"
Private Const HOME_BTN_WIDTH As Double = 90
Private Const HOME_BTN_HEIGHT As Double = 28
Private Const HOME_BTN_LEFT As Double = 6
Private Const HOME_BTN_TOP As Double = 6


' ===============================================================
' NAVIGATION: Startseite aktivieren
' ===============================================================
Public Sub NavigiereZuStartseite()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_STARTMENUE())
    If Not ws Is Nothing Then
        ws.Activate
        ws.Range("A1").Select
    End If
    On Error GoTo 0
End Sub


' ===============================================================
' NAVIGATION: Einzelne Blaetter aktivieren (Button-Handler)
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
        ws.Activate
        ws.Range("A1").Select
    End If
End Sub

Public Sub NavigiereZu_Uebersicht()
    AktiviereTabellenblatt WS_UEBERSICHT()
End Sub

Public Sub NavigiereZu_Dashboard()
    ' Dashboard wird dynamisch erzeugt - Name kann variieren
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dashboard Mitgliederzahlungen")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Das Dashboard wurde noch nicht erstellt." & vbLf & vbLf & _
               "Bitte zuerst die Zahlungs" & ChrW(252) & "bersicht " & _
               "oder das Dashboard generieren.", _
               vbInformation, "Dashboard nicht vorhanden"
        Exit Sub
    End If
    
    ws.Activate
    ws.Range("A1").Select
End Sub

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
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(blattName)
    If Not ws Is Nothing Then
        ws.Activate
        ws.Range("A1").Select
    Else
        MsgBox "Tabellenblatt """ & blattName & """ nicht gefunden.", _
               vbExclamation, "Navigation"
    End If
    On Error GoTo 0
End Sub


' ===============================================================
' HOME-BUTTONS: Auf allen Blaettern erstellen (ausser Startseite)
' Wird bei Workbook_Open aufgerufen
' ===============================================================
Public Sub SetzeHomeButtonsAufAllenBlaettern()
    Dim ws As Worksheet
    Dim startName As String
    startName = WS_STARTMENUE()
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> startName Then
            Call ErstelleHomeButton(ws)
        End If
    Next ws
End Sub


' ===============================================================
' HOME-BUTTON: Einzelnen Button auf Blatt erstellen
' ===============================================================
Private Sub ErstelleHomeButton(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' Bestehenden Button entfernen falls vorhanden
    Call EntferneHomeButton(ws)
    
    On Error GoTo BtnFehler
    
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                  HOME_BTN_LEFT, HOME_BTN_TOP, _
                                  HOME_BTN_WIDTH, HOME_BTN_HEIGHT)
    
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
                .Font.Size = 10
                .Font.Bold = msoTrue
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
        
        .OnAction = "'mod_Navigation.NavigiereZuStartseite'"
        .Placement = xlFreeFloating
    End With
    
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    Exit Sub

BtnFehler:
    Debug.Print "[Navigation] Home-Button auf """ & ws.Name & """ fehlgeschlagen: " & Err.Description
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
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












