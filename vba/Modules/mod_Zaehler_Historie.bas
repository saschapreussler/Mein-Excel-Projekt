Attribute VB_Name = "mod_Zaehler_Historie"
Option Explicit

' ===============================================================
' MODUL: mod_Zaehler_Historie
' Ausgelagert aus mod_ZaehlerLogik
' Enth?lt: SchreibeHistorie, FarbeHistorieEintraege
' ===============================================================

' --- Konstanten (lokal dupliziert) ---
Private Const HIST_SHEET_NAME As String = "Z?hlerhistorie"
Private Const HIST_TABLE_NAME As String = "Tabelle_Zaehlerhistorie"
Private Const PASSWORD As String = ""

Private Const COL_STAND_ANFANG As String = "B"

Private Const COL_HIST_ID As Long = 1
Private Const COL_HIST_DATUM As Long = 2
Private Const COL_HIST_PARZELLE As Long = 3
Private Const COL_HIST_MEDIUM As Long = 4
Private Const COL_HIST_ZAEHLER_ALT As Long = 5
Private Const COL_HIST_STAND_ALT_ANFANG As Long = 6
Private Const COL_HIST_STAND_ALT_ENDE As Long = 7
Private Const COL_HIST_ZAEHLER_NEU As Long = 8
Private Const COL_HIST_STAND_NEU_START As Long = 9
Private Const COL_HIST_VERBRAUCH_ALT As Long = 10
Private Const COL_HIST_BEMERKUNG As Long = 11

Private Const RGB_STROM As Long = 86271
Private Const RGB_WASSER As Long = 16737792
Private Const RGB_EINGABE_ERLAUBT As Long = 7592334
Private Const RGB_GEWECHSELT As Long = 4980735


' ==========================================================
' HISTORIE SCHREIBEN
' ==========================================================
Public Sub SchreibeHistorie( _
    ByVal parzelle As String, _
    ByVal DatumW As Date, _
    ByVal AltEnde As Double, _
    ByVal neuStart As Double, _
    ByVal snNeu As String, _
    ByVal snAlt As String, _
    Optional ByVal bem As String = "", _
    Optional ByVal Medium As String)
    
    Dim ws As Worksheet, lo As ListObject, newRow As ListRow
    Dim lngColor As Long
    Dim wsTarget As Worksheet
    Dim targetRow As Long
    Dim standAnfangAlt As Double, verbrauchAltHistorie As Double
    Dim wasTargetProtected As Boolean, wasHistoryProtected As Boolean
    Dim AltEnde_Geprueft As Double, neuStart_Geprueft As Double
    
    ' I. WERTE PR?FEN & RUNDEN
    If AltEnde = Int(AltEnde) Then AltEnde_Geprueft = AltEnde Else AltEnde_Geprueft = Round(AltEnde, 4)
    If neuStart = Int(neuStart) Then neuStart_Geprueft = neuStart Else neuStart_Geprueft = Round(neuStart, 4)
    
    ' Zielblatt ermitteln
    Select Case Medium
        Case "Strom"
            Set wsTarget = ThisWorkbook.Worksheets("Strom")
        Case "Wasser"
            Set wsTarget = ThisWorkbook.Worksheets("Wasser")
        Case Else
            Exit Sub
    End Select
    
    targetRow = mod_ZaehlerLogik.GetTargetRow(parzelle, Medium)
    
    On Error GoTo Fehler_Handler
    
    If mod_ZaehlerLogik.IsZaehlerLogicRunning Then Err.Raise 9998, "SchreibeHistorie", "Logik ist bereits aktiv (Rekursion). Vorgang abgebrochen."
    mod_ZaehlerLogik.IsZaehlerLogicRunning = True
    Application.EnableEvents = False
    
    ' BLATTSCHUTZ AUFHEBEN (Zielblatt)
    If Not wsTarget Is Nothing Then
        wasTargetProtected = wsTarget.ProtectContents
        If wasTargetProtected Then
            On Error Resume Next
            wsTarget.Unprotect PASSWORD
            On Error GoTo Fehler_Handler
        End If
    End If

    ' 1. SICHERSTELLEN, DASS DAS BLATT/LISTOBJECT EXISTIERT
    Call mod_ZaehlerLogik.PruefeUndErstelleZaehlerhistorie
    
    Set ws = ThisWorkbook.Worksheets(HIST_SHEET_NAME)
    Set lo = ws.ListObjects(HIST_TABLE_NAME)
    
    If lo Is Nothing Then Err.Raise 9999, "mod_Zaehler_Historie.SchreibeHistorie", "ListObject wurde nicht gefunden/erstellt."
    
    ' Historienblatt entsperren
    wasHistoryProtected = ws.ProtectContents
    If wasHistoryProtected Then
        On Error Resume Next
        ws.Unprotect PASSWORD
        On Error GoTo Fehler_Handler
    End If
    
    ' 2. Lese den Startstand des ALTEN Z?hlers
    If targetRow > 0 And Not wsTarget Is Nothing Then
        If IsNumeric(wsTarget.Cells(targetRow, COL_STAND_ANFANG).value) Then
            standAnfangAlt = CDbl(wsTarget.Cells(targetRow, COL_STAND_ANFANG).value)
            If standAnfangAlt <> Int(standAnfangAlt) Then standAnfangAlt = Round(standAnfangAlt, 4)
        Else
            standAnfangAlt = 0
        End If
    Else
        standAnfangAlt = 0
    End If
    
    verbrauchAltHistorie = Round(CDec(AltEnde_Geprueft) - CDec(standAnfangAlt), 4)
    
    ' 3. Daten in Historie speichern
    Set newRow = lo.ListRows.Add(AlwaysInsert:=True)
    
    With newRow.Range
        .Cells(1, COL_HIST_ID).value = lo.ListRows.count
        .Cells(1, COL_HIST_DATUM).value = DatumW
        .Cells(1, COL_HIST_PARZELLE).value = parzelle
        .Cells(1, COL_HIST_MEDIUM).value = Medium
        
        .Cells(1, COL_HIST_ZAEHLER_ALT).value = snAlt
        .Cells(1, COL_HIST_STAND_ALT_ANFANG).value = mod_ZaehlerLogik.CleanNumber(standAnfangAlt)
        .Cells(1, COL_HIST_STAND_ALT_ENDE).value = mod_ZaehlerLogik.CleanNumber(AltEnde_Geprueft)
        .Cells(1, COL_HIST_ZAEHLER_NEU).value = snNeu
        .Cells(1, COL_HIST_STAND_NEU_START).value = mod_ZaehlerLogik.CleanNumber(neuStart_Geprueft)
        .Cells(1, COL_HIST_VERBRAUCH_ALT).value = mod_ZaehlerLogik.CleanNumber(verbrauchAltHistorie)
        .Cells(1, COL_HIST_BEMERKUNG).value = bem
    End With
    
    ' 4. ZIELBLATT-UPDATE (Spalten B, C)
    If targetRow > 0 And Not wsTarget Is Nothing Then
        
        wsTarget.Cells(targetRow, COL_STAND_ANFANG).value = mod_ZaehlerLogik.CleanNumber(neuStart_Geprueft)
        wsTarget.Cells(targetRow, "C").value = mod_ZaehlerLogik.CleanNumber(neuStart_Geprueft)
        
        With wsTarget.Cells(targetRow, COL_STAND_ANFANG)
            .Interior.color = RGB_GEWECHSELT
            .Locked = True
        End With
        
        With wsTarget.Cells(targetRow, "C")
            .Interior.color = RGB_EINGABE_ERLAUBT
            .Locked = False
        End With
        
    End If
    
    ' 5. Farben f?r Historie setzen & Update-Call
    Select Case Medium
        Case "Strom"
            lngColor = RGB_STROM
        Case "Wasser"
            lngColor = RGB_WASSER
        Case Else
            lngColor = xlNone
    End Select
    
    If lngColor <> xlNone Then newRow.Range.Interior.color = lngColor
    
    If Not wsTarget Is Nothing Then Call mod_Zaehler_Berechnung.CalculateAllZaehlerVerbrauch(wsTarget)
    Call FarbeHistorieEintraege
    
CleanUp:
    mod_ZaehlerLogik.IsZaehlerLogicRunning = False
    Application.EnableEvents = True
    
    If Not wsTarget Is Nothing Then
        If wasTargetProtected Then wsTarget.Protect PASSWORD, AllowFormattingCells:=True
    End If
    If Not ws Is Nothing Then
        If wasHistoryProtected Then ws.Protect PASSWORD, AllowFormattingCells:=True
    End If
    
    Exit Sub

Fehler_Handler:
    Dim errNum As Long
    Dim errDesc As String
    If Err.Number <> 0 Then
        errNum = Err.Number
        errDesc = Err.Description
    Else
        errNum = 9997
        errDesc = "Unbekannter Fehler im Fehler-Handler"
    End If

    Resume CleanUp

    Err.Raise errNum, "mod_Zaehler_Historie.SchreibeHistorie", errDesc
End Sub

' ==========================================================
' F?RBT HISTORIEN-EINTR?GE
' ==========================================================
Public Sub FarbeHistorieEintraege()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lr As ListRow
    Dim lngColor As Long
    Dim wasProtected As Boolean
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HIST_SHEET_NAME)
    If ws Is Nothing Then Exit Sub
    Set lo = ws.ListObjects(HIST_TABLE_NAME)
    If lo Is Nothing Then Exit Sub
    
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect PASSWORD
    On Error GoTo 0
    
    For Each lr In lo.ListRows
        If StrComp(Trim(CStr(lr.Range.Cells(1, COL_HIST_MEDIUM).value)), "Strom", vbTextCompare) = 0 Then
            lngColor = RGB_STROM
        ElseIf StrComp(Trim(CStr(lr.Range.Cells(1, COL_HIST_MEDIUM).value)), "Wasser", vbTextCompare) = 0 Then
            lngColor = RGB_WASSER
        Else
            lngColor = xlNone
        End If
        
        lr.Range.Interior.color = lngColor
    Next lr
    
    lo.Range.Borders.LineStyle = xlContinuous
    lo.Range.Borders.color = RGB(0, 0, 0)
    
    If wasProtected Then ws.Protect PASSWORD, AllowFormattingCells:=True
End Sub












































