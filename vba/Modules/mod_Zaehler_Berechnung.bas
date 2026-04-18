Attribute VB_Name = "mod_Zaehler_Berechnung"
Option Explicit

' ===============================================================
' MODUL: mod_Zaehler_Berechnung
' Ausgelagert aus mod_ZaehlerLogik
' Enth?lt: CalculateAllZaehlerVerbrauch, CalculateSingleZaehler
' ===============================================================

' --- Konstanten (lokal dupliziert) ---
Private Const HIST_SHEET_NAME As String = "Z?hlerhistorie"
Private Const PASSWORD As String = ""
Private Const STR_HISTORY_SEPARATOR As String = "--- Z?hlerhistorie Makro-Eintrag ---"

Private Const COL_STAND_ANFANG As String = "B"
Private Const COL_STAND_ENDE As String = "C"
Private Const COL_VERBRAUCH_GESAMT As String = "D"
Private Const COL_BEMERKUNG As String = "E"

Private Const COL_HIST_PARZELLE As Long = 3
Private Const COL_HIST_MEDIUM As Long = 4
Private Const COL_HIST_DATUM As Long = 2
Private Const COL_HIST_ZAEHLER_NEU As Long = 8
Private Const COL_HIST_STAND_NEU_START As Long = 9
Private Const COL_HIST_VERBRAUCH_ALT As Long = 10

Private Const RGB_EINGABE_ERLAUBT As Long = 7592334
Private Const RGB_GEWECHSELT As Long = 4980735


' ==========================================================
' HAUPTPROZEDUR: Berechnung aller Z?hler einer Seite
' ==========================================================
Public Sub CalculateAllZaehlerVerbrauch(wsTarget As Worksheet)
    
    Dim wsHist As Worksheet
    Dim r As Long
    Dim wasProtected As Boolean
    
    If wsTarget Is Nothing Then Exit Sub
    
    ' --- Blattschutz aufheben ---
    wasProtected = wsTarget.ProtectContents
    If wasProtected Then
        On Error Resume Next
        wsTarget.Unprotect PASSWORD
        On Error GoTo 0
    End If
    
    On Error GoTo Fehler_Handler_Berechnung
    
    ' --- Formatierung ---
    wsTarget.Range("8:23").RowHeight = 50

    With wsTarget.Range("B8:D23, F8:I23")
        .ShrinkToFit = True
        .WrapText = False
    End With
    
    With wsTarget.Range("A8:A23")
        .ShrinkToFit = False
        .WrapText = True
    End With

    With wsTarget.Range(COL_BEMERKUNG & "8:" & COL_BEMERKUNG & "23")
        .ShrinkToFit = False
        .WrapText = True
    End With
    
    ' --- Historie laden ---
    On Error Resume Next
    Set wsHist = ThisWorkbook.Worksheets(HIST_SHEET_NAME)
    On Error GoTo Fehler_Handler_Berechnung

    ' ==========================================================
    ' 1. PARZELLENZ?HLER & UNTERZ?HLER
    ' ==========================================================
    
    If LCase(wsTarget.Name) = "strom" Then
        ' STROM: Parzelle 1 bis 12 (Zeilen 8 bis 19)
        For r = 1 To 12
            Call CalculateSingleZaehler(wsTarget, wsHist, "Strom", "Parzelle " & r, r + 7)
            wsTarget.Rows(r + 7).AutoFit
            Call mod_ZaehlerLogik.EnsureMinRowHeight(wsTarget, r + 7)
        Next r
        
        ' STROM: Feste Z?hler
        Call CalculateSingleZaehler(wsTarget, wsHist, "Strom", "Clubwagen", 22)
        wsTarget.Rows(22).AutoFit
        Call mod_ZaehlerLogik.EnsureMinRowHeight(wsTarget, 22)
        
        Call CalculateSingleZaehler(wsTarget, wsHist, "Strom", "K?hltruhe", 23)
        wsTarget.Rows(23).AutoFit
        Call mod_ZaehlerLogik.EnsureMinRowHeight(wsTarget, 23)
        
        Call CalculateSingleZaehler(wsTarget, wsHist, "Strom", "Parzelle 13", 20)
        wsTarget.Rows(20).AutoFit
        Call mod_ZaehlerLogik.EnsureMinRowHeight(wsTarget, 20)
        
        Call CalculateSingleZaehler(wsTarget, wsHist, "Strom", "Parzelle 14", 21)
        wsTarget.Rows(21).AutoFit
        Call mod_ZaehlerLogik.EnsureMinRowHeight(wsTarget, 21)
        
    ElseIf LCase(wsTarget.Name) = "wasser" Then
        ' WASSER: Parzelle 1 bis 14 (Zeilen 10 bis 23)
        For r = 1 To 14
            Call CalculateSingleZaehler(wsTarget, wsHist, "Wasser", "Parzelle " & r, r + 9)
            wsTarget.Rows(r + 9).AutoFit
            Call mod_ZaehlerLogik.EnsureMinRowHeight(wsTarget, r + 9)
        Next r
    End If
    
    ' ==========================================================
    ' 2. HAUPTZ?HLER
    ' ==========================================================
    
    If LCase(wsTarget.Name) = "strom" Then
        Call CalculateSingleZaehler(wsTarget, wsHist, "Strom", "Hauptz?hler", 26)
        wsTarget.Rows(26).AutoFit
        Call mod_ZaehlerLogik.EnsureMinRowHeight(wsTarget, 26)
    End If
    
    If LCase(wsTarget.Name) = "wasser" Then
        Call CalculateSingleZaehler(wsTarget, wsHist, "Wasser", "Hauptz?hler", 29)
        wsTarget.Rows(29).AutoFit
        Call mod_ZaehlerLogik.EnsureMinRowHeight(wsTarget, 29)
    End If

Cleanup_Berechnung:
    If wasProtected Then
        wsTarget.Protect PASSWORD, AllowFormattingCells:=True
    End If
    Exit Sub

Fehler_Handler_Berechnung:
    MsgBox "Ein schwerwiegender Fehler ist w?hrend der Z?hlerberechnung aufgetreten. " & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description, vbCritical, "Fehler in CalculateAllZaehlerVerbrauch"
    Resume Cleanup_Berechnung

End Sub

' ==========================================================
' EINZELBERECHNUNG (Kernlogik)
' ==========================================================
Private Sub CalculateSingleZaehler( _
    wsTarget As Worksheet, _
    wsHist As Worksheet, _
    ByVal ZaehlerTyp As String, _
    ByVal ZaehlerName As String, _
    ByVal targetRow As Long)

    Dim startCell As Range, endCell As Range
    Dim targetCellD As Range
    Dim targetBemerkung As Range
    
    Dim standAnfangCurrent As Double
    Dim standEndeCurrent As Double
    Dim VerbrauchGesamt As Double
    Dim verbrauchAltHistorie_Summe As Double
    Dim verbrauchNeuAktuell As Double
    Dim einheit As String
    Dim f As Range, firstAddr As String
    Dim currentRow As Long
    Dim zyklen As Long
    Dim lastDate As Date
    Dim snNeu_last As String
    Dim standNeuStart_last As Double
    
    Set startCell = wsTarget.Cells(targetRow, COL_STAND_ANFANG)
    Set endCell = wsTarget.Cells(targetRow, COL_STAND_ENDE)
    Set targetCellD = wsTarget.Cells(targetRow, COL_VERBRAUCH_GESAMT)
    Set targetBemerkung = wsTarget.Cells(targetRow, COL_BEMERKUNG)

    einheit = IIf(LCase(ZaehlerTyp) = "strom", "kWh", "m?")
    
    ' 0. Startwerte lesen
    If IsNumeric(startCell.value) And Not isEmpty(startCell.value) Then
        standAnfangCurrent = CDbl(startCell.value)
    Else
        standAnfangCurrent = 0
    End If
    
    If IsNumeric(endCell.value) And Not isEmpty(endCell.value) Then
        standEndeCurrent = CDbl(endCell.value)
    Else
        standEndeCurrent = 0
    End If
    
    ' 1. Vorabpr?fung (Fehler)
    If standEndeCurrent < standAnfangCurrent Then
        targetBemerkung.value = "FEHLER: Endstand (" & Format(standEndeCurrent, "#,##0.00") & ") < Startstand (" & Format(standAnfangCurrent, "#,##0.00") & ")."
        targetCellD.ClearContents
        
        startCell.Interior.color = RGB_EINGABE_ERLAUBT
        startCell.Locked = False
        endCell.Interior.color = RGB_EINGABE_ERLAUBT
        endCell.Locked = False
        
        With targetBemerkung
            .ShrinkToFit = False
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        
        Exit Sub
    End If
    
    
    ' 2. HISTORIE DURCHSUCHEN: SUMMIERE ALLE WECHSEL
    verbrauchAltHistorie_Summe = 0
    zyklen = 0
    lastDate = 0
    standNeuStart_last = 0
    snNeu_last = ""
    
    If Not wsHist Is Nothing Then
        Set f = wsHist.Columns(COL_HIST_PARZELLE).Find( _
            What:=ZaehlerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

        If Not f Is Nothing Then
            firstAddr = f.Address
            Do
                currentRow = f.Row
                If StrComp(Trim(CStr(wsHist.Cells(currentRow, COL_HIST_MEDIUM).value)), ZaehlerTyp, vbTextCompare) = 0 Then
                    
                    zyklen = zyklen + 1
                    If IsNumeric(wsHist.Cells(currentRow, COL_HIST_VERBRAUCH_ALT).value) Then
                        verbrauchAltHistorie_Summe = verbrauchAltHistorie_Summe + CDbl(wsHist.Cells(currentRow, COL_HIST_VERBRAUCH_ALT).value)
                    End If
                    
                    If IsDate(wsHist.Cells(currentRow, COL_HIST_DATUM).value) Then
                        If CDate(wsHist.Cells(currentRow, COL_HIST_DATUM).value) >= lastDate Then
                            lastDate = CDate(wsHist.Cells(currentRow, COL_HIST_DATUM).value)
                            snNeu_last = CStr(wsHist.Cells(currentRow, COL_HIST_ZAEHLER_NEU).value)
                            If IsNumeric(wsHist.Cells(currentRow, COL_HIST_STAND_NEU_START).value) Then
                                standNeuStart_last = CDbl(wsHist.Cells(currentRow, COL_HIST_STAND_NEU_START).value)
                            End If
                        End If
                    End If
                End If
                Set f = wsHist.Columns(COL_HIST_PARZELLE).FindNext(f)
            Loop While Not f Is Nothing And f.Address <> firstAddr
        End If
    End If


    ' 3. BERECHNUNG UND SCHREIBEN IN D, E
    If zyklen > 0 Then ' FALL A: Mindestens ein Z?hlerwechsel
        
        If standAnfangCurrent <> standNeuStart_last Then
            startCell.value = mod_ZaehlerLogik.CleanNumber(standNeuStart_last)
            standAnfangCurrent = standNeuStart_last
        End If
        
        verbrauchNeuAktuell = Round(CDec(standEndeCurrent) - CDec(standAnfangCurrent), 2)
        VerbrauchGesamt = CDec(verbrauchAltHistorie_Summe) + CDec(verbrauchNeuAktuell)
        
        ' Spalte D: Gesamtverbrauch
        If targetRow = 22 And ZaehlerName = "Clubwagen" Then
              targetCellD.value = Round(VerbrauchGesamt, 0)
              targetCellD.NumberFormat = "0;[Red]-0;;"
        Else
              targetCellD.value = VerbrauchGesamt
              targetCellD.NumberFormat = "#,##0.00;[Red]-#,##0.00;;"
        End If
        
        ' ***************************************************************
        ' LOGIK F?R SPALTE E (BEMERKUNG BEI Z?HLERWECHSEL)
        ' ***************************************************************
        Dim oldBemerkung As String
        Dim newHistoryText As String
        Dim posSeparator As Long
        
        newHistoryText = "Letzter Z?hlerwechsel am: " & Format(lastDate, "dd.mm.yyyy") & vbLf & _
                             "Anzahl der Wechsel: " & zyklen & vbLf & _
                             "Gesamtverbrauch gewechselte Z?hler: " & Format(verbrauchAltHistorie_Summe, "#,##0.00") & " " & einheit & vbLf & _
                             "Verbrauch derzeitiger Z?hler: " & Format(verbrauchNeuAktuell, "#,##0.00") & " " & einheit
        
        oldBemerkung = Trim(CStr(targetBemerkung.value))
        
        posSeparator = InStr(1, oldBemerkung, STR_HISTORY_SEPARATOR, vbTextCompare)
        
        If posSeparator > 0 Then
            Dim userText As String
            
            userText = Trim(Left(oldBemerkung, posSeparator - 1))
            
            If Len(userText) > 0 Then
                targetBemerkung.value = userText & vbLf & STR_HISTORY_SEPARATOR & vbLf & newHistoryText
            Else
                targetBemerkung.value = STR_HISTORY_SEPARATOR & vbLf & newHistoryText
            End If
            
        Else
            If Len(oldBemerkung) > 0 Then
                targetBemerkung.value = oldBemerkung & vbLf & STR_HISTORY_SEPARATOR & vbLf & newHistoryText
            Else
                targetBemerkung.value = STR_HISTORY_SEPARATOR & vbLf & newHistoryText
            End If
        End If
        
        With targetBemerkung
            .ShrinkToFit = False
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        
        startCell.Interior.color = RGB_GEWECHSELT
        startCell.Locked = True
        endCell.Interior.color = RGB_EINGABE_ERLAUBT
        endCell.Locked = False
        
    Else ' FALL B: Kein Wechsel (Standardfall)
        
        verbrauchNeuAktuell = Round(CDec(standEndeCurrent) - CDec(standAnfangCurrent), 2)
        VerbrauchGesamt = verbrauchNeuAktuell
        
        targetCellD.value = VerbrauchGesamt
        targetCellD.NumberFormat = "#,##0.00;[Red]-#,##0.00;;"
        
        With targetBemerkung
            .ShrinkToFit = False
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        
        With startCell
            .Locked = False
            .Interior.color = RGB_EINGABE_ERLAUBT
        End With
        
        With endCell
            .Locked = False
            .Interior.color = RGB_EINGABE_ERLAUBT
        End With
        
    End If
    
End Sub






































































