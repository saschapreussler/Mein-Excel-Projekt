VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ==========================================================
' TABELLENBLATT-MODUL: Strom (Tabelle5) - MIT NETTO/BRUTTO-TOGGLE
' ==========================================================
Option Explicit

' ==========================================================
' EVENTS
' ==========================================================

' Wird bei jedem Wechsel auf das Blatt ausgeführt
Private Sub Worksheet_Activate()

    Dim ToggleControl As Object
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    ' 1. Modus setzen (Versuch, den aktuellen Modus des Toggle-Buttons abzurufen)
    On Error Resume Next
    Set ToggleControl = Me.OLEObjects("ToggleNettoBrutto").Object
    On Error GoTo ErrHandler
    
    If Not ToggleControl Is Nothing Then
        
        ' DEBUG-PRÜFUNG: Ist das Blatt geschützt?
        If Me.ProtectContents Then
            Debug.Print "Worksheet_Activate: Blatt ist geschützt. Hebe Schutz für SetModus auf."
            Me.Unprotect PASSWORD:=""
        Else
            Debug.Print "Worksheet_Activate: Blatt ist UNGESCHÜTZT."
        End If
        
        ' Rufen Sie SetModus auf
        If ToggleControl.Value = True Then
            Debug.Print "Worksheet_Activate: Aufruf SetModus('NETTO')"
            Call SetModus("NETTO")
        Else
            Debug.Print "Worksheet_Activate: Aufruf SetModus('BRUTTO')"
            Call SetModus("BRUTTO")
        End If
        
        ' Blattschutz wiederherstellen
        Me.Protect PASSWORD:="", AllowFormattingCells:=True
        Debug.Print "Worksheet_Activate: Blattschutz wiederhergestellt."
        
    Else
        ' Fallback: Wenn das Steuerelement nicht gefunden wurde
        Debug.Print "Worksheet_Activate: Toggle-Control nicht gefunden. Führe nur UpdateStromblatt aus."
        Call UpdateStromblatt
    End If
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    ' WICHTIG: Fehlerbehandlung, um den Blattschutz wiederherzustellen, bevor der Fehler angezeigt wird
    On Error Resume Next
    ' Nur versuchen, Blattschutz wiederherzustellen, wenn wir ihn aufgehoben haben könnten
    If Not Me.ProtectContents Then
        Me.Protect PASSWORD:="", AllowFormattingCells:=True
        Debug.Print "Worksheet_Activate (FEHLER): Blattschutz wiederhergestellt."
    End If
    MsgBox "Fehler beim Aktivieren des Strom-Blattes: " & Err.Description, vbCritical ' Hier kommt die Fehlermeldung aus dem Screenshot her
    On Error GoTo 0
    
    
End Sub

' Wird bei jeder Änderung einer Zelle auf dem Blatt ausgeführt
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim isZaehlerChange As Boolean: isZaehlerChange = False
    Dim isCalculationChange As Boolean: isCalculationChange = False
    Dim wasProtected As Boolean
    Dim ToggleControl As Object
    
    On Error GoTo ErrHandler
    
    ' 1. Prüfen, ob eine Zählerstand-relevante Zelle geändert wurde (Original-Logik B8:C26)
    If Not Intersect(Target, Me.Range("B8:C26")) Is Nothing Then
        isZaehlerChange = True
    End If
    
    ' 2. Prüfen, ob eine Netto/Brutto-relevante Zelle geändert wurde (Neue Logik B5, K2:K3, M2:M3)
    If Not Intersect(Target, Me.Range("B5,K2:K3,M2:M3")) Is Nothing Then
        isCalculationChange = True
    End If
    
    If Not isZaehlerChange And Not isCalculationChange Then Exit Sub
    
    ' Verhindert erneute Auslösung durch Schreiben von Werten
    Application.EnableEvents = False
    
    ' Hole Toggle-Control sicher
    On Error Resume Next
    Set ToggleControl = Me.OLEObjects("ToggleNettoBrutto").Object
    On Error GoTo ErrHandler
    
    ' --- Zellerkennung und Berechnung ---
    
    If isCalculationChange Then
        
        ' Blattschutz-Handling nur für die Berechnungsteile (K2:M3, B5) notwendig
        wasProtected = Me.ProtectContents
        If wasProtected Then
            Me.Unprotect PASSWORD:=""
            Debug.Print "Worksheet_Change: Blattschutz für Calculation-Change aufgehoben."
        End If
        
        If Not ToggleControl Is Nothing Then
            
            ' A. Netto/Brutto-Werte neu berechnen (Neue Logik)
            If ToggleControl.Value = True Then ' NETTO Modus
                If Not Intersect(Target, Me.Range("B5,K2:K3")) Is Nothing Then Call RechneAlleZeilen("NETTO")
            Else ' BRUTTO Modus
                If Not Intersect(Target, Me.Range("B5,M2:M3")) Is Nothing Then Call RechneAlleZeilen("BRUTTO")
            End If
            
        End If
        
        ' Blattschutz wiederherstellen
        If wasProtected Then
            Me.Protect PASSWORD:="", AllowFormattingCells:=True
            Debug.Print "Worksheet_Change: Blattschutz für Calculation-Change wiederhergestellt."
        End If
        
    End If
    
    ' B. Zählerstände aktualisieren (Original-Logik)
    If isZaehlerChange Then
        Call UpdateStromblatt
    End If
    
Cleanup:
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "Worksheet_Change Fehler #" & Err.Number & ": " & Err.Description, vbCritical
    ' Blattschutz wiederherstellen, falls er manuell im Change-Event aufgehoben wurde
    If wasProtected Then
        On Error Resume Next
        Me.Protect PASSWORD:="", AllowFormattingCells:=True
        On Error GoTo 0
    End If
    Resume Cleanup
End Sub

' ==========================================================
' TOGGLE BUTTON FUNKTIONALITÄT (NETTO/BRUTTO)
' ==========================================================
Private Sub ToggleNettoBrutto_Click()
    Dim sht As Worksheet: Set sht = Me
    
    On Error GoTo ErrHandler
    
    Dim wasProtected As Boolean
    wasProtected = sht.ProtectContents
    
    If wasProtected Then
        sht.Unprotect PASSWORD:=""
        Debug.Print "ToggleNettoBrutto_Click: Blattschutz aufgehoben."
    End If
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If ToggleNettoBrutto.Value = True Then
        Call SetModus("NETTO")
    Else
        Call SetModus("BRUTTO")
    End If
    
    Application.Calculate ' Sicherstellen, dass alle Excel-Formeln aktualisiert werden
    
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If wasProtected Then
        sht.Protect PASSWORD:="", AllowFormattingCells:=True
        Debug.Print "ToggleNettoBrutto_Click: Blattschutz wiederhergestellt."
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "Fehler beim Toggle-Button: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Private Sub SetModus(modus As String)
    Dim r As Long
    
    Debug.Print "SetModus(" & UCase(modus) & ") gestartet. Blatt geschützt? " & Me.ProtectContents
    
    For r = 2 To 3 ' Nur für die Zeilen 2 und 3
        If UCase(modus) = "NETTO" Then
            ' NETTO (K) ist Eingabe, MWST (L) und BRUTTO (M) sind Ausgabe
            
            Debug.Print "SetModus(" & r & "): Zelle K" & r & " - Locked=False, Farbe Eingabe"
            Me.Cells(r, "K").Locked = False: Me.Cells(r, "K").Interior.color = RGB(169, 208, 142) ' Hellgrün (Eingabe)
            
            Debug.Print "SetModus(" & r & "): Zelle L" & r & " - Locked=True, Farbe Ausgabe"
            Me.Cells(r, "L").Locked = True: Me.Cells(r, "L").Interior.color = RGB(244, 176, 132)  ' Orange (Ausgabe)
            
            Debug.Print "SetModus(" & r & "): Zelle M" & r & " - Locked=True, Farbe Ausgabe"
            Me.Cells(r, "M").Locked = True: Me.Cells(r, "M").Interior.color = RGB(244, 176, 132)  ' Orange (Ausgabe)
        Else ' BRUTTO Modus
            ' BRUTTO (M) ist Eingabe, NETTO (K) und MWST (L) sind Ausgabe
            
            Debug.Print "SetModus(" & r & "): Zelle M" & r & " - Locked=False, Farbe Eingabe"
            Me.Cells(r, "M").Locked = False: Me.Cells(r, "M").Interior.color = RGB(169, 208, 142) ' Hellgrün (Eingabe)
            
            Debug.Print "SetModus(" & r & "): Zelle L" & r & " - Locked=True, Farbe Ausgabe"
            Me.Cells(r, "L").Locked = True: Me.Cells(r, "L").Interior.color = RGB(244, 176, 132)  ' Orange (Ausgabe)
            
            Debug.Print "SetModus(" & r & "): Zelle K" & r & " - Locked=True, Farbe Ausgabe"
            Me.Cells(r, "K").Locked = True: Me.Cells(r, "K").Interior.color = RGB(244, 176, 132)  ' Orange (Ausgabe)
        End If
    Next r
    
    ' Nach der Modus-Umschaltung:
    ' 1. Netto/Brutto-Berechnung auslösen
    If UCase(modus) = "NETTO" Then
        Call RechneAlleZeilen("NETTO")
    Else
        Call RechneAlleZeilen("BRUTTO")
    End If
    
    ' 2. Zählerstands-Berechnung auslösen
    Call UpdateStromblatt
    Debug.Print "SetModus beendet."
End Sub

Private Sub RechneAlleZeilen(modus As String)
    Dim r As Long
    Dim mwstSatz As Double
    Dim nettoVal As Variant, bruttoVal As Variant, mwstVal As Double
    
    ' MwSt-Satz in B5 lesen
    If IsNumeric(Me.Range("B5").Value) Then
        mwstSatz = CDbl(Me.Range("B5").Value)
        If mwstSatz >= 1 And mwstSatz <= 100 Then
            mwstSatz = mwstSatz / 100
        End If
    Else
        ' Dieser Fehler würde nicht 1004 auslösen
        ' MsgBox "Der MwSt-Satz in B5 ist ungültig. Bitte korrigieren Sie ihn.", vbExclamation
        Exit Sub
    End If
    
    For r = 2 To 3
        ' Prüfung auf Zell-Zusammenführung beibehalten
        If Me.Cells(r, "K").MergeCells Or Me.Cells(r, "L").MergeCells Or Me.Cells(r, "M").MergeCells Then
            MsgBox "Zeile " & r & ": Bitte entferne jegliche Zell-Zusammenführung (MergeCells) in Spalte K, L oder M.", vbCritical
            Exit Sub
        End If
        
        nettoVal = Me.Cells(r, "K").Value
        bruttoVal = Me.Cells(r, "M").Value
        
        If UCase(modus) = "NETTO" Then
            ' Eingabe in NETTO (K)
            If IsNumeric(nettoVal) And Trim(CStr(nettoVal)) <> "" And CDbl(nettoVal) <> 0 Then
                mwstVal = CDbl(nettoVal) * mwstSatz
                Me.Cells(r, "L").Value = mwstVal
                Me.Cells(r, "M").Value = CDbl(nettoVal) + mwstVal
            Else
                Me.Cells(r, "L").ClearContents
                Me.Cells(r, "M").ClearContents
            End If
        ElseIf UCase(modus) = "BRUTTO" Then
            ' Eingabe in BRUTTO (M)
            If IsNumeric(bruttoVal) And Trim(CStr(bruttoVal)) <> "" And CDbl(bruttoVal) <> 0 Then
                nettoVal = CDbl(bruttoVal) / (1 + mwstSatz)
                mwstVal = CDbl(bruttoVal) - nettoVal
                Me.Cells(r, "K").Value = nettoVal
                Me.Cells(r, "L").Value = mwstVal
            Else
                Me.Cells(r, "K").ClearContents
                Me.Cells(r, "L").ClearContents
            End If
        End If
    Next r
End Sub

' ==========================================================
' ZÄHLERWECHSEL STARTEN
' ==========================================================
Private Sub btn_neuerZaehler_Strom_Click()
    On Error GoTo ErrHandler
    mod_ZaehlerLogik.Start_Zaehlerwechsel "Strom"
    Exit Sub
ErrHandler:
    MsgBox "Fehler beim Öffnen des Zählerwechsel-Formulars: " & Err.Description, vbExclamation
End Sub

' ==========================================================
' HAUPT-UPDATE-ROUTINE
' ==========================================================
Public Sub UpdateStromblatt()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    ' Führt Berechnung und Formatierung durch (in mod_ZaehlerLogik)
    Call mod_ZaehlerLogik.CalculateAllZaehlerVerbrauch(Me)
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "FEHLER BEI BERECHNUNG Strom: " & Err.Description, vbCritical
End Sub

' ==========================================================
' OPTIONALE ROUTINEN (Historie-Farben)
' ==========================================================
Public Sub KorrigiereHistorieFarben()
    Call mod_ZaehlerLogik.FarbeHistorieEintraege
End Sub
