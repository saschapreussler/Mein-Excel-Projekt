VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Zaehlerwechsel 
   Caption         =   "neuer Zähler"
   ClientHeight    =   4700
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5760
   OleObjectBlob   =   "frm_Zaehlerwechsel.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frm_Zaehlerwechsel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==========================================================
' VARIABLEN
' ==========================================================
Public m_Medium As String
Private m_targetRow As Long
' Der Dezimaltrenner muss hier bekannt sein für die Formatierung
Private Const DECIMAL_SEP As String = "," ' Standard für DE-Excel-UI

' ==========================================================
' INIT
' ==========================================================
Public Sub InitForm_Runtime(ByVal Medium As String)

    m_Medium = Medium

    ' txt_StandAlt darf NIEMALS numerisch interpretiert werden
    With Me.txt_StandAlt
        .ControlSource = ""
        .Locked = False ' !!! KORRIGIERT: Muss bearbeitbar sein !!!
        .text = ""
    End With

    Dim einheit As String
    Dim color As Long

    Select Case Medium
        Case "Strom"
            einheit = "kWh"
            color = RGB(255, 0, 0)
        Case "Wasser"
            einheit = "m³"
            color = RGB(0, 0, 255)
        Case Else
            einheit = "---"
            color = RGB(0, 0, 0)
    End Select

    Me.fra_Header.Caption = "Zählerwechsel erfassen (" & Medium & ")"
    Me.fra_Header.ForeColor = color

    ' Zuweisung der Einheit für das ALT-Feld und das NEU-Start-Feld
    Me.lbl_EinheitAlt.Caption = einheit
    Me.lbl_EinheitNeuStart.Caption = einheit
 
    ' INITIALISIERUNG: Textbox ausblenden
    Me.txt_Bemerkung.Visible = False
    Me.txt_Bemerkung.text = ""
    Me.chk_Bemerkung.value = False

    Me.txt_Datum.text = Format(Date, "dd.mm.yyyy")

    Populate_cmb_Parzelle
    Me.cmb_Parzelle.SetFocus

End Sub

' ==========================================================
' PARZELLENLISTE
' ==========================================================
Private Sub Populate_cmb_Parzelle()

    Dim i As Long
    Me.cmb_Parzelle.Clear

    For i = 1 To 14
        Me.cmb_Parzelle.AddItem "Parzelle " & i
    Next i

    If m_Medium = "Strom" Then
        Me.cmb_Parzelle.AddItem "Clubwagen"
        Me.cmb_Parzelle.AddItem "Kühltruhe"
    End If

    Me.cmb_Parzelle.AddItem "Hauptzähler"

End Sub

' ==========================================================
' PARZELLE -> VORBELEGUNG
' ==========================================================
Private Sub cmb_Parzelle_Change()

    Dim ws As Worksheet
    Dim standAltValue As Variant
    
    ' Code-Namen verwenden
    If m_Medium = "Strom" Then
        Set ws = Tabelle5 ' Code-Name für das Strom-Blatt
    ElseIf m_Medium = "Wasser" Then
        Set ws = Tabelle6 ' Code-Name für das Wasser-Blatt
    Else
        Exit Sub ' Falls Medium unbekannt
    End If

    ' Zeile bestimmen
    Select Case Me.cmb_Parzelle.value
        Case "Clubwagen"
            m_targetRow = 22
        Case "Kühltruhe"
            m_targetRow = 23
        Case "Hauptzähler"
            m_targetRow = IIf(m_Medium = "Strom", 26, 29)
        Case Else
            If Left(Me.cmb_Parzelle.value, 8) = "Parzelle" Then
                Dim idx As Long
                idx = Val(Mid(Me.cmb_Parzelle.value, 10))
                m_targetRow = IIf(m_Medium = "Strom", idx + 7, idx + 9)
            Else
                Exit Sub
            End If
    End Select

    ' ===== STAND ALT – EXAKT: C-Spalte der Hauptblätter =====
    ' Wir lesen den Rohwert (der potentiell Dezimalstellen hat)
    standAltValue = ws.Cells(m_targetRow, "C").value
    
    If IsNumeric(standAltValue) Then
        ' NEU: CleanNumber verwenden, um unnötige .0 oder ,0 zu entfernen
        Me.txt_StandAlt.text = CleanAndFormatNumber(standAltValue)
    Else
        Me.txt_StandAlt.text = "0"
    End If

    Me.txt_ZaehlerAlt.text = ""
    Me.txt_StandNeuStart.text = "0"
    ' NEU: Formatierung nach der Zuweisung
    Call FormatNumberInput_Enhanced(Me.txt_StandNeuStart)

    Me.txt_ZaehlerNeu.text = ""

End Sub

' ==========================================================
' NEU: ERWEITERTE FORMATIERUNG FÜR ZAHLEN UND DEZIMALEN
' (Ersetzt FormatTausender und FormatNumberInput)
' ==========================================================
Private Function CleanAndFormatNumber(ByVal v As Variant) As String
    ' 1. Wert durch die deterministische Funktion aus mod_ZaehlerLogik bereinigen (String)
    Dim cleanStr As String
    cleanStr = mod_ZaehlerLogik.CleanNumber(v)
    
    ' 2. Konvertiere Dezimaltrenner (falls vorhanden) zu UI-Konvention
    If InStr(cleanStr, Application.International(xlDecimalSeparator)) > 0 Then
        cleanStr = Replace(cleanStr, Application.International(xlDecimalSeparator), DECIMAL_SEP)
    End If
    
    ' 3. Führe Tausender-Formatierung nur für den Vorkomma-Teil durch (optional)
    Dim parts() As String
    Dim preDecimal As String
    Dim postDecimal As String
    
    If InStr(cleanStr, DECIMAL_SEP) > 0 Then
        parts = Split(cleanStr, DECIMAL_SEP)
        preDecimal = parts(0)
        postDecimal = parts(1)
    Else
        preDecimal = cleanStr
        postDecimal = ""
    End If
    
    ' Tausenderformatierung für den Vorkomma-Teil
    Dim i As Long
    Dim tempPreDecimal As String
    tempPreDecimal = ""
    
    For i = Len(preDecimal) To 1 Step -1
        tempPreDecimal = Mid$(preDecimal, i, 1) & tempPreDecimal
        If ((Len(preDecimal) - i + 1) Mod 3 = 0) And i > 1 And Mid$(preDecimal, i, 1) <> "-" Then
            tempPreDecimal = "." & tempPreDecimal
        End If
    Next i
    
    ' 4. Ergebnis zusammensetzen
    If postDecimal <> "" Then
        CleanAndFormatNumber = tempPreDecimal & DECIMAL_SEP & postDecimal
    Else
        CleanAndFormatNumber = tempPreDecimal
    End If
End Function

Private Sub FormatNumberInput_Enhanced(ByVal TxtBox As MSForms.TextBox)

    Dim rawText As String
    
    ' Text reinigen: Tausenderpunkte entfernen, dann UI-Dezimaltrenner durch VBA/Excel-Dezimaltrenner ersetzen
    rawText = Replace(TxtBox.text, ".", "")
    rawText = Replace(rawText, DECIMAL_SEP, Application.International(xlDecimalSeparator))

    If rawText = "" Then
        TxtBox.text = "0"
        Exit Sub
    End If

    ' Wenn der Text numerisch ist
    If IsNumeric(rawText) Then
        ' NEU: Bereinigen und neu formatieren
        TxtBox.text = CleanAndFormatNumber(CDbl(rawText))
    Else
        ' Fallback, wenn nach der Reinigung nicht numerisch (z.B. nur "..." eingegeben)
        TxtBox.text = "0"
    End If

End Sub

' ==========================================================
' EXIT EVENTS
' ==========================================================
Private Sub txt_StandAlt_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' NEU: Erweiterte Formatierung, die CleanNumber im Hintergrund aufruft, falls der Wert korrigiert wurde
    Call FormatNumberInput_Enhanced(Me.txt_StandAlt)
End Sub

Private Sub txt_StandNeuStart_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' NEU: Erweiterte Formatierung, die CleanNumber im Hintergrund aufruft
    Call FormatNumberInput_Enhanced(Me.txt_StandNeuStart)
End Sub


' ==========================================================
' CHECKBOX EVENT (STEUERT SICHTBARKEIT)
' ==========================================================
Private Sub chk_Bemerkung_Click()
    Me.txt_Bemerkung.Visible = Me.chk_Bemerkung.value
    
    If Me.txt_Bemerkung.Visible Then
        Me.txt_Bemerkung.SetFocus
    Else
        Me.txt_Bemerkung.text = ""
    End If
End Sub

' ==========================================================
' SPEICHERN (KORRIGIERT: Logik für editierbaren Alt-Stand und Plausibilitätsprüfung)
' ==========================================================
Private Sub Btn_Speichern_Click()

    Dim ws As Worksheet
    Dim standAltOriginal As Double ' Originaler Wert aus Spalte C des Hauptblatts
    Dim standAltUser As Double     ' Der (korrigierte) Wert aus der Textbox txt_StandAlt
    Dim standNeuStart_Raw As Double
    Dim standNeuStart_Final As Double ' Bereinigter Wert für die Übergabe
    
    ' Die Fehlerbehandlung wird zuerst aktiviert
    On Error GoTo SpeichernErrHandler

    ' Zielblatt setzen
    If m_Medium = "Strom" Then
        Set ws = Tabelle5
    ElseIf m_Medium = "Wasser" Then
        Set ws = Tabelle6
    Else
        MsgBox "Interner Fehler: Unbekanntes Medium beim Speichern.", vbCritical
        Exit Sub
    End If
    
    ' 1. Plausibilitätsprüfungen (unverändert)
    If Me.cmb_Parzelle.value = "" Then
        MsgBox "Bitte wählen Sie eine Parzelle/einen Zähler aus.", vbExclamation
        Me.cmb_Parzelle.SetFocus
        Exit Sub
    End If
    
    If Not mod_ZaehlerLogik.PlausiDatum(Me.txt_Datum.text) Then
        MsgBox "Das eingegebene Datum ist ungültig. Bitte korrigieren. (Prüfen Sie auch das Format tt.mm.jjjj).", vbExclamation
        Me.txt_Datum.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.txt_ZaehlerAlt.text) = "" Or Trim(Me.txt_ZaehlerNeu.text) = "" Then
        If MsgBox("Achtung: Haben Sie Zählernummer Alt/Neu vergessen einzugeben? Trotzdem speichern?", vbYesNo + vbQuestion) = vbNo Then
            Me.txt_ZaehlerAlt.SetFocus
            Exit Sub
        End If
    End If

    ' 2. Daten einlesen und konvertieren (KERNKORREKTUREN)
    
    ' a) Originaler Stand Alt: Liest den Endstand des alten Zählers (der in C steht)
    ' Dient als Grundlage für die Plausibilitätsprüfung des User-korrigierten Wertes.
    standAltOriginal = ws.Cells(m_targetRow, "C").value
    
    ' b) Korrigierter Stand Alt: Liest den Wert aus der Textbox txt_StandAlt und bereinigt ihn
    Dim rawTextAlt As String
    rawTextAlt = Replace(Me.txt_StandAlt.text, ".", "") ' Tausenderpunkt entfernen
    rawTextAlt = Replace(rawTextAlt, DECIMAL_SEP, Application.International(xlDecimalSeparator)) ' UI-Dezimal in VBA-Dezimal konvertieren

    If Not IsNumeric(rawTextAlt) Then
         MsgBox "Der Stand Alt (Ende) ist ungültig oder enthält unzulässige Zeichen.", vbExclamation
         Me.txt_StandAlt.SetFocus
         Exit Sub
    End If
    
    ' Deterministische Bereinigung des Finalwerts
    standAltUser = CDbl(mod_ZaehlerLogik.CleanNumber(CDbl(rawTextAlt)))
    
    ' !!! NEUE PLAUSIBILITÄTSPRÜFUNG: STAND ALT KORRIGIERT GEGEN ORIGINAL !!!
    If standAltUser < standAltOriginal Then
        If MsgBox("Achtung: Der korrigierte Stand Alt (" & Format(standAltUser, "0.####") & _
                  ") ist KLEINER als der letzte Stand aus dem Ableseblatt (" & Format(standAltOriginal, "0.####") & _
                  "). Sind Sie sicher, dass dieser niedrigere Wert korrekt ist?", vbYesNo + vbExclamation) = vbNo Then
            Me.txt_StandAlt.SetFocus
            Exit Sub
        End If
    End If
    
    ' c) Stand Neu Start: Liest den Wert aus der Textbox und bereinigt ihn
    Dim rawTextNeu As String
    rawTextNeu = Replace(Me.txt_StandNeuStart.text, ".", "") ' Tausenderpunkt entfernen
    rawTextNeu = Replace(rawTextNeu, DECIMAL_SEP, Application.International(xlDecimalSeparator)) ' UI-Dezimal in VBA-Dezimal konvertieren
    
    If Not IsNumeric(rawTextNeu) Then
        MsgBox "Der Stand Neu (Start) ist ungültig oder enthält unzulässige Zeichen.", vbExclamation
        Me.txt_StandNeuStart.SetFocus
        Exit Sub
    End If
    
    standNeuStart_Raw = CDbl(rawTextNeu)
    standNeuStart_Final = CDbl(mod_ZaehlerLogik.CleanNumber(standNeuStart_Raw))
    
    ' Optional: Textboxen mit dem endgültigen, bereinigten String aktualisieren
    Me.txt_StandAlt.text = CleanAndFormatNumber(standAltUser)
    Me.txt_StandNeuStart.text = CleanAndFormatNumber(standNeuStart_Final)
    
    
    ' 3. Plausibilitätsprüfung für Stände (NEU GEGEN KORRIGIERTEN ALT)
    If standNeuStart_Final < 0 Then
        MsgBox "Der Startstand des neuen Zählers darf nicht negativ sein.", vbExclamation
        Me.txt_StandNeuStart.SetFocus
        Exit Sub
    End If
    
    ' !!! WARNUNG, WENN DER NEUE ZÄHLERSTAND GRÖSSER IST ALS DER KORRIGIERTE ALTE ZÄHLERSTAND !!!
    ' Wichtig: Wir vergleichen gegen standAltUser (den korrigierten Wert)!
    If standNeuStart_Final > standAltUser Then
        If MsgBox("Achtung: Der neue Zählerstand (" & Format(standNeuStart_Final, "0.####") & ") ist GRÖSSER als der Endstand des alten Zählers (" & Format(standAltUser, "0.####") & "). Fortfahren?", vbYesNo + vbExclamation) = vbNo Then
            Me.txt_StandNeuStart.SetFocus
            Exit Sub
        End If
    End If
    
    ' 4. Aufruf der Logik (Argumente sind korrekt)
    ' WICHTIG: Wir übergeben standAltUser (den korrigierten Wert) als AltEnde.
    Call mod_ZaehlerLogik.SchreibeHistorie( _
        parzelle:=Me.cmb_Parzelle.value, _
        DatumW:=CDate(Me.txt_Datum.text), _
        AltEnde:=standAltUser, _
        neuStart:=standNeuStart_Final, _
        snNeu:=Trim(Me.txt_ZaehlerNeu.text), _
        snAlt:=Trim(Me.txt_ZaehlerAlt.text), _
        bem:=Me.txt_Bemerkung.text, _
        Medium:=m_Medium)

    ' Wenn SchreibeHistorie KEINEN Fehler ausgelöst hat, kommt der Code hier an.
    MsgBox "Zählerwechsel erfolgreich gespeichert.", vbInformation
    Unload Me
    
    Exit Sub ' Erfolgreicher Ausgang

' --- FEHLERBEHANDLUNG ---
SpeichernErrHandler:
    ' Zeigt einen Fehler an, der im UserForm selbst oder in der Sub mod_ZaehlerLogik.SchreibeHistorie aufgetreten ist.
    MsgBox "Ein unerwarteter Fehler ist aufgetreten: " & Err.Description & vbCrLf & _
           "FEHLER CODE: " & Err.Number & vbCrLf & _
           "Vorgang abgebrochen.", vbCritical
    
    Exit Sub
    
End Sub

' ==========================================================
' ABBRECHEN
' ==========================================================
Private Sub Btn_Abbrechen_Click()
    Unload Me
End Sub

