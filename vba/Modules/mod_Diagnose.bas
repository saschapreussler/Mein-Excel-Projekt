Attribute VB_Name = "mod_Diagnose"
Option Explicit

' ===============================================================
' MODUL: mod_Diagnose
' VERSION: 1.0 - 19.04.2026
' ZWECK: Diagnose-Tool zum Testen aller neuen Features
'        Jeder Test zeigt ein MsgBox-Ergebnis
'
' ANLEITUNG:
'   1. VBE oeffnen (Alt+F11)
'   2. Im Direktfenster (Strg+G) eingeben:
'      mod_Diagnose.DiagnoseAlles
'   3. Oder einzelne Tests ausfuehren:
'      mod_Diagnose.Test_01_ModulCheck
'      mod_Diagnose.Test_02_BlattNamen
'      etc.
' ===============================================================

Private m_log As String
Private m_fehler As Long
Private m_ok As Long


' ===============================================================
' HAUPTPROZEDUR: Alle Tests ausfuehren
' ===============================================================
Public Sub DiagnoseAlles()
    m_log = ""
    m_fehler = 0
    m_ok = 0
    
    LogZeile "=========================================="
    LogZeile "  DIAGNOSE - Kassenbuch v6.1"
    LogZeile "  Datum: " & Format$(Now, "dd.mm.yyyy hh:nn:ss")
    LogZeile "=========================================="
    LogZeile ""
    
    Call Test_01_ModulCheck
    Call Test_02_BlattNamen
    Call Test_03_TabellenCodenames
    Call Test_04_EinstellungenWerte
    Call Test_05_Startseite
    Call Test_06_BankkontoFormeln
    Call Test_07_VereinskasseComboBox
    Call Test_08_NavigationsModule
    
    LogZeile ""
    LogZeile "=========================================="
    LogZeile "  ERGEBNIS: " & m_ok & " OK / " & m_fehler & " FEHLER"
    LogZeile "=========================================="
    
    ' Ergebnis anzeigen
    Dim titel As String
    If m_fehler = 0 Then
        titel = ChrW(9989) & " Alle " & m_ok & " Tests bestanden!"
    Else
        titel = ChrW(10060) & " " & m_fehler & " Fehler gefunden!"
    End If
    
    MsgBox m_log, IIf(m_fehler > 0, vbExclamation, vbInformation), titel
    
    ' Auch ins Direktfenster schreiben
    Debug.Print m_log
End Sub


' ===============================================================
' TEST 01: Pruefen ob alle erwarteten Module existieren
' ===============================================================
Public Sub Test_01_ModulCheck()
    LogZeile "--- TEST 01: Module im VBA-Projekt ---"
    
    Dim modulNamen As Variant
    modulNamen = Array( _
        "mod_Const", _
        "mod_Startseite", _
        "mod_Navigation", _
        "mod_Einstellungen", _
        "mod_Banking_Format", _
        "mod_Vereinskasse_Filter", _
        "mod_Formatierung", _
        "mod_Hilfsfunktionen", _
        "mod_Format_Protection", _
        "mod_Mitglieder_UI", _
        "mod_Banking_Report")
    
    Dim i As Long
    Dim comp As Object
    Dim gefunden As Boolean
    
    For i = LBound(modulNamen) To UBound(modulNamen)
        gefunden = False
        On Error Resume Next
        Set comp = ThisWorkbook.VBProject.VBComponents(CStr(modulNamen(i)))
        If Not comp Is Nothing Then gefunden = True
        Set comp = Nothing
        Err.Clear
        On Error GoTo 0
        
        If gefunden Then
            LogOK CStr(modulNamen(i)) & " gefunden"
        Else
            LogFEHLER CStr(modulNamen(i)) & " FEHLT! Modul muss importiert werden."
        End If
    Next i
End Sub


' ===============================================================
' TEST 02: Pruefen ob Blattnamen korrekt sind
' ===============================================================
Public Sub Test_02_BlattNamen()
    LogZeile ""
    LogZeile "--- TEST 02: Blattnamen ---"
    
    Dim namen As Variant
    namen = Array( _
        WS_BANKKONTO, _
        WS_DATEN, _
        WS_MITGLIEDER, _
        WS_EINSTELLUNGEN, _
        WS_VEREINSKASSE, _
        WS_STARTMENUE())
    
    Dim beschreibungen As Variant
    beschreibungen = Array( _
        "Bankkonto", _
        "Daten", _
        "Mitgliederliste", _
        "Einstellungen", _
        "Vereinskasse", _
        "Startmen" & ChrW(252))
    
    Dim i As Long
    Dim ws As Worksheet
    
    For i = LBound(namen) To UBound(namen)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(namen(i)))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            LogOK CStr(beschreibungen(i)) & " -> Tab """ & ws.Name & """ (Code: " & ws.CodeName & ")"
        Else
            LogFEHLER CStr(beschreibungen(i)) & " -> Blatt """ & CStr(namen(i)) & """ NICHT GEFUNDEN!"
        End If
    Next i
    
    ' Alle vorhandenen Blaetter auflisten
    LogZeile "  Vorhandene Bl" & ChrW(228) & "tter:"
    For Each ws In ThisWorkbook.Worksheets
        LogZeile "    -> """ & ws.Name & """ (CodeName: " & ws.CodeName & ")"
    Next ws
End Sub


' ===============================================================
' TEST 03: Tabellen-Codenamen pruefen
'          Besonders wichtig: Welche Tabelle = Vereinskasse?
' ===============================================================
Public Sub Test_03_TabellenCodenames()
    LogZeile ""
    LogZeile "--- TEST 03: Tabellen-Codenamen ---"
    
    Dim ws As Worksheet
    
    ' Vereinskasse-Blatt finden
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_VEREINSKASSE)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        LogOK "Vereinskasse -> CodeName = """ & ws.CodeName & """"
        
        If ws.CodeName = "Tabelle4" Then
            LogOK "Tabelle4.cls ist korrekt f" & ChrW(252) & "r Vereinskasse"
        Else
            LogFEHLER "Vereinskasse hat CodeName """ & ws.CodeName & """, aber Events stehen in Tabelle4.cls!" & vbLf & _
                      "  LOESUNG: Event-Code von Tabelle4 nach " & ws.CodeName & " verschieben"
        End If
    Else
        LogFEHLER "Blatt 'Vereinskasse' existiert nicht!"
    End If
    
    ' Startmenue-Blatt pruefen
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_STARTMENUE())
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        LogOK "Startmen" & ChrW(252) & " -> CodeName = """ & ws.CodeName & """"
    Else
        LogFEHLER "Blatt '" & WS_STARTMENUE() & "' existiert nicht!"
    End If
End Sub


' ===============================================================
' TEST 04: Einstellungen-Werte pruefen
' ===============================================================
Public Sub Test_04_EinstellungenWerte()
    LogZeile ""
    LogZeile "--- TEST 04: Einstellungen-Werte ---"
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    
    If ws Is Nothing Then
        LogFEHLER "Blatt 'Einstellungen' nicht gefunden!"
        Exit Sub
    End If
    
    ' Abrechnungsjahr (C4)
    Dim abrJahr As Variant
    abrJahr = ws.Cells(ES_CFG_ABRECHNUNGSJAHR_ROW, ES_CFG_VALUE_COL).value
    If IsNumeric(abrJahr) And abrJahr <> "" Then
        If CLng(abrJahr) >= 2000 And CLng(abrJahr) <= 2100 Then
            LogOK "Abrechnungsjahr (C" & ES_CFG_ABRECHNUNGSJAHR_ROW & "): " & abrJahr & _
                  " -> InputBox wird NICHT erscheinen (Wert vorhanden)"
        Else
            LogFEHLER "Abrechnungsjahr (C" & ES_CFG_ABRECHNUNGSJAHR_ROW & "): " & abrJahr & " (ungueltig)"
        End If
    Else
        LogOK "Abrechnungsjahr (C" & ES_CFG_ABRECHNUNGSJAHR_ROW & "): LEER -> InputBox WIRD erscheinen"
    End If
    
    ' Kontostand (C5)
    Dim kontostand As Variant
    kontostand = ws.Cells(ES_CFG_KONTOSTAND_ROW, ES_CFG_VALUE_COL).value
    If IsNumeric(kontostand) And kontostand <> "" Then
        If CDbl(kontostand) <> 0 Then
            LogOK "Kontostand (C" & ES_CFG_KONTOSTAND_ROW & "): " & Format$(CDbl(kontostand), "#,##0.00") & _
                  " -> InputBox wird NICHT erscheinen (Wert vorhanden)"
        Else
            LogOK "Kontostand (C" & ES_CFG_KONTOSTAND_ROW & "): 0 -> InputBox WIRD erscheinen"
        End If
    Else
        LogOK "Kontostand (C" & ES_CFG_KONTOSTAND_ROW & "): LEER -> InputBox WIRD erscheinen"
    End If
    
    ' Vereinsname (C16)
    Dim vName As String
    vName = Trim(CStr(ws.Cells(ES_CFG_VEREINSNAME_ROW, ES_CFG_VALUE_COL).value))
    If vName <> "" Then
        LogOK "Vereinsname (C" & ES_CFG_VEREINSNAME_ROW & "): """ & vName & """" & _
              " -> InputBox wird NICHT erscheinen"
    Else
        LogOK "Vereinsname (C" & ES_CFG_VEREINSNAME_ROW & "): LEER -> InputBox WIRD erscheinen"
    End If
    
    ' Adresse
    Dim strasse As String, plz As String, ort As String
    strasse = Trim(CStr(ws.Cells(ES_CFG_STRASSE_ROW, ES_CFG_VALUE_COL).value))
    plz = Trim(CStr(ws.Cells(ES_CFG_PLZ_ORT_ROW, ES_CFG_VALUE_COL).value))
    ort = Trim(CStr(ws.Cells(ES_CFG_PLZ_ORT_ROW, 5).value))
    LogZeile "  Strasse (C" & ES_CFG_STRASSE_ROW & "): """ & strasse & """"
    LogZeile "  PLZ (C" & ES_CFG_PLZ_ORT_ROW & "): """ & plz & """"
    LogZeile "  Ort (E" & ES_CFG_PLZ_ORT_ROW & "): """ & ort & """"
End Sub


' ===============================================================
' TEST 05: Startseite manuell initialisieren (mit Fehleranzeige!)
' ===============================================================
Public Sub Test_05_Startseite()
    LogZeile ""
    LogZeile "--- TEST 05: Startseite ---"
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_STARTMENUE())
    On Error GoTo 0
    
    If ws Is Nothing Then
        LogFEHLER "Blatt '" & WS_STARTMENUE() & "' nicht gefunden!"
        Exit Sub
    End If
    
    LogOK "Blatt gefunden: " & ws.Name & " (Code: " & ws.CodeName & ")"
    
    ' Pruefen ob InitialisiereStartseite aufrufbar ist
    On Error Resume Next
    Err.Clear
    Call mod_Startseite.InitialisiereStartseite
    
    If Err.Number <> 0 Then
        LogFEHLER "InitialisiereStartseite FEHLER " & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        LogOK "InitialisiereStartseite erfolgreich ausgef" & ChrW(252) & "hrt"
    End If
    On Error GoTo 0
End Sub


' ===============================================================
' TEST 06: Bankkonto E4-Formel pruefen
' ===============================================================
Public Sub Test_06_BankkontoFormeln()
    LogZeile ""
    LogZeile "--- TEST 06: Bankkonto-Formeln ---"
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo 0
    
    If ws Is Nothing Then
        LogFEHLER "Blatt 'Bankkonto' nicht gefunden!"
        Exit Sub
    End If
    
    ' E4 pruefen
    Dim formelE2 As String
    On Error Resume Next
    formelE2 = ws.Range("E4").FormulaLocal
    On Error GoTo 0
    
    If formelE2 = "" Then
        LogFEHLER "E4 ist leer (keine Formel)"
    ElseIf InStr(formelE2, "Startmen") > 0 Then
        LogFEHLER "E4 referenziert noch 'Startmen" & ChrW(252) & "' statt 'Einstellungen'!"
        LogZeile "  Aktuelle Formel: " & Left$(formelE2, 80) & "..."
        LogZeile "  -> StelleFormelnWiederHer muss ausgef" & ChrW(252) & "hrt werden"
    ElseIf InStr(formelE2, "Einstellungen") > 0 Then
        LogOK "E4 referenziert korrekt 'Einstellungen'"
        LogZeile "  Formel: " & Left$(formelE2, 80) & "..."
    Else
        LogZeile "  E4 Formel: " & Left$(formelE2, 80) & "..."
    End If
    
    ' E4-Formel manuell wiederherstellen
    LogZeile "  -> Stelle E4-Formel jetzt wieder her..."
    On Error Resume Next
    Err.Clear
    Call mod_Banking_Format.StelleFormelnWiederHer(ws)
    If Err.Number <> 0 Then
        LogFEHLER "StelleFormelnWiederHer FEHLER " & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        LogOK "StelleFormelnWiederHer erfolgreich"
        ' Nochmal pruefen
        formelE2 = ws.Range("E4").FormulaLocal
        If InStr(formelE2, "Einstellungen") > 0 Then
            LogOK "E4 jetzt korrekt: " & Left$(formelE2, 60) & "..."
        End If
    End If
    On Error GoTo 0
End Sub


' ===============================================================
' TEST 07: Vereinskasse ComboBox
' ===============================================================
Public Sub Test_07_VereinskasseComboBox()
    LogZeile ""
    LogZeile "--- TEST 07: Vereinskasse ComboBox ---"
    
    Dim wsVK As Worksheet
    On Error Resume Next
    Set wsVK = ThisWorkbook.Worksheets(WS_VEREINSKASSE)
    On Error GoTo 0
    
    If wsVK Is Nothing Then
        LogFEHLER "Blatt 'Vereinskasse' nicht gefunden!"
        Exit Sub
    End If
    
    LogOK "Vereinskasse gefunden (CodeName: " & wsVK.CodeName & ")"
    
    ' Pruefen ob ComboBox existiert
    Dim oleObj As OLEObject
    On Error Resume Next
    Set oleObj = wsVK.OLEObjects("cbo_MonatFilter_VK")
    On Error GoTo 0
    
    If Not oleObj Is Nothing Then
        LogOK "ComboBox 'cbo_MonatFilter_VK' existiert bereits"
    Else
        LogZeile "  ComboBox fehlt - versuche Erstellung..."
        On Error Resume Next
        Err.Clear
        Call mod_Vereinskasse_Filter.InitialisiereVereinskasseComboBox
        If Err.Number <> 0 Then
            LogFEHLER "InitialisiereVereinskasseComboBox FEHLER " & Err.Number & ": " & Err.Description
            Err.Clear
        Else
            ' Pruefen ob jetzt vorhanden
            Set oleObj = Nothing
            Set oleObj = wsVK.OLEObjects("cbo_MonatFilter_VK")
            If Not oleObj Is Nothing Then
                LogOK "ComboBox erfolgreich erstellt"
            Else
                LogFEHLER "ComboBox konnte nicht erstellt werden (kein Fehler aber nicht vorhanden)"
            End If
        End If
        On Error GoTo 0
    End If
    
    ' Vereinskasse CodeName pruefen
    If wsVK.CodeName <> "Tabelle4" Then
        LogFEHLER "WICHTIG: Vereinskasse CodeName ist """ & wsVK.CodeName & """ aber Events stehen in Tabelle4.cls!" & vbLf & _
                  "  Die ComboBox-Events (Worksheet_Activate, cbo_MonatFilter_VK_Change)" & vbLf & _
                  "  muessen in das Sheet-Modul """ & wsVK.CodeName & """ verschoben werden."
    End If
End Sub


' ===============================================================
' TEST 08: Navigations-Module pruefen
' ===============================================================
Public Sub Test_08_NavigationsModule()
    LogZeile ""
    LogZeile "--- TEST 08: Navigation ---"
    
    ' Pruefen ob mod_Navigation existiert und Funktionen hat
    Dim comp As Object
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents("mod_Navigation")
    On Error GoTo 0
    
    If comp Is Nothing Then
        LogFEHLER "mod_Navigation fehlt!"
        Exit Sub
    End If
    
    LogOK "mod_Navigation vorhanden (Zeilen: " & comp.CodeModule.CountOfLines & ")"
    
    ' Home-Buttons testen
    On Error Resume Next
    Err.Clear
    Call mod_Navigation.SetzeHomeButtonsAufAllenBlaettern
    If Err.Number <> 0 Then
        LogFEHLER "SetzeHomeButtonsAufAllenBlaettern FEHLER " & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        LogOK "Home-Buttons gesetzt"
    End If
    On Error GoTo 0
End Sub


' ===============================================================
' EXTRA: Manuell InputBoxen erzwingen (zum Testen)
' ===============================================================
Public Sub Test_InputBox_Abrechnungsjahr()
    ' Setzt C4 auf leer, dann ruft PruefeAbrechnungsjahr auf
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Einstellungen-Blatt nicht gefunden!", vbCritical
        Exit Sub
    End If
    
    Dim antwort As VbMsgBoxResult
    antwort = MsgBox( _
        "Dieser Test leert Zelle C" & ES_CFG_ABRECHNUNGSJAHR_ROW & _
        " (Abrechnungsjahr) tempor" & ChrW(228) & "r," & vbLf & _
        "damit die InputBox erscheint." & vbLf & vbLf & _
        "Aktueller Wert: " & ws.Cells(ES_CFG_ABRECHNUNGSJAHR_ROW, ES_CFG_VALUE_COL).value & vbLf & vbLf & _
        "Fortfahren?", _
        vbYesNo + vbQuestion, "Test: InputBox Abrechnungsjahr")
    
    If antwort <> vbYes Then Exit Sub
    
    Dim alter_wert As Variant
    alter_wert = ws.Cells(ES_CFG_ABRECHNUNGSJAHR_ROW, ES_CFG_VALUE_COL).value
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    ws.Cells(ES_CFG_ABRECHNUNGSJAHR_ROW, ES_CFG_VALUE_COL).ClearContents
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
    ' InputBox aufrufen
    Call mod_Einstellungen.PruefeAbrechnungsjahr
    
    ' Pruefen ob Wert gesetzt wurde
    Dim neuer_wert As Variant
    neuer_wert = ws.Cells(ES_CFG_ABRECHNUNGSJAHR_ROW, ES_CFG_VALUE_COL).value
    
    If neuer_wert <> "" Then
        MsgBox "Abrechnungsjahr gesetzt auf: " & neuer_wert, vbInformation
    Else
        ' Alten Wert wiederherstellen
        On Error Resume Next
        ws.Unprotect PASSWORD:=PASSWORD
        ws.Cells(ES_CFG_ABRECHNUNGSJAHR_ROW, ES_CFG_VALUE_COL).value = alter_wert
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        On Error GoTo 0
        MsgBox "Abgebrochen. Alter Wert wiederhergestellt: " & alter_wert, vbInformation
    End If
End Sub


Public Sub Test_InputBox_Kontostand()
    ' Setzt C5 auf leer, dann ruft PruefeKontostandVorjahr auf
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Einstellungen-Blatt nicht gefunden!", vbCritical
        Exit Sub
    End If
    
    Dim antwort As VbMsgBoxResult
    antwort = MsgBox( _
        "Dieser Test leert Zelle C" & ES_CFG_KONTOSTAND_ROW & _
        " (Kontostand Vorjahr) tempor" & ChrW(228) & "r," & vbLf & _
        "damit die InputBox erscheint." & vbLf & vbLf & _
        "Aktueller Wert: " & ws.Cells(ES_CFG_KONTOSTAND_ROW, ES_CFG_VALUE_COL).value & vbLf & vbLf & _
        "Fortfahren?", _
        vbYesNo + vbQuestion, "Test: InputBox Kontostand")
    
    If antwort <> vbYes Then Exit Sub
    
    Dim alter_wert As Variant
    alter_wert = ws.Cells(ES_CFG_KONTOSTAND_ROW, ES_CFG_VALUE_COL).value
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    ws.Cells(ES_CFG_KONTOSTAND_ROW, ES_CFG_VALUE_COL).ClearContents
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
    Call mod_Einstellungen.PruefeKontostandVorjahr
    
    Dim neuer_wert As Variant
    neuer_wert = ws.Cells(ES_CFG_KONTOSTAND_ROW, ES_CFG_VALUE_COL).value
    
    If neuer_wert <> "" Then
        MsgBox "Kontostand gesetzt auf: " & Format$(CDbl(neuer_wert), "#,##0.00") & " " & ChrW(8364), vbInformation
    Else
        On Error Resume Next
        ws.Unprotect PASSWORD:=PASSWORD
        ws.Cells(ES_CFG_KONTOSTAND_ROW, ES_CFG_VALUE_COL).value = alter_wert
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        On Error GoTo 0
        MsgBox "Abgebrochen. Alter Wert wiederhergestellt.", vbInformation
    End If
End Sub


Public Sub Test_InputBox_Vereinsdaten()
    ' Setzt C16 auf leer, dann ruft PruefeVereinsdaten auf
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Einstellungen-Blatt nicht gefunden!", vbCritical
        Exit Sub
    End If
    
    Dim antwort As VbMsgBoxResult
    antwort = MsgBox( _
        "Dieser Test leert Zelle C" & ES_CFG_VEREINSNAME_ROW & _
        " (Vereinsname) tempor" & ChrW(228) & "r," & vbLf & _
        "damit die InputBox-Kette erscheint." & vbLf & vbLf & _
        "Aktueller Wert: """ & ws.Cells(ES_CFG_VEREINSNAME_ROW, ES_CFG_VALUE_COL).value & """" & vbLf & vbLf & _
        "Fortfahren?", _
        vbYesNo + vbQuestion, "Test: InputBox Vereinsdaten")
    
    If antwort <> vbYes Then Exit Sub
    
    Dim alter_name As Variant
    alter_name = ws.Cells(ES_CFG_VEREINSNAME_ROW, ES_CFG_VALUE_COL).value
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    ws.Cells(ES_CFG_VEREINSNAME_ROW, ES_CFG_VALUE_COL).ClearContents
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
    Call mod_Einstellungen.PruefeVereinsdaten
    
    Dim neuer_name As Variant
    neuer_name = ws.Cells(ES_CFG_VEREINSNAME_ROW, ES_CFG_VALUE_COL).value
    
    If neuer_name <> "" Then
        MsgBox "Vereinsname gesetzt auf: """ & neuer_name & """", vbInformation
    Else
        On Error Resume Next
        ws.Unprotect PASSWORD:=PASSWORD
        ws.Cells(ES_CFG_VEREINSNAME_ROW, ES_CFG_VALUE_COL).value = alter_name
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
        On Error GoTo 0
        MsgBox "Abgebrochen. Alter Wert wiederhergestellt.", vbInformation
    End If
End Sub


' ===============================================================
' EXTRA: Startseite manuell neu aufbauen (mit sichtbarem Fehler)
' ===============================================================
Public Sub Startseite_Neu_Aufbauen()
    On Error GoTo ErrHandler
    
    MsgBox "Starte Neuaufbau der Startseite..." & vbLf & _
           "Bei Fehler wird eine Meldung angezeigt.", _
           vbInformation, "Startseite"
    
    Call mod_Startseite.InitialisiereStartseite
    
    MsgBox ChrW(9989) & " Startseite erfolgreich aufgebaut!" & vbLf & vbLf & _
           "Wechsle jetzt zum Startmen" & ChrW(252) & " um das Ergebnis zu sehen.", _
           vbInformation, "Startseite"
    
    ' Zum Blatt wechseln
    On Error Resume Next
    ThisWorkbook.Worksheets(WS_STARTMENUE()).Activate
    On Error GoTo 0
    
    Exit Sub

ErrHandler:
    MsgBox ChrW(10060) & " FEHLER beim Aufbau der Startseite!" & vbLf & vbLf & _
           "Fehler " & Err.Number & ": " & Err.Description & vbLf & vbLf & _
           "Zeile: " & Erl, _
           vbCritical, "Startseite - Fehler"
End Sub


' ===============================================================
' EXTRA: Bankkonto E2 manuell reparieren
' ===============================================================
Public Sub BankkontoE2_Reparieren()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Blatt 'Bankkonto' nicht gefunden!", vbCritical
        Exit Sub
    End If
    
    Dim altFormel As String
    altFormel = ws.Range("E2").FormulaLocal
    
    Call mod_Banking_Format.StelleFormelnWiederHer(ws)
    
    Dim neuFormel As String
    neuFormel = ws.Range("E2").FormulaLocal
    
    MsgBox "E2-Formel aktualisiert:" & vbLf & vbLf & _
           "VORHER: " & Left$(altFormel, 100) & vbLf & vbLf & _
           "NACHHER: " & Left$(neuFormel, 100), _
           vbInformation, "Bankkonto E2"
End Sub


' ===============================================================
' EXTRA: Vereinskasse ComboBox manuell erstellen
' ===============================================================
Public Sub VereinskasseComboBox_Erstellen()
    On Error GoTo ErrHandler
    
    Call mod_Vereinskasse_Filter.InitialisiereVereinskasseComboBox
    
    Dim wsVK As Worksheet
    Set wsVK = ThisWorkbook.Worksheets(WS_VEREINSKASSE)
    
    Dim oleObj As OLEObject
    On Error Resume Next
    Set oleObj = wsVK.OLEObjects("cbo_MonatFilter_VK")
    On Error GoTo 0
    
    If Not oleObj Is Nothing Then
        MsgBox ChrW(9989) & " ComboBox erfolgreich erstellt!" & vbLf & vbLf & _
               "Wechsle jetzt zum Blatt 'Vereinskasse'.", _
               vbInformation, "Vereinskasse ComboBox"
    Else
        MsgBox "ComboBox wurde nicht erstellt (kein Fehler aber nicht vorhanden).", _
               vbExclamation, "Vereinskasse ComboBox"
    End If
    
    Exit Sub

ErrHandler:
    MsgBox ChrW(10060) & " FEHLER: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Vereinskasse ComboBox"
End Sub


' ===============================================================
' HELPER: Workbook_Open simulieren (mit Fehleranzeige)
' ===============================================================
Public Sub SimuliereWorkbookOpen()
    Dim protokoll As String
    protokoll = "=== Workbook_Open Simulation ===" & vbLf & vbLf
    
    ' 1. MigriereEinstellungenLayout
    On Error Resume Next
    Err.Clear
    Call mod_Einstellungen.MigriereEinstellungenLayout
    protokoll = protokoll & SchritErgebnis("MigriereEinstellungenLayout", Err)
    
    ' 2. PruefeAbrechnungsjahr
    Err.Clear
    Call mod_Einstellungen.PruefeAbrechnungsjahr
    protokoll = protokoll & SchritErgebnis("PruefeAbrechnungsjahr", Err)
    
    ' 3. PruefeKontostandVorjahr
    Err.Clear
    Call mod_Einstellungen.PruefeKontostandVorjahr
    protokoll = protokoll & SchritErgebnis("PruefeKontostandVorjahr", Err)
    
    ' 4. PruefeVereinsdaten
    Err.Clear
    Call mod_Einstellungen.PruefeVereinsdaten
    protokoll = protokoll & SchritErgebnis("PruefeVereinsdaten", Err)
    
    ' 5. InitialisiereStartseite
    Err.Clear
    Call mod_Startseite.InitialisiereStartseite
    protokoll = protokoll & SchritErgebnis("InitialisiereStartseite", Err)
    
    ' 6. SetzeHomeButtons
    Err.Clear
    Call mod_Navigation.SetzeHomeButtonsAufAllenBlaettern
    protokoll = protokoll & SchritErgebnis("SetzeHomeButtons", Err)
    
    ' 7. Vereinskasse ComboBox
    Err.Clear
    Call mod_Vereinskasse_Filter.InitialisiereVereinskasseComboBox
    protokoll = protokoll & SchritErgebnis("VereinskasseComboBox", Err)
    
    ' 8. Vereinskasse Formeln
    Err.Clear
    Dim wsVK As Worksheet
    Set wsVK = ThisWorkbook.Worksheets(WS_VEREINSKASSE)
    If Not wsVK Is Nothing Then
        Call mod_Vereinskasse_Filter.SetzeVereinskasseFormeln(wsVK)
    End If
    protokoll = protokoll & SchritErgebnis("VereinskasseFormeln", Err)
    
    ' 9. Bankkonto Formeln
    Err.Clear
    Dim wsBK As Worksheet
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    If Not wsBK Is Nothing Then
        Call mod_Banking_Format.StelleFormelnWiederHer(wsBK)
    End If
    protokoll = protokoll & SchritErgebnis("BankkontoFormeln", Err)
    
    On Error GoTo 0
    
    MsgBox protokoll, vbInformation, "Workbook_Open Simulation"
    Debug.Print protokoll
End Sub

Private Function SchritErgebnis(ByVal Name As String, ByVal e As ErrObject) As String
    If e.Number <> 0 Then
        SchritErgebnis = ChrW(10060) & " " & Name & ": FEHLER " & e.Number & " - " & e.Description & vbLf
    Else
        SchritErgebnis = ChrW(9989) & " " & Name & ": OK" & vbLf
    End If
End Function


' ===============================================================
' PRIVATE HELFER
' ===============================================================
Private Sub LogZeile(ByVal text As String)
    m_log = m_log & text & vbLf
End Sub

Private Sub LogOK(ByVal text As String)
    m_log = m_log & "  " & ChrW(9989) & " " & text & vbLf
    m_ok = m_ok + 1
End Sub

Private Sub LogFEHLER(ByVal text As String)
    m_log = m_log & "  " & ChrW(10060) & " " & text & vbLf
    m_fehler = m_fehler + 1
End Sub










