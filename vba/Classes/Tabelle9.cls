VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ===============================================================
' TABELLENBLATT-MODUL: Einstellungen (v5.0)
' Formatiert die Zahlungstermin-Tabelle bei Aenderungen.
' NEU v5.0: Konfigurationsbereich (Zeilen 1-19) auswerten
'           Mitgliedsbeitrag/Pacht -> Zahlungstermine sync
' ===============================================================

Private Sub Worksheet_Activate()
    Call mod_Einstellungen.FormatiereZahlungsterminTabelle(Me)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    
    ' --- Konfigurationsbereich (Zeilen 4-18, Spalte C) ---
    If Target.Cells.CountLarge = 1 Then
        If Target.Column = ES_CFG_VALUE_COL And _
           Target.Row >= ES_CFG_ABRECHNUNGSJAHR_ROW And _
           Target.Row <= ES_CFG_PLZ_ORT_ROW Then
            Call VerarbeiteKonfigAenderung(Target)
            Exit Sub
        End If
        ' PLZ/Ort Spalte D-E
        If Target.Row = ES_CFG_PLZ_ORT_ROW And _
           (Target.Column = 4 Or Target.Column = 5) Then
            Exit Sub  ' Nur Eingabe, keine Aktion noetig
        End If
    End If
    
    ' --- Zahlungstermine-Tabelle (ab Zeile ES_START_ROW) ---
    If Intersect(Target, Me.Range(Me.Cells(ES_START_ROW, ES_COL_START), _
                                   Me.Cells(Me.Rows.count, ES_COL_END))) Is Nothing Then
        Exit Sub
    End If
    
    If Target.Row < ES_START_ROW Then Exit Sub
    If Target.Cells.CountLarge > 1 Then GoTo NurFormatieren
    
    On Error GoTo SafeExit
    Application.EnableEvents = False
    
    On Error Resume Next
    Me.Unprotect PASSWORD:=PASSWORD
    On Error GoTo SafeExit
    
    Dim rowES As Long
    rowES = Target.Row
    
    ' ---------------------------------------------------------------
    ' SPALTE E: Soll-Monat(e) - Eingabe normalisieren
    ' ---------------------------------------------------------------
    If Target.Column = ES_COL_SOLL_MONATE Then
        Dim eingabe As String
        eingabe = Trim(CStr(Target.value))
        
        If eingabe <> "" Then
            ' "alle" / "Alle" -> "01, 02, 03, ..., 12"
            If LCase(eingabe) = "alle" Then
                Target.value = "01, 02, 03, 04, 05, 06, 07, 08, 09, 10, 11, 12"
            Else
                ' Normalisieren: "3,6,9" -> "03, 06, 09"
                Target.value = NormalisiereMonate(eingabe)
            End If
            
            ' D/E vs F Sperre: Wenn E gef?llt, darf F NICHT gef?llt sein
            If Trim(CStr(Me.Cells(rowES, ES_COL_STICHTAG_FIX).value)) <> "" Then
                MsgBox "Spalte F (Soll-Stichtag Fix) ist bereits gef?llt!" & vbLf & _
                       "Entweder Spalte D/E ODER Spalte F verwenden." & vbLf & vbLf & _
                       "Die Eingabe in Spalte E wird zur?ckgesetzt.", vbExclamation, "Konflikt"
                Target.ClearContents
                GoTo WeiterFormatieren
            End If
            
            ' Nur Spalte E gef?llt (kein Tag in D) -> automatisch 31 setzen
            Dim aktuellerTag As String
            aktuellerTag = Trim(CStr(Me.Cells(rowES, ES_COL_SOLL_TAG).value))
            If aktuellerTag = "" Then
                Me.Cells(rowES, ES_COL_SOLL_TAG).value = 31
                MsgBox "Nur Soll-Monat(e) ohne Soll-Tag eingegeben." & vbLf & vbLf & _
                       "Der Soll-Tag wird automatisch auf den 31. gesetzt," & vbLf & _
                       "d.h. es gilt jeweils der letzte Tag im Monat.", _
                       vbInformation, "Letzter Tag im Monat"
            End If
        End If
    End If
    
    ' ---------------------------------------------------------------
    ' SPALTE D: Soll-Tag - D/E vs F Sperre pr?fen
    ' ---------------------------------------------------------------
    If Target.Column = ES_COL_SOLL_TAG Then
        Dim tagWert As String
        tagWert = Trim(CStr(Target.value))
        
        If tagWert <> "" Then
            ' D/E vs F Sperre: Wenn D gef?llt, darf F NICHT gef?llt sein
            If Trim(CStr(Me.Cells(rowES, ES_COL_STICHTAG_FIX).value)) <> "" Then
                MsgBox "Spalte F (Soll-Stichtag Fix) ist bereits gef?llt!" & vbLf & _
                       "Entweder Spalte D/E ODER Spalte F verwenden." & vbLf & vbLf & _
                       "Die Eingabe in Spalte D wird zur?ckgesetzt.", vbExclamation, "Konflikt"
                Target.ClearContents
                GoTo WeiterFormatieren
            End If
        End If
    End If
    
    ' ---------------------------------------------------------------
    ' SPALTE F: Soll-Stichtag (Fix) - D/E vs F Sperre pr?fen
    ' ---------------------------------------------------------------
    If Target.Column = ES_COL_STICHTAG_FIX Then
        Dim stichtagWert As String
        stichtagWert = Trim(CStr(Target.value))
        
        If stichtagWert <> "" Then
            ' F vs D/E Sperre: Wenn F gef?llt, d?rfen D und E NICHT gef?llt sein
            Dim dWert As String
            Dim eWert As String
            dWert = Trim(CStr(Me.Cells(rowES, ES_COL_SOLL_TAG).value))
            eWert = Trim(CStr(Me.Cells(rowES, ES_COL_SOLL_MONATE).value))
            
            If dWert <> "" Or eWert <> "" Then
                MsgBox "Spalte D (Soll-Tag) oder E (Soll-Monate) ist bereits gef?llt!" & vbLf & _
                       "Entweder Spalte D/E ODER Spalte F verwenden." & vbLf & vbLf & _
                       "Die Eingabe in Spalte F wird zur?ckgesetzt.", vbExclamation, "Konflikt"
                Target.ClearContents
                GoTo WeiterFormatieren
            End If
        End If
    End If
    
    ' ---------------------------------------------------------------
    ' SPALTE B: Kategorie - Duplikat pr?fen, sortieren, Cursor setzen
    ' ---------------------------------------------------------------
    If Target.Column = ES_COL_KATEGORIE Then
        Dim katWert As String
        katWert = Trim(CStr(Target.value))
        
        If katWert <> "" Then
            ' Duplikat-Pr?fung: Kommt diese Kategorie bereits in einer anderen Zeile vor?
            Dim pruefZeile As Long
            Dim LetzteZeile As Long
            LetzteZeile = Me.Cells(Me.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
            
            For pruefZeile = ES_START_ROW To LetzteZeile
                If pruefZeile <> rowES Then
                    If StrComp(Trim(CStr(Me.Cells(pruefZeile, ES_COL_KATEGORIE).value)), _
                               katWert, vbTextCompare) = 0 Then
                        MsgBox "Die Kategorie """ & katWert & """ ist bereits in Zeile " & _
                               pruefZeile & " vorhanden!" & vbLf & vbLf & _
                               "Jede Kategorie darf nur einmal vorkommen.", _
                               vbExclamation, "Duplikat"
                        Target.ClearContents
                        GoTo WeiterFormatieren
                    End If
                End If
            Next pruefZeile
            
            Call mod_Einstellungen.SortiereAlphabetisch(Me)
            
            ' Neue Position des Eintrags finden und Cursor setzen
            Dim suchZeile As Long
            Dim letzte As Long
            letzte = Me.Cells(Me.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
            
            For suchZeile = ES_START_ROW To letzte
                If StrComp(Trim(CStr(Me.Cells(suchZeile, ES_COL_KATEGORIE).value)), katWert, vbTextCompare) = 0 Then
                    ' Cursor in die n?chste Spalte (C = Soll-Betrag) setzen
                    Me.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
                    Application.EnableEvents = True
                    Me.Cells(suchZeile, ES_COL_SOLL_BETRAG).Select
                    
                    ' Formatierung trotzdem anwenden (ohne Events)
                    Application.EnableEvents = False
                    Call mod_Einstellungen.FormatiereZahlungsterminTabelle(Me)
                    Application.EnableEvents = True
                    
                    ' Re-Evaluierung ausl?sen
                    Call mod_KategorieEngine_Pipeline.ReEvaluiereAlleNichtManuellen
                    
                    ' v3.0: ?bersicht nach Kategorie-?nderung aktualisieren
                    On Error Resume Next
                    Call mod_Uebersicht_Generator.GeneriereUebersicht(Year(Date), stummModus:=True)
                    Call mod_Uebersicht_Dashboard.GeneriereUebersichtNeu(stummModus:=True)
                    On Error GoTo 0
                    
                    Exit Sub
                End If
            Next suchZeile
        End If
    End If
    
WeiterFormatieren:
    Me.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True

NurFormatieren:
    ' v4.5b: Error-Handler und Events auch bei Multi-Zell-Aenderungen sicherstellen
    On Error GoTo SafeExit
    Application.EnableEvents = False
    
    ' Formatierung aktualisieren
    Call mod_Einstellungen.FormatiereZahlungsterminTabelle(Me)
    
    ' Bankkonto-Kategorien neu evaluieren
    Call mod_KategorieEngine_Pipeline.ReEvaluiereAlleNichtManuellen
    
    ' v3.0: ?bersicht nach Einstellungs-?nderung automatisch aktualisieren
    On Error Resume Next
    Call mod_Uebersicht_Generator.GeneriereUebersicht(Year(Date), stummModus:=True)
    Call mod_Uebersicht_Dashboard.GeneriereUebersichtNeu(stummModus:=True)
    On Error GoTo 0
    
SafeExit:
    Application.EnableEvents = True
    If Err.Number <> 0 Then Err.Clear
    
End Sub


' ===============================================================
' Hilfsfunktion: Monatseingabe normalisieren
' "3,6,9" -> "03, 06, 09"
' "1, 3, 12" -> "01, 03, 12"
' Jeder Monat wird auf 2 Stellen aufgef?llt,
' nach dem Komma folgt ein Leerzeichen.
' Ung?ltige Monate (>12 oder <1) werden ignoriert.
' ===============================================================
Private Function NormalisiereMonate(ByVal eingabe As String) As String
    Dim teile() As String
    Dim i As Long
    Dim monat As Long
    Dim result As String
    
    ' Alle Leerzeichen entfernen, dann am Komma splitten
    eingabe = Replace(eingabe, " ", "")
    teile = Split(eingabe, ",")
    
    result = ""
    For i = LBound(teile) To UBound(teile)
        If IsNumeric(Trim(teile(i))) Then
            monat = CLng(Trim(teile(i)))
            If monat >= 1 And monat <= 12 Then
                If result <> "" Then result = result & ", "
                result = result & Format(monat, "00")
            End If
        End If
    Next i
    
    If result = "" Then result = eingabe  ' Fallback: Original zurueckgeben
    NormalisiereMonate = result
End Function


' ===============================================================
' KONFIG-AENDERUNG: Reaktion auf Eingaben im Config-Bereich
' ===============================================================
Private Sub VerarbeiteKonfigAenderung(ByVal Target As Range)
    On Error GoTo CfgExit
    Application.EnableEvents = False
    
    On Error Resume Next
    Me.Unprotect PASSWORD:=PASSWORD
    On Error GoTo CfgExit
    
    Select Case Target.Row
        Case ES_CFG_MITGLIEDSBEITRAG_ROW
            ' Mitgliedsbeitrag -> Zahlungstermine sync
            Call mod_Einstellungen.SyncMitgliedsbeitragZuTabelle(Me)
            
        Case ES_CFG_MIETE_ROW, ES_CFG_GRUNDSTEUER_ROW
            ' Miete/Grundsteuer geaendert -> Pacht aktualisiert sich per Formel
            ' -> Pacht in Zahlungstermine sync
            Application.Calculate  ' Formeln aktualisieren
            Call mod_Einstellungen.SyncPachtZuTabelle(Me)
            
        Case ES_CFG_PARZELLEN_ROW
            ' Parzellen-Anzahl geaendert -> Pacht aktualisiert sich per Formel
            Application.Calculate
            Call mod_Einstellungen.SyncPachtZuTabelle(Me)
    End Select
    
    Me.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' v6.1: Startseite bei JEDER Konfig-Aenderung aktualisieren
    ' (Abrechnungsjahr, Kontostand, Vereinsname etc.)
    On Error Resume Next
    Call mod_Startseite.InitialisiereStartseite
    On Error GoTo 0
    
CfgExit:
    Application.EnableEvents = True
    If Err.Number <> 0 Then Err.Clear
End Sub

















































































