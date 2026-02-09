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
' TABELLENBLATT-MODUL: Einstellungen (v4.4)
' Formatiert die Zahlungstermin-Tabelle bei Änderungen.
' NEU v4.4: "Ultimo" entfernt – stattdessen 31 (= "31. Tag").
'           Duplikat-Prüfung für Spalte B.
'           D/E vs F Sperre ohne Ultimo-Sonderbehandlung.
' ===============================================================

Private Sub Worksheet_Activate()
    Call mod_Einstellungen.FormatiereZahlungsterminTabelle(Me)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    
    ' Nur reagieren auf Änderungen in Spalten B-I, ab Zeile 4
    If Intersect(Target, Me.Range(Me.Cells(ES_START_ROW, ES_COL_START), _
                                   Me.Cells(Me.Rows.count, ES_COL_END))) Is Nothing Then
        Exit Sub
    End If
    
    If Target.Row < ES_START_ROW Then Exit Sub
    If Target.Cells.count > 1 Then GoTo NurFormatieren
    
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
            
            ' D/E vs F Sperre: Wenn E gefüllt, darf F NICHT gefüllt sein
            If Trim(CStr(Me.Cells(rowES, ES_COL_STICHTAG_FIX).value)) <> "" Then
                MsgBox "Spalte F (Soll-Stichtag Fix) ist bereits gefüllt!" & vbLf & _
                       "Entweder Spalte D/E ODER Spalte F verwenden." & vbLf & vbLf & _
                       "Die Eingabe in Spalte E wird zurückgesetzt.", vbExclamation, "Konflikt"
                Target.ClearContents
                GoTo WeiterFormatieren
            End If
            
            ' Nur Spalte E gefüllt (kein Tag in D) -> automatisch 31 setzen
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
    ' SPALTE D: Soll-Tag - D/E vs F Sperre prüfen
    ' ---------------------------------------------------------------
    If Target.Column = ES_COL_SOLL_TAG Then
        Dim tagWert As String
        tagWert = Trim(CStr(Target.value))
        
        If tagWert <> "" Then
            ' D/E vs F Sperre: Wenn D gefüllt, darf F NICHT gefüllt sein
            If Trim(CStr(Me.Cells(rowES, ES_COL_STICHTAG_FIX).value)) <> "" Then
                MsgBox "Spalte F (Soll-Stichtag Fix) ist bereits gefüllt!" & vbLf & _
                       "Entweder Spalte D/E ODER Spalte F verwenden." & vbLf & vbLf & _
                       "Die Eingabe in Spalte D wird zurückgesetzt.", vbExclamation, "Konflikt"
                Target.ClearContents
                GoTo WeiterFormatieren
            End If
        End If
    End If
    
    ' ---------------------------------------------------------------
    ' SPALTE F: Soll-Stichtag (Fix) - D/E vs F Sperre prüfen
    ' ---------------------------------------------------------------
    If Target.Column = ES_COL_STICHTAG_FIX Then
        Dim stichtagWert As String
        stichtagWert = Trim(CStr(Target.value))
        
        If stichtagWert <> "" Then
            ' F vs D/E Sperre: Wenn F gefüllt, dürfen D und E NICHT gefüllt sein
            Dim dWert As String
            Dim eWert As String
            dWert = Trim(CStr(Me.Cells(rowES, ES_COL_SOLL_TAG).value))
            eWert = Trim(CStr(Me.Cells(rowES, ES_COL_SOLL_MONATE).value))
            
            If dWert <> "" Or eWert <> "" Then
                MsgBox "Spalte D (Soll-Tag) oder E (Soll-Monate) ist bereits gefüllt!" & vbLf & _
                       "Entweder Spalte D/E ODER Spalte F verwenden." & vbLf & vbLf & _
                       "Die Eingabe in Spalte F wird zurückgesetzt.", vbExclamation, "Konflikt"
                Target.ClearContents
                GoTo WeiterFormatieren
            End If
        End If
    End If
    
    ' ---------------------------------------------------------------
    ' SPALTE B: Kategorie - Duplikat prüfen, sortieren, Cursor setzen
    ' ---------------------------------------------------------------
    If Target.Column = ES_COL_KATEGORIE Then
        Dim katWert As String
        katWert = Trim(CStr(Target.value))
        
        If katWert <> "" Then
            ' Duplikat-Prüfung: Kommt diese Kategorie bereits in einer anderen Zeile vor?
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
                    ' Cursor in die nächste Spalte (C = Soll-Betrag) setzen
                    Me.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
                    Application.EnableEvents = True
                    Me.Cells(suchZeile, ES_COL_SOLL_BETRAG).Select
                    
                    ' Formatierung trotzdem anwenden (ohne Events)
                    Application.EnableEvents = False
                    Call mod_Einstellungen.FormatiereZahlungsterminTabelle(Me)
                    Application.EnableEvents = True
                    
                    ' Re-Evaluierung auslösen
                    Call mod_KategorieEngine_Pipeline.ReEvaluiereAlleNichtManuellen
                    Exit Sub
                End If
            Next suchZeile
        End If
    End If
    
WeiterFormatieren:
    Me.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True

NurFormatieren:
    ' Formatierung aktualisieren
    Call mod_Einstellungen.FormatiereZahlungsterminTabelle(Me)
    
    ' Bankkonto-Kategorien neu evaluieren
    Call mod_KategorieEngine_Pipeline.ReEvaluiereAlleNichtManuellen
    
SafeExit:
    Application.EnableEvents = True
    If Err.Number <> 0 Then Err.Clear
    
End Sub


' ===============================================================
' Hilfsfunktion: Monatseingabe normalisieren
' "3,6,9" -> "03, 06, 09"
' "1, 3, 12" -> "01, 03, 12"
' Jeder Monat wird auf 2 Stellen aufgefüllt,
' nach dem Komma folgt ein Leerzeichen.
' Ungültige Monate (>12 oder <1) werden ignoriert.
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
    
    If result = "" Then result = eingabe  ' Fallback: Original zurückgeben
    NormalisiereMonate = result
End Function

