Attribute VB_Name = "mod_Uebersicht_Event"
Option Explicit

' ***************************************************************
' MODUL: mod_Uebersicht_Event
' VERSION: 1.0 - 15.03.2026
' ZWECK: Verarbeitet manuelle Aenderungen auf dem Uebersicht-Blatt
'        - Gelb -> Gruen wenn Nutzer Soll-Betrag eintraegt
'        - MsgBox: Soll-Betrag fuer Folgemonat uebernehmen?
'        - Automatische Uebernahme in Folgemonate gleicher Parzelle+Kategorie
'        - Wird von DieseArbeitsmappe.Workbook_SheetChange aufgerufen
' ***************************************************************

' Konstanten (identisch mit mod_Uebersicht_Generator)
Private Const UEBERSICHT_START_ROW As Long = 4
Private Const UEBERSICHT_HEADER_ROW As Long = 3
Private Const UEB_COL_PARZELLE As Long = 1
Private Const UEB_COL_MITGLIED As Long = 2
Private Const UEB_COL_MONAT As Long = 3
Private Const UEB_COL_KATEGORIE As Long = 4
Private Const UEB_COL_SOLL As Long = 5
Private Const UEB_COL_IST As Long = 6
Private Const UEB_COL_STATUS As Long = 7
Private Const UEB_COL_BEMERKUNG As Long = 8

Private Const FARBE_HELLGELB_MANUELL As Long = 10092543  ' RGB(255, 255, 153)
Private Const AMPEL_GRUEN As Long = 12968900             ' RGB(196, 225, 196)


' ===============================================================
' Wird von Workbook_SheetChange aufgerufen wenn eine Zelle
' auf dem Uebersicht-Blatt geaendert wurde.
' Prueft ob eine gelbe Soll-Zelle (Spalte E) manuell befuellt wurde.
' ===============================================================
Public Sub VerarbeiteUebersichtAenderung(ByVal Target As Range)
    
    On Error GoTo ErrorHandler
    
    ' Nur einzelne Zelle
    If Target.Cells.CountLarge <> 1 Then Exit Sub
    
    ' Nur Spalte E (Soll)
    If Target.Column <> UEB_COL_SOLL Then Exit Sub
    
    ' Nur Datenzeilen (ab Zeile 4)
    If Target.Row < UEBERSICHT_START_ROW Then Exit Sub
    
    ' Nur wenn Zelle aktuell hell-gelb ist (= variabel, editierbar)
    If Target.Interior.color <> FARBE_HELLGELB_MANUELL Then Exit Sub
    
    ' Neuer Wert pruefen
    Dim neuerWert As Double
    neuerWert = 0
    
    If IsNumeric(Target.value) Then
        neuerWert = CDbl(Target.value)
    End If
    
    ' Wenn Wert geloescht oder 0 -> nichts tun (bleibt gelb)
    If neuerWert <= 0 Then Exit Sub
    
    Dim wsUeb As Worksheet
    Set wsUeb = Target.Worksheet
    
    Dim zeile As Long
    zeile = Target.Row
    
    ' Parzelle und Kategorie der geaenderten Zeile ermitteln
    Dim parzelle As String
    parzelle = CStr(wsUeb.Cells(zeile, UEB_COL_PARZELLE).value)
    Dim kategorie As String
    kategorie = CStr(wsUeb.Cells(zeile, UEB_COL_KATEGORIE).value)
    
    ' Events deaktivieren (verhindert Endlosschleife)
    Application.EnableEvents = False
    
    ' Blattschutz temporaer entfernen
    On Error Resume Next
    wsUeb.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    ' 1. Aktuelle Zelle: Gelb -> Gruen + Bemerkung anpassen
    Target.Interior.color = AMPEL_GRUEN
    
    Dim bemerkung As String
    bemerkung = CStr(wsUeb.Cells(zeile, UEB_COL_BEMERKUNG).value)
    
    ' Alte variable Hinweise entfernen
    bemerkung = EntferneTeilBemerkung(bemerkung, "Soll-Betrag variabel")
    bemerkung = EntferneTeilBemerkung(bemerkung, "Soll aus Vormonat")
    
    ' Neuen Hinweis hinzufuegen
    Dim manuellHinweis As String
    manuellHinweis = "Soll manuell gesetzt (" & Format(neuerWert, "#,##0.00") & _
                     " " & ChrW(8364) & ")"
    If bemerkung = "" Then
        bemerkung = manuellHinweis
    Else
        bemerkung = bemerkung & " | " & manuellHinweis
    End If
    wsUeb.Cells(zeile, UEB_COL_BEMERKUNG).value = bemerkung
    
    ' 2. Pruefen ob IST den neuen Soll erreicht -> Status aktualisieren
    Dim istWert As Double
    istWert = val(CStr(wsUeb.Cells(zeile, UEB_COL_IST).value))
    
    If istWert > 0 And Abs(istWert - neuerWert) < 0.01 Then
        wsUeb.Cells(zeile, UEB_COL_STATUS).value = "GR" & ChrW(220) & "N"
        wsUeb.Cells(zeile, UEB_COL_STATUS).Interior.color = AMPEL_GRUEN
    End If
    
    ' 3. Folgemonat-Uebernahme: MsgBox fragen
    Dim lastRow As Long
    lastRow = wsUeb.Cells(wsUeb.Rows.count, UEB_COL_PARZELLE).End(xlUp).Row
    
    ' Pruefen ob es ueberhaupt Folgezeilen fuer diese Parzelle+Kategorie gibt
    Dim hatFolgezeilen As Boolean
    hatFolgezeilen = False
    Dim rCheck As Long
    For rCheck = zeile + 1 To lastRow
        If CStr(wsUeb.Cells(rCheck, UEB_COL_PARZELLE).value) = parzelle Then
            If StrComp(CStr(wsUeb.Cells(rCheck, UEB_COL_KATEGORIE).value), kategorie, vbTextCompare) = 0 Then
                hatFolgezeilen = True
                Exit For
            End If
        End If
    Next rCheck
    
    If hatFolgezeilen Then
        ' MsgBox: Soll-Betrag fuer Folgezahlungen uebernehmen?
        Dim antwort As VbMsgBoxResult
        antwort = MsgBox( _
            "Der Soll-Betrag f" & ChrW(252) & "r '" & kategorie & "' (Parzelle " & parzelle & _
            ") wurde auf " & Format(neuerWert, "#,##0.00") & " " & ChrW(8364) & " gesetzt." & vbLf & vbLf & _
            "Soll dieser Betrag auch f" & ChrW(252) & "r die Folgemonat-Abschlags" & _
            "zahlungen " & ChrW(252) & "bernommen werden?", _
            vbYesNo + vbQuestion, _
            "Soll-Betrag " & ChrW(252) & "bernehmen?")
        
        If antwort = vbYes Then
            ' Alle Folgezeilen mit gleicher Parzelle+Kategorie aktualisieren
            Call UebernehmeSollInFolgemonate(wsUeb, zeile, parzelle, kategorie, neuerWert, lastRow)
        End If
    End If
    
    ' Blattschutz wieder aktivieren
    On Error Resume Next
    wsUeb.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0
    
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    On Error Resume Next
    wsUeb.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0
    Debug.Print "[" & ChrW(220) & "bersicht Event] FEHLER: " & Err.Description
    
End Sub


' ===============================================================
' Uebertraegt den Soll-Betrag in alle Folgezeilen mit gleicher
' Parzelle+Kategorie. Setzt Farbe auf Gruen + Bemerkung.
' ===============================================================
Private Sub UebernehmeSollInFolgemonate(ByVal wsUeb As Worksheet, _
                                         ByVal startZeile As Long, _
                                         ByVal parzelle As String, _
                                         ByVal kategorie As String, _
                                         ByVal sollWert As Double, _
                                         ByVal lastRow As Long)
    
    Dim r As Long
    For r = startZeile + 1 To lastRow
        ' Parzelle + Kategorie muessen uebereinstimmen
        If CStr(wsUeb.Cells(r, UEB_COL_PARZELLE).value) = parzelle Then
            If StrComp(CStr(wsUeb.Cells(r, UEB_COL_KATEGORIE).value), kategorie, vbTextCompare) = 0 Then
                ' Nur wenn Zelle noch gelb ist (= noch nicht manuell gesetzt)
                If wsUeb.Cells(r, UEB_COL_SOLL).Interior.color = FARBE_HELLGELB_MANUELL Then
                    ' Soll-Wert setzen
                    wsUeb.Cells(r, UEB_COL_SOLL).value = sollWert
                    
                    ' Gelb -> Gruen
                    wsUeb.Cells(r, UEB_COL_SOLL).Interior.color = AMPEL_GRUEN
                    
                    ' Bemerkung aktualisieren
                    Dim bem As String
                    bem = CStr(wsUeb.Cells(r, UEB_COL_BEMERKUNG).value)
                    bem = EntferneTeilBemerkung(bem, "Soll-Betrag variabel")
                    bem = EntferneTeilBemerkung(bem, "Soll aus Vormonat")
                    
                    Dim hinweis As String
                    hinweis = "Soll " & ChrW(252) & "bernommen (" & _
                              Format(sollWert, "#,##0.00") & " " & ChrW(8364) & ")"
                    If bem = "" Then
                        bem = hinweis
                    Else
                        bem = bem & " | " & hinweis
                    End If
                    wsUeb.Cells(r, UEB_COL_BEMERKUNG).value = bem
                    
                    ' Status aktualisieren wenn IST passt
                    Dim istW As Double
                    istW = val(CStr(wsUeb.Cells(r, UEB_COL_IST).value))
                    If istW > 0 And Abs(istW - sollWert) < 0.01 Then
                        wsUeb.Cells(r, UEB_COL_STATUS).value = "GR" & ChrW(220) & "N"
                        wsUeb.Cells(r, UEB_COL_STATUS).Interior.color = AMPEL_GRUEN
                    End If
                End If
            End If
        End If
    Next r
    
End Sub


' ===============================================================
' Hilfsfunktion: Entfernt einen Teil-String aus einer
' Pipe-getrennten Bemerkung (z.B. "Soll-Betrag variabel...")
' ===============================================================
Private Function EntferneTeilBemerkung(ByVal bemerkung As String, _
                                       ByVal suchText As String) As String
    
    If bemerkung = "" Then
        EntferneTeilBemerkung = ""
        Exit Function
    End If
    
    Dim teile() As String
    teile = Split(bemerkung, " | ")
    
    Dim ergebnis As String
    ergebnis = ""
    
    Dim i As Long
    For i = 0 To UBound(teile)
        If InStr(1, teile(i), suchText, vbTextCompare) = 0 Then
            If ergebnis = "" Then
                ergebnis = Trim(teile(i))
            Else
                ergebnis = ergebnis & " | " & Trim(teile(i))
            End If
        End If
    Next i
    
    EntferneTeilBemerkung = ergebnis
    
End Function


























































