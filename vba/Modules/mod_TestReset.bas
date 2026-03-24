Attribute VB_Name = "mod_TestReset"
Option Explicit

' ***************************************************************
' MODUL: mod_TestReset
' VERSION: 1.0 - 15.03.2026
' ZWECK: Setzt die Arbeitsmappe in den Zustand VOR dem
'        CSV-Import zurueck. Loescht:
'        1. Bankkonto-Daten (ab Zeile 28)
'        2. Uebersicht-Daten (ab Zeile 4)
'        3. Import-Protokoll (Daten Y500)
'        4. Vorjahr-Speicher (Daten CA-CF)
'
'        Aufruf: Alt+F8 > TestReset_VorCSVImport
'        Oder im Direktfenster: mod_TestReset.TestReset_VorCSVImport
' ***************************************************************

Public Sub TestReset_VorCSVImport()
    
    Dim wsBank As Worksheet
    Dim wsUeb As Worksheet
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim eventsWaren As Boolean
    Dim antwort As VbMsgBoxResult
    
    ' --- Sicherheitsabfrage ---
    antwort = MsgBox("Alle importierten Kontoausz" & ChrW(252) & "ge, die " & ChrW(220) & "bersicht " & _
                     "und das Import-Protokoll werden gel" & ChrW(246) & "scht." & vbCrLf & vbCrLf & _
                     "Die Einstellungen, Mitgliederliste, Kategorie- und " & _
                     "EntityKey-Tabellen bleiben erhalten." & vbCrLf & vbCrLf & _
                     "Fortfahren?", vbYesNo + vbQuestion, "Test-Reset vor CSV-Import")
    
    If antwort <> vbYes Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    eventsWaren = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' =============================================================
    ' 1. BANKKONTO leeren (ab Zeile 28, Spalten A-Z)
    ' =============================================================
    Set wsBank = ThisWorkbook.Worksheets(WS_BANKKONTO)
    wsBank.Unprotect PASSWORD:=PASSWORD
    
    ' AutoFilter entfernen falls aktiv
    If wsBank.AutoFilterMode Then wsBank.AutoFilterMode = False
    
    lastRow = wsBank.Cells(wsBank.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    If lastRow >= BK_START_ROW Then
        wsBank.Rows(BK_START_ROW & ":" & lastRow).Clear
    End If
    
    ' Formeln wiederherstellen (Spalte G, Zusammenfassungen)
    Call mod_Banking_Format.StelleFormelnWiederHer(wsBank)
    
    wsBank.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Debug.Print "[TestReset] Bankkonto: " & _
        IIf(lastRow >= BK_START_ROW, (lastRow - BK_START_ROW + 1) & " Zeilen", "keine Daten") & _
        " gel" & ChrW(246) & "scht."
    
    ' =============================================================
    ' 2. UEBERSICHT leeren (ab Zeile 4, Spalten A-H)
    ' =============================================================
    Set wsUeb = ThisWorkbook.Worksheets(WS_UEBERSICHT())
    wsUeb.Unprotect PASSWORD:=PASSWORD
    
    ' AutoFilter entfernen falls aktiv
    If wsUeb.AutoFilterMode Then wsUeb.AutoFilterMode = False
    
    lastRow = wsUeb.Cells(wsUeb.Rows.count, 1).End(xlUp).Row
    
    If lastRow >= 4 Then
        wsUeb.Rows("4:" & lastRow).Clear
        ' Auch Spalte I (Summe Ist) leeren, falls Zeilen weiter reichen
        wsUeb.Range(wsUeb.Cells(4, 9), wsUeb.Cells(lastRow, 9)).Clear
    End If
    
    wsUeb.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Debug.Print "[TestReset] " & ChrW(220) & "bersicht: " & _
        IIf(lastRow >= 4, (lastRow - 3) & " Zeilen", "keine Daten") & _
        " gel" & ChrW(246) & "scht."
    
    ' =============================================================
    ' 3. IMPORT-PROTOKOLL leeren (Daten Y500)
    ' =============================================================
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    wsDaten.Unprotect PASSWORD:=PASSWORD
    
    wsDaten.Range(CELL_IMPORT_PROTOKOLL).ClearContents
    
    Debug.Print "[TestReset] Import-Protokoll (Y500) gel" & ChrW(246) & "scht."
    
    ' =============================================================
    ' 4. VORJAHR-SPEICHER leeren (Daten CA-CF)
    ' =============================================================
    lastRow = wsDaten.Cells(wsDaten.Rows.count, VJ_COL_DATUM).End(xlUp).Row
    
    If lastRow >= VJ_START_ROW Then
        wsDaten.Range(wsDaten.Cells(VJ_START_ROW, VJ_COL_DATUM), _
                      wsDaten.Cells(lastRow, VJ_COL_ENTITYKEY)).Clear
        Debug.Print "[TestReset] Vorjahr-Speicher: " & _
            (lastRow - VJ_START_ROW + 1) & " Zeilen gel" & ChrW(246) & "scht."
    Else
        Debug.Print "[TestReset] Vorjahr-Speicher: keine Daten."
    End If
    
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    ' =============================================================
    ' 5. Import-Report ListBox aktualisieren (falls sichtbar)
    ' =============================================================
    On Error Resume Next
    Call mod_Banking_Report.Initialize_ImportReport_ListBox
    On Error GoTo ErrorHandler
    
    ' =============================================================
    ' FERTIG
    ' =============================================================
    Application.ScreenUpdating = True
    Application.EnableEvents = eventsWaren
    
    MsgBox "Test-Reset abgeschlossen." & vbCrLf & vbCrLf & _
           "Gel" & ChrW(246) & "scht:" & vbCrLf & _
           "  " & ChrW(8226) & " Bankkonto (alle Kontoausz" & ChrW(252) & "ge)" & vbCrLf & _
           "  " & ChrW(8226) & " " & ChrW(220) & "bersicht (alle Eintr" & ChrW(228) & "ge)" & vbCrLf & _
           "  " & ChrW(8226) & " Import-Protokoll (Y500)" & vbCrLf & _
           "  " & ChrW(8226) & " Vorjahr-Speicher (CA-CF)" & vbCrLf & vbCrLf & _
           "Du kannst jetzt den CSV-Import erneut starten.", _
           vbInformation, "Test-Reset"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = eventsWaren
    MsgBox "Fehler beim Test-Reset:" & vbCrLf & _
           "Nr. " & Err.Number & ": " & Err.Description, vbCritical
End Sub






























