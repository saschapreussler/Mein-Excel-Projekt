Attribute VB_Name = "mod_EntityKey_Ampel"
Option Explicit

' ***************************************************************
' MODUL: mod_EntityKey_Ampel
' ZWECK: Ampelfarben-System (Gruen/Gelb/Rot) fuer EntityKey-Tabelle
' ABGELEITET AUS: mod_EntityKey_Manager (Modularisierung)
' VERSION: 1.0 - 01.03.2026
' FUNKTIONEN:
'   - SetzeAmpelFarbe: Farbe fuer Spalten U-X einer Zeile setzen
'   - SetzeAlleAmpelfarbenNachSortierung: Alle Zeilen nach Sort
'   - BerechneAmpelStatus: Status-Logik (1=Gruen, 2=Gelb, 3=Rot)
' ***************************************************************

' ===============================================================
' Setzt Farbe fuer Spalten U-X einer Zeile
' ampelStatus: 1=Gruen, 2=Gelb, 3=Rot
' ===============================================================
Public Sub SetzeAmpelFarbe(ByRef ws As Worksheet, ByVal zeile As Long, ByVal ampelStatus As Long)
    Dim rngAmpel As Range
    Dim farbe As Long
    
    Set rngAmpel = ws.Range(ws.Cells(zeile, EK_COL_ZUORDNUNG), _
                            ws.Cells(zeile, EK_COL_DEBUG))
    
    Select Case ampelStatus
        Case 1
            farbe = RGB(198, 224, 180)  ' Gruen
        Case 2
            farbe = RGB(255, 230, 153)  ' Gelb
        Case 3
            farbe = RGB(255, 150, 150)  ' Rot
        Case Else
            farbe = RGB(198, 224, 180)  ' Default Gruen
    End Select
    
    rngAmpel.Interior.color = farbe
End Sub

' ===============================================================
' Setzt Ampelfarben fuer ALLE Zeilen NACH Sortierung
' EHEMALIGES MITGLIED in Historie -> GRUEN
' ===============================================================
Public Sub SetzeAlleAmpelfarbenNachSortierung(ByRef wsD As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    Dim entityKey As String
    Dim zuordnung As String
    Dim role As String
    Dim debugTxt As String
    Dim ampel As Long
    Dim kontoname As String
    Dim wsH As Worksheet
    
    On Error Resume Next
    Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
    On Error GoTo 0
    
    lastRow = wsD.Cells(wsD.Rows.count, EK_COL_IBAN).End(xlUp).Row
    Dim lastRowR As Long
    lastRowR = wsD.Cells(wsD.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    If lastRowR > lastRow Then lastRow = lastRowR
    If lastRow < EK_START_ROW Then Exit Sub
    
    For r = EK_START_ROW To lastRow
        entityKey = Trim(wsD.Cells(r, EK_COL_ENTITYKEY).value)
        zuordnung = Trim(wsD.Cells(r, EK_COL_ZUORDNUNG).value)
        role = Trim(wsD.Cells(r, EK_COL_ROLE).value)
        debugTxt = Trim(wsD.Cells(r, EK_COL_DEBUG).value)
        
        ampel = BerechneAmpelStatus(entityKey, zuordnung, role, debugTxt)
        
        ' Bei EHEMALIGES MITGLIED pruefen ob in Historie
        If UCase(role) = "EHEMALIGES MITGLIED" Then
            If Not wsH Is Nothing Then
                kontoname = Trim(CStr(wsD.Cells(r, EK_COL_KONTONAME).value))
                If mod_EntityKey_Matching.PruefeObInHistorie(kontoname, wsH) Then
                    ampel = 1
                Else
                    ampel = 2
                    If InStr(debugTxt, "nicht in Historie") = 0 Then
                        If debugTxt <> "" Then
                            debugTxt = debugTxt & " | nicht in Historie"
                        Else
                            debugTxt = "nicht in Historie"
                        End If
                        wsD.Cells(r, EK_COL_DEBUG).value = debugTxt
                    End If
                End If
            End If
        End If
        
        Call SetzeAmpelFarbe(wsD, r, ampel)
    Next r
End Sub

' ===============================================================
' Berechnet den korrekten Ampelstatus einer Zeile
' GRUEN (1), GELB (2), ROT (3)
' ===============================================================
Public Function BerechneAmpelStatus(ByVal entityKey As String, _
                                      ByVal zuordnung As String, _
                                      ByVal role As String, _
                                      ByVal debugTxt As String) As Long
    Dim debugUpper As String
    debugUpper = UCase(debugTxt)
    
    ' ROT: Kein EntityKey UND kein Role
    If entityKey = "" And role = "" Then
        BerechneAmpelStatus = 3
        Exit Function
    End If
    
    ' ROT: Debug sagt KEIN TREFFER und keine Role gesetzt
    If InStr(debugUpper, "KEIN TREFFER") > 0 And role = "" Then
        BerechneAmpelStatus = 3
        Exit Function
    End If
    
    ' GELB: Nur Nachname gefunden, unsicher
    If InStr(debugUpper, "NUR NACHNAME") > 0 Then
        BerechneAmpelStatus = 2
        Exit Function
    End If
    
    ' GELB: EntityKey fehlt, obwohl Role vorhanden
    If entityKey = "" And role <> "" Then
        BerechneAmpelStatus = 2
        Exit Function
    End If
    
    ' GELB: Role fehlt, obwohl EntityKey vorhanden
    If role = "" And entityKey <> "" Then
        BerechneAmpelStatus = 2
        Exit Function
    End If
    
    ' GELB: Ehemaliges Mitglied (Historie-Check extern)
    If UCase(role) = "EHEMALIGES MITGLIED" Then
        BerechneAmpelStatus = 2
        Exit Function
    End If
    
    ' GRUEN: Alles vorhanden und sicher
    If entityKey <> "" And role <> "" Then
        If zuordnung <> "" Then
            BerechneAmpelStatus = 1
        Else
            BerechneAmpelStatus = 2
        End If
        Exit Function
    End If
    
    ' Default: Gelb
    BerechneAmpelStatus = 2
End Function










































































