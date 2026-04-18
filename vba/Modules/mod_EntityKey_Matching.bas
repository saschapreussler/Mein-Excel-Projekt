Attribute VB_Name = "mod_EntityKey_Matching"
Option Explicit

' ***************************************************************
' MODUL: mod_EntityKey_Matching
' ZWECK: Mitglieder-Suche und Namensabgleich fuer EntityKey-System
' ABGELEITET AUS: mod_EntityKey_Manager (Modularisierung)
' VERSION: 1.0 - 01.03.2026
' FUNKTIONEN:
'   - SucheMitgliederZuKontoname: Hauptsuche in Mitgliederliste
'   - PruefeNamensMatch: Namens-Matching (Vor-/Nachname)
'   - IstMitgliedBereitsGefunden: Duplikatpruefung in Collection
'   - FindeBestenTreffer: Besten Match aus Collection ermitteln
'   - PruefeObInHistorie: Prueft ob Name in Mitgliederhistorie
'   - HoleParzelleFuerEhemaligesAusHistorie: Parzelle aus Historie
'   - HoleAlleParzellen: Alle Parzellen einer MemberID
' ***************************************************************

' ===============================================================
' Sucht Mitglieder im Kontonamen
' ===============================================================
Public Function SucheMitgliederZuKontoname(ByVal kontoname As String, _
                                              ByRef wsM As Worksheet, _
                                              ByRef wsH As Worksheet) As Collection
    
    Dim gefunden As New Collection
    Dim r As Long
    Dim lastRow As Long
    Dim nachname As String
    Dim vorname As String
    Dim memberID As String
    Dim parzelle As String
    Dim funktion As String
    Dim kontoNameNorm As String
    Dim mitgliedInfo(0 To 8) As Variant
    Dim matchResult As Long
    
    Set SucheMitgliederZuKontoname = gefunden
    
    If kontoname = "" Then Exit Function
    
    kontoNameNorm = mod_EntityKey_Normalize.NormalisiereStringFuerVergleich(kontoname)
    If kontoNameNorm = "" Then Exit Function
    
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_NACHNAME).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        If Trim(wsM.Cells(r, M_COL_PACHTENDE).value) = "" Then
            nachname = Trim(wsM.Cells(r, M_COL_NACHNAME).value)
            vorname = Trim(wsM.Cells(r, M_COL_VORNAME).value)
            memberID = Trim(wsM.Cells(r, M_COL_MEMBER_ID).value)
            parzelle = Trim(wsM.Cells(r, M_COL_PARZELLE).value)
            funktion = Trim(wsM.Cells(r, M_COL_FUNKTION).value)
            
            matchResult = PruefeNamensMatch(nachname, vorname, kontoNameNorm)
            
            If matchResult > 0 Then
                If Not IstMitgliedBereitsGefunden(gefunden, memberID, False) Then
                    mitgliedInfo(0) = memberID
                    mitgliedInfo(1) = nachname
                    mitgliedInfo(2) = vorname
                    mitgliedInfo(3) = parzelle
                    mitgliedInfo(4) = funktion
                    mitgliedInfo(5) = r
                    mitgliedInfo(6) = False
                    mitgliedInfo(7) = CDate("01.01.1900")
                    mitgliedInfo(8) = matchResult
                    gefunden.Add mitgliedInfo
                End If
            End If
        End If
    Next r
    
    Set SucheMitgliederZuKontoname = gefunden
End Function

' ===============================================================
' Prueft Namens-Match (0=kein, 1=nur Nachname, 2=Vor+Nachname)
' ===============================================================
Public Function PruefeNamensMatch(ByVal nachname As String, ByVal vorname As String, _
                                     ByVal kontoNameNorm As String) As Long
    Dim nachnameNorm As String
    Dim vornameNorm As String
    
    PruefeNamensMatch = 0
    
    nachnameNorm = mod_EntityKey_Normalize.NormalisiereStringFuerVergleich(nachname)
    vornameNorm = mod_EntityKey_Normalize.NormalisiereStringFuerVergleich(vorname)
    
    If nachnameNorm = "" Or Len(nachnameNorm) < 3 Then Exit Function
    If InStr(kontoNameNorm, nachnameNorm) = 0 Then Exit Function
    
    If vornameNorm <> "" And Len(vornameNorm) >= 2 Then
        If InStr(kontoNameNorm, vornameNorm) > 0 Then
            PruefeNamensMatch = 2
        Else
            PruefeNamensMatch = 1
        End If
    Else
        PruefeNamensMatch = 1
    End If
End Function

' ===============================================================
' Prueft ob MemberID bereits in Collection gefunden
' ===============================================================
Public Function IstMitgliedBereitsGefunden(ByRef col As Collection, _
                                              ByVal memberID As String, _
                                              ByVal istEhemalig As Boolean) As Boolean
    Dim item As Variant
    IstMitgliedBereitsGefunden = False
    
    For Each item In col
        If item(0) = memberID And item(6) = istEhemalig Then
            IstMitgliedBereitsGefunden = True
            Exit Function
        End If
    Next item
End Function

' ===============================================================
' Findet den besten Treffer aus Collection
' ===============================================================
Public Function FindeBestenTreffer(ByRef mitglieder As Collection) As Variant
    Dim bestInfo As Variant
    Dim info As Variant
    Dim bestScore As Long
    Dim i As Long
    
    bestScore = 0
    
    For i = 1 To mitglieder.count
        info = mitglieder(i)
        If CLng(info(8)) > bestScore Then
            bestScore = CLng(info(8))
            bestInfo = info
        End If
    Next i
    
    If bestScore = 0 And mitglieder.count > 0 Then
        bestInfo = mitglieder(1)
    End If
    
    FindeBestenTreffer = bestInfo
End Function

' ===============================================================
' Prueft ob ehemaliges Mitglied in Mitgliederhistorie steht
' ===============================================================
Public Function PruefeObInHistorie(ByVal kontoname As String, ByRef wsH As Worksheet) As Boolean
    Dim r As Long
    Dim lastRow As Long
    Dim nachnameHist As String
    Dim kontoNameNorm As String
    Dim nachnameNorm As String
    
    PruefeObInHistorie = False
    
    If kontoname = "" Then Exit Function
    
    kontoNameNorm = mod_EntityKey_Normalize.NormalisiereStringFuerVergleich(kontoname)
    If kontoNameNorm = "" Then Exit Function
    
    lastRow = wsH.Cells(wsH.Rows.count, H_COL_NAME_EHEM_PAECHTER).End(xlUp).Row
    
    For r = H_START_ROW To lastRow
        nachnameHist = Trim(wsH.Cells(r, H_COL_NAME_EHEM_PAECHTER).value)
        If nachnameHist <> "" Then
            nachnameNorm = mod_EntityKey_Normalize.NormalisiereStringFuerVergleich(nachnameHist)
            If nachnameNorm <> "" And Len(nachnameNorm) >= 3 Then
                If InStr(kontoNameNorm, nachnameNorm) > 0 Then
                    PruefeObInHistorie = True
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

' ===============================================================
' Holt Parzelle fuer ehemaliges Mitglied aus Historie
' ===============================================================
Public Function HoleParzelleFuerEhemaligesAusHistorie(ByVal kontoname As String, ByRef wsH As Worksheet) As String
    Dim r As Long
    Dim lastRow As Long
    Dim nachnameHist As String
    Dim kontoNameNorm As String
    Dim nachnameNorm As String
    
    HoleParzelleFuerEhemaligesAusHistorie = ""
    
    If kontoname = "" Then Exit Function
    
    kontoNameNorm = mod_EntityKey_Normalize.NormalisiereStringFuerVergleich(kontoname)
    If kontoNameNorm = "" Then Exit Function
    
    lastRow = wsH.Cells(wsH.Rows.count, H_COL_NAME_EHEM_PAECHTER).End(xlUp).Row
    
    For r = H_START_ROW To lastRow
        nachnameHist = Trim(wsH.Cells(r, H_COL_NAME_EHEM_PAECHTER).value)
        If nachnameHist <> "" Then
            nachnameNorm = mod_EntityKey_Normalize.NormalisiereStringFuerVergleich(nachnameHist)
            If nachnameNorm <> "" And Len(nachnameNorm) >= 3 Then
                If InStr(kontoNameNorm, nachnameNorm) > 0 Then
                    HoleParzelleFuerEhemaligesAusHistorie = Trim(CStr(wsH.Cells(r, H_COL_PARZELLE).value))
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

' ===============================================================
' Holt alle Parzellen fuer eine MemberID (kommagetrennt)
' ===============================================================
Public Function HoleAlleParzellen(ByVal memberID As String, _
                                    ByRef wsM As Worksheet) As String
    Dim r As Long
    Dim lastRow As Long
    Dim parzelle As String
    Dim result As String
    Dim dictParzellen As Object
    
    Set dictParzellen = CreateObject("Scripting.Dictionary")
    result = ""
    
    If memberID = "" Then
        HoleAlleParzellen = ""
        Exit Function
    End If
    
    lastRow = wsM.Cells(wsM.Rows.count, M_COL_MEMBER_ID).End(xlUp).Row
    
    For r = M_START_ROW To lastRow
        If Trim(wsM.Cells(r, M_COL_MEMBER_ID).value) = memberID Then
            parzelle = Trim(wsM.Cells(r, M_COL_PARZELLE).value)
            If parzelle <> "" Then
                If Not dictParzellen.Exists(parzelle) Then
                    dictParzellen.Add parzelle, True
                    If result <> "" Then result = result & ", "
                    result = result & parzelle
                End If
            End If
        End If
    Next r
    
    HoleAlleParzellen = result
End Function
















































































