Attribute VB_Name = "mod_EntityKey_UI"
Option Explicit

' ***************************************************************
' MODUL: mod_EntityKey_UI
' ZWECK: UI-Interaktion und manuelle Zuordnung fuer EntityKey-System
' ABGELEITET AUS: mod_EntityKey_Manager (Modularisierung)
' VERSION: 1.0 - 01.03.2026
' FUNKTIONEN:
'   - VerarbeiteManuelleRoleAenderung: Event-Handler fuer Spalte W
'   - SetupEntityRoleDropdown: DropDown fuer EntityRole-Spalte
'   - SetupParzelleDropdown: DropDown fuer Parzelle-Spalte
'   - FormatiereEntityKeyZeile: Kompatibilitaets-Stub
' ***************************************************************

' ===============================================================
' Lokale Konstanten (Prefixes und Roles)
' ===============================================================
Private Const PREFIX_SHARE As String = "SHARE-"
Private Const PREFIX_VERSORGER As String = "VERS-"
Private Const PREFIX_BANK As String = "BANK-"
Private Const PREFIX_SHOP As String = "SHOP-"
Private Const PREFIX_EHEMALIG As String = "EX-"
Private Const PREFIX_SONSTIGE As String = "SONST-"

Private Const ROLE_MITGLIED_MIT_PACHT As String = "MITGLIED MIT PACHT"
Private Const ROLE_MITGLIED_OHNE_PACHT As String = "MITGLIED OHNE PACHT"
Private Const ROLE_EHEMALIGES_MITGLIED As String = "EHEMALIGES MITGLIED"
Private Const ROLE_VERSORGER As String = "VERSORGER"
Private Const ROLE_BANK As String = "BANK"
Private Const ROLE_SHOP As String = "SHOP"
Private Const ROLE_SONSTIGE As String = "SONSTIGE"

' ===============================================================
' Verarbeitet manuelle Role-Aenderung in Spalte W
' ===============================================================
Public Sub VerarbeiteManuelleRoleAenderung(ByVal Target As Range)
    Dim wsDaten As Worksheet
    Dim wsM As Worksheet
    Dim wsH As Worksheet
    Dim zeile As Long
    Dim neueRole As String
    Dim kontoname As String
    Dim currentEntityKey As String
    Dim neuerEntityKey As String
    Dim neueZuordnung As String
    Dim neueParzelle As String
    Dim neuerDebug As String
    Dim ampelStatus As Long
    Dim correctPrefix As String
    
    On Error GoTo ErrorHandler
    
    If Target.Column <> EK_COL_ROLE Then Exit Sub
    If Target.Row < EK_START_ROW Then Exit Sub
    
    Set wsDaten = Target.Worksheet
    zeile = Target.Row
    neueRole = UCase(Trim(CStr(Target.value)))
    kontoname = mod_EntityKey_Normalize.EntferneMehrfacheLeerzeichen(Trim(CStr(wsDaten.Cells(zeile, EK_COL_KONTONAME).value)))
    currentEntityKey = Trim(wsDaten.Cells(zeile, EK_COL_ENTITYKEY).value)
    
    Application.EnableEvents = False
    
    On Error Resume Next
    wsDaten.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    Select Case neueRole
        Case "MITGLIED MIT PACHT", "MITGLIED OHNE PACHT", "MITGLIED"
            Set wsM = ThisWorkbook.Worksheets(WS_MITGLIEDER)
            Dim mitglieder As Collection
            Set mitglieder = mod_EntityKey_Matching.SucheMitgliederZuKontoname(kontoname, wsM, _
                              ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE))
            
            If mitglieder.count > 0 Then
                Dim bestMatch As Variant
                bestMatch = mod_EntityKey_Matching.FindeBestenTreffer(mitglieder)
                
                neuerEntityKey = CStr(bestMatch(0))
                neueZuordnung = bestMatch(1) & ", " & bestMatch(2)
                neueParzelle = mod_EntityKey_Matching.HoleAlleParzellen(CStr(bestMatch(0)), wsM)
                neuerDebug = "Manuell: " & neueRole & " -> Mitglied gefunden (" & Format(Now, "dd.mm.yyyy") & ")"
                ampelStatus = 1
            Else
                neuerEntityKey = currentEntityKey
                neueZuordnung = mod_EntityKey_Normalize.ExtrahiereAnzeigeName(kontoname)
                neueParzelle = ""
                neuerDebug = "Manuell: " & neueRole & " -> KEIN Mitglied gefunden (" & Format(Now, "dd.mm.yyyy") & ")"
                ampelStatus = 2
            End If
            
        Case "EHEMALIGES MITGLIED"
            correctPrefix = PREFIX_EHEMALIG
            If Left(UCase(currentEntityKey), Len(correctPrefix)) <> correctPrefix Then
                neuerEntityKey = correctPrefix & mod_EntityKey_Manager.CreateGUID()
            Else
                neuerEntityKey = currentEntityKey
            End If
            neueZuordnung = mod_EntityKey_Normalize.ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            ampelStatus = 2
            
            On Error Resume Next
            Set wsH = ThisWorkbook.Worksheets(WS_MITGLIEDER_HISTORIE)
            On Error GoTo ErrorHandler
            If Not wsH Is Nothing Then
                If mod_EntityKey_Matching.PruefeObInHistorie(kontoname, wsH) Then
                    ampelStatus = 1
                    Dim historieParzelle As String
                    historieParzelle = mod_EntityKey_Matching.HoleParzelleFuerEhemaligesAusHistorie(kontoname, wsH)
                    If historieParzelle <> "" Then
                        neueParzelle = historieParzelle
                    End If
                    neuerDebug = "Manuell: EHEMALIGES MITGLIED - in Historie gefunden; " & Format(Now, "dd.mm.yyyy")
                Else
                    ampelStatus = 2
                    
                    Dim eingabe As String
                    Dim parzelleGueltig As Boolean
                    Dim parzelleNr As Long
                    
                    parzelleGueltig = False
                    Do
                        eingabe = InputBox("Welche Parzelle belegte das ehemalige Mitglied?" & vbCrLf & vbCrLf & _
                                           "Bitte eine Zahl von 1 bis 14 eingeben:" & vbCrLf & _
                                           "(Abbrechen = keine Parzelle zuweisen)", _
                                           "Parzelle f" & ChrW(252) & "r ehemaliges Mitglied", "")
                        
                        If eingabe = "" Then
                            Exit Do
                        End If
                        
                        If IsNumeric(eingabe) Then
                            parzelleNr = CLng(eingabe)
                            If parzelleNr >= 1 And parzelleNr <= 14 Then
                                parzelleGueltig = True
                            Else
                                MsgBox "Ung" & ChrW(252) & "ltige Eingabe! Bitte eine Zahl zwischen 1 und 14 eingeben.", vbExclamation, "Ung" & ChrW(252) & "ltige Parzelle"
                            End If
                        Else
                            MsgBox "Ung" & ChrW(252) & "ltige Eingabe! Bitte nur eine Zahl eingeben.", vbExclamation, "Ung" & ChrW(252) & "ltige Eingabe"
                        End If
                    Loop Until parzelleGueltig
                    
                    If parzelleGueltig Then
                        neueParzelle = CStr(parzelleNr)
                        neuerDebug = "Manuell: EHEMALIGES MITGLIED - Parzelle " & neueParzelle & "; nicht in Historie; " & Format(Now, "dd.mm.yyyy")
                    Else
                        neuerDebug = "Manuell: EHEMALIGES MITGLIED; nicht in Historie; " & Format(Now, "dd.mm.yyyy")
                    End If
                End If
            End If
        Case "VERSORGER"
            correctPrefix = PREFIX_VERSORGER
            If Left(UCase(currentEntityKey), Len(correctPrefix)) <> UCase(correctPrefix) Then
                neuerEntityKey = correctPrefix & mod_EntityKey_Manager.CreateGUID()
            Else
                neuerEntityKey = currentEntityKey
            End If
            neueZuordnung = mod_EntityKey_Normalize.ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            neuerDebug = "Manuell: VERSORGER (" & Format(Now, "dd.mm.yyyy") & ")"
            ampelStatus = 1
            
        Case "BANK"
            correctPrefix = PREFIX_BANK
            If Left(UCase(currentEntityKey), Len(correctPrefix)) <> UCase(correctPrefix) Then
                neuerEntityKey = correctPrefix & mod_EntityKey_Manager.CreateGUID()
            Else
                neuerEntityKey = currentEntityKey
            End If
            neueZuordnung = mod_EntityKey_Normalize.ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            neuerDebug = "Manuell: BANK (" & Format(Now, "dd.mm.yyyy") & ")"
            ampelStatus = 1
            
        Case "SHOP"
            correctPrefix = PREFIX_SHOP
            If Left(UCase(currentEntityKey), Len(correctPrefix)) <> UCase(correctPrefix) Then
                neuerEntityKey = correctPrefix & mod_EntityKey_Manager.CreateGUID()
            Else
                neuerEntityKey = currentEntityKey
            End If
            neueZuordnung = mod_EntityKey_Normalize.ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            neuerDebug = "Manuell: SHOP (" & Format(Now, "dd.mm.yyyy") & ")"
            ampelStatus = 1
            
        Case "SONSTIGE"
            correctPrefix = PREFIX_SONSTIGE
            If Left(UCase(currentEntityKey), Len(correctPrefix)) <> UCase(correctPrefix) Then
                neuerEntityKey = correctPrefix & mod_EntityKey_Manager.CreateGUID()
            Else
                neuerEntityKey = currentEntityKey
            End If
            neueZuordnung = mod_EntityKey_Normalize.ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            neuerDebug = "Manuell: SONSTIGE (" & Format(Now, "dd.mm.yyyy") & ")"
            ampelStatus = 1
            
        Case ""
            neuerEntityKey = ""
            neueZuordnung = ""
            neueParzelle = ""
            neuerDebug = ""
            ampelStatus = 3
            
        Case Else
            correctPrefix = PREFIX_SONSTIGE
            If Left(UCase(currentEntityKey), Len(correctPrefix)) <> UCase(correctPrefix) Then
                neuerEntityKey = correctPrefix & mod_EntityKey_Manager.CreateGUID()
            Else
                neuerEntityKey = currentEntityKey
            End If
            neueZuordnung = mod_EntityKey_Normalize.ExtrahiereAnzeigeName(kontoname)
            neueParzelle = ""
            neuerDebug = "Manuell: " & neueRole & " (" & Format(Now, "dd.mm.yyyy") & ")"
            ampelStatus = 2
    End Select
    
    ' EntityKey setzen
    If neuerEntityKey <> "" Then
        wsDaten.Cells(zeile, EK_COL_ENTITYKEY).value = neuerEntityKey
    ElseIf neueRole = "" Then
        wsDaten.Cells(zeile, EK_COL_ENTITYKEY).value = ""
    End If
    
    ' Zuordnung
    If neueZuordnung <> "" Then
        Dim aktuelleZuordnung As String
        aktuelleZuordnung = Trim(wsDaten.Cells(zeile, EK_COL_ZUORDNUNG).value)
        If aktuelleZuordnung = "" Then
            wsDaten.Cells(zeile, EK_COL_ZUORDNUNG).value = neueZuordnung
        End If
    ElseIf neueRole = "" Then
        wsDaten.Cells(zeile, EK_COL_ZUORDNUNG).value = ""
    End If
    
    ' Parzelle
    If mod_EntityKey_Classifier.DarfParzelleHaben(neueRole) Then
        If neueParzelle <> "" Then
            wsDaten.Cells(zeile, EK_COL_PARZELLE).value = neueParzelle
        End If
    Else
        If neueRole = "EHEMALIGES MITGLIED" And neueParzelle <> "" Then
            wsDaten.Cells(zeile, EK_COL_PARZELLE).value = neueParzelle
        Else
            wsDaten.Cells(zeile, EK_COL_PARZELLE).value = ""
        End If
    End If
    
    ' Debug-Spalte X
    wsDaten.Cells(zeile, EK_COL_DEBUG).value = neuerDebug
    
    ' Ampelfarbe
    Call mod_EntityKey_Ampel.SetzeAmpelFarbe(wsDaten, zeile, ampelStatus)
    
    ' Dropdowns
    Call SetupEntityRoleDropdown(wsDaten, zeile)
    
    If neueRole = "EHEMALIGES MITGLIED" Or neueRole = "SONSTIGE" Then
        Call SetupParzelleDropdown(wsDaten, zeile)
    End If
    
    ' U, W, X immer editierbar
    wsDaten.Cells(zeile, EK_COL_ZUORDNUNG).Locked = False
    wsDaten.Cells(zeile, EK_COL_ROLE).Locked = False
    wsDaten.Cells(zeile, EK_COL_DEBUG).Locked = False
    
    ' Sortierung + Ampelfarben sofort nach manueller Aenderung
    Call mod_Formatierung.FormatEntityKeyTableComplete(wsDaten)
    Call mod_EntityKey_Ampel.SetzeAlleAmpelfarbenNachSortierung(wsDaten)
    
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Application.EnableEvents = True
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    On Error Resume Next
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    Debug.Print "FEHLER in VerarbeiteManuelleRoleAenderung: " & Err.Description
End Sub

' ===============================================================
' Setzt EntityRole-Dropdown fuer eine Zeile
' ===============================================================
Public Sub SetupEntityRoleDropdown(ByRef ws As Worksheet, ByVal zeile As Long)
    Dim lastRowDD As Long
    
    On Error Resume Next
    
    lastRowDD = ws.Cells(ws.Rows.count, DATA_COL_DD_ENTITYROLE).End(xlUp).Row
    If lastRowDD < DATA_START_ROW Then lastRowDD = DATA_START_ROW
    
    ws.Cells(zeile, EK_COL_ROLE).Validation.Delete
    With ws.Cells(zeile, EK_COL_ROLE).Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:="=" & WS_DATEN & "!$AD$" & DATA_START_ROW & ":$AD$" & lastRowDD
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    
    On Error GoTo 0
End Sub

' ===============================================================
' Setzt Parzellen-Dropdown fuer eine Zeile
' ===============================================================
Public Sub SetupParzelleDropdown(ByRef ws As Worksheet, ByVal zeile As Long)
    Dim lastRowParzelle As Long
    
    On Error Resume Next
    
    lastRowParzelle = ws.Cells(ws.Rows.count, DATA_COL_DD_PARZELLE).End(xlUp).Row
    If lastRowParzelle < DATA_START_ROW Then lastRowParzelle = DATA_START_ROW
    
    ws.Cells(zeile, EK_COL_PARZELLE).Validation.Delete
    With ws.Cells(zeile, EK_COL_PARZELLE).Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:="=" & WS_DATEN & "!$F$" & DATA_START_ROW & ":$F$" & lastRowParzelle
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    
    ws.Cells(zeile, EK_COL_PARZELLE).Locked = False
    
    On Error GoTo 0
End Sub

' ===============================================================
' Kompatibilitaets-Stub (bewusst leer)
' ===============================================================
Public Sub FormatiereEntityKeyZeile(ByVal zeile As Long, Optional ByVal ws As Worksheet = Nothing)
    ' BEWUSST LEER
End Sub






