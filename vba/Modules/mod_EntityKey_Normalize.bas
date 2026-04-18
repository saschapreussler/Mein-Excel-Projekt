Attribute VB_Name = "mod_EntityKey_Normalize"
Option Explicit

' ***************************************************************
' MODUL: mod_EntityKey_Normalize
' ZWECK: String-Normalisierung und Hilfsfunktionen fuer EntityKey-System
' ABGELEITET AUS: mod_EntityKey_Manager (Modularisierung)
' VERSION: 1.0 - 01.03.2026
' FUNKTIONEN:
'   - EntferneMehrfacheLeerzeichen: Kollabiert Leerzeichen
'   - NormalisiereIBAN: IBAN-Standardisierung
'   - NormalisiereStringFuerVergleich: Fuzzy-Vergleich-Normalisierung
'   - ExtrahiereAnzeigeName: Erste Zeile eines Kontonamens
' ***************************************************************

' ===============================================================
' Entfernt mehrfache Leerzeichen
' ===============================================================
Public Function EntferneMehrfacheLeerzeichen(ByVal s As String) As String
    Dim result As String
    result = s
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    EntferneMehrfacheLeerzeichen = Trim(result)
End Function

' ===============================================================
' Normalisiert IBAN (Grossbuchstaben, keine Leerzeichen/Bindestriche)
' ===============================================================
Public Function NormalisiereIBAN(ByVal iban As Variant) As String
    Dim result As String
    
    If IsNull(iban) Or isEmpty(iban) Then
        NormalisiereIBAN = ""
        Exit Function
    End If
    
    result = UCase(Trim(CStr(iban)))
    result = Replace(result, " ", "")
    result = Replace(result, "-", "")
    
    NormalisiereIBAN = result
End Function

' ===============================================================
' Normalisiert String fuer Vergleich (Umlaute ersetzen, Kleinbuchstaben)
' ===============================================================
Public Function NormalisiereStringFuerVergleich(ByVal s As String) As String
    Dim result As String
    
    result = LCase(Trim(s))
    result = Replace(result, ",", " ")
    result = Replace(result, ".", " ")
    result = Replace(result, "-", " ")
    result = Replace(result, ChrW(228), "ae")
    result = Replace(result, ChrW(246), "oe")
    result = Replace(result, ChrW(252), "ue")
    result = Replace(result, ChrW(223), "ss")
    result = Replace(result, "ae", "a")
    result = Replace(result, "oe", "o")
    result = Replace(result, "ue", "u")
    
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    NormalisiereStringFuerVergleich = Trim(result)
End Function

' ===============================================================
' Extrahiert Anzeigename (erste Zeile, max. 50 Zeichen)
' ===============================================================
Public Function ExtrahiereAnzeigeName(ByVal kontoname As String) As String
    Dim zeilen() As String
    Dim erstesElement As String
    
    If kontoname = "" Then
        ExtrahiereAnzeigeName = ""
        Exit Function
    End If
    
    zeilen = Split(kontoname, vbLf)
    erstesElement = EntferneMehrfacheLeerzeichen(Trim(zeilen(0)))
    
    If Len(erstesElement) > 50 Then
        erstesElement = Left(erstesElement, 50) & "..."
    End If
    
    ExtrahiereAnzeigeName = erstesElement
End Function







































