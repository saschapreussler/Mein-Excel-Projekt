Attribute VB_Name = "mod_EntityKey_Kontoname"
Option Explicit

' ***************************************************************
' MODUL: mod_EntityKey_Kontoname
' ZWECK: Kontonamen-Deduplizierung und -Bereinigung
' ABGELEITET AUS: mod_EntityKey_Manager (Modularisierung)
' VERSION: 1.0 - 01.03.2026
' FUNKTIONEN:
'   - IstKontonameRedundant: Semantische Redundanzpruefung
'   - ZerlegeInWorte: Name in normalisierte Wortmenge zerlegen
'   - SindWortmengenGleich: Wortmengen-Gleichheit pruefen
'   - IstTeilmenge: Teilmengen-Pruefung
'   - BereinigeKontonamen: Dictionary von Redundanzen bereinigen
'   - SammelKontonamen: Dictionary zu String zusammenfuegen
' ***************************************************************

' ===============================================================
' Prueft ob ein Kontoname semantisch redundant ist
' ===============================================================
Public Function IstKontonameRedundant(ByRef dictNames As Object, ByVal neuerName As String) As Boolean
    Dim key As Variant
    Dim bestehenderName As String
    Dim neueWorte As Object
    Dim bestehendeWorte As Object
    
    IstKontonameRedundant = False
    
    Set neueWorte = ZerlegeInWorte(neuerName)
    
    For Each key In dictNames.keys
        bestehenderName = CStr(dictNames(key))
        Set bestehendeWorte = ZerlegeInWorte(bestehenderName)
        
        If SindWortmengenGleich(neueWorte, bestehendeWorte) Then
            IstKontonameRedundant = True
            Exit Function
        End If
        
        If IstTeilmenge(neueWorte, bestehendeWorte) Then
            IstKontonameRedundant = True
            Exit Function
        End If
    Next key
End Function

' ===============================================================
' Zerlegt einen Namen in normalisierte Worte (Dictionary)
' ===============================================================
Public Function ZerlegeInWorte(ByVal Name As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim parts() As String
    Dim wort As String
    Dim i As Long
    
    Name = UCase(Trim(Name))
    Name = Replace(Name, ",", " ")
    Name = Replace(Name, ".", " ")
    Name = Replace(Name, "-", " ")
    Name = mod_EntityKey_Normalize.EntferneMehrfacheLeerzeichen(Name)
    
    If Name = "" Then
        Set ZerlegeInWorte = dict
        Exit Function
    End If
    
    parts = Split(Name, " ")
    
    For i = LBound(parts) To UBound(parts)
        wort = Trim(parts(i))
        If wort <> "" And wort <> "UND" And wort <> "U" Then
            If Not dict.Exists(wort) Then
                dict.Add wort, True
            End If
        End If
    Next i
    
    Set ZerlegeInWorte = dict
End Function

' ===============================================================
' Prueft ob zwei Wortmengen identisch sind
' ===============================================================
Public Function SindWortmengenGleich(ByRef dict1 As Object, ByRef dict2 As Object) As Boolean
    Dim key As Variant
    
    SindWortmengenGleich = False
    
    If dict1.count <> dict2.count Then Exit Function
    If dict1.count = 0 Then Exit Function
    
    For Each key In dict1.keys
        If Not dict2.Exists(key) Then Exit Function
    Next key
    
    SindWortmengenGleich = True
End Function

' ===============================================================
' Prueft ob dict1 eine Teilmenge von dict2 ist
' ===============================================================
Public Function IstTeilmenge(ByRef dictKlein As Object, ByRef dictGross As Object) As Boolean
    Dim key As Variant
    
    IstTeilmenge = False
    
    If dictKlein.count = 0 Then Exit Function
    If dictKlein.count >= dictGross.count Then Exit Function
    
    For Each key In dictKlein.keys
        If Not dictGross.Exists(key) Then Exit Function
    Next key
    
    IstTeilmenge = True
End Function

' ===============================================================
' Bereinigt Dictionary von redundanten Kontonamen
' ===============================================================
Public Function BereinigeKontonamen(ByRef dictNames As Object) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    Dim keys() As Variant
    Dim values() As Variant
    Dim istRedundant() As Boolean
    Dim i As Long, j As Long
    Dim cnt As Long
    Dim worteI As Object, worteJ As Object
    
    cnt = dictNames.count
    If cnt = 0 Then
        Set BereinigeKontonamen = result
        Exit Function
    End If
    
    If cnt = 1 Then
        Dim singleKey As Variant
        For Each singleKey In dictNames.keys
            result.Add singleKey, dictNames(singleKey)
        Next singleKey
        Set BereinigeKontonamen = result
        Exit Function
    End If
    
    ReDim keys(0 To cnt - 1)
    ReDim values(0 To cnt - 1)
    ReDim istRedundant(0 To cnt - 1)
    
    i = 0
    Dim k As Variant
    For Each k In dictNames.keys
        keys(i) = k
        values(i) = dictNames(k)
        istRedundant(i) = False
        i = i + 1
    Next k
    
    For i = 0 To cnt - 1
        If Not istRedundant(i) Then
            Set worteI = ZerlegeInWorte(CStr(values(i)))
            For j = i + 1 To cnt - 1
                If Not istRedundant(j) Then
                    Set worteJ = ZerlegeInWorte(CStr(values(j)))
                    
                    If SindWortmengenGleich(worteI, worteJ) Then
                        If Len(CStr(values(i))) >= Len(CStr(values(j))) Then
                            istRedundant(j) = True
                        Else
                            istRedundant(i) = True
                            Exit For
                        End If
                    ElseIf IstTeilmenge(worteI, worteJ) Then
                        istRedundant(i) = True
                        Exit For
                    ElseIf IstTeilmenge(worteJ, worteI) Then
                        istRedundant(j) = True
                    End If
                End If
            Next j
        End If
    Next i
    
    For i = 0 To cnt - 1
        If Not istRedundant(i) Then
            result.Add keys(i), values(i)
        End If
    Next i
    
    Set BereinigeKontonamen = result
End Function

' ===============================================================
' Sammelt alle Kontonamen aus Dictionary zu String (LF-getrennt)
' ===============================================================
Public Function SammelKontonamen(ByRef dictNames As Object) As String
    Dim key As Variant
    Dim result As String
    Dim cleanName As String
    
    result = ""
    
    For Each key In dictNames.keys
        cleanName = mod_EntityKey_Normalize.EntferneMehrfacheLeerzeichen(Trim(CStr(dictNames(key))))
        If cleanName <> "" Then
            If result <> "" Then
                result = result & vbLf & cleanName
            Else
                result = cleanName
            End If
        End If
    Next key
    
    SammelKontonamen = result
End Function





































