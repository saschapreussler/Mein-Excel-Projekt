Attribute VB_Name = "mod_Mapping_Tools"
Option Explicit

' ==========================================================
' MODUL: mod_Mapping_Tools (FINAL KORRIGIERT)
' Zweck: Bereitstellung von Hilfsfunktionen für Normalisierung und Fuzzy-Suche
' **********************************************************

' Definitionen für Match-Typen (Intern)
Private Const MATCH_NONE As Long = 0
Private Const MATCH_PARTIAL As Long = 1 ' Nur Vor- ODER Nachname gefunden (Gelb)
Private Const MATCH_FULL As Long = 2 ' Vor- UND Nachname gefunden (Grün)


Private Function NormalizeString(ByVal inputStr As String) As String
    ' Normalisiert Strings für tolerante Vergleiche (Umlaute, Groß-/Kleinschreibung, ß)
    
    Dim tempStr As String
    tempStr = LCase(Trim(inputStr))
    
    ' Ersetzung der Umlaute (doppelte Schreibweise)
    tempStr = Replace(tempStr, "ae", "a")
    tempStr = Replace(tempStr, "oe", "o")
    tempStr = Replace(tempStr, "ue", "u")
    
    ' Ersetzung der Umlaute (einfache Schreibweise)
    tempStr = Replace(tempStr, "ä", "a")
    tempStr = Replace(tempStr, "ö", "o")
    tempStr = Replace(tempStr, "ü", "u")
    
    ' Ersetzung von scharfem S
    tempStr = Replace(tempStr, "ß", "ss")
    
    ' Entfernen von Interpunktion und unnötigen Leerzeichen
    tempStr = Replace(tempStr, ",", "")
    tempStr = Replace(tempStr, ".", "")
    tempStr = Replace(tempStr, "-", "")
    tempStr = Replace(tempStr, "/", "")
    
    ' Mehrere Leerzeichen durch eines ersetzen
    Do While InStr(tempStr, "  ") > 0
        tempStr = Replace(tempStr, "  ", " ")
    Loop
    
    NormalizeString = tempStr
End Function

Public Function FuzzyMemberSearch(ByVal nameToSearch As String, ByVal wsMembers As Worksheet, ByRef parzelleRange As Range) As String
    ' Sucht nach einem Mitglied und gibt den besten Match zurück.
    ' Rückgabe: String mit nur den besten, einzigartigen Treffern (mit vbLf getrennt).
    
    Dim lastRowM As Long
    Dim r As Long
    Dim memberLastName As String, memberFirstName As String
    Dim parzelle As String
    
    Dim normSearchName As String
    
    ' --- Variablen zur Speicherung des besten Matches ---
    ' memberName -> MatchStatus (2=FULL, 1=PARTIAL)
    Dim dictAllMatches As Object
    ' memberName -> Parzelle(n)
    Dim dictParzellenMap As Object
    
    Dim bestMatchStatus As Long ' 0: Keiner, 1: Teil, 2: Voll
    
    Set dictAllMatches = CreateObject("Scripting.Dictionary")
    Set dictParzellenMap = CreateObject("Scripting.Dictionary")
    
    bestMatchStatus = MATCH_NONE
    
    normSearchName = NormalizeString(nameToSearch)
    
    ' HINWEIS: Wir verwenden nun die in mod_Const neu definierten MEMBER_COL_ Konstanten.
    lastRowM = wsMembers.Cells(wsMembers.Rows.count, MEMBER_COL_NACHNAME).End(xlUp).Row
    
    If lastRowM < M_START_ROW Then GoTo EndSearch
    
    Dim currentMemberFullnameString As String
    
    For r = M_START_ROW To lastRowM
        ' Konstanten M_COL_NACHNAME, M_COL_VORNAME wurden durch MEMBER_COL_... ersetzt
        memberLastName = Trim(wsMembers.Cells(r, MEMBER_COL_NACHNAME).value)
        memberFirstName = Trim(wsMembers.Cells(r, MEMBER_COL_VORNAME).value)
        parzelle = Trim(wsMembers.Cells(r, MEMBER_COL_PARZELLE).value)
        
        If memberLastName = "" And memberFirstName = "" Then GoTo NextMember
        
        currentMemberFullnameString = Trim(memberFirstName & " " & memberLastName)
        
        Dim normMemberLast As String
        Dim normMemberFirst As String
        
        normMemberLast = NormalizeString(memberLastName)
        normMemberFirst = NormalizeString(memberFirstName)
        
        Dim currentMatchStatus As Long
        Dim matchFoundName As String
        currentMatchStatus = MATCH_NONE
        
        ' -------------------------------------------------------------
        ' LOGIK: STATUS BESTIMMEN
        ' -------------------------------------------------------------
        
        ' --- PRÜFUNG: Voll-Match (Vor- UND Nachname enthalten) ---
        If InStr(normSearchName, normMemberLast) > 0 And normMemberLast <> "" And _
           InStr(normSearchName, normMemberFirst) > 0 And normMemberFirst <> "" Then
            
            currentMatchStatus = MATCH_FULL
            matchFoundName = currentMemberFullnameString ' Rückgabe des kompletten Originalnamens
            
        ' --- PRÜFUNG: Teil-Match (Nur Vor- ODER Nachname enthalten) ---
        ElseIf InStr(normSearchName, normMemberLast) > 0 And normMemberLast <> "" Then
            ' Nur Nachname gefunden
            currentMatchStatus = MATCH_PARTIAL
            matchFoundName = memberLastName ' Rückgabe des Original-Nachnamens
        
        ElseIf InStr(normSearchName, normMemberFirst) > 0 And normMemberFirst <> "" Then
            ' Nur Vorname gefunden
            currentMatchStatus = MATCH_PARTIAL
            matchFoundName = memberFirstName ' Rückgabe des Original-Vornamens
        End If
        
        
        If currentMatchStatus > MATCH_NONE Then
            
            ' Führt den besten Match-Status nach
            If currentMatchStatus > bestMatchStatus Then
                bestMatchStatus = currentMatchStatus
            End If
            
            ' Alle Matches speichern, damit wir später nur die besten aggregieren können
            If Not dictAllMatches.Exists(matchFoundName) Then
                dictAllMatches.Add matchFoundName, currentMatchStatus
            End If
            
            ' Parzelle(n) zum Mitglied hinzufügen (für Spalte W)
            If dictParzellenMap.Exists(matchFoundName) Then
                 ' Wenn bereits ein Eintrag existiert, Parzelle mit Komma/vbLf anhängen
                If InStr(dictParzellenMap.item(matchFoundName), parzelle) = 0 Then
                    dictParzellenMap.item(matchFoundName) = dictParzellenMap.item(matchFoundName) & vbLf & parzelle
                End If
            Else
                dictParzellenMap.Add matchFoundName, parzelle
            End If
            
        End If

NextMember:
    Next r
    
    
EndSearch:
    ' ------------------------------------------------------------------
    ' ERGEBNISFILTERUNG: Nur die Treffer mit dem höchsten Status zählen
    ' ------------------------------------------------------------------
    Dim finalZuordnung As String
    Dim finalParzellen As String
    Dim memberName As Variant
    
    If bestMatchStatus = MATCH_NONE Then
         parzelleRange.value = "" ' Leert W, wenn kein Treffer
         FuzzyMemberSearch = ""
         Exit Function
    End If
    
    ' Listen zur Vermeidung von Duplikaten in der Endausgabe
    Dim listUniqueNames As Object: Set listUniqueNames = CreateObject("Scripting.Dictionary")
    Dim listUniqueParzellen As Object: Set listUniqueParzellen = CreateObject("Scripting.Dictionary")
    
    For Each memberName In dictAllMatches.Keys
        ' **HIER IST DIE KORREKTUR:** Wir akzeptieren nur Treffer, die dem besten Status entsprechen
        If dictAllMatches.item(memberName) = bestMatchStatus Then
            
            ' Nur einzigartige Namen in V sammeln
            If Not listUniqueNames.Exists(memberName) Then
                If finalZuordnung <> "" Then finalZuordnung = finalZuordnung & vbLf
                finalZuordnung = finalZuordnung & memberName
                listUniqueNames.Add memberName, True
            End If
            
            ' Parzellen für W sammeln (aus dem Parzellen-Dictionary)
            If dictParzellenMap.Exists(memberName) Then
                 Dim parts() As String
                 parts = Split(dictParzellenMap.item(memberName), vbLf)
                 
                 Dim part As Variant
                 For Each part In parts
                      If Not listUniqueParzellen.Exists(part) Then
                          If finalParzellen <> "" Then finalParzellen = finalParzellen & vbLf
                          finalParzellen = finalParzellen & part
                          listUniqueParzellen.Add part, True
                      End If
                 Next part
            End If
            
        End If
    Next
    
    ' Ergebnis setzen
    parzelleRange.value = finalParzellen
    FuzzyMemberSearch = finalZuordnung
    
End Function

