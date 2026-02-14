Attribute VB_Name = "mod_KategorieEngine_Scoring"
Option Explicit

' =====================================================
' KATEGORIE-ENGINE - SCORING & FILTER
' Ausgelagert aus mod_KategorieEngine_Evaluator
' Enth�lt: Keyword-Matching, Score-Boni, EntityRole-Filter
' =====================================================


' =====================================================
' MULTI-WORD-MATCHING (v7.0)
' Pr�ft ob ALLE W�rter des Keywords im Text vorkommen.
' Reihenfolge ist egal. Zusammengeschriebene W�rter
' werden ebenfalls erkannt (Substring-Matching je Wort).
' =====================================================
Public Function MatchKeyword(ByVal normText As String, _
                              ByVal normKeyword As String) As Boolean
    
    If InStr(normKeyword, " ") = 0 Then
        MatchKeyword = (InStr(normText, normKeyword) > 0)
        Exit Function
    End If
    
    Dim woerter() As String
    woerter = Split(normKeyword, " ")
    
    Dim w As Long
    For w = LBound(woerter) To UBound(woerter)
        If Len(woerter(w)) > 0 Then
            If InStr(normText, woerter(w)) = 0 Then
                MatchKeyword = False
                Exit Function
            End If
        End If
    Next w
    
    MatchKeyword = True
End Function

' =====================================================
' ExactMatchBonus (v8.0)
' Gibt Bonuspunkte wenn das normalisierte Keyword als
' zusammenh�ngender Substring im Text vorkommt.
' =====================================================
Public Function ExactMatchBonus(ByVal normText As String, _
                                 ByVal normKeyword As String) As Long
    If InStr(normText, normKeyword) > 0 Then
        ExactMatchBonus = 10
    Else
        ExactMatchBonus = 0
    End If
End Function

' =====================================================
' WordCountBonus (v9.3 - ersetzt CoverageBonus)
' Z�hlt die W�rter im normalisierten Keyword und
' gibt pro Wort 5 Punkte Bonus. L�ngere/spezifischere
' Keywords mit mehr W�rtern bekommen dadurch mehr Punkte.
'
' Beispiel: normText = "max mustermann stvom wasser parz 9 gutschrift"
'   Keyword "stvom wasser parz 9" -> 4 W�rter -> +20
'   Keyword "wasser parz 9"       -> 3 W�rter -> +15
'   Keyword "wasser"               -> 1 Wort   -> +5
'
' Zusammen mit dem erh�hten Prio-Bonus (10-prio)*8
' ergibt sich bei Prio1 vs Prio3 eine Differenz von
' 16 (Prio) + 5 (WordCount) = 21 >= SCHWELLE 20
' =====================================================
Public Function WordCountBonus(ByVal normKeyword As String) As Long
    If Len(normKeyword) = 0 Then
        WordCountBonus = 0
        Exit Function
    End If
    
    Dim woerter() As String
    woerter = Split(normKeyword, " ")
    
    Dim anzahl As Long
    Dim w As Long
    anzahl = 0
    For w = LBound(woerter) To UBound(woerter)
        If Len(woerter(w)) > 0 Then
            anzahl = anzahl + 1
        End If
    Next w
    
    WordCountBonus = anzahl * 5
End Function


' =====================================================
' FILTER: Strenge EntityRole-Kategorie-Trennung (v7.0)
' Detaillierte Logik wiederhergestellt!
' =====================================================
Public Function PasstEntityRoleZuKategorie(ByVal ctx As Object, _
                                            ByVal category As String, _
                                            ByVal einAus As String) As Boolean
    
    Dim catLower As String
    catLower = LCase(category)
    Dim role As String
    role = ctx("EntityRole")
    
    PasstEntityRoleZuKategorie = True
    
    If role = "" Then Exit Function
    
    ' --- VERSORGER: Nur Versorger-typische Kategorien ---
    If ctx("IsVersorger") Then
        If catLower Like "*mitglied*" Then PasstEntityRoleZuKategorie = False: Exit Function
        
        If catLower Like "*pacht*" And catLower Like "*mitglied*" Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
        
        If catLower Like "*endabrechnung*" And catLower Like "*mitglied*" Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
        If catLower Like "*vorauszahlung*" And catLower Like "*mitglied*" Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
        If catLower Like "*spende*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*beitrag*" And Not catLower Like "*verband*" Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
        If catLower Like "*sammelzahlung*" Then PasstEntityRoleZuKategorie = False: Exit Function
    End If
    
    ' --- MITGLIED: Nur Mitglieder-typische Kategorien ---
    If ctx("IsMitglied") Then
        If catLower Like "*versorger*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*stadtwerke*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*energieversorger*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*wasserwerk*" Then PasstEntityRoleZuKategorie = False: Exit Function
        
        If catLower Like "*rueckzahlung*versorger*" Or _
           catLower Like "*r" & ChrW(252) & "ckzahlung*versorger*" Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
        
        If catLower Like "*miete*" And (catLower Like "*grundst" & ChrW(252) & "ck*" Or catLower Like "*grundstueck*") Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
        
        If catLower Like "*entgeltabschluss*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*kontof" & ChrW(252) & "hrung*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*kontofuehrung*" Then PasstEntityRoleZuKategorie = False: Exit Function
        
        ' Mitglied bei Ausgabe = nur R�ckerstattung/Auszahlung/Guthaben
        If ctx("IsAusgabe") Then
            If Not (catLower Like "*r" & ChrW(252) & "ck*" Or catLower Like "*rueck*" Or _
                    catLower Like "*erstattung*" Or catLower Like "*gutschrift*" Or _
                    catLower Like "*auszahlung*" Or catLower Like "*guthaben*") Then
                PasstEntityRoleZuKategorie = False: Exit Function
            End If
        End If
    End If
    
    ' --- BANK: Nur Bank-typische Kategorien ---
    If ctx("IsBank") Then
        If Not (catLower Like "*bank*" Or _
                catLower Like "*entgelt*" Or _
                catLower Like "*geb" & ChrW(252) & "hr*" Or catLower Like "*gebuehr*" Or _
                catLower Like "*kontof" & ChrW(252) & "hrung*" Or catLower Like "*kontofuehrung*" Or _
                catLower Like "*zins*") Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
    End If
    
    ' --- EHEMALIGES MITGLIED ---
    If ctx("IsEhemaligesMitglied") Then
        If catLower Like "*versorger*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*stadtwerke*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*entgeltabschluss*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*kontof" & ChrW(252) & "hrung*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*kontofuehrung*" Then PasstEntityRoleZuKategorie = False: Exit Function
        If catLower Like "*miete*" And (catLower Like "*grundst" & ChrW(252) & "ck*" Or catLower Like "*grundstueck*") Then
            PasstEntityRoleZuKategorie = False: Exit Function
        End If
        
        ' Ehemalige bei Ausgabe: Auszahlung/Guthaben/R�ckzahlung erlaubt
        If ctx("IsAusgabe") Then
            If Not (catLower Like "*r" & ChrW(252) & "ck*" Or catLower Like "*rueck*" Or _
                    catLower Like "*erstattung*" Or catLower Like "*gutschrift*" Or _
                    catLower Like "*auszahlung*" Or catLower Like "*guthaben*" Or _
                    catLower Like "*endabrechnung*") Then
                PasstEntityRoleZuKategorie = False: Exit Function
            End If
        End If
    End If
    
End Function
