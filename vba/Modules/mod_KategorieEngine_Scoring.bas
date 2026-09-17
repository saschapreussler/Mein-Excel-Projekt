Attribute VB_Name = "mod_KategorieEngine_Scoring"
Option Explicit

' =====================================================
' KATEGORIE-ENGINE - SCORING & FILTER
' Ausgelagert aus mod_KategorieEngine_Evaluator
' enthält: Keyword-Matching, Score-Boni, EntityRole-Filter
' =====================================================

' Mindestlänge eines Schlüsselworts für die gelockerte
' Rückfallprüfung. Kürzere Schlüsselwörter bleiben streng, weil
' bei ihnen Zufallstreffer zu wahrscheinlich sind.
Private Const MIN_LEN_TOLERANT As Long = 6

' Zwischenspeicher der Textvarianten der aktuellen Buchung.
Private m_cacheText As String
Private m_cacheGefuellt As Boolean
Private m_ohneLeerzeichen As String
Private m_grundform As String


' =====================================================
' MULTI-WORD-MATCHING (v7.0)
' prüft ob ALLE Wörter des Keywords im Text vorkommen.
' Reihenfolge ist egal. Zusammengeschriebene Wörter
' werden ebenfalls erkannt (Substring-Matching je Wort).
' =====================================================
Public Function MatchKeyword(ByVal normText As String, _
                              ByVal normKeyword As String, _
                              Optional ByRef istTolerant As Boolean) As Boolean

    istTolerant = False

    ' Stufe 1: unveränderte strenge Prüfung wie bisher.
    If InStr(normKeyword, " ") = 0 Then
        If InStr(normText, normKeyword) > 0 Then
            MatchKeyword = True
            Exit Function
        End If
        MatchKeyword = MatchKeywordTolerant(normText, normKeyword)
        istTolerant = MatchKeyword
        Exit Function
    End If
    
    Dim woerter() As String
    woerter = Split(normKeyword, " ")
    
    Dim w As Long
    For w = LBound(woerter) To UBound(woerter)
        If Len(woerter(w)) > 0 Then
            If InStr(normText, woerter(w)) = 0 Then
                ' Ein Wort fehlt -> gelockerte Rückfallprüfung versuchen.
                MatchKeyword = MatchKeywordTolerant(normText, normKeyword)
                istTolerant = MatchKeyword
                Exit Function
            End If
        End If
    Next w
    
    MatchKeyword = True
End Function


' =====================================================
' GELOCKERTE RÜCKFALLPRÜFUNG (Stufe 2 und 3)
' Wird ausschließlich aufgerufen, wenn die strenge Prüfung
' nichts gefunden hat. Dadurch kann kein bisher funktionierender
' Treffer verloren gehen, es können nur neue hinzukommen.
'
' Stufe 2: Vergleich ohne Leerzeichen. Erkennt getrennt
'          geschriebene Wörter wie "Fix Kosten".
' Stufe 3: Vergleich ohne Leerzeichen und ohne angehängtes "e"
'          je Wort. Erkennt gebeugte Adjektive, also zum Beispiel
'          "fixe kosten" gegen das Schlüsselwort "fixkosten".
'
' Ein Treffer aus dieser Funktion gilt als unsicher. Der Evaluator
' setzt ihn deshalb auf GELB zur Bestätigung und niemals
' automatisch auf GRÜN.
' =====================================================
Private Function MatchKeywordTolerant(ByVal normText As String, _
                                       ByVal normKeyword As String) As Boolean

    Dim flachKeyword As String

    flachKeyword = Replace(normKeyword, " ", "")
    If Len(flachKeyword) < MIN_LEN_TOLERANT Then Exit Function

    BaueTextVarianten normText

    If InStr(m_ohneLeerzeichen, flachKeyword) > 0 Then
        MatchKeywordTolerant = True
        Exit Function
    End If

    If InStr(m_grundform, flachKeyword) > 0 Then
        MatchKeywordTolerant = True
    End If
End Function


' =====================================================
' Baut die beiden Textvarianten für die Rückfallprüfung.
' Der Evaluator prüft alle Regelzeilen einer Buchung
' hintereinander, deshalb genügt ein einziger Zwischenspeicher,
' um die Varianten nur einmal je Buchung zu erzeugen statt
' einmal je Regelzeile.
' =====================================================
Private Sub BaueTextVarianten(ByVal normText As String)

    Dim woerter() As String
    Dim w As Long
    Dim wort As String
    Dim grund As String

    If m_cacheGefuellt Then
        If m_cacheText = normText Then Exit Sub
    End If

    m_cacheText = normText
    m_ohneLeerzeichen = Replace(normText, " ", "")

    grund = ""
    woerter = Split(normText, " ")
    For w = LBound(woerter) To UBound(woerter)
        wort = woerter(w)
        ' Nur ein einzelnes angehängtes "e" entfernen. Die Endungen
        ' "en", "er" und "es" bleiben stehen, damit Wörter wie
        ' "kosten" oder "wasser" unverändert erhalten bleiben.
        If Len(wort) >= 4 Then
            If Right$(wort, 1) = "e" Then wort = Left$(wort, Len(wort) - 1)
        End If
        grund = grund & wort
    Next w

    m_grundform = grund
    m_cacheGefuellt = True
End Sub

' =====================================================
' ExactMatchBonus (v8.0)
' Gibt Bonuspunkte wenn das normalisierte Keyword als
' zusammenhängender Substring im Text vorkommt.
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
' zählt die Wörter im normalisierten Keyword und
' gibt pro Wort 5 Punkte Bonus. Längere/spezifischere
' Keywords mit mehr Wörtern bekommen dadurch mehr Punkte.
'
' Beispiel: normText = "max mustermann stvom wasser parz 9 gutschrift"
'   Keyword "stvom wasser parz 9" -> 4 Wörter -> +20
'   Keyword "wasser parz 9"       -> 3 Wörter -> +15
'   Keyword "wasser"               -> 1 Wort   -> +5
'
' Zusammen mit dem erhöhten Prio-Bonus (10-prio)*8
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
        
        ' Mitglied bei Ausgabe = nur Rückerstattung/Auszahlung/Guthaben
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
        
        ' Ehemalige bei Ausgabe: Auszahlung/Guthaben/Rückzahlung erlaubt
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






















































































































































