Attribute VB_Name = "mod_KategorieEngine_Normalize"
Option Explicit

Public Function NormalizeBankkontoZeile(ByVal wsBK As Worksheet, _
                                         ByVal rowBK As Long) As String
    Dim rawText As String
    rawText = _
        Trim(wsBK.Cells(rowBK, BK_COL_NAME).value) & " " & _
        Trim(wsBK.Cells(rowBK, BK_COL_VERWENDUNGSZWECK).value) & " " & _
        Trim(wsBK.Cells(rowBK, BK_COL_BUCHUNGSTEXT).value)

    NormalizeBankkontoZeile = NormalizeText(rawText)
End Function

Public Function NormalizeText(ByVal inputText As String) As String
    Dim txt As String
    txt = LCase(Trim(inputText))
    If txt = "" Then NormalizeText = "": Exit Function

    ' Umlaute ersetzen (für Matching-Normalisierung)
    txt = Replace(txt, ChrW(228), "ae")  ' ä
    txt = Replace(txt, ChrW(246), "oe")  ' ö
    txt = Replace(txt, ChrW(252), "ue")  ' ü
    txt = Replace(txt, ChrW(223), "ss")  ' ß
    txt = Replace(txt, ChrW(196), "ae")  ' Ä
    txt = Replace(txt, ChrW(214), "oe")  ' Ö
    txt = Replace(txt, ChrW(220), "ue")  ' Ü

    ' Typische Tippfehler korrigieren
    txt = Replace(txt, "mitgliets", "mitglieds")
    txt = Replace(txt, "mitgliedbetrag", "mitgliedsbeitrag")
    txt = Replace(txt, "mitglied beitrag", "mitgliedsbeitrag")
    txt = Replace(txt, "beitragsgeb hr", "beitragsgebuehr")
    txt = Replace(txt, "entgelt abschluss", "entgeltabschluss")

    ' WICHTIG: "abschlag" -> "abschlagszahlung" Expansion
    ' ABER: Nur wenn "abschlagszahlung" noch NICHT im Text steht!
    ' Sonst entsteht "abschlagszahlungszahlung" (doppelte Expansion)
    If InStr(txt, "abschlagszahlung") = 0 Then
        txt = Replace(txt, "abschlag", "abschlagszahlung")
    End If

    ' Sonderzeichen entfernen (nur a-z, 0-9, Leerzeichen behalten)
    Dim i As Long
    For i = 1 To Len(txt)
        Select Case Mid$(txt, i, 1)
            Case "a" To "z", "0" To "9", " "
            Case Else
                Mid$(txt, i, 1) = " "
        End Select
    Next i

    ' Mehrfache Leerzeichen zusammenfassen
    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop

    NormalizeText = Trim(txt)
End Function

