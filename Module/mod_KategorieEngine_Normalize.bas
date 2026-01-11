Attribute VB_Name = "mod_KategorieEngine_Normalize"
Option Explicit

Public Function NormalizeBankkontoZeile(ByVal wsBK As Worksheet, _
                                         ByVal rowBK As Long) As String
    Dim rawText As String
    rawText = _
        Trim(wsBK.Cells(rowBK, BK_COL_NAME).Value) & " " & _
        Trim(wsBK.Cells(rowBK, BK_COL_VERWENDUNGSZWECK).Value) & " " & _
        Trim(wsBK.Cells(rowBK, BK_COL_BUCHUNGSTEXT).Value)

    NormalizeBankkontoZeile = NormalizeText(rawText)
End Function

Public Function NormalizeText(ByVal inputText As String) As String
    Dim txt As String
    txt = LCase(Trim(inputText))
    If txt = "" Then NormalizeText = "": Exit Function

    txt = Replace(txt, "ä", "ae")
    txt = Replace(txt, "ö", "oe")
    txt = Replace(txt, "ü", "ue")
    txt = Replace(txt, "ß", "ss")
    txt = Replace(txt, "mitgliets", "mitglieds")
    txt = Replace(txt, "mitgliedbetrag", "mitgliedsbeitrag")
    txt = Replace(txt, "mitglied beitrag", "mitgliedsbeitrag")
    txt = Replace(txt, "beitragsgeb hr", "beitragsgebuehr")
    txt = Replace(txt, "abschlag", "abschlagszahlung")
    txt = Replace(txt, "entgelt abschluss", "entgeltabschluss")

    Dim i As Long
    For i = 1 To Len(txt)
        Select Case Mid$(txt, i, 1)
            Case "a" To "z", "0" To "9", " "
            Case Else
                Mid$(txt, i, 1) = " "
        End Select
    Next i

    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop

    NormalizeText = Trim(txt)
End Function


