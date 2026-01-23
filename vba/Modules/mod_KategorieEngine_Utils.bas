Attribute VB_Name = "mod_KategorieEngine_Utils"
Option Explicit

' ========================================================
' ApplyKategorie
' Setzt den Wert, die Ampelfarbe und Schriftfarbe in einer Zelle
' category: Text der Kategorie
' confidence: "GRUEN", "GELB" oder "ROT"
' ========================================================

Public Sub ApplyKategorie(ByVal targetCell As Range, _
                           ByVal category As String, _
                           ByVal confidence As String)

    With targetCell
        .value = category
        .Font.color = vbBlack
        .Interior.Pattern = xlSolid

        Select Case confidence
            Case "GRUEN"
                .Interior.color = RGB(198, 239, 206)
                .Font.color = vbBlack

            Case "GELB"
                .Interior.color = RGB(255, 235, 156)
                .Font.color = vbBlack

            Case "ROT"
                .Interior.color = RGB(255, 199, 206)
                .Font.color = vbRed
        End Select
    End With

End Sub


' ========================================================
' Prüft, ob ein Betrag einem Fixbetrag entspricht
' ========================================================
Public Function IsFixbetrag(ByVal amount As Double, _
                            ByVal fixValue As Double, _
                            Optional ByVal tolerance As Double = 0.01) As Boolean
    IsFixbetrag = (Abs(Abs(amount) - fixValue) <= tolerance)
End Function


