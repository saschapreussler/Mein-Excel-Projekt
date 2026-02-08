Attribute VB_Name = "mod_KategorieEngine_Apply"
Option Explicit

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


