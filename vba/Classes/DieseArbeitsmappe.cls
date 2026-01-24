VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    Call Initialize_ImportReport_ListBox

    ' Blattschutz vorerst deaktiviert - bedingte Formatierung funktioniert besser ohne
    ' With Worksheets(WS_MITGLIEDER)
    '     .Protect PASSWORD:="", UserInterfaceOnly:=True
    ' End With

    ' ApplyMitgliederDropdowns nur manuell aufrufen wenn nötig
    ' Call mod_Mitglieder_UI.ApplyMitgliederDropdowns

End Sub
