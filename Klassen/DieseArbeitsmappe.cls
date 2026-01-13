VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()

    ' WICHTIG: Stellen Sie sicher, dass WS_MITGLIEDER in mod_Hilfsfunktionen definiert ist!

    Call Initialize_ImportReport_ListBox ' Annahme: Diese Prozedur existiert

    With Worksheets(WS_MITGLIEDER)
        ' HINWEIS: Das Passwort muss hier über die Konstante in mod_Hilfsfunktionen übergeben werden,
        ' da wir dort Public Const PASSWORD definiert haben.
        .Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    End With

    Call ApplyMitgliederDropdowns ' <<< DAS fehlte (Annahme: Diese Prozedur existiert)

End Sub
