VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ===============================================================
' TABELLENBLATT-MODUL: Vereinskasse
' VERSION: 1.0 - 18.04.2026
' ZWECK: Monatsfilter (ComboBox cbo_MonatFilter_VK)
'        Analog zu Bankkonto (Tabelle3)
' HINWEIS: Dieses Modul gehoert zum Blatt "Vereinskasse".
'          Falls der VBA-Codename nicht "Tabelle4" ist,
'          bitte den Code in das korrekte Sheet-Modul verschieben.
' ===============================================================

Private Sub Worksheet_Activate()
    ' ComboBox befuellen und Filter zuruecksetzen
    Call mod_Vereinskasse_Filter.BefuelleVereinskasseComboBox(Me)
End Sub


Private Sub cbo_MonatFilter_VK_Change()
    ' Filter anwenden basierend auf ComboBox-Auswahl
    If Me.cbo_MonatFilter_VK.ListIndex = -1 Then Exit Sub
    
    Call mod_Vereinskasse_Filter.WendeVereinskasseFilterAn(Me, Me.cbo_MonatFilter_VK.value)
End Sub


