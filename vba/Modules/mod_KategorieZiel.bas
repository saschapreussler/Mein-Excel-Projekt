Attribute VB_Name = "mod_KategorieZiel"
Option Explicit

' ===============================================================
' MODUL: mod_KategorieZiel
' ZWECK:
'   Dynamisches DropDown für Zielspalte (Daten!Spalte N)
'   abhängig von Einnahme / Ausgabe (Spalte K)
' ===============================================================

Private Const RULE_START_ROW As Long = 4
Private Const MAX_RULE_ROWS As Long = 1000

Private Const COL_EIN_AUS As Long = 11    ' Spalte K
Private Const COL_ZIELSPALTE As Long = 14 ' Spalte N

' ---------------------------------------------------------------
' Initialisierung für kompletten Bereich (manuell aufrufbar)
' ---------------------------------------------------------------
Public Sub Init_ZielspaltenDropdowns()

    Dim wsRules As Worksheet
    Set wsRules = ThisWorkbook.Worksheets(WS_DATEN)

    Dim r As Long
    For r = RULE_START_ROW To RULE_START_ROW + MAX_RULE_ROWS
        ApplyZielspaltenDropdown wsRules, r
    Next r

End Sub

' ---------------------------------------------------------------
' Kernlogik: Dropdown je Zeile setzen
' ---------------------------------------------------------------
Public Sub ApplyZielspaltenDropdown(ByVal ws As Worksheet, ByVal rowNr As Long)

    Dim ea As String
    ea = Trim(ws.Cells(rowNr, COL_EIN_AUS).value)

    Dim targetCell As Range
    Set targetCell = ws.Cells(rowNr, COL_ZIELSPALTE)

    ' Immer zuerst löschen
    targetCell.Validation.Delete

    If ea = "" Then Exit Sub

    Dim listFormula As String
    listFormula = GetZielspaltenListe(ea)

    If listFormula = "" Then Exit Sub

    With targetCell.Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:=listFormula
        .InCellDropdown = True
    End With

End Sub

' ---------------------------------------------------------------
' Ermittelt passende Überschriftenliste
' ---------------------------------------------------------------
Private Function GetZielspaltenListe(ByVal einAus As String) As String

    Dim wsBK As Worksheet
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)

    Select Case UCase(einAus)

        Case "E"   ' Einnahmen
            GetZielspaltenListe = _
                "='" & WS_BANKKONTO & "'!" & _
                wsBK.Range("M27:S27").Address

        Case "A"   ' Ausgaben
            GetZielspaltenListe = _
                "='" & WS_BANKKONTO & "'!" & _
                wsBK.Range("T27:Z27").Address

        Case Else
            GetZielspaltenListe = ""

    End Select

End Function




