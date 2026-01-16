Attribute VB_Name = "mod_KategorieRegeln"
Option Explicit

' ***************************************************************
' MODUL: mod_KategorieRegeln
' ZWECK: Kategorieverwaltung + Konsistenzprüfung
' ***************************************************************

Public Sub Initialisiere_Kategorie_Regeln()

    Dim ws As Worksheet
    Const FIRST_DATA_ROW As Long = 4
    Const MAX_CATEGORY_ROWS As Long = 1000

    Set ws = ThisWorkbook.Worksheets("Daten")

    ' Auswahlfelder auf festen Bereich anwenden
    Call SetListValidationRange(ws.Range("K" & FIRST_DATA_ROW & ":K" & FIRST_DATA_ROW + MAX_CATEGORY_ROWS), "lst_EinnahmeAusgabe")
    Call SetListValidationRange(ws.Range("M" & FIRST_DATA_ROW & ":M" & FIRST_DATA_ROW + MAX_CATEGORY_ROWS), "lst_Prioritaet")
    Call SetListValidationRange(ws.Range("O" & FIRST_DATA_ROW & ":O" & FIRST_DATA_ROW + MAX_CATEGORY_ROWS), "lst_JaNein")
    Call SetListValidationRange(ws.Range("P" & FIRST_DATA_ROW & ":P" & FIRST_DATA_ROW + MAX_CATEGORY_ROWS), "lst_Faelligkeit")

End Sub

Private Sub SetListValidationRange(targetRange As Range, listName As String)
    With targetRange.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=" & listName
        .InCellDropdown = True
    End With
End Sub

' =====================================================
' Konsistenzprüfung und automatische Vorbelegung bei gleichen Kategorien
' =====================================================
Public Sub PruefeUndSynchronisiere_Kategorie(ByVal ws As Worksheet, ByVal changedRow As Long)

    Dim katName As String
    katName = Trim(ws.Cells(changedRow, "J").Value)
    If katName = "" Then Exit Sub

    Dim refRow As Long
    refRow = FindeErsteKategorieZeile(ws, katName, changedRow)

    If refRow = 0 Then Exit Sub ' erste Kategorie ? nichts zu prüfen

    Application.EnableEvents = False

    ' Referenzwerte übernehmen
    ws.Cells(changedRow, "K").Value = ws.Cells(refRow, "K").Value
    ws.Cells(changedRow, "N").Value = ws.Cells(refRow, "N").Value
    ws.Cells(changedRow, "O").Value = ws.Cells(refRow, "O").Value
    ws.Cells(changedRow, "P").Value = ws.Cells(refRow, "P").Value

    ' Hinweis für den Nutzer
    MsgBox _
        "Die Kategorie '" & katName & "' existiert bereits." & vbCrLf & vbCrLf & _
        "Einnahme/Ausgabe, Zielspalte, Guthabenfähigkeit und Fälligkeit " & _
        "wurden automatisch übernommen, um Inkonsistenzen zu vermeiden.", _
        vbInformation, "Kategorie vereinheitlicht"

    Application.EnableEvents = True

End Sub

' ---------------------------------------------------------------
' Erste Vorkommnis-Zeile finden
' ---------------------------------------------------------------
Private Function FindeErsteKategorieZeile(ws As Worksheet, _
                                         katName As String, _
                                         excludeRow As Long) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

    Dim r As Long
    For r = 4 To lastRow
        If r <> excludeRow Then
            If Trim(ws.Cells(r, "J").Value) = katName Then
                FindeErsteKategorieZeile = r
                Exit Function
            End If
        End If
    Next r

    FindeErsteKategorieZeile = 0
End Function




