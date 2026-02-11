Attribute VB_Name = "mod_KategorieRegeln"
Option Explicit

' ***************************************************************
' MODUL: mod_KategorieRegeln
' ZWECK: Kategorieverwaltung + Konsistenzpruefung
' AKTUALISIERT: Spalte O (Guthabenfaehig) wurde entfernt
' ***************************************************************

Public Sub Initialisiere_Kategorie_Regeln()

    Dim ws As Worksheet
    Const FIRST_DATA_ROW As Long = 4
    Const MAX_CATEGORY_ROWS As Long = 1000

    Set ws = ThisWorkbook.Worksheets("Daten")

    ' Auswahlfelder auf festen Bereich anwenden
    ' AKTUALISIERT: Spalte O ist jetzt Faelligkeit (war vorher Guthabenfaehig)
    Call SetListValidationRange(ws.Range("K" & FIRST_DATA_ROW & ":K" & FIRST_DATA_ROW + MAX_CATEGORY_ROWS), "lst_EinnahmeAusgabe")
    Call SetListValidationRange(ws.Range("M" & FIRST_DATA_ROW & ":M" & FIRST_DATA_ROW + MAX_CATEGORY_ROWS), "lst_Prioritaet")
    Call SetListValidationRange(ws.Range("O" & FIRST_DATA_ROW & ":O" & FIRST_DATA_ROW + MAX_CATEGORY_ROWS), "lst_Faelligkeit")

End Sub

Private Sub SetListValidationRange(targetRange As Range, listName As String)
    On Error Resume Next
    With targetRange.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=" & listName
        .InCellDropdown = True
    End With
    On Error GoTo 0
End Sub

' =====================================================
' Konsistenzpruefung bei gleichen Kategorien
' =====================================================
Public Sub PruefeUndSynchronisiere_Kategorie(ByVal ws As Worksheet, ByVal changedRow As Long)

    Dim katName As String
    katName = Trim(ws.Cells(changedRow, DATA_CAT_COL_KATEGORIE).value)
    If katName = "" Then Exit Sub

    Dim refRow As Long
    refRow = FindeErsteKategorieZeile(ws, katName, changedRow)

    If refRow = 0 Then Exit Sub

    Application.EnableEvents = False

    ' Referenzwerte uebernehmen (AKTUALISIERT - ohne Guthabenfaehig)
    ws.Cells(changedRow, DATA_CAT_COL_EINAUS).value = ws.Cells(refRow, DATA_CAT_COL_EINAUS).value
    ws.Cells(changedRow, DATA_CAT_COL_ZIELSPALTE).value = ws.Cells(refRow, DATA_CAT_COL_ZIELSPALTE).value
    ws.Cells(changedRow, DATA_CAT_COL_FAELLIGKEIT).value = ws.Cells(refRow, DATA_CAT_COL_FAELLIGKEIT).value

    MsgBox _
        "Die Kategorie '" & katName & "' existiert bereits." & vbCrLf & vbCrLf & _
        "Einnahme/Ausgabe, Zielspalte und Faelligkeit " & _
        "wurden automatisch uebernommen, um Inkonsistenzen zu vermeiden.", _
        vbInformation, "Kategorie vereinheitlicht"

    Application.EnableEvents = True

End Sub

Private Function FindeErsteKategorieZeile(ws As Worksheet, _
                                         katName As String, _
                                         excludeRow As Long) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        If r <> excludeRow Then
            If Trim(ws.Cells(r, DATA_CAT_COL_KATEGORIE).value) = katName Then
                FindeErsteKategorieZeile = r
                Exit Function
            End If
        End If
    Next r

    FindeErsteKategorieZeile = 0
End Function



