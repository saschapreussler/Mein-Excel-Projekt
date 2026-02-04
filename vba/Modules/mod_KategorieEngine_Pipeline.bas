Attribute VB_Name = "mod_KategorieEngine_Pipeline"
Option Explicit

' ===============================================================
' KATEGORIEENGINE PIPELINE
' ===============================================================
Public Sub KategorieEngine_Pipeline(Optional ByVal wsBK As Worksheet)

    Dim wsData As Worksheet
    Dim rngRules As Range
    Dim lastRowBK As Long
    Dim r As Long

    If wsBK Is Nothing Then Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsData = ThisWorkbook.Worksheets(WS_DATEN)
    Set rngRules = wsData.Range(RANGE_KATEGORIE_REGELN)

    lastRowBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRowBK < BK_START_ROW Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For r = BK_START_ROW To lastRowBK

        Dim normText As String
        normText = NormalizeBankkontoZeile(wsBK, r)
        If normText = "" Then GoTo NextRow

        ' ------------------------------
        ' Kategorie ermitteln
        ' ------------------------------
        EvaluateKategorieEngineRow wsBK, r, rngRules

        ' ------------------------------
        ' Betrag sofort zuordnen
        ' (nur wenn Kategorie GRÜN)
        ' ------------------------------
        ApplyBetragsZuordnung wsBK, r

NextRow:
    Next r

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub




