Attribute VB_Name = "mod_Uebersicht_Filter"
Option Explicit

' ***************************************************************
' MODUL: mod_Uebersicht_Filter
' VERSION: 1.1 - 15.03.2026
' ZWECK: Monats-Register (Shapes) auf dem Blatt "Uebersicht"
'        Erstellt 13 Registerreiter-Shapes ("Alle" + 12 Monate)
'        Klick auf Shape -> AutoFilter auf Spalte C (Monat)
'        v1.1: Zebra auf sichtbare Zeilen nach Filter anwenden
'
' FUNKTIONEN:
'   - ErstelleMonatsRegister: Shapes erzeugen/aktualisieren
'   - FilterUebersichtNachMonat: AutoFilter anwenden
'   - EntferneMonatsRegister: Shapes loeschen
' ***************************************************************

' Farben fuer Register-Tabs
Private Const REG_FARBE_AKTIV As Long = 10053120     ' RGB(0, 112, 153) - aktiv
Private Const REG_FARBE_INAKTIV As Long = 14408667    ' RGB(219, 229, 219) - inaktiv
Private Const REG_SCHRIFT_AKTIV As Long = 16777215    ' Weiss
Private Const REG_SCHRIFT_INAKTIV As Long = 0          ' Schwarz
Private Const REG_SHAPE_PREFIX As String = "regMonat_"

' Zeile/Spalte fuer Shape-Platzierung
Private Const REG_TOP_ROW As Long = 1    ' Zeile 1
Private Const REG_HEIGHT As Double = 22
Private Const REG_WIDTH As Double = 72
Private Const REG_SPACING As Double = 2
Private Const REG_LEFT_START As Double = 5

' v1.1: Konstanten fuer Zebra-Reapply (identisch mit mod_Uebersicht_Generator)
Private Const ZEBRA_COLOR As Long = &HDEE5E3
Private Const FARBE_HELLGELB_MANUELL As Long = 10092543
Private Const UEB_COL_PARZELLE As Long = 1
Private Const UEB_COL_SOLL As Long = 5
Private Const UEB_COL_STATUS As Long = 7
Private Const UEB_COL_BEMERKUNG As Long = 8
Private Const UEB_COL_SUMME_IST As Long = 9
Private Const FARBE_SUMME As Long = 16247773
Private Const FARBE_SUMME_ZEBRA As Long = 15790320
Private Const UEBERSICHT_START_ROW As Long = 4


' ===============================================================
' Erstellt oder aktualisiert die 13 Monats-Register-Shapes
' Wird am Ende von GeneriereUebersicht aufgerufen
' ===============================================================
Public Sub ErstelleMonatsRegister()

    Dim wsUeb As Worksheet
    On Error Resume Next
    Set wsUeb = ThisWorkbook.Worksheets(WS_UEBERSICHT())
    On Error GoTo 0
    If wsUeb Is Nothing Then Exit Sub

    ' Blattschutz temporaer entfernen
    On Error Resume Next
    wsUeb.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0

    ' Alte Register-Shapes entfernen
    Dim shp As Shape
    Dim delNames As New Collection
    For Each shp In wsUeb.Shapes
        If Left(shp.Name, Len(REG_SHAPE_PREFIX)) = REG_SHAPE_PREFIX Then
            delNames.Add shp.Name
        End If
    Next shp
    Dim dName As Variant
    For Each dName In delNames
        wsUeb.Shapes(CStr(dName)).Delete
    Next dName

    ' Register-Beschriftungen: "Alle" + 12 Monate
    Dim labels(0 To 12) As String
    labels(0) = "Alle"
    Dim m As Long
    For m = 1 To 12
        labels(m) = MonthName(m)
    Next m

    ' Shapes erstellen
    Dim leftPos As Double
    leftPos = REG_LEFT_START
    Dim topPos As Double
    topPos = wsUeb.Rows(REG_TOP_ROW).Top + 2

    For m = 0 To 12
        Dim newShp As Shape
        Set newShp = wsUeb.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            leftPos, topPos, REG_WIDTH, REG_HEIGHT)

        newShp.Name = REG_SHAPE_PREFIX & m

        ' Aussehen
        With newShp
            .TextFrame2.TextRange.text = labels(m)
            .TextFrame2.TextRange.Font.Size = 9
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .TextFrame2.WordWrap = msoFalse

            ' Standard: "Alle" ist aktiv
            If m = 0 Then
                .Fill.ForeColor.RGB = REG_FARBE_AKTIV
                .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = REG_SCHRIFT_AKTIV
            Else
                .Fill.ForeColor.RGB = REG_FARBE_INAKTIV
                .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = REG_SCHRIFT_INAKTIV
            End If

            .Line.Visible = msoFalse
            .Shadow.Visible = msoFalse

            ' Makro zuweisen
            .OnAction = "'mod_Uebersicht_Filter.FilterUebersichtNachMonat " & m & "'"

            ' Shape soll nicht druckbar sein und nicht verschoben werden
            .Placement = xlFreeFloating
        End With

        leftPos = leftPos + REG_WIDTH + REG_SPACING
    Next m

    ' Blattschutz wieder aktivieren
    On Error Resume Next
    wsUeb.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0

End Sub


' ===============================================================
' Filtert die Uebersicht nach Monat (0 = Alle anzeigen)
' Wird per Shape-OnAction aufgerufen
' ===============================================================
Public Sub FilterUebersichtNachMonat(ByVal monatIndex As Long)

    Dim wsUeb As Worksheet
    On Error Resume Next
    Set wsUeb = ThisWorkbook.Worksheets(WS_UEBERSICHT())
    On Error GoTo 0
    If wsUeb Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Blattschutz temporaer entfernen
    On Error Resume Next
    wsUeb.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0

    ' Bestehenden AutoFilter zuruecksetzen
    If wsUeb.AutoFilterMode Then wsUeb.AutoFilterMode = False

    ' Letzte Datenzeile ermitteln
    Dim lastRow As Long
    lastRow = wsUeb.Cells(wsUeb.Rows.count, 1).End(xlUp).Row
    If lastRow < 4 Then lastRow = 4

    If monatIndex = 0 Then
        ' "Alle" -> Filter-Kriterien entfernen, AutoFilter-Dropdowns behalten
        wsUeb.Range("A3:H" & lastRow).AutoFilter
    Else
        ' Filter auf Spalte C (Monat) anwenden
        ' Filterkriterium: "*MonatName*" (enthaelt den Monatsnamen)
        Dim filterMonat As String
        filterMonat = MonthName(monatIndex)

        wsUeb.Range("A3:H" & lastRow).AutoFilter _
            Field:=3, _
            Criteria1:="=*" & filterMonat & "*", _
            Operator:=xlAnd
    End If

    ' Register-Farben aktualisieren
    Dim shp As Shape
    Dim shpMonat As Long
    For Each shp In wsUeb.Shapes
        If Left(shp.Name, Len(REG_SHAPE_PREFIX)) = REG_SHAPE_PREFIX Then
            shpMonat = CLng(Mid(shp.Name, Len(REG_SHAPE_PREFIX) + 1))

            If shpMonat = monatIndex Then
                ' Aktiver Tab
                shp.Fill.ForeColor.RGB = REG_FARBE_AKTIV
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = REG_SCHRIFT_AKTIV
            Else
                ' Inaktiver Tab
                shp.Fill.ForeColor.RGB = REG_FARBE_INAKTIV
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = REG_SCHRIFT_INAKTIV
            End If
        End If
    Next shp

    ' v1.1: Zebra auf sichtbare Zeilen neu anwenden
    Call WendeZebraAufSichtbareZeilenAn(wsUeb, UEBERSICHT_START_ROW, lastRow)

    ' Blattschutz wieder aktivieren
    On Error Resume Next
    wsUeb.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0

    Application.ScreenUpdating = True

End Sub


' ===============================================================
' v1.1: Zebra-Muster basierend auf SICHTBAREN Zeilen anwenden
' Zaehlt nur sichtbare Zeilen fuer das Mod-2-Muster.
' Ueberspringt Status-Spalte (Ampel) und gelbe Soll-Zellen.
' ===============================================================
Private Sub WendeZebraAufSichtbareZeilenAn(ByVal wsUeb As Worksheet, _
                                            ByVal startRow As Long, _
                                            ByVal endRow As Long)
    
    If endRow < startRow Then Exit Sub
    
    Dim r As Long
    Dim visibleIdx As Long
    Dim c As Long
    
    visibleIdx = 0
    
    For r = startRow To endRow
        ' Nur sichtbare Zeilen beruecksichtigen
        If wsUeb.Rows(r).Hidden = False Then
            If visibleIdx Mod 2 = 1 Then
                ' Gerade sichtbare Zeile -> Zebra-Farbe
                For c = UEB_COL_PARZELLE To UEB_COL_SUMME_IST
                    If c = UEB_COL_STATUS Then
                        ' Status-Spalte behaelt Ampelfarbe
                    ElseIf c = UEB_COL_SOLL Then
                        If wsUeb.Cells(r, c).Interior.color <> FARBE_HELLGELB_MANUELL And _
                           wsUeb.Cells(r, c).Interior.color <> RGB(196, 225, 196) Then
                            wsUeb.Cells(r, c).Interior.color = ZEBRA_COLOR
                        End If
                    ElseIf c = UEB_COL_SUMME_IST Then
                        wsUeb.Cells(r, c).Interior.color = FARBE_SUMME_ZEBRA
                    Else
                        wsUeb.Cells(r, c).Interior.color = ZEBRA_COLOR
                    End If
                Next c
            Else
                ' Ungerade sichtbare Zeile -> weiss
                For c = UEB_COL_PARZELLE To UEB_COL_SUMME_IST
                    If c = UEB_COL_STATUS Then
                        ' Status-Spalte behaelt Ampelfarbe
                    ElseIf c = UEB_COL_SOLL Then
                        If wsUeb.Cells(r, c).Interior.color <> FARBE_HELLGELB_MANUELL And _
                           wsUeb.Cells(r, c).Interior.color <> RGB(196, 225, 196) Then
                            wsUeb.Cells(r, c).Interior.ColorIndex = xlNone
                        End If
                    ElseIf c = UEB_COL_SUMME_IST Then
                        wsUeb.Cells(r, c).Interior.color = FARBE_SUMME
                    Else
                        wsUeb.Cells(r, c).Interior.ColorIndex = xlNone
                    End If
                Next c
            End If
            
            visibleIdx = visibleIdx + 1
        End If
    Next r
    
End Sub


' ===============================================================
' Entfernt alle Monats-Register-Shapes (fuer Reset)
' ===============================================================
Public Sub EntferneMonatsRegister()

    Dim wsUeb As Worksheet
    On Error Resume Next
    Set wsUeb = ThisWorkbook.Worksheets(WS_UEBERSICHT())
    On Error GoTo 0
    If wsUeb Is Nothing Then Exit Sub

    On Error Resume Next
    wsUeb.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0

    ' AutoFilter zuruecksetzen
    If wsUeb.AutoFilterMode Then wsUeb.AutoFilterMode = False

    ' Shapes loeschen
    Dim shp As Shape
    Dim delNames As New Collection
    For Each shp In wsUeb.Shapes
        If Left(shp.Name, Len(REG_SHAPE_PREFIX)) = REG_SHAPE_PREFIX Then
            delNames.Add shp.Name
        End If
    Next shp
    Dim dName As Variant
    For Each dName In delNames
        wsUeb.Shapes(CStr(dName)).Delete
    Next dName

    On Error Resume Next
    wsUeb.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0

End Sub


































































