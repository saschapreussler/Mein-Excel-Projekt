Attribute VB_Name = "mod_Testdaten"
Option Explicit

' ===============================================================
' mod_Testdaten
' ---------------------------------------------------------------
' Setzt die Arbeitsmappe auf den Zustand vor dem ersten CSV-Import
' zurueck, ohne die Einrichtung zu verlieren.
'
' Zweck ist das Durchspielen des gesamten Ablaufs: Import, Erkennung
' der Kategorien, Perioden, Rueckfragen und Auswertungen. Bisher
' musste dafuer von Hand geloescht werden, und dabei blieben die
' unsichtbaren Speicher auf dem Blatt Daten stehen. Genau die sind
' der Grund, warum nach einem solchen Handgriff alte Entscheidungen
' und Bestaetigungen wieder auftauchen, obwohl keine Buchung mehr da
' ist.
'
' Geleert werden:
'   - Bankkonto, alle Buchungszeilen
'   - Vereinskasse, alle Buchungszeilen
'   - Zahlungsuebersicht
'   - Dashboard Mitgliederzahlungen
'   - Finanz-Uebersicht
'   - Zuordnungstabelle auf dem Blatt Daten (Spalten R bis X)
'   - saemtliche Entscheidungsspeicher auf dem Blatt Daten (CA bis CX):
'     Vorjahresbuchungen, manuelle Vorjahrentscheidungen,
'     Guthabenverrechnungen, Saeumnisbestaetigungen, Sonderzuordnungen
'
' Erhalten bleiben:
'   - Mitgliederliste und Mitgliederhistorie
'   - Kategorietabelle und Hilfslisten auf dem Blatt Daten
'   - Blatt Einstellungen mit Zahlungsterminen
'   - Strom, Wasser und Zaehlerhistorie
' ===============================================================


' ===============================================================
' Bedieneinstieg. Fragt zweimal nach, bevor irgendetwas passiert.
' ===============================================================
Public Sub LeereTestdaten()

    Dim antwort As VbMsgBoxResult
    Dim bestaetigung As String
    Dim bericht As String
    Dim anzahl As Long

    antwort = MsgBox( _
        "Testlauf vorbereiten" & vbCrLf & vbCrLf & _
        "Diese Funktion setzt die Mappe auf den Zustand vor dem ersten" & vbCrLf & _
        "CSV-Import zur" & ChrW(252) & "ck. Sie ist zum Ausprobieren gedacht, nicht" & vbCrLf & _
        "f" & ChrW(252) & "r den laufenden Betrieb." & vbCrLf & vbCrLf & _
        "GEL" & ChrW(214) & "SCHT werden:" & vbCrLf & _
        "  - alle Buchungen auf Bankkonto und Vereinskasse" & vbCrLf & _
        "  - Zahlungs" & ChrW(252) & "bersicht, Dashboard und Finanz-" & ChrW(220) & "bersicht" & vbCrLf & _
        "  - die Zuordnungstabelle auf dem Blatt Daten" & vbCrLf & _
        "  - alle gespeicherten Entscheidungen, Best" & ChrW(228) & "tigungen" & vbCrLf & _
        "    und Guthabenverrechnungen" & vbCrLf & vbCrLf & _
        "ERHALTEN bleiben:" & vbCrLf & _
        "  - Mitgliederliste und Mitgliederhistorie" & vbCrLf & _
        "  - Kategorien und Zahlungstermine" & vbCrLf & _
        "  - Z" & ChrW(228) & "hlerst" & ChrW(228) & "nde und Einstellungen" & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbExclamation + vbYesNo + vbDefaultButton2, "Testlauf vorbereiten")

    If antwort <> vbYes Then Exit Sub

    bestaetigung = InputBox( _
        "Sicherheitsabfrage." & vbCrLf & vbCrLf & _
        "Bitte das Wort LEEREN eintippen, um das L" & ChrW(246) & "schen auszul" & ChrW(246) & "sen." & vbCrLf & vbCrLf & _
        "Legen Sie vorher eine Sicherungskopie der Mappe an.", _
        "Testlauf vorbereiten")

    If StrComp(Trim$(bestaetigung), "LEEREN", vbTextCompare) <> 0 Then
        MsgBox "Abgebrochen. Es wurde nichts ge" & ChrW(228) & "ndert.", _
               vbInformation, "Testlauf vorbereiten"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Fehler

    anzahl = LoescheBuchungszeilen(WS_BANKKONTO, BK_START_ROW)
    bericht = "Bankkonto: " & anzahl & " Zeilen"

    anzahl = LoescheBuchungszeilen(WS_VEREINSKASSE, VK_START_ROW)
    bericht = bericht & vbCrLf & "Vereinskasse: " & anzahl & " Zeilen"

    anzahl = LoescheBuchungszeilen(WS_UEBERSICHT(), UEBERSICHT_START_ROW)
    bericht = bericht & vbCrLf & "Zahlungs" & ChrW(252) & "bersicht: " & anzahl & " Zeilen"

    anzahl = LoescheBuchungszeilen("Dashboard Mitgliederzahlungen", DASH_MATRIX_START_ROW)
    bericht = bericht & vbCrLf & "Dashboard: " & anzahl & " Zeilen"

    Call LeereFinanzUebersichtTest
    bericht = bericht & vbCrLf & "Finanz-" & ChrW(220) & "bersicht: geleert"

    anzahl = LeereDatenBloecke()
    bericht = bericht & vbCrLf & "Blatt Daten: " & anzahl & " Speicherzeilen"

    Call EntladeAlleZwischenspeicher

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Die Mappe ist zur" & ChrW(252) & "ckgesetzt." & vbCrLf & vbCrLf & _
           bericht & vbCrLf & vbCrLf & _
           "Sie k" & ChrW(246) & "nnen jetzt mit dem CSV-Import beginnen." & vbCrLf & _
           "Bitte die Mappe speichern, wenn der Stand erhalten bleiben soll.", _
           vbInformation, "Testlauf vorbereiten"

    Exit Sub

Fehler:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Zur" & ChrW(252) & "cksetzen:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Testlauf vorbereiten"

End Sub


' ===============================================================
' Loescht auf einem Blatt alle Zeilen ab der ersten Datenzeile.
'
' Bewusst ganze Zeilen und nicht nur die Inhalte: Die Ampelfarben der
' Kategoriespalte und die Hervorhebungen der Uebersicht haengen an der
' Zellformatierung. Wuerden nur die Inhalte geloescht, bliebe ein
' buntes Gerippe stehen.
' ===============================================================
Private Function LoescheBuchungszeilen(ByVal blattName As String, _
                                       ByVal startZeile As Long) As Long

    Dim ws As Worksheet
    Dim letzteZeile As Long
    Dim warGeschuetzt As Boolean

    LoescheBuchungszeilen = 0

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(blattName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    warGeschuetzt = ws.ProtectContents
    On Error Resume Next
    If warGeschuetzt Then ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0

    letzteZeile = ws.UsedRange.Row + ws.UsedRange.Rows.count - 1
    If letzteZeile >= startZeile Then
        ws.Range(ws.Rows(startZeile), ws.Rows(letzteZeile)).Delete
        LoescheBuchungszeilen = letzteZeile - startZeile + 1
    End If

    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    If warGeschuetzt Then
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    On Error GoTo 0

End Function


' ===============================================================
' Leert die Finanz-Uebersicht vollstaendig. Sie wird bei jedem Aufruf
' ohnehin neu erzeugt und traegt keine Eingaben des Nutzers.
' ===============================================================
Private Sub LeereFinanzUebersichtTest()

    Dim ws As Worksheet
    Dim warGeschuetzt As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_FINANZ_UEBERSICHT())
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    warGeschuetzt = ws.ProtectContents
    On Error Resume Next
    If warGeschuetzt Then ws.Unprotect PASSWORD:=PASSWORD
    ws.UsedRange.ClearContents
    If warGeschuetzt Then
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    On Error GoTo 0

End Sub


' ===============================================================
' Leert die Zuordnungstabelle und alle Entscheidungsspeicher auf dem
' Blatt Daten.
'
' Die Kategorietabelle (Spalten J bis P) und die Hilfslisten bleiben
' unberuehrt, denn sie gehoeren zur Einrichtung und nicht zu den
' Bewegungsdaten. Geleert wird der zusammenhaengende Block von der
' Zuordnungstabelle sowie der Bereich CA bis CX, in dem das Programm
' Vorjahresbuchungen, manuelle Entscheidungen, Guthabenverrechnungen,
' Saeumnisbestaetigungen und Sonderzuordnungen ablegt.
' ===============================================================
Private Function LeereDatenBloecke() As Long

    Dim ws As Worksheet
    Dim warGeschuetzt As Boolean
    Dim letzteZeile As Long
    Dim gesamt As Long

    LeereDatenBloecke = 0

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    warGeschuetzt = ws.ProtectContents
    On Error Resume Next
    If warGeschuetzt Then ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0

    ' Zuordnungstabelle R bis X
    letzteZeile = ws.Cells(ws.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    If letzteZeile >= EK_START_ROW Then
        ws.Range(ws.Cells(EK_START_ROW, EK_COL_ENTITYKEY), _
                 ws.Cells(letzteZeile, EK_COL_DEBUG)).ClearContents
        gesamt = gesamt + (letzteZeile - EK_START_ROW + 1)
    End If

    ' Entscheidungsspeicher CA bis CX. Der Block wird in einem Zug
    ' geleert, damit kein Teilspeicher uebersehen wird, wenn spaeter
    ' eine weitere Spalte dazukommt.
    letzteZeile = ErmittleLetzteSpeicherzeile(ws)
    If letzteZeile >= VJ_START_ROW Then
        ws.Range(ws.Cells(VJ_START_ROW, VJ_COL_DATUM), _
                 ws.Cells(letzteZeile, SZ_COL_ERFASST)).ClearContents
        gesamt = gesamt + (letzteZeile - VJ_START_ROW + 1)
    End If

    On Error Resume Next
    If warGeschuetzt Then
        ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    On Error GoTo 0

    LeereDatenBloecke = gesamt

End Function


' ===============================================================
' Sucht die unterste belegte Zeile ueber alle Speicherspalten hinweg.
' ===============================================================
Private Function ErmittleLetzteSpeicherzeile(ByVal ws As Worksheet) As Long

    Dim spalte As Long
    Dim zeile As Long

    ErmittleLetzteSpeicherzeile = 0

    For spalte = VJ_COL_DATUM To SZ_COL_ERFASST
        zeile = ws.Cells(ws.Rows.count, spalte).End(xlUp).Row
        If zeile > ErmittleLetzteSpeicherzeile Then ErmittleLetzteSpeicherzeile = zeile
    Next spalte

End Function


' ===============================================================
' Wirft alle Zwischenspeicher weg.
'
' Ohne diesen Schritt wuerde der naechste Lauf noch mit den Werten
' rechnen, die vor dem Leeren im Arbeitsspeicher lagen.
' ===============================================================
Private Sub EntladeAlleZwischenspeicher()

    On Error Resume Next
    Call mod_Zahlungspruefung.EntladeEinstellungenCacheZP
    Call mod_Sonderzuordnung.EntladeSonderzuordnungsCache
    On Error GoTo 0

End Sub


' ===============================================================
' Legt die Schaltflaeche fuer den Testlauf auf dem Blatt Einstellungen
' an, deutlich unterhalb der Zahlungstermine.
'
' Bewusst nicht auf der Startseite und nicht neben dem Import: Die
' Funktion loescht Daten und soll nicht versehentlich getroffen
' werden. Das Blatt Einstellungen ist der Ort, an dem ohnehin nur
' bewusst gearbeitet wird.
' ===============================================================
Public Sub ErstelleTestdatenButton(Optional ByVal wsEinst As Worksheet = Nothing)

    Dim shp As Shape
    Dim warGeschuetzt As Boolean
    Dim anker As Range

    On Error GoTo Aufraeumen

    If wsEinst Is Nothing Then
        Set wsEinst = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    End If
    If wsEinst Is Nothing Then Exit Sub

    warGeschuetzt = wsEinst.ProtectContents
    If warGeschuetzt Then wsEinst.Unprotect PASSWORD:=PASSWORD

    Set anker = wsEinst.Range("B34")

    On Error Resume Next
    wsEinst.Shapes("btn_Testlauf").Delete
    On Error GoTo Aufraeumen

    wsEinst.Cells(32, 2).value = "Testbetrieb"
    wsEinst.Cells(32, 2).Font.Bold = True
    wsEinst.Cells(33, 2).value = "Setzt Buchungen und Auswertungen zur" & ChrW(252) & _
                                 "ck. Mitglieder, Kategorien, Zahlungstermine und " & _
                                 "Z" & ChrW(228) & "hlerst" & ChrW(228) & "nde bleiben erhalten."
    wsEinst.Cells(33, 2).Font.Italic = True

    Set shp = wsEinst.Shapes.AddShape(msoShapeRoundedRectangle, _
              anker.Left, anker.Top + 2, 210, 26)
    With shp
        .Name = "btn_Testlauf"
        .TextFrame2.TextRange.text = ChrW(9888) & "   Testlauf vorbereiten"
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Fill.ForeColor.RGB = RGB(163, 48, 44)
        .Line.Visible = msoFalse
        .OnAction = "'mod_Testdaten.LeereTestdaten'"
        .Placement = xlFreeFloating
    End With

Aufraeumen:
    On Error Resume Next
    If warGeschuetzt And Not wsEinst Is Nothing Then
        wsEinst.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    On Error GoTo 0

End Sub
