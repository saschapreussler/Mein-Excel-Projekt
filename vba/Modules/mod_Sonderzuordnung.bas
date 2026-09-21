Attribute VB_Name = "mod_Sonderzuordnung"
Option Explicit

' ***************************************************************
' MODUL: mod_Sonderzuordnung
' VERSION: 1.0 - 21.09.2026
' ZWECK: Ordnet einzelne Buchungen einem anderen Mitglied zu als
'        dem, zu dem die IBAN gehört.
'
' Hintergrund:
' Die Zahlungsprüfung erkennt normalerweise über die IBAN, wem eine
' Zahlung gutzuschreiben ist. Das genügt nicht, wenn ein Dritter
' zahlt. Zwei Fälle kommen in der Praxis vor:
'
'   1. Ein Mitglied überweist ausnahmsweise für ein Mitglied einer
'      anderen Parzelle. Der Verwendungszweck verrät den Grund,
'      die Automatik kann ihn aber nicht zuverlässig deuten.
'   2. Ein Mitglied verstirbt. Endabrechnung und Betriebskosten-
'      abrechnung müssen trotzdem bezahlt werden. Die Angehörigen
'      zahlen von ihrem eigenen Konto unter ihrem eigenen Namen.
'
' Lösung:
' Der Nutzer ordnet die betroffene Buchung einmalig von Hand einem
' Mitglied zu. Die Zuordnung wird dauerhaft im Blatt Daten abgelegt
' und beim Aufbau des Bankkonto-Zwischenspeichers ausgewertet. Die
' Buchung zählt dann für das Zielmitglied statt für den Zahler.
'
' Schlüssel einer Buchung:
' Datum, Betrag, IBAN und Verwendungszweck. Das ist derselbe
' Schlüssel, mit dem der CSV-Import Dubletten erkennt. Die interne
' Nummer in Spalte J taugt nicht, weil sie bei jedem Lauf neu
' vergeben wird.
'
' Speicherort: Blatt Daten, Spalten CT bis CX (siehe mod_Const).
' ***************************************************************

' Zwischenspeicher, damit der Bankkonto-Cache nicht für jede Zeile
' erneut das Blatt Daten liest.
Private m_Cache As Object
Private m_CacheGeladen As Boolean


' ===============================================================
' Bildet den Schlüssel einer Buchung.
'
' Identisch zum Dublettenschlüssel des CSV-Imports in
' mod_Banking_Data, damit beide Seiten dieselbe Buchung meinen.
' ===============================================================
Public Function BuchungsSchluessel(ByVal datum As Variant, _
                                   ByVal betrag As Variant, _
                                   ByVal iban As Variant, _
                                   ByVal verwendungszweck As Variant) As String

    Dim dDatum As Date
    Dim dBetrag As Double

    BuchungsSchluessel = ""
    If Not IsDate(datum) Then Exit Function
    dDatum = CDate(datum)

    If IsNumeric(betrag) Then dBetrag = CDbl(betrag)

    BuchungsSchluessel = Format$(dDatum, "YYYYMMDD") & "|" & dBetrag & "|" & _
                         Replace(Trim$(CStr(iban)), " ", "") & "|" & _
                         Trim$(CStr(verwendungszweck))

End Function


' ===============================================================
' Liefert den Ziel-Zuordnungsschlüssel einer Buchung.
'
' Rückgabe ist ein Leerstring, wenn für die Buchung keine
' Sonderzuordnung hinterlegt ist. Das ist der Normalfall.
' ===============================================================
Public Function ZielZuordnungsschluessel(ByVal schluessel As String) As String

    ZielZuordnungsschluessel = ""
    If schluessel = "" Then Exit Function

    If Not m_CacheGeladen Then Call LadeCache
    If m_Cache Is Nothing Then Exit Function
    If Not m_Cache.exists(schluessel) Then Exit Function

    ZielZuordnungsschluessel = CStr(m_Cache(schluessel)(0))

End Function


' ===============================================================
' Gibt an, ob überhaupt Sonderzuordnungen bestehen.
'
' Der Bankkonto-Cache kann sich damit die Schlüsselbildung je Zeile
' sparen, solange keine einzige Zuordnung hinterlegt ist.
' ===============================================================
Public Function HatSonderzuordnungen() As Boolean

    If Not m_CacheGeladen Then Call LadeCache
    HatSonderzuordnungen = False
    If m_Cache Is Nothing Then Exit Function
    HatSonderzuordnungen = (m_Cache.count > 0)

End Function


' ===============================================================
' Verwirft den Zwischenspeicher.
'
' Wird nach jeder Änderung und am Ende eines Generatorlaufs
' aufgerufen, damit kein veralteter Stand weiterwirkt.
' ===============================================================
Public Sub EntladeSonderzuordnungsCache()

    Set m_Cache = Nothing
    m_CacheGeladen = False

End Sub


Private Sub LadeCache()

    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim schluessel As String

    m_CacheGeladen = True
    Set m_Cache = CreateObject("Scripting.Dictionary")
    m_Cache.CompareMode = vbTextCompare

    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    If wsDaten Is Nothing Then Exit Sub

    lastRow = wsDaten.Cells(wsDaten.Rows.count, SZ_COL_KEY).End(xlUp).Row
    If lastRow < SZ_START_ROW Then Exit Sub

    For r = SZ_START_ROW To lastRow
        schluessel = Trim$(CStr(wsDaten.Cells(r, SZ_COL_KEY).value))
        If schluessel <> "" Then
            m_Cache(schluessel) = Array( _
                Trim$(CStr(wsDaten.Cells(r, SZ_COL_ZIEL_KEY).value)), _
                Trim$(CStr(wsDaten.Cells(r, SZ_COL_ZIEL_NAME).value)), _
                Trim$(CStr(wsDaten.Cells(r, SZ_COL_GRUND).value)))
        End If
    Next r

End Sub


' ===============================================================
' BEDIENUNG: Markierte Buchung einem Mitglied zuordnen.
'
' Aufruf über Entwicklertools, Makros. Vorher im Blatt Bankkonto
' die Zeile der Buchung markieren.
' ===============================================================
Public Sub OrdneBuchungEinemMitgliedZu()

    Dim wsBK As Worksheet
    Dim zeile As Long
    Dim titel As String
    Dim datum As Variant
    Dim betrag As Double
    Dim zahler As String
    Dim iban As String
    Dim zweck As String
    Dim schluessel As String
    Dim zielKey As String
    Dim zielName As String
    Dim grund As String
    Dim bisher As String

    titel = "Buchung einem Mitglied zuordnen"

    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo 0
    If wsBK Is Nothing Then Exit Sub

    zeile = ErmittleMarkierteBuchungszeile(wsBK, titel)
    If zeile = 0 Then Exit Sub

    datum = wsBK.Cells(zeile, BK_COL_DATUM).value
    If IsNumeric(wsBK.Cells(zeile, BK_COL_BETRAG).value) Then
        betrag = CDbl(wsBK.Cells(zeile, BK_COL_BETRAG).value)
    End If
    zahler = Trim$(CStr(wsBK.Cells(zeile, BK_COL_NAME).value))
    iban = Replace(Trim$(CStr(wsBK.Cells(zeile, BK_COL_IBAN).value)), " ", "")
    zweck = Trim$(CStr(wsBK.Cells(zeile, BK_COL_VERWENDUNGSZWECK).value))

    schluessel = BuchungsSchluessel(datum, betrag, iban, zweck)
    If schluessel = "" Then
        MsgBox "Diese Zeile hat kein gültiges Buchungsdatum.", vbExclamation, titel
        Exit Sub
    End If

    ' Bereits vorhandene Zuordnung anzeigen, damit der Nutzer weiß,
    ' dass er eine bestehende Angabe überschreibt.
    bisher = ""
    If Not m_CacheGeladen Then Call LadeCache
    If m_Cache.exists(schluessel) Then
        bisher = vbCrLf & "Bisher zugeordnet an: " & CStr(m_Cache(schluessel)(1)) & vbCrLf
    End If

    zielName = WaehleZielmitglied(zahler, zweck, betrag, datum, bisher, zielKey, titel)
    If zielName = "" Then Exit Sub

    grund = InputBox("Grund der Sonderzuordnung, zum Beispiel" & vbCrLf & _
                     "Erbfall oder Zahlung durch Nachbarn." & vbCrLf & vbCrLf & _
                     "Der Text dient nur der Nachvollziehbarkeit.", titel, _
                     "Zahlung durch Dritte")

    Call SpeichereSonderzuordnung(schluessel, zielKey, zielName, grund)
    Call VermerkeInBemerkung(wsBK, zeile, zielName)
    Call EntladeSonderzuordnungsCache

    MsgBox "Die Buchung zählt jetzt für " & zielName & "." & vbCrLf & vbCrLf & _
           "Betrag: " & Format$(betrag, "#,##0.00") & " " & ChrW(8364) & vbCrLf & _
           "Gezahlt von: " & zahler & vbCrLf & vbCrLf & _
           "Bitte die Zahlungsübersicht neu aufbauen, damit die" & vbCrLf & _
           "Zuordnung wirksam wird.", vbInformation, titel

End Sub


' ===============================================================
' BEDIENUNG: Bestehende Sonderzuordnungen ansehen und zurücknehmen.
' ===============================================================
Public Sub ZeigeSonderzuordnungen()

    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim anzahl As Long
    Dim liste As String
    Dim titel As String
    Dim eingabe As String
    Dim nummern As Object
    Dim zielRow As Long

    titel = "Sonderzuordnungen"

    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    If wsDaten Is Nothing Then Exit Sub

    Set nummern = CreateObject("Scripting.Dictionary")
    lastRow = wsDaten.Cells(wsDaten.Rows.count, SZ_COL_KEY).End(xlUp).Row

    For r = SZ_START_ROW To lastRow
        If Trim$(CStr(wsDaten.Cells(r, SZ_COL_KEY).value)) <> "" Then
            anzahl = anzahl + 1
            nummern(anzahl) = r
            liste = liste & vbCrLf & anzahl & ". " & _
                    Trim$(CStr(wsDaten.Cells(r, SZ_COL_ZIEL_NAME).value)) & _
                    "  (" & BeschreibeSchluessel(CStr(wsDaten.Cells(r, SZ_COL_KEY).value)) & ")"
            If Trim$(CStr(wsDaten.Cells(r, SZ_COL_GRUND).value)) <> "" Then
                liste = liste & vbCrLf & "     Grund: " & _
                        Trim$(CStr(wsDaten.Cells(r, SZ_COL_GRUND).value))
            End If
        End If
    Next r

    If anzahl = 0 Then
        MsgBox "Es ist keine Sonderzuordnung hinterlegt.", vbInformation, titel
        Exit Sub
    End If

    eingabe = InputBox("Hinterlegte Sonderzuordnungen:" & vbCrLf & liste & vbCrLf & vbCrLf & _
                       "Zum Zurücknehmen die Nummer eingeben." & vbCrLf & _
                       "Leer lassen und OK drücken, um nichts zu ändern.", titel)

    If Trim$(eingabe) = "" Then Exit Sub
    If Not IsNumeric(eingabe) Then Exit Sub
    If Not nummern.exists(CLng(eingabe)) Then
        MsgBox "Diese Nummer gibt es nicht.", vbExclamation, titel
        Exit Sub
    End If

    zielRow = nummern(CLng(eingabe))

    On Error Resume Next
    wsDaten.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0

    wsDaten.Range(wsDaten.Cells(zielRow, SZ_COL_KEY), _
                  wsDaten.Cells(zielRow, SZ_COL_ERFASST)).ClearContents

    On Error Resume Next
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0

    Call EntladeSonderzuordnungsCache

    MsgBox "Die Sonderzuordnung wurde zurückgenommen." & vbCrLf & _
           "Die Buchung zählt wieder für den Kontoinhaber." & vbCrLf & vbCrLf & _
           "Bitte die Zahlungsübersicht neu aufbauen.", vbInformation, titel

End Sub


' ===============================================================
' Ermittelt die markierte Buchungszeile und prüft, ob der Nutzer
' überhaupt im Blatt Bankkonto steht.
' ===============================================================
Private Function ErmittleMarkierteBuchungszeile(ByVal wsBK As Worksheet, _
                                                 ByVal titel As String) As Long

    Dim lastRow As Long
    Dim zeile As Long

    ErmittleMarkierteBuchungszeile = 0

    If Not TypeOf Selection Is Range Or Not ActiveSheet Is wsBK Then
        MsgBox "Bitte zuerst im Blatt " & WS_BANKKONTO & " die Zeile der" & vbCrLf & _
               "Buchung markieren, die einem anderen Mitglied" & vbCrLf & _
               "gutgeschrieben werden soll.", vbExclamation, titel
        Exit Function
    End If

    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    zeile = Selection.Cells(1, 1).Row

    If zeile < BK_START_ROW Or zeile > lastRow Then
        MsgBox "Die Markierung liegt nicht auf einer Buchungszeile.", vbExclamation, titel
        Exit Function
    End If

    ErmittleMarkierteBuchungszeile = zeile

End Function


' ===============================================================
' Lässt den Nutzer das Zielmitglied auswählen.
'
' Gesucht wird in der Zuordnungstabelle des Blattes Daten, weil
' dort auch ausgetretene und verstorbene Mitglieder noch stehen.
' Genau die braucht man im Erbfall.
' ===============================================================
Private Function WaehleZielmitglied(ByVal zahler As String, _
                                    ByVal zweck As String, _
                                    ByVal betrag As Double, _
                                    ByVal datum As Variant, _
                                    ByVal bisher As String, _
                                    ByRef outZielKey As String, _
                                    ByVal titel As String) As String

    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim suche As String
    Dim treffer As Object
    Dim anzahl As Long
    Dim liste As String
    Dim zuordnung As String
    Dim eingabe As String
    Dim gewaehlt As Long

    WaehleZielmitglied = ""
    outZielKey = ""

    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    If wsDaten Is Nothing Then Exit Function

    suche = InputBox("Buchung vom " & Format$(datum, "dd.mm.yyyy") & vbCrLf & _
                     "Betrag: " & Format$(betrag, "#,##0.00") & " " & ChrW(8364) & vbCrLf & _
                     "Gezahlt von: " & zahler & vbCrLf & _
                     "Verwendungszweck: " & Left$(zweck, 120) & vbCrLf & bisher & vbCrLf & _
                     "Name des Mitglieds, dem diese Zahlung" & vbCrLf & _
                     "gutgeschrieben werden soll." & vbCrLf & _
                     "Ein Namensteil genügt.", titel)

    suche = Trim$(suche)
    If suche = "" Then Exit Function

    Set treffer = CreateObject("Scripting.Dictionary")
    lastRow = wsDaten.Cells(wsDaten.Rows.count, DATA_MAP_COL_ENTITYKEY).End(xlUp).Row

    For r = DATA_START_ROW To lastRow
        zuordnung = Trim$(CStr(wsDaten.Cells(r, DATA_MAP_COL_ZUORDNUNG).value))
        If zuordnung = "" Then
            zuordnung = Trim$(CStr(wsDaten.Cells(r, DATA_MAP_COL_KTONAME).value))
        End If

        If zuordnung <> "" And Trim$(CStr(wsDaten.Cells(r, DATA_MAP_COL_ENTITYKEY).value)) <> "" Then
            If InStr(1, zuordnung, suche, vbTextCompare) > 0 Then
                anzahl = anzahl + 1
                treffer(anzahl) = Array( _
                    Trim$(CStr(wsDaten.Cells(r, DATA_MAP_COL_ENTITYKEY).value)), zuordnung)
                liste = liste & vbCrLf & anzahl & ". " & zuordnung & _
                        "   [Parzelle " & Trim$(CStr(wsDaten.Cells(r, DATA_MAP_COL_PARZELLE).value)) & _
                        ", " & Trim$(CStr(wsDaten.Cells(r, DATA_MAP_COL_ENTITYROLE).value)) & "]"
            End If
        End If
    Next r

    If anzahl = 0 Then
        MsgBox "Zu """ & suche & """ wurde in der Zuordnungstabelle" & vbCrLf & _
               "kein Eintrag gefunden.", vbExclamation, titel
        Exit Function
    End If

    If anzahl = 1 Then
        gewaehlt = 1
    Else
        eingabe = InputBox("Mehrere Einträge passen zu """ & suche & """:" & vbCrLf & _
                           liste & vbCrLf & vbCrLf & "Bitte die Nummer eingeben.", titel)
        If Not IsNumeric(eingabe) Then Exit Function
        gewaehlt = CLng(eingabe)
        If Not treffer.exists(gewaehlt) Then Exit Function
    End If

    outZielKey = CStr(treffer(gewaehlt)(0))
    WaehleZielmitglied = CStr(treffer(gewaehlt)(1))

End Function


Private Sub SpeichereSonderzuordnung(ByVal schluessel As String, _
                                     ByVal zielKey As String, _
                                     ByVal zielName As String, _
                                     ByVal grund As String)

    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim zielRow As Long
    Dim r As Long

    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    If wsDaten Is Nothing Then Exit Sub

    On Error Resume Next
    wsDaten.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0

    lastRow = wsDaten.Cells(wsDaten.Rows.count, SZ_COL_KEY).End(xlUp).Row
    If lastRow >= SZ_START_ROW Then
        For r = SZ_START_ROW To lastRow
            If StrComp(Trim$(CStr(wsDaten.Cells(r, SZ_COL_KEY).value)), schluessel, vbTextCompare) = 0 Then
                zielRow = r
                Exit For
            End If
        Next r
    End If
    If zielRow = 0 Then
        zielRow = IIf(lastRow < SZ_START_ROW, SZ_START_ROW, lastRow + 1)
    End If

    wsDaten.Cells(SZ_HEADER_ROW, SZ_COL_KEY).value = "SZ Buchungsschl" & ChrW(252) & "ssel"
    wsDaten.Cells(SZ_HEADER_ROW, SZ_COL_ZIEL_KEY).value = "SZ Ziel-Zuordnungsschl" & ChrW(252) & "ssel"
    wsDaten.Cells(SZ_HEADER_ROW, SZ_COL_ZIEL_NAME).value = "SZ Zielmitglied"
    wsDaten.Cells(SZ_HEADER_ROW, SZ_COL_GRUND).value = "SZ Grund"
    wsDaten.Cells(SZ_HEADER_ROW, SZ_COL_ERFASST).value = "SZ Erfasst am"

    wsDaten.Cells(zielRow, SZ_COL_KEY).value = schluessel
    wsDaten.Cells(zielRow, SZ_COL_ZIEL_KEY).value = zielKey
    wsDaten.Cells(zielRow, SZ_COL_ZIEL_NAME).value = zielName
    wsDaten.Cells(zielRow, SZ_COL_GRUND).value = grund
    wsDaten.Cells(zielRow, SZ_COL_ERFASST).value = Date

    On Error Resume Next
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0

End Sub


' ===============================================================
' Schreibt einen Vermerk in die Bemerkungsspalte der Buchung, damit
' der Sonderfall auch auf dem Blatt Bankkonto sichtbar bleibt.
' ===============================================================
Private Sub VermerkeInBemerkung(ByVal wsBK As Worksheet, _
                                ByVal zeile As Long, _
                                ByVal zielName As String)

    Dim vermerk As String
    Dim bemerkung As String
    Dim pos As Long

    vermerk = "Sonderzuordnung: Zahlung f" & ChrW(252) & "r " & zielName
    bemerkung = Trim$(CStr(wsBK.Cells(zeile, BK_COL_BEMERKUNG).value))

    ' Einen früheren Vermerk ersetzen statt mehrfach anhängen.
    pos = InStr(1, bemerkung, "Sonderzuordnung:", vbTextCompare)
    If pos > 0 Then
        bemerkung = Trim$(Left$(bemerkung, pos - 1))
        Do While Right$(bemerkung, 1) = "|"
            bemerkung = Trim$(Left$(bemerkung, Len(bemerkung) - 1))
        Loop
    End If

    On Error Resume Next
    If bemerkung = "" Then
        wsBK.Cells(zeile, BK_COL_BEMERKUNG).value = vermerk
    Else
        wsBK.Cells(zeile, BK_COL_BEMERKUNG).value = bemerkung & " | " & vermerk
    End If
    On Error GoTo 0

End Sub


' ===============================================================
' Macht aus dem technischen Buchungsschlüssel eine lesbare Angabe.
' ===============================================================
Private Function BeschreibeSchluessel(ByVal schluessel As String) As String

    Dim teile() As String

    BeschreibeSchluessel = schluessel
    teile = Split(schluessel, "|")
    If UBound(teile) < 1 Then Exit Function
    If Len(teile(0)) <> 8 Then Exit Function

    BeschreibeSchluessel = mid$(teile(0), 7, 2) & "." & mid$(teile(0), 5, 2) & "." & _
                           Left$(teile(0), 4) & ", " & teile(1) & " " & ChrW(8364)

End Function


' ===============================================================
' Legt die beiden Bedienschaltflächen auf dem Blatt Bankkonto an.
'
' Die Sonderzuordnung soll ohne Umweg über Entwicklerwerkzeuge oder
' das Makrofenster erreichbar sein. Die Schaltflächen werden bei jedem
' Öffnen der Arbeitsmappe neu aufgebaut, damit sie nicht verloren
' gehen können. Das Blatt wird dafür kurz entsperrt und danach
' wieder mit denselben Rechten geschützt wie sonst auch.
'
' Sie sitzen bewusst im leeren Feld rechts neben dem Importprotokoll
' (Spalten I bis K, oberhalb der Kennzahlen). Dort verdecken sie weder
' die Kontoführungsangaben links noch die Auszugsangaben rechts.
' ===============================================================
Public Sub ErstelleSonderzuordnungButtons(Optional ByVal wsBK As Worksheet = Nothing)

    Const REIHE_HOEHE As Single = 26

    Dim warGeschuetzt As Boolean
    Dim ankerLinks As Single
    Dim ankerOben As Single

    On Error GoTo Aufraeumen

    If wsBK Is Nothing Then
        Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    End If
    If wsBK Is Nothing Then Exit Sub

    warGeschuetzt = wsBK.ProtectContents
    If warGeschuetzt Then wsBK.Unprotect PASSWORD:=PASSWORD

    ankerLinks = wsBK.Range("I2").Left
    ankerOben = wsBK.Range("I2").Top

    Call ZeichneBedienschaltflaeche(wsBK, "btn_ZahlungZuordnen", _
         ChrW(8644) & "   Zahlung zuordnen", _
         ankerLinks, ankerOben, 166, REIHE_HOEHE, RGB(33, 156, 170), _
         "'mod_Sonderzuordnung.OrdneBuchungEinemMitgliedZu'")

    Call ZeichneBedienschaltflaeche(wsBK, "btn_ZuordnungenZeigen", _
         "Zuordnungen", _
         ankerLinks + 172, ankerOben, 100, REIHE_HOEHE, RGB(82, 88, 94), _
         "'mod_Sonderzuordnung.ZeigeSonderzuordnungen'")

Aufraeumen:
    On Error Resume Next
    If warGeschuetzt And Not wsBK Is Nothing Then
        wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    On Error GoTo 0

End Sub


' ===============================================================
' Zeichnet eine einzelne Schaltfläche im Stil der übrigen Kacheln.
'
' Eine vorhandene Schaltfläche gleichen Namens wird vorher entfernt,
' damit mehrfaches Aufrufen keine Stapel übereinanderliegender
' Formen erzeugt.
' ===============================================================
Private Sub ZeichneBedienschaltflaeche(ByVal ws As Worksheet, _
                                       ByVal formName As String, _
                                       ByVal beschriftung As String, _
                                       ByVal x As Single, _
                                       ByVal y As Single, _
                                       ByVal breite As Single, _
                                       ByVal hoehe As Single, _
                                       ByVal farbe As Long, _
                                       ByVal makro As String)

    Dim shp As Shape

    On Error Resume Next
    ws.Shapes(formName).Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, breite, hoehe)
    With shp
        .Name = formName
        .TextFrame2.TextRange.text = beschriftung
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Fill.ForeColor.RGB = farbe
        .Line.Visible = msoFalse
        .OnAction = makro
        .Placement = xlFreeFloating
    End With

End Sub
