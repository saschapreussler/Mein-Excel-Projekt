Attribute VB_Name = "mod_Uebersicht_Generator"
Option Explicit

' ***************************************************************
' MODUL: mod_Uebersicht_Generator
' VERSION: 2.0 - 12.02.2026
' ZWECK: Generiert Übersichtsblatt (Variante 2: Lange Tabelle)
'        - 14 Mitglieder (Parzellen 1-14)
'        - 12 Monate (Januar - Dezember)
'        - Kategorien DYNAMISCH aus Einstellungen-Blatt (Spalte B)
'        - Zeigt Soll/Ist/Status für jede Kombination
'        - Behandelt Parzelle 5 (2 Personen, getrennte Konten) und
'          Parzelle 2 (2 Personen, Gemeinschaftskonto) korrekt
'        - Bei Kategorien OHNE festen Soll-Betrag:
'          Soll-Zelle bleibt leer + hell-gelb + editierbar
'          Nur Zahlungstermin-Prüfung (pünktlich / Säumnis)
'        - Säumnis-Gebühren werden in Bemerkung angezeigt
' FIX v1.1: InitialisiereNachDezemberCache -> InitialisiereNachDezemberCacheZP
' FIX v1.2: Val() statt CDbl() für systemunabhängiges Parsen
' FIX v1.3: "Typen unverträglich" behoben (Variant, StrComp, etc.)
' FIX v1.4: ChrW() in Const nicht erlaubt -> Private Variablen
' NEU v2.0: Kategorien DYNAMISCH aus Einstellungen-Blatt
'           - Keine hart kodierten Kategorienamen mehr
'           - Soll-Betrag 0 -> Zelle leer + hell-gelb + editierbar
'           - Zahlungstermin-Prüfung auch ohne Soll-Betrag
'           - Säumnis-Gebühren in Bemerkung
' ***************************************************************

' ===============================================================
' KONSTANTEN
' ===============================================================
Private Const UEBERSICHT_START_ROW As Long = 4
Private Const UEBERSICHT_HEADER_ROW As Long = 3

' Spalten im Übersichtsblatt
Private Const UEB_COL_PARZELLE As Long = 1      ' A - Parzelle
Private Const UEB_COL_MITGLIED As Long = 2      ' B - Mitglied
Private Const UEB_COL_MONAT As Long = 3         ' C - Monat
Private Const UEB_COL_KATEGORIE As Long = 4     ' D - Kategorie
Private Const UEB_COL_SOLL As Long = 5          ' E - Soll
Private Const UEB_COL_IST As Long = 6           ' F - Ist
Private Const UEB_COL_STATUS As Long = 7        ' G - Status (GRÜN/GELB/ROT)
Private Const UEB_COL_BEMERKUNG As Long = 8     ' H - Bemerkung

' Ampelfarben
Private Const AMPEL_GRUEN As Long = 12968900    ' RGB(196, 225, 196)
Private Const AMPEL_GELB As Long = 10086143     ' RGB(255, 235, 156)
Private Const AMPEL_ROT As Long = 9871103       ' RGB(255, 199, 206)

' Hell-gelb für "bitte manuell befüllen" (Soll-Betrag variabel)
Private Const FARBE_HELLGELB_MANUELL As Long = 10092543  ' RGB(255, 255, 153)

' Status-String für GRÜN (Encoding-sicher, wird in Init gesetzt)
Private m_STATUS_GRUEN As String
Private m_StatusInitialisiert As Boolean


' ===============================================================
' Type für eine dynamische Kategorie aus Einstellungen
' ===============================================================
Private Type UebKategorie
    Name As String
    SollBetrag As Double
    HatFestenSoll As Boolean      ' True wenn Spalte C > 0
    SaeumnisGebuehr As Double     ' Spalte I auf Einstellungen
    SollMonate As String          ' Spalte E: "03, 06, 09" oder leer = alle
End Type


' ===============================================================
' Initialisiert Status-String (Encoding-sicher)
' ===============================================================
Private Sub InitStatus()
    
    If m_StatusInitialisiert Then Exit Sub
    
    m_STATUS_GRUEN = "GR" & ChrW(220) & "N"
    m_StatusInitialisiert = True
    
End Sub


' ===============================================================
' HAUPTFUNKTION: Generiert komplettes Übersichtsblatt
' v2.0: Kategorien DYNAMISCH aus Einstellungen-Blatt
' ===============================================================
Public Sub GeneriereUebersicht(Optional ByVal jahr As Long = 0)
    
    On Error GoTo ErrorHandler
    
    ' Status initialisieren (Encoding-sicher)
    Call InitStatus
    
    Dim wsUeb As Worksheet
    Dim wsMitgl As Worksheet
    Dim startTime As Double
    Dim monat As Long
    Dim kategorie As String
    Dim mitglieder As Collection
    Dim mitglied As Object
    Dim entityKey As String
    Dim ergebnis As String
    Dim teile() As String
    Dim soll As Double
    Dim ist As Double
    Dim status As String
    Dim rowIdx As Long
    
    startTime = Timer
    
    ' Jahr-Parameter validieren
    If jahr = 0 Then jahr = Year(Date)
    
    ' =============================================
    ' v2.0: Kategorien DYNAMISCH aus Einstellungen laden
    ' =============================================
    Dim kategorien() As UebKategorie
    Dim anzahlKat As Long
    Call LadeKategorienAusEinstellungen(kategorien, anzahlKat)
    
    If anzahlKat = 0 Then
        MsgBox "Keine Kategorien im Einstellungen-Blatt (Spalte B) gefunden!" & vbLf & _
               "Bitte mindestens eine Kategorie mit Zahlungstermin anlegen.", _
               vbCritical, "Fehler"
        Exit Sub
    End If
    
    ' Worksheets holen
    On Error Resume Next
    Set wsUeb = ThisWorkbook.Worksheets(WS_UEBERSICHT)
    Set wsMitgl = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    On Error GoTo ErrorHandler
    
    If wsUeb Is Nothing Or wsMitgl Is Nothing Then
        MsgBox "Blatt '" & ChrW(220) & "bersicht' oder 'Mitgliederliste' nicht gefunden!", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Blatt entsperren
    On Error Resume Next
    wsUeb.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    ' Alten Inhalt löschen (ab Zeile 4)
    wsUeb.Range(wsUeb.Cells(UEBERSICHT_START_ROW, 1), _
                wsUeb.Cells(wsUeb.Rows.count, UEB_COL_BEMERKUNG)).ClearContents
    wsUeb.Range(wsUeb.Cells(UEBERSICHT_START_ROW, 1), _
                wsUeb.Cells(wsUeb.Rows.count, UEB_COL_BEMERKUNG)).Interior.ColorIndex = xlNone
    
    ' Header setzen
    Call SetzeUebersichtHeader(wsUeb)
    
    ' Einstellungen-Cache laden (Performance)
    Call mod_Zahlungspruefung.LadeEinstellungenCacheZP
    
    ' Dezember-Cache initialisieren (für Vorauszahlungen)
    Call mod_Zahlungspruefung.InitialisiereNachDezemberCacheZP(jahr)
    
    ' Mitgliederliste laden (nur aktive Mitglieder mit Parzelle)
    Set mitglieder = HoleAktiveMitglieder(wsMitgl)
    
    ' Daten generieren
    rowIdx = UEBERSICHT_START_ROW
    
    For Each mitglied In mitglieder
        Dim parzelleWert As Variant
        parzelleWert = mitglied("Parzelle")
        entityKey = mitglied("EntityKey")
        Dim mitgliedName As String
        mitgliedName = mitglied("Name")
        
        For monat = 1 To 12
            Dim k As Long
            For k = 0 To anzahlKat - 1
                
                ' Prüfen ob diese Kategorie in diesem Monat fällig ist
                If Not IstKategorieImMonatFaellig(kategorien(k), monat) Then
                    GoTo NextKat
                End If
                
                kategorie = kategorien(k).Name
                
                ' Zahlung prüfen (mod_Zahlungspruefung)
                ergebnis = mod_Zahlungspruefung.PruefeZahlungen(entityKey, kategorie, monat, jahr)
                
                ' Ergebnis parsen: "GRÜN|Soll:50.00|Ist:50.00"
                ' WICHTIG: Dezimaltrenner ist IMMER Punkt (.)
                soll = 0
                ist = 0
                status = "ROT"
                
                teile = Split(ergebnis, "|")
                If UBound(teile) >= 2 Then
                    status = teile(0)
                    
                    ' Soll parsen: "Soll:50.00" -> "50.00"
                    Dim sollTeile() As String
                    sollTeile = Split(teile(1), ":")
                    If UBound(sollTeile) >= 1 Then
                        soll = Val(sollTeile(1))
                    End If
                    
                    ' Ist parsen: "Ist:50.00" -> "50.00"
                    Dim istTeile() As String
                    istTeile = Split(teile(2), ":")
                    If UBound(istTeile) >= 1 Then
                        ist = Val(istTeile(1))
                    End If
                ElseIf UBound(teile) >= 0 Then
                    status = teile(0)
                End If
                
                ' Zeile schreiben
                wsUeb.Cells(rowIdx, UEB_COL_PARZELLE).value = parzelleWert
                wsUeb.Cells(rowIdx, UEB_COL_MITGLIED).value = mitgliedName
                wsUeb.Cells(rowIdx, UEB_COL_MONAT).value = Format(DateSerial(jahr, monat, 1), "MMMM YYYY")
                wsUeb.Cells(rowIdx, UEB_COL_KATEGORIE).value = kategorie
                
                ' =============================================
                ' v2.0: Soll-Betrag Logik
                ' =============================================
                If kategorien(k).HatFestenSoll Then
                    ' Fester Soll-Betrag aus Einstellungen
                    wsUeb.Cells(rowIdx, UEB_COL_SOLL).value = soll
                Else
                    ' KEIN fester Soll-Betrag -> Zelle leer + hell-gelb
                    ' Nutzer kann hier pro Parzelle den individuellen Betrag eintragen
                    wsUeb.Cells(rowIdx, UEB_COL_SOLL).value = ""
                    wsUeb.Cells(rowIdx, UEB_COL_SOLL).Interior.color = FARBE_HELLGELB_MANUELL
                    wsUeb.Cells(rowIdx, UEB_COL_SOLL).Locked = False
                    
                    ' Status bei variablem Betrag: nur Termin-Prüfung
                    ' Wenn Ist > 0 -> Zahlung eingegangen -> GRÜN
                    ' Wenn Ist = 0 -> Keine Zahlung -> ROT oder GELB
                    If ist > 0 Then
                        status = m_STATUS_GRUEN
                    End If
                End If
                
                wsUeb.Cells(rowIdx, UEB_COL_IST).value = ist
                wsUeb.Cells(rowIdx, UEB_COL_STATUS).value = status
                
                ' Farbe setzen
                If StrComp(status, m_STATUS_GRUEN, vbTextCompare) = 0 Then
                    wsUeb.Cells(rowIdx, UEB_COL_STATUS).Interior.color = AMPEL_GRUEN
                ElseIf StrComp(status, "GELB", vbTextCompare) = 0 Then
                    wsUeb.Cells(rowIdx, UEB_COL_STATUS).Interior.color = AMPEL_GELB
                Else
                    wsUeb.Cells(rowIdx, UEB_COL_STATUS).Interior.color = AMPEL_ROT
                End If
                
                ' =============================================
                ' v2.0: Bemerkung mit Säumnis-Info
                ' =============================================
                Dim bemerkung As String
                bemerkung = ""
                
                ' Zusatzinfo aus Ergebnis (4. Teil)
                If UBound(teile) >= 3 Then
                    bemerkung = teile(3)
                End If
                
                ' Säumnis-Gebühr anhängen wenn Status ROT und Gebühr definiert
                If StrComp(status, "ROT", vbTextCompare) = 0 Then
                    If kategorien(k).SaeumnisGebuehr > 0 Then
                        Dim saeumnisText As String
                        saeumnisText = "S" & ChrW(228) & "umnis-Geb" & ChrW(252) & "hr: " & _
                                       Format(kategorien(k).SaeumnisGebuehr, "#,##0.00") & _
                                       " " & ChrW(8364)
                        If bemerkung = "" Then
                            bemerkung = saeumnisText
                        Else
                            bemerkung = bemerkung & " | " & saeumnisText
                        End If
                    End If
                End If
                
                ' Kein fester Soll -> Hinweis
                If Not kategorien(k).HatFestenSoll Then
                    Dim variabelHinweis As String
                    variabelHinweis = "Soll-Betrag variabel (bitte manuell eintragen)"
                    If bemerkung = "" Then
                        bemerkung = variabelHinweis
                    Else
                        bemerkung = bemerkung & " | " & variabelHinweis
                    End If
                End If
                
                wsUeb.Cells(rowIdx, UEB_COL_BEMERKUNG).value = bemerkung
                
                rowIdx = rowIdx + 1
                
NextKat:
            Next k
        Next monat
    Next mitglied
    
    ' Formatierung anwenden
    Call FormatiereUebersicht(wsUeb, UEBERSICHT_START_ROW, rowIdx - 1)
    
    ' Einstellungen-Cache freigeben
    Call mod_Zahlungspruefung.EntladeEinstellungenCacheZP
    
    ' Blatt schützen (Soll-Zellen ohne festen Betrag bleiben editierbar)
    On Error Resume Next
    wsUeb.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo ErrorHandler
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Dim endTime As Double
    endTime = Timer
    
    MsgBox ChrW(220) & "bersicht erfolgreich generiert!" & vbLf & vbLf & _
           "Zeilen: " & (rowIdx - UEBERSICHT_START_ROW) & vbLf & _
           "Kategorien: " & anzahlKat & " (dynamisch aus Einstellungen)" & vbLf & _
           "Dauer: " & Format(endTime - startTime, "0.00") & " Sekunden", _
           vbInformation, "Fertig"
    
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Generieren der " & ChrW(220) & "bersicht:" & vbLf & vbLf & _
           Err.Description, vbCritical, "Fehler"
    
End Sub


' ===============================================================
' v2.0: Lädt Kategorien DYNAMISCH aus Einstellungen-Blatt
' Liest Spalte B (Kategorie), C (Soll-Betrag), E (Soll-Monate),
' I (Säumnis-Gebühr)
' Gibt eindeutige Kategorien zurück (keine Duplikate)
' ===============================================================
Private Sub LadeKategorienAusEinstellungen(ByRef kategorien() As UebKategorie, _
                                            ByRef anzahl As Long)
    
    Dim wsEinst As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim katName As String
    Dim dict As Object
    
    anzahl = 0
    
    On Error Resume Next
    Set wsEinst = ThisWorkbook.Worksheets(WS_EINSTELLUNGEN)
    On Error GoTo 0
    
    If wsEinst Is Nothing Then Exit Sub
    
    lastRow = wsEinst.Cells(wsEinst.Rows.count, ES_COL_KATEGORIE).End(xlUp).Row
    If lastRow < ES_START_ROW Then Exit Sub
    
    ' Dictionary für Eindeutigkeit
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Zuerst zählen für ReDim
    For r = ES_START_ROW To lastRow
        katName = Trim(CStr(wsEinst.Cells(r, ES_COL_KATEGORIE).value))
        If katName <> "" Then
            If Not dict.Exists(katName) Then
                dict.Add katName, r  ' Merke Zeilennummer für späteres Lesen
            End If
        End If
    Next r
    
    anzahl = dict.count
    If anzahl = 0 Then Exit Sub
    
    ReDim kategorien(0 To anzahl - 1)
    
    Dim idx As Long
    idx = 0
    Dim key As Variant
    
    For Each key In dict.keys
        r = dict(key)  ' Zeilennummer aus Dictionary
        
        With kategorien(idx)
            .Name = CStr(key)
            
            ' Soll-Betrag aus Spalte C
            Dim sollWert As Variant
            sollWert = wsEinst.Cells(r, ES_COL_SOLL_BETRAG).value
            If IsNumeric(sollWert) Then
                .SollBetrag = CDbl(sollWert)
            Else
                .SollBetrag = 0
            End If
            .HatFestenSoll = (.SollBetrag > 0)
            
            ' Säumnis-Gebühr aus Spalte I
            Dim saeumnisWert As Variant
            saeumnisWert = wsEinst.Cells(r, ES_COL_SAEUMNIS).value
            If IsNumeric(saeumnisWert) Then
                .SaeumnisGebuehr = CDbl(saeumnisWert)
            Else
                .SaeumnisGebuehr = 0
            End If
            
            ' Soll-Monate aus Spalte E (z.B. "03, 06, 09" oder leer = alle)
            .SollMonate = Trim(CStr(wsEinst.Cells(r, ES_COL_SOLL_MONATE).value))
        End With
        
        idx = idx + 1
    Next key
    
    Set dict = Nothing
    
End Sub


' ===============================================================
' v2.0: Prüft ob eine Kategorie in einem bestimmten Monat fällig ist
' Wenn SollMonate leer -> gilt für ALLE Monate
' Wenn SollMonate = "03, 06, 09" -> nur in diesen Monaten
' ===============================================================
Private Function IstKategorieImMonatFaellig(ByRef kat As UebKategorie, _
                                             ByVal monat As Long) As Boolean
    
    ' Keine Monate definiert -> gilt für ALLE Monate
    If kat.SollMonate = "" Then
        IstKategorieImMonatFaellig = True
        Exit Function
    End If
    
    ' Monate parsen: "03, 06, 09"
    Dim monate() As String
    monate = Split(kat.SollMonate, ",")
    
    Dim m As Long
    For m = LBound(monate) To UBound(monate)
        If IsNumeric(Trim(monate(m))) Then
            If CLng(Trim(monate(m))) = monat Then
                IstKategorieImMonatFaellig = True
                Exit Function
            End If
        End If
    Next m
    
    IstKategorieImMonatFaellig = False
    
End Function


' ===============================================================
' Header im Übersichtsblatt setzen
' ===============================================================
Private Sub SetzeUebersichtHeader(ByVal wsUeb As Worksheet)
    
    With wsUeb
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_PARZELLE).value = "Parzelle"
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_MITGLIED).value = "Mitglied"
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_MONAT).value = "Monat"
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_KATEGORIE).value = "Kategorie"
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_SOLL).value = "Soll"
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_IST).value = "Ist"
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_STATUS).value = "Status"
        .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_BEMERKUNG).value = "Bemerkung"
        
        ' Header formatieren
        Dim rngHeader As Range
        Set rngHeader = .Range(.Cells(UEBERSICHT_HEADER_ROW, UEB_COL_PARZELLE), _
                                .Cells(UEBERSICHT_HEADER_ROW, UEB_COL_BEMERKUNG))
        
        With rngHeader
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.color = RGB(217, 217, 217)  ' Hellgrau
            .Borders.LineStyle = xlContinuous
        End With
    End With
    
End Sub


' ===============================================================
' Holt alle aktiven Mitglieder mit Parzelle aus Mitgliederliste
' ===============================================================
Private Function HoleAktiveMitglieder(ByVal wsMitgl As Worksheet) As Collection
    
    Dim col As Collection
    Set col = New Collection
    
    Dim lastRow As Long
    lastRow = wsMitgl.Cells(wsMitgl.Rows.count, M_COL_PARZELLE).End(xlUp).Row
    
    Dim r As Long
    Dim parzelleWert As Variant
    Dim parzelleNr As Long
    Dim pachtende As String
    Dim entityKey As String
    Dim mitgliedName As String
    Dim dict As Object
    
    For r = M_START_ROW To lastRow
        parzelleWert = wsMitgl.Cells(r, M_COL_PARZELLE).value
        
        ' Prüfen ob Parzelle numerisch ist
        If isEmpty(parzelleWert) Then GoTo NextMitglRow
        If Not IsNumeric(parzelleWert) Then GoTo NextMitglRow
        
        parzelleNr = CLng(parzelleWert)
        
        ' Nur Parzellen 1-14
        If parzelleNr < 1 Or parzelleNr > 14 Then GoTo NextMitglRow
        
        pachtende = Trim(CStr(wsMitgl.Cells(r, M_COL_PACHTENDE).value))
        entityKey = Trim(CStr(wsMitgl.Cells(r, M_COL_ENTITY_KEY).value))
        
        ' Nur aktive Mitglieder (kein Pachtende) mit EntityKey
        If pachtende = "" And entityKey <> "" Then
            mitgliedName = Trim(CStr(wsMitgl.Cells(r, M_COL_VORNAME).value)) & " " & _
                            Trim(CStr(wsMitgl.Cells(r, M_COL_NACHNAME).value))
            
            Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "Parzelle", parzelleNr
            dict.Add "EntityKey", entityKey
            dict.Add "Name", mitgliedName
            
            col.Add dict
        End If
        
NextMitglRow:
    Next r
    
    Set HoleAktiveMitglieder = col
    
End Function


' ===============================================================
' Formatierung des Übersichtsblatts
' ===============================================================
Private Sub FormatiereUebersicht(ByVal wsUeb As Worksheet, _
                                   ByVal startRow As Long, _
                                   ByVal endRow As Long)
    
    Dim r As Long
    Dim rngTable As Range
    
    If endRow < startRow Then Exit Sub
    
    Set rngTable = wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_PARZELLE), _
                                wsUeb.Cells(endRow, UEB_COL_BEMERKUNG))
    
    ' Zebramuster (jede 2. Zeile hellgrau)
    ' ACHTUNG: Nicht überschreiben wenn Zelle bereits hell-gelb ist (variabler Soll)
    For r = startRow To endRow
        If (r - startRow) Mod 2 = 0 Then
            Dim c As Long
            For c = UEB_COL_PARZELLE To UEB_COL_BEMERKUNG
                ' Nur Zebra setzen wenn Zelle NICHT bereits speziell gefärbt ist
                ' (hell-gelb für variablen Soll, Ampelfarben für Status)
                If c <> UEB_COL_SOLL And c <> UEB_COL_STATUS Then
                    wsUeb.Cells(r, c).Interior.color = RGB(242, 242, 242)
                ElseIf c = UEB_COL_SOLL Then
                    ' Nur Zebra wenn NICHT hell-gelb (variabel)
                    If wsUeb.Cells(r, c).Interior.color <> FARBE_HELLGELB_MANUELL Then
                        wsUeb.Cells(r, c).Interior.color = RGB(242, 242, 242)
                    End If
                End If
                ' Status-Spalte (G) behält immer ihre Ampelfarbe
            Next c
        End If
    Next r
    
    ' Rahmen
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' Spaltenbreiten
    wsUeb.Columns(UEB_COL_PARZELLE).ColumnWidth = 10
    wsUeb.Columns(UEB_COL_MITGLIED).ColumnWidth = 25
    wsUeb.Columns(UEB_COL_MONAT).ColumnWidth = 18
    wsUeb.Columns(UEB_COL_KATEGORIE).ColumnWidth = 22
    wsUeb.Columns(UEB_COL_SOLL).ColumnWidth = 14
    wsUeb.Columns(UEB_COL_IST).ColumnWidth = 14
    wsUeb.Columns(UEB_COL_STATUS).ColumnWidth = 10
    wsUeb.Columns(UEB_COL_BEMERKUNG).ColumnWidth = 45
    
    ' Deutsches Zahlenformat mit Euro-Zeichen
    wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_SOLL), _
                wsUeb.Cells(endRow, UEB_COL_IST)).NumberFormat = "#.##0,00 " & ChrW(8364)
    
    ' Ausrichtung
    wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_PARZELLE), _
                wsUeb.Cells(endRow, UEB_COL_PARZELLE)).HorizontalAlignment = xlCenter
    wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_STATUS), _
                wsUeb.Cells(endRow, UEB_COL_STATUS)).HorizontalAlignment = xlCenter
    
    ' Vertikale Zentrierung
    rngTable.VerticalAlignment = xlCenter
    
End Sub

