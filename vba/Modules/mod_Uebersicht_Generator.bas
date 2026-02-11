Attribute VB_Name = "mod_Uebersicht_Generator"
Option Explicit

' ***************************************************************
' MODUL: mod_Uebersicht_Generator
' VERSION: 1.1 - 11.02.2026
' ZWECK: Generiert Übersichtsblatt (Variante 2: Lange Tabelle)
'        - 14 Mitglieder (Parzellen 1-14)
'        - 12 Monate (Januar - Dezember)
'        - 5 Kategorien (Mitgliedsbeitrag, Pachtgebühr, Wasser, Strom, Müll)
'        - Zeigt Soll/Ist/Status für jede Kombination
'        - Behandelt Parzelle 5 (2 Personen, getrennte Konten) und
'          Parzelle 2 (2 Personen, Gemeinschaftskonto) korrekt
' FIX v1.1: InitialisiereNachDezemberCache -> InitialisiereNachDezemberCacheZP
'           MsgBox-Text: 'Uebersicht' -> 'Übersicht' (Umlaut-Vorgabe)
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

' Kategorien (müssen mit Einstellungen übereinstimmen!)
Private Const KAT_MITGLIEDSBEITRAG As String = "Mitgliedsbeitrag"
Private Const KAT_PACHTGEBUEHR As String = "Pachtgebühr"
Private Const KAT_WASSER As String = "Wasserkosten"
Private Const KAT_STROM As String = "Stromkosten"
Private Const KAT_MUELL As String = "Müllgebühren"

' Ampelfarben
Private Const AMPEL_GRUEN As Long = 12968900
Private Const AMPEL_GELB As Long = 10086143
Private Const AMPEL_ROT As Long = 9871103


' ===============================================================
' HAUPTFUNKTION: Generiert komplettes Übersichtsblatt
' ===============================================================
Public Sub GeneriereUebersicht(Optional ByVal jahr As Long = 0)
    
    On Error GoTo ErrorHandler
    
    Dim wsUeb As Worksheet
    Dim wsMitgl As Worksheet
    Dim startTime As Double
    Dim r As Long
    Dim parzelle As Long
    Dim monat As Long
    Dim kategorie As String
    Dim kategorien(1 To 5) As String
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
    
    ' Kategorien definieren
    kategorien(1) = KAT_MITGLIEDSBEITRAG
    kategorien(2) = KAT_PACHTGEBUEHR
    kategorien(3) = KAT_WASSER
    kategorien(4) = KAT_STROM
    kategorien(5) = KAT_MUELL
    
    ' Worksheets holen
    On Error Resume Next
    Set wsUeb = ThisWorkbook.Worksheets(WS_UEBERSICHT)
    Set wsMitgl = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    On Error GoTo ErrorHandler
    
    If wsUeb Is Nothing Or wsMitgl Is Nothing Then
        MsgBox "Blatt 'Übersicht' oder 'Mitgliederliste' nicht gefunden!", vbCritical
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
    ' FIX v1.1: Korrekter Prozedurname mit Suffix ZP
    Call mod_Zahlungspruefung.InitialisiereNachDezemberCacheZP(jahr)
    
    ' Mitgliederliste laden (nur aktive Mitglieder mit Parzelle)
    Set mitglieder = HoleAktiveMitglieder(wsMitgl)
    
    ' Daten generieren
    rowIdx = UEBERSICHT_START_ROW
    
    For Each mitglied In mitglieder
        parzelle = mitglied("Parzelle")
        entityKey = mitglied("EntityKey")
        Dim mitgliedName As String
        mitgliedName = mitglied("Name")
        
        For monat = 1 To 12
            Dim i As Long
            For i = 1 To 5
                kategorie = kategorien(i)
                
                ' Zahlung prüfen (mod_Zahlungspruefung)
                ergebnis = mod_Zahlungspruefung.PruefeZahlungen(entityKey, kategorie, monat, jahr)
                
                ' Ergebnis parsen: "GRÜN|Soll:50.00|Ist:50.00"
                teile = Split(ergebnis, "|")
                If UBound(teile) >= 2 Then
                    status = teile(0)
                    soll = CDbl(Replace(Split(teile(1), ":")(1), ",", "."))
                    ist = CDbl(Replace(Split(teile(2), ":")(1), ",", "."))
                Else
                    status = "ROT"
                    soll = 0
                    ist = 0
                End If
                
                ' Zeile schreiben
                wsUeb.Cells(rowIdx, UEB_COL_PARZELLE).value = parzelle
                wsUeb.Cells(rowIdx, UEB_COL_MITGLIED).value = mitgliedName
                wsUeb.Cells(rowIdx, UEB_COL_MONAT).value = Format(DateSerial(jahr, monat, 1), "MMMM YYYY")
                wsUeb.Cells(rowIdx, UEB_COL_KATEGORIE).value = kategorie
                wsUeb.Cells(rowIdx, UEB_COL_SOLL).value = soll
                wsUeb.Cells(rowIdx, UEB_COL_IST).value = ist
                wsUeb.Cells(rowIdx, UEB_COL_STATUS).value = status
                
                ' Farbe setzen
                Select Case status
                    Case "GRÜN"
                        wsUeb.Cells(rowIdx, UEB_COL_STATUS).Interior.color = AMPEL_GRUEN
                    Case "GELB"
                        wsUeb.Cells(rowIdx, UEB_COL_STATUS).Interior.color = AMPEL_GELB
                    Case "ROT"
                        wsUeb.Cells(rowIdx, UEB_COL_STATUS).Interior.color = AMPEL_ROT
                End Select
                
                rowIdx = rowIdx + 1
            Next i
        Next monat
    Next mitglied
    
    ' Formatierung anwenden
    Call FormatiereUebersicht(wsUeb, UEBERSICHT_START_ROW, rowIdx - 1)
    
    ' Einstellungen-Cache freigeben
    Call mod_Zahlungspruefung.EntladeEinstellungenCacheZP
    
    ' Blatt schützen
    On Error Resume Next
    wsUeb.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo ErrorHandler
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Dim endTime As Double
    endTime = Timer
    
    MsgBox "Übersicht erfolgreich generiert!" & vbLf & vbLf & _
           "Zeilen: " & (rowIdx - UEBERSICHT_START_ROW) & vbLf & _
           "Dauer: " & Format(endTime - startTime, "0.00") & " Sekunden", _
           vbInformation, "Fertig"
    
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Generieren der Übersicht:" & vbLf & vbLf & _
           Err.Description, vbCritical, "Fehler"
    
End Sub


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
' Behandelt Sonderfälle:
' - Parzelle 5: 2 Personen, getrennte Konten
' - Parzelle 2: 2 Personen, Gemeinschaftskonto
' ===============================================================
Private Function HoleAktiveMitglieder(ByVal wsMitgl As Worksheet) As Collection
    
    Dim col As Collection
    Set col = New Collection
    
    Dim lastRow As Long
    lastRow = wsMitgl.Cells(wsMitgl.Rows.count, M_COL_PARZELLE).End(xlUp).Row
    
    Dim r As Long
    Dim parzelle As Long
    Dim pachtende As String
    Dim entityKey As String
    Dim mitgliedName As String
    Dim dict As Object
    
    For r = M_START_ROW To lastRow
        parzelle = wsMitgl.Cells(r, M_COL_PARZELLE).value
        pachtende = Trim(CStr(wsMitgl.Cells(r, M_COL_PACHTENDE).value))
        entityKey = Trim(CStr(wsMitgl.Cells(r, M_COL_ENTITY_KEY).value))
        
        ' Nur aktive Mitglieder (kein Pachtende)
        If pachtende = "" And parzelle >= 1 And parzelle <= 14 And entityKey <> "" Then
            mitgliedName = Trim(wsMitgl.Cells(r, M_COL_VORNAME).value) & " " & _
                            Trim(wsMitgl.Cells(r, M_COL_NACHNAME).value)
            
            Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "Parzelle", parzelle
            dict.Add "EntityKey", entityKey
            dict.Add "Name", mitgliedName
            
            col.Add dict
        End If
    Next r
    
    Set HoleAktiveMitglieder = col
    
End Function


' ===============================================================
' Formatierung des Übersichtsblatts (Zebramuster, Rahmen, Spaltenbreiten)
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
    For r = startRow To endRow
        If (r - startRow) Mod 2 = 0 Then
            wsUeb.Range(wsUeb.Cells(r, UEB_COL_PARZELLE), _
                        wsUeb.Cells(r, UEB_COL_BEMERKUNG)).Interior.color = RGB(242, 242, 242)
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
    wsUeb.Columns(UEB_COL_KATEGORIE).ColumnWidth = 20
    wsUeb.Columns(UEB_COL_SOLL).ColumnWidth = 12
    wsUeb.Columns(UEB_COL_IST).ColumnWidth = 12
    wsUeb.Columns(UEB_COL_STATUS).ColumnWidth = 10
    wsUeb.Columns(UEB_COL_BEMERKUNG).ColumnWidth = 30
    
    ' Zahlenformat
    wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_SOLL), _
                wsUeb.Cells(endRow, UEB_COL_IST)).NumberFormat = "#,##0.00 "
    
    ' Ausrichtung
    wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_PARZELLE), _
                wsUeb.Cells(endRow, UEB_COL_PARZELLE)).HorizontalAlignment = xlCenter
    wsUeb.Range(wsUeb.Cells(startRow, UEB_COL_STATUS), _
                wsUeb.Cells(endRow, UEB_COL_STATUS)).HorizontalAlignment = xlCenter
    
    ' Vertikale Zentrierung
    rngTable.VerticalAlignment = xlCenter
    
End Sub

