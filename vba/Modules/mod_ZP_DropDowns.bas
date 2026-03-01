Attribute VB_Name = "mod_ZP_DropDowns"
Option Explicit

' ===============================================================
' MODUL: mod_ZP_DropDowns
' Ausgelagert aus mod_Zahlungspruefung
' Enthält: DropDown-Logik für Bankkonto-Blatt (Spalte H + I),
'          Hilfsspalten AF/AG, Spaltenentsperrung
' ===============================================================

Private Const FARBE_HELLGRUEN_MANUELL As Long = 13565382


' ===============================================================
' OEFFENTLICH: Setzt ALLE DropDowns auf dem Bankkonto-Blatt
' Wird von Tabelle3.Worksheet_Activate UND nach CSV-Import aufgerufen.
' Setzt:
'   - Spalte H (Kategorie): E- oder A-Kategorien je nach Betrag
'   - Spalte I (Monat/Periode): Januar bis Dezember
'   - Entsperrt editierbare Spalten (H, I, J, L)
' ===============================================================
Public Sub SetzeBankkontoDropDowns(ByVal wsBK As Worksheet)
    
    Dim lastRow As Long
    
    If wsBK Is Nothing Then Exit Sub
    
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Sub
    
    ' Hilfsspalten auf Daten-Blatt aktualisieren (AF + AG)
    Call AktualisiereKategorieHilfsspalten
    
    ' Blattschutz aufheben (noetig fuer Data Validation)
    On Error Resume Next
    wsBK.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' DropDowns setzen
    Call SetzeKategorieDropDowns(wsBK, lastRow)
    Call SetzeMonatDropDowns(wsBK, lastRow)
    
    ' Spalten entsperren fuer Nutzereingaben
    Call EntsperreSpaltenFuerNutzer(wsBK, lastRow)
    
    ' Blattschutz wieder aktivieren
    On Error Resume Next
    wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
End Sub


' ===============================================================
' Befuellt Hilfsspalten AF (32) + AG (33) auf Blatt "Daten"
' mit eindeutigen Kategorienamen, getrennt nach E und A.
' AF = Einnahmen-Kategorien (K = "E")
' AG = Ausgaben-Kategorien (K = "A")
' Quelle: Spalte J (DATA_CAT_COL_KATEGORIE = 10)
'         Spalte K (DATA_CAT_COL_EINAUS = 11)
' ===============================================================
Public Sub AktualisiereKategorieHilfsspalten()
    
    Dim wsDaten As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim katName As String
    Dim einAus As String
    
    Dim dictE As Object
    Dim dictA As Object
    
    Set dictE = CreateObject("Scripting.Dictionary")
    Set dictA = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    
    On Error Resume Next
    wsDaten.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    lastRow = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then GoTo ProtectAndExit
    
    ' Eindeutige Kategorien sammeln
    For r = DATA_START_ROW To lastRow
        katName = Trim(CStr(wsDaten.Cells(r, DATA_CAT_COL_KATEGORIE).value))
        If katName = "" Then GoTo NextHilfsRow
        
        einAus = UCase(Trim(CStr(wsDaten.Cells(r, DATA_CAT_COL_EINAUS).value)))
        
        If einAus = "E" Then
            If Not dictE.Exists(katName) Then dictE.Add katName, katName
        ElseIf einAus = "A" Then
            If Not dictA.Exists(katName) Then dictA.Add katName, katName
        End If
        
NextHilfsRow:
    Next r
    
    ' Hilfsspalten leeren (ab Zeile 4, max 200 Zeilen sicherheitshalber)
    Dim maxClear As Long
    maxClear = wsDaten.Cells(wsDaten.Rows.count, DATA_COL_KAT_EINNAHMEN).End(xlUp).Row
    If maxClear < DATA_START_ROW + 200 Then maxClear = DATA_START_ROW + 200
    
    wsDaten.Range(wsDaten.Cells(DATA_START_ROW, DATA_COL_KAT_EINNAHMEN), _
                  wsDaten.Cells(maxClear, DATA_COL_KAT_EINNAHMEN)).ClearContents
    wsDaten.Range(wsDaten.Cells(DATA_START_ROW, DATA_COL_KAT_AUSGABEN), _
                  wsDaten.Cells(maxClear, DATA_COL_KAT_AUSGABEN)).ClearContents
    
    ' Einnahmen in Spalte AF (DATA_COL_KAT_EINNAHMEN = 32) schreiben
    Dim idx As Long
    idx = DATA_START_ROW
    Dim key As Variant
    For Each key In dictE.keys
        wsDaten.Cells(idx, DATA_COL_KAT_EINNAHMEN).value = CStr(key)
        idx = idx + 1
    Next key
    
    ' Ausgaben in Spalte AG (DATA_COL_KAT_AUSGABEN = 33) schreiben
    idx = DATA_START_ROW
    For Each key In dictA.keys
        wsDaten.Cells(idx, DATA_COL_KAT_AUSGABEN).value = CStr(key)
        idx = idx + 1
    Next key
    
ProtectAndExit:
    On Error Resume Next
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
    Set dictE = Nothing
    Set dictA = Nothing
    
End Sub


' ===============================================================
' Setzt DropDown-Listen in Spalte H (Kategorie)
' Fuer jede Zeile: Betrag > 0 -> Einnahmen (AF), Betrag < 0 -> Ausgaben (AG)
' Referenziert dynamisch auf den befuellten Bereich in AF bzw. AG
' ===============================================================
Private Sub SetzeKategorieDropDowns(ByVal ws As Worksheet, ByVal lastRow As Long)
    
    If lastRow < BK_START_ROW Then Exit Sub
    
    Dim wsDaten As Worksheet
    
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    
    ' Letzter befuellter Eintrag in AF und AG ermitteln
    Dim lastE As Long
    lastE = wsDaten.Cells(wsDaten.Rows.count, DATA_COL_KAT_EINNAHMEN).End(xlUp).Row
    If lastE < DATA_START_ROW Then lastE = DATA_START_ROW
    
    Dim lastA As Long
    lastA = wsDaten.Cells(wsDaten.Rows.count, DATA_COL_KAT_AUSGABEN).End(xlUp).Row
    If lastA < DATA_START_ROW Then lastA = DATA_START_ROW
    
    ' Spaltenbuchstaben fuer Validation-Formeln berechnen
    Dim spalteBuchstabeE As String
    spalteBuchstabeE = SpalteNrZuBuchstabe(DATA_COL_KAT_EINNAHMEN)
    
    Dim spalteBuchstabeA As String
    spalteBuchstabeA = SpalteNrZuBuchstabe(DATA_COL_KAT_AUSGABEN)
    
    ' Daten-Blattname fuer Formel
    Dim datenName As String
    datenName = wsDaten.Name
    
    ' Validation-Formeln: =Daten!$AF$4:$AF$xx
    Dim formelEinnahmen As String
    formelEinnahmen = "=" & datenName & "!$" & spalteBuchstabeE & "$" & DATA_START_ROW & _
                      ":$" & spalteBuchstabeE & "$" & lastE
    
    Dim formelAusgaben As String
    formelAusgaben = "=" & datenName & "!$" & spalteBuchstabeA & "$" & DATA_START_ROW & _
                     ":$" & spalteBuchstabeA & "$" & lastA
    
    ' Pro Zeile die passende Validation setzen
    Dim r As Long
    Dim betrag As Double
    Dim formel As String
    
    On Error Resume Next
    
    For r = BK_START_ROW To lastRow
        betrag = 0
        If IsNumeric(ws.Cells(r, BK_COL_BETRAG).value) Then
            betrag = CDbl(ws.Cells(r, BK_COL_BETRAG).value)
        End If
        
        If betrag > 0 Then
            formel = formelEinnahmen
        ElseIf betrag < 0 Then
            formel = formelAusgaben
        Else
            ' Betrag = 0 oder leer: Einnahmen als Default
            formel = formelEinnahmen
        End If
        
        With ws.Cells(r, BK_COL_KATEGORIE).Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertInformation, _
                 Operator:=xlBetween, _
                 Formula1:=formel
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = False
        End With
    Next r
    
    On Error GoTo 0
    
End Sub


' ===============================================================
' Hilfsfunktion: Spaltennummer -> Spaltenbuchstabe
' (1="A", 26="Z", 27="AA", 28="AB", 32="AF", 33="AG" etc.)
' ===============================================================
Private Function SpalteNrZuBuchstabe(ByVal spalte As Long) As String
    Dim temp As String
    temp = ""
    Do While spalte > 0
        Dim rest As Long
        rest = (spalte - 1) Mod 26
        temp = Chr(65 + rest) & temp
        spalte = (spalte - 1) \ 26
    Loop
    SpalteNrZuBuchstabe = temp
End Function


' ===============================================================
' DropDown-Listen (Januar-Dezember) auf Spalte I setzen
' ===============================================================
Private Sub SetzeMonatDropDowns(ByVal ws As Worksheet, ByVal lastRow As Long)
    
    If lastRow < BK_START_ROW Then Exit Sub
    
    Dim monatsListe As String
    monatsListe = "Januar,Februar,M" & ChrW(228) & "rz,April,Mai,Juni," & _
                  "Juli,August,September,Oktober,November,Dezember"
    
    Dim rngMonat As Range
    Set rngMonat = ws.Range(ws.Cells(BK_START_ROW, BK_COL_MONAT_PERIODE), _
                            ws.Cells(lastRow, BK_COL_MONAT_PERIODE))
    
    On Error Resume Next
    With rngMonat.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertInformation, _
             Operator:=xlBetween, _
             Formula1:=monatsListe
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = False
    End With
    On Error GoTo 0
    
End Sub


' ===============================================================
' Spalten H, I, J, L entsperren fuer Nutzereingaben
' ===============================================================
Private Sub EntsperreSpaltenFuerNutzer(ByVal ws As Worksheet, ByVal lastRow As Long)
    
    If lastRow < BK_START_ROW Then Exit Sub
    
    On Error Resume Next
    
    ' Spalte H (Kategorie)
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_KATEGORIE), _
             ws.Cells(lastRow, BK_COL_KATEGORIE)).Locked = False
    
    ' Spalte I (Monat/Periode)
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_MONAT_PERIODE), _
             ws.Cells(lastRow, BK_COL_MONAT_PERIODE)).Locked = False
    
    ' Spalte J (Interne Nr) = Spalte 10
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_INTERNE_NR), _
             ws.Cells(lastRow, BK_COL_INTERNE_NR)).Locked = False
    
    ' Spalte L (Bemerkung)
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_BEMERKUNG), _
             ws.Cells(lastRow, BK_COL_BEMERKUNG)).Locked = False
    
    On Error GoTo 0
    
End Sub


