Attribute VB_Name = "mod_BK_KA_Nummern"
Option Explicit

' ***************************************************************
' MODUL: mod_BK_KA_Nummern
' VERSION: 1.0 - 15.05.2026
' ZWECK: Punkt 11 - Vergibt fortlaufende Nummern f?r
'        Bankkonto-Ausgaben (BK NN) und Vereinskassen-Eintr?ge (KA NN).
'        Bei Bargeldauszahlung wird zus?tzlich ein VK-Eintrag mit
'        Bezug zur BK-Nummer angelegt.
'
' Schema (Spalte J auf Bankkonto, Spalte F auf Vereinskasse):
'   - "BK 01", "BK 02", ... f?r Ausgaben (Betrag < 0) je Jahr
'   - "KA 01", "KA 02", ... f?r Vereinskassen-Eintr?ge je Jahr
'   - Bei Kategorie "Bargeldauszahlung":
'        Bankkonto Spalte J = "BK 03 / KA 01"
'        Vereinskasse Spalte F = "KA 01 / BK 03"
'
' Trigger:
'   - Worksheet_Change auf Bankkonto Spalte H (Kategorie)
'   - Nach jedem CSV-Import (BelegeAlleBKNummern aufrufen)
' ***************************************************************

Public Const KAT_BARGELDAUSZAHLUNG As String = "Bargeldauszahlung"

' Reentrancy-Schutz
Private m_IsRunning As Boolean


' ===============================================================
' Public: Vergibt BK-Nummer f?r eine einzelne Bankkonto-Zeile.
' Wird vom Worksheet_Change auf Bankkonto Spalte H aufgerufen.
' ===============================================================
Public Sub VergebeBKNummerFuerZeile(ByVal wsBK As Worksheet, ByVal zeile As Long)
    If m_IsRunning Then Exit Sub
    If wsBK Is Nothing Then Exit Sub
    If zeile < BK_START_ROW Then Exit Sub
    
    On Error GoTo CleanUp
    m_IsRunning = True
    Application.EnableEvents = False
    
    Call NeuberechneAlleBKNummern(wsBK)
    
CleanUp:
    Application.EnableEvents = True
    m_IsRunning = False
End Sub


' ===============================================================
' Public: Neuberechnung aller BK/KA-Nummern f?r das aktuelle
' Abrechnungsjahr. Wird nach CSV-Import oder bei ?nderung der
' Kategorie aufgerufen.
' ===============================================================
Public Sub NeuberechneAlleBKNummern(Optional ByVal wsBK As Worksheet = Nothing)
    On Error GoTo CleanUp
    
    Dim eigenerLauf As Boolean
    eigenerLauf = Not m_IsRunning
    If eigenerLauf Then
        m_IsRunning = True
        Application.EnableEvents = False
    End If
    
    If wsBK Is Nothing Then
        On Error Resume Next
        Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
        On Error GoTo CleanUp
        If wsBK Is Nothing Then GoTo CleanUp
    End If
    
    Dim wsVK As Worksheet
    On Error Resume Next
    Set wsVK = ThisWorkbook.Worksheets(WS_VEREINSKASSE)
    On Error GoTo CleanUp
    
    ' Blattschutz aufheben
    On Error Resume Next
    wsBK.Unprotect PASSWORD:=PASSWORD
    If Not wsVK Is Nothing Then wsVK.Unprotect PASSWORD:=PASSWORD
    On Error GoTo CleanUp
    
    Dim lastRow As Long
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then GoTo CleanUp
    
    ' --- Schritt 1: BK-Spalte J f?r das Abrechnungsjahr leeren ---
    '     (Eintr?ge aus Vorjahren bleiben unangetastet)
    Dim abrJahr As Long
    abrJahr = HoleAbrechnungsjahr
    
    Dim r As Long
    For r = BK_START_ROW To lastRow
        If IsDate(wsBK.Cells(r, BK_COL_DATUM).value) Then
            If Year(CDate(wsBK.Cells(r, BK_COL_DATUM).value)) = abrJahr Then
                wsBK.Cells(r, BK_COL_INTERNE_NR).value = ""
            End If
        End If
    Next r
    
    ' --- Schritt 2: BK-Zeilen sortiert nach Datum sammeln (nur Ausgaben) ---
    Dim sortIdx() As Long
    Dim sortDat() As Date
    Dim cnt As Long
    cnt = 0
    ReDim sortIdx(0 To lastRow - BK_START_ROW)
    ReDim sortDat(0 To lastRow - BK_START_ROW)
    
    For r = BK_START_ROW To lastRow
        If IsDate(wsBK.Cells(r, BK_COL_DATUM).value) Then
            Dim dat As Date
            dat = CDate(wsBK.Cells(r, BK_COL_DATUM).value)
            If Year(dat) = abrJahr Then
                If IsNumeric(wsBK.Cells(r, BK_COL_BETRAG).value) Then
                    If CDbl(wsBK.Cells(r, BK_COL_BETRAG).value) < 0 Then
                        sortIdx(cnt) = r
                        sortDat(cnt) = dat
                        cnt = cnt + 1
                    End If
                End If
            End If
        End If
    Next r
    
    If cnt = 0 Then GoTo NumeriereKA
    
    ReDim Preserve sortIdx(0 To cnt - 1)
    ReDim Preserve sortDat(0 To cnt - 1)
    
    ' Insertion-Sort nach Datum aufsteigend
    Dim i As Long, j As Long
    Dim tmpIdx As Long, tmpDat As Date
    For i = 1 To cnt - 1
        tmpIdx = sortIdx(i)
        tmpDat = sortDat(i)
        j = i - 1
        Do While j >= 0
            If sortDat(j) > tmpDat Then
                sortIdx(j + 1) = sortIdx(j)
                sortDat(j + 1) = sortDat(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        sortIdx(j + 1) = tmpIdx
        sortDat(j + 1) = tmpDat
    Next i
    
    ' --- Schritt 3: BK-Nummern vergeben ---
    Dim bkNr As Long
    bkNr = 0
    For i = 0 To cnt - 1
        bkNr = bkNr + 1
        wsBK.Cells(sortIdx(i), BK_COL_INTERNE_NR).value = "BK " & Format(bkNr, "00")
    Next i
    
NumeriereKA:
    ' --- Schritt 4: VK-Eintr?ge f?r Bargeldauszahlung sicherstellen + KA-Nummern vergeben ---
    If wsVK Is Nothing Then GoTo Schutz
    
    Call SyncBargeldauszahlungenZuVK(wsBK, wsVK, abrJahr)
    Call NumeriereVKEintraege(wsBK, wsVK, abrJahr)
    
Schutz:
    On Error Resume Next
    wsBK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    If Not wsVK Is Nothing Then wsVK.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0
    
CleanUp:
    If eigenerLauf Then
        Application.EnableEvents = True
        m_IsRunning = False
    End If
End Sub


' ===============================================================
' Synchronisiert Bargeldauszahlungen vom Bankkonto in die Vereinskasse:
' - F?r jede Bankkonto-Zeile mit Kategorie "Bargeldauszahlung" und
'   BK-Nummer im aktuellen Jahr wird ein VK-Eintrag sichergestellt
'   (anhand von Datum + Betrag identifiziert).
' - Bestehende VK-Eintr?ge werden NICHT dupliziert.
' ===============================================================
Private Sub SyncBargeldauszahlungenZuVK(ByVal wsBK As Worksheet, _
                                        ByVal wsVK As Worksheet, _
                                        ByVal jahr As Long)
    Dim lastBK As Long, lastVK As Long
    lastBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    lastVK = wsVK.Cells(wsVK.Rows.count, VK_COL_DATUM).End(xlUp).Row
    If lastVK < VK_START_ROW Then lastVK = VK_START_ROW - 1
    
    Dim r As Long, v As Long
    For r = BK_START_ROW To lastBK
        Dim kat As String
        kat = CStr(wsBK.Cells(r, BK_COL_KATEGORIE).value)
        If StrComp(kat, KAT_BARGELDAUSZAHLUNG, vbTextCompare) <> 0 Then GoTo NextBK
        If Not IsDate(wsBK.Cells(r, BK_COL_DATUM).value) Then GoTo NextBK
        Dim bkDatum As Date
        bkDatum = CDate(wsBK.Cells(r, BK_COL_DATUM).value)
        If Year(bkDatum) <> jahr Then GoTo NextBK
        If Not IsNumeric(wsBK.Cells(r, BK_COL_BETRAG).value) Then GoTo NextBK
        Dim bkBetrag As Double
        bkBetrag = CDbl(wsBK.Cells(r, BK_COL_BETRAG).value)
        If bkBetrag >= 0 Then GoTo NextBK
        
        Dim bkNrStr As String
        bkNrStr = CStr(wsBK.Cells(r, BK_COL_INTERNE_NR).value)
        ' Falls der Eintrag noch keine BK-Nr hat, ?berspringen (kommt in naechstem Lauf)
        If LenB(bkNrStr) = 0 Then GoTo NextBK
        ' Nur den BK-Teil verwenden (BK 03 / KA xx -> BK 03)
        Dim p As Long
        p = InStr(1, bkNrStr, "/")
        If p > 0 Then bkNrStr = Trim(Left$(bkNrStr, p - 1))
        
        ' Pr?fen ob VK-Eintrag bereits existiert (Datum + Betrag positiv)
        Dim gefunden As Boolean
        gefunden = False
        For v = VK_START_ROW To lastVK
            If IsDate(wsVK.Cells(v, VK_COL_DATUM).value) Then
                If CDate(wsVK.Cells(v, VK_COL_DATUM).value) = bkDatum Then
                    If IsNumeric(wsVK.Cells(v, VK_COL_BETRAG).value) Then
                        If Abs(CDbl(wsVK.Cells(v, VK_COL_BETRAG).value) - Abs(bkBetrag)) < 0.005 Then
                            ' Zus?tzlich pr?fen ob Beschreibung auf Bargeldauszahlung hindeutet
                            Dim besch As String
                            besch = LCase$(CStr(wsVK.Cells(v, VK_COL_BESCHREIBUNG).value))
                            If InStr(besch, LCase$(KAT_BARGELDAUSZAHLUNG)) > 0 Or _
                               InStr(LCase$(CStr(wsVK.Cells(v, VK_COL_INTERNE_NR).value)), "bk ") > 0 Then
                                gefunden = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next v
        
        If Not gefunden Then
            lastVK = lastVK + 1
            wsVK.Cells(lastVK, VK_COL_DATUM).value = bkDatum
            wsVK.Cells(lastVK, VK_COL_DATUM).NumberFormat = "DD.MM.YYYY"
            wsVK.Cells(lastVK, VK_COL_BESCHREIBUNG).value = KAT_BARGELDAUSZAHLUNG & " (" & bkNrStr & ")"
            wsVK.Cells(lastVK, VK_COL_NAME).value = "Bankkonto"
            wsVK.Cells(lastVK, VK_COL_BETRAG).value = Abs(bkBetrag)
            wsVK.Cells(lastVK, VK_COL_BETRAG).NumberFormat = "#,##0.00 " & ChrW(8364)
            ' Interne Nr wird in NumeriereVKEintraege gesetzt
        End If
        
NextBK:
    Next r
End Sub


' ===============================================================
' Numeriert alle VK-Eintr?ge im Abrechnungsjahr nach Datum
' fortlaufend mit "KA 01", "KA 02", ...
' Wenn der VK-Eintrag einer Bankkonto-Bargeldauszahlung entspricht,
' wird die KA-Nr UND BK-Nr verschraenkt eingetragen:
'   Bankkonto J: "BK 03 / KA 01"
'   Vereinskasse F: "KA 01 / BK 03"
' ===============================================================
Private Sub NumeriereVKEintraege(ByVal wsBK As Worksheet, _
                                 ByVal wsVK As Worksheet, _
                                 ByVal jahr As Long)
    Dim lastVK As Long
    lastVK = wsVK.Cells(wsVK.Rows.count, VK_COL_DATUM).End(xlUp).Row
    If lastVK < VK_START_ROW Then Exit Sub
    
    ' VK-Zeilen sammeln + sortieren nach Datum
    Dim cnt As Long
    cnt = 0
    Dim sortIdx() As Long
    Dim sortDat() As Date
    ReDim sortIdx(0 To lastVK - VK_START_ROW)
    ReDim sortDat(0 To lastVK - VK_START_ROW)
    
    Dim r As Long
    For r = VK_START_ROW To lastVK
        If IsDate(wsVK.Cells(r, VK_COL_DATUM).value) Then
            If Year(CDate(wsVK.Cells(r, VK_COL_DATUM).value)) = jahr Then
                sortIdx(cnt) = r
                sortDat(cnt) = CDate(wsVK.Cells(r, VK_COL_DATUM).value)
                cnt = cnt + 1
            End If
        End If
    Next r
    
    If cnt = 0 Then Exit Sub
    
    ReDim Preserve sortIdx(0 To cnt - 1)
    ReDim Preserve sortDat(0 To cnt - 1)
    
    Dim i As Long, j As Long, tmpI As Long, tmpD As Date
    For i = 1 To cnt - 1
        tmpI = sortIdx(i): tmpD = sortDat(i)
        j = i - 1
        Do While j >= 0
            If sortDat(j) > tmpD Then
                sortIdx(j + 1) = sortIdx(j)
                sortDat(j + 1) = sortDat(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        sortIdx(j + 1) = tmpI: sortDat(j + 1) = tmpD
    Next i
    
    ' KA-Nummern vergeben + Verschraenkung mit BK-Nr
    Dim lastBK As Long
    lastBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    Dim kaNr As Long
    kaNr = 0
    For i = 0 To cnt - 1
        kaNr = kaNr + 1
        Dim kaStr As String
        kaStr = "KA " & Format(kaNr, "00")
        Dim vRow As Long
        vRow = sortIdx(i)
        
        ' Korrespondierende BK-Zeile suchen (Bargeldauszahlung gleiches Datum + Betrag)
        Dim bkNrStr As String
        bkNrStr = ""
        Dim bkRowMatch As Long
        bkRowMatch = 0
        
        Dim vDatum As Date
        vDatum = CDate(wsVK.Cells(vRow, VK_COL_DATUM).value)
        Dim vBetrag As Double
        vBetrag = 0
        If IsNumeric(wsVK.Cells(vRow, VK_COL_BETRAG).value) Then
            vBetrag = CDbl(wsVK.Cells(vRow, VK_COL_BETRAG).value)
        End If
        
        Dim b As Long
        For b = BK_START_ROW To lastBK
            If IsDate(wsBK.Cells(b, BK_COL_DATUM).value) Then
                If CDate(wsBK.Cells(b, BK_COL_DATUM).value) = vDatum Then
                    If StrComp(CStr(wsBK.Cells(b, BK_COL_KATEGORIE).value), KAT_BARGELDAUSZAHLUNG, vbTextCompare) = 0 Then
                        If IsNumeric(wsBK.Cells(b, BK_COL_BETRAG).value) Then
                            If Abs(Abs(CDbl(wsBK.Cells(b, BK_COL_BETRAG).value)) - vBetrag) < 0.005 Then
                                Dim raw As String
                                raw = CStr(wsBK.Cells(b, BK_COL_INTERNE_NR).value)
                                Dim slashPos As Long
                                slashPos = InStr(1, raw, "/")
                                If slashPos > 0 Then
                                    bkNrStr = Trim(Left$(raw, slashPos - 1))
                                Else
                                    bkNrStr = Trim(raw)
                                End If
                                bkRowMatch = b
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next b
        
        If LenB(bkNrStr) > 0 Then
            wsVK.Cells(vRow, VK_COL_INTERNE_NR).value = kaStr & " / " & bkNrStr
            If bkRowMatch > 0 Then
                wsBK.Cells(bkRowMatch, BK_COL_INTERNE_NR).value = bkNrStr & " / " & kaStr
            End If
        Else
            wsVK.Cells(vRow, VK_COL_INTERNE_NR).value = kaStr
        End If
    Next i
End Sub


' ===============================================================
' (Abrechnungsjahr wird aus mod_Const.HoleAbrechnungsjahr bezogen)
' ===============================================================






















































