Attribute VB_Name = "mod_Dashboard_Matrix"
Option Explicit

' ***************************************************************
' MODUL: mod_Dashboard_Matrix
' VERSION: 2.0 - 27.03.2026
' ZWECK: Zahlungsmatrix, Verzugsdetail und Spaltenanpassung
'        fuer das Dashboard "Uebersicht (neu)"
' ABHAENGIGKEITEN: mod_Uebersicht_Dashboard (Types, Farben),
'                  mod_Uebersicht_Daten, mod_Zahlungspruefung
' ***************************************************************


' ============================================================
'  MATRIX MIT DATEN SCHREIBEN
'  Sammelt gleichzeitig KPI-Werte und Verzug-Eintraege
' ============================================================
Public Sub SchreibeMatrixMitDaten(ByVal ws As Worksheet, _
                                    ByVal jahr As Long, _
                                    ByRef kategorien() As UebKategorie, _
                                    ByVal anzKat As Long, _
                                    ByRef parzellen() As ParzelleInfo, _
                                    ByVal anzParz As Long, _
                                    ByVal mitglieder As Collection, _
                                    ByVal sollDict As Object, _
                                    ByRef matrixEndRow As Long, _
                                    ByRef kpiSummeIst As Double, _
                                    ByRef kpiSummeSoll As Double, _
                                    ByRef kpiSummeSaeumnis As Double, _
                                    ByRef kpiAnzahlOffen As Long, _
                                    ByRef kpiAnzahlBezahlt As Long, _
                                    ByRef kpiAnzahlSaeumnis As Long, _
                                    ByRef kpiOffenOhneSoll As Long, _
                                    ByRef kpiOffenBetrag As Double, _
                                    ByRef verzugListe() As VerzugEintrag, _
                                    ByRef anzVerzug As Long)
    
    Dim statusGruen As String
    statusGruen = "GR" & ChrW(220) & "N"
    
    Dim importierteMonate() As Boolean
    importierteMonate = mod_Uebersicht_Daten.ErmittleImportierteMonate(jahr)
    
    ' --- PUNKT 12: Kategorien klassifizieren (monatlich vs. jaehrlich) ---
    Dim katIstJaehrlich() As Boolean
    Dim katSpalte() As Long
    ReDim katIstJaehrlich(0 To anzKat)
    ReDim katSpalte(0 To anzKat)
    Dim anzMonatlich As Long: anzMonatlich = 0
    Dim anzJaehrlich As Long: anzJaehrlich = 0
    Dim kk As Long
    For kk = 0 To anzKat - 1
        katIstJaehrlich(kk) = IstJaehrlicheKategorie(kategorien(kk))
        If katIstJaehrlich(kk) Then
            katSpalte(kk) = 0    ' wird in Sammelspalte geschrieben
            anzJaehrlich = anzJaehrlich + 1
        Else
            anzMonatlich = anzMonatlich + 1
            katSpalte(kk) = 2 + anzMonatlich   ' 3, 4, 5, ...
        End If
    Next kk
    
    ' --- Header (ohne Nr-Spalte) ---
    Dim headerRow As Long
    headerRow = DASH_MATRIX_HEADER_ROW
    
    ws.Cells(headerRow, 1).value = "Parzelle"
    ws.Cells(headerRow, 2).value = "Mitglied(er)"
    
    Dim k As Long
    For k = 0 To anzKat - 1
        If Not katIstJaehrlich(k) Then
            ws.Cells(headerRow, katSpalte(k)).value = kategorien(k).Name
        End If
    Next k
    
    ' Jahresposten-Spalte (Punkt 12)
    Dim colJahresposten As Long
    colJahresposten = 0
    If anzJaehrlich > 0 Then
        colJahresposten = 3 + anzMonatlich
        ws.Cells(headerRow, colJahresposten).value = "Jahresposten"
    End If
    
    Dim colGesamt As Long
    colGesamt = 3 + anzMonatlich + IIf(anzJaehrlich > 0, 1, 0)
    ws.Cells(headerRow, colGesamt).value = "Gesamt"
    ws.Cells(headerRow, colGesamt + 1).value = "Quote"
    
    Dim letzteSpalte As Long
    letzteSpalte = colGesamt + 1
    
    ' Header formatieren
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, letzteSpalte))
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .Font.color = m_CLR_WEISS
        .Interior.color = m_CLR_HEADER_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 28
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.color = m_CLR_WEISS
        .Borders.Weight = xlThin
    End With
    
    ' --- Datenzeilen ---
    Dim rowIdx As Long
    rowIdx = DASH_MATRIX_START_ROW
    
    kpiSummeIst = 0
    kpiSummeSoll = 0
    kpiSummeSaeumnis = 0
    kpiAnzahlOffen = 0
    kpiAnzahlBezahlt = 0
    kpiAnzahlSaeumnis = 0
    kpiOffenOhneSoll = 0
    kpiOffenBetrag = 0
    
    Dim p As Long
    For p = 1 To anzParz
        ' Parzelle + Mitglied(er)
        With ws.Cells(rowIdx, 1)
            .value = parzellen(p).parzNr
            .Font.Name = "Calibri"
            .Font.Size = 10
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        With ws.Cells(rowIdx, 2)
            .value = parzellen(p).mitgliedNamen
            .Font.Name = "Calibri"
            .Font.Size = 10
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        
        Dim zeileSoll As Double
        Dim zeileIst As Double
        zeileSoll = 0
        zeileIst = 0
        
        ' EntityKeys und Rollen dieser Parzelle
        Dim eKeys() As String
        eKeys = Split(parzellen(p).entityKeys, ",")
        Dim eRollen() As String
        eRollen = Split(parzellen(p).roles, ",")
        ' v5.2: Eintrittsdaten parallel zu EntityKeys
        Dim eEintritte() As String
        eEintritte = Split(parzellen(p).eintritte, ",")
        Dim alleRollen As String
        alleRollen = UCase(parzellen(p).roles)
        
        ' Punkt 12: Aggregator fuer Jahresposten-Sammelspalte (pro Mitgliederzeile)
        Dim jpFaellig As Long: jpFaellig = 0
        Dim jpBezahlt As Long: jpBezahlt = 0
        Dim jpSoll As Double: jpSoll = 0
        Dim jpIst As Double: jpIst = 0
        Dim jpHatRot As Boolean: jpHatRot = False
        Dim jpHatGelb As Boolean: jpHatGelb = False
        
        For k = 0 To anzKat - 1
            Dim kategorie As String
            kategorie = kategorien(k).Name
            Dim katCol As Long
            katCol = katSpalte(k)   ' Punkt 12: 0 wenn jaehrlich, sonst echte Spalte
            
            ' OHNE PACHT: nur Mitgliedsbeitrag
            Dim istNurMitgliedsbeitrag As Boolean
            istNurMitgliedsbeitrag = (InStr(alleRollen, "MIT PACHT") = 0 And _
                                      InStr(alleRollen, "OHNE PACHT") > 0 And _
                                      StrComp(kategorie, "Mitgliedsbeitrag", vbTextCompare) <> 0)
            If istNurMitgliedsbeitrag Then
                If katCol > 0 Then Call SchreibeNichtAnwendbar(ws, rowIdx, katCol)
                GoTo NextKatDash
            End If
            
            Dim istMB As Boolean
            istMB = (StrComp(kategorie, "Mitgliedsbeitrag", vbTextCompare) = 0)
            
            ' Ehrenmitglied-Pruefung fuer Mitgliedsbeitrag
            If istMB Then
                Dim mbZahler As Long: mbZahler = 0
                Dim mbEhren As Long: mbEhren = 0
                Dim eChk As Long
                For eChk = LBound(eKeys) To UBound(eKeys)
                    Dim chkRole As String: chkRole = ""
                    If eChk <= UBound(eRollen) Then chkRole = UCase(Trim(eRollen(eChk)))
                    If InStr(chkRole, "EHREN") > 0 Then
                        mbEhren = mbEhren + 1
                    Else
                        mbZahler = mbZahler + 1
                    End If
                Next eChk
                
                If mbZahler = 0 And mbEhren > 0 Then
                    ' Alle Ehrenmitglieder -> Befreit (gruen)
                    If katCol > 0 Then
                    With ws.Cells(rowIdx, katCol)
                        .value = ChrW(10004) & " Befreit"
                        .Font.Name = "Calibri"
                        .Font.Size = 9
                        .Font.color = m_CLR_TEXT_GRUEN
                        .Font.Italic = True
                        .Font.Bold = True
                        .Interior.color = m_CLR_ZELLE_GRUEN
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With
                    End If
                    GoTo NextKatDash
                End If
            End If
            
            ' --- Monatsschleife ---
            Dim bezahltMonate As Long: bezahltMonate = 0
            Dim faelligMonate As Long: faelligMonate = 0
            Dim katSoll As Double: katSoll = 0
            Dim katIst As Double: katIst = 0
            Dim katSaeumnis As Double: katSaeumnis = 0
            Dim katHatRot As Boolean: katHatRot = False
            Dim katHatGelb As Boolean: katHatGelb = False
            
            Dim monat As Long
            For monat = 1 To 12
                If Not IstKatImMonatFaellig(kategorien(k), monat) Then
                    GoTo NextMonatDash
                End If
                
                ' v5.4: Monat nur relevant wenn CSV-Import vorliegt
                ' (keine Frist-basierte Fallback-Logik mehr)
                If Not importierteMonate(monat) Then GoTo NextMonatDash
                
                faelligMonate = faelligMonate + 1
                
                ' PruefeZahlungen fuer EntityKeys
                Dim mIst As Double: mIst = 0
                Dim mSoll As Double: mSoll = 0
                Dim mBestStatus As String: mBestStatus = "ROT"
                Dim mBem As String: mBem = ""
                
                Dim eIdx As Long
                For eIdx = LBound(eKeys) To UBound(eKeys)
                    Dim ek As String
                    ek = Trim(eKeys(eIdx))
                    If ek = "" Then GoTo NextEKDash
                    
                    ' MB: Ehrenmitglieder ueberspringen
                    If istMB Then
                        Dim eRole As String: eRole = ""
                        If eIdx <= UBound(eRollen) Then eRole = UCase(Trim(eRollen(eIdx)))
                        If InStr(eRole, "EHREN") > 0 Then GoTo NextEKDash
                    End If

                    ' v7.4: Eintrittsdatum-Filter fuer ALLE Kategorien
                    ' Mitglied zahlt erst ab seinem Eintrittsmonat fuer alle
                    ' Gebuehren (Mitgliedsbeitrag, Pacht, Wasser-/Strom-Abschlag etc.)
                    If eIdx <= UBound(eEintritte) Then
                        Dim eEintritt As String
                        eEintritt = Trim(eEintritte(eIdx))
                        If Len(eEintritt) = 8 Then
                            Dim eJahr As Long, eMon As Long
                            eJahr = val(Left(eEintritt, 4))
                            eMon = val(Mid(eEintritt, 5, 2))
                            If eJahr = jahr And monat < eMon Then GoTo NextEKDash
                        End If
                    End If
                    
                    Dim ergebnis As String
                    ergebnis = mod_Zahlungspruefung.PruefeZahlungen(ek, kategorie, monat, jahr)
                    
                    ' Parsen: "STATUS|Soll:X.XX|Ist:Y.YY|Bemerkung"
                    Dim teile() As String
                    teile = Split(ergebnis, "|")
                    Dim tmpS As String: tmpS = "ROT"
                    Dim tmpSoll As Double: tmpSoll = 0
                    Dim tmpIst As Double: tmpIst = 0
                    Dim tmpBem As String: tmpBem = ""
                    
                    If UBound(teile) >= 2 Then
                        tmpS = teile(0)
                        Dim sT() As String: sT = Split(teile(1), ":")
                        If UBound(sT) >= 1 Then tmpSoll = val(sT(1))
                        Dim iT() As String: iT = Split(teile(2), ":")
                        If UBound(iT) >= 1 Then tmpIst = val(iT(1))
                    ElseIf UBound(teile) >= 0 Then
                        tmpS = teile(0)
                    End If
                    If UBound(teile) >= 3 Then tmpBem = teile(3)
                    
                    If istMB Then
                        ' Per-Mitglied: Summieren
                        mSoll = mSoll + tmpSoll
                        mIst = mIst + tmpIst
                    Else
                        ' Per-Parzelle: Bestes Ergebnis
                        If tmpIst > mIst Then mIst = tmpIst
                        If tmpSoll > mSoll Then mSoll = tmpSoll
                        If StrComp(tmpS, statusGruen, vbTextCompare) = 0 Then
                            mBestStatus = statusGruen
                        ElseIf StrComp(tmpS, "GELB", vbTextCompare) = 0 Then
                            If StrComp(mBestStatus, statusGruen, vbTextCompare) <> 0 Then
                                mBestStatus = "GELB"
                            End If
                        End If
                    End If
                    If tmpBem <> "" Then mBem = tmpBem
                    
NextEKDash:
                Next eIdx
                
                ' v5.4: MB-Soll anpassen fuer Mitglieder ohne eigenen EntityKey
                ' Wenn mehr Mitglieder auf der Parzelle sind als zahlende EntityKeys,
                ' muss der Soll auf die tatsaechliche Mitgliederzahl hochgerechnet werden.
                If istMB Then
                    Dim tatsaechlicheMB As Long
                    tatsaechlicheMB = parzellen(p).anzMitglieder - mbEhren
                    If tatsaechlicheMB < 1 Then tatsaechlicheMB = 1
                    If tatsaechlicheMB > mbZahler And kategorien(k).SollBetrag > 0 Then
                        mSoll = kategorien(k).SollBetrag * tatsaechlicheMB
                    End If
                End If
                
                ' MB: Status aus Summen berechnen
                If istMB Then
                    If mSoll > 0 And mIst >= mSoll - 0.01 Then
                        mBestStatus = statusGruen
                    ElseIf mIst > 0 Then
                        mBestStatus = "GELB"
                    Else
                        mBestStatus = "ROT"
                    End If
                End If
                
                ' Soll aus Uebersicht-Blatt nachladen (nur per-Parzelle)
                If Not istMB And mSoll = 0 Then
                    If Not sollDict Is Nothing Then
                        Dim uKey As String
                        uKey = CStr(parzellen(p).parzNr) & "|" & kategorie
                        If sollDict.exists(uKey) Then
                            mSoll = CDbl(sollDict(uKey))
                            If mIst >= mSoll - 0.01 Then
                                mBestStatus = statusGruen
                            ElseIf mIst > 0 Then
                                mBestStatus = "GELB"
                            Else
                                mBestStatus = "ROT"
                            End If
                        End If
                    End If
                End If
                
                ' Keine Saeumnis -> ROT wird zu GELB herabgestuft
                If StrComp(mBestStatus, "ROT", vbTextCompare) = 0 Then
                    If kategorien(k).saeumnisGebuehr = 0 Then
                        mBestStatus = "GELB"
                    End If
                End If
                
                ' Aggregieren
                katSoll = katSoll + mSoll
                katIst = katIst + mIst
                
                If StrComp(mBestStatus, statusGruen, vbTextCompare) = 0 Then
                    bezahltMonate = bezahltMonate + 1
                    kpiAnzahlBezahlt = kpiAnzahlBezahlt + 1
                ElseIf StrComp(mBestStatus, "GELB", vbTextCompare) = 0 Then
                    katHatGelb = True
                    If mIst > 0 Then
                        bezahltMonate = bezahltMonate + 1
                        kpiAnzahlBezahlt = kpiAnzahlBezahlt + 1
                    Else
                        kpiAnzahlOffen = kpiAnzahlOffen + 1
                        If mSoll = 0 Then kpiOffenOhneSoll = kpiOffenOhneSoll + 1
                        ' v5.3: Offenen Betrag akkumulieren
                        Dim offenPosten As Double
                        offenPosten = mSoll
                        If mSoll = 0 And kategorien(k).SollBetrag > 0 Then
                            If istMB Then
                                offenPosten = kategorien(k).SollBetrag * mbZahler
                            Else
                                offenPosten = kategorien(k).SollBetrag
                            End If
                        End If
                        If offenPosten > 0 Then kpiOffenBetrag = kpiOffenBetrag + offenPosten
                    End If
                Else
                    katHatRot = True
                    kpiAnzahlOffen = kpiAnzahlOffen + 1
                    If mSoll = 0 Then kpiOffenOhneSoll = kpiOffenOhneSoll + 1
                    
                    ' v5.3: Offenen Betrag akkumulieren
                    offenPosten = mSoll - mIst
                    If mSoll = 0 And kategorien(k).SollBetrag > 0 Then
                        If istMB Then
                            offenPosten = kategorien(k).SollBetrag * mbZahler - mIst
                        Else
                            offenPosten = kategorien(k).SollBetrag - mIst
                        End If
                    End If
                    If offenPosten > 0 Then kpiOffenBetrag = kpiOffenBetrag + offenPosten
                    
                    If kategorien(k).saeumnisGebuehr > 0 Then
                        katSaeumnis = katSaeumnis + kategorien(k).saeumnisGebuehr
                        kpiAnzahlSaeumnis = kpiAnzahlSaeumnis + 1
                    End If
                    
                    ' Verzug-Eintrag
                    If anzVerzug <= UBound(verzugListe) Then
                        With verzugListe(anzVerzug)
                            .parzNr = parzellen(p).parzNr
                            ' v5.1: Namen untereinander (vbLf) statt "/" getrennt
                            .mitglied = parzellen(p).mitgliedNamen
                            .kategorie = kategorie
                            .monatNr = monat
                            .monatText = MonthName(monat) & " " & jahr
                            .soll = mSoll
                            .ist = mIst
                            .differenz = mSoll - mIst
                            .saeumnis = kategorien(k).saeumnisGebuehr
                            .bemerkung = mBem
                            Dim sDat As Date, vl As Long, nl As Long, sg As Double
                            sDat = mod_Zahlungspruefung.BerechneSollDatumZP(kategorie, monat, jahr)
                            Call mod_Zahlungspruefung.HoleToleranzZP(kategorie, vl, nl, sg)
                            Dim fDat As Date
                            fDat = DateAdd("d", nl, sDat)
                            If Date > fDat Then
                                .tageVerzug = DateDiff("d", fDat, Date)
                            Else
                                .tageVerzug = 0
                            End If
                        End With
                        anzVerzug = anzVerzug + 1
                    End If
                End If
                
NextMonatDash:
            Next monat
            
            ' Zelle schreiben + KPI aggregieren
            kpiSummeSoll = kpiSummeSoll + katSoll
            kpiSummeIst = kpiSummeIst + katIst
            kpiSummeSaeumnis = kpiSummeSaeumnis + katSaeumnis
            zeileSoll = zeileSoll + katSoll
            zeileIst = zeileIst + katIst
            
            ' Punkt 12: Bei jaehrlichen Kategorien in Sammel-Aggregator statt eigene Zelle
            If katCol = 0 Then
                jpFaellig = jpFaellig + faelligMonate
                jpBezahlt = jpBezahlt + bezahltMonate
                jpSoll = jpSoll + katSoll
                jpIst = jpIst + katIst
                If katHatRot Then jpHatRot = True
                If katHatGelb Then jpHatGelb = True
                GoTo NextKatDash
            End If
            
            Call SchreibeMatrixZelle(ws, rowIdx, katCol, _
                                     faelligMonate, bezahltMonate, _
                                     katSoll, katIst, katHatRot, katHatGelb)
            
NextKatDash:
        Next k
        
        ' Punkt 12: Jahresposten-Sammelspalte schreiben
        If colJahresposten > 0 Then
            Call SchreibeMatrixZelle(ws, rowIdx, colJahresposten, _
                                     jpFaellig, jpBezahlt, _
                                     jpSoll, jpIst, jpHatRot, jpHatGelb, _
                                     alsPunkte:=True)
        End If
        
        ' Gesamt-Spalte
        With ws.Cells(rowIdx, colGesamt)
            .value = zeileIst
            .NumberFormat = "#,##0.00 " & ChrW(8364)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 10
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
        End With
        
        ' Quote-Spalte
        With ws.Cells(rowIdx, colGesamt + 1)
            .Font.Name = "Calibri"
            .Font.Size = 10
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            If zeileSoll > 0 Then
                Dim quote As Double
                quote = zeileIst / zeileSoll
                .value = quote
                .NumberFormat = "0%"
                If quote >= 1 Then
                    .Font.color = m_CLR_TEXT_GRUEN
                    .Interior.color = m_CLR_ZELLE_GRUEN
                ElseIf quote >= 0.5 Then
                    .Font.color = RGB(120, 100, 0)
                    .Interior.color = m_CLR_ZELLE_GELB
                Else
                    .Font.color = m_CLR_TEXT_DUNKELROT
                    .Interior.color = m_CLR_ZELLE_ROT
                End If
            Else
                .value = ChrW(8212)
                .Font.color = RGB(180, 180, 180)
                .Interior.color = m_CLR_ZELLE_GRAU
            End If
        End With
        
        ' Zebra-Streifen
        If p Mod 2 = 0 Then
            ws.Range(ws.Cells(rowIdx, 1), ws.Cells(rowIdx, 2)).Interior.color = RGB(245, 245, 250)
        End If
        
        ' Rahmen + Zeilenhoehe
        With ws.Range(ws.Cells(rowIdx, 1), ws.Cells(rowIdx, letzteSpalte))
            .Borders.LineStyle = xlContinuous
            .Borders.color = RGB(220, 220, 220)
            .Borders.Weight = xlThin
        End With
        Dim rowH As Long
        rowH = 13 * parzellen(p).anzMitglieder + 8
        If rowH < 26 Then rowH = 26
        ws.Rows(rowIdx).RowHeight = rowH
        
        rowIdx = rowIdx + 1
    Next p
    
    matrixEndRow = rowIdx - 1
    
    ' Summenzeile
    rowIdx = matrixEndRow + 1
    With ws.Cells(rowIdx, 2)
        .value = "SUMME"
        .Font.Bold = True
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    With ws.Cells(rowIdx, colGesamt)
        .value = kpiSummeIst
        .NumberFormat = "#,##0.00 " & ChrW(8364)
        .Font.Bold = True
        .Font.Name = "Calibri"
        .VerticalAlignment = xlCenter
    End With
    
    Dim gesamtQuote As Double
    If kpiSummeSoll > 0 Then
        gesamtQuote = kpiSummeIst / kpiSummeSoll
    Else
        gesamtQuote = 0
    End If
    With ws.Cells(rowIdx, colGesamt + 1)
        If kpiSummeSoll > 0 Then
            .value = gesamtQuote
            .NumberFormat = "0%"
        Else
            .value = ChrW(8212)
        End If
        .Font.Bold = True
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range(ws.Cells(rowIdx, 1), ws.Cells(rowIdx, letzteSpalte))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeTop).color = m_CLR_HEADER_BG
        .RowHeight = 24
    End With
    
    matrixEndRow = rowIdx
    
    ' v5.3: Leere Kategorie-Spalten ausblenden
    '   Spalte wird versteckt wenn ALLE Datenzellen nur ChrW(8212) enthalten
    Dim kCheck As Long
    For kCheck = 0 To anzKat - 1
        Dim prufCol As Long
        prufCol = 3 + kCheck
        Dim alleLeer As Boolean
        alleLeer = True
        Dim rCheck As Long
        For rCheck = DASH_MATRIX_START_ROW To matrixEndRow - 1
            Dim zellWert As String
            zellWert = CStr(ws.Cells(rCheck, prufCol).value)
            If zellWert <> ChrW(8212) Then
                alleLeer = False
                Exit For
            End If
        Next rCheck
        If alleLeer Then ws.Columns(prufCol).Hidden = True
    Next kCheck
    
    ' Datenbalken fuer Quote-Spalte
    On Error Resume Next
    Dim rngQ As Range
    Set rngQ = ws.Range(ws.Cells(DASH_MATRIX_START_ROW, colGesamt + 1), _
                         ws.Cells(matrixEndRow - 1, colGesamt + 1))
    rngQ.FormatConditions.Delete
    Dim db As Object
    Set db = rngQ.FormatConditions.AddDatabar
    If Not db Is Nothing Then
        db.BarColor.color = RGB(41, 128, 185)
        db.BarFillType = xlDataBarFillGradient
        db.MinPoint.Modify newtype:=xlConditionValueNumber, newValue:=0
        db.MaxPoint.Modify newtype:=xlConditionValueNumber, newValue:=1
        db.BarBorder.Type = xlDataBarBorderSolid
        db.BarBorder.color.color = RGB(41, 128, 185)
        db.ShowValue = True
    End If
    On Error GoTo 0
    
End Sub


' ============================================================
'  MATRIX-ZELLE SCHREIBEN
' ============================================================
Private Sub SchreibeMatrixZelle(ByVal ws As Worksheet, _
                                 ByVal zeile As Long, _
                                 ByVal spalte As Long, _
                                 ByVal faellig As Long, _
                                 ByVal bezahlt As Long, _
                                 ByVal soll As Double, _
                                 ByVal ist As Double, _
                                 ByVal hatRot As Boolean, _
                                 ByVal hatGelb As Boolean, _
                                 Optional ByVal alsPunkte As Boolean = False)
    
    With ws.Cells(zeile, spalte)
        .Font.Name = "Calibri"
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        Dim euroFmt As String
        euroFmt = "#,##0.00 " & ChrW(8364)
        
        ' PUNKT 12: Kreis-Anzeige fuer Sammelspalte Jahresposten
        Dim punkte As String: punkte = ""
        If alsPunkte And faellig > 0 Then
            Dim ii As Long
            For ii = 1 To faellig
                If ii <= bezahlt Then
                    punkte = punkte & ChrW(9679)   ' gefuellt
                Else
                    punkte = punkte & ChrW(9675)   ' leer
                End If
            Next ii
            punkte = punkte & " "
        End If
        
        If faellig = 0 Then
            .value = ChrW(8212)
            .Font.color = RGB(180, 180, 180)
            .Interior.color = m_CLR_ZELLE_GRAU
            
        ElseIf bezahlt >= faellig And Not hatRot Then
            If alsPunkte Then
                .value = punkte & CStr(bezahlt) & "/" & CStr(faellig) & " " & ChrW(8226) & " " & Format(ist, euroFmt)
            Else
                .value = ChrW(10004) & " " & Format(ist, euroFmt)
            End If
            .Font.color = m_CLR_TEXT_GRUEN
            .Font.Bold = True
            .Interior.color = m_CLR_ZELLE_GRUEN
            
        ElseIf hatRot And bezahlt = 0 Then
            If alsPunkte Then
                .value = punkte & "0/" & CStr(faellig) & " " & ChrW(8226) & " " & Format(soll, euroFmt)
            ElseIf faellig = 1 Then
                .value = ChrW(10008) & " offen " & Format(soll, euroFmt)
            Else
                .value = ChrW(10008) & " " & Format(soll, euroFmt)
            End If
            .Font.color = m_CLR_TEXT_DUNKELROT
            .Font.Bold = True
            .Interior.color = m_CLR_ZELLE_ROT
            
        ElseIf hatRot Then
            .value = punkte & CStr(bezahlt) & "/" & CStr(faellig) & " " & ChrW(8226) & " " & _
                     Format(ist, euroFmt)
            .Font.color = m_CLR_TEXT_DUNKELROT
            .Interior.color = m_CLR_ZELLE_ROT
            
        ElseIf hatGelb Then
            If alsPunkte Then
                .value = punkte & CStr(bezahlt) & "/" & CStr(faellig) & " " & ChrW(8226) & " " & Format(ist, euroFmt)
            ElseIf faellig = 1 Then
                .value = ChrW(9888) & " " & Format(ist, euroFmt)
            Else
                .value = CStr(bezahlt) & "/" & CStr(faellig) & " " & ChrW(8226) & " " & _
                         Format(ist, euroFmt)
            End If
            .Font.color = RGB(120, 100, 0)
            .Interior.color = m_CLR_ZELLE_GELB
            
        Else
            .value = Format(ist, euroFmt)
            .Font.color = m_CLR_TEXT_GRUEN
            .Interior.color = m_CLR_ZELLE_GRUEN
        End If
    End With
    
End Sub


' ============================================================
'  NICHT-ANWENDBAR ZELLE (grauer Strich)
' ============================================================
Private Sub SchreibeNichtAnwendbar(ByVal ws As Worksheet, _
                                    ByVal zeile As Long, _
                                    ByVal spalte As Long)
    With ws.Cells(zeile, spalte)
        .value = ChrW(8212)
        .Font.color = RGB(180, 180, 180)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.color = m_CLR_ZELLE_GRAU
    End With
End Sub


' ============================================================
'  PUNKT 12: JAEHRLICHE KATEGORIE ERKENNEN
' ============================================================
Public Function IstJaehrlicheKategorie(ByRef kat As UebKategorie) As Boolean
    Dim fl As String
    fl = LCase(Trim(kat.faelligkeit))
    If fl = "j" & ChrW(228) & "hrlich" Or fl = "jaehrlich" Or fl = "jahr" Or fl = "annual" Then
        IstJaehrlicheKategorie = True
        Exit Function
    End If
    
    ' SollMonate hat genau einen Monat -> jaehrlich
    If kat.SollMonate <> "" Then
        Dim teile() As String
        teile = Split(kat.SollMonate, ",")
        If UBound(teile) = 0 Then
            IstJaehrlicheKategorie = True
            Exit Function
        End If
    End If
    
    ' Klassiker als Fallback
    Dim nm As String: nm = LCase(Trim(kat.Name))
    Select Case nm
        Case "pacht", "betriebskosten", "endabrechnung", "fixkosten"
            IstJaehrlicheKategorie = True
        Case Else
            IstJaehrlicheKategorie = False
    End Select
End Function


' ============================================================
'  FAELLIGKEIT PRUEFEN
' ============================================================
Public Function IstKatImMonatFaellig(ByRef kat As UebKategorie, _
                                       ByVal monat As Long) As Boolean
    If kat.SollMonate <> "" Then
        IstKatImMonatFaellig = mod_KategorieEngine_Zeitraum.IstMonatInListe(monat, kat.SollMonate)
        Exit Function
    End If
    
    Dim fl As String
    fl = kat.faelligkeit
    If fl = "" Or fl = "monatlich" Then
        IstKatImMonatFaellig = True
        Exit Function
    End If
    
    IstKatImMonatFaellig = False
End Function


' ============================================================
'  VERZUG SORTIEREN (Bubble Sort nach TageVerzug absteigend)
' ============================================================
Public Sub SortiereVerzug(ByRef liste() As VerzugEintrag, ByVal anzahl As Long)
    Dim i As Long, j As Long
    Dim temp As VerzugEintrag
    
    For i = 0 To anzahl - 2
        For j = 0 To anzahl - 2 - i
            If liste(j).tageVerzug < liste(j + 1).tageVerzug Then
                temp = liste(j)
                liste(j) = liste(j + 1)
                liste(j + 1) = temp
            End If
        Next j
    Next i
End Sub


' ============================================================
'  VERZUGSDETAIL SCHREIBEN
' ============================================================
Public Sub SchreibeVerzugsdetail(ByVal ws As Worksheet, _
                                    ByVal startRow As Long, _
                                    ByRef liste() As VerzugEintrag, _
                                    ByVal anzahl As Long, _
                                    ByRef endRow As Long)
    
    Dim titelCol As Long
    titelCol = ws.Cells(DASH_TITEL_ROW, 1).MergeArea.Columns.count
    If titelCol < 10 Then titelCol = 10
    
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, titelCol))
        .Merge
        .value = ChrW(9632) & " VERZUGSDETAIL " & ChrW(8212) & _
                 " OFFENE ZAHLUNGEN (" & anzahl & ")"
        .Font.Name = "Calibri"
        .Font.Size = 13
        .Font.Bold = True
        .Font.color = m_CLR_KPI_ROT
        .VerticalAlignment = xlCenter
        .RowHeight = 28
    End With
    
    Dim hRow As Long
    hRow = startRow + 1
    
    Dim headers As Variant
    headers = Array("Parzelle", "Mitglied(er)", "Kategorie", "Monat", _
                    "Soll", "Ist", "Differenz", _
                    "S" & ChrW(228) & "umnis", "Tage Verzug", "Bemerkung")
    
    Dim c As Long
    For c = 0 To 9
        ws.Cells(hRow, c + 1).value = headers(c)
    Next c
    
    With ws.Range(ws.Cells(hRow, 1), ws.Cells(hRow, 10))
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .Font.color = m_CLR_WEISS
        .Interior.color = m_CLR_KPI_ROT
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 24
        .Borders.LineStyle = xlContinuous
        .Borders.color = m_CLR_WEISS
        .Borders.Weight = xlThin
    End With
    
    Dim dRow As Long
    dRow = hRow + 1
    
    Dim i As Long
    For i = 0 To anzahl - 1
        With liste(i)
            ws.Cells(dRow, 1).value = .parzNr
            ws.Cells(dRow, 2).value = .mitglied
            ws.Cells(dRow, 2).WrapText = True
            ws.Cells(dRow, 3).value = .kategorie
            ws.Cells(dRow, 4).value = .monatText
            ws.Cells(dRow, 5).value = .soll
            ws.Cells(dRow, 5).NumberFormat = "#,##0.00"
            ws.Cells(dRow, 6).value = .ist
            ws.Cells(dRow, 6).NumberFormat = "#,##0.00"
            ws.Cells(dRow, 7).value = .differenz
            ws.Cells(dRow, 7).NumberFormat = "#,##0.00"
            ws.Cells(dRow, 8).value = .saeumnis
            ws.Cells(dRow, 8).NumberFormat = "#,##0.00"
            ws.Cells(dRow, 9).value = .tageVerzug
            ws.Cells(dRow, 10).value = .bemerkung
        End With
        
        With ws.Range(ws.Cells(dRow, 1), ws.Cells(dRow, 10))
            .Font.Name = "Calibri"
            .Font.Size = 9
            .VerticalAlignment = xlCenter
            ' v5.1: Zeilenhoehe dynamisch je nach Anzahl Namen (vbLf-getrennt)
            Dim nameCount As Long
            nameCount = UBound(Split(liste(i).mitglied, vbLf)) + 1
            If nameCount > 1 Then
                .RowHeight = 13 * nameCount + 8
            Else
                .RowHeight = 22
            End If
            .Borders.LineStyle = xlContinuous
            .Borders.color = RGB(220, 220, 220)
            .Borders.Weight = xlThin
        End With
        
        ws.Cells(dRow, 1).HorizontalAlignment = xlCenter
        ws.Cells(dRow, 4).HorizontalAlignment = xlCenter
        ws.Cells(dRow, 9).HorizontalAlignment = xlCenter
        
        ' Farbe je nach Verzugstagen
        If liste(i).tageVerzug > 60 Then
            ws.Range(ws.Cells(dRow, 1), ws.Cells(dRow, 10)).Interior.color = RGB(255, 220, 220)
            ws.Cells(dRow, 9).Font.Bold = True
            ws.Cells(dRow, 9).Font.color = m_CLR_TEXT_DUNKELROT
        ElseIf liste(i).tageVerzug > 30 Then
            ws.Range(ws.Cells(dRow, 1), ws.Cells(dRow, 10)).Interior.color = m_CLR_ZELLE_ROT
        ElseIf liste(i).tageVerzug > 0 Then
            ws.Range(ws.Cells(dRow, 1), ws.Cells(dRow, 10)).Interior.color = m_CLR_ZELLE_GELB
        End If
        
        ' Balken-Effekt
        If liste(i).tageVerzug > 0 Then
            Dim bLen As Long
            bLen = liste(i).tageVerzug \ 10
            If bLen > 10 Then bLen = 10
            If bLen < 1 Then bLen = 1
            ws.Cells(dRow, 9).value = CStr(liste(i).tageVerzug) & " " & _
                                       Application.WorksheetFunction.Rept(ChrW(9608), bLen)
        End If
        
        If i Mod 2 = 1 And liste(i).tageVerzug = 0 Then
            ws.Range(ws.Cells(dRow, 1), ws.Cells(dRow, 10)).Interior.color = RGB(250, 250, 252)
        End If
        
        dRow = dRow + 1
    Next i
    
    endRow = dRow - 1
    
    ' Summenzeile
    If anzahl > 0 Then
        ws.Cells(dRow, 4).value = "SUMME:"
        ws.Cells(dRow, 4).Font.Bold = True
        ws.Cells(dRow, 4).HorizontalAlignment = xlRight
        ws.Cells(dRow, 4).Font.Name = "Calibri"
        ws.Cells(dRow, 4).VerticalAlignment = xlCenter
        
        ws.Cells(dRow, 5).Formula = "=SUM(" & ws.Cells(hRow + 1, 5).Address & _
                                    ":" & ws.Cells(dRow - 1, 5).Address & ")"
        ws.Cells(dRow, 5).NumberFormat = "#,##0.00"
        ws.Cells(dRow, 5).Font.Bold = True
        
        ws.Cells(dRow, 7).Formula = "=SUM(" & ws.Cells(hRow + 1, 7).Address & _
                                    ":" & ws.Cells(dRow - 1, 7).Address & ")"
        ws.Cells(dRow, 7).NumberFormat = "#,##0.00"
        ws.Cells(dRow, 7).Font.Bold = True
        ws.Cells(dRow, 7).Font.color = m_CLR_TEXT_DUNKELROT
        
        ws.Cells(dRow, 8).Formula = "=SUM(" & ws.Cells(hRow + 1, 8).Address & _
                                    ":" & ws.Cells(dRow - 1, 8).Address & ")"
        ws.Cells(dRow, 8).NumberFormat = "#,##0.00"
        ws.Cells(dRow, 8).Font.Bold = True
        ws.Cells(dRow, 8).Font.color = m_CLR_KPI_ORANGE
        
        With ws.Range(ws.Cells(dRow, 1), ws.Cells(dRow, 10))
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeTop).color = m_CLR_KPI_ROT
            .RowHeight = 22
        End With
        
        endRow = dRow
    End If
    
End Sub


' ============================================================
'  SPALTENBREITEN ANPASSEN
' ============================================================
Public Sub PasseSpaltenAn(ByVal ws As Worksheet, ByVal anzKat As Long)
    
    ws.Columns(1).ColumnWidth = 10   ' Parzelle
    ws.Columns(2).ColumnWidth = 26   ' Mitglied(er)
    
    ' PUNKT 12: Spalten anhand der Header-Zelle erkennen statt fixer Indizes
    Dim hdrRow As Long: hdrRow = DASH_MATRIX_HEADER_ROW
    Dim col As Long
    Dim maxCol As Long
    maxCol = ws.Cells(hdrRow, ws.Columns.count).End(xlToLeft).Column
    If maxCol < 3 Then maxCol = 3 + anzKat + 1
    
    For col = 3 To maxCol
        Dim hdr As String
        hdr = CStr(ws.Cells(hdrRow, col).value)
        Select Case hdr
            Case "Gesamt"
                ws.Columns(col).ColumnWidth = 16
            Case "Quote"
                ws.Columns(col).ColumnWidth = 10
            Case "Jahresposten"
                ws.Columns(col).ColumnWidth = 28
            Case Else
                ws.Columns(col).AutoFit
                If ws.Columns(col).ColumnWidth < 18 Then ws.Columns(col).ColumnWidth = 18
                ws.Columns(col).ColumnWidth = ws.Columns(col).ColumnWidth + 2
        End Select
    Next col
    
    ' Verzugsdetail: Bemerkungsspalte breiter
    On Error Resume Next
    If ws.Cells(ws.Rows.count, 10).End(xlUp).Row > DASH_MATRIX_START_ROW + 20 Then
        ws.Columns(10).ColumnWidth = 35
    End If
    On Error GoTo 0
    
End Sub








































































