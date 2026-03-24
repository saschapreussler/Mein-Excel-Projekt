Attribute VB_Name = "mod_Uebersicht_Daten"
Option Explicit

' ***************************************************************
' MODUL: mod_Uebersicht_Daten
' VERSION: 1.2 - 16.03.2026
' ZWECK: Datenquellen und Hilfsfunktionen fuer die Uebersicht
'        - Kategorien aus Einstellungen laden (inkl. Faelligkeit)
'        - Aktive Mitglieder aus Daten-Blatt holen
'        - Jahr und importierte Monate aus Bankkonto ermitteln
'        - Vorjahr-Speicher (Okt-Dez Puffer auf Daten CA-CF)
' QUELLE: Extrahiert aus mod_Uebersicht_Generator v4.1
' NEU v1.1: HoleAktiveMitglieder gleicht Role live mit
'           Mitgliederliste Spalte O ab (Ehrenmitglied-Fix)
' NEU v1.2: Spalte B = Zuordnung (Spalte U) statt Kontoname (Spalte T)
' ***************************************************************


' ===============================================================
' Laedt Kategorien DYNAMISCH aus Einstellungen-Blatt
' Liest Spalte B (Kategorie), C (Soll-Betrag), E (Soll-Monate),
' I (Saeumnis-Gebuehr) + Faelligkeit aus Daten Spalte O
' Gibt eindeutige Kategorien zurueck (keine Duplikate)
' ===============================================================
Public Sub LadeKategorienAusEinstellungen(ByRef kategorien() As UebKategorie, _
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
    
    ' Dictionary fuer Eindeutigkeit
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Zuerst zaehlen fuer ReDim
    For r = ES_START_ROW To lastRow
        katName = Trim(CStr(wsEinst.Cells(r, ES_COL_KATEGORIE).value))
        If katName <> "" Then
            If Not dict.Exists(katName) Then
                dict.Add katName, r  ' Merke Zeilennummer fuer spaeteres Lesen
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
            
            ' Saeumnis-Gebuehr aus Spalte I
            Dim saeumnisWert As Variant
            saeumnisWert = wsEinst.Cells(r, ES_COL_SAEUMNIS).value
            If IsNumeric(saeumnisWert) Then
                .saeumnisGebuehr = CDbl(saeumnisWert)
            Else
                .saeumnisGebuehr = 0
            End If
            
            ' Soll-Monate aus Spalte E (z.B. "03, 06, 09" oder leer = alle)
            .SollMonate = Trim(CStr(wsEinst.Cells(r, ES_COL_SOLL_MONATE).value))
            
            ' Faelligkeit aus Daten-Blatt Spalte O (Kategorie-Tabelle)
            .faelligkeit = ""
        End With
        
        idx = idx + 1
    Next key
    
    ' Faelligkeit aus Daten-Blatt nachladen (Spalte O)
    Dim wsDatenKat As Worksheet
    On Error Resume Next
    Set wsDatenKat = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If Not wsDatenKat Is Nothing Then
        Dim lastRowDaten As Long
        lastRowDaten = wsDatenKat.Cells(wsDatenKat.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
        
        Dim ki As Long
        For ki = 0 To anzahl - 1
            Dim rD As Long
            For rD = DATA_START_ROW To lastRowDaten
                If StrComp(Trim(CStr(wsDatenKat.Cells(rD, DATA_CAT_COL_KATEGORIE).value)), _
                           kategorien(ki).Name, vbTextCompare) = 0 Then
                    kategorien(ki).faelligkeit = LCase(Trim(CStr( _
                        wsDatenKat.Cells(rD, DATA_CAT_COL_FAELLIGKEIT).value)))
                    Exit For
                End If
            Next rD
        Next ki
    End If
    
    Set dict = Nothing
    
End Sub


' ===============================================================
' Holt alle aktiven Mitglieder aus Daten-Blatt (EntityKey-Tabelle)
' Spalten: R=EntityKey, S=IBAN, T=Kontoname, U=Zuordnung, V=Parzelle, W=Role
' Bei SHARE-Keys koennen mehrere Parzellen in V stehen (z.B. "2, 5")
' Mehrere Mitglieder pro Parzelle erlaubt (z.B. MIT + OHNE PACHT)
' Dedup ueber EntityKey+Parzelle (nicht nur Parzelle)
' Name aus Spalte T (Kontoname), Fallback auf Spalte U (Zuordnung)
' v4.7: Role wird live aus Mitgliederliste Spalte O abgeglichen,
'       damit Ehren-/Funktions-Aenderungen sofort wirken
' ===============================================================
Public Function HoleAktiveMitglieder(ByVal wsDaten As Worksheet) As Collection
    
    Dim col As Collection
    Set col = New Collection
    
    Dim lastRow As Long
    lastRow = wsDaten.Cells(wsDaten.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
    
    If lastRow < EK_START_ROW Then
        Set HoleAktiveMitglieder = col
        Exit Function
    End If
    
    ' Dictionary fuer bereits verarbeitete EntityKey+Parzelle-Kombinationen
    Dim verarbeiteteKombis As Object
    Set verarbeiteteKombis = CreateObject("Scripting.Dictionary")
    
    ' v4.7: Funktions-Cache aus Mitgliederliste aufbauen
    ' Schluessel: EntityKey -> Funktion (Spalte O)
    ' Damit wird die Role live aktualisiert, z.B. bei "Ehrenmitglied"
    Dim funktionsCache As Object
    Set funktionsCache = CreateObject("Scripting.Dictionary")
    funktionsCache.CompareMode = vbTextCompare
    
    Dim wsML As Worksheet
    On Error Resume Next
    Set wsML = ThisWorkbook.Worksheets(WS_MITGLIEDER)
    On Error GoTo 0
    
    If Not wsML Is Nothing Then
        Dim lastRowML As Long
        lastRowML = wsML.Cells(wsML.Rows.count, M_COL_MEMBER_ID).End(xlUp).Row
        
        Dim rML As Long
        For rML = M_START_ROW To lastRowML
            Dim mlEntityKey As String
            mlEntityKey = Trim(CStr(wsML.Cells(rML, M_COL_ENTITY_KEY).value))
            If mlEntityKey <> "" Then
                Dim mlFunktion As String
                mlFunktion = Trim(CStr(wsML.Cells(rML, M_COL_FUNKTION).value))
                If mlFunktion <> "" Then
                    ' Nur den ersten Eintrag pro EntityKey verwenden
                    If Not funktionsCache.Exists(mlEntityKey) Then
                        funktionsCache.Add mlEntityKey, mlFunktion
                    End If
                End If
            End If
        Next rML
    End If
    
    Dim r As Long
    Dim entityKey As String
    Dim zuordnung As String
    Dim parzelleWert As String
    Dim roleWert As String
    Dim dict As Object
    
    For r = EK_START_ROW To lastRow
        entityKey = Trim(CStr(wsDaten.Cells(r, EK_COL_ENTITYKEY).value))
        If entityKey = "" Then GoTo NextDatenRow
        
        ' Role pruefen: nur aktive Mitglieder
        ' "MITGLIED MIT PACHT" und "MITGLIED OHNE PACHT" -> ja
        ' "EHEMALIGES MITGLIED" -> nein (ausschliessen)
        roleWert = UCase(Trim(CStr(wsDaten.Cells(r, EK_COL_ROLE).value)))
        
        ' v4.7: Role live aus Mitgliederliste aktualisieren
        ' Falls in Mitgliederliste Spalte O eine Aenderung erfolgte
        ' (z.B. "Ehrenmitglied"), wird die Role hier korrekt abgeleitet,
        ' auch wenn Spalte W auf dem Daten-Blatt noch den alten Wert hat.
        If funktionsCache.Exists(entityKey) Then
            Dim liveRole As String
            liveRole = UCase(mod_EntityKey_Classifier.ErmittleEntityRoleVonFunktion( _
                       funktionsCache(entityKey)))
            If liveRole <> roleWert Then
                Debug.Print "[" & ChrW(220) & "bersicht] Role-Update: " & entityKey & _
                            " W=" & roleWert & " -> ML=" & liveRole
                roleWert = liveRole
            End If
        End If
        
        If InStr(roleWert, "MITGLIED") = 0 Then GoTo NextDatenRow
        If InStr(roleWert, "EHEMALIGES") > 0 Then GoTo NextDatenRow
        
        ' Parzelle(n) lesen (kann "2" oder "2, 5" sein bei SHARE-Keys)
        parzelleWert = Trim(CStr(wsDaten.Cells(r, EK_COL_PARZELLE).value))
        If parzelleWert = "" Then GoTo NextDatenRow
        
        ' Zuordnung aus Spalte U - der zugeordnete Name
        zuordnung = Trim(CStr(wsDaten.Cells(r, EK_COL_ZUORDNUNG).value))
        ' Falls Zuordnung leer -> Fallback auf Kontoname (Spalte T)
        If zuordnung = "" Then
            zuordnung = Trim(CStr(wsDaten.Cells(r, EK_COL_KONTONAME).value))
        End If
        
        ' Parzelle(n) aufteilen (bei SHARE-Keys: "2, 5" -> 2 Eintraege)
        Dim parzellen() As String
        parzellen = Split(parzelleWert, ",")
        
        Dim p As Long
        For p = LBound(parzellen) To UBound(parzellen)
            Dim einzelParzelle As String
            einzelParzelle = Trim(parzellen(p))
            
            If IsNumeric(einzelParzelle) Then
                Dim parzelleNr As Long
                parzelleNr = CLng(einzelParzelle)
                
                ' Nur Parzellen 1-14
                If parzelleNr >= 1 And parzelleNr <= 14 Then
                    ' Duplikat-Pruefung: EntityKey+Parzelle nur einmal
                    Dim kombiKey As String
                    kombiKey = entityKey & "_" & parzelleNr
                    
                    If Not verarbeiteteKombis.Exists(kombiKey) Then
                        verarbeiteteKombis.Add kombiKey, True
                        
                        Set dict = CreateObject("Scripting.Dictionary")
                        dict.Add "Parzelle", parzelleNr
                        dict.Add "EntityKey", entityKey
                        dict.Add "Name", zuordnung
                        dict.Add "Role", roleWert
                        
                        col.Add dict
                    End If
                End If
            End If
        Next p
        
NextDatenRow:
    Next r
    
    Set verarbeiteteKombis = Nothing
    Set HoleAktiveMitglieder = col
    
End Function


' ===============================================================
' Ermittelt das haeufigste Jahr aus Bankkonto-Daten
' Scannt Spalte A (Datum) und zaehlt welches Jahr am meisten
' vorkommt. Gibt 0 zurueck wenn keine Daten vorhanden.
' ===============================================================
Public Function ErmittleJahrAusBankkonto() As Long
    
    Dim wsBK As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim zellWert As Variant
    Dim buchDatum As Date
    Dim jahrZaehler As Object
    Dim jahrKey As String
    
    ErmittleJahrAusBankkonto = 0
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo 0
    
    If wsBK Is Nothing Then Exit Function
    
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Function
    
    Set jahrZaehler = CreateObject("Scripting.Dictionary")
    
    For r = BK_START_ROW To lastRow
        zellWert = wsBK.Cells(r, BK_COL_DATUM).value
        
        If IsDate(zellWert) Then
            buchDatum = CDate(zellWert)
            jahrKey = CStr(Year(buchDatum))
            
            If jahrZaehler.Exists(jahrKey) Then
                jahrZaehler(jahrKey) = jahrZaehler(jahrKey) + 1
            Else
                jahrZaehler.Add jahrKey, 1
            End If
        End If
    Next r
    
    ' Haeufigtes Jahr finden
    If jahrZaehler.count = 0 Then
        Set jahrZaehler = Nothing
        Exit Function
    End If
    
    Dim maxAnzahl As Long
    Dim maxJahr As String
    Dim key As Variant
    maxAnzahl = 0
    
    For Each key In jahrZaehler.keys
        If jahrZaehler(key) > maxAnzahl Then
            maxAnzahl = jahrZaehler(key)
            maxJahr = CStr(key)
        End If
    Next key
    
    ErmittleJahrAusBankkonto = CLng(maxJahr)
    
    Debug.Print "[" & ChrW(220) & "bersicht] Jahr aus Bankkonto erkannt: " & maxJahr & _
                " (" & maxAnzahl & " Buchungen)"
    
    Set jahrZaehler = Nothing
    
End Function


' ===============================================================
' Ermittelt welche Monate im Bankkonto CSV-Daten haben
' Scannt Spalte A (Datum) ab BK_START_ROW und setzt True
' fuer jeden Monat der mindestens eine Buchung enthaelt
' Gibt Boolean-Array(1 To 12) zurueck
' ===============================================================
Public Function ErmittleImportierteMonate(ByVal jahr As Long) As Boolean()
    
    Dim result() As Boolean
    ReDim result(1 To 12)
    Dim wsBK As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim zellWert As Variant
    Dim buchDatum As Date
    Dim m As Long
    
    ' Array initialisieren (alles False - ReDim setzt bereits auf False)
    For m = 1 To 12
        result(m) = False
    Next m
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    On Error GoTo 0
    
    If wsBK Is Nothing Then
        Debug.Print "[" & ChrW(220) & "bersicht] WARNUNG: Blatt 'Bankkonto' nicht gefunden!"
        ErmittleImportierteMonate = result
        Exit Function
    End If
    
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    Debug.Print "[" & ChrW(220) & "bersicht] Bankkonto lastRow=" & lastRow & _
                " (BK_START_ROW=" & BK_START_ROW & ")"
    
    If lastRow < BK_START_ROW Then
        Debug.Print "[" & ChrW(220) & "bersicht] Keine Buchungen im Bankkonto gefunden."
        ErmittleImportierteMonate = result
        Exit Function
    End If
    
    For r = BK_START_ROW To lastRow
        zellWert = wsBK.Cells(r, BK_COL_DATUM).value
        
        If IsDate(zellWert) Then
            buchDatum = CDate(zellWert)
            
            If Year(buchDatum) = jahr Then
                result(Month(buchDatum)) = True
            End If
        End If
    Next r
    
    ErmittleImportierteMonate = result
    
End Function


' ===============================================================
' VORJAHR-SPEICHER: Okt-Dez des Vorjahres cachen
' Kopiert relevante Bankkonto-Buchungen (Okt-Dez Vorjahr) in den
' Hilfsspeicher auf Blatt Daten ab Spalte CA.
' Zweck: Dezember-Zahlungen die fuer Januar gelten erkennen
' ===============================================================
Public Sub BefuelleVorjahrSpeicher(ByVal vorjahr As Long)
    
    Dim wsBK As Worksheet
    Dim wsDaten As Worksheet
    Dim lastRowBK As Long
    Dim r As Long
    Dim vjRow As Long
    Dim zahlDatum As Date
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsBK Is Nothing Or wsDaten Is Nothing Then Exit Sub
    
    ' Zuerst alten Speicher loeschen
    Call LoescheVorjahrSpeicher
    
    ' Header setzen
    wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_DATUM).value = "VJ Datum"
    wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_BETRAG).value = "VJ Betrag"
    wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_IBAN).value = "VJ IBAN"
    wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_KATEGORIE).value = "VJ Kategorie"
    wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_MONAT_PERIODE).value = "VJ Monat/Periode"
    wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_ENTITYKEY).value = "VJ EntityKey"
    
    ' Header formatieren
    Dim rngVJHeader As Range
    Set rngVJHeader = wsDaten.Range(wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_DATUM), _
                                     wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_ENTITYKEY))
    rngVJHeader.Font.Bold = True
    rngVJHeader.Interior.color = RGB(217, 217, 217)
    
    lastRowBK = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    vjRow = VJ_START_ROW
    
    For r = BK_START_ROW To lastRowBK
        If Not IsDate(wsBK.Cells(r, BK_COL_DATUM).value) Then GoTo NextVJRow
        zahlDatum = CDate(wsBK.Cells(r, BK_COL_DATUM).value)
        
        ' Nur Okt-Dez des Vorjahres
        If Year(zahlDatum) <> vorjahr Then GoTo NextVJRow
        If Month(zahlDatum) < 10 Then GoTo NextVJRow
        
        ' Nur wenn Kategorie und IBAN vorhanden
        Dim vjKat As String
        vjKat = Trim(CStr(wsBK.Cells(r, BK_COL_KATEGORIE).value))
        If vjKat = "" Then GoTo NextVJRow
        
        Dim vjIBAN As String
        vjIBAN = Replace(Trim(CStr(wsBK.Cells(r, BK_COL_IBAN).value)), " ", "")
        If vjIBAN = "" Then GoTo NextVJRow
        
        ' In Speicher schreiben
        wsDaten.Cells(vjRow, VJ_COL_DATUM).value = zahlDatum
        wsDaten.Cells(vjRow, VJ_COL_DATUM).NumberFormat = "DD.MM.YYYY"
        wsDaten.Cells(vjRow, VJ_COL_BETRAG).value = wsBK.Cells(r, BK_COL_BETRAG).value
        wsDaten.Cells(vjRow, VJ_COL_IBAN).value = vjIBAN
        wsDaten.Cells(vjRow, VJ_COL_KATEGORIE).value = vjKat
        wsDaten.Cells(vjRow, VJ_COL_MONAT_PERIODE).value = _
            Trim(CStr(wsBK.Cells(r, BK_COL_MONAT_PERIODE).value))
        
        ' EntityKey via IBAN aufloesen (ueber EntityKey-Tabelle)
        Dim vjEK As String
        vjEK = ""
        Dim ek As Long
        Dim ekLastRow As Long
        ekLastRow = wsDaten.Cells(wsDaten.Rows.count, EK_COL_ENTITYKEY).End(xlUp).Row
        For ek = EK_START_ROW To ekLastRow
            Dim ekIBAN As String
            ekIBAN = Replace(Trim(CStr(wsDaten.Cells(ek, EK_COL_IBAN).value)), " ", "")
            If StrComp(ekIBAN, vjIBAN, vbTextCompare) = 0 Then
                vjEK = Trim(CStr(wsDaten.Cells(ek, EK_COL_ENTITYKEY).value))
                Exit For
            End If
        Next ek
        wsDaten.Cells(vjRow, VJ_COL_ENTITYKEY).value = vjEK
        
        vjRow = vjRow + 1
        
NextVJRow:
    Next r
    
    Debug.Print "[" & ChrW(220) & "bersicht] Vorjahr-Speicher: " & _
                (vjRow - VJ_START_ROW) & " Buchungen aus Okt-Dez " & vorjahr & " gecached"
    
End Sub


' ===============================================================
' Loescht den Vorjahr-Speicher auf Blatt Daten (ab CA)
' ===============================================================
Public Sub LoescheVorjahrSpeicher()
    
    Dim wsDaten As Worksheet
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = wsDaten.Cells(wsDaten.Rows.count, VJ_COL_DATUM).End(xlUp).Row
    
    If lastRow >= VJ_HEADER_ROW Then
        wsDaten.Range(wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_DATUM), _
                       wsDaten.Cells(lastRow, VJ_COL_ENTITYKEY)).ClearContents
        wsDaten.Range(wsDaten.Cells(VJ_HEADER_ROW, VJ_COL_DATUM), _
                       wsDaten.Cells(lastRow, VJ_COL_ENTITYKEY)).Interior.ColorIndex = xlNone
    End If
    
    Debug.Print "[" & ChrW(220) & "bersicht] Vorjahr-Speicher gel" & ChrW(246) & "scht"
    
End Sub


' ===============================================================
' Prueft automatisch ob Vorjahr-Speicher geloescht werden soll
' Ab August des Folgejahres wird der Speicher automatisch geleert
' ===============================================================
Public Sub PruefeVorjahrSpeicherAblauf()
    
    If Month(Date) >= 8 Then
        Dim wsDaten As Worksheet
        On Error Resume Next
        Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
        On Error GoTo 0
        
        If wsDaten Is Nothing Then Exit Sub
        
        ' Pruefen ob noch Daten im Speicher sind
        Dim ersteDatum As Variant
        ersteDatum = wsDaten.Cells(VJ_START_ROW, VJ_COL_DATUM).value
        
        If IsDate(ersteDatum) Then
            If Year(CDate(ersteDatum)) < Year(Date) - 1 Then
                ' Daten sind aelter als Vorjahr -> loeschen
                Call LoescheVorjahrSpeicher
            ElseIf Year(CDate(ersteDatum)) = Year(Date) - 1 Then
                ' Vorjahr-Daten und wir sind >= August -> loeschen
                Call LoescheVorjahrSpeicher
            End If
        End If
    End If
    
End Sub


' ===============================================================
' Holt Vorjahr-Zahlungsbetrag aus dem Speicher
' Prueft ob fuer den EntityKey + Kategorie eine Dezember-Zahlung
' vorliegt, die fuer Januar des Folgejahres gelten koennte
' (basierend auf Monat/Periode in Spalte CE)
' ===============================================================
Public Function HoleVorjahrZahlung(ByVal entityKey As String, _
                                    ByVal kategorie As String, _
                                    ByVal monat As Long) As Double
    HoleVorjahrZahlung = 0
    
    ' Nur fuer fruehe Monate relevant (Jan-Maerz)
    If monat > 3 Then Exit Function
    
    Dim wsDaten As Worksheet
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Function
    
    Dim lastRow As Long
    lastRow = wsDaten.Cells(wsDaten.Rows.count, VJ_COL_DATUM).End(xlUp).Row
    If lastRow < VJ_START_ROW Then Exit Function
    
    Dim r As Long
    Dim vjMonatPeriode As String
    Dim erwarteterMonat As String
    erwarteterMonat = MonthName(monat)
    
    For r = VJ_START_ROW To lastRow
        ' EntityKey pruefen
        If StrComp(Trim(CStr(wsDaten.Cells(r, VJ_COL_ENTITYKEY).value)), _
                   entityKey, vbTextCompare) <> 0 Then GoTo NextVJPruefRow
        
        ' Kategorie pruefen
        If StrComp(Trim(CStr(wsDaten.Cells(r, VJ_COL_KATEGORIE).value)), _
                   kategorie, vbTextCompare) <> 0 Then GoTo NextVJPruefRow
        
        ' Monat/Periode pruefen
        vjMonatPeriode = Trim(CStr(wsDaten.Cells(r, VJ_COL_MONAT_PERIODE).value))
        
        If StrComp(vjMonatPeriode, erwarteterMonat, vbTextCompare) = 0 Then
            ' Direkt-Match: Monat/Periode = "Januar"
            HoleVorjahrZahlung = HoleVorjahrZahlung + Abs(wsDaten.Cells(r, VJ_COL_BETRAG).value)
        End If
        
NextVJPruefRow:
    Next r
    
End Function


' ===============================================================
' v4.6: Prueft ob Vorjahr-Daten im Speicher vorhanden sind
' Gibt True zurueck wenn mindestens eine Zeile in CA-CF existiert
' ===============================================================
Public Function HatVorjahrDaten() As Boolean
    HatVorjahrDaten = False
    
    Dim wsDaten As Worksheet
    On Error Resume Next
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsDaten Is Nothing Then Exit Function
    
    Dim lastRow As Long
    lastRow = wsDaten.Cells(wsDaten.Rows.count, VJ_COL_DATUM).End(xlUp).Row
    HatVorjahrDaten = (lastRow >= VJ_START_ROW)
End Function


























