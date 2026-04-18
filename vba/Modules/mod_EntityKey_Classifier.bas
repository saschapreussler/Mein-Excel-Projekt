Attribute VB_Name = "mod_EntityKey_Classifier"
Option Explicit

' ***************************************************************
' MODUL: mod_EntityKey_Classifier
' ZWECK: Klassifizierung von Kontonamen (Shop/Versorger/Bank/etc.)
' ABGELEITET AUS: mod_EntityKey_Manager (Modularisierung)
' VERSION: 1.0 - 01.03.2026
' FUNKTIONEN:
'   - DarfParzelleHaben: Prueft ob Role eine Parzelle haben darf
'   - IstShop: Shop-Erkennung ueber Keyword-Liste
'   - ErmittleVersorgerZweck: Versorger-Erkennung mit Zweck-Rueckgabe
'   - IstBank: Bank-Erkennung ueber Keyword-Liste
'   - IstGeldautomatAbhebung: GA-Erkennung
'   - IstBankAbschluss: Bankabschluss-Erkennung
'   - ErmittleEntityRoleVonFunktion: Funktion -> EntityRole
' ***************************************************************

' ===============================================================
' Prueft ob Role eine Parzelle haben darf
' ===============================================================
Public Function DarfParzelleHaben(ByVal role As String) As Boolean
    Dim normRole As String
    
    If Trim(role) = "" Then
        DarfParzelleHaben = False
        Exit Function
    End If
    
    normRole = UCase(Trim(role))
    
    If InStr(normRole, "MITGLIED") > 0 Then
        DarfParzelleHaben = True
    ElseIf normRole = "SONSTIGE" Then
        DarfParzelleHaben = True
    Else
        DarfParzelleHaben = False
    End If
End Function

' ===============================================================
' IstShop - Keyword-Liste (~60 Eintraege)
' ===============================================================
Public Function IstShop(ByVal kontoname As String) As Boolean
    Dim n As String
    Dim keywords As Variant
    Dim i As Long
    
    n = UCase(Trim(kontoname))
    IstShop = False
    If Len(n) = 0 Then Exit Function
    
    keywords = Array( _
        "LIDL", "ALDI", "REWE", "EDEKA", "PENNY", "NETTO", "KAUFLAND", _
        "NORMA", "REAL", "ROSSMANN", "DM-DROGERIE", "MUELLER DROGERIE", _
        "BAUHAUS", "HORNBACH", "OBI", "HAGEBAU", "TOOM", "HELLWEG", _
        "GLOBUS BAUMARKT", "BAYWA", "RAIFFEISEN MARKT", _
        "AMAZON", "EBAY", "ZALANDO", "OTTO", "MEDIAMARKT", "SATURN", _
        "CONRAD ELECTRONIC", "ALTERNATE", "NOTEBOOKSBILLIGER", _
        "IKEA", "POCO", "ROLLER", "XXX LUTZ", _
        "DEHNER", "PFLANZEN KOELLE", "OVERKAMP", _
        "ARAL", "SHELL", "TOTAL", "ESSO", "JET TANKSTELLE", "TANKSTELLE", _
        "PAYPAL", "KLARNA", "SUMUP", _
        "FRESSNAPF", "ZOOPLUS", "DAS FUTTERHAUS", _
        "APOTHEKE", "FIELMANN", "APOLLO OPTIK", _
        "ACTION", "TEDI", "WOOLWORTH", "KIK", _
        "DECATHLON", "INTERSPORT", _
        "H&M", "C&A", "PRIMARK", "DEICHMANN" _
    )
    
    For i = LBound(keywords) To UBound(keywords)
        If InStr(n, CStr(keywords(i))) > 0 Then
            IstShop = True
            Exit Function
        End If
    Next i
End Function

' ===============================================================
' ErmittleVersorgerZweck - Gibt Versorger-Zweck zurueck oder ""
' ===============================================================
Public Function ErmittleVersorgerZweck(ByVal kontoname As String) As String
    Dim n As String
    
    n = UCase(Trim(kontoname))
    ErmittleVersorgerZweck = ""
    If Len(n) = 0 Then Exit Function
    
    ' --- Wasser / Abwasser ---
    If InStr(n, "WAZV") > 0 Then
        ErmittleVersorgerZweck = "Wasser/Abwasser Zweckverband"
        Exit Function
    End If
    If InStr(n, "BRAUCHWASSER") > 0 Or InStr(n, "EIGENBETRIEB") > 0 Then
        ErmittleVersorgerZweck = "Brauchwasserversorgung"
        Exit Function
    End If
    If InStr(n, "WASSER") > 0 Or InStr(n, "ABWASSER") > 0 Then
        ErmittleVersorgerZweck = "Wasser/Abwasser"
        Exit Function
    End If
    If InStr(n, "BWB") > 0 Or InStr(n, "BERLINER WASSERBETRIEBE") > 0 Then
        ErmittleVersorgerZweck = "Wasser/Abwasser"
        Exit Function
    End If
    If InStr(n, "ZWECKVERBAND") > 0 Then
        ErmittleVersorgerZweck = "Zweckverband"
        Exit Function
    End If
    
    ' --- Strom / Energie ---
    If InStr(n, "STADTWERK") > 0 Or InStr(n, "ENERGIE") > 0 Or InStr(n, "STROM") > 0 Then
        ErmittleVersorgerZweck = "Strom/Energie"
        Exit Function
    End If
    If InStr(n, "VATTENFALL") > 0 Or InStr(n, "E.ON") > 0 Or InStr(n, "EON") > 0 Then
        ErmittleVersorgerZweck = "Strom/Energie"
        Exit Function
    End If
    If InStr(n, "RWE") > 0 Or InStr(n, "ENVIA") > 0 Or InStr(n, "ENVIAM") > 0 Then
        ErmittleVersorgerZweck = "Strom/Energie"
        Exit Function
    End If
    If InStr(n, "ENBW") > 0 Or InStr(n, "MAINOVA") > 0 Or InStr(n, "ENTEGA") > 0 Then
        ErmittleVersorgerZweck = "Strom/Energie"
        Exit Function
    End If
    
    ' --- Gas / Heizung ---
    If InStr(n, "GASAG") > 0 Or InStr(n, "GAS") > 0 Then
        ErmittleVersorgerZweck = "Gas/Heizung"
        Exit Function
    End If
    If InStr(n, "FERNWAERME") > 0 Or InStr(n, "HEIZUNG") > 0 Then
        ErmittleVersorgerZweck = "Fernw" & ChrW(228) & "rme/Heizung"
        Exit Function
    End If
    
    ' --- Versicherung ---
    If InStr(n, "VERSICHERUNG") > 0 Or InStr(n, "ALLIANZ") > 0 Or InStr(n, "DEVK") > 0 Then
        ErmittleVersorgerZweck = "Versicherung"
        Exit Function
    End If
    If InStr(n, "HUK") > 0 Or InStr(n, "HDI") > 0 Or InStr(n, "ERGO") > 0 Then
        ErmittleVersorgerZweck = "Versicherung"
        Exit Function
    End If
    If InStr(n, "GENERALI") > 0 Or InStr(n, "AXA") > 0 Or InStr(n, "ZURICH") > 0 Then
        ErmittleVersorgerZweck = "Versicherung"
        Exit Function
    End If
    If InStr(n, "WUERTTEMBERGISCHE") > 0 Then
        ErmittleVersorgerZweck = "Versicherung"
        Exit Function
    End If
    
    ' --- Telekommunikation ---
    If InStr(n, "TELEKOM") > 0 Or InStr(n, "VODAFONE") > 0 Or InStr(n, "1&1") > 0 Then
        ErmittleVersorgerZweck = "Telekommunikation"
        Exit Function
    End If
    If InStr(n, "O2") > 0 Or InStr(n, "TELEFONICA") > 0 Then
        ErmittleVersorgerZweck = "Telekommunikation"
        Exit Function
    End If
    If InStr(n, "KABEL DEUTSCHLAND") > 0 Or InStr(n, "UNITYMEDIA") > 0 Then
        ErmittleVersorgerZweck = "Telekommunikation"
        Exit Function
    End If
    
    ' --- Abfall / Entsorgung ---
    If InStr(n, "BSR") > 0 Or InStr(n, "ENTSORGUNG") > 0 Or InStr(n, "STADTREINIGUNG") > 0 Then
        ErmittleVersorgerZweck = "Abfallwirtschaft/Entsorgung"
        Exit Function
    End If
    If InStr(n, "ABFALLWIRTSCHAFT") > 0 Or InStr(n, "ABFALL") > 0 Then
        ErmittleVersorgerZweck = "Abfallwirtschaft/Entsorgung"
        Exit Function
    End If
    If InStr(n, "LANDKREIS") > 0 Then
        ErmittleVersorgerZweck = "Abfallwirtschaft (Landkreis)"
        Exit Function
    End If
    
    ' --- Grundsteuer / Finanzamt ---
    If InStr(n, "GRUNDSTEUER") > 0 Or InStr(n, "FINANZAMT") > 0 Then
        ErmittleVersorgerZweck = "Grundsteuer/Steuern"
        Exit Function
    End If
    If InStr(n, "STADT WERDER") > 0 Or InStr(n, "STADT WERDER (HAVEL)") > 0 Then
        ErmittleVersorgerZweck = "Grundsteuer (Stadt)"
        Exit Function
    End If
    If InStr(n, "ABGABE") > 0 Then
        ErmittleVersorgerZweck = "Abgaben"
        Exit Function
    End If
    
    ' --- Rundfunk ---
    If InStr(n, "RUNDFUNK") > 0 Or InStr(n, "BEITRAGSSERVICE") > 0 Or InStr(n, "ARD ZDF") > 0 Then
        ErmittleVersorgerZweck = "Rundfunkbeitrag"
        Exit Function
    End If
    
    ' --- Verband ---
    If InStr(n, "VERBAND") > 0 Or InStr(n, "BEZIRKSVERBAND") > 0 Or InStr(n, "LANDESVERBAND") > 0 Then
        ErmittleVersorgerZweck = "Verband/Verb" & ChrW(228) & "nde"
        Exit Function
    End If
    If InStr(n, "VERPACHTUNG") > 0 Or InStr(n, "KLEINGARTENVERBAND") > 0 Then
        ErmittleVersorgerZweck = "Verpachtung/Kleingartenverband"
        Exit Function
    End If
    
    ' --- Miete / Grundstueck ---
    If InStr(n, "GRUNDSTUECKSGESELLSCHAFT") > 0 Or InStr(n, "GRUNDSTUCKSGESELLSCHAFT") > 0 Then
        ErmittleVersorgerZweck = "Grundst" & ChrW(252) & "cks-Miete"
        Exit Function
    End If
    If InStr(n, "HAUSVERWALTUNG") > 0 Then
        ErmittleVersorgerZweck = "Hausverwaltung/Miete"
        Exit Function
    End If
    If InStr(n, "HAUS- UND GRUNDSTUECK") > 0 Or InStr(n, "HAUS UND GRUNDSTUECK") > 0 Then
        ErmittleVersorgerZweck = "Grundst" & ChrW(252) & "cks-Miete"
        Exit Function
    End If
    If InStr(n, "HUG ") > 0 Or InStr(n, "H.U.G") > 0 Then
        ErmittleVersorgerZweck = "Grundst" & ChrW(252) & "cks-Miete"
        Exit Function
    End If
    If InStr(n, "MIETE") > 0 Or InStr(n, "MIETVERTRAG") > 0 Then
        ErmittleVersorgerZweck = "Miete"
        Exit Function
    End If
    If InStr(n, "PACHT") > 0 Then
        ErmittleVersorgerZweck = "Pacht"
        Exit Function
    End If
    
    ' Kein Versorger erkannt
    ErmittleVersorgerZweck = ""
End Function

' ===============================================================
' IstBank - Keyword-Liste
' ===============================================================
Public Function IstBank(ByVal kontoname As String) As Boolean
    Dim n As String
    Dim keywords As Variant
    Dim i As Long
    
    n = UCase(Trim(kontoname))
    IstBank = False
    If Len(n) = 0 Then Exit Function
    
    keywords = Array( _
        "SPARKASSE", "VOLKSBANK", "RAIFFEISENBANK", "RAIFFEISEN", _
        "COMMERZBANK", "DEUTSCHE BANK", "POSTBANK", _
        "ING DIBA", "ING-DIBA", "DKB", "DEUTSCHE KREDITBANK", _
        "COMDIRECT", "CONSORSBANK", "TARGOBANK", _
        "HYPOVEREINSBANK", "UNICREDIT", _
        "BERLINER BANK", "LANDESBANK", "SPARDA", _
        "PSD BANK", "NORISBANK", "N26", _
        "BANK" _
    )
    
    For i = LBound(keywords) To UBound(keywords)
        If InStr(n, CStr(keywords(i))) > 0 Then
            IstBank = True
            Exit Function
        End If
    Next i
End Function

' ===============================================================
' Prueft ob Kontoname eine Geldautomat-Abhebung ist
' Muster: IBAN="0", Name beginnt mit "GA " und enthaelt "BLZ"
' ===============================================================
Public Function IstGeldautomatAbhebung(ByVal iban As String, ByVal kontoname As String) As Boolean
    Dim normIBAN As String
    Dim nameUpper As String
    
    IstGeldautomatAbhebung = False
    normIBAN = mod_EntityKey_Normalize.NormalisiereIBAN(iban)
    
    If normIBAN <> "0" Then Exit Function
    
    nameUpper = UCase(Trim(kontoname))
    
    If Left(nameUpper, 3) = "GA " And InStr(nameUpper, "BLZ") > 0 Then
        IstGeldautomatAbhebung = True
    End If
End Function

' ===============================================================
' Prueft ob IBAN eine Bank-Abschluss-IBAN ist
' ===============================================================
Public Function IstBankAbschluss(ByVal iban As String, ByRef wsBK As Worksheet) As Boolean
    Dim normIBAN As String
    Dim r As Long
    Dim lastRow As Long
    Dim bkIBAN As String
    Dim buchungstext As String
    Dim bkKontoname As String
    
    IstBankAbschluss = False
    normIBAN = mod_EntityKey_Normalize.NormalisiereIBAN(iban)
    
    If normIBAN <> "0" And normIBAN <> "3529000972" Then Exit Function
    
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    
    For r = BK_START_ROW To lastRow
        bkIBAN = mod_EntityKey_Normalize.NormalisiereIBAN(wsBK.Cells(r, BK_COL_IBAN).value)
        If bkIBAN = normIBAN Then
            ' Geldautomat ausschliessen
            bkKontoname = Trim(CStr(wsBK.Cells(r, BK_COL_NAME).value))
            If IstGeldautomatAbhebung(CStr(wsBK.Cells(r, BK_COL_IBAN).value), bkKontoname) Then
                GoTo NaechsteZeile
            End If
            
            buchungstext = UCase(Trim(CStr(wsBK.Cells(r, BK_COL_BUCHUNGSTEXT).value)))
            If InStr(buchungstext, "ABSCHLUSS") > 0 Or _
               InStr(buchungstext, "ENTGELTABSCHLUSS") > 0 Then
                IstBankAbschluss = True
                Exit Function
            End If
        End If
NaechsteZeile:
    Next r
End Function

' ===============================================================
' Ermittelt EntityRole aus Funktion
' ===============================================================
Public Function ErmittleEntityRoleVonFunktion(ByVal funktion As String) As String
    Dim funktionUpper As String
    funktionUpper = UCase(funktion)
    
    If InStr(funktionUpper, "OHNE PACHT") > 0 Then
        ErmittleEntityRoleVonFunktion = "MITGLIED OHNE PACHT"
    ElseIf InStr(funktionUpper, "EHEMALIG") > 0 Then
        ErmittleEntityRoleVonFunktion = "EHEMALIGES MITGLIED"
    ElseIf InStr(funktionUpper, "EHREN") > 0 Then
        ErmittleEntityRoleVonFunktion = "EHRENMITGLIED"
    Else
        ErmittleEntityRoleVonFunktion = "MITGLIED MIT PACHT"
    End If
End Function














































































