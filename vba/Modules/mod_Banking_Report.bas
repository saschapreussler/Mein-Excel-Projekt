Attribute VB_Name = "mod_Banking_Report"
Option Explicit

' ===============================================================
' MODUL: mod_Banking_Report
' Ausgelagert aus mod_Banking_Data
' enthält: Import Report ListBox (ActiveX) Verwaltung
'          Protokoll-Speicher (Daten!Y500), Farbcodierung
' ===============================================================

' Farb-Konstanten für ListBox-Hintergrund (OLE_COLOR / BGR)
Private Const LB_COLOR_GRUEN As Long = &HC0FFC0     ' hellgruen
Private Const LB_COLOR_GELB As Long = &HC0FFFF      ' hellgelb
Private Const LB_COLOR_ROT As Long = &HC0C0FF       ' hellrot
Private Const LB_COLOR_WEISS As Long = &HFFFFFF     ' weiss

' Trennzeichen für Serialisierung in Zelle Y500
Private Const PROTO_SEP As String = "||"

' Protokoll-Speicher: Zelle Y500 auf dem Daten-Blatt
Private Const PROTO_ZEILE As Long = 500
Private Const PROTO_SPALTE As Long = 25              ' Spalte Y

' Maximale Anzahl Import-Bloecke im Speicher (je 5 Zeilen)
Private Const MAX_BLOECKE As Long = 100
' 100 x 5 = 500 Zeilen maximal
Private Const MAX_ZEILEN As Long = 500


' ---------------------------------------------------------------
' Initialize: Liest Y500, beFuellt ActiveX ListBox,
'     setzt Hintergrundfarbe.
'     Aufruf: Workbook_Open, Worksheet_Activate, nach Löschen
' ---------------------------------------------------------------
Public Sub Initialize_ImportReport_ListBox()
    
    Dim wsBK As Worksheet
    Dim wsDaten As Worksheet
    Dim gespeichert As String
    Dim zeilen() As String
    Dim anzahl As Long
    Dim i As Long
    Dim reportText As String
    Dim reportFarbe As Long
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsBK Is Nothing Or wsDaten Is Nothing Then Exit Sub
    
    ' Blattschutz aufheben
    On Error Resume Next
    wsBK.Unprotect PASSWORD:=PASSWORD
    wsBK.OLEObjects(FORM_LISTBOX_NAME).Delete
    On Error GoTo 0
    reportText = ""
    reportFarbe = LB_COLOR_WEISS
    
    ' Wenn kein Bankkonto-Auszug vorhanden ist, immer den
    ' gewünschten Leerstates-Text anzeigen.
    If Not HatBankkontoDaten(wsBK) Then
        reportText = "Bisher kein Import durchgef" & ChrW(252) & "hrt."
        reportFarbe = LB_COLOR_WEISS
        Call SchreibeReportNachH8(wsBK, reportText, reportFarbe)

        On Error Resume Next
        Call mod_Uebersicht_Dashboard.LeereDashboardWennKeineDaten
        On Error GoTo 0

        Call mod_Banking_Format.Schuetze_BankkontoBlatt(wsBK)
        Exit Sub
    End If

    ' Gespeichertes Protokoll aus Y500 lesen
    gespeichert = CStr(wsDaten.Cells(PROTO_ZEILE, PROTO_SPALTE).value)
    
    If gespeichert = "" Or gespeichert = "0" Then
        reportText = "Bisher kein Import durchgef" & ChrW(252) & "hrt."
    Else
        ' Protokoll-Zeilen aus Y500 lesen (neuester Block zuerst)
        zeilen = Split(gespeichert, PROTO_SEP)
        anzahl = UBound(zeilen) + 1
        If anzahl > 5 Then anzahl = 5
        
        For i = 0 To anzahl - 1
            If reportText = "" Then
                reportText = zeilen(i)
            Else
                reportText = reportText & vbLf & zeilen(i)
            End If
        Next i

        reportFarbe = BestimmeReportFarbe(zeilen)
    End If

    Call SchreibeReportNachH8(wsBK, reportText, reportFarbe)

    ' Wenn weder Bankkonto noch Zahlungsübersicht Daten haben,
    ' darf auch das Dashboard keine alten Einträge mehr zeigen.
    On Error Resume Next
    Call mod_Uebersicht_Dashboard.LeereDashboardWennKeineDaten
    On Error GoTo 0
    
    ' Blattschutz wiederherstellen (mit AllowFiltering für Zeile 29)
    Call mod_Banking_Format.Schuetze_BankkontoBlatt(wsBK)
    
End Sub

' ---------------------------------------------------------------
' True, wenn mindestens eine Datenzeile im Bankkonto vorhanden ist.
' ---------------------------------------------------------------
Private Function HatBankkontoDaten(ByVal wsBK As Worksheet) As Boolean
    HatBankkontoDaten = False
    If wsBK Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = wsBK.Cells(wsBK.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then Exit Function

    Dim r As Long
    For r = BK_START_ROW To lastRow
        If Trim(CStr(wsBK.Cells(r, BK_COL_DATUM).value)) <> "" Or _
           Trim(CStr(wsBK.Cells(r, BK_COL_BETRAG).value)) <> "" Or _
           Trim(CStr(wsBK.Cells(r, BK_COL_VERWENDUNGSZWECK).value)) <> "" Then
            HatBankkontoDaten = True
            Exit Function
        End If
    Next r
End Function

' ---------------------------------------------------------------
' Update: Neuen 5-Zeilen-Block OBEN einfügen,
'     in Y500 serialisiert speichern, ListBox aktualisieren.
' ---------------------------------------------------------------
Public Sub Update_ImportReport_ListBox(ByVal totalRows As Long, ByVal imported As Long, _
                                         ByVal dupes As Long, ByVal failed As Long)
    
    Dim wsBK As Worksheet
    Dim wsDaten As Worksheet
    Dim altGespeichert As String
    Dim neuerBlock As String
    Dim gesamt As String
    Dim zeilen() As String
    Dim anzahl As Long
    Dim i As Long
    Dim eventsWaren As Boolean
    Dim reportText As String
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsBK Is Nothing Or wsDaten Is Nothing Then Exit Sub
    
    ' Blattschutz aufheben
    On Error Resume Next
    wsBK.Unprotect PASSWORD:=PASSWORD
    wsBK.OLEObjects(FORM_LISTBOX_NAME).Delete
    On Error GoTo 0
    
    ' --- 5-Zeilen-Block zusammenbauen ---
    neuerBlock = "Import: " & Format(Now, "DD.MM.YYYY  HH:MM:SS") & _
                 PROTO_SEP & _
                 imported & " / " & totalRows & " Datensätze importiert" & _
                 PROTO_SEP & _
                 dupes & " Duplikate erkannt" & _
                 PROTO_SEP & _
                 failed & " Fehler" & _
                 PROTO_SEP & _
                 "--------------------------------------"
    
    ' --- WICHTIG: Events deaktivieren BEVOR in Daten geschrieben wird ---
    eventsWaren = Application.EnableEvents
    Application.EnableEvents = False
    
    ' --- Daten-Blatt entsperren ---
    On Error Resume Next
    wsDaten.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0
    
    ' --- Alten Inhalt aus Y500 laden ---
    altGespeichert = CStr(wsDaten.Cells(PROTO_ZEILE, PROTO_SPALTE).value)
    
    If altGespeichert = "" Or altGespeichert = "0" Then
        gesamt = neuerBlock
    Else
        gesamt = neuerBlock & PROTO_SEP & altGespeichert
    End If
    
    ' --- Auf MAX_ZEILEN begrenzen ---
    zeilen = Split(gesamt, PROTO_SEP)
    anzahl = UBound(zeilen) + 1
    If anzahl > MAX_ZEILEN Then
        gesamt = zeilen(0)
        For i = 1 To MAX_ZEILEN - 1
            gesamt = gesamt & PROTO_SEP & zeilen(i)
        Next i
        anzahl = MAX_ZEILEN
    End If
    
    ' --- In Y500 speichern (eine einzige Zelle!) ---
    wsDaten.Cells(PROTO_ZEILE, PROTO_SPALTE).value = gesamt
    
    ' --- Daten-Blatt schützen ---
    On Error Resume Next
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
    ' --- Events wieder herstellen ---
    Application.EnableEvents = eventsWaren
    
    ' --- Report in Bankkonto H8 aktualisieren ---
    reportText = ""
    zeilen = Split(neuerBlock, PROTO_SEP)
    For i = 0 To UBound(zeilen)
        If reportText = "" Then
            reportText = zeilen(i)
        Else
            reportText = reportText & vbLf & zeilen(i)
        End If
    Next i
    Call SchreibeReportNachH8(wsBK, reportText, _
                              IIf(failed > 0, LB_COLOR_ROT, IIf(dupes > 0, LB_COLOR_GELB, LB_COLOR_GRUEN)))
    
    ' Blattschutz wiederherstellen (mit AllowFiltering für Zeile 29)
    Call mod_Banking_Format.Schuetze_BankkontoBlatt(wsBK)
    
End Sub

' ---------------------------------------------------------------
' Schreibt den Importstatus formatiert in Bankkonto!H8
' ---------------------------------------------------------------
Private Sub SchreibeReportNachH8(ByVal wsBK As Worksheet, _
                                 ByVal reportText As String, _
                                 ByVal bgColor As Long)
    With wsBK.Range("H8")
        .value = reportText
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .Interior.color = bgColor
        .Font.Bold = True
        .Font.Size = 9
    End With
    wsBK.Rows(8).RowHeight = 72
End Sub

' ---------------------------------------------------------------
' Farbe für den aktuellsten Report-Block bestimmen
' ---------------------------------------------------------------
Private Function BestimmeReportFarbe(ByRef zeilen() As String) As Long
    BestimmeReportFarbe = LB_COLOR_WEISS
    If UBound(zeilen) < 3 Then Exit Function
    
    Dim dupes As Long
    Dim failed As Long
    dupes = ExtrahiereZahl(CStr(zeilen(2)))
    failed = ExtrahiereZahl(CStr(zeilen(3)))
    
    If failed > 0 Then
        BestimmeReportFarbe = LB_COLOR_ROT
    ElseIf dupes > 0 Then
        BestimmeReportFarbe = LB_COLOR_GELB
    Else
        BestimmeReportFarbe = LB_COLOR_GRUEN
    End If
End Function

' ---------------------------------------------------------------
' Zahl am Anfang eines Strings extrahieren
'     "123 Duplikate erkannt" -> 123
' ---------------------------------------------------------------
Public Function ExtrahiereZahl(ByVal text As String) As Long
    
    Dim i As Long
    Dim zahlStr As String
    
    zahlStr = ""
    For i = 1 To Len(text)
        If mid(text, i, 1) >= "0" And mid(text, i, 1) <= "9" Then
            zahlStr = zahlStr & mid(text, i, 1)
        Else
            If zahlStr <> "" Then Exit For
        End If
    Next i
    
    If zahlStr <> "" Then
        ExtrahiereZahl = CLng(zahlStr)
    Else
        ExtrahiereZahl = 0
    End If
    
End Function


















































































































































