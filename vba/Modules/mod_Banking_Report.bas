Attribute VB_Name = "mod_Banking_Report"
Option Explicit

' ===============================================================
' MODUL: mod_Banking_Report
' Ausgelagert aus mod_Banking_Data
' Enth?lt: Import Report ListBox (ActiveX) Verwaltung
'          Protokoll-Speicher (Daten!Y500), Farbcodierung
' ===============================================================

' Farb-Konstanten f?r ListBox-Hintergrund (OLE_COLOR / BGR)
Private Const LB_COLOR_GRUEN As Long = &HC0FFC0     ' hellgr?n
Private Const LB_COLOR_GELB As Long = &HC0FFFF      ' hellgelb
Private Const LB_COLOR_ROT As Long = &HC0C0FF       ' hellrot
Private Const LB_COLOR_WEISS As Long = &HFFFFFF     ' wei?

' Trennzeichen f?r Serialisierung in Zelle Y500
Private Const PROTO_SEP As String = "||"

' Protokoll-Speicher: Zelle Y500 auf dem Daten-Blatt
Private Const PROTO_ZEILE As Long = 500
Private Const PROTO_SPALTE As Long = 25              ' Spalte Y

' Maximale Anzahl Import-Bl?cke im Speicher (je 5 Zeilen)
Private Const MAX_BLOECKE As Long = 100
' 100 x 5 = 500 Zeilen maximal
Private Const MAX_ZEILEN As Long = 500


' ---------------------------------------------------------------
' Initialize: Liest Y500, bef?llt ActiveX ListBox,
'     setzt Hintergrundfarbe.
'     Aufruf: Workbook_Open, Worksheet_Activate, nach L?schen
' ---------------------------------------------------------------
Public Sub Initialize_ImportReport_ListBox()
    
    Dim wsBK As Worksheet
    Dim wsDaten As Worksheet
    Dim lb As MSForms.ListBox
    Dim oleObj As OLEObject
    Dim gespeichert As String
    Dim zeilen() As String
    Dim anzahl As Long
    Dim i As Long
    Dim savLeft As Double, savTop As Double
    Dim savWidth As Double, savHeight As Double
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsBK Is Nothing Or wsDaten Is Nothing Then Exit Sub
    
    ' OLEObject holen und Position/Gr??e VORHER sichern
    On Error Resume Next
    Set oleObj = wsBK.OLEObjects(FORM_LISTBOX_NAME)
    On Error GoTo 0
    If oleObj Is Nothing Then Exit Sub
    
    savLeft = oleObj.Left
    savTop = oleObj.Top
    savWidth = oleObj.Width
    savHeight = oleObj.Height
    
    ' Placement auf freifliegend setzen
    On Error Resume Next
    oleObj.Placement = xlFreeFloating
    On Error GoTo 0
    
    ' ActiveX ListBox holen
    On Error Resume Next
    Set lb = oleObj.Object
    On Error GoTo 0
    If lb Is Nothing Then Exit Sub
    
    ' ListBox leeren
    lb.Clear
    
    ' Gespeichertes Protokoll aus Y500 lesen
    gespeichert = CStr(wsDaten.Cells(PROTO_ZEILE, PROTO_SPALTE).value)
    
    If gespeichert = "" Or gespeichert = "0" Then
        ' Kein Protokoll vorhanden - Standardtext
        lb.AddItem "Kein Status Report"
        lb.AddItem "vorhanden."
        lb.BackColor = LB_COLOR_WEISS
    Else
        ' Protokoll-Zeilen aus Y500 deserialisieren und einf?gen
        zeilen = Split(gespeichert, PROTO_SEP)
        anzahl = UBound(zeilen) + 1
        If anzahl > MAX_ZEILEN Then anzahl = MAX_ZEILEN
        
        For i = 0 To anzahl - 1
            lb.AddItem zeilen(i)
        Next i
        
        ' Farbe aus j?ngstem Block bestimmen
        Call FaerbeListBoxAusProtokoll(lb, zeilen)
    End If
    
    ' Position und Gr??e WIEDERHERSTELLEN (AddItem kann sie ?ndern)
    On Error Resume Next
    oleObj.Left = savLeft
    oleObj.Top = savTop
    oleObj.Width = savWidth
    oleObj.Height = savHeight
    On Error GoTo 0
    
End Sub

' ---------------------------------------------------------------
' Update: Neuen 5-Zeilen-Block OBEN einf?gen,
'     in Y500 serialisiert speichern, ListBox aktualisieren.
' ---------------------------------------------------------------
Public Sub Update_ImportReport_ListBox(ByVal totalRows As Long, ByVal imported As Long, _
                                         ByVal dupes As Long, ByVal failed As Long)
    
    Dim wsBK As Worksheet
    Dim wsDaten As Worksheet
    Dim lb As MSForms.ListBox
    Dim oleObj As OLEObject
    Dim altGespeichert As String
    Dim neuerBlock As String
    Dim gesamt As String
    Dim zeilen() As String
    Dim anzahl As Long
    Dim i As Long
    Dim eventsWaren As Boolean
    Dim savLeft As Double, savTop As Double
    Dim savWidth As Double, savHeight As Double
    
    On Error Resume Next
    Set wsBK = ThisWorkbook.Worksheets(WS_BANKKONTO)
    Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
    On Error GoTo 0
    
    If wsBK Is Nothing Or wsDaten Is Nothing Then Exit Sub
    
    ' OLEObject holen und Position/Gr??e VORHER sichern
    On Error Resume Next
    Set oleObj = wsBK.OLEObjects(FORM_LISTBOX_NAME)
    On Error GoTo 0
    If oleObj Is Nothing Then Exit Sub
    
    savLeft = oleObj.Left
    savTop = oleObj.Top
    savWidth = oleObj.Width
    savHeight = oleObj.Height
    
    ' Placement auf freifliegend setzen
    On Error Resume Next
    oleObj.Placement = xlFreeFloating
    On Error GoTo 0
    
    ' --- 5-Zeilen-Block zusammenbauen ---
    neuerBlock = "Import: " & Format(Now, "DD.MM.YYYY  HH:MM:SS") & _
                 PROTO_SEP & _
                 imported & " / " & totalRows & " Datens?tze importiert" & _
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
    
    ' --- Daten-Blatt sch?tzen ---
    On Error Resume Next
    wsDaten.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
    ' --- Events wieder herstellen ---
    Application.EnableEvents = eventsWaren
    
    ' --- ActiveX ListBox aktualisieren ---
    On Error Resume Next
    Set lb = oleObj.Object
    On Error GoTo 0
    
    If Not lb Is Nothing Then
        lb.Clear
        zeilen = Split(gesamt, PROTO_SEP)
        For i = 0 To anzahl - 1
            lb.AddItem zeilen(i)
        Next i
        
        ' Farbcodierung
        Call FaerbeListBoxNachImport(lb, imported, dupes, failed)
    End If
    
    ' Position und Gr??e WIEDERHERSTELLEN (AddItem kann sie ?ndern)
    On Error Resume Next
    oleObj.Left = savLeft
    oleObj.Top = savTop
    oleObj.Width = savWidth
    oleObj.Height = savHeight
    On Error GoTo 0
    
End Sub

' ---------------------------------------------------------------
' Farbcodierung nach Import-Ergebnis (direkt auf ListBox)
'     GR?N   = Alles OK (dupes = 0, failed = 0)
'     GELB   = Duplikate vorhanden (dupes > 0, failed = 0)
'     ROT    = Fehler vorhanden (failed > 0)
' ---------------------------------------------------------------
Private Sub FaerbeListBoxNachImport(ByVal lb As MSForms.ListBox, _
                                     ByVal imported As Long, _
                                     ByVal dupes As Long, _
                                     ByVal failed As Long)
    
    If failed > 0 Then
        lb.BackColor = LB_COLOR_ROT
    ElseIf dupes > 0 Then
        lb.BackColor = LB_COLOR_GELB
    Else
        lb.BackColor = LB_COLOR_GRUEN
    End If
    
End Sub

' ---------------------------------------------------------------
' Farbcodierung aus gespeichertem Protokoll bestimmen
'     Liest Index 2: "X Duplikate erkannt"
'     Liest Index 3: "X Fehler"
' ---------------------------------------------------------------
Private Sub FaerbeListBoxAusProtokoll(ByVal lb As MSForms.ListBox, ByRef zeilen() As String)
    
    Dim dupes As Long
    Dim failed As Long
    
    If UBound(zeilen) < 3 Then
        lb.BackColor = LB_COLOR_WEISS
        Exit Sub
    End If
    
    dupes = ExtrahiereZahl(CStr(zeilen(2)))
    failed = ExtrahiereZahl(CStr(zeilen(3)))
    
    If failed > 0 Then
        lb.BackColor = LB_COLOR_ROT
    ElseIf dupes > 0 Then
        lb.BackColor = LB_COLOR_GELB
    Else
        lb.BackColor = LB_COLOR_GRUEN
    End If
    
End Sub

' ---------------------------------------------------------------
' Zahl am Anfang eines Strings extrahieren
'     "123 Duplikate erkannt" -> 123
' ---------------------------------------------------------------
Public Function ExtrahiereZahl(ByVal text As String) As Long
    
    Dim i As Long
    Dim zahlStr As String
    
    zahlStr = ""
    For i = 1 To Len(text)
        If Mid(text, i, 1) >= "0" And Mid(text, i, 1) <= "9" Then
            zahlStr = zahlStr & Mid(text, i, 1)
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


























