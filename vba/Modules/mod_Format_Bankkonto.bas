Attribute VB_Name = "mod_Format_Bankkonto"
Option Explicit

' ***************************************************************
' MODUL: mod_Format_Bankkonto
' ZWECK: Bankkonto-Blatt Formatierung + DropDown-Listen
' ABGELEITET AUS: mod_Formatierung (Modularisierung)
' VERSION: 1.0 - 01.03.2026
' FUNKTIONEN:
'   - FormatiereBlattBankkonto: Komplett-Formatierung Bankkonto
'   - NamedRangeExists: Prueft ob Named Range existiert
' ***************************************************************

' ===============================================================
' BANKKONTO-BLATT FORMATIEREN
' ===============================================================
Public Sub FormatiereBlattBankkonto()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim euroFormat As String
    Dim r As Long
    
    Set ws = ThisWorkbook.Worksheets(WS_BANKKONTO)
    
    euroFormat = "#,##0.00 " & ChrW(8364)
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo ErrorHandler
    
    ws.Cells.VerticalAlignment = xlCenter
    
    lastRow = ws.Cells(ws.Rows.count, BK_COL_DATUM).End(xlUp).Row
    If lastRow < BK_START_ROW Then lastRow = BK_START_ROW
    
    With ws.Range(ws.Cells(BK_START_ROW, BK_COL_BEMERKUNG), _
                  ws.Cells(lastRow, BK_COL_BEMERKUNG))
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    
    ws.Rows(BK_START_ROW & ":" & lastRow).AutoFit
    
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_BETRAG), _
             ws.Cells(lastRow, BK_COL_BETRAG)).NumberFormat = euroFormat
    
    ws.Range(ws.Cells(BK_START_ROW, BK_COL_MITGL_BEITR), _
             ws.Cells(lastRow, BK_COL_AUSZAHL_KASSE)).NumberFormat = euroFormat
    
    ' DropDown-Listen fuer Spalte H (Kategorie)
    Dim hatEinnahmen As Boolean
    Dim hatAusgaben As Boolean
    hatEinnahmen = NamedRangeExists("lst_KategorienEinnahmen")
    hatAusgaben = NamedRangeExists("lst_KategorienAusgaben")
    
    If hatEinnahmen Or hatAusgaben Then
        Dim wsDaten As Worksheet
        Set wsDaten = ThisWorkbook.Worksheets(WS_DATEN)
        
        Dim dictAlleKat As Object
        Set dictAlleKat = CreateObject("Scripting.Dictionary")
        
        Dim lastRowKat As Long
        lastRowKat = wsDaten.Cells(wsDaten.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
        
        Dim katName As String
        If lastRowKat >= DATA_START_ROW Then
            For r = DATA_START_ROW To lastRowKat
                katName = Trim(CStr(wsDaten.Cells(r, DATA_CAT_COL_KATEGORIE).value))
                If katName <> "" Then
                    If Not dictAlleKat.Exists(katName) Then
                        dictAlleKat.Add katName, katName
                    End If
                End If
            Next r
        End If
        
        Dim katListe As String
        katListe = ""
        Dim kk As Variant
        For Each kk In dictAlleKat.keys
            If katListe <> "" Then katListe = katListe & ","
            katListe = katListe & CStr(kk)
        Next kk
        
        If katListe <> "" Then
            For r = BK_START_ROW To lastRow
                On Error Resume Next
                ws.Cells(r, BK_COL_KATEGORIE).Validation.Delete
                On Error GoTo ErrorHandler
                
                With ws.Cells(r, BK_COL_KATEGORIE).Validation
                    .Add Type:=xlValidateList, _
                         AlertStyle:=xlValidAlertWarning, _
                         Formula1:=katListe
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = False
                    .ShowError = False
                End With
                
                ws.Cells(r, BK_COL_KATEGORIE).Locked = False
            Next r
        End If
    End If
    
    ' DropDown-Listen fuer Spalte I (Monat/Periode)
    Dim hatMonatListe As Boolean
    hatMonatListe = NamedRangeExists("lst_MonatPeriode")
    
    Dim monatFormel As String
    If hatMonatListe Then
        monatFormel = "=lst_MonatPeriode"
    Else
        monatFormel = "Januar,Februar,M" & ChrW(228) & "rz,April,Mai,Juni," & _
                      "Juli,August,September,Oktober,November,Dezember"
    End If
    
    For r = BK_START_ROW To lastRow
        On Error Resume Next
        ws.Cells(r, BK_COL_MONAT_PERIODE).Validation.Delete
        On Error GoTo ErrorHandler
        
        With ws.Cells(r, BK_COL_MONAT_PERIODE).Validation
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertWarning, _
                 Formula1:=monatFormel
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = False
        End With
        
        ws.Cells(r, BK_COL_MONAT_PERIODE).Locked = False
    Next r
    
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    Application.ScreenUpdating = True
    
    MsgBox "Formatierung des Bankkonto-Blatts abgeschlossen!" & vbCrLf & vbCrLf & _
           "- Alle Zellen vertikal zentriert" & vbCrLf & _
           "- Spalte L mit Textumbruch" & vbCrLf & _
           "- Zeilenh" & ChrW(246) & "he angepasst" & vbCrLf & _
           "- W" & ChrW(228) & "hrung mit Euro-Zeichen" & vbCrLf & _
           "- DropDown-Listen in Spalte H (Kategorie)" & vbCrLf & _
           "- DropDown-Listen in Spalte I (Monat/Periode)", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    MsgBox "Fehler bei der Formatierung: " & Err.Description, vbCritical
End Sub

' ===============================================================
' Prueft ob Named Range existiert
' ===============================================================
Public Function NamedRangeExists(ByVal rangeName As String) As Boolean
    Dim nm As Name
    NamedRangeExists = False
    
    On Error Resume Next
    Set nm = ThisWorkbook.Names(rangeName)
    If Not nm Is Nothing Then
        NamedRangeExists = True
    End If
    On Error GoTo 0
End Function




























































