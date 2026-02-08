VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ***************************************************************
' WORKSHEET: Tabelle8 (Daten)
' ZWECK: Automatische Formatierung bei Datenaenderungen
' VERSION: 4.0 - 08.02.2026
' FIX: OnDatenChange entfernt (war redundant mit FormatKategorieTableComplete)
' FIX: ShrinkToFit=False + Font.Color=vbBlack fuer Spalte L
' NEU: FindeKategorieZeile - Cursor bleibt nach Sort auf editierter Zeile
' NEU: Bei Spalte W (EntityRole) Aenderung -> Bankkonto neu evaluieren
' ***************************************************************

Private Sub Worksheet_Change(ByVal Target As Range)
    
    Dim changedCol As Long
    
    On Error GoTo ErrorHandler
    
    ' Nur Einzelzellen-Aenderungen verarbeiten
    If Target.Cells.CountLarge > 100 Then GoTo ErrorHandler
    If Target.Row < 4 Then GoTo ErrorHandler
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    changedCol = Target.Column
    
    ' === EINZELSPALTEN B, D, F, H ===
    If changedCol = 2 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 2)
        GoTo CleanUp
    End If
    
    If changedCol = 4 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 4)
        GoTo CleanUp
    End If
    
    If changedCol = 6 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 6)
        GoTo CleanUp
    End If
    
    If changedCol = 8 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, 8)
        GoTo CleanUp
    End If
    
    ' === KATEGORIE-TABELLE J-P (10-16) ===
    If changedCol >= 10 And changedCol <= 16 Then
        ' Merke aktuelle Kategorie + Keyword VOR der Sortierung
        Dim merkKategorie As String
        Dim merkKeyword As String
        merkKategorie = Trim(CStr(Me.Cells(Target.Row, DATA_CAT_COL_KATEGORIE).value))
        merkKeyword = Trim(CStr(Me.Cells(Target.Row, DATA_CAT_COL_KEYWORD).value))
        
        ' Formatierung + Sortierung
        Call mod_Formatierung.FormatKategorieTableComplete(Me)
        
        ' Konsistenzpruefung bei J, K oder L
        If changedCol = 10 Or changedCol = 11 Or changedCol = 12 Then
            Call mod_Formatierung.OnKategorieChange(Target)
        End If
        
        ' ShrinkToFit deaktivieren + Schriftfarbe schwarz fuer Spalte L
        Dim lastRowKat As Long
        lastRowKat = Me.Cells(Me.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
        If lastRowKat >= DATA_START_ROW Then
            Dim rngKeywords As Range
            Set rngKeywords = Me.Range(Me.Cells(DATA_START_ROW, DATA_CAT_COL_KEYWORD), _
                                       Me.Cells(lastRowKat, DATA_CAT_COL_KEYWORD))
            rngKeywords.ShrinkToFit = False
            rngKeywords.WrapText = False
            rngKeywords.Font.color = vbBlack
        End If
        
        ' Cursor auf die editierte Zeile zuruecksetzen (nach Sort)
        If merkKategorie <> "" And merkKeyword <> "" Then
            Dim gefundeneZeile As Long
            gefundeneZeile = FindeKategorieZeile(merkKategorie, merkKeyword)
            If gefundeneZeile > 0 Then
                ' Naechste leere oder aktuelle Spalte anwaehlen
                Dim zielSpalte As Long
                zielSpalte = changedCol
                If zielSpalte < DATA_CAT_COL_START Then zielSpalte = DATA_CAT_COL_START
                If zielSpalte > DATA_CAT_COL_END Then zielSpalte = DATA_CAT_COL_END
                
                Application.EnableEvents = True
                Application.ScreenUpdating = True
                Me.Cells(gefundeneZeile, zielSpalte).Select
                Application.EnableEvents = False
                Application.ScreenUpdating = False
            End If
        ElseIf merkKategorie <> "" Then
            ' Nur Kategorie bekannt (Keyword noch leer) -> erste passende Zeile finden
            Dim suchZeile As Long
            Dim lrSuch As Long
            lrSuch = Me.Cells(Me.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
            For suchZeile = DATA_START_ROW To lrSuch
                If StrComp(Trim(Me.Cells(suchZeile, DATA_CAT_COL_KATEGORIE).value), _
                           merkKategorie, vbTextCompare) = 0 Then
                    If Trim(Me.Cells(suchZeile, DATA_CAT_COL_KEYWORD).value) = "" Then
                        Application.EnableEvents = True
                        Application.ScreenUpdating = True
                        Me.Cells(suchZeile, DATA_CAT_COL_KEYWORD).Select
                        Application.EnableEvents = False
                        Application.ScreenUpdating = False
                        Exit For
                    End If
                End If
            Next suchZeile
        End If
        
        GoTo CleanUp
    End If
    
    ' === ENTITYKEY-TABELLE R-X (18-24) ===
    If changedCol >= 18 And changedCol <= 24 Then
        Call mod_Formatierung.FormatEntityKeyTableComplete(Me)
        If changedCol = 23 Then
            ' EntityRole wurde geaendert
            Call mod_EntityKey_Manager.VerarbeiteManuelleRoleAenderung(Target)
            
            ' NEU: Bankkonto-Kategorien fuer diese IBAN neu evaluieren
            ' Nur wenn eine Role tatsaechlich gesetzt wurde (nicht geloescht)
            Dim neueRole As String
            neueRole = Trim(CStr(Target.value))
            If neueRole <> "" Then
                Dim geaenderteIBAN As String
                geaenderteIBAN = Trim(CStr(Me.Cells(Target.Row, EK_COL_IBAN).value))
                If geaenderteIBAN <> "" Then
                    Call mod_KategorieEngine_Pipeline.ReEvaluiereNachEntityRoleAenderung(geaenderteIBAN)
                End If
            End If
        End If
        GoTo CleanUp
    End If
    
    ' === HELPER-SPALTEN Z bis AH (26-34) ===
    If changedCol >= 26 And changedCol <= 34 Then
        Call mod_Formatierung.FormatSingleColumnComplete(Me, changedCol)
        GoTo CleanUp
    End If

CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------
' Findet eine Kategorie-Zeile nach Sortierung anhand von
' Kategorie + Keyword (beide muessen uebereinstimmen)
' Gibt die Zeilennummer zurueck, oder 0 wenn nicht gefunden
' ---------------------------------------------------------------
Private Function FindeKategorieZeile(ByVal suchKategorie As String, _
                                      ByVal suchKeyword As String) As Long
    Dim lastRow As Long
    Dim r As Long
    
    FindeKategorieZeile = 0
    lastRow = Me.Cells(Me.Rows.count, DATA_CAT_COL_KATEGORIE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Function
    
    For r = DATA_START_ROW To lastRow
        If StrComp(Trim(Me.Cells(r, DATA_CAT_COL_KATEGORIE).value), _
                   suchKategorie, vbTextCompare) = 0 Then
            If StrComp(Trim(Me.Cells(r, DATA_CAT_COL_KEYWORD).value), _
                       suchKeyword, vbTextCompare) = 0 Then
                FindeKategorieZeile = r
                Exit Function
            End If
        End If
    Next r
End Function

