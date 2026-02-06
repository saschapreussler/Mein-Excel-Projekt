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

'================================================================================
' TABELLE8 (DATEN) - WORKSHEET EVENT HANDLER
' VERSION: 2.0 - 06.02.2026
' ZWECK: Automatische Wartung aller Tabellen bei Datenaenderungen
'
' BEHANDELTE BEREICHE:
'   - B (col 2): Vereinsfunktionen - Zebra + Rahmen
'   - D (col 4): Anredeformen - Zebra + Rahmen
'   - F (col 6): Parzelle - Zebra + Rahmen
'   - H (col 8): Seite - Zebra + Rahmen
'   - J-P (10-16): Kategorie-Tabelle - Dropdown-Updates + Sortierung
'   - R-X (18-24): EntityKey-Tabelle - Role-Verarbeitung + Sortierung
'   - Z-AH (26-34): Helper-Spalten - Unabhaengige Formatierung pro Spalte
'
' WICHTIG:
'   - Spalten U-X erhalten NIE Zebra-Formatierung (Ampel-Bereich!)
'   - Helper-Spalten Z-AH werden als 9 unabhaengige Einheiten behandelt
'================================================================================

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rng_KategorieRegeln As Range
    Dim cat As String
    Dim r As Range
    Dim refValue As Variant
    Dim targetValue As Variant
    Dim zeile As Long
    Dim spalte As Long
    Dim needsKategorieSortAndFormat As Boolean
    Dim needsEntityKeySortAndFormat As Boolean

    On Error GoTo ExitPoint
    If Application.EnableEvents = False Then Exit Sub
    If Target.Cells.CountLarge <> 1 Then Exit Sub

    zeile = Target.Row
    spalte = Target.Column
    
    ' Nur Datenzeilen bearbeiten (ab Zeile 4)
    If zeile < DATA_START_ROW Then Exit Sub
    
    needsKategorieSortAndFormat = False
    needsEntityKeySortAndFormat = False
    
    ' =====================================================================
    ' ABSCHNITT 1: EINZELSPALTEN B, D, F, H (Zebra + Rahmen)
    ' =====================================================================
    If spalte = 2 Or spalte = 4 Or spalte = 6 Or spalte = 8 Then
        Application.EnableEvents = False
        Call mod_Formatierung.RefreshTableFormattingAfterChange(Me, spalte, zeile)
        Application.EnableEvents = True
        GoTo ExitPoint
    End If
    
    ' =====================================================================
    ' ABSCHNITT 2: ENTITYKEY-TABELLE (Spalten R-X = 18-24)
    ' =====================================================================
    If spalte >= EK_COL_ENTITYKEY And spalte <= EK_COL_DEBUG Then
        
        Application.EnableEvents = False
        
        ' Bei Role-Aenderung (Spalte W = 23): EntityKey anpassen
        If spalte = EK_COL_ROLE Then
            Call mod_EntityKey_Manager.VerarbeiteManuelleRoleAenderung(Target)
            ' Nach Role-Aenderung: Dropdown aktualisieren
            Call mod_Formatierung.UpdateEntityRoleDropdown(Me)
            needsEntityKeySortAndFormat = True
        End If
        
        ' Zeile formatieren (Stub in mod_EntityKey_Manager)
        Call mod_EntityKey_Manager.FormatiereEntityKeyZeile(zeile, Me)
        
        ' Bei EntityKey-Aenderung (Spalte R = 18): Sortierung triggern
        If spalte = EK_COL_ENTITYKEY Then
            needsEntityKeySortAndFormat = True
        End If
        
        ' Sortierung und Formatierung wenn erforderlich
        If needsEntityKeySortAndFormat Then
            Call mod_Formatierung.SortEntityKeyTable(Me)
        End If
        
        Application.EnableEvents = True
        GoTo ExitPoint
    End If
    
    ' =====================================================================
    ' ABSCHNITT 3: KATEGORIE-TABELLE (Spalten J-P = 10-16)
    ' =====================================================================
    If spalte >= DATA_CAT_COL_START And spalte <= DATA_CAT_COL_END Then
        
        Application.EnableEvents = False
        
        ' Bei Kategorie oder E/A Aenderung: Dropdowns aktualisieren
        If spalte = DATA_CAT_COL_KATEGORIE Or spalte = DATA_CAT_COL_EINAUS Then
            Call mod_Formatierung.OnKategorieChange(Target)
            needsKategorieSortAndFormat = True
        End If
        
        ' Sortierung und Formatierung wenn erforderlich
        If needsKategorieSortAndFormat Then
            Call mod_Formatierung.SortKategorieTable(Me)
            Call mod_Formatierung.FormatiereKategorieTabelle(Me)
        End If
        
        Application.EnableEvents = True
        
        ' === KATEGORIE-REGELN VALIDIERUNG ===
        Application.EnableEvents = True  ' Fuer MsgBox
        
        On Error Resume Next
        Set rng_KategorieRegeln = Me.Range("rng_KategorieRegeln")
        On Error GoTo ExitPoint
        
        If rng_KategorieRegeln Is Nothing Then GoTo ExitPoint
        
        Select Case spalte
            Case DATA_CAT_COL_EINAUS, DATA_CAT_COL_ZIELSPALTE, DATA_CAT_COL_FAELLIGKEIT
                If Intersect(Target, rng_KategorieRegeln) Is Nothing Then
                    GoTo ExitPoint
                End If
            Case Else
                GoTo ExitPoint
        End Select
        
        cat = Trim(CStr(Me.Cells(Target.Row, DATA_CAT_COL_KATEGORIE).value))
        If cat = "" Then GoTo ExitPoint
        
        targetValue = Me.Cells(Target.Row, Target.Column).value
        refValue = ""
        
        For Each r In Intersect(Me.Columns(DATA_CAT_COL_KATEGORIE), rng_KategorieRegeln).Cells
            
            If Trim(CStr(r.value)) = cat Then
                
                Dim currentValue As Variant
                currentValue = Me.Cells(r.Row, Target.Column).value
                
                If Not IsEmpty(currentValue) And Len(Trim(CStr(currentValue))) > 0 Then
                    If r.Row <> Target.Row Then
                        refValue = currentValue
                        Exit For
                    End If
                End If
            End If
        Next r
        
        If Not IsEmpty(refValue) And Len(Trim(CStr(refValue))) > 0 Then
            If Trim(CStr(targetValue)) <> Trim(CStr(refValue)) Then
                
                Application.EnableEvents = False
                Application.Undo
                
                MsgBox "Die Kategorie '" & cat & "' ist bereits mit dem Wert '" & refValue & "' definiert." & vbCrLf & _
                       "Unterschiedliche Werte in Spalte " & ColName(Target.Column) & " sind fuer diese Kategorie nicht erlaubt.", _
                       vbCritical, "Kategorie-Regel verletzt"
            End If
        End If
        
        GoTo ExitPoint
    End If
    
    ' =====================================================================
    ' ABSCHNITT 4: HELPER-SPALTEN Z-AH (26-34) - UNABHAENGIG!
    ' =====================================================================
    If spalte >= 26 And spalte <= 34 Then
        Application.EnableEvents = False
        ' WICHTIG: Jede Spalte wird als unabhaengige Einheit behandelt!
        Call mod_Formatierung.RefreshTableFormattingAfterChange(Me, spalte, zeile)
        Application.EnableEvents = True
        GoTo ExitPoint
    End If

ExitPoint:
    Application.EnableEvents = True
End Sub

'================================================================================
' Hilfsfunktion: Spaltenname aus Spaltennummer
'================================================================================
Private Function ColName(ByVal colNum As Long) As String
    Dim temp As String
    Do
        temp = Chr(((colNum - 1) Mod 26) + 65) & temp
        colNum = (colNum - 1) \ 26
    Loop While colNum > 0
    ColName = temp
End Function

