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

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rng_KategorieRegeln As Range
    Dim cat As String
    Dim r As Range
    Dim refValue As Variant
    Dim targetValue As Variant
    Dim zeile As Long
    Dim spalte As Long
    
    ' AKTUALISIERT: Spaltennummern nach Loeschung von Spalte O (Guthabenfaehig)
    ' Neue Struktur: J=Kategorie, K=E/A, L=Keyword, M=Prio, N=Zielspalte, O=Faelligkeit, P=Kommentar
    Const COL_KATEGORIE As Long = 10  ' J
    Const COL_EINAUS As Long = 11     ' K
    Const COL_ZIELSPALTE As Long = 14 ' N
    Const COL_FAELLIGKEIT As Long = 15 ' O
    Const COL_KOMMENTAR As Long = 16   ' P

    On Error GoTo ExitPoint
    If Application.EnableEvents = False Then Exit Sub
    If Target.Cells.CountLarge <> 1 Then Exit Sub

    zeile = Target.Row
    spalte = Target.Column
    
    ' === ENTITYKEY-TABELLE VERARBEITUNG (Spalten R-X, ab Zeile 4) ===
    ' KORRIGIERT: Jetzt R-X (18-24) statt S-Y (19-25)
    If zeile >= EK_START_ROW And spalte >= EK_COL_ENTITYKEY And spalte <= EK_COL_DEBUG Then
        
        If spalte = EK_COL_ROLE Then
            Call mod_EntityKey_Manager.VerarbeiteManuelleRoleAenderung(zeile)
        End If
        
        Call mod_EntityKey_Manager.FormatiereEntityKeyZeile(zeile)
        GoTo ExitPoint
    End If
    
    ' === KATEGORIE-TABELLE VERARBEITUNG (Spalten J-P, ab Zeile 4) ===
    If zeile >= DATA_START_ROW And spalte >= DATA_CAT_COL_KATEGORIE And spalte <= DATA_CAT_COL_END Then
        ' Bei Aenderung in Kategorie oder E/A: DropDown-Listen aktualisieren
        If spalte = COL_KATEGORIE Or spalte = COL_EINAUS Then
            Call mod_Formatierung.OnKategorieChange(Target)
        End If
    End If
    
    ' === KATEGORIE-REGELN VERARBEITUNG ===
    
    On Error Resume Next
    Set rng_KategorieRegeln = Me.Range("rng_KategorieRegeln")
    On Error GoTo ExitPoint
    
    If rng_KategorieRegeln Is Nothing Then Exit Sub

    ' Pruefen ob Aenderung in K, N, O liegt (relevante Sync-Spalten)
    Select Case Target.Column
        Case COL_EINAUS, COL_ZIELSPALTE, COL_FAELLIGKEIT
            If Intersect(Target, rng_KategorieRegeln) Is Nothing Then
                Exit Sub
            End If
        Case Else
            Exit Sub
    End Select
    
    ' Kategorie ermitteln
    cat = Trim(CStr(Me.Cells(Target.Row, COL_KATEGORIE).value))
    If cat = "" Then Exit Sub
    
    targetValue = Me.Cells(Target.Row, Target.Column).value

    ' Referenzwert suchen
    refValue = ""
    
    For Each r In Intersect(Me.Columns(COL_KATEGORIE), rng_KategorieRegeln).Cells
        
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

    ' Vergleich und Rueckgaengigmachung
    If Not IsEmpty(refValue) And Len(Trim(CStr(refValue))) > 0 Then
        If Trim(CStr(targetValue)) <> Trim(CStr(refValue)) Then

            Application.EnableEvents = False
            Application.Undo

            MsgBox "Die Kategorie '" & cat & "' ist bereits mit dem Wert '" & refValue & "' definiert." & vbCrLf & _
                   "Unterschiedliche Werte in Spalte " & ColName(Target.Column) & " sind fuer diese Kategorie nicht erlaubt.", _
                   vbCritical, "Kategorie-Regel verletzt"
        End If
    End If

ExitPoint:
    Application.EnableEvents = True
End Sub

Private Function ColName(ByVal colNum As Long) As String
    Dim temp As String
    Do
        temp = Chr(((colNum - 1) Mod 26) + 65) & temp
        colNum = (colNum - 1) \ 26
    Loop While colNum > 0
    ColName = temp
End Function
