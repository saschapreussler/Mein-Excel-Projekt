VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rng_KategorieRegeln As Range
    Dim cat As String
    Dim r As Range
    Dim refValue As Variant
    Dim targetValue As Variant
    Dim zeile As Long
    Dim spalte As Long

    On Error GoTo ExitPoint
    If Application.EnableEvents = False Then Exit Sub
    If Target.Cells.CountLarge <> 1 Then Exit Sub

    zeile = Target.Row
    spalte = Target.Column
    
    ' === ENTITYKEY-TABELLE VERARBEITUNG (Spalten R-X, ab Zeile 4) ===
    If zeile >= EK_START_ROW And spalte >= EK_COL_ENTITYKEY And spalte <= EK_COL_DEBUG Then
        
        ' Bei Role-Aenderung: EntityKey anpassen
        If spalte = EK_COL_ROLE Then
            Call mod_EntityKey_Manager.VerarbeiteManuelleRoleAenderung(Target)
        End If
        
        ' Zeile formatieren
        Call mod_EntityKey_Manager.VerarbeiteManuelleRoleAenderung(Target)
        GoTo ExitPoint
    End If
    
    ' === KATEGORIE-TABELLE VERARBEITUNG (Spalten J-P, ab Zeile 4) ===
    If zeile >= DATA_START_ROW And spalte >= DATA_CAT_COL_KATEGORIE And spalte <= DATA_CAT_COL_END Then
        If spalte = DATA_CAT_COL_KATEGORIE Or spalte = DATA_CAT_COL_EINAUS Then
            Call mod_Formatierung.OnKategorieChange(Target)
        End If
    End If
    
    ' === KATEGORIE-REGELN VERARBEITUNG ===
    
    On Error Resume Next
    Set rng_KategorieRegeln = Me.Range("rng_KategorieRegeln")
    On Error GoTo ExitPoint
    
    If rng_KategorieRegeln Is Nothing Then Exit Sub

    Select Case Target.Column
        Case DATA_CAT_COL_EINAUS, DATA_CAT_COL_ZIELSPALTE, DATA_CAT_COL_FAELLIGKEIT
            If Intersect(Target, rng_KategorieRegeln) Is Nothing Then
                Exit Sub
            End If
        Case Else
            Exit Sub
    End Select
    
    cat = Trim(CStr(Me.Cells(Target.Row, DATA_CAT_COL_KATEGORIE).value))
    If cat = "" Then Exit Sub

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

