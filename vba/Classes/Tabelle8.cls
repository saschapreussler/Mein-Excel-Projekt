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
    Dim r As Range ' Einzelne Zelle in der Kategorie-Spalte J
    Dim refValue As Variant
    Dim targetValue As Variant
    
    ' Die Spaltennummern: J=10, K=11, N=14, O=15, P=16
    Const COL_KATEGORIE As Long = 10 ' J
    Const COL_K As Long = 11
    Const COL_N As Long = 14
    Const COL_O As Long = 15
    Const COL_P As Long = 16

    ' --- Fehlerbehandlung & Events ausschalten ---
    On Error GoTo ExitPoint
    If Application.EnableEvents = False Then Exit Sub
    If Target.Cells.CountLarge <> 1 Then Exit Sub ' Nur Einzelzellen-Änderungen behandeln

    ' --- 1. Bereich definieren und Prüfspalten identifizieren ---
    
    ' Definierten Bereich festlegen
    On Error Resume Next
    Set rng_KategorieRegeln = Me.Range("rng_KategorieRegeln")
    On Error GoTo ExitPoint
    
    If rng_KategorieRegeln Is Nothing Then Exit Sub ' Bereich nicht vorhanden

    ' Prüfen, ob die Änderung in K, N, O oder P UND innerhalb des definierten Bereichs liegt
    Select Case Target.Column
        Case COL_K, COL_N, COL_O, COL_P
            ' Die Spalte ist relevant, nun prüfen, ob die Zeile im Bereich liegt
            If Intersect(Target, rng_KategorieRegeln) Is Nothing Then
                Exit Sub ' Änderung außerhalb des relevanten Bereichs
            End If
        Case Else
            Exit Sub ' Änderung in einer irrelevanten Spalte
    End Select
    
    ' --- 2. Kategorie und neuen Wert ermitteln ---
    
    ' Kategorie ist immer in Spalte J (10) in der gleichen Zeile wie Target
    cat = Trim(CStr(Me.Cells(Target.Row, COL_KATEGORIE).value))
    If cat = "" Then Exit Sub ' Keine Kategorie, keine Prüfung
    
    targetValue = Me.Cells(Target.Row, Target.Column).value

    ' --- 3. Referenzwert suchen ---
    
    refValue = ""
    
    ' Schleife durch alle Zellen in Spalte J, die Teil des Bereichs sind
    For Each r In Intersect(Me.Columns(COL_KATEGORIE), rng_KategorieRegeln).Cells
        
        If Trim(CStr(r.value)) = cat Then ' Gleiche Kategorie gefunden
            
            Dim currentValue As Variant
            ' Wert in der aktuellen Prüfspalte (Target.Column) der Zeile von r
            currentValue = Me.Cells(r.Row, Target.Column).value

            ' Wenn ein nicht leerer Wert gefunden wird, der NICHT die gerade geänderte Zelle ist
            If Not IsEmpty(currentValue) And Len(Trim(CStr(currentValue))) > 0 Then
                If r.Row <> Target.Row Then
                    ' Dies ist der Referenzwert aus einer anderen Zeile
                    refValue = currentValue
                    Exit For ' Erster gefundener nicht leerer Wert ist unser Referenzwert
                End If
            End If
        End If
    Next r

    ' --- 4. Vergleich und Rückgängigmachung ---
    
    If Not IsEmpty(refValue) And Len(Trim(CStr(refValue))) > 0 Then
        ' Vergleich mit dem Referenzwert (nach Trimmen, um Leerzeichen zu ignorieren)
        If Trim(CStr(targetValue)) <> Trim(CStr(refValue)) Then

            Application.EnableEvents = False
            Application.Undo ' Änderung rückgängig machen

            MsgBox "Die Kategorie '" & cat & "' ist bereits mit dem Wert '" & refValue & "' definiert." & vbCrLf & _
                   "Unterschiedliche Werte in Spalte " & ColName(Target.Column) & " sind für diese Kategorie nicht erlaubt.", _
                   vbCritical, "Kategorie-Regel verletzt"
        End If
    End If

ExitPoint:
    Application.EnableEvents = True
End Sub

' Hilfsfunktion, um die Spaltennummer in den Spaltenbuchstaben umzuwandeln (z.B. 11 -> K)
Function ColName(ColNum As Long) As String
    ColName = Split(Columns(ColNum).Address(False, False), ":")(0)
End Function

