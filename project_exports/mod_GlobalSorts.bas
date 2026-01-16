Attribute VB_Name = "mod_GlobalSorts"
Option Explicit

' ***************************************************************
' MODUL: mod_GlobalSorts
' ZWECK: Global sort helper for worksheet data
' ***************************************************************

' TODO: Replace with original complex implementation if available
' This is a minimal implementation that sorts a sheet by the specified key range
Public Sub GlobalSort(ByVal sh As Worksheet, ByVal keyRange As String)
    On Error GoTo SortError
    
    Dim sortKey As Range
    Dim dataRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' Validate inputs
    If sh Is Nothing Then
        MsgBox "Invalid worksheet provided to GlobalSort", vbCritical
        Exit Sub
    End If
    
    If Trim(keyRange) = "" Then
        MsgBox "Invalid key range provided to GlobalSort", vbCritical
        Exit Sub
    End If
    
    ' Set the sort key
    Set sortKey = sh.Range(keyRange)
    
    ' Find the data range boundaries
    lastRow = sh.Cells(sh.Rows.Count, sortKey.Column).End(xlUp).Row
    lastCol = sh.Cells(sortKey.Row, sh.Columns.Count).End(xlToLeft).Column
    
    ' Validate that we have data to sort
    If lastRow <= sortKey.Row Then
        Exit Sub ' No data to sort
    End If
    
    ' Define the data range (from key row to last row, all columns)
    Set dataRange = sh.Range(sh.Cells(sortKey.Row, 1), sh.Cells(lastRow, lastCol))
    
    ' Perform the sort
    With sh.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sortKey, _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SetRange dataRange
        .Header = xlNo
        .Apply
    End With
    
    Exit Sub

SortError:
    MsgBox "Error in GlobalSort: " & Err.Description, vbCritical
End Sub
