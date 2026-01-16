Attribute VB_Name = "Modul1"
Option Explicit
Public Sub ExportAllVBA()
    Dim vbComp As Object
    Dim folder As String
    folder = ThisWorkbook.Path & "\project_exports"
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        On Error Resume Next
        vbComp.Export folder & "\" & vbComp.Name & GetExtensionForComponent(vbComp)
        On Error GoTo 0
    Next vbComp
    MsgBox "Export abgeschlossen nach: " & folder, vbInformation
End Sub

Private Function GetExtensionForComponent(vbComp As Object) As String
    ' VBIDE.VBComponentType values: 1=StandardModule, 2=ClassModule, 3=MSForm, 100=Document (Sheet/ThisWorkbook)
    Select Case vbComp.Type
        Case 1: GetExtensionForComponent = ".bas"
        Case 2: GetExtensionForComponent = ".cls"
        Case 3: GetExtensionForComponent = ".frm" ' .frx will be exported automatically by VBIDE.Export
        Case Else: GetExtensionForComponent = ".bas"
    End Select
End Function
