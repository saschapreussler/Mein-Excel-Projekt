Attribute VB_Name = "mod_VBA_Export"
Option Explicit

' ***************************************************************
' MODUL: mod_VBA_Export
' ZWECK: Exportiert alle VBA-Komponenten in Repository-Ordner
' VERSION: 1.0 - 04.02.2026
' ***************************************************************

' ===============================================================
' ZIELORDNER FÜR EXPORT
' ===============================================================
Private Const EXPORT_PATH_CLASSES As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\Classes\"
Private Const EXPORT_PATH_USERFORMS As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\UserForms\"
Private Const EXPORT_PATH_MODULES As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\Modules\"

' ===============================================================
' HAUPTPROZEDUR: Exportiert alle VBA-Komponenten
' ===============================================================
Public Sub ExportiereAlleVBAKomponenten()
    
    Dim vbProj As Object
    Dim vbComp As Object
    Dim exportPath As String
    Dim fileName As String
    Dim extension As String
    
    Dim countModules As Long
    Dim countClasses As Long
    Dim countForms As Long
    Dim countSkipped As Long
    Dim errors As String
    
    On Error GoTo ErrorHandler
    
    ' Prüfe ob Zugriff auf VBA-Projekt erlaubt ist
    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "FEHLER: Zugriff auf VBA-Projekt nicht erlaubt!" & vbCrLf & vbCrLf & _
               "Bitte aktivieren Sie in Excel:" & vbCrLf & _
               "Datei > Optionen > Trust Center > Einstellungen für das Trust Center" & vbCrLf & _
               "> Makroeinstellungen > 'Zugriff auf das VBA-Projektobjektmodell vertrauen'", _
               vbCritical, "Zugriff verweigert"
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Prüfe ob Zielordner existieren
    If Not OrdnerExistiert(EXPORT_PATH_CLASSES) Then
        errors = errors & "Ordner nicht gefunden: " & EXPORT_PATH_CLASSES & vbCrLf
    End If
    If Not OrdnerExistiert(EXPORT_PATH_USERFORMS) Then
        errors = errors & "Ordner nicht gefunden: " & EXPORT_PATH_USERFORMS & vbCrLf
    End If
    If Not OrdnerExistiert(EXPORT_PATH_MODULES) Then
        errors = errors & "Ordner nicht gefunden: " & EXPORT_PATH_MODULES & vbCrLf
    End If
    
    If errors <> "" Then
        MsgBox "FEHLER: Zielordner nicht gefunden!" & vbCrLf & vbCrLf & errors, vbCritical, "Export abgebrochen"
        Exit Sub
    End If
    
    ' Zähler initialisieren
    countModules = 0
    countClasses = 0
    countForms = 0
    countSkipped = 0
    errors = ""
    
    ' Alle Komponenten durchlaufen
    For Each vbComp In vbProj.VBComponents
        
        fileName = vbComp.Name
        
        Select Case vbComp.Type
            
            Case 1 ' vbext_ct_StdModule - Standard-Modul (.bas)
                exportPath = EXPORT_PATH_MODULES
                extension = ".bas"
                
                On Error Resume Next
                vbComp.Export exportPath & fileName & extension
                If Err.Number = 0 Then
                    countModules = countModules + 1
                Else
                    errors = errors & "Modul: " & fileName & " - " & Err.Description & vbCrLf
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
                
            Case 2 ' vbext_ct_ClassModule - Klassenmodul (.cls)
                exportPath = EXPORT_PATH_CLASSES
                extension = ".cls"
                
                On Error Resume Next
                vbComp.Export exportPath & fileName & extension
                If Err.Number = 0 Then
                    countClasses = countClasses + 1
                Else
                    errors = errors & "Klasse: " & fileName & " - " & Err.Description & vbCrLf
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
                
            Case 3 ' vbext_ct_MSForm - UserForm (.frm + .frx)
                exportPath = EXPORT_PATH_USERFORMS
                extension = ".frm"
                
                On Error Resume Next
                vbComp.Export exportPath & fileName & extension
                If Err.Number = 0 Then
                    countForms = countForms + 1
                Else
                    errors = errors & "UserForm: " & fileName & " - " & Err.Description & vbCrLf
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
                
            Case 100 ' vbext_ct_Document - Dokument/Arbeitsblatt (.cls)
                ' Arbeitsblätter und ThisWorkbook als Klassen exportieren
                exportPath = EXPORT_PATH_CLASSES
                extension = ".cls"
                
                On Error Resume Next
                vbComp.Export exportPath & fileName & extension
                If Err.Number = 0 Then
                    countClasses = countClasses + 1
                Else
                    errors = errors & "Dokument: " & fileName & " - " & Err.Description & vbCrLf
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
                
            Case Else
                countSkipped = countSkipped + 1
                
        End Select
        
    Next vbComp
    
    ' Ergebnis anzeigen
    Dim msg As String
    msg = "VBA-Export abgeschlossen!" & vbCrLf & vbCrLf
    msg = msg & "Exportiert:" & vbCrLf
    msg = msg & "  Module:    " & countModules & " -> " & EXPORT_PATH_MODULES & vbCrLf
    msg = msg & "  Klassen:   " & countClasses & " -> " & EXPORT_PATH_CLASSES & vbCrLf
    msg = msg & "  UserForms: " & countForms & " -> " & EXPORT_PATH_USERFORMS & vbCrLf
    
    If countSkipped > 0 Then
        msg = msg & vbCrLf & "Übersprungen: " & countSkipped
    End If
    
    If errors <> "" Then
        msg = msg & vbCrLf & vbCrLf & "FEHLER:" & vbCrLf & errors
        MsgBox msg, vbExclamation, "Export mit Fehlern"
    Else
        msg = msg & vbCrLf & "Alle Komponenten erfolgreich exportiert!"
        MsgBox msg, vbInformation, "Export erfolgreich"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Unerwarteter Fehler beim Export:" & vbCrLf & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description, _
           vbCritical, "Export fehlgeschlagen"
End Sub

' ===============================================================
' HILFSFUNKTION: Prüft ob Ordner existiert
' ===============================================================
Private Function OrdnerExistiert(ByVal pfad As String) As Boolean
    On Error Resume Next
    OrdnerExistiert = (GetAttr(pfad) And vbDirectory) = vbDirectory
    If Err.Number <> 0 Then OrdnerExistiert = False
    On Error GoTo 0
End Function

' ===============================================================
' BONUS: Nur Module exportieren
' ===============================================================
Public Sub ExportiereNurModule()
    
    Dim vbProj As Object
    Dim vbComp As Object
    Dim count As Long
    
    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "Zugriff auf VBA-Projekt nicht erlaubt!", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    If Not OrdnerExistiert(EXPORT_PATH_MODULES) Then
        MsgBox "Ordner nicht gefunden: " & EXPORT_PATH_MODULES, vbCritical
        Exit Sub
    End If
    
    count = 0
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type = 1 Then ' Standard-Modul
            On Error Resume Next
            vbComp.Export EXPORT_PATH_MODULES & vbComp.Name & ".bas"
            If Err.Number = 0 Then count = count + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next vbComp
    
    MsgBox count & " Module exportiert nach:" & vbCrLf & EXPORT_PATH_MODULES, vbInformation
    
End Sub

' ===============================================================
' BONUS: Nur Klassen exportieren
' ===============================================================
Public Sub ExportiereNurKlassen()
    
    Dim vbProj As Object
    Dim vbComp As Object
    Dim count As Long
    
    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "Zugriff auf VBA-Projekt nicht erlaubt!", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    If Not OrdnerExistiert(EXPORT_PATH_CLASSES) Then
        MsgBox "Ordner nicht gefunden: " & EXPORT_PATH_CLASSES, vbCritical
        Exit Sub
    End If
    
    count = 0
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type = 2 Or vbComp.Type = 100 Then ' Klasse oder Dokument
            On Error Resume Next
            vbComp.Export EXPORT_PATH_CLASSES & vbComp.Name & ".cls"
            If Err.Number = 0 Then count = count + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next vbComp
    
    MsgBox count & " Klassen exportiert nach:" & vbCrLf & EXPORT_PATH_CLASSES, vbInformation
    
End Sub

