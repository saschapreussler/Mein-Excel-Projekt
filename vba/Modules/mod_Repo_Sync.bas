Attribute VB_Name = "mod_Repo_Sync"
Option Explicit

' ***************************************************************
' MODUL: mod_Repo_Sync
' ZWECK: Importiert alle VBA-Komponenten aus dem Repository
' VERSION: 1.1 - 01.03.2026
' ***************************************************************

' ===============================================================
' QUELLORDNER FÜR IMPORT (REPOSITORY)
' ===============================================================
Private Const REPO_PATH_CLASSES As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\Classes\"
Private Const REPO_PATH_USERFORMS As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\UserForms\"
Private Const REPO_PATH_MODULES As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\Modules\"

' ===============================================================
' HAUPTPROZEDUR: Synchronisiert das VBA-Projekt mit dem Repo
' ===============================================================
Public Sub SyncVBAVomRepository()
    
    Dim vbProj As Object
    Dim vbComp As Object
    Dim fso As Object
    
    Dim countModules As Long
    Dim countClasses As Long
    Dim countForms As Long
    Dim skippedDocs As String
    
    On Error GoTo ErrorHandler
    
    ' 1. Prüfe Zugriff auf VBA-Projekt
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
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 2. Prüfe ob Quellordner existieren
    If Not fso.FolderExists(REPO_PATH_CLASSES) Or _
       Not fso.FolderExists(REPO_PATH_USERFORMS) Or _
       Not fso.FolderExists(REPO_PATH_MODULES) Then
        MsgBox "FEHLER: Mindestens ein Quellordner im Repo wurde nicht gefunden!", vbCritical, "Ordner fehlt"
        Exit Sub
    End If
    
    ' 3. Alte Mammut-Module gezielt entfernen (Platz schaffen für neue Splits)
    ' Wir löschen diese, da sie im Repo nun in Untermodule aufgeteilt sind.
    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents("mod_EntityKey_Manager")
    vbProj.VBComponents.Remove vbProj.VBComponents("mod_Formatierung")
    vbProj.VBComponents.Remove vbProj.VBComponents("mod_Banking_Data")
    vbProj.VBComponents.Remove vbProj.VBComponents("mod_ZaehlerLogik")
    vbProj.VBComponents.Remove vbProj.VBComponents("mod_Einstellungen")
    vbProj.VBComponents.Remove vbProj.VBComponents("mod_Zahlungspruefung")
    On Error GoTo ErrorHandler
    
    ' Zähler initialisieren
    countModules = 0
    countClasses = 0
    countForms = 0
    skippedDocs = ""
    
    ' 4. IMPORT-VORGÄNGE (Module, Klassen, UserForms)
    
    ' A) Standard-Module (.bas)
    ImportiereDateienAusOrdner fso, vbProj, REPO_PATH_MODULES, "bas", countModules, skippedDocs
    
    ' B) Klassen-Module (.cls)
    ImportiereDateienAusOrdner fso, vbProj, REPO_PATH_CLASSES, "cls", countClasses, skippedDocs
    
    ' C) UserForms (.frm)
    ImportiereDateienAusOrdner fso, vbProj, REPO_PATH_USERFORMS, "frm", countForms, skippedDocs
    
    ' 5. Ergebnis anzeigen
    Dim msg As String
    msg = "VBA-Synchronisierung abgeschlossen!" & vbCrLf & vbCrLf
    msg = msg & "Importiert aus Repo:" & vbCrLf
    msg = msg & "  Module:    " & countModules & vbCrLf
    msg = msg & "  Klassen:   " & countClasses & vbCrLf
    msg = msg & "  UserForms: " & countForms & vbCrLf
    
    If skippedDocs <> "" Then
        msg = msg & vbCrLf & "HINWEIS: Dokument-Module (z.B. TabelleX, DieseArbeitsmappe) wurden übersprungen, " & _
              "da diese in Excel fest verbaut sind." & vbCrLf
    End If
    
    msg = msg & vbCrLf & "Das Projekt ist nun auf dem Stand des Repositories." & vbCrLf & _
          "WICHTIG: Bitte führen Sie jetzt 'Debuggen -> Kompilieren' aus."
          
    MsgBox msg, vbInformation, "Synchronisierung erfolgreich"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Unerwarteter Fehler beim Import:" & vbCrLf & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description, _
           vbCritical, "Sync fehlgeschlagen"
End Sub

' ===============================================================
' HILFSPROZEDUR: Importiert alle Dateien eines Typs aus einem Ordner
' ===============================================================
Private Sub ImportiereDateienAusOrdner(fso As Object, vbProj As Object, pfad As String, ext As String, ByRef counter As Long, ByRef skipped As String)
    Dim folder As Object
    Dim file As Object
    Dim compName As String
    Dim vbComp As Object
    
    Set folder = fso.GetFolder(pfad)
    
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = LCase(ext) Then
            compName = fso.GetBaseName(file.Name)
            
            ' Überspringe die Synchronisierungs-Module selbst und den Exporteur
            If compName <> "mod_Repo_Sync" And compName <> "mod_VBA_Export" Then
                
                On Error Resume Next
                Set vbComp = vbProj.VBComponents(compName)
                
                If Err.Number = 0 Then
                    ' Komponente existiert bereits
                    If vbComp.Type = 100 Then
                        ' Dokument-Module (Tabelle, DieseArbeitsmappe)
                        skipped = skipped & compName & ", "
                        On Error GoTo 0
                    Else
                        ' Standard-Modul, Klasse oder Form löschen und neu importieren
                        vbProj.VBComponents.Remove vbComp
                        vbProj.VBComponents.Import file.Path
                        counter = counter + 1
                    End If
                Else
                    ' Komponente existiert noch nicht: Neu importieren
                    Err.Clear
                    vbProj.VBComponents.Import file.Path
                    counter = counter + 1
                End If
                On Error GoTo 0
                
            End If
        End If
    Next file
End Sub

