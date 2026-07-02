Attribute VB_Name = "mod_VBA_Export"
Option Explicit

' ***************************************************************
' MODUL: mod_VBA_Export
' VERSION: 1.2 - 02.07.2026
'
' ZWECK: Exportiert alle VBA-Komponenten in das Repository und
'        konvertiert .bas/.cls-Dateien anschliessend explizit
'        von Windows-1252 (VBA-Export-Format) nach UTF-8 mit BOM,
'        damit VS Code und Git die Umlaute korrekt anzeigen.
'
' WICHTIG: .frm-Dateien werden NICHT konvertiert! Sie muessen fuer
'          den Re-Import im ANSI-Format (Windows-1252, ohne BOM)
'          bleiben, sonst schlaegt VBComponents.Import fehl.
'
' AUFRUF:
'        ExportiereAlleVBAKomponenten   Voller Export
'        ExportiereNurModule            Nur .bas
'        ExportiereNurKlassen           Nur .cls
' ***************************************************************

' ===============================================================
' ZIELORDNER FUER EXPORT
' ===============================================================
Private Const EXPORT_PATH_CLASSES   As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\Classes\"
Private Const EXPORT_PATH_USERFORMS As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\UserForms\"
Private Const EXPORT_PATH_MODULES   As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\Modules\"


' ===============================================================
' HAUPTPROZEDUR: Exportiert alle VBA-Komponenten
' ===============================================================
Public Sub ExportiereAlleVBAKomponenten()

    Dim vbProj As Object
    Dim vbComp As Object
    Dim zielDatei As String
    Dim countModules As Long, countClasses As Long, countForms As Long
    Dim countSkipped As Long
    Dim errors As String

    On Error GoTo ErrorHandler

    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "FEHLER: Zugriff auf VBA-Projekt nicht erlaubt!" & vbCrLf & vbCrLf & _
               "Bitte in Excel aktivieren:" & vbCrLf & _
               "Datei > Optionen > Trust Center > Trust-Center-Einstellungen" & vbCrLf & _
               "> Makroeinstellungen > 'Zugriff auf das VBA-Projektobjektmodell vertrauen'", _
               vbCritical, "Zugriff verweigert"
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    If Not OrdnerExistiert(EXPORT_PATH_CLASSES) Then errors = errors & "Ordner fehlt: " & EXPORT_PATH_CLASSES & vbCrLf
    If Not OrdnerExistiert(EXPORT_PATH_USERFORMS) Then errors = errors & "Ordner fehlt: " & EXPORT_PATH_USERFORMS & vbCrLf
    If Not OrdnerExistiert(EXPORT_PATH_MODULES) Then errors = errors & "Ordner fehlt: " & EXPORT_PATH_MODULES & vbCrLf
    If errors <> "" Then
        MsgBox "FEHLER: Zielordner fehlen!" & vbCrLf & vbCrLf & errors, vbCritical, "Export abgebrochen"
        Exit Sub
    End If

    For Each vbComp In vbProj.VBComponents

        Select Case vbComp.Type

            Case 1  ' vbext_ct_StdModule - .bas
                zielDatei = EXPORT_PATH_MODULES & vbComp.Name & ".bas"
                On Error Resume Next
                vbComp.Export zielDatei
                If Err.Number = 0 Then
                    KonvertiereDateiZuUtf8BOM zielDatei
                    countModules = countModules + 1
                Else
                    errors = errors & "Modul: " & vbComp.Name & " - " & Err.Description & vbCrLf
                    Err.Clear
                End If
                On Error GoTo ErrorHandler

            Case 2   ' vbext_ct_ClassModule - .cls
                zielDatei = EXPORT_PATH_CLASSES & vbComp.Name & ".cls"
                On Error Resume Next
                vbComp.Export zielDatei
                If Err.Number = 0 Then
                    KonvertiereDateiZuUtf8BOM zielDatei
                    countClasses = countClasses + 1
                Else
                    errors = errors & "Klasse: " & vbComp.Name & " - " & Err.Description & vbCrLf
                    Err.Clear
                End If
                On Error GoTo ErrorHandler

            Case 3   ' vbext_ct_MSForm - .frm + .frx
                zielDatei = EXPORT_PATH_USERFORMS & vbComp.Name & ".frm"
                On Error Resume Next
                vbComp.Export zielDatei
                ' KEINE Encoding-Konvertierung! .frm muss ANSI bleiben,
                ' sonst schlaegt der Re-Import fehl.
                If Err.Number = 0 Then
                    countForms = countForms + 1
                Else
                    errors = errors & "UserForm: " & vbComp.Name & " - " & Err.Description & vbCrLf
                    Err.Clear
                End If
                On Error GoTo ErrorHandler

            Case 100 ' vbext_ct_Document - Tabellenmodule + DieseArbeitsmappe (.cls)
                zielDatei = EXPORT_PATH_CLASSES & vbComp.Name & ".cls"
                On Error Resume Next
                vbComp.Export zielDatei
                If Err.Number = 0 Then
                    KonvertiereDateiZuUtf8BOM zielDatei
                    countClasses = countClasses + 1
                Else
                    errors = errors & "Dokument: " & vbComp.Name & " - " & Err.Description & vbCrLf
                    Err.Clear
                End If
                On Error GoTo ErrorHandler

            Case Else
                countSkipped = countSkipped + 1

        End Select

    Next vbComp

    Dim msg As String
    msg = "VBA-Export abgeschlossen! (v1.2 - UTF-8+BOM)" & vbCrLf & vbCrLf & _
          "Exportiert:" & vbCrLf & _
          "  Module:    " & countModules & vbCrLf & _
          "  Klassen:   " & countClasses & vbCrLf & _
          "  UserForms: " & countForms & " (ANSI, unveraendert)" & vbCrLf
    If countSkipped > 0 Then msg = msg & vbCrLf & "Uebersprungen: " & countSkipped
    If errors <> "" Then
        msg = msg & vbCrLf & vbCrLf & "FEHLER:" & vbCrLf & errors
        MsgBox msg, vbExclamation, "Export mit Fehlern"
    Else
        msg = msg & vbCrLf & "Alle Komponenten erfolgreich exportiert."
        MsgBox msg, vbInformation, "Export erfolgreich"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Unerwarteter Fehler beim Export:" & vbCrLf & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description, _
           vbCritical, "Export fehlgeschlagen"
End Sub


' ===============================================================
' KonvertiereDateiZuUtf8BOM (v1.2 Kernstueck)
' ---------------------------------------------------------------
' Liest die soeben von VBA.Export geschriebene ANSI-Datei
' explizit als Windows-1252 und schreibt sie als UTF-8 mit BOM
' zurueck. Damit bleiben Umlaute im Repo erhalten und VS Code
' zeigt sie korrekt an.
'
' NUR fuer .bas und .cls aufrufen -- .frm-Dateien muessen ANSI
' bleiben!
' ===============================================================
Public Sub KonvertiereDateiZuUtf8BOM(ByVal pfad As String)
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(pfad) Then Exit Sub

    Dim st As Object
    Set st = CreateObject("ADODB.Stream")

    ' 1) Inhalt EXPLIZIT als Windows-1252 lesen
    st.Type = 2                         ' adTypeText
    st.Charset = "Windows-1252"
    st.Open
    st.LoadFromFile pfad
    Dim txt As String
    txt = st.ReadText(-1)
    st.Close

    ' 2) Als UTF-8 (ADODB fuegt automatisch BOM hinzu) zurueckschreiben
    st.Open
    st.Charset = "UTF-8"
    st.WriteText txt
    st.SaveToFile pfad, 2               ' adSaveCreateOverWrite
    st.Close

    Set st = Nothing
    Set fso = Nothing

    On Error GoTo 0
End Sub


' ===============================================================
' Nur Module exportieren
' ===============================================================
Public Sub ExportiereNurModule()
    Dim vbProj As Object, vbComp As Object, ziel As String, count As Long

    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "Zugriff auf VBA-Projekt nicht erlaubt!", vbCritical: Exit Sub
    End If
    On Error GoTo 0

    If Not OrdnerExistiert(EXPORT_PATH_MODULES) Then
        MsgBox "Ordner fehlt: " & EXPORT_PATH_MODULES, vbCritical: Exit Sub
    End If

    For Each vbComp In vbProj.VBComponents
        If vbComp.Type = 1 Then
            ziel = EXPORT_PATH_MODULES & vbComp.Name & ".bas"
            On Error Resume Next
            vbComp.Export ziel
            If Err.Number = 0 Then
                KonvertiereDateiZuUtf8BOM ziel
                count = count + 1
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next vbComp

    MsgBox count & " Module exportiert (UTF-8+BOM)" & vbCrLf & EXPORT_PATH_MODULES, vbInformation
End Sub


' ===============================================================
' Nur Klassen exportieren
' ===============================================================
Public Sub ExportiereNurKlassen()
    Dim vbProj As Object, vbComp As Object, ziel As String, count As Long

    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "Zugriff auf VBA-Projekt nicht erlaubt!", vbCritical: Exit Sub
    End If
    On Error GoTo 0

    If Not OrdnerExistiert(EXPORT_PATH_CLASSES) Then
        MsgBox "Ordner fehlt: " & EXPORT_PATH_CLASSES, vbCritical: Exit Sub
    End If

    For Each vbComp In vbProj.VBComponents
        If vbComp.Type = 2 Or vbComp.Type = 100 Then
            ziel = EXPORT_PATH_CLASSES & vbComp.Name & ".cls"
            On Error Resume Next
            vbComp.Export ziel
            If Err.Number = 0 Then
                KonvertiereDateiZuUtf8BOM ziel
                count = count + 1
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next vbComp

    MsgBox count & " Klassen exportiert (UTF-8+BOM)" & vbCrLf & EXPORT_PATH_CLASSES, vbInformation
End Sub


' ===============================================================
' HILFSFUNKTION: Ordner-Existenz
' ===============================================================
Private Function OrdnerExistiert(ByVal pfad As String) As Boolean
    On Error Resume Next
    OrdnerExistiert = (GetAttr(pfad) And vbDirectory) = vbDirectory
    If Err.Number <> 0 Then OrdnerExistiert = False
    On Error GoTo 0
End Function
