Attribute VB_Name = "Bootstrap_Reimport"
' ========================================================================
' Bootstrap_Reimport.bas
'
' ZWECK:
'   Einmaliger Bootstrap-Helfer, um die VBA-Quelldateien aus dem
'   Repository (vba\Modules\*.bas + vba\Classes\*.cls) nach einem
'   Umlaut-Fix sauber zurueck in die Arbeitsmappe zu importieren.
'
' HINTERGRUND:
'   Die Repo-Dateien sind UTF-8 mit BOM. Der VBA-Editor liest beim
'   File-Import nur ANSI. Wir muessen also UTF-8 nach ANSI konvertieren
'   bevor wir importieren. Das eigentliche Sync-Modul (mod_Repo_Sync)
'   kann das, aber wir muessen es ZUERST aktualisieren -- mit einer
'   bootstrap-fest verdrahteten ANSI-Konvertierung in diesem Modul.
'
' VERWENDUNG:
'   1. Excel-Mappe oeffnen (Programm Kassenbuch 2018_v2.7.4.xlsm)
'   2. Alt+F11 -> VBE oeffnen
'   3. Menue Datei -> Datei importieren ->
'      tools\Bootstrap_Reimport.bas auswaehlen
'   4. Im Direktbereich (Strg+G) tippen: BootstrapReimport
'      und Enter druecken
'   5. Excel zeigt am Ende die Sync-Ergebnis-MsgBox
'   6. Modul Bootstrap_Reimport kann dann wieder entfernt werden
'      (rechtsklick -> Entfernen -> Nein)
'
' WICHTIG:
'   Diese Datei MUSS ASCII-clean bleiben (keine Umlaute, keine
'   gross/klein-Kollisionen), damit sie selbst per File-Import
'   ohne Encoding-Probleme einspielbar ist.
' ========================================================================
Option Explicit

' ----- ANPASSEN, falls die Mappe an einem anderen Ort liegt --------------
Private Const REPO_ROOT      As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\"
Private Const SUB_MODULES    As String = "Modules\"
Private Const SUB_CLASSES    As String = "Classes\"
Private Const SYNC_MODULE    As String = "mod_Repo_Sync"
Private Const SYNC_FILE      As String = "mod_Repo_Sync.bas"
Private Const SYNC_SUB       As String = "SyncVBAVomRepository"

' Module, die VOR dem Sync entfernt werden muessen (z.B. weil sie mit
' anderen Modulen zusammengefuehrt wurden und sonst Ambiguous-Name-
' Konflikte verursachen).
Private Const MODULES_TO_REMOVE As String = "mod_TestReset"
' -------------------------------------------------------------------------

' Diagnose-Variablen fuer detaillierte Fehlermeldungen
Private gLastStep As String
Private gLastErrNum As Long
Private gLastErrDesc As String


Public Sub BootstrapReimport()
    Dim fso As Object
    Dim quellPfad As String
    Dim tempPfad As String
    Dim ansiPfad As String
    Dim vbProj As Object
    Dim vbComp As Object
    Dim antwort As VbMsgBoxResult

    On Error GoTo Fehler

    ' --- Pfade pruefen ----------------------------------------------------
    Set fso = CreateObject("Scripting.FileSystemObject")
    quellPfad = REPO_ROOT & SUB_MODULES & SYNC_FILE
    If Not fso.FileExists(quellPfad) Then
        MsgBox "Quelldatei nicht gefunden:" & vbCrLf & quellPfad & vbCrLf & vbCrLf & _
               "Bitte Konstante REPO_ROOT in Bootstrap_Reimport.bas anpassen.", _
               vbCritical, "Bootstrap abgebrochen"
        Exit Sub
    End If

    ' --- Sicherheitsabfrage ----------------------------------------------
    antwort = MsgBox( _
        "Bootstrap-Reimport startet jetzt:" & vbCrLf & vbCrLf & _
        "  1. Veraltete Module werden entfernt:" & vbCrLf & _
        "     " & MODULES_TO_REMOVE & vbCrLf & _
        "  2. mod_Repo_Sync wird aus dem Repo neu eingespielt" & vbCrLf & _
        "     (UTF-8 -> ANSI konvertiert)" & vbCrLf & _
        "  3. SyncVBAVomRepository wird gestartet, das alle anderen" & vbCrLf & _
        "     Module und Klassen aus dem Repo importiert." & vbCrLf & vbCrLf & _
        "Vor dem Start sollte die Mappe gespeichert sein!" & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbYesNo + vbQuestion, "Bootstrap-Reimport")
    If antwort <> vbYes Then Exit Sub

    Application.ScreenUpdating = False

    Set vbProj = ThisWorkbook.VBProject

    ' --- Schritt 0: Veraltete Module entfernen ----------------------------
    Dim moduleList() As String
    moduleList = Split(MODULES_TO_REMOVE, ",")
    Dim i As Long
    For i = LBound(moduleList) To UBound(moduleList)
        Dim modName As String
        modName = Trim(moduleList(i))
        If Len(modName) > 0 Then
            On Error Resume Next
            Set vbComp = Nothing
            Set vbComp = vbProj.VBComponents(modName)
            If Not vbComp Is Nothing Then
                vbProj.VBComponents.Remove vbComp
                Debug.Print "[Bootstrap] Entfernt: " & modName
            End If
            On Error GoTo Fehler
        End If
    Next i

    ' --- Schritt 1: mod_Repo_Sync nach ANSI konvertieren ------------------
    tempPfad = Environ$("TEMP") & "\bootstrap_reimport_" & Format(Now, "yyyymmdd_hhnnss") & "\"
    If Not fso.FolderExists(tempPfad) Then fso.CreateFolder tempPfad
    ansiPfad = tempPfad & SYNC_FILE

    If Not KonvertiereUTF8nachANSI(quellPfad, ansiPfad, fso) Then
        MsgBox "Konvertierung UTF-8 -> ANSI fehlgeschlagen fuer:" & vbCrLf & _
               quellPfad & vbCrLf & vbCrLf & _
               "Letzter Fehler:" & vbCrLf & _
               "  Schritt: " & gLastStep & vbCrLf & _
               "  Err " & gLastErrNum & ": " & gLastErrDesc, _
               vbCritical, "Bootstrap fehlgeschlagen"
        GoTo Aufraeumen
    End If

    ' --- Schritt 2: bestehendes mod_Repo_Sync im Projekt austauschen ------
    On Error Resume Next
    Set vbComp = Nothing
    Set vbComp = vbProj.VBComponents(SYNC_MODULE)
    On Error GoTo Fehler

    If Not vbComp Is Nothing Then
        ' Code direkt im CodeModule ersetzen (kein Remove noetig)
        If Not ErsetzeCodeAusANSIDatei(vbComp, ansiPfad) Then
            MsgBox "Konnte mod_Repo_Sync nicht aktualisieren.", _
                   vbCritical, "Bootstrap fehlgeschlagen"
            GoTo Aufraeumen
        End If
    Else
        ' nicht vorhanden -> neu importieren
        vbProj.VBComponents.Import ansiPfad
    End If

    ' --- Schritt 3: SyncVBAVomRepository starten --------------------------
    Application.ScreenUpdating = True
    Application.Run SYNC_SUB

Aufraeumen:
    On Error Resume Next
    If fso.FolderExists(tempPfad) Then fso.DeleteFolder tempPfad, True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub

Fehler:
    Application.ScreenUpdating = True
    MsgBox "Unerwarteter Fehler:" & vbCrLf & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description, _
           vbCritical, "Bootstrap fehlgeschlagen"
End Sub


' ========================================================================
' Konvertiert eine UTF-8-Datei (mit oder ohne BOM) in eine ANSI-Datei
' (Windows-1252). Loest das BOM-Problem beim VBA-Import.
' Verwendet ADODB.Stream fuer Lesen UND Schreiben - keine FSO-Mischung.
' Rueckgabe: True bei Erfolg. Bei Fehler werden gLastStep / gLastErrNum
' / gLastErrDesc gesetzt, damit der Aufrufer eine sinnvolle Meldung
' anzeigen kann.
' ========================================================================
Private Function KonvertiereUTF8nachANSI(quellPfad As String, _
                                          zielPfad As String, _
                                          fso As Object) As Boolean
    Dim sIn As Object
    Dim sOut As Object
    Dim inhalt As String

    gLastStep = ""
    gLastErrNum = 0
    gLastErrDesc = ""

    ' --- Schritt 1: Quelldatei pruefen ----------------------------------
    gLastStep = "Quelldatei pruefen"
    If Not fso.FileExists(quellPfad) Then
        gLastErrDesc = "Datei existiert nicht."
        Exit Function
    End If

    ' --- Schritt 2: ADODB.Stream erstellen (Lesen) ----------------------
    gLastStep = "ADODB.Stream (Lesen) erstellen"
    On Error Resume Next
    Set sIn = CreateObject("ADODB.Stream")
    If sIn Is Nothing Then
        gLastErrNum = Err.Number
        gLastErrDesc = "CreateObject(ADODB.Stream) fehlgeschlagen: " & Err.Description
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' --- Schritt 3: Quelldatei als UTF-8 lesen --------------------------
    gLastStep = "Quelldatei als UTF-8 lesen"
    On Error Resume Next
    With sIn
        .Type = 2                ' adTypeText
        .Charset = "utf-8"
        .Open
        .LoadFromFile quellPfad
        inhalt = .ReadText(-1)
        .Close
    End With
    If Err.Number <> 0 Then
        gLastErrNum = Err.Number
        gLastErrDesc = Err.Description
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' BOM entfernen (falls noch vorhanden)
    If Len(inhalt) > 0 Then
        If AscW(Left(inhalt, 1)) = &HFEFF Then
            inhalt = Mid(inhalt, 2)
        End If
    End If

    ' --- Schritt 4: Zieldatei loeschen falls vorhanden ------------------
    gLastStep = "Zieldatei vorbereiten"
    On Error Resume Next
    If fso.FileExists(zielPfad) Then fso.DeleteFile zielPfad, True
    Err.Clear
    On Error GoTo 0

    ' --- Schritt 5: ADODB.Stream erstellen (Schreiben) ------------------
    gLastStep = "ADODB.Stream (Schreiben) erstellen"
    On Error Resume Next
    Set sOut = CreateObject("ADODB.Stream")
    On Error GoTo 0

    ' --- Schritt 6: Als Windows-1252 (ANSI) schreiben -------------------
    gLastStep = "Als Windows-1252 schreiben"
    On Error Resume Next
    With sOut
        .Type = 2                ' adTypeText
        .Charset = "windows-1252"
        .Open
        .WriteText inhalt
        .SaveToFile zielPfad, 2  ' adSaveCreateOverWrite
        .Close
    End With
    If Err.Number <> 0 Then
        gLastErrNum = Err.Number
        gLastErrDesc = Err.Description
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' --- Schritt 7: BOM aus ANSI-Datei entfernen ------------------------
    ' ADODB.Stream schreibt mit Charset=windows-1252 KEIN BOM,
    ' aber sicherheitshalber pruefen
    gLastStep = "BOM aus ANSI-Datei entfernen"
    On Error Resume Next
    Dim pruef As Object
    Set pruef = CreateObject("ADODB.Stream")
    pruef.Type = 1               ' adTypeBinary
    pruef.Open
    pruef.LoadFromFile zielPfad
    Dim allBytes() As Byte
    If pruef.Size > 0 Then
        allBytes = pruef.Read()
        pruef.Close
        ' Wenn erste 3 Bytes EF BB BF sind -> entfernen
        If UBound(allBytes) >= 2 Then
            If allBytes(0) = &HEF And allBytes(1) = &HBB And allBytes(2) = &HBF Then
                Dim neueBytes() As Byte
                Dim n As Long
                n = UBound(allBytes) - 2
                ReDim neueBytes(n)
                Dim k As Long
                For k = 0 To n
                    neueBytes(k) = allBytes(k + 3)
                Next k
                Dim s2 As Object
                Set s2 = CreateObject("ADODB.Stream")
                s2.Type = 1
                s2.Open
                s2.Write neueBytes
                s2.SaveToFile zielPfad, 2
                s2.Close
            End If
        End If
    Else
        pruef.Close
    End If
    On Error GoTo 0

    KonvertiereUTF8nachANSI = True
End Function


' ========================================================================
' Ersetzt den Code im CodeModule einer existierenden VBComponent
' durch den Inhalt einer ANSI-.bas-Datei. Header-Zeilen werden
' uebersprungen (Attribute VB_Name etc.).
' ========================================================================
Private Function ErsetzeCodeAusANSIDatei(vbComp As Object, _
                                          dateipfad As String) As Boolean
    Dim ts As Object
    Dim fso As Object
    Dim alleZeilen As String
    Dim zeilen() As String
    Dim i As Long
    Dim startIdx As Long
    Dim codeText As String

    On Error GoTo Fehler

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(dateipfad, 1, False, 0)  ' ForReading, ASCII
    If Not ts.AtEndOfStream Then
        alleZeilen = ts.ReadAll
    Else
        alleZeilen = ""
    End If
    ts.Close

    If InStr(alleZeilen, vbCrLf) > 0 Then
        zeilen = Split(alleZeilen, vbCrLf)
    Else
        zeilen = Split(alleZeilen, vbLf)
    End If

    ' Header-Zeile "Attribute VB_Name = ..." ueberspringen
    startIdx = -1
    For i = LBound(zeilen) To UBound(zeilen)
        If Left(Trim(zeilen(i)), 9) <> "Attribute" And Trim(zeilen(i)) <> "" Then
            startIdx = i
            Exit For
        End If
        If Left(Trim(zeilen(i)), 9) = "Attribute" Then
            ' weiter
        End If
    Next i

    If startIdx < 0 Then startIdx = 0

    codeText = ""
    For i = startIdx To UBound(zeilen)
        If i > startIdx Then codeText = codeText & vbCrLf
        codeText = codeText & zeilen(i)
    Next i

    With vbComp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(Trim(codeText)) > 0 Then .AddFromString codeText
    End With

    ErsetzeCodeAusANSIDatei = True
    Exit Function

Fehler:
    ErsetzeCodeAusANSIDatei = False
End Function


