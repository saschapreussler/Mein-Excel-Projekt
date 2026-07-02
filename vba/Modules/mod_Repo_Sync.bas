Attribute VB_Name = "mod_Repo_Sync"
Option Explicit

' ***************************************************************
' MODUL: mod_Repo_Sync
' VERSION: 3.3 - 02.07.2026
'
' ZWECK: Importiert ALLE VBA-Komponenten aus dem Repository
'        (Modules, Classes, UserForms) in das laufende VBA-Projekt.
'
' NEUE FEATURES gegenueber v3.0:
'   - BOM-basierte Encoding-Erkennung (kein "?"-Heuristik-Bug mehr)
'   - .frm/.frx per Binaerkopie (KEINE Encoding-Konvertierung)
'   - Explizit Windows-1252 fuer den ANSI-Schreib-Zwischenschritt
'   - Public-Wrapper fuer das Direktfenster:
'         RepoNachExcel        Repo -> Excel
'         ExcelNachRepo        Excel -> Repo (UTF-8+BOM)
'         SyncMitVerification  Legacy-Alias
'
' STRATEGIE ("CodeModule first"):
'   1. Doubletten bereinigen (mod_XYZ1, mod_XYZ2 usw.)
'   2. Bestehende .bas/.cls: CodeModule-Ersetzung (in-place,
'      kein Remove -> keine "Zugriff verweigert"-Fehler)
'   3. Neue Module: VBComponents.Import ueber ANSI-Temp
'   4. UserForms: Remove + Import per Binaerkopie
'
' REPO-DATEIFORMAT:
'   .bas/.cls  UTF-8 mit BOM  (VS Code + Git-freundlich)
'   .frm/.frx  ANSI ohne BOM  (VBA-Import-Format)
'
' HINWEIS: mod_Repo_Sync und mod_VBA_Export werden beim Import
'          uebersprungen, um sich nicht selbst zu ueberschreiben.
' ***************************************************************

' ===============================================================
' QUELLORDNER FUER IMPORT (REPOSITORY)
' ===============================================================
Private Const REPO_PATH_CLASSES   As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\Classes\"
Private Const REPO_PATH_USERFORMS As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\UserForms\"
Private Const REPO_PATH_MODULES   As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\Modules\"

Private Const TEMP_SUBFOLDER As String = "VBA_Repo_Sync_Temp"


' ===============================================================
' HAUPTPROZEDUR: Synchronisiert das VBA-Projekt mit dem Repo
' ===============================================================
Public Sub SyncVBAVomRepository()

    Dim vbProj As Object
    Dim fso As Object
    Dim tempPfad As String

    Dim countModules As Long, countKlassen As Long, countForms As Long
    Dim countDokumente As Long, countDoubletten As Long
    Dim fehlerListe As String

    On Error GoTo ErrorHandler

    ' ---------------------------------------------------------
    ' 1. Zugriff auf VBA-Projekt pruefen
    ' ---------------------------------------------------------
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

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Quellordner pruefen
    If Not fso.FolderExists(REPO_PATH_MODULES) Or _
       Not fso.FolderExists(REPO_PATH_USERFORMS) Or _
       Not fso.FolderExists(REPO_PATH_CLASSES) Then
        MsgBox "FEHLER: Repository-Ordner fehlen!" & vbCrLf & vbCrLf & _
               "  " & REPO_PATH_MODULES & vbCrLf & _
               "  " & REPO_PATH_CLASSES & vbCrLf & _
               "  " & REPO_PATH_USERFORMS, _
               vbCritical, "Repo nicht gefunden"
        Exit Sub
    End If

    ' Temp-Ordner (fuer ANSI-Zwischenkopien beim .bas/.cls-Import)
    tempPfad = Environ("TEMP") & "\" & TEMP_SUBFOLDER & "\"
    If Not fso.FolderExists(tempPfad) Then fso.CreateFolder tempPfad

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ---------------------------------------------------------
    ' 2. Doubletten bereinigen (mod_XYZ1, mod_XYZ2, ...)
    ' ---------------------------------------------------------
    Application.StatusBar = "VBA-Sync: Bereinige Doubletten..."
    BereinigeDoubletten fso, vbProj, countDoubletten, fehlerListe

    ' ---------------------------------------------------------
    ' 3. .bas Standard-Module (CodeModule-Ersetzung + Neu-Import)
    ' ---------------------------------------------------------
    Application.StatusBar = "VBA-Sync: Importiere Standard-Module..."
    ImportiereStandardDateien fso, vbProj, REPO_PATH_MODULES, tempPfad, _
                              "bas", 1, countModules, fehlerListe

    ' ---------------------------------------------------------
    ' 4. .cls Klassen (inkl. Dokument-Module Type=100)
    ' ---------------------------------------------------------
    Application.StatusBar = "VBA-Sync: Importiere Klassen..."
    ImportiereKlassenModule fso, vbProj, REPO_PATH_CLASSES, tempPfad, _
                            countKlassen, countDokumente, fehlerListe

    ' ---------------------------------------------------------
    ' 5. .frm UserForms (Remove + Import per Binaerkopie)
    ' ---------------------------------------------------------
    Application.StatusBar = "VBA-Sync: Importiere UserForms..."
    ImportiereUserForms fso, vbProj, REPO_PATH_USERFORMS, tempPfad, _
                        countForms, fehlerListe

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ' ---------------------------------------------------------
    ' Ergebnis-MsgBox
    ' ---------------------------------------------------------
    Dim msg As String
    msg = "VBA-Synchronisierung abgeschlossen! (v3.3)" & vbCrLf & vbCrLf & _
          "Importiert aus Repo:" & vbCrLf & _
          "  Standard-Module: " & countModules & vbCrLf & _
          "  Klassen-Module:  " & countKlassen & vbCrLf & _
          "  Dokument-Module: " & countDokumente & " (Code ueberschrieben)" & vbCrLf & _
          "  UserForms:       " & countForms & vbCrLf

    If countDoubletten > 0 Then
        msg = msg & vbCrLf & "Doubletten entfernt: " & countDoubletten & vbCrLf
    End If

    If fehlerListe <> "" Then
        msg = msg & vbCrLf & "FEHLER bei folgenden Dateien:" & vbCrLf & fehlerListe & _
              vbCrLf & "Das Projekt ist NICHT vollstaendig synchronisiert."
        MsgBox msg, vbExclamation, "Synchronisierung mit Warnungen"
    Else
        msg = msg & vbCrLf & "Das Projekt ist nun auf dem Stand des Repositories." & vbCrLf & _
              "WICHTIG: Bitte fuehre jetzt 'Debuggen > Kompilieren' aus."
        MsgBox msg, vbInformation, "Synchronisierung erfolgreich"
    End If

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Unerwarteter Fehler bei Sync:" & vbCrLf & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description, _
           vbCritical, "Sync fehlgeschlagen"
End Sub


' ===============================================================
' Importiert .bas-Dateien
' ===============================================================
Private Sub ImportiereStandardDateien(fso As Object, vbProj As Object, _
                                       pfad As String, tempPfad As String, _
                                       ext As String, componentType As Integer, _
                                       ByRef counter As Long, _
                                       ByRef Fehler As String)
    Dim folder As Object, file As Object
    Dim compName As String
    Dim vbComp As Object
    Dim inhalt As String
    Dim ansiDatei As String

    Set folder = fso.GetFolder(pfad)
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = LCase(ext) Then
            compName = fso.GetBaseName(file.Name)

            ' Selbst-Schutz
            If compName = "mod_Repo_Sync" Or compName = "mod_VBA_Export" Then
                GoTo NaechsteDatei
            End If

            On Error Resume Next

            ' Existierendes Modul: CodeModule ersetzen (in-place)
            Set vbComp = Nothing
            Set vbComp = vbProj.VBComponents(compName)
            If Not vbComp Is Nothing Then
                inhalt = LeseDateiOhneKlassenHeader(file.Path)
                ErsetzeCodeInDokumentModul vbComp, inhalt
                If Err.Number = 0 Then
                    counter = counter + 1
                Else
                    Fehler = Fehler & "  " & file.Name & " (CodeModule: " & Err.Description & ")" & vbCrLf
                    Err.Clear
                End If
            Else
                ' Neu: Import ueber ANSI-Temp
                ansiDatei = KonvertiereUTF8zuAnsi(file.Path, tempPfad & file.Name, fso)
                If ansiDatei <> "" Then
                    vbProj.VBComponents.Import ansiDatei
                Else
                    vbProj.VBComponents.Import file.Path
                End If
                If Err.Number = 0 Then
                    counter = counter + 1
                Else
                    Fehler = Fehler & "  " & file.Name & " (Import: " & Err.Description & ")" & vbCrLf
                    Err.Clear
                End If
            End If

            On Error GoTo 0
NaechsteDatei:
        End If
    Next file
End Sub


' ===============================================================
' Importiert .cls-Dateien (Klassen + Dokument-Module Type=100)
' ===============================================================
Private Sub ImportiereKlassenModule(fso As Object, vbProj As Object, _
                                     pfad As String, tempPfad As String, _
                                     ByRef counterKlassen As Long, _
                                     ByRef counterDokumente As Long, _
                                     ByRef Fehler As String)
    Dim folder As Object, file As Object
    Dim compName As String
    Dim vbComp As Object
    Dim inhalt As String
    Dim ansiDatei As String

    Set folder = fso.GetFolder(pfad)
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "cls" Then
            compName = fso.GetBaseName(file.Name)

            On Error Resume Next
            Set vbComp = Nothing
            Set vbComp = vbProj.VBComponents(compName)

            If Not vbComp Is Nothing Then
                ' Existierendes Element -> Code ersetzen
                inhalt = LeseDateiOhneKlassenHeader(file.Path)
                ErsetzeCodeInDokumentModul vbComp, inhalt
                If Err.Number = 0 Then
                    If vbComp.Type = 100 Then
                        counterDokumente = counterDokumente + 1
                    Else
                        counterKlassen = counterKlassen + 1
                    End If
                Else
                    Fehler = Fehler & "  " & file.Name & " (CodeModule: " & Err.Description & ")" & vbCrLf
                    Err.Clear
                End If
            Else
                ' Neue Klasse -> Import ueber ANSI-Temp
                ansiDatei = KonvertiereUTF8zuAnsi(file.Path, tempPfad & file.Name, fso)
                If ansiDatei <> "" Then
                    vbProj.VBComponents.Import ansiDatei
                Else
                    vbProj.VBComponents.Import file.Path
                End If
                If Err.Number = 0 Then
                    counterKlassen = counterKlassen + 1
                Else
                    Fehler = Fehler & "  " & file.Name & " (Import: " & Err.Description & ")" & vbCrLf
                    Err.Clear
                End If
            End If

            On Error GoTo 0
        End If
    Next file
End Sub


' ===============================================================
' Importiert UserForms (.frm + .frx) per BINAERKOPIE
' ---------------------------------------------------------------
' WICHTIG: KEINE Encoding-Konvertierung! .frm-Dateien sind bereits
' ANSI (Windows-1252) — VBA erwartet genau dieses Format. Ein
' UTF-8->ANSI-Roundtrip zerstoert einzelne Umlaut-Bytes.
' ===============================================================
Private Sub ImportiereUserForms(fso As Object, vbProj As Object, _
                                 pfad As String, tempPfad As String, _
                                 ByRef counter As Long, _
                                 ByRef Fehler As String)
    Dim folder As Object, file As Object
    Dim compName As String
    Dim vbComp As Object
    Dim frmQuelle As String, frmZiel As String
    Dim frxQuelle As String, frxZiel As String

    Set folder = fso.GetFolder(pfad)
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "frm" Then
            compName = fso.GetBaseName(file.Name)

            On Error Resume Next

            ' Bestehende UserForm entfernen
            Set vbComp = Nothing
            Set vbComp = vbProj.VBComponents(compName)
            If Not vbComp Is Nothing Then
                vbProj.VBComponents.Remove vbComp
                DoEvents
            End If
            Err.Clear

            ' 1:1 Binaerkopie ins Temp
            frmQuelle = file.Path
            frmZiel = tempPfad & compName & ".frm"
            fso.CopyFile frmQuelle, frmZiel, True

            frxQuelle = fso.BuildPath(pfad, compName & ".frx")
            frxZiel = tempPfad & compName & ".frx"
            If fso.FileExists(frxQuelle) Then
                fso.CopyFile frxQuelle, frxZiel, True
            End If

            Err.Clear
            vbProj.VBComponents.Import frmZiel

            If Err.Number = 0 Then
                counter = counter + 1
            Else
                Fehler = Fehler & "  " & file.Name & " (" & Err.Description & ")" & vbCrLf
                Err.Clear
            End If

            On Error GoTo 0
        End If
    Next file
End Sub


' ===============================================================
' Bereinigt Doubletten (mod_XYZ1, mod_XYZ2, ...)
' ===============================================================
Private Sub BereinigeDoubletten(fso As Object, vbProj As Object, _
                                 ByRef countEntfernt As Long, _
                                 ByRef Fehler As String)

    Dim repoNamen As Object
    Set repoNamen = CreateObject("Scripting.Dictionary")
    repoNamen.CompareMode = vbTextCompare

    Dim folder As Object, file As Object
    For Each folder In Array(fso.GetFolder(REPO_PATH_MODULES), _
                              fso.GetFolder(REPO_PATH_CLASSES), _
                              fso.GetFolder(REPO_PATH_USERFORMS))
        For Each file In folder.Files
            Dim ext As String, baseName As String
            ext = LCase(fso.GetExtensionName(file.Name))
            baseName = fso.GetBaseName(file.Name)
            If ext = "bas" Or ext = "cls" Or ext = "frm" Then
                If Not repoNamen.Exists(baseName) Then repoNamen.Add baseName, True
            End If
        Next file
    Next folder

    Dim vbComp As Object
    Dim toRemove As Collection
    Set toRemove = New Collection

    For Each vbComp In vbProj.VBComponents
        If vbComp.Type <> 100 Then
            Dim compName As String
            compName = vbComp.Name
            Dim basisName As String
            basisName = EntferneNachgestellteZiffern(compName)
            If basisName <> compName Then
                If repoNamen.Exists(basisName) Then
                    toRemove.Add vbComp
                End If
            End If
        End If
    Next vbComp

    Dim comp As Object
    For Each comp In toRemove
        On Error Resume Next
        vbProj.VBComponents.Remove comp
        If Err.Number = 0 Then
            countEntfernt = countEntfernt + 1
        Else
            Fehler = Fehler & "  Doublette " & comp.Name & " (" & Err.Description & ")" & vbCrLf
            Err.Clear
        End If
        On Error GoTo 0
    Next comp
End Sub

Private Function EntferneNachgestellteZiffern(name As String) As String
    Dim i As Long
    For i = Len(name) To 1 Step -1
        If Not IsNumeric(Mid(name, i, 1)) Then
            EntferneNachgestellteZiffern = Left(name, i)
            Exit Function
        End If
    Next i
    EntferneNachgestellteZiffern = name
End Function


' ===============================================================
' Ersetzt den kompletten Code eines VBComponent-CodeModules
' ===============================================================
Private Sub ErsetzeCodeInDokumentModul(vbComp As Object, neuerCode As String)
    With vbComp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(neuerCode) > 0 Then .AddFromString neuerCode
    End With
End Sub


' ===============================================================
' Liest Datei und entfernt VBA-Header-Zeilen ("Attribute VB_...",
' "VERSION 1.0 CLASS", "BEGIN" ... "END", "MultiUse = -1", etc.)
' ===============================================================
Private Function LeseDateiOhneKlassenHeader(dateipfad As String) As String
    Dim gesamtInhalt As String
    gesamtInhalt = LeseDateiMitEncodingErkennung(dateipfad)

    Dim zeilen() As String
    zeilen = Split(gesamtInhalt, vbCrLf)

    Dim ausgabe As String
    Dim i As Long
    Dim inKlassenHeader As Boolean
    inKlassenHeader = False
    Dim headerFertig As Boolean
    headerFertig = False

    For i = 0 To UBound(zeilen)
        Dim zeile As String
        zeile = zeilen(i)

        If Not headerFertig Then
            Dim trimZeile As String
            trimZeile = Trim(zeile)

            If Left(trimZeile, 10) = "VERSION 1." And InStr(trimZeile, "CLASS") > 0 Then
                inKlassenHeader = True
            ElseIf Left(trimZeile, 5) = "BEGIN" Then
                inKlassenHeader = True
            ElseIf trimZeile = "END" And inKlassenHeader Then
                inKlassenHeader = False
                GoTo NaechsteZeile
            ElseIf Left(trimZeile, 10) = "Attribute " Then
                GoTo NaechsteZeile
            ElseIf inKlassenHeader Then
                GoTo NaechsteZeile
            ElseIf Left(trimZeile, 7) = "Option " Or _
                   Left(trimZeile, 4) = "Sub " Or _
                   Left(trimZeile, 8) = "Public " Or _
                   Left(trimZeile, 8) = "Private " Or _
                   Left(trimZeile, 4) = "Dim " Or _
                   Left(trimZeile, 6) = "Const " Or _
                   Left(trimZeile, 9) = "Function " Or _
                   Left(trimZeile, 5) = "Type " Or _
                   Left(trimZeile, 5) = "Enum " Or _
                   Left(trimZeile, 1) = "'" Or _
                   Left(trimZeile, 1) = "#" Then
                headerFertig = True
                ausgabe = ausgabe & zeile & vbCrLf
                GoTo NaechsteZeile
            ElseIf Len(trimZeile) = 0 Then
                GoTo NaechsteZeile
            End If
        Else
            ausgabe = ausgabe & zeile & vbCrLf
        End If

NaechsteZeile:
    Next i

    If Right(ausgabe, 2) = vbCrLf Then ausgabe = Left(ausgabe, Len(ausgabe) - 2)
    LeseDateiOhneKlassenHeader = ausgabe
End Function


' ===============================================================
' Liest eine Datei mit robuster Encoding-Erkennung
' ---------------------------------------------------------------
' STRATEGIE (v3.3, BOM-first):
'   1. BOM EF BB BF -> UTF-8 mit BOM (Standard-Repo-Format)
'   2. Sonst als UTF-8 lesen (Charset "utf-8" liefert bei
'      ungueltigen Sequenzen U+FFFD als Marker)
'   3. Wenn U+FFFD im Ergebnis -> Datei war ANSI, als
'      Windows-1252 neu lesen
'   4. Sonst UTF-8-Ergebnis verwenden
'
' Fruehere Versionen suchten mit InStr(text, "?") — das war
' buggy, weil legitime "?"-Zeichen in Strings/Kommentaren das
' ANSI-Fallback faelschlich triggerten und Umlaute zerstoerten.
' ===============================================================
Private Function LeseDateiMitEncodingErkennung(dateipfad As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    ' ---- Schritt 1: BOM-Pruefung (Binaer, ohne Dekodierung) ----
    Dim hatBOM As Boolean
    hatBOM = False
    On Error Resume Next
    Dim byteStream As Object
    Set byteStream = CreateObject("ADODB.Stream")
    byteStream.Type = 1                     ' adTypeBinary
    byteStream.Open
    byteStream.LoadFromFile dateipfad
    If byteStream.Size >= 3 Then
        Dim bomBytes() As Byte
        bomBytes = byteStream.Read(3)
        If bomBytes(0) = &HEF And bomBytes(1) = &HBB And bomBytes(2) = &HBF Then
            hatBOM = True
        End If
    End If
    byteStream.Close
    On Error GoTo 0

    ' ---- Schritt 2: als UTF-8 lesen ---------------------------
    Dim utf8Inhalt As String
    With stream
        .Type = 2                           ' adTypeText
        .Charset = "utf-8"
        .Open
        .LoadFromFile dateipfad
        utf8Inhalt = .ReadText(-1)
        .Close
    End With

    ' fuehrendes BOM-Zeichen entfernen (ADODB liest es als U+FEFF)
    If Len(utf8Inhalt) > 0 Then
        If AscW(Left(utf8Inhalt, 1)) = &HFEFF Then
            utf8Inhalt = Mid(utf8Inhalt, 2)
        End If
    End If

    ' Wenn BOM vorhanden war -> Ergebnis ist definitiv korrekt
    If hatBOM Then
        LeseDateiMitEncodingErkennung = utf8Inhalt
        Exit Function
    End If

    ' ---- Schritt 3: FFFD-Check (U+FFFD = ungueltiges UTF-8) --
    ' Wenn FFFD im Ergebnis vorkommt, war die Datei nicht UTF-8.
    If InStr(utf8Inhalt, ChrW(&HFFFD)) > 0 Then
        Dim ansiInhalt As String
        With stream
            .Charset = "windows-1252"
            .Open
            .LoadFromFile dateipfad
            ansiInhalt = .ReadText(-1)
            .Close
        End With
        Debug.Print "[Sync] Encoding: ANSI (Windows-1252) fuer " & _
                    Mid(dateipfad, InStrRev(dateipfad, "\") + 1)
        LeseDateiMitEncodingErkennung = ansiInhalt
        Exit Function
    End If

    ' Kein BOM, aber gueltiges UTF-8 (oder reines ASCII)
    LeseDateiMitEncodingErkennung = utf8Inhalt
End Function


' ===============================================================
' Konvertiert eine Datei in ANSI-Temp-Kopie (Windows-1252)
' ---------------------------------------------------------------
' Wird nur beim NEU-Import per VBComponents.Import benoetigt.
' Existierende Module gehen ueber ErsetzeCodeInDokumentModul und
' brauchen keine Encoding-Konvertierung.
' ===============================================================
Private Function KonvertiereUTF8zuAnsi(quellPfad As String, _
                                        zielPfad As String, _
                                        fso As Object) As String
    On Error GoTo FallbackKopie

    Dim inhalt As String
    inhalt = LeseDateiMitEncodingErkennung(quellPfad)

    ' BOM entfernen (falls doch noch drin)
    If Len(inhalt) > 0 Then
        If AscW(Left(inhalt, 1)) = &HFEFF Then inhalt = Mid(inhalt, 2)
    End If

    ' Explizit als Windows-1252 schreiben (KEIN FSO.CreateTextFile —
    ' das nutzt die System-Codepage und ist unzuverlaessig).
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "Windows-1252"
    st.Open
    st.WriteText inhalt
    st.SaveToFile zielPfad, 2               ' adSaveCreateOverWrite
    st.Close
    Set st = Nothing

    KonvertiereUTF8zuAnsi = zielPfad
    Exit Function

FallbackKopie:
    On Error Resume Next
    fso.CopyFile quellPfad, zielPfad, True
    If Err.Number = 0 Then
        KonvertiereUTF8zuAnsi = zielPfad
    Else
        KonvertiereUTF8zuAnsi = ""
    End If
    On Error GoTo 0
End Function


' ===============================================================
' KURZBEFEHLE FUER DAS DIREKTFENSTER (v3.3)
' ---------------------------------------------------------------
' Alle drei sind Public Subs ohne Parameter -- im Direktfenster
' einfach den Namen tippen und Enter druecken (KEIN "?" davor!).
'
'   RepoNachExcel        Repo -> Excel  (Import)
'   ExcelNachRepo        Excel -> Repo  (Export, UTF-8+BOM)
'   SyncMitVerification  Legacy-Alias fuer RepoNachExcel
'
' WICHTIG: Nach RepoNachExcel die Mappe mit Strg+S speichern,
' sonst sind die neuen Subs beim naechsten Oeffnen wieder weg.
' ===============================================================
Public Sub RepoNachExcel()
    Call SyncVBAVomRepository
End Sub

Public Sub ExcelNachRepo()
    Call mod_VBA_Export.ExportiereAlleVBAKomponenten
End Sub

Public Sub SyncMitVerification()
    Call SyncVBAVomRepository
End Sub
