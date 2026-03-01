Attribute VB_Name = "mod_Repo_Sync"
Option Explicit

' ***************************************************************
' MODUL: mod_Repo_Sync
' VERSION: 3.0 - 01.03.2026
' ZWECK: Importiert ALLE VBA-Komponenten aus dem Repository
'        inkl. Dokument-Module (DieseArbeitsmappe, TabelleX)
'
'        Unterstützte Dateitypen:
'        -------------------------------------------------------
'        .bas  Standard-Module:
'              - Existierende: CodeModule-Ersetzung (in-place)
'              - Neue: Import nach ANSI-Konvertierung
'        .cls  Klassen-Module:
'              - Dokument-Module (Type=100): CodeModule-Ersetzung
'              - Reguläre Klassen: CodeModule-Ersetzung (in-place)
'              - Neue Klassen: Import nach ANSI-Konvertierung
'        .frm  UserForms: Löschen + Neu-Import (inkl. .frx)
'
'        STRATEGIE (v3.0 - "CodeModule first"):
'        1. BEREINIGUNG: Doubletten entfernen (mod_XYZ1,
'           mod_XYZ2 usw.), die durch fehlgeschlagene
'           Remove+Import-Zyklen entstanden sind.
'        2. IMPORT: Für bestehende Module wird der Code
'           direkt im CodeModule überschrieben (DeleteLines +
'           AddFromString). Kein Remove nötig, daher kein
'           "Zugriff verweigert". Nur für NEUE Module wird
'           VBComponents.Import nach ANSI-Konvertierung
'           verwendet.
'
'        ENCODING:
'        Dateien aus dem Repo (VS Code) sind UTF-8 kodiert.
'        VBA erwartet für den Import ANSI (Windows-1252).
'        Dieses Modul konvertiert automatisch UTF-8 → ANSI,
'        damit Umlaute (ä, ö, ü, ß) korrekt übernommen werden.
'
' HINWEIS: Dieses Modul und mod_VBA_Export werden beim Import
'          übersprungen, um sich nicht selbst zu überschreiben.
' ***************************************************************

' ===============================================================
' QUELLORDNER FÜR IMPORT (REPOSITORY)
' ===============================================================
Private Const REPO_PATH_CLASSES As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\Classes\"
Private Const REPO_PATH_USERFORMS As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\UserForms\"
Private Const REPO_PATH_MODULES As String = "C:\Users\DELL Latitude 7490\Desktop\Mein Projekt\vba\Modules\"

' Temporärer Unterordner für ANSI-konvertierte Dateien
Private Const TEMP_SUBFOLDER As String = "VBA_Repo_Sync_Temp"


' ===============================================================
' HAUPTPROZEDUR: Synchronisiert das VBA-Projekt mit dem Repo
' ===============================================================
Public Sub SyncVBAVomRepository()
    
    Dim vbProj As Object
    Dim fso As Object
    Dim tempPfad As String
    
    Dim countModules As Long
    Dim countKlassen As Long
    Dim countForms As Long
    Dim countDokumente As Long
    Dim countDoubletten As Long
    Dim fehlerListe As String
    
    On Error GoTo ErrorHandler
    
    ' ---------------------------------------------------------
    ' 1. Prüfe Zugriff auf VBA-Projekt
    ' ---------------------------------------------------------
    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "FEHLER: Zugriff auf VBA-Projekt nicht erlaubt!" & vbCrLf & vbCrLf & _
               "Bitte aktiviere in Excel:" & vbCrLf & _
               "Datei > Optionen > Trust Center > Einstellungen f" & ChrW(252) & "r das Trust Center" & vbCrLf & _
               "> Makroeinstellungen > 'Zugriff auf das VBA-Projektobjektmodell vertrauen'", _
               vbCritical, "Zugriff verweigert"
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ---------------------------------------------------------
    ' 2. Prüfe ob Quellordner existieren
    ' ---------------------------------------------------------
    If Not fso.FolderExists(REPO_PATH_CLASSES) Or _
       Not fso.FolderExists(REPO_PATH_USERFORMS) Or _
       Not fso.FolderExists(REPO_PATH_MODULES) Then
        MsgBox "FEHLER: Mindestens ein Quellordner im Repo wurde nicht gefunden!" & vbCrLf & _
               "Erwartete Pfade:" & vbCrLf & _
               "  " & REPO_PATH_MODULES & vbCrLf & _
               "  " & REPO_PATH_CLASSES & vbCrLf & _
               "  " & REPO_PATH_USERFORMS, _
               vbCritical, "Ordner fehlt"
        Exit Sub
    End If
    
    ' ---------------------------------------------------------
    ' 3. Temporären Ordner für ANSI-Konvertierung erstellen
    ' ---------------------------------------------------------
    tempPfad = Environ("TEMP") & "\" & TEMP_SUBFOLDER & "\"
    On Error Resume Next
    If fso.FolderExists(tempPfad) Then
        fso.DeleteFolder Left(tempPfad, Len(tempPfad) - 1), True
    End If
    fso.CreateFolder tempPfad
    On Error GoTo ErrorHandler
    
    ' Zähler initialisieren
    countModules = 0
    countKlassen = 0
    countForms = 0
    countDokumente = 0
    countDoubletten = 0
    fehlerListe = ""
    
    ' ---------------------------------------------------------
    ' BEREINIGUNG: Doubletten entfernen
    '   Durch mehrfaches Syncen entstandene Kopien wie
    '   mod_Format_Spalten1, mod_KategorieRegeln2 usw.
    ' ---------------------------------------------------------
    Application.StatusBar = "VBA-Sync: Bereinige Doubletten..."
    BereinigeDoubletten fso, vbProj, countDoubletten, fehlerListe
    
    Application.StatusBar = "VBA-Sync: Importiere Standard-Module..."
    
    ' ---------------------------------------------------------
    ' 4. IMPORT: Standard-Module (.bas)
    ' ---------------------------------------------------------
    ImportiereStandardDateien fso, vbProj, REPO_PATH_MODULES, "bas", tempPfad, countModules, fehlerListe
    
    Application.StatusBar = "VBA-Sync: Importiere Klassen-Module..."
    
    ' ---------------------------------------------------------
    ' 5. IMPORT: Klassen-Module (.cls) inkl. Dokument-Module
    ' ---------------------------------------------------------
    ImportiereKlassenModule fso, vbProj, REPO_PATH_CLASSES, tempPfad, countKlassen, countDokumente, fehlerListe
    
    Application.StatusBar = "VBA-Sync: Importiere UserForms..."
    
    ' ---------------------------------------------------------
    ' 6. IMPORT: UserForms (.frm + .frx)
    ' ---------------------------------------------------------
    ImportiereUserForms fso, vbProj, REPO_PATH_USERFORMS, tempPfad, countForms, fehlerListe
    
    ' ---------------------------------------------------------
    ' 7. Temporären Ordner aufräumen
    ' ---------------------------------------------------------
    On Error Resume Next
    If fso.FolderExists(tempPfad) Then
        fso.DeleteFolder Left(tempPfad, Len(tempPfad) - 1), True
    End If
    On Error GoTo ErrorHandler
    
    Application.StatusBar = False
    
    ' ---------------------------------------------------------
    ' 8. Ergebnis anzeigen
    ' ---------------------------------------------------------
    Dim msg As String
    msg = "VBA-Synchronisierung abgeschlossen! (v3.0)" & vbCrLf & vbCrLf
    If countDoubletten > 0 Then
        msg = msg & "Bereinigt:" & vbCrLf
        msg = msg & "  Doubletten entfernt: " & countDoubletten & vbCrLf & vbCrLf
    End If
    msg = msg & "Importiert aus Repo:" & vbCrLf
    msg = msg & "  Standard-Module:  " & countModules & vbCrLf
    msg = msg & "  Klassen-Module:   " & countKlassen & vbCrLf
    msg = msg & "  Dokument-Module:  " & countDokumente & " (Code " & ChrW(252) & "berschrieben)" & vbCrLf
    msg = msg & "  UserForms:        " & countForms & vbCrLf
    
    If fehlerListe <> "" Then
        msg = msg & vbCrLf & "FEHLER bei folgenden Dateien:" & vbCrLf & fehlerListe
    End If
    
    msg = msg & vbCrLf & "Das Projekt ist nun auf dem Stand des Repositories." & vbCrLf & _
          "WICHTIG: Bitte f" & ChrW(252) & "hre jetzt 'Debuggen > Kompilieren' aus."
    
    If fehlerListe <> "" Then
        MsgBox msg, vbExclamation, "Synchronisierung mit Warnungen"
    Else
        MsgBox msg, vbInformation, "Synchronisierung erfolgreich"
    End If
    
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    
    ' Temporären Ordner aufräumen bei Fehler
    On Error Resume Next
    If Not fso Is Nothing Then
        Dim tmpClean As String
        tmpClean = Environ("TEMP") & "\" & TEMP_SUBFOLDER
        If fso.FolderExists(tmpClean) Then fso.DeleteFolder tmpClean, True
    End If
    On Error GoTo 0
    
    MsgBox "Unerwarteter Fehler beim Import:" & vbCrLf & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description, _
           vbCritical, "Sync fehlgeschlagen"
End Sub


' ===============================================================
' Importiert Standard-Module (.bas) aus dem Repository
'
' STRATEGIE (v2.2 - "CodeModule first"):
'   - Existierendes Modul: Code wird direkt im CodeModule
'     überschrieben (DeleteLines + AddFromString).
'     KEIN Remove, dadurch kein "Zugriff verweigert".
'   - Neues Modul: ANSI-konvertierte Datei wird importiert.
' ===============================================================
Private Sub ImportiereStandardDateien(fso As Object, vbProj As Object, _
                                      pfad As String, ext As String, _
                                      tempPfad As String, _
                                      ByRef counter As Long, _
                                      ByRef fehler As String)
    Dim folder As Object
    Dim file As Object
    Dim compName As String
    Dim vbComp As Object
    Dim ansiDatei As String
    
    Set folder = fso.GetFolder(pfad)
    
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = LCase(ext) Then
            compName = fso.GetBaseName(file.Name)
            
            ' Überspringe dieses Modul selbst und den Exporteur
            If compName = "mod_Repo_Sync" Or compName = "mod_VBA_Export" Then
                GoTo NaechsteStandardDatei
            End If
            
            On Error Resume Next
            
            ' Bestehende Komponente suchen
            Set vbComp = Nothing
            Set vbComp = vbProj.VBComponents(compName)
            Err.Clear
            On Error GoTo 0
            
            If Not vbComp Is Nothing Then
                ' -----------------------------------------------
                ' MODUL EXISTIERT → Code direkt überschreiben
                ' (kein Remove nötig, funktioniert auch zur Laufzeit)
                ' -----------------------------------------------
                If ErsetzeCodeInDokumentModul(vbComp, file.Path) Then
                    counter = counter + 1
                Else
                    fehler = fehler & "  " & file.Name & _
                             " (Code-Ersetzung fehlgeschlagen)" & vbCrLf
                End If
            Else
                ' -----------------------------------------------
                ' MODUL EXISTIERT NICHT → Neu importieren
                ' -----------------------------------------------
                On Error Resume Next
                ansiDatei = KonvertiereUTF8zuAnsi(file.Path, tempPfad & file.Name, fso)
                
                If ansiDatei <> "" Then
                    vbProj.VBComponents.Import ansiDatei
                Else
                    ' Fallback: Original importieren
                    vbProj.VBComponents.Import file.Path
                End If
                
                If Err.Number = 0 Then
                    counter = counter + 1
                Else
                    fehler = fehler & "  " & file.Name & " (" & Err.Description & ")" & vbCrLf
                    Err.Clear
                End If
                On Error GoTo 0
            End If
            
NaechsteStandardDatei:
        End If
    Next file
End Sub


' ===============================================================
' Importiert Klassen-Module (.cls) aus dem Repository
'
' STRATEGIE (v2.2 - "CodeModule first"):
'   - Dokument-Module (Type=100): CodeModule-Ersetzung (einzige Option)
'   - Reguläre Klassen (existierend): CodeModule-Ersetzung (in-place)
'   - Neue Klassen: ANSI-konvertierte Datei wird importiert
' ===============================================================
Private Sub ImportiereKlassenModule(fso As Object, vbProj As Object, _
                                     pfad As String, tempPfad As String, _
                                     ByRef countKlassen As Long, _
                                     ByRef countDokumente As Long, _
                                     ByRef fehler As String)
    Dim folder As Object
    Dim file As Object
    Dim compName As String
    Dim vbComp As Object
    Dim ansiDatei As String
    
    Set folder = fso.GetFolder(pfad)
    
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "cls" Then
            compName = fso.GetBaseName(file.Name)
            
            On Error Resume Next
            Set vbComp = Nothing
            Set vbComp = vbProj.VBComponents(compName)
            Err.Clear
            On Error GoTo 0
            
            If Not vbComp Is Nothing Then
                ' -----------------------------------------------
                ' KOMPONENTE EXISTIERT → Code direkt überschreiben
                ' (funktioniert für Type=100 UND reguläre Klassen)
                ' -----------------------------------------------
                If ErsetzeCodeInDokumentModul(vbComp, file.Path) Then
                    If vbComp.Type = 100 Then
                        countDokumente = countDokumente + 1
                    Else
                        countKlassen = countKlassen + 1
                    End If
                Else
                    fehler = fehler & "  " & file.Name & _
                             " (Code-Ersetzung fehlgeschlagen)" & vbCrLf
                End If
            Else
                ' -----------------------------------------------
                ' KLASSE EXISTIERT NICHT → Neu importieren
                ' -----------------------------------------------
                On Error Resume Next
                ansiDatei = KonvertiereUTF8zuAnsi(file.Path, tempPfad & file.Name, fso)
                If ansiDatei <> "" Then
                    vbProj.VBComponents.Import ansiDatei
                Else
                    vbProj.VBComponents.Import file.Path
                End If
                
                If Err.Number = 0 Then
                    countKlassen = countKlassen + 1
                Else
                    fehler = fehler & "  " & file.Name & " (" & Err.Description & ")" & vbCrLf
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
    Next file
End Sub


' ===============================================================
' Importiert UserForms (.frm + .frx) aus dem Repository
' Die .frm-Datei wird UTF-8 → ANSI konvertiert.
' Die zugehörige .frx-Datei (Binärdaten der Steuerelemente)
' wird direkt in den Temp-Ordner kopiert, da VBA beim Import
' beide Dateien im selben Ordner erwartet.
' ===============================================================
Private Sub ImportiereUserForms(fso As Object, vbProj As Object, _
                                 pfad As String, tempPfad As String, _
                                 ByRef counter As Long, _
                                 ByRef fehler As String)
    Dim folder As Object
    Dim file As Object
    Dim compName As String
    Dim vbComp As Object
    Dim ansiDatei As String
    Dim frxQuelle As String
    Dim frxZiel As String
    
    Set folder = fso.GetFolder(pfad)
    
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "frm" Then
            compName = fso.GetBaseName(file.Name)
            
            ' Überspringe diese Module
            If compName = "mod_Repo_Sync" Or compName = "mod_VBA_Export" Then
                GoTo NaechsteForm
            End If
            
            On Error Resume Next
            
            ' Bestehende UserForm löschen (falls vorhanden)
            Set vbComp = Nothing
            Set vbComp = vbProj.VBComponents(compName)
            If Not vbComp Is Nothing Then
                vbProj.VBComponents.Remove vbComp
                DoEvents  ' VBA Zeit geben, die Löschung zu verarbeiten
            End If
            Err.Clear
            
            ' .frm-Datei konvertieren (UTF-8 → ANSI)
            ansiDatei = KonvertiereUTF8zuAnsi(file.Path, tempPfad & file.Name, fso)
            
            ' .frx-Datei (Binärdaten) in den Temp-Ordner kopieren
            frxQuelle = fso.BuildPath(pfad, compName & ".frx")
            frxZiel = tempPfad & compName & ".frx"
            If fso.FileExists(frxQuelle) Then
                fso.CopyFile frxQuelle, frxZiel, True
            End If
            Err.Clear
            
            ' Import der konvertierten .frm (+ .frx im selben Ordner)
            If ansiDatei <> "" Then
                vbProj.VBComponents.Import ansiDatei
            Else
                ' Fallback: direkt aus Repo importieren
                vbProj.VBComponents.Import file.Path
            End If
            
            If Err.Number = 0 Then
                counter = counter + 1
            Else
                fehler = fehler & "  " & file.Name & " (" & Err.Description & ")" & vbCrLf
                Err.Clear
            End If
            
            On Error GoTo 0
            
NaechsteForm:
        End If
    Next file
End Sub


' ===============================================================
' Bereinigt Doubletten im VBA-Projekt
'
' Wenn VBA beim Import eine Komponente nicht loeschen konnte
' (z.B. wegen "Zugriff verweigert"), erstellt es beim Import
' automatisch eine Kopie mit angehaengter Ziffer:
'   mod_Format_Spalten  -> mod_Format_Spalten1 (Doublette)
'   mod_KategorieRegeln -> mod_KategorieRegeln3 (Doublette)
'
' Diese Prozedur findet solche Doubletten und entfernt sie:
'   1. Sammelt alle erwarteten Modulnamen aus dem Repository
'   2. Iteriert alle VBComponents im Projekt
'   3. Wenn ein Modul-Name = Basisname + Ziffern ist und
'      der Basisname im Repo existiert -> Doublette -> entfernen
' ===============================================================
Private Sub BereinigeDoubletten(fso As Object, vbProj As Object, _
                                 ByRef countEntfernt As Long, _
                                 ByRef fehler As String)
    
    ' 1. Alle erwarteten Modulnamen aus dem Repo sammeln
    Dim repoNamen As Object
    Set repoNamen = CreateObject("Scripting.Dictionary")
    repoNamen.CompareMode = vbTextCompare
    
    Dim folder As Object, file As Object
    
    ' Standard-Module (.bas)
    If fso.FolderExists(REPO_PATH_MODULES) Then
        Set folder = fso.GetFolder(REPO_PATH_MODULES)
        For Each file In folder.Files
            If LCase(fso.GetExtensionName(file.Name)) = "bas" Then
                repoNamen(fso.GetBaseName(file.Name)) = True
            End If
        Next file
    End If
    
    ' Klassen-Module (.cls)
    If fso.FolderExists(REPO_PATH_CLASSES) Then
        Set folder = fso.GetFolder(REPO_PATH_CLASSES)
        For Each file In folder.Files
            If LCase(fso.GetExtensionName(file.Name)) = "cls" Then
                repoNamen(fso.GetBaseName(file.Name)) = True
            End If
        Next file
    End If
    
    ' UserForms (.frm)
    If fso.FolderExists(REPO_PATH_USERFORMS) Then
        Set folder = fso.GetFolder(REPO_PATH_USERFORMS)
        For Each file In folder.Files
            If LCase(fso.GetExtensionName(file.Name)) = "frm" Then
                repoNamen(fso.GetBaseName(file.Name)) = True
            End If
        Next file
    End If
    
    ' 2. Alle VBComponents pruefen und Doubletten sammeln
    Dim vbComp As Object
    Dim compName As String
    Dim basisName As String
    Dim zuEntfernen As New Collection
    
    For Each vbComp In vbProj.VBComponents
        compName = vbComp.Name
        
        ' Dokument-Module (DieseArbeitsmappe, TabelleX) nie entfernen
        If vbComp.Type = 100 Then GoTo NaechsteKomponente
        
        ' Wenn der Name ein bekannter Repo-Name ist -> kein Doublette
        If repoNamen.Exists(compName) Then GoTo NaechsteKomponente
        
        ' Pruefen ob der Name = Basisname + Ziffern ist
        basisName = EntferneNachgestellteZiffern(compName)
        
        ' Wenn Basisname anders UND Basisname existiert im Repo -> Doublette!
        If basisName <> compName And repoNamen.Exists(basisName) Then
            zuEntfernen.Add vbComp
        End If
        
NaechsteKomponente:
    Next vbComp
    
    ' 3. Gesammelte Doubletten entfernen
    Dim i As Long
    For i = 1 To zuEntfernen.Count
        Set vbComp = zuEntfernen(i)
        On Error Resume Next
        vbProj.VBComponents.Remove vbComp
        DoEvents
        If Err.Number = 0 Then
            countEntfernt = countEntfernt + 1
        Else
            ' Fallback: Code leeren (Modul bleibt, stoert aber nicht mehr)
            Err.Clear
            With vbComp.CodeModule
                If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
            End With
            If Err.Number = 0 Then
                countEntfernt = countEntfernt + 1
            Else
                fehler = fehler & "  " & vbComp.Name & _
                         " (Doublette konnte nicht entfernt werden)" & vbCrLf
                Err.Clear
            End If
        End If
        On Error GoTo 0
    Next i
    
End Sub


' ===============================================================
' Entfernt nachgestellte Ziffern von einem Komponentennamen.
' Beispiele:
'   "mod_Format_Spalten12" -> "mod_Format_Spalten"
'   "mod_KategorieRegeln3" -> "mod_KategorieRegeln"
'   "mod_Format_Spalten"   -> "mod_Format_Spalten" (unveraendert)
' ===============================================================
Private Function EntferneNachgestellteZiffern(ByVal modulName As String) As String
    Do While Len(modulName) > 0 And IsNumeric(Right(modulName, 1))
        modulName = Left(modulName, Len(modulName) - 1)
    Loop
    EntferneNachgestellteZiffern = modulName
End Function


' ===============================================================
' Ersetzt den Code in einem bestehenden VBA-Modul (in-place).
' Funktioniert für ALLE Modultypen:
'   - Dokument-Module (Type=100)
'   - Standard-Module (.bas)
'   - Reguläre Klassen (.cls)
'
' Liest die Datei als UTF-8, entfernt den Header
' (VERSION, BEGIN...END, Attribute-Zeilen) und schreibt den
' eigentlichen Code direkt ins bestehende CodeModule.
' Dadurch wird kein Remove benötigt → kein "Zugriff verweigert".
'
' Rückgabe: True bei Erfolg, False bei Fehler
' ===============================================================
Private Function ErsetzeCodeInDokumentModul(vbComp As Object, _
                                             dateipfad As String) As Boolean
    
    On Error GoTo FehlerHandler
    
    ' Dateiinhalt als UTF-8 lesen und Header entfernen
    Dim codeInhalt As String
    codeInhalt = LeseDateiOhneKlassenHeader(dateipfad)
    
    ' Bestehenden Code komplett löschen
    With vbComp.CodeModule
        If .CountOfLines > 0 Then
            .DeleteLines 1, .CountOfLines
        End If
        
        ' Neuen Code einfügen (nur wenn Inhalt vorhanden)
        If Len(Trim(codeInhalt)) > 0 Then
            .AddFromString codeInhalt
        End If
    End With
    
    ErsetzeCodeInDokumentModul = True
    Exit Function
    
FehlerHandler:
    ErsetzeCodeInDokumentModul = False
End Function


' ===============================================================
' Liest eine .cls-Datei als UTF-8 und entfernt den Klassen-Header:
'   VERSION 1.0 CLASS
'   BEGIN
'     MultiUse = -1  'True
'   END
'   Attribute VB_Name = "..."
'   Attribute VB_GlobalNameSpace = ...
'   Attribute VB_Creatable = ...
'   Attribute VB_PredeclaredId = ...
'   Attribute VB_Exposed = ...
'
' Gibt nur den eigentlichen Code zurück (ab "Option Explicit").
' Verwendet ADODB.Stream für korrekte UTF-8-Behandlung.
' ===============================================================
Private Function LeseDateiOhneKlassenHeader(dateipfad As String) As String
    
    Dim stream As Object
    Dim gesamtInhalt As String
    Dim zeilen() As String
    Dim i As Long
    Dim codeStart As Long
    Dim ergebnis As String
    
    ' Datei als UTF-8 lesen
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2          ' adTypeText
        .Charset = "utf-8"
        .Open
        .LoadFromFile dateipfad
        gesamtInhalt = .ReadText(-1)  ' adReadAll
        .Close
    End With
    
    ' BOM entfernen (falls vorhanden)
    If Len(gesamtInhalt) > 0 Then
        If AscW(Left(gesamtInhalt, 1)) = &HFEFF Then
            gesamtInhalt = Mid(gesamtInhalt, 2)
        End If
    End If
    
    ' In Zeilen aufteilen (Windows: vbCrLf, Unix: vbLf)
    If InStr(gesamtInhalt, vbCrLf) > 0 Then
        zeilen = Split(gesamtInhalt, vbCrLf)
    Else
        zeilen = Split(gesamtInhalt, vbLf)
    End If
    
    ' Header-Zeilen überspringen und erste Code-Zeile finden
    codeStart = -1
    For i = LBound(zeilen) To UBound(zeilen)
        Dim trimZeile As String
        trimZeile = Trim(zeilen(i))
        
        ' Bekannte Header-Zeilen überspringen
        If Left(trimZeile, 7) = "VERSION" Then GoTo WeiterSuchen
        If Left(trimZeile, 5) = "BEGIN" Then GoTo WeiterSuchen
        If trimZeile = "END" Then GoTo WeiterSuchen
        If Left(trimZeile, 8) = "MultiUse" Then GoTo WeiterSuchen
        If Left(trimZeile, 9) = "Attribute" Then GoTo WeiterSuchen
        
        ' Leerzeilen zwischen Header und Code überspringen
        If trimZeile = "" And codeStart = -1 Then GoTo WeiterSuchen
        
        ' Erste echte Code-Zeile gefunden!
        codeStart = i
        Exit For
        
WeiterSuchen:
    Next i
    
    ' Code ab der gefundenen Startzeile zusammenbauen
    If codeStart >= 0 Then
        ergebnis = ""
        For i = codeStart To UBound(zeilen)
            If i > codeStart Then ergebnis = ergebnis & vbCrLf
            ergebnis = ergebnis & zeilen(i)
        Next i
    Else
        ergebnis = ""
    End If
    
    LeseDateiOhneKlassenHeader = ergebnis
    
End Function


' ===============================================================
' Konvertiert eine UTF-8-Datei in eine ANSI-Datei (Windows-1252)
' Dies ist nötig, weil VBComponents.Import ANSI erwartet und
' VS Code die Dateien als UTF-8 speichert.
'
' Ablauf:
'   1. Quelldatei mit ADODB.Stream als UTF-8 lesen
'   2. Inhalt mit FSO als ANSI (System-Codepage) schreiben
'
' Rückgabe: Pfad der ANSI-Datei, oder "" bei Fehler
' ===============================================================
Private Function KonvertiereUTF8zuAnsi(quellPfad As String, _
                                        zielPfad As String, _
                                        fso As Object) As String
    
    On Error GoTo FallbackKopie
    
    ' Quelldatei als UTF-8 lesen
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2          ' adTypeText
        .Charset = "utf-8"
        .Open
        .LoadFromFile quellPfad
        Dim inhalt As String
        inhalt = .ReadText(-1)  ' adReadAll
        .Close
    End With
    
    ' BOM entfernen (falls vorhanden)
    If Len(inhalt) > 0 Then
        If AscW(Left(inhalt, 1)) = &HFEFF Then
            inhalt = Mid(inhalt, 2)
        End If
    End If
    
    ' Als ANSI schreiben (FSO nutzt System-Codepage = Windows-1252)
    Dim ts As Object
    Set ts = fso.CreateTextFile(zielPfad, True, False)  ' Overwrite, NICHT Unicode
    ts.Write inhalt
    ts.Close
    
    KonvertiereUTF8zuAnsi = zielPfad
    Exit Function
    
FallbackKopie:
    ' Fallback: Datei direkt kopieren (kein Encoding-Wechsel)
    On Error Resume Next
    fso.CopyFile quellPfad, zielPfad, True
    If Err.Number = 0 Then
        KonvertiereUTF8zuAnsi = zielPfad
    Else
        KonvertiereUTF8zuAnsi = ""
    End If
    On Error GoTo 0
End Function

