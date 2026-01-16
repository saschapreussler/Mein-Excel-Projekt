VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Mitgliederverwaltung 
   Caption         =   "Mitgliederverwaltung"
   ClientHeight    =   9720.001
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   18740
   OleObjectBlob   =   "frm_Mitgliederverwaltung.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frm_Mitgliederverwaltung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==========================================================
' Prozedur: cmd_MitgliedEdit_Click
' ==========================================================
Private Sub cmd_MitgliedEdit_Click()
    If Me.lst_Mitgliederliste.ListIndex >= 0 Then
        Call OeffneMitgliedsDetails
    Else
        MsgBox "Bitte wählen Sie zuerst ein Mitglied aus der Liste aus.", vbExclamation
    End If
End Sub

' ==========================================================
' Prozedur: cmd_NeuesMitglied_Click
' Erstellt ein neues Mitglied und übergibt Vorbelegungen an frm_Mitgliedsdaten
' ==========================================================
Private Sub cmd_NeuesMitglied_Click()

    Dim ws As Worksheet
    Dim Parzelle As String
    Dim r As Long
    Dim foundCount As Long
    Dim vorhandeneListe As String
    Dim ersteZeile As Long
    Dim uebernahme As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Worksheets("Mitgliederliste")
    
    ' 1) Parzellenauswahl (ersetzt frm_ParzelleAuswahl durch InputBox)
    Parzelle = InputBox("Bitte geben Sie die Parzellennummer für das neue Mitglied ein (z.B. 1, 12a, 35b):", "Parzellenauswahl")
    
    Parzelle = Trim(Parzelle)

    If Parzelle = "" Then
        ' Abbrechen gedrückt oder keine Eingabe -> Vorgang abbrechen
        Exit Sub
    End If
    
    ' 2) Mitglieder mit dieser Parzelle suchen (Spalte B)
    foundCount = 0
    vorhandeneListe = ""
    ersteZeile = 0
    
    ' ACHTUNG: Die Spalten-Indizes sind in den Konstanten in mod_Hilfsfunktionen definiert
    ' Hier werden hartkodierte Spalten (2, 5, 6 etc.) verwendet.
    ' Wir sollten diese im Zuge der nächsten Überarbeitung durch Konstanten ersetzen.
    
    For r = 6 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        If StrComp(Trim(ws.Cells(r, 2).Value), Parzelle, vbTextCompare) = 0 Then
            foundCount = foundCount + 1
            If vorhandeneListe <> "" Then vorhandeneListe = vorhandeneListe & vbCrLf
            vorhandeneListe = vorhandeneListe & " - " & Trim(ws.Cells(r, 6).Value) & " " & Trim(ws.Cells(r, 5).Value) ' Vorname Nachname
            If ersteZeile = 0 Then ersteZeile = r ' Zeile des ersten gefundenen Mitglieds merken (für Adressübernahme)
        End If
    Next r
    
    ' 3) Falls bereits Mitglieder auf dieser Parzelle: Liste anzeigen & Entscheidung erfragen
    If foundCount > 0 Then
        Dim Antwort As VbMsgBoxResult
        Antwort = MsgBox("Auf Parzelle " & Parzelle & " sind bereits " & foundCount & " Mitglied(er) eingetragen:" & vbCrLf & vbCrLf & _
                             vorhandeneListe & vbCrLf & vbCrLf & "Möchten Sie trotzdem ein weiteres Mitglied anlegen?", vbYesNo + vbQuestion, "Parzelle belegt")
        If Antwort = vbNo Then Exit Sub
        
        ' Frage: Adresse von vorhandenem übernehmen?
        uebernahme = MsgBox("Sollen die Adressdaten (Straße, Nr., PLZ, Ort, Telefon etc.) des ersten vorhandenen Mitglieds übernommen werden?", vbYesNo + vbQuestion, "Adresse übernehmen")
    End If
    
    ' 4) Jetzt vorbereitetes Formular öffnen und Felder vorbelegen:
    With frm_Mitgliedsdaten
        
        ' Ruft die Initialisierung auf (RowIndex = 0 bedeutet "NEU")
        ' Alle Vorbelegungen geschehen nun IM Init_MemberData,
        ' aber wir müssen die Adress-Übernahme von hier nach dort verschieben
        ' oder die Felder hier füllen, BEVOR wir Init_MemberData aufrufen, um den Modus zu setzen.
        ' Der sauberste Weg ist, die Übernahme-Logik HIER zu machen und dann das Formular zu zeigen.
        
        ' 5. Initialisierung: Setzt den Modus auf "NEU" und belegt Pachtbeginn
        ' Wir müssen hier Init_MemberData aufrufen. Da diese Prozedur keinen Rückgabewert hat,
        ' können wir sie nicht mit einer Zuweisung verwenden, sondern nur mit Call oder direktem Aufruf.
        ' ACHTUNG: Init_MemberData erwartet nur einen RowIndex.
        Call .Init_MemberData(0) ' RowIndex = 0 signalisiert Neuanlage
        
        ' 6. Daten, die HIER vorliegen, in die Textboxen schreiben (da sie jetzt sichtbar sind)
        .cbo_Parzelle.Value = Parzelle ' Parzelle immer setzen

        If foundCount > 0 And uebernahme = vbYes Then
            ' Adressübernahme (Spalten 7 bis 11)
            .txt_Strasse.Value = ws.Cells(ersteZeile, 7).Value
            .txt_Nummer.Value = ws.Cells(ersteZeile, 8).Value
            .txt_PLZ.Value = ws.Cells(ersteZeile, 9).Value
            .txt_Wohnort.Value = ws.Cells(ersteZeile, 10).Value
            .txt_Telefon.Value = ws.Cells(ersteZeile, 11).Value
            '.cbo_Seite.Value = ws.Cells(ersteZeile, 3).Value ' cbo_Seite entfernt
        Else
            ' Leeren der Felder, die in Init_MemberData(0) nicht geleert wurden
            .txt_Strasse.Value = ""
            .txt_Nummer.Value = ""
            .txt_PLZ.Value = ""
            .txt_Wohnort.Value = ""
            .txt_Telefon.Value = ""
        End If
        
        ' Mobil und E-Mail immer setzen/leeren
        .txt_Mobil.Value = ""
        .txt_Email.Value = ""
        
        ' Zeige Formular
        .Show
    End With

End Sub


' ==========================================================
' Prozedur: lst_Mitgliederliste_DblClick
' Öffnet die Detailansicht
' ==========================================================
Private Sub lst_Mitgliederliste_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call OeffneMitgliedsDetails
End Sub
    
' ***************************************************************
' NEUE PROZEDUR: Oeffnet die Details des ausgewaehlten Mitglieds
' ***************************************************************
Public Sub OeffneMitgliedsDetails()

    If Me.lst_Mitgliederliste.ListIndex < 0 Then Exit Sub
    
    Dim ws As Worksheet
    Dim lRow As Long
    
    Set ws = ThisWorkbook.Worksheets("Mitgliederliste")
    
    Dim ParzelleToFind As String
    Dim NachnameToFind As String
    
    ParzelleToFind = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 0)
    NachnameToFind = Me.lst_Mitgliederliste.List(Me.lst_Mitgliederliste.ListIndex, 4)
    lRow = 0
    
    ' Suche die korrekte Zeile: Suche nach Parzelle (Spalte B) und Nachname (Spalte E)
    Dim r As Long
    For r = 6 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        If StrComp(Trim(ws.Cells(r, 2).Value), ParzelleToFind, vbTextCompare) = 0 And _
           StrComp(Trim(ws.Cells(r, 5).Value), NachnameToFind, vbTextCompare) = 0 Then
             lRow = r
             Exit For ' Zeile gefunden
        End If
    Next r

    If lRow < 6 Then
        MsgBox "Fehler: Datenzeile in der Tabelle nicht gefunden.", vbCritical
        Exit Sub
    End If
    
    ' ***************************************************************
    ' KORREKTUR: Verwende eine explizite Variable, um die UserForm zu instanziieren.
    ' Dies umgeht mögliche Probleme mit der globalen Standardinstanz, die möglicherweise
    ' nicht korrekt initialisiert wird, bevor Init_MemberData aufgerufen wird.
    ' ***************************************************************
    Dim Frm As frm_Mitgliedsdaten
    Set Frm = New frm_Mitgliedsdaten

    With Frm
        
        ' HIER IST DIE WICHTIGE ÄNDERUNG:
        ' Statt alle Labels und Textboxen hier manuell zu befüllen,
        ' rufen wir nur Init_MemberData mit der gefundenen Zeilennummer auf.
        ' Init_MemberData übernimmt das Laden, Setzen der Labels/Textboxen
        ' und das Umschalten in den ANSICHT-Modus.
        
        Call .Init_MemberData(lRow)
        
        .Show
        
    End With
    
End Sub


' ==========================================================
' NEUE PRIVATE PROZEDUR: Befüllt die ListBox (UNVERÄNDERT)
' ==========================================================
Private Sub LoadListBoxData()
    ' (Der Code von LoadListBoxData bleibt unverändert)
    ' ...
    Dim iZeile As Long
    Dim AnzArr As Long
    Dim arr() As Variant

    lst_Mitgliederliste.ColumnCount = 14
    lst_Mitgliederliste.ColumnHeads = False
    
    AnzArr = 0
    
    With Worksheets("Mitgliederliste")
        For iZeile = 6 To .Cells(.Rows.Count, 2).End(xlUp).Row
            If Trim(.Cells(iZeile, 2).Value) <> "" And StrComp(Trim(.Cells(iZeile, 2).Value), "Verein", vbTextCompare) <> 0 Then
                AnzArr = AnzArr + 1
            End If
        Next iZeile
        
        If AnzArr > 0 Then
             ReDim arr(0 To AnzArr - 1, 0 To 13)
        Else
             lst_Mitgliederliste.Clear
             Exit Sub
        End If
        
        AnzArr = 0
        
        For iZeile = 6 To .Cells(.Rows.Count, 2).End(xlUp).Row
            If Trim(.Cells(iZeile, 2).Value) <> "" And StrComp(Trim(.Cells(iZeile, 2).Value), "Verein", vbTextCompare) <> 0 Then
                
                arr(AnzArr, 0) = .Cells(iZeile, 2).Value    ' B: Parzelle
                arr(AnzArr, 1) = .Cells(iZeile, 3).Value    ' C: Seite
                arr(AnzArr, 2) = .Cells(iZeile, 4).Value    ' D: Anrede
                arr(AnzArr, 3) = .Cells(iZeile, 6).Value    ' F: Vorname
                arr(AnzArr, 4) = .Cells(iZeile, 5).Value    ' E: Nachname
                arr(AnzArr, 5) = .Cells(iZeile, 7).Value    ' G: Strasse
                arr(AnzArr, 6) = .Cells(iZeile, 8).Value    ' H: Nummer
                arr(AnzArr, 7) = .Cells(iZeile, 9).Value    ' I: PLZ
                arr(AnzArr, 8) = .Cells(iZeile, 10).Value ' J: Wohnort
                arr(AnzArr, 9) = .Cells(iZeile, 11).Value ' K: Telefon
                arr(AnzArr, 10) = .Cells(iZeile, 12).Value ' L: Mobil
                arr(AnzArr, 11) = .Cells(iZeile, 13).Value ' M: Geburtstag
                arr(AnzArr, 12) = .Cells(iZeile, 14).Value ' N: Email
                arr(AnzArr, 13) = .Cells(iZeile, 15).Value ' O: Funktion
                
                AnzArr = AnzArr + 1
            End If
        Next iZeile
        
        lst_Mitgliederliste.List = arr
    End With
End Sub

' ==========================================================
' Prozedur: UserForm_Initialize (UNVERÄNDERT)
' ==========================================================
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 1
    With Me
        ' ACHTUNG: Hier werden hartkodierte Bereiche verwendet (D2).
        .lbl_ListDatum.Caption = "Mitgliederliste vom:  " & Worksheets("Mitgliederliste").Range("D2")
    End With
    Call LoadListBoxData
End Sub

' ==========================================================
' Prozedur: RefreshMitgliederListe (UNVERÄNDERT)
' ==========================================================
Public Sub RefreshMitgliederListe()
    Call LoadListBoxData
End Sub



