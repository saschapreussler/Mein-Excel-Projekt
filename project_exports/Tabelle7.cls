VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo Fehlerbehandlung

    ' ***************************************************************
    ' KERN-KONSTANTEN
    ' ***************************************************************
    ' HINWEIS: M_START_ROW MUSS definiert sein. Standard ist 6.
    ' Falls abweichend, muss dieser Wert angepasst werden.
    Const M_START_ROW As Long = 6      ' ANNAHME: Erste Datenzeile
    
    Const M_COL_PARZELLE As Long = 2   ' Spalte B
    Const M_COL_SEITE As Long = 3      ' Spalte C
    Const M_COL_ANREDE As Long = 4     ' Spalte D
    Const M_COL_FUNKTION As Long = 15  ' Spalte O
    Const M_COL_AUSTRITT As Long = 16  ' Spalte P
    
    Const RELEVANT_COLUMNS As String = "B:B,D:D,O:O"
    Const AUSTRITT_STATUS As String = "Ehemaliges Mitglied"
    Const PARZELLE_VEREIN As String = "Verein"
    Const ANREDE_KGA As String = "KGA"

    ' ------------------------------------------------------------
    ' PRÜFUNG & INITIALISIERUNG
    ' ------------------------------------------------------------
    If Intersect(Target, Me.Range(RELEVANT_COLUMNS)) Is Nothing Then Exit Sub
    If Target.Row < M_START_ROW Then Exit Sub

    Application.EnableEvents = False
    
    ' Blattschutz aufheben (mit leerem Passwort)
    Me.Unprotect PASSWORD:=""

    ' ------------------------------------------------------------
    ' 1. Automatische Logik (Parzelle geändert: B -> C / B -> D)
    ' ------------------------------------------------------------
    If Not Intersect(Target, Me.Columns(M_COL_PARZELLE)) Is Nothing Then

        Dim rCell As Range
        Dim parzelle As String
        Dim seite As String
        Dim nr As Long

        For Each rCell In Intersect(Target, Me.Columns(M_COL_PARZELLE)).Cells

            If rCell.Row >= M_START_ROW Then

                parzelle = Trim(CStr(rCell.Value))
                seite = ""
                
                ' NEUE LOGIK: Wenn Parzelle "Verein", setze Anrede auf "KGA" (Spalte D)
                ' Verwendung von StrComp für robuste, fallunabhängige Prüfung
                If StrComp(parzelle, PARZELLE_VEREIN, vbTextCompare) = 0 Then
                    Me.Cells(rCell.Row, M_COL_ANREDE).Value = ANREDE_KGA
                End If

                ' Logik zur Füllung Spalte C (Seite)
                If IsNumeric(parzelle) Then
                    nr = CLng(parzelle)

                    If nr >= 1 And nr <= 9 Then
                        seite = "rechts"
                    ElseIf nr >= 10 And nr <= 14 Then
                        seite = "links"
                    End If

                ' Verwendung von StrComp für robuste, fallunabhängige Prüfung
                ElseIf StrComp(parzelle, PARZELLE_VEREIN, vbTextCompare) = 0 Then
                    seite = "mitte"
                End If

                Me.Cells(rCell.Row, M_COL_SEITE).Value = seite
            End If
        Next rCell
    End If
    
    ' ------------------------------------------------------------
    ' 2. Logik: Austrittsdatum setzen/löschen (Spalte P)
    ' ------------------------------------------------------------
    If Not Intersect(Target, Me.Columns(M_COL_FUNKTION)) Is Nothing Then
        
        Dim cCell As Range
        
        For Each cCell In Intersect(Target, Me.Columns(M_COL_FUNKTION)).Cells
            
            If cCell.Row >= M_START_ROW Then
            
                Dim AustrittsZelle As Range
                Set AustrittsZelle = Me.Cells(cCell.Row, M_COL_AUSTRITT)
                
                If StrComp(Trim(CStr(cCell.Value)), AUSTRITT_STATUS, vbTextCompare) = 0 Then
                    If IsEmpty(AustrittsZelle.Value) Then
                        AustrittsZelle.Value = Date
                        AustrittsZelle.NumberFormat = "DD.MM.YYYY"
                    End If
                Else
                    If Not IsEmpty(AustrittsZelle.Value) Then
                        AustrittsZelle.ClearContents
                    End If
                End If
            End If
        Next cCell
    End If


    ' ------------------------------------------------------------
    ' 3. Sortierung (AKTIVIERT)
    ' ------------------------------------------------------------
    ' Dies sollte nur ausgeführt werden, wenn sich die Parzelle geändert hat und
    ' das Modul mod_Mitglieder_UI die Unprotect/Protect Logik enthält.
    If Not Intersect(Target, Me.Columns(M_COL_PARZELLE)) Is Nothing Then
        Call mod_Mitglieder_UI.Sortiere_Mitgliederliste_Nach_Parzelle
    End If
    

    ' ------------------------------------------------------------
    ' 4. Datenstand aktualisieren (AKTIVIERT)
    ' ------------------------------------------------------------
    Call mod_Mitglieder_UI.AktualisiereDatenstand

    ' ------------------------------------------------------------
    ' 5. Bestehende Logiken (DEAKTIVIERT - NUR FÜR STABILITÄT)
    ' ------------------------------------------------------------
    ' Diese Aufrufe bleiben auskommentiert, bis die Sortierung funktioniert
    ' und die Laufzeitfehler 424/1004 dauerhaft behoben sind.
    ' Call mod_Datenpflege.Aktualisiere_Parzellen_Mapping_Final
    ' Call mod_Datenpflege.Sortiere_Tabellen_Daten
    ' Call mod_ZaehlerLogik.Ermittle_Kennzahlen_Mitgliederliste
    ' Call mod_ZaehlerLogik.ErzeugeParzellenUebersicht
    ' Call mod_ZaehlerLogik.AktualisiereZaehlerTabellenSpalteA

Fertig:
    Me.Protect PASSWORD:="", UserInterfaceOnly:=True
    Application.EnableEvents = True
    Exit Sub

Fehlerbehandlung:
    ' Sicherstellen, dass der Blattschutz wiederhergestellt und Events aktiviert werden
    Me.Protect PASSWORD:="", UserInterfaceOnly:=True
    Application.EnableEvents = True

    MsgBox "Fehler bei Änderung der Mitgliederliste." & vbCrLf & _
           "Fehler-Nr: " & Err.Number & vbCrLf & _
           "Beschreibung: " & Err.Description, vbCritical
End Sub

