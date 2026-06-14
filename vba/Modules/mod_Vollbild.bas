Attribute VB_Name = "mod_Vollbild"
Option Explicit

' ===============================================================
' Modul: mod_Vollbild
' v8.0: Vollbildmodus fuer Startmenue + Dashboard
'       - kein Ribbon (Menueband)
'       - keine Zeilen-/Spaltenkoepfe
'       - keine Bearbeitungsleiste (FormulaBar)
'       - keine Statusleiste
'       - keine Blattregister, keine Scrollleisten
'       Beim Verlassen wird alles wieder eingeblendet.
'       Application.DisplayFullScreen schaltet den ribbonlosen
'       Vollbildmodus von Excel automatisch ein/aus.
' ===============================================================

Private m_VollbildAktiv As Boolean

Public Sub SetzeVollbildModus(ByVal aktiv As Boolean)
    On Error Resume Next
    
    ' Idempotent: wenn schon im gewuenschten Zustand -> nichts tun
    If aktiv = m_VollbildAktiv Then
        ' Trotzdem ActiveWindow-Headings korrekt setzen, falls jemand sie geaendert hat
        If Not ActiveWindow Is Nothing Then
            ActiveWindow.DisplayHeadings = Not aktiv
        End If
        Exit Sub
    End If
    
    If aktiv Then
        ' --- Vollbild aktivieren ---
        Application.DisplayFullScreen = True
        Application.DisplayFormulaBar = False
        Application.DisplayStatusBar = False
        
        If Not ActiveWindow Is Nothing Then
            ActiveWindow.DisplayHeadings = False
            ActiveWindow.DisplayWorkbookTabs = False
            ActiveWindow.DisplayHorizontalScrollBar = False
            ActiveWindow.DisplayVerticalScrollBar = False
            ' Fenster maximieren (an aktuelle Aufloesung anpassen)
            ActiveWindow.WindowState = xlMaximized
        End If
    Else
        ' --- Vollbild verlassen ---
        Application.DisplayFullScreen = False
        Application.DisplayFormulaBar = True
        Application.DisplayStatusBar = True
        
        If Not ActiveWindow Is Nothing Then
            ActiveWindow.DisplayHeadings = True
            ActiveWindow.DisplayWorkbookTabs = True
            ActiveWindow.DisplayHorizontalScrollBar = True
            ActiveWindow.DisplayVerticalScrollBar = True
        End If
    End If
    
    m_VollbildAktiv = aktiv
End Sub

' Hilfsfunktion fuer Reset (z.B. aus Workbook_BeforeClose)
Public Sub ResetVollbildState()
    m_VollbildAktiv = False
End Sub
















