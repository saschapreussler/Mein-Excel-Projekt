Attribute VB_Name = "mod_Const"
Option Explicit

' ***************************************************************
' MODUL: mod_Const
' ZWECK: Zentrale Konstanten für das gesamte Projekt
' ***************************************************************

' ===============================================================
' A. ARBEITSBLATTNAMEN
' ===============================================================
Public Const WS_BANKKONTO As String = "Bankkonto"
Public Const WS_DATEN As String = "Daten"
Public Const WS_MITGLIEDER As String = "Mitgliederliste"
Public Const WS_UEBERSICHT As String = "Übersicht"
Public Const WS_MITGLIEDER_HISTORIE As String = "Mitgliederhistorie"

' ===============================================================
' B. NAMED RANGES
' ===============================================================
Public Const RANGE_KATEGORIE_REGELN As String = "rng_KategorieRegeln"

' ===============================================================
' C. DATEN – TEMPORÄRE HILFSSPALTEN (CRUD)
' ===============================================================
Public Const DATA_TEMP_COL_KEY As Long = 26       ' Z
Public Const DATA_TEMP_COL_NAME As Long = 27      ' AA
Public Const DATA_TEMP_COL_KONTONAME As Long = 28 ' AB
Public Const DATA_TEMP_COL_IBAN As Long = 29      ' AC

' ===============================================================
' D. MITGLIEDERLISTE – STRUKTUR (M_COL_HAUSNR/M_COL_NUMMER korrigiert und MemberID hinzugefügt)
' ===============================================================
Public Const M_HEADER_ROW As Long = 5
Public Const M_START_ROW As Long = 6

Public Const M_COL_MEMBER_ID As Long = 1   ' NEU: Spalte A (Eindeutiger Schlüssel des Mitglieds)
Public Const M_COL_PARZELLE As Long = 2    ' Spalte B
Public Const M_COL_SEITE As Long = 3       ' Spalte C
Public Const M_COL_ANREDE As Long = 4      ' Spalte D
Public Const M_COL_NACHNAME As Long = 5    ' Spalte E
Public Const M_COL_VORNAME As Long = 6     ' Spalte F
Public Const M_COL_STRASSE As Long = 7     ' Spalte G
Public Const M_COL_HAUSNR As Long = 8      ' Spalte H (Hausnummer)
Public Const M_COL_NUMMER As Long = 8      ' Spalte H (Alias für Hausnummer, zur Behebung des vorherigen Fehlers)
Public Const M_COL_PLZ As Long = 9         ' Spalte I
Public Const M_COL_WOHNORT As Long = 10    ' Spalte J
Public Const M_COL_TELEFON As Long = 11    ' Spalte K
Public Const M_COL_MOBIL As Long = 12      ' Spalte L
Public Const M_COL_GEBURTSTAG As Long = 13 ' Spalte M
Public Const M_COL_EMAIL As Long = 14      ' Spalte N
Public Const M_COL_FUNKTION As Long = 15   ' Spalte O

' --- NEU FÜR DEN DATENSTAND (Stand: 23.12.2025 in D2) ---
Public Const M_STAND_ROW As Long = 2       ' Zeile 2
Public Const M_STAND_COL As Long = 4       ' Spalte D (Zelle D2)

' --- NEU FÜR MITGLIEDSSTATUS / AUSTRETEN ---
Public Const M_COL_PACHTENDE As Long = 16  ' Spalte P (Neu: Datum des Pachtendes)
Public Const M_COL_ENTITY_KEY As Long = 17 ' Spalte Q (Neu: Eindeutiger Schlüssel des Mitglieds)
' ACHTUNG: Die Konstante M_COL_MAX muss mindestens so hoch sein wie die letzte genutzte Spalte.
Public Const M_COL_MAX As Long = 17

Public Const M_UPDATE_ROW As Long = M_STAND_ROW
Public Const M_UPDATE_COL As Long = M_STAND_COL

' --- HILFSKONSTANTEN FÜR DAS MAPPING ---
Public Const MEMBER_COL_PARZELLE As Long = M_COL_PARZELLE    ' = 2 (B)
Public Const MEMBER_COL_VORNAME As Long = M_COL_VORNAME      ' = 6 (F)
Public Const MEMBER_COL_NACHNAME As Long = M_COL_NACHNAME    ' = 5 (E)

' ===============================================================
' E. BANKKONTO
' ===============================================================
Public Const BK_START_ROW As Long = 28
Public Const BK_COL_DATUM As Long = 1
Public Const BK_COL_BETRAG As Long = 2
Public Const BK_COL_NAME As Long = 3
Public Const BK_COL_IBAN As Long = 4
Public Const BK_COL_VERWENDUNGSZWECK As Long = 5
Public Const BK_COL_BUCHUNGSTEXT As Long = 6
Public Const BK_COL_KATEGORIE As Long = 8
Public Const BK_COL_STATUS As Long = 11
Public Const BK_COL_ENTITY_KEY As Long = 22

' Neue Spalten für mod_BetragsZuordnung
Public Const BK_COL_INTERNE_NR As Long = 10      ' Spalte J
Public Const BK_COL_MITGL_BEITR As Long = 13     ' Spalte M
Public Const BK_COL_SONST_EINN As Long = 19      ' Spalte S
Public Const BK_COL_UNTERHALT As Long = 20       ' Spalte T
Public Const BK_COL_BEMERKUNG As Long = 12       ' Spalte L

' ===============================================================
' F. DATEN – ENTITY / MAPPING
' ===============================================================
Public Const DATA_START_ROW As Long = 4

' Primäre Mapping-Spalten
Public Const DATA_MAP_COL_ENTITYKEY As Long = 19     ' S
Public Const DATA_MAP_COL_IBAN_OLD As Long = 20      ' T
Public Const DATA_MAP_COL_KTONAME As Long = 21       ' U
Public Const DATA_MAP_COL_ZUORDNUNG As Long = 22     ' V
Public Const DATA_MAP_COL_PARZELLE As Long = 23      ' W
Public Const DATA_MAP_COL_ENTITYROLE As Long = 24    ' X
Public Const DATA_MAP_COL_DEBUG As Long = 25         ' Y
Public Const DATA_MAP_COL_LAST As Long = 25

' ===============================================================
' G. ABWÄRTSKOMPATIBLE ALIASE (NICHT ENTFERNEN!)
' ===============================================================
Public Const DATA_MAP_COL_PARZ_KEY As Long = DATA_MAP_COL_PARZELLE
Public Const DATA_MAP_COL_NAME As Long = DATA_MAP_COL_ZUORDNUNG
Public Const DATA_MAP_COL_KONTONAME As Long = DATA_MAP_COL_KTONAME
Public Const DATA_MAP_COL_IBAN As Long = DATA_MAP_COL_IBAN_OLD

' ===============================================================
' H. KATEGORIE-ENGINE
' ===============================================================
Public Const DATA_CAT_COL_START As Long = 10
Public Const DATA_CAT_COL_END As Long = 17
Public Const DATA_CAT_COL_KATEGORIE As Long = 10

' ===============================================================
' I. CSV-IMPORT (SPARKASSE)
' ===============================================================
Public Const CSV_COL_BUCHUNGSDATUM As Long = 2
Public Const CSV_COL_STATUS As Long = 4
Public Const CSV_COL_VERWENDUNGSZWECK As Long = 5
Public Const CSV_COL_NAME As Long = 12
Public Const CSV_COL_IBAN As Long = 13
Public Const CSV_COL_BETRAG As Long = 15

' ===============================================================
' J. ZÄHLERLOGIK
' ===============================================================
Public Const COL_PARZELLE As Long = 1
Public Const COL_STAND_ANFANG As Long = 2
Public Const COL_STAND_ENDE As Long = 3
Public Const COL_VERBRAUCH As Long = 4
Public Const COL_RECHNUNG_FORMEL As Long = 5
Public Const COL_BEMERKUNG As Long = 9

Public Const HIST_SHEET_NAME As String = "Zaehlerhistorie"
Public Const HIST_TABLE_NAME As String = "tbl_Historie"

Public Const STR_HISTORY_SEPARATOR As String = "--- MA-INFO ---"

' ===============================================================
' K. LISTBOX / PROTOKOLL
' ===============================================================
Public Const FORM_LISTBOX_NAME As String = "lst_ImportReport"
Public Const WS_PROTOCOL_TEMP As String = "Protokoll_Temp"
Public Const PROTOCOL_RANGE_START As String = "A1"
Public Const MAX_LISTBOX_LINES As Long = 60

Public Const CELL_IMPORT_PROTOKOLL As String = "Z100"

' ===============================================================
' L. ENTITY ROLE
' ===============================================================
Public Const ROLE_RANGE As String = "AE4:AE8"

' ===============================================================
' M. MITGLIEDERHISTORIE – STRUKTUR (NEU)
' ===============================================================
Public Const H_HEADER_ROW As Long = 3
Public Const H_START_ROW As Long = 4

Public Const H_COL_PARZELLE As Long = 1        ' A: Abgegebene Parzelle
Public Const H_COL_MITGL_ID As Long = 2        ' B: ID des ausscheidenden Mitglieds (EntityKey)
Public Const H_COL_NACHNAME As Long = 3        ' C: Nachname (Lesbarkeit)
Public Const H_COL_AUST_DATUM As Long = 4      ' D: Stichtag der Endabrechnung
Public Const H_COL_NEUER_PAECHTER_ID As Long = 5 ' E: ID des Nachpächters
Public Const H_COL_GRUND As Long = 6           ' F: Grund (Austritt/Wechsel)
Public Const H_COL_SYSTEMZEIT As Long = 7      ' G: Eintragungszeitpunkt

' ===============================================================
' N. SICHERHEIT & SONSTIGES
' ===============================================================
Public Const PASSWORD As String = ""
Public Const PARZELLE_VEREIN As String = "Verein"
Public Const ANREDE_KGA As String = "KGA"
Public Const AUSTRITT_STATUS As String = "Ehemaliges Mitglied"



