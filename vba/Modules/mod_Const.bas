Attribute VB_Name = "mod_Const"
Option Explicit

' ***************************************************************
' MODUL: mod_Const
' ZWECK: Zentrale Konstanten fuer das gesamte Projekt
' VERSION: 2.1 - 01.02.2026
' AENDERUNG: EntityKey-Tabelle R-X (18-24), EntityRole-DropDown auf AD (30)
' ***************************************************************

' ===============================================================
' A. ARBEITSBLATTNAMEN
' ===============================================================
Public Const WS_BANKKONTO As String = "Bankkonto"
Public Const WS_DATEN As String = "Daten"
Public Const WS_MITGLIEDER As String = "Mitgliederliste"
Public Const WS_UEBERSICHT As String = "Uebersicht"
Public Const WS_MITGLIEDER_HISTORIE As String = "Mitgliederhistorie"
Public Const WS_EINSTELLUNGEN As String = "Einstellungen"
Public Const WS_VEREINSKASSE As String = "Vereinskasse"

' ===============================================================
' B. NAMED RANGES
' ===============================================================
Public Const RANGE_KATEGORIE_REGELN As String = "rng_KategorieRegeln"

' ===============================================================
' C. DATEN - TEMPORAERE HILFSSPALTEN (CRUD)
' ===============================================================
Public Const DATA_TEMP_COL_KEY As Long = 26
Public Const DATA_TEMP_COL_NAME As Long = 27
Public Const DATA_TEMP_COL_KONTONAME As Long = 28
Public Const DATA_TEMP_COL_IBAN As Long = 29

' ===============================================================
' D. MITGLIEDERLISTE - STRUKTUR
' ===============================================================
Public Const M_HEADER_ROW As Long = 5
Public Const M_START_ROW As Long = 6

Public Const M_COL_MEMBER_ID As Long = 1
Public Const M_COL_PARZELLE As Long = 2
Public Const M_COL_SEITE As Long = 3
Public Const M_COL_ANREDE As Long = 4
Public Const M_COL_NACHNAME As Long = 5
Public Const M_COL_VORNAME As Long = 6
Public Const M_COL_STRASSE As Long = 7
Public Const M_COL_HAUSNR As Long = 8
Public Const M_COL_NUMMER As Long = 8
Public Const M_COL_PLZ As Long = 9
Public Const M_COL_WOHNORT As Long = 10
Public Const M_COL_TELEFON As Long = 11
Public Const M_COL_MOBIL As Long = 12
Public Const M_COL_GEBURTSTAG As Long = 13
Public Const M_COL_EMAIL As Long = 14
Public Const M_COL_FUNKTION As Long = 15
Public Const M_COL_PACHTANFANG As Long = 16
Public Const M_COL_PACHTENDE As Long = 17
Public Const M_COL_ENTITY_KEY As Long = 18
Public Const M_COL_MAX As Long = 18

Public Const M_STAND_ROW As Long = 2
Public Const M_STAND_COL As Long = 4
Public Const M_UPDATE_ROW As Long = M_STAND_ROW
Public Const M_UPDATE_COL As Long = M_STAND_COL

Public Const MEMBER_COL_PARZELLE As Long = M_COL_PARZELLE
Public Const MEMBER_COL_VORNAME As Long = M_COL_VORNAME
Public Const MEMBER_COL_NACHNAME As Long = M_COL_NACHNAME

' ===============================================================
' E. BANKKONTO - SPALTENSTRUKTUR
' ===============================================================
Public Const BK_START_ROW As Long = 28
Public Const BK_HEADER_ROW As Long = 27

' Spalte A-G: Import-Daten
Public Const BK_COL_DATUM As Long = 1
Public Const BK_COL_BETRAG As Long = 2
Public Const BK_COL_NAME As Long = 3
Public Const BK_COL_IBAN As Long = 4
Public Const BK_COL_VERWENDUNGSZWECK As Long = 5
Public Const BK_COL_BUCHUNGSTEXT As Long = 6
Public Const BK_COL_IM_AUSWERTUNGSMONAT As Long = 7

' Spalte H-L: Kategorisierung und Verwaltung
Public Const BK_COL_KATEGORIE As Long = 8
Public Const BK_COL_MONAT_PERIODE As Long = 9
Public Const BK_COL_INTERNE_NR As Long = 10
Public Const BK_COL_STATUS As Long = 11
Public Const BK_COL_BEMERKUNG As Long = 12

' Spalte M-S: EINNAHMEN
Public Const BK_COL_MITGL_BEITR As Long = 13
Public Const BK_COL_SPENDEN As Long = 14
Public Const BK_COL_ZUSCHUESSE As Long = 15
Public Const BK_COL_VERWALTUNG_E As Long = 16
Public Const BK_COL_VERMOEGEN As Long = 17
Public Const BK_COL_VERANSTALT_E As Long = 18
Public Const BK_COL_SONST_EINN As Long = 19

' Spalte T-Z: AUSGABEN
Public Const BK_COL_UNTERHALT As Long = 20
Public Const BK_COL_FORTBILDUNG As Long = 21
Public Const BK_COL_VERANSTALT_A As Long = 22
Public Const BK_COL_BUEROBETRIEB As Long = 23
Public Const BK_COL_AUFWANDSENTSCH As Long = 24
Public Const BK_COL_SONST_AUSG As Long = 25
Public Const BK_COL_AUSZAHL_KASSE As Long = 26

' Bereichsgrenzen
Public Const BK_COL_EINNAHMEN_START As Long = 13
Public Const BK_COL_EINNAHMEN_ENDE As Long = 19
Public Const BK_COL_AUSGABEN_START As Long = 20
Public Const BK_COL_AUSGABEN_ENDE As Long = 26

' Legacy-Alias
Public Const BK_COL_ENTITY_KEY As Long = 22

' ===============================================================
' F. DATEN - ENTITY / MAPPING (Spalten R-X nach Loeschung von O!)
' ===============================================================
' WICHTIG: Nach Loeschung von Spalte O sind alle EntityKey-Spalten
' um 1 nach links verschoben! R-X statt S-Y (18-24 statt 19-25)
' ===============================================================
Public Const DATA_START_ROW As Long = 4

' KORRIGIERT: Spalten 18-24 (R-X) statt 19-25 (S-Y)
Public Const DATA_MAP_COL_ENTITYKEY As Long = 18    ' R - EntityKey (GUID)
Public Const DATA_MAP_COL_IBAN_OLD As Long = 19     ' S - IBAN
Public Const DATA_MAP_COL_KTONAME As Long = 20      ' T - Zahler/Empfaenger
Public Const DATA_MAP_COL_ZUORDNUNG As Long = 21    ' U - Mitglied(er)/Zuordnung
Public Const DATA_MAP_COL_PARZELLE As Long = 22     ' V - Parzelle(n)
Public Const DATA_MAP_COL_ENTITYROLE As Long = 23   ' W - EntityRole
Public Const DATA_MAP_COL_DEBUG As Long = 24        ' X - Debug Zuordnung
Public Const DATA_MAP_COL_LAST As Long = 24

' Aliase fuer Kompatibilitaet
Public Const DATA_MAP_COL_PARZ_KEY As Long = DATA_MAP_COL_PARZELLE
Public Const DATA_MAP_COL_NAME As Long = DATA_MAP_COL_ZUORDNUNG
Public Const DATA_MAP_COL_KONTONAME As Long = DATA_MAP_COL_KTONAME
Public Const DATA_MAP_COL_IBAN As Long = DATA_MAP_COL_IBAN_OLD

' Aliase fuer EntityKey (EK_) - ZENTRAL hier definiert!
Public Const EK_START_ROW As Long = DATA_START_ROW
Public Const EK_HEADER_ROW As Long = 3
Public Const EK_COL_ENTITYKEY As Long = DATA_MAP_COL_ENTITYKEY   ' 18 = R
Public Const EK_COL_IBAN As Long = DATA_MAP_COL_IBAN_OLD         ' 19 = S
Public Const EK_COL_KONTONAME As Long = DATA_MAP_COL_KTONAME     ' 20 = T
Public Const EK_COL_ZUORDNUNG As Long = DATA_MAP_COL_ZUORDNUNG   ' 21 = U
Public Const EK_COL_PARZELLE As Long = DATA_MAP_COL_PARZELLE     ' 22 = V
Public Const EK_COL_ROLE As Long = DATA_MAP_COL_ENTITYROLE       ' 23 = W
Public Const EK_COL_DEBUG As Long = DATA_MAP_COL_DEBUG           ' 24 = X

' ===============================================================
' G. KATEGORIE-TABELLE (Daten! Spalte J-P)
' ===============================================================
' Struktur: J=Kategorie, K=E/A, L=Keyword, M=Prioritaet, N=Zielspalte, O=Faelligkeit, P=Kommentar

Public Const DATA_CAT_COL_START As Long = 10
Public Const DATA_CAT_COL_KATEGORIE As Long = 10    ' J
Public Const DATA_CAT_COL_EINAUS As Long = 11       ' K
Public Const DATA_CAT_COL_KEYWORD As Long = 12      ' L
Public Const DATA_CAT_COL_PRIORITAET As Long = 13   ' M
Public Const DATA_CAT_COL_ZIELSPALTE As Long = 14   ' N
Public Const DATA_CAT_COL_FAELLIGKEIT As Long = 15  ' O
Public Const DATA_CAT_COL_KOMMENTAR As Long = 16    ' P
Public Const DATA_CAT_COL_END As Long = 16

' ===============================================================
' H. HILFSSPALTEN AUF DATEN! (DropDown-Listen etc.)
' ===============================================================
Public Const DATA_COL_IMPORT_STATUS As Long = 26    ' Z - Import-Protokoll
Public Const DATA_COL_DD_KATEGORIE_START As Long = 27  ' AA - DropDown-Liste Kategorien (alt)
Public Const DATA_COL_DD_KATEGORIE_END As Long = 31    ' AE - DropDown-Liste Ende (alt)
Public Const DATA_COL_DD_ROLE As Long = 30          ' AD - EntityRole DropDown (NEU!)
Public Const DATA_COL_EINNAHMEN As Long = 32        ' AF - Einnahmen-Kategorien
Public Const DATA_COL_AUSGABEN As Long = 33         ' AG - Ausgaben-Kategorien
Public Const DATA_COL_MONAT_PERIODE As Long = 34    ' AH - Monat/Periode Liste

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
' J. ZAEHLERLOGIK
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
' L. ENTITY ROLE - KORRIGIERT: Jetzt Spalte AD!
' ===============================================================
Public Const ROLE_RANGE As String = "AD4:AD10"

' ===============================================================
' M. MITGLIEDERHISTORIE - STRUKTUR
' ===============================================================
Public Const H_HEADER_ROW As Long = 3
Public Const H_START_ROW As Long = 4

Public Const H_COL_PARZELLE As Long = 1
Public Const H_COL_MITGL_ID As Long = 2
Public Const H_COL_MEMBER_ID_ALT As Long = 2
Public Const H_COL_NAME_EHEM_PAECHTER As Long = 3
Public Const H_COL_NACHNAME As Long = 3
Public Const H_COL_AUST_DATUM As Long = 4
Public Const H_COL_GRUND As Long = 5
Public Const H_COL_NAME_NEUER_PAECHTER As Long = 6
Public Const H_COL_NACHPAECHTER_NAME As Long = 6
Public Const H_COL_NEUER_PAECHTER_ID As Long = 7
Public Const H_COL_NACHPAECHTER_ID As Long = 7
Public Const H_COL_KOMMENTAR As Long = 8
Public Const H_COL_ENDABRECHNUNG As Long = 9
Public Const H_COL_SYSTEMZEIT As Long = 10

' ===============================================================
' N. SICHERHEIT & SONSTIGES
' ===============================================================
Public Const PASSWORD As String = ""
Public Const PARZELLE_VEREIN As String = "Verein"
Public Const ANREDE_KGA As String = "KGA"
Public Const AUSTRITT_STATUS As String = "Ehemaliges Mitglied"

' ===============================================================
' O. ERLAUBTE FUNKTIONEN FUER PARZELLENPACHT
' ===============================================================
Public Const FUNKTION_MITGLIED_MIT_PACHT As String = "Mitglied mit Pacht"
Public Const FUNKTION_1_VORSITZENDER As String = "1. Vorsitzende(r)"
Public Const FUNKTION_2_VORSITZENDER As String = "2. Vorsitzende(r)"
Public Const FUNKTION_KASSIERER As String = "Kassierer"
Public Const FUNKTION_SCHRIFTFUEHRER As String = "Schriftfuehrer"

Public Function GetErlaubteFunktionenFuerParzelle() As Variant
    GetErlaubteFunktionenFuerParzelle = Array( _
        FUNKTION_MITGLIED_MIT_PACHT, _
        FUNKTION_1_VORSITZENDER, _
        FUNKTION_2_VORSITZENDER, _
        FUNKTION_KASSIERER, _
        FUNKTION_SCHRIFTFUEHRER _
    )
End Function



