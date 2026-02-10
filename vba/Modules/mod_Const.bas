Attribute VB_Name = "mod_Const"
Option Explicit

' ***************************************************************
' MODUL: mod_Const
' ZWECK: Zentrale Konstanten für das gesamte Projekt
' VERSION: 2.7 - 09.02.2026
' ÄNDERUNG: ES_COL_SOLL_MONATE hinzugefügt (Spalte E),
'           alle folgenden ES-Spalten +1 verschoben,
'           ES_COL_END von 8 auf 9 geändert
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
' C. DATEN - TEMPORÄRE HILFSSPALTEN
' ===============================================================
Public Const DATA_TEMP_COL_KEY As Long = 25
Public Const DATA_TEMP_COL_NAME As Long = 26
Public Const DATA_TEMP_COL_KONTONAME As Long = 27
Public Const DATA_TEMP_COL_IBAN As Long = 28

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
' F. DATEN - KATEGORIE-TABELLE (Spalten J-P)
' ===============================================================
Public Const DATA_START_ROW As Long = 4
Public Const DATA_HEADER_ROW As Long = 3

Public Const DATA_CAT_COL_START As Long = 10
Public Const DATA_CAT_COL_KATEGORIE As Long = 10
Public Const DATA_CAT_COL_EINAUS As Long = 11
Public Const DATA_CAT_COL_KEYWORD As Long = 12
Public Const DATA_CAT_COL_PRIORITAET As Long = 13
Public Const DATA_CAT_COL_ZIELSPALTE As Long = 14
Public Const DATA_CAT_COL_FAELLIGKEIT As Long = 15
Public Const DATA_CAT_COL_KOMMENTAR As Long = 16
Public Const DATA_CAT_COL_END As Long = 16

' ===============================================================
' G. DATEN - ENTITYKEY-TABELLE (Spalten R-X)
' ===============================================================
Public Const DATA_MAP_COL_ENTITYKEY As Long = 18
Public Const DATA_MAP_COL_IBAN As Long = 19
Public Const DATA_MAP_COL_KTONAME As Long = 20
Public Const DATA_MAP_COL_ZUORDNUNG As Long = 21
Public Const DATA_MAP_COL_PARZELLE As Long = 22
Public Const DATA_MAP_COL_ENTITYROLE As Long = 23
Public Const DATA_MAP_COL_DEBUG As Long = 24
Public Const DATA_MAP_COL_LAST As Long = 24

' Aliase für Kompatibilität
Public Const DATA_MAP_COL_IBAN_OLD As Long = 19
Public Const DATA_MAP_COL_PARZ_KEY As Long = 22
Public Const DATA_MAP_COL_NAME As Long = 21
Public Const DATA_MAP_COL_KONTONAME As Long = 20

' ===============================================================
' G2. ENTITYKEY-TABELLE - ALIASE FÜR mod_Formatierung
' ===============================================================
Public Const EK_START_ROW As Long = 4
Public Const EK_COL_ENTITYKEY As Long = 18
Public Const EK_COL_IBAN As Long = 19
Public Const EK_COL_KONTONAME As Long = 20
Public Const EK_COL_ZUORDNUNG As Long = 21
Public Const EK_COL_PARZELLE As Long = 22
Public Const EK_COL_ROLE As Long = 23
Public Const EK_COL_DEBUG As Long = 24

' ===============================================================
' H. DATEN - DROPDOWN-FÜLLBEREICHE (Spalten Y-AH)
' ===============================================================
Public Const DATA_COL_IMPORT_STATUS As Long = 25
Public Const CELL_IMPORT_PROTOKOLL As String = "Y500"

Public Const DATA_COL_DD_EINAUS As Long = 26
Public Const DATA_COL_DD_PRIORITAET As Long = 27
Public Const DATA_COL_DD_JANEIN As Long = 28
Public Const DATA_COL_DD_FAELLIGKEIT As Long = 29
Public Const DATA_COL_DD_ENTITYROLE As Long = 30

Public Const DATA_COL_HILFSZELLE_FILTER As Long = 31

Public Const DATA_COL_KAT_EINNAHMEN As Long = 32
Public Const DATA_COL_KAT_AUSGABEN As Long = 33
Public Const DATA_COL_MONAT_PERIODE As Long = 34

' Legacy-Aliase
Public Const DATA_COL_EINNAHMEN As Long = 32
Public Const DATA_COL_AUSGABEN As Long = 33
Public Const DATA_COL_DD_ROLE As Long = 30
Public Const DATA_COL_DD_PARZELLE As Long = 6

' Hilfsspalte BA auf Daten für Einstellungen-DropDown Fallback
Public Const DATA_COL_ES_HILF As Long = 53      ' Spalte BA

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
Public Const MAX_LISTBOX_LINES As Long = 500

' ===============================================================
' L. ENTITY ROLE
' ===============================================================
Public Const ROLE_RANGE As String = "AD4:AD8"

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
' O. ERLAUBTE FUNKTIONEN FÜR PARZELLENPACHT
' ===============================================================
Public Const FUNKTION_MITGLIED_MIT_PACHT As String = "Mitglied mit Pacht"
Public Const FUNKTION_1_VORSITZENDER As String = "1. Vorsitzende(r)"
Public Const FUNKTION_2_VORSITZENDER As String = "2. Vorsitzende(r)"
Public Const FUNKTION_KASSIERER As String = "Kassierer"
Public Const FUNKTION_SCHRIFTFUEHRER As String = "Schriftfuehrer"

' ===============================================================
' P. EINSTELLUNGEN - ZAHLUNGSTERMINE (Spalten B-I)
'    v2.7: Neue Spalte E = Soll-Monat(e), alle folgenden +1
' ===============================================================
Public Const ES_HEADER_ROW As Long = 3
Public Const ES_START_ROW As Long = 4

Public Const ES_COL_KATEGORIE As Long = 2       ' Spalte B - Referenz Kategorie
Public Const ES_COL_SOLL_BETRAG As Long = 3     ' Spalte C - Soll-Betrag
Public Const ES_COL_SOLL_TAG As Long = 4        ' Spalte D - Soll-Tag (1-31)
Public Const ES_COL_SOLL_MONATE As Long = 5     ' Spalte E - Soll-Monat(e) z.B. "03, 06, 09"
Public Const ES_COL_STICHTAG_FIX As Long = 6    ' Spalte F - Soll-Stichtag (Fix) TT.MM.
Public Const ES_COL_VORLAUF As Long = 7         ' Spalte G - Vorlauf-Toleranz (Tage)
Public Const ES_COL_NACHLAUF As Long = 8        ' Spalte H - Nachlauf-Toleranz (Tage)
Public Const ES_COL_SAEUMNIS As Long = 9        ' Spalte I - Säumnis-Gebühr

Public Const ES_COL_START As Long = 2           ' Erste Datenspalte (B)
Public Const ES_COL_END As Long = 9             ' Letzte Datenspalte (I)


Public Function GetErlaubteFunktionenFuerParzelle() As Variant
    GetErlaubteFunktionenFuerParzelle = Array( _
        FUNKTION_MITGLIED_MIT_PACHT, _
        FUNKTION_1_VORSITZENDER, _
        FUNKTION_2_VORSITZENDER, _
        FUNKTION_KASSIERER, _
        FUNKTION_SCHRIFTFUEHRER _
    )
End Function


