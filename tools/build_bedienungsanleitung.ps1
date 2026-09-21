<#
.SYNOPSIS
    Erzeugt doku/Bedienungsanleitung.docx aus doku/Bedienungsanleitung.md.

.DESCRIPTION
    Die Textfassung der Anleitung liegt als Markdown im Projekt, weil sich
    Aenderungen daran sauber versionieren und vergleichen lassen. Dieses
    Skript macht daraus ein Word-Dokument mit Titelblatt, Inhaltsverzeichnis,
    Kopf- und Fusszeile.

    Unterstuetzte Auszeichnungen:
      #, ##, ###     Ueberschriften
      ---            Seitenumbruch
      ![alt](pfad)   Bild
      | a | b |      Tabelle, zweite Zeile ist die Trennzeile
      -              Aufzaehlung
      1.             nummerierte Liste
      >              Hinweiskasten
      **fett**       fetter Text
      `code`         nichtproportionale Schrift

    Word muss installiert sein. Das Skript aendert die Arbeitsmappe nicht.

.NOTES
    Nur auf ausdruecklichen Auftrag ausfuehren.
#>

[CmdletBinding()]
param(
    [string] $Quelle,
    [string] $Ziel,
    # Zusaetzlich eine PDF-Fassung erzeugen. Praktisch zum Weitergeben und
    # zum schnellen Sichtpruefen des Umbruchs.
    [switch] $AlsPdf
)

$ErrorActionPreference = 'Stop'

$skriptOrdner = Split-Path -Parent $MyInvocation.MyCommand.Definition
$wurzel       = Split-Path -Parent $skriptOrdner
if (-not $Quelle) { $Quelle = Join-Path $wurzel 'doku\Bedienungsanleitung.md' }
if (-not $Ziel)   { $Ziel   = Join-Path $wurzel 'doku\Bedienungsanleitung.docx' }

$Quelle = (Resolve-Path $Quelle).Path
$dokuOrdner = Split-Path -Parent $Quelle

Write-Host "Quelle : $Quelle"
Write-Host "Ziel   : $Ziel"

# Word-Konstanten, sprachunabhaengig ueber die eingebauten Indizes.
$wdStyleNormal     = -1
$wdStyleHeading1   = -2
$wdStyleHeading2   = -3
$wdStyleHeading3   = -4
$wdStyleListBullet = -5
$wdStyleListNumber = -6
$wdStyleTitle      = -63
$wdStoryEnd        = 6
$wdPageBreak       = 7
$wdAlignCenter     = 1
$wdFieldPage       = 33
$wdSeekHeader      = 9
$wdSeekFooter      = 10
$wdSeekMainDoc     = 0

$zeilen = [System.IO.File]::ReadAllLines($Quelle, [System.Text.Encoding]::UTF8)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Add()
    $sel = $word.Selection

    # ---- Seitenraender ----
    $doc.PageSetup.TopMargin    = $word.CentimetersToPoints(2.2)
    $doc.PageSetup.BottomMargin = $word.CentimetersToPoints(2.0)
    $doc.PageSetup.LeftMargin   = $word.CentimetersToPoints(2.4)
    $doc.PageSetup.RightMargin  = $word.CentimetersToPoints(2.0)

    $nutzbareBreite = $doc.PageSetup.PageWidth - $doc.PageSetup.LeftMargin - $doc.PageSetup.RightMargin

    # ---- Hilfsfunktionen ----

    # Schreibt eine Zeile und wertet dabei **fett** und `code` aus.
    function Write-Inline {
        param([string] $text)

        $teile = [regex]::Split($text, '(\*\*.+?\*\*|`.+?`)')
        foreach ($teil in $teile) {
            if ($teil -eq '') { continue }
            if ($teil -like '**?*' -and $teil.StartsWith('**') -and $teil.EndsWith('**')) {
                $sel.Font.Bold = $true
                $sel.TypeText($teil.Substring(2, $teil.Length - 4))
                $sel.Font.Bold = $false
            }
            elseif ($teil.StartsWith('`') -and $teil.EndsWith('`') -and $teil.Length -gt 2) {
                $altName = $sel.Font.Name
                $sel.Font.Name = 'Consolas'
                $sel.TypeText($teil.Substring(1, $teil.Length - 2))
                $sel.Font.Name = $altName
            }
            else {
                $sel.TypeText($teil)
            }
        }
    }

    function Set-Absatzstil {
        param([int] $stil)
        $sel.EndKey($wdStoryEnd) | Out-Null
        $sel.Style = $doc.Styles.Item($stil)
    }

    # ---- Titelblatt ----
    $sel.Style = $doc.Styles.Item($wdStyleTitle)
    $sel.ParagraphFormat.Alignment = $wdAlignCenter
    $sel.TypeText('Kassenbuch der Kleingartenanlage')
    $sel.TypeParagraph()

    $sel.Style = $doc.Styles.Item($wdStyleNormal)
    $sel.ParagraphFormat.Alignment = $wdAlignCenter
    $sel.Font.Size = 16
    $sel.TypeText('Bedienungs- und Betriebsanleitung')
    $sel.TypeParagraph()
    $sel.Font.Size = 11
    $sel.TypeText('Programm Kassenbuch 2018, Fassung v2.7.8')
    $sel.TypeParagraph()
    $sel.TypeText('Stand: ' + (Get-Date -Format 'dd.MM.yyyy'))
    $sel.TypeParagraph()
    $sel.ParagraphFormat.Alignment = 0
    $sel.InsertBreak($wdPageBreak)

    # ---- Inhaltsverzeichnis ----
    # Bewusst ohne Ueberschriftenformat, sonst taucht "Inhalt" als erster
    # Eintrag im eigenen Verzeichnis auf.
    $sel.Style = $doc.Styles.Item($wdStyleNormal)
    $sel.Font.Size = 18
    $sel.Font.Bold = $true
    $sel.TypeText('Inhalt')
    $sel.TypeParagraph()
    $sel.Font.Bold = $false
    $sel.Font.Size = 11

    $tocBereich = $sel.Range
    $toc = $doc.TablesOfContents.Add($tocBereich, $true, 1, 3)
    $sel.EndKey($wdStoryEnd) | Out-Null
    $sel.TypeParagraph()
    $sel.InsertBreak($wdPageBreak)

    # ---- Hauptteil ----
    $i = 0
    $imTitelblock = $true     # alles vor dem ersten --- steht schon auf dem Deckblatt
    $bilder = 0
    $tabellen = 0

    while ($i -lt $zeilen.Count) {

        $zeile = $zeilen[$i]
        $trim  = $zeile.Trim()

        # Der Kopf der Markdown-Datei wiederholt nur den Titel.
        if ($imTitelblock) {
            if ($trim -eq '---') { $imTitelblock = $false }
            $i++
            continue
        }

        if ($trim -eq '') { $i++; continue }

        # ---- Seitenumbruch ----
        if ($trim -eq '---') {
            $sel.EndKey($wdStoryEnd) | Out-Null
            $sel.InsertBreak($wdPageBreak)
            $i++
            continue
        }

        # ---- Bild ----
        if ($trim -match '^!\[(?<alt>.*?)\]\((?<pfad>.+?)\)$') {
            $bildPfad = Join-Path $dokuOrdner $Matches['pfad']
            $alt = $Matches['alt']
            if (Test-Path $bildPfad) {
                Set-Absatzstil $wdStyleNormal
                $sel.ParagraphFormat.Alignment = $wdAlignCenter
                $shape = $sel.InlineShapes.AddPicture((Resolve-Path $bildPfad).Path, $false, $true)
                $shape.LockAspectRatio = -1
                if ($shape.Width -gt $nutzbareBreite) { $shape.Width = $nutzbareBreite }
                $sel.EndKey($wdStoryEnd) | Out-Null
                $sel.TypeParagraph()
                $sel.Font.Size = 9
                $sel.Font.Italic = $true
                $sel.TypeText($alt)
                $sel.Font.Italic = $false
                $sel.Font.Size = 11
                $sel.TypeParagraph()
                $sel.ParagraphFormat.Alignment = 0
                $bilder++
            }
            else {
                Write-Warning "Bild fehlt: $bildPfad"
            }
            $i++
            continue
        }

        # ---- Ueberschriften ----
        if ($trim -match '^(?<raute>#{1,3})\s+(?<text>.+)$') {
            $tiefe = $Matches['raute'].Length
            $stil = switch ($tiefe) { 1 { $wdStyleHeading1 } 2 { $wdStyleHeading2 } default { $wdStyleHeading3 } }
            Set-Absatzstil $stil
            $sel.TypeText($Matches['text'])
            $sel.TypeParagraph()
            Set-Absatzstil $wdStyleNormal
            $i++
            continue
        }

        # ---- Tabelle ----
        if ($trim.StartsWith('|') -and ($i + 1) -lt $zeilen.Count -and $zeilen[$i + 1].Trim() -match '^\|[\s\-\|:]+\|$') {

            $tabellenZeilen = @()
            while ($i -lt $zeilen.Count -and $zeilen[$i].Trim().StartsWith('|')) {
                $roh = $zeilen[$i].Trim()
                if ($roh -notmatch '^\|[\s\-\|:]+\|$') {
                    $zellen = $roh.Trim('|').Split('|') | ForEach-Object { $_.Trim() }
                    $tabellenZeilen += , $zellen
                }
                $i++
            }

            if ($tabellenZeilen.Count -gt 0) {
                $anzZeilen = $tabellenZeilen.Count
                $anzSpalten = ($tabellenZeilen | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum

                $sel.EndKey($wdStoryEnd) | Out-Null
                $tbl = $doc.Tables.Add($sel.Range, $anzZeilen, $anzSpalten)
                $tbl.Borders.Enable = $true
                $tbl.Range.Font.Size = 10
                $tbl.Rows.Item(1).HeadingFormat = $true
                $tbl.Rows.Item(1).Range.Font.Bold = $true
                $tbl.Rows.Item(1).Shading.BackgroundPatternColor = 15132390   # helles Grau

                for ($r = 0; $r -lt $anzZeilen; $r++) {
                    for ($c = 0; $c -lt $anzSpalten; $c++) {
                        $wert = ''
                        if ($c -lt $tabellenZeilen[$r].Count) { $wert = $tabellenZeilen[$r][$c] }
                        # Auszeichnungen in Tabellenzellen schlicht entfernen.
                        $wert = $wert -replace '\*\*', '' -replace '`', ''
                        $tbl.Cell($r + 1, $c + 1).Range.Text = $wert
                    }
                }

                $tbl.Columns.AutoFit() | Out-Null
                $tbl.PreferredWidthType = 2       # wdPreferredWidthPercent
                $tbl.PreferredWidth = 100

                $sel.EndKey($wdStoryEnd) | Out-Null
                $sel.TypeParagraph()
                Set-Absatzstil $wdStyleNormal
                $tabellen++
            }
            continue
        }

        # ---- Hinweiskasten ----
        if ($trim.StartsWith('>')) {
            $text = $trim.Substring(1).Trim()
            $i++
            while ($i -lt $zeilen.Count -and $zeilen[$i].Trim().StartsWith('>')) {
                $text += ' ' + $zeilen[$i].Trim().Substring(1).Trim()
                $i++
            }

            Set-Absatzstil $wdStyleNormal
            $sel.ParagraphFormat.LeftIndent  = $word.CentimetersToPoints(0.6)
            $sel.ParagraphFormat.RightIndent = $word.CentimetersToPoints(0.6)
            $sel.ParagraphFormat.SpaceBefore = 6
            $sel.ParagraphFormat.SpaceAfter  = 6
            $sel.ParagraphFormat.Shading.BackgroundPatternColor = 14545663   # helles Gelb
            $sel.Font.Italic = $true
            Write-Inline $text
            $sel.Font.Italic = $false
            $sel.TypeParagraph()
            $sel.ParagraphFormat.Shading.BackgroundPatternColor = -16777216  # wdColorAutomatic
            $sel.ParagraphFormat.LeftIndent  = 0
            $sel.ParagraphFormat.RightIndent = 0
            continue
        }

        # ---- Aufzaehlung ----
        if ($trim -match '^-\s+(?<text>.+)$') {
            Set-Absatzstil $wdStyleListBullet
            Write-Inline $Matches['text']
            $sel.TypeParagraph()
            Set-Absatzstil $wdStyleNormal
            $i++
            continue
        }

        # ---- nummerierte Liste ----
        if ($trim -match '^\d+\.\s+(?<text>.+)$') {
            Set-Absatzstil $wdStyleListNumber
            Write-Inline $Matches['text']
            $sel.TypeParagraph()
            Set-Absatzstil $wdStyleNormal
            $i++
            continue
        }

        # ---- normaler Absatz ----
        Set-Absatzstil $wdStyleNormal
        Write-Inline $trim
        $sel.TypeParagraph()
        $i++
    }

    # ---- Kopf- und Fusszeile ----
    foreach ($abschnitt in $doc.Sections) {
        $kopf = $abschnitt.Headers.Item(1)
        $kopf.Range.Text = 'Kassenbuch der Kleingartenanlage - Bedienungsanleitung'
        $kopf.Range.Font.Size = 8
        $kopf.Range.ParagraphFormat.Alignment = $wdAlignCenter

        $fuss = $abschnitt.Footers.Item(1)
        $fuss.Range.ParagraphFormat.Alignment = $wdAlignCenter
        $fuss.Range.Font.Size = 8
        $fuss.Range.Text = 'Seite '
        $fuss.Range.Collapse(0) | Out-Null
        $doc.Fields.Add($fuss.Range, $wdFieldPage) | Out-Null
    }

    # ---- Verzeichnisse aktualisieren ----
    $toc.Update()
    $doc.Repaginate()
    $toc.Update()

    if (Test-Path $Ziel) { Remove-Item $Ziel -Force }
    $doc.SaveAs2($Ziel, 16)   # wdFormatDocumentDefault = docx
    $seiten = $doc.ComputeStatistics(2)   # wdStatisticPages

    $pdfZiel = ''
    if ($AlsPdf) {
        # Im selben Lauf exportieren, solange das Dokument noch offen ist.
        $pdfZiel = [System.IO.Path]::ChangeExtension($Ziel, '.pdf')
        if (Test-Path $pdfZiel) { Remove-Item $pdfZiel -Force }
        $doc.ExportAsFixedFormat($pdfZiel, 17, $false, 0, 0, 1, 1, 0, $true, $true, 1)
    }

    $doc.Close($false)

    Write-Host ""
    Write-Host "Fertig."
    Write-Host "  Seiten   : $seiten"
    Write-Host "  Bilder   : $bilder"
    Write-Host "  Tabellen : $tabellen"
    if ($pdfZiel) { Write-Host "  PDF      : $pdfZiel" }
}
finally {
    $word.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    [GC]::Collect()
}
