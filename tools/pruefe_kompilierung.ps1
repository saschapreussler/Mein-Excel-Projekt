# -----------------------------------------------------------------
# pruefe_kompilierung.ps1
#
# Prueft, ob das gesamte VBA-Projekt der Arbeitsmappe fehlerfrei
# kompiliert. Gedacht als Kontrolle nach jedem Repo -> Excel-Sync.
#
# Warum nicht einfach "Debuggen -> Kompilieren" per COM ausloesen:
# Der Menuepunkt meldet sein Ergebnis nicht zurueck, und ein
# Fehlerdialog im VBA-Editor wuerde eine unsichtbare Excel-Instanz
# blockieren. Stattdessen wird die Editor-Option "Kompilieren bei
# Bedarf" kurzzeitig abgeschaltet. Dann uebersetzt VBA vor der ersten
# Ausfuehrung das komplette Projekt, und ein Fehler in irgendeinem
# Modul laesst den Makroaufruf scheitern - abfangbar und ohne Dialog.
#
# Rueckgabe: Exit-Code 0 = fehlerfrei, 1 = Kompilierfehler.
# -----------------------------------------------------------------

param(
    [string]$Workbook = ''
)

$ErrorActionPreference = 'Stop'

$skriptOrdner = Split-Path -Parent $MyInvocation.MyCommand.Definition
$projektOrdner = Split-Path -Parent $skriptOrdner

if ([string]::IsNullOrWhiteSpace($Workbook)) {
    $Workbook = Join-Path $projektOrdner 'excel\Programm Kassenbuch 2018_v2.7.8.xlsm'
}

if (-not (Test-Path $Workbook)) {
    Write-Host "Arbeitsmappe nicht gefunden: $Workbook" -ForegroundColor Red
    exit 1
}

# Der Testaufruf muss eine oeffentliche, parameterlose und
# nebenwirkungsfreie Funktion sein. IsGenerating liest nur ein Flag.
$testAufruf = 'mod_Uebersicht_Generator.IsGenerating'

$vbaSchluessel = 'HKCU:\Software\Microsoft\VBA\7.1\Common'
$alterWert = $null
if (Test-Path $vbaSchluessel) {
    $alterWert = (Get-ItemProperty $vbaSchluessel -Name CompileOnDemand -ErrorAction SilentlyContinue).CompileOnDemand
    Set-ItemProperty $vbaSchluessel -Name CompileOnDemand -Value 0
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.AutomationSecurity = 1

$erfolg = $false
$meldung = ''

try {
    $mappe = $excel.Workbooks.Open((Resolve-Path $Workbook).Path, $false, $true)
    try {
        [void]$excel.Run($testAufruf)
        $erfolg = $true
    }
    catch {
        $meldung = $_.Exception.Message
    }
    $mappe.Close($false)
}
catch {
    $meldung = $_.Exception.Message
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    if ($null -ne $alterWert) {
        Set-ItemProperty $vbaSchluessel -Name CompileOnDemand -Value $alterWert
    }
}

if ($erfolg) {
    Write-Host 'VBA-Projekt kompiliert fehlerfrei.' -ForegroundColor Green
    exit 0
}

Write-Host 'Das VBA-Projekt kompiliert NICHT.' -ForegroundColor Red
Write-Host "Der Testaufruf $testAufruf schlug fehl:"
Write-Host "  $meldung"
Write-Host ''
Write-Host 'Naechster Schritt: Arbeitsmappe oeffnen und im VBA-Editor'
Write-Host 'Debuggen -> Kompilieren ausfuehren. Dort steht die Fundstelle.'
exit 1
