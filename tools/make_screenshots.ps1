<#
.SYNOPSIS
    Erzeugt die Bildschirmfotos fuer die Bedienungsanleitung.

.DESCRIPTION
    Oeffnet die Arbeitsmappe sichtbar, ruft nacheinander jedes Blatt auf und
    fotografiert das Excel-Fenster. So zeigt die Anleitung genau das Bild, das
    der Nutzer auch vor sich hat, samt Schaltflaechen und Menueband.

    Die Bilder landen in doku/screenshots. Vorhandene Dateien werden ersetzt.

    Waehrend des Laufs darf der Rechner nicht anderweitig benutzt werden, weil
    das Excel-Fenster im Vordergrund stehen muss.

.NOTES
    Nur auf ausdruecklichen Auftrag ausfuehren. Das Skript aendert keine
    Arbeitsmappe und speichert nicht.
#>

[CmdletBinding()]
param(
    [string] $Workbook,
    [string] $OutDir,
    [int]    $WaitMs = 1400
)

$ErrorActionPreference = 'Stop'

# Projektwurzel unabhaengig vom Aufrufort bestimmen.
$skriptOrdner = Split-Path -Parent $MyInvocation.MyCommand.Definition
$wurzel       = Split-Path -Parent $skriptOrdner
if (-not $Workbook) { $Workbook = Join-Path $wurzel 'excel\Programm Kassenbuch 2018_v2.7.8.xlsm' }
if (-not $OutDir)   { $OutDir   = Join-Path $wurzel 'doku\screenshots' }

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

if (-not ('Win32ScreenShot' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public class Win32ScreenShot {
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
'@
}

# Blattname -> Dateiname. Die Reihenfolge entspricht der Anleitung.
$blaetter = [ordered]@{
    'Startmenü'                     = '01-startmenue'
    'Bankkonto'                     = '02-bankkonto'
    'Zahlungsübersicht'             = '03-zahlungsuebersicht'
    'Dashboard Mitgliederzahlungen' = '04-dashboard'
    'Mitgliederliste'               = '05-mitgliederliste'
    'Daten'                         = '06-daten'
    'Einstellungen'                 = '07-einstellungen'
    'Vereinskasse'                  = '08-vereinskasse'
    'Strom'                         = '09-strom'
    'Wasser'                        = '10-wasser'
    'Finanz-Übersicht'              = '11-finanz-uebersicht'
    'Mitgliederhistorie'            = '12-mitgliederhistorie'
}

$wbPath = (Resolve-Path $Workbook).Path
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$OutDir = (Resolve-Path $OutDir).Path

Write-Host "Arbeitsmappe : $wbPath"
Write-Host "Zielordner   : $OutDir"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1   # msoAutomationSecurityLow: Makros zulassen

$erzeugt = 0
try {
    $book = $xl.Workbooks.Open($wbPath, $false, $true)   # schreibgeschuetzt oeffnen

    $xl.WindowState = -4137                              # xlMaximized
    [void][Win32ScreenShot]::ShowWindow([IntPtr]$xl.Hwnd, 3)
    [void][Win32ScreenShot]::SetForegroundWindow([IntPtr]$xl.Hwnd)
    Start-Sleep -Milliseconds 1200

    foreach ($blatt in $blaetter.Keys) {

        $ws = $null
        try { $ws = $book.Worksheets($blatt) } catch { $ws = $null }
        if ($null -eq $ws) {
            Write-Warning "Blatt nicht gefunden, wird uebersprungen: $blatt"
            continue
        }

        if (-not $ws.Visible) {
            Write-Warning "Blatt ausgeblendet, wird uebersprungen: $blatt"
            continue
        }

        $ws.Activate()

        # Ansicht so einpassen, dass moeglichst viel des Blattes sichtbar ist.
        # Bei sehr breiten Blaettern wuerde der Zoom unlesbar klein, deshalb
        # die Untergrenze. Die Mappe ist schreibgeschuetzt geoeffnet, der Zoom
        # wird also nicht zurueckgeschrieben.
        try {
            [void]$ws.UsedRange.Select()
            $xl.ActiveWindow.Zoom = $true
            if ($xl.ActiveWindow.Zoom -lt 55)  { $xl.ActiveWindow.Zoom = 55 }
            if ($xl.ActiveWindow.Zoom -gt 100) { $xl.ActiveWindow.Zoom = 100 }
        } catch { }

        try { [void]$ws.Range('A1').Select() } catch { }
        try { $xl.ActiveWindow.ScrollRow = 1; $xl.ActiveWindow.ScrollColumn = 1 } catch { }

        [void][Win32ScreenShot]::SetForegroundWindow([IntPtr]$xl.Hwnd)
        Start-Sleep -Milliseconds $WaitMs

        $rect = New-Object Win32ScreenShot+RECT
        [void][Win32ScreenShot]::GetWindowRect([IntPtr]$xl.Hwnd, [ref]$rect)

        $breite = $rect.Right - $rect.Left
        $hoehe  = $rect.Bottom - $rect.Top
        if ($breite -le 0 -or $hoehe -le 0) {
            Write-Warning "Fenstermasse unbrauchbar bei: $blatt"
            continue
        }

        $bmp = New-Object System.Drawing.Bitmap $breite, $hoehe
        $gfx = [System.Drawing.Graphics]::FromImage($bmp)
        $gfx.CopyFromScreen($rect.Left, $rect.Top, 0, 0, $bmp.Size)
        $gfx.Dispose()

        $ziel = Join-Path $OutDir ($blaetter[$blatt] + '.png')
        $bmp.Save($ziel, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()

        $erzeugt++
        Write-Host ("  {0,-32} -> {1}" -f $blatt, (Split-Path $ziel -Leaf))
    }

    $book.Close($false)
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
}

Write-Host ""
Write-Host "Fertig. $erzeugt Bildschirmfotos erzeugt."
