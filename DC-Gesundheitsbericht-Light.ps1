#Requires -Version 5.1
<#
.SYNOPSIS
    LIGHT-VERSION (5 von 19 Prüfpunkten, kostenlos). DC-Gesundheitsbericht: prüft alle Domänencontroller einer Domäne und erzeugt einen HTML-Bericht zu
    Replikation, SYSVOL, DNS, Zeit, Diensten, Speicherplatz, Sicherung, Updates und Ereignisprotokollen.

.DESCRIPTION
    Das Skript arbeitet ausschließlich LESEND. Es ändert nichts im Active Directory und nichts an den
    Domänencontrollern, startet keine Dienste neu, stößt keine Replikation an und überträgt keine Daten.
    Es schreibt genau eine HTML-Datei (und auf Wunsch CSV-Dateien) in den angegebenen Ausgabeordner.

    Voraussetzungen: Windows PowerShell 5.1, das Modul "ActiveDirectory" und ein Konto mit Administratorrechten
    auf den Domänencontrollern (Domänen-Admin). Am einfachsten direkt auf einem Domänencontroller in einer
    PowerShell "Als Administrator" starten. Was sich nicht abfragen lässt, erscheint im Bericht als Befund -
    es wird nie stillschweigend als "in Ordnung" gewertet.

.PARAMETER Ausgabe
    Pfad der HTML-Datei. Standard: DC-Gesundheitsbericht_<Domäne>_<Datum>.html im aktuellen Ordner.

.PARAMETER MinFreiProzent
    Laufwerke mit weniger freiem Speicher werden gemeldet (Standard 15).

.PARAMETER SicherungTage
    Ab wie vielen Tagen ohne registrierte Sicherung des Verzeichnisses gewarnt wird (Standard 7).

.PARAMETER UpdateTage
    Ab wie vielen Tagen ohne installiertes Update bzw. ohne Neustart gewarnt wird (Standard 60).

.PARAMETER EreignisStunden
    Zeitraum, aus dem Fehler der Ereignisprotokolle ausgewertet werden (Standard 24).

.PARAMETER Server
    Bestimmter Domänencontroller oder Domäne für die Abfrage des Verzeichnisses (optional).

.PARAMETER OhneDcdiag
    Überspringt den Lauf von dcdiag (spart je Domänencontroller etwa eine Minute).

.PARAMETER Csv
    Schreibt zusätzlich je Prüfpunkt mit Treffern eine CSV-Datei neben den Bericht.

.PARAMETER NichtOeffnen
    Öffnet den fertigen Bericht nicht automatisch im Browser.

.PARAMETER PassThru
    Gibt die Befunde zusätzlich als Objekte zurück (für eigene Auswertungen oder geplante Aufgaben).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\DC-Gesundheitsbericht.ps1

.EXAMPLE
    .\DC-Gesundheitsbericht.ps1 -OhneDcdiag -EreignisStunden 72 -Ausgabe C:\Berichte\dc.html

.NOTES
    IT Admin Tools by MM - Version 1.0 (10.2026)
    Kein Ersatz für eine Überwachung oder Sicherheitsprüfung. Nutzung auf eigene Verantwortung.
#>
[CmdletBinding()]
param(
    [string]$Ausgabe,
    [ValidateRange(1, 90)][int]$MinFreiProzent = 15,
    [ValidateRange(1, 3650)][int]$SicherungTage = 7,
    [ValidateRange(1, 3650)][int]$UpdateTage = 60,
    [ValidateRange(1, 720)][int]$EreignisStunden = 24,
    [string]$Server,
    [switch]$OhneDcdiag,
    [switch]$Csv,
    [switch]$NichtOeffnen,
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'
$Version = '1.0'

# CSV-Dateien öffnen sich in Excel: Werte, die mit = + - @ beginnen, würden dort als Formel ausgeführt (Namen von Dateien, Ordnern, Konten oder Computern
# stammen von Benutzern). Diese Funktion ersetzt für dieses Skript den Befehl Export-Csv und setzt vor solche Werte ein Hochkomma.
function Export-Csv {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline = $true)]$InputObject, [Parameter(Mandatory)][string]$Path, [switch]$NoTypeInformation, [string]$Delimiter = ';', [string]$Encoding = 'UTF8')
    begin { $alle = New-Object System.Collections.Generic.List[object] }
    process { if ($null -ne $InputObject) { $alle.Add($InputObject) } }
    end {
        $sicher = foreach ($o in $alle) {
            $n = [ordered]@{}
            foreach ($p in $o.PSObject.Properties) { $v = $p.Value; if (($v -is [string]) -and ($v -match '^[=+@\t\r]|^-.')) { $v = "'" + $v }; $n[$p.Name] = $v }
            [pscustomobject]$n
        }
        $sicher | Microsoft.PowerShell.Utility\Export-Csv -Path $Path -NoTypeInformation:$NoTypeInformation -Delimiter $Delimiter -Encoding $Encoding
    }
}

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host 'Das PowerShell-Modul "ActiveDirectory" fehlt.' -ForegroundColor Red
    Write-Host 'Windows 10/11: Einstellungen > Optionale Features > "RSAT: Active Directory Domain Services" hinzufügen.'
    Write-Host 'Windows Server:  Install-WindowsFeature RSAT-AD-PowerShell'
    exit 2
}
Import-Module ActiveDirectory

$ad = @{}
if ($Server) { $ad.Server = $Server }

# ------------------------------------------------------------------ Hilfsfunktionen
function Datum($d, $leer = 'nie') { if ($d) { ([datetime]$d).ToString('dd.MM.yyyy HH:mm') } else { $leer } }
function JaNein($b) { if ($b) { 'Ja' } else { 'Nein' } }
function Html($t) { [System.Net.WebUtility]::HtmlEncode([string]$t) }
function Kurz($t, $n = 220) { $t = ([string]$t -replace '\s+', ' ').Trim(); if ($t.Length -gt $n) { $t.Substring(0, $n) + ' ...' } else { $t } }
function Zeile($dc, $was, $details = '') { [pscustomobject]@{ 'Domänencontroller' = $dc; Feststellung = $was; Details = $details } }
# Eine fehlgeschlagene Abfrage ist selbst eine Feststellung - sonst sähe der Prüfpunkt fälschlich sauber aus
function Fehlzeile($dc, $fehler) { Zeile $dc 'Abfrage fehlgeschlagen - nicht geprüft' (Kurz $fehler.Exception.Message) }

$Befunde = New-Object System.Collections.Generic.List[object]
function Befund {
    param([string]$Id, [string]$Bereich, [string]$Titel, [ValidateSet('Hoch', 'Mittel', 'Niedrig', 'Info')][string]$Stufe,
        $Objekte, [string]$Erklaerung, [string]$Empfehlung)
    $liste = @($Objekte | Where-Object { $null -ne $_ })
    $status = $Stufe
    if ($liste.Count -eq 0 -and $Stufe -ne 'Info') { $status = 'OK' }
    $Befunde.Add([pscustomobject]@{
            Id = $Id; Bereich = $Bereich; Titel = $Titel; Stufe = $status; Anzahl = $liste.Count
            Objekte = $liste; Erklaerung = $Erklaerung; Empfehlung = $Empfehlung
        })
}

# Verbindung für WMI-Abfragen: lokal ohne Netz, sonst WinRM und ersatzweise DCOM
$Sitzungen = @{}
function Wmi($dc) {
    if ($dc.Lokal) { return @{} }
    if (-not $Sitzungen.ContainsKey($dc.Name)) {
        try { $s = New-CimSession -ComputerName $dc.HostName -OperationTimeoutSec 30 -ErrorAction Stop }
        catch { $s = New-CimSession -ComputerName $dc.HostName -OperationTimeoutSec 30 -SessionOption (New-CimSessionOption -Protocol Dcom) -ErrorAction Stop }
        $Sitzungen[$dc.Name] = $s
    }
    @{ CimSession = $Sitzungen[$dc.Name] }
}
$HKLM = [uint32]2147483650
function RegText($dc, $schluessel, $wert) {
    $w = Wmi $dc
    (Invoke-CimMethod -Namespace root\cimv2 -ClassName StdRegProv -MethodName GetStringValue -Arguments @{ hDefKey = $HKLM; sSubKeyName = $schluessel; sValueName = $wert } @w).sValue
}
function RegSchluesselDa($dc, $schluessel) {
    $w = Wmi $dc
    (Invoke-CimMethod -Namespace root\cimv2 -ClassName StdRegProv -MethodName EnumValues -Arguments @{ hDefKey = $HKLM; sSubKeyName = $schluessel } @w).ReturnValue -eq 0
}

# ------------------------------------------------------------------ Daten lesen
Write-Host 'DC-Gesundheitsbericht - es wird nur gelesen, nichts verändert.' -ForegroundColor Cyan
Write-Host '  Domäne und Domänencontroller ...'
$dom = Get-ADDomain @ad
$wald = Get-ADForest @ad
$dn = $dom.DistinguishedName
$jetzt = Get-Date
$ich = $env:COMPUTERNAME
$dcs = @(Get-ADDomainController -Filter * @ad | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{
            Name = $_.Name; HostName = $_.HostName; IP = $_.IPv4Address; Standort = $_.Site; OS = $_.OperatingSystem
            GC = $_.IsGlobalCatalog; RODC = $_.IsReadOnly; Rollen = @($_.OperationMasterRoles | ForEach-Object { "$_" })
            Lokal = ($_.Name -eq $ich); Erreichbar = $false; Neustart = $null
        }
    })
$dcIPs = @($dcs | ForEach-Object { $_.IP } | Where-Object { $_ })


$e = @()
if ($dcs.Count -eq 1) { $e = Zeile $dcs[0].Name 'Einziger Domänencontroller der Domäne' 'Kein Ausfallschutz für Anmeldung und DNS' }
Befund 'E2' 'Erreichbarkeit' 'Nur ein Domänencontroller vorhanden' 'Mittel' $e `
    'Fällt der einzige Domänencontroller aus, kann sich niemand mehr anmelden, und die Namensauflösung steht. Die Wiederherstellung hängt dann allein an der Sicherung.' `
    'Einen zweiten Domänencontroller einrichten (auch als virtuelle Maschine auf anderer Hardware) - oder das Risiko bewusst akzeptieren und die Wiederherstellung regelmäßig üben.'

# --- SYSVOL
Write-Host '  SYSVOL und DFS-Replikation ...'
$y1 = @(); $y3 = @()
foreach ($dc in $da) {
    foreach ($frei in 'SYSVOL', 'NETLOGON') {
        try { if (-not (Test-Path -LiteralPath "\\$($dc.HostName)\$frei" -ErrorAction Stop)) { $y1 += Zeile $dc.Name "Freigabe $frei fehlt oder ist nicht erreichbar" } }
        catch { $y1 += Fehlzeile $dc.Name $_ }
    }
    try {
        $w = Wmi $dc
        $info = @(Get-CimInstance -Namespace root\MicrosoftDFS -ClassName DfsrReplicatedFolderInfo -Filter "ReplicatedFolderName='SYSVOL Share'" -ErrorAction Stop @w)
        $zustand = @{ 0 = 'Nicht initialisiert'; 1 = 'Initialisiert'; 2 = 'Erstabgleich läuft'; 3 = 'Automatische Wiederherstellung'; 4 = 'Normal'; 5 = 'Fehler' }
        if ($info.Count -eq 0) { $y3 += Zeile $dc.Name 'Kein Zustand der SYSVOL-Replikation abfragbar - nicht geprüft' 'Läuft der Dienst DFS-Replikation? Bei Abgleich über FRS siehe Y2' }
        foreach ($i in $info) {
            if ($i.State -ne 4) { $y3 += Zeile $dc.Name "SYSVOL-Replikation im Zustand '$($zustand[[int]$i.State])'" 'Erwartet: Normal' }
        }
    } catch { $y3 += Fehlzeile $dc.Name $_ }
}
Befund 'Y1' 'SYSVOL' 'Freigabe SYSVOL oder NETLOGON fehlt' 'Hoch' $y1 `
    'Über diese Freigaben verteilt ein Domänencontroller Gruppenrichtlinien und Anmeldeskripte. Fehlen sie, gilt der Server nicht als vollwertiger Domänencontroller.' `
    'Zustand der SYSVOL-Replikation prüfen (Ereignisprotokoll "DFS-Replikation"); häufig hat der Erstabgleich nach der Einrichtung nie abgeschlossen.'

# --- Zeit
Write-Host '  Zeit ...'
$z1 = @(); $z2 = @()
foreach ($dc in ($da | Where-Object { -not $_.Lokal })) {
    try {
        $aus = (& w32tm.exe /stripchart "/computer:$($dc.HostName)" /samples:1 /dataonly 2>&1 | Out-String)
        if ($aus -match ',\s*([+-]\d+[.,]\d+)s') {
            $diff = [double]($Matches[1] -replace ',', '.')
            if ([math]::Abs($diff) -gt 10) { $z1 += Zeile $dc.Name ("Uhrzeit weicht um {0:N0} Sekunden ab" -f $diff) "gemessen gegen $ich" }
        } else { $z1 += Zeile $dc.Name 'Abfrage fehlgeschlagen - nicht geprüft' (Kurz $aus) }
    } catch { $z1 += Fehlzeile $dc.Name $_ }
}
Befund 'Z1' 'Zeit' 'Uhrzeit weicht zwischen Domänencontrollern ab' 'Hoch' $z1 `
    'Die Anmeldung über Kerberos verträgt höchstens fünf Minuten Abweichung. Gemessen wird gegen den Rechner, auf dem das Skript läuft; gemeldet wird ab 10 Sekunden.' `
    'Zeitabgleich prüfen: "w32tm /query /status" auf dem betroffenen Server. Bei virtuellen Maschinen den Zeitabgleich mit dem Host abschalten, wenn die Domäne die Zeit vorgibt.'

Befund 'P1' 'Speicherplatz' "Laufwerke mit weniger als $MinFreiProzent % freiem Speicher" 'Hoch' $p1 `
    'Läuft das Laufwerk mit der Verzeichnisdatenbank oder dem Protokoll voll, bleibt der Verzeichnisdienst stehen.' `
    'Speicher freigeben (alte Protokolle, Updates bereinigen, Dateien verlagern) oder das Laufwerk vergrößern.'
Befund 'U2' 'Updates' 'Neustart steht aus' 'Mittel' $u2 `
    'Windows hat Updates installiert, die erst nach einem Neustart wirksam werden.' `
    'Server im nächsten Wartungsfenster neu starten - nie alle Domänencontroller gleichzeitig.'

foreach ($s in $Sitzungen.Values) { try { Remove-CimSession -CimSession $s } catch { } }

# Besteht ein Prüfpunkt nur aus "nicht geprüft"-Zeilen (Abfrage fehlgeschlagen), ist das kein Mangel dieses Prüfpunkts: als Info führen statt
# mit seiner Stufe, damit es in den Kacheln nicht als hoher Befund zählt.
$platzTexte = @('Abfrage fehlgeschlagen - nicht geprüft', 'Kein Zustand der SYSVOL-Replikation abfragbar - nicht geprüft', 'PDC-Emulator nicht erreichbar - Zeitquelle nicht geprüft', 'Kein beschreibbarer Domänencontroller erreichbar - nicht geprüft', 'dcdiag.exe ist auf diesem Rechner nicht vorhanden - nicht geprüft')
foreach ($x in ($Befunde | Where-Object { $_.Anzahl -gt 0 -and $_.Stufe -ne 'Info' })) {
    # nur der feste Text, den dieses Skript selbst erzeugt (keine Teilstring-Suche in Namen oder Beschreibungen)
    if (-not @($x.Objekte | Where-Object { -not (@($_.PSObject.Properties.Value | Where-Object { $_ -is [string] -and $platzTexte -ccontains $_ }).Count -gt 0) }).Count) { $x.Stufe = 'Info' }
}

# ------------------------------------------------------------------ Bericht
Write-Host '  Bericht schreiben ...'
if (-not $Ausgabe) { $Ausgabe = Join-Path (Get-Location).Path ("DC-Gesundheitsbericht_{0}_{1:yyyy-MM-dd_HHmm}.html" -f $dom.DNSRoot, $jetzt) }
$Ausgabe = [System.IO.Path]::GetFullPath($Ausgabe)
$ordner = Split-Path $Ausgabe -Parent
if (-not (Test-Path $ordner)) { New-Item -ItemType Directory -Path $ordner -Force | Out-Null }

$rang = @{ Hoch = 0; Mittel = 1; Niedrig = 2; Info = 3; OK = 4 }
$zahl = @{}
foreach ($s in 'Hoch', 'Mittel', 'Niedrig', 'OK') { $zahl[$s] = @($Befunde | Where-Object Stufe -eq $s).Count }
$wer = "$env:USERDOMAIN\$env:USERNAME"

$css = @'
body{font-family:Segoe UI,Arial,sans-serif;color:#2b2f36;margin:0;background:#f4f6fa}
.seite{max-width:1100px;margin:0 auto;padding:28px 24px 60px}
h1{color:#1f5fae;margin:0 0 4px;font-size:28px} h2{color:#1f5fae;font-size:19px;margin:34px 0 10px;border-bottom:2px solid #d0d5dd;padding-bottom:6px}
.meta{color:#6b7280;font-size:13px;margin-bottom:22px}
.kacheln{display:flex;gap:14px;flex-wrap:wrap;margin:18px 0}
.kachel{flex:1;min-width:150px;background:#fff;border-radius:10px;padding:14px 18px;border-left:8px solid #999;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.kachel b{display:block;font-size:30px} .kachel span{font-size:13px;color:#6b7280}
table{border-collapse:collapse;width:100%;background:#fff;font-size:13px;margin:8px 0 4px}
th{background:#1f5fae;color:#fff;text-align:left;padding:7px 9px;font-weight:600} td{padding:6px 9px;border-bottom:1px solid #e5e7eb;vertical-align:top}
tr:nth-child(even) td{background:#f9fafb}
.marke{display:inline-block;min-width:58px;text-align:center;padding:2px 9px;border-radius:11px;color:#fff;font-size:12px;font-weight:600}
.m-Hoch{background:#d92d20}.m-Mittel{background:#f79009}.m-Niedrig{background:#ca8a04}.m-OK{background:#12b76a}.m-Info{background:#1f5fae}
.befund{background:#fff;border-radius:10px;padding:16px 20px;margin:14px 0;border-left:8px solid #999;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.befund h3{margin:0 0 6px;font-size:16px} .befund p{margin:5px 0;font-size:14px;line-height:1.45} .befund .emp{color:#14532d}
.Hoch{border-color:#d92d20}.Mittel{border-color:#f79009}.Niedrig{border-color:#eab308}.OK{border-color:#12b76a}.Info{border-color:#1f5fae}
.klein{color:#6b7280;font-size:12px} a{color:#1f5fae;text-decoration:none}
pre{background:#fff;border:1px solid #d0d5dd;border-radius:8px;padding:12px 14px;font-size:12px;white-space:pre-wrap;word-break:break-word}
.fuss{margin-top:40px;color:#6b7280;font-size:12px;border-top:1px solid #d0d5dd;padding-top:12px}
@media print{body{background:#fff}.befund,.kachel{box-shadow:none;border:1px solid #d0d5dd;border-left-width:8px;break-inside:avoid}}
'@

$h = New-Object System.Text.StringBuilder
[void]$h.Append("<!DOCTYPE html><html lang='de'><head><meta charset='utf-8'><title>DC-Gesundheitsbericht $(Html $dom.DNSRoot)</title><style>$css</style></head><body><div class='seite'>")
[void]$h.Append("<h1>DC-Gesundheitsbericht Light</h1><div class='meta'>Domäne <b>$(Html $dom.DNSRoot)</b> · erstellt am $($jetzt.ToString('dd.MM.yyyy HH:mm')) von $(Html $wer) auf $(Html $ich)</div>")
[void]$h.Append("<div class='kacheln'>")
foreach ($s in 'Hoch', 'Mittel', 'Niedrig', 'OK') {
    $text = @{ Hoch = 'Befunde mit hoher Priorität'; Mittel = 'mit mittlerer Priorität'; Niedrig = 'mit niedriger Priorität'; OK = 'Prüfpunkte ohne Befund' }[$s]
    [void]$h.Append("<div class='kachel $s'><b>$($zahl[$s])</b><span>$text</span></div>")
}
[void]$h.Append("</div>")

[void]$h.Append("<h2>Domänencontroller</h2>")
[void]$h.Append((($dcs | ForEach-Object {
                [pscustomobject][ordered]@{
                    Name = $_.Name; Standort = $_.Standort; 'IP-Adresse' = $_.IP; Betriebssystem = $_.OS
                    'Globaler Katalog' = JaNein $_.GC; 'Schreibgeschützt' = JaNein $_.RODC; Erreichbar = JaNein $_.Erreichbar
                    'Letzter Start' = Datum $_.Neustart 'unbekannt'
                }
            } | ConvertTo-Html -Fragment) -join "`n"))

[void]$h.Append("<h2>Umgebung</h2><table><tr><th>Eigenschaft</th><th>Wert</th></tr>")
$umgebung = [ordered]@{
    'Domäne (NetBIOS)'                       = "$($dom.DNSRoot) ($($dom.NetBIOSName))"
    'Funktionsebene Domäne / Gesamtstruktur' = "$($dom.DomainMode) / $($wald.ForestMode)"
    'Standorte'                              = ($dcs | ForEach-Object { $_.Standort } | Sort-Object -Unique) -join ', '
}
foreach ($r in $rollen.Keys) { $umgebung[$r] = $rollen[$r] }
$umgebung['Schwellen dieses Berichts'] = "Speicher unter $MinFreiProzent % frei, Sicherung älter als $SicherungTage Tage, Updates/Neustart älter als $UpdateTage Tage, Ereignisse der letzten $EreignisStunden Stunden"
foreach ($k in $umgebung.Keys) { [void]$h.Append("<tr><td>$(Html $k)</td><td>$(Html $umgebung[$k])</td></tr>") }
[void]$h.Append("</table>")

[void]$h.Append("<h2>Übersicht der Prüfpunkte</h2><table><tr><th>Nr.</th><th>Bereich</th><th>Prüfpunkt</th><th>Treffer</th><th>Bewertung</th></tr>")
foreach ($x in $Befunde) {
    [void]$h.Append("<tr><td>$($x.Id)</td><td>$(Html $x.Bereich)</td><td><a href='#$($x.Id)'>$(Html $x.Titel)</a></td><td>$($x.Anzahl)</td><td><span class='marke m-$($x.Stufe)'>$($x.Stufe)</span></td></tr>")
}
[void]$h.Append("</table>")

[void]$h.Append("<h2>Befunde</h2>")
foreach ($x in ($Befunde | Where-Object { $_.Stufe -ne 'OK' } | Sort-Object { $rang[$_.Stufe] }, { [regex]::Replace($_.Id, '\d+', '') }, { [int]('0' + [regex]::Match($_.Id, '\d+').Value) })) {
    [void]$h.Append("<div class='befund $($x.Stufe)' id='$($x.Id)'><h3><span class='marke m-$($x.Stufe)'>$($x.Stufe)</span> $($x.Id) · $(Html $x.Titel) <span class='klein'>($($x.Anzahl) Treffer)</span></h3>")
    [void]$h.Append("<p>$(Html $x.Erklaerung)</p><p class='emp'><b>Empfehlung:</b> $(Html $x.Empfehlung)</p>")
    if ($x.Anzahl -gt 0) { [void]$h.Append((($x.Objekte | ConvertTo-Html -Fragment) -join "`n")) }
    [void]$h.Append("</div>")
}
if (-not @($Befunde | Where-Object { $_.Stufe -ne 'OK' }).Count) { [void]$h.Append("<p>Keine Auffälligkeiten gefunden.</p>") }
$ok = @($Befunde | Where-Object Stufe -eq 'OK')
if ($ok.Count) {
    [void]$h.Append("<h2>Ohne Befund</h2><table><tr><th>Nr.</th><th>Prüfpunkt</th></tr>")
    foreach ($x in $ok) { [void]$h.Append("<tr id='$($x.Id)'><td>$($x.Id)</td><td>$(Html $x.Titel)</td></tr>") }
    [void]$h.Append("</table>")
}
if ($dcdiag.Count) {
    [void]$h.Append("<h2>Ausgabe von dcdiag</h2><p class='klein'>Unveränderte Ausgabe im stillen Modus (nur Auffälligkeiten).</p>")
    foreach ($k in ($dcdiag.Keys | Sort-Object)) { [void]$h.Append("<h3>$(Html $k)</h3><pre>$(Html ($dcdiag[$k] -join "`n"))</pre>") }
}
[void]$h.Append("<div class='fuss'>DC-Gesundheitsbericht Light (5 Prüfpunkte, Vollversion: https://mueller86.gumroad.com/l/bseljr) Version $Version · IT Admin Tools by MM · Das Skript hat ausschließlich gelesen und nichts verändert. Der Bericht zeigt Auffälligkeiten anhand fester Regeln zum Zeitpunkt des Laufs; er ersetzt keine laufende Überwachung und erhebt keinen Anspruch auf Vollständigkeit. Dieser Bericht enthält Servernamen und Adressen - bitte vertraulich behandeln.</div>")
[void]$h.Append("</div></body></html>")
[System.IO.File]::WriteAllText($Ausgabe, $h.ToString(), (New-Object System.Text.UTF8Encoding($true)))

if ($Csv) {
    foreach ($x in ($Befunde | Where-Object { $_.Anzahl -gt 0 })) {
        $x.Objekte | Export-Csv -Path ([System.IO.Path]::ChangeExtension($Ausgabe, $null).TrimEnd('.') + "_$($x.Id).csv") -NoTypeInformation -Delimiter ';' -Encoding UTF8
    }
}

Write-Host ''
Write-Host ("Fertig: {0} hoch, {1} mittel, {2} niedrig, {3} ohne Befund." -f $zahl.Hoch, $zahl.Mittel, $zahl.Niedrig, $zahl.OK) -ForegroundColor Green
Write-Host "Bericht: $Ausgabe"
if (-not $NichtOeffnen -and [Environment]::UserInteractive) { try { Start-Process $Ausgabe } catch { } }
if ($PassThru) { $Befunde }
