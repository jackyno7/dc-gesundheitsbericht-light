<#
.SYNOPSIS
    LIGHT VERSION (5 of 19 checks, free). DC health report: checks all domain controllers of a domain and creates an HTML report on
    replication, SYSVOL, DNS, time, services, disk space, backup, updates and event logs.

    Deutsch:
    LIGHT-VERSION (5 von 19 Prüfpunkten, kostenlos). DC-Gesundheitsbericht: prüft alle Domänencontroller einer Domäne und erzeugt einen HTML-Bericht zu
    Replikation, SYSVOL, DNS, Zeit, Diensten, Speicherplatz, Sicherung, Updates und Ereignisprotokollen.

.DESCRIPTION
    The script works exclusively READ-ONLY. It changes nothing in Active Directory and nothing on the
    domain controllers, restarts no services, triggers no replication and transfers no data.
    It writes exactly one HTML file (and, on request, CSV files) to the specified output folder.

    Requirements: Windows PowerShell 5.1, the "ActiveDirectory" module and an account with administrator rights
    on the domain controllers (domain admin). The easiest way is to start it directly on a domain controller in a
    PowerShell "as administrator". Anything that cannot be queried appears in the report as a finding -
    it is never silently rated "OK".

    Deutsch:
    Das Skript arbeitet ausschließlich LESEND. Es ändert nichts im Active Directory und nichts an den
    Domänencontrollern, startet keine Dienste neu, stößt keine Replikation an und überträgt keine Daten.
    Es schreibt genau eine HTML-Datei (und auf Wunsch CSV-Dateien) in den angegebenen Ausgabeordner.

    Voraussetzungen: Windows PowerShell 5.1, das Modul "ActiveDirectory" und ein Konto mit Administratorrechten
    auf den Domänencontrollern (Domänen-Admin). Am einfachsten direkt auf einem Domänencontroller in einer
    PowerShell "Als Administrator" starten. Was sich nicht abfragen lässt, erscheint im Bericht als Befund -
    es wird nie stillschweigend als "in Ordnung" gewertet.

.PARAMETER Ausgabe
    Path of the HTML file. Default: DC-Gesundheitsbericht_<domain>_<date>.html in the current folder.

    Deutsch:
    Pfad der HTML-Datei. Standard: DC-Gesundheitsbericht_<Domäne>_<Datum>.html im aktuellen Ordner.

.PARAMETER MinFreiProzent
    Drives with less free space are reported (default 15).

    Deutsch:
    Laufwerke mit weniger freiem Speicher werden gemeldet (Standard 15).

.PARAMETER SicherungTage
    Number of days without a registered backup of the directory after which a warning is issued (default 7).

    Deutsch:
    Ab wie vielen Tagen ohne registrierte Sicherung des Verzeichnisses gewarnt wird (Standard 7).

.PARAMETER UpdateTage
    Number of days without an installed update or without a restart after which a warning is issued (default 60).

    Deutsch:
    Ab wie vielen Tagen ohne installiertes Update bzw. ohne Neustart gewarnt wird (Standard 60).

.PARAMETER EreignisStunden
    Period from which errors in the event logs are evaluated (default 24).

    Deutsch:
    Zeitraum, aus dem Fehler der Ereignisprotokolle ausgewertet werden (Standard 24).

.PARAMETER Server
    A specific domain controller or domain for querying the directory (optional).

    Deutsch:
    Bestimmter Domänencontroller oder Domäne für die Abfrage des Verzeichnisses (optional).

.PARAMETER OhneDcdiag
    Skips running dcdiag (saves about a minute per domain controller).

    Deutsch:
    Überspringt den Lauf von dcdiag (spart je Domänencontroller etwa eine Minute).

.PARAMETER Csv
    Additionally writes a CSV file for each check point with hits next to the report.

    Deutsch:
    Schreibt zusätzlich je Prüfpunkt mit Treffern eine CSV-Datei neben den Bericht.

.PARAMETER NichtOeffnen
    Does not automatically open the finished report in the browser.

    Deutsch:
    Öffnet den fertigen Bericht nicht automatisch im Browser.

.PARAMETER PassThru
    Additionally returns the findings as objects (for your own evaluations or scheduled tasks).

    Deutsch:
    Gibt die Befunde zusätzlich als Objekte zurück (für eigene Auswertungen oder geplante Aufgaben).

.PARAMETER Language
    Language of the report and the console output: 'de' (German) or 'en' (English). Default: German if the Windows display language
    is German, otherwise English.

    Deutsch:
    Sprache von Bericht und Konsolenausgabe: 'de' (Deutsch) oder 'en' (Englisch). Standard: Deutsch, wenn die Anzeigesprache von
    Windows Deutsch ist, sonst Englisch.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\DC-Gesundheitsbericht.ps1

.EXAMPLE
    .\DC-Gesundheitsbericht.ps1 -OhneDcdiag -EreignisStunden 72 -Ausgabe C:\Berichte\dc.html

.NOTES
    IT Admin Tools by MM - Version 1.0 (10.2026)
    Kein Ersatz für eine Überwachung oder Sicherheitsprüfung. Nutzung auf eigene Verantwortung.
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('de', 'en')][string]$Language = $(if ((Get-UICulture).TwoLetterISOLanguageName -eq 'de') { 'de' } else { 'en' }),
    [Alias('Output')][string]$Ausgabe,
    [ValidateRange(1, 90)][int]$MinFreiProzent = 15,
    [ValidateRange(1, 3650)][int]$SicherungTage = 7,
    [ValidateRange(1, 3650)][int]$UpdateTage = 60,
    [ValidateRange(1, 720)][int]$EreignisStunden = 24,
    [string]$Server,
    [switch]$OhneDcdiag,
    [switch]$Csv,
    [Alias('NoOpen')][switch]$NichtOeffnen,
    [switch]$PassThru
)
# Sprache des Berichts: de oder en (Parameter -Language). T 'deutsch' 'english' liefert den Text der gewählten Sprache.
$En = ($Language -eq 'en')
function T([string]$de, [string]$en) { if ($script:En) { $en } else { $de } }

#@SPRACHE@
# Englische Berichte: Spaltenüberschriften aus der Übersetzungstabelle (intern bleiben die deutschen Namen, der Code liest sie weiter)
$SpaltenEn = @{
    'Abfrage' = 'Query'
    'Administrator' = 'Administrator'
    'Adresse' = 'Address'
    'Aktiv' = 'Active'
    'Aktiv verknüpft mit' = 'Active linked to'
    'Aktive Verknüpfungen' = 'Active links'
    'Alter in Tagen' = 'Age in days'
    'Angelegt' = 'Created'
    'Anmeldung' = 'Last logon'
    'Anmeldungen' = 'Logons'
    'Anteil (%)' = 'Share (%)'
    'Anzahl' = 'Count'
    'Art' = 'Account type'
    'Auf Domänencontroller' = 'On domain controller'
    'Aufräumen' = 'Cleanup'
    'Auslastung' = 'Utilization'
    'Auslösender Rechner' = 'Triggering computer'
    'Auslösender Rechner (Anzahl)' = 'Triggering computers (count)'
    'Aussteller' = 'Issuer'
    'Beispiele' = 'Examples'
    'Belegt' = 'Used'
    'Belegt (%)' = 'Used (%)'
    'Benutzer und Computer (auch verschachtelt)' = 'Users and computers (also nested)'
    'Benutzereinstellungen' = 'User settings'
    'Benutzername' = 'Username'
    'Berechtigter' = 'Principal'
    'Bereich' = 'Scope'
    'Bereiche' = 'Scopes'
    'Beschreibung' = 'Description'
    'Besitzer' = 'Owner'
    'Betriebssystem' = 'OS'
    'Betroffene Konten' = 'Affected accounts'
    'Bewertung' = 'Rating'
    'Bis' = 'To'
    'BitLocker' = 'BitLocker'
    'Computer' = 'Computer'
    'Computereinstellungen' = 'Computer settings'
    'Darf anwenden' = 'Can apply'
    'Darf delegieren an' = 'Can delegate to'
    'Darf sich hier als Benutzer ausgeben' = 'Can log on as user'
    'Datei' = 'File'
    'Dateien' = 'Files'
    'Dateisystem' = 'File system'
    'Dateityp' = 'File type'
    'Datenbank' = 'Database'
    'Detail' = 'Detail'
    'Details' = 'Details'
    'Direkte Mitglieder' = 'Direct members'
    'Domäne' = 'Domain'
    'Domänencontroller' = 'Domain controller'
    'Domänennamenmaster' = 'Domain naming master'
    'Dynamische Updates' = 'Dynamic updates'
    'E-Mail' = 'Email'
    'Eigenschaft' = 'Property'
    'Einstellung' = 'Setting'
    'Empfohlen' = 'Recommended'
    'Enthaltene Gruppen' = 'Included groups'
    'Ereignisse' = 'Events'
    'Erreichbar' = 'Reachable'
    'Erste' = 'First'
    'Erster' = 'First'
    'Erzwungen' = 'Forced'
    'Failover' = 'Failover'
    'Falsches Kennwort' = 'Wrong password'
    'Fehlanmeldungen am DC' = 'Failed logons at DC'
    'Fehlversuche' = 'Failed attempts'
    'Feststellung' = 'Finding'
    'Fingerabdruck' = 'Fingerprint'
    'Firewall' = 'Firewall'
    'Frei' = 'Free'
    'Freigabe' = 'Share'
    'Freigaberecht' = 'Share permission'
    'Funktion' = 'Function'
    'Gelesen von' = 'Read by'
    'Geprüft' = 'Checked'
    'Gesamtgröße' = 'Total size'
    'Geschützte Benutzer' = 'Protected users'
    'Gesperrt seit' = 'Locked since'
    'Geändert' = 'Changed'
    'Gilt' = 'Applies to'
    'Gilt für' = 'Applies to'
    'Globaler Katalog' = 'Global catalog'
    'Grenze' = 'Limit'
    'Groesse' = 'Size'
    'Grund' = 'Reason'
    'Gruppe' = 'Group'
    'Gruppen' = 'Groups'
    'Gruppen (auch verschachtelt)' = 'Groups (including nested)'
    'Gruppenrichtlinie' = 'Group policy'
    'Größe' = 'Size'
    'Größe (GB)' = 'Size (GB)'
    'Gültig ab' = 'Valid from'
    'Gültig bis' = 'Valid to'
    'Herkunft' = 'Origin'
    'Hersteller' = 'Vendor'
    'Hinweis' = 'Note'
    'Häufigster Grund' = 'Most common reason'
    'IP-Adresse' = 'IP address'
    'Im AD gespeichert' = 'Stored in AD'
    'In der Domäne vorhanden' = 'Present in domain'
    'Infrastrukturmaster' = 'Infrastructure master'
    'Inhalt' = 'Content'
    'Ist' = 'Is'
    'Kennung' = 'ID'
    'Kennwort gesetzt' = 'Password set'
    'Kennwort läuft nie ab' = 'Password never expires'
    'Kerberos' = 'Kerberos'
    'Konten' = 'Accounts'
    'Konten (auch verschachtelt)' = 'Accounts (also nested)'
    'Konto' = 'Account'
    'Konto für DNS-Einträge' = 'Account for DNS entries'
    'Kopie bleibt im Postfach' = 'Copy remains in mailbox'
    'LAPS' = 'LAPS'
    'LAPS-Ablauf' = 'LAPS expiry'
    'Laufwerk' = 'Drive'
    'Lease-Dauer' = 'Lease duration'
    'Leitet weiter an' = 'Forwards to'
    'Leseweg' = 'Read path'
    'Letzte Aktualisierung' = 'Last update'
    'Letzte Anmeldung' = 'Last logon'
    'Letzte Vollsicherung' = 'Last full backup'
    'Letzte inkrementelle Sicherung' = 'Last incremental backup'
    'Letzter' = 'Last'
    'Letzter Aufräumlauf' = 'Last cleanup run'
    'Letzter Start' = 'Last start'
    'Lokale Admins' = 'Local admins'
    'Mindestlänge' = 'Minimum length'
    'Mitglied' = 'Member'
    'Mitglieder' = 'Members'
    'NLA' = 'NLA'
    'NTLM' = 'NTLM'
    'Name' = 'Name'
    'Nr.' = 'No.'
    'Objekt' = 'Object'
    'Ordner' = 'Folder'
    'Organisationseinheit' = 'OU'
    'Ort' = 'Location'
    'PDC-Emulator' = 'PDC emulator'
    'Pfad' = 'Path'
    'Postfach' = 'Mailbox'
    'Postfach-Typ' = 'Mailbox type'
    'Postfachtyp' = 'Mailbox type'
    'Postfächer' = 'Mailboxes'
    'Privilegierte Gruppen' = 'Privileged groups'
    'Programm' = 'Program'
    'Prüfpunkt' = 'Checkpoint'
    'Quelle' = 'Source'
    'Quelle (Anzahl)' = 'Sources (count)'
    'RID-Master' = 'RID master'
    'Recht' = 'Permission'
    'Rechte' = 'Permissions'
    'Regel' = 'Rule'
    'Reihenfolge' = 'Order'
    'SPN' = 'SPN'
    'Schemamaster' = 'Schema master'
    'Schlüssel' = 'Key'
    'Schreibgeschützt' = 'Read-only'
    'Sendesperre ab' = 'Send blocked from'
    'Server' = 'Server'
    'Sid' = 'SID'
    'Signatur' = 'Signature'
    'Speicher' = 'Storage'
    'Sperrungen' = 'Lockouts'
    'Stand' = 'Status'
    'Standort' = 'Site'
    'Status' = 'Status'
    'Support-Ende' = 'End of support'
    'Support-Ende (spätestes)' = 'End of support (latest)'
    'Tage bis Ablauf' = 'Days until expiry'
    'Tage her' = 'Days ago'
    'Tage ohne Anmeldung' = 'Days without logon'
    'Tage seit Ablauf' = 'Days since expiry'
    'Treffer' = 'Hits'
    'Typ' = 'Type'
    'Vererbung' = 'Inheritance'
    'Verschiedene Konten' = 'Distinct accounts'
    'Verschlüsselt' = 'Encrypted'
    'Version' = 'Version'
    'Version C/B' = 'Version C/B'
    'Verweigert für' = 'Denied for'
    'Verwendung' = 'Use'
    'Virenschutz' = 'Antivirus'
    'Von' = 'From'
    'WMI-Filter' = 'WMI filter'
    'Weg' = 'Path'
    'Weiterleitungen' = 'Forwarding'
    'Wert' = 'Value'
    'Wirksame Richtlinie' = 'Effective policy'
    'Wirksames Recht' = 'Effective permission'
    'Wirkung' = 'Effect'
    'Zeit' = 'Time'
    'Zeitpunkt' = 'Time'
    'Zeitraum' = 'Period'
    'Zertifikat' = 'Certificate'
    'Zone' = 'Zone'
    'Zonen' = 'Zones'
    'Zonenübertragung' = 'Zone transfer'
    'Zugriff für große Gruppen' = 'Access for large groups'
    'Zugriffsbasierte Aufzählung' = 'Access-based enumeration'
    'Zuletzt' = 'Last'
    'Zustand' = 'State'
    'davon Größe (GB)' = 'Of which size (GB)'
    'davon aktiv' = 'Of which active'
    'davon inaktiv' = 'Of which inactive'
    'davon unbekannte Konten' = 'Of which unknown accounts'
    'ist Mitglied von' = 'Member of'
    'Älter als 5 Jahre' = 'Older than 5 years'
    'Älter als 5 Jahren' = 'Older than 5 years'
    'Ältester Eintrag' = 'Oldest entry'
    'Über' = 'Via'
    'Überwachungsprotokoll' = 'Audit log'
}
# Werte, die der Code intern auf Deutsch vergleicht (z. B. Art = 'Benutzer'), erscheinen im englischen Bericht übersetzt (nur Anzeige und CSV)
$WerteEn = New-Object System.Collections.Specialized.OrderedDictionary([System.StringComparer]::Ordinal)   # Groß-/Kleinschreibung zählt (Verwaist / verwaist)
$WerteEn.Add('Alle erweiterten Rechte (u. a. Kennwort zurücksetzen)', 'All extended rights (including password reset)')
$WerteEn.Add('Kein Serverbetriebssystem (Support nicht bewertet)', 'No server operating system (support not assessed)')
$WerteEn.Add('Rechte je Ordner (nur dort, wo sie gesetzt werden)', 'Permissions per folder (only where set)')
$WerteEn.Add('Mehrfach freigegebene Ordner (nur einmal gezählt)', 'Folders shared multiple times (counted only once)')
$WerteEn.Add('Übersprungene Verknüpfungen (Junctions/Symlinks)', 'Skipped links (Junctions/Symlinks)')
$WerteEn.Add('Der Ordner der Freigabe existiert nicht mehr.', 'The share folder no longer exists.')
$WerteEn.Add('DHCP-Überwachungsprotokoll ist ausgeschaltet', 'DHCP audit log is disabled')
$WerteEn.Add('Vollzugriff und "Senden als" im Überblick', 'Full access and "Send as" overview')
$WerteEn.Add('Ordner der Freigabe existiert nicht mehr', 'Share folder no longer exists')
$WerteEn.Add('Enthält außer Sicherheitseinstellungen:', 'Contains except security settings:')
$WerteEn.Add('Nicht geprüft - Postfach nicht gefunden', 'Not checked - mailbox not found')
$WerteEn.Add('WMI-Filter nicht lesbar - nicht geprüft', 'WMI filter not readable - not checked')
$WerteEn.Add('Abfrage fehlgeschlagen - nicht geprüft', 'Query failed - not checked')
$WerteEn.Add('Nicht im Active Directory gespeichert', 'Not stored in Active Directory')
$WerteEn.Add('keine im Active Directory autorisiert', 'None authorized in Active Directory')
$WerteEn.Add('Zugriff auf den Ordner nicht möglich', 'Access to folder not possible')
$WerteEn.Add('Zertifikate (öffentliche Schlüssel)', 'Certificates (public keys)')
$WerteEn.Add('Domänencontroller der Organisation', 'Organization domain controllers')
$WerteEn.Add('Remotedesktop erlaubt Verbindungen', 'Remote Desktop allows connections')
$WerteEn.Add('autorisierte DHCP-Server ermitteln', 'Determine authorized DHCP servers')
$WerteEn.Add('Windows Server 2008 R2 und älter', 'Windows Server 2008 R2 and older')
$WerteEn.Add('Konto deaktiviert oder gesperrt', 'Account disabled or locked')
$WerteEn.Add('Vollzugriff (Zusatz: ObjectType', 'Full access (additional: ObjectType')
$WerteEn.Add('Kennwort oder Konto abgelaufen', 'Password or account expired')
$WerteEn.Add('Lokale Administratorengruppen', 'Local administrator groups')
$WerteEn.Add('Berechtigungen/Besitz ändern', 'Change permissions/ownership')
$WerteEn.Add('Freigabepfade nicht sichtbar', 'Share paths not visible')
$WerteEn.Add('Aktive privilegierte Konten', 'Active privileged accounts')
$WerteEn.Add('Gruppe hat keine Mitglieder', 'Group has no members')
$WerteEn.Add('Keine Verbindung zum Server', 'No connection to server')
$WerteEn.Add('unerlaubte Zeichen im Namen', 'Illegal characters in name')
$WerteEn.Add('Geprüfte Domänencontroller', 'Checked domain controllers')
$WerteEn.Add('Authentifizierte Benutzer', 'Authenticated Users')
$WerteEn.Add('Schwellen dieses Berichts', 'Thresholds of this report')
$WerteEn.Add('Benutzer (inetOrgPerson)', 'User (inetOrgPerson)')
$WerteEn.Add('Datenbestand je Freigabe', 'Data inventory per share')
$WerteEn.Add('Kontosperrungsrichtlinie', 'Account lockout policy')
$WerteEn.Add('Sicherheitseinstellungen', 'Security settings')
$WerteEn.Add('Alle erweiterten Rechte', 'All extended rights')
$WerteEn.Add('Hinweise zur Auswertung', 'Notes for evaluation')
$WerteEn.Add('Nicht abgefragte Server', 'Servers not queried')
$WerteEn.Add('Unterordner und Dateien', 'Subfolders and files')
$WerteEn.Add('Bestätigtes Schreiben:', 'Validated write:')
$WerteEn.Add('Installierte Programme', 'Installed programs')
$WerteEn.Add('Konto für DNS-Einträge', 'Account for DNS records')
$WerteEn.Add('Kontosperrungsschwelle', 'Account lockout threshold')
$WerteEn.Add('Nicht lesbare Bereiche', 'Unreadable areas')
$WerteEn.Add('Berechtigungen ändern', 'Change permissions')
$WerteEn.Add('Hinweise zum Auslesen', 'Notes for reading')
$WerteEn.Add('Windows 8.1 und älter', 'Windows 8.1 and older')
$WerteEn.Add('Überwachungsprotokoll', 'Audit log')
$WerteEn.Add('Nicht berücksichtigt', 'Not considered')
$WerteEn.Add('Anonymous-Anmeldung', 'Anonymous logon')
$WerteEn.Add('Infrastrukturmaster', 'Infrastructure master')
$WerteEn.Add('Secure Boot ist aus', 'Secure Boot is off')
$WerteEn.Add('Domänennamenmaster', 'Domain naming master')
$WerteEn.Add('Kennwortrichtlinie', 'Password policy')
$WerteEn.Add('Komponentenwartung', 'Component servicing')
$WerteEn.Add('alle Eigenschaften', 'all properties')
$WerteEn.Add('Besitz übernehmen', 'Take ownership')
$WerteEn.Add('Domänencontroller', 'Domain Controllers')
$WerteEn.Add('Falsches Kennwort', 'Incorrect password')
$WerteEn.Add('Konten-Operatoren', 'Account Operators')
$WerteEn.Add('Nur dieser Ordner', 'Only this folder')
$WerteEn.Add('Domänen-Benutzer', 'Domain Users')
$WerteEn.Add('Format unbekannt', 'Format unknown')
$WerteEn.Add('Sonstige Dateien', 'Other files')
$WerteEn.Add('Ältester Eintrag', 'Oldest entry')
$WerteEn.Add('Domänencomputer', 'Domain Computers')
$WerteEn.Add('Konto unbekannt', 'Account unknown')
$WerteEn.Add('Standard (alle)', 'Default (all)')
$WerteEn.Add('Computerkonten', 'Computer accounts')
$WerteEn.Add('Konto gesperrt', 'Account locked')
$WerteEn.Add('Nicht erreicht', 'Not reached')
$WerteEn.Add('DHCP-Bereiche', 'DHCP scopes')
$WerteEn.Add('Dieser Ordner', 'This folder')
$WerteEn.Add('Gruppe/System', 'Group/system')
$WerteEn.Add('Remotedesktop', 'Remote Desktop')
$WerteEn.Add('Arbeitsplatz', 'Workstation')
$WerteEn.Add('Schemamaster', 'Schema master')
$WerteEn.Add('DHCP-Server', 'DHCP server')
$WerteEn.Add('Datenbanken', 'Databases')
$WerteEn.Add('Sonderkonto', 'Special account')
$WerteEn.Add('Unterordner', 'Subfolders')
$WerteEn.Add('Vollzugriff', 'Full control')
$WerteEn.Add('Zugriff von', 'Access from')
$WerteEn.Add('DNS-Server', 'DNS server')
$WerteEn.Add('Inaktiv ab', 'Inactive from')
$WerteEn.Add('Schreiben:', 'Write:')
$WerteEn.Add('Verweigern', 'Deny')
$WerteEn.Add('DNS-Zonen', 'DNS zones')
$WerteEn.Add('Schreiben', 'Write')
$WerteEn.Add('Sonstiges', 'Other')
$WerteEn.Add('Standorte', 'Sites')
$WerteEn.Add('Unbekannt', 'Unknown')
$WerteEn.Add('unbekannt', 'unknown')
$WerteEn.Add('(Wurzel)', '(root)')
$WerteEn.Add('Anlegen:', 'Create:')
$WerteEn.Add('Benutzer', 'User')
$WerteEn.Add('Erlauben', 'Allow')
$WerteEn.Add('Netzwerk', 'Network')
$WerteEn.Add('Verwaist', 'Orphaned')
$WerteEn.Add('verwaist', 'orphaned')
$WerteEn.Add('Dateien', 'Files')
$WerteEn.Add('Gruppen', 'Groups')
$WerteEn.Add('Anonym', 'Anonymous')
$WerteEn.Add('Domäne', 'Domain')
$WerteEn.Add('Gruppe', 'Group')
$WerteEn.Add('Umfang', 'Scope')
$WerteEn.Add('direkt', 'direct')
$WerteEn.Add('Ändern', 'Change')
$WerteEn.Add('Lesen', 'Read')
$WerteEn.Add('über', 'via')
$WerteMuster = @()
foreach ($k in ($WerteEn.Keys | Sort-Object Length -Descending)) { $WerteMuster += , @([regex]::new('(?<![\p{L}\d])' + [regex]::Escape($k) + '(?![\p{L}\d])'), $WerteEn[$k]) }
function WertAnzeige($v) {
    if (-not $script:En -or $v -isnot [string] -or -not $v) { return $v }
    if ($v.Contains('\')) { return $v }   # Kontonamen (DOMÄNE\Name) und Pfade bleiben, wie sie sind
    foreach ($m in $WerteMuster) { $v = $m[0].Replace($v, $m[1]) }
    $v
}
function SpaltenAnzeige($liste) {
    if (-not $script:En -or -not @($liste).Count) { return $liste }
    foreach ($o in @($liste)) {
        $n = [ordered]@{}
        foreach ($p in $o.PSObject.Properties) { $k = $p.Name; if ($SpaltenEn.ContainsKey($k)) { $k = $SpaltenEn[$k] }; $i = 2; $k0 = $k; while ($n.Contains($k)) { $k = "$k0 $i"; $i++ }; $n[$k] = WertAnzeige $p.Value }
        [pscustomobject]$n
    }
}
$ErrorActionPreference = 'Stop'
$Version = '1.1'

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
        if ($alle.Count) { $sicher | Microsoft.PowerShell.Utility\Export-Csv -Path $Path -NoTypeInformation:$NoTypeInformation -Delimiter $Delimiter -Encoding $Encoding }   # leere Liste: keine Datei (sie hätte keine Kopfzeile)
    }
}

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host (T 'Das PowerShell-Modul "ActiveDirectory" fehlt.' 'The PowerShell module "ActiveDirectory" is missing.') -ForegroundColor Red
    Write-Host (T 'Windows 10/11: Einstellungen > Optionale Features > "RSAT: Active Directory Domain Services" hinzufügen.' 'Windows 10/11: Settings > Optional features > add "RSAT: Active Directory Domain Services".')
    Write-Host (T 'Windows Server:  Install-WindowsFeature RSAT-AD-PowerShell' 'Windows Server: Install-WindowsFeature RSAT-AD-PowerShell')
    exit 2
}
Import-Module ActiveDirectory

$ad = @{}
if ($Server) { $ad.Server = $Server }

# ------------------------------------------------------------------ Hilfsfunktionen
function Datum($d, $leer = $null) { if ($null -eq $leer) { $leer = T 'nie' 'never' }; if ($d) { ([datetime]$d).ToString((T 'dd.MM.yyyy HH:mm' 'yyyy-MM-dd HH:mm')) } else { $leer } }
function JaNein($b) { if ($b) { T 'Ja' 'Yes' } else { T 'Nein' 'No' } }
function Html($t) { [System.Net.WebUtility]::HtmlEncode([string]$t) }
function Kurz($t, $n = 220) { $t = ([string]$t -replace '\s+', ' ').Trim(); if ($t.Length -gt $n) { $t.Substring(0, $n) + ' ...' } else { $t } }
function Zeile($dc, $was, $details = '') { [pscustomobject]@{ 'Domänencontroller' = $dc; Feststellung = $was; Details = $details } }
# Eine fehlgeschlagene Abfrage ist selbst eine Feststellung - sonst sähe der Prüfpunkt fälschlich sauber aus
function Fehlzeile($dc, $fehler) { Zeile $dc (T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked') (Kurz $fehler.Exception.Message) }

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
Write-Host (T 'DC-Gesundheitsbericht - es wird nur gelesen, nichts verändert.' 'DC health report - read-only, nothing is changed.') -ForegroundColor Cyan
Write-Host (T '  Domäne und Domänencontroller ...' 'Domain and domain controllers ...')
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

# --- Erreichbarkeit: LDAP-Port 389
Write-Host (T '  Erreichbarkeit ...' '  Reachability ...')
$e = @()
foreach ($dc in $dcs) {
    $tcp = New-Object System.Net.Sockets.TcpClient
    try {
        $warte = $tcp.BeginConnect($dc.HostName, 389, $null, $null)
        if ($warte.AsyncWaitHandle.WaitOne(4000) -and $tcp.Connected) { $dc.Erreichbar = $true }
    } catch { }
    finally { $tcp.Close() }
    if (-not $dc.Erreichbar) { $e += Zeile $dc.Name (T 'Nicht erreichbar (LDAP, Port 389)' 'Not reachable (LDAP, port 389)') (T "Adresse laut Verzeichnis: $($dc.IP)" "Address according to directory: $($dc.IP)") }
}
$da = @($dcs | Where-Object Erreichbar)

$e = @()
if ($dcs.Count -eq 1) { $e = Zeile $dcs[0].Name 'Einziger Domänencontroller der Domäne' (T 'Kein Ausfallschutz für Anmeldung und DNS' 'No failover protection for logon and DNS') }
Befund 'E2' (T 'Erreichbarkeit' 'Reachability') (T 'Nur ein Domänencontroller vorhanden' 'Only one domain controller present') 'Mittel' $e `
    (T 'Fällt der einzige Domänencontroller aus, kann sich niemand mehr anmelden, und die Namensauflösung steht. Die Wiederherstellung hängt dann allein an der Sicherung.' 'If the only domain controller fails, no one can log on anymore, and name resolution stops. Recovery then depends solely on the backup.') `
    (T 'Einen zweiten Domänencontroller einrichten (auch als virtuelle Maschine auf anderer Hardware) - oder das Risiko bewusst akzeptieren und die Wiederherstellung regelmäßig üben.' 'Set up a second domain controller (also as a virtual machine on different hardware) - or deliberately accept the risk and regularly test recovery.')

# --- Betriebsmasterrollen
$rollen = [ordered]@{
    'PDC-Emulator' = $dom.PDCEmulator; 'RID-Master' = $dom.RIDMaster; 'Infrastrukturmaster' = $dom.InfrastructureMaster
    'Schemamaster' = $wald.SchemaMaster; 'Domänennamenmaster' = $wald.DomainNamingMaster
}
$f = @()
foreach ($r in $rollen.Keys) {
    $inhaber = $dcs | Where-Object { $_.HostName -eq $rollen[$r] }
    if ($inhaber -and -not $inhaber.Erreichbar) { $f += Zeile $inhaber.Name (T "Inhaber der Rolle $r ist nicht erreichbar" "Owner of role $r is not reachable") }
}

# --- SYSVOL
Write-Host (T '  SYSVOL und DFS-Replikation ...' 'SYSVOL and DFS replication ...')
$y1 = @(); $y3 = @()
foreach ($dc in $da) {
    foreach ($frei in 'SYSVOL', 'NETLOGON') {
        try { if (-not (Test-Path -LiteralPath "\\$($dc.HostName)\$frei" -ErrorAction Stop)) { $y1 += Zeile $dc.Name (T "Freigabe $frei fehlt oder ist nicht erreichbar" "Share $frei is missing or not reachable") } }
        catch { $y1 += Fehlzeile $dc.Name $_ }
    }
    try {
        $w = Wmi $dc
        $info = @(Get-CimInstance -Namespace root\MicrosoftDFS -ClassName DfsrReplicatedFolderInfo -Filter "ReplicatedFolderName='SYSVOL Share'" -ErrorAction Stop @w)
        $zustand = @{ 0 = (T 'Nicht initialisiert' 'Not initialized'); 1 = 'Initialisiert'; 2 = (T 'Erstabgleich läuft' 'Initial replication is running'); 3 = (T 'Automatische Wiederherstellung' 'Automatic recovery'); 4 = 'Normal'; 5 = (T 'Fehler' 'Error') }
        if ($info.Count -eq 0) { $y3 += Zeile $dc.Name (T 'Kein Zustand der SYSVOL-Replikation abfragbar - nicht geprüft' 'SYSVOL replication state cannot be queried - not checked') (T 'Läuft der Dienst DFS-Replikation? Bei Abgleich über FRS siehe Y2' 'Is the DFS Replication service running? For replication via FRS, see Y2') }
        foreach ($i in $info) {
            if ($i.State -ne 4) { $y3 += Zeile $dc.Name "SYSVOL-Replikation im Zustand '$($zustand[[int]$i.State])'" (T 'Erwartet: Normal' 'Expected: Normal') }
        }
    } catch { $y3 += Fehlzeile $dc.Name $_ }
}
Befund 'Y1' 'SYSVOL' (T 'Freigabe SYSVOL oder NETLOGON fehlt' 'SYSVOL or NETLOGON share is missing') 'Hoch' $y1 `
    (T 'Über diese Freigaben verteilt ein Domänencontroller Gruppenrichtlinien und Anmeldeskripte. Fehlen sie, gilt der Server nicht als vollwertiger Domänencontroller.' 'A domain controller distributes group policies and logon scripts over these shares. If they are missing, the server is not considered a full-fledged domain controller.') `
    (T 'Zustand der SYSVOL-Replikation prüfen (Ereignisprotokoll "DFS-Replikation"); häufig hat der Erstabgleich nach der Einrichtung nie abgeschlossen.' 'Check the SYSVOL replication state (Event log "DFS Replication"); often the initial synchronization after setup never completed.')

# --- Zeit
Write-Host (T '  Zeit ...' '  Time ...')
$z1 = @(); $z2 = @()
foreach ($dc in ($da | Where-Object { -not $_.Lokal })) {
    try {
        $aus = (& w32tm.exe /stripchart "/computer:$($dc.HostName)" /samples:1 /dataonly 2>&1 | Out-String)
        if ($aus -match ',\s*([+-]\d+[.,]\d+)s') {
            $diff = [double]($Matches[1] -replace ',', '.')
            if ([math]::Abs($diff) -gt 10) { $z1 += Zeile $dc.Name ("Uhrzeit weicht um {0:N0} Sekunden ab" -f $diff) (T "gemessen gegen $ich" "measured against $ich") }
        } else { $z1 += Zeile $dc.Name (T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked') (Kurz $aus) }
    } catch { $z1 += Fehlzeile $dc.Name $_ }
}
Befund 'Z1' (T 'Zeit' 'Time') (T 'Uhrzeit weicht zwischen Domänencontrollern ab' 'Time differs between domain controllers') 'Hoch' $z1 `
    (T 'Die Anmeldung über Kerberos verträgt höchstens fünf Minuten Abweichung. Gemessen wird gegen den Rechner, auf dem das Skript läuft; gemeldet wird ab 10 Sekunden.' 'Kerberos logon tolerates a maximum deviation of five minutes. It is measured against the computer on which the script is running; it is reported from 10 seconds.') `
    (T 'Zeitabgleich prüfen: "w32tm /query /status" auf dem betroffenen Server. Bei virtuellen Maschinen den Zeitabgleich mit dem Host abschalten, wenn die Domäne die Zeit vorgibt.' 'Check time synchronization: "w32tm /query /status" on the affected server. For virtual machines, disable time synchronization with the host if the domain provides the time.')

$pdc = $dcs | Where-Object { $_.HostName -eq $dom.PDCEmulator }
if ($dom.DNSRoot -eq $wald.RootDomain -and $pdc -and $pdc.Erreichbar) {
    try {
        $typ = RegText $pdc 'SYSTEM\CurrentControlSet\Services\W32Time\Parameters' 'Type'
        $quelle = RegText $pdc 'SYSTEM\CurrentControlSet\Services\W32Time\Parameters' 'NtpServer'
        if ($typ -notin 'NTP', 'AllSync') { $z2 = Zeile $pdc.Name (T 'PDC-Emulator hat keine externe Zeitquelle eingestellt' 'PDC emulator has no external time source configured') (T "Einstellung Type: $typ, NtpServer: $quelle" "Setting Type: $typ, NtpServer: $quelle") }
    } catch { $z2 = Fehlzeile $pdc.Name $_ }
} elseif ($dom.DNSRoot -eq $wald.RootDomain) {
    $z2 = Zeile $dom.PDCEmulator (T 'PDC-Emulator nicht erreichbar - Zeitquelle nicht geprüft' 'PDC emulator not reachable - time source not checked')
}

# --- Dienste, Speicher, Updates, Neustart, Betriebssystem
Write-Host (T '  Dienste, Speicherplatz, Updates ...' 'Services, storage, updates ...')
$s1 = @(); $p1 = @(); $u1 = @(); $u2 = @(); $u3 = @()
$pflicht = 'NTDS', 'Netlogon', 'Kdc', 'W32Time', 'DFSR', 'ADWS', 'LanmanServer', 'DNS'
foreach ($dc in $da) {
    $w = $null
    try { $w = Wmi $dc } catch { $fz = Fehlzeile $dc.Name $_; $s1 += $fz; $p1 += $fz; $u1 += $fz; $u2 += $fz; continue }
    try {
        $dienste = @(Get-CimInstance -ClassName Win32_Service -ErrorAction Stop @w | Where-Object { $_.Name -in $pflicht })
        foreach ($d in $dienste) {
            if ($d.StartMode -eq 'Disabled') { if ($d.Name -ne 'DNS') { $s1 += Zeile $dc.Name (T "Dienst $($d.DisplayName) ($($d.Name)) ist deaktiviert" "Service $($d.DisplayName) ($($d.Name)) is disabled") } }
            elseif ($d.State -ne 'Running') { $s1 += Zeile $dc.Name (T "Dienst $($d.DisplayName) ($($d.Name)) läuft nicht" "Service $($d.DisplayName) ($($d.Name)) is not running") (T "Zustand: $($d.State), Startart: $($d.StartMode)" "State: $($d.State), Startup type: $($d.StartMode)") }
        }
        foreach ($n in ($pflicht | Where-Object { $_ -ne 'DNS' -and $_ -notin ($dienste | ForEach-Object { $_.Name }) })) { $s1 += Zeile $dc.Name (T "Dienst $n ist nicht vorhanden" "Service $n is not present") }
    } catch { $s1 += Fehlzeile $dc.Name $_ }
    try {
        foreach ($l in @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop @w)) {
            if ($l.Size -gt 0) {
                $frei = [math]::Round(100 * $l.FreeSpace / $l.Size, 1)
                if ($frei -lt $MinFreiProzent) { $p1 += Zeile $dc.Name ((T "Laufwerk {0} nur noch {1:N1} % frei" "Drive {0} has only {1:N1} % free") -f $l.DeviceID, $frei) ((T "{0:N1} GB frei von {1:N1} GB" "{0:N1} GB free of {1:N1} GB") -f ($l.FreeSpace / 1GB), ($l.Size / 1GB)) }
            }
        }
    } catch { $p1 += Fehlzeile $dc.Name $_ }
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop @w
        $dc.Neustart = $os.LastBootUpTime
        if ($os.LastBootUpTime -lt $jetzt.AddDays(-$UpdateTage)) { $u1 += Zeile $dc.Name "Seit $([int]($jetzt - $os.LastBootUpTime).TotalDays) Tagen nicht neu gestartet" "Letzter Start: $(Datum $os.LastBootUpTime)" }
        $neueste = @(Get-CimInstance -ClassName Win32_QuickFixEngineering -ErrorAction Stop @w | Where-Object { $_.InstalledOn } | Sort-Object InstalledOn -Descending)[0]
        if (-not $neueste) { $u1 += Zeile $dc.Name (T 'Kein installiertes Update mit Datum gefunden' 'No installed update with a date found') }
        elseif ($neueste.InstalledOn -lt $jetzt.AddDays(-$UpdateTage)) { $u1 += Zeile $dc.Name (T "Letztes Update vor $([int]($jetzt - $neueste.InstalledOn).TotalDays) Tagen" "Last update $([int]($jetzt - $neueste.InstalledOn).TotalDays) days ago") (T "$($neueste.HotFixID) am $(([datetime]$neueste.InstalledOn).ToString('dd.MM.yyyy'))" "$($neueste.HotFixID) on $(([datetime]$neueste.InstalledOn).ToString('yyyy-MM-dd'))") }
    } catch { $u1 += Fehlzeile $dc.Name $_ }
    try {
        $gruende = @()
        if (RegSchluesselDa $dc 'SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $gruende += 'Komponentenwartung' }
        if (RegSchluesselDa $dc 'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $gruende += 'Windows Update' }
        if ($gruende.Count) { $u2 += Zeile $dc.Name (T 'Neustart steht aus' 'Restart pending') ($gruende -join ', ') }
    } catch { $u2 += Fehlzeile $dc.Name $_ }
}
foreach ($dc in $dcs) {
    if ($dc.OS -match 'Server.? (2000|2003|2008|2012)') { $u3 += Zeile $dc.Name (T 'Betriebssystem wird nicht mehr unterstützt' 'Operating system is no longer supported') $dc.OS }
}
Befund 'P1' (T 'Speicherplatz' 'Disk space') (T "Laufwerke mit weniger als $MinFreiProzent % freiem Speicher" "Drives with less than $MinFreiProzent % free space") 'Hoch' $p1 `
    (T 'Läuft das Laufwerk mit der Verzeichnisdatenbank oder dem Protokoll voll, bleibt der Verzeichnisdienst stehen.' 'If the drive with the directory database or log fills up, the directory service stops.') `
    (T 'Speicher freigeben (alte Protokolle, Updates bereinigen, Dateien verlagern) oder das Laufwerk vergrößern.' 'Free up disk space (clean up old logs, updates, move files) or expand the drive.')
Befund 'U2' 'Updates' (T 'Neustart steht aus' 'Restart pending') 'Mittel' $u2 `
    (T 'Windows hat Updates installiert, die erst nach einem Neustart wirksam werden.' 'Windows has installed updates that take effect only after a restart.') `
    (T 'Server im nächsten Wartungsfenster neu starten - nie alle Domänencontroller gleichzeitig.' 'Restart the server in the next maintenance window - never restart all domain controllers at the same time.')

# --- Sicherung: das Verzeichnis vermerkt jede anwendungskonsistente Sicherung am Attribut dSASignature
Write-Host (T '  Sicherung ...' '  Backup ...')
$b = @()
$zielDc = @($da | Where-Object { -not $_.RODC })[0]
if ($zielDc) {
    try {
        $meta = Get-ADReplicationAttributeMetadata -Object $dn -Server $zielDc.HostName -Properties dSASignature -ErrorAction Stop
        $letzte = $meta.LastOriginatingChangeTime
        # Version 1 ist der Stand vom Anlegen der Domäne: seitdem wurde nie eine Sicherung vermerkt
        if (-not $letzte -or $meta.Version -le 1) { $b = Zeile $dom.DNSRoot (T 'Im Verzeichnis ist keine Sicherung vermerkt' 'No backup is recorded in the directory') (T "Domäne besteht seit $(Datum $letzte 'unbekannt')" "Domain has existed since $(Datum $letzte 'unbekannt')") }
        elseif ($letzte -lt $jetzt.AddDays(-$SicherungTage)) { $b = Zeile $dom.DNSRoot "Letzte vermerkte Sicherung vor $([int]($jetzt - $letzte).TotalDays) Tagen" (T "am $(Datum $letzte)" "on $(Datum $letzte)") }
    } catch { $b = Fehlzeile $dom.DNSRoot $_ }
}
if (-not $zielDc) { $b = Zeile $dom.DNSRoot (T 'Kein beschreibbarer Domänencontroller erreichbar - nicht geprüft' 'No writable domain controller reachable - not checked') }

# --- dcdiag
$dcdiag = @{}
$t = @()
if (-not $OhneDcdiag) {
    if (Get-Command dcdiag.exe -ErrorAction SilentlyContinue) {
        # dcdiag schreibt in der OEM-Codepage der Konsole; steht die Konsole auf UTF-8 (Wrapper, geplante Aufgabe), würden Umlaute zu Ersatzzeichen
        $alteKodierung = [Console]::OutputEncoding
        try { [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding([System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage) } catch { }
        foreach ($dc in $da) {
            Write-Host (T "  dcdiag für $($dc.Name) (etwa eine Minute) ..." "dcdiag for $($dc.Name) (about one minute) ...")
            try {
                $zeilen = @(& dcdiag.exe "/s:$($dc.HostName)" /q 2>&1 | ForEach-Object { "$_".TrimEnd() } | Where-Object { $_ })
                if ($zeilen.Count) { $dcdiag[$dc.Name] = $zeilen; $t += Zeile $dc.Name (T "dcdiag meldet $($zeilen.Count) Zeilen mit Fehlern oder Warnungen" "dcdiag reports $($zeilen.Count) lines with errors or warnings") (T 'Ausgabe siehe Abschnitt "Ausgabe von dcdiag"' 'See the "dcdiag output" section') }
            } catch { $t += Fehlzeile $dc.Name $_ }
        }
        try { [Console]::OutputEncoding = $alteKodierung } catch { }
    } else { $t = Zeile $ich (T 'dcdiag.exe ist auf diesem Rechner nicht vorhanden - nicht geprüft' 'dcdiag.exe is not present on this computer - not checked') (T 'Auf einem Domänencontroller starten oder die RSAT-Werkzeuge für AD DS installieren' 'Run on a domain controller or install the RSAT tools for AD DS') }
}
if (-not $OhneDcdiag) {
}

foreach ($s in $Sitzungen.Values) { try { Remove-CimSession -CimSession $s } catch { } }

# Besteht ein Prüfpunkt nur aus "nicht geprüft"-Zeilen (Abfrage fehlgeschlagen), ist das kein Mangel dieses Prüfpunkts: als Info führen statt
# mit seiner Stufe, damit es in den Kacheln nicht als hoher Befund zählt.
$platzTexte = @((T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked'), (T 'Kein Zustand der SYSVOL-Replikation abfragbar - nicht geprüft' 'SYSVOL replication state cannot be queried - not checked'), (T 'PDC-Emulator nicht erreichbar - Zeitquelle nicht geprüft' 'PDC emulator not reachable - time source not checked'), (T 'Kein beschreibbarer Domänencontroller erreichbar - nicht geprüft' 'No writable domain controller reachable - not checked'), (T 'dcdiag.exe ist auf diesem Rechner nicht vorhanden - nicht geprüft' 'dcdiag.exe is not present on this computer - not checked'))
foreach ($x in ($Befunde | Where-Object { $_.Anzahl -gt 0 -and $_.Stufe -ne 'Info' })) {
    # nur der feste Text, den dieses Skript selbst erzeugt (keine Teilstring-Suche in Namen oder Beschreibungen)
    if (-not @($x.Objekte | Where-Object { -not (@($_.PSObject.Properties.Value | Where-Object { $_ -is [string] -and $platzTexte -ccontains $_ }).Count -gt 0) }).Count) { $x.Stufe = 'Info' }
}

# ------------------------------------------------------------------ Bericht
Write-Host (T '  Bericht schreiben ...' 'Writing report ...')
if (-not $Ausgabe) { $Ausgabe = Join-Path (Get-Location).Path ("DC-Gesundheitsbericht_{0}_{1:yyyy-MM-dd_HHmm}.html" -f $dom.DNSRoot, $jetzt) }
$Ausgabe = [System.IO.Path]::GetFullPath($Ausgabe)
$ordner = Split-Path $Ausgabe -Parent
if (-not (Test-Path $ordner)) { New-Item -ItemType Directory -Path $ordner -Force | Out-Null }

$stufeText = @{ Hoch = (T 'Hoch' 'High'); Mittel = (T 'Mittel' 'Medium'); Niedrig = (T 'Niedrig' 'Low'); Info = 'Info'; OK = 'OK' }
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
[void]$h.Append((T "<!DOCTYPE html><html lang='$Language'><head><meta charset='utf-8'><title>DC-Gesundheitsbericht Light $(Html $dom.DNSRoot)</title><style>$css</style></head><body><div class='seite'>" "<!DOCTYPE html><html lang='$Language'><head><meta charset='utf-8'><title>DC health report Light $(Html $dom.DNSRoot)</title><style>$css</style></head><body><div class='seite'>"))
[void]$h.Append((T "<h1>DC-Gesundheitsbericht Light</h1><div class='meta'>Domäne <b>$(Html $dom.DNSRoot)</b> · erstellt am $($jetzt.ToString('dd.MM.yyyy HH:mm')) von $(Html $wer) auf $(Html $ich)</div>" "<h1>DC health report Light</h1><div class='meta'>Domain <b>$(Html $dom.DNSRoot)</b> · created on $($jetzt.ToString('dd.MM.yyyy HH:mm')) by $(Html $wer) on $(Html $ich)</div>"))
[void]$h.Append("<div class='kacheln'>")
foreach ($s in 'Hoch', 'Mittel', 'Niedrig', 'OK') {
    $text = @{ Hoch = (T 'Befunde mit hoher Priorität' 'High-priority findings'); Mittel = (T 'mit mittlerer Priorität' 'with medium priority'); Niedrig = (T 'mit niedriger Priorität' 'with low priority'); OK = (T 'Prüfpunkte ohne Befund' 'Checks with no findings') }[$s]
    [void]$h.Append("<div class='kachel $s'><b>$($zahl[$s])</b><span>$text</span></div>")
}
[void]$h.Append("</div>")

[void]$h.Append("<h2>$(T 'Domänencontroller' 'Domain controllers')</h2>")
[void]$h.Append(((SpaltenAnzeige @($dcs | ForEach-Object {
                [pscustomobject][ordered]@{
                    Name = $_.Name; Standort = $_.Standort; 'IP-Adresse' = $_.IP; Betriebssystem = $_.OS
                    'Globaler Katalog' = JaNein $_.GC; 'Schreibgeschützt' = JaNein $_.RODC; Erreichbar = JaNein $_.Erreichbar
                    'Letzter Start' = Datum $_.Neustart (T 'unbekannt' 'unknown')
                }
            }) | ConvertTo-Html -Fragment) -join "`n"))

[void]$h.Append((T "<h2>Umgebung</h2><table><tr><th>Eigenschaft</th><th>Wert</th></tr>" "<h2>Environment</h2><table><tr><th>Property</th><th>Value</th></tr>"))
$umgebung = [ordered]@{
    (T 'Domäne (NetBIOS)' 'Domain (NetBIOS)')                       = "$($dom.DNSRoot) ($($dom.NetBIOSName))"
    (T 'Funktionsebene Domäne / Gesamtstruktur' 'Domain / forest functional level') = "$($dom.DomainMode) / $($wald.ForestMode)"
    'Standorte'                              = ($dcs | ForEach-Object { $_.Standort } | Sort-Object -Unique) -join ', '
}
foreach ($r in $rollen.Keys) { $umgebung[$r] = $rollen[$r] }
$umgebung['Schwellen dieses Berichts'] = (T "Speicher unter $MinFreiProzent % frei, Sicherung älter als $SicherungTage Tage, Updates/Neustart älter als $UpdateTage Tage, Ereignisse der letzten $EreignisStunden Stunden" "Disk space under $MinFreiProzent % free, backup older than $SicherungTage days, updates/restart older than $UpdateTage days, events from the last $EreignisStunden hours")
foreach ($k in $umgebung.Keys) { [void]$h.Append("<tr><td>$(Html (WertAnzeige $k))</td><td>$(Html (WertAnzeige $umgebung[$k]))</td></tr>") }
[void]$h.Append("</table>")

[void]$h.Append((T "<h2>Übersicht der Prüfpunkte</h2><table><tr><th>Nr.</th><th>Bereich</th><th>Prüfpunkt</th><th>Treffer</th><th>Bewertung</th></tr>" "<h2>Overview of checks</h2><table><tr><th>No.</th><th>Area</th><th>Check</th><th>Hits</th><th>Rating</th></tr>"))
foreach ($x in $Befunde) {
    [void]$h.Append("<tr><td>$($x.Id)</td><td>$(Html $x.Bereich)</td><td><a href='#$($x.Id)'>$(Html $x.Titel)</a></td><td>$($x.Anzahl)</td><td><span class='marke m-$($x.Stufe)'>$($stufeText[$x.Stufe])</span></td></tr>")
}
[void]$h.Append("</table>")

[void]$h.Append((T "<h2>Befunde</h2>" "<h2>Findings</h2>"))
foreach ($x in ($Befunde | Where-Object { $_.Stufe -ne 'OK' } | Sort-Object { $rang[$_.Stufe] }, { [regex]::Replace($_.Id, '\d+', '') }, { [int]('0' + [regex]::Match($_.Id, '\d+').Value) })) {
    [void]$h.Append((T "<div class='befund $($x.Stufe)' id='$($x.Id)'><h3><span class='marke m-$($x.Stufe)'>$($stufeText[$x.Stufe])</span> $($x.Id) · $(Html $x.Titel) <span class='klein'>($($x.Anzahl) Treffer)</span></h3>" "<div class='befund $($x.Stufe)' id='$($x.Id)'><h3><span class='marke m-$($x.Stufe)'>$($stufeText[$x.Stufe])</span> $($x.Id) · $(Html $x.Titel) <span class='klein'>($($x.Anzahl) hits)</span></h3>"))
    [void]$h.Append((T "<p>$(Html $x.Erklaerung)</p><p class='emp'><b>Empfehlung:</b> $(Html $x.Empfehlung)</p>" "<p>$(Html $x.Erklaerung)</p><p class='emp'><b>Recommendation:</b> $(Html $x.Empfehlung)</p>"))
    if ($x.Anzahl -gt 0) { [void]$h.Append(((SpaltenAnzeige $x.Objekte | ConvertTo-Html -Fragment) -join "`n")) }
    [void]$h.Append("</div>")
}
if (-not @($Befunde | Where-Object { $_.Stufe -ne 'OK' }).Count) { [void]$h.Append((T "<p>Keine Auffälligkeiten gefunden.</p>" "<p>No anomalies found.</p>")) }
$ok = @($Befunde | Where-Object Stufe -eq 'OK')
if ($ok.Count) {
    [void]$h.Append((T "<h2>Ohne Befund</h2><table><tr><th>Nr.</th><th>Prüfpunkt</th></tr>" "<h2>No findings</h2><table><tr><th>No.</th><th>Check</th></tr>"))
    foreach ($x in $ok) { [void]$h.Append("<tr id='$($x.Id)'><td>$($x.Id)</td><td>$(Html $x.Titel)</td></tr>") }
    [void]$h.Append("</table>")
}
if ($dcdiag.Count) {
    [void]$h.Append((T "<h2>Ausgabe von dcdiag</h2><p class='klein'>Unveränderte Ausgabe im stillen Modus (nur Auffälligkeiten).</p>" "<h2>dcdiag output</h2><p class='klein'>Unchanged output in quiet mode (only anomalies).</p>"))
    foreach ($k in ($dcdiag.Keys | Sort-Object)) { [void]$h.Append("<h3>$(Html $k)</h3><pre>$(Html ($dcdiag[$k] -join "`n"))</pre>") }
}
[void]$h.Append((T "<div class='fuss'>DC-Gesundheitsbericht Light (5 Prüfpunkte, Vollversion: https://mueller86.gumroad.com/l/bseljr) Version $Version · IT Admin Tools by MM · Das Skript hat ausschließlich gelesen und nichts verändert. Der Bericht zeigt Auffälligkeiten anhand fester Regeln zum Zeitpunkt des Laufs; er ersetzt keine laufende Überwachung und erhebt keinen Anspruch auf Vollständigkeit. Dieser Bericht enthält Servernamen und Adressen - bitte vertraulich behandeln.</div>" "<div class='fuss'>DC health report Light (5 checks, full version: https://mueller86.gumroad.com/l/bseljr) version $Version · IT Admin Tools by MM · The script performed read-only operations and made no changes. The report shows anomalies based on fixed rules at the time of the run; it does not replace ongoing monitoring and makes no claim to completeness. This report contains server names and addresses - please treat as confidential.</div>"))
[void]$h.Append("</div></body></html>")
[System.IO.File]::WriteAllText($Ausgabe, $h.ToString(), (New-Object System.Text.UTF8Encoding($true)))

if ($Csv) {
    foreach ($x in ($Befunde | Where-Object { $_.Anzahl -gt 0 })) {
        SpaltenAnzeige $x.Objekte | Export-Csv -Path ([System.IO.Path]::ChangeExtension($Ausgabe, $null).TrimEnd('.') + "_$($x.Id).csv") -NoTypeInformation -Delimiter ';' -Encoding UTF8
    }
}

Write-Host ''
Write-Host ((T "Fertig: {0} hoch, {1} mittel, {2} niedrig, {3} ohne Befund." "Done: {0} high, {1} medium, {2} low, {3} without findings.") -f $zahl.Hoch, $zahl.Mittel, $zahl.Niedrig, $zahl.OK) -ForegroundColor Green
Write-Host (T "Bericht: $Ausgabe" "Report: $Ausgabe")
if (-not $NichtOeffnen -and [Environment]::UserInteractive) { try { Start-Process $Ausgabe } catch { } }
if ($PassThru) { $Befunde }
