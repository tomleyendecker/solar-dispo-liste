<#
    SolarDispo.ps1 - Leyendecker Solar Dispo Liste

    Internes Web-Tool fuer Projektleitung, Montageplanung, Lager und Einkauf von Leyendecker Solar.
    Gleicher Aufbau wie die Oezer-Liste (Lorenz): ein PowerShell-Dienst (HttpListener) haelt als
    einziger die Datenbankverbindung (nur Lesen) und liefert fertiges HTML. Grundsatz: KEINE Preise.

    Daten:
      - Sage 100, Datenbank LeyendeckerSolar, Mandant 1, Login readonly_user (Passwort: Umgebungsvariable SAGE_RO_PW)
      - Montagetermine aus dem Online-Kalender (kalender.digital), iCal-Export. Der Link gibt ohne Anmeldung
        Zugriff auf alle Termine - er steht deshalb NICHT im Skript, sondern in der Umgebungsvariablen
        SOLAR_KALENDER_ICS (wie ein Passwort behandeln).

    Start:  powershell -ExecutionPolicy Bypass -File SolarDispo.ps1        (Port 8082)
    Test:   . .\SolarDispo.ps1   (dot-sourcen laedt nur die Funktionen, startet keinen Dienst)
#>

# ---------- Konfiguration ----------
$Port     = 8082
$Server   = "192.168.200.226\SAGEOL"
$Database = "LeyendeckerSolar"
$DbUser   = "readonly_user"
$Mandant  = 1
function Umgebung($name) {
    foreach ($ziel in 'Process', 'User', 'Machine') { $v = [Environment]::GetEnvironmentVariable($name, $ziel); if ($v) { return $v } }
    return $null
}
$DbPass      = Umgebung 'SAGE_RO_PW'
$KalenderUrl = Umgebung 'SOLAR_KALENDER_ICS'
$KalenderCacheMinuten = 15

# Artikelgruppen (KHKArtikel.Artikelgruppe): Material wird auf Verfuegbarkeit geprueft, Leistungen nur angezeigt
$MaterialGruppen = [ordered]@{
    '50' = 'Module'; '10' = 'Montagesystem'; '13' = 'Montagesystem'; '60' = 'WR &amp; Speicher'
    '64' = 'Wallbox'; '70' = 'Elektro'; '40' = 'Mini-PV'
}
$LeistungsGruppen = @('12', '14', '72', '80', '82', '90')   # Dachmontage, Elektromontage, Geruest, Sonstiges, Foerderung, Zuschlaege

# ---------- Hilfsfunktionen ----------
function Run-Query($sqlText, $params) {
    $conn = New-Object System.Data.SqlClient.SqlConnection "Server=$Server;Database=$Database;User Id=$DbUser;Password=$DbPass;Connection Timeout=15"
    $conn.Open()
    try {
        $cmd = $conn.CreateCommand(); $cmd.CommandText = $sqlText; $cmd.CommandTimeout = 60
        if ($params) { foreach ($key in $params.Keys) { [void]$cmd.Parameters.AddWithValue($key, $params[$key]) } }
        $table = New-Object System.Data.DataTable
        [void](New-Object System.Data.SqlClient.SqlDataAdapter $cmd).Fill($table)
        return ,$table
    } finally { $conn.Close() }
}
function HtmlEncode($s) { if ($null -eq $s) { return "" }; return [System.Net.WebUtility]::HtmlEncode([string]$s) }
function ToNum($v) { if ($null -eq $v -or $v -is [System.DBNull]) { return $null }; try { return [double]$v } catch { return $null } }
function FmtZahl($v) {
    $n = ToNum $v; if ($null -eq $n) { return "" }
    if ([math]::Floor($n) -eq $n) { return ([long]$n).ToString() }
    return ("{0:N2}" -f $n)
}
function StandText { return "Stand " + (Get-Date).ToString("HH:mm") }
function Parse-Query($rawQuery, $name) {
    if ($rawQuery -match "(^|[?&])$name=([^&]*)") { return [System.Uri]::UnescapeDataString($matches[2].Replace('+', ' ')) }
    return $null
}
$WochentagKurz = @('So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa')
function Datum-Kurz($d) { if ($d -isnot [datetime]) { return "" }; return "$($WochentagKurz[[int]$d.DayOfWeek]) $($d.ToString('dd.MM.'))" }
# Vergleichsform fuer Namen: klein, Umlaute ausgeschrieben
function Name-Norm($s) { return (([string]$s).ToLower() -replace [char]0x00e4, 'ae' -replace [char]0x00f6, 'oe' -replace [char]0x00fc, 'ue' -replace [char]0x00df, 'ss') }

# ---------- Montagetermine aus dem Online-Kalender (iCal) ----------
# Termine: Titel = Kundenname + Zusatz ("Juchem, Wilhelm (29 kWp ...)"), Kategorie = Monteur/Team.
# Abruf hoechstens alle $KalenderCacheMinuten Minuten; schlaegt er fehl, bleibt der letzte Stand stehen.
$script:KalCache = $null; $script:KalZeit = [datetime]::MinValue; $script:KalFehler = ""
function Ical-Datum($wert) {
    if ($wert -notmatch '(\d{8})(T(\d{6})(Z?))?') { return $null }
    if (-not $matches[2]) { return [datetime]::ParseExact($matches[1], 'yyyyMMdd', $null) }
    $d = [datetime]::ParseExact($matches[1] + $matches[3], 'yyyyMMddHHmmss', $null)
    if ($matches[4] -eq 'Z') { $d = [datetime]::SpecifyKind($d, 'Utc').ToLocalTime() }
    return $d
}
function Ical-Text($s) { return ([string]$s -replace '\\n', ' ' -replace '\\N', ' ' -replace '\\,', ',' -replace '\\;', ';' -replace '\\\\', '\').Trim() }
function Get-Kalender {
    if ($script:KalCache -and ((Get-Date) - $script:KalZeit).TotalMinutes -lt $KalenderCacheMinuten) { return $script:KalCache }
    if (-not $KalenderUrl) { $script:KalFehler = "Kalender-Link nicht hinterlegt (SOLAR_KALENDER_ICS)"; return @() }
    try {
        $wc = New-Object System.Net.WebClient; $wc.Encoding = [System.Text.Encoding]::UTF8
        $ics = $wc.DownloadString($KalenderUrl)
        $ics = $ics -replace "`r`n[ `t]", "" -replace "`n[ `t]", ""          # Fortsetzungszeilen (RFC 5545)
        $liste = New-Object System.Collections.Generic.List[object]
        foreach ($m in [regex]::Matches($ics, "(?s)BEGIN:VEVENT(.*?)END:VEVENT")) {
            $e = $m.Groups[1].Value
            $feld = { param($n) ([regex]::Match($e, "(?m)^$n(;[^:\r\n]*)?:(.*?)\r?$")).Groups[2].Value }
            $kat = Ical-Text (& $feld 'CATEGORIES')
            if ($kat -eq 'Feiertag') { continue }
            $start = Ical-Datum (& $feld 'DTSTART'); if (-not $start) { continue }
            $ende = Ical-Datum (& $feld 'DTEND')
            $ganztag = ((& $feld 'DTSTART') -notmatch 'T\d{6}')
            if (-not $ende) { $ende = $start }
            if ($ganztag -and $ende -gt $start) { $ende = $ende.AddDays(-1) }     # DTEND ist bei ganztaegigen Terminen exklusiv
            $titel = Ical-Text (& $feld 'SUMMARY')
            $liste.Add([pscustomobject]@{ Start = $start; Ende = $ende; Ganztag = $ganztag; Titel = $titel; Norm = (Name-Norm $titel)
                                         Team = (($kat -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }); Ort = (Ical-Text (& $feld 'LOCATION')) })
        }
        $script:KalCache = $liste.ToArray(); $script:KalZeit = Get-Date; $script:KalFehler = ""
    } catch {
        $script:KalFehler = "Kalender nicht erreichbar: $($_.Exception.Message)"
        if ($script:KalCache) { return $script:KalCache }
        return @()
    }
    return $script:KalCache
}

# Termine zu einem Kunden: Nachname muss im Titel stehen. Steht dort "Nachname, Vorname" mit einem anderen
# Vornamen, ist es eine andere Person. Sicher = auch der Vorname steht im Titel (oder Firma).
$FirmenMuster = '(?i)gmbh|gbr|\bkg\b|\bag\b|e\.v\.|stiftung|verwaltung|beteiligung|\bohg\b|\bug\b'
$NamenStopp = @('gmbh', 'und', 'co', 'kg', 'gbr', 'grundstuecks', 'beteiligungs', 'verwaltungs', 'energie', 'solar', 'herr', 'frau', 'familie', 'eheleute')
function Kunden-Termine($name, $termine) {
    $woerter = @(((Name-Norm $name) -split '[^a-z]+') | Where-Object { $_.Length -ge 3 -and $NamenStopp -notcontains $_ })
    if ($woerter.Count -eq 0) { return @() }
    $firma = ([string]$name -match $FirmenMuster)
    $vorname = if (-not $firma -and $woerter.Count -ge 2) { $woerter[0] } else { $null }
    $nach = if ($vorname) { $woerter[1..($woerter.Count - 1)] } else { $woerter }
    $erg = foreach ($t in $termine) {
        $passt = $false
        foreach ($n in $nach) {
            if ($t.Norm -match "\b$n\b") {
                $passt = $true
                if ($vorname -and $t.Norm -match "\b$n,\s*([a-z]+)" -and $matches[1] -ne $vorname) { $passt = $false }
            }
        }
        if ($passt) { [pscustomobject]@{ T = $t; Sicher = ($firma -or ($vorname -and $t.Norm -match "\b$vorname\b")) } }
    }
    return @($erg)
}

# ---------- Offene Auftraege (AB, Reservierungsbestaetigung, AB aus Reservierung) ----------
# Offen-Logik wie bei Lorenz: Vorgang und Vorgangsposition nicht erfuellt, Restmenge > 0.
$SQL_OffeneAuftraege = @"
SELECT k.BelID, k.Belegart, CAST(k.Belegjahr AS varchar) + '-' + CAST(k.Belegnummer AS varchar) AS Beleg, k.Belegdatum,
       k.A0Name1, k.A0Name2, k.A0Strasse, k.A0PLZ, k.A0Ort, k.A1Name1, k.A1Strasse, k.A1PLZ, k.A1Ort,
       p.Position, p.Artikelnummer, p.Bezeichnung1 AS PosBezeichnung, p.Liefertermin,
       (vp.GGBestellt - vp.GGGeliefert) AS Offen, a.Artikelgruppe, a.Lagerfuehrung, a.Stuecklistentyp, a.Bezeichnung1 AS ArtBezeichnung
FROM KHKVKBelege k
INNER JOIN KHKVKBelegePositionen p ON p.BelID = k.BelID AND p.Mandant = k.Mandant
INNER JOIN KHKVKVorgaengePositionen vp ON vp.VorPosID = p.VorPosID AND vp.Mandant = p.Mandant
INNER JOIN KHKVKVorgaenge v ON v.VorID = vp.VorID AND v.Mandant = vp.Mandant
LEFT JOIN KHKArtikel a ON a.Artikelnummer = p.Artikelnummer AND a.Mandant = p.Mandant
WHERE k.Mandant = $Mandant AND p.PosType = 1
  AND (k.Belegart LIKE 'Auftragsbest%' OR k.Belegart LIKE 'Reservierungsbest%' OR k.Belegart = 'AB (aus Res)')
  AND v.GGErfuellt = 0 AND vp.GGErfuellt = 0 AND (vp.GGBestellt - vp.GGGeliefert) > 0
ORDER BY k.BelID, p.Position
"@
$SQL_Stuecklisten = @"
SELECT s.Stueckliste, s.Element, s.Menge, a.Bezeichnung1 AS Bezeichnung, a.Artikelgruppe, a.Lagerfuehrung
FROM KHKArtikelStueckliste s LEFT JOIN KHKArtikel a ON a.Artikelnummer = s.Element AND a.Mandant = s.Mandant
WHERE s.Mandant = $Mandant AND s.Menge > 0
"@
$SQL_Bestand = "SELECT Artikelnummer, SUM(Lagerbestand) AS Bestand FROM KHKArtikelVarianten WHERE Mandant = $Mandant GROUP BY Artikelnummer"
$SQL_OffeneBestellungen = @"
SELECT p.Artikelnummer, (vp.GGBestellt - vp.GGGeliefert) AS Offen, p.Liefertermin
FROM KHKEKBelege b
INNER JOIN KHKEKBelegePositionen p ON p.BelID = b.BelID AND p.Mandant = b.Mandant
INNER JOIN KHKEKVorgaengePositionen vp ON vp.VorPosID = p.VorPosID AND vp.Mandant = p.Mandant
INNER JOIN KHKEKVorgaenge v ON v.VorID = vp.VorID AND v.Mandant = vp.Mandant
WHERE b.Mandant = $Mandant AND b.Belegart = 'Bestellung' AND v.GGErfuellt = 0 AND vp.GGErfuellt = 0 AND (vp.GGBestellt - vp.GGGeliefert) > 0
"@

# Liest alle offenen Auftraege, ordnet Kalendertermine zu und prueft das Material.
# Rueckgabe: Liste von Auftraegen @{ BelID; Beleg; Art; Kunde; Baustelle; Liefertermin; Termine; Naechster; Material; Leistungen }
function Get-Auftraege {
    $heute = (Get-Date).Date
    $termine = Get-Kalender
    $pos = Run-Query $SQL_OffeneAuftraege $null
    $stl = @{}; foreach ($r in (Run-Query $SQL_Stuecklisten $null).Rows) { $k = [string]$r["Stueckliste"]; if (-not $stl.ContainsKey($k)) { $stl[$k] = New-Object System.Collections.ArrayList }; [void]$stl[$k].Add($r) }
    $bestand = @{}; foreach ($r in (Run-Query $SQL_Bestand $null).Rows) { $bestand[[string]$r["Artikelnummer"]] = [double](ToNum $r["Bestand"]) }
    $ek = @{}; foreach ($r in (Run-Query $SQL_OffeneBestellungen $null).Rows) {
        $a = [string]$r["Artikelnummer"]; if (-not $ek.ContainsKey($a)) { $ek[$a] = New-Object System.Collections.ArrayList }
        [void]$ek[$a].Add(@{ Rest = [double](ToNum $r["Offen"]); Termin = $r["Liefertermin"] })
    }

    $auftraege = [ordered]@{}
    foreach ($r in $pos.Rows) {
        $id = [string]$r["BelID"]
        if (-not $auftraege.Contains($id)) {
            $ba = [string]$r["Belegart"]
            $art = if ($ba -like 'Reservierungs*') { 'RES' } elseif ($ba -eq 'AB (aus Res)') { 'AB/RES' } else { 'AB' }
            $bsName = ([string]$r["A1Name1"]).Trim(); $bsStr = ([string]$r["A1Strasse"]).Trim(); $bsOrt = ("$($r['A1PLZ']) $($r['A1Ort'])").Trim()
            if (-not $bsStr -and -not $bsOrt) { $bsStr = ([string]$r["A0Strasse"]).Trim(); $bsOrt = ("$($r['A0PLZ']) $($r['A0Ort'])").Trim() }
            $kunde = (@(([string]$r["A0Name1"]).Trim(), ([string]$r["A0Name2"]).Trim()) | Where-Object { $_ }) -join ' '
            $auftraege[$id] = @{ BelID = $id; Beleg = [string]$r["Beleg"]; Art = $art; Datum = $r["Belegdatum"]; Kunde = $kunde; Name1 = [string]$r["A0Name1"]
                                 Baustelle = (@($bsStr, $bsOrt) | Where-Object { $_ }) -join ', '; BaustelleName = $bsName
                                 Liefertermin = $null; Material = (New-Object System.Collections.ArrayList); Leistungen = (New-Object System.Collections.ArrayList) }
        }
        $au = $auftraege[$id]
        $lt = $r["Liefertermin"]; if ($lt -is [datetime] -and ($null -eq $au.Liefertermin -or $lt -lt $au.Liefertermin)) { $au.Liefertermin = $lt }
        $an = [string]$r["Artikelnummer"]; $gr = [string]$r["Artikelgruppe"]; $menge = [double](ToNum $r["Offen"])
        $bez = ([string]$r["PosBezeichnung"]).Trim(); if (-not $bez) { $bez = [string]$r["ArtBezeichnung"] }
        if ($LeistungsGruppen -contains $gr -or (-not $MaterialGruppen.Contains($gr) -and (ToNum $r["Lagerfuehrung"]) -ne -1 -and -not $stl.ContainsKey($an))) {
            [void]$au.Leistungen.Add(@{ Art = $an; Bez = $bez; Menge = $menge }); continue
        }
        $gName = if ($MaterialGruppen.Contains($gr)) { $MaterialGruppen[$gr] } else { 'Sonstiges' }
        if ($stl.ContainsKey($an)) {
            # Handelsstueckliste (z.B. Speicher): die lagergefuehrten Bestandteile pruefen
            foreach ($k in $stl[$an]) {
                [void]$au.Material.Add(@{ Gruppe = $gName; Art = [string]$k["Element"]; Bez = [string]$k["Bezeichnung"]; Menge = $menge * [double]$k["Menge"]
                                          Lager = ((ToNum $k["Lagerfuehrung"]) -eq -1); Aus = "$an $bez"; Status = ''; Fehl = 0; EkTermin = $null })
            }
        } else {
            [void]$au.Material.Add(@{ Gruppe = $gName; Art = $an; Bez = $bez; Menge = $menge; Lager = ((ToNum $r["Lagerfuehrung"]) -eq -1); Aus = ''; Status = ''; Fehl = 0; EkTermin = $null })
        }
    }

    # Kalendertermine je Auftrag, naechster Termin ab heute
    foreach ($au in $auftraege.Values) {
        $treffer = @(Kunden-Termine $au.Name1 $termine)
        if ($treffer.Count -eq 0 -and $au.BaustelleName -and $au.BaustelleName -ne $au.Name1) { $treffer = @(Kunden-Termine $au.BaustelleName $termine) }
        $au.Termine = @($treffer | Sort-Object { $_.T.Start })
        $au.Naechster = @($au.Termine | Where-Object { $_.T.Ende.Date -ge $heute }) | Select-Object -First 1
        $au.Planung = if ($au.Naechster) { $au.Naechster.T.Start.Date } elseif ($au.Liefertermin) { $au.Liefertermin.Date } else { [datetime]::MaxValue }
    }

    # Material zuteilen: Bestand und offene Bestellungen gehen an den Auftrag mit dem fruehesten Montage-/Liefertermin
    $rest = @{}; foreach ($k in $bestand.Keys) { $rest[$k] = [math]::Max(0, $bestand[$k]) }
    foreach ($au in @($auftraege.Values | Sort-Object { $_.Planung }, { $_.BelID })) {
        foreach ($m in $au.Material) {
            if (-not $m.Lager) { $m.Status = 'kein'; continue }
            $bedarf = $m.Menge; $vomLager = [math]::Min($bedarf, [double]$rest[$m.Art]); $rest[$m.Art] = [double]$rest[$m.Art] - $vomLager; $bedarf -= $vomLager
            if ($bedarf -le 1e-9) { $m.Status = 'da'; continue }
            if ($ek.ContainsKey($m.Art)) {
                foreach ($b in ($ek[$m.Art] | Sort-Object { if ($_.Termin -is [datetime]) { $_.Termin } else { [datetime]::MaxValue } })) {
                    if ($bedarf -le 1e-9) { break }
                    $nimm = [math]::Min($bedarf, $b.Rest); if ($nimm -le 0) { continue }
                    $b.Rest -= $nimm; $bedarf -= $nimm; $m.EkTermin = $b.Termin
                }
            }
            if ($bedarf -le 1e-9) { $m.Status = 'bestellt' } else { $m.Status = 'fehlt'; $m.Fehl = $bedarf }
        }
    }
    # Anzeige: zuerst Auftraege mit kommendem Montagetermin (nach Termin), dann ohne Termin (nach Liefertermin)
    return @($auftraege.Values | Sort-Object { if ($_.Naechster) { 0 } else { 1 } }, { $_.Planung }, { $_.BelID })
}

# ---------- Seitenrahmen ----------
function Page-Head($titel) {
    return @"
<!DOCTYPE html>
<html lang='de'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>$titel</title>
<style>
* { box-sizing: border-box; }
body { font-family: Segoe UI, Roboto, Arial, sans-serif; margin:0; padding:24px; background:#f4f5f7; color:#1a1a1a; }
.kopf { display:flex; align-items:baseline; justify-content:space-between; flex-wrap:wrap; gap:8px; margin-bottom:16px; }
h1 { font-size:20px; margin:0; }
h2 { font-size:16px; margin:20px 0 10px; }
.meta { font-size:13px; color:#666; }
.steuer { display:flex; gap:12px; flex-wrap:wrap; margin-bottom:12px; align-items:center; }
.suche, .btn { padding:8px 12px; border:1px solid #ccc; border-radius:6px; font-size:14px; }
.suche { width:100%; max-width:420px; }
.btn { background:#2c3e50; color:#fff; text-decoration:none; border:none; cursor:pointer; display:inline-block; }
.btn2 { background:#fff; color:#2c3e50; border:1px solid #c5ccd4; }
table { border-collapse:collapse; width:100%; background:#fff; box-shadow:0 1px 3px rgba(0,0,0,.1); border-radius:8px; overflow:hidden; font-size:13px; }
th, td { padding:7px 10px; text-align:left; border-bottom:1px solid #eee; vertical-align:top; }
th { background:#2c3e50; color:#fff; position:sticky; top:0; white-space:nowrap; }
tr:hover td { background:#f0f7ff; }
.num { text-align:right; font-variant-numeric:tabular-nums; }
.warn { color:#c0392b; font-weight:600; }
.leise { color:#888; font-size:12px; }
.fehler { background:#fdecea; color:#b71c1c; padding:12px 16px; border-radius:8px; margin-bottom:12px; }
.hinweisbox { background:#fff8e1; border:1px solid #ffe082; padding:8px 12px; border-radius:6px; font-size:13px; margin-bottom:12px; }
a.artlink { color:#1565c0; text-decoration:none; font-weight:600; }
a.artlink:hover { text-decoration:underline; }
.hauptnav { display:flex; flex-wrap:wrap; align-items:center; gap:4px 22px; margin:-24px -24px 20px; padding:8px 24px; background:#2c3e50; box-shadow:0 1px 3px rgba(0,0,0,.15); }
.hauptnav .marke { color:#fff; font-weight:700; font-size:16px; text-decoration:none; margin-right:6px; }
.hauptnav a.nav { color:#dfe6ed; text-decoration:none; font-size:13px; padding:5px 9px; border-radius:5px; white-space:nowrap; }
.hauptnav a.nav:hover { background:#34495e; color:#fff; }
.hauptnav a.nav.aktiv { background:#fff; color:#2c3e50; font-weight:600; }
.art { display:inline-block; padding:1px 7px; border-radius:10px; font-size:11px; font-weight:700; background:#e8f4fd; color:#1565c0; white-space:nowrap; }
.art.res { background:#fff3e0; color:#e65100; }
.team { display:inline-block; padding:0 6px; border-radius:8px; background:#2c3e50; color:#fff; font-size:10px; font-weight:700; line-height:16px; margin:1px 2px 1px 0; white-space:nowrap; }
.unsicher { color:#e65100; font-weight:700; cursor:help; }
.chip { display:inline-block; padding:1px 8px; border-radius:10px; font-size:11px; font-weight:600; margin:1px 3px 1px 0; white-space:nowrap; }
.chip.da { background:#d6f0dc; color:#1b5e20; }
.chip.bestellt { background:#e3f2fd; color:#0d47a1; }
.chip.fehlt { background:#fdecea; color:#b71c1c; }
.chip.kein { background:#eceff1; color:#546e7a; }
</style>
</head>
<body>
<nav class='hauptnav'><a class='marke' href='/'>Solar Dispo</a>
  <a class='nav$(if ($pfad -eq '/' -or $pfad -eq '/auftrag') { ' aktiv' })' href='/'>Offene Auftr&auml;ge</a>
</nav>
"@
}

# Status einer Materialgruppe aus ihren Zeilen: fehlt > bestellt > da > kein Lagerartikel
function Gruppen-Chip($name, $zeilen) {
    $fehlt = @($zeilen | Where-Object { $_.Status -eq 'fehlt' }); $best = @($zeilen | Where-Object { $_.Status -eq 'bestellt' })
    $da = @($zeilen | Where-Object { $_.Status -eq 'da' })
    $tip = HtmlEncode ((@($zeilen | ForEach-Object { "$(FmtZahl $_.Menge) x $($_.Bez): $(switch ($_.Status) { 'da' { 'da' } 'bestellt' { 'bestellt' } 'fehlt' { "fehlt $(FmtZahl $_.Fehl)" } default { 'kein Lagerartikel' } })" })) -join "`n")
    if ($fehlt.Count) { return "<span class='chip fehlt' title='$tip'>$name fehlt</span>" }
    if ($best.Count)  { return "<span class='chip bestellt' title='$tip'>$name bestellt</span>" }
    if ($da.Count)    { return "<span class='chip da' title='$tip'>$name &#10003;</span>" }
    return "<span class='chip kein' title='$tip'>$name</span>"
}
function Termin-Html($tr, $mitTitel) {
    $t = $tr.T
    $zeit = if ($t.Ende.Date -gt $t.Start.Date) { "$(Datum-Kurz $t.Start)&ndash;$($t.Ende.ToString('dd.MM.'))" } else { Datum-Kurz $t.Start }
    $teams = ($t.Team | ForEach-Object { "<span class='team'>$(HtmlEncode $_)</span>" }) -join ''
    $u = if (-not $tr.Sicher) { " <span class='unsicher' title='Nur der Nachname passt &ndash; bitte pr&uuml;fen: $(HtmlEncode $t.Titel)'>?</span>" } else { "" }
    $titel = if ($mitTitel) { "<div class='leise'>$(HtmlEncode $t.Titel)</div>" } else { "" }
    return "<div title='$(HtmlEncode $t.Titel)'><b>$zeit</b>$u $teams$titel</div>"
}

function Build-Auftraege($liste, $fehler) {
    $head = Page-Head "Solar Dispo - Offene Auftraege"
    $hinweis = if ($script:KalFehler) { "<div class='hinweisbox'>$(HtmlEncode $script:KalFehler)</div>" } else { "" }
    if ($fehler) { $inhalt = "<div class='fehler'>Fehler beim Laden der Daten:<br><pre>$(HtmlEncode $fehler)</pre></div>" }
    elseif ($liste.Count -eq 0) { $inhalt = "<p class='leise'>Keine offenen Auftr&auml;ge.</p>" }
    else {
        $rows = foreach ($au in $liste) {
            $naechste = @($au.Termine | Where-Object { $_.T.Ende.Date -ge (Get-Date).Date } | Select-Object -First 2)
            $montage = if ($naechste.Count) { ($naechste | ForEach-Object { Termin-Html $_ $false }) -join '' } else { "<span class='leise'>kein Termin</span>" }
            $chips = foreach ($g in @($MaterialGruppen.Values | Select-Object -Unique) + 'Sonstiges') {
                $z = @($au.Material | Where-Object { $_.Gruppe -eq $g }); if ($z.Count) { Gruppen-Chip $g $z }
            }
            $artCls = if ($au.Art -like '*RES*') { 'art res' } else { 'art' }
            $lt = if ($au.Liefertermin) { $au.Liefertermin.ToString('dd.MM.yyyy') } else { "" }
            "<tr><td style='white-space:nowrap'>$montage</td>" +
            "<td style='white-space:nowrap'><span class='$artCls'>$($au.Art)</span> <a class='artlink' href='/auftrag?bel=$($au.BelID)'>$(HtmlEncode $au.Beleg)</a></td>" +
            "<td>$(HtmlEncode $au.Kunde)</td><td>$(HtmlEncode $au.Baustelle)</td><td style='white-space:nowrap'>$lt</td><td>$($chips -join '')</td></tr>"
        }
        $inhalt = @"
<table id='tab'>
<thead><tr><th>Montage</th><th>Auftrag</th><th>Kunde</th><th>Baustelle</th><th>Liefertermin</th><th>Material</th></tr></thead>
<tbody>$($rows -join '')</tbody>
</table>
"@
    }
    return @"
$head
<div class='kopf'><h1>Offene Auftr&auml;ge</h1><div class='meta'>$(StandText) &middot; $($liste.Count) Auftr&auml;ge</div></div>
<div class='steuer'><input class='suche' id='suche' placeholder='Filtern ...' onkeyup='filter()'></div>
$hinweis
$inhalt
<script>
function filter(){
  var q=document.getElementById('suche').value.toLowerCase();
  document.querySelectorAll('#tab tbody tr').forEach(function(tr){ tr.style.display = tr.innerText.toLowerCase().indexOf(q)>-1 ? '' : 'none'; });
}
</script>
</body></html>
"@
}

function Build-AuftragDetail($au, $fehler) {
    $head = Page-Head "Solar Dispo - Auftrag"
    if ($fehler) { return "$head<div class='fehler'>Fehler beim Laden der Daten:<br><pre>$(HtmlEncode $fehler)</pre></div></body></html>" }
    if (-not $au) { return "$head<p class='leise'>Auftrag nicht gefunden oder nicht mehr offen.</p></body></html>" }
    $termine = if ($au.Termine.Count) {
        "<table><thead><tr><th>Termin</th><th>Team</th><th>Kalender</th></tr></thead><tbody>" + (($au.Termine | ForEach-Object {
            $t = $_.T; $zeit = if ($t.Ende.Date -gt $t.Start.Date) { "$(Datum-Kurz $t.Start) &ndash; $(Datum-Kurz $t.Ende) $($t.Ende.Year)" } else { "$(Datum-Kurz $t.Start) $($t.Start.Year)" }
            if (-not $t.Ganztag) { $zeit += " $($t.Start.ToString('HH:mm'))" }
            $vergangen = if ($t.Ende.Date -lt (Get-Date).Date) { " style='color:#999'" } else { "" }
            $u = if (-not $_.Sicher) { " <span class='unsicher' title='Nur der Nachname passt &ndash; bitte pr&uuml;fen'>?</span>" } else { "" }
            "<tr$vergangen><td style='white-space:nowrap'>$zeit$u</td><td>$(($t.Team | ForEach-Object { "<span class='team'>$(HtmlEncode $_)</span>" }) -join '')</td><td>$(HtmlEncode $t.Titel)</td></tr>"
        }) -join '') + "</tbody></table>"
    } else { "<p class='leise'>Im Kalender kein Termin zu diesem Kunden gefunden.</p>" }
    $material = if ($au.Material.Count) {
        "<table><thead><tr><th>Gruppe</th><th>Artikel</th><th>Bezeichnung</th><th class='num'>Menge</th><th>Status</th></tr></thead><tbody>" + (($au.Material | ForEach-Object {
            $st = switch ($_.Status) {
                'da' { "<span class='chip da'>da</span>" }
                'bestellt' { "<span class='chip bestellt'>bestellt$(if ($_.EkTermin -is [datetime]) { ' &middot; ' + $_.EkTermin.ToString('dd.MM.') })</span>" }
                'fehlt' { "<span class='chip fehlt'>fehlt $(FmtZahl $_.Fehl)</span>" }
                default { "<span class='chip kein'>kein Lagerartikel</span>" }
            }
            $aus = if ($_.Aus) { "<div class='leise'>aus $(HtmlEncode $_.Aus)</div>" } else { "" }
            "<tr><td>$($_.Gruppe)</td><td>$(HtmlEncode $_.Art)</td><td>$(HtmlEncode $_.Bez)$aus</td><td class='num'>$(FmtZahl $_.Menge)</td><td>$st</td></tr>"
        }) -join '') + "</tbody></table>"
    } else { "<p class='leise'>Kein Material im Auftrag.</p>" }
    $leistungen = if ($au.Leistungen.Count) {
        "<table><thead><tr><th>Leistung</th><th class='num'>Menge</th></tr></thead><tbody>" + (($au.Leistungen | ForEach-Object { "<tr><td>$(HtmlEncode $_.Bez)</td><td class='num'>$(FmtZahl $_.Menge)</td></tr>" }) -join '') + "</tbody></table>"
    } else { "<p class='leise'>Keine Leistungen.</p>" }
    $artCls = if ($au.Art -like '*RES*') { 'art res' } else { 'art' }
    return @"
$head
<div class='kopf'><h1><span class='$artCls'>$($au.Art)</span> $(HtmlEncode $au.Beleg) &ndash; $(HtmlEncode $au.Kunde)</h1><div class='meta'>$(StandText)</div></div>
<div class='steuer'><a class='btn btn2' href='/'>&larr; Offene Auftr&auml;ge</a></div>
<table style='max-width:620px;margin-bottom:8px'><tbody>
<tr><td>Kunde</td><td>$(HtmlEncode $au.Kunde)</td></tr>
<tr><td>Baustelle</td><td>$(HtmlEncode (@($au.BaustelleName, $au.Baustelle) | Where-Object { $_ }) -join ', ')</td></tr>
<tr><td>Auftragsdatum</td><td>$(if ($au.Datum -is [datetime]) { $au.Datum.ToString('dd.MM.yyyy') })</td></tr>
<tr><td>Liefertermin</td><td>$(if ($au.Liefertermin) { $au.Liefertermin.ToString('dd.MM.yyyy') })</td></tr>
</tbody></table>
<h2>Montagetermine</h2>
$termine
<h2>Material</h2>
$material
<h2>Leistungen</h2>
$leistungen
</body></html>
"@
}

# ---------- Dienst ----------
if ($MyInvocation.InvocationName -eq '.') { return }   # dot-gesourct (Test): nur Funktionen laden

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:$Port/")
try { $listener.Start() }
catch {
    Write-Host "`nFEHLER beim Start auf Port $Port (Berechtigung?)." -ForegroundColor Red
    Write-Host "Einmalig als Administrator:  netsh http add urlacl url=http://+:$Port/ user=Jeder" -ForegroundColor Yellow
    Write-Host $_.Exception.Message
    exit 1
}
Write-Host ("=" * 60)
Write-Host "  Leyendecker Solar Dispo Liste laeuft auf Port $Port"
Write-Host ("=" * 60)
if (-not $DbPass) { Write-Host "  ACHTUNG: SAGE_RO_PW nicht gesetzt." -ForegroundColor Yellow }
if (-not $KalenderUrl) { Write-Host "  ACHTUNG: SOLAR_KALENDER_ICS nicht gesetzt - keine Montagetermine." -ForegroundColor Yellow }

while ($listener.IsListening) {
    try {
        $context = $listener.GetContext()
        $request = $context.Request; $response = $context.Response
        $pfad = $request.Url.AbsolutePath.ToLower(); $query = $request.Url.Query
        $ausgabe = ""
        if ($pfad -eq "/") {
            try { $ausgabe = Build-Auftraege (Get-Auftraege) $null }
            catch { $ausgabe = Build-Auftraege @() ($_.Exception.ToString()) }
        } elseif ($pfad -eq "/auftrag") {
            $bel = ([string](Parse-Query $query "bel")) -replace '[^0-9]', ''
            try { $ausgabe = Build-AuftragDetail (@(Get-Auftraege) | Where-Object { $_.BelID -eq $bel } | Select-Object -First 1) $null }
            catch { $ausgabe = Build-AuftragDetail $null ($_.Exception.ToString()) }
        } else {
            $response.StatusCode = 404; $ausgabe = "Nicht gefunden"
        }
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($ausgabe)
        $response.ContentType = "text/html; charset=utf-8"
        $response.ContentLength64 = $buffer.Length
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
        $response.OutputStream.Close()
    } catch {
        Write-Host "Fehler bei Anfrage: $($_.Exception.Message)" -ForegroundColor Red
    }
}
