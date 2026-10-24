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
$StartVariablen = @(Get-Variable -Scope Script | ForEach-Object { $_.Name })   # alles danach bekommt auch der Chat-Runspace
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

# ---- Assistent (Chat) ----
# API-Schluessel als Umgebungsvariable ANTHROPIC_API_KEY (Benutzer oder Maschine) auf dem Server,
# wird bei jeder Anfrage frisch gelesen. Ohne Schluessel wird das Chatfenster nicht angezeigt.
$ChatModell    = "claude-sonnet-5-5"
$ChatMaxRunden = 8      # max. Werkzeug-Runden je Frage
$ChatProtokollPfad = Join-Path $PSScriptRoot "chat-protokoll.jsonl"    # je Frage eine JSON-Zeile (ohne Antwort)
$ChatBewertungPfad = Join-Path $PSScriptRoot "chat-bewertungen.jsonl"  # Daumen hoch/runter inkl. Frage, Antwort, Kommentar
# $false = Testbetrieb: Chat nur in Browsern, die einmal /chat-an aufgerufen haben (Cookie). $true = fuer alle.
$ChatFuerAlle  = $true

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
function Name-Norm($s) { return (Name-Umlaut $s).ToLower() }
# Umlaute ausschreiben, Gross-/Kleinschreibung behalten (fuer die Vornamen-Pruefung)
function Name-Umlaut($s) { return ([string]$s -creplace [char]0x00e4, 'ae' -creplace [char]0x00f6, 'oe' -creplace [char]0x00fc, 'ue' -creplace [char]0x00c4, 'Ae' -creplace [char]0x00d6, 'Oe' -creplace [char]0x00dc, 'Ue' -replace [char]0x00df, 'ss') }

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
            $liste.Add([pscustomobject]@{ Start = $start; Ende = $ende; Ganztag = $ganztag; Titel = $titel; Norm = (Name-Norm $titel); Fall = (Name-Umlaut $titel)
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

# Termine zu einem Kunden: Nachname (letztes Wort, bei 'Nachname, Vorname' das erste) muss im Titel stehen. Steht dort "Nachname, Vorname" mit einem anderen grossgeschriebenen
# Vornamen, ist es eine andere Person. Sicher = auch der Vorname steht im Titel (oder Firma).
$FirmenMuster = '(?i)gmbh|gbr|\bkg\b|\bag\b|e\.v\.|stiftung|verwaltung|beteiligung|\bohg\b|\bug\b'
$NamenStopp = @('gmbh', 'und', 'co', 'kg', 'gbr', 'grundstuecks', 'beteiligungs', 'verwaltungs', 'energie', 'solar', 'herr', 'frau', 'familie', 'eheleute')
function Kunden-Termine($name, $termine) {
    $woerter = @(((Name-Norm $name) -split '[^a-z]+') | Where-Object { $_.Length -ge 3 -and $NamenStopp -notcontains $_ })
    if ($woerter.Count -eq 0) { return @() }
    $firma = ([string]$name -match $FirmenMuster)
    # Person: Nachname = letztes Wort ("Jenny und Julian Alloway" -> alloway), bei "Nachname, Vorname" das erste
    $vorname = $null; $nach = $woerter
    if (-not $firma -and $woerter.Count -ge 2) {
        if ([string]$name -match ',') { $nach = @($woerter[0]); $vorname = $woerter[1] }
        else { $nach = @($woerter[-1]); $vorname = $woerter[0] }
    }
    $erg = foreach ($t in $termine) {
        $passt = $false
        foreach ($n in $nach) {
            if ($t.Norm -match "\b$n\b") {
                $passt = $true
                # "Nachname, Vorname" mit anderem, grossgeschriebenem Vornamen = andere Person ("Schmitz, Bodo");
                # kleingeschrieben ist es Text ("Halama, nach Dachsanierung")
                if ($vorname -and $t.Fall -match "(?i)\b$n,\s*([A-Za-z]+)" -and [char]::IsUpper($matches[1][0]) -and $matches[1].ToLower() -ne $vorname) { $passt = $false }
            }
        }
        if ($passt) { [pscustomobject]@{ T = $t; Sicher = ($firma -or ($vorname -and $t.Norm -match "\b$vorname\b")) } }
    }
    return @($erg)
}

# Modulleistung in Watt aus der Bezeichnung einer Modulposition (Artikelgruppe 50), z.B. "LUXOR Eco Line M108 465W"
function Modul-Wp($gruppe, $bez) {
    if ([string]$gruppe -ne '50') { return $null }
    if ([string]$bez -match '(\d{3})\s*Wp?\b') { return [int]$matches[1] }
    return $null
}
function Fmt-kWp($k) { if (-not $k) { return "" }; return ([math]::Round($k, 1)).ToString('0.#', [Globalization.CultureInfo]'de-DE') + " kWp" }

# ---------- Offene Auftraege (AB, Reservierungsbestaetigung, AB aus Reservierung) ----------
# Offen-Logik wie bei Lorenz: Vorgang und Vorgangsposition nicht erfuellt, Restmenge > 0.
# Sage zaehlt Reservierung und die daraus erzeugte "AB (aus Res)" beide in GGBestellt - die Reservierungsmenge
# wird dann abgezogen und die Reservierung selbst nicht mehr gelistet (die AB ersetzt sie).
$SQL_ResAbzug = @"
OUTER APPLY (
    SELECT SUM(rp.MengeBasis) AS Menge FROM KHKVKBelegePositionen rp
    INNER JOIN KHKVKBelege rk ON rk.BelID = rp.BelID AND rk.Mandant = rp.Mandant
    WHERE rp.Mandant = p.Mandant AND rp.VorPosID = p.VorPosID AND rk.Belegart LIKE 'Reservierungsbest%'
      AND EXISTS (SELECT 1 FROM KHKVKBelegePositionen ap INNER JOIN KHKVKBelege ak ON ak.BelID = ap.BelID AND ak.Mandant = ap.Mandant
                  WHERE ap.Mandant = p.Mandant AND ap.VorPosID = p.VorPosID AND ak.Belegart = 'AB (aus Res)')
) res
CROSS APPLY (SELECT vp.GGBestellt - ISNULL(res.Menge, 0) AS Menge) soll
"@
$SQL_OffeneAuftraege = @"
SELECT k.BelID, k.VorID, k.A0Empfaenger, k.Belegart, CAST(k.Belegjahr AS varchar) + '-' + CAST(k.Belegnummer AS varchar) AS Beleg, k.Belegdatum,
       k.A0Name1, k.A0Name2, k.A0Strasse, k.A0PLZ, k.A0Ort, k.A1Name1, k.A1Strasse, k.A1PLZ, k.A1Ort,
       p.BelPosID, p.Position, p.Artikelnummer, p.Bezeichnung1 AS PosBezeichnung, p.Liefertermin,
       (soll.Menge - vp.GGGeliefert) AS Offen, a.Artikelgruppe, a.Lagerfuehrung, a.Stuecklistentyp, a.Bezeichnung1 AS ArtBezeichnung
FROM KHKVKBelege k
INNER JOIN KHKVKBelegePositionen p ON p.BelID = k.BelID AND p.Mandant = k.Mandant
INNER JOIN KHKVKVorgaengePositionen vp ON vp.VorPosID = p.VorPosID AND vp.Mandant = p.Mandant
INNER JOIN KHKVKVorgaenge v ON v.VorID = vp.VorID AND v.Mandant = vp.Mandant
LEFT JOIN KHKArtikel a ON a.Artikelnummer = p.Artikelnummer AND a.Mandant = p.Mandant
$SQL_ResAbzug
WHERE k.Mandant = $Mandant AND p.PosType = 1
  AND (k.Belegart LIKE 'Auftragsbest%' OR k.Belegart LIKE 'Reservierungsbest%' OR k.Belegart = 'AB (aus Res)')
  AND NOT (k.Belegart LIKE 'Reservierungsbest%' AND res.Menge IS NOT NULL)
  AND v.GGErfuellt = 0 AND vp.GGErfuellt = 0 AND (soll.Menge - vp.GGGeliefert) > 0
ORDER BY k.BelID, p.Position
"@
# Liefer-/Rechnungsstand je Auftrag ueber alle Positionen (auch die schon erledigten); {BELIDS} = Liste der BelIDs
$SQL_AuftragStand = @"
SELECT p.BelID, COUNT(*) AS Pos,
       SUM(CASE WHEN vp.GGGeliefert >= soll.Menge - 0.0001 THEN 1 ELSE 0 END) AS Gel, SUM(CASE WHEN vp.GGGeliefert > 0 THEN 1 ELSE 0 END) AS GelAnf,
       SUM(CASE WHEN vp.GGBerechnet >= soll.Menge - 0.0001 THEN 1 ELSE 0 END) AS Ber, SUM(CASE WHEN vp.GGBerechnet > 0 THEN 1 ELSE 0 END) AS BerAnf
FROM KHKVKBelegePositionen p
INNER JOIN KHKVKVorgaengePositionen vp ON vp.VorPosID = p.VorPosID AND vp.Mandant = p.Mandant
$SQL_ResAbzug
WHERE p.Mandant = $Mandant AND p.PosType = 1 AND p.BelID IN ({BELIDS}) AND soll.Menge > 0
GROUP BY p.BelID
"@
# Alle Positionen eines Auftrags (auch schon gelieferte) fuer die Detailseite
$SQL_AuftragPositionen = @"
SELECT p.BelPosID, p.Artikelnummer, p.Bezeichnung1 AS PosBezeichnung, a.Bezeichnung1 AS ArtBezeichnung, a.Artikelgruppe, a.Lagerfuehrung,
       soll.Menge AS Soll, vp.GGGeliefert AS Geliefert, vp.GGBerechnet AS Berechnet
FROM KHKVKBelegePositionen p
INNER JOIN KHKVKVorgaengePositionen vp ON vp.VorPosID = p.VorPosID AND vp.Mandant = p.Mandant
LEFT JOIN KHKArtikel a ON a.Artikelnummer = p.Artikelnummer AND a.Mandant = p.Mandant
$SQL_ResAbzug
WHERE p.Mandant = $Mandant AND p.BelID = @bel AND p.PosType = 1 AND soll.Menge > 0
ORDER BY p.BelPosID
"@
# Lieferscheine und Rechnungen der Vorgaenge (fuer die Detailseite); {VORIDS} = Liste der VorIDs
$SQL_AuftragFolgebelege = @"
SELECT VorID, Belegart, CAST(Belegjahr AS varchar) + '-' + CAST(Belegnummer AS varchar) AS Beleg, Belegdatum
FROM KHKVKBelege
WHERE Mandant = $Mandant AND VorID IN ({VORIDS}) AND (Belegart LIKE '%lieferschein' OR Belegart LIKE '%rechnung' OR Belegart = 'Gutschrift')
ORDER BY Belegdatum, BelID
"@
$SQL_Stuecklisten = @"
SELECT s.Stueckliste, s.Element, s.Menge, a.Bezeichnung1 AS Bezeichnung, a.Artikelgruppe, a.Lagerfuehrung
FROM KHKArtikelStueckliste s LEFT JOIN KHKArtikel a ON a.Artikelnummer = s.Element AND a.Mandant = s.Mandant
WHERE s.Mandant = $Mandant AND s.Menge > 0
"@
$SQL_Bestand = "SELECT Artikelnummer, SUM(Lagerbestand) AS Bestand FROM KHKArtikelVarianten WHERE Mandant = $Mandant GROUP BY Artikelnummer"
$SQL_OffeneBestellungen = @"
SELECT b.BelID, CAST(b.Belegjahr AS varchar) + '-' + CAST(b.Belegnummer AS varchar) AS Beleg, b.A0Name1 AS Lieferant, b.Belegdatum,
       p.BelPosID, p.Artikelnummer, p.Bezeichnung1 AS Bezeichnung, vp.GGBestellt AS Bestellt, (vp.GGBestellt - vp.GGGeliefert) AS Offen, p.Liefertermin
FROM KHKEKBelege b
INNER JOIN KHKEKBelegePositionen p ON p.BelID = b.BelID AND p.Mandant = b.Mandant
INNER JOIN KHKEKVorgaengePositionen vp ON vp.VorPosID = p.VorPosID AND vp.Mandant = p.Mandant
INNER JOIN KHKEKVorgaenge v ON v.VorID = vp.VorID AND v.Mandant = vp.Mandant
WHERE b.Mandant = $Mandant AND b.Belegart = 'Bestellung' AND v.GGErfuellt = 0 AND vp.GGErfuellt = 0 AND (vp.GGBestellt - vp.GGGeliefert) > 0
ORDER BY p.Liefertermin, b.BelID, p.BelPosID
"@

# Liest alle offenen Auftraege, ordnet Kalendertermine zu und prueft das Material.
# Rueckgabe: Liste von Auftraegen @{ BelID; Beleg; Art; Kunde; Baustelle; Liefertermin; Termine; Naechster; Material; Leistungen }
function Get-Auftraege {
    $heute = (Get-Date).Date
    $termine = Get-Kalender
    $pos = Run-Query $SQL_OffeneAuftraege $null
    $stl = @{}; foreach ($r in (Run-Query $SQL_Stuecklisten $null).Rows) { $k = [string]$r["Stueckliste"]; if (-not $stl.ContainsKey($k)) { $stl[$k] = New-Object System.Collections.ArrayList }; [void]$stl[$k].Add($r) }
    $bestand = @{}; foreach ($r in (Run-Query $SQL_Bestand $null).Rows) { $bestand[[string]$r["Artikelnummer"]] = [double](ToNum $r["Bestand"]) }
    # offene Bestellpositionen; Fuer = Auftraege, denen die Position zugeteilt wird (Einkaufsseite)
    $ek = @{}; $script:EkPositionen = New-Object System.Collections.ArrayList
    foreach ($r in (Run-Query $SQL_OffeneBestellungen $null).Rows) {
        $a = [string]$r["Artikelnummer"]; if (-not $ek.ContainsKey($a)) { $ek[$a] = New-Object System.Collections.ArrayList }
        $e = @{ Rest = [double](ToNum $r["Offen"]); Termin = $r["Liefertermin"]; Row = $r; Fuer = (New-Object System.Collections.ArrayList) }
        [void]$ek[$a].Add($e); [void]$script:EkPositionen.Add($e)
    }
    $script:Bestand = $bestand; $script:EkSumme = @{}; $script:EkTermin = @{}
    foreach ($a in $ek.Keys) {
        $script:EkSumme[$a] = ($ek[$a] | ForEach-Object { $_.Rest } | Measure-Object -Sum).Sum
        $t = @($ek[$a] | ForEach-Object { $_.Termin } | Where-Object { $_ -is [datetime] } | Sort-Object); if ($t.Count) { $script:EkTermin[$a] = $t[0] }
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
            $auftraege[$id] = @{ BelID = $id; VorID = [string]$r["VorID"]; Kto = [string]$r["A0Empfaenger"]; Plz = @(([string]$r["A1PLZ"]).Trim(), ([string]$r["A0PLZ"]).Trim()); StrasseRoh = $bsStr; Cn = $null; Stand = $null; Folgebelege = @(); Beleg = [string]$r["Beleg"]; Art = $art; Datum = $r["Belegdatum"]; Kunde = $kunde; Name1 = [string]$r["A0Name1"]
                                 Baustelle = (@($bsStr, $bsOrt) | Where-Object { $_ }) -join ', '; BaustelleName = $bsName
                                 Liefertermin = $null; kWp = 0.0; Material = (New-Object System.Collections.ArrayList); Leistungen = (New-Object System.Collections.ArrayList) }
        }
        $au = $auftraege[$id]
        $lt = $r["Liefertermin"]; if ($lt -is [datetime] -and ($null -eq $au.Liefertermin -or $lt -lt $au.Liefertermin)) { $au.Liefertermin = $lt }
        $an = [string]$r["Artikelnummer"]; $gr = [string]$r["Artikelgruppe"]; $menge = [double](ToNum $r["Offen"])
        $bez = ([string]$r["PosBezeichnung"]).Trim(); if (-not $bez) { $bez = [string]$r["ArtBezeichnung"] }
        if ($LeistungsGruppen -contains $gr -or (-not $MaterialGruppen.Contains($gr) -and (ToNum $r["Lagerfuehrung"]) -ne -1 -and -not $stl.ContainsKey($an))) {
            [void]$au.Leistungen.Add(@{ Art = $an; Bez = $bez; Menge = $menge }); continue
        }
        $gName = if ($MaterialGruppen.Contains($gr)) { $MaterialGruppen[$gr] } else { 'Sonstiges' }
        $wp = Modul-Wp $gr $bez; if ($wp) { $au.kWp += $menge * $wp / 1000 }
        if ($stl.ContainsKey($an)) {
            # Handelsstueckliste (z.B. Speicher): die lagergefuehrten Bestandteile pruefen
            foreach ($k in $stl[$an]) {
                [void]$au.Material.Add(@{ Gruppe = $gName; Art = [string]$k["Element"]; Bez = [string]$k["Bezeichnung"]; Menge = $menge * [double]$k["Menge"]; KitArt = $an; KitMenge = $menge; KitPos = [string]$r["Position"]
                                          Lager = ((ToNum $k["Lagerfuehrung"]) -eq -1); Aus = "$an $bez"; Status = ''; Fehl = 0; EkTermin = $null; PosID = [string]$r["BelPosID"] })
            }
        } else {
            [void]$au.Material.Add(@{ Gruppe = $gName; Art = $an; Bez = $bez; Menge = $menge; Lager = ((ToNum $r["Lagerfuehrung"]) -eq -1); Aus = ''; Status = ''; Fehl = 0; EkTermin = $null; PosID = [string]$r["BelPosID"] })
        }
    }

    # Liefer-/Rechnungsstand und Folgebelege (Lieferscheine, Rechnungen) je Auftrag
    if ($auftraege.Count) {
        foreach ($r in (Run-Query ($SQL_AuftragStand -replace '\{BELIDS\}', (@($auftraege.Keys) -join ',')) $null).Rows) {
            $auftraege[[string]$r["BelID"]].Stand = @{ Pos = [int]$r["Pos"]; Gel = [int]$r["Gel"]; GelAnf = [int]$r["GelAnf"]; Ber = [int]$r["Ber"]; BerAnf = [int]$r["BerAnf"] }
        }
        $vorIds = @($auftraege.Values | ForEach-Object { $_.VorID } | Where-Object { $_ -match '^[1-9][0-9]*$' } | Select-Object -Unique)
        if ($vorIds.Count) {
            $folge = @((Run-Query ($SQL_AuftragFolgebelege -replace '\{VORIDS\}', ($vorIds -join ',')) $null).Rows)
            foreach ($au in $auftraege.Values) { $au.Folgebelege = @($folge | Where-Object { [string]$_["VorID"] -eq $au.VorID }) }
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
                    $b.Rest -= $nimm; $bedarf -= $nimm; $m.EkTermin = $b.Termin; [void]$b.Fuer.Add(@{ Au = $au; Menge = $nimm })
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
<link rel='icon' href='/static/logo.svg'>
<style>
/* Farben und Schrift wie leyendecker.solar: Gelb #FFBC00, Anthrazit #343A40, Hellgrau #F8F9FA, Barlow */
@font-face { font-family:'Barlow'; font-weight:400; font-display:swap; src:url('/static/barlow-regular.woff2') format('woff2'); }
@font-face { font-family:'Barlow'; font-weight:500; font-display:swap; src:url('/static/barlow-500.woff2') format('woff2'); }
@font-face { font-family:'Barlow'; font-weight:600; font-display:swap; src:url('/static/barlow-600.woff2') format('woff2'); }
@font-face { font-family:'Barlow'; font-weight:700; font-display:swap; src:url('/static/barlow-700.woff2') format('woff2'); }
* { box-sizing: border-box; }
body { font-family: Barlow, Segoe UI, Arial, sans-serif; font-size:15px; margin:0; padding:24px; background:#F8F9FA; color:#343A40; }
.kopf { display:flex; align-items:baseline; justify-content:space-between; flex-wrap:wrap; gap:8px; margin-bottom:16px; }
h1 { font-size:26px; font-weight:700; margin:0; text-transform:uppercase; letter-spacing:.3px; color:#343A40; }
h2 { font-size:18px; font-weight:700; margin:22px 0 10px; }
.meta { font-size:14px; color:#5B5B5A; }
.steuer { display:flex; gap:10px; flex-wrap:wrap; margin-bottom:14px; align-items:center; }
.suche, .btn { padding:8px 16px; border:1px solid #d6d8db; border-radius:999px; font-family:inherit; font-size:14px; }
.suche { width:100%; max-width:420px; background:#fff; color:#343A40; }
.suche:focus { outline:none; border-color:#FFBC00; box-shadow:0 0 0 3px rgba(255,188,0,.25); }
.btn { background:#343A40; color:#fff; font-weight:600; text-decoration:none; border:1px solid #343A40; cursor:pointer; display:inline-block; }
.btn:hover { background:#23272b; }
.btn2 { background:#fff; color:#343A40; border:1px solid #c9ccd0; font-weight:500; }
.btn2:hover { background:#fff; border-color:#FFBC00; }
table { border-collapse:collapse; width:100%; background:#fff; box-shadow:0 1px 3px rgba(0,0,0,.08); border-radius:10px; overflow:hidden; font-size:14px; }
th, td { padding:8px 11px; text-align:left; border-bottom:1px solid #eceef0; vertical-align:top; }
th { background:#343A40; color:#fff; font-weight:600; position:sticky; top:0; white-space:nowrap; }
tr:hover td { background:#fff8e0; }
.num { text-align:right; font-variant-numeric:tabular-nums; }
.warn { color:#c0392b; font-weight:600; }
.leise { color:#868e96; font-size:13px; }
.fehler { background:#fdecea; color:#b71c1c; padding:12px 16px; border-radius:10px; margin-bottom:12px; }
.hinweisbox { background:#fff8e0; border:1px solid #ffe08a; padding:8px 12px; border-radius:10px; font-size:14px; margin-bottom:12px; }
a.artlink { color:#086ED5; text-decoration:none; font-weight:600; }
a.artlink:hover { text-decoration:underline; }
/* Kopf: weiss mit Logo, darunter die gelbe Doppellinie der Webseite */
.hauptnav { display:flex; flex-wrap:wrap; align-items:center; gap:4px 26px; margin:-24px -24px 26px; padding:10px 24px; background:#fff;
            border-bottom:1px solid #FFBC00; box-shadow:0 3px 0 #fff, 0 4px 0 #FFBC00; }
.hauptnav .marke { display:flex; align-items:center; gap:12px; text-decoration:none; margin-right:10px; }
.hauptnav .marke img { height:42px; display:block; }
.hauptnav .marke span { font-size:13px; font-weight:600; letter-spacing:2px; text-transform:uppercase; color:#5B5B5A; border-left:1px solid #d6d8db; padding-left:12px; }
.hauptnav a.nav { color:#343A40; text-decoration:none; font-size:15px; font-weight:600; padding:6px 2px; border-bottom:3px solid transparent; white-space:nowrap; }
.hauptnav a.nav:hover { border-bottom-color:#ffe08a; }
.hauptnav a.nav.aktiv { border-bottom-color:#FFBC00; }
.art { display:inline-block; padding:1px 8px; border-radius:10px; font-size:12px; font-weight:700; background:#e7f1fc; color:#086ED5; white-space:nowrap; }
.art.res { background:#fff3d6; color:#b26a00; }
.team { display:inline-block; padding:0 7px; border-radius:8px; background:#343A40; color:#fff; font-size:11px; font-weight:600; line-height:17px; margin:1px 2px 1px 0; white-space:nowrap; }
.unsicher { color:#e65100; font-weight:700; cursor:help; }
.chip { display:inline-block; padding:1px 9px; border-radius:10px; font-size:12px; font-weight:600; margin:1px 3px 1px 0; white-space:nowrap; }
.chip.da { background:#d6f0dc; color:#1b5e20; }
.chip.bestellt { background:#e3f2fd; color:#0d47a1; }
.chip.fehlt { background:#fdecea; color:#b71c1c; }
.chip.kein { background:#eceff1; color:#546e7a; }
tr.erledigt td, tr.erledigt td a.artlink { color:#9aa0a6; }
a.cn { display:inline-block; padding:1px 9px; border-radius:10px; font-size:12px; font-weight:600; text-decoration:none; white-space:nowrap; background:#eceff1; color:#455a64; }
a.cn:hover { text-decoration:underline; }
a.cn-auftrag { background:#e7f1fc; color:#086ED5; }
a.cn-foerderung { background:#fff3d6; color:#b26a00; }
a.cn-inbetrieb, a.cn-fertig { background:#d6f0dc; color:#1b5e20; }
a.cn-absage { background:#fdecea; color:#b71c1c; }
a.cn-service { background:#f3e5f5; color:#6a1b9a; }
details summary { cursor:pointer; }
</style>
</head>
<body>
<nav class='hauptnav'><a class='marke' href='/'><img src='/static/logo.svg' alt='Leyendecker Solar'><span>Dispo</span></a>
  <a class='nav$(if ($pfad -eq '/' -or $pfad -eq '/auftrag') { ' aktiv' })' href='/'>Offene Auftr&auml;ge</a>
  <a class='nav$(if ($pfad -eq '/material') { ' aktiv' })' href='/material'>Materialbedarf</a>
  <a class='nav$(if ($pfad -eq '/angebote') { ' aktiv' })' href='/angebote'>Angebote</a>
  <a class='nav$(if ($pfad -eq '/einkauf' -or $pfad -eq '/ek') { ' aktiv' })' href='/einkauf'>Einkauf</a>
  <a class='nav$(if ($pfad -eq '/dispo') { ' aktiv' })' href='/dispo'>Dispo</a>
  <a class='nav$(if ($pfad -eq '/bewegungen') { ' aktiv' })' href='/bewegungen'>Bewegungen</a>
  <a class='nav$(if ($pfad -eq '/abgleich') { ' aktiv' })' href='/abgleich'>Abgleich</a>
</nav>
$(if (Test-ChatErlaubt) { $ChatWidgetHtml })
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

# Liefer- bzw. Rechnungsstand eines Auftrags: - (nichts), x/y Positionen, Haken (alles)
function Stand-Html($stand, $was) {
    if (-not $stand -or $stand.Pos -eq 0) { return "" }
    $voll = if ($was -eq 'Gel') { $stand.Gel } else { $stand.Ber }
    $anf  = if ($was -eq 'Gel') { $stand.GelAnf } else { $stand.BerAnf }
    $wort = if ($was -eq 'Gel') { 'geliefert' } else { 'berechnet' }
    if ($anf -eq 0) { return "<span class='leise' title='nichts $wort'>&ndash;</span>" }
    if ($voll -ge $stand.Pos) { return "<span class='chip da' title='alle Positionen $wort'>&#10003;</span>" }
    return "<span title='$voll von $($stand.Pos) Positionen vollst&auml;ndig $wort'>$voll/$($stand.Pos)</span>"
}
function Build-Auftraege($liste, $fehler) {
    $head = Page-Head "Solar Dispo - Offene Auftraege"
    $hinweis = (@(if ($script:KalFehler) { "<div class='hinweisbox'>$(HtmlEncode $script:KalFehler)</div>" }; if ($script:CnFehler) { "<div class='hinweisbox'>$($script:CnFehler)</div>" }) -join '')
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
            "<td>$(HtmlEncode $au.Kunde)</td><td>$(HtmlEncode $au.Baustelle)</td><td class='num' style='white-space:nowrap'>$(Fmt-kWp $au.kWp)</td><td style='white-space:nowrap'>$lt</td>" +
            "<td style='text-align:center'>$(Stand-Html $au.Stand 'Gel')</td><td style='text-align:center'>$(Stand-Html $au.Stand 'Ber')</td><td>$(CN-Chip $au.Cn)</td><td>$($chips -join '')</td></tr>"
        }
        $inhalt = @"
<table id='tab'>
<thead><tr><th>Montage</th><th>Auftrag</th><th>Kunde</th><th>Baustelle</th><th class='num'>Anlage</th><th>Liefertermin</th><th>Geliefert</th><th>Berechnet</th><th>Craftnote</th><th>Material</th></tr></thead>
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

# Positionen fuer die Auftragsseite: Material (Stuecklisten aufgeloest) und Leistungen, je mit Soll, geliefert, berechnet
# und dem Status aus der Zuteilung (nur fuer den noch offenen Teil)
function Get-AuftragPositionen($au) {
    $t = Run-Query $SQL_AuftragPositionen @{ "@bel" = [int]$au.BelID }
    $stl = @{}; foreach ($r in (Run-Query $SQL_Stuecklisten $null).Rows) { $k = [string]$r["Stueckliste"]; if (-not $stl.ContainsKey($k)) { $stl[$k] = New-Object System.Collections.ArrayList }; [void]$stl[$k].Add($r) }
    $mat = New-Object System.Collections.ArrayList; $lei = New-Object System.Collections.ArrayList
    foreach ($r in $t.Rows) {
        $an = [string]$r["Artikelnummer"]; $gr = [string]$r["Artikelgruppe"]; $posId = [string]$r["BelPosID"]
        $bez = ([string]$r["PosBezeichnung"]).Trim(); if (-not $bez) { $bez = [string]$r["ArtBezeichnung"] }
        $soll = [double](ToNum $r["Soll"]); $gel = [double](ToNum $r["Geliefert"]); $ber = [double](ToNum $r["Berechnet"])
        if ($LeistungsGruppen -contains $gr -or (-not $MaterialGruppen.Contains($gr) -and (ToNum $r["Lagerfuehrung"]) -ne -1 -and -not $stl.ContainsKey($an))) {
            [void]$lei.Add(@{ Bez = $bez; Soll = $soll; Gel = $gel; Ber = $ber }); continue
        }
        $gName = if ($MaterialGruppen.Contains($gr)) { $MaterialGruppen[$gr] } else { 'Sonstiges' }
        $teile = if ($stl.ContainsKey($an)) { @($stl[$an] | ForEach-Object { @{ Art = [string]$_["Element"]; Bez = [string]$_["Bezeichnung"]; F = [double]$_["Menge"]; Lager = ((ToNum $_["Lagerfuehrung"]) -eq -1); Aus = "$an $bez" } }) }
                 else { @(@{ Art = $an; Bez = $bez; F = 1.0; Lager = ((ToNum $r["Lagerfuehrung"]) -eq -1); Aus = '' }) }
        foreach ($x in $teile) {
            $offen = @($au.Material | Where-Object { $_.PosID -eq $posId -and $_.Art -eq $x.Art }) | Select-Object -First 1
            [void]$mat.Add(@{ Gruppe = $gName; Art = $x.Art; Bez = $x.Bez; Aus = $x.Aus; Lager = $x.Lager; Soll = $soll * $x.F; Gel = $gel * $x.F; Ber = $ber * $x.F; Offen = $offen })
        }
    }
    return @{ Material = $mat; Leistungen = $lei }
}
# Zelle "geliefert" / "berechnet": - (nichts), Teilmenge, Haken (vollstaendig)
function Mengen-Zelle($ist, $soll, $wort) {
    if ($ist -le 1e-9) { return "<td class='num leise'>&ndash;</td>" }
    if ($ist -ge $soll - 1e-9) { return "<td class='num'><span class='chip da' title='vollst&auml;ndig $wort'>&#10003;</span></td>" }
    return "<td class='num' title='$(FmtZahl $ist) von $(FmtZahl $soll) $wort'>$(FmtZahl $ist)</td>"
}
function Build-AuftragDetail($au, $pos, $fehler) {
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
    if (-not $pos) { $pos = @{ Material = @(); Leistungen = @() } }
    $material = if ($pos.Material.Count) {
        "<table><thead><tr><th>Gruppe</th><th>Artikel</th><th>Bezeichnung</th><th class='num'>Menge</th><th class='num'>Geliefert</th><th class='num'>Berechnet</th><th>Status</th></tr></thead><tbody>" + (($pos.Material | ForEach-Object {
            $fertig = $_.Gel -ge $_.Soll - 1e-9
            $st = ""
            if (-not $fertig) {
                $o = $_.Offen
                $st = if (-not $_.Lager) { "<span class='chip kein'>kein Lagerartikel</span>" }
                      elseif (-not $o) { "" }
                      else { switch ($o.Status) {
                          'da' { "<span class='chip da'>da</span>" }
                          'bestellt' { "<span class='chip bestellt'>bestellt$(if ($o.EkTermin -is [datetime]) { ' &middot; ' + $o.EkTermin.ToString('dd.MM.') })</span>" }
                          'fehlt' { "<span class='chip fehlt'>fehlt $(FmtZahl $o.Fehl)</span>" }
                          default { "<span class='chip kein'>kein Lagerartikel</span>" } } }
            }
            $aus = if ($_.Aus) { "<div class='leise'>aus $(HtmlEncode $_.Aus)</div>" } else { "" }
            "<tr$(if ($fertig) { " class='erledigt'" })><td>$($_.Gruppe)</td><td><a class='artlink' href='/dispo?art=$(HtmlEncode $_.Art)'>$(HtmlEncode $_.Art)</a></td><td>$(HtmlEncode $_.Bez)$aus</td>" +
            "<td class='num'>$(FmtZahl $_.Soll)</td>$(Mengen-Zelle $_.Gel $_.Soll 'geliefert')$(Mengen-Zelle $_.Ber $_.Soll 'berechnet')<td>$st</td></tr>"
        }) -join '') + "</tbody></table>"
    } else { "<p class='leise'>Kein Material im Auftrag.</p>" }
    $leistungen = if ($pos.Leistungen.Count) {
        "<table><thead><tr><th>Leistung</th><th class='num'>Menge</th><th class='num'>Geliefert</th><th class='num'>Berechnet</th></tr></thead><tbody>" + (($pos.Leistungen | ForEach-Object {
            "<tr$(if ($_.Gel -ge $_.Soll - 1e-9) { " class='erledigt'" })><td>$(HtmlEncode $_.Bez)</td><td class='num'>$(FmtZahl $_.Soll)</td>$(Mengen-Zelle $_.Gel $_.Soll 'geliefert')$(Mengen-Zelle $_.Ber $_.Soll 'berechnet')</tr>"
        }) -join '') + "</tbody></table>"
    } else { "<p class='leise'>Keine Leistungen.</p>" }
    $artCls = if ($au.Art -like '*RES*') { 'art res' } else { 'art' }
    $folge = { param($muster)
        (@($au.Folgebelege | Where-Object { [string]$_["Belegart"] -like $muster } | ForEach-Object {
            "$(HtmlEncode $_['Belegart']) $(HtmlEncode $_['Beleg'])$(if ($_['Belegdatum'] -is [datetime]) { ' (' + $_['Belegdatum'].ToString('dd.MM.yyyy') + ')' })"
        })) -join ', ' }
    $standZeile = { param($was, $muster)
        $txt = ""
        if ($au.Stand -and $au.Stand.Pos) {
            $voll = if ($was -eq 'Gel') { $au.Stand.Gel } else { $au.Stand.Ber }
            $anf  = if ($was -eq 'Gel') { $au.Stand.GelAnf } else { $au.Stand.BerAnf }
            $txt = if ($anf -eq 0) { "nichts" } elseif ($voll -ge $au.Stand.Pos) { "<span class='chip da'>&#10003;</span> alle Positionen" } else { "$voll von $($au.Stand.Pos) Positionen" }
        }
        $bel = & $folge $muster
        "$txt$(if ($bel) { "<div class='leise'>$bel</div>" })" }
    return @"
$head
<div class='kopf'><h1><span class='$artCls'>$($au.Art)</span> $(HtmlEncode $au.Beleg) &ndash; $(HtmlEncode $au.Kunde)</h1><div class='meta'>$(StandText)</div></div>
<div class='steuer'><a class='btn btn2' href='/'>&larr; Offene Auftr&auml;ge</a>$(if ($au.Cn) { " <a class='btn btn2' href='$CraftnoteWeb$($au.Cn.P.Id)' target='_blank'>Craftnote &#8599;</a>" })</div>
<table style='max-width:620px;margin-bottom:8px'><tbody>
<tr><td>Kunde</td><td>$(HtmlEncode $au.Kunde)</td></tr>
<tr><td>Baustelle</td><td>$(HtmlEncode (@($au.BaustelleName, $au.Baustelle) | Where-Object { $_ }) -join ', ')</td></tr>
<tr><td>Auftragsdatum</td><td>$(if ($au.Datum -is [datetime]) { $au.Datum.ToString('dd.MM.yyyy') })</td></tr>
<tr><td>Liefertermin</td><td>$(if ($au.Liefertermin) { $au.Liefertermin.ToString('dd.MM.yyyy') })</td></tr>
<tr><td>Anlage</td><td>$(Fmt-kWp $au.kWp)</td></tr>
<tr><td>Geliefert</td><td>$(& $standZeile 'Gel' '*lieferschein')</td></tr>
<tr><td>Berechnet</td><td>$(& $standZeile 'Ber' '*rechnung')</td></tr>
</tbody></table>
<h2>Montagetermine</h2>
$termine
<h2>Craftnote</h2>
$(CN-DetailHtml $au.Cn)
<h2>Material</h2>
$material
<h2>Leistungen</h2>
$leistungen
</body></html>
"@
}

# ---------- Materialbedarf (Lager / Einkauf) ----------
# Je Lagerartikel: Bedarf aus allen offenen Auftraegen, Bestand, offene Bestellungen und der erste Termin, an dem
# es fehlt. Gleiche Zuteilung wie auf "Offene Auftraege" (Auftraege nach Montage-/Liefertermin). Dazu die Positionen
# ohne Lagerfuehrung (z.B. Montagesystem LORENZ, Sonderteile), die je Auftrag beschafft werden muessen.
function Build-Material($liste, $nur, $tage, $fehler) {
    $head = Page-Head "Solar Dispo - Materialbedarf"
    $heute = (Get-Date).Date
    $bis = if ($tage -gt 0) { $heute.AddDays($tage) } else { [datetime]::MaxValue }
    $inhalt = ""
    if ($fehler) { $inhalt = "<div class='fehler'>Fehler beim Laden der Daten:<br><pre>$(HtmlEncode $fehler)</pre></div>" }
    else {
        $art = [ordered]@{}; $ohneLager = New-Object System.Collections.ArrayList
        foreach ($au in $liste) {
            if ($au.Planung -gt $bis) { continue }
            foreach ($m in $au.Material) {
                if (-not $m.Lager) {
                    if ($m.Gruppe -ne 'Elektro') { [void]$ohneLager.Add(@{ Au = $au; M = $m }) }
                    continue
                }
                if (-not $art.Contains($m.Art)) { $art[$m.Art] = @{ Art = $m.Art; Bez = $m.Bez; Gruppe = $m.Gruppe; Bedarf = 0.0; Fehl = 0.0; Bestellt = $false; Erster = $null; ErsterAu = $null; Zeilen = (New-Object System.Collections.ArrayList) } }
                $a = $art[$m.Art]; $a.Bedarf += $m.Menge
                if ($m.Status -eq 'fehlt') { $a.Fehl += $m.Fehl; if (-not $a.Erster) { $a.Erster = $au.Planung; $a.ErsterAu = $au } }
                if ($m.Status -eq 'bestellt') { $a.Bestellt = $true }
                [void]$a.Zeilen.Add(@{ Au = $au; M = $m })
            }
        }
        $alle = @($art.Values)
        $anzFehlt = @($alle | Where-Object { $_.Fehl -gt 0 }).Count
        $zeige = if ($nur -eq 'alle') { $alle } else { @($alle | Where-Object { $_.Fehl -gt 0 }) }
        $zeige = @($zeige | Sort-Object { if ($_.Fehl -gt 0) { 0 } elseif ($_.Bestellt) { 1 } else { 2 } }, { if ($_.Erster) { $_.Erster } else { [datetime]::MaxValue } }, { $_.Gruppe }, { $_.Art })
        $rows = foreach ($a in $zeige) {
            $st = if ($a.Fehl -gt 0) { "<span class='chip fehlt'>fehlt $(FmtZahl $a.Fehl)</span>" } elseif ($a.Bestellt) { "<span class='chip bestellt'>bestellt</span>" } else { "<span class='chip da'>da</span>" }
            $best = ToNum $script:Bestand[$a.Art]; if ($null -eq $best) { $best = 0 }
            $ekM = ToNum $script:EkSumme[$a.Art]
            $ekTxt = if ($ekM) { "$(FmtZahl $ekM)$(if ($script:EkTermin[$a.Art]) { " <span class='leise'>ab $($script:EkTermin[$a.Art].ToString('dd.MM.'))</span>" })" } else { "<span class='leise'>&ndash;</span>" }
            $erster = if ($a.ErsterAu) { "$(Datum-Kurz $a.Erster) &middot; <a class='artlink' href='/auftrag?bel=$($a.ErsterAu.BelID)'>$(HtmlEncode $a.ErsterAu.Kunde)</a>" } else { "" }
            $tip = HtmlEncode ((@($a.Zeilen | ForEach-Object {
                $d = if ($_.Au.Planung -ne [datetime]::MaxValue) { $_.Au.Planung.ToString('dd.MM.') } else { 'ohne Termin' }
                "$d  $($_.Au.Beleg)  $($_.Au.Kunde): $(FmtZahl $_.M.Menge)$(if ($_.M.Status -eq 'fehlt') { ' (fehlt ' + (FmtZahl $_.M.Fehl) + ')' })"
            })) -join "`n")
            "<tr><td>$st</td><td>$($a.Gruppe)</td><td><a class='artlink' href='/dispo?art=$(HtmlEncode $a.Art)'>$(HtmlEncode $a.Art)</a></td><td>$(HtmlEncode $a.Bez)</td>" +
            "<td class='num'>$(FmtZahl $a.Bedarf)</td><td class='num'>$(FmtZahl $best)</td><td class='num' style='white-space:nowrap'>$ekTxt</td>" +
            "<td style='white-space:nowrap'>$erster</td><td class='num' title='$tip' style='cursor:help'>$($a.Zeilen.Count)</td></tr>"
        }
        $tab1 = if ($zeige.Count) {
            "<table class='filtertab'><thead><tr><th>Status</th><th>Gruppe</th><th>Artikel</th><th>Bezeichnung</th><th class='num'>Bedarf</th><th class='num'>Bestand</th><th class='num'>bestellt</th><th>fehlt ab</th><th class='num' title='Anzahl Auftr&auml;ge; Maus dar&uuml;ber zeigt sie'>Auftr.</th></tr></thead><tbody>$($rows -join '')</tbody></table>"
        } else { "<p class='leise'>Nichts fehlt in diesem Zeitraum.</p>" }
        # Positionen ohne Lagerfuehrung je Auftrag
        $rows2 = foreach ($x in ($ohneLager | Sort-Object { $_.Au.Planung }, { $_.Au.BelID })) {
            $d = if ($x.Au.Planung -ne [datetime]::MaxValue) { Datum-Kurz $x.Au.Planung } else { "<span class='leise'>ohne Termin</span>" }
            "<tr><td style='white-space:nowrap'>$d</td><td><a class='artlink' href='/auftrag?bel=$($x.Au.BelID)'>$(HtmlEncode $x.Au.Beleg)</a></td><td>$(HtmlEncode $x.Au.Kunde)</td>" +
            "<td>$($x.M.Gruppe)</td><td><a class='artlink' href='/dispo?art=$(HtmlEncode $x.M.Art)'>$(HtmlEncode $x.M.Art)</a></td><td>$(HtmlEncode $x.M.Bez)</td><td class='num'>$(FmtZahl $x.M.Menge)</td></tr>"
        }
        $tab2 = if ($ohneLager.Count) {
            "<h2>Ohne Lagerf&uuml;hrung &ndash; je Auftrag zu beschaffen</h2><table class='filtertab'><thead><tr><th>Termin</th><th>Auftrag</th><th>Kunde</th><th>Gruppe</th><th>Artikel</th><th>Bezeichnung</th><th class='num'>Menge</th></tr></thead><tbody>$($rows2 -join '')</tbody></table>"
        } else { "" }
        $inhalt = "$tab1`n$tab2"
    }
    $knopf = { param($n, $t, $text) $cls = if ($n -eq $nur -and $t -eq $tage) { 'btn' } else { 'btn btn2' }; "<a class='$cls' href='/material?nur=$n&amp;tage=$t'>$text</a>" }
    $zeitKnoepfe = (@(@(14, '2 Wochen'), @(28, '4 Wochen'), @(0, 'Alle Termine')) | ForEach-Object { & $knopf $nur $_[0] $_[1] }) -join ' '
    $nurKnoepfe = (& $knopf 'fehlt' $tage "Nur Fehlendes ($anzFehlt)") + ' ' + (& $knopf 'alle' $tage 'Alle Artikel')
    return @"
$head
<div class='kopf'><h1>Materialbedarf</h1><div class='meta'>$(StandText)</div></div>
<div class='steuer'>$zeitKnoepfe <span style='width:12px'></span> $nurKnoepfe <input class='suche' id='suche' placeholder='Filtern ...' onkeyup='filter()'></div>
$(if ($script:KalFehler) { "<div class='hinweisbox'>$(HtmlEncode $script:KalFehler)</div>" })
$inhalt
<script>
function filter(){
  var q=document.getElementById('suche').value.toLowerCase();
  document.querySelectorAll('.filtertab tbody tr').forEach(function(tr){ tr.style.display = tr.innerText.toLowerCase().indexOf(q)>-1 ? '' : 'none'; });
}
</script>
</body></html>
"@
}

# ---------- Offene Angebote (Projektleitung) ----------
# In dieser Datenbank bekommt ein Angebot erst mit der Beauftragung einen Vorgang: offen = VorID 0.
# Hat der Kunde danach einen Auftrag (z.B. ueber eine andere Angebotsvariante), wird das Angebot ausgegraut.
$SQL_OffeneAngebote = @"
SELECT k.BelID, CAST(k.Belegjahr AS varchar) + '-' + CAST(k.Belegnummer AS varchar) AS Beleg, k.Belegdatum,
       k.A0Name1, k.A0Name2, k.A0Ort, k.A1Ort, k.A1Strasse, k.A0Empfaenger, k.A0PLZ, k.A1PLZ, k.A0Strasse,
       p.Artikelnummer, p.Bezeichnung1 AS PosBezeichnung, p.Menge, a.Artikelgruppe, a.Stuecklistentyp,
       ab.Beleg AS AuftragDanach
FROM KHKVKBelege k
LEFT JOIN KHKVKBelegePositionen p ON p.BelID = k.BelID AND p.Mandant = k.Mandant AND p.PosType = 1
LEFT JOIN KHKArtikel a ON a.Artikelnummer = p.Artikelnummer AND a.Mandant = p.Mandant
OUTER APPLY (SELECT TOP 1 CAST(x.Belegjahr AS varchar) + '-' + CAST(x.Belegnummer AS varchar) AS Beleg FROM KHKVKBelege x
             WHERE x.Mandant = k.Mandant AND x.A0Empfaenger = k.A0Empfaenger AND x.Belegdatum >= k.Belegdatum
               AND (x.Belegart LIKE 'Auftragsbest%' OR x.Belegart LIKE 'Reservierungsbest%') ORDER BY x.Belegdatum) ab
WHERE k.Mandant = $Mandant AND k.Belegart = 'Angebot' AND k.VorID = 0 AND k.Belegdatum >= @von
ORDER BY k.Belegdatum DESC, k.BelID DESC
"@
function Build-Angebote($monate, $fehler) {
    $head = Page-Head "Solar Dispo - Angebote"
    $heute = (Get-Date).Date
    $inhalt = ""; $anzahl = 0
    if ($fehler) { $inhalt = "<div class='fehler'>Fehler beim Laden der Daten:<br><pre>$(HtmlEncode $fehler)</pre></div>" }
    else {
        $termine = Get-Kalender
        $ang = [ordered]@{}
        foreach ($r in (Run-Query $SQL_OffeneAngebote @{ "@von" = $heute.AddMonths(-$monate) }).Rows) {
            $id = [string]$r["BelID"]
            if (-not $ang.Contains($id)) {
                $ort = ([string]$r["A1Ort"]).Trim(); if (-not $ort) { $ort = ([string]$r["A0Ort"]).Trim() }
                $ang[$id] = @{ BelID = $id; Beleg = [string]$r["Beleg"]; Datum = $r["Belegdatum"]; Kunde = ((@(([string]$r["A0Name1"]).Trim(), ([string]$r["A0Name2"]).Trim()) | Where-Object { $_ }) -join ' ')
                               Name1 = [string]$r["A0Name1"]; Ort = $ort; Kto = [string]$r["A0Empfaenger"]; Plz = @(([string]$r["A1PLZ"]).Trim(), ([string]$r["A0PLZ"]).Trim()); Str = $(if (([string]$r["A1Strasse"]).Trim()) { [string]$r["A1Strasse"] } else { [string]$r["A0Strasse"] }); kWp = 0.0; Module = 0.0; Teile = @{}; AuftragDanach = $(if ($r["AuftragDanach"] -is [System.DBNull]) { "" } else { [string]$r["AuftragDanach"] }) }
            }
            $g = $ang[$id]; $gr = [string]$r["Artikelgruppe"]; $m = [double](ToNum $r["Menge"]); $bez = [string]$r["PosBezeichnung"]
            if ($r["Artikelnummer"] -is [System.DBNull]) { continue }
            $wp = Modul-Wp $gr $bez; if ($wp) { $g.kWp += $m * $wp / 1000; $g.Module += $m }
            switch ($gr) {
                '60' { if ((ToNum $r["Stuecklistentyp"]) -eq 2 -or $bez -match '(?i)batter|speicher|reserva') { $g.Teile['Speicher'] = 1 } else { $g.Teile['Wechselrichter'] = 1 } }
                '64' { $g.Teile['Wallbox'] = 1 }
                '12' { if ($bez -match '(?i)flachdach') { $g.Teile['Flachdach'] = 1 } elseif ($bez -match '(?i)steildach') { $g.Teile['Steildach'] = 1 } }
                '40' { $g.Teile['Mini-PV'] = 1 }
            }
        }
        $anzahl = $ang.Count
        $cn = try { Get-Craftnote } catch { $null }
        $rows = foreach ($g in $ang.Values) {
            $alter = ($heute - ([datetime]$g.Datum).Date).Days
            $tr = @(Kunden-Termine $g.Name1 $termine | Sort-Object { $_.T.Start })
            $naechster = @($tr | Where-Object { $_.T.Ende.Date -ge $heute }) | Select-Object -First 1
            $letzter = @($tr | Where-Object { $_.T.Ende.Date -lt $heute }) | Select-Object -Last 1
            $terminTxt = if ($naechster) { Termin-Html $naechster $true } elseif ($letzter) { "<span class='leise'>zuletzt</span> " + (Termin-Html $letzter $true) } else { "" }
            $chips = (@('Wechselrichter', 'Speicher', 'Wallbox', 'Steildach', 'Flachdach', 'Mini-PV') | Where-Object { $g.Teile.ContainsKey($_) } | ForEach-Object { "<span class='chip kein'>$_</span>" }) -join ''
            $anlage = if ($g.kWp) { "$(Fmt-kWp $g.kWp) <span class='leise'>($(FmtZahl $g.Module) Module)</span>" } else { "" }
            $stil = if ($g.AuftragDanach) { " style='color:#999'" } else { "" }
            $hinweis = if ($g.AuftragDanach) { " <span class='chip da' title='Der Kunde hat nach diesem Angebot einen Auftrag erteilt (z.B. &uuml;ber eine andere Variante)'>Auftrag $(HtmlEncode $g.AuftragDanach)</span>" } else { "" }
            "<tr$stil><td style='white-space:nowrap'>$(([datetime]$g.Datum).ToString('dd.MM.yyyy'))</td><td class='num'>$alter</td><td style='white-space:nowrap'>$(HtmlEncode $g.Beleg)$hinweis</td>" +
            "<td>$(HtmlEncode $g.Kunde)</td><td>$(HtmlEncode $g.Ort)</td><td style='white-space:nowrap'>$anlage</td><td>$chips</td><td>$(CN-Chip (CN-Finde $cn $g.Name1 $g.Kto $g.Plz $g.Str))</td><td>$terminTxt</td></tr>"
        }
        $inhalt = if ($anzahl) {
            "<table class='filtertab'><thead><tr><th>Datum</th><th class='num'>Tage</th><th>Angebot</th><th>Kunde</th><th>Ort</th><th>Anlage</th><th>Umfang</th><th>Craftnote</th><th>Kalender</th></tr></thead><tbody>$($rows -join '')</tbody></table>"
        } else { "<p class='leise'>Keine offenen Angebote in diesem Zeitraum.</p>" }
    }
    $knoepfe = (@(3, 6, 12) | ForEach-Object { $cls = if ($_ -eq $monate) { 'btn' } else { 'btn btn2' }; "<a class='$cls' href='/angebote?monate=$_'>$_ Monate</a>" }) -join ' '
    return @"
$head
<div class='kopf'><h1>Offene Angebote</h1><div class='meta'>$(StandText) &middot; $anzahl Angebote</div></div>
<div class='steuer'>$knoepfe <input class='suche' id='suche' placeholder='Filtern ...' onkeyup='filter()'></div>
$(if ($script:CnFehler) { "<div class='hinweisbox'>$($script:CnFehler)</div>" })
$(if ($script:KalFehler) { "<div class='hinweisbox'>$(HtmlEncode $script:KalFehler)</div>" })
$inhalt
<script>
function filter(){
  var q=document.getElementById('suche').value.toLowerCase();
  document.querySelectorAll('.filtertab tbody tr').forEach(function(tr){ tr.style.display = tr.innerText.toLowerCase().indexOf(q)>-1 ? '' : 'none'; });
}
</script>
</body></html>
"@
}


# ---------- Dispo-Diagramm, Dokumente, Bestandsverlauf, Verbrauch (aus der Oezer-Liste uebernommen) ----------
# Zeitstrahl der Dispo-Seite (Stufenlinie des verfuegbaren Bestands, Meldebestand als Linie,
# Fadenkreuz-Tooltip). Daten kommen als DISPO = {melde, p:[{d, v, ev:[{t,b,i,m}]}]}.
# Einfache Anfuehrungszeichen: keine PowerShell-Ersetzung; nur ASCII (Umlaute als \u-Escapes).
$DispoChartJs = @'
(function () {
  var el = document.getElementById('chart'), tip = document.getElementById('charttip');
  if (!el || !window.DISPO || !DISPO.p.length) return;
  var P = DISPO.p, melde = DISPO.melde, DAY = 864e5;
  var C = { serie: '#2a78d6', flaeche: 'rgba(42,120,214,0.10)', minus: 'rgba(208,59,59,0.14)',
            melde: '#eda100', grid: '#e1e0d9', achse: '#c3c2b7', surface: '#fcfcfb', kreuz: '#898781' };
  function tag(s) { var a = s.split('-'); return new Date(+a[0], a[1] - 1, +a[2]).getTime(); }
  function fmt(n) { return n.toLocaleString('de-DE', { maximumFractionDigits: 2 }); }
  function fmtDat(t, kurz) {
    var d = new Date(t), s = ('0' + d.getDate()).slice(-2) + '.' + ('0' + (d.getMonth() + 1)).slice(-2) + '.';
    return kurz ? s : s + d.getFullYear();
  }
  P.forEach(function (p) { p.t = tag(p.d); });

  var t0 = P[0].t, tL = P[P.length - 1].t;
  var t1 = t0 + Math.max(tL - t0, 21 * DAY) * 1.06;
  var vMin = Math.min(0, melde), vMax = Math.max(0, melde);
  P.forEach(function (p) { vMin = Math.min(vMin, p.v); vMax = Math.max(vMax, p.v); });
  if (vMax === vMin) vMax = vMin + 1;
  function niceStep(range, n) {
    var raw = range / n, mag = Math.pow(10, Math.floor(Math.log(raw) / Math.LN10)), r = raw / mag;
    return (r <= 1 ? 1 : r <= 2 ? 2 : r <= 2.5 ? 2.5 : r <= 5 ? 5 : 10) * mag;
  }
  var step = niceStep(vMax - vMin, 5);
  var y0 = Math.floor(vMin / step) * step;
  var y1 = Math.ceil((vMax + (vMax - vMin) * 0.06) / step) * step;

  function zeichne() {
    var W = el.clientWidth, H = el.clientHeight || 280, m = { l: 60, r: 132, t: 14, b: 26 };
    var iw = Math.max(W - m.l - m.r, 50), ih = H - m.t - m.b;
    function x(t) { return m.l + (t - t0) / (t1 - t0) * iw; }
    function y(v) { return m.t + (y1 - v) / (y1 - y0) * ih; }
    var s = [], i, v;
    s.push('<svg width="' + W + '" height="' + H + '" role="img" aria-label="Verf\u00fcgbarer Bestand im Zeitverlauf">');
    s.push('<defs><clipPath id="cpPlus"><rect x="0" y="0" width="' + W + '" height="' + y(0) + '"/></clipPath>' +
           '<clipPath id="cpMinus"><rect x="0" y="' + y(0) + '" width="' + W + '" height="' + H + '"/></clipPath></defs>');
    for (v = y0; v <= y1 + step / 2; v += step) {
      s.push('<line x1="' + m.l + '" x2="' + (m.l + iw) + '" y1="' + y(v) + '" y2="' + y(v) + '" stroke="' + C.grid + '" stroke-width="1"/>');
      s.push('<text x="' + (m.l - 8) + '" y="' + (y(v) + 4) + '" text-anchor="end">' + fmt(v) + '</text>');
    }
    var tage = (t1 - t0) / DAY, k = tage <= 45 ? 7 : tage <= 120 ? 14 : tage <= 300 ? 30 : 61;
    var d = new Date(t0);
    if (k < 30) { while (d.getDay() !== 1) d.setDate(d.getDate() + 1); } else { d = new Date(d.getFullYear(), d.getMonth() + 1, 1); }
    while (d.getTime() <= t1) {
      var tt = d.getTime();
      s.push('<line x1="' + x(tt) + '" x2="' + x(tt) + '" y1="' + (H - m.b) + '" y2="' + (H - m.b + 4) + '" stroke="' + C.achse + '"/>');
      s.push('<text x="' + x(tt) + '" y="' + (H - m.b + 16) + '" text-anchor="middle">' + fmtDat(tt, true) + '</text>');
      if (k >= 30) d.setMonth(d.getMonth() + (k === 61 ? 2 : 1)); else d.setDate(d.getDate() + k);
    }
    var pfad = 'M' + x(P[0].t) + ',' + y(P[0].v);
    for (i = 1; i < P.length; i++) pfad += 'H' + x(P[i].t) + 'V' + y(P[i].v);
    pfad += 'H' + x(t1);
    var flaeche = pfad + 'V' + y(0) + 'H' + x(P[0].t) + 'Z';
    s.push('<path d="' + flaeche + '" fill="' + C.flaeche + '" clip-path="url(#cpPlus)"/>');
    s.push('<path d="' + flaeche + '" fill="' + C.minus + '" clip-path="url(#cpMinus)"/>');
    s.push('<line x1="' + m.l + '" x2="' + (m.l + iw) + '" y1="' + y(0) + '" y2="' + y(0) + '" stroke="' + C.achse + '" stroke-width="1"/>');
    if (melde > 0) {
      s.push('<line x1="' + m.l + '" x2="' + (m.l + iw) + '" y1="' + y(melde) + '" y2="' + y(melde) + '" stroke="' + C.melde + '" stroke-width="2"/>');
      s.push('<text class="lbl" x="' + (m.l + iw + 8) + '" y="' + (y(melde) + 4) + '">Meldebestand ' + fmt(melde) + '</text>');
    }
    s.push('<path d="' + pfad + '" fill="none" stroke="' + C.serie + '" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>');
    function punkt(px, py) { return '<circle cx="' + px + '" cy="' + py + '" r="4" fill="' + C.serie + '" stroke="' + C.surface + '" stroke-width="2"/>'; }
    var last = P[P.length - 1], ly = y(last.v) + 4;
    if (melde > 0 && Math.abs(y(last.v) - y(melde)) < 14) ly = y(last.v) > y(melde) ? y(melde) + 18 : y(melde) - 10;
    s.push(punkt(x(t1), y(last.v)));
    s.push('<text class="lbl" x="' + (m.l + iw + 8) + '" y="' + ly + '">' + fmt(last.v) + '</text>');
    var iMin = 0;
    for (i = 1; i < P.length; i++) if (P[i].v < P[iMin].v) iMin = i;
    if (iMin > 0 && iMin < P.length - 1) {
      var mx = x(P[iMin].t), my = y(P[iMin].v);
      s.push(punkt(mx, my));
      s.push('<text class="lbl" x="' + mx + '" y="' + (my > H - m.b - 26 ? my - 10 : my + 18) + '" text-anchor="middle">tiefster Stand ' + fmt(P[iMin].v) + '</text>');
    }
    s.push('<line id="kreuz" x1="0" x2="0" y1="' + m.t + '" y2="' + (H - m.b) + '" stroke="' + C.kreuz + '" stroke-width="1" visibility="hidden"/>');
    s.push('<circle id="kreuzpunkt" r="5" fill="' + C.serie + '" stroke="' + C.surface + '" stroke-width="2" visibility="hidden"/>');
    s.push('<rect id="treffer" x="' + m.l + '" y="0" width="' + iw + '" height="' + H + '" fill="transparent"/>');
    s.push('</svg>');
    el.innerHTML = s.join('');

    var kreuz = document.getElementById('kreuz'), kp = document.getElementById('kreuzpunkt');
    var treffer = document.getElementById('treffer');
    function zeile(links, rechts) {
      var z = document.createElement('div'), a = document.createElement('span'), b = document.createElement('span');
      z.className = 'zeile'; a.textContent = links; b.textContent = rechts; z.appendChild(a); z.appendChild(b); return z;
    }
    treffer.addEventListener('pointermove', function (evt) {
      var r = el.getBoundingClientRect(), px0 = evt.clientX - r.left, best = 0, bd = 1e18;
      for (var j = 0; j < P.length; j++) { var dd = Math.abs(x(P[j].t) - px0); if (dd < bd) { bd = dd; best = j; } }
      var p = P[best], px = x(p.t), py = y(p.v);
      kreuz.setAttribute('x1', px); kreuz.setAttribute('x2', px); kreuz.setAttribute('visibility', 'visible');
      kp.setAttribute('cx', px); kp.setAttribute('cy', py); kp.setAttribute('visibility', 'visible');
      tip.textContent = '';
      var w = document.createElement('div'); w.className = 'wert'; w.textContent = fmt(p.v); tip.appendChild(w);
      var dz = document.createElement('div');
      dz.textContent = (best === 0 ? 'heute, ' : '') + fmtDat(p.t) + ' \u2013 verf\u00fcgbar nach diesem Tag';
      tip.appendChild(dz);
      p.ev.slice(0, 10).forEach(function (e) {
        var txt = e.t + (e.b ? ' ' + e.b : '') + (e.i ? ' \u00b7 ' + e.i : '');
        tip.appendChild(zeile(txt, (e.t !== 'Bestand' && e.m > 0 ? '+' : '') + fmt(e.m)));
      });
      if (p.ev.length > 10) tip.appendChild(zeile('\u2026 ' + (p.ev.length - 10) + ' weitere', ''));
      tip.hidden = false;
      var karte = tip.parentNode, left = el.offsetLeft + px + 14;
      if (left + tip.offsetWidth > karte.clientWidth - 6) left = el.offsetLeft + px - tip.offsetWidth - 14;
      tip.style.left = Math.max(6, left) + 'px';
      tip.style.top = (el.offsetTop + Math.max(0, Math.min(py - 24, H - tip.offsetHeight))) + 'px';
    });
    treffer.addEventListener('pointerleave', function () {
      kreuz.setAttribute('visibility', 'hidden'); kp.setAttribute('visibility', 'hidden'); tip.hidden = true;
    });
  }
  zeichne();
  var rt; window.addEventListener('resize', function () { clearTimeout(rt); rt = setTimeout(zeichne, 150); });
})();
'@

# ---------- Dokumente aus der Sage-Sammelmappe ----------
# KHKSammelmappen: Typ 2 = Artikel (Schluessel = Artikelnummer), Quelle = Dateipfad (UNC), Mappe = Ordner.
# Die Ordnernamen stehen nicht in der Datenbank - hier aus dem Inhalt abgeleitet (bei Bedarf anpassen).
$SammelmappenNamen = @{ '1' = 'Datenbl&auml;tter' }
$SammelmappenReihenfolge = @('1')
$SQL_ArtikelDokumente = @"
SELECT ID, Mappe, Titel, Quelle, Datum, Benutzer FROM KHKSammelmappen
WHERE Mandant = $Mandant AND Typ = 2 AND Schluessel = @art
ORDER BY Mappe, Datum DESC
"@
$DokuMaxMB = 60   # groessere Dateien werden nicht ueber den (einthreadigen) Webserver ausgeliefert

function Get-DateiTyp($pfad) {
    switch -regex ([System.IO.Path]::GetExtension([string]$pfad).ToLower()) {
        '^\.pdf$'           { return @{ Kurz = 'PDF'; Mime = 'application/pdf'; Bild = $false; Inline = $true } }
        '^\.(jpg|jpeg)$'    { return @{ Kurz = 'JPG'; Mime = 'image/jpeg'; Bild = $true; Inline = $true } }
        '^\.png$'           { return @{ Kurz = 'PNG'; Mime = 'image/png'; Bild = $true; Inline = $true } }
        '^\.gif$'           { return @{ Kurz = 'GIF'; Mime = 'image/gif'; Bild = $true; Inline = $true } }
        '^\.mp4$'           { return @{ Kurz = 'MP4'; Mime = 'video/mp4'; Bild = $false; Inline = $true } }
        '^\.(tif|tiff)$'    { return @{ Kurz = 'TIF'; Mime = 'image/tiff'; Bild = $false; Inline = $false } }
        '^\.xlsx$'          { return @{ Kurz = 'XLSX'; Mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'; Bild = $false; Inline = $false } }
        default             { return @{ Kurz = ([System.IO.Path]::GetExtension([string]$pfad).TrimStart('.').ToUpper()); Mime = 'application/octet-stream'; Bild = $false; Inline = $false } }
    }
}

function Build-Dokumente($artikel) {
    $t = Run-Query $SQL_ArtikelDokumente @{ "@art" = $artikel }
    if ($t.Rows.Count -eq 0) {
        return @{ Anzahl = 0; Html = "<div class='dokkarte' id='dokumente'><h3>Dokumente</h3><p style='color:#666;margin:0'>In der Sage-Sammelmappe sind f&uuml;r diesen Artikel keine Dokumente hinterlegt.</p></div>" }
    }
    $gruppen = @{}
    foreach ($r in $t.Rows) {
        $m = [string]$r["Mappe"]
        if (-not $gruppen.ContainsKey($m)) { $gruppen[$m] = New-Object System.Collections.ArrayList }
        [void]$gruppen[$m].Add($r)
    }
    $reihe = @($SammelmappenReihenfolge | Where-Object { $gruppen.ContainsKey($_) }) + @($gruppen.Keys | Where-Object { $SammelmappenReihenfolge -notcontains $_ } | Sort-Object)
    $teile = foreach ($m in $reihe) {
        $name = if ($SammelmappenNamen.ContainsKey($m)) { $SammelmappenNamen[$m] } else { "Mappe $m" }
        $bilder = @(); $liste = @()
        foreach ($r in $gruppen[$m]) {
            $pfad = ([string]$r["Quelle"]).Trim()
            $typ = Get-DateiTyp $pfad
            $titel = ([string]$r["Titel"]).Trim(); if (-not $titel) { $titel = [System.IO.Path]::GetFileName($pfad) }
            $datum = if ($r["Datum"] -is [datetime]) { $r["Datum"].ToString("dd.MM.yyyy") } else { "" }
            $meta = (@($datum, ([string]$r["Benutzer"]).Trim()) | Where-Object { $_ }) -join ' &middot; '
            $vorhanden = $false
            if ($pfad) { try { $vorhanden = Test-Path -LiteralPath $pfad -PathType Leaf } catch { $vorhanden = $false } }
            $kopie = if ($pfad) { " <a class='dokpfad' href='javascript:void(0)' data-pfad='$(HtmlEncode $pfad)' onclick='pfadKopieren(this)' title='$(HtmlEncode $pfad)'>Pfad kopieren</a>" } else { "" }
            if (-not $vorhanden) {
                $liste += "<div class='dok fehlt'><span class='dokext'>$(HtmlEncode $typ.Kurz)</span> <span>$(HtmlEncode $titel)</span> <span class='dokmeta'>$meta &middot; Datei nicht gefunden</span>$kopie</div>"
            } elseif ($typ.Bild) {
                $bilder += "<a class='dokbild' href='/datei?id=$($r['ID'])' target='_blank' title='$(HtmlEncode $titel)'><img src='/datei?id=$($r['ID'])' loading='lazy' alt=''><span>$(HtmlEncode $titel)</span></a>"
            } else {
                $liste += "<div class='dok'><span class='dokext ext-$($typ.Kurz.ToLower())'>$(HtmlEncode $typ.Kurz)</span> <a class='artlink' href='/datei?id=$($r['ID'])' target='_blank'>$(HtmlEncode $titel)</a> <span class='dokmeta'>$meta</span>$kopie</div>"
            }
        }
        $inhalt = ""
        if ($bilder.Count) { $inhalt += "<div class='dokbilder'>$($bilder -join '')</div>" }
        if ($liste.Count) { $inhalt += ($liste -join '') }
        "<div class='dokgruppe'><h4>$name <span style='color:#888;font-weight:400'>($($gruppen[$m].Count))</span></h4>$inhalt</div>"
    }
    $html = @"
<div class='dokkarte' id='dokumente'>
<h3>Dokumente</h3>
$($teile -join '')
</div>
<script>
function pfadKopieren(a) {
  var p = a.getAttribute('data-pfad'), ta = document.createElement('textarea');
  ta.value = p; document.body.appendChild(ta); ta.select();
  try { document.execCommand('copy'); a.textContent = 'kopiert'; setTimeout(function(){ a.textContent = 'Pfad kopieren'; }, 1500); } catch (e) { prompt('Pfad:', p); }
  document.body.removeChild(ta);
}
</script>
"@
    return @{ Anzahl = $t.Rows.Count; Html = $html }
}
# ---------- Historischer Lagerbestand (letzte 12 Wochen) ----------
# Rueckrechnung: Bestand Ende Woche = heutiger Bestand (KHKArtikelVarianten) minus alle spaeteren
# bestandswirksamen Lagerplatzbuchungen. Bestandswirkung (+1/-1/0) steht in KHKLagerbewegungsarten;
# Umbuchungen (IU/SU/US/IV) haben 0 und zaehlen nicht. Geprueft 14.10.2026: Summe aller Buchungen
# = Lagerbestand bei 3211 von 3486 Artikeln, Rest meist Rundung / Altbestand vor Buchungsbeginn.
# Sage speichert nur das Buchungsdatum; Buchungen mit Datum nach heute zaehlen zur laufenden Woche.
$HistorieWochen = 12
$SQL_BestandHistorie = @"
SELECT b.Bewegungsdatum AS Datum, b.Bewegungsart, ba.Bestandswirkung, SUM(b.MengeBasis) AS Menge
FROM KHKLagerplatzbuchungen b
INNER JOIN KHKLagerbewegungsarten ba ON ba.Bewegungsart = b.Bewegungsart
WHERE b.Mandant = $Mandant AND b.Artikelnummer = @art AND b.MengeBasis <> 0
  AND ba.Bestandswirkung <> 0 AND b.Bewegungsdatum >= @von
GROUP BY b.Bewegungsdatum, b.Bewegungsart, ba.Bestandswirkung
"@

# Wochendiagramm auf einer Achse: Bestand am Wochenende (Linie), Zu-/Abgaenge je Woche (zwei Balken nebeneinander).
# Daten: HIST = {start, w:[{kw, von, bis, lfd, b, zu, ab, zuD:[[text,menge]], abD:[...]}]}
# Einfache Anfuehrungszeichen: keine PowerShell-Ersetzung; nur ASCII (Umlaute als \u-Escapes).
$HistorieChartJs = @'
(function () {
  var el = document.getElementById('hchart'), tip = document.getElementById('hcharttip');
  if (!el || !window.HIST || !HIST.w.length) return;
  var W8 = HIST.w, n = W8.length;
  var C = { bestand: '#2a78d6', zu: '#1baf7a', ab: '#eb6834', grid: '#e1e0d9', achse: '#c3c2b7',
            surface: '#fcfcfb', band: 'rgba(11,11,11,0.05)' };
  function fmt(v) { return v.toLocaleString('de-DE', { maximumFractionDigits: 2 }); }
  function niceStep(range, k) {
    if (range <= 0) return 1;
    var raw = range / k, mag = Math.pow(10, Math.floor(Math.log(raw) / Math.LN10)), r = raw / mag;
    return (r <= 1 ? 1 : r <= 2 ? 2 : r <= 2.5 ? 2.5 : r <= 5 ? 5 : 10) * mag;
  }
  // Eine Achse fuer Bestand und Bewegungen, ab 0 (nur bei negativem Bestand darunter)
  var lo = Math.min(0, HIST.start), hi = Math.max(0, HIST.start);
  W8.forEach(function (w) { lo = Math.min(lo, w.b); hi = Math.max(hi, w.b, w.zu, w.ab); });
  hi *= 1.08; if (hi === lo) hi = lo + 1;
  var st = niceStep(hi - lo, 5), S = { lo: Math.floor(lo / st) * st, hi: Math.ceil(hi / st) * st };

  function balken(x, w, y0, y1, r) {
    // Balken mit 4px gerundetem Datenende, an der Nulllinie gerade
    var h = y0 - y1; r = Math.min(r, h, w / 2);
    if (h < 0.5) return '';
    return 'M' + x + ',' + y0 + 'V' + (y1 + r) + 'Q' + x + ',' + y1 + ' ' + (x + r) + ',' + y1 +
           'H' + (x + w - r) + 'Q' + (x + w) + ',' + y1 + ' ' + (x + w) + ',' + (y1 + r) + 'V' + y0 + 'Z';
  }

  function zeichne() {
    var W = el.clientWidth, H = el.clientHeight || 320, m = { l: 64, r: 18, t: 22, b: 26 };
    var iw = Math.max(W - m.l - m.r, 120), ih = H - m.t - m.b, bw = iw / n;
    function xEnde(i) { return m.l + bw * (i + 1); }
    function xMitte(i) { return m.l + bw * (i + 0.5); }
    function y(v) { return m.t + (S.hi - v) / (S.hi - S.lo) * ih; }
    var s = [], i, v;
    s.push('<svg width="' + W + '" height="' + H + '" role="img" aria-label="Lagerbestand, Zug\u00e4nge und Abg\u00e4nge der letzten ' + n + ' Wochen">');
    s.push('<rect id="hband" x="0" y="' + (m.t - 18) + '" width="' + bw + '" height="' + (H - m.b - m.t + 18) + '" fill="' + C.band + '" visibility="hidden"/>');
    for (v = S.lo; v <= S.hi + st / 2; v += st) {
      s.push('<line x1="' + m.l + '" x2="' + (m.l + iw) + '" y1="' + y(v) + '" y2="' + y(v) + '" stroke="' + (Math.abs(v) < st / 2 ? C.achse : C.grid) + '"/>');
      s.push('<text x="' + (m.l - 8) + '" y="' + (y(v) + 4) + '" text-anchor="end">' + fmt(v) + '</text>');
    }
    var jede = bw < 44 ? 2 : 1;
    for (i = 0; i < n; i++) {
      if ((n - 1 - i) % jede) continue;
      s.push('<text x="' + xMitte(i) + '" y="' + (H - m.b + 16) + '" text-anchor="middle">' + W8[i].kw + '</text>');
    }
    // je Woche zwei Balken nebeneinander (Zugang | Abgang), 2px Abstand
    var bb = Math.min(22, (bw * 0.7 - 2) / 2);
    for (i = 0; i < n; i++) {
      var x0 = xMitte(i) - bb - 1;
      if (W8[i].zu > 0) s.push('<path d="' + balken(x0, bb, y(0), y(W8[i].zu), 4) + '" fill="' + C.zu + '"/>');
      if (W8[i].ab > 0) s.push('<path d="' + balken(x0 + bb + 2, bb, y(0), y(W8[i].ab), 4) + '" fill="' + C.ab + '"/>');
    }
    // Bestandslinie: Startwert am linken Rand, dann je Wochenende
    var pfad = 'M' + m.l + ',' + y(HIST.start);
    for (i = 0; i < n; i++) pfad += 'L' + xEnde(i) + ',' + y(W8[i].b);
    s.push('<path d="' + pfad + '" fill="none" stroke="' + C.surface + '" stroke-width="5" stroke-linejoin="round" stroke-linecap="round"/>');
    s.push('<path d="' + pfad + '" fill="none" stroke="' + C.bestand + '" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>');
    for (i = 0; i < n; i++) s.push('<circle cx="' + xEnde(i) + '" cy="' + y(W8[i].b) + '" r="4" fill="' + C.bestand + '" stroke="' + C.surface + '" stroke-width="2"/>');
    var last = W8[n - 1];
    s.push('<text class="lbl" x="' + (xEnde(n - 1) - 8) + '" y="' + (y(last.b) - 10) + '" text-anchor="end" stroke="' + C.surface + '" stroke-width="4" paint-order="stroke">heute ' + fmt(last.b) + '</text>');
    s.push('<rect id="htreffer" x="' + m.l + '" y="0" width="' + iw + '" height="' + H + '" fill="transparent"/>');
    s.push('</svg>');
    el.innerHTML = s.join('');

    var band = document.getElementById('hband'), treffer = document.getElementById('htreffer');
    function zeile(links, rechts, cls) {
      var z = document.createElement('div'), a = document.createElement('span'), b = document.createElement('span');
      z.className = 'zeile' + (cls ? ' ' + cls : ''); a.textContent = links; b.textContent = rechts;
      z.appendChild(a); z.appendChild(b); return z;
    }
    function kopfzeile(farbe, links, rechts) {
      var z = zeile(links, rechts, 'fett'), p = document.createElement('i');
      p.className = 'hpunkt'; p.style.background = farbe; z.firstChild.insertBefore(p, z.firstChild.firstChild);
      tip.appendChild(z);
    }
    treffer.addEventListener('pointermove', function (evt) {
      var r = el.getBoundingClientRect(), px = evt.clientX - r.left;
      var k = Math.max(0, Math.min(n - 1, Math.floor((px - m.l) / bw))), w = W8[k];
      band.setAttribute('x', m.l + bw * k); band.setAttribute('visibility', 'visible');
      tip.textContent = '';
      var t = document.createElement('div'); t.className = 'wert';
      t.textContent = w.kw + (w.lfd ? ' (laufend)' : ''); tip.appendChild(t);
      var d = document.createElement('div'); d.textContent = w.von + ' \u2013 ' + w.bis; tip.appendChild(d);
      kopfzeile(C.bestand, 'Bestand ' + (w.lfd ? 'heute' : 'am Wochenende'), fmt(w.b));
      kopfzeile(C.zu, 'Zug\u00e4nge', '+' + fmt(w.zu));
      w.zuD.forEach(function (x) { tip.appendChild(zeile(x[0], '+' + fmt(x[1]), 'detail')); });
      kopfzeile(C.ab, 'Abg\u00e4nge', '\u2212' + fmt(w.ab));
      w.abD.forEach(function (x) { tip.appendChild(zeile(x[0], '\u2212' + fmt(x[1]), 'detail')); });
      tip.hidden = false;
      var karte = tip.parentNode, mx = xMitte(k), left = el.offsetLeft + mx + bw / 2 + 8;
      if (left + tip.offsetWidth > karte.clientWidth - 6) left = el.offsetLeft + mx - bw / 2 - tip.offsetWidth - 8;
      tip.style.left = Math.max(6, left) + 'px';
      tip.style.top = (el.offsetTop + 4) + 'px';
    });
    treffer.addEventListener('pointerleave', function () { band.setAttribute('visibility', 'hidden'); tip.hidden = true; });
  }
  zeichne();
  var rt; window.addEventListener('resize', function () { clearTimeout(rt); rt = setTimeout(zeichne, 150); });
})();
'@

function Build-BestandHistorie($artikel, $bestandHeute) {
    $inv = [Globalization.CultureInfo]::InvariantCulture
    $heute = (Get-Date).Date
    $montag = $heute.AddDays(-(([int]$heute.DayOfWeek + 6) % 7))
    $von = $montag.AddDays(-7 * ($HistorieWochen - 1))
    $t = Run-Query $SQL_BestandHistorie @{ "@art" = $artikel; "@von" = $von }

    # Je Woche: Zu-/Abgang gesamt und nach Bewegungsart
    $wochen = for ($i = 0; $i -lt $HistorieWochen; $i++) { @{ Zu = 0.0; Ab = 0.0; ZuD = @{}; AbD = @{} } }
    foreach ($r in $t.Rows) {
        $i = [int][math]::Floor(($r["Datum"].Date - $von).TotalDays / 7)
        if ($i -ge $HistorieWochen) { $i = $HistorieWochen - 1 }
        $menge = [double]$r["Menge"]; $ba = [string]$r["Bewegungsart"]
        $text = if ($BewegungsArten.ContainsKey($ba)) { [System.Net.WebUtility]::HtmlDecode($BewegungsArten[$ba][1]) } else { $ba }
        if ([int]$r["Bestandswirkung"] -gt 0) { $wochen[$i].Zu += $menge; $wochen[$i].ZuD[$text] += $menge }
        else { $wochen[$i].Ab += $menge; $wochen[$i].AbD[$text] += $menge }
    }

    # Rueckrechnung vom heutigen Bestand
    $b = ToNum $bestandHeute; if ($null -eq $b) { $b = 0 }
    $ende = New-Object double[] $HistorieWochen
    $ende[$HistorieWochen - 1] = [double]$b
    for ($i = $HistorieWochen - 1; $i -gt 0; $i--) { $ende[$i - 1] = $ende[$i] - $wochen[$i].Zu + $wochen[$i].Ab }
    $start = $ende[0] - $wochen[0].Zu + $wochen[0].Ab

    # JSON (Texte mit \u-Escapes, damit die Ausgabe unabhaengig vom Encoding bleibt)
    $js = { param($s) -join @(([string]$s).ToCharArray() | ForEach-Object { if ([int]$_ -lt 128 -and $_ -ne '"' -and $_ -ne '\' -and $_ -ne '<') { [string]$_ } else { '\u{0:x4}' -f [int]$_ } }) }
    $num = { param($x) ([double]$x).ToString($inv) }
    $det = { param($h) (@($h.Keys | Sort-Object { -$h[$_] }) | ForEach-Object { "[""$(& $js $_)"",$(& $num $h[$_])]" }) -join ',' }
    $kal = $inv.Calendar
    $wJson = @(); $summeZu = 0.0; $summeAb = 0.0
    for ($i = 0; $i -lt $HistorieWochen; $i++) {
        $mo = $von.AddDays(7 * $i); $so = $mo.AddDays(6); $lfd = ($i -eq $HistorieWochen - 1)
        $kw = "KW " + $kal.GetWeekOfYear($mo.AddDays(3), [Globalization.CalendarWeekRule]::FirstFourDayWeek, [DayOfWeek]::Monday)
        $w = $wochen[$i]; $summeZu += $w.Zu; $summeAb += $w.Ab
        $wJson += "{""kw"":""$kw"",""von"":""$($mo.ToString('dd.MM.yyyy'))"",""bis"":""$($so.ToString('dd.MM.yyyy'))"",""lfd"":$(if ($lfd) { 'true' } else { 'false' }),""b"":$(& $num $ende[$i]),""zu"":$(& $num $w.Zu),""ab"":$(& $num $w.Ab),""zuD"":[$(& $det $w.ZuD)],""abD"":[$(& $det $w.AbD)]}"
    }
    $titel = "<h3>Lagerbestand der letzten $HistorieWochen Wochen</h3>"
    if ($t.Rows.Count -eq 0) {
        return "<div class='dokkarte' id='historie'>$titel<p style='color:#666;margin:0'>Keine Lagerbewegungen seit $($von.ToString('dd.MM.yyyy')). Bestand heute: $(FmtZahl $b).</p></div>"
    }
    $daten = "{""start"":$(& $num $start),""w"":[$($wJson -join ',')]}"
    return @"
<div class='dokkarte hkarte' id='historie'>
$titel
<div class='hlegende'><span><i class='hpunkt' style='background:#2a78d6'></i>Bestand am Wochenende</span><span><i class='hquadrat' style='background:#1baf7a'></i>Zug&auml;nge</span><span><i class='hquadrat' style='background:#eb6834'></i>Abg&auml;nge</span>
<span class='hsumme'><b class='sign-plus'>+$(FmtZahl $summeZu)</b> / <b class='sign-minus'>&minus;$(FmtZahl $summeAb)</b> &middot; Bestand $($von.ToString('dd.MM.')) <b>$(FmtZahl $start)</b> &rarr; heute <b>$(FmtZahl $b)</b></span></div>
<div id='hchart'></div>
<div id='hcharttip' class='charttip' hidden></div>
</div>
<script>
var HIST = $daten;
$HistorieChartJs
</script>
"@
}

# ---------- Verbrauch ueber 36 Monate (Dispo) ----------
# Verbrauch = Lieferung an Kunde (EA) + manuelle Entnahme (EM) - Storno Lieferung (ZS) (EF Fertigung gibt es bei Solar nicht),
# wie auf der Engpass-Seite. Nur abgeschlossene Monate. Kennzahlen als Durchschnitt je Woche (Tage/7);
# Reichweite = heutiger Bestand / gewichteter Durchschnitt der letzten 6 Monate (Gewichte 6..1, wie die Linie im Diagramm).
$VerbrauchMonate = 36
$SQL_VerbrauchMonate = @"
SELECT YEAR(b.Bewegungsdatum) AS J, MONTH(b.Bewegungsdatum) AS M, b.Bewegungsart, SUM(b.MengeBasis) AS Menge
FROM KHKLagerplatzbuchungen b
WHERE b.Mandant = $Mandant AND b.Artikelnummer = @art AND b.Bewegungsart IN ('EA','EF','EM','ZS')
  AND b.Bewegungsdatum >= @von AND b.Bewegungsdatum < @bis
GROUP BY YEAR(b.Bewegungsdatum), MONTH(b.Bewegungsdatum), b.Bewegungsart
"@

# Monatsbalken (Durchschnitt je Woche des Monats), Linie = gewichteter gleitender Durchschnitt ueber 6 Monate
# (Gewichte 6..1, juengster Monat am staerksten; am Anfang ueber die vorhandenen Monate).
# Daten: VERB = {avg12, m:[{l, k, j, w, s, d:[[text, menge]]}]}. Nur ASCII (Umlaute als \u-Escapes).
$VerbrauchChartJs = @'
(function () {
  var el = document.getElementById('vchart'), tip = document.getElementById('vcharttip');
  if (!el || !window.VERB || !VERB.m.length) return;
  var M = VERB.m, n = M.length, avg = VERB.avg12;
  var C = { balken: '#eb6834', linie: '#2a78d6', grid: '#e1e0d9', achse: '#c3c2b7', band: 'rgba(11,11,11,0.05)' };
  function fmt(v) { return v.toLocaleString('de-DE', { maximumFractionDigits: 0 }); }
  function niceStep(range, k) {
    if (range <= 0) return 1;
    var raw = range / k, mag = Math.pow(10, Math.floor(Math.log(raw) / Math.LN10)), r = raw / mag;
    return (r <= 1 ? 1 : r <= 2 ? 2 : r <= 2.5 ? 2.5 : r <= 5 ? 5 : 10) * mag;
  }
  // gewichteter gleitender Durchschnitt (WMA) ueber 6 Monate
  var GEW = 6, wma = M.map(function (x, i) {
    var sw = 0, sg = 0;
    for (var j = 0; j < GEW && i - j >= (VERB.ab || 0); j++) { sw += M[i - j].w * (GEW - j); sg += GEW - j; }
    return sg ? sw / sg : 0;
  });
  // Ausreisser kappen: Achse hoechstens bis zum 3-fachen Median bzw. knapp ueber den viertgroessten Monat
  var werte = M.map(function (x) { return x.w; }).sort(function (a, b) { return b - a; });
  var cap = Math.max(werte[Math.floor(werte.length / 2)] * 3, (werte[3] || 0) * 1.1), hi = Math.max(Math.max.apply(null, wma), Math.min(werte[0], cap)) * 1.08 || 1;
  var st = niceStep(hi, 4), top = Math.ceil(hi / st) * st;

  function zeichne() {
    var W = el.clientWidth, H = el.clientHeight || 260, m = { l: 56, r: 18, t: 24, b: 40 };
    var iw = Math.max(W - m.l - m.r, 120), ih = H - m.t - m.b, bw = iw / n;
    function y(v) { return m.t + (top - v) / top * ih; }
    var s = [], i, v;
    s.push('<svg width="' + W + '" height="' + H + '" role="img" aria-label="Verbrauch je Woche, Monate">');
    s.push('<rect id="vband" x="0" y="' + m.t + '" width="' + bw + '" height="' + ih + '" fill="' + C.band + '" visibility="hidden"/>');
    for (v = 0; v <= top + st / 2; v += st) {
      s.push('<line x1="' + m.l + '" x2="' + (m.l + iw) + '" y1="' + y(v) + '" y2="' + y(v) + '" stroke="' + (v === 0 ? C.achse : C.grid) + '"/>');
      s.push('<text x="' + (m.l - 8) + '" y="' + (y(v) + 4) + '" text-anchor="end">' + fmt(v) + '</text>');
    }
    var bb = Math.max(2, Math.min(18, bw - 3));
    for (i = 0; i < n; i++) {
      var gekappt = M[i].w > top, wv = gekappt ? top : M[i].w;
      var x0 = m.l + bw * i + (bw - bb) / 2, h = y(0) - y(wv), r = Math.min(3, h, bb / 2);
      if (h >= 0.5) s.push('<path d="M' + x0 + ',' + y(0) + 'V' + (y(wv) + r) + 'Q' + x0 + ',' + y(wv) + ' ' + (x0 + r) + ',' + y(wv) +
                           'H' + (x0 + bb - r) + 'Q' + (x0 + bb) + ',' + y(wv) + ' ' + (x0 + bb) + ',' + (y(wv) + r) + 'V' + y(0) + 'Z" fill="' + C.balken + '"/>');
      if (gekappt) {
        // Bruchmarke und echter Wert ueber dem gekappten Balken
        s.push('<line x1="' + (x0 - 2) + '" x2="' + (x0 + bb + 2) + '" y1="' + (y(wv) + 10) + '" y2="' + (y(wv) + 6) + '" stroke="#fff" stroke-width="3"/>');
        s.push('<text class="lbl" x="' + (x0 + bb / 2) + '" y="' + (y(wv) - 4) + '" text-anchor="middle">' + fmt(M[i].w) + '</text>');
      }
      s.push('<text x="' + (m.l + bw * (i + 0.5)) + '" y="' + (H - m.b + 14) + '" text-anchor="middle">' + M[i].k + '</text>');
      if (i === 0 || M[i].k === 'J' && M[i].l.indexOf('Jan') === 0)
        s.push('<text class="lbl" x="' + (m.l + bw * i + 2) + '" y="' + (H - m.b + 30) + '">' + M[i].j + '</text>');
    }
    // gewichteter gleitender Durchschnitt als Linie durch die Monatsmitten (heller Rand gegen die Balken)
    var pfad = '';
    for (i = 0; i < n; i++) pfad += (i ? 'L' : 'M') + (m.l + bw * (i + 0.5)) + ',' + y(Math.min(wma[i], top));
    s.push('<path d="' + pfad + '" fill="none" stroke="#fff" stroke-width="5" stroke-linejoin="round"/>');
    s.push('<path d="' + pfad + '" fill="none" stroke="' + C.linie + '" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>');
    var xl = m.l + bw * (n - 0.5), yl = y(Math.min(wma[n - 1], top));
    s.push('<circle cx="' + xl + '" cy="' + yl + '" r="4" fill="' + C.linie + '" stroke="#fff" stroke-width="2"/>');
    s.push('<text class="lbl" x="' + (xl - 8) + '" y="' + (yl - 10) + '" text-anchor="end" stroke="#fff" stroke-width="4" paint-order="stroke">gewichteter \u00d8 ' + fmt(wma[n - 1]) + '</text>');
    s.push('<rect id="vtreffer" x="' + m.l + '" y="0" width="' + iw + '" height="' + H + '" fill="transparent"/>');
    s.push('</svg>');
    el.innerHTML = s.join('');

    var band = document.getElementById('vband'), treffer = document.getElementById('vtreffer');
    function zeile(links, rechts, cls) {
      var z = document.createElement('div'), a = document.createElement('span'), b = document.createElement('span');
      z.className = 'zeile' + (cls ? ' ' + cls : ''); a.textContent = links; b.textContent = rechts;
      z.appendChild(a); z.appendChild(b); return z;
    }
    treffer.addEventListener('pointermove', function (evt) {
      var rc = el.getBoundingClientRect(), k = Math.max(0, Math.min(n - 1, Math.floor((evt.clientX - rc.left - m.l) / bw))), x = M[k];
      band.setAttribute('x', m.l + bw * k); band.setAttribute('visibility', 'visible');
      tip.textContent = '';
      var t = document.createElement('div'); t.className = 'wert'; t.textContent = x.l; tip.appendChild(t);
      tip.appendChild(zeile('\u00d8 je Woche', fmt(x.w), 'fett'));
      tip.appendChild(zeile('gewichteter \u00d8 (6 Monate)', fmt(wma[k]), 'fett'));
      tip.appendChild(zeile('Verbrauch im Monat', fmt(x.s), 'fett'));
      x.d.forEach(function (d) { tip.appendChild(zeile(d[0], (d[1] < 0 ? '\u2212' : '') + fmt(Math.abs(d[1])), 'detail')); });
      tip.hidden = false;
      var karte = tip.parentNode, left = el.offsetLeft + m.l + bw * (k + 1) + 8;
      if (left + tip.offsetWidth > karte.clientWidth - 6) left = el.offsetLeft + m.l + bw * k - tip.offsetWidth - 8;
      tip.style.left = Math.max(6, left) + 'px';
      tip.style.top = (el.offsetTop + 4) + 'px';
    });
    treffer.addEventListener('pointerleave', function () { band.setAttribute('visibility', 'hidden'); tip.hidden = true; });
  }
  zeichne();
  var rt; window.addEventListener('resize', function () { clearTimeout(rt); rt = setTimeout(zeichne, 150); });
})();
'@

function Build-Verbrauch($artikel, $bestandHeute) {
    $inv = [Globalization.CultureInfo]::InvariantCulture; $de = [Globalization.CultureInfo]'de-DE'
    $heute = (Get-Date).Date
    $bis = New-Object DateTime $heute.Year, $heute.Month, 1
    $von = $bis.AddMonths(-$VerbrauchMonate)
    $t = Run-Query $SQL_VerbrauchMonate @{ "@art" = $artikel; "@von" = $von; "@bis" = $bis }
    $titel = "<h3>Verbrauch</h3>"
    if ($t.Rows.Count -eq 0) {
        return "<div class='dokkarte' id='verbrauch'>$titel<p style='color:#666;margin:0'>Kein Verbrauch seit $($von.ToString('MM/yyyy')).</p></div>"
    }

    # Monate 0..35 (0 = aeltester), je Bewegungsart summiert
    $mon = for ($i = 0; $i -lt $VerbrauchMonate; $i++) { @{ EA = 0.0; EF = 0.0; EM = 0.0; ZS = 0.0 } }
    foreach ($r in $t.Rows) {
        $i = ([int]$r["J"] - $von.Year) * 12 + [int]$r["M"] - $von.Month
        if ($i -ge 0 -and $i -lt $VerbrauchMonate) { $mon[$i][[string]$r["Bewegungsart"]] += [double]$r["Menge"] }
    }
    $summe = { param($a, $b) $x = 0.0; for ($i = $a; $i -lt $b; $i++) { $x += $mon[$i].EA + $mon[$i].EF + $mon[$i].EM - $mon[$i].ZS }; [math]::Max(0, $x) }
    $tage = { param($a, $b) ($von.AddMonths($b) - $von.AddMonths($a)).TotalDays }
    $n = $VerbrauchMonate
    # Durchschnitte erst ab dem ersten Monat mit Verbrauch (neue Artikel), $null = Zeitraum liegt ganz davor
    $i0 = 0; while ($i0 -lt $n - 1 -and ($mon[$i0].EA + $mon[$i0].EF + $mon[$i0].EM) -le 0) { $i0++ }
    $proWo = { param($a, $b) $a = [math]::Max($a, $i0); if ($a -ge $b) { return $null }; (& $summe $a $b) / ((& $tage $a $b) / 7) }
    $w3 = & $proWo ($n - 3) $n; $w6 = & $proWo ($n - 6) $n; $w12 = & $proWo ($n - 12) $n; $wVj = & $proWo ($n - 24) ($n - 12)
    $seit = $von.AddMonths($i0).ToString('MM/yyyy')
    # gewichteter Durchschnitt je Woche: letzte 6 Monate (ab erstem Verbrauch), juengster Monat Gewicht 6
    $sw = 0.0; $sg = 0
    for ($j = 0; $j -lt 6 -and $n - 1 - $j -ge $i0; $j++) { $sw += (& $proWo ($n - 1 - $j) ($n - $j)) * (6 - $j); $sg += 6 - $j }
    $wGew = if ($sg) { $sw / $sg } else { 0.0 }

    $b = ToNum $bestandHeute; if ($null -eq $b) { $b = 0 }
    $zahl = { param($x) if ($null -eq $x) { return "&ndash;" }; ([math]::Round($x)).ToString('N0', $de) }
    $reichweite = { param($w) if ($null -eq $w -or $w -le 0) { $null } elseif ($b -le 0) { 0 } else { $b / $w } }
    $rwG = & $reichweite $wGew; $rw12 = & $reichweite $w12
    $rwTxt = if ($null -eq $rwG) { "&ndash;" } elseif ($rwG -ge 100) { "&gt; 100 Wochen" } else { "$($rwG.ToString('N1', $de)) Wochen" }
    $rwTitle = "Bestand $(& $zahl $b) / gewichteter &Oslash; $(& $zahl $wGew) je Woche" + $(if ($null -ne $rw12) { " &middot; bei &Oslash; 12 Monate: $($rw12.ToString('N1', $de)) Wochen" } else { "" })
    $trend = ""
    if ($null -ne $wVj -and $null -ne $w12 -and $wVj -gt 0) {
        $p = [math]::Round(($w12 - $wVj) / $wVj * 100)
        $trend = "<div class='vk-s' title='&Oslash; 12 Monate gegen&uuml;ber den 12 Monaten davor'>$(if ($p -ge 0) { '&#9650; +' } else { '&#9660; &minus;' })$([math]::Abs($p)) %</div>"
    }
    $kachel = { param($label, $wert, $sub, $title) "<div class='vkachel'$(if ($title) { " title='$title'" })><div class='vk-l'>$label</div><div class='vk-w'>$wert</div>$sub</div>" }
    # Zeitraum reicht vor den ersten Verbrauch: Hinweis "seit MM/yyyy"
    $abSub = { param($monate) if ($n - $monate -lt $i0 -and $n - $i0 -lt $monate) { "<div class='vk-s'>seit $seit</div>" } else { '' } }
    $kacheln = (& $kachel '&Oslash; 3 Monate' "$(& $zahl $w3) <small>/Wo</small>" (& $abSub 3) '') +
               (& $kachel '&Oslash; 6 Monate' "$(& $zahl $w6) <small>/Wo</small>" (& $abSub 6) '') +
               (& $kachel '&Oslash; 12 Monate' "$(& $zahl $w12) <small>/Wo</small>" (& $abSub 12) '') +
               (& $kachel 'Vorjahr' $(if ($null -eq $wVj) { '&ndash;' } else { "$(& $zahl $wVj) <small>/Wo</small>" }) $trend $(if ($null -eq $wVj) { "Erster Verbrauch $seit" } else { '&Oslash; der 12 Monate davor' })) +
               (& $kachel 'Reichweite' $rwTxt "<div class='vk-s'>bei gewichtetem &Oslash; $(& $zahl $wGew) /Wo</div>" $rwTitle)

    # JSON fuer das Diagramm (Texte mit \u-Escapes)
    $js = { param($s) -join @(([string]$s).ToCharArray() | ForEach-Object { if ([int]$_ -lt 128 -and $_ -ne '"' -and $_ -ne '\' -and $_ -ne '<') { [string]$_ } else { '\u{0:x4}' -f [int]$_ } }) }
    $num = { param($x) ([double]$x).ToString($inv) }
    $texte = @{ EA = 'Lieferung an Kunde'; EF = 'Verbrauch Fertigung'; EM = 'Entnahme manuell'; ZS = 'Storno Lieferung' }
    $cs = [math]::Min($i0, $n - 12)
    $mJson = for ($i = $cs; $i -lt $n; $i++) {
        $d0 = $von.AddMonths($i); $mm = $mon[$i]
        $s = [math]::Max(0, $mm.EA + $mm.EF + $mm.EM - $mm.ZS)
        $det = @(foreach ($k in 'EA', 'EF', 'EM', 'ZS') { if ($mm[$k] -ne 0) { "[""$($texte[$k])"",$(& $num $(if ($k -eq 'ZS') { -$mm[$k] } else { $mm[$k] }))]" } }) -join ','
        $l = $de.DateTimeFormat.GetMonthName($d0.Month) + " " + $d0.Year
        "{""l"":""$(& $js $l)"",""k"":""$(& $js $de.DateTimeFormat.GetMonthName($d0.Month).Substring(0,1))"",""j"":$($d0.Year),""w"":$(& $num ($s / ($d0.AddMonths(1) - $d0).TotalDays * 7)),""s"":$(& $num $s),""d"":[$det]}"
    }
    $daten = "{""avg12"":$(& $num $(if ($null -eq $w12) { 0 } else { $w12 })),""ab"":$($i0 - $cs),""m"":[$($mJson -join ',')]}"
    return @"
<div class='dokkarte hkarte' id='verbrauch'>
$titel
<div class='vkacheln'>$kacheln</div>
<div id='vchart'></div>
<div id='vcharttip' class='charttip' hidden></div>
</div>
<script>
var VERB = $daten;
$VerbrauchChartJs
</script>
"@
}

# Bewegungsart -> Gruppe (Filter), Klartext, Farbe
$BewegungsArten = @{
    'ZB' = @('zu', 'Wareneingang Bestellung', 'zu'); 'ZM' = @('zu', 'Zugang manuell', 'zu'); 'SZ' = @('zu', 'Zugang manuell (schwebend)', 'zu')
    'ZR' = @('zu', 'R&uuml;cklieferung Kunde', 'zu'); 'ZS' = @('zu', 'Storno Lieferung', 'zu'); 'ZV' = @('zu', 'Inventur Zugang', 'inv')
    'EA' = @('ab', 'Lieferung an Kunde', 'ab'); 'EM' = @('ab', 'Entnahme manuell', 'ab'); 'ES' = @('ab', 'Storno Zugang', 'ab'); 'EV' = @('ab', 'Inventur Abgang', 'inv')
    'ER' = @('ab', 'R&uuml;cksendung an Lieferant', 'ab'); 'IU' = @('um', 'Umbuchung', 'um')
}
$DispoCss = @"
<style>
.badge { display:inline-block; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:600; }
.badge-bestand { background:#e8f4fd; color:#086ED5; }
.badge-vk      { background:#fde8e8; color:#b71c1c; }
.badge-mb      { background:#f3e5f5; color:#6a1b9a; }
.badge-ek      { background:#fff3e0; color:#e65100; }
.sign-plus     { color:#27ae60; font-weight:600; text-align:right; }
.sign-minus    { color:#c0392b; font-weight:600; text-align:right; }
.fazit         { padding:10px 14px; border-radius:10px; margin:0 0 12px; background:#fff; border-left:4px solid #FFBC00; box-shadow:0 1px 3px rgba(0,0,0,.08); font-size:15px; }
.warn-box      { background:#fdecea; color:#b71c1c; border-left-color:#c0392b; }
.ok-box        { background:#e8f5e9; color:#1b5e20; border-left-color:#27ae60; }
.chartkarte    { background:#fcfcfb; border-radius:8px; box-shadow:0 1px 3px rgba(0,0,0,.1); padding:14px 16px 8px; margin:0 0 16px; position:relative; }
.charttitel    { font-size:16px; font-weight:700; color:#343A40; }
.chartsub      { font-size:12px; color:#52514e; margin:2px 0 6px; }
#chart         { position:relative; height:280px; }
#chart svg     { display:block; overflow:visible; }
#chart text    { font-family:inherit; font-size:11px; fill:#898781; font-variant-numeric:tabular-nums; }
#chart text.lbl { fill:#52514e; font-size:12px; }
.charttip      { position:absolute; pointer-events:none; background:#fff; border:1px solid rgba(11,11,11,.12); border-radius:6px;
                 box-shadow:0 4px 14px rgba(0,0,0,.12); padding:8px 10px; font-size:12px; color:#52514e; min-width:180px; max-width:340px; z-index:5; }
.charttip .wert { font-size:15px; font-weight:700; color:#0b0b0b; }
.charttip .zeile { display:flex; gap:8px; justify-content:space-between; margin-top:3px; }
.charttip .zeile span:last-child { font-variant-numeric:tabular-nums; color:#0b0b0b; white-space:nowrap; }
.dokkarte      { background:#fff; border-radius:8px; box-shadow:0 1px 3px rgba(0,0,0,.1); padding:14px 16px; margin:20px 0 16px; }
.dokkarte h3   { margin:0 0 8px; font-size:17px; font-weight:700; }
.dokgruppe     { margin:10px 0 4px; }
.dokgruppe h4  { margin:0 0 6px; font-size:13px; color:#343A40; }
.dok           { font-size:13px; padding:4px 0; border-bottom:1px solid #f0f0f0; display:flex; flex-wrap:wrap; gap:6px; align-items:center; }
.dok.fehlt     { color:#999; }
.dokext        { display:inline-block; min-width:34px; text-align:center; padding:1px 5px; border-radius:4px; font-size:10px; font-weight:700; background:#eceff1; color:#455a64; }
.dokext.ext-pdf { background:#fdecea; color:#b71c1c; }
.dokmeta       { color:#888; font-size:12px; }
.dokpfad       { font-size:11px; color:#888; margin-left:auto; }
.dokbilder     { display:flex; flex-wrap:wrap; gap:10px; margin:4px 0 8px; }
.dokbild       { display:flex; flex-direction:column; width:170px; text-decoration:none; color:#333; font-size:11px; }
.dokbild img   { width:170px; height:120px; object-fit:contain; background:#f4f5f7; border:1px solid #e1e4e8; border-radius:6px; }
.dokbild span  { margin-top:3px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.dokbild:hover img { border-color:#086ED5; }
.doksprung     { font-size:12px; font-weight:400; margin-left:12px; }
.hkarte       { position:relative; }
.hlegende     { display:flex; flex-wrap:wrap; gap:6px 16px; align-items:center; font-size:12px; color:#52514e; margin:2px 0 6px; }
.hlegende span { display:inline-flex; align-items:center; gap:6px; }
.hsumme       { margin-left:auto; }
.hpunkt       { display:inline-block; width:10px; height:10px; border-radius:50%; }
.hquadrat     { display:inline-block; width:10px; height:10px; border-radius:2px; }
#hchart       { position:relative; height:320px; }
#hchart svg   { display:block; overflow:visible; }
#hchart text  { font-family:inherit; font-size:11px; fill:#898781; font-variant-numeric:tabular-nums; }
#hchart text.lbl { fill:#52514e; font-size:12px; }
#vchart       { position:relative; height:260px; }
#vchart svg   { display:block; overflow:visible; }
#vchart text  { font-family:inherit; font-size:11px; fill:#898781; font-variant-numeric:tabular-nums; }
#vchart text.lbl { fill:#52514e; font-size:12px; }
.vkacheln     { display:grid; grid-template-columns:repeat(auto-fit, minmax(140px, 1fr)); gap:10px; margin:4px 0 10px; }
.vkachel      { background:#f6f7f9; border-radius:6px; padding:8px 12px; }
.vk-l         { font-size:12px; color:#52514e; }
.vk-w         { font-size:20px; font-weight:700; color:#0b0b0b; font-variant-numeric:tabular-nums; }
.vk-w small   { font-size:12px; font-weight:400; color:#52514e; }
.vk-s         { font-size:12px; color:#52514e; }
.vw-kopf      { display:flex; align-items:center; gap:16px; flex-wrap:wrap; margin-bottom:6px; }
.vw-kopf h3   { margin:0; }
.vw-leg       { font-size:12px; color:#52514e; display:inline-flex; align-items:center; gap:6px; }
.vw-leg i     { display:inline-block; width:10px; height:10px; border-radius:2px; margin-left:8px; }
.vw-wahl      { margin-left:auto; display:inline-flex; border:1px solid #d0d4d9; border-radius:6px; overflow:hidden; }
.vw-wahl button { border:0; background:#fff; color:#52514e; font:inherit; font-size:12px; padding:3px 10px; cursor:pointer; }
.vw-wahl button + button { border-left:1px solid #d0d4d9; }
.vw-wahl button.an { background:#343A40; color:#fff; }
.vw-sum       { font-size:12px; color:#52514e; margin:2px 0 4px; }
.vw-stapel    { display:flex; gap:2px; height:22px; border-radius:4px; overflow:hidden; margin-bottom:10px; }
.vw-seg       { height:100%; display:flex; align-items:center; justify-content:center; color:#fff; font-size:11px; font-weight:700; min-width:3px; }
.vw-zeile     { display:grid; grid-template-columns:minmax(260px, 2.2fr) 3fr 80px 50px 80px; gap:10px; align-items:center; font-size:13px; padding:4px 0; border-bottom:1px solid #f0f0f0; position:relative; }
.vw-bar       { height:10px; background:#f0efec; border-radius:5px; overflow:hidden; }
.vw-bar span  { display:block; height:100%; border-radius:5px; }
.vw-m, .vw-p  { text-align:right; font-variant-numeric:tabular-nums; }
.vw-n         { color:#888; font-size:12px; }
.vw-weitere   { color:#666; cursor:default; }
.vw-weitere .vw-bar { background:none; }
/* position:fixed - ein aufgeklappter Tooltip verlaengert die Seite nicht (sonst Flackern am Seitenende) */
.vw-tip       { display:none; position:fixed; z-index:6; pointer-events:none; background:#fff; border:1px solid rgba(11,11,11,.12); border-radius:6px;
                box-shadow:0 4px 14px rgba(0,0,0,.12); padding:8px 10px; font-size:12px; color:#52514e; min-width:320px; max-width:560px; }
.vw-tip div   { display:flex; justify-content:space-between; gap:12px; margin:2px 0; }
.vw-tip b     { color:#0b0b0b; font-weight:600; white-space:nowrap; }
.charttip .zeile.fett { margin-top:6px; color:#0b0b0b; font-weight:600; }
.charttip .zeile.fett span:first-child { display:inline-flex; align-items:center; gap:6px; }
.charttip .zeile.detail { padding-left:16px; font-size:11px; margin-top:1px; }
.charttip .zeile.detail span:last-child { color:#52514e; }
.autobox { position:relative; width:100%; max-width:520px; }
.vorschlaege { position:absolute; left:0; right:0; top:100%; background:#fff; border:1px solid #ccc; border-top:none;
    border-radius:0 0 6px 6px; max-height:280px; overflow-y:auto; z-index:10; display:none; }
.vorschlaege div { padding:8px 12px; cursor:pointer; font-size:13px; border-bottom:1px solid #f0f0f0; display:flex; gap:10px; line-height:1.35; }
.vorschlaege div:hover, .vorschlaege div.sel { background:#f0f7ff; }
.vorschlaege .nr { color:#086ED5; font-weight:600; flex-shrink:0; }
/* Vorschlagsliste darf breiter als das Eingabefeld werden, lange Bezeichnungen brechen um */
.vorschlaege { right:auto; min-width:100%; width:max-content; max-width:min(720px, calc(100vw - 48px)); max-height:360px;
    border-top:1px solid #ccc; border-radius:0 6px 6px 6px; box-shadow:0 4px 12px rgba(0,0,0,.12); }
.ok { color:#27ae60; }
.badge-fehlt { display:inline-block; padding:1px 7px; border-radius:10px; font-size:11px; font-weight:700; background:#fdecea; color:#b71c1c; }
</style>
"@

# ---------- Dispo je Artikel ----------
# Bedarfe kommen aus Get-Auftraege (gleiche Zuteilung und Termine wie "Offene Auftraege" / "Materialbedarf"):
# Termin = naechster Montagetermin, sonst Liefertermin. Zugaenge = offene Bestellungen.
function JsEncode($s) {
    if ($null -eq $s) { return "" }
    return ([string]$s).Replace('\', '\\').Replace('"', '\"').Replace("`r", "").Replace("`n", " ")
}
$SQL_DispoEk = @"
SELECT b.BelID, CAST(b.Belegjahr AS varchar) + '-' + CAST(b.Belegnummer AS varchar) AS Beleg, b.A0Name1 AS Lieferant,
       p.Liefertermin, (vp.GGBestellt - vp.GGGeliefert) AS Offen
FROM KHKEKBelege b
INNER JOIN KHKEKBelegePositionen p ON p.BelID = b.BelID AND p.Mandant = b.Mandant
INNER JOIN KHKEKVorgaengePositionen vp ON vp.VorPosID = p.VorPosID AND vp.Mandant = p.Mandant
INNER JOIN KHKEKVorgaenge v ON v.VorID = vp.VorID AND v.Mandant = vp.Mandant
WHERE b.Mandant = $Mandant AND b.Belegart = 'Bestellung' AND p.Artikelnummer = @art
  AND v.GGErfuellt = 0 AND vp.GGErfuellt = 0 AND (vp.GGBestellt - vp.GGGeliefert) > 0
"@
$SQL_DispoArtikel = @"
SELECT TOP 1 a.Bezeichnung1 AS Bezeichnung, a.Stuecklistentyp, a.Lagerfuehrung, a.Artikelgruppe,
       (SELECT SUM(Lagerbestand) FROM KHKArtikelVarianten WHERE Mandant = a.Mandant AND Artikelnummer = a.Artikelnummer) AS Bestand,
       (SELECT SUM(Meldebestand) FROM KHKArtikelVarianten WHERE Mandant = a.Mandant AND Artikelnummer = a.Artikelnummer) AS Meldebestand
FROM KHKArtikel a WHERE a.Mandant = $Mandant AND a.Artikelnummer = @art
"@
$SQL_DispoKomponenten = @"
SELECT s.Element, s.Menge, a.Bezeichnung1 AS Bezeichnung, a.Lagerfuehrung
FROM KHKArtikelStueckliste s LEFT JOIN KHKArtikel a ON a.Artikelnummer = s.Element AND a.Mandant = s.Mandant
WHERE s.Mandant = $Mandant AND s.Stueckliste = @art AND s.Menge > 0
ORDER BY s.Element
"@

# Termin eines Bedarfs: Montagetermin (mit Team) oder Liefertermin
function Dispo-Termin($au) {
    if ($au.Naechster) { return @{ Datum = $au.Naechster.T.Start.Date; Montage = $true; Tr = $au.Naechster } }
    if ($au.Liefertermin) { return @{ Datum = $au.Liefertermin.Date; Montage = $false; Tr = $null } }
    return @{ Datum = $null; Montage = $false; Tr = $null }
}

# Handelsstueckliste (z.B. Speicher-Set): Bestandteile mit Bestand, offenem Bedarf und Status
function Build-Komponenten($artikel, $liste) {
    $t = Run-Query $SQL_DispoKomponenten @{ "@art" = $artikel }
    if ($t.Rows.Count -eq 0) { return "" }
    $rows = foreach ($r in $t.Rows) {
        $el = [string]$r["Element"]; $elEnc = HtmlEncode $el
        $mat = @($liste | ForEach-Object { $_.Material } | Where-Object { $_.Art -eq $el })
        $bedarf = 0.0; foreach ($m in $mat) { $bedarf += $m.Menge }
        $bst = if ($script:Bestand.ContainsKey($el)) { $script:Bestand[$el] } else { 0 }
        $ek = if ($script:EkSumme.ContainsKey($el)) { $script:EkSumme[$el] } else { 0 }
        $fehl = 0.0; foreach ($m in $mat) { if ($m.Status -eq 'fehlt') { $fehl += $m.Fehl } }
        $st = if ((ToNum $r["Lagerfuehrung"]) -ne -1) { "<span class='chip kein'>kein Lagerartikel</span>" }
              elseif ($fehl -gt 0) { "<span class='chip fehlt'>fehlt $(FmtZahl $fehl)</span>" }
              elseif (@($mat | Where-Object { $_.Status -eq 'bestellt' }).Count) { "<span class='chip bestellt'>bestellt</span>" }
              elseif ($bedarf -gt 0) { "<span class='chip da'>da</span>" } else { "" }
        "<tr><td><a class='artlink' href='/dispo?art=$elEnc'>$elEnc</a></td><td>$(HtmlEncode $r['Bezeichnung'])</td><td class='num'>$(FmtZahl $r['Menge'])</td>" +
        "<td class='num'>$(FmtZahl $bst)</td><td class='num'>$(if ($ek) { FmtZahl $ek })</td><td class='num'>$(if ($bedarf) { FmtZahl $bedarf })</td><td>$st</td></tr>"
    }
    return @"
<div class='dokkarte' id='komponenten'>
<h3>Bestandteile</h3>
<table><thead><tr><th>Artikel</th><th>Bezeichnung</th><th class='num'>je Stk</th><th class='num'>Bestand</th><th class='num'>bestellt</th><th class='num' title='Offener Bedarf aus allen Auftr&auml;gen'>Bedarf</th><th>Status</th></tr></thead>
<tbody>$($rows -join '')</tbody></table>
</div>
"@
}

# Deckungstext "bestellt": rot, wenn ohne Termin, ueberfaellig oder erst nach dem Bedarfstermin
$DispoBestelltTxt = { param($et, $dat)
    if ($et -isnot [datetime]) { return "<span class='warn'>bestellt, ohne Termin</span>" }
    $txt = "bestellt &middot; $($et.ToString('dd.MM.yyyy'))"
    if ($et.Date -lt (Get-Date).Date) { return "<span class='warn' title='Liefertermin der Bestellung &uuml;berschritten'>$txt</span>" }
    if ($dat -is [datetime] -and $et.Date -gt $dat) { return "<span class='warn' title='Wareneingang erst nach dem Termin'>$txt</span>" }
    return $txt
}

function Build-Dispo($artikel, $info, $liste, $ekTab, $fehler, $zusatz, $sprung) {
    $head = Page-Head "Solar Dispo - Artikel"
    $heute = (Get-Date).Date
    $inv = [Globalization.CultureInfo]::InvariantCulture

    if ($fehler) {
        $inhalt = "<div class='fehler'>Fehler beim Laden der Daten:<br><pre>$(HtmlEncode $fehler)</pre></div>"
    } elseif ([string]::IsNullOrEmpty($artikel)) {
        $inhalt = "<p class='leise'>Bitte oben einen Artikel suchen.</p>"
    } elseif (-not $info -or $info.Rows.Count -eq 0) {
        $inhalt = "<p class='leise'>Artikel $(HtmlEncode $artikel) nicht gefunden.</p>"
    } else {
        $ir = $info.Rows[0]
        $bez = HtmlEncode $ir["Bezeichnung"]
        $istHstl = (ToNum $ir["Stuecklistentyp"]) -eq 2
        $istLager = (ToNum $ir["Lagerfuehrung"]) -eq -1
        $bestand = ToNum $ir["Bestand"]; if ($null -eq $bestand) { $bestand = 0 }
        $melde = ToNum $ir["Meldebestand"]; if ($null -eq $melde) { $melde = 0 }

        $mitBestand = $istLager -and -not $istHstl

        # Ereignisse: Bedarfe aus den offenen Auftraegen, Zugaenge aus offenen Bestellungen
        $ev = New-Object System.Collections.ArrayList
        foreach ($au in $liste) {
            if ($istHstl) {
                # Stueckliste: je Auftragsposition eine Zeile, Status ueber alle Bestandteile
                $kit = [ordered]@{}
                foreach ($m in $au.Material) {
                    if ($m.KitArt -ne $artikel) { continue }
                    $k = "$($m.KitPos)"; if (-not $kit.Contains($k)) { $kit[$k] = @{ Menge = $m.KitMenge; Teile = (New-Object System.Collections.ArrayList) } }
                    [void]$kit[$k].Teile.Add($m)
                }
                foreach ($k in $kit.Values) {
                    $tm = Dispo-Termin $au
                    [void]$ev.Add(@{ Typ = 'VK'; Datum = $tm.Datum; Tm = $tm; Menge = [double]$k.Menge; Au = $au; Teile = $k.Teile; Beleg = $au.Beleg; Info = $au.Kunde })
                }
                continue
            }
            foreach ($m in $au.Material) {
                if ($m.Art -ne $artikel) { continue }
                $tm = Dispo-Termin $au
                [void]$ev.Add(@{ Typ = $(if ($m.Aus) { 'VS' } else { 'VK' }); Datum = $tm.Datum; Tm = $tm; Menge = [double]$m.Menge; Au = $au; M = $m
                                 Beleg = $au.Beleg; Info = $au.Kunde + $(if ($m.Aus) { " (aus $($m.Aus))" } else { "" }) })
            }
        }
        if ($ekTab) {
            foreach ($r in $ekTab.Rows) {
                $lt = $r["Liefertermin"]
                [void]$ev.Add(@{ Typ = 'EK'; Datum = $(if ($lt -is [datetime]) { $lt.Date } else { $null }); Menge = [double](ToNum $r["Offen"]); Beleg = [string]$r["Beleg"]; BelID = [string]$r["BelID"]; Info = [string]$r["Lieferant"] })
            }
        }
        $sortiert = @($ev | Sort-Object @{ e = { if ($_.Datum -is [datetime]) { $_.Datum } else { [datetime]::MaxValue } } }, @{ e = { if ($_.Typ -eq 'EK') { 0 } else { 1 } } })

        $kumuliert = $bestand
        $fehltAb = $null; $minimum = $bestand; $minDatum = $null
        $unterMeldeAb = $null; if ($melde -gt 0 -and $bestand -lt $melde) { $unterMeldeAb = "heute" }
        $punkte = New-Object System.Collections.Specialized.OrderedDictionary
        $heuteKey = $heute.ToString("yyyy-MM-dd")
        $punkte[$heuteKey] = @{ V = $bestand; Ev = [System.Collections.ArrayList]::new() }
        [void]$punkte[$heuteKey].Ev.Add("{""t"":""Bestand"",""b"":"""",""i"":"""",""m"":$(([double]$bestand).ToString($inv))}")
        $ohneTermin = 0
        $bCss = if ($bestand -lt 0) { " class='num warn'" } else { " class='num'" }
        $tRows = @()
        if ($mitBestand) { $tRows += "<tr><td><b>Heute</b></td><td><span class='badge badge-bestand'>Bestand</span></td><td></td><td></td><td class='num'>$(FmtZahl $bestand)</td><td$bCss>$(FmtZahl $kumuliert)</td><td></td></tr>" }

        foreach ($e in $sortiert) {
            $zugang = ($e.Typ -eq 'EK'); $m = $e.Menge
            $kumuliert += $(if ($zugang) { $m } else { -$m })
            $dat = $e.Datum
            $datTxt = if ($dat -is [datetime]) { "$(Datum-Kurz $dat) $($dat.Year)" } else { "(kein Termin)" }
            if ($dat -is [datetime] -and $dat -lt $heute) { $datTxt = "<span class='warn' title='Termin &uuml;berschritten, noch offen'>$datTxt</span>" }
            $datFazit = if ($dat -is [datetime]) { $dat.ToString('dd.MM.yyyy') } else { "(kein Termin)" }
            if ($mitBestand) {
                if ($kumuliert -lt 0 -and $null -eq $fehltAb) { $fehltAb = $datFazit }
                if ($kumuliert -lt $minimum) { $minimum = $kumuliert; $minDatum = $datFazit }
                if ($melde -gt 0 -and $kumuliert -lt $melde -and $null -eq $unterMeldeAb) { $unterMeldeAb = $datFazit }
            }
            # Terminspalte: Montagetermin mit Team, sonst Liefertermin
            $terminZelle = $datTxt
            if ($e.Tm -and $e.Tm.Montage) {
                $teams = ($e.Tm.Tr.T.Team | ForEach-Object { "<span class='team'>$(HtmlEncode $_)</span>" }) -join ''
                $u = if (-not $e.Tm.Tr.Sicher) { " <span class='unsicher' title='Nur der Nachname passt &ndash; bitte pr&uuml;fen: $(HtmlEncode $e.Tm.Tr.T.Titel)'>?</span>" } else { "" }
                $terminZelle = "<span title='Montage: $(HtmlEncode $e.Tm.Tr.T.Titel)'>$datTxt</span>$u $teams"
            } elseif ($e.Typ -ne 'EK' -and $dat -is [datetime]) {
                $terminZelle = "$datTxt <span class='leise'>Liefertermin</span>"
            }

            if ($dat -is [datetime]) {
                $pk = $(if ($dat -lt $heute) { $heute } else { $dat }).ToString("yyyy-MM-dd")
                if (-not $punkte.Contains($pk)) { $punkte[$pk] = @{ V = 0; Ev = [System.Collections.ArrayList]::new() } }
                $punkte[$pk].V = $kumuliert
                $tName = switch ($e.Typ) { 'VK' { 'Auftrag' } 'VS' { 'Auftrag (Stueckl.)' } default { 'Bestellung' } }
                $ueberf = $(if ($dat -lt $heute) { " (ueberf. " + $dat.ToString("dd.MM.") + ")" } else { "" })
                $mSigned = $(if ($zugang) { $m } else { -$m })
                [void]$punkte[$pk].Ev.Add("{""t"":""$tName"",""b"":""$(JsEncode $e.Beleg)$ueberf"",""i"":""$(JsEncode $e.Info)"",""m"":$(([double]$mSigned).ToString($inv))}")
            } else { $ohneTermin++ }

            switch ($e.Typ) {
                'VK' { $badge = "<span class='$(if ($e.Au.Art -like '*RES*') { 'art res' } else { 'art' })'>$($e.Au.Art)</span>" }
                'VS' { $badge = "<span class='$(if ($e.Au.Art -like '*RES*') { 'art res' } else { 'art' })' title='Bestandteil einer St&uuml;ckliste'>$($e.Au.Art) St&uuml;ckl.</span>" }
                default { $badge = "<span class='badge badge-ek'>Bestellung</span>" }
            }
            $belegZelle = if ($e.Typ -eq 'EK') { "<a class='artlink' href='/ek?bel=$($e.BelID)'>$(HtmlEncode $e.Beleg)</a>" } else { "<a class='artlink' href='/auftrag?bel=$($e.Au.BelID)'>$(HtmlEncode $e.Beleg)</a>" }
            $mCss = if ($zugang) { " class='num sign-plus'" } else { " class='num sign-minus'" }
            $mTxt = if ($zugang) { "+" + (FmtZahl $m) } else { "&minus;" + (FmtZahl $m) }
            $kCss = if ($kumuliert -lt 0) { " class='num warn'" } else { " class='num'" }

            # Deckung wie auf "Offene Auftraege" (Zuteilung nach Montage-/Liefertermin ueber alle Auftraege)
            $deckTxt = ""
            if ($e.Teile) {
                # Stueckliste: schlechtester Status der Bestandteile
                $fehlen = @($e.Teile | Where-Object { $_.Status -eq 'fehlt' })
                $best = @($e.Teile | Where-Object { $_.Status -eq 'bestellt' })
                if ($fehlen.Count) {
                    $tip = HtmlEncode (($fehlen | ForEach-Object { "$($_.Art) $($_.Bez): fehlt $(FmtZahl $_.Fehl)" }) -join "`n")
                    $deckTxt = "<span class='badge-fehlt' title='$tip'>$(if ($fehlen.Count -eq $e.Teile.Count) { 'fehlt' } else { "$($fehlen.Count) von $($e.Teile.Count) Teilen fehlen" })</span>"
                } elseif ($best.Count) {
                    $et = @($best | ForEach-Object { $_.EkTermin } | Where-Object { $_ -is [datetime] } | Sort-Object -Descending)
                    $deckTxt = & $DispoBestelltTxt $(if ($et.Count) { $et[0] } else { $null }) $dat
                } elseif (@($e.Teile | Where-Object { $_.Status -eq 'da' }).Count) { $deckTxt = "<span class='ok'>ab Lager</span>" }
            } elseif (-not $zugang) {
                $deckTxt = switch ($e.M.Status) {
                    'da' { "<span class='ok'>ab Lager</span>" }
                    'bestellt' { & $DispoBestelltTxt $e.M.EkTermin $dat }
                    'fehlt' { "<span class='badge-fehlt'>fehlt $(FmtZahl $e.M.Fehl)</span>" }
                    default { "" }
                }
            }
            $kumZelle = if ($mitBestand) { "<td$kCss>$(FmtZahl $kumuliert)</td>" } else { "" }
            $deckZelle = if ($istLager -or $istHstl) { "<td>$deckTxt</td>" } else { "" }
            $tRows += "<tr><td style='white-space:nowrap'>$terminZelle</td><td style='white-space:nowrap'>$badge</td><td style='white-space:nowrap'>$belegZelle</td><td class='leise'>$(HtmlEncode $e.Info)</td><td$mCss>$mTxt</td>$kumZelle$deckZelle</tr>"
        }

        $meldeTxt = ""
        if ($istLager -and -not $istHstl -and $melde -gt 0 -and $unterMeldeAb) {
            $meldeTxt = if ($unterMeldeAb -eq "heute") { " &middot; <b>unter Meldebestand</b> ($(FmtZahl $melde))" }
                        else { " &middot; unter Meldebestand ($(FmtZahl $melde)) ab <b>$unterMeldeAb</b>" }
        }
        $bedarfe = @($sortiert | Where-Object { $_.Typ -ne 'EK' })
        if ($istHstl) {
            $summe = 0.0; foreach ($b in $bedarfe) { $summe += $b.Menge }
            $nFehlt = @($bedarfe | Where-Object { @($_.Teile | Where-Object { $_.Status -eq 'fehlt' }).Count }).Count
            $fazit = if (-not $bedarfe.Count) { "<div class='fazit'>St&uuml;ckliste &middot; keine offenen Auftr&auml;ge</div>" }
                     elseif ($nFehlt) { "<div class='fazit warn-box'>St&uuml;ckliste &middot; offen <b>$(FmtZahl $summe)</b> &middot; bei <b>$nFehlt</b> von $($bedarfe.Count) Auftr&auml;gen fehlen Bestandteile</div>" }
                     else { "<div class='fazit ok-box'>St&uuml;ckliste &middot; offen <b>$(FmtZahl $summe)</b> &middot; Bestandteile f&uuml;r alle $($bedarfe.Count) Auftr&auml;ge da oder bestellt</div>" }
        } elseif (-not $istLager) {
            $summe = 0.0; foreach ($b in $bedarfe) { $summe += $b.Menge }
            $fazit = if ($bedarfe.Count) { "<div class='fazit'>Kein Lagerartikel &ndash; wird je Auftrag beschafft &middot; offen <b>$(FmtZahl $summe)</b> in $(if ($bedarfe.Count -eq 1) { '1 Auftrag' } else { "$($bedarfe.Count) Auftr&auml;gen" })</div>" }
                     else { "<div class='fazit'>Kein Lagerartikel &middot; keine offenen Auftr&auml;ge</div>" }
        } elseif ($null -ne $fehltAb) {
            $fazit = "<div class='fazit warn-box'><b>Fehlt ab $fehltAb</b> &middot; tiefster Stand <b>$(FmtZahl $minimum)</b> am $minDatum$meldeTxt</div>"
        } elseif ($sortiert.Count -gt 0) {
            $fazit = "<div class='fazit ok-box'>Alles gedeckt &middot; tiefster Stand <b>$(FmtZahl $minimum)</b>$(if ($minDatum) { " am $minDatum" })$meldeTxt</div>"
        } else {
            $fazit = "<div class='fazit'>Keine offenen Auftr&auml;ge oder Bestellungen$meldeTxt</div>"
        }

        $chartHtml = ""
        if ($istLager -and -not $istHstl -and $sortiert.Count -gt 0) {
            $pJson = ($punkte.Keys | ForEach-Object { "{""d"":""$_"",""v"":$(([double]$punkte[$_].V).ToString($inv)),""ev"":[$($punkte[$_].Ev -join ',')]}" }) -join ','
            $chartDaten = ("{""melde"":$(([double]$melde).ToString($inv)),""p"":[$pJson]}") -replace '<', '<'
            $ohneTxt = if ($ohneTermin -gt 0) { "$ohneTermin ohne Termin nicht im Diagramm" } else { "" }
            $chartHtml = @"
<div class='chartkarte'>
  <div class='charttitel'>Verf&uuml;gbarer Bestand im Zeitverlauf</div>
$(if ($ohneTxt) { "  <div class='chartsub'>$ohneTxt</div>" })
  <div id='chart'></div>
  <div id='charttip' class='charttip' hidden></div>
</div>
<script>
var DISPO = $chartDaten;
$DispoChartJs
</script>
"@
        }
        $tabelle = if ($sortiert.Count -gt 0 -or $mitBestand) { @"
<table id='tab'>
<thead><tr>
  <th>Termin</th><th>Typ</th><th>Beleg</th><th>Kunde / Lieferant</th>
  <th class='num'>Menge</th>$(if ($mitBestand) { "<th class='num'>Kumuliert</th>" })$(if ($istLager -or $istHstl) { "<th title='Bestand und Bestellungen gehen an den Auftrag mit dem fr&uuml;hesten Montage- bzw. Liefertermin'>Deckung</th>" })
</tr></thead>
<tbody>$($tRows -join '')</tbody>
</table>
"@ } else { "" }
        $inhalt = @"
<h2 style='font-size:16px;margin:4px 0 12px'>Artikel $(HtmlEncode $artikel) &ndash; $bez$sprung</h2>
$fazit
$chartHtml
$tabelle
$zusatz
"@
    }

    return @"
$head
$DispoCss
<div class='kopf'><h1>Dispositions&uuml;bersicht</h1><div class='meta'>$(StandText)</div></div>
<div class='steuer'>
  <button class='btn btn2' onclick='history.back()'>&larr; zur&uuml;ck</button>
  <div class='autobox'>
    <input class='suche' id='artsuche' autocomplete='off' placeholder='Artikel suchen ...' oninput='vorschlage()' onkeydown='navi(event)'>
    <div class='vorschlaege' id='vorschlaege'></div>
  </div>
</div>
$(if ($script:KalFehler) { "<div class='hinweisbox'>$(HtmlEncode $script:KalFehler)</div>" })
<script>
$ArtikelsucheJs
</script>
$inhalt
</body></html>
"@
}

# Artikelsuche mit Vorschlaegen aus /api/artikelsuche; Auswahl oeffnet /dispo?art=
$ArtikelsucheJs = @'
var sTimer = null, selIdx = -1;
function vorschlage() {
  var q = document.getElementById('artsuche').value.trim();
  if (sTimer) clearTimeout(sTimer);
  if (q.length < 2) { schliesse(); return; }
  sTimer = setTimeout(function () {
    fetch('/api/artikelsuche?q=' + encodeURIComponent(q)).then(function (r) { return r.json(); }).then(zeige).catch(schliesse);
  }, 200);
}
function zeige(daten) {
  selIdx = -1;
  var box = document.getElementById('vorschlaege');
  box.textContent = '';
  if (!daten || daten.length === 0) { schliesse(); return; }
  daten.forEach(function (d) {
    var div = document.createElement('div'), nr = document.createElement('span');
    div.setAttribute('data-nr', d.nr);
    nr.className = 'nr'; nr.textContent = d.nr;
    div.appendChild(nr); div.appendChild(document.createTextNode(d.bez));
    div.onclick = function () { waehle(d.nr); };
    box.appendChild(div);
  });
  box.style.display = 'block';
}
function schliesse() { var box = document.getElementById('vorschlaege'); box.style.display = 'none'; box.textContent = ''; selIdx = -1; }
function waehle(nr) { window.location = '/dispo?art=' + encodeURIComponent(nr); }
function navi(e) {
  var items = document.getElementById('vorschlaege').querySelectorAll('div');
  if (e.key === 'Enter') {
    e.preventDefault();
    if (selIdx >= 0 && items[selIdx]) { waehle(items[selIdx].getAttribute('data-nr')); }
    else { var q = document.getElementById('artsuche').value.trim(); if (q) waehle(q); }
  } else if (e.key === 'ArrowDown') { e.preventDefault(); if (selIdx < items.length - 1) selIdx++; markiere(items); }
  else if (e.key === 'ArrowUp') { e.preventDefault(); if (selIdx > 0) selIdx--; markiere(items); }
  else if (e.key === 'Escape') { schliesse(); }
}
function markiere(items) { items.forEach(function (it, i) { it.className = (i === selIdx) ? 'sel' : ''; }); }
document.addEventListener('click', function (e) { if (!e.target.closest('.autobox')) schliesse(); });
'@

# ---------- Verwendung: an welche Kunden ging der Artikel (3/6/12 Monate) ----------
# EA Lieferung an Kunde (Referenz = Lieferschein "Jahr-Nummer" -> Kunde), ZS Storno Lieferung (abgezogen), EM manuelle Entnahme.
# Dazu die Stuecklisten, in denen der Artikel steckt.
$SQL_VerwendungIst = @"
SELECT b.Bewegungsdatum AS Datum, b.Bewegungsart AS BA, b.MengeBasis AS Menge, b.Referenz, j.Benutzer, kd.Kunde
FROM KHKLagerplatzbuchungen b
INNER JOIN KHKLagerplatzbuchungenJobs j ON j.Job = b.Job AND j.Mandant = b.Mandant
OUTER APPLY (SELECT TOP 1 v.A0Name1 AS Kunde FROM KHKVKBelege v
             WHERE b.Referenz LIKE '[0-9][0-9][0-9][0-9]-%' AND v.Mandant = b.Mandant
               AND v.Belegjahr = TRY_CAST(LEFT(b.Referenz, 4) AS int) AND v.Belegnummer = TRY_CAST(SUBSTRING(b.Referenz, 6, 20) AS int)
             ORDER BY CASE v.Belegkennzeichen WHEN 'VLL' THEN 0 WHEN 'VFR' THEN 1 ELSE 2 END) kd
WHERE b.Mandant = $Mandant AND b.Artikelnummer = @art AND b.Bewegungsart IN ('EA', 'EM', 'ZS')
  AND b.MengeBasis <> 0 AND b.Bewegungsdatum >= @von
"@
$SQL_VerwendungStueckliste = @"
SELECT s.Stueckliste, s.Menge, a.Bezeichnung1 AS Bezeichnung
FROM KHKArtikelStueckliste s INNER JOIN KHKArtikel a ON a.Artikelnummer = s.Stueckliste AND a.Mandant = s.Mandant AND a.Aktiv = -1
WHERE s.Mandant = $Mandant AND s.Element = @art AND s.Menge > 0
ORDER BY s.Stueckliste
"@
$VerwendungZeitraeume = @(3, 6, 12)
$VerwendungFarben = @{ K = '#1baf7a'; M = '#eb6834' }
$VerwendungMaxKunden = 12

function Build-Verwendung($artikel) {
    $de = [Globalization.CultureInfo]'de-DE'; $inv = [Globalization.CultureInfo]::InvariantCulture
    $heute = (Get-Date).Date
    $ist = Run-Query $SQL_VerwendungIst @{ "@art" = $artikel; "@von" = $heute.AddMonths(-($VerwendungZeitraeume | Measure-Object -Maximum).Maximum) }
    $stl = Run-Query $SQL_VerwendungStueckliste @{ "@art" = $artikel }
    if ($ist.Rows.Count -eq 0 -and $stl.Rows.Count -eq 0) { return "" }

    $zahl = { param($x) ([math]::Round([double]$x)).ToString('N0', $de) }
    $stlZeile = ""
    if ($stl.Rows.Count) {
        $liste = @($stl.Rows | ForEach-Object { "<div><span>$(HtmlEncode $_['Stueckliste']) $(HtmlEncode $_['Bezeichnung'])</span><b>$(FmtZahl $_['Menge']) je Stk</b></div>" })
        $stlZeile = "<div class='vw-zeile vw-weitere'><div class='vw-name'>Enthalten in St&uuml;cklisten ($($stl.Rows.Count))</div><div class='vw-bar'></div><div class='vw-m'></div><div class='vw-p'></div><div class='vw-n'>&#9432;</div><div class='vw-tip'>$($liste -join '')</div></div>"
    }
    $perioden = foreach ($mon in $VerwendungZeitraeume) {
        $von = $heute.AddMonths(-$mon)
        # Gruppen: K|<Kunde>, M (manuell)
        $gr = [ordered]@{}
        foreach ($r in $ist.Rows) {
            if ($r["Datum"] -lt $von) { continue }
            $ba = [string]$r["BA"]; $m = [double]$r["Menge"]; if ($ba -eq 'ZS') { $m = -$m }
            $datum = ([datetime]$r['Datum']).ToString('dd.MM.yyyy')
            if ($ba -eq 'EM') {
                $key = "M"; $bez = "Manuelle Entnahme"; $detKey = [string]$r["Benutzer"]; $detTxt = HtmlEncode $r["Benutzer"]
            } else {
                $k = if ($r["Kunde"] -isnot [System.DBNull]) { [string]$r["Kunde"] } else { "Beleg $($r['Referenz'])" }
                $key = "K|$k"; $bez = $k; $detKey = [string]$r["Referenz"]; $detTxt = "Lieferschein $(HtmlEncode $r['Referenz']) &middot; $datum"
            }
            if (-not $gr.Contains($key)) { $gr[$key] = @{ Bez = $bez; Menge = 0.0; Det = [ordered]@{} } }
            $g = $gr[$key]; $g.Menge += $m
            if (-not $g.Det.Contains($detKey)) { $g.Det[$detKey] = @{ Txt = $detTxt; M = 0.0 } }
            $g.Det[$detKey].M += $m
        }
        $gruppen = @($gr.Keys | Where-Object { $gr[$_].Menge -gt 0 } | Sort-Object { -$gr[$_].Menge })
        $gesamt = 0.0; foreach ($k in $gruppen) { $gesamt += $gr[$k].Menge }
        $kunden = @($gruppen | Where-Object { $_ -like 'K|*' })
        # zu viele Kunden: die kleinen zu einer Zeile zusammenfassen
        $zeigen = @($gruppen | Where-Object { $_ -eq 'M' -or ($kunden.IndexOf($_) -ge 0 -and $kunden.IndexOf($_) -lt $VerwendungMaxKunden) })
        $rest = @($kunden | Select-Object -Skip $VerwendungMaxKunden)
        $maxM = if ($gruppen.Count) { $gr[$gruppen[0]].Menge } else { 1 }
        $zeile = { param($name, $menge, $typ, $anz, $tip)
            $bar = if ($VerwendungFarben.ContainsKey($typ)) { "<span style='width:$([math]::Max(0.5, 100 * $menge / $maxM).ToString('0.###', $inv))%;background:$($VerwendungFarben[$typ])'></span>" } else { "" }
            "<div class='vw-zeile$(if (-not $bar) { ' vw-weitere' })'><div class='vw-name'>$name</div><div class='vw-bar'>$bar</div>" +
            "<div class='vw-m'>$(& $zahl $menge)</div><div class='vw-p'>$([math]::Round(100 * $menge / $gesamt).ToString('N0', $de)) %</div><div class='vw-n'>$anz</div><div class='vw-tip'>$tip</div></div>" }
        $zeilen = foreach ($k in $zeigen) {
            $g = $gr[$k]; $typ = $k.Substring(0, 1)
            $det = @($g.Det.Values | Where-Object { $_.M -ne 0 } | Sort-Object { -$_.M } | Select-Object -First 15 | ForEach-Object { "<div><span>$($_.Txt)</span><b>$(& $zahl $_.M)</b></div>" })
            $anz = if ($typ -eq 'K') { "$($g.Det.Count) LS" } else { "$($g.Det.Count) Pers." }
            & $zeile (HtmlEncode $g.Bez) $g.Menge $typ $anz ($det -join '')
        }
        if ($rest.Count) {
            $rm = 0.0; foreach ($k in $rest) { $rm += $gr[$k].Menge }
            $tip = @($rest | Select-Object -First 30 | ForEach-Object { "<div><span>$(HtmlEncode $gr[$_].Bez)</span><b>$(& $zahl $gr[$_].Menge)</b></div>" }) -join ''
            $zeilen = @($zeilen) + (& $zeile "weitere $($rest.Count) Kunden" $rm 'R' '&#9432;' $tip)
        }
        $inhalt = if ($gruppen.Count) {
            "<div class='vw-sum'><b>$(& $zahl $gesamt)</b> Stk$(if ($kunden.Count) { " an $($kunden.Count) Kunde$(if ($kunden.Count -ne 1) { 'n' })" })</div>$($zeilen -join '')$stlZeile"
        } else { "<p class='leise' style='margin:6px 0'>Keine Entnahmen in den letzten $mon Monaten.</p>$stlZeile" }
        "<div class='vw-periode' data-m='$mon'$(if ($mon -ne $VerwendungZeitraeume[0]) { " hidden" })>$inhalt</div>"
    }
    $knoepfe = ($VerwendungZeitraeume | ForEach-Object { "<button type='button' data-m='$_'$(if ($_ -eq $VerwendungZeitraeume[0]) { " class='an'" })>$_ Monate</button>" }) -join ''
    return @"
<div class='dokkarte' id='verwendung'>
<div class='vw-kopf'><h3>Verwendung</h3>
  <span class='vw-leg'><i style='background:$($VerwendungFarben.K)'></i>Lieferung an Kunde <i style='background:$($VerwendungFarben.M)'></i>Manuell</span>
  <span class='vw-wahl'>$knoepfe</span></div>
$($perioden -join '')
</div>
<script>
(function () {
  var karte = document.getElementById('verwendung'); if (!karte) return;
  function zeige(m) {
    karte.querySelectorAll('.vw-periode').forEach(function (p) { p.hidden = p.getAttribute('data-m') !== m; });
    karte.querySelectorAll('.vw-wahl button').forEach(function (b) { b.className = b.getAttribute('data-m') === m ? 'an' : ''; });
  }
  karte.querySelectorAll('.vw-wahl button').forEach(function (b) {
    b.addEventListener('click', function () { var m = b.getAttribute('data-m'); zeige(m); try { localStorage.setItem('vwMonate', m); } catch (e) {} });
  });
  // Tooltip unter der Zeile, bei Platzmangel darueber; seitlich im Fenster halten
  karte.querySelectorAll('.vw-zeile').forEach(function (z) {
    var tip = z.querySelector('.vw-tip'); if (!tip || !tip.innerHTML) return;
    z.addEventListener('mouseenter', function () {
      tip.style.display = 'block';
      var r = z.getBoundingClientRect(), h = tip.offsetHeight, w = tip.offsetWidth;
      var top = r.bottom + 4; if (top + h > window.innerHeight - 8) top = Math.max(8, r.top - h - 4);
      var left = Math.min(r.left + r.width * 0.3, window.innerWidth - w - 8);
      tip.style.top = top + 'px'; tip.style.left = Math.max(8, left) + 'px';
    });
    z.addEventListener('mouseleave', function () { tip.style.display = ''; });
  });
  window.addEventListener('scroll', function () { karte.querySelectorAll('.vw-tip').forEach(function (t) { t.style.display = ''; }); }, { passive: true });
  try { var gemerkt = localStorage.getItem('vwMonate'); if (gemerkt && karte.querySelector(".vw-periode[data-m='" + gemerkt + "']")) zeige(gemerkt); } catch (e) {}
})();
</script>
"@
}


# Komplette Dispo-Seite eines Artikels (Route /dispo)
function Dispo-Seite($art) {
    try {
        if (-not $art) { return Build-Dispo "" $null @() $null $null "" "" }
        else {
            $info = Run-Query $SQL_DispoArtikel @{ "@art" = $art }
            $liste = @(Get-Auftraege)
            $ekTab = Run-Query $SQL_DispoEk @{ "@art" = $art }
            $artBestand = if ($info.Rows.Count) { $info.Rows[0]["Bestand"] } else { 0 }
            # Zusatzbereiche einzeln abfangen - ein Fehler dort soll die Dispo nicht verhindern
            $zus = ""; $sprung = ""
            $teile = @(
                @{ Id = 'komponenten'; Name = 'Bestandteile'; Bau = { Build-Komponenten $art $liste } },
                @{ Id = 'dokumente'; Name = 'Dokumente'; Bau = { $d = Build-Dokumente $art; if ($d.Anzahl -gt 0) { $d.Html } } },
                @{ Id = 'historie'; Name = 'Bestandsverlauf'; Lager = $true; Bau = { Build-BestandHistorie $art $artBestand } },
                @{ Id = 'verbrauch'; Name = 'Verbrauch'; Lager = $true; Bau = { Build-Verbrauch $art $artBestand } },
                @{ Id = 'verwendung'; Name = 'Verwendung'; Bau = { Build-Verwendung $art } }
            )
            if ($info.Rows.Count) {
                # Bestandsverlauf und Verbrauch nur fuer Lagerartikel (nicht fuer Stuecklisten)
                $mitLager = (ToNum $info.Rows[0]["Lagerfuehrung"]) -eq -1 -and (ToNum $info.Rows[0]["Stuecklistentyp"]) -ne 2
                foreach ($t in $teile) {
                    if ($t.Lager -and -not $mitLager) { continue }
                    try { $h = & $t.Bau; if ($h) { $zus += $h; $sprung += "<a class='artlink doksprung' href='#$($t.Id)'>$($t.Name) &darr;</a>" } }
                    catch { $zus += "<div class='fehler'>$($t.Name) konnte nicht geladen werden: $(HtmlEncode $_.Exception.Message)</div>" }
                }
            }
            return Build-Dispo $art $info $liste $ekTab $null $zus $sprung
        }
    } catch { return Build-Dispo $art $null @() $null ($_.Exception.ToString()) "" "" }
}


# ---------- Lagerbewegungsprotokoll (aus der Oezer-Liste) ----------
# ---- SQL: Lagerbewegungsprotokoll ----
# Jede Lagerplatzbuchung gehoert zu einem Job; dort steht der Sage-Benutzer (Kuerzel wie "AKr"
# oder bei Scannerbuchungen "MDE <Nachname>"). Sage speichert nur das Datum - eine Uhrzeit gibt
# es lediglich fuer MDE-Kommissionierbuchungen (IaswebdeMdeKommissionierbuchungen).
# Buchungen mit Menge 0 sind interne Gegenbuchungen ohne Bestandswirkung.
$SQL_Bewegungen = @"
SELECT TOP (@max)
    b.BuchungsID, b.Bewegungsdatum, b.Bewegungsart, b.Artikelnummer, a.Bezeichnung1 AS Bezeichnung,
    b.MengeBasis AS Menge, b.Referenz, j.Standardtext, j.Memo, j.Benutzer,
    hp.Lagerkennung AS HKennung, hp.Dimension1 AS HD1, hp.Dimension2 AS HD2, hp.Dimension3 AS HD3,
    zp.Lagerkennung AS ZKennung, zp.Dimension1 AS ZD1, zp.Dimension2 AS ZD2, zp.Dimension3 AS ZD3,
    mde.Zeit, kd.Name AS Kunde, lf.Name AS Lieferant
FROM KHKLagerplatzbuchungen b
INNER JOIN KHKLagerplatzbuchungenJobs j ON j.Job = b.Job AND j.Mandant = b.Mandant
LEFT JOIN KHKArtikel a       ON a.Artikelnummer = b.Artikelnummer AND a.Mandant = b.Mandant
LEFT JOIN KHKLagerplaetze hp ON hp.PlatzID = b.HerkunftsLPKennung AND hp.Mandant = b.Mandant
LEFT JOIN KHKLagerplaetze zp ON zp.PlatzID = b.ZielLPKennung AND zp.Mandant = b.Mandant
LEFT JOIN (SELECT BuchungsId, MIN([Timestamp]) AS Zeit FROM IaswebdeMdeKommissionierbuchungen
           WHERE Mandant = $Mandant GROUP BY BuchungsId) mde ON mde.BuchungsId = b.BuchungsID
OUTER APPLY (SELECT TOP 1 v.A0Name1 AS Name FROM KHKVKBelege v
             WHERE b.Bewegungsart IN ('EA', 'ZS', 'ZR', 'SR') AND b.Referenz LIKE '[0-9][0-9][0-9][0-9]-%' AND v.Mandant = b.Mandant
               AND v.Belegjahr = TRY_CAST(LEFT(b.Referenz, 4) AS int) AND v.Belegnummer = TRY_CAST(SUBSTRING(b.Referenz, 6, 20) AS int)
             ORDER BY CASE v.Belegkennzeichen WHEN 'VLL' THEN 0 WHEN 'VFR' THEN 1 ELSE 2 END) kd
OUTER APPLY (SELECT TOP 1 e.A0Name1 AS Name FROM KHKEKBelege e
             WHERE b.Bewegungsart IN ('ZB', 'SB', 'ES', 'ER') AND b.Referenz LIKE '[0-9][0-9][0-9][0-9]-%' AND e.Mandant = b.Mandant
               AND e.Belegjahr = TRY_CAST(LEFT(b.Referenz, 4) AS int) AND e.Belegnummer = TRY_CAST(SUBSTRING(b.Referenz, 6, 20) AS int)) lf
WHERE b.Mandant = $Mandant AND b.MengeBasis <> 0
  AND b.Bewegungsdatum >= @von AND b.Bewegungsdatum < @bis
  {FILTER}
ORDER BY b.Bewegungsdatum DESC, b.BuchungsID DESC
"@

# ---- SQL: Benutzer mit Lagerbuchungen (fuer die Personenauswahl) und Klarnamen ----
$SQL_BewegungenBenutzer = @"
SELECT DISTINCT j.Benutzer FROM KHKLagerplatzbuchungenJobs j
WHERE j.Mandant = $Mandant AND j.Erfassungsdatum > DATEADD(day, -365, GETDATE()) AND ISNULL(j.Benutzer, '') <> ''
"@
$SQL_MitarbeiterNamen = @"
SELECT UserName, Vorname, Nachname FROM KHKMitarbeiter WHERE Mandant = $Mandant
"@

# Sage-Benutzer -> lesbarer Name. Kuerzel ueber KHKMitarbeiter.UserName, Scannerbenutzer
# "MDE <Nachname>" ueber einen eindeutigen Nachnamen.
function Get-BenutzerNamen($benutzerTab, $maTab) {
    $perUser = @{}; $perNachname = @{}
    if ($maTab) {
        foreach ($m in $maTab.Rows) {
            $voll = ("$($m['Vorname']) $($m['Nachname'])").Trim()
            if ($voll -eq '') { continue }
            $u = [string]$m["UserName"]; if ($u -ne '') { $perUser[$u] = $voll }
            $n = ([string]$m["Nachname"]).Trim()
            if ($n -ne '') { if ($perNachname.ContainsKey($n)) { $perNachname[$n] = $null } else { $perNachname[$n] = $voll } }
        }
    }
    $namen = @{}
    if ($benutzerTab) {
        foreach ($r in $benutzerTab.Rows) {
            $b = [string]$r["Benutzer"]
            if ($perUser.ContainsKey($b)) { $namen[$b] = @{ Name = $perUser[$b]; Scanner = $false } }
            elseif ($b -match '^MDE\s+(.+)$') {
                $nn = $matches[1].Trim()
                $voll = if ($perNachname.ContainsKey($nn) -and $perNachname[$nn]) { $perNachname[$nn] } else { $nn }
                $namen[$b] = @{ Name = $voll; Scanner = $true }
            }
            elseif ($b -eq 'mde') { $namen[$b] = @{ Name = 'Scanner (allgemein)'; Scanner = $true } }
            else { $namen[$b] = @{ Name = $b; Scanner = $false } }
        }
    }
    return $namen
}

function Fmt-Platz($kennung, $d1, $d2, $d3) {
    if ($kennung -is [System.DBNull] -or $null -eq $kennung) { return "" }
    $a = if ($d1 -is [System.DBNull]) { 0 } else { [int]$d1 }
    $b = if ($d2 -is [System.DBNull]) { 0 } else { [int]$d2 }
    $c = if ($d3 -is [System.DBNull]) { 0 } else { [int]$d3 }
    # Nullen am Ende weglassen: "Regal A 1/2/3", "IASKommlag 7", "Kran H1"
    if ($c -ne 0) { return "$kennung $a/$b/$c" }
    if ($b -ne 0) { return "$kennung $a/$b" }
    if ($a -ne 0) { return "$kennung $a" }
    return [string]$kennung
}

# Lagerplatz-Filter der Bewegungen: "Regal A" = ganzes Lager (Kennung beginnt so),
# "Regal A;1;2;3" = genau dieser Platz (Lagerkennung + Dimension 1-3)
function Platz-Wert($kennung, $d1, $d2, $d3) {
    $z = foreach ($d in @($d1, $d2, $d3)) { if ($d -is [System.DBNull] -or $null -eq $d) { 0 } else { [int]$d } }
    return "$kennung;" + ($z -join ';')
}
function Platz-Text($wert) {
    if ([string]$wert -notlike '*;*') { return [string]$wert }
    $t = ([string]$wert).Split(';')
    $z = for ($i = 1; $i -le 3; $i++) { $n = 0; if ($t.Count -gt $i) { [void][int]::TryParse($t[$i], [ref]$n) }; $n }
    return Fmt-Platz $t[0] $z[0] $z[1] $z[2]
}
function Bew-Url($f, $aenderung) {
    $p = [ordered]@{ von = $f.von.ToString('yyyy-MM-dd'); bis = $f.bis.ToString('yyyy-MM-dd'); typ = $f.typ; wer = $f.wer; art = $f.art; platz = $f.platz }
    foreach ($k in $aenderung.Keys) { $p[$k] = $aenderung[$k] }
    $teile = foreach ($k in $p.Keys) { if (-not [string]::IsNullOrEmpty($p[$k])) { "$k=" + [System.Uri]::EscapeDataString([string]$p[$k]) } }
    return HtmlEncode ("/bewegungen?" + ($teile -join '&'))
}

function Build-Bewegungen($table, $f, $namen, $max, $fehler) {
    $head = Page-Head "Lagerbewegungen"
    $anzahl = 0

    if ($fehler) {
        $tabelle = "<div class='fehler'>Fehler beim Laden der Daten:<br><pre>$(HtmlEncode $fehler)</pre></div>"
    } elseif ($null -eq $table -or $table.Rows.Count -eq 0) {
        $tabelle = "<p style='color:#666'>Keine Buchungen f&uuml;r diese Auswahl gefunden.</p>"
    } else {
        $anzahl = $table.Rows.Count
        $generisch = @('Manuelle Buchungen', 'Manuelle Umbuchung', 'IAS MDE Lagerbuchung', 'Inventurbuchungen')
        # Filterlinks je Zeile: aktuelle Auswahl ohne den jeweiligen Parameter einmal vorab bauen
        # (Bew-Url je Zelle waere bei >1000 Zeilen spuerbar langsam)
        $basis = @{}
        foreach ($k in 'wer', 'art', 'platz') { $basis[$k] = ([System.Net.WebUtility]::HtmlDecode((Bew-Url $f @{ $k = '' }))) + "&$k=" }
        # Zellen fuer Person, Artikel und Lagerplatz wiederholen sich oft - je Wert nur einmal bauen
        $he = [System.Net.WebUtility]
        $cache = @{}
        $link = { param($k, $v, $text, $titel)
            $key = "$k|$v|$text"
            if (-not $cache.ContainsKey($key)) {
                $cache[$key] = "<a class='filterlink' href='" + $he::HtmlEncode($basis[$k] + [System.Uri]::EscapeDataString([string]$v)) + "'$titel>" + $he::HtmlEncode($text) + "</a>"
            }
            $cache[$key]
        }
        $rows = foreach ($r in $table.Rows) {
            $ba = [string]$r["Bewegungsart"]
            $meta = if ($BewegungsArten.ContainsKey($ba)) { $BewegungsArten[$ba] } else { @('um', $ba, 'um') }
            $gruppe = $meta[0]

            $dat = $r["Bewegungsdatum"]
            $datTxt = if ($dat -is [datetime]) { $dat.ToString("dd.MM.yyyy") } else { "" }
            $zeit = $r["Zeit"]
            if ($zeit -is [datetime]) { $datTxt += " <span class='zeit'>$($zeit.ToString('HH:mm'))</span>" }

            $b = [string]$r["Benutzer"]
            $n = if ($namen.ContainsKey($b)) { $namen[$b] } else { @{ Name = $b; Scanner = $false } }
            $werTxt = & $link 'wer' $n.Name $n.Name (" title='Sage-Benutzer: " + $he::HtmlEncode($b) + "'")
            if ($n.Scanner) { $werTxt += " <span class='badge-scan' title='Gebucht mit dem MDE-Scanner'>Scanner</span>" }

            $an = [string]$r["Artikelnummer"]
            $artTxt = & $link 'art' $an $an ''

            $m = $r["Menge"]
            $mZahl = if ($m -is [System.DBNull]) { "" } elseif ([math]::Floor($m) -eq $m) { ([decimal]$m).ToString('0') } else { "{0:N2}" -f $m }
            $mTxt = if ($gruppe -eq 'zu') { "+" + $mZahl } elseif ($gruppe -eq 'ab') { "&minus;" + $mZahl } else { $mZahl }

            $von = ""; $nach = ""
            if ($r["HKennung"] -isnot [System.DBNull]) { $von = & $link 'platz' (Platz-Wert $r["HKennung"] $r["HD1"] $r["HD2"] $r["HD3"]) (Fmt-Platz $r["HKennung"] $r["HD1"] $r["HD2"] $r["HD3"]) '' }
            if ($r["ZKennung"] -isnot [System.DBNull]) { $nach = & $link 'platz' (Platz-Wert $r["ZKennung"] $r["ZD1"] $r["ZD2"] $r["ZD3"]) (Fmt-Platz $r["ZKennung"] $r["ZD1"] $r["ZD2"] $r["ZD3"]) '' }
            $ref = ([string]$r["Referenz"]).Trim()
            $std = ([string]$r["Standardtext"]).Trim()
            $memo = ([string]$r["Memo"]).Trim()
            $bezug = ""
            if ($ref -match '^\d{4}-\d+$') {
                $wort = switch ($ba) { { 'EA','ZS' -contains $_ } { 'Lieferschein' } { 'ZR','SR' -contains $_ } { 'R&uuml;cklieferung' } default { 'Wareneingang' } }
                $name = if ($r["Kunde"] -isnot [System.DBNull]) { [string]$r["Kunde"] } elseif ($r["Lieferant"] -isnot [System.DBNull]) { [string]$r["Lieferant"] } else { "" }
                $bezug = "$wort $(HtmlEncode $ref)$(if ($name) { " &middot; $(HtmlEncode $name)" })"
            }
            elseif ($std -ne '' -and $generisch -notcontains $std) { $bezug = HtmlEncode $std }
            if ($std -match "'\d{4}-\d+'$" -and $bezug) { $memo = '' }
            if ($memo -ne '' -and $memo -ne $std -and $generisch -notcontains $memo -and $memo -notmatch '^FA \d+ Bewegungsart') {
                $bezug += $(if ($bezug) { " &middot; " } else { "" }) + "<span class='memo'>$(HtmlEncode $memo)</span>"
            }

            "<tr><td style='white-space:nowrap'>$datTxt</td><td style='white-space:nowrap'>$werTxt</td>" +
            "<td><span class='bew bew-$($meta[2])' title='Sage-Bewegungsart $ba'>$($meta[1])</span></td>" +
            "<td>$artTxt</td><td>$($he::HtmlEncode([string]$r['Bezeichnung']))</td>" +
            "<td class='num'>$mTxt</td><td>$von</td><td>$nach</td><td>$bezug</td></tr>"
        }
        $hinweis = if ($anzahl -ge $max) { "<p class='hinweis'>Es werden nur die neuesten $max Buchungen angezeigt &ndash; bitte Zeitraum oder Filter eingrenzen.</p>" } else { "" }
        $tabelle = @"
$hinweis
<table id='tab'>
<thead><tr>
  <th>Datum</th><th>Wer</th><th>Art</th><th>Artikel</th><th>Bezeichnung</th>
  <th class='num'>Menge</th><th>Von Lagerplatz</th><th>Nach Lagerplatz</th><th>Bezug</th>
</tr></thead>
<tbody>$($rows -join '')</tbody>
</table>
"@
    }

    # Schnellfilter Art und Zeitraum
    $typLinks = (@(@('', 'Alle'), @('um', 'Umbuchungen'), @('zu', 'Zug&auml;nge'), @('ab', 'Entnahmen')) | ForEach-Object {
        $cls = if ($_[0] -eq $f.typ) { "btn" } else { "btn btn2" }
        "<a class='$cls' href='$(Bew-Url $f @{ typ = $_[0] })'>$($_[1])</a>"
    }) -join " "
    $heute = (Get-Date).Date
    $zeitLinks = (@(@(0, 'Heute'), @(6, '7 Tage'), @(29, '30 Tage')) | ForEach-Object {
        $v = $heute.AddDays(-$_[0])
        $cls = if ($f.von -eq $v -and $f.bis -eq $heute) { "btn" } else { "btn btn2" }
        "<a class='$cls' href='$(Bew-Url $f @{ von = $v.ToString('yyyy-MM-dd'); bis = $heute.ToString('yyyy-MM-dd') })'>$($_[1])</a>"
    }) -join " "

    $personen = @($namen.Values | ForEach-Object { $_.Name } | Sort-Object -Unique)
    $optionen = "<option value=''>Alle Personen</option>" + (($personen | ForEach-Object {
        $sel = if ($_ -eq $f.wer) { " selected" } else { "" }
        "<option value='$(HtmlEncode $_)'$sel>$(HtmlEncode $_)</option>"
    }) -join '')

    $aktiveFilter = @()
    if ($f.wer)   { $aktiveFilter += "Person: <b>$(HtmlEncode $f.wer)</b>" }
    if ($f.art)   { $aktiveFilter += "Artikel: <b>$(HtmlEncode $f.art)</b>" }
    if ($f.platz) { $aktiveFilter += $(if ($f.platz -like '*;*') { "Lagerplatz: <b>$(HtmlEncode (Platz-Text $f.platz))</b>" } else { "Lager: <b>$(HtmlEncode $f.platz)</b>" }) }
    $filterZeile = if ($aktiveFilter.Count -gt 0) {
        "<div class='aktivfilter'>Gefiltert nach $($aktiveFilter -join ', ') &middot; <a href='$(Bew-Url $f @{ wer = ''; art = ''; platz = '' })'>Filter entfernen</a></div>"
    } else { "" }

    return @"
$head
<style>
.bew { display:inline-block; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:600; white-space:nowrap; }
.bew-um  { background:#e8f4fd; color:#086ED5; }
.bew-zu  { background:#e8f5e9; color:#1b5e20; }
.bew-ab  { background:#fdecea; color:#b71c1c; }
.bew-inv { background:#eceff1; color:#455a64; }
.badge-scan { display:inline-block; padding:0 6px; border-radius:8px; font-size:10px; font-weight:600; background:#f3e5f5; color:#6a1b9a; vertical-align:middle; }
.zeit { color:#888; font-size:12px; }
.memo { color:#888; font-size:12px; }
a.filterlink { color:inherit; text-decoration:none; border-bottom:1px dotted #9aa5b1; }
a.filterlink:hover { color:#086ED5; border-bottom-color:#086ED5; }
.bewform { display:flex; gap:10px; flex-wrap:wrap; align-items:center; margin-bottom:12px; }
.bewform input[type=date] { padding:7px 8px; border:1px solid #ccc; border-radius:999px; font-family:inherit; font-size:14px; }
.bewform .suche { max-width:230px; }
.bewform .autobox { max-width:380px; }
.bewform .autobox .suche { max-width:none; }
.bewform select { min-width:200px; padding:8px 14px; border:1px solid #d6d8db; border-radius:999px; background:#fff; color:#343A40; font-family:inherit; font-size:14px; }
.aktivfilter { background:#fff8e1; border:1px solid #ffe082; padding:7px 12px; border-radius:6px; font-size:13px; margin-bottom:12px; }
.hinweis { color:#b26a00; font-size:13px; margin:0 0 8px; }
.trenner { width:10px; }
</style>
<div class='kopf'>
  <h1>Lagerbewegungen</h1>
  <div class='meta'>$(StandText) &middot; $anzahl Buchungen</div>
</div>
<div class='steuer'>
  $zeitLinks
  <span class='trenner'></span>
  $typLinks
</div>
<form class='bewform' method='get' action='/bewegungen'>
  <input type='hidden' name='typ' value='$(HtmlEncode $f.typ)'>
  <label>von <input type='date' name='von' value='$($f.von.ToString('yyyy-MM-dd'))'></label>
  <label>bis <input type='date' name='bis' value='$($f.bis.ToString('yyyy-MM-dd'))'></label>
  <select name='wer' onchange='this.form.submit()'>$optionen</select>
  <div class='autobox'>
    <input class='suche' id='artsuche' name='art' autocomplete='off' value='$(HtmlEncode $f.art)' placeholder='Artikel'
           oninput='vorschlage()' onkeydown='navi(event)'>
    <div class='vorschlaege' id='vorschlaege'></div>
  </div>
  <div class='autobox'>
    <input class='suche' id='platzsuche' autocomplete='off' value='$(HtmlEncode (Platz-Text $f.platz))' placeholder='Lagerplatz'
           oninput='platzVorschlage()' onkeydown='platzNavi(event)'>
    <input type='hidden' name='platz' id='platzwert' value='$(HtmlEncode $f.platz)'>
    <div class='vorschlaege' id='platzvorschlaege'></div>
  </div>
  <button class='btn' type='submit'>Anzeigen</button>
</form>
$filterZeile
$tabelle
<script>
var timer = null, selIdx = -1;
function vorschlage() {
  var q = document.getElementById('artsuche').value.trim();
  if (timer) clearTimeout(timer);
  if (q.length < 2) { schliesse(); return; }
  timer = setTimeout(function () {
    fetch('/api/artikelsuche?q=' + encodeURIComponent(q))
      .then(function (r) { return r.json(); })
      .then(zeige)
      .catch(schliesse);
  }, 200);
}
function zeige(daten) {
  var box = document.getElementById('vorschlaege');
  selIdx = -1; box.textContent = '';
  if (!daten || daten.length === 0) { schliesse(); return; }
  daten.forEach(function (d) {
    var div = document.createElement('div'), nr = document.createElement('span'), bez = document.createElement('span');
    div.setAttribute('data-nr', d.nr); nr.className = 'nr'; nr.textContent = d.nr; bez.textContent = d.bez;
    div.appendChild(nr); div.appendChild(bez);
    div.onclick = function () { waehle(d.nr); };
    box.appendChild(div);
  });
  box.style.display = 'block';
}
function schliesse() { var box = document.getElementById('vorschlaege'); box.style.display = 'none'; box.textContent = ''; selIdx = -1; }
function waehle(nr) { var ein = document.getElementById('artsuche'); ein.value = nr; schliesse(); ein.form.submit(); }
function navi(e) {
  var items = document.getElementById('vorschlaege').querySelectorAll('div');
  if (e.key === 'Enter' && selIdx >= 0 && items[selIdx]) { e.preventDefault(); waehle(items[selIdx].getAttribute('data-nr')); }
  else if (e.key === 'ArrowDown' && items.length) { e.preventDefault(); if (selIdx < items.length - 1) selIdx++; markiere(items); }
  else if (e.key === 'ArrowUp' && items.length) { e.preventDefault(); if (selIdx > 0) selIdx--; markiere(items); }
  else if (e.key === 'Escape') { schliesse(); }
}
function markiere(items) { for (var i = 0; i < items.length; i++) items[i].className = (i === selIdx) ? 'sel' : ''; }
document.addEventListener('click', function (e) { if (!e.target.closest('.autobox')) schliesse(); });
// Lagerplatz: Vorschlaege "ganzes Lager" und einzelne Plaetze; Wert (Kennung;D1;D2;D3) im versteckten Feld
var pTimer = null, pSel = -1;
function platzVorschlage() {
  var q = document.getElementById('platzsuche').value.trim();
  document.getElementById('platzwert').value = q;   // frei getippt = ganzes Lager mit diesem Anfang
  if (pTimer) clearTimeout(pTimer);
  if (q.length < 2) { platzSchliesse(); return; }
  pTimer = setTimeout(function () {
    fetch('/api/lagerplatzsuche?mit=alle&q=' + encodeURIComponent(q))
      .then(function (r) { return r.json(); }).then(platzZeige).catch(platzSchliesse);
  }, 200);
}
function platzZeige(daten) {
  var box = document.getElementById('platzvorschlaege');
  pSel = -1; box.textContent = '';
  if (!daten || daten.length === 0) { platzSchliesse(); return; }
  daten.forEach(function (d) {
    var div = document.createElement('div'), t = document.createElement('span');
    div.setAttribute('data-v', d.v); div.setAttribute('data-t', d.t);
    t.textContent = d.t; div.appendChild(t);
    if (d.g) { var z = document.createElement('span'); z.textContent = ' ganzes Lager'; z.style.color = '#888'; div.appendChild(z); }
    div.onclick = function () { platzWaehle(d.v, d.t); };
    box.appendChild(div);
  });
  box.style.display = 'block';
}
function platzSchliesse() { var box = document.getElementById('platzvorschlaege'); box.style.display = 'none'; box.textContent = ''; pSel = -1; }
function platzWaehle(v, t) { document.getElementById('platzsuche').value = t; document.getElementById('platzwert').value = v; platzSchliesse(); document.getElementById('platzwert').form.submit(); }
function platzNavi(e) {
  var items = document.getElementById('platzvorschlaege').querySelectorAll('div');
  if (e.key === 'Enter' && pSel >= 0 && items[pSel]) { e.preventDefault(); platzWaehle(items[pSel].getAttribute('data-v'), items[pSel].getAttribute('data-t')); }
  else if (e.key === 'ArrowDown' && items.length) { e.preventDefault(); if (pSel < items.length - 1) pSel++; for (var i = 0; i < items.length; i++) items[i].className = (i === pSel) ? 'sel' : ''; }
  else if (e.key === 'ArrowUp' && items.length) { e.preventDefault(); if (pSel > 0) pSel--; for (var j = 0; j < items.length; j++) items[j].className = (j === pSel) ? 'sel' : ''; }
  else if (e.key === 'Escape') { platzSchliesse(); }
}
document.addEventListener('click', function (e) { if (!e.target.closest('.autobox')) platzSchliesse(); });
</script>
</body></html>
"@
}


# ============================================================
#  ASSISTENT (Chat mit Zugriff auf die Sage-Daten) - aus der Oezer-Liste uebernommen
# ============================================================
# Ablauf: Browser schickt den Gespraechsverlauf an POST /api/chat -> Hintergrund-Runspace
# fragt das Modell (Anthropic Messages API) mit festen Werkzeugen. Das Modell liest Sage nur
# ueber diese Werkzeuge (dieselben Abfragen wie die Seiten, kein freies SQL, keine Preise).
# Der Browser fragt den Stand per GET /api/chat?id=... ab, damit der (einthreadige)
# Webserver waehrend der Antwortzeit nicht blockiert.

$ChatSystemPrompt = @'
Du bist der Assistent der Solar Dispo Liste von Leyendecker Solar (Planung und Montage von Photovoltaikanlagen mit Speichern und Wallboxen). Du beantwortest Fragen von Projektleitung, Montageplanung, Lager und Einkauf zu Kundenauftraegen, Montageterminen, Material, Lagerbestaenden und Bestellungen. Datenquelle sind die Sage-Warenwirtschaft und der Montagekalender - du erreichst sie NUR ueber die Werkzeuge.

Regeln:
- Nenne nur Zahlen und Fakten aus Werkzeugergebnissen. Erfinde nichts. Wenn die Werkzeuge etwas nicht hergeben, sag das offen und nenne, welche Seite der Solar Dispo Liste weiterhelfen koennte.
- Keine Preise, Betraege oder Werte - die Werkzeuge liefern bewusst keine.
- Wird ein Artikel nur beschrieben (nicht per Nummer genannt), zuerst artikel_suchen. Bei mehreren passenden Treffern kurz die Kandidaten nennen und nachfragen, statt zu raten.
- Antworte auf Deutsch mit korrekten Umlauten, kurz und konkret. Zuerst die Antwort in ein bis zwei Saetzen, dann bei Bedarf Details als kurze Liste. Tabellen nur wenn wirklich hilfreich (max. 6 Spalten). Datum als TT.MM.JJJJ.
- Verlinke passende Seiten als Markdown-Links mit den relativen Pfaden aus den Werkzeugergebnissen (Zeilen "Link:" oder Pfade in Klammern), z.B. [2026-40147](/auftrag?bel=9612). Craftnote-Projekte als Link auf https://app.mycraftnote.de/... wie im Werkzeugergebnis, sonst keine externen Links.
- Du kannst nur lesen. Buchungen, Aenderungen, Termine oder Bestellungen kannst du nicht ausfuehren.
- Craftnote ist die Projekt-App (Projektstatus von Anfrage ueber Angebot, Foerderung, Auftrag terminiert, EVU-Unterlagen, Zaehlerwechsel, Inbetriebnahme bis abgeschlossen; dazu Notizen, Dateien, Zeiten). Craftnote-Projekte werden den Sage-Auftraegen ueber die E-Mail-Adresse des Kunden zugeordnet, sonst ueber Name und PLZ (dann ggf. unsicher). Aufgaben und Chat aus Craftnote sind nicht verfuegbar.

Begriffe: AB = Auftragsbestaetigung, RES = Reservierungsbestaetigung (zaehlt wie ein Auftrag), AB/RES = AB aus einer Reservierung. Montagetermine stehen im Online-Kalender (Titel = Kundenname, Kategorie = Monteur/Team) und werden den Auftraegen ueber den Namen zugeordnet; "unsicher" heisst, nur der Nachname passt - dann darauf hinweisen. Planungstermin eines Auftrags = naechster Montagetermin, sonst der Liefertermin. Material-Status: da = aus Lagerbestand gedeckt, bestellt = durch eine offene Bestellung gedeckt (mit Liefertermin), fehlt = weder Bestand noch Bestellung, kein Lagerartikel = wird je Auftrag beschafft (z.B. Montagesystem LORENZ). Bestand und Bestellungen werden den Auftraegen in der Reihenfolge ihrer Planungstermine zugeteilt. Stueckliste = Set (z.B. Speicher), dessen Bestand und Bedarf ueber die Bestandteile laufen. kWp = Anlagenleistung aus den Modulpositionen. Geliefert/berechnet = wie viele Positionen eines Auftrags schon geliefert bzw. in Rechnung gestellt sind.

Heute ist {DATUM}.
'@

$ChatToolsJson = @'
[
 {"name":"artikel_suchen","description":"Sucht aktive Artikel nach Artikelnummer oder Bezeichnung (alle Suchwoerter muessen vorkommen). Nutzen, wenn der Nutzer einen Artikel beschreibt oder die Nummer unsicher ist.",
  "input_schema":{"type":"object","properties":{"suchtext":{"type":"string","description":"Nummer, Nummernteil oder Woerter aus der Bezeichnung, z.B. 'BYD 5.1' oder 'Fronius Symo 10'"}},"required":["suchtext"]}},
 {"name":"offene_auftraege","description":"Offene Kundenauftraege mit naechstem Montagetermin (Team), Liefertermin, Anlagengroesse, Liefer-/Rechnungsstand und Materialstatus je Gruppe (Module, Montagesystem, WR & Speicher, Wallbox, Elektro). Sortiert nach Planungstermin. Fuer Fragen wie: was steht als naechstes an, bei welchen Auftraegen fehlt Material, wie ist der Stand bei Kunde X.",
  "input_schema":{"type":"object","properties":{
    "kunde":{"type":"string","description":"Kundenname oder Teil davon"},
    "beleg":{"type":"string","description":"Auftragsnummer, Format JJJJ-Nummer, z.B. 2026-40147"},
    "planung_bis":{"type":"string","description":"nur Auftraege mit Planungstermin (Montage oder Liefertermin) bis einschliesslich JJJJ-MM-TT"},
    "nur_mit_fehlendem_material":{"type":"boolean"}}}},
 {"name":"auftrag","description":"Alle Details eines Auftrags: Kunde, Baustelle, alle Montagetermine, jede Material- und Leistungsposition mit Menge, geliefert, berechnet und Status, Lieferscheine und Rechnungen.",
  "input_schema":{"type":"object","properties":{"beleg":{"type":"string","description":"Auftragsnummer JJJJ-Nummer"},"kunde":{"type":"string","description":"alternativ Kundenname, wenn die Nummer unbekannt ist"}}}},
 {"name":"materialbedarf","description":"Was fehlt fuer die anstehenden Auftraege: je Lagerartikel fehlende Menge, Bestand, offene Bestellmenge mit fruehestem Liefertermin und der erste betroffene Auftrag. Dazu die Positionen ohne Lagerfuehrung, die je Auftrag beschafft werden muessen (z.B. Montagesystem).",
  "input_schema":{"type":"object","properties":{"tage":{"type":"integer","description":"nur Auftraege mit Planungstermin bis heute plus N Tage, Standard 28; 0 = alle"}}}},
 {"name":"montagetermine","description":"Termine aus dem Montagekalender in einem Zeitraum mit Team/Monteur und dem zugeordneten Auftrag. Fuer Fragen wie: wer ist naechste Woche wo, welche Montagen hat Danny, wann wird bei Kunde X montiert.",
  "input_schema":{"type":"object","properties":{
    "von":{"type":"string","description":"JJJJ-MM-TT, Standard heute"},
    "bis":{"type":"string","description":"JJJJ-MM-TT, Standard heute + 7 Tage"},
    "team":{"type":"string","description":"Name des Monteurs/Teams oder Teil davon"},
    "text":{"type":"string","description":"Text im Termintitel, z.B. Kundenname"}}}},
 {"name":"artikel_dispo","description":"Disposition eines Artikels: Lagerbestand und alle offenen Bedarfe (Auftraege mit Montage-/Liefertermin) und Zugaenge (Bestellungen) in Terminreihenfolge mit kumuliertem Bestand und Deckung. Bei Stuecklisten (z.B. Speicher-Sets) die Auftraege mit Status der Bestandteile. Fuer Fragen wie: reicht der Bestand, wann kommt Nachschub, wer braucht den Artikel.",
  "input_schema":{"type":"object","properties":{"artikelnummer":{"type":"string"}},"required":["artikelnummer"]}},
 {"name":"lagerbestand","description":"Aktueller Bestand eines Artikels je Lagerplatz. Bei Stuecklisten die Bestandteile mit ihrem Bestand.",
  "input_schema":{"type":"object","properties":{"artikelnummer":{"type":"string"}},"required":["artikelnummer"]}},
 {"name":"offene_bestellungen","description":"Offene Einkaufsbestellungen mit Lieferant, Artikel, offener Menge, Liefertermin und fuer welche Auftraege (mit Planungstermin) die Ware gebraucht wird.",
  "input_schema":{"type":"object","properties":{
    "artikelnummer":{"type":"string"},
    "lieferant":{"type":"string","description":"Lieferantenname oder Teil davon"},
    "nur_ueberfaellig":{"type":"boolean","description":"nur Bestellungen mit Liefertermin vor heute"}}}},
 {"name":"lagerbewegungen","description":"Protokoll der Lagerbuchungen: wer hat wann welchen Artikel von welchem Lagerplatz wohin gebucht (Wareneingang, Lieferung an Kunde, Umbuchung, Inventur), mit Lieferschein und Kunde bzw. Wareneingang und Lieferant. Neueste zuerst, max. 150 Zeilen. Sage speichert nur das Datum.",
  "input_schema":{"type":"object","properties":{
    "von":{"type":"string","description":"JJJJ-MM-TT, Standard: vor 7 Tagen"},
    "bis":{"type":"string","description":"JJJJ-MM-TT, Standard: heute"},
    "artikelnummer":{"type":"string","description":"Artikelnummer (Anfang) oder Text aus der Bezeichnung"},
    "lager":{"type":"string","description":"Lagerkennung oder Platz, z.B. A, A 1/2/3, Warenlager"},
    "person":{"type":"string","description":"Name der Person"},
    "art":{"type":"string","enum":["um","zu","ab"],"description":"um = Umbuchungen, zu = Zugaenge, ab = Entnahmen"}}}},
 {"name":"craftnote_projekte","description":"Projekte aus Craftnote (Projekt-App) mit Status, Ort, letzter Aenderung, Notiz, Link und zugeordnetem offenen Sage-Auftrag; dazu die Anzahl je Status. Fuer Fragen wie: welche Projekte warten auf Foerderung, wo sind EVU-Unterlagen eingereicht, wie viele Angebote sind versendet, Stand bei Kunde X in Craftnote.",
  "input_schema":{"type":"object","properties":{
    "status":{"type":"string","description":"Teil des Statusnamens, z.B. Foerderung, EVU, Angebot, Auftrag vorhanden, Inbetrieb, abgeschlossen"},
    "text":{"type":"string","description":"Teil von Projektname, Ort oder PLZ"}}}},
 {"name":"lagerplatz_inhalt","description":"Welche Artikel mit welchem Bestand auf Lagerplaetzen liegen, deren Lagerkennung so beginnt (oder genau auf einem Platz wie A 1/2/3).",
  "input_schema":{"type":"object","properties":{"lager":{"type":"string","description":"Lagerkennung oder Platz, z.B. A, A 1/2, Warenlager"}},"required":["lager"]}}
]
'@

# JSON-Serialisierung ohne die Eigenheiten von ConvertTo-Json in PS 5.1 (einelementige
# Arrays, leere Objekte). Nicht-ASCII wird als \uXXXX geschrieben.
function ConvertTo-JsonText($o) {
    if ($null -eq $o -or $o -is [System.DBNull]) { return 'null' }
    if ($o -is [string] -or $o -is [char]) {
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append('"')
        foreach ($ch in ([string]$o).ToCharArray()) {
            $c = [int]$ch
            if ($ch -eq '"') { [void]$sb.Append('\"') }
            elseif ($ch -eq '\') { [void]$sb.Append('\\') }
            elseif ($c -eq 10) { [void]$sb.Append('\n') }
            elseif ($c -eq 13) { [void]$sb.Append('\r') }
            elseif ($c -eq 9) { [void]$sb.Append('\t') }
            elseif ($c -lt 32 -or $c -gt 126) { [void]$sb.Append('\u').Append($c.ToString('x4')) }
            else { [void]$sb.Append($ch) }
        }
        [void]$sb.Append('"')
        return $sb.ToString()
    }
    if ($o -is [bool]) { if ($o) { return 'true' } else { return 'false' } }
    if ($o -is [datetime]) { return '"' + $o.ToString('yyyy-MM-ddTHH:mm:ss') + '"' }
    if ($o -is [int] -or $o -is [long] -or $o -is [int16] -or $o -is [byte] -or $o -is [decimal] -or $o -is [double] -or $o -is [single]) {
        return ([System.IConvertible]$o).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    }
    if ($o -is [System.Collections.IDictionary]) {
        $teile = foreach ($k in $o.Keys) { (ConvertTo-JsonText ([string]$k)) + ':' + (ConvertTo-JsonText $o[$k]) }
        return '{' + ($teile -join ',') + '}'
    }
    if ($o -is [System.Management.Automation.PSCustomObject]) {
        $teile = foreach ($p in $o.PSObject.Properties) { (ConvertTo-JsonText $p.Name) + ':' + (ConvertTo-JsonText $p.Value) }
        return '{' + ($teile -join ',') + '}'
    }
    if ($o -is [System.Collections.IEnumerable]) {
        $teile = foreach ($x in $o) { ConvertTo-JsonText $x }
        return '[' + ($teile -join ',') + ']'
    }
    return ConvertTo-JsonText ([string]$o)
}

function Get-ChatApiKey {
    foreach ($ziel in 'Process', 'User', 'Machine') {
        $k = [Environment]::GetEnvironmentVariable('ANTHROPIC_API_KEY', $ziel)
        if (-not [string]::IsNullOrWhiteSpace($k)) { return $k.Trim() }
    }
    return $null
}

# Chat fuer diese Anfrage anzeigen/erlauben? (Schluessel vorhanden + ggf. Testfreigabe per Cookie)
function Test-ChatErlaubt {
    if ($null -eq $ChatPool) { return $false }
    if ($null -eq $script:ChatKeyGeprueft -or ((Get-Date) - $script:ChatKeyGeprueft).TotalSeconds -gt 60) {
        $script:ChatKeyVorhanden = [bool](Get-ChatApiKey); $script:ChatKeyGeprueft = Get-Date
    }
    if (-not $script:ChatKeyVorhanden) { return $false }
    if ($ChatFuerAlle) { return $true }
    $ck = $request.Cookies['solarchat']
    return ($null -ne $ck -and $ck.Value -eq '1')
}

# ---- Werkzeuge ----
function ChatArg($inp, $name) {
    if ($null -eq $inp) { return '' }
    $p = $inp.PSObject.Properties[$name]
    if ($null -eq $p -or $null -eq $p.Value) { return '' }
    return ([string]$p.Value).Trim()
}
function ChatDat($d) { if ($d -is [datetime] -and $d -ne [datetime]::MaxValue) { return $d.ToString('dd.MM.yyyy') } else { return '-' } }
function ChatDatum($s, $standard) {
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact([string]$s, 'yyyy-MM-dd', $null, 'None', [ref]$d)) { return $d }
    return $standard
}
function ChatKuerzen($text, $max) {
    if ($text.Length -le $max) { return $text }
    return $text.Substring(0, $max) + "`n... (gekuerzt)"
}
function ChatText($s) { return [System.Net.WebUtility]::HtmlDecode([string]$s) }
# Offene Auftraege fuer die Werkzeuge: innerhalb einer Frage nur einmal laden (60 s)
function Chat-Auftraege {
    if ($null -eq $script:ChatAuZeit -or ((Get-Date) - $script:ChatAuZeit).TotalSeconds -gt 60) {
        $script:ChatAu = @(Get-Auftraege); $script:ChatAuZeit = Get-Date
    }
    return $script:ChatAu
}
function ChatTermin($tr) {
    if (-not $tr) { return 'kein Montagetermin' }
    $t = $tr.T
    $zeit = if ($t.Ende.Date -gt $t.Start.Date) { "$($t.Start.ToString('dd.MM.yyyy'))-$($t.Ende.ToString('dd.MM.yyyy'))" } else { $t.Start.ToString('dd.MM.yyyy') }
    return "$zeit ($(@($t.Team) -join ', '))$(if (-not $tr.Sicher) { ' [unsicher zugeordnet]' })"
}
function ChatMatStatus($m) {
    switch ($m.Status) {
        'da' { 'da' }
        'bestellt' { "bestellt$(if ($m.EkTermin -is [datetime]) { ' Liefertermin ' + $m.EkTermin.ToString('dd.MM.yyyy') + $(if ($m.EkTermin.Date -lt (Get-Date).Date) { ' (ueberfaellig)' }) })" }
        'fehlt' { "fehlt $(FmtZahl $m.Fehl)" }
        default { 'kein Lagerartikel' }
    }
}
function ChatStand($st, $was) {
    if (-not $st -or -not $st.Pos) { return '-' }
    $voll = if ($was -eq 'Gel') { $st.Gel } else { $st.Ber }; $anf = if ($was -eq 'Gel') { $st.GelAnf } else { $st.BerAnf }
    if ($anf -eq 0) { return 'nichts' }
    if ($voll -ge $st.Pos) { return 'alles' }
    return "$voll/$($st.Pos) Pos."
}
function ChatAuftragFinden($liste, $beleg, $kunde) {
    if ($beleg) { return @($liste | Where-Object { $_.Beleg -eq $beleg }) }
    if ($kunde) { return @($liste | Where-Object { $_.Kunde -like "*$kunde*" -or $_.BaustelleName -like "*$kunde*" }) }
    return @()
}

# SQL-Bedingung fuer eine Lagerangabe. "A 4/2/2" = Lagerkennung "A" mit Dimensionen 4/2/2 (so zeigt
# Fmt-Platz die Plaetze an); sonst Anfang der Lagerkennung. Ist der ganze Text eine Lagerkennung, gilt er als solche.
# $aliase: Liste von @(AliasLagerkennung, AliasDimensionen), Bedingungen werden ODER-verknuepft.
function ChatPlatzBedingung($lager, $aliase, $p) {
    $exakt = $false
    $m = $null
    if ($lager -match '^(.+?)\s+(\d+)(?:/(\d+))?(?:/(\d+))?$') {
        $m = $matches.Clone()
        $istKennung = Run-Query "SELECT TOP 1 1 AS x FROM KHKLagerplaetze WHERE Mandant = $Mandant AND Lagerkennung = @k" @{ "@k" = $lager }
        $exakt = ($istKennung.Rows.Count -eq 0)
    }
    if (-not $exakt) {
        $p['@platz'] = $lager
        return '(' + (($aliase | ForEach-Object { "$($_[0]).Lagerkennung LIKE @platz + '%'" }) -join ' OR ') + ')'
    }
    $p['@lk'] = $m[1]
    $dims = @()
    for ($i = 2; $i -le 4; $i++) { if ($m[$i]) { $p["@d$($i - 1)"] = [int]$m[$i]; $dims += $i - 1 } }
    return '(' + (($aliase | ForEach-Object {
        $a = $_
        "($($a[0]).Lagerkennung = @lk" + (($dims | ForEach-Object { " AND $($a[1]).Dimension$_ = @d$_" }) -join '') + ')'
    }) -join ' OR ') + ')'
}

function Invoke-ChatTool($name, $inp) {
    $heute = (Get-Date).Date
    switch ($name) {
        'artikel_suchen' {
            $worte = @((ChatArg $inp 'suchtext') -split '\s+' | Where-Object { $_ -ne '' } | Select-Object -First 5)
            if ($worte.Count -eq 0) { return 'Kein Suchtext angegeben.' }
            $p = @{}; $bed = for ($i = 0; $i -lt $worte.Count; $i++) {
                $p["@w$i"] = "%$($worte[$i])%"; "(Artikelnummer LIKE @w$i OR Bezeichnung1 LIKE @w$i OR ISNULL(Bezeichnung2,'') LIKE @w$i)"
            }
            $t = Run-Query "SELECT TOP 25 Artikelnummer, Bezeichnung1, Stuecklistentyp, Lagerfuehrung FROM KHKArtikel WHERE Mandant = $Mandant AND Aktiv = -1 AND $($bed -join ' AND ') ORDER BY Artikelnummer" $p
            if ($t.Rows.Count -eq 0) { return "Keine aktiven Artikel gefunden fuer: $($worte -join ' ')" }
            $z = foreach ($r in $t.Rows) { "$($r['Artikelnummer']) | $($r['Bezeichnung1'])$(if ((ToNum $r['Stuecklistentyp']) -eq 2) { ' | Stueckliste' } elseif ((ToNum $r['Lagerfuehrung']) -ne -1) { ' | kein Lagerartikel' })" }
            return "Artikelnummer | Bezeichnung ($($t.Rows.Count) Treffer, max. 25). Dispo je Artikel: /dispo?art=<Nummer>`n" + ($z -join "`n")
        }
        'offene_auftraege' {
            $liste = Chat-Auftraege
            $kunde = ChatArg $inp 'kunde'; $beleg = ChatArg $inp 'beleg'
            $bis = ChatDatum (ChatArg $inp 'planung_bis') $null
            $nurFehlt = (ChatArg $inp 'nur_mit_fehlendem_material') -eq 'True'
            $treffer = @($liste | Where-Object {
                (-not $kunde -or $_.Kunde -like "*$kunde*" -or $_.BaustelleName -like "*$kunde*") -and (-not $beleg -or $_.Beleg -eq $beleg) -and
                (-not $bis -or $_.Planung -le $bis) -and (-not $nurFehlt -or @($_.Material | Where-Object { $_.Status -eq 'fehlt' }).Count)
            })
            if ($treffer.Count -eq 0) { return "Keine offenen Auftraege fuer diese Auswahl ($($liste.Count) offene Auftraege insgesamt)." }
            try { $null = Craftnote-Zuordnen $treffer } catch {}
            $gruppen = @($MaterialGruppen.Values | Select-Object -Unique) + 'Sonstiges'
            $z = foreach ($au in ($treffer | Select-Object -First 60)) {
                $mat = foreach ($g in $gruppen) {
                    $zl = @($au.Material | Where-Object { $_.Gruppe -eq $g }); if (-not $zl.Count) { continue }
                    $f = @($zl | Where-Object { $_.Status -eq 'fehlt' }); $b = @($zl | Where-Object { $_.Status -eq 'bestellt' })
                    $gt = ChatText $g
                    if ($f.Count) { "$gt FEHLT (" + (($f | ForEach-Object { "$($_.Art) $($_.Bez) fehlt $(FmtZahl $_.Fehl)" }) -join '; ') + ')' }
                    elseif ($b.Count) { "$gt bestellt (" + (($b | ForEach-Object { "$($_.Art) $(ChatMatStatus $_)" }) -join '; ') + ')' }
                    elseif (@($zl | Where-Object { $_.Status -eq 'da' }).Count) { "$gt da" } else { "$gt kein Lagerartikel" }
                }
                "$(ChatTermin $au.Naechster) | $($au.Art) $($au.Beleg) (/auftrag?bel=$($au.BelID)) | $($au.Kunde) | $($au.Baustelle) | $(if ($au.kWp) { [math]::Round($au.kWp, 1) }) | $(ChatDat $au.Liefertermin) | $(ChatStand $au.Stand 'Gel') | $(ChatStand $au.Stand 'Ber') | $(if ($au.Cn) { "$($au.Cn.P.Status)$(if (-not $au.Cn.Sicher) { ' (unsicher)' })" } else { '-' }) | $(@($mat) -join ', ')"
            }
            return "$($treffer.Count) offene Auftraege$(if ($treffer.Count -gt 60) { ', die ersten 60' }), sortiert: zuerst mit kommendem Montagetermin, dann nach Liefertermin. Link Liste: /`nNaechster Montagetermin (Team) | Auftrag (Link) | Kunde | Baustelle | kWp | Liefertermin | geliefert | berechnet | Craftnote-Status | Material`n" + ($z -join "`n")
        }
        'auftrag' {
            $liste = Chat-Auftraege
            $tr = @(ChatAuftragFinden $liste (ChatArg $inp 'beleg') (ChatArg $inp 'kunde'))
            if ($tr.Count -eq 0) { return 'Kein offener Auftrag gefunden (abgeschlossene Auftraege sind nicht enthalten). Tipp: offene_auftraege mit Kundenname.' }
            if ($tr.Count -gt 1) { return "Mehrere offene Auftraege passen: " + (($tr | ForEach-Object { "$($_.Beleg) $($_.Kunde)" }) -join '; ') + '. Bitte Auftragsnummer angeben.' }
            $au = $tr[0]
            $pos = Get-AuftragPositionen $au
            $termine = if ($au.Termine.Count) { ($au.Termine | ForEach-Object { "- $(ChatTermin $_)$(if ($_.T.Ende.Date -lt $heute) { ' (vergangen)' }): $($_.T.Titel)" }) -join "`n" } else { '- keine Kalendertermine gefunden' }
            $mat = foreach ($m in $pos.Material) {
                $fertig = $m.Gel -ge $m.Soll - 1e-9
                $st = if ($fertig) { 'geliefert' } elseif (-not $m.Lager) { 'kein Lagerartikel' } elseif ($m.Offen) { ChatMatStatus $m.Offen } else { '-' }
                "$(ChatText $m.Gruppe) | $($m.Art) | $($m.Bez)$(if ($m.Aus) { " (aus $($m.Aus))" }) | $(FmtZahl $m.Soll) | $(FmtZahl $m.Gel) | $(FmtZahl $m.Ber) | $st"
            }
            $lei = foreach ($l in $pos.Leistungen) { "$($l.Bez) | $(FmtZahl $l.Soll) | $(FmtZahl $l.Gel) | $(FmtZahl $l.Ber)" }
            $folge = (@($au.Folgebelege | ForEach-Object { "$($_['Belegart']) $($_['Beleg']) ($(ChatDat $_['Belegdatum']))" })) -join '; '
            try { $null = Craftnote-Zuordnen @($au) } catch {}
            $cnTxt = if ($au.Cn) { $p = $au.Cn.P; "Projekt $($p.Name), Status $($p.Status), zugeordnet ueber $($au.Cn.Wie)$(if (-not $au.Cn.Sicher) { ' (unsicher)' }), geaendert $(ChatDat $p.Geaendert), Link $CraftnoteWeb$($p.Id)$(if ($p.Notiz) { "`nCraftnote-Notiz ($($p.NotizAutor)): $($p.Notiz)" })" } else { "kein Projekt gefunden$(if ($script:CnFehler) { ' (' + (ChatText $script:CnFehler) + ')' })" }
            return "$($au.Art) $($au.Beleg) - $($au.Kunde)`nLink: /auftrag?bel=$($au.BelID)`nCraftnote: $cnTxt`nBaustelle: $((@($au.BaustelleName, $au.Baustelle) | Where-Object { $_ }) -join ', ')`nAuftragsdatum $(ChatDat $au.Datum), Liefertermin $(ChatDat $au.Liefertermin), Anlage $(if ($au.kWp) { [math]::Round($au.kWp, 1) } else { '-' }) kWp`nGeliefert: $(ChatStand $au.Stand 'Gel'), berechnet: $(ChatStand $au.Stand 'Ber'). Lieferscheine/Rechnungen: $(if ($folge) { $folge } else { 'keine' })`nMontagetermine:`n$termine`nMaterial: Gruppe | Artikel | Bezeichnung | Menge | geliefert | berechnet | Status`n" + ($mat -join "`n") + "`nLeistungen: Leistung | Menge | geliefert | berechnet`n" + ($lei -join "`n")
        }
        'materialbedarf' {
            $liste = Chat-Auftraege
            $tage = 28; $tv = ChatArg $inp 'tage'; if ($tv -match '^\d{1,4}$') { $tage = [int]$tv }
            $bis = if ($tage -gt 0) { $heute.AddDays($tage) } else { [datetime]::MaxValue }
            $art = [ordered]@{}; $ohne = New-Object System.Collections.ArrayList
            foreach ($au in $liste) {
                if ($au.Planung -gt $bis) { continue }
                foreach ($m in $au.Material) {
                    if (-not $m.Lager) { if ($m.Gruppe -ne 'Elektro') { [void]$ohne.Add("$(ChatDat $au.Planung) | $($au.Beleg) (/auftrag?bel=$($au.BelID)) $($au.Kunde) | $($m.Art) $($m.Bez) | $(FmtZahl $m.Menge)") }; continue }
                    if ($m.Status -ne 'fehlt') { continue }
                    if (-not $art.Contains($m.Art)) { $art[$m.Art] = @{ Bez = $m.Bez; Fehl = 0.0; Erster = $au; Anz = 0 } }
                    $art[$m.Art].Fehl += $m.Fehl; $art[$m.Art].Anz++
                }
            }
            $z = foreach ($k in $art.Keys) {
                $a = $art[$k]; $ek = ToNum $script:EkSumme[$k]; $bst = ToNum $script:Bestand[$k]
                "$k (/dispo?art=$k) | $($a.Bez) | fehlt $(FmtZahl $a.Fehl) | Bestand $(FmtZahl $bst) | bestellt $(if ($ek) { "$(FmtZahl $ek) ab $(ChatDat $script:EkTermin[$k])" } else { '-' }) | erster: $(ChatDat $a.Erster.Planung) $($a.Erster.Beleg) $($a.Erster.Kunde) | $($a.Anz) Auftraege"
            }
            $zeitraum = if ($tage -gt 0) { "Auftraege mit Planungstermin bis $($bis.ToString('dd.MM.yyyy'))" } else { 'alle offenen Auftraege' }
            return "$zeitraum. Link: /material?nur=fehlt&tage=$(if (@(14, 28, 0) -contains $tage) { $tage } else { 28 })`nFehlende Lagerartikel ($($art.Count)): Artikel | Bezeichnung | fehlt | Bestand | offene Bestellung | erster Auftrag | Anzahl`n" + $(if ($art.Count) { $z -join "`n" } else { 'keine' }) +
                   "`nOhne Lagerfuehrung, je Auftrag zu beschaffen ($($ohne.Count)): Planung | Auftrag | Artikel | Menge`n" + $(if ($ohne.Count) { (@($ohne) | Select-Object -First 60) -join "`n" } else { 'keine' })
        }
        'montagetermine' {
            $von = ChatDatum (ChatArg $inp 'von') $heute; $bis = ChatDatum (ChatArg $inp 'bis') $heute.AddDays(7)
            $team = ChatArg $inp 'team'; $text = ChatArg $inp 'text'
            $termine = @(Get-Kalender | Where-Object { $_.Ende.Date -ge $von -and $_.Start.Date -le $bis -and
                (-not $team -or (@($_.Team) -join ',') -like "*$team*") -and (-not $text -or $_.Titel -like "*$text*") } | Sort-Object Start)
            if ($script:KalFehler) { $hinweis = "Hinweis: $($script:KalFehler)`n" } else { $hinweis = '' }
            if ($termine.Count -eq 0) { return "$($hinweis)Keine Kalendertermine $($von.ToString('dd.MM.yyyy')) bis $($bis.ToString('dd.MM.yyyy')) fuer diese Auswahl." }
            # Zuordnung Termin -> Auftrag (wie auf der Seite Offene Auftraege)
            $zu = @{}
            foreach ($au in (Chat-Auftraege)) { foreach ($tr in $au.Termine) { $zu["$($tr.T.Start.Ticks)|$($tr.T.Titel)"] = @{ Au = $au; Sicher = $tr.Sicher } } }
            $z = foreach ($t in ($termine | Select-Object -First 120)) {
                $zeit = if ($t.Ende.Date -gt $t.Start.Date) { "$($t.Start.ToString('ddd dd.MM.', [Globalization.CultureInfo]'de-DE'))-$($t.Ende.ToString('dd.MM.'))" } else { $t.Start.ToString('ddd dd.MM.yyyy', [Globalization.CultureInfo]'de-DE') }
                if (-not $t.Ganztag) { $zeit += ' ' + $t.Start.ToString('HH:mm') }
                $a = $zu["$($t.Start.Ticks)|$($t.Titel)"]
                $auTxt = if ($a) { "$($a.Au.Beleg) (/auftrag?bel=$($a.Au.BelID))$(if (-not $a.Sicher) { ' [unsicher]' })" } else { '-' }
                "$zeit | $(@($t.Team) -join ', ') | $($t.Titel) | $auTxt"
            }
            return "$hinweis$($termine.Count) Termine $($von.ToString('dd.MM.yyyy')) bis $($bis.ToString('dd.MM.yyyy'))$(if ($termine.Count -gt 120) { ', die ersten 120' }). Zugeordnet werden nur offene Auftraege.`nDatum | Team | Termin | offener Auftrag (Link)`n" + ($z -join "`n")
        }
        'artikel_dispo' {
            $art = ChatArg $inp 'artikelnummer'
            $info = Run-Query $SQL_DispoArtikel @{ "@art" = $art }
            if ($info.Rows.Count -eq 0) { return "Artikel $art nicht gefunden. Tipp: artikel_suchen verwenden." }
            $ir = $info.Rows[0]
            $istHstl = (ToNum $ir['Stuecklistentyp']) -eq 2; $istLager = (ToNum $ir['Lagerfuehrung']) -eq -1
            $bestand = ToNum $ir['Bestand']; if ($null -eq $bestand) { $bestand = 0 }
            $kopf = "Artikel $art - $($ir['Bezeichnung'])`nLink: /dispo?art=$art"
            $liste = Chat-Auftraege
            $ev = New-Object System.Collections.ArrayList
            foreach ($au in $liste) {
                $tm = Dispo-Termin $au
                foreach ($m in $au.Material) {
                    if ($istHstl) { if ($m.KitArt -ne $art) { continue } }
                    elseif ($m.Art -ne $art) { continue }
                    [void]$ev.Add(@{ Datum = $tm.Datum; Montage = $tm.Montage; Au = $au; M = $m })
                }
            }
            if ($istHstl) {
                # je Auftrag + Position eine Zeile, schlechtester Status der Bestandteile
                $kits = [ordered]@{}
                foreach ($e in $ev) { $k = "$($e.Au.BelID)|$($e.M.KitPos)"; if (-not $kits.Contains($k)) { $kits[$k] = @{ E = $e; Teile = (New-Object System.Collections.ArrayList) } }; [void]$kits[$k].Teile.Add($e.M) }
                $z = foreach ($k in $kits.Values) {
                    $e = $k.E; $f = @($k.Teile | Where-Object { $_.Status -eq 'fehlt' }); $b = @($k.Teile | Where-Object { $_.Status -eq 'bestellt' })
                    $st = if ($f.Count) { 'fehlt: ' + (($f | ForEach-Object { "$($_.Art) $($_.Bez) fehlt $(FmtZahl $_.Fehl)" }) -join '; ') } elseif ($b.Count) { 'bestellt: ' + (($b | ForEach-Object { "$($_.Art) $(ChatMatStatus $_)" }) -join '; ') } else { 'da' }
                    "$(ChatDat $e.Datum)$(if ($e.Montage) { ' Montage' } else { ' Liefertermin' }) | $($e.Au.Beleg) (/auftrag?bel=$($e.Au.BelID)) | $($e.Au.Kunde) | $(FmtZahl $e.M.KitMenge) | $st"
                }
                $komp = Run-Query $SQL_DispoKomponenten @{ "@art" = $art }
                $kz = foreach ($r in $komp.Rows) { "$($r['Element']) | $($r['Bezeichnung']) | $(FmtZahl $r['Menge']) je Set | Bestand $(FmtZahl $script:Bestand[[string]$r['Element']])" }
                return "$kopf`nStueckliste (Set) ohne eigenen Bestand. Bestandteile:`n" + ($kz -join "`n") + "`nOffene Auftraege mit diesem Set ($($kits.Count)): Termin | Auftrag (Link) | Kunde | Menge | Status der Bestandteile`n" + $(if ($kits.Count) { $z -join "`n" } else { 'keine' })
            }
            $ekTab = Run-Query $SQL_DispoEk @{ "@art" = $art }
            foreach ($r in $ekTab.Rows) { $lt = $r['Liefertermin']; [void]$ev.Add(@{ Datum = $(if ($lt -is [datetime]) { $lt.Date } else { $null }); EK = $r }) }
            $sortiert = @($ev | Sort-Object @{ e = { if ($_.Datum -is [datetime]) { $_.Datum } else { [datetime]::MaxValue } } }, @{ e = { if ($_.EK) { 0 } else { 1 } } })
            $kum = [double]$bestand; $min = $kum; $minDat = 'heute'; $fehltAb = $null
            $z = foreach ($e in $sortiert) {
                if ($e.EK) {
                    $m = [double](ToNum $e.EK['Offen']); $kum += $m
                    "$(ChatDat $e.Datum)$(if ($e.Datum -is [datetime] -and $e.Datum -lt $heute) { ' (ueberfaellig)' }) | Bestellung $($e.EK['Beleg']) | $($e.EK['Lieferant']) | +$(FmtZahl $m) | $(FmtZahl $kum) | -"
                } else {
                    $m = [double]$e.M.Menge; $kum -= $m
                    if ($istLager -and $kum -lt $min) { $min = $kum; $minDat = ChatDat $e.Datum }
                    if ($istLager -and $kum -lt 0 -and $null -eq $fehltAb) { $fehltAb = ChatDat $e.Datum }
                    "$(ChatDat $e.Datum)$(if ($e.Montage) { ' Montage' } else { ' Liefertermin' }) | $($e.Au.Art) $($e.Au.Beleg) (/auftrag?bel=$($e.Au.BelID)) | $($e.Au.Kunde)$(if ($e.M.Aus) { " (aus Set $($e.M.Aus))" }) | -$(FmtZahl $m) | $(if ($istLager) { FmtZahl $kum } else { '-' }) | $(ChatMatStatus $e.M)"
                }
            }
            $fazit = if (-not $istLager) { 'Kein Lagerartikel - wird je Auftrag beschafft.' }
                     elseif ($null -ne $fehltAb) { "Bestand reicht NICHT: fehlt ab $fehltAb, tiefster Stand $(FmtZahl $min) am $minDat." }
                     elseif ($sortiert.Count) { "Alle offenen Bedarfe gedeckt, tiefster Stand $(FmtZahl $min) am $minDat." } else { 'Keine offenen Bedarfe oder Bestellungen.' }
            return "$kopf`nLagerbestand heute: $(FmtZahl $bestand)`n$fazit`nTermine vor heute sind ueberfaellig, aber noch offen.`nTermin | Beleg (Link) | Kunde/Lieferant | Menge | Kumuliert | Deckung`n" + ((@($z) | Select-Object -First 80) -join "`n") + $(if ($sortiert.Count -gt 80) { "`n... $($sortiert.Count - 80) weitere, siehe Link" })
        }
        'lagerbestand' {
            $art = ChatArg $inp 'artikelnummer'
            $info = Run-Query $SQL_DispoArtikel @{ "@art" = $art }
            if ($info.Rows.Count -eq 0) { return "Artikel $art nicht gefunden. Tipp: artikel_suchen verwenden." }
            $kopf = "Artikel $art - $($info.Rows[0]['Bezeichnung'])`nLink: /dispo?art=$art"
            if ((ToNum $info.Rows[0]['Stuecklistentyp']) -eq 2) {
                $null = Chat-Auftraege   # fuellt $script:Bestand
                $k = Run-Query $SQL_DispoKomponenten @{ "@art" = $art }
                $z = foreach ($r in $k.Rows) { "$($r['Element']) | $($r['Bezeichnung']) | $(FmtZahl $r['Menge']) je Set | $(FmtZahl $script:Bestand[[string]$r['Element']])" }
                return "$kopf`nStueckliste (kein eigener Bestand). Bestandteil | Bezeichnung | Menge | Bestand`n" + ($z -join "`n")
            }
            if ((ToNum $info.Rows[0]['Lagerfuehrung']) -ne -1) { return "$kopf`nKein Lagerartikel (wird je Auftrag beschafft), daher kein Bestand." }
            $t = Run-Query @"
SELECT lpb.Lagerkennung, lp.Dimension1, lp.Dimension2, lp.Dimension3, SUM(lpb.Bestand) AS Bestand
FROM KHKLagerplatzbestaende lpb
LEFT JOIN KHKLagerplaetze lp ON lp.PlatzID = lpb.PlatzID AND lp.Mandant = lpb.Mandant
WHERE lpb.Mandant = $Mandant AND lpb.Artikelnummer = @art AND lpb.Bestand <> 0
GROUP BY lpb.Lagerkennung, lp.Dimension1, lp.Dimension2, lp.Dimension3
ORDER BY SUM(lpb.Bestand) DESC
"@ @{ "@art" = $art }
            if ($t.Rows.Count -eq 0) { return "$kopf`nKein Bestand auf Lagerplaetzen." }
            $summe = 0; $z = foreach ($r in $t.Rows) { $summe += [double]$r['Bestand']; "$(Fmt-Platz $r['Lagerkennung'] $r['Dimension1'] $r['Dimension2'] $r['Dimension3']) | $(FmtZahl $r['Bestand'])" }
            return "$kopf`nGesamt: $(FmtZahl $summe). Wie viel davon schon fuer Auftraege verplant ist, zeigt artikel_dispo.`nLagerplatz | Bestand`n" + ($z -join "`n")
        }
        'offene_bestellungen' {
            $p = @{}; $filter = ''
            $art = ChatArg $inp 'artikelnummer'; if ($art) { $filter += ' AND p.Artikelnummer = @art'; $p['@art'] = $art }
            $lief = ChatArg $inp 'lieferant'; if ($lief) { $filter += " AND b.A0Name1 LIKE '%' + @lief + '%'"; $p['@lief'] = $lief }
            if ((ChatArg $inp 'nur_ueberfaellig') -eq 'True') { $filter += ' AND p.Liefertermin < CAST(GETDATE() AS date)' }
            $t = Run-Query @"
SELECT TOP 120 b.BelID, p.BelPosID, CAST(b.Belegjahr AS varchar) + '-' + CAST(b.Belegnummer AS varchar) AS Beleg, b.A0Name1 AS Lieferant, p.Artikelnummer, p.Bezeichnung1 AS Bezeichnung,
       (vp.GGBestellt - vp.GGGeliefert) AS Offen, p.Liefertermin
FROM KHKEKBelege b
INNER JOIN KHKEKBelegePositionen p ON p.BelID = b.BelID AND p.Mandant = b.Mandant
INNER JOIN KHKEKVorgaengePositionen vp ON vp.VorPosID = p.VorPosID AND vp.Mandant = p.Mandant
INNER JOIN KHKEKVorgaenge v ON v.VorID = vp.VorID AND v.Mandant = vp.Mandant
WHERE b.Mandant = $Mandant AND b.Belegart = 'Bestellung' AND v.GGErfuellt = 0 AND vp.GGErfuellt = 0 AND (vp.GGBestellt - vp.GGGeliefert) > 0 $filter
ORDER BY p.Liefertermin, b.Belegnummer
"@ $p
            if ($t.Rows.Count -eq 0) { return 'Keine offenen Bestellungen fuer diese Auswahl.' }
            $null = Chat-Auftraege
            $fuer = @{}; foreach ($e in $script:EkPositionen) { $fuer[[string]$e.Row['BelPosID']] = $e.Fuer }
            $z = foreach ($r in $t.Rows) {
                $fl = $fuer[[string]$r['BelPosID']]
                $fTxt = if ($fl -and $fl.Count) { (@($fl | ForEach-Object { "$($_.Au.Beleg) $($_.Au.Kunde) $(FmtZahl $_.Menge) (Planung $(ChatDat $_.Au.Planung))" }) -join '; ') } else { 'kein offener Auftrag (Lager)' }
                "$(ChatDat $r['Liefertermin'])$(if ($r['Liefertermin'] -is [datetime] -and $r['Liefertermin'].Date -lt $heute) { ' (ueberfaellig)' }) | $($r['Beleg']) (/ek?bel=$($r['BelID'])) | $($r['Lieferant']) | $($r['Artikelnummer']) | $($r['Bezeichnung']) | $(FmtZahl $r['Offen']) | $fTxt"
            }
            return "$($t.Rows.Count) offene Bestellpositionen$(if ($t.Rows.Count -ge 120) { ' (Limit 120 erreicht)' }). Link: /einkauf (ueberfaellig: /einkauf?nur=ueberfaellig, zu spaet fuer Montage: /einkauf?nur=spaet)`nLiefertermin | Bestellung (Link) | Lieferant | Artikel | Bezeichnung | Offen | fuer Auftraege (Zuteilung wie Offene Auftraege)`n" + ($z -join "`n")
        }
        'lagerbewegungen' {
            $von = ChatDatum (ChatArg $inp 'von') $heute.AddDays(-6)
            $bis = ChatDatum (ChatArg $inp 'bis') $heute
            $p = @{ "@max" = 150; "@von" = $von; "@bis" = $bis.AddDays(1) }
            $filter = @()
            $typ = ChatArg $inp 'art'
            if (@('um','zu','ab') -contains $typ) {
                $codes = @($BewegungsArten.Keys | Where-Object { $BewegungsArten[$_][0] -eq $typ } | ForEach-Object { "'$_'" }) -join ','
                $filter += "AND b.Bewegungsart IN ($codes)"
            }
            $art = ChatArg $inp 'artikelnummer'
            if ($art) { $filter += "AND (b.Artikelnummer LIKE @art + '%' OR a.Bezeichnung1 LIKE '%' + @art + '%')"; $p['@art'] = $art }
            $lager = ChatArg $inp 'lager'
            if ($lager) { $filter += "AND " + (ChatPlatzBedingung $lager @(@('hp', 'hp'), @('zp', 'zp')) $p) }
            $namen = Get-BenutzerNamen (Run-Query $SQL_BewegungenBenutzer $null) (Run-Query $SQL_MitarbeiterNamen $null)
            $person = ChatArg $inp 'person'
            if ($person) {
                $codes = @($namen.Keys | Where-Object { $namen[$_].Name -like "*$person*" -or $_ -like "*$person*" })
                if ($codes.Count -eq 0) { return "Keine Person mit Lagerbuchungen passt zu '$person'. Personen: " + ((@($namen.Values | ForEach-Object { $_.Name }) | Sort-Object -Unique) -join '; ') }
                $ph = for ($i = 0; $i -lt $codes.Count; $i++) { $p["@w$i"] = $codes[$i]; "@w$i" }
                $filter += "AND j.Benutzer IN ($($ph -join ','))"
            }
            $t = Run-Query ($SQL_Bewegungen -replace '\{FILTER\}', ($filter -join "`n  ")) $p
            $link = "/bewegungen?von=$($von.ToString('yyyy-MM-dd'))&bis=$($bis.ToString('yyyy-MM-dd'))$(if ($typ) { "&typ=$typ" })$(if ($art) { "&art=$([uri]::EscapeDataString($art))" })$(if ($lager) { "&platz=$([uri]::EscapeDataString($(if ($p['@lk']) { $p['@lk'] } else { $lager })))" })"
            if ($t.Rows.Count -eq 0) { return "Keine Lagerbuchungen $($von.ToString('dd.MM.yyyy')) bis $($bis.ToString('dd.MM.yyyy')) fuer diese Auswahl.`nLink: $link" }
            $z = foreach ($r in $t.Rows) {
                $b = [string]$r['Benutzer']; $n = if ($namen.ContainsKey($b)) { $namen[$b].Name } else { $b }
                $ba = [string]$r['Bewegungsart']; $bez = if ($BewegungsArten.ContainsKey($ba)) { ChatText $BewegungsArten[$ba][1] } else { $ba }
                $ref = ([string]$r['Referenz']).Trim(); if (-not $ref) { $ref = ([string]$r['Standardtext']).Trim() }
                $wer = if ($r['Kunde'] -isnot [System.DBNull]) { " Kunde $($r['Kunde'])" } elseif ($r['Lieferant'] -isnot [System.DBNull]) { " Lieferant $($r['Lieferant'])" } else { '' }
                "$(ChatDat $r['Bewegungsdatum']) | $n | $bez | $($r['Artikelnummer']) $($r['Bezeichnung']) | $(FmtZahl $r['Menge']) | $(Fmt-Platz $r['HKennung'] $r['HD1'] $r['HD2'] $r['HD3']) | $(Fmt-Platz $r['ZKennung'] $r['ZD1'] $r['ZD2'] $r['ZD3']) | $ref$wer"
            }
            return "$($t.Rows.Count) Buchungen$(if ($t.Rows.Count -ge 150) { ' (Limit 150 erreicht, es gibt mehr - Zeitraum eingrenzen)' }), $($von.ToString('dd.MM.yyyy')) bis $($bis.ToString('dd.MM.yyyy')), neueste zuerst.`nLink: $link`nDatum | Wer | Art | Artikel | Menge | Von | Nach | Bezug`n" + ($z -join "`n")
        }
        'lagerplatz_inhalt' {
            $lager = ChatArg $inp 'lager'
            if (-not $lager) { return 'Kein Lager angegeben.' }
            $p = @{}; $bed = ChatPlatzBedingung $lager @(,@('lpb', 'lp')) $p
            $t = Run-Query @"
SELECT TOP 200 lpb.Lagerkennung, lp.Dimension1, lp.Dimension2, lp.Dimension3, lpb.Artikelnummer, a.Bezeichnung1, SUM(lpb.Bestand) AS Bestand
FROM KHKLagerplatzbestaende lpb
LEFT JOIN KHKLagerplaetze lp ON lp.PlatzID = lpb.PlatzID AND lp.Mandant = lpb.Mandant
LEFT JOIN KHKArtikel a ON a.Artikelnummer = lpb.Artikelnummer AND a.Mandant = lpb.Mandant
WHERE lpb.Mandant = $Mandant AND lpb.Bestand <> 0 AND $bed
GROUP BY lpb.Lagerkennung, lp.Dimension1, lp.Dimension2, lp.Dimension3, lpb.Artikelnummer, a.Bezeichnung1
ORDER BY lpb.Lagerkennung, lp.Dimension1, lp.Dimension2, lp.Dimension3, lpb.Artikelnummer
"@ $p
            if ($t.Rows.Count -eq 0) { return "Kein Bestand auf Lagerplaetzen, die mit '$lager' beginnen." }
            $z = foreach ($r in $t.Rows) { "$(Fmt-Platz $r['Lagerkennung'] $r['Dimension1'] $r['Dimension2'] $r['Dimension3']) | $($r['Artikelnummer']) | $($r['Bezeichnung1']) | $(FmtZahl $r['Bestand'])" }
            return "$($t.Rows.Count) Eintraege$(if ($t.Rows.Count -ge 200) { ' (Limit 200 erreicht)' }).`nLagerplatz | Artikel | Bezeichnung | Bestand`n" + ($z -join "`n")
        }
        'craftnote_projekte' {
            $cn = Get-Craftnote
            if (-not $cn) { return "Craftnote nicht verfuegbar: $(ChatText $script:CnFehler)" }
            $stF = ChatArg $inp 'status'; $txt = ChatArg $inp 'text'
            $au = @(Chat-Auftraege); try { $null = Craftnote-Zuordnen $au } catch {}
            $zuAu = @{}; foreach ($a in $au) { if ($a.Cn) { $zuAu[$a.Cn.P.Id] = $a } }
            $alle = @($cn.Projekte | Where-Object { -not $_.Archiviert })
            $zaehl = (@($alle | Group-Object Status | Sort-Object Count -Descending | ForEach-Object { "$($_.Count) $(if ($_.Name) { $_.Name } else { 'ohne Status' })" })) -join '; '
            $t = @($alle | Where-Object { (-not $stF -or $_.Status -like "*$stF*") -and (-not $txt -or $_.Name -like "*$txt*" -or $_.Ort -like "*$txt*" -or $_.Plz -like "$txt*") } | Sort-Object { - $(if ($_.Geaendert) { $_.Geaendert.Ticks } else { 0 }) })
            $z = foreach ($p in ($t | Select-Object -First 80)) {
                $a = $zuAu[$p.Id]
                "$($p.Name) | $($p.Plz) $($p.Ort) | $($p.Status) | $(ChatDat $p.Geaendert) | $CraftnoteWeb$($p.Id) | $(if ($a) { "$($a.Beleg) (/auftrag?bel=$($a.BelID))" } else { '-' })$(if ($p.Notiz) { ' | Notiz: ' + $(if ($p.Notiz.Length -gt 150) { $p.Notiz.Substring(0, 150) + '...' } else { $p.Notiz }) })"
            }
            return "Craftnote-Projekte gesamt (nicht archiviert): $($alle.Count). Je Status: $zaehl`nTreffer: $($t.Count)$(if ($t.Count -gt 80) { ', die neuesten 80' }). Abgleich mit Sage: /abgleich`nProjekt | Ort | Status | zuletzt geaendert | Craftnote-Link | offener Sage-Auftrag`n" + ($z -join "`n")
        }
        default { return "Unbekanntes Werkzeug: $name" }
    }
}

# Fortschrittstext fuer das Chatfenster (HTML, wird per innerHTML gesetzt)
function ChatSchrittText($name, $inp) {
    $a = [System.Net.WebUtility]::HtmlEncode((ChatArg $inp 'artikelnummer'))
    switch ($name) {
        'artikel_suchen'      { "Suche Artikel &bdquo;$([System.Net.WebUtility]::HtmlEncode((ChatArg $inp 'suchtext')))&ldquo; &hellip;" }
        'offene_auftraege'    { "Lese offene Auftr&auml;ge &hellip;" }
        'auftrag'             { "Lese Auftrag &hellip;" }
        'materialbedarf'      { "Pr&uuml;fe Materialbedarf &hellip;" }
        'montagetermine'      { "Schaue in den Montagekalender &hellip;" }
        'artikel_dispo'       { "Schaue in die Dispo von $a &hellip;" }
        'lagerbestand'        { "Pr&uuml;fe Lagerbestand von $a &hellip;" }
        'offene_bestellungen' { "Lese offene Bestellungen &hellip;" }
        'lagerbewegungen'     { "Lese Lagerbewegungen &hellip;" }
        'lagerplatz_inhalt'   { "Lese Lagerplatz &hellip;" }
        'craftnote_projekte'  { "Schaue in Craftnote &hellip;" }
        default               { "Frage Sage ab &hellip;" }
    }
}

function Invoke-AnthropicApi($key, $body) {
    for ($versuch = 1; $versuch -le 3; $versuch++) {
        $wc = New-Object System.Net.WebClient
        $wc.Encoding = [System.Text.Encoding]::UTF8
        $wc.Headers.Add('x-api-key', $key)
        $wc.Headers.Add('anthropic-version', '2023-06-01')
        $wc.Headers.Add('content-type', 'application/json')
        try {
            $txt = $wc.UploadString('https://api.anthropic.com/v1/messages', 'POST', $body)
            return ($txt | ConvertFrom-Json)
        } catch {
            $ex = $_.Exception
            while ($ex -and $ex -isnot [System.Net.WebException]) { $ex = $ex.InnerException }
            $code = 0; $detail = $_.Exception.Message
            if ($ex -and $ex.Response) {
                $code = [int]$ex.Response.StatusCode
                $sr = New-Object System.IO.StreamReader($ex.Response.GetResponseStream(), [System.Text.Encoding]::UTF8)
                $detail = $sr.ReadToEnd(); $sr.Close()
            }
            # Ueberlastung / Rate-Limit: kurz warten und erneut versuchen
            if ($versuch -lt 3 -and @(429, 500, 502, 503, 529) -contains $code) { Start-Sleep -Seconds (3 * $versuch); continue }
            throw "API-Fehler $code`: $detail"
        } finally { $wc.Dispose() }
    }
}

# Laeuft im Hintergrund-Runspace. $status ist eine synchronisierte Hashtable fuer den Fortschritt.
function Invoke-ChatAnfrage($verlauf, $status) {
    $erg = @{ Antwort = ''; Fehler = ''; TokensEin = 0; CacheLesen = 0; CacheSchreiben = 0; TokensAus = 0; Werkzeuge = (New-Object System.Collections.ArrayList) }
    try {
        $key = Get-ChatApiKey
        if (-not $key) { $erg.Fehler = 'Kein API-Schl&uuml;ssel hinterlegt (ANTHROPIC_API_KEY).'; return $erg }
        $script:ChatAuZeit = $null   # pro Frage frische Auftragsdaten
        $msgs = New-Object System.Collections.ArrayList
        foreach ($m in $verlauf) { [void]$msgs.Add([ordered]@{ role = $m.role; content = $m.content }) }
        $de = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
        $system = $ChatSystemPrompt.Replace('{DATUM}', (Get-Date).ToString('dddd, dd.MM.yyyy', $de))
        $kopfJson = '{"model":' + (ConvertTo-JsonText $ChatModell) + ',"max_tokens":2000,' +
                    '"system":[{"type":"text","text":' + (ConvertTo-JsonText $system) + ',"cache_control":{"type":"ephemeral"}}],' +
                    '"tools":' + $ChatToolsJson + ',"messages":'
        for ($runde = 1; $runde -le $ChatMaxRunden; $runde++) {
            $status.Schritt = if ($runde -eq 1) { 'Denke nach &hellip;' } else { 'Werte die Daten aus &hellip;' }
            $resp = Invoke-AnthropicApi $key ($kopfJson + (ConvertTo-JsonText $msgs) + '}')
            $u = $resp.usage
            if ($u) {
                $erg.TokensEin += [int]$u.input_tokens
                $erg.CacheLesen += [int]$u.cache_read_input_tokens
                $erg.CacheSchreiben += [int]$u.cache_creation_input_tokens
                $erg.TokensAus += [int]$u.output_tokens
            }
            [void]$msgs.Add([ordered]@{ role = 'assistant'; content = @($resp.content) })
            if ($resp.stop_reason -eq 'tool_use') {
                $ergebnisse = New-Object System.Collections.ArrayList
                foreach ($blk in @($resp.content)) {
                    if ($blk.type -ne 'tool_use') { continue }
                    [void]$erg.Werkzeuge.Add($blk.name)
                    $status.Schritt = ChatSchrittText $blk.name $blk.input
                    $istFehler = $false
                    try { $txt = ChatKuerzen ([string](Invoke-ChatTool $blk.name $blk.input)) 20000 }
                    catch { $txt = "Fehler beim Abfragen: $($_.Exception.Message)"; $istFehler = $true }
                    [void]$ergebnisse.Add([ordered]@{ type = 'tool_result'; tool_use_id = $blk.id; content = $txt; is_error = $istFehler })
                }
                [void]$msgs.Add([ordered]@{ role = 'user'; content = $ergebnisse })
                continue
            }
            $texte = @($resp.content | Where-Object { $_.type -eq 'text' } | ForEach-Object { $_.text })
            $erg.Antwort = ($texte -join "`n").Trim()
            if ($resp.stop_reason -eq 'max_tokens') { $erg.Antwort += "`n`n_(Antwort wurde abgeschnitten.)_" }
            return $erg
        }
        $erg.Fehler = 'Die Frage braucht zu viele Abfragen auf einmal &ndash; bitte konkreter oder in Teilfragen stellen.'
    } catch {
        $erg.Fehler = [System.Net.WebUtility]::HtmlEncode($_.Exception.Message)
    }
    return $erg
}

# Chatfenster (unten rechts). Einfaches HTML/ES5, Verlauf je Browser-Tab in sessionStorage.
# Antworten werden als kleines Markdown-Subset gerendert (Fett, Listen, Tabellen, interne Links).
$ChatWidgetHtml = @'
<style>
#chatknopf { position:fixed; right:20px; bottom:20px; z-index:1000; background:#343A40; color:#fff; border:2px solid #FFBC00; border-radius:24px; font-family:inherit;
  padding:11px 18px; font-size:14px; font-weight:600; cursor:pointer; box-shadow:0 3px 10px rgba(0,0,0,.25); display:flex; align-items:center; gap:8px; }
#chatknopf:hover { background:#23272b; }
#chatpanel { position:fixed; right:20px; bottom:20px; z-index:1001; width:440px; max-width:calc(100vw - 32px); height:640px; max-height:calc(100vh - 40px);
  background:#fff; border-radius:12px; box-shadow:0 8px 30px rgba(0,0,0,.3); display:none; flex-direction:column; overflow:hidden; font-size:14px; }
#chatpanel.offen { display:flex; }
.chatkopf { background:#343A40; color:#fff; padding:10px 14px; display:flex; align-items:center; gap:8px; }
.chatkopf b { flex:1; font-size:15px; }
.chatkopf .beta { font-size:10px; background:#FFBC00; color:#343A40; border-radius:8px; padding:1px 6px; font-weight:700; }
.chatkopf button { background:none; border:1px solid #5d7286; color:#dfe6ed; border-radius:5px; padding:3px 8px; cursor:pointer; font-size:12px; }
.chatkopf button:hover { background:#23272b; color:#fff; }
#chatverlauf { flex:1; overflow-y:auto; padding:12px; background:#F8F9FA; display:flex; flex-direction:column; gap:10px; }
.cm { max-width:92%; padding:8px 12px; border-radius:10px; line-height:1.45; overflow-wrap:anywhere; }
.cm-user { align-self:flex-end; background:#343A40; color:#fff; white-space:pre-wrap; }
.cm-bot { align-self:flex-start; background:#fff; border:1px solid #e1e4e8; }
.cm-bot p { margin:0 0 6px; } .cm-bot p:last-child { margin-bottom:0; }
.cm-bot ul, .cm-bot ol { margin:4px 0 6px; padding-left:20px; }
.cm-bot a { color:#086ED5; font-weight:600; text-decoration:none; } .cm-bot a:hover { text-decoration:underline; }
.cm-bot table { font-size:12px; box-shadow:none; margin:4px 0; border:1px solid #e1e4e8; }
.cm-bot th { position:static; background:#fff8e0; color:#1a1a1a; cursor:default; padding:4px 6px; }
.cm-bot td { padding:4px 6px; }
.cm-bot h4 { margin:6px 0 4px; font-size:14px; }
.cm-warte { align-self:flex-start; color:#666; font-style:italic; background:none; padding:4px 2px; }
.cm-fehler { align-self:flex-start; background:#fdecea; color:#b71c1c; }
.chatfuss { border-top:1px solid #e1e4e8; padding:8px; display:flex; gap:8px; align-items:flex-end; }
#chateingabe { flex:1; resize:none; border:1px solid #ccc; border-radius:6px; padding:8px; font:inherit; font-size:14px; max-height:120px; }
#chatsenden { background:#343A40; color:#fff; border:none; border-radius:999px; font-family:inherit; padding:9px 14px; cursor:pointer; font-weight:600; }
#chatsenden:disabled { opacity:.5; cursor:default; }
.chathinweis { font-size:11px; color:#888; padding:0 10px 8px; }
.chatleer { color:#666; font-size:13px; }
.chatleer a { display:block; margin:6px 0; color:#086ED5; cursor:pointer; }
.bewertung { margin-top:6px; padding-top:6px; border-top:1px solid #eef0f2; display:flex; flex-wrap:wrap; gap:6px; align-items:center; font-size:12px; color:#888; }
.bewertung button { background:#fff; border:1px solid #d0d5db; border-radius:12px; padding:1px 9px; cursor:pointer; font-size:13px; line-height:1.6; }
.bewertung button:hover { background:#fff8e0; }
.bewertung textarea { width:100%; border:1px solid #ccc; border-radius:6px; padding:6px; font:inherit; font-size:13px; resize:vertical; }
.bewertung .senden { background:#343A40; color:#fff; border-color:#343A40; border-radius:6px; font-size:12px; padding:3px 10px; }
</style>
<button id='chatknopf' type='button' onclick='chatAuf(true)'>
  <svg width='18' height='18' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z'/></svg>
  Frage stellen
</button>
<div id='chatpanel' role='dialog' aria-label='Assistent'>
  <div class='chatkopf'><b>Assistent</b><span class='beta'>TEST</span>
    <button type='button' onclick='chatNeu()' title='Gespr&auml;ch leeren'>Neu</button>
    <button type='button' onclick='chatAuf(false)' title='Schlie&szlig;en'>&times;</button></div>
  <div id='chatverlauf'></div>
  <div class='chatfuss'>
    <textarea id='chateingabe' rows='2' placeholder='Frage zu Auftr&auml;gen, Montage, Material, Lager ...'></textarea>
    <button id='chatsenden' type='button' onclick='chatSenden()'>Senden</button>
  </div>
  <div class='chathinweis'>Antworten basieren auf Sage-Daten, k&ouml;nnen aber Fehler enthalten &ndash; wichtige Zahlen &uuml;ber die Links pr&uuml;fen.</div>
</div>
<script>
(function(){
  var verlauf = [], laeuft = false;
  function lies(k){ try { return sessionStorage.getItem(k); } catch(e) { return null; } }
  function schreib(k, v){ try { sessionStorage.setItem(k, v); } catch(e) {} }
  try { verlauf = JSON.parse(lies('solarChat') || '[]') || []; } catch(e) { verlauf = []; }
  var panel = document.getElementById('chatpanel'), box = document.getElementById('chatverlauf'),
      ein = document.getElementById('chateingabe'), btn = document.getElementById('chatsenden'), knopf = document.getElementById('chatknopf');

  function esc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }
  function inline(s){
    s = esc(s);
    s = s.replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>');
    s = s.replace(/(^|[^*])\*([^*\s][^*]*)\*/g, '$1<i>$2</i>');
    s = s.replace(/`([^`]+)`/g, '<code>$1</code>');
    s = s.replace(/\[([^\]]+)\]\((\/[^)\s]*)\)/g, function(m, t, h){ return "<a href='" + h.replace(/'/g,'%27') + "'>" + t + '</a>'; });
    s = s.replace(/\[([^\]]+)\]\((https:\/\/app\.mycraftnote\.de\/[^)\s]*)\)/g, function(m, t, h){ return "<a href='" + h.replace(/'/g,'%27') + "' target='_blank'>" + t + ' &#8599;</a>'; });
    return s;
  }
  function md(text){
    var zeilen = String(text).split('\n'), html = '', i = 0, liste = null;
    function listeZu(){ if (liste) { html += '</' + liste + '>'; liste = null; } }
    while (i < zeilen.length) {
      var z = zeilen[i], m;
      if (/^\s*\|.*\|\s*$/.test(z)) {
        listeZu();
        var t = '<table>', kopf = true;
        while (i < zeilen.length && /^\s*\|.*\|\s*$/.test(zeilen[i])) {
          var zz = zeilen[i].trim();
          if (/^\|[\s:|-]+\|$/.test(zz)) { i++; continue; }
          var zellen = zz.replace(/^\||\|$/g, '').split('|'), tag = kopf ? 'th' : 'td';
          t += '<tr>';
          for (var c = 0; c < zellen.length; c++) t += '<' + tag + '>' + inline(zellen[c].trim()) + '</' + tag + '>';
          t += '</tr>'; kopf = false; i++;
        }
        html += t + '</table>'; continue;
      }
      if ((m = /^\s*[-*]\s+(.*)$/.exec(z))) { if (liste !== 'ul') { listeZu(); html += '<ul>'; liste = 'ul'; } html += '<li>' + inline(m[1]) + '</li>'; i++; continue; }
      if ((m = /^\s*\d+[.)]\s+(.*)$/.exec(z))) { if (liste !== 'ol') { listeZu(); html += '<ol>'; liste = 'ol'; } html += '<li>' + inline(m[1]) + '</li>'; i++; continue; }
      listeZu();
      if ((m = /^#{1,6}\s+(.*)$/.exec(z))) { html += '<h4>' + inline(m[1]) + '</h4>'; i++; continue; }
      if (z.trim() === '') { i++; continue; }
      var absatz = [z]; i++;
      while (i < zeilen.length && zeilen[i].trim() !== '' && !/^\s*([-*]|\d+[.)]|\||#)/.test(zeilen[i])) { absatz.push(zeilen[i]); i++; }
      html += '<p>' + absatz.map(inline).join('<br>') + '</p>';
    }
    listeZu();
    return html;
  }
  function blase(cls, inhaltHtml){ var d = document.createElement('div'); d.className = 'cm ' + cls; d.innerHTML = inhaltHtml; box.appendChild(d); box.scrollTop = box.scrollHeight; return d; }
  function bewerten(idx, wert, kommentar){
    var e = verlauf[idx]; if (!e) return;
    e.bewertet = wert; schreib('solarChat', JSON.stringify(verlauf));
    var frage = (idx > 0 && verlauf[idx - 1].role === 'user') ? verlauf[idx - 1].content : '';
    var x = new XMLHttpRequest();
    x.open('POST', '/api/chat-bewertung'); x.setRequestHeader('Content-Type', 'application/json; charset=utf-8');
    x.send(JSON.stringify({ id: e.id || '', bewertung: wert, frage: frage, antwort: e.content, kommentar: kommentar || '' }));
  }
  function leiste(d, idx){
    var e = verlauf[idx]; if (!e.id) return;
    var l = document.createElement('div'); l.className = 'bewertung'; d.appendChild(l);
    if (e.bewertet) { l.innerHTML = (e.bewertet === 'gut' ? '&#128077;' : '&#128078;') + ' Danke f&uuml;r die R&uuml;ckmeldung'; return; }
    l.innerHTML = 'Hilfreich? <button type="button" title="Antwort war richtig und hilfreich">&#128077;</button><button type="button" title="Antwort war falsch oder nicht hilfreich">&#128078;</button>';
    var b = l.getElementsByTagName('button');
    b[0].onclick = function(){ bewerten(idx, 'gut', ''); l.innerHTML = '&#128077; Danke f&uuml;r die R&uuml;ckmeldung'; };
    b[1].onclick = function(){
      l.innerHTML = '<textarea rows="2" placeholder="Was war falsch oder hat gefehlt? (optional)"></textarea><button type="button" class="senden">Absenden</button>';
      var ta = l.getElementsByTagName('textarea')[0]; ta.focus();
      l.getElementsByTagName('button')[0].onclick = function(){ bewerten(idx, 'schlecht', ta.value.trim()); l.innerHTML = '&#128078; Danke &ndash; wir schauen uns das an.'; };
      box.scrollTop = box.scrollHeight;
    };
  }
  function botBlase(idx){ var d = blase('cm-bot', md(verlauf[idx].content)); leiste(d, idx); box.scrollTop = box.scrollHeight; }
  function zeichne(){
    box.innerHTML = '';
    if (verlauf.length === 0) {
      var d = blase('chatleer', 'Ich beantworte Fragen zu Auftr&auml;gen, Montageterminen, Material und Lager mit den aktuellen Daten aus Sage und dem Montagekalender. Zum Beispiel:');
      ['Was fehlt f\u00fcr die Montagen der n\u00e4chsten zwei Wochen?', 'Welche Montagen stehen n\u00e4chste Woche an und mit welchem Team?', 'Wie ist der Stand beim Auftrag Kafitz?', 'Wann ist der Speicher BYD HVS 5.1 wieder vollst\u00e4ndig da?'].forEach(function(q){
        var a = document.createElement('a'); a.textContent = q; a.onclick = function(){ ein.value = q; chatSenden(); }; d.appendChild(a);
      });
      return;
    }
    for (var i = 0; i < verlauf.length; i++) {
      if (verlauf[i].role === 'user') blase('cm-user', esc(verlauf[i].content));
      else botBlase(i);
    }
  }
  window.chatAuf = function(auf){
    panel.className = auf ? 'offen' : ''; knopf.style.display = auf ? 'none' : 'flex';
    schreib('solarChatOffen', auf ? '1' : '0');
    if (auf) { zeichne(); setTimeout(function(){ ein.focus(); }, 50); }
  };
  window.chatNeu = function(){ if (laeuft) return; verlauf = []; schreib('solarChat', '[]'); zeichne(); };
  window.chatSenden = function(){
    var q = ein.value.trim(); if (!q || laeuft) return;
    if (verlauf.length === 0) box.innerHTML = '';
    verlauf.push({ role: 'user', content: q }); schreib('solarChat', JSON.stringify(verlauf));
    blase('cm-user', esc(q)); ein.value = '';
    laeuft = true; btn.disabled = true;
    var warte = blase('cm-warte', 'Denke nach &hellip;');
    function ende(){ laeuft = false; btn.disabled = false; if (warte.parentNode) warte.parentNode.removeChild(warte); }
    function fehler(t){ ende(); verlauf.pop(); schreib('solarChat', JSON.stringify(verlauf)); blase('cm-fehler', t); }
    var x = new XMLHttpRequest();
    x.open('POST', '/api/chat'); x.setRequestHeader('Content-Type', 'application/json; charset=utf-8');
    x.onload = function(){
      var r; try { r = JSON.parse(x.responseText); } catch(e) { return fehler('Unerwartete Antwort vom Server.'); }
      if (!r.id) return fehler(r.fehler || 'Chat ist nicht verf&uuml;gbar.');
      var start = new Date().getTime();
      (function frage(){
        var y = new XMLHttpRequest();
        y.open('GET', '/api/chat?id=' + encodeURIComponent(r.id));
        y.onload = function(){
          var s; try { s = JSON.parse(y.responseText); } catch(e) { return fehler('Unerwartete Antwort vom Server.'); }
          if (s.status === 'fertig') {
            ende(); verlauf.push({ role: 'assistant', content: s.antwort, id: r.id }); schreib('solarChat', JSON.stringify(verlauf));
            botBlase(verlauf.length - 1); return;
          }
          if (s.status === 'fehler') return fehler(s.fehler || 'Fehler.');
          if (new Date().getTime() - start > 180000) return fehler('Keine Antwort nach 3 Minuten &ndash; bitte erneut versuchen.');
          if (s.schritt) warte.innerHTML = s.schritt;
          setTimeout(frage, 1000);
        };
        y.onerror = function(){ fehler('Verbindung zum Server unterbrochen.'); };
        y.send();
      })();
    };
    x.onerror = function(){ fehler('Server nicht erreichbar.'); };
    x.send(JSON.stringify({ messages: verlauf.slice(-12).map(function(m){ return { role: m.role, content: m.content }; }) }));
  };
  ein.addEventListener('keydown', function(e){ if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); chatSenden(); } });
  if (lies('solarChatOffen') === '1') chatAuf(true);
})();
</script>
'@


# ---------- Einkauf: offene Bestellungen ----------
# Je offene Bestellposition auch, fuer welche Auftraege die Ware gebraucht wird: gleiche Zuteilung wie auf
# "Offene Auftraege" (Bestand zuerst, dann Bestellungen nach Liefertermin an die Auftraege nach Montage-/Liefertermin).
# "Zu spaet" = Liefertermin liegt nach dem ersten Termin, fuer den die Position gebraucht wird.
$SQL_EkBeleg = @"
SELECT b.BelID, CAST(b.Belegjahr AS varchar) + '-' + CAST(b.Belegnummer AS varchar) AS Beleg, b.A0Name1 AS Lieferant, b.Belegdatum,
       p.BelPosID, p.Position, p.Artikelnummer, p.Bezeichnung1 AS Bezeichnung, p.Liefertermin, vp.GGBestellt AS Bestellt, vp.GGGeliefert AS Geliefert
FROM KHKEKBelege b
INNER JOIN KHKEKBelegePositionen p ON p.BelID = b.BelID AND p.Mandant = b.Mandant
INNER JOIN KHKEKVorgaengePositionen vp ON vp.VorPosID = p.VorPosID AND vp.Mandant = p.Mandant
WHERE b.Mandant = $Mandant AND b.BelID = @bel AND p.PosType = 1
ORDER BY p.BelPosID
"@

# Bewertung einer offenen Bestellposition: erster Bedarfstermin, ueberfaellig, zu spaet
function Ek-Lage($e) {
    $heute = (Get-Date).Date
    $lt = $e.Row["Liefertermin"]
    $plan = @($e.Fuer | ForEach-Object { $_.Au.Planung } | Where-Object { $_ -ne [datetime]::MaxValue } | Sort-Object)
    $erster = if ($plan.Count) { $plan[0] } else { $null }
    $ueberf = ($lt -is [datetime] -and $lt.Date -lt $heute)
    $spaet = ($null -ne $erster -and ($lt -isnot [datetime] -or $lt.Date -gt $erster))
    return @{ Erster = $erster; Ueberfaellig = $ueberf; Spaet = $spaet }
}
# Zelle "fuer Auftraege": erster Auftrag als Link, weitere im Tooltip
function Ek-FuerZelle($e) {
    if (-not $e -or -not $e.Fuer.Count) { return "<span class='leise' title='Keinem offenen Auftrag zugeteilt &ndash; f&uuml;rs Lager'>Lager</span>" }
    $liste = @($e.Fuer | Sort-Object { $_.Au.Planung })
    $tip = HtmlEncode ((@($liste | ForEach-Object { "$(if ($_.Au.Planung -ne [datetime]::MaxValue) { $_.Au.Planung.ToString('dd.MM.') } else { 'ohne Termin' })  $($_.Au.Beleg)  $($_.Au.Kunde): $(FmtZahl $_.Menge)" })) -join "`n")
    $a = $liste[0].Au
    $d = if ($a.Planung -ne [datetime]::MaxValue) { "$(Datum-Kurz $a.Planung) " } else { "" }
    $mehr = if ($liste.Count -gt 1) { " <span class='leise'>+$($liste.Count - 1)</span>" } else { "" }
    return "<span title='$tip'>$d<a class='artlink' href='/auftrag?bel=$($a.BelID)'>$(HtmlEncode $a.Kunde)</a>$mehr</span>"
}

function Build-Einkauf($eks, $nur, $art, $fehler) {
    $head = Page-Head "Solar Dispo - Einkauf"
    $heute = (Get-Date).Date
    $inhalt = ""; $nAlle = 0; $nUeb = 0; $nSpaet = 0
    if ($fehler) { $inhalt = "<div class='fehler'>Fehler beim Laden der Daten:<br><pre>$(HtmlEncode $fehler)</pre></div>" }
    else {
        $liste = @(foreach ($e in $eks) { if ($art -and [string]$e.Row["Artikelnummer"] -ne $art) { continue }; @{ E = $e; L = (Ek-Lage $e) } })
        $nAlle = $liste.Count; $nUeb = @($liste | Where-Object { $_.L.Ueberfaellig }).Count; $nSpaet = @($liste | Where-Object { $_.L.Spaet }).Count
        $zeige = switch ($nur) { 'ueberfaellig' { @($liste | Where-Object { $_.L.Ueberfaellig }) } 'spaet' { @($liste | Where-Object { $_.L.Spaet }) } default { $liste } }
        if ($zeige.Count -eq 0) { $inhalt = "<p class='leise'>Keine offenen Bestellungen f&uuml;r diese Auswahl.</p>" }
        else {
            $rows = foreach ($x in $zeige) {
                $r = $x.E.Row; $l = $x.L; $an = HtmlEncode $r["Artikelnummer"]
                $lt = $r["Liefertermin"]
                $ltTxt = if ($lt -is [datetime]) { "$(Datum-Kurz $lt) $($lt.Year)" } else { "<span class='leise'>ohne Termin</span>" }
                if ($l.Ueberfaellig) { $ltTxt = "<span class='warn' title='$(($heute - $lt.Date).Days) Tage &uuml;berf&auml;llig'>$ltTxt</span>" }
                $st = if ($l.Spaet) { "<span class='chip fehlt' title='Gebraucht ab $($l.Erster.ToString('dd.MM.yyyy'))'>zu sp&auml;t</span>" } else { "" }
                "<tr><td style='white-space:nowrap'>$ltTxt</td><td><a class='artlink' href='/ek?bel=$($r['BelID'])'>$(HtmlEncode $r['Beleg'])</a></td><td>$(HtmlEncode $r['Lieferant'])</td>" +
                "<td><a class='artlink' href='/dispo?art=$an'>$an</a></td><td>$(HtmlEncode $r['Bezeichnung'])</td>" +
                "<td class='num'>$(FmtZahl $r['Bestellt'])</td><td class='num'><b>$(FmtZahl $r['Offen'])</b></td><td style='white-space:nowrap'>$(Ek-FuerZelle $x.E)</td><td>$st</td></tr>"
            }
            $inhalt = @"
<table id='tab'>
<thead><tr><th>Liefertermin</th><th>Bestellung</th><th>Lieferant</th><th>Artikel</th><th>Bezeichnung</th><th class='num'>Bestellt</th><th class='num'>Offen</th>
<th title='Zuteilung wie auf &bdquo;Offene Auftr&auml;ge&ldquo;: erster Auftrag mit Montage- bzw. Liefertermin, weitere im Tooltip'>F&uuml;r Auftr&auml;ge</th><th></th></tr></thead>
<tbody>$($rows -join '')</tbody>
</table>
"@
        }
    }
    $artQ = if ($art) { "&amp;art=$([uri]::EscapeDataString($art))" } else { "" }
    $knopf = { param($n, $text, $anz) $cls = if ($n -eq $nur) { 'btn' } else { 'btn btn2' }; "<a class='$cls' href='/einkauf?nur=$n$artQ'>$text ($anz)</a>" }
    $artChip = if ($art) { "<span class='hinweisbox' style='margin:0'>Artikel <b>$(HtmlEncode $art)</b> &middot; <a href='/einkauf$(if ($nur) { "?nur=$nur" })'>alle Artikel</a></span>" } else { "" }
    return @"
$head
<div class='kopf'><h1>Einkauf &ndash; offene Bestellungen</h1><div class='meta'>$(StandText) &middot; $nAlle Positionen</div></div>
<div class='steuer'>$(& $knopf '' 'Alle' $nAlle) $(& $knopf 'ueberfaellig' '&Uuml;berf&auml;llig' $nUeb) $(& $knopf 'spaet' 'Zu sp&auml;t f&uuml;r Montage' $nSpaet)
  <input class='suche' id='suche' placeholder='Filtern ...' onkeyup='filter()'> $artChip</div>
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

function Build-EkDetail($tab, $eks, $fehler) {
    $head = Page-Head "Solar Dispo - Bestellung"
    if ($fehler) { return "$head<div class='fehler'>Fehler beim Laden der Daten:<br><pre>$(HtmlEncode $fehler)</pre></div></body></html>" }
    if (-not $tab -or $tab.Rows.Count -eq 0) { return "$head<p class='leise'>Bestellung nicht gefunden.</p></body></html>" }
    $k = $tab.Rows[0]
    $offen = @{}; foreach ($e in $eks) { $offen[[string]$e.Row["BelPosID"]] = $e }
    $rows = foreach ($r in $tab.Rows) {
        $soll = [double](ToNum $r["Bestellt"]); $gel = [double](ToNum $r["Geliefert"]); $rest = $soll - $gel
        $e = $offen[[string]$r["BelPosID"]]
        $an = HtmlEncode $r["Artikelnummer"]
        $lt = $r["Liefertermin"]; $ltTxt = if ($lt -is [datetime]) { $lt.ToString('dd.MM.yyyy') } else { "" }
        $st = ""
        if ($e) {
            $l = Ek-Lage $e
            if ($l.Ueberfaellig) { $ltTxt = "<span class='warn'>$ltTxt</span>" }
            if ($l.Spaet) { $st = "<span class='chip fehlt' title='Gebraucht ab $($l.Erster.ToString('dd.MM.yyyy'))'>zu sp&auml;t</span>" }
        }
        $fertig = $rest -le 1e-9
        "<tr$(if ($fertig) { " class='erledigt'" })><td>$(HtmlEncode $r['Position'])</td><td><a class='artlink' href='/dispo?art=$an'>$an</a></td><td>$(HtmlEncode $r['Bezeichnung'])</td>" +
        "<td style='white-space:nowrap'>$ltTxt</td><td class='num'>$(FmtZahl $soll)</td>$(Mengen-Zelle $gel $soll 'geliefert' $false)<td class='num'>$(if (-not $fertig) { "<b>$(FmtZahl $rest)</b>" })</td>" +
        "<td style='white-space:nowrap'>$(if (-not $fertig) { Ek-FuerZelle $e })</td><td>$st</td></tr>"
    }
    $bd = $k["Belegdatum"]
    return @"
$head
<div class='kopf'><h1>Bestellung $(HtmlEncode $k['Beleg']) &ndash; $(HtmlEncode $k['Lieferant'])</h1><div class='meta'>$(StandText)</div></div>
<div class='steuer'><a class='btn btn2' href='/einkauf'>&larr; Einkauf</a></div>
<table style='max-width:520px;margin-bottom:16px'><tbody>
<tr><td>Lieferant</td><td>$(HtmlEncode $k['Lieferant'])</td></tr>
<tr><td>Bestelldatum</td><td>$(if ($bd -is [datetime]) { $bd.ToString('dd.MM.yyyy') })</td></tr>
</tbody></table>
<table>
<thead><tr><th>Pos.</th><th>Artikel</th><th>Bezeichnung</th><th>Liefertermin</th><th class='num'>Bestellt</th><th class='num'>Geliefert</th><th class='num'>Offen</th><th>F&uuml;r Auftr&auml;ge</th><th></th></tr></thead>
<tbody>$($rows -join '')</tbody>
</table>
</body></html>
"@
}


# ---------- Craftnote (Projekt-App): Status, Notizen, Dateien, Zeiten - NUR LESEN ----------
# API-Schluessel als Umgebungsvariable CRAFTNOTE_API_KEY (wie ein Passwort behandeln, nie ins Skript).
# Der Schluessel darf in Craftnote auch schreiben - hier werden ausschliesslich GET-Abfragen verwendet.
# Aufgaben und Chat-Nachrichten brauchen den kostenpflichtigen "Advanced API access" und fehlen daher.
# Zuordnung zu Sage: 1. gleiche E-Mail-Adresse (Kunde in Sage / Kontakt im Projekt), 2. Nachname + PLZ.
$CraftnoteBasis = 'https://functions.craftnote.cloud/api/v1'
$CraftnoteWeb = 'https://app.mycraftnote.de/#/project?id='
$CraftnoteCacheMinuten = 15
$CraftnoteEigeneMails = '(?i)@(leyendecker\.solar|lorenz-ms\.de)$'   # eigene Mitarbeiter als Projektkontakt - nicht zuordnen
$SQL_KundenMails = @"
SELECT kk.Kto, LOWER(LTRIM(RTRIM(ad.EMail))) AS Mail FROM KHKKontokorrent kk
INNER JOIN KHKAdressen ad ON ad.Adresse = kk.Adresse AND ad.Mandant = kk.Mandant
WHERE kk.Mandant = $Mandant AND ad.EMail LIKE '%@%'
UNION
SELECT kk.Kto, LOWER(LTRIM(RTRIM(ap.EMail))) FROM KHKKontokorrent kk
INNER JOIN KHKAnsprechpartner ap ON ap.Adresse = kk.Adresse AND ap.Mandant = kk.Mandant
WHERE kk.Mandant = $Mandant AND ap.EMail LIKE '%@%'
"@
# letzter Auftrag je Kunde (fuer den Abgleich: Craftnote "Auftrag terminiert", in Sage aber nichts offen)
$SQL_LetzteAuftraege = @"
SELECT x.Kto, x.Beleg, x.BelID, x.Belegdatum, x.Erledigt FROM (
    SELECT k.A0Empfaenger AS Kto, CAST(k.Belegjahr AS varchar) + '-' + CAST(k.Belegnummer AS varchar) AS Beleg, k.BelID, k.Belegdatum,
           CASE WHEN v.GGErfuellt = 0 THEN 0 ELSE 1 END AS Erledigt,
           ROW_NUMBER() OVER (PARTITION BY k.A0Empfaenger ORDER BY k.Belegdatum DESC, k.BelID DESC) AS nr
    FROM KHKVKBelege k INNER JOIN KHKVKVorgaenge v ON v.VorID = k.VorID AND v.Mandant = k.Mandant
    WHERE k.Mandant = $Mandant AND (k.Belegart LIKE 'Auftragsbest%' OR k.Belegart LIKE 'Reservierungsbest%' OR k.Belegart = 'AB (aus Res)')
) x WHERE x.nr = 1
"@

function CN-Abruf($pfad) {
    $key = Umgebung 'CRAFTNOTE_API_KEY'
    if (-not $key) { throw 'Craftnote-Schluessel nicht hinterlegt (CRAFTNOTE_API_KEY)' }
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    $wc = New-Object System.Net.WebClient; $wc.Encoding = [System.Text.Encoding]::UTF8
    $wc.Headers.Add('X-CN-API-KEY', $key.Trim()); $wc.Headers.Add('Accept', 'application/json')
    try { return ($wc.DownloadString("$CraftnoteBasis$pfad") | ConvertFrom-Json) } finally { $wc.Dispose() }
}
# Liste ueber alle Seiten (Token); das Listenfeld heisst je Endpunkt anders (data, files, times, members)
function CN-Liste($pfad) {
    $erg = New-Object System.Collections.ArrayList; $nach = $null
    for ($i = 0; $i -lt 100; $i++) {
        $trenn = if ($pfad -match '\?') { '&' } else { '?' }
        $r = CN-Abruf "$pfad$($trenn)paginationMode=token&limit=100$(if ($nach) { "&startAfter=$nach" })"
        $feld = @($r.PSObject.Properties | Where-Object { @('data', 'files', 'times', 'members', 'projects') -contains $_.Name } | Select-Object -First 1)
        $items = if ($feld.Count) { @($feld[0].Value) } else { @() }
        foreach ($x in $items) { if ($null -ne $x) { [void]$erg.Add($x) } }
        $nach = if ($r.PSObject.Properties['startAfter']) { [string]$r.startAfter } else { '' }
        if (-not $nach -or $items.Count -lt 100) { break }
    }
    return $erg.ToArray()
}
function CN-Datum($u) {
    if (-not $u) { return $null }
    $v = [double]$u; if ($v -gt 1e11) { $v = $v / 1000 }
    return [DateTimeOffset]::FromUnixTimeSeconds([long]$v).LocalDateTime
}

# Alle Projekte + Status + Kunden-Mails aus Sage, 15 Minuten zwischengespeichert. Bei Fehler bleibt der letzte Stand.
$script:CnCache = $null; $script:CnZeit = [datetime]::MinValue; $script:CnFehler = ''
function Get-Craftnote {
    if ($script:CnCache -and ((Get-Date) - $script:CnZeit).TotalMinutes -lt $CraftnoteCacheMinuten) { return $script:CnCache }
    if (-not $script:CnCache -and ((Get-Date) - $script:CnZeit).TotalMinutes -lt 1) { return $null }   # nach Fehler kurz warten
    if (-not (Umgebung 'CRAFTNOTE_API_KEY')) { $script:CnFehler = 'Craftnote-Schl&uuml;ssel nicht hinterlegt (CRAFTNOTE_API_KEY)'; return $null }
    try {
        $status = [ordered]@{}; foreach ($s in @((CN-Abruf '/company/settings/projectstatus').data)) { $status[[string]$s.id] = [string]$s.name }
        $proj = foreach ($p in (CN-Liste '/projects')) {
            if ($p.projectType -ne 'PROJECT') { continue }
            $pm = @(@($p.billingEmail) + @($p.contacts | ForEach-Object { $_.emails }) | Where-Object { $_ } | ForEach-Object { ([string]$_).Trim().ToLower() } | Where-Object { $_ -notmatch $CraftnoteEigeneMails } | Select-Object -Unique)
            [pscustomobject]@{ Id = [string]$p.id; Name = [string]$p.name; Norm = (Name-Norm $p.name); Plz = ([string]$p.zipcode).Trim(); Strasse = (Name-Norm $p.street); Ort = ([string]$p.city).Trim()
                Mails = $pm; StatusId = [string]$p.statusId; Status = $(if ($p.statusId -and $status.Contains([string]$p.statusId)) { $status[[string]$p.statusId] } else { '' })
                Archiviert = [bool]$p.archived; Erstellt = (CN-Datum $p.creationDate); Geaendert = (CN-Datum $p.lastEditedDate); Start = (CN-Datum $p.startDate)
                Notiz = ([string]$p.noteContent).Trim(); NotizAutor = [string]$p.noteAuthor; NotizZeit = (CN-Datum $p.noteTimestamp) }
        }
        # Indizes fuer die Zuordnung
        $nachMail = @{}; $nachPlz = @{}
        foreach ($p in $proj) {
            if ($p.Archiviert) { continue }
            foreach ($m in $p.Mails) { if (-not $nachMail.ContainsKey($m)) { $nachMail[$m] = New-Object System.Collections.ArrayList }; [void]$nachMail[$m].Add($p) }
            if ($p.Plz) { if (-not $nachPlz.ContainsKey($p.Plz)) { $nachPlz[$p.Plz] = New-Object System.Collections.ArrayList }; [void]$nachPlz[$p.Plz].Add($p) }
        }
        $kMails = @{}; $mailKto = @{}
        foreach ($r in (Run-Query $SQL_KundenMails $null).Rows) {
            $k = [string]$r['Kto']; $m = [string]$r['Mail']
            if (-not $kMails.ContainsKey($k)) { $kMails[$k] = New-Object System.Collections.ArrayList }; [void]$kMails[$k].Add($m)
            if (-not $mailKto.ContainsKey($m)) { $mailKto[$m] = $k }
        }
        $script:CnCache = @{ Projekte = @($proj); Status = $status; NachMail = $nachMail; NachPlz = $nachPlz; KundenMails = $kMails; MailKto = $mailKto }
        $script:CnZeit = Get-Date; $script:CnFehler = ''
    } catch {
        $script:CnFehler = "Craftnote nicht erreichbar: $(HtmlEncode $_.Exception.Message)"
        if (-not $script:CnCache) { $script:CnZeit = Get-Date }
    }
    return $script:CnCache
}

# Statusgruppe fuer Farbe, Reihenfolge und Abgleich (Statusnamen kommen aus Craftnote)
function CN-Gruppe($status) {
    switch -regex ([string]$status) {
        '^$' { return '' }
        'abgeschlossen' { return 'fertig' }
        'Absage|nicht gew' { return 'absage' }
        'Inbetrieb' { return 'inbetrieb' }
        'rderung' { return 'foerderung' }
        'Anfrage|Angebot' { return 'vertrieb' }
        'Service' { return 'service' }
        default { return 'auftrag' }
    }
}
$CnStatusRang = @{ 'auftrag' = 0; 'inbetrieb' = 1; 'foerderung' = 2; 'service' = 3; 'vertrieb' = 4; 'fertig' = 5; 'absage' = 6; '' = 7 }
function CN-Kurz($s) {
    return ([string]$s) -replace 'Auftrag vorhanden und Terminiert', 'Auftrag terminiert' -replace 'Angebot erstellt und versendet', 'Angebot versendet' `
        -replace 'Termin zum Z(.)hlerwechsel vereinbart', 'Z$1hlerwechsel vereinbart' -replace 'EVU Unterlagen eingereicht', 'EVU eingereicht' -replace '^zu weit oder zu klein, ', ''
}

# Craftnote-Projekt zu einem Sage-Kunden. Rueckgabe @{ P; Sicher; Wie; Anzahl } oder $null.
function CN-Finde($cn, $name1, $kto, $plzListe, $strasse) {
    if (-not $cn) { return $null }
    $wahl = { param($k) @($k | Sort-Object { $CnStatusRang[(CN-Gruppe $_.Status)] }, { if ($_.Geaendert) { - $_.Geaendert.Ticks } else { 0 } }) }
    # 1. E-Mail
    $k = New-Object System.Collections.ArrayList
    if ($kto -and $cn.KundenMails.ContainsKey([string]$kto)) {
        foreach ($m in $cn.KundenMails[[string]$kto]) { if ($cn.NachMail.ContainsKey($m)) { foreach ($p in $cn.NachMail[$m]) { if (-not $k.Contains($p)) { [void]$k.Add($p) } } } }
    }
    if ($k.Count) { return @{ P = (& $wahl $k)[0]; Sicher = $true; Wie = 'E-Mail'; Anzahl = $k.Count } }
    # 2. Nachname + PLZ (Firmen: ein Namenswort + PLZ)
    $woerter = @(((Name-Norm $name1) -split '[^a-z]+') | Where-Object { $_.Length -ge 3 -and $NamenStopp -notcontains $_ })
    if (-not $woerter.Count) { return $null }
    $firma = ([string]$name1 -match $FirmenMuster)
    $nach = if ($firma) { $woerter } elseif ([string]$name1 -match ',') { @($woerter[0]) } else { @($woerter[-1]) }
    $vor = if (-not $firma -and $woerter.Count -ge 2) { if ([string]$name1 -match ',') { $woerter[1] } else { $woerter[0] } } else { $null }
    $k = New-Object System.Collections.ArrayList
    foreach ($plz in @($plzListe | Where-Object { $_ } | ForEach-Object { ([string]$_).Trim() } | Select-Object -Unique)) {
        if (-not $cn.NachPlz.ContainsKey($plz)) { continue }
        foreach ($p in $cn.NachPlz[$plz]) { foreach ($n in $nach) { if ($p.Norm -match "\b$n\b" -and -not $k.Contains($p)) { [void]$k.Add($p) } } }
    }
    if ($k.Count -eq 0) { return $null }
    $k = @($k)
    if ($k.Count -gt 1 -and $vor) { $v = @($k | Where-Object { $_.Norm -match "\b$vor\b" }); if ($v.Count) { $k = $v } }
    if ($k.Count -gt 1 -and $strasse) {
        $sn = (Name-Norm $strasse) -replace '[^a-z]', ''
        if ($sn.Length -ge 4) { $an = $sn.Substring(0, [math]::Min(8, $sn.Length)); $s = @($k | Where-Object { ($_.Strasse -replace '[^a-z]', '') -like "$an*" }); if ($s.Count) { $k = $s } }
    }
    return @{ P = (& $wahl $k)[0]; Sicher = ($k.Count -eq 1); Wie = 'Name und PLZ'; Anzahl = $k.Count }
}
function Craftnote-Zuordnen($liste) {
    $cn = Get-Craftnote
    foreach ($au in $liste) { $au.Cn = $(if ($cn) { CN-Finde $cn $au.Name1 $au.Kto $au.Plz $au.StrasseRoh } else { $null }) }
    return $cn
}

# Status als Chip mit Link in die Craftnote-Web-App
function CN-Chip($z) {
    if (-not $z) { return "" }
    $p = $z.P; $g = CN-Gruppe $p.Status
    $u = if (-not $z.Sicher) { " <span class='unsicher' title='$($z.Anzahl) Craftnote-Projekte passen zu Name und PLZ &ndash; bitte pr&uuml;fen'>?</span>" } else { "" }
    return "<a class='cn cn-$g' href='$CraftnoteWeb$($p.Id)' target='_blank' title='Craftnote: $(HtmlEncode $p.Name) &ndash; $(HtmlEncode $p.Status) (zugeordnet &uuml;ber $($z.Wie))'>$(HtmlEncode (CN-Kurz $(if ($p.Status) { $p.Status } else { 'ohne Status' })))</a>$u"
}

# Bereich "Craftnote" auf der Auftragsseite: Projekt, Status, Notiz, Mitglieder, Dateien je Ordner, Zeiten
function CN-DetailHtml($z) {
    if (-not $z) {
        if ($script:CnFehler) { return "<p class='leise'>$($script:CnFehler)</p>" }
        return "<p class='leise'>Kein passendes Craftnote-Projekt gefunden (weder &uuml;ber die E-Mail-Adresse noch &uuml;ber Name und PLZ).</p>"
    }
    $p = $z.P; $d = { param($x) if ($x -is [datetime]) { $x.ToString('dd.MM.yyyy') } else { '' } }
    $zeilen = New-Object System.Collections.ArrayList
    [void]$zeilen.Add("<tr><td>Projekt</td><td><a class='artlink' href='$CraftnoteWeb$($p.Id)' target='_blank'>$(HtmlEncode $p.Name) &#8599;</a> <span class='leise'>zugeordnet &uuml;ber $($z.Wie)$(if ($z.Anzahl -gt 1) { ", $($z.Anzahl) Projekte passen &ndash; bitte pr&uuml;fen" })</span></td></tr>")
    [void]$zeilen.Add("<tr><td>Status</td><td>$(CN-Chip $z)</td></tr>")
    [void]$zeilen.Add("<tr><td>Angelegt / ge&auml;ndert</td><td>$(& $d $p.Erstellt) / $(& $d $p.Geaendert)$(if ($p.Start) { " &middot; Start $(& $d $p.Start)" })</td></tr>")
    if ($p.Notiz) { [void]$zeilen.Add("<tr><td>Notiz</td><td style='white-space:pre-wrap'>$(HtmlEncode $p.Notiz)<div class='leise'>$(HtmlEncode $p.NotizAutor) $(& $d $p.NotizZeit)</div></td></tr>") }
    try {
        $mit = @(CN-Liste "/projects/$($p.Id)/members" | ForEach-Object { ("$($_.name) $($_.lastname)").Trim() } | Where-Object { $_ })
        if ($mit.Count) { [void]$zeilen.Add("<tr><td>Projektmitglieder</td><td>$(HtmlEncode ($mit -join ', '))</td></tr>") }
        $zeiten = @(CN-Liste "/projects/$($p.Id)/times" | Where-Object { $_.state -ne 'deleted' })
        if ($zeiten.Count) {
            $jeP = [ordered]@{}; foreach ($t in $zeiten) { $n = [string]$t.userName; $jeP[$n] = [double]$jeP[$n] + [double]$t.workDuration }
            $summe = 0.0; foreach ($v in $jeP.Values) { $summe += $v }
            $txt = (@($jeP.Keys | Sort-Object { - $jeP[$_] } | ForEach-Object { "$(HtmlEncode $_) $(([math]::Round($jeP[$_] / 3600, 1)).ToString('0.#', [Globalization.CultureInfo]'de-DE')) h" })) -join ' &middot; '
            [void]$zeilen.Add("<tr><td>Erfasste Zeit</td><td><b>$(([math]::Round($summe / 3600, 1)).ToString('0.#', [Globalization.CultureInfo]'de-DE')) h</b> &middot; $txt</td></tr>")
        }
        $dateien = @(CN-Liste "/projects/$($p.Id)/files")
        $ordner = @{}; foreach ($x in $dateien) { if ($x.type -eq 'FOLDER') { $ordner[[string]$x.id] = $x } }
        $oben = { param($fid) $n = 0; while ($fid -and $ordner.ContainsKey($fid) -and $ordner[$fid].folderId -and $n -lt 10) { $fid = [string]$ordner[$fid].folderId; $n++ }; $fid }
        $gruppen = [ordered]@{}
        foreach ($x in ($dateien | Where-Object { $_.type -ne 'FOLDER' } | Sort-Object { - [double]$_.creationTimestamp })) {
            $o = & $oben ([string]$x.folderId); $name = if ($o -and $ordner.ContainsKey($o)) { [string]$ordner[$o].name } else { 'ohne Ordner' }
            if (-not $gruppen.Contains($name)) { $gruppen[$name] = New-Object System.Collections.ArrayList }; [void]$gruppen[$name].Add($x)
        }
        if ($gruppen.Count) {
            $teile = foreach ($name in ($gruppen.Keys | Sort-Object)) {
                $liste = @($gruppen[$name] | ForEach-Object { "<div><a class='artlink' href='/cndatei?id=$($_.id)' target='_blank'>$(HtmlEncode $_.name)</a> <span class='leise'>$(& $d (CN-Datum $_.creationTimestamp))</span></div>" })
                "<details><summary>$(HtmlEncode $name) <span class='leise'>($($gruppen[$name].Count))</span></summary><div style='padding:4px 0 6px 16px'>$($liste -join '')</div></details>"
            }
            [void]$zeilen.Add("<tr><td>Dateien</td><td>$($teile -join '')</td></tr>")
        } else { [void]$zeilen.Add("<tr><td>Dateien</td><td class='leise'>keine</td></tr>") }
    } catch { [void]$zeilen.Add("<tr><td>Details</td><td class='leise'>Craftnote-Details nicht abrufbar: $(HtmlEncode $_.Exception.Message)</td></tr>") }
    return "<table style='max-width:900px'><tbody>$($zeilen -join '')</tbody></table>"
}

# ---------- Abgleich Sage <-> Craftnote ----------
function Build-Abgleich($liste, $cn, $fehler) {
    $head = Page-Head "Solar Dispo - Abgleich Craftnote"
    $heute = (Get-Date).Date
    if ($fehler) { return "$head<div class='fehler'>Fehler beim Laden der Daten:<br><pre>$(HtmlEncode $fehler)</pre></div></body></html>" }
    if (-not $cn) { return "$head<div class='kopf'><h1>Abgleich Craftnote</h1></div><div class='hinweisbox'>$($script:CnFehler)</div></body></html>" }
    $auZelle = { param($au) "<span class='$(if ($au.Art -like '*RES*') { 'art res' } else { 'art' })'>$($au.Art)</span> <a class='artlink' href='/auftrag?bel=$($au.BelID)'>$(HtmlEncode $au.Beleg)</a>" }
    $planZelle = { param($au) if ($au.Naechster) { Termin-Html $au.Naechster $false } elseif ($au.Liefertermin) { "<span class='leise'>Liefertermin</span> $($au.Liefertermin.ToString('dd.MM.yyyy'))" } else { "" } }
    # A: in Sage offen, in Craftnote schon fertig/abgesagt bzw. noch im Vertrieb
    $a = @($liste | Where-Object { $_.Cn -and @('fertig', 'absage', 'vertrieb') -contains (CN-Gruppe $_.Cn.P.Status) })
    $rowsA = foreach ($au in $a) { "<tr><td style='white-space:nowrap'>$(& $auZelle $au)</td><td>$(HtmlEncode $au.Kunde)</td><td style='white-space:nowrap'>$(& $planZelle $au)</td><td style='text-align:center'>$(Stand-Html $au.Stand 'Gel')</td><td style='text-align:center'>$(Stand-Html $au.Stand 'Ber')</td><td>$(CN-Chip $au.Cn)</td></tr>" }
    # B: Craftnote "Auftrag terminiert", aber kein offener Sage-Auftrag dazu
    $zugeordnet = @{}; foreach ($au in $liste) { if ($au.Cn) { $zugeordnet[$au.Cn.P.Id] = 1 } }
    $letzte = @{}; foreach ($r in (Run-Query $SQL_LetzteAuftraege $null).Rows) { $letzte[[string]$r['Kto']] = $r }
    $b = @($cn.Projekte | Where-Object { -not $_.Archiviert -and $_.Status -match 'Auftrag vorhanden' -and -not $zugeordnet.ContainsKey($_.Id) } | Sort-Object { - $(if ($_.Geaendert) { $_.Geaendert.Ticks } else { 0 }) })
    $rowsB = foreach ($p in $b) {
        $kto = $null; foreach ($m in $p.Mails) { if ($cn.MailKto.ContainsKey($m)) { $kto = $cn.MailKto[$m]; break } }
        $sage = if (-not $kto) { "<span class='leise'>kein Sage-Kunde mit dieser E-Mail</span>" }
                elseif (-not $letzte.ContainsKey($kto)) { "<span class='leise'>Kunde vorhanden, kein Auftrag</span>" }
                else { $l = $letzte[$kto]; "letzter Auftrag $(HtmlEncode $l['Beleg']) vom $(([datetime]$l['Belegdatum']).ToString('dd.MM.yyyy'))$(if ([int]$l['Erledigt'] -eq 1) { ' &middot; <b>erledigt</b>' } else { ' &middot; offen' })" }
        "<tr><td><a class='artlink' href='$CraftnoteWeb$($p.Id)' target='_blank'>$(HtmlEncode $p.Name) &#8599;</a></td><td>$(HtmlEncode "$($p.Plz) $($p.Ort)")</td><td>$(if ($p.Geaendert) { $p.Geaendert.ToString('dd.MM.yyyy') })</td><td>$sage</td></tr>"
    }
    # C: offene Sage-Auftraege ohne Craftnote-Projekt
    $c = @($liste | Where-Object { -not $_.Cn })
    $rowsC = foreach ($au in $c) { "<tr><td style='white-space:nowrap'>$(& $auZelle $au)</td><td>$(HtmlEncode $au.Kunde)</td><td>$(HtmlEncode $au.Baustelle)</td><td style='white-space:nowrap'>$(& $planZelle $au)</td></tr>" }
    $unsicher = @($liste | Where-Object { $_.Cn -and -not $_.Cn.Sicher }).Count
    $tab = { param($kopf, $rows, $leer) if (@($rows).Count) { "<table><thead><tr>$kopf</tr></thead><tbody>$($rows -join '')</tbody></table>" } else { "<p class='leise'>$leer</p>" } }
    return @"
$head
<div class='kopf'><h1>Abgleich Craftnote</h1><div class='meta'>$(StandText) &middot; $($liste.Count) offene Auftr&auml;ge, $(@($liste | Where-Object { $_.Cn }).Count) zugeordnet$(if ($unsicher) { ", davon $unsicher unsicher" })</div></div>
$(if ($script:CnFehler) { "<div class='hinweisbox'>$($script:CnFehler)</div>" })
<h2>In Sage offen, in Craftnote abgeschlossen, abgesagt oder noch im Angebot ($($a.Count))</h2>
$(& $tab "<th>Auftrag</th><th>Kunde</th><th>Montage</th><th>Geliefert</th><th>Berechnet</th><th>Craftnote</th>" $rowsA 'Keine Abweichungen.')
<h2>In Craftnote &bdquo;Auftrag terminiert&ldquo;, in Sage kein offener Auftrag ($($b.Count))</h2>
$(& $tab "<th>Craftnote-Projekt</th><th>Ort</th><th>Ge&auml;ndert</th><th>Sage</th>" $rowsB 'Keine Abweichungen.')
<h2>Offene Auftr&auml;ge ohne Craftnote-Projekt ($($c.Count))</h2>
$(& $tab "<th>Auftrag</th><th>Kunde</th><th>Baustelle</th><th>Montage</th>" $rowsC 'Alle offenen Auftr&auml;ge haben ein Craftnote-Projekt.')
</body></html>
"@
}

# ---------- Dienst ----------
if ($MyInvocation.InvocationName -eq '.') { return }   # dot-gesourct (Test): nur Funktionen laden

# Hintergrund-Runspaces fuer den Chat: bekommen alle Funktionen und alle Einstellungen/SQL-Texte dieses Skripts
$ChatPool = $null
$ChatJobs = @{}
try {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    $chatVarNamen = @(Get-Variable -Scope Script | Where-Object { $StartVariablen -notcontains $_.Name -and @('StartVariablen', 'ChatPool', 'ChatJobs') -notcontains $_.Name } | ForEach-Object { $_.Name })
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $skriptAst = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref]$null, [ref]$null)
    foreach ($st in $skriptAst.EndBlock.Statements) {
        if ($st -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { continue }
        $rumpf = $st.Body.Extent.Text; $rumpf = $rumpf.Substring(1, $rumpf.Length - 2)
        if ($st.Parameters) { $rumpf = "param(" + (($st.Parameters | ForEach-Object { $_.Extent.Text }) -join ', ') + ")`n" + $rumpf }
        $iss.Commands.Add((New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry($st.Name, $rumpf)))
    }
    foreach ($n in $chatVarNamen) {
        $iss.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry($n, (Get-Variable -Scope Script -Name $n).Value, '')))
    }
    $ChatPool = [runspacefactory]::CreateRunspacePool(1, 3, $iss, $Host)
    $ChatPool.Open()
} catch {
    Write-Host "Assistent nicht verfuegbar: $($_.Exception.Message)" -ForegroundColor Yellow
    $ChatPool = $null
}

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
        $ausgabe = ""; $contentType = "text/html; charset=utf-8"; $bereitsGesendet = $false
        if ($pfad -eq "/") {
            try { $liste = @(Get-Auftraege); try { $null = Craftnote-Zuordnen $liste } catch { $script:CnFehler = HtmlEncode $_.Exception.Message }; $ausgabe = Build-Auftraege $liste $null }
            catch { $ausgabe = Build-Auftraege @() ($_.Exception.ToString()) }
        } elseif ($pfad -eq "/auftrag") {
            $bel = ([string](Parse-Query $query "bel")) -replace '[^0-9]', ''
            try {
                $au = @(Get-Auftraege) | Where-Object { $_.BelID -eq $bel } | Select-Object -First 1
                if ($au) { try { $null = Craftnote-Zuordnen @($au) } catch { $script:CnFehler = HtmlEncode $_.Exception.Message } }
                $ausgabe = Build-AuftragDetail $au $(if ($au) { Get-AuftragPositionen $au }) $null
            } catch { $ausgabe = Build-AuftragDetail $null $null ($_.Exception.ToString()) }
        } elseif ($pfad -eq "/material") {
            $nur = Parse-Query $query "nur"; if (@('fehlt', 'alle') -notcontains $nur) { $nur = 'fehlt' }
            $tg = Parse-Query $query "tage"; $tage = if (@('14', '28', '0') -contains $tg) { [int]$tg } else { 28 }
            try { $ausgabe = Build-Material (Get-Auftraege) $nur $tage $null }
            catch { $ausgabe = Build-Material @() $nur $tage ($_.Exception.ToString()) }
        } elseif ($pfad -eq "/angebote") {
            $mn = Parse-Query $query "monate"; $monate = if (@('3', '6', '12') -contains $mn) { [int]$mn } else { 3 }
            try { $ausgabe = Build-Angebote $monate $null }
            catch { $ausgabe = Build-Angebote $monate ($_.Exception.ToString()) }
        } elseif ($pfad -match '^/static/([a-z0-9_-]+\.(woff2|svg))$') {
            $sDatei = Join-Path (Join-Path $PSScriptRoot 'static') $matches[1]
            if (Test-Path -LiteralPath $sDatei -PathType Leaf) {
                $bytes = [System.IO.File]::ReadAllBytes($sDatei)
                $response.ContentType = if ($sDatei -like '*.svg') { 'image/svg+xml' } else { 'font/woff2' }
                $response.AddHeader("Cache-Control", "public, max-age=86400")
                $response.ContentLength64 = $bytes.Length
                $response.OutputStream.Write($bytes, 0, $bytes.Length); $response.OutputStream.Close()
                $bereitsGesendet = $true
            } else { $response.StatusCode = 404; $ausgabe = "Nicht gefunden" }
        } elseif ($pfad -eq "/dispo") {
            $art = ([string](Parse-Query $query "art")).Trim() -replace "[';]", ""
            $ausgabe = Dispo-Seite $art
        } elseif ($pfad -eq "/api/artikelsuche") {
            # Autovervollstaendigung: JSON
            $contentType = "application/json; charset=utf-8"
            $q = ([string](Parse-Query $query "q")) -replace "[';]", ""
            try {
                $sql = "SELECT TOP 50 Artikelnummer, Bezeichnung1 AS Bezeichnung FROM KHKArtikel " +
                       "WHERE Mandant = $Mandant AND Aktiv = -1 AND (Artikelnummer LIKE @suche OR Bezeichnung1 LIKE @suche) " +
                       "ORDER BY CASE WHEN Artikelnummer LIKE @anfang THEN 0 WHEN Bezeichnung1 LIKE @anfang THEN 1 ELSE 2 END, Artikelnummer"
                $tab = Run-Query $sql @{ "@suche" = "%$q%"; "@anfang" = "$q%" }
                $ausgabe = "[" + ((@($tab.Rows) | ForEach-Object { "{""nr"":""$(JsEncode $_['Artikelnummer'])"",""bez"":""$(JsEncode $_['Bezeichnung'])""}" }) -join ",") + "]"
            } catch { $ausgabe = "[]" }
        } elseif ($pfad -eq "/api/lagerplatzsuche") {
            # Autovervollstaendigung Lagerplatz (ganze Lager + einzelne Plaetze): JSON
            $contentType = "application/json; charset=utf-8"
            $qLpClean = ([string](Parse-Query $query "q")) -replace "[';]", ""
            # ganze Lager + alle Plaetze (auch ohne Bestand), Treffer auch auf "A 1/2"
            try {
                $teile = New-Object System.Collections.ArrayList
                foreach ($r in (Run-Query "SELECT TOP 5 Lagerkennung FROM KHKLagerorte WHERE Mandant = $Mandant AND Lagerkennung LIKE @s ORDER BY Lagerkennung" @{ "@s" = "%$qLpClean%" }).Rows) {
                    [void]$teile.Add("{""v"":""$(JsEncode $r['Lagerkennung'])"",""t"":""$(JsEncode $r['Lagerkennung'])"",""g"":1}")
                }
                $sqlPl = @"
SELECT TOP 20 Lagerkennung AS K, ISNULL(Dimension1, 0) AS D1, ISNULL(Dimension2, 0) AS D2, ISNULL(Dimension3, 0) AS D3
FROM KHKLagerplaetze
WHERE Mandant = $Mandant
  AND (Lagerkennung + ' ' + CAST(ISNULL(Dimension1, 0) AS varchar) + '/' + CAST(ISNULL(Dimension2, 0) AS varchar) + '/' + CAST(ISNULL(Dimension3, 0) AS varchar)) LIKE @s
  AND NOT (ISNULL(Dimension1, 0) = 0 AND ISNULL(Dimension2, 0) = 0 AND ISNULL(Dimension3, 0) = 0)
ORDER BY Lagerkennung, ISNULL(Dimension1, 0), ISNULL(Dimension2, 0), ISNULL(Dimension3, 0)
"@
                foreach ($r in (Run-Query $sqlPl @{ "@s" = "%$qLpClean%" }).Rows) {
                    [void]$teile.Add("{""v"":""$(JsEncode (Platz-Wert $r['K'] $r['D1'] $r['D2'] $r['D3']))"",""t"":""$(JsEncode (Fmt-Platz $r['K'] $r['D1'] $r['D2'] $r['D3']))""}")
                }
                $ausgabe = "[" + ($teile -join ",") + "]"
            } catch { $ausgabe = "[]" }
        } elseif ($pfad -eq "/bewegungen") {
            # --- Lagerbewegungsprotokoll: wer hat wann was wohin gebucht ---
            # Werte kommen aus einem GET-Formular: "+" steht fuer Leerzeichen
            $fw = { param($n) if ($query -match "(^|[?&])$n=([^&]*)") { ([System.Uri]::UnescapeDataString($matches[2].Replace('+', ' '))).Trim() } else { "" } }
            $heute = (Get-Date).Date
            $f = @{ von = $heute.AddDays(-6); bis = $heute; typ = (& $fw 'typ'); wer = (& $fw 'wer'); art = (& $fw 'art'); platz = (& $fw 'platz') }
            $d = [datetime]::MinValue
            if ([datetime]::TryParseExact((& $fw 'von'), 'yyyy-MM-dd', $null, 'None', [ref]$d)) { $f.von = $d }
            if ([datetime]::TryParseExact((& $fw 'bis'), 'yyyy-MM-dd', $null, 'None', [ref]$d)) { $f.bis = $d }
            if ($f.bis -lt $f.von) { $t = $f.von; $f.von = $f.bis; $f.bis = $t }
            if (@('um','zu','ab') -notcontains $f.typ) { $f.typ = '' }
            $max = 2000
            $namen = @{}
            try {
                # Namensliste aendert sich selten - 15 Minuten zwischenspeichern
                if ($null -eq $script:BewNamenCache -or ((Get-Date) - $script:BewNamenZeit).TotalMinutes -gt 15) {
                    $script:BewNamenCache = Get-BenutzerNamen (Run-Query $SQL_BewegungenBenutzer $null) (Run-Query $SQL_MitarbeiterNamen $null)
                    $script:BewNamenZeit = Get-Date
                }
                $namen = $script:BewNamenCache
                $p = @{ "@max" = $max; "@von" = $f.von; "@bis" = $f.bis.AddDays(1) }
                $filter = @()
                if ($f.typ) {
                    $codes = @($BewegungsArten.Keys | Where-Object { $BewegungsArten[$_][0] -eq $f.typ } | ForEach-Object { "'$_'" }) -join ','
                    $filter += "AND b.Bewegungsart IN ($codes)"
                }
                if ($f.art)   { $filter += "AND (b.Artikelnummer LIKE @art + '%' OR a.Bezeichnung1 LIKE '%' + @art + '%')"; $p["@art"] = $f.art }
                if ($f.platz -like '*;*') {
                    # genau ein Platz: Lagerkennung + Dimension 1-3 (Herkunft oder Ziel)
                    $tp = $f.platz.Split(';'); $p["@pk"] = $tp[0]
                    for ($i = 1; $i -le 3; $i++) { $n = 0; if ($tp.Count -gt $i) { [void][int]::TryParse($tp[$i], [ref]$n) }; $p["@pd$i"] = $n }
                    $filter += "AND ((hp.Lagerkennung = @pk AND ISNULL(hp.Dimension1, 0) = @pd1 AND ISNULL(hp.Dimension2, 0) = @pd2 AND ISNULL(hp.Dimension3, 0) = @pd3) OR (zp.Lagerkennung = @pk AND ISNULL(zp.Dimension1, 0) = @pd1 AND ISNULL(zp.Dimension2, 0) = @pd2 AND ISNULL(zp.Dimension3, 0) = @pd3))"
                } elseif ($f.platz) { $filter += "AND (hp.Lagerkennung LIKE @platz + '%' OR zp.Lagerkennung LIKE @platz + '%')"; $p["@platz"] = $f.platz }
                if ($f.wer) {
                    # Eine Person kann mehrere Sage-Benutzer haben (Kuerzel und Scanner)
                    $codes = @($namen.Keys | Where-Object { $namen[$_].Name -eq $f.wer })
                    if ($codes.Count -eq 0) { $codes = @($f.wer) }
                    $platzhalter = for ($i = 0; $i -lt $codes.Count; $i++) { $p["@w$i"] = $codes[$i]; "@w$i" }
                    $filter += "AND j.Benutzer IN ($($platzhalter -join ','))"
                }
                $sql = $SQL_Bewegungen -replace '\{FILTER\}', ($filter -join "`n  ")
                $tab = Run-Query $sql $p
                $ausgabe = Build-Bewegungen $tab $f $namen $max $null
            } catch {
                $ausgabe = Build-Bewegungen $null $f $namen $max ($_.Exception.ToString())
            }
        } elseif ($pfad -eq "/chat-an" -or $pfad -eq "/chat-aus") {
            # --- Testfreigabe des Assistenten fuer diesen Browser (Cookie, 1 Jahr) ---
            $an = ($pfad -eq "/chat-an")
            $ck = New-Object System.Net.Cookie('solarchat', $(if ($an) { '1' } else { '0' }), '/')
            $ck.Expires = (Get-Date).AddDays($(if ($an) { 365 } else { -1 }))
            $response.Cookies.Add($ck)
            $hinweis = if ($null -eq $ChatPool) { "Der Assistent konnte beim Start nicht geladen werden (siehe Dienst-Log)." }
                       elseif (-not (Get-ChatApiKey)) { "Achtung: Auf dem Server ist noch kein API-Schl&uuml;ssel (ANTHROPIC_API_KEY) hinterlegt &ndash; das Chatfenster erscheint erst danach." }
                       else { "" }
            $ausgabe = "$(Page-Head 'Assistent')<div class='kopf'><h1>Assistent $(if ($an) { 'freigeschaltet' } else { 'ausgeschaltet' })</h1></div>" +
                       "<p>$(if ($an) { 'In diesem Browser erscheint jetzt unten rechts der Knopf &bdquo;Frage stellen&ldquo;. Ausschalten mit <a href=''/chat-aus''>/chat-aus</a>.' } else { 'Das Chatfenster wird in diesem Browser nicht mehr angezeigt.' })</p>" +
                       "$(if ($hinweis) { "<p class='warn'>$hinweis</p>" })<p><a class='btn' href='/'>Zur Solar Dispo</a></p></body></html>"

        } elseif ($pfad -eq "/api/chat") {
            # --- Assistent: POST startet eine Anfrage im Hintergrund, GET ?id= liefert den Stand ---
            $contentType = "application/json; charset=utf-8"
            # alte, nicht abgeholte Anfragen aufraeumen
            foreach ($altId in @($ChatJobs.Keys)) {
                if (((Get-Date) - $ChatJobs[$altId].Start).TotalMinutes -gt 10) { try { $ChatJobs[$altId].PS.Dispose() } catch {}; $ChatJobs.Remove($altId) }
            }
            if (-not (Test-ChatErlaubt)) {
                $ausgabe = '{"fehler":"Der Assistent ist f\u00fcr diesen Browser nicht freigeschaltet."}'
            } elseif ($request.HttpMethod -eq 'POST') {
                try {
                    $sr = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
                    $daten = $sr.ReadToEnd() | ConvertFrom-Json; $sr.Close()
                    # nur Text-Nachrichten user/assistant, beginnend mit user, abwechselnd
                    $verlauf = New-Object System.Collections.ArrayList
                    foreach ($m in @($daten.messages)) {
                        $rolle = [string]$m.role; $inhalt = [string]$m.content
                        if (@('user', 'assistant') -notcontains $rolle -or [string]::IsNullOrWhiteSpace($inhalt)) { continue }
                        if ($verlauf.Count -eq 0 -and $rolle -ne 'user') { continue }
                        if ($verlauf.Count -gt 0 -and $verlauf[$verlauf.Count - 1].role -eq $rolle) { $verlauf.RemoveAt($verlauf.Count - 1) }
                        if ($inhalt.Length -gt 6000) { $inhalt = $inhalt.Substring(0, 6000) }
                        [void]$verlauf.Add(@{ role = $rolle; content = $inhalt })
                    }
                    if ($verlauf.Count -eq 0 -or $verlauf[$verlauf.Count - 1].role -ne 'user') { throw "keine Frage" }
                    $id = [guid]::NewGuid().ToString('N')
                    $status = [hashtable]::Synchronized(@{ Schritt = 'Denke nach &hellip;' })
                    $ps = [powershell]::Create()
                    $ps.RunspacePool = $ChatPool
                    [void]$ps.AddScript('param($v, $s) Invoke-ChatAnfrage $v $s').AddArgument($verlauf).AddArgument($status)
                    $ChatJobs[$id] = @{ PS = $ps; Handle = $ps.BeginInvoke(); Status = $status; Start = Get-Date; Frage = $verlauf[$verlauf.Count - 1].content }
                    $ausgabe = '{"id":"' + $id + '"}'
                } catch {
                    $ausgabe = '{"fehler":' + (ConvertTo-JsonText ([System.Net.WebUtility]::HtmlEncode("Anfrage konnte nicht gestartet werden: $($_.Exception.Message)"))) + '}'
                }
            } else {
                $id = Parse-Query $query "id"
                if (-not $id -or -not $ChatJobs.ContainsKey($id)) {
                    $ausgabe = '{"status":"fehler","fehler":"Anfrage nicht gefunden (Server neu gestartet?) \u2013 bitte erneut senden."}'
                } else {
                    $job = $ChatJobs[$id]
                    if (-not $job.Handle.IsCompleted) {
                        $ausgabe = '{"status":"laeuft","schritt":' + (ConvertTo-JsonText $job.Status.Schritt) + '}'
                    } else {
                        $erg = $null
                        try { $erg = @($job.PS.EndInvoke($job.Handle))[0] } catch { $erg = @{ Fehler = [System.Net.WebUtility]::HtmlEncode($_.Exception.Message); Werkzeuge = @() } }
                        if ($null -eq $erg) { $erg = @{ Fehler = 'Keine Antwort erhalten.'; Werkzeuge = @() } }
                        $dauer = [int]((Get-Date) - $job.Start).TotalSeconds
                        try { $job.PS.Dispose() } catch {}
                        $ChatJobs.Remove($id)
                        # Protokoll (Frage, Werkzeuge, Tokens - keine Antworten)
                        try {
                            $frageLog = [string]$job.Frage; if ($frageLog.Length -gt 1000) { $frageLog = $frageLog.Substring(0, 1000) }
                            $eintrag = [ordered]@{ zeit = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); id = $id; ip = [string]$request.RemoteEndPoint.Address; dauer = $dauer
                                                   ein = [int]$erg.TokensEin; cacheLesen = [int]$erg.CacheLesen; cacheSchreiben = [int]$erg.CacheSchreiben; aus = [int]$erg.TokensAus
                                                   werkzeuge = @($erg.Werkzeuge); frage = $frageLog; fehler = [string]$erg.Fehler }
                            [System.IO.File]::AppendAllText($ChatProtokollPfad, (ConvertTo-JsonText $eintrag) + "`r`n", [System.Text.Encoding]::UTF8)
                        } catch {}
                        if ($erg.Fehler) { $ausgabe = '{"status":"fehler","fehler":' + (ConvertTo-JsonText $erg.Fehler) + '}' }
                        else { $ausgabe = '{"status":"fertig","antwort":' + (ConvertTo-JsonText $erg.Antwort) + '}' }
                    }
                }
            }

        } elseif ($pfad -eq "/api/chat-bewertung") {
            # --- Daumen hoch/runter zu einer Antwort ---
            $contentType = "application/json; charset=utf-8"
            $ausgabe = '{"ok":false}'
            if ($request.HttpMethod -eq 'POST') {
                try {
                    $sr = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
                    $b = $sr.ReadToEnd() | ConvertFrom-Json; $sr.Close()
                    $kurz = { param($s, $n) $s = [string]$s; if ($s.Length -gt $n) { $s.Substring(0, $n) } else { $s } }
                    if (@('gut', 'schlecht') -contains [string]$b.bewertung) {
                        $eintrag = [ordered]@{ zeit = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                                               id = $(if ([string]$b.id -match '^[0-9a-f]{32}$') { [string]$b.id } else { '' })
                                               ip = [string]$request.RemoteEndPoint.Address; bewertung = [string]$b.bewertung
                                               frage = (& $kurz $b.frage 2000); antwort = (& $kurz $b.antwort 8000); kommentar = (& $kurz $b.kommentar 2000) }
                        [System.IO.File]::AppendAllText($ChatBewertungPfad, (ConvertTo-JsonText $eintrag) + "`r`n", [System.Text.Encoding]::UTF8)
                        $ausgabe = '{"ok":true}'
                    }
                } catch {}
            }

        } elseif ($pfad -eq "/einkauf") {
            $nur = Parse-Query $query "nur"; if (@('ueberfaellig', 'spaet') -notcontains $nur) { $nur = '' }
            $art = ([string](Parse-Query $query "art")).Trim() -replace "[';]", ""
            try { $null = Get-Auftraege; $ausgabe = Build-Einkauf $script:EkPositionen $nur $art $null }
            catch { $ausgabe = Build-Einkauf @() $nur $art ($_.Exception.ToString()) }
        } elseif ($pfad -eq "/ek") {
            $bel = ([string](Parse-Query $query "bel")) -replace '[^0-9]', ''
            try {
                if (-not $bel) { throw "Keine Bestellung angegeben." }
                $null = Get-Auftraege
                $ausgabe = Build-EkDetail (Run-Query $SQL_EkBeleg @{ "@bel" = [int]$bel }) $script:EkPositionen $null
            } catch { $ausgabe = Build-EkDetail $null @() ($_.Exception.ToString()) }
        } elseif ($pfad -eq "/abgleich") {
            try { $liste = @(Get-Auftraege); $cn = Craftnote-Zuordnen $liste; $ausgabe = Build-Abgleich $liste $cn $null }
            catch { $ausgabe = Build-Abgleich @() $null ($_.Exception.ToString()) }
        } elseif ($pfad -eq "/cndatei") {
            $fid = [string](Parse-Query $query "id")
            $ziel = $null
            if ($fid -match '^[0-9a-fA-F-]{36}$') {
                try {
                    $req = [System.Net.HttpWebRequest]::Create("$CraftnoteBasis/files/$fid`?download=true"); $req.AllowAutoRedirect = $false
                    $req.Headers.Add('X-CN-API-KEY', ([string](Umgebung 'CRAFTNOTE_API_KEY')).Trim())
                    try { $res = $req.GetResponse() } catch [System.Net.WebException] { $res = $_.Exception.Response }
                    if ($res) { if ([int]$res.StatusCode -ge 300 -and [int]$res.StatusCode -lt 400) { $ziel = $res.Headers['Location'] }; $res.Close() }
                } catch {}
            }
            if ($ziel -and $ziel -match '^https://') { $response.Redirect($ziel); $response.Close(); $bereitsGesendet = $true }
            else { $response.StatusCode = 404; $ausgabe = "Datei nicht gefunden" }
        } elseif ($pfad -eq "/datei") {
            # Datei aus der Sage-Sammelmappe ausliefern - nur Pfade, die in KHKSammelmappen stehen
            $dId = 0; [void][int]::TryParse(((Parse-Query $query "id") -replace '[^0-9]', ''), [ref]$dId)
            $dPfad = ""
            try { $dt = Run-Query "SELECT Quelle FROM KHKSammelmappen WHERE Mandant = $Mandant AND ID = @id" @{ "@id" = $dId }; if ($dt.Rows.Count) { $dPfad = ([string]$dt.Rows[0]["Quelle"]).Trim() } } catch {}
            $contentType = "text/plain; charset=utf-8"
            if (-not $dPfad -or -not (Test-Path -LiteralPath $dPfad -PathType Leaf)) {
                $response.StatusCode = 404; $ausgabe = "Datei nicht gefunden: $dPfad"
            } else {
                $dInfo = Get-Item -LiteralPath $dPfad
                if ($dInfo.Length -gt $DokuMaxMB * 1MB) {
                    $response.StatusCode = 413; $ausgabe = "Datei ist zu gross fuer die Anzeige im Browser ($([math]::Round($dInfo.Length / 1MB)) MB). Pfad: $dPfad"
                } else {
                    $dTyp = Get-DateiTyp $dPfad
                    $response.ContentType = $dTyp.Mime
                    $response.AddHeader("Content-Disposition", "$(if ($dTyp.Inline) { 'inline' } else { 'attachment' }); filename*=UTF-8''$([System.Uri]::EscapeDataString($dInfo.Name))")
                    $response.AddHeader("Cache-Control", "private, max-age=600")
                    $response.ContentLength64 = $dInfo.Length
                    $fs = [System.IO.File]::OpenRead($dPfad)
                    try { $fs.CopyTo($response.OutputStream) } finally { $fs.Close(); $response.OutputStream.Close() }
                    $bereitsGesendet = $true
                }
            }
        } else {
            $response.StatusCode = 404; $ausgabe = "Nicht gefunden"
        }
        if ($bereitsGesendet) { continue }
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($ausgabe)
        $response.ContentType = $contentType
        $response.ContentLength64 = $buffer.Length
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
        $response.OutputStream.Close()
    } catch {
        Write-Host "Fehler bei Anfrage: $($_.Exception.Message)" -ForegroundColor Red
    }
}
