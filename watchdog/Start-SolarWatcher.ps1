<#
    Start-SolarWatcher.ps1 (Kopie von Start-OezerlisteWatcher.ps1)

    Startet SolarDispo.ps1 und ueberwacht die Datei auf Aenderungen.
    Bei einer Aenderung wird die Syntax geprueft und nur bei gueltigem
    Skript neu gestartet - so ersetzt ein fehlerhafter Zwischenstand
    (z.B. waehrend des Speicherns) nie die laufende, funktionierende Version.
    Stuerzt der Dienst ab (oder laesst sich z.B. wegen eines noch nicht
    verfuegbaren Netzlaufwerks kurz nach dem Systemstart nicht starten),
    wird es im naechsten Pruefintervall automatisch erneut versucht.

    Protokoll: watchdog.log (Watchdog) und dienst.log / dienst-fehler.log
    (Ausgabe von Oezerliste.ps1) im selben Ordner wie dieses Skript.

    Zusaetzlich startet der Watchdog stuendlich timecard_sync.ps1 (Abwesenheiten aus timeCard fuer den
    Kalender) als eigenen Prozess unter demselben Konto - ein haengender Abgleich wird nach 10 Minuten
    beendet und blockiert den Dienst nie. Protokoll: timecard-sync.log.

    Gedacht zum Einrichten als geplante Aufgabe auf lorenzserver5
    (Trigger: "Beim Systemstart", Konto mit Zugriff auf den Netzwerk-Share).
#>
param(
    # UNC-Pfad statt Q:\ - Laufwerksbuchstaben sind nur in einer angemeldeten
    # Benutzersitzung gemappt, eine geplante Aufgabe "ohne angemeldeten Benutzer" sieht sie nicht.
    [string]$ScriptPfad = "\\Lorenzserver2\Daten intern\LORENZ\LORENZ Produktion\30_Produktionstool Sage\Solar Dispo Liste\SolarDispo.ps1",
    [int]$PruefIntervallSek = 5,
    [int]$DebounceSek = 3,
    [int]$SyncIntervallMin = 60
)

$ErrorActionPreference = "Stop"
$proc = $null
$letzterStand = $null
$LogDatei     = Join-Path $PSScriptRoot "watchdog.log"
$DienstLog    = Join-Path $PSScriptRoot "dienst.log"
$DienstFehler = Join-Path $PSScriptRoot "dienst-fehler.log"
$SyncSkript   = Join-Path (Split-Path $ScriptPfad -Parent) "timecard_sync.ps1"
$SyncLog      = Join-Path $PSScriptRoot "timecard-sync.log"
$syncProc = $null; $naechsterSync = Get-Date   # erster Abgleich direkt nach dem Start

function Log {
    param([string]$Text)
    $zeile = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  [$env:COMPUTERNAME/$env:USERNAME PID $PID]  $Text"
    Write-Host $zeile
    try {
        if ((Test-Path -LiteralPath $LogDatei) -and (Get-Item -LiteralPath $LogDatei).Length -gt 1MB) {
            Move-Item -LiteralPath $LogDatei -Destination "$LogDatei.alt" -Force
        }
        Add-Content -LiteralPath $LogDatei -Value $zeile -Encoding UTF8
    } catch { }
}

function Test-PsSyntax {
    param([string]$Pfad)
    $inhalt = Get-Content -Raw -LiteralPath $Pfad
    $fehler = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($inhalt, [ref]$null, [ref]$fehler)
    return ($fehler.Count -eq 0)
}

# Beendet laufende Oezerliste.ps1-Prozesse, die nicht vom aktuellen Watchdog stammen.
# Hintergrund: Wird der Watchdog beendet (Aufgabe beendet, Zeitlimit der Aufgabenplanung,
# Neustart der Aufgabe), laeuft der von ihm gestartete Dienst weiter und belegt Port 8080.
# Ein neuer Watchdog kann dann nie starten und Aenderungen werden nicht mehr uebernommen.
function Stop-VerwaisteDienste {
    $eigene = if ($proc -and -not $proc.HasExited) { $proc.Id } else { -1 }
    $kandidaten = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" |
        Where-Object { $_.CommandLine -like ('*\' + [IO.Path]::GetFileName($ScriptPfad) + '*') -and $_.ProcessId -ne $eigene }
    foreach ($k in $kandidaten) {
        try {
            Stop-Process -Id $k.ProcessId -Force
            Log "Verwaisten Dienst beendet: PID $($k.ProcessId), gestartet $($k.CreationDate)"
        } catch {
            Log "Verwaisten Dienst NICHT beendbar (PID $($k.ProcessId), gestartet $($k.CreationDate)): $($_.Exception.Message)"
        }
    }
}

function Start-Dienst {
    Stop-VerwaisteDienste
    Log "Starte Oezerliste-Dienst ..."
    $script:proc = Start-Process powershell.exe `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPfad`"" `
        -PassThru -NoNewWindow `
        -RedirectStandardOutput $DienstLog -RedirectStandardError $DienstFehler
    Log "Dienst gestartet, PID $($proc.Id)"
}

function Stop-Dienst {
    if ($proc -and -not $proc.HasExited) {
        Log "Stoppe Oezerliste-Dienst (PID $($proc.Id)) ..."
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        try { $proc.WaitForExit(10000) | Out-Null } catch { }
    }
}

# Zustand beim Start protokollieren - der Watchdog startet bei jedem Systemstart, so ist
# nachvollziehbar, was nach einem Neustart von lorenzserver5 laeuft (alte Prozesse, fremde
# Autostarts, Zeitlimit der Aufgabe ...).
function Log-Startdiagnose {
    try { Log "Diagnose: letzter Systemstart $((Get-CimInstance Win32_OperatingSystem).LastBootUpTime)" } catch { }
    try {
        foreach ($t in (Get-ScheduledTask -ErrorAction Stop | Where-Object { $_.TaskName -match 'Solar|Oezer|Watchdog' })) {
            Log ("Diagnose: Aufgabe '{0}' Status={1} Konto={2} Anmeldetyp={3} Zeitlimit={4} Mehrfach={5} Trigger={6}" -f `
                $t.TaskName, $t.State, $t.Principal.UserId, $t.Principal.LogonType, $t.Settings.ExecutionTimeLimit,
                $t.Settings.MultipleInstances, (($t.Triggers | ForEach-Object { $_.CimClass.CimClassName -replace 'MSFT_Task|Trigger', '' }) -join ','))
        }
    } catch { Log "Diagnose: Aufgaben nicht lesbar: $($_.Exception.Message)" }
    try {
        foreach ($p in (Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'")) {
            if ($p.ProcessId -eq $PID) { continue }
            Log "Diagnose: Prozess PID $($p.ProcessId) gestartet $($p.CreationDate): $($p.CommandLine)"
        }
    } catch { }
    try {
        $q = (netsh http show servicestate view=requestq) -join "`n"
        # nur an nicht eingerueckten Warteschlangen-Koepfen trennen (eingerueckt = Unterabschnitt URL-Gruppe)
        foreach ($block in ($q -split "(?m)(?=^Anforderungswarteschlangenname|^Request queue name)")) {
            if ($block -match ':8080') {
                $pids = [regex]::Matches($block, '(?m)^\s*(?:ID|Prozess-ID|Process ID)\s*:\s*(\d+)') | ForEach-Object { $_.Groups[1].Value }
                Log "Diagnose: Port 8080 wird bedient von PID(s): $($pids -join ', ')"
            }
        }
    } catch { }
}

# Bei schnellem Ende des Dienstes (z.B. Port belegt) die Ausgabe ins Watchdog-Log uebernehmen.
function Log-DienstAusgabe {
    foreach ($f in @($DienstFehler, $DienstLog)) {
        try {
            $zeilen = Get-Content -LiteralPath $f -Tail 8 -ErrorAction Stop | Where-Object { $_.Trim() }
            if ($zeilen) { Log "Ausgabe $([IO.Path]::GetFileName($f)): $($zeilen -join ' | ')" }
        } catch { }
    }
}

# Abwesenheiten aus timeCard: timecard_sync.ps1 als eigener Prozess, hoechstens einer gleichzeitig
function Pruefe-Sync {
    if ($syncProc) {
        if (-not $syncProc.HasExited) {
            if (((Get-Date) - $syncProc.StartTime).TotalMinutes -gt 10) {
                Stop-Process -Id $syncProc.Id -Force -ErrorAction SilentlyContinue
                Log "timeCard-Abgleich nach 10 Minuten abgebrochen (PID $($syncProc.Id))."
                $script:syncProc = $null
            }
            return
        }
        $ende = try { (Get-Content -LiteralPath $SyncLog -Tail 3 -ErrorAction Stop | Where-Object { $_.Trim() }) -join ' | ' } catch { '' }
        if ($null -ne $syncProc.ExitCode -and $syncProc.ExitCode -ne 0) { Log "timeCard-Abgleich fehlgeschlagen (Exitcode $($syncProc.ExitCode)): $ende" }
        $script:syncProc = $null
    }
    if ((Get-Date) -lt $naechsterSync -or -not (Test-Path -LiteralPath $SyncSkript)) { return }
    $script:naechsterSync = (Get-Date).AddMinutes($SyncIntervallMin)
    try {
        $script:syncProc = Start-Process powershell.exe `
            -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$SyncSkript`"" `
            -PassThru -NoNewWindow -RedirectStandardOutput $SyncLog -RedirectStandardError "$SyncLog.fehler"
        $null = $script:syncProc.Handle   # Handle sofort holen, sonst liefert ExitCode spaeter nichts
    } catch { Log "timeCard-Abgleich nicht startbar: $($_.Exception.Message)" }
}
Log "Watchdog gestartet. Ueberwache $ScriptPfad"
Log-Startdiagnose

while ($true) {
    try {
        # Kein Dienst aktiv (Erststart, Absturz, oder vorheriger Startversuch
        # gescheitert z.B. weil das Netzlaufwerk noch nicht bereit war) -> versuchen zu starten
        if (-not $proc -or $proc.HasExited) {
            if ($proc) {
                Log "Dienst nicht mehr aktiv (PID $($proc.Id), Exitcode $($proc.ExitCode)), starte neu."
                Log-DienstAusgabe
                Log-Startdiagnose
                Start-Sleep -Seconds 25   # bei dauerhaftem Fehler (Port belegt) das Log nicht fluten
            }
            if (Test-PsSyntax -Pfad $ScriptPfad) {
                $letzterStand = (Get-Item -LiteralPath $ScriptPfad).LastWriteTimeUtc
                Start-Dienst
            } else {
                Log "ACHTUNG: Skript hat einen Syntaxfehler - warte auf eine gueltige Version."
            }
        } else {
            $aktuellerStand = (Get-Item -LiteralPath $ScriptPfad).LastWriteTimeUtc
            if ($aktuellerStand -ne $letzterStand) {
                # Aenderung erkannt: kurz warten, ob noch weiter gespeichert wird (Debounce)
                Start-Sleep -Seconds $DebounceSek
                $aktuellerStand2 = (Get-Item -LiteralPath $ScriptPfad).LastWriteTimeUtc
                if ($aktuellerStand2 -eq $aktuellerStand) {
                    $letzterStand = $aktuellerStand2
                    if (Test-PsSyntax -Pfad $ScriptPfad) {
                        Log "Aenderung erkannt (Dateistand $($aktuellerStand2.ToLocalTime())), starte Dienst neu."
                        Stop-Dienst
                        Start-Dienst
                    } else {
                        Log "NEUE VERSION HAT SYNTAXFEHLER - Neustart uebersprungen, alte Version laeuft weiter!"
                    }
                }
                # sonst: wird noch gespeichert, naechste Runde erneut pruefen
            }
        }
    }
    catch {
        Log "Watchdog-Fehler, naechster Versuch in $PruefIntervallSek s: $($_.Exception.Message)"
    }
    try { Pruefe-Sync } catch { Log "timeCard-Abgleich: $($_.Exception.Message)" }

    Start-Sleep -Seconds $PruefIntervallSek
}
