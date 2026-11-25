# =====================================================================
#  CoE Cockpit – .sppkg bauen
#  Aufruf: Doppelklick auf build.cmd (oder dieses Skript in PowerShell starten)
# =====================================================================

# ---------------- Parameter ----------------
$ProjectPath      = $PSScriptRoot                     # Ordner mit package.json (Standard: Ordner dieses Skripts)
$RequiredNodeMajor = 22                               # SPFx 1.21.1 benötigt Node.js 22
$SolutionFileName = "coe-cockpit.sppkg"               # Name der erzeugten Datei
$OutputFolder     = "sharepoint\solution"             # Ausgabeordner relativ zum Projekt
$CleanInstall     = $false                            # $true = node_modules vorher löschen
$OpenFolderAtEnd  = $true                             # Explorer mit der .sppkg öffnen
# -------------------------------------------

$ErrorActionPreference = "Stop"

function Write-Step($text) { Write-Host ""; Write-Host "==> $text" -ForegroundColor Cyan }
function Fail($text) {
    Write-Host ""
    Write-Host "FEHLER: $text" -ForegroundColor Red
    Write-Host ""
    exit 1
}

Set-Location $ProjectPath
Write-Host "Projektordner: $ProjectPath"

# 1) Projektordner prüfen
if (-not (Test-Path (Join-Path $ProjectPath "package.json"))) {
    Fail "Im Ordner '$ProjectPath' gibt es keine package.json. Das Skript muss im Projektordner (mit src, config, package.json) liegen."
}

# 2) Node.js prüfen
Write-Step "Node.js prüfen"
$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) {
    Fail "Node.js ist nicht installiert. Bitte Node.js $RequiredNodeMajor (LTS) von https://nodejs.org installieren und das Skript erneut starten."
}
$nodeVersion = (& node --version).Trim()
$nodeMajor = [int]($nodeVersion.TrimStart('v').Split('.')[0])
Write-Host "Gefunden: Node.js $nodeVersion"
if ($nodeMajor -ne $RequiredNodeMajor) {
    Fail "Gefunden ist Node.js $nodeVersion, benötigt wird Version $RequiredNodeMajor.x. Mit nvm-windows z. B.: nvm install $RequiredNodeMajor && nvm use $RequiredNodeMajor"
}

# 3) Optional sauber neu installieren
if ($CleanInstall -and (Test-Path "node_modules")) {
    Write-Step "node_modules löschen"
    Remove-Item "node_modules" -Recurse -Force
}

# 4) Pakete installieren
Write-Step "Pakete installieren (npm install) – das dauert beim ersten Mal einige Minuten"
& npm install
if ($LASTEXITCODE -ne 0) { Fail "npm install ist fehlgeschlagen (siehe Meldungen oben)." }

# 5) Bundle (Produktion)
Write-Step "Bundle erstellen (gulp bundle --ship)"
& npx gulp bundle --ship
if ($LASTEXITCODE -ne 0) { Fail "gulp bundle ist fehlgeschlagen (siehe Meldungen oben)." }

# 6) Paket erstellen
Write-Step "Paket erstellen (gulp package-solution --ship)"
& npx gulp package-solution --ship
if ($LASTEXITCODE -ne 0) { Fail "gulp package-solution ist fehlgeschlagen (siehe Meldungen oben)." }

# 7) Ergebnis prüfen
$sppkg = Join-Path $ProjectPath (Join-Path $OutputFolder $SolutionFileName)
if (-not (Test-Path $sppkg)) { Fail "Build lief durch, aber '$sppkg' wurde nicht gefunden." }

$size = [math]::Round((Get-Item $sppkg).Length / 1KB, 0)
Write-Host ""
Write-Host "FERTIG: $sppkg ($size KB)" -ForegroundColor Green
Write-Host "Diese Datei in den SharePoint App Catalog hochladen."
Write-Host ""

if ($OpenFolderAtEnd) {
    Start-Process explorer.exe "/select,`"$sppkg`""
}
