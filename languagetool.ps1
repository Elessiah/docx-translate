<#
.SYNOPSIS
  Installe, demarre, arrete ou interroge le correcteur LanguageTool local.

.DESCRIPTION
  LanguageTool est un correcteur a base de regles. Il ne comprend pas ce qu'il
  lit : il rate tout ce qui demande de savoir qui parle ou de quel genre est le
  narrateur. En echange, ce qu'il trouve, il le situe exactement, il ne reecrit
  jamais rien, et il traite un texte entier en quelques secondes.

  proofread.ps1 -ReportOnly s'en sert automatiquement : il demarre le serveur
  s'il le faut et l'arrete a la fin. Ce script-ci n'est la que pour les cas ou
  on veut la main : installer, verifier, ou couper un serveur oublie.

  RIEN N'EST INSTALLE DANS WINDOWS. Pas de service, pas d'entree au demarrage,
  aucune modification systeme : tout tient dans tools\, et le serveur est un
  java.exe qu'on lance et qu'on tue. La machine redemarre sans rien de tout ca.

.EXAMPLE
  # Premiere fois : telecharge le correcteur et son java (environ 300 Mo)
  .\languagetool.ps1 -Install

.EXAMPLE
  # Verifier que rien ne tourne
  .\languagetool.ps1 -Status

.EXAMPLE
  # Couper un serveur reste en memoire
  .\languagetool.ps1 -Stop
#>

param(
    [switch]$Install,
    [switch]$Start,
    [switch]$Stop,
    [switch]$Status,
    [string]$Endpoint = "http://localhost:8081/v2"
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\lib-lmstudio.ps1"

# URLs officielles. Le java est celui d'Eclipse Temurin, en archive : il
# s'extrait dans tools\jre et ne s'enregistre nulle part dans le systeme.
$LT_URL  = "https://languagetool.org/download/LanguageTool-stable.zip"
$JRE_URL = "https://api.adoptium.net/v3/binary/latest/21/ga/windows/x64/jre/hotspot/normal/eclipse"

function Show-Status {
    $inst = Get-LanguageToolInstall -Root $PSScriptRoot
    Write-Host ""
    if ($inst.Ok) {
        Write-Host "  Installe  : oui" -ForegroundColor Green
        Write-Host "  Dossier   : $($inst.Dir)" -ForegroundColor DarkGray
        Write-Host "  Java      : $($inst.Java)" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  Installe  : non" -ForegroundColor DarkYellow
        Write-Host "  Lancez .\languagetool.ps1 -Install" -ForegroundColor DarkGray
    }

    $up = Test-LanguageToolServer -Endpoint $Endpoint
    if ($up) {
        Write-Host "  Serveur   : en cours d'execution sur $Endpoint" -ForegroundColor Yellow
        Write-Host "  Pour le couper : .\languagetool.ps1 -Stop" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  Serveur   : arrete (rien en memoire)" -ForegroundColor Green
    }

    $j = @(Get-Process java -ErrorAction SilentlyContinue)
    Write-Host ("  java.exe  : {0} processus" -f $j.Count) -ForegroundColor DarkGray
    Write-Host ""
}

function Get-Archive {
    param([string]$Url, [string]$Dest, [string]$Label)
    Write-Host "  $Label..." -NoNewline -ForegroundColor DarkGray
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    # Invoke-WebRequest suit les redirections et affiche sa propre barre de
    # progression, qui divise le debit par dix sur les gros fichiers.
    $old = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    try {
        Invoke-WebRequest -Uri $Url -OutFile $Dest -UseBasicParsing
    }
    finally { $ProgressPreference = $old }
    $mo = (Get-Item $Dest).Length / 1MB
    Write-Host (" {0:N0} Mo en {1:N0} s" -f $mo, $sw.Elapsed.TotalSeconds) -ForegroundColor DarkGray
}

# ------------------------------------------------------------------ install
if ($Install) {
    $tools = Join-Path $PSScriptRoot 'tools'
    if (-not (Test-Path $tools)) { New-Item -ItemType Directory -Path $tools | Out-Null }

    $inst = Get-LanguageToolInstall -Root $PSScriptRoot
    if ($inst.Ok) {
        Write-Host "Deja installe. Rien a faire." -ForegroundColor Green
        Show-Status
        return
    }

    Write-Host ""
    Write-Host "Installation de LanguageTool dans $tools" -ForegroundColor Cyan
    Write-Host "Environ 300 Mo a telecharger, 400 Mo sur le disque." -ForegroundColor DarkGray
    Write-Host "Rien n'est ecrit ailleurs, rien n'est enregistre dans Windows." -ForegroundColor DarkGray
    Write-Host ""

    $jreZip = Join-Path $tools 'jre.zip'
    $ltZip  = Join-Path $tools 'lt.zip'

    if (-not (Test-Path (Join-Path $tools 'jre\bin\java.exe'))) {
        Get-Archive -Url $JRE_URL -Dest $jreZip -Label "java (Eclipse Temurin 21)"
        Write-Host "  extraction..." -NoNewline -ForegroundColor DarkGray
        Expand-Archive -Path $jreZip -DestinationPath (Join-Path $tools 'jre-tmp') -Force
        $src = Get-ChildItem (Join-Path $tools 'jre-tmp') -Directory | Select-Object -First 1
        Move-Item $src.FullName (Join-Path $tools 'jre')
        Remove-Item (Join-Path $tools 'jre-tmp') -Recurse -Force
        Remove-Item $jreZip -Force
        Write-Host " ok" -ForegroundColor DarkGray
    }

    if (-not (Test-Path (Join-Path $tools 'languagetool\languagetool-server.jar'))) {
        Get-Archive -Url $LT_URL -Dest $ltZip -Label "LanguageTool"
        Write-Host "  extraction..." -NoNewline -ForegroundColor DarkGray
        Expand-Archive -Path $ltZip -DestinationPath (Join-Path $tools 'lt-tmp') -Force
        $src = Get-ChildItem (Join-Path $tools 'lt-tmp') -Directory | Select-Object -First 1
        Move-Item $src.FullName (Join-Path $tools 'languagetool')
        Remove-Item (Join-Path $tools 'lt-tmp') -Recurse -Force
        Remove-Item $ltZip -Force
        Write-Host " ok" -ForegroundColor DarkGray
    }

    Write-Host ""
    Write-Host "Verification : demarrage puis arret du serveur..." -ForegroundColor DarkGray
    if (Start-LanguageToolServer -Endpoint $Endpoint) {
        Write-Host "Le correcteur repond." -ForegroundColor Green
        Stop-LanguageToolServer
    }
    else {
        Write-Host "Le serveur n'a pas repondu. Verifiez tools\ et relancez -Status." -ForegroundColor Red
    }
    Show-Status
    return
}

# -------------------------------------------------------------------- start
if ($Start) {
    if (Test-LanguageToolServer -Endpoint $Endpoint) {
        Write-Host "Deja demarre sur $Endpoint" -ForegroundColor Green
        return
    }
    Write-Host "Demarrage..." -NoNewline -ForegroundColor DarkGray
    if (Start-LanguageToolServer -Endpoint $Endpoint) {
        Write-Host " ok" -ForegroundColor Green
        Write-Host "Il restera en memoire jusqu'a .\languagetool.ps1 -Stop" -ForegroundColor DarkYellow
    }
    else {
        Write-Host " echec" -ForegroundColor Red
        Write-Host "Installe ? .\languagetool.ps1 -Status" -ForegroundColor DarkGray
    }
    return
}

# --------------------------------------------------------------------- stop
if ($Stop) {
    # -Force : ici l'utilisateur demande explicitement l'arret, meme d'un
    # serveur lance par une autre session.
    Stop-LanguageToolServer -Force
    $j = @(Get-Process java -ErrorAction SilentlyContinue)
    foreach ($p in $j) {
        try {
            $cl = (Get-CimInstance Win32_Process -Filter "ProcessId=$($p.Id)").CommandLine
            if ($cl -match 'languagetool-server\.jar') {
                Stop-Process -Id $p.Id -Force
                Write-Host "Serveur arrete (pid $($p.Id))." -ForegroundColor Green
            }
        }
        catch { }
    }
    if (-not (Test-LanguageToolServer -Endpoint $Endpoint)) {
        Write-Host "Plus rien n'ecoute sur $Endpoint" -ForegroundColor Green
    }
    return
}

# ------------------------------------------------------------------- status
Show-Status
