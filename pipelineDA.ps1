<#
.SYNOPSIS
  Chaine complete : .docx francais -> texte anglais pret a publier sur DeviantArt.

.DESCRIPTION
  Une seule commande, un seul argument. Le script se charge du reste :

    0. Demarre le serveur LM Studio s'il est arrete, et charge le modele avec
       un contexte de 20 000 tokens (le chargement JIT n'en donne que 8 192).
    1. translate.ps1  -> <nom>_EN.docx
    2. review.ps1     -> <nom>_EN_revised.docx  + <nom>_EN_revised_notes.md
    3. format-deviantart.ps1 -> <nom>_EN_revised_DA.txt

  Tout atterrit dans E:\AIs\Output par defaut.

  Deux sorties, deux usages :
    .docx  archivage. Reouvrable dans Word, c'est la version a conserver.
    .txt   publication. Les lignes vides entre paragraphes sont litterales,
           donc rien ne peut les reflower ni les avaler comme le ferait Word.
           A ouvrir puis Ctrl+A / Ctrl+C, ou directement -Clipboard.

  -Format html reste disponible, mais n'apporte rien ici : il sert a preserver
  italiques et gras, or la traduction produit du texte brut sans mise en forme.

.EXAMPLE
  .\pipelineDA.ps1 "D:\textes\mon-recit.docx"

.EXAMPLE
  .\pipelineDA.ps1 "textes\mon-recit.docx" -ContextFile "contexte.md"

.EXAMPLE
  .\pipelineDA.ps1 "textes\mon-recit.docx" -OutDir "D:\Publications" -KeepIntermediates
#>

param(
    [Parameter(Mandatory = $true, Position = 0)][string]$Path,
    [string]$OutDir = "E:\AIs\Output",
    [string]$Context,
    [string[]]$ContextFile,
    [int]$MaxWords = 400,
    [int]$MaxParagraphs = 8,
    [int]$ContextLength = 20000,
    [ValidateSet('txt', 'html', 'both')][string]$Format = 'both',
    [string]$Model,
    [string]$Endpoint = "http://localhost:1234/v1/chat/completions",
    [switch]$SkipProofread,
    [switch]$SkipReview,
    [switch]$SkipFrench,
    [switch]$NoTypography,
    [switch]$KeepIntermediates,
    [switch]$Clipboard
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\lib-lmstudio.ps1"
if (-not $Model) { $Model = Get-DefaultModel }

if (-not (Test-Path $Path)) { throw "Fichier introuvable : $Path" }
$src = (Get-Item $Path).FullName
$name = [System.IO.Path]::GetFileNameWithoutExtension($src)
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
$OutDir = (Get-Item $OutDir).FullName

$frTxt      = Join-Path $OutDir "$name`_FR.txt"
$frDocx     = Join-Path $OutDir "$name`_FR.docx"
$frNotes    = Join-Path $OutDir "$name`_FR_notes.md"
$enTxt      = Join-Path $OutDir "$name`_EN.txt"
$enDocx     = Join-Path $OutDir "$name`_EN.docx"
$revTxt     = Join-Path $OutDir "$name`_EN_revised.txt"
$revDocx    = Join-Path $OutDir "$name`_EN_revised.docx"
$notes      = Join-Path $OutDir "$name`_EN_revised_notes.md"

$total = [System.Diagnostics.Stopwatch]::StartNew()

Write-Host ""
Write-Host "=== 1. LM Studio ===" -ForegroundColor Magenta
Initialize-LMStudioSession -Model $Model -Endpoint $Endpoint -ContextLength $ContextLength

# ---------------------------------------------------------------- 2. correction FR
# En tete de chaine : tout ce qui suit part d'un francais deja corrige.
$frSource = $src
if (-not $SkipProofread) {
    Write-Host ""
    Write-Host "=== 2. Correction du francais ===" -ForegroundColor Magenta
    $pArgs = @{
        Path          = $src
        Out           = $frTxt
        ReportPath    = $frNotes
        MaxWords      = $MaxWords
        MaxParagraphs = $MaxParagraphs
        Marks         = $true
        Model         = $Model
        Endpoint      = $Endpoint
    }
    if ($Context)     { $pArgs.Context     = $Context }
    if ($ContextFile) { $pArgs.ContextFile = $ContextFile }
    & "$PSScriptRoot\proofread.ps1" @pArgs

    if (-not (Test-Path $frTxt)) { throw "La correction n'a produit aucun fichier." }
    Write-DocxFromText -Path $frDocx -Text ([System.IO.File]::ReadAllText($frTxt, [System.Text.Encoding]::UTF8))
    Write-Host "-> $frDocx" -ForegroundColor Green
    $frSource = $frDocx
}

# ------------------------------------------------- 3. version francaise a publier
if (-not $SkipFrench) {
    Write-Host ""
    Write-Host "=== 3. Mise en forme francaise ===" -ForegroundColor Magenta
    # Le fichier corrige s'appelle deja <nom>_FR : eviter <nom>_FR_FR_DA.
    $frSuffix = if ([System.IO.Path]::GetFileNameWithoutExtension($frSource) -match '_FR$') { "_DA" } else { "_FR_DA" }
    $frArgs = @{
        Path        = $frSource
        OutDir      = $OutDir
        Suffix      = $frSuffix
        Format      = 'both'   # le .html conserve les italiques
        French      = $true
        SmartQuotes = $true
        Force       = $true
    }
    & "$PSScriptRoot\format-deviantart.ps1" @frArgs
}

# ---------------------------------------------------------------- 4. traduction
Write-Host ""
Write-Host "=== 4. Traduction ===" -ForegroundColor Magenta
$tArgs = @{
    Path          = $frSource
    Out           = $enTxt
    MaxWords      = $MaxWords
    MaxParagraphs = $MaxParagraphs
    Marks         = $true
    Model         = $Model
    Endpoint      = $Endpoint
}
if ($Context)     { $tArgs.Context     = $Context }
if ($ContextFile) { $tArgs.ContextFile = $ContextFile }
& "$PSScriptRoot\translate.ps1" @tArgs

if (-not (Test-Path $enTxt)) { throw "La traduction n'a produit aucun fichier." }
Write-DocxFromText -Path $enDocx -Text ([System.IO.File]::ReadAllText($enTxt, [System.Text.Encoding]::UTF8))
Write-Host "-> $enDocx" -ForegroundColor Green

# ---------------------------------------------------------------- 2. revision
$toFormat = $enDocx
if (-not $SkipReview) {
    Write-Host ""
    Write-Host "=== 5. Revision ===" -ForegroundColor Magenta
    $rArgs = @{
        Source        = $frSource
        Translation   = $enDocx
        Marks         = $true
        Out           = $revTxt
        ReportPath    = $notes
        MaxWords      = $MaxWords
        MaxParagraphs = $MaxParagraphs
        Model         = $Model
        Endpoint      = $Endpoint
    }
    if ($Context)     { $rArgs.Context     = $Context }
    if ($ContextFile) { $rArgs.ContextFile = $ContextFile }
    & "$PSScriptRoot\review.ps1" @rArgs

    if (-not (Test-Path $revTxt)) { throw "La revision n'a produit aucun fichier." }
    Write-DocxFromText -Path $revDocx -Text ([System.IO.File]::ReadAllText($revTxt, [System.Text.Encoding]::UTF8))
    Write-Host "-> $revDocx" -ForegroundColor Green
    $toFormat = $revDocx
}

# ---------------------------------------------------------------- 3. DeviantArt
Write-Host ""
Write-Host "=== 6. Mise en forme anglaise ===" -ForegroundColor Magenta
$fArgs = @{
    Path   = $toFormat
    OutDir = $OutDir
    Format = $Format
    Model  = $Model
    Endpoint = $Endpoint
    Force  = $true
}
if (-not $NoTypography) { $fArgs.English = $true; $fArgs.SmartQuotes = $true }
# L'auteur garde les tirets de dialogue a la francaise : ne pas les signaler.
$fArgs.KeepDashes = $true
if ($Clipboard)         { $fArgs.Clipboard = $true }
& "$PSScriptRoot\format-deviantart.ps1" @fArgs

# ---------------------------------------------------------------- bilan
if (-not $KeepIntermediates) {
    foreach ($f in @($frTxt, $enTxt, $revTxt)) {
        if (Test-Path $f) { Remove-Item $f -Force }
    }
}

$total.Stop()
Write-Host ""
Write-Host "=== Termine en $([math]::Round($total.Elapsed.TotalMinutes,1)) min ===" -ForegroundColor Yellow
Get-ChildItem $OutDir -Filter "$name`_*" | Sort-Object Name |
    Format-Table @{n='Fichier';e={$_.Name}}, @{n='Ko';e={[math]::Round($_.Length/1KB,1)}} -AutoSize
