<#
.SYNOPSIS
  Chaine complete : .docx francais -> texte anglais pret a publier sur DeviantArt.

.DESCRIPTION
  Une seule commande, un seul argument. Le script se charge du reste :

    1. LM Studio demarre si besoin (correction, ou traduction -Engine local),
       avec un contexte de 20 000 tokens (le chargement JIT n'en donne que 8 192).
    2. proofread.ps1         -> <nom>_FR.docx            francais corrige
    3. format-deviantart.ps1 -> <nom>_FR_DA[_pN].txt/.html  version francaise
    4. traduction            -> <nom>_EN.docx
         -Engine deepl (defaut) : deepl-api.ps1, glossaire applique
         -Engine local          : translate.ps1, puis revision review.ps1
    5. format-deviantart.ps1 -> <nom>_EN[_revised]_DA[_pN].txt/.html

  Avec DeepL, pas de revision par le modele local : mesuree, elle degradait la
  traduction (paragraphes recopies en francais sans que le controle le voie).

  Tout atterrit dans E:\AIs\Output par defaut.

  Deux sorties, deux usages :
    .html  publication. Ctrl+A / Ctrl+C dans le navigateur puis coller dans
           DeviantArt : italique et gras conserves.
    .txt   texte brut, mise en forme perdue. -Clipboard le copie directement.

.EXAMPLE
  .\pipelineDA.ps1 "D:\textes\mon-recit.docx"

.EXAMPLE
  .\pipelineDA.ps1 "textes\mon-recit.docx" -ContextFile "contexte.md"

.EXAMPLE
  .\pipelineDA.ps1 "textes\mon-recit.docx" -OutDir "D:\Publications" -KeepIntermediates

.EXAMPLE
  .\pipelineDA.ps1 "textes\mon-recit.docx" -SplitChapters -FrHeader entete-fr.txt -Header entete-en.txt
  Correction, decoupage par chapitre et mise en forme en francais, puis
  traduction DeepL et la meme chose en anglais.

.EXAMPLE
  .\pipelineDA.ps1 "textes\mon-recit.docx" -CheckNarrator -ContextFile contexte.md
  Recit a la premiere personne : liste les accords qui ne correspondent pas au
  genre du narrateur dans <nom>_narrateur.md, puis S'ARRETE. Corriger dans Word,
  et relancer sans -CheckNarrator. Le genre est lu dans la fiche
  ("Narratrice : ..." ou "Genre du narrateur : F"), ou force par -Narrator.

.EXAMPLE
  .\pipelineDA.ps1 "textes\mon-recit.docx" -FrenchOnly -SplitChapters -FrHeader entete-fr.txt
  Corrige le francais, le coupe en une deviation par chapitre, le met en forme,
  et s'arrete la : ni traduction ni revision.

  Publication (transmis a format-deviantart.ps1) :
    -SplitChapters / -SplitWords  decoupage en plusieurs deviations
    -Header / -Footer             entete et pied de la version anglaise
    -FrHeader / -FrFooter         entete et pied de la version francaise
    -FrenchOnly                   s'arrete apres la mise en forme francaise
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
    [ValidateSet('deepl', 'local')][string]$Engine = 'deepl',
    [string]$Model,
    [string]$Endpoint = "http://localhost:1234/v1/chat/completions",
    [switch]$SkipProofread,
    [switch]$SkipReview,
    [switch]$SkipFrench,
    [switch]$NoTypography,
    [switch]$KeepIntermediates,
    [switch]$SplitChapters,
    [int]$SplitWords = 0,
    [string]$Header,
    [string]$Footer,
    [string]$FrHeader,
    [string]$FrFooter,
    [switch]$FrenchOnly,
    [switch]$CheckNarrator,
    [ValidateSet('', 'feminin', 'masculin')][string]$Narrator = '',
    [switch]$Clipboard
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\lib-lmstudio.ps1"
if (-not $Model) { $Model = Get-DefaultModel }

if (-not (Test-Path $Path)) { throw "Fichier introuvable : $Path" }
if ($FrenchOnly -and $SkipFrench) { throw "-FrenchOnly et -SkipFrench s'excluent : il ne resterait rien a faire." }
foreach ($h in @($Header, $Footer, $FrHeader, $FrFooter)) {
    # Un chemin mal tape serait colle tel quel comme texte libre : on prefere s'arreter.
    if ($h -and $h -match '\.(txt|md)$' -and -not (Test-Path -LiteralPath $h)) { throw "Fichier d'entete ou de pied introuvable : $h" }
}
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
# Le modele local ne sert qu'a la correction et a la traduction -Engine local.
$needLocal = $CheckNarrator -or (-not $SkipProofread) -or ($Engine -eq 'local' -and -not $FrenchOnly)
if ($needLocal) { Initialize-LMStudioSession -Model $Model -Endpoint $Endpoint -ContextLength $ContextLength }
else { Write-Host "non requis" -ForegroundColor DarkGray }

# ------------------------------------------------------ verification du narrateur
# Un rapport, pas une correction : c'est l'auteur qui tranche. La chaine s'arrete
# donc ici, pour qu'il corrige le .docx avant de relancer sans -CheckNarrator.
if ($CheckNarrator) {
    Write-Host ""
    Write-Host "=== Accords avec le narrateur ===" -ForegroundColor Magenta
    $narReport = Join-Path $OutDir "$name`_narrateur.md"
    $nArgs = @{
        Path         = $src
        NarratorOnly = $true
        ReportPath   = $narReport
        Model        = $Model
        Endpoint     = $Endpoint
    }
    if ($Narrator)    { $nArgs.Narrator    = $Narrator }
    if ($Context)     { $nArgs.Context     = $Context }
    if ($ContextFile) { $nArgs.ContextFile = $ContextFile }
    & "$PSScriptRoot\proofread.ps1" @nArgs

    $total.Stop()
    Write-Host ""
    Write-Host "=== Termine en $([math]::Round($total.Elapsed.TotalMinutes,1)) min ===" -ForegroundColor Yellow
    Write-Host "Rapport : $narReport" -ForegroundColor Green
    Write-Host "Corrige le .docx dans Word, puis relance la pipeline sans -CheckNarrator." -ForegroundColor Yellow
    return
}

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
    if ($SplitChapters)    { $frArgs.SplitChapters = $true }
    if ($SplitWords -gt 0) { $frArgs.SplitWords = $SplitWords }
    if ($FrHeader)         { $frArgs.Header = $FrHeader }
    if ($FrFooter)         { $frArgs.Footer = $FrFooter }
    & "$PSScriptRoot\format-deviantart.ps1" @frArgs
}

if ($FrenchOnly) {
    if (-not $KeepIntermediates -and (Test-Path $frTxt)) { Remove-Item $frTxt -Force }
    $total.Stop()
    Write-Host ""
    Write-Host "=== Termine en $([math]::Round($total.Elapsed.TotalMinutes,1)) min (francais seul) ===" -ForegroundColor Yellow
    Get-ChildItem $OutDir -Filter "$name`_*" | Sort-Object Name |
        Format-Table @{n='Fichier';e={$_.Name}}, @{n='Ko';e={[math]::Round($_.Length/1KB,1)}} -AutoSize
    return
}

# ---------------------------------------------------------------- 4. traduction
Write-Host ""
Write-Host ("=== 4. Traduction ({0}) ===" -f $Engine) -ForegroundColor Magenta
if ($Engine -eq 'deepl') {
    # Ecrit <nom>_EN.docx avec italiques et gras, et <nom>_EN.txt a cote.
    $dArgs = @{ Path = $frSource; Out = $enDocx; OutDir = $OutDir }
    if ($Context)     { $dArgs.Context     = $Context }
    if ($ContextFile) { $dArgs.ContextFile = $ContextFile }
    & "$PSScriptRoot\deepl-api.ps1" translate @dArgs
    if (-not (Test-Path $enDocx)) { throw "La traduction DeepL n'a produit aucun fichier." }
}
else {
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
}
Write-Host "-> $enDocx" -ForegroundColor Green

# ---------------------------------------------------------------- 5. revision
# Reservee a la traduction locale : sur DeepL, la revision degradait la sortie.
$toFormat = $enDocx
if ($Engine -eq 'local' -and -not $SkipReview) {
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
if ($SplitChapters)     { $fArgs.SplitChapters = $true }
if ($SplitWords -gt 0)  { $fArgs.SplitWords = $SplitWords }
if ($Header)            { $fArgs.Header = $Header }
if ($Footer)            { $fArgs.Footer = $Footer }
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
