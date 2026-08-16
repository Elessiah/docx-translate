<#
.SYNOPSIS
  Corrige un texte francais avant traduction : orthographe, grammaire, typographie.

.DESCRIPTION
  Passe le texte au modele local bloc par bloc et corrige ce qui est objectivement
  fautif, sans reecrire le style de l'auteur :

    - orthographe, accords, conjugaison
    - participe passe, homophones (a/a, ou/ou, ces/ses, se/ce, leur/leurs)
    - ponctuation, majuscules, espaces insecables avant ? ! ; :
    - guillemets et tirets de dialogue coherents
    - repetitions involontaires a quelques mots d'intervalle

  Ce qu'il ne touche PAS : le vocabulaire, le registre, le rythme, les phrases
  nominales, les tournures orales, le fond du propos. Un texte litteraire n'est pas
  une dissertation : les libertes de style sont volontaires.

  Produit le texte corrige et un rapport listant chaque correction.

  Comme translate.ps1 et review.ps1 : garde-fou sur le nombre de paragraphes avec
  relance automatique, et -Marks pour transporter gras et italique.

  Mode reflexion DESACTIVE par defaut. Mesure sur un texte de 1168 mots :
  avec reflexion le texte perd des paragraphes (69 -> 66, un bloc en echec meme
  apres relance) pour 20 min de traitement ; sans, 69 -> 69 en 12 min.
  Comme pour la traduction, la reflexion casse les taches qui reproduisent un
  texte entier. -Think reste disponible pour essayer.

  La typographie francaise (apostrophes courbes, espaces insecables) est
  retablie mecaniquement apres coup : le modele la degrade malgre la consigne.

.EXAMPLE
  .\proofread.ps1 -Path "textes\mon-recit.docx"

.EXAMPLE
  .\proofread.ps1 -Path "textes\mon-recit.docx" -Out "textes\mon-recit_corrige.txt" -Marks
#>

param(
    [Parameter(Mandatory = $true, Position = 0)][string]$Path,
    [string]$Out,
    [string]$ReportPath,
    [int]$MaxWords = 400,
    [int]$MaxParagraphs = 8,
    [int]$MaxRetries = 3,
    [string]$Context,
    [string[]]$ContextFile,
    [string]$Model,
    [string]$Endpoint = "http://localhost:1234/v1/chat/completions",
    [double]$Temperature = 0.2,
    [string]$PromptFile,
    [int]$ContextLength = 20000,
    [switch]$Marks,
    [switch]$Think
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\lib-lmstudio.ps1"
if (-not $Model) { $Model = Get-DefaultModel }

# Demarre le serveur et charge le modele si besoin : un script lance seul
# ne doit pas echouer juste parce que LM Studio s'est arrete.
Initialize-LMStudioSession -Model $Model -Endpoint $Endpoint -ContextLength $ContextLength

$item = Get-Item $Path
if ($Marks -and $item.Extension.ToLower() -eq '.docx') {
    $raw = Read-DocxMarkedText -File $Path
}
else {
    $raw = Read-SourceText -File $Path
}

if (-not $Out) {
    $ext = $item.Extension.ToLower()
    if ($ext -eq '.docx') { $ext = '.txt' }
    $Out = Join-Path $item.DirectoryName ($item.BaseName + "_corrige" + $ext)
}
if (-not $ReportPath) {
    $ReportPath = [System.IO.Path]::ChangeExtension($Out, $null).TrimEnd('.') + "_notes.md"
}

$systemPrompt = Get-PromptText -Name 'proofread' -PromptFile $PromptFile

if ($Marks) {
    $systemPrompt += @'


BALISES DE MISE EN FORME (regle absolue) :
Le texte contient des balises [[i]] ... [[/i]] (italique) et [[b]] ... [[/b]] (gras).
Ce ne sont PAS du texte : ne les corrige pas, ne les supprime pas, ne les commente pas.
Recopie-les telles quelles autour des memes passages. Leur nombre doit etre
identique dans ta reponse et dans le texte fourni.
'@
}

$ctx = ""
foreach ($cf in $ContextFile) { if ($cf) { $ctx = ($ctx + "`n`n" + (Read-SourceText -File $cf).Trim()).Trim() } }
if ($Context)     { $ctx = ($ctx + "`n" + $Context).Trim() }
if ($ctx) {
    $systemPrompt += "`n`nCONTEXTE DE LA NOUVELLE (fourni par l'auteur, a respecter imperativement) :`n" +
                     "---`n$ctx`n---`n" +
                     "Les noms propres et le vocabulaire qui y figurent ne sont pas des fautes."
}

# ---------------------------------------------------------------- decoupage
$paragraphs = Get-Paragraphs -Text $raw
$chunkParas = Group-ParagraphsIntoChunks -Paragraphs $paragraphs -MaxWords $MaxWords -MaxParagraphs $MaxParagraphs

Write-Host "Source  : $Path" -ForegroundColor Cyan
Write-Host "Sortie  : $Out" -ForegroundColor Cyan
Write-Host ("{0} mots, {1} paragraphes -> {2} bloc(s)" -f `
    (Measure-Words $raw), $paragraphs.Count, $chunkParas.Count) -ForegroundColor Cyan
if ($ctx) { Write-Host ("Contexte : {0} mots" -f (Measure-Words $ctx)) -ForegroundColor Cyan }
Write-Host ""

# ---------------------------------------------------------------- correction
$results = @(); $allNotes = @(); $mismatches = @(); $noOpTotal = 0; $templateEchoes = 0
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$genTokens = 0

for ($i = 0; $i -lt $chunkParas.Count; $i++) {
    $expected = $chunkParas[$i].Count
    $chunkText = $chunkParas[$i] -join "`n`n"
    Write-Host ("[{0}/{1}] {2} paragraphe(s)..." -f ($i + 1), $chunkParas.Count, $expected) -NoNewline

    # On garde la MEILLEURE tentative, pas la derniere : une relance peut rendre
    # bien pire, et la conserver perdrait du texte.
    $attempt = 0; $body = $null; $notes = $null; $got = -1; $marksOk = $true
    $best = $null; $bestNotes = $null; $bestGot = -1; $bestMarks = $false; $bestScore = [int]::MaxValue
    while ($attempt -le $MaxRetries) {
        $sysTry = $systemPrompt
        if ($attempt -gt 0) {
            $why = @()
            if ($got -ne $expected) { $why += "tu as rendu $got paragraphes au lieu de $expected" }
            if (-not $marksOk)      { $why += "tu n'as pas restitue les balises [[i]] / [[b]] a l'identique" }
            $sysTry += "`n`nATTENTION : ta tentative precedente etait invalide : " +
                       ($why -join ' ; ') + ". Recommence en rendant EXACTEMENT " +
                       "$expected paragraphes separes par une ligne vide, avant la ligne ===NOTES===."
        }

        $r = Invoke-LMStudioChat -System $sysTry -User $chunkText -Model $Model `
                -Endpoint $Endpoint -Temperature $Temperature -Think:$Think `
                -MaxTokens ([Math]::Max(2048, $MaxWords * 5))

        $genTokens += $r.Tokens
        $parts = [regex]::Split($r.Text, '(?m)^\s*=+\s*NOTES\s*=+\s*$')

        if ($parts.Count -ge 2) { $body = $parts[0].Trim(); $notes = $parts[1].Trim() }
        else {
            $body = $r.Text.Trim()
            $notes = "(separateur ===NOTES=== absent dans la reponse du modele pour ce bloc)"
        }

        $echo = Remove-TemplateEcho -Text $body
        if ($echo.Removed) { $templateEchoes += $echo.Removed }
        $body = $echo.Text

        $got = (Get-Paragraphs -Text $body).Count
        $marksOk = (-not $Marks) -or (Test-MarksIntact -Before $chunkText -After $body)
        if ($got -eq $expected -and $marksOk) {
            $best = $body; $bestNotes = $notes; $bestGot = $got; $bestMarks = $true; $bestScore = 0
            break
        }

        $score = [Math]::Abs($got - $expected) +
                 $(if ($marksOk) { 0 } else { 5 }) +
                 [Math]::Max(0, [int](((Measure-Words $chunkText) - (Measure-Words $body)) / 20))
        if ($score -lt $bestScore) {
            $bestScore = $score; $best = $body; $bestNotes = $notes
            $bestGot = $got; $bestMarks = $marksOk
        }
        $attempt++
    }
    if ($bestScore -gt 0 -and $best) {
        $body = $best; $notes = $bestNotes; $got = $bestGot; $marksOk = $bestMarks
    }

    if ($got -eq $expected -and $marksOk) {
        if ($attempt -eq 0) { Write-Host " ok" -ForegroundColor Green }
        else { Write-Host (" ok apres {0} relance(s)" -f $attempt) -ForegroundColor DarkYellow }
    }
    else {
        $what = @()
        if ($got -ne $expected) { $what += ("{0} paragraphes au lieu de {1}" -f $got, $expected) }
        if (-not $marksOk)      { $what += "balises de mise en forme perdues" }
        Write-Host (" ECART : " + ($what -join ', ')) -ForegroundColor Red
        $mismatches += ("bloc {0} : {1}" -f ($i + 1), ($what -join ', '))
    }

    $results += $body

    $clean = Remove-NoOpNotes -Notes $notes
    $noOpTotal += $clean.Dropped
    if (-not $clean.Notes) { $clean.Notes = "- RAS" }
    $allNotes += "## Bloc $($i + 1)`n`n$($clean.Notes)"
}

$sw.Stop()

$final = $results -join "`n`n"

# Le modele rend l'apostrophe droite et avale les insecables malgre la consigne :
# on retablit la typographie francaise mecaniquement.
$typo = Repair-FrenchTypography -Text $final
$final = $typo.Text
Write-Utf8NoBom -Path $Out -Content $final

$header = "# Rapport de correction (francais)`n`n" +
          "- Source  : $Path`n" +
          "- Corrige : $Out`n" +
          "- Date    : $(Get-Date -Format 'yyyy-MM-dd HH:mm')`n`n"
Write-Utf8NoBom -Path $ReportPath -Content ($header + ($allNotes -join "`n`n"))

Write-Host ""
Write-Host ("Termine en {0:N0} s - {1} tokens - {2:N1} tok/s" -f `
    $sw.Elapsed.TotalSeconds, $genTokens, ($genTokens / $sw.Elapsed.TotalSeconds)) -ForegroundColor Yellow
Write-Host ("Mots : {0} -> {1}" -f (Measure-Words $raw), (Measure-Words $final)) -ForegroundColor Yellow
Write-Host ("Paragraphes : {0} -> {1}" -f $paragraphs.Count, (Get-Paragraphs -Text $final).Count) -ForegroundColor Yellow
if ($typo.ApostrophesFixed -or $typo.NbspAdded) {
    Write-Host ("Typographie retablie : {0} apostrophe(s) courbe(s), {1} espace(s) insecable(s)" -f `
        $typo.ApostrophesFixed, $typo.NbspAdded) -ForegroundColor DarkYellow
}
if ($templateEchoes) {
    Write-Host ("Gabarit de prompt recopie et retire : {0} ligne(s)" -f $templateEchoes) -ForegroundColor DarkGray
}
if ($noOpTotal) {
    Write-Host ("Notes vides filtrees : {0}" -f $noOpTotal) -ForegroundColor DarkGray
}
if ($mismatches.Count) {
    Write-Host ""
    Write-Warning "Blocs encore en ecart apres relance - a relire en priorite :"
    $mismatches | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
}
Write-Host "Corrige : $Out" -ForegroundColor Yellow
Write-Host "Rapport : $ReportPath" -ForegroundColor Yellow
