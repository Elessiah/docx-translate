<#
.SYNOPSIS
  Revise une traduction EN en la confrontant a l'original FR.

.DESCRIPTION
  Prend l'original francais ET sa traduction anglaise, les aligne paragraphe par
  paragraphe, et fait relire chaque bloc par le modele. Produit :
    - un fichier de traduction corrigee
    - un rapport .md listant chaque correction et sa raison

  Ce que la relecture cherche : contresens, omissions, ajouts, passages
  affaiblis, ponctuation de dialogue, calques d'idiomes,
  incoherences de temps / pronoms / terminologie, derives de registre.

  Formats acceptes : .docx, .txt, .md. Les .pdf doivent etre convertis d'abord.

.EXAMPLE
  .\review.ps1 -Source "textes\mon-recit.docx" -Translation "textes\mon-recit_EN.txt"
  .\review.ps1 -Source a.docx -Translation a_EN.txt -Out final.txt -MaxWords 500
#>

param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Translation,
    [string]$Out,
    [string]$ReportPath,
    [int]$MaxWords = 400,
    [int]$MaxParagraphs = 8,
    [int]$MaxRetries = 3,
    [string]$Model,
    [string]$Endpoint = "http://localhost:1234/v1/chat/completions",
    [double]$Temperature = 0.2,
    [string]$PromptFile,
    [int]$ContextLength = 20000,
    [string]$Context,
    [string[]]$ContextFile,
    [switch]$Marks,
    [switch]$NoThink
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\lib-lmstudio.ps1"
if (-not $Model) { $Model = Get-DefaultModel }

# Demarre le serveur et charge le modele si besoin : un script lance seul
# ne doit pas echouer juste parce que LM Studio s'est arrete.
Initialize-LMStudioSession -Model $Model -Endpoint $Endpoint -ContextLength $ContextLength

if ($Marks) {
    # gras et italique transportes sous forme de balises [[i]] / [[b]]
    $frText = if ((Get-Item $Source).Extension -eq '.docx')      { Read-DocxMarkedText -File $Source }      else { Read-SourceText -File $Source }
    $enText = if ((Get-Item $Translation).Extension -eq '.docx') { Read-DocxMarkedText -File $Translation } else { Read-SourceText -File $Translation }
}
else {
    $frText = Read-SourceText -File $Source
    $enText = Read-SourceText -File $Translation
}

$transItem = Get-Item $Translation
if (-not $Out) {
    $ext = $transItem.Extension.ToLower()
    if ($ext -eq '.docx') { $ext = '.txt' }
    $Out = Join-Path $transItem.DirectoryName ($transItem.BaseName + "_revised" + $ext)
}
if (-not $ReportPath) {
    $ReportPath = [System.IO.Path]::ChangeExtension($Out, $null).TrimEnd('.') + "_notes.md"
}

# ---------------------------------------------------------------- alignement
$frParas = Get-Paragraphs -Text $frText
$enParas = Get-Paragraphs -Text $enText

Write-Host "Original    : $Source  ($(Measure-Words $frText) mots, $($frParas.Count) paragraphes)" -ForegroundColor Cyan
Write-Host "Traduction  : $Translation  ($(Measure-Words $enText) mots, $($enParas.Count) paragraphes)" -ForegroundColor Cyan

if ($frParas.Count -eq $enParas.Count) {
    # Alignement 1:1, regroupement par budget de mots
    $frChunks = @(); $enChunks = @()
    $curFr = @(); $curEn = @(); $count = 0
    for ($i = 0; $i -lt $frParas.Count; $i++) {
        $w = Measure-Words $frParas[$i]
        # Memes trois contraintes que translate.ps1 : mots, paragraphes, et jamais
        # de rupture de chapitre au milieu d'un bloc.
        $mustSplit = $curFr.Count -gt 0 -and (
            (Test-StructuralBreak -Paragraph $frParas[$i]) -or
            ($count + $w) -gt $MaxWords -or
            $curFr.Count -ge $MaxParagraphs
        )
        if ($mustSplit) {
            $frChunks += , ($curFr -join "`n`n"); $enChunks += , ($curEn -join "`n`n")
            $curFr = @(); $curEn = @(); $count = 0
        }
        $curFr += $frParas[$i]; $curEn += $enParas[$i]; $count += $w
    }
    if ($curFr.Count) { $frChunks += , ($curFr -join "`n`n"); $enChunks += , ($curEn -join "`n`n") }
    Write-Host "Alignement  : 1:1 sur les paragraphes -> $($frChunks.Count) bloc(s)" -ForegroundColor Cyan
}
else {
    # Comptes differents : decoupage proportionnel, meme nombre de blocs des deux cotes
    $msg = "Nombre de paragraphes different ({0} vs {1}) : le decoupage sera proportionnel " +
           "et non strictement aligne. Verifie le resultat de pres."
    Write-Warning ($msg -f $frParas.Count, $enParas.Count)
    $frChunks = Group-IntoChunks -Paragraphs $frParas -MaxWords $MaxWords
    $n = $frChunks.Count
    $per = [Math]::Max(1, [Math]::Ceiling($enParas.Count / $n))
    $enChunks = @()
    for ($i = 0; $i -lt $n; $i++) {
        $slice = $enParas | Select-Object -Skip ($i * $per) -First $per
        $enChunks += , (($slice -join "`n`n"))
    }
    Write-Host "Alignement  : proportionnel -> $n bloc(s)" -ForegroundColor Cyan
}
Write-Host ""

$systemPrompt = Get-PromptText -Name 'review' -PromptFile $PromptFile

if ($Marks) {
    $systemPrompt += @'


BALISES DE MISE EN FORME (regle absolue) :
Le texte contient des balises [[i]] ... [[/i]] (italique) et [[b]] ... [[/b]] (gras).
Ce ne sont PAS du texte : ne les corrige pas, ne les commente pas, ne les supprime pas.
Recopie-les telles quelles autour des memes passages. Leur nombre doit etre
identique dans ta reponse et dans la traduction fournie. Ne les mentionne jamais
dans les NOTES.
'@
}

# Contexte auteur : personnages, genres, univers, terminologie a respecter.
$ctx = ""
foreach ($cf in $ContextFile) { if ($cf) { $ctx = ($ctx + "`n`n" + (Read-SourceText -File $cf).Trim()).Trim() } }
if ($Context)     { $ctx = ($ctx + "`n" + $Context).Trim() }
if ($ctx) {
    $systemPrompt += "`n`nCONTEXTE DE LA NOUVELLE (fourni par l'auteur, a respecter imperativement) :`n" +
                     "---`n$ctx`n---"
    Write-Host "Contexte    : $(Measure-Words $ctx) mots" -ForegroundColor Cyan
}

# ---------------------------------------------------------------- revision
$revised = @(); $allNotes = @()
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$genTokens = 0

$mismatches = @(); $noOpTotal = 0; $templateEchoes = 0

for ($i = 0; $i -lt $frChunks.Count; $i++) {
    $expected = (Get-Paragraphs -Text $enChunks[$i]).Count
    Write-Host ("[{0}/{1}] {2} paragraphe(s)..." -f ($i + 1), $frChunks.Count, $expected) -NoNewline

    $user = "=== ORIGINAL (francais) ===`n" + $frChunks[$i] +
            "`n`n=== TRADUCTION (anglais) ===`n" + $enChunks[$i]

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

        $r = Invoke-LMStudioChat -System $sysTry -User $user -Model $Model `
                -Endpoint $Endpoint -Temperature $Temperature -Think:(-not $NoThink) `
                -MaxTokens ([Math]::Max(2048, $MaxWords * 5))

        $genTokens += $r.Tokens
        $parts = [regex]::Split($r.Text, '(?m)^\s*=+\s*NOTES\s*=+\s*$')

        if ($parts.Count -ge 2) {
            $body = $parts[0].Trim()
            $notes = $parts[1].Trim()
        }
        else {
            # Le modele n'a pas respecte le separateur : on garde tout comme texte
            $body = $r.Text.Trim()
            $notes = "(separateur ===NOTES=== absent dans la reponse du modele pour ce bloc)"
        }

        # Le modele recopie parfois le gabarit du prompt en tete de reponse.
        $echo = Remove-TemplateEcho -Text $body
        if ($echo.Removed) { $templateEchoes += $echo.Removed }
        $body = $echo.Text

        $got = (Get-Paragraphs -Text $body).Count
        $marksOk = (-not $Marks) -or (Test-MarksIntact -Before $enChunks[$i] -After $body)
        if ($got -eq $expected -and $marksOk) {
            $best = $body; $bestNotes = $notes; $bestGot = $got; $bestMarks = $true; $bestScore = 0
            break
        }

        # Ecart de paragraphes, penalite si balises perdues, penalite si du texte manque.
        $score = [Math]::Abs($got - $expected) +
                 $(if ($marksOk) { 0 } else { 5 }) +
                 [Math]::Max(0, [int](((Measure-Words $enChunks[$i]) - (Measure-Words $body)) / 20))
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

    $revised += $body

    # Le modele signale parfois des "corrections" ou avant et apres sont identiques.
    $clean = Remove-NoOpNotes -Notes $notes
    $noOpTotal += $clean.Dropped
    if (-not $clean.Notes) { $clean.Notes = "- RAS" }
    $allNotes += "## Bloc $($i + 1)`n`n$($clean.Notes)"
}

$sw.Stop()

$rep = Repair-DialogueQuotes -Text ($revised -join "`n`n")
$revisedText = $rep.Text
Write-Utf8NoBom -Path $Out -Content $revisedText

$header = "# Rapport de revision`n`n" +
          "- Original   : $Source`n" +
          "- Traduction : $Translation`n" +
          "- Revision   : $Out`n" +
          "- Date       : $(Get-Date -Format 'yyyy-MM-dd HH:mm')`n`n"
Write-Utf8NoBom -Path $ReportPath -Content ($header + ($allNotes -join "`n`n"))

$before = Measure-Words $enText
$after = Measure-Words $revisedText

Write-Host ""
Write-Host ("Termine en {0:N0} s - {1} tokens - {2:N1} tok/s" -f `
    $sw.Elapsed.TotalSeconds, $genTokens, ($genTokens / $sw.Elapsed.TotalSeconds)) -ForegroundColor Yellow
Write-Host ("Mots : {0} -> {1}" -f $before, $after) -ForegroundColor Yellow
Write-Host ("Paragraphes : {0} -> {1}" -f $enParas.Count, (Get-Paragraphs -Text $revisedText).Count) -ForegroundColor Yellow

if ($mismatches.Count) {
    Write-Host ""
    Write-Warning "Blocs encore en ecart apres relance - a relire en priorite :"
    $mismatches | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
}

if ($templateEchoes) {
    Write-Host ("Gabarit de prompt recopie et retire : {0} ligne(s)" -f $templateEchoes) -ForegroundColor DarkGray
}
if ($noOpTotal) {
    Write-Host ("Notes vides filtrees : {0} (correction annoncee mais identique)" -f $noOpTotal) -ForegroundColor DarkGray
}

if ($rep.Fixed) {
    Write-Host ("Guillemets ouvrants restaures : {0} paragraphe(s)" -f $rep.Fixed) -ForegroundColor DarkYellow
}
$badBefore = Get-UnbalancedQuoteParagraphs -Text $enText
$badAfter  = Get-UnbalancedQuoteParagraphs -Text $revisedText
Write-Host ("Guillemets desequilibres : {0} paragraphe(s) avant -> {1} apres" -f `
    $badBefore.Count, $badAfter.Count) -ForegroundColor Yellow
if ($badAfter.Count) {
    Write-Warning ("Restent a corriger, paragraphe(s) : " + ($badAfter -join ', '))
}

Write-Host "Revision : $Out" -ForegroundColor Yellow
Write-Host "Rapport  : $ReportPath" -ForegroundColor Yellow
