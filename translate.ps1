<#
.SYNOPSIS
  Traduit un fichier FR -> EN via LM Studio en local.

.DESCRIPTION
  Formats acceptes : .docx, .txt, .md et tout fichier texte brut.
  Les .pdf ne sont pas supportes : convertir en .docx ou .txt d'abord.

  Decoupe le texte aux frontieres de paragraphes et traduit bloc par bloc,
  ce qui garantit que chaque bloc tient en contexte. Evite ainsi le
  basculement silencieux en RAG de l'interface graphique.

  GARDE-FOU ANTI-OUBLI : le nombre de paragraphes est compare en entree et en
  sortie de chaque bloc. En cas d'ecart, le bloc est automatiquement relance
  avec une consigne correctrice. Les blocs encore en ecart sont signales a la fin.

  Sortie : .txt UTF-8 (la mise en forme Word n'est pas conservee).

.EXAMPLE
  .\translate.ps1 -Path "D:\textes\mon-recit.docx"
  .\translate.ps1 -Path chapitre1.md -Out chapter1_EN.md -MaxWords 700
#>

param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$Out,
    [int]$MaxWords = 400,
    [int]$MaxParagraphs = 8,
    [int]$MaxRetries = 3,
    [string]$Model,
    [string]$Endpoint = "http://localhost:1234/v1/chat/completions",
    [double]$Temperature = 0.3,
    [string]$PromptFile,
    [int]$ContextLength = 20000,
    [string]$Context,
    [string[]]$ContextFile,
    [switch]$Think,
    [switch]$Marks,
    [switch]$NoSummary,
    [int]$SummaryFrom = 4,
    [switch]$Review,
    [string]$ReviewOut
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\lib-lmstudio.ps1"
if (-not $Model) { $Model = Get-DefaultModel }

# Demarre le serveur et charge le modele si besoin : un script lance seul
# ne doit pas echouer juste parce que LM Studio s'est arrete.
Initialize-LMStudioSession -Model $Model -Endpoint $Endpoint -ContextLength $ContextLength
Reset-TokenLedger -Label 'translate'

$item = Get-Item $Path
if ($Marks -and $item.Extension.ToLower() -eq '.docx') {
    # gras et italique transportes sous forme de balises [[i]] / [[b]]
    $raw = Read-DocxMarkedText -File $Path
}
else {
    $raw = Read-SourceText -File $Path
}
if (-not $Out) {
    $outExt = $item.Extension.ToLower()
    if ($outExt -eq '.docx') { $outExt = '.txt' }
    $Out = Join-Path $item.DirectoryName ($item.BaseName + "_EN" + $outExt)
}

$systemPrompt = Get-PromptText -Name 'translate' -PromptFile $PromptFile

# Regle injectee UNIQUEMENT dans les blocs qui contiennent des balises. La mettre
# partout apprend au modele a en produire, et il en inventait la ou il n'y en avait pas.
$marksRule = @'


BALISES DE MISE EN FORME (regle absolue) :
Ce passage contient des balises [[i]] ... [[/i]] (italique) et [[b]] ... [[/b]] (gras).
Ce ne sont PAS du texte : ne les traduis pas, ne les commente pas, ne les supprime pas.
Recopie-les telles quelles autour de la traduction du passage qu'elles encadraient.
N'en AJOUTE jamais autour d'un passage qui n'en avait pas. Le nombre de balises
doit etre exactement identique dans ta reponse et dans le texte source.
  Il lisait [[i]]Le Monde[[/i]] hier.  ->  He was reading [[i]]Le Monde[[/i]] yesterday.
'@

# Contexte auteur : personnages, genres, univers, terminologie a respecter.
# -Context "texte libre"  et/ou  -ContextFile "chemin.txt|.md|.docx"
$ctx = ""
foreach ($cf in $ContextFile) { if ($cf) { $ctx = ($ctx + "`n`n" + (Read-SourceText -File $cf).Trim()).Trim() } }
if ($Context)     { $ctx = ($ctx + "`n" + $Context).Trim() }
if ($ctx) {
    $systemPrompt += "`n`nCONTEXTE DE LA NOUVELLE (fourni par l'auteur, a respecter imperativement) :`n" +
                     "---`n$ctx`n---`n" +
                     "Ce contexte prime sur tes suppositions : respecte les genres, les noms, " +
                     "les niveaux de langue et la terminologie qui y sont indiques."
}

# ---------------------------------------------------------------- decoupage
# On garde les blocs sous forme de listes de paragraphes pour pouvoir les compter.
$paragraphs = Get-Paragraphs -Text $raw
$chunkParas = Group-ParagraphsIntoChunks -Paragraphs $paragraphs -MaxWords $MaxWords -MaxParagraphs $MaxParagraphs

Write-Host "Source  : $Path" -ForegroundColor Cyan
Write-Host "Sortie  : $Out" -ForegroundColor Cyan
Write-Host ("{0} mots, {1} paragraphes -> {2} bloc(s) de {3} mots max" -f `
    (Measure-Words $raw), $paragraphs.Count, $chunkParas.Count, $MaxWords) -ForegroundColor Cyan
if ($ctx) { Write-Host ("Contexte : {0} mots" -f (Measure-Words $ctx)) -ForegroundColor Cyan }
if ($Think) { Write-Host "Mode reflexion : actif (plus lent, plus fidele)" -ForegroundColor Cyan }

# Au-dela de quelques blocs, la queue de 60 mots ne suffit plus a tenir la
# coherence : on entretient une fiche de continuite reinjectee dans chaque bloc.
$useSummary = (-not $NoSummary) -and ($chunkParas.Count -ge $SummaryFrom)
if ($useSummary) {
    Write-Host ("Fiche de continuite : activee ({0} blocs)" -f $chunkParas.Count) -ForegroundColor Cyan
}
$sheet = ""
Write-Host ""

# ---------------------------------------------------------------- traduction
$results = @(); $tail = ""; $mismatches = @(); $truncated = 0; $maxPromptSeen = 0
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$genTokens = 0

for ($i = 0; $i -lt $chunkParas.Count; $i++) {
    $expected = $chunkParas[$i].Count
    $chunkText = $chunkParas[$i] -join "`n`n"
    Write-Host ("[{0}/{1}] {2} paragraphe(s)..." -f ($i + 1), $chunkParas.Count, $expected) -NoNewline

    $chunkHasMarks = $Marks -and ($chunkText -match $script:MarkRegex)

    $sys = $systemPrompt
    if ($chunkHasMarks) { $sys += $marksRule }
    if ($sheet) {
        $sys += "`n`nFICHE DE CONTINUITE (etat du recit et traductions deja retenues plus haut " +
                "dans le texte). Respecte imperativement les correspondances de TERMS et les " +
                "genres de CHARACTERS :`n---`n$sheet`n---"
    }
    if ($tail) {
        $sys += "`n`nFin de la traduction du passage precedent, pour assurer la continuite " +
                "(NE PAS la retraduire, NE PAS la repeter) :`n---`n$tail`n---"
    }

    # On garde la MEILLEURE tentative, pas la derniere : une relance peut rendre
    # un resultat bien pire que le precedent, et le conserver perdrait du texte.
    $attempt = 0; $txt = $null; $got = -1; $marksOk = $true
    $best = $null; $bestGot = -1; $bestMarks = $false; $bestScore = [int]::MaxValue
    while ($attempt -le $MaxRetries) {
        $sysTry = $sys
        if ($attempt -gt 0) {
            $why = @()
            if ($got -ne $expected) {
                $why += "tu as rendu $got paragraphes au lieu de $expected (texte omis ou fusionne)"
            }
            if (-not $marksOk) {
                $why += "tu n'as pas restitue les balises [[i]] / [[b]] a l'identique"
            }
            $sysTry += "`n`nATTENTION : ta tentative precedente etait invalide : " +
                       ($why -join ' ; ') + ". Recommence en rendant EXACTEMENT " +
                       "$expected paragraphes separes par une ligne vide, sans rien omettre."
        }

        $r = Invoke-LMStudioChat -System $sysTry -User $chunkText -Model $Model `
                -Endpoint $Endpoint -Temperature $Temperature -Think:$Think `
                -MaxTokens ([Math]::Max(1024, $MaxWords * 3))

        $genTokens += $r.Tokens
        $txt = $r.Text
        if ($r.Truncated) {
            $truncated++
            Write-Host " [COUPE par max_tokens]" -NoNewline -ForegroundColor Red
        }
        if ($r.PromptTokens -gt $maxPromptSeen) { $maxPromptSeen = $r.PromptTokens }
        # Filet deterministe : si la source n'a aucune balise, toute balise en
        # sortie est une invention du modele. On la retire, sans risque d'erreur.
        if ($Marks -and -not $chunkHasMarks) { $txt = Remove-Marks -Text $txt }

        $got = (Get-Paragraphs -Text $txt).Count
        $marksOk = (-not $Marks) -or (Test-MarksIntact -Before $chunkText -After $txt)
        if ($got -eq $expected -and $marksOk) { $best = $txt; $bestGot = $got; $bestMarks = $true; $bestScore = 0; break }

        # Ecart de paragraphes, plus une penalite forte si du texte manque.
        $score = [Math]::Abs($got - $expected) +
                 $(if ($marksOk) { 0 } else { 5 }) +
                 [Math]::Max(0, [int](((Measure-Words $chunkText) - (Measure-Words $txt)) / 20))
        if ($score -lt $bestScore) {
            $bestScore = $score; $best = $txt; $bestGot = $got; $bestMarks = $marksOk
        }
        $attempt++
    }
    if ($bestScore -gt 0) { $txt = $best; $got = $bestGot; $marksOk = $bestMarks }

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

    $results += $txt
    $words = ($txt -split '\s+' | Where-Object { $_ })
    $tail = ($words | Select-Object -Last 60) -join ' '

    # Inutile de mettre a jour la fiche apres le dernier bloc : personne ne la lira.
    if ($useSummary -and $i -lt ($chunkParas.Count - 1)) {
        Write-Host "        mise a jour de la fiche de continuite..." -NoNewline -ForegroundColor DarkGray
        $u = Update-ContinuitySheet -Current $sheet -SourceChunk $chunkText -TranslatedChunk $txt `
                -Model $Model -Endpoint $Endpoint
        $sheet = $u.Sheet
        $genTokens += $u.Tokens
        Write-Host (" {0} mots" -f (Measure-Words $sheet)) -ForegroundColor DarkGray
    }
}

$sw.Stop()

$final = $results -join "`n`n"

# L'auteur garde les tirets de dialogue a la francaise. Le modele met des
# guillemets malgre la consigne : on retablit le tiret d'apres la source, ce qui
# est exact puisque l'information ne vient pas d'une deduction.
$rep = Repair-DialogueDashes -SourceParagraphs $paragraphs -Text $final
$final = $rep.Text
if (-not $rep.Aligned) {
    Write-Warning "Alignement source/traduction impossible : tirets de dialogue non verifies."
}
Write-Utf8NoBom -Path $Out -Content $final

# La fiche est ecrite a cote : relisible, editable, reutilisable pour une suite.
if ($useSummary -and $sheet) {
    $sheetPath = [System.IO.Path]::ChangeExtension($Out, $null).TrimEnd('.') + "_continuite.md"
    Write-Utf8NoBom -Path $sheetPath -Content ("# Fiche de continuite`n`n- Source : $Path`n`n" + $sheet + "`n")
    Write-Host "Fiche de continuite : $sheetPath" -ForegroundColor DarkGray
}

$outParas = (Get-Paragraphs -Text $final).Count

Write-Host ""
Write-Host ("Termine en {0:N0} s - {1} tokens generes - {2:N1} tok/s" -f `
    $sw.Elapsed.TotalSeconds, $genTokens, ($genTokens / $sw.Elapsed.TotalSeconds)) -ForegroundColor Yellow
Write-Host ("Paragraphes : {0} en entree -> {1} en sortie" -f $paragraphs.Count, $outParas) -ForegroundColor Yellow
Write-Host ("Mots : {0} en entree -> {1} en sortie" -f (Measure-Words $raw), (Measure-Words $final)) -ForegroundColor Yellow
Write-Host ("Prompt le plus lourd : {0} tokens (contexte charge : verifier lms ps)" -f $maxPromptSeen) -ForegroundColor DarkGray
if ($truncated) {
    Write-Warning ("{0} reponse(s) COUPEE(S) par max_tokens : augmente -MaxWords/max_tokens ou reduis la taille des blocs" -f $truncated)
}

if ($mismatches.Count) {
    Write-Host ""
    Write-Warning "Blocs encore en ecart apres relance - a relire en priorite :"
    $mismatches | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
}

if ($rep.Fixed) {
    Write-Host ("Tirets de dialogue retablis : {0} paragraphe(s)" -f $rep.Fixed) -ForegroundColor DarkYellow
}
$outParas2 = Get-Paragraphs -Text $final
$stillQuoted = @($outParas2 | Where-Object { $_ -match '^\s*["\u201C]' }).Count
if ($stillQuoted) {
    Write-Warning ("{0} paragraphe(s) commencent encore par un guillemet alors que la source n'avait pas de tiret - a verifier" -f $stillQuoted)
}

Write-Host "Ecrit dans : $Out" -ForegroundColor Yellow

# ---------------------------------------------------------------- revision enchainee
if ($Review) {
    Write-Host ""
    Write-Host "=== Revision enchainee (mode reflexion) ===" -ForegroundColor Magenta
    $revArgs = @{
        Source      = $Path
        Translation = $Out
        Model       = $Model
        Endpoint    = $Endpoint
    }
    if ($Context)     { $revArgs.Context     = $Context }
    if ($ContextFile) { $revArgs.ContextFile = $ContextFile }
    if ($ReviewOut)   { $revArgs.Out         = $ReviewOut }
    & "$PSScriptRoot\review.ps1" @revArgs
}
Write-TokenSummary
Write-TokenLog -Task 'traduction' -Model $Model -Source $Path
