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
  # Mode reperage : ne modifie rien, produit la liste des fautes a verifier
  .\proofread.ps1 -Path "textes\mon-recit.docx" -ReportOnly

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
    [switch]$Think,
    [switch]$ReportOnly,
    [switch]$ByBlock,
    [int]$Passes = 2
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\lib-lmstudio.ps1"
if (-not $Model) { $Model = Get-DefaultModel }

# Demarre le serveur et charge le modele si besoin : un script lance seul
# ne doit pas echouer juste parce que LM Studio s'est arrete.
Initialize-LMStudioSession -Model $Model -Endpoint $Endpoint -ContextLength $ContextLength
Reset-TokenLedger -Label 'proofread'

$item = Get-Item $Path
if ($ReportOnly) { $Marks = $false }
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
if ($ReportOnly -and -not $ReportPath) {
    $ReportPath = Join-Path $item.DirectoryName ($item.BaseName + "_reperage.md")
}
if (-not $ReportPath) {
    $ReportPath = [System.IO.Path]::ChangeExtension($Out, $null).TrimEnd('.') + "_notes.md"
}

# La reflexion casse les taches ou le modele REPRODUIT un texte entier, elle
# aide celles ou il JUGE. En mode reperage il ne reproduit plus rien : on
# l'active par defaut, sauf si l'appelant a tranche lui-meme.
if ($ReportOnly -and -not $PSBoundParameters.ContainsKey('Think')) { $Think = $true }

$promptName = 'proofread'
if ($ReportOnly) { $promptName = 'proofread-signal' }
$systemPrompt = Get-PromptText -Name $promptName -PromptFile $PromptFile

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

# ------------------------------------------------------ mode reperage seul
# Le modele ne rend plus le texte, seulement une liste. Il ne peut donc rien
# casser : le garde-fou sur le nombre de paragraphes devient sans objet.
# Le prix a payer se deplace sur l'exactitude des reperes, traitee par
# Resolve-FindingAnchor : une citation absente du texte est ecartee.
if ($ReportOnly) {

    # Unites d'examen. Par defaut UN paragraphe a la fois, ses voisins fournis
    # en lecture seule : l'attention n'est plus diluee sur huit paragraphes, et
    # l'ancre ne peut plus glisser sur un voisin. Mesure sur un extrait connu :
    # par blocs il fallait 3 passes pour retrouver ce qu'une seule trouve ici.
    # -ByBlock revient au decoupage par blocs, plus rapide, moins couvrant.
    $units = @()
    if ($ByBlock) {
        $acc = 0
        foreach ($cp in $chunkParas) {
            $units += @{
                From = $acc; To = ($acc + $cp.Count - 1)
                Body = ($cp -join "`n`n"); Before = @(); After = @()
            }
            $acc += $cp.Count
        }
    }
    else {
        for ($p = 0; $p -lt $paragraphs.Count; $p++) {
            $before = @(); $w = 0
            for ($k = $p - 1; ($k -ge 0) -and ($k -ge $p - 3); $k--) {
                $w += Measure-Words $paragraphs[$k]
                if ($w -gt 200) { break }
                $before = @($paragraphs[$k]) + $before
            }
            $after = @()
            if (($p + 1) -lt $paragraphs.Count) { $after = @($paragraphs[$p + 1]) }
            $units += @{
                From = $p; To = $p
                Body = $paragraphs[$p]; Before = $before; After = $after
            }
        }
    }

    if (-not $ByBlock) {
        $systemPrompt += "`n`nON NE TE DONNE QU'UN SEUL PARAGRAPHE A EXAMINER, precede et " +
                         "suivi de son contexte.`nTu ne signales RIEN dans le contexte : il " +
                         "n'est la que pour que tu saches de quoi parle le passage, qui parle, " +
                         "et a quel temps le recit est mene.`nExamine le paragraphe mot par mot."
    }

    # Le genre du narrateur est l'information la plus rentable de la fiche :
    # sans elle le modele devine, et signale des accords corrects comme fautifs.
    if ($ctx) {
        $systemPrompt += "`n`nSERS-TOI DE LA FICHE DE CONTEXTE POUR VERIFIER :`n" +
            "- le GENRE du narrateur et de chaque personnage. Tous les participes " +
            "passes et adjectifs qui s'y rapportent doivent s'accorder avec lui.`n" +
            "- le TEMPS du recit. Un verbe qui en sort est a signaler en categorie temps.`n" +
            "- QUI PARLE dans chaque incise de dialogue, pour le genre du verbe.`n" +
            "Un nom propre ou un terme qui figure dans la fiche n'est jamais une faute. " +
            "Si la fiche ne dit rien sur un point, ne l'inventes pas : classe en doute."
    }

    Write-Host "Source   : $Path" -ForegroundColor Cyan
    Write-Host "Rapport  : $ReportPath" -ForegroundColor Cyan
    $grain = 'paragraphe par paragraphe'
    if ($ByBlock) { $grain = 'par blocs' }
    Write-Host ("{0} mots, {1} paragraphes -> {2} unite(s) {3}, {4} passe(s)" -f `
        (Measure-Words $raw), $paragraphs.Count, $units.Count, $grain, $Passes) -ForegroundColor Cyan
    if ($ctx) { Write-Host ("Contexte : {0} mots" -f (Measure-Words $ctx)) -ForegroundColor Cyan }
    Write-Host "Aucun mot du texte ne sera modifie." -ForegroundColor DarkGray
    Write-Host ""

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $genTokens = 0
    $found = @(); $orphans = @(); $seen = @{}

    for ($i = 0; $i -lt $units.Count; $i++) {
        $u = $units[$i]
        $from = $u.From; $to = $u.To

        if ($ByBlock) {
            $userMsg = $u.Body
        }
        else {
            $userMsg = ""
            if ($u.Before.Count) {
                $userMsg += "CONTEXTE QUI PRECEDE (ne rien signaler ici) :`n---`n" +
                            ($u.Before -join "`n`n") + "`n---`n`n"
            }
            $userMsg += "PARAGRAPHE A EXAMINER (signale uniquement ici) :`n---`n" +
                        $u.Body + "`n---"
            if ($u.After.Count) {
                $userMsg += "`n`nCONTEXTE QUI SUIT (ne rien signaler ici) :`n---`n" +
                            ($u.After -join "`n`n") + "`n---"
            }
        }

        Write-Host ("[{0}/{1}] paragraphe {2}..." -f ($i + 1), $units.Count, ($from + 1)) -NoNewline

        $kept = 0; $lost = 0
        for ($pass = 1; $pass -le $Passes; $pass++) {
            # Deux lectures identiques ne trouvent pas plus qu'une seule :
            # on decale la temperature pour que la passe suivante apporte
            # vraiment quelque chose.
            $temp = $Temperature + (0.15 * ($pass - 1))
            $r = Invoke-LMStudioChat -System $systemPrompt -User $userMsg -Model $Model `
                    -Endpoint $Endpoint -Temperature $temp -Think:$Think -MaxTokens 1536
            $genTokens += $r.Tokens

            foreach ($f in (ConvertFrom-FindingLines -Text $r.Text)) {
                if (-not $f.Fragment) {
                    $orphans += @{ Para = $from; Raw = $f.Raw; Why = 'ligne hors format' }
                    $lost++
                    continue
                }
                $key = (ConvertTo-MatchKey -Text ($f.Fragment + '|' + $f.Suggestion)).Key + "|$from"
                if ($seen.ContainsKey($key)) { continue }
                $seen[$key] = $true

                $a = Resolve-FindingAnchor -Fragment $f.Fragment -Paragraphs $paragraphs `
                        -HintFrom $from -HintTo $to
                if (-not $a.Found) {
                    $orphans += @{ Para = $from; Raw = $f.Raw; Why = 'citation absente du texte' }
                    $lost++
                    continue
                }
                # En mode paragraphe, une ancre qui tombe ailleurs vient du
                # contexte : c'est ce qu'on a explicitement interdit de signaler.
                if ((-not $ByBlock) -and ($a.Para -ne $from)) {
                    $orphans += @{ Para = $from; Raw = $f.Raw; Why = 'signale dans le contexte, pas dans le paragraphe' }
                    $lost++
                    continue
                }
                $found += @{
                    Para = $a.Para; Start = $a.Start; Length = $a.Length
                    Suggestion = $f.Suggestion; Reason = $f.Reason; Source = 'modele'
                }
                $kept++
            }
        }
        if ($lost) {
            Write-Host (" {0} retenu(s), {1} ecarte(s)" -f $kept, $lost) -ForegroundColor DarkYellow
        }
        elseif ($kept) {
            Write-Host (" {0} retenu(s)" -f $kept) -ForegroundColor Green
        }
        else {
            Write-Host " rien" -ForegroundColor DarkGray
        }
    }

    # Ce qui se detecte sans modele se detecte mieux sans modele.
    foreach ($m in (Find-MechanicalIssues -Paragraphs $paragraphs)) {
        $key = "MECA|$($m.Para)|$($m.Start)|$($m.Reason)"
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $found += @{
            Para = $m.Para; Start = $m.Start; Length = $m.Length
            Suggestion = $m.Suggestion; Reason = $m.Reason; Source = 'mecanique'
        }
    }

    $sw.Stop()
    $found = @($found | Sort-Object @{ Expression = { $_.Para } }, @{ Expression = { $_.Start } })

    # Deux passes signalent souvent la meme faute avec un fragment plus ou moins
    # large. On garde le repere le plus precis plutot que d'afficher deux fois
    # la meme chose : c'est du bruit sans information.
    $dedup = @()
    foreach ($f in $found) {
        $dup = $false
        for ($k = 0; $k -lt $dedup.Count; $k++) {
            $d = $dedup[$k]
            if ($d.Para -ne $f.Para) { continue }
            $overlap = [Math]::Min($d.Start + $d.Length, $f.Start + $f.Length) -
                       [Math]::Max($d.Start, $f.Start)
            if ($overlap -le 0) { continue }
            $shortest = [Math]::Min($d.Length, $f.Length)
            if ($overlap -lt ($shortest * 0.8)) { continue }
            $dup = $true
            # on conserve le fragment le plus court, donc le plus precis
            if ($f.Length -lt $d.Length) {
                $f.Reason = Join-Reasons -First $f.Reason -Second $d.Reason
                $dedup[$k] = $f
            }
            else {
                $d.Reason = Join-Reasons -First $d.Reason -Second $f.Reason
            }
            break
        }
        if (-not $dup) { $dedup += $f }
    }
    $merged = $found.Count - $dedup.Count
    $found = $dedup

    $typo = Measure-TypographyGaps -Text $raw

    $BT = [string][char]0x60
    $lines = @()
    $lines += "# Reperage de fautes - francais"
    $lines += ""
    $lines += "- Source : $Path"
    $lines += "- Date   : $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    $lines += "- $($paragraphs.Count) paragraphes, $Passes passe(s), $($found.Count) signalement(s) avec repere"
    $lines += ""
    $lines += "**Aucun mot de ton texte n'a ete modifie.** Chaque entree donne une chaine"
    $lines += "verifiee presente dans le fichier : colle-la dans Ctrl+F pour tomber dessus."
    $lines += "Les signalements sont donnes dans l'ordre du texte."
    $lines += ""
    $lines += "La colonne de droite est ce que le modele **propose**. Elle est souvent"
    $lines += "fausse meme quand le repere est bon : sers-toi du repere, juge la correction."
    $lines += ""

    $lastPara = -1
    foreach ($f in $found) {
        if ($f.Para -ne $lastPara) {
            $lines += "## Paragraphe $($f.Para + 1)"
            $lines += ""
            $lastPara = $f.Para
        }
        $p = $paragraphs[$f.Para]
        $search  = Get-UniqueSearchString -Paragraph $p -Start $f.Start -Length $f.Length -AllParagraphs $paragraphs
        $excerpt = Get-AnchorExcerpt -Paragraph $p -Start $f.Start -Length $f.Length

        $head = $f.Reason
        if ($f.Source -eq 'mecanique') { $head = $head + ' (detection mecanique)' }
        if ($f.Suggestion) { $head = $head + '  ->  ' + $f.Suggestion }

        $lines += "**Ctrl+F** $BT$search$BT"
        $lines += ""
        $lines += $head
        $lines += ""
        $lines += "> $excerpt"
        $lines += ""
    }

    if ($typo.StraightApostrophes -or $typo.MissingNbsp) {
        $lines += "## Typographie"
        $lines += ""
        $lines += "Comptee, pas listee : ces ecarts se comptent en masse et noieraient le"
        $lines += "reste. Le mode correction les repare mecaniquement."
        $lines += ""
        $lines += "- apostrophes droites au lieu de courbes : $($typo.StraightApostrophes)"
        $lines += "- ponctuations doubles sans espace insecable : $($typo.MissingNbsp)"
        $lines += ""
    }

    if ($orphans.Count) {
        $lines += "## Signalements ecartes ($($orphans.Count))"
        $lines += ""
        $lines += "Citation introuvable dans le paragraphe examine : impossible de dire ou le"
        $lines += "modele voulait en venir, donc impossible d'en faire un repere. Listes ici"
        $lines += "plutot que jetes, mais a lire avec mefiance."
        $lines += ""
        foreach ($o in $orphans) { $lines += "- paragraphe $($o.Para + 1) - $($o.Why) : $($o.Raw)" }
        $lines += ""
    }

    Write-Utf8NoBom -Path $ReportPath -Content ($lines -join "`n")

    Write-Host ""
    Write-Host ("Termine en {0:N0} s - {1} tokens" -f $sw.Elapsed.TotalSeconds, $genTokens) -ForegroundColor Yellow
    $meca = @($found | Where-Object { $_.Source -eq 'mecanique' }).Count
    Write-Host ("Signalements avec repere : {0} dont {1} mecanique(s)" -f $found.Count, $meca) -ForegroundColor Yellow
    if ($merged) { Write-Host ("Doublons fusionnes : {0}" -f $merged) -ForegroundColor DarkGray }
    if ($orphans.Count) {
        Write-Host ("Ecartes faute de repere : {0}" -f $orphans.Count) -ForegroundColor DarkYellow
    }
    Write-Host "Rapport : $ReportPath" -ForegroundColor Yellow
    Write-TokenSummary
    Write-TokenLog -Task 'reperage' -Model $Model -Source $Path `
        -Note "$Passes passe(s), $($found.Count) signalement(s)"
    return
}


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
Write-TokenSummary
Write-TokenLog -Task 'correction' -Model $Model -Source $Path
