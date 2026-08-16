<#
.SYNOPSIS
  Aller-retour DeepL gratuit : decoupe pour le copier-coller, puis retablit le
  gras et l'italique perdus au passage.

.DESCRIPTION
  DeepL gratuit s'utilise par copier-coller et perd toute mise en forme.
  Ce script encadre l'operation en deux temps.

  MODE EXPORT
    Lit le .docx francais et ecrit des fichiers texte decoupes sous la limite
    de caracteres de DeepL, aux frontieres de paragraphes. Les balises de mise
    en forme sont retirees : DeepL ne doit voir que du texte propre.

  MODE RESTORE
    Reprend l'original .docx (qui porte la mise en forme) et la traduction
    collee depuis DeepL, aligne les paragraphes, et replace le gras et
    l'italique :
      - paragraphe entierement formate  -> report direct, exact
      - portion de paragraphe           -> le modele local situe le passage

    Garde-fou : apres placement des balises, le texte anglais debarrasse de ses
    balises doit etre IDENTIQUE au caractere pres a ce que DeepL avait rendu.
    Sinon le paragraphe est laisse sans formatage. Le script ne peut donc pas
    alterer ta traduction.

  Sortie : un .docx avec la mise en forme, et le .txt correspondant.

.EXAMPLE
  .\deepl-roundtrip.ps1 export -Source "textes\mon-recit.docx"

.EXAMPLE
  .\deepl-roundtrip.ps1 restore -Source "mon-recit.docx" -Translation part_1_EN.txt,part_2_EN.txt
#>

param(
    [Parameter(Mandatory = $true, Position = 0)][ValidateSet('export', 'restore')][string]$Mode,
    [Parameter(Mandatory = $true)][string]$Source,
    [string[]]$Translation,
    [string]$OutDir,
    [string]$Out,
    [int]$MaxChars = 4500,
    [string]$Model,
    [string]$Endpoint = "http://localhost:1234/v1/chat/completions",
    [int]$ContextLength = 20000
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\lib-lmstudio.ps1"
if (-not $Model) { $Model = Get-DefaultModel }

$srcItem = Get-Item $Source
$name = $srcItem.BaseName
if (-not $OutDir) { $OutDir = $srcItem.DirectoryName }
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }

$marked = Read-DocxMarkedText -File $srcItem.FullName
$srcParas = Get-Paragraphs -Text $marked

# ==================================================================== EXPORT
if ($Mode -eq 'export') {
    $files = @(); $cur = @(); $n = 1
    foreach ($p in $srcParas) {
        $clean = Remove-Marks -Text $p
        # Longueur reelle du fichier tel qu'il sera ecrit, separateurs CRLFCRLF
        # compris : une estimation ferait depasser la limite de DeepL.
        $candidate = (($cur + $clean) -join "`r`n`r`n")
        if ($cur.Count -gt 0 -and $candidate.Length -gt $MaxChars) {
            $f = Join-Path $OutDir ("{0}_deepl_{1}.txt" -f $name, $n)
            Write-Utf8NoBom -Path $f -Content ($cur -join "`r`n`r`n")
            $files += $f; $cur = @(); $n++
        }
        $cur += $clean
    }
    if ($cur.Count) {
        $f = Join-Path $OutDir ("{0}_deepl_{1}.txt" -f $name, $n)
        Write-Utf8NoBom -Path $f -Content ($cur -join "`r`n`r`n")
        $files += $f
    }

    Write-Host ("{0} : {1} paragraphes, {2:N0} caracteres" -f $name, $srcParas.Count, (Remove-Marks -Text $marked).Length) -ForegroundColor Cyan
    Write-Host ("-> {0} fichier(s) de {1} caracteres max" -f $files.Count, $MaxChars) -ForegroundColor Cyan
    Write-Host ""
    $files | ForEach-Object { Write-Host ("  {0}  ({1:N0} car.)" -f $_, (Get-Item $_).Length) -ForegroundColor Green }
    Write-Host ""
    Write-Host "Colle chaque fichier dans DeepL, puis enregistre la traduction sous" -ForegroundColor Yellow
    Write-Host "le meme nom suffixe _EN.txt, en gardant les lignes vides entre paragraphes." -ForegroundColor Yellow
    Write-Host ("Ensuite :  .\deepl-roundtrip.ps1 restore -Source `"{0}`" -Translation <fichiers _EN.txt>" -f $srcItem.FullName) -ForegroundColor Yellow
    return
}

# =================================================================== RESTORE
if (-not $Translation) { throw "-Translation est requis en mode restore." }

$enText = ""
foreach ($t in $Translation) {
    if (-not (Test-Path $t)) { throw "Fichier introuvable : $t" }
    $enText = ($enText + "`n`n" + (Read-SourceText -File $t).Trim()).Trim()
}
$enParas = Get-Paragraphs -Text $enText

Write-Host ("Original   : {0} paragraphes" -f $srcParas.Count) -ForegroundColor Cyan
Write-Host ("Traduction : {0} paragraphes" -f $enParas.Count) -ForegroundColor Cyan

if ($srcParas.Count -ne $enParas.Count) {
    throw ("Alignement impossible : {0} paragraphes en francais contre {1} en anglais.`n" -f $srcParas.Count, $enParas.Count) +
          "DeepL a fusionne ou coupe des paragraphes. Verifie que les lignes vides ont ete conservees au collage."
}

$formatted = @($srcParas | Where-Object { $_ -match $script:MarkRegex }).Count
Write-Host ("Paragraphes portant une mise en forme : {0}" -f $formatted) -ForegroundColor Cyan
Write-Host ""

# On ne fait PAS reecrire le paragraphe au modele : il perdait un guillemet en
# recopiant, et tout etait rejete. On lui demande seulement d'identifier le
# passage anglais correspondant ; les balises sont ensuite inserees par le script
# dans la traduction d'origine, qui n'est donc jamais regeneree.
$sys = @'
On te donne un passage FRANCAIS et sa TRADUCTION anglaise complete.
Dans le passage francais, une portion est encadree par >>> et <<<.

Reponds UNIQUEMENT par la portion de la TRADUCTION ANGLAISE qui correspond a
cette portion francaise.

Recopie-la EXACTEMENT telle qu'elle apparait dans la traduction : memes mots,
meme casse, meme ponctuation. N'ajoute ni guillemets, ni explication, ni
ponctuation absente. Ne renvoie que ces quelques mots, rien d'autre.
'@

function Add-MarksToTranslation {
    <# Fait localiser chaque portion balisee par le modele, puis insere les
       balises dans la traduction d'origine. Retourne $null si une portion n'a
       pas pu etre situee : mieux vaut pas de mise en forme qu'une mauvaise. #>
    param([string]$FrMarked, [string]$EnPlain)

    $result = $EnPlain
    foreach ($m in [regex]::Matches($FrMarked, '\[\[([ib])\]\](.*?)\[\[/\1\]\]')) {
        $tag = $m.Groups[1].Value
        $frSpan = $m.Groups[2].Value.Trim()
        if (-not $frSpan) { continue }

        $frHighlighted = $FrMarked -replace '\[\[/?[ib]\]\]', ''
        $frHighlighted = $frHighlighted -replace [regex]::Escape($frSpan), (">>>$frSpan<<<")

        $r = Invoke-LMStudioChat -System $sys `
                -User ("=== FRANCAIS ===`n$frHighlighted`n`n=== TRADUCTION ANGLAISE ===`n$EnPlain") `
                -Model $Model -Endpoint $Endpoint -Temperature 0.1 -MaxTokens 256
        $span = (Remove-TemplateEcho -Text $r.Text).Text.Trim().Trim('"').Trim()
        if (-not $span) { return $null }

        $idx = $result.IndexOf($span, [StringComparison]::Ordinal)
        if ($idx -lt 0) { $idx = $result.IndexOf($span, [StringComparison]::OrdinalIgnoreCase) }
        if ($idx -lt 0) { return $null }

        $result = $result.Substring(0, $idx) + "[[$tag]]" + $result.Substring($idx, $span.Length) +
                  "[[/$tag]]" + $result.Substring($idx + $span.Length)
    }
    $result
}

$outParas = @(); $direct = 0; $viaModel = 0; $failed = 0; $needModel = $false
$sw = [System.Diagnostics.Stopwatch]::StartNew()

# Le modele n'est demarre que si au moins un paragraphe en a besoin.
foreach ($i in 0..($srcParas.Count - 1)) {
    if ($srcParas[$i] -match $script:MarkRegex) {
        $stripped = (Remove-Marks -Text $srcParas[$i]).Trim()
        $m = [regex]::Match($srcParas[$i], '^\s*\[\[([ib])\]\](.*?)\[\[/\1\]\]\s*$')
        if (-not ($m.Success -and $m.Groups[2].Value.Trim() -eq $stripped)) { $needModel = $true; break }
    }
}
if ($needModel) { Initialize-LMStudioSession -Model $Model -Endpoint $Endpoint -ContextLength $ContextLength }

for ($i = 0; $i -lt $srcParas.Count; $i++) {
    $src = $srcParas[$i]
    $en = $enParas[$i]

    if ($src -notmatch $script:MarkRegex) { $outParas += $en; continue }

    # Cas 1 : le paragraphe entier est formate -> report exact, sans modele.
    $stripped = (Remove-Marks -Text $src).Trim()
    $m = [regex]::Match($src, '^\s*\[\[([ib])\]\](.*?)\[\[/\1\]\]\s*$')
    if ($m.Success -and $m.Groups[2].Value.Trim() -eq $stripped) {
        $tag = $m.Groups[1].Value
        $outParas += "[[$tag]]$en[[/$tag]]"
        $direct++
        continue
    }

    # Cas 2 : portion de paragraphe -> le modele situe le passage.
    Write-Host ("  paragraphe {0} : placement des balises..." -f ($i + 1)) -NoNewline -ForegroundColor DarkGray
    try { $cand = Add-MarksToTranslation -FrMarked $src -EnPlain $en }
    catch { $cand = $null }

    # Le texte n'a pas ete regenere : il suffit de verifier qu'aucun caractere
    # n'a bouge une fois les balises retirees.
    if ($cand -and (Remove-Marks -Text $cand) -eq $en) {
        $outParas += $cand
        $viaModel++
        Write-Host " ok" -ForegroundColor Green
    }
    else {
        $outParas += $en
        $failed++
        Write-Host " passage non situe - paragraphe laisse sans mise en forme" -ForegroundColor Yellow
    }
}

$sw.Stop()
$final = $outParas -join "`n`n"

if (-not $Out) { $Out = Join-Path $OutDir ($name + "_EN_format.docx") }
$outTxt = [System.IO.Path]::ChangeExtension($Out, '.txt')

Write-DocxFromText -Path $Out -Text $final
Write-Utf8NoBom -Path $outTxt -Content (Remove-Marks -Text $final)

Write-Host ""
Write-Host ("Termine en {0:N0} s" -f $sw.Elapsed.TotalSeconds) -ForegroundColor Yellow
Write-Host ("Mise en forme reportee : {0} paragraphe(s) entier(s), {1} portion(s)" -f $direct, $viaModel) -ForegroundColor Yellow
if ($failed) {
    Write-Host ("{0} paragraphe(s) laisse(s) sans mise en forme (texte qui aurait ete altere)" -f $failed) -ForegroundColor Yellow
}
Write-Host ("Mots : {0} (traduction inchangee)" -f (Measure-Words (Remove-Marks -Text $final))) -ForegroundColor Yellow
Write-Host "Avec mise en forme : $Out" -ForegroundColor Yellow
Write-Host "Texte brut         : $outTxt" -ForegroundColor Yellow
