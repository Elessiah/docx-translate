<#
.SYNOPSIS
  Prepare un texte pour DeviantArt : une ligne vide entre chaque paragraphe.

.DESCRIPTION
  Lit des .docx (et aussi .txt / .md), extrait les paragraphes, et les reecrit
  separes par une ligne vide. C'est ce qui aere le texte dans l'editeur de
  DeviantArt, qui colle les paragraphes les uns aux autres sans cette respiration.

  Mise en forme du texte :
    - un paragraphe Word = un paragraphe de sortie
    - les sauts de ligne manuels (Maj+Entree) deviennent de vrais paragraphes
    - les paragraphes vides du .docx sont supprimes (sinon on obtiendrait
      deux ou trois lignes vides d'affilee)
    - espaces insecables, tabulations et espaces multiples normalises (-NoCleanup)
    - les revisions supprimees (suivi des modifications) sont ignorees
    - les separateurs de scene (***, ---, ~~~...) sont uniformises et centres

  Nettoyage typographique, optionnel :
    -English       supprime l'espace avant ? ! : ; et convertit les guillemets
                   francais : residus courants d'une traduction FR->EN
    -SmartQuotes   guillemets et apostrophes courbes, ... -> ellipse, -- -> tiret cadratin

  Controle qualite avant publication : signale les guillemets non refermes et les
  tirets cadratins de dialogue francais qui auraient survecu a la traduction.

  Publication :
    -Header / -Footer  bloc colle avant / apres le texte (fichier ou texte libre)
    -SplitWords        decoupe en plusieurs deviations, aux ruptures de scene
    -Sheet             fiche .md : mots, temps de lecture, resume et tags
                       generes par le modele local

  Deux formats de sortie :
    txt   texte brut, a coller dans l'editeur DeviantArt. L'italique est perdu.
    html  page autonome a ouvrir dans le navigateur : Ctrl+A / Ctrl+C puis coller
          dans l'editeur DeviantArt CONSERVE l'italique et le gras du .docx.

  Traite un fichier, plusieurs fichiers, un dossier ou un motif joker.

.EXAMPLE
  .\format-deviantart.ps1 -Path "textes\mon-recit_EN_revised.txt" -English -SmartQuotes -Clipboard
  .\format-deviantart.ps1 -Path "textes\mon-recit.docx" -Format html -English -SmartQuotes -Sheet
  .\format-deviantart.ps1 -Path "SS" -Format both -SplitWords 4000 -Footer pied.txt -Force
#>

param(
    [Parameter(Mandatory = $true, Position = 0)][string[]]$Path,
    [string]$OutDir,
    [string]$Suffix = "_DA",
    [ValidateSet('txt', 'html', 'both')][string]$Format = 'txt',
    [ValidateRange(1, 5)][int]$BlankLines = 1,
    [switch]$English,
    [switch]$French,
    [switch]$KeepDashes,
    [switch]$SmartQuotes,
    [string]$SceneBreak = "* * *",
    [string]$Header,
    [string]$Footer,
    [int]$SplitWords = 0,
    [switch]$Sheet,
    [int]$SheetWords = 1500,
    [string]$Model,
    [string]$Endpoint = "http://localhost:1234/v1/chat/completions",
    [switch]$KeepEmpty,
    [switch]$NoCleanup,
    [switch]$Recurse,
    [switch]$Clipboard,
    [switch]$Force
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\lib-lmstudio.ps1"   # Measure-Words, Write-Utf8NoBom, Get-UnbalancedQuoteParagraphs, LM Studio
if (-not $Model) { $Model = Get-DefaultModel }

$W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# Ce fichier reste volontairement en pur ASCII : Windows PowerShell 5.1 lit les
# .ps1 sans BOM comme de l'ANSI, un caractere typographique ecrit en clair y
# serait corrompu en silence. D'ou les [char] ci-dessous.
$ELLIPSIS = [char]0x2026   # ...
$EMDASH   = [char]0x2014   # tiret cadratin
$ENDASH   = [char]0x2013   # tiret demi-cadratin
$LDQUO    = [char]0x201C   # guillemet anglais ouvrant
$RDQUO    = [char]0x201D   # guillemet anglais fermant
$RSQUO    = [char]0x2019   # apostrophe courbe
$LAQUO    = [char]0x00AB   # guillemet francais ouvrant
$RAQUO    = [char]0x00BB   # guillemet francais fermant
$BULLET   = [char]0x2022
$MIDDOT   = [char]0x00B7
$ASTERISM = [char]0x2042
$NBSP     = [char]0x00A0   # espace insecable

if ($English -and $French) {
    throw "-English et -French s'excluent : le premier convertit la typographie " +
          "francaise vers l'anglais, le second la preserve."
}

# ---------------------------------------------------------------- lecture docx

function Test-RunFlag {
    <# Vrai si l'attribut de mise en forme (i, b) est actif sur ce run. #>
    param($Rpr, [string]$Name, $Ns)
    if (-not $Rpr) { return $false }
    $node = $Rpr.SelectSingleNode("w:$Name", $Ns)
    if (-not $node) { return $false }
    $val = $node.GetAttribute('val', $script:W)
    # <w:i/> sans attribut = actif ; <w:i w:val="0"/> = desactive
    return ($val -eq '' -or $val -notin @('0', 'false', 'off'))
}

function New-Run {
    param([string]$Text, [bool]$Italic = $false, [bool]$Bold = $false)
    [pscustomobject]@{ Text = $Text; Italic = $Italic; Bold = $Bold }
}

function New-Para {
    param($Runs, [bool]$Center = $false, [bool]$IsBreak = $false)
    [pscustomobject]@{
        Runs    = @($Runs)
        Center  = $Center
        IsBreak = $IsBreak
        Text    = (@($Runs) | ForEach-Object { $_.Text }) -join ''
    }
}

function Read-DocxRichParagraphs {
    <# Retourne des paragraphes sous forme de listes de runs { Text; Italic; Bold }.
       Les <w:br> coupent le paragraphe en deux. #>
    param([Parameter(Mandatory)][string]$File)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($File)
    try {
        # certains outils zippent avec des antislashs, Word non : on tolere les deux
        $entry = $zip.Entries | Where-Object { ($_.FullName -replace '\\', '/') -eq 'word/document.xml' }
        if (-not $entry) { throw "Ce .docx ne contient pas word/document.xml (fichier corrompu ?)" }
        $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        $xmlText = $reader.ReadToEnd()
        $reader.Close()
        $bullets = Get-DocxBullets -Zip $zip
    }
    finally { $zip.Dispose() }

    $xml = [xml]$xmlText
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('w', $W)

    # w:t / w:tab / w:br pris uniquement comme enfants directs d'un run : evite les
    # taquets de tabulation declares dans w:pPr/w:tabs, qui ne sont pas du texte.
    $textNodes = './/w:r[not(ancestor::w:del)]/*[self::w:t or self::w:tab or self::w:br]'

    $paragraphs = @()
    foreach ($p in $xml.SelectNodes('//w:body//w:p', $ns)) {
        $jc = $p.SelectSingleNode('w:pPr/w:jc', $ns)
        $centered = ($jc -and $jc.GetAttribute('val', $W) -eq 'center')

        $runs = @()
        # Puce de liste = marqueur de dialogue de l'auteur, absent du texte.
        $numId = $p.SelectSingleNode('w:pPr/w:numPr/w:numId', $ns)
        if ($numId) {
            $id = $numId.GetAttribute('val', $W)
            $bullet = if ($bullets -and $bullets.ContainsKey($id)) { $bullets[$id] } else { '-' }
            if ($bullet) { $runs += (New-Run -Text ($bullet + ' ')) }
        }
        foreach ($n in $p.SelectNodes($textNodes, $ns)) {
            if ($n.LocalName -eq 'br') {
                # saut de ligne manuel : on ferme le paragraphe courant
                $paragraphs += (New-Para -Runs $runs -Center $centered)
                $runs = @()
                continue
            }
            $piece = if ($n.LocalName -eq 'tab') { "`t" } else { $n.InnerText }
            $rPr = $n.ParentNode.SelectSingleNode('w:rPr', $ns)
            $runs += (New-Run -Text $piece `
                              -Italic (Test-RunFlag $rPr 'i' $ns) `
                              -Bold   (Test-RunFlag $rPr 'b' $ns))
        }
        $paragraphs += (New-Para -Runs $runs -Center $centered)
    }
    , $paragraphs
}

function Read-PlainParagraphs {
    <# .txt / .md : lignes vides comme separateur si le fichier en contient,
       sinon une ligne = un paragraphe. #>
    param([Parameter(Mandatory)][string]$File)

    $text = [System.IO.File]::ReadAllText($File, [System.Text.Encoding]::UTF8)
    if ($text -match '\r?\n[ \t]*\r?\n') { $parts = [regex]::Split($text, '\r?\n[ \t]*\r?\n') }
    else { $parts = [regex]::Split($text, '\r?\n') }

    , @($parts | ForEach-Object { New-Para -Runs @(New-Run -Text ($_ -replace '\r?\n', ' ')) })
}

# ---------------------------------------------------------------- nettoyage

function Format-RunText {
    param([string]$Text)
    if ($NoCleanup) { return $Text }
    # En francais l'espace insecable porte du sens : avant ? ! ; : et dans les
    # guillemets. On la normalise en U+00A0 au lieu de l'aplatir comme en anglais,
    # ou elle n'est qu'un residu de mise en page.
    if ($French) { $t = $Text -replace '[\u202F\u2007\u2009]', $NBSP }
    else         { $t = $Text -replace '[\u00A0\u202F\u2007\u2009]', ' ' }
    $t = $t -replace "\t", ' '
    $t -replace ' {2,}', ' '
}

function Resolve-Paragraphs {
    <# Nettoie les runs, calcule le texte plein, jette les paragraphes vides. #>
    param($Paragraphs)

    $out = @()
    foreach ($p in $Paragraphs) {
        $runs = @()
        foreach ($r in $p.Runs) {
            $runs += (New-Run -Text (Format-RunText $r.Text) -Italic $r.Italic -Bold $r.Bold)
        }
        # Les blancs de bord relevent de la mise en page, pas du texte. Ils peuvent
        # s'etaler sur plusieurs runs (tabulation d'alinea + espaces), d'ou les boucles.
        while ($runs.Count) {
            $runs[0].Text = $runs[0].Text.TrimStart()
            if ($runs[0].Text -ne '') { break }
            $runs = @($runs | Select-Object -Skip 1)
        }
        while ($runs.Count) {
            $runs[-1].Text = $runs[-1].Text.TrimEnd()
            if ($runs[-1].Text -ne '') { break }
            $runs = @($runs | Select-Object -SkipLast 1)
        }
        $runs = @($runs | Where-Object { $_.Text -ne '' })
        $out += (New-Para -Runs $runs -Center $p.Center)
    }
    , $out
}

# ---------------------------------------------------------------- separateurs de scene

function Test-SceneBreak {
    <# Un paragraphe fait uniquement de caracteres de separation : *** , * * * ,
       --- , ~~~ , ... au moins trois signes. #>
    param([string]$Text)
    $s = $Text -replace '\s', ''
    if ($s.Length -lt 3) { return $false }
    $class = '^[\*\-_=~#' + $EMDASH + $ENDASH + $BULLET + $MIDDOT + $ASTERISM + ']+$'
    return ($s -match $class)
}

function Set-SceneBreaks {
    <# Marque et uniformise les ruptures de scene. -SceneBreak "" les laisse telles quelles. #>
    param($Paragraphs)
    $out = @()
    foreach ($p in $Paragraphs) {
        if (Test-SceneBreak $p.Text) {
            $marker = if ($SceneBreak -ne '') { $SceneBreak } else { $p.Text }
            $out += (New-Para -Runs @(New-Run -Text $marker) -Center $true -IsBreak $true)
        }
        else { $out += $p }
    }
    , $out
}

# ---------------------------------------------------------------- typographie

function Convert-Punctuation {
    <# Residus de ponctuation francaise dans un texte destine a etre lu en anglais. #>
    param([string]$Text)
    $t = $Text
    $t = $t -replace ('\s+([?!:;' + $RAQUO + '])'), '$1'     # espace avant ponctuation double
    $t = $t -replace ($LAQUO + '\s+'), $LAQUO                # espace apres le chevron ouvrant
    $t = $t -replace ($LAQUO + '|' + $RAQUO), '"'            # chevrons -> guillemets droits
    $t = $t -replace ('\s+' + $EMDASH + '\s+'), ($EMDASH)    # tiret cadratin sans espaces
    $t
}

function Convert-RunsTypography {
    <# Transformations locales puis, en un seul passage, les guillemets courbes :
       leur sens depend du caractere precedent, qui peut appartenir au run d'avant. #>
    param($Runs)

    $out = @()
    foreach ($r in $Runs) {
        $t = $r.Text
        if ($English)     { $t = Convert-Punctuation $t }
        if ($SmartQuotes) {
            $t = $t -replace '\.\.\.', $ELLIPSIS
            $t = $t -replace '\s*-{2,3}\s*', $EMDASH
        }
        $out += (New-Run -Text $t -Italic $r.Italic -Bold $r.Bold)
    }

    if (-not $SmartQuotes) { return , $out }

    $prev = ''
    foreach ($r in $out) {
        $sb = New-Object System.Text.StringBuilder
        foreach ($c in $r.Text.ToCharArray()) {
            $piece = [string]$c
            if ($c -eq '"') {
                # ouvrant en debut de paragraphe, apres un blanc, une parenthese ou un tiret
                $isOpen = ($prev -eq '' -or $prev -match '[\s\(\[\{]' -or $prev -eq $EMDASH -or $prev -eq $ENDASH)
                if ($French) {
                    # chevrons avec l'espace insecable qu'impose la typographie francaise
                    $piece = if ($isOpen) { [string]$LAQUO + $NBSP } else { [string]$NBSP + $RAQUO }
                }
                else {
                    $piece = if ($isOpen) { [string]$LDQUO } else { [string]$RDQUO }
                }
            }
            elseif ($c -eq "'") {
                # toujours l'apostrophe : dans ce corpus (don't, 'em, Amelie's) elle est
                # bien plus frequente que le guillemet simple imbrique
                $piece = [string]$RSQUO
            }
            [void]$sb.Append($piece)
            $prev = $piece[$piece.Length - 1]
        }
        $r.Text = $sb.ToString()
    }
    , $out
}

# ---------------------------------------------------------------- controle qualite

function Get-DashDialogueParagraphs {
    <# Numeros des paragraphes ouvrant sur un tiret de dialogue francais : la
       traduction aurait du les convertir en guillemets anglais. #>
    param($Paragraphs)
    $bad = @()
    for ($i = 0; $i -lt $Paragraphs.Count; $i++) {
        if ($Paragraphs[$i].IsBreak) { continue }
        if ($Paragraphs[$i].Text -match ('^[' + $EMDASH + $ENDASH + '\-]\s')) { $bad += ($i + 1) }
    }
    , $bad
}

function Write-NumberList {
    <# "1, 4, 9 et 12 autres" : evite de deverser 200 numeros dans la console. #>
    param([int[]]$Numbers, [int]$Max = 12)
    if ($Numbers.Count -le $Max) { return ($Numbers -join ', ') }
    (($Numbers | Select-Object -First $Max) -join ', ') + (" et {0} autre(s)" -f ($Numbers.Count - $Max))
}

# ---------------------------------------------------------------- entete / pied

function Resolve-Block {
    <# -Header / -Footer accepte un chemin de fichier ou du texte libre. #>
    param([string]$Value)
    if (-not $Value) { return @() }
    $text = if (Test-Path -LiteralPath $Value -PathType Leaf) {
        [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $Value), [System.Text.Encoding]::UTF8)
    } else { $Value }

    , @([regex]::Split($text.Trim(), '\r?\n[ \t]*\r?\n') |
        Where-Object { $_.Trim() } |
        ForEach-Object { New-Para -Runs @(New-Run -Text ($_.Trim() -replace '\r?\n', ' ')) })
}

function Expand-Placeholders {
    <# {title} {part} {parts} dans l'entete et le pied. #>
    param($Block, [string]$Title, [int]$Part, [int]$Parts)
    , @($Block | ForEach-Object {
        $t = $_.Text -replace '\{title\}', $Title -replace '\{part\}', $Part -replace '\{parts\}', $Parts
        New-Para -Runs @(New-Run -Text $t) -Center $_.Center
    })
}

# ---------------------------------------------------------------- decoupage

function Split-IntoParts {
    <# Coupe en parties sous un budget de mots, de preference a la derniere rupture
       de scene rencontree : le separateur devient la frontiere et disparait. #>
    param($Paragraphs, [int]$MaxWords)

    if ($MaxWords -le 0) { return , @(, $Paragraphs) }

    $parts = @(); $cur = @(); $count = 0; $breakIdx = -1
    foreach ($p in $Paragraphs) {
        $w = Measure-Words $p.Text
        if ($count -gt 0 -and ($count + $w) -gt $MaxWords) {
            if ($breakIdx -gt 0) {
                $head = @($cur[0..($breakIdx - 1)])
                # NB : un @() renvoye par un bloc if ne produit aucune sortie et
                # laisserait $tail a $null. D'ou l'affectation en deux temps.
                $tail = @()
                if ($breakIdx -lt $cur.Count - 1) { $tail = @($cur[($breakIdx + 1)..($cur.Count - 1)]) }
                $parts += , $head
                $cur = @($tail)
                $sum = ($cur | ForEach-Object { Measure-Words $_.Text } | Measure-Object -Sum).Sum
                $count = if ($sum) { $sum } else { 0 }
            }
            else {
                $parts += , @($cur)
                $cur = @(); $count = 0
            }
            $breakIdx = -1
        }
        $cur += $p
        if ($p.IsBreak) { $breakIdx = $cur.Count - 1 }
        $count += $w
    }
    if ($cur.Count) { $parts += , @($cur) }
    , $parts
}

# ---------------------------------------------------------------- sorties

function ConvertTo-DaText {
    param($Paragraphs, [int]$Blank)
    # CRLF : le fichier est destine a etre ouvert puis copie sous Windows
    $sep = "`r`n" * (1 + $Blank)
    ($Paragraphs | ForEach-Object { $_.Text }) -join $sep
}

function ConvertTo-DaHtml {
    param($Paragraphs, [string]$Title)

    $body = foreach ($p in $Paragraphs) {
        $inner = foreach ($r in $p.Runs) {
            $s = [System.Net.WebUtility]::HtmlEncode($r.Text)
            if ($r.Italic) { $s = "<em>$s</em>" }
            if ($r.Bold) { $s = "<strong>$s</strong>" }
            $s
        }
        $cls = if ($p.Center) { ' class="c"' } else { '' }
        "<p$cls>" + ($inner -join '') + "</p>"
    }

    @"
<!doctype html>
<meta charset="utf-8">
<title>$([System.Net.WebUtility]::HtmlEncode($Title))</title>
<style>
  body { max-width: 42em; margin: 3em auto; padding: 0 1.5em;
         font: 1.05em/1.7 Georgia, "Times New Roman", serif; }
  p { margin: 0 0 1.2em; }
  p.c { text-align: center; }
</style>
<!-- Ctrl+A puis Ctrl+C ici, puis coller dans l'editeur DeviantArt :
     l'italique et le gras sont conserves. -->
$($body -join "`r`n")
"@
}

function Get-CleanTitle {
    <# "MonRecit_EN_revised" -> "MonRecit" #>
    param([string]$BaseName)
    $t = $BaseName -replace '(_EN|_revised|_final|_DA)+$', ''
    ($t -replace '[_]+', ' ').Trim()
}

function New-PublicationSheet {
    <# Fiche a coller dans la description de la deviation. Le resume et les tags
       viennent du modele local ; sans serveur, la fiche est ecrite sans eux. #>
    param(
        [string]$Title, $Paragraphs, [int]$Words, [int]$Parts, [string]$OutPath
    )

    $minutes = [Math]::Max(1, [Math]::Ceiling($Words / 200))
    $lines = @(
        "# $Title"
        ""
        "- Mots : {0:N0}" -f $Words
        "- Temps de lecture : ~$minutes min"
        "- Paragraphes : $($Paragraphs.Count)"
    )
    if ($Parts -gt 1) { $lines += "- Publie en $Parts parties" }
    $lines += ""

    $excerpt = (($Paragraphs | Where-Object { -not $_.IsBreak } | ForEach-Object { $_.Text }) -join "`n`n")
    $excerptWords = $excerpt -split '\s+' | Where-Object { $_ }
    if ($excerptWords.Count -gt $SheetWords) {
        $excerpt = ($excerptWords | Select-Object -First $SheetWords) -join ' '
    }

    $sys = @'
Tu rediges la fiche de publication d'une nouvelle destinee a DeviantArt.
Le texte est une oeuvre de fiction : ne le censure pas, ne moralise pas, ne commente pas.
A partir de l'extrait fourni, produis EXACTEMENT ce format, et rien d'autre :

SUMMARY: <deux phrases en anglais, accrocheuses, au present, qui donnent envie de lire, sans reveler la fin>
TAGS: <8 tags en anglais, en minuscules, separes par des virgules, sans diese>
'@

    $ok = $true
    try { Assert-LMStudio -Endpoint $Endpoint -Model $Model }
    catch {
        Write-Warning "  fiche : serveur LM Studio injoignable, resume et tags non generes (lms server start)"
        $ok = $false
    }

    if ($ok) {
        $r = Invoke-LMStudioChat -System $sys -User $excerpt -Model $Model `
                -Endpoint $Endpoint -Temperature 0.6 -MaxTokens 512
        $sum = [regex]::Match($r.Text, '(?ims)^\s*SUMMARY\s*:\s*(.*?)(?=^\s*TAGS\s*:|\z)')
        $tag = [regex]::Match($r.Text, '(?im)^\s*TAGS\s*:\s*(.+)$')

        $lines += "## Resume"
        $lines += ""
        $lines += if ($sum.Success) { $sum.Groups[1].Value.Trim() } else { $r.Text.Trim() }
        $lines += ""
        if ($tag.Success) {
            $tags = @($tag.Groups[1].Value -split ',' | ForEach-Object { $_.Trim().TrimStart('#') } | Where-Object { $_ })
            $lines += "## Tags suggeres"
            $lines += ""
            $lines += ($tags -join ', ')
            $lines += ""
        }
        $lines += "> Resume et tags generes par $Model - a relire avant publication."
    }

    Write-Utf8NoBom -Path $OutPath -Content (($lines -join "`r`n") + "`r`n")
}

# ---------------------------------------------------------------- fichiers

$supported = @('.docx', '.txt', '.md')
$files = @()

foreach ($spec in $Path) {
    if (Test-Path -LiteralPath $spec -PathType Container) {
        $files += Get-ChildItem -LiteralPath $spec -File -Recurse:$Recurse |
                  Where-Object { $supported -contains $_.Extension.ToLower() }
    }
    else {
        $found = @(Get-ChildItem -Path $spec -File -ErrorAction SilentlyContinue)
        if (-not $found) { throw "Fichier introuvable : $spec" }
        $files += $found
    }
}

$files = @($files |
    Where-Object { $supported -contains $_.Extension.ToLower() } |
    Where-Object { $_.BaseName -notlike "*$Suffix*" } |   # sorties deja produites : _DA, _DA_p2, _DA_fiche
    Sort-Object FullName -Unique)

if (-not $files.Count) {
    throw "Aucun fichier .docx / .txt / .md a traiter dans : $($Path -join ', ')"
}

if ($OutDir -and -not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir | Out-Null
}

if ($Clipboard -and $files.Count -gt 1) {
    Write-Warning "-Clipboard ignore : $($files.Count) fichiers a traiter, un seul tient dans le presse-papier."
    $Clipboard = $false
}

$headerBlock = Resolve-Block $Header
$footerBlock = Resolve-Block $Footer

# ---------------------------------------------------------------- traitement

Write-Host ("{0} fichier(s) a traiter - format : {1}" -f $files.Count, $Format) -ForegroundColor Cyan
$opts = @()
if ($English)          { $opts += "ponctuation anglaise" }
if ($French)           { $opts += "typographie francaise preservee" }
if ($KeepDashes)       { $opts += "tirets de dialogue conserves" }
if ($SmartQuotes)      { $opts += "typographie courbe" }
if ($SplitWords -gt 0) { $opts += "decoupage a $SplitWords mots" }
if ($headerBlock.Count){ $opts += "entete" }
if ($footerBlock.Count){ $opts += "pied" }
if ($Sheet)            { $opts += "fiche de publication" }
if ($opts.Count) { Write-Host ("Options : " + ($opts -join ', ')) -ForegroundColor Cyan }
Write-Host ""

$done = 0
foreach ($file in $files) {
    Write-Host ("{0}" -f $file.Name) -ForegroundColor White

    if ($file.Extension.ToLower() -eq '.docx') { $raw = Read-DocxRichParagraphs -File $file.FullName }
    else { $raw = Read-PlainParagraphs -File $file.FullName }

    $all = Resolve-Paragraphs $raw
    if ($KeepEmpty) { $paras = $all }
    else { $paras = @($all | Where-Object { $_.Text -ne '' }) }

    if (-not $paras.Count) {
        Write-Warning "  aucun texte extrait - ignore"
        continue
    }

    $paras = Set-SceneBreaks $paras
    if ($English -or $SmartQuotes) {
        $paras = @($paras | ForEach-Object {
            if ($_.IsBreak) { $_ }
            else { New-Para -Runs (Convert-RunsTypography $_.Runs) -Center $_.Center }
        })
    }

    $title = Get-CleanTitle $file.BaseName
    $breaks = @($paras | Where-Object { $_.IsBreak }).Count
    $words = Measure-Words (($paras | Where-Object { -not $_.IsBreak } | ForEach-Object { $_.Text }) -join ' ')
    $dropped = $all.Count - @($all | Where-Object { $_.Text -ne '' }).Count
    $italic = @($paras | Where-Object { $_.Runs | Where-Object { $_.Italic } }).Count

    Write-Host ("  {0} paragraphes, {1:N0} mots" -f $paras.Count, $words) -ForegroundColor DarkGray
    if ($dropped -gt 0) { Write-Host ("  {0} paragraphe(s) vide(s) supprime(s)" -f $dropped) -ForegroundColor DarkGray }
    if ($breaks -gt 0)  { Write-Host ("  {0} rupture(s) de scene" -f $breaks) -ForegroundColor DarkGray }
    if ($italic -gt 0 -and $Format -eq 'txt') {
        Write-Host ("  {0} paragraphe(s) en italique - perdus en .txt, utilise -Format html" -f $italic) -ForegroundColor DarkYellow
    }

    # ---- controle qualite
    $badQuotes = Get-UnbalancedQuoteParagraphs -Text (($paras | ForEach-Object { $_.Text }) -join "`n`n")
    if ($badQuotes.Count) {
        Write-Warning ("  guillemets non refermes, paragraphe(s) : " + (Write-NumberList $badQuotes))
        Write-Host "    (normal si un dialogue se poursuit sur plusieurs paragraphes)" -ForegroundColor DarkGray
    }
    # En francais le tiret cadratin EST la ponctuation de dialogue attendue :
    # le controle ne vaut que pour un texte destine a etre lu en anglais.
    if (-not ($French -or $KeepDashes)) {
        $badDashes = Get-DashDialogueParagraphs $paras
        if ($badDashes.Count) {
            Write-Warning ("  tiret de dialogue francais non converti, paragraphe(s) : " + (Write-NumberList $badDashes))
        }
    }

    # ---- decoupage et ecriture
    $parts = Split-IntoParts -Paragraphs $paras -MaxWords $SplitWords
    $dir = if ($OutDir) { $OutDir } else { $file.DirectoryName }
    $written = @()

    for ($i = 0; $i -lt $parts.Count; $i++) {
        $body = @()
        if ($headerBlock.Count) { $body += Expand-Placeholders $headerBlock $title ($i + 1) $parts.Count }
        $body += $parts[$i]
        if ($footerBlock.Count) { $body += Expand-Placeholders $footerBlock $title ($i + 1) $parts.Count }

        $partSuffix = if ($parts.Count -gt 1) { "_p{0}" -f ($i + 1) } else { "" }
        $base = Join-Path $dir ($file.BaseName + $Suffix + $partSuffix)
        $partTitle = if ($parts.Count -gt 1) { "{0} ({1}/{2})" -f $title, ($i + 1), $parts.Count } else { $title }

        if ($Format -in @('txt', 'both')) {
            $outTxt = "$base.txt"
            if ((Test-Path -LiteralPath $outTxt) -and -not $Force) {
                throw "$outTxt existe deja. Utilise -Force pour ecraser, ou -OutDir / -Suffix."
            }
            $textOut = ConvertTo-DaText -Paragraphs $body -Blank $BlankLines
            Write-Utf8NoBom -Path $outTxt -Content $textOut
            $written += $outTxt
            if ($Clipboard -and $parts.Count -eq 1) {
                Set-Clipboard -Value $textOut
                Write-Host "  copie dans le presse-papier" -ForegroundColor Magenta
            }
        }

        if ($Format -in @('html', 'both')) {
            $outHtml = "$base.html"
            if ((Test-Path -LiteralPath $outHtml) -and -not $Force) {
                throw "$outHtml existe deja. Utilise -Force pour ecraser, ou -OutDir / -Suffix."
            }
            Write-Utf8NoBom -Path $outHtml -Content (ConvertTo-DaHtml -Paragraphs $body -Title $partTitle)
            $written += $outHtml
        }
    }

    if ($Clipboard -and $parts.Count -gt 1) {
        Write-Warning "  -Clipboard ignore : le texte est decoupe en $($parts.Count) parties."
    }
    if ($parts.Count -gt 1) {
        Write-Host ("  decoupe en {0} parties" -f $parts.Count) -ForegroundColor DarkGray
    }

    # ---- fiche de publication
    if ($Sheet) {
        $sheetPath = Join-Path $dir ($file.BaseName + $Suffix + "_fiche.md")
        if ((Test-Path -LiteralPath $sheetPath) -and -not $Force) {
            throw "$sheetPath existe deja. Utilise -Force pour ecraser."
        }
        New-PublicationSheet -Title $title -Paragraphs $paras -Words $words `
                             -Parts $parts.Count -OutPath $sheetPath
        $written += $sheetPath
    }

    $written | ForEach-Object { Write-Host ("  -> {0}" -f $_) -ForegroundColor Green }
    Write-Host ""
    $done++
}

Write-Host ("Termine : {0}/{1} fichier(s)" -f $done, $files.Count) -ForegroundColor Yellow
