<#
.SYNOPSIS
  Traduction FR->EN via l'API DeepL, mise en forme et dialogues preserves.

.DESCRIPTION
  Ce que l'API permet et que le copier-coller interdit :

    tag_handling=xml   gras et italique passent comme balises <i> <b> et
                       reviennent exactement a leur place. Aucun modele requis.
    plusieurs textes   les paragraphes partent en lot et reviennent DANS L'ORDRE :
                       l'alignement est garanti par construction.
    glossary_id        ta terminologie imposee est appliquee par DeepL lui-meme,
                       au lieu d'etre corrigee apres coup.
    context            oriente la traduction sans etre traduit, et n'est pas
                       facture.

  Les marqueurs de dialogue (puces de liste Word) sont restitues par la lecture
  du .docx et traverses tels quels.

  LA CLE N'EST JAMAIS ECRITE DANS CE FICHIER. Elle est lue depuis la variable
  d'environnement DEEPL_API_KEY, sinon depuis E:\AIs\.deepl-key.
  Le point d'entree (gratuit ou payant) est deduit du suffixe ":fx".

.EXAMPLE
  .\deepl-api.ps1 translate -Path "textes\mon-recit.docx"

.EXAMPLE
  .\deepl-api.ps1 glossary -GlossaryFile "E:\AIs\glossaire.tsv"

.EXAMPLE
  .\deepl-api.ps1 usage
#>

param(
    [Parameter(Mandatory = $true, Position = 0)][ValidateSet('translate', 'glossary', 'usage')][string]$Mode,
    [string]$Path,
    [string]$Out,
    [string]$OutDir = "E:\AIs\Output",
    [string[]]$ContextFile,
    [string]$Context,
    [string]$GlossaryFile,
    [string]$GlossaryName = "fr-en",
    [switch]$NoGlossary,
    [int]$MaxCharsPerRequest = 100000
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\lib-lmstudio.ps1"

# ------------------------------------------------------------------ clef
function Get-DeeplKey {
    if ($env:DEEPL_API_KEY) { return $env:DEEPL_API_KEY.Trim() }

    # setx n'alimente que les processus lances ENSUITE : on va lire la valeur
    # persistee, sinon il faudrait rouvrir un terminal apres chaque changement.
    foreach ($scope in @('User', 'Machine')) {
        $v = [System.Environment]::GetEnvironmentVariable('DEEPL_API_KEY', $scope)
        if ($v) { return $v.Trim() }
    }

    $file = Join-Path $PSScriptRoot ".deepl-key"
    if (Test-Path $file) {
        $k = ([System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)).Trim()
        if ($k) { return $k }
    }
    throw "Clef DeepL introuvable.`n" +
          "  Variable d'environnement :  setx DEEPL_API_KEY `"ta-clef`"  (puis nouveau terminal)`n" +
          "  ou fichier                :  $file"
}

$key = Get-DeeplKey
# Les clefs gratuites se terminent par ":fx" et visent un autre hote.
$base = if ($key -match ':fx$') { 'https://api-free.deepl.com/v2' } else { 'https://api.deepl.com/v2' }
$headers = @{ Authorization = "DeepL-Auth-Key $key" }
$plan = if ($key -match ':fx$') { 'Free' } else { 'Pro' }

function Invoke-Deepl {
    param([string]$Endpoint, [string]$Method = 'POST', $Body, [string]$ContentType = 'application/x-www-form-urlencoded')
    try {
        # Invoke-RestMethod ne respecte pas toujours le charset en PS 5.1 et rend
        # des guillemets courbes en mojibake. On decode le flux brut en UTF-8.
        if ($Method -eq 'GET') {
            $resp = Invoke-WebRequest -Uri "$base/$Endpoint" -Headers $headers -Method GET `
                        -TimeoutSec 120 -UseBasicParsing
        }
        else {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Body)
            $resp = Invoke-WebRequest -Uri "$base/$Endpoint" -Headers $headers -Method $Method `
                        -Body $bytes -ContentType "$ContentType; charset=utf-8" `
                        -TimeoutSec 600 -UseBasicParsing
        }
        $json = [System.Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())
        if (-not $json) { return $null }
        $json | ConvertFrom-Json
    }
    catch {
        $resp = $_.Exception.Response
        $code = if ($resp) { [int]$resp.StatusCode } else { 0 }
        $detail = ""
        if ($resp) {
            try {
                $sr = New-Object System.IO.StreamReader($resp.GetResponseStream())
                $detail = $sr.ReadToEnd(); $sr.Close()
            } catch { }
        }
        switch ($code) {
            403 { throw "DeepL 403 : clef refusee. Verifie DEEPL_API_KEY. $detail" }
            456 { throw "DeepL 456 : quota de caracteres epuise pour ce mois. $detail" }
            429 { throw "DeepL 429 : trop de requetes, reessaie dans quelques instants." }
            default { throw "DeepL $code : $detail" }
        }
    }
}

# ------------------------------------------------------------------ usage
if ($Mode -eq 'usage') {
    $u = Invoke-Deepl -Endpoint 'usage' -Method GET
    $pct = if ($u.character_limit) { 100 * $u.character_count / $u.character_limit } else { 0 }
    Write-Host "Plan      : $plan  ($base)" -ForegroundColor Cyan
    Write-Host ("Caracteres: {0:N0} / {1:N0}  ({2:N1} %)" -f $u.character_count, $u.character_limit, $pct) -ForegroundColor Cyan
    return
}

# ------------------------------------------------------------------ glossaire
$glossaryStore = Join-Path $PSScriptRoot ".deepl-glossary-id"

if ($Mode -eq 'glossary') {
    if (-not $GlossaryFile) { throw "-GlossaryFile est requis : un .tsv de lignes 'francais<TAB>anglais'." }
    if (-not (Test-Path $GlossaryFile)) { throw "Fichier introuvable : $GlossaryFile" }

    $tsv = [System.IO.File]::ReadAllText($GlossaryFile, [System.Text.Encoding]::UTF8)
    $entries = @($tsv -split "`r?`n" | Where-Object { $_ -match "`t" -and $_ -notmatch '^\s*#' })
    if (-not $entries.Count) { throw "Aucune entree valide dans $GlossaryFile (attendu : terme<TAB>traduction)." }

    # Un glossaire DeepL est immuable : on supprime l'ancien avant d'en creer un.
    if (Test-Path $glossaryStore) {
        $old = ([System.IO.File]::ReadAllText($glossaryStore, [System.Text.Encoding]::UTF8)).Trim()
        if ($old) {
            try { Invoke-Deepl -Endpoint "glossaries/$old" -Method DELETE | Out-Null; Write-Host "Ancien glossaire supprime." -ForegroundColor DarkGray }
            catch { Write-Host "Ancien glossaire deja absent." -ForegroundColor DarkGray }
        }
    }

    $body = @{
        name           = $GlossaryName
        source_lang    = 'fr'
        target_lang    = 'en'
        entries        = ($entries -join "`n")
        entries_format = 'tsv'
    }
    $g = Invoke-Deepl -Endpoint 'glossaries' -Body $body
    Write-Utf8NoBom -Path $glossaryStore -Content $g.glossary_id
    Write-Host ("Glossaire cree : {0} entrees" -f $g.entry_count) -ForegroundColor Green
    Write-Host ("Identifiant enregistre dans {0}" -f $glossaryStore) -ForegroundColor DarkGray
    return
}

# ------------------------------------------------------------------ traduction
if (-not $Path) { throw "-Path est requis en mode translate." }
if (-not (Test-Path $Path)) { throw "Fichier introuvable : $Path" }
$item = Get-Item $Path
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
if (-not $Out) { $Out = Join-Path $OutDir ($item.BaseName + "_EN_deepl.docx") }

# Les balises internes deviennent du XML, que DeepL replace exactement.
$marked = if ($item.Extension -eq '.docx') { Read-DocxMarkedText -File $item.FullName }
          else { Read-SourceText -File $item.FullName }
$paras = Get-Paragraphs -Text $marked

$toXml = { param($s)
    $s = [System.Security.SecurityElement]::Escape($s)
    $s = $s -replace '\[\[i\]\]', '<i>' -replace '\[\[/i\]\]', '</i>'
    $s = $s -replace '\[\[b\]\]', '<b>' -replace '\[\[/b\]\]', '</b>'
    $s
}
$fromXml = { param($s)
    $s = $s -replace '<i>', '[[i]]' -replace '</i>', '[[/i]]'
    $s = $s -replace '<b>', '[[b]]' -replace '</b>', '[[/b]]'
    [System.Net.WebUtility]::HtmlDecode($s)
}

$xmlParas = @($paras | ForEach-Object { & $toXml $_ })

# Contexte : oriente la traduction, n'est pas traduit, n'est pas facture.
$ctx = ""
foreach ($cf in $ContextFile) { if ($cf) { $ctx = ($ctx + "`n`n" + (Read-SourceText -File $cf).Trim()).Trim() } }
if ($Context) { $ctx = ($ctx + "`n" + $Context).Trim() }

# Avec tag_handling=xml, DeepL parse AUSSI le contexte. Les chevrons du modele
# de fiche (<titre>, <homme|femme>) le feraient echouer : on les echappe.
if ($ctx) { $ctx = [System.Security.SecurityElement]::Escape($ctx) }

$glossaryId = $null
if (-not $NoGlossary -and (Test-Path $glossaryStore)) {
    $glossaryId = ([System.IO.File]::ReadAllText($glossaryStore, [System.Text.Encoding]::UTF8)).Trim()
}

Write-Host "Plan      : $plan" -ForegroundColor Cyan
Write-Host ("Source    : {0}  ({1} paragraphes, {2:N0} mots)" -f $item.Name, $paras.Count, (Measure-Words $marked)) -ForegroundColor Cyan
if ($ctx)        { Write-Host ("Contexte  : {0} mots (non facture)" -f (Measure-Words $ctx)) -ForegroundColor Cyan }
if ($glossaryId) { Write-Host "Glossaire : applique" -ForegroundColor Cyan }
else             { Write-Host "Glossaire : aucun (lance le mode glossary pour en creer un)" -ForegroundColor DarkGray }

# Lots sous la limite de taille de requete.
$batches = @(); $cur = @(); $len = 0
foreach ($x in $xmlParas) {
    if ($cur.Count -gt 0 -and ($len + $x.Length) -gt $MaxCharsPerRequest) {
        $batches += , $cur; $cur = @(); $len = 0
    }
    $cur += $x; $len += $x.Length
}
if ($cur.Count) { $batches += , $cur }
Write-Host ("Envoi     : {0} requete(s)" -f $batches.Count) -ForegroundColor Cyan
Write-Host ""

$translated = @()
$sw = [System.Diagnostics.Stopwatch]::StartNew()
for ($b = 0; $b -lt $batches.Count; $b++) {
    Write-Host ("[{0}/{1}] {2} paragraphes..." -f ($b + 1), $batches.Count, $batches[$b].Count) -NoNewline

    # Corps encode a la main : le parametre text est repete, un par paragraphe.
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("source_lang=FR&target_lang=EN-US&tag_handling=xml&preserve_formatting=1")
    if ($glossaryId) { [void]$sb.Append("&glossary_id=" + [uri]::EscapeDataString($glossaryId)) }
    if ($ctx)        { [void]$sb.Append("&context=" + [uri]::EscapeDataString($ctx)) }
    foreach ($t in $batches[$b]) { [void]$sb.Append("&text=" + [uri]::EscapeDataString($t)) }

    $r = Invoke-Deepl -Endpoint 'translate' -Body $sb.ToString()
    if ($r.translations.Count -ne $batches[$b].Count) {
        throw ("DeepL a renvoye {0} traductions pour {1} paragraphes envoyes." -f $r.translations.Count, $batches[$b].Count)
    }
    $translated += @($r.translations | ForEach-Object { & $fromXml $_.text })
    Write-Host " ok" -ForegroundColor Green
}
$sw.Stop()

if ($translated.Count -ne $paras.Count) {
    throw ("Alignement rompu : {0} paragraphes traduits pour {1} envoyes." -f $translated.Count, $paras.Count)
}

$final = $translated -join "`n`n"

# DeepL ajoute des guillemets anglais par-dessus les tirets de dialogue de
# l'auteur. On les ote en confrontant a la source, ligne par ligne.
$rep = Repair-DialogueDashes -SourceParagraphs $paras -Text $final
$final = $rep.Text

Write-DocxFromText -Path $Out -Text $final
$outTxt = [System.IO.Path]::ChangeExtension($Out, '.txt')
Write-Utf8NoBom -Path $outTxt -Content (Remove-Marks -Text $final)

$srcMarks = Get-MarkCounts -Text $marked
$outMarks = Get-MarkCounts -Text $final
$srcDash = (($marked -split "`n") | Where-Object { $_ -match '^\s*[-\u2013\u2014]\s*\S' }).Count
$outDash = (($final  -split "`n") | Where-Object { $_ -match '^\s*[-\u2013\u2014]\s*\S' }).Count

Write-Host ""
Write-Host ("Termine en {0:N0} s" -f $sw.Elapsed.TotalSeconds) -ForegroundColor Yellow
Write-Host ("Paragraphes : {0} -> {1}" -f $paras.Count, (Get-Paragraphs -Text $final).Count) -ForegroundColor Yellow
Write-Host ("Repliques   : {0} -> {1}" -f $srcDash, $outDash) -ForegroundColor Yellow
if ($rep.Stripped) {
    Write-Host ("Guillemets ajoutes par DeepL et retires : {0}" -f $rep.Stripped) -ForegroundColor DarkYellow
}
if (-not $rep.Aligned) { Write-Warning "Alignement source/traduction rompu : dialogues non verifies." }
Write-Host ("Italiques   : {0} -> {1}   Gras : {2} -> {3}" -f `
    $srcMarks.IOpen, $outMarks.IOpen, $srcMarks.BOpen, $outMarks.BOpen) -ForegroundColor Yellow
Write-Host "Avec mise en forme : $Out" -ForegroundColor Yellow
Write-Host "Texte brut         : $outTxt" -ForegroundColor Yellow
