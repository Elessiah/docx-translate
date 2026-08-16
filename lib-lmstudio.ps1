<#
  Fonctions communes a translate.ps1 et review.ps1.
  Dot-source :   . "$PSScriptRoot\lib-lmstudio.ps1"
#>

# Configuration : valeurs par defaut, puis surcharge locale non versionnee.
foreach ($c in @('config.ps1', 'config.local.ps1')) {
    $f = Join-Path $PSScriptRoot $c
    if (Test-Path $f) { . $f }
}

function Get-DefaultModel { if ($script:DefaultModel) { $script:DefaultModel } else { 'local-model' } }

function Get-DocxBullets {
    <# Les repliques de dialogue sont souvent des paragraphes de LISTE Word : le
       tiret est une puce de numerotation, pas un caractere du texte. Sans cette
       table, l'extraction perd tous les marqueurs de dialogue.
       Retourne une table numId -> caractere de puce. #>
    param([Parameter(Mandatory)]$Zip)

    $map = @{}
    $entry = $Zip.Entries | Where-Object { ($_.FullName -replace '\\', '/') -eq 'word/numbering.xml' }
    if (-not $entry) { return $map }

    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    $xmlText = $reader.ReadToEnd()
    $reader.Close()

    $W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $xml = [xml]$xmlText
    $ns2 = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns2.AddNamespace('w', $W)

    $abstract = @{}
    foreach ($a in $xml.SelectNodes('//w:abstractNum', $ns2)) {
        $id = $a.GetAttribute('abstractNumId', $W)
        $lvl = $a.SelectSingleNode("w:lvl[@w:ilvl='0']", $ns2)
        if (-not $lvl) { continue }
        $fmt = $lvl.SelectSingleNode('w:numFmt', $ns2)
        $txt = $lvl.SelectSingleNode('w:lvlText', $ns2)
        if (-not $txt) { continue }
        $char = $txt.GetAttribute('val', $W)
        # Les listes numerotees portent des %1 : on ne garde que les puces.
        if ($fmt -and $fmt.GetAttribute('val', $W) -ne 'bullet') { continue }
        if ($char -match '%') { continue }
        $abstract[$id] = $char
    }

    foreach ($n in $xml.SelectNodes('//w:num', $ns2)) {
        $numId = $n.GetAttribute('numId', $W)
        $ref = $n.SelectSingleNode('w:abstractNumId', $ns2)
        if (-not $ref) { continue }
        $aid = $ref.GetAttribute('val', $W)
        if ($abstract.ContainsKey($aid)) { $map[$numId] = $abstract[$aid] }
    }
    $map
}

function Read-DocxText {
    <# Texte brut, marqueurs de dialogue compris. #>
    param([Parameter(Mandatory)][string]$File)
    Remove-Marks -Text (Read-DocxMarkedText -File $File)
}

function Read-SourceText {
    <# Lit .docx, .txt, .md, ou tout texte brut. Rejette .pdf. #>
    param([Parameter(Mandatory)][string]$File)

    if (-not (Test-Path $File)) { throw "Fichier introuvable : $File" }
    $item = Get-Item $File
    $ext = $item.Extension.ToLower()

    if ($ext -eq '.pdf') {
        throw "Les .pdf ne sont pas supportes. Convertis en .docx ou .txt d'abord " +
              "(Word : Fichier > Ouvrir le PDF, puis Enregistrer sous .docx)."
    }

    if ($ext -eq '.docx') { $text = Read-DocxText -File $item.FullName }
    else { $text = [System.IO.File]::ReadAllText($item.FullName, [System.Text.Encoding]::UTF8) }

    if (-not $text.Trim()) { throw "Aucun texte extrait de $File" }
    $text
}

function Get-Paragraphs {
    param([Parameter(Mandatory)][string]$Text)
    , ([regex]::Split($Text, '\r?\n\s*\r?\n') | Where-Object { $_.Trim() })
}

function Measure-Words {
    param([string]$Text)
    if (-not $Text) { return 0 }
    ($Text -split '\s+' | Where-Object { $_ }).Count
}

function Group-IntoChunks {
    <# Regroupe des paragraphes en blocs sous un budget de mots. #>
    param(
        [Parameter(Mandatory)][string[]]$Paragraphs,
        [int]$MaxWords = 900
    )
    $chunks = @(); $current = @(); $count = 0
    foreach ($p in $Paragraphs) {
        $w = Measure-Words $p
        if ($count -gt 0 -and ($count + $w) -gt $MaxWords) {
            $chunks += , ($current -join "`n`n")
            $current = @(); $count = 0
        }
        $current += $p
        $count += $w
    }
    if ($current.Count) { $chunks += , ($current -join "`n`n") }
    , $chunks
}

function Test-StructuralBreak {
    <# Vrai si le paragraphe est une frontiere de structure : ligne de separation
       (---, ***, ~~~) ou titre de chapitre. Le modele traite ces marqueurs comme
       une fin de tache et s'arrete de traduire : un bloc ne doit jamais en
       contenir un ailleurs qu'a son tout debut. #>
    param([Parameter(Mandatory)][string]$Paragraph)

    foreach ($line in ($Paragraph -split '\r?\n')) {
        $l = $line.Trim()
        if (-not $l) { continue }
        # ligne faite uniquement de caracteres de separation
        $bare = $l -replace '\s', ''
        if ($bare.Length -ge 3 -and $bare -match '^[\*\-_=~#\u2014\u2013\u2022\u00B7]+$') { return $true }
        # titre de chapitre explicite ; u00E9 couvre epilogue accentue
        if ($l -match '^(chapitre|chapter|partie|part|prologue|\u00E9pilogue|epilogue)\b') { return $true }
    }
    $false
}

function Group-ParagraphsIntoChunks {
    <# Decoupe une liste de paragraphes en blocs sous TROIS contraintes :
       budget de mots, nombre de paragraphes, et jamais de frontiere structurelle
       au milieu d'un bloc. Retourne un tableau de tableaux de paragraphes. #>
    param(
        [Parameter(Mandatory)][string[]]$Paragraphs,
        [int]$MaxWords = 400,
        [int]$MaxParagraphs = 8
    )

    $chunks = @(); $current = @(); $count = 0
    foreach ($p in $Paragraphs) {
        $w = Measure-Words $p
        $isBreak = Test-StructuralBreak -Paragraph $p

        $mustSplit = $current.Count -gt 0 -and (
            $isBreak -or
            ($count + $w) -gt $MaxWords -or
            $current.Count -ge $MaxParagraphs
        )
        if ($mustSplit) {
            $chunks += , $current
            $current = @(); $count = 0
        }
        $current += $p
        $count += $w
    }
    if ($current.Count) { $chunks += , $current }
    , $chunks
}

function Get-PromptText {
    <# Charge un prompt systeme depuis prompts\<nom>.txt.
       Si prompts\<nom>.local.txt existe, il a la priorite : c'est la version
       personnelle, non versionnee. Le depot ne porte que la version par defaut. #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$PromptFile
    )
    if ($PromptFile) {
        if (-not (Test-Path $PromptFile)) { throw "Prompt introuvable : $PromptFile" }
        return ([System.IO.File]::ReadAllText($PromptFile, [System.Text.Encoding]::UTF8)).TrimEnd()
    }
    $dir = Join-Path $PSScriptRoot 'prompts'
    foreach ($candidate in @("$Name.local.txt", "$Name.txt")) {
        $f = Join-Path $dir $candidate
        if (Test-Path $f) {
            return ([System.IO.File]::ReadAllText($f, [System.Text.Encoding]::UTF8)).TrimEnd()
        }
    }
    throw "Aucun prompt trouve pour '$Name' dans $dir"
}

function Assert-LMStudio {
    <# Verifie que le serveur repond et que le modele est disponible. #>
    param(
        [Parameter(Mandatory)][string]$Endpoint,
        [Parameter(Mandatory)][string]$Model
    )
    $modelsUrl = $Endpoint -replace '/chat/completions\s*$', '/models'
    try {
        $available = (Invoke-RestMethod -Uri $modelsUrl -TimeoutSec 15).data.id
    }
    catch {
        throw "Serveur LM Studio injoignable sur $modelsUrl`n" +
              "Demarre-le avec :   lms server start"
    }
    if ($available -notcontains $Model) {
        throw "Modele '$Model' introuvable. Modeles disponibles :`n  " + ($available -join "`n  ")
    }
}

function Get-LmsPath {
    <# Chemin du CLI lms : d'abord le PATH utilisateur, sinon le binaire embarque. #>
    $candidates = @(
        (Join-Path $env:USERPROFILE ".lmstudio\bin\lms.exe"),
        "E:\AIs\LMStudio\resources\app\.webpack\lms.exe"
    )
    foreach ($c in $candidates) { if (Test-Path $c) { return $c } }
    throw "CLI lms introuvable. LM Studio est-il installe ?"
}

function Invoke-Lms {
    <# Appelle le CLI lms en renvoyant sa sortie standard.
       Windows PowerShell 5.1 transforme chaque ligne de stderr d'un executable
       natif en NativeCommandError des qu'on la redirige : avec
       $ErrorActionPreference = 'Stop', une commande qui REUSSIT ferait planter
       l'appelant (lms unload ecrit son message de succes sur stderr).
       D'ou 2>$null et le relachement temporaire de la preference. #>
    param([Parameter(Mandatory)][string[]]$Arguments)

    $lms = Get-LmsPath
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { $out = & $lms @Arguments 2>$null | Out-String }
    finally { $ErrorActionPreference = $prev }
    $out
}

function Initialize-LMStudioSession {
    <# Demarre le serveur s'il est arrete, puis charge le modele s'il n'est pas
       deja charge avec au moins le contexte demande. Le chargement JIT n'utilise
       que le contexte par defaut (8192), souvent insuffisant pour la revision. #>
    param(
        [Parameter(Mandatory)][string]$Model,
        [Parameter(Mandatory)][string]$Endpoint,
        [int]$ContextLength = 20000
    )

    $modelsUrl = $Endpoint -replace '/chat/completions\s*$', '/models'

    # 1. Serveur
    $up = $false
    try { $null = Invoke-RestMethod $modelsUrl -TimeoutSec 5; $up = $true } catch { }
    if (-not $up) {
        Write-Host "Serveur arrete -> demarrage..." -ForegroundColor DarkYellow
        Invoke-Lms -Arguments @('server', 'start') | Out-Null
        for ($i = 0; $i -lt 20; $i++) {
            Start-Sleep -Seconds 2
            try { $null = Invoke-RestMethod $modelsUrl -TimeoutSec 5; $up = $true; break } catch { }
        }
        if (-not $up) { throw "Le serveur LM Studio n'a pas demarre." }
    }
    Write-Host "Serveur   : OK" -ForegroundColor DarkGray

    # 2. Modele charge avec assez de contexte ?
    $psOut = Invoke-Lms -Arguments @('ps')
    $line = ($psOut -split '\r?\n') | Where-Object { $_ -match [regex]::Escape($Model) } | Select-Object -First 1

    $loadedCtx = 0
    if ($line) {
        $cols = @(($line -split '\s{2,}') | Where-Object { $_.Trim() })
        $num = $cols | Where-Object { $_ -match '^\d+$' } | Select-Object -First 1
        if ($num) { $loadedCtx = [int]$num }
    }

    if ($loadedCtx -ge $ContextLength) {
        Write-Host "Modele    : deja charge (contexte $loadedCtx)" -ForegroundColor DarkGray
        return
    }

    if ($line) {
        Write-Host "Modele    : recharge (contexte $loadedCtx -> $ContextLength)" -ForegroundColor DarkYellow
        Invoke-Lms -Arguments @('unload', $Model) | Out-Null
    }
    else {
        Write-Host "Modele    : chargement (contexte $ContextLength)..." -ForegroundColor DarkYellow
    }
    Invoke-Lms -Arguments @('load', $Model, '-c', "$ContextLength", '-y') | Out-Null

    $psOut = Invoke-Lms -Arguments @('ps')
    if ($psOut -notmatch [regex]::Escape($Model)) {
        throw "Echec du chargement du modele '$Model'."
    }
    Write-Host "Modele    : OK" -ForegroundColor DarkGray
}

function Invoke-LMStudioChat {
    <# Un aller-retour /v1/chat/completions. Retourne @{ Text; Tokens }. #>
    param(
        [Parameter(Mandatory)][string]$System,
        [Parameter(Mandatory)][string]$User,
        [Parameter(Mandatory)][string]$Model,
        [Parameter(Mandatory)][string]$Endpoint,
        [double]$Temperature = 0.3,
        [int]$MaxTokens = 4096,
        [switch]$Think
    )
    # Qwen3 est un modele hybride : /think et /no_think pilotent le raisonnement.
    # En mode Think il faut de la marge en tokens, le raisonnement s'impute dessus.
    if ($Think) { $switchTag = "/think"; $MaxTokens = $MaxTokens * 3 }
    else        { $switchTag = "/no_think" }

    $body = @{
        model       = $Model
        messages    = @(
            @{ role = "system"; content = $System },
            @{ role = "user";   content = ($User + "`n`n" + $switchTag) }
        )
        temperature = $Temperature
        max_tokens  = $MaxTokens
    } | ConvertTo-Json -Depth 5 -Compress

    $r = Invoke-RestMethod -Uri $Endpoint -Method Post `
            -ContentType "application/json; charset=utf-8" `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -TimeoutSec 3600

    $txt = $r.choices[0].message.content
    # Qwen3 peut emettre un bloc de raisonnement malgre /no_think
    $txt = [regex]::Replace($txt, '(?s)<think>.*?</think>', '').Trim()

    # finish_reason = "length" : la reponse a ete COUPEE par max_tokens. C'est la
    # seule facon de distinguer une omission du modele d'une troncature technique.
    $finish = $r.choices[0].finish_reason
    @{
        Text        = $txt
        Tokens      = $r.usage.completion_tokens
        Finish      = $finish
        Truncated   = ($finish -eq 'length')
        PromptTokens = $r.usage.prompt_tokens
    }
}

# Classes de caracteres en echappement \u : le fichier peut etre relu en ANSI
# par PowerShell 5.1, des guillemets courbes litteraux seraient corrompus.
$script:QuoteClass = '["\u201C\u201D]'
$script:OpenQuote  = '^\s*["\u201C]'

function Get-UnbalancedQuoteParagraphs {
    <# Numeros des paragraphes dont le nombre de guillemets est impair,
       c'est-a-dire une replique de dialogue jamais refermee. #>
    param([Parameter(Mandatory)][string]$Text)
    $bad = @()
    $paras = Get-Paragraphs -Text $Text
    for ($i = 0; $i -lt $paras.Count; $i++) {
        $n = ([regex]::Matches($paras[$i], $script:QuoteClass)).Count
        if ($n % 2 -ne 0) { $bad += ($i + 1) }
    }
    , $bad
}

function Repair-DialogueDashes {
    <# Les nouvelles n'ont pas toutes la meme convention : certaines ouvrent les
       repliques par un tiret, d'autres non. La regle est donc de REPRODUIRE la
       source, jamais d'imposer un style.

       Les repliques sont separees par des sauts de ligne manuels A L'INTERIEUR
       d'un meme paragraphe Word : la comparaison se fait ligne par ligne, pas
       bloc par bloc. #>
    param(
        [Parameter(Mandatory)][string[]]$SourceParagraphs,
        [Parameter(Mandatory)][string]$Text
    )

    $out = Get-Paragraphs -Text $Text
    if ($out.Count -ne $SourceParagraphs.Count) {
        return @{ Text = $Text; Fixed = 0; Stripped = 0; Aligned = $false }
    }

    $dashStart   = '^(\s*)([-\u2013\u2014])\s*'
    $startsDash  = '^\s*[-\u2013\u2014]\s'
    $startsQuote = '^(\s*)["\u201C]\s*'

    $fixed = 0; $stripped = 0
    for ($i = 0; $i -lt $out.Count; $i++) {
        $srcLines = $SourceParagraphs[$i] -split "`n"
        $outLines = $out[$i] -split "`n"
        # Sans correspondance ligne a ligne, on ne touche a rien.
        if ($srcLines.Count -ne $outLines.Count) { continue }

        for ($j = 0; $j -lt $outLines.Count; $j++) {
            $m = [regex]::Match($srcLines[$j], $dashStart)
            $hasQuote = $outLines[$j] -match $startsQuote

            if ($m.Success) {
                # La source ouvre par un tiret : la traduction doit faire pareil.
                if ($outLines[$j] -match $startsDash) {
                    # Tiret present mais guillemets ajoutes par-dessus : on les ote.
                    # Tiret present mais guillemets ajoutes par-dessus. Si la source
                    # n'en contient aucun, la traduction ne doit en avoir aucun :
                    # DeepL encadre chaque phrase, donc compter les paires ne suffit pas.
                    if ($srcLines[$j] -notmatch '["\u201C\u201D\u00AB\u00BB]') {
                        $before = ([regex]::Matches($outLines[$j], '["\u201C\u201D\u00AB\u00BB]')).Count
                        if ($before -gt 0) {
                            $outLines[$j] = $outLines[$j] -replace '["\u201C\u201D\u00AB\u00BB]', ''
                            $outLines[$j] = $outLines[$j] -replace '  +', ' '
                            $stripped += $before
                        }
                    }
                    continue
                }
                $mark = $m.Groups[2].Value
                if ($hasQuote) {
                    $n = ([regex]::Matches($outLines[$j], '["\u201C\u201D]')).Count
                    $outLines[$j] = [regex]::Replace($outLines[$j], $startsQuote, "`$1$mark ")
                    if ($n -eq 2) { $outLines[$j] = [regex]::Replace($outLines[$j], '["\u201C\u201D]', '', 1) }
                }
                else { $outLines[$j] = "$mark " + $outLines[$j].TrimStart() }
                $fixed++
            }
            elseif ($hasQuote) {
                # La source n'a pas de marqueur : le modele en a invente un.
                $n = ([regex]::Matches($outLines[$j], '["\u201C\u201D]')).Count
                if ($n -eq 2) {
                    $outLines[$j] = [regex]::Replace($outLines[$j], $startsQuote, '$1')
                    $outLines[$j] = [regex]::Replace($outLines[$j], '["\u201C\u201D]', '', 1)
                    $stripped++
                }
            }
        }
        $out[$i] = $outLines -join "`n"
    }
    @{ Text = ($out -join "`n`n"); Fixed = $fixed; Stripped = $stripped; Aligned = $true }
}

function Repair-DialogueQuotes {
    <# Le modele oublie systematiquement le guillemet OUVRANT quand il convertit
       un tiret cadratin francais :  Texte," dit-elle.
       Repare les paragraphes a nombre impair de guillemets qui ne commencent
       pas par un guillemet. Les autres cas sont laisses tels quels. #>
    param([Parameter(Mandatory)][string]$Text)
    # PAS de @() ici : Get-Paragraphs renvoie deja , $tableau (protection contre
    # le deroulement). Un @() supplementaire donnerait un tableau imbrique.
    $paras = Get-Paragraphs -Text $Text
    $fixed = 0
    for ($i = 0; $i -lt $paras.Count; $i++) {
        $n = ([regex]::Matches($paras[$i], $script:QuoteClass)).Count
        if (($n % 2 -ne 0) -and ($paras[$i] -notmatch $script:OpenQuote)) {
            $paras[$i] = '"' + $paras[$i].TrimStart()
            $fixed++
        }
    }
    @{ Text = ($paras -join "`n`n"); Fixed = $fixed }
}

function Repair-FrenchTypography {
    <# Le modele rend l'apostrophe droite et avale les espaces insecables meme
       quand on le lui interdit. On les retablit apres coup : c'est mecanique,
       donc fiable, la ou la consigne ne l'est pas.
         - apostrophe droite entre deux lettres -> apostrophe courbe
         - espace insecable avant ? ! ; et avant : suivi d'un blanc
         - espace insecable a l'interieur des guillemets francais #>
    param([Parameter(Mandatory)][string]$Text)

    # [string] et non [char] : passe un char en 3e argument de [regex]::Replace et
    # .NET resout vers la surcharge MatchEvaluator, qui echoue.
    $NBSP  = [string][char]0x00A0
    $RSQUO = [string][char]0x2019
    $LAQUO = [string][char]0x00AB
    $RAQUO = [string][char]0x00BB

    $before = @{
        Straight = ([regex]::Matches($Text, "'")).Count
        Nbsp     = ([regex]::Matches($Text, $NBSP)).Count
    }

    # apostrophe d'elision ou de possession : toujours courbe en francais
    $t = [regex]::Replace($Text, "(?<=\p{L})'(?=\p{L})", $RSQUO)

    # ponctuation double : espace insecable, jamais d'espace ordinaire ni rien
    $t = [regex]::Replace($t, "(?<=\S)[ $NBSP]*([?!;])", "$NBSP`$1")
    $t = [regex]::Replace($t, "(?<=\S)[ $NBSP]*:(?=\s|$)", "$NBSP" + ':')

    # guillemets francais
    $t = [regex]::Replace($t, "$LAQUO[ $NBSP]*", "$LAQUO$NBSP")
    $t = [regex]::Replace($t, "[ $NBSP]*$RAQUO", "$NBSP$RAQUO")

    $after = @{
        Straight = ([regex]::Matches($t, "'")).Count
        Nbsp     = ([regex]::Matches($t, $NBSP)).Count
    }

    @{
        Text          = $t
        ApostrophesFixed = [Math]::Max(0, $before.Straight - $after.Straight)
        NbspAdded        = [Math]::Max(0, $after.Nbsp - $before.Nbsp)
    }
}

function Remove-TemplateEcho {
    <# Le modele recopie parfois le gabarit du prompt en tete de reponse, par
       exemple "<la traduction corrigee complete, rien d'autre>". Ces lignes
       entierement entourees de chevrons ne font jamais partie d'une traduction :
       on les retire du debut de la reponse. #>
    param([Parameter(Mandatory)][string]$Text)

    $lines = $Text -split '\r?\n'
    $i = 0
    $removed = 0
    while ($i -lt $lines.Count) {
        $l = $lines[$i].Trim()
        if ($l -eq '') { $i++; continue }
        if ($l -match '^<[^>]{0,160}>$') { $removed++; $i++; continue }
        break
    }
    @{ Text = (($lines | Select-Object -Skip $i) -join "`n").Trim(); Removed = $removed }
}

function Remove-NoOpNotes {
    <# Supprime les lignes de note du type  - "X" -> "X" : ...  ou avant et apres
       sont identiques : le modele signale une correction qu'il n'a pas faite. #>
    param([Parameter(Mandatory)][string]$Notes)
    $kept = @()
    $dropped = 0
    # Le modele ecrit tantot  - "x" -> "y" : raison  tantot  - x -> y : raison.
    # On capture les deux : guillemets optionnels, retires avant comparaison.
    $pattern = '^\s*[-*]\s*(.+?)\s*->\s*(.+?)\s*:\s'
    foreach ($line in ($Notes -split "`r?`n")) {
        $m = [regex]::Match($line, $pattern)
        if ($m.Success) {
            $a = ($m.Groups[1].Value -replace $script:QuoteClass, '' -replace '\s+', ' ').Trim()
            $b = ($m.Groups[2].Value -replace $script:QuoteClass, '' -replace '\s+', ' ').Trim()
            if ($a -eq $b) { $dropped++; continue }
        }
        $kept += $line
    }
    @{ Notes = (($kept -join "`n").Trim()); Dropped = $dropped }
}

# Balises de mise en forme transportees dans le texte a travers la traduction.
# Choisies pour ne jamais apparaitre dans de la prose et survivre a un LLM :
# doubles crochets, ASCII pur, pas de caractere que le modele voudrait traduire.
$script:MarkRegex = '(\[\[/?[ib]\]\])'

function Read-DocxMarkedText {
    <# Comme Read-DocxText, mais entoure les passages en italique de [[i]]...[[/i]]
       et ceux en gras de [[b]]...[[/b]]. Les runs voisins de meme mise en forme
       sont fusionnes pour ne pas semer des balises a chaque mot. #>
    param([Parameter(Mandatory)][string]$File)

    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
    try { $zip = [System.IO.Compression.ZipFile]::OpenRead($File) }
    catch [System.IO.IOException] {
        throw "Impossible d'ouvrir $File : le fichier est verrouille. Ferme-le dans Word puis relance."
    }
    $W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    try {
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

    function Test-Flag($rPr, [string]$name) {
        if (-not $rPr) { return $false }
        $node = $rPr.SelectSingleNode("w:$name", $ns)
        if (-not $node) { return $false }
        $val = $node.GetAttribute('val', $W)
        return ($val -eq '' -or $val -notin @('0', 'false', 'off'))
    }

    $paragraphs = @()
    foreach ($p in $xml.SelectNodes('//w:body//w:p', $ns)) {
        $sb = New-Object System.Text.StringBuilder
        $curI = $false; $curB = $false

        # Paragraphe de liste : la puce est le marqueur de dialogue de l'auteur.
        # Elle n'est pas dans le texte, il faut la restituer.
        $numId = $p.SelectSingleNode('w:pPr/w:numPr/w:numId', $ns)
        if ($numId) {
            $id = $numId.GetAttribute('val', $W)
            $bullet = if ($bullets.ContainsKey($id)) { $bullets[$id] } else { '-' }
            if ($bullet) { [void]$sb.Append($bullet + ' ') }
        }

        foreach ($n in $p.SelectNodes('.//w:r[not(ancestor::w:del)]/*[self::w:t or self::w:tab or self::w:br]', $ns)) {
            if ($n.LocalName -eq 'br') {
                if ($curI) { [void]$sb.Append('[[/i]]'); $curI = $false }
                if ($curB) { [void]$sb.Append('[[/b]]'); $curB = $false }
                [void]$sb.Append("`n")
                continue
            }
            $piece = if ($n.LocalName -eq 'tab') { "`t" } else { $n.InnerText }
            if ($piece -eq '') { continue }

            $rPr = $n.ParentNode.SelectSingleNode('w:rPr', $ns)
            $isI = Test-Flag $rPr 'i'
            $isB = Test-Flag $rPr 'b'

            # on ne ferme puis rouvre que si la mise en forme change reellement
            if ($curB -and -not $isB) { [void]$sb.Append('[[/b]]'); $curB = $false }
            if ($curI -and -not $isI) { [void]$sb.Append('[[/i]]'); $curI = $false }
            if ($isI -and -not $curI) { [void]$sb.Append('[[i]]');  $curI = $true }
            if ($isB -and -not $curB) { [void]$sb.Append('[[b]]');  $curB = $true }

            [void]$sb.Append($piece)
        }
        if ($curI) { [void]$sb.Append('[[/i]]') }
        if ($curB) { [void]$sb.Append('[[/b]]') }

        $paragraphs += $sb.ToString()
    }
    ($paragraphs -join "`n`n")
}

function Get-MarkCounts {
    <# Comptage des balises, pour verifier qu'elles survivent a la traduction. #>
    param([Parameter(Mandatory)][string]$Text)
    @{
        IOpen  = ([regex]::Matches($Text, '\[\[i\]\]')).Count
        IClose = ([regex]::Matches($Text, '\[\[/i\]\]')).Count
        BOpen  = ([regex]::Matches($Text, '\[\[b\]\]')).Count
        BClose = ([regex]::Matches($Text, '\[\[/b\]\]')).Count
    }
}

function Test-MarksIntact {
    <# Vrai si la sortie porte autant de balises que l'entree, et qu'elles sont
       appariees. Sert de critere de relance au meme titre que les paragraphes. #>
    param([Parameter(Mandatory)][string]$Before, [Parameter(Mandatory)][string]$After)
    $a = Get-MarkCounts -Text $Before
    $b = Get-MarkCounts -Text $After
    ($a.IOpen -eq $b.IOpen) -and ($a.IClose -eq $b.IClose) -and
    ($a.BOpen -eq $b.BOpen) -and ($a.BClose -eq $b.BClose) -and
    ($b.IOpen -eq $b.IClose) -and ($b.BOpen -eq $b.BClose)
}

function Remove-Marks {
    param([Parameter(Mandatory)][string]$Text)
    $Text -replace '\[\[/?[ib]\]\]', ''
}

function Update-ContinuitySheet {
    <# Fiche de continuite pour les textes longs : sur plusieurs blocs, les
       60 derniers mots du bloc precedent ne suffisent plus. La fiche retient les
       personnages, les traductions deja fixees et l'etat du recit, et repart dans
       le prompt de chaque bloc suivant.
       Volontairement courte : elle est reinjectee a chaque appel. #>
    param(
        [string]$Current,
        [Parameter(Mandatory)][string]$SourceChunk,
        [Parameter(Mandatory)][string]$TranslatedChunk,
        [Parameter(Mandatory)][string]$Model,
        [Parameter(Mandatory)][string]$Endpoint,
        [int]$MaxWords = 220
    )

    $sys = @'
Tu tiens a jour une fiche de continuite pour une traduction FR->EN en cours.
On te donne la fiche actuelle, puis le passage source et sa traduction.

Rends la fiche MISE A JOUR, en anglais, sous ce format exact et rien d'autre :

CHARACTERS: un par ligne - nom, genre (he/she/they), role en trois mots
TERMS: un par ligne - terme francais = traduction anglaise retenue
       (noms propres, lieux, objets, surnoms, jargon recurrent)
PLOT: trois phrases maximum sur ou en est le recit
VOICE: personne narrative, temps verbal, registre

Regles :
- CONSERVE toutes les entrees existantes, ajoute seulement les nouvelles.
- Ne corrige une entree que si le passage la contredit franchement.
- Pas de commentaire, pas de recopie du texte, pas de titre.
'@

    $user = "=== FICHE ACTUELLE ===`n"
    $user += if ($Current) { $Current } else { "(vide, premier passage)" }
    $user += "`n`n=== PASSAGE SOURCE (francais) ===`n$SourceChunk"
    $user += "`n`n=== SA TRADUCTION (anglais) ===`n$TranslatedChunk"

    try {
        $r = Invoke-LMStudioChat -System $sys -User $user -Model $Model -Endpoint $Endpoint `
                -Temperature 0.2 -MaxTokens ($MaxWords * 4)
        $t = (Remove-TemplateEcho -Text $r.Text).Text.Trim()
        if ($t) {
            # Plafond applique pour de vrai : demande dans le prompt, la limite
            # est ignoree et la fiche enfle jusqu'a noyer le prompt de traduction.
            $w = $t -split '\s+' | Where-Object { $_ }
            if ($w.Count -gt $MaxWords) { $t = ($w | Select-Object -First $MaxWords) -join ' ' }
            return @{ Sheet = $t; Tokens = $r.Tokens }
        }
    }
    catch {
        Write-Warning "  fiche de continuite non mise a jour : $($_.Exception.Message)"
    }
    @{ Sheet = $Current; Tokens = 0 }
}

function Write-DocxFromText {
    <# Ecrit un .docx minimal mais valide (OOXML) a partir de texte brut :
       un paragraphe separe par une ligne vide = un <w:p>, les sauts de ligne
       internes deviennent des <w:br/>. Pas de mise en forme (ni italique ni gras) :
       la sortie des scripts de traduction est du texte brut. #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Text
    )

    Add-Type -AssemblyName System.IO.Compression | Out-Null
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

    $W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<w:document xmlns:w="' + $W + '"><w:body>')

    foreach ($p in (Get-Paragraphs -Text $Text)) {
        [void]$sb.Append('<w:p>')
        $lines = $p -split '\r?\n'
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($i -gt 0) { [void]$sb.Append('<w:r><w:br/></w:r>') }

            # Les balises [[i]] / [[b]] redeviennent de vrais attributs de run.
            $italic = $false; $bold = $false
            foreach ($tok in [regex]::Split($lines[$i], $script:MarkRegex)) {
                switch ($tok) {
                    '[[i]]'  { $italic = $true;  continue }
                    '[[/i]]' { $italic = $false; continue }
                    '[[b]]'  { $bold   = $true;  continue }
                    '[[/b]]' { $bold   = $false; continue }
                }
                if ($tok -eq '' -or $tok -match '^\[\[/?[ib]\]\]$') { continue }

                $rPr = ''
                if ($italic -or $bold) {
                    $rPr = '<w:rPr>'
                    if ($bold)   { $rPr += '<w:b/>' }
                    if ($italic) { $rPr += '<w:i/>' }
                    $rPr += '</w:rPr>'
                }
                $esc = [System.Security.SecurityElement]::Escape($tok)
                [void]$sb.Append('<w:r>' + $rPr + '<w:t xml:space="preserve">' + $esc + '</w:t></w:r>')
            }
        }
        [void]$sb.Append('</w:p>')
    }
    [void]$sb.Append('<w:sectPr/></w:body></w:document>')

    $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>'

    $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '</Relationships>'

    $dir = Split-Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    if (Test-Path $Path) { Remove-Item $Path -Force }

    $zip = [System.IO.Compression.ZipFile]::Open($Path, 'Create')
    try {
        foreach ($part in @(
            @{ Name = '[Content_Types].xml'; Body = $contentTypes },
            @{ Name = '_rels/.rels';         Body = $rels },
            @{ Name = 'word/document.xml';   Body = $sb.ToString() }
        )) {
            $entry = $zip.CreateEntry($part.Name)
            $stream = $entry.Open()
            try {
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($part.Body)
                $stream.Write($bytes, 0, $bytes.Length)
            }
            finally { $stream.Dispose() }
        }
    }
    finally { $zip.Dispose() }
}

function Write-Utf8NoBom {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Content)
    [System.IO.File]::WriteAllText($Path, $Content, (New-Object System.Text.UTF8Encoding($false)))
}
