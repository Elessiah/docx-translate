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

# ============================================================================
#  Comptabilite des tokens.
#  On se contente de lire le champ "usage" que le serveur renvoie deja dans
#  chaque reponse : aucune requete supplementaire, aucun changement dans ce qui
#  est envoye au modele. Le suivi ne peut donc pas influencer la generation.
# ============================================================================

$script:TokenLedger = $null

function Reset-TokenLedger {
    param([string]$Label = '')
    $script:TokenLedger = [ordered]@{
        Label      = $Label
        Calls      = 0
        Prompt     = 0
        Completion = 0
        Total      = 0
        Truncated  = 0
        Seconds    = 0.0
        MaxPrompt  = 0
        Started    = (Get-Date)
    }
}

function Add-TokenLedgerEntry {
    <# Appelee par Invoke-LMStudioChat apres chaque aller-retour. #>
    param(
        [int]$Prompt = 0,
        [int]$Completion = 0,
        [double]$Seconds = 0,
        [switch]$WasTruncated
    )
    if (-not $script:TokenLedger) { Reset-TokenLedger }
    $l = $script:TokenLedger
    $l.Calls++
    $l.Prompt     += $Prompt
    $l.Completion += $Completion
    $l.Total      += ($Prompt + $Completion)
    $l.Seconds    += $Seconds
    if ($Prompt -gt $l.MaxPrompt) { $l.MaxPrompt = $Prompt }
    if ($WasTruncated) { $l.Truncated++ }
}

function Get-TokenLedger {
    if (-not $script:TokenLedger) { Reset-TokenLedger }
    $script:TokenLedger
}

function Write-TokenSummary {
    <# Resume lisible en fin de script. Les tokens d'entree sont separes des
       tokens generes : en traitement paragraphe par paragraphe, le prompt
       systeme et la fiche de contexte sont renvoyes a chaque appel et pesent
       souvent plus lourd que tout ce que le modele ecrit. #>
    param([string]$Prefix = '')

    try { $l = Get-TokenLedger } catch { return }
    if (-not $l -or -not $l.Calls) { return }

    $genRate = 0.0
    if ($l.Seconds -gt 0) { $genRate = $l.Completion / $l.Seconds }
    $share = 0
    if ($l.Total -gt 0) { $share = [int](100 * $l.Prompt / $l.Total) }

    Write-Host ""
    Write-Host ($Prefix + ("Tokens : {0:N0} au total sur {1} appel(s)" -f $l.Total, $l.Calls)) -ForegroundColor Yellow
    Write-Host ($Prefix + ("  entree  : {0,9:N0}  ({1} % du total, pointe a {2:N0} par appel)" -f `
        $l.Prompt, $share, $l.MaxPrompt)) -ForegroundColor DarkGray
    Write-Host ($Prefix + ("  generes : {0,9:N0}  ({1:N1} tok/s)" -f $l.Completion, $genRate)) -ForegroundColor DarkGray
    if ($l.Truncated) {
        Write-Host ($Prefix + ("  ATTENTION : {0} reponse(s) coupee(s) par max_tokens" -f $l.Truncated)) -ForegroundColor Red
    }
}

function Write-TokenLog {
    <# Ajoute une ligne au journal cumulatif, pour suivre la consommation d'une
       execution a l'autre. Le fichier contient des chemins de travail : il est
       exclu du depot. #>
    param(
        [string]$Path,
        [string]$Task = '',
        [string]$Model = '',
        [string]$Source = '',
        [string]$Note = ''
    )
    try { $l = Get-TokenLedger } catch { return }
    if (-not $l -or -not $l.Calls) { return }
    if (-not $Path) { $Path = Join-Path $PSScriptRoot 'token-usage.csv' }

    $row = [pscustomobject][ordered]@{
        Date       = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Tache      = $Task
        Modele     = $Model
        Source     = $Source
        Appels     = $l.Calls
        Entree     = $l.Prompt
        Generes    = $l.Completion
        Total      = $l.Total
        Secondes   = [Math]::Round($l.Seconds, 1)
        Tronquees  = $l.Truncated
        Note       = $Note
    }
    # Le fichier peut etre ouvert ailleurs au moment ou on ecrit. On abandonne
    # la ligne plutot que de faire echouer un traitement deja termine.
    try {
        $exists = Test-Path $Path
        $csv = $row | ConvertTo-Csv -NoTypeInformation
        if ($exists) { $csv = $csv | Select-Object -Skip 1 }
        Add-Content -Path $Path -Value $csv -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-Host "  (journal de tokens non ecrit : $($_.Exception.Message))" -ForegroundColor DarkGray
    }
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

    $callWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $r = Invoke-RestMethod -Uri $Endpoint -Method Post `
            -ContentType "application/json; charset=utf-8" `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -TimeoutSec 3600
    $callWatch.Stop()

    $txt = $r.choices[0].message.content
    # Qwen3 peut emettre un bloc de raisonnement malgre /no_think
    $txt = [regex]::Replace($txt, '(?s)<think>.*?</think>', '').Trim()

    # finish_reason = "length" : la reponse a ete COUPEE par max_tokens. C'est la
    # seule facon de distinguer une omission du modele d'une troncature technique.
    $finish = $r.choices[0].finish_reason

    # Lecture seule de ce que le serveur a deja renvoye : le suivi n'ajoute
    # aucune requete et ne modifie pas ce qui est envoye au modele.
    # Enferme dans un try : une comptabilite d'agrement n'a pas le droit de
    # faire echouer une traduction, et ce bloc est sur le chemin de CHAQUE appel.
    try {
        Add-TokenLedgerEntry -Prompt ([int]$r.usage.prompt_tokens) `
            -Completion ([int]$r.usage.completion_tokens) `
            -Seconds $callWatch.Elapsed.TotalSeconds `
            -WasTruncated:($finish -eq 'length')
    }
    catch { }

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

# ============================================================================
#  Reperage sans reecriture : le modele signale, il ne corrige pas.
#  Chaque signalement est ancre mecaniquement dans le texte source. Ce qui ne
#  s'y retrouve pas est marque comme tel, jamais presente comme un fait : un
#  repere faux coute plus cher a l'auteur qu'un oubli.
# ============================================================================

function ConvertTo-MatchKey {
    <# Forme normalisee d'un texte pour la recherche, avec la table de
       correspondance vers les positions d'origine. Neutralise la casse, les
       apostrophes courbes, les guillemets et les blancs multiples : le modele
       recopie rarement une citation au caractere pres. #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    $sb  = New-Object System.Text.StringBuilder
    $map = New-Object System.Collections.Generic.List[int]
    $prevSpace = $true
    for ($i = 0; $i -lt $Text.Length; $i++) {
        $c = $Text[$i]
        if ([char]::IsWhiteSpace($c)) {
            if ($prevSpace) { continue }
            [void]$sb.Append(' '); $map.Add($i); $prevSpace = $true
            continue
        }
        $prevSpace = $false
        $lc = [char]::ToLowerInvariant($c)
        switch ([int]$lc) {
            0x2019 { $lc = [char]0x27 }
            0x02BC { $lc = [char]0x27 }
            0x201C { $lc = [char]0x22 }
            0x201D { $lc = [char]0x22 }
            0x2013 { $lc = [char]0x2D }
            0x2014 { $lc = [char]0x2D }
        }
        [void]$sb.Append($lc); $map.Add($i)
    }
    @{ Key = $sb.ToString(); Map = $map }
}

function Measure-KeyOccurrences {
    <# Nombre d'endroits du texte ou la chaine apparait, comparaison normalisee. #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Needle,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Paragraphs
    )
    $nk = (ConvertTo-MatchKey -Text $Needle).Key.Trim()
    if (-not $nk) { return 0 }
    $n = 0
    foreach ($p in $Paragraphs) {
        $pk = (ConvertTo-MatchKey -Text $p).Key
        $pos = 0
        while ($true) {
            $j = $pk.IndexOf($nk, $pos)
            if ($j -lt 0) { break }
            $n++; $pos = $j + 1
        }
    }
    $n
}

function Resolve-FindingAnchor {
    <# Retrouve un fragment cite par le modele dans le texte source. C'est le
       garde-fou central du mode reperage : ce qui ne se retrouve pas ici n'a
       pas ete lu dans le texte, donc n'est pas un repere utilisable.
       HintFrom / HintTo restreignent au bloc traite, pour ne pas ancrer sur
       une occurrence identique situee ailleurs. #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Fragment,
        [Parameter(Mandatory)][string[]]$Paragraphs,
        [int]$HintFrom = -1,
        [int]$HintTo = -1
    )
    $fk = (ConvertTo-MatchKey -Text $Fragment).Key.Trim()
    if ($fk.Length -lt 3) { return @{ Found = $false } }

    $hits = @()
    for ($i = 0; $i -lt $Paragraphs.Count; $i++) {
        $k = ConvertTo-MatchKey -Text $Paragraphs[$i]
        $pos = 0
        while ($true) {
            $j = $k.Key.IndexOf($fk, $pos)
            if ($j -lt 0) { break }
            $st = $k.Map[$j]
            $li = [Math]::Min($j + $fk.Length - 1, $k.Map.Count - 1)
            $en = $k.Map[$li]
            $hits += @{ Para = $i; Start = $st; Length = ($en - $st + 1) }
            $pos = $j + 1
        }
    }
    if (-not $hits.Count) { return @{ Found = $false } }

    $pick = $hits[0]
    if ($HintFrom -ge 0) {
        $inChunk = @($hits | Where-Object { $_.Para -ge $HintFrom -and $_.Para -le $HintTo })
        if ($inChunk.Count) { $pick = $inChunk[0] }
    }
    @{ Found = $true; Para = $pick.Para; Start = $pick.Start
       Length = $pick.Length; Count = $hits.Count }
}

function Get-UniqueSearchString {
    <# Chaine a taper dans Ctrl+F : le fragment tel qu'il est reellement ecrit
       dans le fichier, elargi aux mots voisins jusqu'a ne designer qu'un seul
       endroit du texte. #>
    param(
        [Parameter(Mandatory)][string]$Paragraph,
        [Parameter(Mandatory)][int]$Start,
        [Parameter(Mandatory)][int]$Length,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$AllParagraphs
    )
    $s = $Start; $e = $Start + $Length
    while ($s -gt 0 -and [char]::IsLetterOrDigit($Paragraph[$s - 1])) { $s-- }
    while ($e -lt $Paragraph.Length -and [char]::IsLetterOrDigit($Paragraph[$e])) { $e++ }

    for ($round = 0; $round -lt 6; $round++) {
        $cand = $Paragraph.Substring($s, $e - $s).Trim()
        if ((Measure-KeyOccurrences -Needle $cand -Paragraphs $AllParagraphs) -le 1) { return $cand }
        if (($e - $s) -gt 90) { return $cand }
        $s = [Math]::Max(0, $s - 14)
        $e = [Math]::Min($Paragraph.Length, $e + 14)
        while ($s -gt 0 -and [char]::IsLetterOrDigit($Paragraph[$s - 1])) { $s-- }
        while ($e -lt $Paragraph.Length -and [char]::IsLetterOrDigit($Paragraph[$e])) { $e++ }
    }
    $Paragraph.Substring($s, $e - $s).Trim()
}

function Get-AnchorExcerpt {
    <# Extrait du paragraphe avec le fragment mis en gras, pour reconnaitre le
       passage sans ouvrir le fichier. #>
    param(
        [Parameter(Mandatory)][string]$Paragraph,
        [Parameter(Mandatory)][int]$Start,
        [Parameter(Mandatory)][int]$Length,
        [int]$Around = 70
    )
    $s = [Math]::Max(0, $Start - $Around)
    $e = [Math]::Min($Paragraph.Length, $Start + $Length + $Around)
    while ($s -gt 0 -and [char]::IsLetterOrDigit($Paragraph[$s - 1])) { $s-- }
    while ($e -lt $Paragraph.Length -and [char]::IsLetterOrDigit($Paragraph[$e])) { $e++ }

    $pre  = $Paragraph.Substring($s, $Start - $s)
    $hit  = $Paragraph.Substring($Start, $Length)
    $post = $Paragraph.Substring($Start + $Length, $e - ($Start + $Length))
    $txt = $pre + '**' + $hit + '**' + $post
    if ($s -gt 0) { $txt = '[...]' + $txt }
    if ($e -lt $Paragraph.Length) { $txt = $txt + '[...]' }
    ($txt -replace '\s+', ' ').Trim()
}

function ConvertFrom-FindingLines {
    <# Lit les signalements du modele. Une par ligne :
         fragment exact ||| correction proposee ||| motif
       Une ligne hors format est conservee brute et marquee : en mode reperage
       rien n'est jete en silence, c'est l'auteur qui tranche. #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    $trimChars = [char[]]@([char]0x22, [char]0x201C, [char]0x201D, [char]0x27,
                           [char]0x2019, [char]0x60, ' ')
    $out = @()
    foreach ($line in ($Text -split '\r?\n')) {
        $l = $line.Trim()
        $l = $l -replace '^\s*[-*]\s+', ''
        $l = $l -replace '^\s*\d+[.)]\s+', ''
        if (-not $l) { continue }
        if ($l -match '^#') { continue }
        if ($l -match '^(RAS|R\.A\.S|aucune?|rien a signaler)\b') { continue }

        $parts = $l -split '\s*\|\|\|\s*'
        if ($parts.Count -ge 3) {
            # Le modele repond parfois RAS dans la colonne du motif tout en
            # remplissant les deux autres : la ligne a la forme d'un
            # signalement mais n'en est pas un.
            $why = (($parts[2..($parts.Count - 1)]) -join ' ').Trim()
            if ($why -match '^(RAS|R\.A\.S|aucune?|rien)\b') { continue }
            $out += @{
                Fragment   = $parts[0].Trim($trimChars)
                Suggestion = $parts[1].Trim($trimChars)
                Reason     = $why
                Raw        = $l
            }
        }
        else {
            $out += @{ Fragment = ''; Suggestion = ''; Reason = ''; Raw = $l }
        }
    }
    , $out
}

function Find-MechanicalIssues {
    <# Fautes reperables sans modele. Sur ces categories la position est exacte
       et il n'y a pas d'oubli possible : autant ne pas les confier a un LLM. #>
    param([Parameter(Mandatory)][string[]]$Paragraphs)

    # mots qui se redoublent legitimement en francais
    $twice = 'nous|vous|si|tres|bien|tout|toute|non|oui|ha|ah|oh|eh|hi|ho'
    $out = @()

    for ($i = 0; $i -lt $Paragraphs.Count; $i++) {
        $p = $Paragraphs[$i]

        foreach ($m in [regex]::Matches($p, '\b(\p{L}{2,})(\s+)\1\b', 'IgnoreCase')) {
            if ($m.Groups[1].Value -match "^($twice)$") { continue }
            $out += @{ Para = $i; Start = $m.Index; Length = $m.Length
                       Fragment = $m.Value; Suggestion = $m.Groups[1].Value
                       Reason = 'mot repete a l identique' }
        }
        foreach ($m in [regex]::Matches($p, '\s+,')) {
            $out += @{ Para = $i; Start = $m.Index; Length = $m.Length
                       Fragment = $m.Value; Suggestion = ','
                       Reason = 'espace avant la virgule' }
        }
        foreach ($m in [regex]::Matches($p, '(?<![.\d])\s+\.(?!\.)')) {
            $out += @{ Para = $i; Start = $m.Index; Length = $m.Length
                       Fragment = $m.Value; Suggestion = '.'
                       Reason = 'espace avant le point' }
        }
        foreach ($m in [regex]::Matches($p, '(?<=\S)  +(?=\S)')) {
            $out += @{ Para = $i; Start = $m.Index; Length = $m.Length
                       Fragment = $m.Value; Suggestion = ' '
                       Reason = 'espaces multiples' }
        }

        $op = ([regex]::Matches($p, '\(')).Count
        $cl = ([regex]::Matches($p, '\)')).Count
        if ($op -ne $cl) {
            $out += @{ Para = $i; Start = 0; Length = [Math]::Min(40, $p.Length)
                       Fragment = ''; Suggestion = ''
                       Reason = "parentheses non appariees : $op ouvrante(s), $cl fermante(s)" }
        }
    }
    , $out
}

function Measure-TypographyGaps {
    <# Compte les ecarts typographiques sans les lister un par un : ils se
       comptent par centaines et noieraient les vraies fautes. Le mode
       correction les repare mecaniquement. #>
    param([Parameter(Mandatory)][string]$Text)
    $NBSP = [string][char]0x00A0
    @{
        StraightApostrophes = ([regex]::Matches($Text, "(?<=\p{L})'(?=\p{L})")).Count
        MissingNbsp         = ([regex]::Matches($Text, "(?<=\S)(?<![$NBSP])[?!;](?=\s|`$)")).Count
    }
}

function Join-Reasons {
    <# Fusionne deux motifs de signalement sans repeter le meme deux fois.
       Deux passes rendent souvent le meme motif a l'accent pres : on compare
       sans accent ni casse, sinon le rapport affiche des doublons. #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$First,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Second
    )
    $seen = @{}
    $out  = @()
    foreach ($chunk in @($First, $Second)) {
        foreach ($r in ($chunk -split '\s+/\s+')) {
            $t = $r.Trim()
            if (-not $t) { continue }
            $flat = $t.Normalize([System.Text.NormalizationForm]::FormD)
            $sb = New-Object System.Text.StringBuilder
            foreach ($ch in $flat.ToCharArray()) {
                if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne
                    [System.Globalization.UnicodeCategory]::NonSpacingMark) {
                    [void]$sb.Append($ch)
                }
            }
            $key = ($sb.ToString().ToLowerInvariant() -replace '[^a-z0-9]', '')
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true
            $out += $t
        }
    }
    $out -join ' / '
}

# ============================================================================
#  LanguageTool : correcteur a base de regles, execute en local.
#
#  Complementaire du modele, pas concurrent. Il ne comprend pas l'histoire,
#  donc il rate tout ce qui demande de savoir qui parle ou de quel genre est
#  la narratrice. En echange il donne la position exacte de ce qu'il trouve,
#  sans jamais rien reecrire, et en une seconde au lieu d'une demi-heure.
#
#  Le serveur ne tourne QUE pendant un controle. Rien n'est installe dans
#  Windows, aucun service, rien au demarrage de la machine : c'est un java.exe
#  lance a la demande et arrete des que le controle est fini.
# ============================================================================

$script:LTStartedByUs   = $false
$script:LTServerProcess = $null

# Regles ecartees par defaut. Deux motifs seulement :
#   - la negation orale ("je peux plus", "j'ai pas") est un choix d'auteur,
#     pas une faute ; c'est exactement la correction qu'on refuse de subir ;
#   - le reste fait doublon avec ce que le rapport compte deja en masse, ou
#     se declenche sur chaque tiret de dialogue et noierait le reste.
$script:LTDefaultDisabled = @(
    'NEGATION_PLUS', 'P_V_PAS', 'NE_MANQUANT', 'FR_NE_MANQUANT',
    'UPPERCASE_SENTENCE_START',
    'WHITESPACE_RULE', 'FRENCH_WHITESPACE', 'FRENCH_WHITESPACE_STRICT',
    'APOS_TYP', 'TYPOGRAPHIC_APOSTROPHES', 'APOS_M'
)

function Get-LanguageToolInstall {
    <# Cherche le serveur et son java dans tools\ a cote des scripts. A defaut
       du java livre avec le projet, accepte celui du PATH s'il y en a un. #>
    param([string]$Root)
    if (-not $Root) { $Root = $PSScriptRoot }

    $ltDir = Join-Path $Root 'tools\languagetool'
    $jar   = Join-Path $ltDir 'languagetool-server.jar'
    $java  = Join-Path $Root 'tools\jre\bin\java.exe'

    if (-not (Test-Path $java)) {
        $cmd = Get-Command java.exe -ErrorAction SilentlyContinue
        if ($cmd) { $java = $cmd.Source } else { $java = $null }
    }

    @{
        Dir  = $ltDir
        Jar  = $jar
        Java = $java
        Ok   = ((Test-Path $jar) -and $java -and (Test-Path $java))
    }
}

function Test-LanguageToolServer {
    param([string]$Endpoint = 'http://localhost:8081/v2')
    try {
        $null = Invoke-RestMethod -Uri "$Endpoint/languages" -TimeoutSec 3 -ErrorAction Stop
        return $true
    }
    catch { return $false }
}

function Start-LanguageToolServer {
    <# Rend $true si un serveur repond a la fin, qu'on l'ait demarre ou non.
       On note si c'est nous qui l'avons lance : on n'arretera que le notre. #>
    param(
        [string]$Endpoint = 'http://localhost:8081/v2',
        [int]$TimeoutSeconds = 90
    )
    if (Test-LanguageToolServer -Endpoint $Endpoint) { return $true }

    $inst = Get-LanguageToolInstall
    if (-not $inst.Ok) { return $false }

    $port = 8081
    if ($Endpoint -match ':(\d+)') { $port = [int]$Matches[1] }

    try {
        $p = Start-Process -FilePath $inst.Java -PassThru -WindowStyle Hidden -WorkingDirectory $inst.Dir -ArgumentList @('-cp', 'languagetool-server.jar', 'org.languagetool.server.HTTPServer', '--port', "$port")
    }
    catch { return $false }

    $script:LTServerProcess = $p
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        Start-Sleep -Milliseconds 400
        if (Test-LanguageToolServer -Endpoint $Endpoint) {
            $script:LTStartedByUs = $true
            return $true
        }
        if ($p.HasExited) { return $false }
    }
    return $false
}

function Stop-LanguageToolServer {
    <# N'arrete que le serveur demarre par ce script : si un serveur tournait
       deja pour autre chose, on ne le coupe pas. -Force passe outre. #>
    param([switch]$Force)
    if (-not $Force -and -not $script:LTStartedByUs) { return }
    try {
        if ($script:LTServerProcess -and -not $script:LTServerProcess.HasExited) {
            Stop-Process -Id $script:LTServerProcess.Id -Force -ErrorAction Stop
        }
    }
    catch { }
    $script:LTStartedByUs   = $false
    $script:LTServerProcess = $null
}

function ConvertTo-LanguageToolReason {
    <# Ramene les categories de LanguageTool sur le vocabulaire du rapport,
       pour que les deux sources se lisent de la meme facon.
       Les categories francaises ne portent pas les memes noms que les
       categories generiques (CAT_GRAMMAIRE, PONCTUATION_VIRGULE...), d'ou la
       reconnaissance par motif. Une categorie inconnue garde son propre nom :
       mieux vaut un libelle inhabituel qu'une information perdue. #>
    param([string]$Category, [string]$Message)

    $c = ''
    switch -Regex ($Category) {
        'HOMONYM|PARONYM|CONFUSED'          { $c = 'homophone';   break }
        'AGREEMENT|ACCORD'                  { $c = 'accord';      break }
        'CONJUGAISON'                       { $c = 'conjugaison'; break }
        'GRAMMAIRE|^GRAMMAR'                { $c = 'grammaire';   break }
        '^TYPOS|ORTHOGRAPH'                 { $c = 'orthographe'; break }
        'PONCTUATION|^PUNCTUATION'          { $c = 'ponctuation'; break }
        'MAJUSCULE|^CASING'                 { $c = 'majuscule';   break }
        'TYPOGRAPH|INSECABLE|TRAITS_UNION'  { $c = 'typographie'; break }
        'REPETITION|^REDUNDANCY'            { $c = 'repetition';  break }
    }
    if (-not $c) {
        $c = ($Category -replace '^CAT_', '' -replace '_', ' ').ToLowerInvariant()
        if (-not $c) { $c = 'doute' }
    }

    $m = ($Message -replace '\s+', ' ').Trim()
    $c + ' : ' + $m
}

function Invoke-LanguageToolCheck {
    <# Un envoi par paragraphe : les positions rendues sont alors directement
       des index dans le paragraphe, sans recalcul et sans ancrage a verifier.
       C'est toute la difference avec un modele : ici la position est donnee,
       pas devinee. #>
    param(
        [Parameter(Mandatory)][string[]]$Paragraphs,
        [string]$Endpoint = 'http://localhost:8081/v2',
        [string]$Language = 'fr',
        [string]$Level = 'picky',
        [string[]]$DisabledRules
    )
    if ($null -eq $DisabledRules) { $DisabledRules = $script:LTDefaultDisabled }

    $out = @()
    for ($i = 0; $i -lt $Paragraphs.Count; $i++) {
        $p = $Paragraphs[$i]
        if (-not $p.Trim()) { continue }

        # Tout est encode en pourcent ici : le corps envoye reste en ASCII pur,
        # ce qui evite la question de l'encodage du body en PowerShell 5.1.
        $body = 'language=' + [System.Uri]::EscapeDataString($Language) + '&level=' + [System.Uri]::EscapeDataString($Level) + '&text=' + [System.Uri]::EscapeDataString($p)
        if ($DisabledRules -and $DisabledRules.Count) {
            $body += '&disabledRules=' + [System.Uri]::EscapeDataString(($DisabledRules -join ','))
        }

        # LanguageTool repond en application/json SANS charset. PowerShell 5.1
        # decode alors en Latin-1 et rend des accents doubles : on lit les
        # octets bruts et on impose UTF-8 nous-memes.
        try {
            $w = Invoke-WebRequest -Uri "$Endpoint/check" -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded' -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
            $r = [System.Text.Encoding]::UTF8.GetString($w.RawContentStream.ToArray()) | ConvertFrom-Json
        }
        catch { continue }

        foreach ($m in $r.matches) {
            $start = [int]$m.offset
            $len   = [int]$m.length
            if ($start -lt 0 -or ($start + $len) -gt $p.Length) { continue }

            $sug = ''
            if ($m.replacements) {
                $sug = (@($m.replacements | Select-Object -First 2 | ForEach-Object { $_.value }) -join ' / ')
            }
            $out += @{
                Para       = $i
                Start      = $start
                Length     = $len
                Fragment   = $p.Substring($start, $len)
                Suggestion = $sug
                Reason     = (ConvertTo-LanguageToolReason -Category $m.rule.category.id -Message $m.message)
                Rule       = [string]$m.rule.id
            }
        }
    }
    , $out
}

function Remove-Diacritics {
    <# Enleve les accents sans toucher aux lettres. Sert partout ou l'on
       compare des mots ecrits par le modele, qui accentue au hasard. #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)
    if (-not $Text) { return '' }
    $flat = $Text.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $flat.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne
            [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    $sb.ToString()
}

# ============================================================================
#  Accord avec le narrateur.
#
#  Repartition volontaire du travail : le modele DESIGNE, le code TRANCHE.
#
#  Decider quels mots se rapportent au narrateur demande de comprendre le
#  texte : seul le modele sait le faire. Decider si "emmitoufle" est masculin
#  ou feminin est de la morphologie : le code le fait sans se tromper et sans
#  changer d'avis. On ne demande donc jamais au modele de proposer une forme -
#  sur un document de 36 paragraphes il a motive trois fois par "narratrice
#  femme" avant de suggerer un masculin.
#
#  Le fichier reste en ASCII pur : les lettres accentuees des formes proposees
#  sont construites par code, jamais ecrites en clair.
# ============================================================================

$script:AccE = [string][char]0x00E9   # e accent aigu
$script:AccG = [string][char]0x00E8   # e accent grave
$script:AccC = [string][char]0x00E7   # c cedille
$script:AccI = [string][char]0x00EE   # i circonflexe

# Formes verbales de la 1e personne dont la terminaison imite un adjectif.
# Sans cette liste, "je reponds" produirait "reponse".
$script:VerbForms1P = @(
    'suis', 'ai', 'vais', 'sais', 'fais', 'dis', 'mets', 'peux', 'veux',
    'prends', 'reponds', 'attends', 'entends', 'vends', 'tends', 'descends',
    'rends', 'perds', 'mords', 'sens', 'viens', 'tiens', 'reviens',
    'deviens', 'obtiens', 'maintiens', 'soutiens',
    'cours', 'sors', 'dors', 'pars', 'sers', 'meurs', 'souris',
    'crois', 'vois', 'bois', 'dois', 'recois', 'ris', 'ecris', 'lis',
    'connais', 'parais', 'plais', 'tais', 'nais', 'mens', 'sens',
    'etais', 'avais', 'faisais', 'disais', 'voulais', 'pouvais', 'devais',
    'serais', 'aurais', 'irais', 'ferais', 'dirais', 'voudrais', 'pourrais',
    'es', 'est', 'as', 'a', 'vas', 'fus', 'eus'
)

# Mots-outils. Le modele n'est pas cense les citer, mais une regle de
# morphologie appliquee a "mon" produit "monne" : on ferme la porte.
$script:FunctionWords = @(
    'rien', 'lui', 'moi', 'toi', 'soi', 'eux', 'mon', 'ton', 'son',
    'mes', 'tes', 'ses', 'nos', 'vos', 'leur', 'leurs', 'des', 'les',
    'ils', 'nous', 'vous', 'qui', 'quoi', 'dont', 'mais', 'donc', 'car',
    'dans', 'sur', 'sous', 'vers', 'avec', 'sans', 'pour', 'par', 'chez',
    'entre', 'apres', 'avant', 'depuis', 'pendant', 'cela', 'ceci', 'ces',
    'cet', 'quel', 'quels', 'quand', 'comme', 'ici', 'puis', 'encore'
)

# Invariables en genre dont la terminaison imiterait un masculin.
$script:GenderInvariable = @(
    'mieux', 'pis', 'plus', 'moins', 'bien', 'mal', 'vite', 'ainsi',
    'debout', 'ensemble', 'expres', 'tot', 'tard', 'trop', 'tres',
    'tout', 'tous', 'chacun', 'meme', 'aussi', 'assez', 'jamais',
    'toujours', 'parfois', 'soudain', 'enfin', 'alors', 'depuis'
)

function Get-FeminineIrregular {
    <# Table des feminins que les regles generales rateraient. Construite en
       fonction pour que les accents restent hors du source. #>
    $t = @{}
    $t['faux']    = 'fausse';   $t['doux']   = 'douce'
    $t['roux']    = 'rousse';   $t['vieux']  = 'vieille'
    $t['beau']    = 'belle';    $t['nouveau']= 'nouvelle'
    $t['fou']     = 'folle';    $t['mou']    = 'molle'
    $t['blanc']   = 'blanche';  $t['franc']  = 'franche'
    $t['sec']     = 's' + $script:AccG + 'che'
    $t['public']  = 'publique'; $t['long']   = 'longue'
    $t['frais']   = 'fra' + $script:AccI + 'che'
    $t['favori']  = 'favorite'; $t['gentil'] = 'gentille'
    $t['nul']     = 'nulle';    $t['bas']    = 'basse'
    $t['gros']    = 'grosse'
    $t['epais']   = $script:AccE + 'paisse'
    $t['las']     = 'lasse';    $t['gras']   = 'grasse'
    $t['metis']   = 'm' + $script:AccE + 'tisse'
    $t['complet'] = 'compl' + $script:AccG + 'te'
    $t['secret']  = 's' + $script:AccE + 'cr' + $script:AccG + 'te'
    $t['inquiet'] = 'inqui' + $script:AccG + 'te'
    $t['discret'] = 'discr' + $script:AccG + 'te'
    $t['concret'] = 'concr' + $script:AccG + 'te'
    $t['sot']     = 'sotte';    $t['idiot']  = 'idiote'
    $t['mort']    = 'morte';    $t['malin']  = 'maligne'
    $t['benin']   = 'b' + $script:AccE + 'nigne'
    $t['fier']    = 'fi' + $script:AccG + 're'
    $t['tiers']   = 'tierce'
    $t['pret']    = 'pr' + [string][char]0x00EA + 'te'
    # -teur fait -teuse par defaut ; ceux-ci font -trice.
    foreach ($n in @('acteur', 'directeur', 'lecteur', 'spectateur',
                     'createur', 'auteur', 'narrateur', 'seducteur',
                     'conducteur', 'instituteur', 'protecteur',
                     'inspecteur', 'traducteur', 'animateur',
                     'observateur', 'organisateur', 'formateur')) {
        $t[$n] = ($n -replace 'teur$', 'trice')
    }
    $t
}
$script:FeminineIrregular = Get-FeminineIrregular

function Restore-WordCase {
    <# Rend $Value avec la casse de $Model : le rapport cite le mot tel qu'il
       apparait dans le texte, majuscule comprise. #>
    param([Parameter(Mandatory)][string]$Model, [Parameter(Mandatory)][string]$Value)
    if ($Model -cmatch '^\p{Lu}') {
        return ($Value.Substring(0, 1).ToUpperInvariant() + $Value.Substring(1))
    }
    $Value
}

function Get-FeminineForm {
    <# Forme feminine d'un mot masculin, ou chaine vide quand il n'y a rien a
       dire : mot deja feminin, invariable en genre, ou forme verbale. Le
       silence est le comportement par defaut : on ne signale que du sur.

       Les motifs sont testes sur la version SANS accent, mais la forme rendue
       est construite sur le mot d'origine : "emmitoufle" finit par un e accent
       aigu, qu'aucune classe de caracteres ASCII ne reconnait. #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Word)

    $w = $Word.Trim()
    if (-not $w) { return '' }
    if ($w -notmatch '^\p{L}+$') { return '' }

    $lower = $w.ToLowerInvariant()
    $plain = (Remove-Diacritics -Text $lower)

    # Un mot de deux lettres n'est jamais une cible d'accord qui vaille
    # un signalement - sauf "nu", qui en est une vraie.
    if ($plain.Length -lt 3 -and $plain -ne 'nu')  { return '' }
    if ($script:VerbForms1P -contains $plain)      { return '' }
    if ($script:GenderInvariable -contains $plain) { return '' }
    if ($script:FunctionWords -contains $plain)    { return '' }

    # Deja feminin, ou epicene : presque tout ce qui finit par un e non
    # accentue l'est. "emmitouflee", "calme", "triste", "assise", "fiere".
    if ($lower -match 'e$' -or $lower -match 'es$') { return '' }

    if ($script:FeminineIrregular.ContainsKey($plain)) {
        return (Restore-WordCase -Model $w -Value $script:FeminineIrregular[$plain])
    }

    $fem = ''
    switch -Regex ($plain) {
        'eur$'   { $fem = $lower -replace '(?i)eur$',  'euse';  break }
        'eux$'   { $fem = $lower -replace '(?i)eux$',  'euse';  break }
        'if$'    { $fem = $lower -replace '(?i)if$',   'ive';   break }
        'er$'    { $fem = $lower -replace '(?i)er$',   ($script:AccG + 're'); break }
        'ien$'   { $fem = $lower + 'ne';  break }
        'on$'    { $fem = $lower + 'ne';  break }
        'el$'    { $fem = $lower + 'le';  break }
        'eil$'   { $fem = $lower + 'le';  break }
        'et$'    { $fem = $lower + 'te';  break }
        'x$'     { $fem = $lower -replace '(?i)x$', 'se'; break }
        # Participes, adjectifs et noms ordinaires : le gros du contingent.
        # Le test porte sur $plain, donc un e accentue final compte comme e.
        default  {
            if ($plain -match '(e|i|u|t|d|l|n|r|c|s)$') { $fem = $lower + 'e' }
        }
    }

    if (-not $fem -or $fem -eq $lower) { return '' }
    Restore-WordCase -Model $w -Value $fem
}

function Get-MasculineForm {
    <# Chemin inverse, volontairement plus prudent : on ne retire une marque
       de feminin que la ou la regle ne laisse pas de place au doute. #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Word)

    $w = $Word.Trim()
    if (-not $w -or $w -notmatch '^\p{L}+$') { return '' }
    $lower = $w.ToLowerInvariant()
    $plain = (Remove-Diacritics -Text $lower)
    if ($plain.Length -lt 3)                    { return '' }
    if ($script:VerbForms1P -contains $plain)   { return '' }
    if ($script:FunctionWords -contains $plain) { return '' }

    $masc = ''
    switch -Regex ($plain) {
        'trice$' { $masc = $lower -replace '(?i)trice$', 'teur'; break }
        'euse$'  { $masc = $lower -replace '(?i)euse$',  'eux';  break }
        'ienne$' { $masc = $lower -replace '(?i)ienne$', 'ien';  break }
        'onne$'  { $masc = $lower -replace '(?i)onne$',  'on';   break }
        'ette$'  { $masc = $lower -replace '(?i)ette$',  'et';   break }
        'elle$'  { $masc = $lower -replace '(?i)elle$',  'el';   break }
        'ive$'   { $masc = $lower -replace '(?i)ive$',   'if';   break }
        'ee$'    { $masc = $lower.Substring(0, $lower.Length - 1); break }
        'ie$'    { $masc = $lower.Substring(0, $lower.Length - 1); break }
        'ue$'    { $masc = $lower.Substring(0, $lower.Length - 1); break }
        default  { return '' }
    }
    if (-not $masc -or $masc -eq $lower) { return '' }
    Restore-WordCase -Model $w -Value $masc
}

function Test-NarratorAgreement {
    <# Verdict sur un mot designe par le modele, connaissant le genre du
       narrateur. Rend @{ Mismatch; Expected; Why }.

       Mismatch faux ne veut pas dire "correct" : le plus souvent cela veut
       dire "rien de verifiable ici". La difference compte, et c'est pour
       cela qu'on ne rend jamais un simple booleen sans motif. #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Word,
        [Parameter(Mandatory)][ValidateSet('feminin', 'masculin')][string]$Gender
    )

    $none = @{ Mismatch = $false; Expected = ''; Why = '' }
    $w = $Word.Trim()
    if (-not $w) { return $none }

    if ($Gender -eq 'feminin') {
        # Une forme feminine existe et differe : le mot est donc au masculin.
        $fem = Get-FeminineForm -Word $w
        if (-not $fem) { return $none }
        return @{
            Mismatch = $true
            Expected = $fem
            Why      = 'accord narrateur : forme masculine, narratrice au feminin'
        }
    }

    $masc = Get-MasculineForm -Word $w
    if (-not $masc) { return $none }
    @{
        Mismatch = $true
        Expected = $masc
        Why      = 'accord narrateur : forme feminine, narrateur au masculin'
    }
}

function Resolve-NarratorGender {
    <# Cherche le genre du narrateur dans la fiche de contexte. On ne devine
       pas : sans mention explicite on rend une chaine vide, et l'appelant
       demandera. Lancer un controle d'accord sur un genre suppose ferait
       exactement les degats qu'on cherche a eviter. #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Context)
    if (-not $Context) { return '' }
    $c = (Remove-Diacritics -Text $Context).ToLowerInvariant()

    if ($c -match 'narrat\w*\s*:?\s*(une\s+)?(femme|feminin|fille)')  { return 'feminin' }
    if ($c -match 'narrat\w*\s*:?\s*(un\s+)?(homme|masculin|garcon)') { return 'masculin' }
    if ($c -match 'genre\s*(du|de la)?\s*narrat\w*\s*:?\s*f')         { return 'feminin' }
    if ($c -match 'genre\s*(du|de la)?\s*narrat\w*\s*:?\s*(m|h)')     { return 'masculin' }
    if ($c -match '\bnarratrice\b')                                   { return 'feminin' }
    if ($c -match 'narrat\w*\s+(est\s+)?(une\s+)?femme')              { return 'feminin' }
    if ($c -match 'narrat\w*\s+(est\s+)?(un\s+)?homme')               { return 'masculin' }
    ''
}

# ============================================================================
#  Artefacts de dictee vocale.
#
#  Un texte dicte ne contient pas de fautes d'orthographe : le logiciel ecrit
#  toujours des mots corrects. Ce qu'il produit, ce sont des mots JUSTES au
#  mauvais endroit. Deux familles se reperent sans modele, et celles-la
#  n'ont aucune raison d'etre confiees a un modele.
# ============================================================================

function Find-DictationArtifacts {
    <# Ce qu'une dictee laisse derriere elle et qu'une regexp voit mieux qu'un
       modele : les commandes de dictee restees dans le texte, et les
       fragments dits deux fois de suite. #>
    param([Parameter(Mandatory)][string[]]$Paragraphs)

    # Commandes de dictee. Uniquement les formules a plusieurs mots : "virgule"
    # ou "point" seuls sont des mots ordinaires du francais et les signaler
    # noierait le rapport.
    $commands = @(
        'nouveau paragraphe', 'nouvelle ligne', 'a la ligne',
        'retour a la ligne', 'point final', 'point virgule',
        'point d.interrogation', 'point d.exclamation',
        'ouvrez les guillemets', 'fermez les guillemets',
        'ouvrir les guillemets', 'fermer les guillemets',
        'entre parentheses', 'nouveau chapitre'
    )

    $out = @()
    for ($i = 0; $i -lt $Paragraphs.Count; $i++) {
        $p = $Paragraphs[$i]
        $flat = Remove-Diacritics -Text $p

        foreach ($c in $commands) {
            foreach ($m in [regex]::Matches($flat, '\b' + $c + '\b', 'IgnoreCase')) {
                # Les index de la version sans accents coincident avec ceux du
                # texte : enlever un accent ne change pas le nombre de lettres.
                $out += @{
                    Para = $i; Start = $m.Index; Length = $m.Length
                    Fragment = $p.Substring($m.Index, $m.Length)
                    Suggestion = ''
                    Reason = 'commande : formule de dictee restee dans le texte'
                }
            }
        }

        # Fragment repete a l'identique juste apres lui-meme, trois mots ou
        # plus. En dessous de trois, la repetition peut etre un effet de style.
        $rx = '\b((?:\p{L}[\p{L}' + [char]0x2019 + "'" + ']*(?:\s+|$)){3,8})\1'
        foreach ($m in [regex]::Matches($p, $rx, 'IgnoreCase')) {
            $out += @{
                Para = $i; Start = $m.Index; Length = $m.Length
                Fragment = $m.Value
                Suggestion = $m.Groups[1].Value.Trim()
                Reason = 'repetition : le meme fragment est dit deux fois de suite'
            }
        }
    }
    , $out
}

function Get-NarratorCandidates {
    <# Tous les mots d'un paragraphe dont la forme ne correspond PAS au genre
       du narrateur. C'est la liste des fautes possibles, etablie par la seule
       morphologie : si un mot n'y est pas, il ne peut pas etre un desaccord.

       La recall ne depend donc plus du modele. Il ne lui restera qu'a dire,
       pour chacun, s'il se rapporte au narrateur ou a quelqu'un d'autre -
       une question fermee, la seule qu'il traite de facon fiable. #>
    param(
        [Parameter(Mandatory)][string]$Paragraph,
        [Parameter(Mandatory)][ValidateSet('feminin', 'masculin')][string]$Gender
    )

    $out = @()
    foreach ($m in [regex]::Matches($Paragraph, '\p{L}+')) {
        $w = $m.Value
        # Les adverbes en -ment ne s'accordent jamais : autant ne pas les
        # soumettre au modele, chaque candidat en trop est du bruit.
        if ((Remove-Diacritics -Text $w.ToLowerInvariant()) -match 'ment$') { continue }

        $verdict = Test-NarratorAgreement -Word $w -Gender $Gender
        if (-not $verdict.Mismatch) { continue }

        $out += @{
            Word     = $w
            Start    = $m.Index
            Length   = $m.Length
            Expected = $verdict.Expected
            Why      = $verdict.Why
        }
    }
    , $out
}
