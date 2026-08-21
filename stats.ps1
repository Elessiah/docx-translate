<#
.SYNOPSIS
  Consommation cumulee des modeles locaux, lue dans le journal des executions.

.DESCRIPTION
  Chaque execution de translate.ps1, review.ps1 et proofread.ps1 ajoute une
  ligne a token-usage.csv. Ce script la relit et n'affiche que des totaux :
  il ne contacte aucun serveur et ne modifie rien.

  Les tokens d'entree sont comptes a part des tokens generes. L'ecart est
  instructif : en traitement paragraphe par paragraphe, le prompt systeme et
  la fiche de contexte repartent a chaque appel et pesent souvent plus lourd
  que tout ce que le modele ecrit.

.EXAMPLE
  .\stats.ps1

.EXAMPLE
  .\stats.ps1 -Detail
#>

param(
    [string]$Path,
    [switch]$Detail
)

$ErrorActionPreference = "Stop"
if (-not $Path) { $Path = Join-Path $PSScriptRoot 'token-usage.csv' }

if (-not (Test-Path $Path)) {
    Write-Host "Aucun journal pour l'instant : $Path" -ForegroundColor DarkGray
    Write-Host "Il se remplit tout seul a chaque execution des scripts." -ForegroundColor DarkGray
    return
}

$rows = @(Import-Csv -Path $Path -Encoding UTF8)
if (-not $rows.Count) {
    Write-Host "Journal vide." -ForegroundColor DarkGray
    return
}

# Les nombres sont ecrits avec le separateur decimal local : on repasse par
# InvariantCulture apres avoir normalise la virgule, sinon les secondes sont
# lues de travers d'une machine a l'autre.
function ConvertTo-Num {
    param([string]$Value)
    if (-not $Value) { return 0.0 }
    $v = $Value -replace ',', '.'
    [double]::Parse($v, [System.Globalization.CultureInfo]::InvariantCulture)
}

$totIn   = 0.0; $totOut = 0.0; $totSec = 0.0; $totCalls = 0
foreach ($r in $rows) {
    $totIn    += ConvertTo-Num $r.Entree
    $totOut   += ConvertTo-Num $r.Generes
    $totSec   += ConvertTo-Num $r.Secondes
    $totCalls += [int](ConvertTo-Num $r.Appels)
}
$total = $totIn + $totOut

$span = [TimeSpan]::FromSeconds($totSec)
$duree = "{0:N0} h {1:00} min" -f [Math]::Floor($span.TotalHours), $span.Minutes
if ($span.TotalHours -lt 1) { $duree = "{0:N0} min" -f $span.TotalMinutes }

Write-Host ""
Write-Host "  CONSOMMATION CUMULEE" -ForegroundColor Cyan
Write-Host "  --------------------" -ForegroundColor DarkGray
Write-Host ("  {0,14:N0} tokens au total" -f $total) -ForegroundColor Yellow
Write-Host ("  {0,14:N0} en entree" -f $totIn) -ForegroundColor DarkGray
Write-Host ("  {0,14:N0} generes par le modele" -f $totOut) -ForegroundColor DarkGray
Write-Host ""
Write-Host ("  {0,14} executions, {1:N0} appels" -f $rows.Count, $totCalls) -ForegroundColor DarkGray
Write-Host ("  {0,14} de calcul" -f $duree) -ForegroundColor DarkGray
if ($totSec -gt 0) {
    Write-Host ("  {0,14:N1} tokens generes par seconde" -f ($totOut / $totSec)) -ForegroundColor DarkGray
}

# --------------------------------------------------------------- par tache
Write-Host ""
Write-Host "  PAR TACHE" -ForegroundColor Cyan
Write-Host "  ---------" -ForegroundColor DarkGray
$rows | Group-Object Tache | Sort-Object { -($_.Group | ForEach-Object { ConvertTo-Num $_.Total } | Measure-Object -Sum).Sum } | ForEach-Object {
    $sum = ($_.Group | ForEach-Object { ConvertTo-Num $_.Total } | Measure-Object -Sum).Sum
    $pct = 0
    if ($total -gt 0) { $pct = [int](100 * $sum / $total) }
    $bar = '#' * [Math]::Max(1, [int]($pct / 4))
    Write-Host ("  {0,-12} {1,12:N0}  {2,3} %  {3}" -f $_.Name, $sum, $pct, $bar) -ForegroundColor Yellow
}

# --------------------------------------------------------------- par mois
Write-Host ""
Write-Host "  PAR MOIS" -ForegroundColor Cyan
Write-Host "  --------" -ForegroundColor DarkGray
$rows | Group-Object { ($_.Date -split ' ')[0].Substring(0, 7) } | Sort-Object Name | ForEach-Object {
    $sum = ($_.Group | ForEach-Object { ConvertTo-Num $_.Total } | Measure-Object -Sum).Sum
    Write-Host ("  {0,-12} {1,12:N0}  ({2} execution(s))" -f $_.Name, $sum, $_.Group.Count) -ForegroundColor Yellow
}

# --------------------------------------------------------------- records
$biggest = $rows | Sort-Object { -(ConvertTo-Num $_.Total) } | Select-Object -First 1
Write-Host ""
Write-Host ("  Plus gros traitement : {0:N0} tokens - {1} du {2}" -f `
    (ConvertTo-Num $biggest.Total), $biggest.Tache, $biggest.Date) -ForegroundColor DarkGray

$cut = @($rows | Where-Object { [int](ConvertTo-Num $_.Tronquees) -gt 0 })
if ($cut.Count) {
    Write-Host ("  {0} execution(s) avec des reponses coupees par max_tokens" -f $cut.Count) -ForegroundColor DarkYellow
}

if ($Detail) {
    Write-Host ""
    Write-Host "  DETAIL" -ForegroundColor Cyan
    Write-Host "  ------" -ForegroundColor DarkGray
    $rows | Select-Object -Last 20 | ForEach-Object {
        Write-Host ("  {0}  {1,-11} {2,10:N0}  {3}" -f `
            $_.Date, $_.Tache, (ConvertTo-Num $_.Total),
            (Split-Path $_.Source -Leaf)) -ForegroundColor DarkGray
    }
}
Write-Host ""
