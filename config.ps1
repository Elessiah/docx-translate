# Valeurs par defaut du projet.
# Pour les adapter sans toucher au depot, creez config.local.ps1 a cote :
# il est charge apres celui-ci et n'est pas versionne.
#
#   $script:DefaultModel = "mon-modele-local"
#
$script:DefaultModel    = "local-model"
$script:DefaultEndpoint = "http://localhost:1234/v1/chat/completions"
$script:DefaultContextLength = 20000