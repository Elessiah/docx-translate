# Chaîne de traduction FR→EN pour documents Word

Traduire des textes longs de façon **fiable** : sans perdre de paragraphes, sans
casser la structure, sans que le gras et l'italique disparaissent en route.

Deux moteurs interchangeables :

- **API DeepL** — rapide, précise, gratuite jusqu'à 500 000 caractères par mois
- **Modèle local** via LM Studio — entièrement hors ligne, sans quota ni compte

Le tout en PowerShell, sans dépendance à installer.

---

## Pourquoi

Traduire un document long avec un LLM échoue de façons discrètes : des
paragraphes fusionnent, des phrases disparaissent, la mise en forme s'évapore, un
terme est rendu différemment au chapitre 1 et au chapitre 7. Rien ne le signale —
le fichier de sortie a l'air correct.

Ce projet répond par des **contraintes mécaniques** plutôt que par des consignes
au modèle. À chaque fois qu'une règle pouvait être vérifiée ou réparée par le
code, elle l'est. Le modèle n'est sollicité que là où il apporte quelque chose.

---

## Installation

```powershell
git clone <url-du-depot> E:\AIs
```

**Pour l'API DeepL** — une clé gratuite sur [deepl.com/pro-api](https://www.deepl.com/pro-api) :

```powershell
setx DEEPL_API_KEY "votre-cle"
```

Le script détecte seul s'il s'agit d'une clé gratuite ou payante, et sélectionne
le bon point d'entrée. La clé n'est jamais écrite dans le dépôt.

**Pour le moteur local** — [LM Studio](https://lmstudio.ai), puis un modèle GGUF :

```powershell
lms get "https://huggingface.co/<depot>/<modele>-GGUF@Q4_K_M" -y
lms server start
```

Les scripts démarrent le serveur et chargent le modèle eux-mêmes si nécessaire.

---

## Usage

### Traduction via l'API — le plus simple

```powershell
.\deepl-api.ps1 usage                                   # vérifie la clé et le quota
.\deepl-api.ps1 translate -Path "mon-recit.docx"
```

Sortie : un `.docx` avec la mise en forme, et le `.txt` correspondant.

### Chaîne complète en local

```powershell
.\pipelineDA.ps1 "mon-recit.docx"
```

Enchaîne correction du français, mise en forme de la version française,
traduction, révision, mise en forme de la version anglaise.

### Relire un texte sans le laisser modifier

```powershell
.\proofread.ps1 -Path "mon-recit.docx" -ReportOnly -ContextFile "contexte.md"
```

Ne touche pas au fichier. Produit un `_reperage.md` : une liste de fautes
possibles, chacune avec une chaine a coller dans Ctrl+F pour retrouver le
passage. C'est vous qui tranchez.

Deux moteurs alimentent ce rapport : le modèle local, et **LanguageTool**, un
correcteur à base de règles qui tourne lui aussi sur la machine. Le second est
utilisé par défaut ; `-NoLanguageTool` s'en passe.

### La passe rapide, sans modèle

```powershell
.\languagetool.ps1 -Install       # une seule fois, environ 300 Mo
.\proofread.ps1 -Path "mon-recit.docx" -ReportOnly -NoModel
```

LanguageTool seul : une trentaine de secondes pour un texte court, aucun modèle
chargé, aucune carte graphique sollicitée. Il attrape les accords, la
conjugaison et la ponctuation, et **rate** tout ce qui demande de comprendre
l'histoire — le genre d'une narratrice, qui parle dans une incise.

C'est la vérification qu'on lance en cours d'écriture. La passe complète avec le
modèle reste pour la relecture finale.

**Rien ne démarre avec Windows.** Aucun service n'est installé, aucune entrée
n'est ajoutée au démarrage : le serveur est un `java.exe` lancé au début du
contrôle et tué à la fin. `.\languagetool.ps1 -Status` le confirme, `-Stop`
coupe un serveur resté en mémoire.

### Scripts individuels

| Script | Rôle |
|---|---|
| `deepl-api.ps1` | Traduction par l'API DeepL, avec glossaire et contexte |
| `deepl-roundtrip.ps1` | Découpe pour le copier-coller DeepL gratuit, puis restaure la mise en forme |
| `translate.ps1` | Traduction par modèle local, bloc par bloc |
| `review.ps1` | Révise une traduction existante en la confrontant à l'original |
| `proofread.ps1` | Corrige la langue du texte source, ou la signale sans rien reecrire (`-ReportOnly`) |
| `format-deviantart.ps1` | Prépare le texte pour un éditeur web (une ligne vide entre paragraphes) |
| `lib-lmstudio.ps1` | Fonctions communes — lecture `.docx`, découpage, réparations |
| `stats.ps1` | Consommation cumulée des modèles locaux |
| `languagetool.ps1` | Installe, démarre et arrête le correcteur local LanguageTool |

### Suivre la consommation

Chaque exécution ajoute une ligne à `token-usage.csv`. Rien à activer.

```powershell
.\stats.ps1
```

Totaux cumulés, répartition par tâche et par mois. Le journal est local et
exclu du dépôt : il contient vos chemins de travail.

Le comptage se contente de lire le champ `usage` que le serveur renvoie déjà
dans chaque réponse — aucune requête supplémentaire, rien de changé dans ce qui
est envoyé au modèle. L'enregistrement et l'écriture du journal sont enfermés
dans des `try` : une comptabilité d'agrément n'a pas à faire échouer une
traduction de quarante minutes parce que le CSV était ouvert dans Excel.

Les tokens d'entrée sont comptés séparément des tokens générés, et l'écart
surprend : en traitement paragraphe par paragraphe, le prompt système et la
fiche de contexte repartent à chaque appel et pèsent plus lourd que tout ce que
le modèle écrit.

---

## Ce qui rend la sortie fiable

**Alignement garanti.** Via l'API, les paragraphes partent en lot et reviennent
dans l'ordre : la correspondance est structurelle. En local, le nombre de
paragraphes est compté à l'entrée et à la sortie de chaque bloc ; en cas
d'écart, le bloc est relancé, et c'est la **meilleure** tentative qui est
retenue, jamais la dernière.

**Mise en forme transportée.** Le gras et l'italique deviennent des balises qui
traversent la traduction, puis redeviennent de vrais attributs Word. Via l'API,
`tag_handling=xml` les replace exactement.

**Marqueurs de dialogue préservés.** Word encode souvent les tirets de dialogue
comme des puces de liste — ils n'existent pas dans le texte du fichier. La
lecture les restitue depuis `numbering.xml`, et un contrôle final vérifie que la
traduction reproduit la convention de la source.

**Découpage conscient de la structure.** Un bloc ne traverse jamais un
séparateur de scène ni un titre de chapitre : le modèle lit ces marqueurs comme
une fin de tâche et s'arrête d'y traduire.

**Terminologie tenue.** Un glossaire DeepL applique vos équivalences imposées à
la source. En local, une fiche de continuité accumule personnages, genres et
termes déjà traduits, et repart dans chaque bloc suivant.

**Repères vérifiés.** En mode signalement, chaque citation du modèle est
recherchée dans le fichier avant d'entrer dans le rapport. Ce qui ne s'y
retrouve pas part dans une annexe, signalé comme non vérifiable — un modèle
cite volontiers un passage qu'il a reformulé, et un repère faux coûte plus
cher qu'un oubli.

**Typographie française rétablie.** Apostrophes courbes et espaces insécables
sont restaurés après coup — les modèles les dégradent malgré la consigne.

---

## Ce que les mesures ont montré

Résultats obtenus sur des textes réels, pas des suppositions.

### Le mode réflexion : oui pour comparer, non pour produire

| Tâche | Sans réflexion | Avec réflexion |
|---|---|---|
| Traduction | **36/36 paragraphes** | 13/36 |
| Correction | **69/69**, 12 min | 69/66, 20 min |
| Révision | rate les omissions | **les repère** |

La règle, vérifiée trois fois : réflexion **oui** quand le modèle doit *comparer*
deux textes, **non** quand il doit *reproduire un texte entier*.

### Les ruptures de chapitre coupent la traduction

Un bloc à cheval sur un séparateur rendait **6 paragraphes au lieu de 20**, deux
fois de suite. À correctif isolé :

| | Avant | Après |
|---|---|---|
| Blocs réussis | 2/7 | **17/19** |
| Paragraphes conservés | 103/126 | **125/126** |

### Faire réviser une traduction par le modèle : abandonné

Sur un chapitre de 76 paragraphes, soumettre une traduction DeepL à la révision
du modèle local a produit une sortie que le garde-fou a déclarée valide — 76
paragraphes en entrée, 76 en sortie — alors que **29 paragraphes sur 76 étaient
la source française recopiée**, avec deux lignes de service du modèle insérées
comme paragraphes et un décalage d'alignement derrière elles.

Deux blocs en écart, l'un à `+1` et l'autre à `-1`, s'étaient annulés.

La leçon porte moins sur le modèle que sur le contrôle : il comptait les
paragraphes, mais ne vérifiait ni la langue de sortie, ni la présence de
méta-commentaire, ni l'écart bloc par bloc plutôt que cumulé. **Un contrôle qui
mesure la mauvaise grandeur ne protège de rien** — et il rassure, ce qui est
pire que pas de contrôle du tout.

### Signaler plutôt que corriger

Le correcteur qui réécrit produit aussi de fausses corrections, et les plus
coûteuses visent le style : restituer un `ne` de négation dans un dialogue oral,
basculer un verbe au passé simple au milieu d'un récit au présent. Ces dégâts-là
ne se voient pas à la relecture rapide.

D'où le mode `-ReportOnly` : le modèle n'écrit plus le texte, seulement une
liste. Il ne peut donc rien casser, et le nombre de paragraphes cesse d'être un
enjeu. Le problème se déplace sur l'exactitude des repères, qui est vérifiable
par le code — donc traitée par le code.

Mesure sur un extrait dont les fautes étaient connues à l'avance :

| Configuration | Fautes réelles pointées |
|---|---|
| Par blocs, 1 passe | 3 / 8 |
| Par blocs, 3 passes | 6 / 8 |
| Paragraphe par paragraphe + fiche de contexte, 2 passes | **8 / 8** |

Deux choses font la différence : examiner **un seul paragraphe à la fois**, ses
voisins fournis en lecture seule, et donner une **fiche de contexte** indiquant
le genre des personnages et le temps du récit. Sans elle le modèle devine, et
signale comme fautifs des accords corrects.

Les corrections *proposées*, elles, restent souvent fausses. Vérification sur
un document entier de 36 paragraphes : 31 signalements, dont **zéro citation
introuvable** — tous les repères tombaient juste. Mais le modèle a proposé trois
fois une forme masculine pour une narratrice qu'il venait lui-même de décrire
comme féminine dans son motif. Le repère est fiable, la correction ne l'est pas.
Le rapport le dit en tête : servez-vous du repère, jugez la correction.

### Deux moteurs qui ne ratent pas les mêmes choses

LanguageTool applique des règles de grammaire. Il ne devine pas une position :
il la **renvoie**. Aucun ancrage à reconstruire, aucune citation à vérifier —
et donc aucune possibilité qu'un repère soit faux.

Sur le même document de 36 paragraphes :

| | Signalements | Durée | Position |
|---|---|---|---|
| Modèle local, 2 passes | 31 | 38 min | citation à retrouver, vérifiée par le code |
| LanguageTool | 7 | 30 s | rendue par l'outil, exacte par construction |

Les sept de LanguageTool étaient tous dans les trente-et-un du modèle : sur ce
texte il n'apporte rien de neuf, il confirme vite. L'intérêt est ailleurs — ses
sept-là sont les plus sûrs du lot, et ils coûtent trente secondes.

Deux règles sont désactivées par défaut : celles qui restituent le `ne` de
négation. `« je peux plus »`, `« j'ai pas »` sont des choix d'auteur dans un
dialogue, pas des fautes — c'est exactement la correction que ce mode existe
pour ne plus subir.

### API DeepL contre modèle local

Sur un chapitre de 1 796 mots : **1 seconde** contre 319 s. Sur la précision,
DeepL l'emporte nettement — un échantillon de 16 paragraphes donne 9 rendus
supérieurs contre 3, et zéro erreur franche contre deux pour le modèle local.

L'intérêt du moteur local est ailleurs : aucun quota, aucun compte, aucune donnée
qui sort de la machine.

---

## Limites connues

- La mise en forme Word au-delà du gras et de l'italique n'est pas transportée.
- Les `.pdf` doivent être convertis en `.docx` au préalable.
- Le correcteur local produit aussi de **fausses** corrections, y compris sur
  le style. Préférez `-ReportOnly`, qui ne modifie rien.
- En mode signalement, le repère est fiable, la correction proposée ne l'est pas.
- Une relecture humaine reste nécessaire, en particulier sur les idiomes.

---

## Personnalisation

Les prompts système sont dans `prompts/`. Pour les adapter sans toucher au code
ni au dépôt, créez un fichier `<nom>.local.txt` à côté : il prend la priorité et
n'est pas versionné.

```
prompts/translate.txt          # version par defaut, versionnee
prompts/translate.local.txt    # votre version, ignoree par git
```

Un glossaire se crée depuis un fichier `.tsv` de lignes `terme<TAB>traduction` :

```powershell
.\deepl-api.ps1 glossary -GlossaryFile "glossaire.tsv"
```

> Attention aux entrées ambiguës. Un glossaire remplace le terme partout, sans
> regarder le sens : `bas → stockings` transformerait « en **bas** de l'escalier ».
> Préférez les termes composés et les mots sans autre acception.

---

## Environnement

Windows PowerShell 5.1. Les `.ps1` sont volontairement en **pur ASCII** : PowerShell
lit un `.ps1` sans BOM comme de l'ANSI, et un caractère typographique écrit en
clair y serait corrompu en silence. Les caractères spéciaux passent par
`[char]0x2014` ou `\u2014` en expression régulière.
