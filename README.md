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

### Scripts individuels

| Script | Rôle |
|---|---|
| `deepl-api.ps1` | Traduction par l'API DeepL, avec glossaire et contexte |
| `deepl-roundtrip.ps1` | Découpe pour le copier-coller DeepL gratuit, puis restaure la mise en forme |
| `translate.ps1` | Traduction par modèle local, bloc par bloc |
| `review.ps1` | Révise une traduction existante en la confrontant à l'original |
| `proofread.ps1` | Corrige la langue du texte source avant traduction |
| `format-deviantart.ps1` | Prépare le texte pour un éditeur web (une ligne vide entre paragraphes) |
| `lib-lmstudio.ps1` | Fonctions communes — lecture `.docx`, découpage, réparations |

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
- Le correcteur local produit aussi de **fausses** corrections : le rapport
  `_notes.md` doit être relu, ce n'est pas un correcteur automatique de confiance.
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
