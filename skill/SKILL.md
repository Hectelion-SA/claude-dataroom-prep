---
name: dataroom-prep
description: Trie et structure des documents en vrac en dataroom M&A-ready (11 dossiers + sous-dossiers multilingues), avec classification IA, détection des doublons exacts (SHA-256), détection des versions obsolètes, et checklist contextuelle enrichie via scraping site web + sources juridiques/fiscales/comptables (Légifrance, data.gouv, Fedlex). Multilingue FR/EN/DE/IT/ES. Utiliser quand l'utilisateur dit "structure cette dataroom", "trie ces documents", "prépare cette dataroom DD", "/dataroom-prep", ou demande de transformer un dossier en vrac en dataroom M&A.
---

# dataroom-prep — Skill de préparation de dataroom M&A

## Quand utiliser ce skill

Déclencher quand l'utilisateur demande de :
- Trier / organiser / structurer des documents en vrac en dataroom DD
- Préparer une dataroom M&A pré-upload VDR (Ansarada, Datasite, Drooms, etc.)
- Renommer en masse des documents avec convention `Project X - Intitulé du document - yyyymmdd`
- Détecter et archiver les doublons / versions obsolètes d'un dossier
- Générer une checklist des documents à demander client

## Confidentialité & zéro exfiltration (RÈGLE ABSOLUE)

> 🔒 **Les documents et leur contenu ne quittent JAMAIS le PC de l'utilisateur.** Toute la
> persistance se fait en local (workdir + dossier destination) ; **rien n'est envoyé vers un
> serveur externe.**

Règles non négociables :

1. **Traitement local par défaut.** Extraction, classification, dedup, build, Excel mapping et
   le Top 50 des documents manquants sont **100% hors-ligne** (scripts Python, listes DD locales
   dans `data_sources/`). Aucun appel réseau n'est requis pour produire une dataroom complète.
2. **Ne JAMAIS envoyer de contenu de document — ni un fichier, ni un snippet `extracted.json`,
   ni un titre/chemin issu des documents — vers un outil externe** : pas de Firecrawl, pas de
   data.gouv, pas de WebSearch/WebFetch, pas de MCP tiers, pas d'upload. Ces données restent
   strictement dans la session et sur le disque local.
3. **Aucune donnée client dans la mémoire persistante.** Ne rien écrire dans `MEMORY.md` ni dans
   les fichiers `memory/` qui contienne un nom de société cible, un chiffre, un titre de
   document ou tout élément du mandat. La mémoire ne sert qu'aux notes de process génériques.
4. **Seule exception réseau = Étape 8 (checklist enrichie), strictement opt-in.** Cette étape
   appelle `firecrawl_scrape` sur le **site web public** de la société + des sources juridiques
   (data.gouv / Légifrance / Fedlex). Elle n'envoie QUE l'URL publique et des requêtes
   secteur/droit — **jamais** le contenu des documents. Avant de la lancer, prévenir
   l'utilisateur : « cette étape contacte des serveurs externes (le site public de la société +
   sources légales) ; veux-tu la lancer ou rester 100% hors-ligne ? ». Si l'utilisateur veut le
   zéro-réseau total, **sauter l'Étape 8** et se contenter du Top 50 local (étape 5bis), qui
   couvre déjà la liste des documents à demander sans aucune connexion.
5. **Nettoyage des résidus.** Écrire les fichiers de travail (`extracted.json`,
   `decisions.json`) dans un **workdir temporaire isolé**, jamais dans `scripts/`. Ne pas
   laisser traîner de fichier contenant des titres/snippets de documents client dans le dossier
   du skill après le run.

## Pipeline complet (8 étapes)

### Étape 1 — Questions obligatoires (TOUT poser au début, en une seule passe)

> ⚠️ **RÈGLE DE FLUX : poser TOUTES les questions au tout début, avant de lancer le moindre
> script.** L'objectif est de pouvoir enchaîner l'intégralité du pipeline (extraction →
> classification → dedup → build → Excel → checklist enrichie) sans interruption ni nouvelle
> question. **Ne PAS découper en « Phase 1 » / « Phase 2 ».** La décision sur la checklist
> enrichie (et ses sous-questions) est demandée ICI, en amont, pas après le build.

**Utiliser `AskUserQuestion`** pour poser dans cet ordre (regrouper en quelques écrans de
questions multiples ; idéalement tout collecter avant la première action) :

1. **"Où sont tes documents source ?"** — un chemin (dossier ou ZIP). **Question obligatoire,
   toujours posée explicitement** si le chemin n'a pas déjà été donné dans le prompt.
2. **"Où veux-tu créer la dataroom organisée ?"** — chemin destination. **Question obligatoire,
   toujours posée explicitement** si le chemin n'a pas déjà été donné dans le prompt. Proposer
   par défaut un sous-dossier à côté de la source (ex. `<dossier parent source>/<Project name>`).
3. **"Nom du projet ?"** — utilisé pour préfixer les fichiers (ex: "Project Medicaps")
4. **"Dans quelle langue veux-tu la structure de dossiers ?"** — FR / EN / DE / IT / ES
5. **"Dans quelle langue veux-tu le fichier Excel mapping ?"** — FR / EN / DE / IT / ES
6. ~~"Quels noms anonymiser ?"~~ — **NE JAMAIS POSER cette question.** L'anonymisation est
   TOUJOURS active par défaut, sans confirmation (« bien sûr que oui »). Le nom réel de la
   société cible + ses marques/filiales/entités liées sont **détectés automatiquement par
   Claude** depuis le contenu des documents extraits, puis passés à `--scrub` de façon
   strictement interne. Ces noms ne doivent JAMAIS être écrits, affichés, ni dans une question,
   ni dans le chat, ni dans le raisonnement (cf. RÈGLE ABSOLUE plus bas).
7. **"Provenance de la documentation (pays) ?"** — France (FR) / Suisse (CH). Choisit la liste DD
   de référence pour le **Top 50 des documents manquants** (cf. étape 5bis).
8. **"Secteur d'activité ?"** — onglet de la liste DD : BTP / Industrie / Services & Conseil /
   Tech / Distribution / Santé / Immobilier / Hôtellerie-Restauration / Transport & Logistique /
   Agroalimentaire.
9. **"Veux-tu aussi la checklist enrichie des documents à demander au client ?"** — Oui / Non.
   **Posée ici, dès le début** (et non après le build). Si **Oui**, enchaîner immédiatement dans
   la même passe de questions les sous-questions de la checklist enrichie :
   - **Localisation de la société** : France / Suisse / Belgique / Luxembourg / autre
   - **Site web de la société** : URL (sera scrapée pour détecter secteur, taille, marchés, conformités)
   - **Secteur d'activité connu** (optionnel — vide = détection auto via site web)
   - **Sources à consulter** : data.gouv / Légifrance / Fedlex / OpenLaw / aucune

> ⚠️ **PIÈGE À NE PLUS REFAIRE : ne jamais reporter Q1/Q2/Q3 (chemins + nom de projet) à plus
> tard sous prétexte que `AskUserQuestion` exige 2 à 4 options et ne fait pas de texte libre
> nativement.** Ces trois questions sont des chemins/texte libre : elles DOIVENT quand même
> apparaître dans le tout premier envoi de questions, pas dans un second aller-retour. Technique
> à utiliser : donner 2 options factices qui couvrent les cas réels (ex. pour la destination,
> `"À côté de la source"` vs `"Autre emplacement"` — l'utilisateur tape le chemin réel via
> l'option "Other", toujours disponible automatiquement) plutôt que d'omettre la question du
> premier lot. Un lot de questions qui ne contient QUE des choix fermés (langue, pays, secteur…)
> et laisse le chemin source pour « après » est un flux raté — corriger avant d'envoyer.

Si l'utilisateur a déjà précisé tout ou partie de ces infos dans le prompt, ne pas re-demander.
**Une fois toutes les réponses collectées, dérouler le pipeline complet sans poser d'autre
question.**

### Étape 2 — Extraction de texte

Lancer `scripts/extract.py` avec :
- `--source <path_user>` — chemin source
- `--out <skill_workdir>/extracted.json` — workdir temporaire

Le script :
- Scanne récursivement les `.pdf .docx .xlsx .pptx .doc .xls .msg`
- Exclut les artefacts (code, builds, `.cache`, `node_modules`, `_files/`)
- Extrait un snippet de 2-3K caractères par doc
- Sort un JSON avec `relpath`, `filename`, `parent_folder`, `snippet`, `readable`

### Étape 3 — Classification IA (règles + parent_folder)

Lancer `scripts/classify.py` avec :
- `--extracted extracted.json`
- `--language <FR/EN/DE/IT/ES>` — langue de la structure
- `--templates templates/folder_structure_<lang>.json`

Le script :
- Charge le template multilingue
- Pour chaque doc, classe en (folder_N1, theme_N2, confidence) via heuristiques `parent_folder` + `filename` + `snippet`
- LOW confidence → forcé en `11_Document à trier`
- Génère un nouveau nom : `{Project} - {Theme} - {Subject} - {yyyymm}.{ext}`

### Étape 4 — Détection doublons + versions

Lancer `scripts/dedup.py` :
- Hash SHA-256 sur chaque fichier → groupes de doublons exacts
- Clustering par filename normalisé (retire dates, versions, suffixes) → versions
- Garde la version au score le plus élevé (final > signed > v3 > v2 > v1 > draft, et date la plus récente)

### Étape 5 — Build dataroom N1/N2

Lancer `scripts/build_dataroom.py` :
- Crée la structure 11 dossiers N1 dans le destination user
- Sous-dossiers N2 par theme (NDA, Brevets, Bilans annuels, Commercial - Client X, etc.)
- Copie + renomme les UNIQUE dans la dataroom principale
- Archive les DOUBLON_EXACT dans `_98_Doublons exacts/`
- Archive les VERSION_OBSOLETE dans `_99_Versions anciennes/{N1}/{N2}/`

### Étape 6 — Excel mapping multilingue

Lancer `scripts/generate_mapping.py --lang <excel_lang>` :
- Colonnes : `Original | Chemin source | Nouveau nom | Dossier N1 | Sous-dossier N2 | Confiance | Statut dedup | Chemin final`
- Headers traduits selon `excel_lang`
- Charte Hectelion : navy `#182E4E`, cellules input cyan `#72C7E7`, police Calibri 9-11 pt
- 5 onglets : Dashboard | Mapping | Structure dataroom | Doublons & versions | (placeholder) Checklist

### Étape 7 — Branchement checklist enrichie (PAS de nouvelle question)

> ⚠️ **Ne RIEN re-demander ici.** La décision (Oui/Non) et ses 4 sous-réponses (localisation,
> site web, secteur connu, sources) ont déjà été collectées à l'**Étape 1**. Se contenter de
> brancher :
> - Si l'utilisateur a répondu **Non** à la checklist → terminer le pipeline après l'Étape 6.
> - Si l'utilisateur a répondu **Oui** → enchaîner directement l'Étape 8 avec les réponses déjà
>   en main (localisation, URL site web, secteur connu, sources).

### Étape 8 — Enrichissement checklist

Lancer `scripts/enrich_checklist.py` :

1. **Scrape site web** via Firecrawl (`firecrawl_scrape`) → texte de la page d'accueil + about + produits/services
2. **Analyse contextuelle** : utiliser Claude (via prompt structuré) pour identifier :
   - Secteur précis (MedTech, SaaS, immobilier, retail, etc.)
   - Présence internationale (filiales, exports)
   - Effectif estimé
   - Conformités requises (CE, FDA, ISO, RGPD, HIPAA, etc.)
3. **Cross-référence avec sources** :
   - Si France → consulter Légifrance + data.gouv (obligations sectorielles)
   - Si Suisse → consulter Fedlex (admin.ch) + obligations OFEN/FINMA
4. **Génère un Excel "Project X - Liste documents à demander.xlsx"** avec 4 colonnes :
   - Catégorie (par dossier N1/N2)
   - Document/Information demandé
   - Pourquoi (référence légale ou contextuelle)
   - Statut (✓ déjà présent / ⚠ partiel / ✗ à demander)

Le tout sauvegardé au chemin destination user.

## Templates de structure de dossiers multilingues

Chaque langue a son fichier `templates/folder_structure_<lang>.json` avec :
- Noms des 11 dossiers N1
- Mapping des themes N2 fréquents (Bilans annuels, NDA, Brevets, KBIS, etc.)

Langues supportées : FR, EN, DE, IT, ES.
Pour autre langue : utiliser Claude pour traduire les noms à la volée.

## Sources de données externes

| Source | Langue | Type | Statut |
|---|---|---|---|
| Firecrawl (`firecrawl_scrape`) | toutes | Scraping site web | Actif |
| data.gouv (`mcp__datagouv__*`) | FR | Datasets publics français | Actif |
| Légifrance (API publique gratuite) | FR | Lois, codes, jurisprudence | À implémenter (V2) |
| Fedlex (Suisse) | DE/FR/IT | Droit fédéral suisse | À implémenter (V2) |
| OpenLaw | FR | Jurisprudence ouverte | À implémenter (V2) |

Pour V1, focus sur Firecrawl + data.gouv. Les autres en V2.

## Livrables finaux

À la fin du pipeline complet, l'utilisateur trouve dans son dossier destination :

```
<destination>/
├── <Project name>/                                    ← dataroom organisée
│   ├── 01_<Informations générales>/                  (langue choisie)
│   ├── 02_<Finance>/
│   │   ├── <Bilans annuels>/
│   │   ├── <Facturation commerciale>/
│   │   └── ...
│   ├── 03_<Légal>/
│   │   ├── <NDA & confidentialité>/
│   │   ├── <Brevets>/
│   │   └── ...
│   ├── ... (11 dossiers N1)
│   ├── _98_<Doublons exacts>/
│   ├── _99_<Versions anciennes>/
│   ├── _Rapport DataPrep.xlsx                        ← mapping + structure + dedup
│   └── _Liste documents à demander.xlsx              ← si Phase 2 activée
└── <Project name>.zip                                  ← ZIP livrable
```

## Conventions de renommage

Pattern obligatoire : **`{Project} - {Intitulé réel du document} - {yyyymmdd}.{ext}`**

> ⚠️ **RÈGLE ABSOLUE — ANONYMISATION (toujours active, jamais négociée).** Le **nom réel de la
> société cible, de ses marques, filiales et entités liées ne doit JAMAIS apparaître** nulle
> part — y compris :
> - dans les noms de fichiers, sous-dossiers, Excel (y compris colonne « Nom original ») et titres ;
> - **dans tes messages de chat à l'utilisateur** ;
> - **dans tes questions** (`AskUserQuestion` ou autres) ;
> - **dans ton raisonnement / réflexion visible**, y compris pendant l'étape où tu cherches à
>   identifier ce nom pour le scrubber.
>
> **Seul le nom de PROJET est visible, partout** (ex. « la société cible », « Project X »).
> **Ne jamais demander si l'on doit anonymiser** : la réponse est toujours oui, implicitement.
> Lors du titrage IA, décrire le document de façon générique (ex. « entité holding » et non le
> vrai nom).
>
> **Détection automatique + filet technique :** Claude identifie lui-même les noms réels depuis
> le contenu extrait (sans jamais les afficher), puis les passe à `--scrub` de `run_pipeline.py`,
> qui retire automatiquement toute occurrence (insensible casse, limites de mots) des noms
> interdits dans les fichiers/dossiers/Excel générés. Pour éviter que le nom réel transparaisse
> même dans une commande shell affichée, écrire la liste de scrub dans un fichier local non
> commenté (`--scrub-file <chemin>`) plutôt que de l'inliner dans `--scrub "..."` lorsque c'est
> possible.

Quatre règles strictes (édictées par l'utilisateur) :

1. **Ne JAMAIS reprendre le vieux nom de fichier brut.** L'intitulé doit être un titre
   propre déduit du *contenu* du document (lu via extraction / OCR), pas le nom d'origine.
   ❌ `... - PV AG fin rem 08 VD - 2024.pdf`  ❌ `... - StatutsMai2020 - undated.pdf`
   ✅ `... - Procès-verbal AGO cessation rémunération du Président - 20240830.pdf`
2. **Ne JAMAIS répéter le nom du dossier / thème dans le nom de fichier** : l'arborescence
   porte déjà le thème. ❌ `... - Marketing & commercial - Plan marketing ...`
   ✅ `... - Plan marketing 2021-2022.docx`
3. **Date au format `yyyymmdd`** (ou `yyyymm` / `yyyy` si le jour/mois est inconnu, ou rien
   si le document n'a pas de date pertinente).
4. **L'intitulé doit rester COURT et descriptif** — l'essentiel du document en quelques mots
   (viser ~3 à 7 mots, pas de phrase entière ni de sous-clauses). Garder le type de document +
   l'élément qui le distingue des autres du même type, rien de plus.
   ❌ `... - Procès-verbal de l'Assemblée Générale Ordinaire Annuelle statuant sur l'approbation des comptes de l'exercice clos le 31 décembre 2020 - 20210630.pdf`
   ✅ `... - PV AGO approbation des comptes 2020 - 20210630.pdf`

Format final à retenir : **Nom [Projet] + Nom du document (court, descriptif) + yyyymmdd**.

Exemples :
- `Project Hectelion - Extrait Kbis - 20250506.pdf`
- `Project Hectelion - Term Sheet signé - 20220428.pdf`
- `Project Hectelion - Comptes annuels - attestation de conformité - 20201231.pdf`
- `Project Hectelion - Table de capitalisation.xlsx`  (sans date)

Pour les documents en `11_Document à trier` (confiance LOW), garder le nom original sans renommage.

### Mécanisme de titrage par contenu (lecture + OCR)

Le titre propre ne peut pas être deviné depuis le nom de fichier : il faut **comprendre le
contenu**. Le pipeline procède donc en deux temps :

1. **Extraction enrichie** (`run_pipeline.py`) :
   - PDF lus via **PyMuPDF** (couche texte native, bien plus fiable que pypdf).
   - PDF **scannés** (sans couche texte) → **fallback OCR** via `pytesseract` + Pillow
     (rendu image puis OCR `fra+eng`). **Avant toute extraction réelle (mode `--source`),
     le pipeline vérifie automatiquement que `pytesseract`/Pillow/le binaire `tesseract`
     sont présents et, sinon, tente de les installer tout seul** (`pip install pytesseract
     pillow` + `winget install UB-Mannheim.TesseractOCR` sur Windows, `brew install
     tesseract` sur macOS, `apt-get install tesseract-ocr` sur Linux) — silencieux et non
     bloquant : si l'installation échoue (pas d'admin, pas de gestionnaire de paquets, pas
     de réseau), le doc est marqué `SCANNED_NO_OCR` et l'extraction reste gracieuse (pas de
     crash). Ne jamais sauter cette vérification ni la rendre optionnelle : l'objectif est
     qu'un maximum de PDF scannés soient lus, à chaque lancement du skill, sans intervention
     manuelle.
   - Word (`python-docx`) et Excel (`openpyxl`) lus nativement.
2. **Décisions IA** : après extraction, Claude lit les snippets (et, pour les PDF scannés
   non couverts par tesseract, **rasterise la page en PNG via PyMuPDF puis la lit comme
   image** — OCR visuel) et produit un fichier `decisions.json` :
   ```json
   { "<relpath>": { "n1": "03", "n2": "PV Assemblées",
                    "title": "Procès-verbal AGO ...", "date": "20240830",
                    "confidence": "HIGH" } }
   ```
   Le pipeline est relancé avec `--decisions-json` : ces titres/classements **surchargent**
   les heuristiques. Le champ `date` fourni (même vide) est respecté littéralement.

Flux recommandé :
```
python run_pipeline.py --source ... --extract-only extracted.json   # 1. extraire
# 2. Claude lit extracted.json (+ OCR visuel des PDF scannés) -> decisions.json
python run_pipeline.py --source ... --decisions-json decisions.json --make-zip   # 3. build
```

**Dépendances** : `pymupdf` (requis pour l'extraction enrichie + le rendu OCR visuel) ;
`pytesseract` + `Pillow` + binaire `tesseract` pour l'OCR automatique en lot — **auto-installés
par le pipeline lui-même à chaque lancement s'ils manquent** (voir ci-dessus), donc à ne plus
jamais demander à l'utilisateur d'installer manuellement en amont.

## Documents manquants — Top 50 par catégorie (étape 5bis)

Après le tri, le rapport Excel gagne un onglet **« Documents à demander »** : le **Top 50 des
documents importants MANQUANTS**, groupés par catégorie et triés par priorité.

Source : les listes DD de référence Hectelion (issues du SaaS VDR prep), une par pays, avec un
onglet par secteur (~344 exigences chacune, colonne `Priorité` Haute/Moyenne/Basse) :
- `data_sources/Hectelion - Listes DD - FRANCE.xlsx`
- `data_sources/Hectelion - Listes DD - SUISSE.xlsx`

Mécanique (`run_pipeline.py --country FR|CH --sector <onglet>`) :
1. Charge la liste DD du pays + secteur choisis.
2. Marque chaque exigence **présent/manquant** par recouvrement de tokens avec les documents
   réellement classés dans la dataroom (statut « présents (estimés) » — à revérifier).
3. Écrit l'onglet, colonnes `Catégorie | Réf. | Document à demander | Priorité | Période | Base légale`.

**Répartition PROPORTIONNELLE obligatoire** : les 50 slots sont distribués entre catégories
(apportionnement au plus fort reste, pondéré par le nb de manquants de chaque section, **min 1
par catégorie représentée**), au lieu de tout concentrer sur les grosses sections (Finance).
Ex. ~10 Finance, ~7 Juridique, ~6 RH, ~5 Fiscal, ~3 Opérations… À l'intérieur de chaque
catégorie, priorité Haute d'abord. Objectif : une demande client équilibrée couvrant tous les
volets de la DD.

C'est la version **template-driven** de la checklist Phase 2 (plus fiable que le scraping) :
le pays oriente les obligations légales, le secteur la liste métier.

## Charte graphique des livrables Excel

Toujours appliquer la charte Hectelion (skill `hectelion-brand` auto-chargé) :
- Headers : fond `#182E4E`, texte blanc, Calibri 11 pt Bold
- Cellules input : fond cyan `#72C7E7`
- Cellules données : Calibri 9-11 pt, blanc
- Titres éditoriaux d'onglet : Cardo 25 pt / 18 pt
- Statut dedup : MINT vert / WARNING orange / ALERT rouge

## Anti-patterns à éviter

- ❌ **Inventer une classification** quand le document est illisible → forcer `11_Document à trier`
- ❌ **Renommer un doc qui va en 11_À trier** → garder son nom original
- ❌ **Ignorer les doublons exacts** → toujours les détecter et archiver séparément
- ❌ **Mettre tous les docs en racine d'un dossier N1** → toujours créer la structure N2
- ❌ **Hardcoder la langue FR** → respecter le choix utilisateur

## Exemple d'invocation

```
User: /dataroom-prep
Claude: [Pose TOUTES les questions au début via AskUserQuestion :
         Q1 source, Q2 destination, Q3 nom projet, Q4 langue structure,
         Q5 langue Excel, Q7 pays, Q8 secteur, Q9 checklist enrichie Oui/Non
         (+ si Oui : localisation, site web, secteur connu, sources)]
User: [répond à tout]
Claude: [déroule TOUT le pipeline sans interruption :
         extraction → classification → dedup → build dataroom + Excel
         → si checklist=Oui : enrich_checklist.py → scrape + analyse + Excel checklist]
Claude: [livre dataroom + Excel(s) + ZIP]
```

Pour usage automatisé/scriptable, voir `scripts/run_pipeline.py` qui orchestre les 8 étapes.
