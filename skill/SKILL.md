---
name: dataroom-prep
description: Trie et structure des documents en vrac en dataroom M&A-ready (11 dossiers + sous-dossiers multilingues), avec classification IA, détection des doublons exacts (SHA-256), détection des versions obsolètes, et checklist contextuelle enrichie via scraping site web + sources juridiques/fiscales/comptables (Légifrance, data.gouv, Fedlex). Multilingue FR/EN/DE/IT/ES. Utiliser quand l'utilisateur dit "structure cette dataroom", "trie ces documents", "prépare cette dataroom DD", "/dataroom-prep", ou demande de transformer un dossier en vrac en dataroom M&A.
---

# dataroom-prep — Skill de préparation de dataroom M&A

## Quand utiliser ce skill

Déclencher quand l'utilisateur demande de :
- Trier / organiser / structurer des documents en vrac en dataroom DD
- Préparer une dataroom M&A pré-upload VDR (Ansarada, Datasite, Drooms, etc.)
- Renommer en masse des documents avec convention `Project X - Theme - Subject - yyyymm`
- Détecter et archiver les doublons / versions obsolètes d'un dossier
- Générer une checklist des documents à demander client

## Pipeline complet (8 étapes)

### Étape 1 — Questions obligatoires (Phase 1)

**Utiliser `AskUserQuestion`** pour poser dans cet ordre :

1. **"Où sont tes documents source ?"** — un chemin (dossier ou ZIP)
2. **"Où veux-tu créer la dataroom organisée ?"** — chemin destination
3. **"Nom du projet ?"** — utilisé pour préfixer les fichiers (ex: "Project Acme")
4. **"Dans quelle langue veux-tu la structure de dossiers ?"** — FR / EN / DE / IT / ES
5. **"Dans quelle langue veux-tu le fichier Excel mapping ?"** — FR / EN / DE / IT / ES

Si l'utilisateur a déjà précisé tout ou partie de ces infos dans le prompt, ne pas re-demander.

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

### Étape 7 — Question Phase 2 (checklist enrichie)

Demander via `AskUserQuestion` :

> "Veux-tu que je génère la checklist des documents à demander au client (basée sur le contexte société + obligations légales) ?"

Options :
- Oui — génère checklist enrichie (étape 8)
- Non — termine ici

Si **OUI**, demander 3 infos en une seule série de questions :

1. **Localisation de la société** : France / Suisse / Belgique / Luxembourg / autre (auto-complète obligations légales)
2. **Site web de la société** : URL — sera scrapée pour détecter secteur, taille, marchés, conformités spécifiques
3. **Secteur d'activité connu** (optionnel — laisser vide pour détection auto via site web)
4. **Sources à consulter** : data.gouv (FR datasets publics) / Légifrance (lois FR) / Fedlex (lois CH) / OpenLaw / aucune

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

Pattern obligatoire : **`{Project} - {Theme N2} - {Subject} - {yyyymm}.{ext}`**

Exemples :
- `Project Acme - KBIS - 202505.pdf`
- `Project Beta - Term sheets - Investor X - 202404.pdf`
- `Project Acme - Bilans annuels - 2022.pdf`

Pour les documents en `11_Document à trier` (confiance LOW), garder le nom original sans renommage.

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
Claude: [Pose Q1-Q5 via AskUserQuestion]
User: [répond]
Claude: [lance pipeline scripts, livre dataroom + Excel]
Claude: [pose Q6: checklist enrichie ?]
User: Oui
Claude: [pose Q7-Q10: localisation, site web, secteur, sources]
User: [répond]
Claude: [lance enrich_checklist.py → scrape + analyse + Excel checklist]
```

Pour usage automatisé/scriptable, voir `scripts/run_pipeline.py` qui orchestre les 8 étapes.
