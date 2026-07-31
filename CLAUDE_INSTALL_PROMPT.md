# Installation via Claude Code

Ce prompt s'utilise avec **Claude Code** (pas Claude.ai web). Claude Code
exécute les commandes lui-même : tu approuves chaque action en cliquant,
mais tu n'as **rien à copier-coller** côté code.

**Pré-requis utilisateur** :

- Claude Code installé ([claude.com/claude-code](https://claude.com/claude-code))
- Compte Anthropic actif (Claude Pro/Max ou crédits API)
- Windows 10/11, macOS ou Linux
- Python 3.9+ sur PATH
- Git (pour cloner le repo — déjà installé si tu as installé Claude Code)

---

## Comment utiliser

1. Ouvre un terminal (PowerShell, Terminal macOS, ou bash Linux) dans n'importe quel dossier.
2. Lance Claude Code : `claude`
3. Copie tout le bloc ci-dessous **entre les lignes ```** et colle-le dans la conversation.
4. Approuve les commandes au fur et à mesure (Claude Code te demande pour chaque action).
5. Réponds aux 5 questions de config quand elles arrivent.

L'install complète prend 2-3 minutes.

À partir de l'install, tape `/dataroom-prep` dans **n'importe quel projet Claude Code** pour lancer le questionnaire et structurer une dataroom à partir d'un dossier en vrac.

---

```
Je veux installer le skill Claude Code "dataroom-prep" depuis le repo public
https://github.com/Hectelion-SA/claude-dataroom-prep

Voici ce que je voudrais que tu fasses pour moi, étape par étape :

1. Clone le repo dans mon dossier home :
   - Vérifie d'abord si le dossier existe déjà : ~/Github/dataroom-prep
     (Windows : %USERPROFILE%\Github\dataroom-prep)
   - Si oui : cd dedans et fais `git pull`
   - Sinon : clone-le avec `git clone https://github.com/Hectelion-SA/claude-dataroom-prep.git ~/Github/dataroom-prep`
     (Windows : git clone ... %USERPROFILE%\Github\dataroom-prep)

2. Installe les dépendances Python (dont `pymupdf` + `pytesseract` + `pillow` pour l'OCR
   des PDF scannés — voir plus bas) :
   `python -m pip install --quiet pypdf python-docx openpyxl pyyaml pymupdf pytesseract pillow`

3. Lance le script d'installation interactif :
   - Windows (PowerShell) : `& "$env:USERPROFILE\Github\dataroom-prep\install.ps1"`
   - macOS/Linux : `bash ~/Github/dataroom-prep/install.sh`

4. Le script va te poser 5 questions de config :
   - Dossier de sauvegarde par défaut des datarooms
   - Langue de la structure de dossiers (fr/en/de/it/es)
   - Langue de l'Excel mapping (fr/en/de/it/es)
   - Localisation société par défaut (France/Suisse/Belgique/Luxembourg/autre)
   - Branding (firm name + couleur + police, ou "default" pour Hectelion SA)

5. Une fois l'install terminée, confirme-moi avec :
   - Le chemin où le skill a été installé
   - Le chemin du fichier config.yaml créé
   - Que je peux maintenant taper /dataroom-prep dans n'importe quel projet

Important :
- Si une commande échoue, lis l'erreur et propose-moi une correction.
- Ne pousse rien sur Git, ne crée pas de commit, n'envoie rien sur Internet
  à part le git clone et le pip install.
- Ne modifie aucun de mes fichiers existants en dehors de :
   ~/Github/dataroom-prep/ (clone du repo)
   ~/.claude/skills/dataroom-prep/ (skill installé)
```

---

## Si quelque chose se passe mal

### Python pas trouvé

Installe Python 3.9+ depuis [python.org](https://www.python.org/downloads/).
Sur Windows, **coche "Add Python to PATH"** pendant l'installation.

### Erreur lors du git clone

Si tu as une erreur d'authentification GitHub, configure Git d'abord :

```bash
git config --global user.name "Ton Nom"
git config --global user.email "ton@email.com"
```

Le repo étant **public**, aucune authentification n'est nécessaire pour le clone.

### Le skill `/dataroom-prep` n'apparaît pas dans Claude Code

Redémarre Claude Code (`/exit` puis `claude`). Si toujours pas visible, vérifie
manuellement que le dossier `~/.claude/skills/dataroom-prep/` existe et contient
`SKILL.md`.

---

## Utilisation après installation

Dans **n'importe quel projet Claude Code**, tape :

```
/dataroom-prep
```

Le skill pose **toutes ses questions en une seule passe**, dès le départ (jamais découpé en
« Phase 1 / Phase 2 ») :

1. **Où sont tes documents source ?** (chemin ou ZIP)
2. **Où veux-tu créer la dataroom ?** (chemin destination)
3. **Nom du projet ?** (ex: "Project Acme")
4. **Langue de la structure de dossiers ?** (fr/en/de/it/es)
5. **Langue de l'Excel mapping ?** (fr/en/de/it/es)
6. **Provenance de la documentation (pays) ?** (France/Suisse — choisit la liste DD de référence)
7. **Secteur d'activité ?** (onglet de la liste DD)
8. **Veux-tu aussi la checklist enrichie des documents à demander au client ?** (Oui/Non — si
   oui, enchaîne immédiatement localisation, site web, secteur connu, sources à consulter)

L'anonymisation (nom réel de la société cible + marques/filiales) est **toujours active**,
sans confirmation à demander : aucune de ces informations ne doit apparaître dans les fichiers,
l'Excel, ou la conversation. Les PDF scannés sont lus par OCR automatiquement (le skill installe
lui-même `pytesseract` + Tesseract au premier lancement réel si absents — rien à faire).

Livrables finaux dans le dossier destination :

```
<Project name>/
├── 01_<General Information>/
├── 02_<Finance>/
│   ├── <Annual Financial Statements>/
│   └── ...
├── 03_<Legal>/
│   ├── <NDA & Confidentiality Agreements>/
│   ├── <Patents>/
│   └── ...
├── 04_<Tax>/
├── 05_<IT>/
├── 06_<Real Estate>/
├── 07_<Insurance>/
├── 08_<HR>/
├── 09_<Operations>/
├── 10_<Processes>/
├── 11_<To Sort>/
├── _98_<Exact Duplicates>/        ← doublons SHA-256 archivés
├── _99_<Old Versions>/             ← versions obsolètes archivées
└── _Dataprep Report.xlsx           ← Dashboard + Mapping + Structure + Doublons/versions
                                      + Documents à demander (Top 50, si pays/secteur fournis)
```

Plus le ZIP final si activé dans la config.

---

## Désinstaller

Supprime simplement le dossier `~/.claude/skills/dataroom-prep/` :

- Windows (PowerShell) :
  ```
  Remove-Item -Recurse -Force "$env:USERPROFILE\.claude\skills\dataroom-prep"
  ```
- macOS / Linux :
  ```
  rm -rf ~/.claude/skills/dataroom-prep
  ```

Et le clone du repo si tu n'en as plus besoin :

- Windows : `Remove-Item -Recurse -Force "$env:USERPROFILE\Github\dataroom-prep"`
- macOS / Linux : `rm -rf ~/Github/dataroom-prep`
