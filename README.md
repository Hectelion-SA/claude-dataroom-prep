# Dataroom Prep — AI-powered M&A Dataroom Builder

> Turn a messy folder of documents into a structured M&A dataroom in 10 minutes instead of 2 days — intelligent classification, deduplication, version detection, multilingual folder structure, and contextual missing-docs checklist. Installed end-to-end as a Claude Code skill.

[![Hectelion SA](https://img.shields.io/badge/built%20by-Hectelion%20SA-182E4E)](https://www.hectelion.com)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Windows](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-blue)](#prerequisites)
[![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue)](https://www.python.org/)
[![Claude Code skill](https://img.shields.io/badge/Claude%20Code-skill-orange)](https://claude.com/claude-code)

Every M&A advisor, lawyer, notary, banker or founder doing a deal runs the same loop: collect 200-1000 messy documents from the seller, rename them all to match a clean convention, dedupe Word/PDF clones from OneDrive sync hell, organize them across 11 due diligence categories with sub-folders, then build an Excel mapping for the buyer side. **2-3 days of junior associate work, every deal.**

This **Claude Code skill** automates the entire loop:

1. **Asks every question upfront, in one single pass** — source folder, destination, project name, folder/Excel language, country + sector (for the DD checklist), and whether to also run the optional enriched checklist — never split across a first pass and a follow-up
2. **Extracts text** from every PDF/DOCX/XLSX/PPTX, with automatic **OCR fallback for scanned PDFs** (PyMuPDF + Tesseract, installed by the skill itself on first run if missing — nothing to set up)
3. **Classifies** using parent_folder + filename + content heuristics into 11 standard DD folders × topical sub-folders
4. **Detects exact duplicates** via SHA-256 hashing → archives them separately
5. **Detects obsolete versions** by clustering filename bases (v1/v2/final/signed/OneDrive copies) → archives only the obsolete ones, keeps the most recent
6. **Renames** with a short, content-derived title (never the raw filename, never the folder theme) following the pattern `Project Name - Short Title - yyyymmdd.ext`
7. **Anonymizes automatically** — the target company's real name, brands, and affiliated entities never appear in filenames, the Excel, or the conversation
8. **Generates an Excel report** (multilingual) showing Original → Renamed for every file, with dedup status, final path, and a Top-50-missing-documents tab (offline, from your firm's own DD reference lists — see [`skill/data_sources/README.md`](skill/data_sources/README.md))
9. **Enriched checklist (optional, opt-in)** — the only step that reaches the internet: scrapes the company's public website + legal sources to refine the missing-documents checklist by location + sector

**Built for**: M&A boutiques, investment banks, corporate lawyers, transactional notaries, accounting firms, family office advisors, solo dealmakers preparing the pre-VDR stage.

**Built by**: [Hectelion SA](https://www.hectelion.com) — independent M&A advisory, Lausanne (Switzerland).

---

## Installation in 3 minutes — driven by Claude Code

**No copy-paste of code, no PowerShell typing.** You paste one prompt into Claude Code, click "Allow" to approve each step, answer 5 config questions, done.

### How it works

1. Have **Claude Code** installed and signed in ([claude.com/claude-code](https://claude.com/claude-code)).
2. Open a terminal in any folder. Run `claude`.
3. Copy-paste the prompt from [`CLAUDE_INSTALL_PROMPT.md`](./CLAUDE_INSTALL_PROMPT.md) into the conversation.
4. Send. Claude Code clones the repo, copies the skill to `~/.claude/skills/dataroom-prep/`, installs Python dependencies, asks 5 config questions, writes `config.yaml`.
5. From then on, in **any** Claude Code project, type `/dataroom-prep` to launch the workflow and build a new dataroom.

### Prerequisites

- Windows 10/11, macOS, or Linux
- Python 3.9+ on `PATH` ([python.org](https://python.org))
- Claude Code installed + active Anthropic plan
- *(Optional)* Firecrawl MCP for the enriched checklist's website scraping step

### Manual install (without Claude Code prompt)

```bash
git clone https://github.com/Hectelion-SA/claude-dataroom-prep.git ~/Github/dataroom-prep
cd ~/Github/dataroom-prep
# Windows
./install.ps1
# macOS / Linux
bash install.sh
```

---

## What you get

### The conversation — every question asked upfront, in one pass

No "come back later for more questions": everything below is asked together, before any file
is touched, including the enriched-checklist sub-questions if you say yes to it.

| Question | Example answer |
|---|---|
| Where are your source documents? | `Z:\Projects\Acme M&A\Source docs` |
| Where to create the dataroom? | `D:\Datarooms` |
| Project name? | `Project Acme` |
| Language for folder structure? | `en` (or `fr` / `de` / `it` / `es`) |
| Language for Excel report? | `en` |
| Country (for the DD reference list)? | `France` / `Switzerland` |
| Sector? | `MedTech` / `SaaS` / `Real Estate` / `Industrial` / ... |
| Also run the optional enriched checklist? | Yes / No |
| *(if Yes)* Company location? | `France` / `Switzerland` / `Belgium` / `Luxembourg` / other |
| *(if Yes)* Company website URL? | `https://acme.com` (auto-scrapes for sector detection) |
| *(if Yes)* Sector already known? | optional — auto-detected from the website if omitted |
| *(if Yes)* Sources to consult? | data.gouv / Légifrance / Fedlex / OpenLaw / none |

### The output structure — 2-level folder tree (multilingual)

```
Project Acme/
├── 01_General Information/
├── 02_Finance/
│   ├── Annual Financial Statements/
│   ├── Business Plan/
│   ├── Cap Table/
│   ├── Commercial Invoicing/
│   └── Funding Financials/
├── 03_Legal/
│   ├── NDA & Confidentiality Agreements/
│   ├── Patents/
│   ├── Trademarks/
│   ├── Articles of Association/
│   ├── General Meetings/
│   ├── Share Transfers/
│   ├── Term Sheets/
│   ├── Beneficial Owners/
│   ├── KBIS/
│   └── ...
├── 04_Tax/
│   └── Tax Return Package/
├── 05_IT/
│   └── IT Certifications/
├── 06_Real Estate/
├── 07_Insurance/
├── 08_HR/
├── 09_Operations/
│   ├── Commercial - Customer A/
│   ├── Commercial - Customer B/
│   ├── Product Certifications/
│   └── Marketing & Commercial/
├── 10_Processes/
├── 11_To Sort/                       ← unreadable scans, LOW confidence files
├── _98_Exact Duplicates/              ← SHA-256 hash matched
├── _99_Old Versions/                  ← obsolete versions, organized by N1/N2
│   ├── 03_Legal/
│   │   └── NDA & Confidentiality Agreements/
│   └── ...
├── _Dataprep Report.xlsx              ← see below (always generated)
└── _Liste documents à demander.xlsx   ← only if you opted into the enriched checklist
```

### The Excel — `_Dataprep Report.xlsx` (5 tabs, multilingual headers, always generated)

**Tab 1 — Dashboard** — KPIs : docs processed, HIGH/LOW confidence split, duplicates archived, version clusters detected

**Tab 2 — Mapping** — 9 columns showing the **Original → Renamed** mapping for every file:

```
# | Original Name | Source Path | Renamed File | Folder L1 | Sub-folder L2 | Confidence | Dedup Status | Final Path
```

Colored cells: confidence (green/orange/red), dedup status (green/orange/red).

**Tab 3 — Structure** — Tree of the generated dataroom with doc count per N1/N2

**Tab 4 — Duplicates & versions** — full audit trail of what was archived and why

**Tab 5 — Documents à demander (Top 50)** — offline gap analysis against your firm's DD
reference list for the chosen country/sector (see
[`skill/data_sources/README.md`](skill/data_sources/README.md) — these lists are firm
methodology and are not bundled in this public repo; without them, this tab is skipped and
everything else still runs)

### Optional — `_Liste documents à demander.xlsx` (enriched checklist)

The skill cross-references your dataroom against:

- **60+ base M&A checklist items** (bilans, statuts, PV AG, top clients, IP, contentieux, RGPD, etc.)
- **Location-specific obligations**:
  - France: DUERP, Index égalité, Sapin II, PV CSE
  - Switzerland: AVS/LPP/SUVA, LAA, OFM permits, RC cantonal
  - Belgium: Bilan social BNB, précomptes
  - Luxembourg: RCS LU bilans, IRC/ICC
- **Sector-specific compliance**:
  - MedTech: marquage CE MDR (Règlement 2017/745), ISO 13485, DMR, PMS
  - SaaS: DPA, hébergement HDS, audit pentest, MRR/ARR/churn cohorts
  - Real Estate: DPE, amiante, plomb, Carnet d'entretien, Loi Carrez
  - Industrial: ICPE, études environnementales
- **Website scraping** (via Firecrawl) to auto-detect sector and add bespoke items

Output: **80–120 line Excel** with columns:

```
# | Section | Document/Information requested | Criticality | Status (Present / To request) | Rationale | Source (Base / Location / Sector)
```

---

## Real-world benchmark

Tested on a 4076-file project folder:

| Step | Result |
|---|---|
| Documents extracted (PDF/DOCX/XLSX) | **785** (3291 noise files filtered out: code, build artifacts) |
| HIGH confidence classification | **99.4%** (780/785) |
| Exact duplicates detected (SHA-256) | **99 files** archived in `_98_Exact Duplicates/` |
| Obsolete versions detected | **70 files** archived in `_99_Old Versions/{N1}/{N2}/` |
| Noise reduction | **22%** of corpus cleaned |
| Final dataroom size | **618 documents** in clean 2-level structure |
| Total pipeline runtime | **3 minutes** (extraction + classification + dedup + build + Excel) |
| Match vs human-classified VDR | **9/9** on root-level documents |

Equivalent manual work: **2-3 days of junior associate** at 600 CHF/day = **1500-1800 CHF of value per dataroom**.

---

## Supported languages

| Language | Code | Folder structure | Excel headers | DD checklist |
|---|---|---|---|---|
| French | `fr` | ✅ | ✅ | ✅ |
| English | `en` | ✅ | ✅ | ✅ |
| German | `de` | ✅ | ✅ | partial |
| Italian | `it` | ✅ | ✅ | partial |
| Spanish | `es` | ✅ | ✅ | partial |

---

## Headless / scripted usage

For batch processing (CI, automation):

```bash
# 1. Extract only (produces extracted.json for content-based titling)
python ~/.claude/skills/dataroom-prep/scripts/run_pipeline.py \
  --source "/path/to/messy/folder" \
  --destination "/path/to/output" \
  --project-name "Project Acme" \
  --extract-only "/path/to/output/extracted.json"

# 2. Build, with content-derived titles/classification (decisions.json), anonymization,
#    and the offline Top-50 gap analysis for a given country/sector
python ~/.claude/skills/dataroom-prep/scripts/run_pipeline.py \
  --destination "/path/to/output" \
  --project-name "Project Acme" \
  --folder-lang en \
  --excel-lang en \
  --decisions-json "/path/to/output/decisions.json" \
  --scrub-file "/path/to/scrub_names.txt" \
  --country FR \
  --sector "Tech" \
  --make-zip

# 3. Optional — enriched checklist (website + legal sources, the only network step)
python ~/.claude/skills/dataroom-prep/scripts/enrich_checklist.py \
  --dataroom "/path/to/output/Project Acme" \
  --extracted-json "/path/to/output/extracted.json" \
  --location "France" \
  --sector "MedTech"
```

Scanned PDFs are OCR'd automatically (`ensure_ocr_stack()` installs `pytesseract` + Tesseract +
French/English language data on first real run if missing — safe to run offline afterwards).

---

## Privacy & Security

- 🔒 **Local processing only** — extraction, classification, dedup, OCR, and the Top-50 gap
  analysis all run 100% offline on your machine; no document content ever leaves it
- 🔒 **No third-party uploads by default** — the only network call in the whole pipeline is
  the opt-in enriched checklist (website + legal sources), and only the public URL and
  sector/legal queries are sent — never document content, and always confirmed with you first
- 🔒 **Automatic anonymization** — the target company's real name, brands, and affiliated
  entities are detected from the extracted content and scrubbed from every generated
  filename, folder, and Excel cell (case-insensitive, word-boundary matching) — on by default,
  never asked as a question
- 🔒 **GDPR-friendly architecture** — no persistent storage, files read but never sent externally
- 🔒 **Audit trail** — every file movement documented in `_Dataprep Report.xlsx`
- 🔒 **Original files untouched** — the skill copies, never moves or modifies source files

---

## Roadmap

- [x] Phase 1: extract + classify + dedup + build N1/N2 dataroom + Excel mapping
- [x] Phase 2: contextual checklist (location + sector + website)
- [ ] V2: Légifrance API integration (FR legal references in checklist rationale)
- [ ] V2: Fedlex (Swiss federal law) integration
- [ ] V2: Auto-redaction of PII (regex + LLM for contextual names) for pre-NDA sharing
- [ ] V2: Document summaries + Q&A pre-generated per key document (buyer experience)
- [ ] V2: Email gateway (`dataprep+project@your-domain.com`) for auto-classification on email forward
- [ ] V3: Direct VDR push to Ansarada, Datasite, Drooms (API integration)
- [ ] V3: Watermarking dynamique per recipient (forensic anti-leak)

---

## About Hectelion SA

[Hectelion SA](https://www.hectelion.com) is an independent Franco-Swiss M&A advisory firm based in Lausanne, focused on **mid-cap transactions** (CHF 2M to CHF 500M):

- **Sell-side & buy-side mandates** for SMEs and family-owned businesses
- **Company valuations** (DCF, comparable transactions, comparable trading, intangibles)
- **Due diligence** (financial, legal, operational) and dataroom preparation
- **Transaction structuring** (share deal vs asset deal, MBO/LBO, earn-out, ratchets)

For tailored DD checklist templates by industry, enterprise deployment, or integration with your existing VDR — [book a 30-min discovery call](https://calendly.com/aristide-ruot-hectelion).

Founder: [Aristide Ruot](https://www.linkedin.com/in/aristideruot/) — `aristide.ruot@hectelion.com`

---

## License

MIT — see [LICENSE](LICENSE). Use it, fork it, modify it, deploy it for your firm. Attribution appreciated but not required.

---

## Keywords (for GitHub search)

`m-and-a` · `mergers-and-acquisitions` · `due-diligence` · `dataroom` · `virtual-data-room` · `vdr` · `vdr-preparation` · `document-classification` · `ai-document-organization` · `pdf-classification` · `claude-code-skill` · `claude-skill` · `legal-tech` · `m-and-a-tools` · `dd-prep` · `dataroom-automation` · `notary` · `lawyer-tools` · `corporate-finance` · `private-equity` · `investment-banking` · `transactional-tools` · `data-deduplication` · `version-control-documents` · `franco-swiss` · `hectelion`
