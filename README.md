# Dataroom Prep — AI-powered M&A Dataroom Builder

> Turn a messy folder of documents into a structured M&A dataroom in 10 minutes instead of 2 days — intelligent classification, deduplication, version detection, multilingual folder structure, and contextual missing-docs checklist. Installed end-to-end as a Claude Code skill.

[![Hectelion SA](https://img.shields.io/badge/built%20by-Hectelion%20SA-182E4E)](https://www.hectelion.com)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Windows](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-blue)](#prerequisites)
[![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue)](https://www.python.org/)
[![Claude Code skill](https://img.shields.io/badge/Claude%20Code-skill-orange)](https://claude.com/claude-code)

Every M&A advisor, lawyer, notary, banker or founder doing a deal runs the same loop: collect 200-1000 messy documents from the seller, rename them all to match a clean convention, dedupe Word/PDF clones from OneDrive sync hell, organize them across 11 due diligence categories with sub-folders, then build an Excel mapping for the buyer side. **2-3 days of junior associate work, every deal.**

This **Claude Code skill** automates the entire loop:

1. **Asks 5 short questions** — source folder, destination, project name, language for folders, language for Excel mapping
2. **Extracts text** from every PDF/DOCX/XLSX/PPTX
3. **Classifies** using parent_folder + filename heuristics into 11 standard DD folders × topical sub-folders
4. **Detects exact duplicates** via SHA-256 hashing → archives them separately
5. **Detects obsolete versions** by clustering filename bases (v1/v2/final/signed/OneDrive copies) → archives only the obsolete ones, keeps the most recent
6. **Renames** following the pattern `Project Name - Theme - Subject - YYYYMM.ext`
7. **Generates an Excel mapping** (multilingual) showing Original → Renamed for every file, with dedup status and final path
8. **Phase 2 (optional)** — builds a contextual checklist of documents to request from client, enriched by location (FR / CH / BE / LU / other) + sector + website scraping

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
- *(Optional)* Firecrawl MCP for Phase 2 website scraping

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

### The conversation — 5 questions Phase 1 + 3 optional Phase 2

| Phase | Question | Example answer |
|---|---|---|
| **1** | Where are your source documents? | `Z:\Projects\Acme M&A\Source docs` |
| **1** | Where to create the dataroom? | `D:\Datarooms` |
| **1** | Project name? | `Project Acme` |
| **1** | Language for folder structure? | `en` (or `fr` / `de` / `it` / `es`) |
| **1** | Language for Excel mapping? | `en` |
| **2** | Generate enriched DD checklist? | Yes / No |
| **2** | Company location? | `France` / `Switzerland` / `Belgium` / `Luxembourg` / other |
| **2** | Company website URL? | `https://acme.com` (auto-scrapes for sector detection) |
| **2** | Sector? | `MedTech` / `SaaS` / `Real Estate` / `Industrial` / ... (auto-detect from website if omitted) |

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
├── _Mapping report.xlsx               ← see below
└── _Documents to request.xlsx         ← Phase 2 only
```

### The Excel — `_Mapping report.xlsx` (5 tabs, multilingual headers)

**Tab 1 — Dashboard** — KPIs : docs processed, HIGH/LOW confidence split, duplicates archived, version clusters detected

**Tab 2 — Mapping** — 9 columns showing the **Original → Renamed** mapping for every file:

```
# | Original Name | Source Path | Renamed File | Folder L1 | Sub-folder L2 | Confidence | Dedup Status | Final Path
```

Colored cells: confidence (green/orange/red), dedup status (green/orange/red).

**Tab 3 — Structure** — Tree of the generated dataroom with doc count per N1/N2

**Tab 4 — DD Checklist** — (only with Phase 2) — items present / to request, with criticality

**Tab 5 — Duplicates & versions** — full audit trail of what was archived and why

### Phase 2 — `_Documents to request.xlsx` (enriched checklist)

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
python ~/.claude/skills/dataroom-prep/scripts/run_pipeline.py \
  --source "/path/to/messy/folder" \
  --destination "/path/to/output" \
  --project-name "Project Acme" \
  --folder-lang en \
  --excel-lang en \
  --make-zip

python ~/.claude/skills/dataroom-prep/scripts/enrich_checklist.py \
  --dataroom "/path/to/output/Project Acme" \
  --extracted-json "/path/to/output/extracted.json" \
  --location "France" \
  --sector "MedTech"
```

---

## Privacy & Security

- 🔒 **Local processing only** — all extraction, classification, dedup happens on your machine
- 🔒 **No third-party uploads** by default (Firecrawl only used if you opt-in for Phase 2 website scraping)
- 🔒 **GDPR-friendly architecture** — no persistent storage, files read but never sent externally
- 🔒 **Audit trail** — every file movement documented in `_Mapping report.xlsx`
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
