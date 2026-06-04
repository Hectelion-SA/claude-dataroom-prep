"""dataroom-prep — Pipeline principal (extract + classify + dedup + build + Excel).

CLI:
    python run_pipeline.py \
        --source "<chemin source>" \
        --destination "<chemin destination>" \
        --project-name "Project Acme" \
        --folder-lang fr \
        --excel-lang fr

Compatible Windows, macOS, Linux. Force UTF-8 stdout.
"""
import argparse
import hashlib
import json
import os
import re
import shutil
import stat
import sys
import unicodedata
import zipfile
from collections import defaultdict
from pathlib import Path

import pypdf
import docx
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

# ============================================================================
# CONSTANTS
# ============================================================================

SKILL_DIR = Path(__file__).parent.parent
TEMPLATES_DIR = SKILL_DIR / "templates"

MAX_CHARS = 2500
DD_EXTS = {".pdf", ".docx", ".xlsx", ".xlsm", ".doc", ".xls", ".pptx", ".ppt", ".msg", ".rtf", ".txt"}
EXCLUDE_SUBSTRINGS = [
    "vdr - 202", "r&d [clean]", "linux-unpacked", "node_modules",
    "smartmirrorv2_data", "ai pod unity builds", "checkuponly",
    "/temp/", "_files\\", "\\bin\\", "\\obj\\", ".git\\", "$recycle.bin",
]

# Hectelion brand
NAVY = "182E4E"
INPUT_CYAN = "72C7E7"
SOFT_SKY = "D3E7FF"
MINT = "6FCF9A"
WARNING = "FFC000"
ALERT = "C00000"
WHITE = "FFFFFF"


# ============================================================================
# EXTRACTION
# ============================================================================

def _excluded(path: Path) -> bool:
    p = str(path).lower()
    return any(sub in p for sub in EXCLUDE_SUBSTRINGS)


def _extract_pdf(path: Path) -> tuple[str, str]:
    try:
        reader = pypdf.PdfReader(str(path))
        n_pages = len(reader.pages)
        text = ""
        for i in [0, min(1, n_pages - 1)]:
            try:
                text += reader.pages[i].extract_text() + "\n"
            except Exception:
                pass
            if len(text) > MAX_CHARS:
                break
        return text[:MAX_CHARS], f"pdf:{n_pages}p"
    except Exception as e:
        return "", f"pdf:ERROR:{type(e).__name__}"


def _extract_docx(path: Path) -> tuple[str, str]:
    try:
        d = docx.Document(str(path))
        parts = []
        for p in d.paragraphs:
            t = p.text.strip()
            if t:
                parts.append(t)
            if sum(len(x) for x in parts) > MAX_CHARS:
                break
        return "\n".join(parts)[:MAX_CHARS], f"docx:{len(d.paragraphs)}p"
    except Exception as e:
        return "", f"docx:ERROR:{type(e).__name__}"


def _extract_xlsx(path: Path) -> tuple[str, str]:
    try:
        wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
        out = []
        for sheet_name in wb.sheetnames[:2]:
            ws = wb[sheet_name]
            out.append(f"=== Sheet: {sheet_name} ===")
            for row in ws.iter_rows(max_row=20, values_only=True):
                row_str = " | ".join(str(c) for c in row if c is not None)
                if row_str.strip():
                    out.append(row_str)
                if sum(len(x) for x in out) > MAX_CHARS:
                    break
        return "\n".join(out)[:MAX_CHARS], f"xlsx:{len(wb.sheetnames)}s"
    except Exception as e:
        return "", f"xlsx:ERROR:{type(e).__name__}"


def _extract_one(path: Path, source_root: Path) -> dict:
    ext = path.suffix.lower()
    try:
        size_kb = round(path.stat().st_size / 1024, 1)
    except Exception:
        size_kb = 0
    if ext == ".pdf":
        text, meta = _extract_pdf(path)
    elif ext == ".docx":
        text, meta = _extract_docx(path)
    elif ext in (".xlsx", ".xlsm"):
        text, meta = _extract_xlsx(path)
    else:
        text, meta = "", f"unsupported:{ext}"
    rel = path.relative_to(source_root)
    return {
        "relpath": str(rel),
        "filename": path.name,
        "parent_folder": str(rel.parent),
        "size_kb": size_kb,
        "meta": meta,
        "snippet": text.strip(),
        "readable": bool(text.strip()),
        "abspath": str(path),
    }


def extract_all(source: Path) -> list:
    """Extrait recursivement tous les docs DD du dossier source."""
    docs = []
    for f in source.rglob("*"):
        if not f.is_file():
            continue
        if f.suffix.lower() not in DD_EXTS:
            continue
        if _excluded(f):
            continue
        try:
            docs.append(_extract_one(f, source))
        except Exception as e:
            pass
    return docs


# ============================================================================
# CLASSIFICATION (multilingue via templates)
# ============================================================================

def load_template(lang: str) -> dict:
    path = TEMPLATES_DIR / f"folder_structure_{lang.lower()}.json"
    if not path.exists():
        # fallback FR
        path = TEMPLATES_DIR / "folder_structure_fr.json"
    return json.loads(path.read_text(encoding="utf-8"))


# Themes neutres (clés internes, traduites par langue dans une table en fin de fichier)
# Pour le POC, on utilise les noms FR du template comme thèmes "canoniques"
# et on les traduit pour l'affichage final

def classify(doc: dict, template: dict) -> tuple[str, str, str]:
    """Retourne (folder_n1, theme_n2, confidence). Folder_n1 et theme_n2 dans la langue du template."""
    parent = doc["parent_folder"].lower().replace("\\", "/")
    fname = doc["filename"].lower()
    readable = doc.get("readable", False)
    folders = template["folders_n1"]
    lang = template["language"]

    informative_filename = bool(re.search(r"[a-z]{4,}", fname.replace(".pdf", "").replace(".docx", "")))
    if not readable and not informative_filename:
        return (folders["11"], "", "LOW")

    # Themes par langue (mini-dictionnaire pour les patterns frequents)
    THEMES = {
        "fr": {
            "ip": "Propriété intellectuelle", "nda": "NDA & accords confidentialité",
            "marque": "Marque", "brevet": "Brevets", "admin_ip": "Administratif IP",
            "marque_bodyo": "Marque (dédiée)",
            "ag": "Assemblées générales", "cession": "Cession de parts",
            "bilans": "Bilans annuels", "funding": "Financials Funding",
            "partnership": "Financials Partnership", "facturation": "Facturation commerciale",
            "financials_com": "Financials commerciaux",
            "liasse": "Liasse fiscale",
            "term_sheets": "Term sheets", "kbis": "KBIS", "strike": "Radiation entité",
            "statuts": "Statuts", "marketing": "Marketing & commercial",
            "sieges": "Sièges sociaux", "benefic": "Bénéficiaires effectifs",
            "pv_ag": "PV Assemblées", "cap_table": "Table de capitalisation",
            "soft_dev": "Certifications IT", "infra": "Infrastructure IT",
            "risk": "Risk management", "doc_cert": "Documentation certifications",
            "machines": "Suivi machines", "cert_produit": "Certifications produit",
            "commercial": "Commercial",
            "nda_com": "NDA commerciales", "to_sort": "Non classé",
        },
        "en": {
            "ip": "Intellectual Property", "nda": "NDA & Confidentiality Agreements",
            "marque": "Trademark", "brevet": "Patents", "admin_ip": "IP Administration",
            "marque_bodyo": "Trademark (dedicated)",
            "ag": "General Meetings", "cession": "Share Transfers",
            "bilans": "Annual Financial Statements", "funding": "Funding Financials",
            "partnership": "Partnership Financials", "facturation": "Commercial Invoicing",
            "financials_com": "Commercial Financials",
            "liasse": "Tax Return Package",
            "term_sheets": "Term Sheets", "kbis": "Commercial Register Extract", "strike": "Strike Off",
            "statuts": "Articles of Association", "marketing": "Marketing & Commercial",
            "sieges": "Registered Offices", "benefic": "Beneficial Owners",
            "pv_ag": "Meeting Minutes", "cap_table": "Cap Table",
            "soft_dev": "IT Certifications", "infra": "IT Infrastructure",
            "risk": "Risk Management", "doc_cert": "Certifications Documentation",
            "machines": "Machine Tracking", "cert_produit": "Product Certifications",
            "commercial": "Commercial",
            "nda_com": "Commercial NDAs", "to_sort": "Unsorted",
        },
        "de": {
            "ip": "Geistiges Eigentum", "nda": "NDA & Vertraulichkeitsvereinbarungen",
            "marque": "Marke", "brevet": "Patente", "admin_ip": "IP Verwaltung",
            "marque_bodyo": "Marke (dediziert)",
            "ag": "Generalversammlungen", "cession": "Anteilsübertragungen",
            "bilans": "Jahresabschlüsse", "funding": "Finanzierung",
            "partnership": "Partnerschaftsfinanzen", "facturation": "Geschäftsabrechnungen",
            "financials_com": "Kommerzielle Finanzen",
            "liasse": "Steuererklärungspaket",
            "term_sheets": "Term Sheets", "kbis": "Handelsregisterauszug", "strike": "Löschung",
            "statuts": "Satzung", "marketing": "Marketing & Vertrieb",
            "sieges": "Eingetragene Sitze", "benefic": "Wirtschaftlich Berechtigte",
            "pv_ag": "Sitzungsprotokolle", "cap_table": "Kapitalisierungstabelle",
            "soft_dev": "IT Zertifizierungen", "infra": "IT Infrastruktur",
            "risk": "Risikomanagement", "doc_cert": "Zertifizierungsdokumentation",
            "machines": "Maschinenverfolgung", "cert_produit": "Produktzertifizierungen",
            "commercial": "Vertrieb",
            "nda_com": "Geschäftliche NDAs", "to_sort": "Unsortiert",
        },
        "it": {
            "ip": "Proprietà intellettuale", "nda": "NDA & Riservatezza",
            "marque": "Marchio", "brevet": "Brevetti", "admin_ip": "Amministrazione PI",
            "marque_bodyo": "Marchio (dedicato)",
            "ag": "Assemblee generali", "cession": "Cessione di quote",
            "bilans": "Bilanci annuali", "funding": "Finanziamenti",
            "partnership": "Finanze partnership", "facturation": "Fatturazione commerciale",
            "financials_com": "Finanze commerciali",
            "liasse": "Dichiarazione fiscale",
            "term_sheets": "Term Sheets", "kbis": "Visura camerale", "strike": "Cancellazione",
            "statuts": "Statuto", "marketing": "Marketing & commerciale",
            "sieges": "Sedi sociali", "benefic": "Titolari effettivi",
            "pv_ag": "Verbali assemblea", "cap_table": "Cap Table",
            "soft_dev": "Certificazioni IT", "infra": "Infrastruttura IT",
            "risk": "Gestione del rischio", "doc_cert": "Documentazione certificazioni",
            "machines": "Tracciabilità macchine", "cert_produit": "Certificazioni prodotto",
            "commercial": "Commerciale",
            "nda_com": "NDA commerciali", "to_sort": "Da classificare",
        },
        "es": {
            "ip": "Propiedad intelectual", "nda": "NDA & Confidencialidad",
            "marque": "Marca", "brevet": "Patentes", "admin_ip": "Administración PI",
            "marque_bodyo": "Marca (dedicada)",
            "ag": "Asambleas generales", "cession": "Cesión de participaciones",
            "bilans": "Balances anuales", "funding": "Financiación",
            "partnership": "Financiación partenariado", "facturation": "Facturación comercial",
            "financials_com": "Finanzas comerciales",
            "liasse": "Declaración fiscal",
            "term_sheets": "Term Sheets", "kbis": "Extracto registro mercantil", "strike": "Disolución",
            "statuts": "Estatutos", "marketing": "Marketing & comercial",
            "sieges": "Sedes sociales", "benefic": "Beneficiarios efectivos",
            "pv_ag": "Actas asamblea", "cap_table": "Cap Table",
            "soft_dev": "Certificaciones IT", "infra": "Infraestructura IT",
            "risk": "Gestión de riesgos", "doc_cert": "Documentación certificaciones",
            "machines": "Seguimiento máquinas", "cert_produit": "Certificaciones producto",
            "commercial": "Comercial",
            "nda_com": "NDAs comerciales", "to_sort": "Por clasificar",
        },
    }
    T = THEMES.get(lang, THEMES["fr"])

    # === 03_Légal : IP, Brevets, Marques, NDA ===
    if "11_assets/b - propriete intel" in parent or "intellectual property" in parent or "propriete intel" in parent:
        if "nda" in parent or "confident" in parent:
            return (folders["03"], T["nda"], "HIGH")
        if "marque" in parent:
            return (folders["03"], T["marque"], "HIGH")
        if "brevet" in parent or "patent" in fname:
            return (folders["03"], T["brevet"], "HIGH")
        return (folders["03"], T["ip"], "HIGH")
    if "11_assets/a - marque" in parent:
        return (folders["03"], T["marque_bodyo"], "HIGH")
    if "11_assets/c - administratif" in parent or "11_assets\\c" in parent:
        return (folders["03"], T["admin_ip"], "HIGH")
    if "11_assets" in parent:
        return (folders["03"], T["ip"], "HIGH")
    if "confidentialite" in parent or "confidendialite" in parent or "nda" in parent:
        return (folders["03"], T["nda"], "HIGH")
    if "assemblees principales" in parent or parent.startswith("assemblees"):
        return (folders["03"], T["ag"], "HIGH")
    if "cession de parts" in parent:
        return (folders["03"], T["cession"], "HIGH")

    # === 02_Finance : Bilans ===
    if re.search(r"\bbilan", parent) or re.search(r"^bilan", parent) or "balancecompta" in parent:
        return (folders["02"], T["bilans"], "HIGH")
    if "financials funding" in parent:
        return (folders["02"], T["funding"], "HIGH")
    if "financials partnership" in parent:
        return (folders["02"], T["partnership"], "HIGH")

    # === 04_Fiscal : Liasse ===
    if parent == "liasse" or parent.startswith("liasse/"):
        return (folders["04"], T["liasse"], "HIGH")

    # === 09_Opérations : Commercial ===
    if "commercial" in parent:
        if any(k in parent for k in ["invoice", "facture"]):
            return (folders["02"], T["facturation"], "HIGH")
        if any(k in parent for k in ["financial", "balance"]):
            return (folders["02"], T["financials_com"], "HIGH")
        if "nda" in parent or "mnda" in fname:
            return (folders["03"], T["nda_com"], "HIGH")
        m = re.search(r"commercial/(clos|en cours)/([^/]+)", parent)
        if m:
            return (folders["09"], f"{T['commercial']} - {m.group(2).title()}", "HIGH")
        return (folders["09"], T["commercial"], "HIGH")

    # === 12_Certifications ===
    if "12_certifications" in parent:
        if any(k in parent for k in ["software developpement", "rgpd", "angular", "infratructure", "network"]):
            return (folders["05"], T["soft_dev"], "HIGH")
        if "risk management" in parent:
            return (folders["09"], T["risk"], "HIGH")
        if "documentation" in parent or "rapports de tests" in parent:
            return (folders["09"], T["doc_cert"], "HIGH")
        if "suivi machines" in parent:
            return (folders["09"], T["machines"], "HIGH")
        return (folders["09"], T["cert_produit"], "HIGH")

    # === Racine ===
    if parent == "." or parent == "":
        if "term sheet" in fname:
            return (folders["03"], T["term_sheets"], "HIGH") if readable else (folders["11"], "", "LOW")
        if "kbis" in fname:
            return (folders["03"], T["kbis"], "HIGH")
        if "strike off" in fname:
            return (folders["03"], T["strike"], "MEDIUM") if readable else (folders["11"], "", "LOW")
        if "actes" in fname or "statuts" in fname:
            return (folders["03"], T["statuts"], "HIGH")
        if "marketing" in fname:
            return (folders["09"], T["marketing"], "HIGH")
        if "registered offices" in fname:
            return (folders["03"], T["sieges"], "HIGH")
        if "beneficial owners" in fname or "benefic" in fname:
            return (folders["03"], T["benefic"], "HIGH")
        if "pv ag" in fname or "proces verbal" in fname or "assemblee" in fname:
            return (folders["03"], T["pv_ag"], "HIGH")
        if "capitalisation" in fname or "cap table" in fname:
            return (folders["03"], T["cap_table"], "HIGH")

    return (folders["11"], "", "LOW")


# ============================================================================
# DEDUPLICATION
# ============================================================================

VERSION_PATTERNS = [
    re.compile(r"\bv(\d+(?:\.\d+)?)\b", re.I),
    re.compile(r"\b(rev\d+)\b", re.I),
    re.compile(r"\b(final|signed|signe[s]?)\b", re.I),
    re.compile(r"\b(draft|brouillon)\b", re.I),
    re.compile(r"\(\d+\)"),
    re.compile(r"-APPLE.s MacBook Pro", re.I),
    re.compile(r"\b(20\d{2})[\.\-_/]?(0\d|1[0-2])[\.\-_/]?(0\d|[12]\d|3[01])?\b"),
]

VERSION_RANK = {
    "draft": 0, "brouillon": 0,
    "v1": 10, "v2": 20, "v3": 30, "v4": 40, "v5": 50,
    "rev1": 10, "rev2": 20, "rev3": 30,
    "final": 100, "signed": 200, "signe": 200, "signes": 200,
}


def file_hash(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()[:16]


def normalize_for_clustering(filename: str) -> str:
    base = Path(filename).stem.lower()
    for rx in VERSION_PATTERNS:
        base = rx.sub("", base)
    base = re.sub(r"[\s_\-\.]+", " ", base).strip()
    return base


def version_score(filename: str) -> int:
    fname = filename.lower()
    score = 0
    for keyword, points in VERSION_RANK.items():
        if keyword in fname:
            score = max(score, points)
    m = re.search(r"\bv(\d+)\b", fname)
    if m:
        score += int(m.group(1)) * 5
    m = re.search(r"\b(20\d{2})[\.\-_/]?(0\d|1[0-2])?", fname)
    if m:
        year = int(m.group(1))
        month = int(m.group(2)) if m.group(2) else 6
        score += (year - 2018) * 12 + month
    return score


def detect_duplicates(docs: list) -> dict:
    by_hash = defaultdict(list)
    for d in docs:
        if d.get("abspath"):
            try:
                h = file_hash(Path(d["abspath"]))
                d["sha"] = h
                by_hash[h].append(d)
            except Exception:
                d["sha"] = None

    exact_dups = []
    dup_set = set()
    for h, group in by_hash.items():
        if len(group) > 1:
            keep = min(group, key=lambda x: (len(x["filename"]), x["filename"]))
            exact_dups.append({
                "hash": h,
                "files": [d["relpath"] for d in group],
                "keep": keep["relpath"],
                "remove": [d["relpath"] for d in group if d["relpath"] != keep["relpath"]],
            })
            for d in group:
                if d["relpath"] != keep["relpath"]:
                    dup_set.add(d["relpath"])

    # Version clusters (exclut les exact dups deja traites)
    by_cluster = defaultdict(list)
    for d in docs:
        if d["relpath"] in dup_set:
            continue
        key = (d["parent_folder"], normalize_for_clustering(d["filename"]))
        by_cluster[key].append(d)

    version_clusters = []
    obs_set = {}
    for (parent, base), group in by_cluster.items():
        if len(group) > 1 and base:
            scored = sorted(group, key=lambda x: -version_score(x["filename"]))
            keep = scored[0]
            obsolete = scored[1:]
            version_clusters.append({
                "parent_folder": parent,
                "base_name": base,
                "count": len(group),
                "keep": keep["relpath"],
                "obsolete": [{"path": d["relpath"], "score": version_score(d["filename"])} for d in obsolete],
            })
            for o in obsolete:
                obs_set[o["relpath"]] = keep["relpath"]

    for d in docs:
        if d["relpath"] in dup_set:
            d["status_dedup"] = "DOUBLON_EXACT"
        elif d["relpath"] in obs_set:
            d["status_dedup"] = "VERSION_OBSOLETE"
            d["kept_version"] = obs_set[d["relpath"]]
        else:
            d["status_dedup"] = "UNIQUE"

    return {
        "exact_duplicates": exact_dups,
        "version_clusters": version_clusters,
        "total_exact_dups": len(dup_set),
        "total_obsolete": len(obs_set),
    }


# ============================================================================
# BUILD DATAROOM
# ============================================================================

def _force_remove(func, path, _excinfo):
    os.chmod(path, stat.S_IWRITE)
    func(path)


def sanitize(name: str, max_len: int = 80) -> str:
    name = unicodedata.normalize("NFC", name)
    name = re.sub(r'[<>:"/\\|?*]', " - ", name)
    name = re.sub(r"\s+", " ", name).strip().rstrip(". ")
    return name[:max_len] if name else "Divers"


YYYYMM_RX = re.compile(r"(20\d{2})[\.\-_\s]?(0\d|1[0-2])")
YYYY_RX = re.compile(r"\b(20\d{2})\b")
DDMMYYYY_RX = re.compile(r"\b(0\d|[12]\d|3[01])[\.\-_/](0\d|1[0-2])[\.\-_/](20\d{2})\b")


def extract_date(fname: str, snippet: str) -> str:
    text = fname + " " + (snippet[:300] if snippet else "")
    m = DDMMYYYY_RX.search(text)
    if m:
        return f"{m.group(3)}{m.group(2)}"
    m = YYYYMM_RX.search(text)
    if m:
        return f"{m.group(1)}{m.group(2)}"
    m = YYYY_RX.search(text)
    if m:
        return m.group(1)
    return "undated"


def clean_subject(fname: str, max_len: int = 60) -> str:
    base = Path(fname).stem
    base = re.sub(r"\b\d{1,2}[\.\-_/]\d{1,2}[\.\-_/]\d{2,4}\b", "", base)
    base = re.sub(r"\b20\d{2}[\.\-_/]?\d{0,2}\b", "", base)
    base = re.sub(r"-APPLE.s MacBook Pro", "", base, flags=re.I)
    base = re.sub(r"\bv\d+(\.\d+)?\b", "", base, flags=re.I)
    base = re.sub(r"\bfinal\b|\bsigned\b|\bunsigned\b|\bsignés?\b", "", base, flags=re.I)
    base = re.sub(r"[_\s]+", " ", base).strip(" -._")
    if not base:
        base = "document"
    return base[:max_len]


def build_new_name(doc: dict, theme: str, project: str) -> str:
    ext = Path(doc["filename"]).suffix.lower()
    date = extract_date(doc["filename"], doc.get("snippet", ""))
    subject = clean_subject(doc["filename"])
    parts = [project, theme, subject, date]
    name = " - ".join(p for p in parts if p)
    return sanitize(name + ext, max_len=200)


def build_dataroom(docs: list, destination: Path, project: str, template: dict) -> dict:
    """Build dataroom structure N1/N2 + archives dedup."""
    out_dir = destination / project
    if out_dir.exists():
        shutil.rmtree(out_dir, onerror=_force_remove)
    out_dir.mkdir(parents=True, exist_ok=True)

    folders_n1 = template["folders_n1"]
    archives = template["archives"]

    for f in folders_n1.values():
        (out_dir / f).mkdir(exist_ok=True)

    archive_dup = out_dir / archives["duplicates"]
    archive_ver = out_dir / archives["versions"]
    archive_dup.mkdir(exist_ok=True)
    archive_ver.mkdir(exist_ok=True)

    copied_main, copied_dup, copied_ver = 0, 0, 0
    name_dedup = defaultdict(int)

    for d in docs:
        if not d.get("abspath"):
            d["copy_status"] = "MISSING SOURCE"
            continue
        src = Path(d["abspath"])

        if d["status_dedup"] == "DOUBLON_EXACT":
            dst = archive_dup / unicodedata.normalize("NFC", d["filename"])
            target_key = str(dst)
            name_dedup[target_key] += 1
            if name_dedup[target_key] > 1:
                dst = archive_dup / f"{Path(d['filename']).stem} ({name_dedup[target_key]}){Path(d['filename']).suffix}"
            try:
                shutil.copy2(src, dst)
                d["final_path"] = str(dst.relative_to(out_dir))
                d["copy_status"] = "OK"
                copied_dup += 1
            except Exception as e:
                d["copy_status"] = f"ERROR: {e}"

        elif d["status_dedup"] == "VERSION_OBSOLETE":
            sub = archive_ver / d["folder_n1"] / (d["folder_n2"] or "Divers")
            sub.mkdir(parents=True, exist_ok=True)
            dst = sub / unicodedata.normalize("NFC", d["filename"])
            target_key = str(dst)
            name_dedup[target_key] += 1
            if name_dedup[target_key] > 1:
                dst = sub / f"{Path(d['filename']).stem} ({name_dedup[target_key]}){Path(d['filename']).suffix}"
            try:
                shutil.copy2(src, dst)
                d["final_path"] = str(dst.relative_to(out_dir))
                d["copy_status"] = "OK"
                copied_ver += 1
            except Exception as e:
                d["copy_status"] = f"ERROR: {e}"

        else:
            if d["confidence"] == "LOW":
                dst_dir = out_dir / folders_n1["11"]
            else:
                dst_dir = out_dir / d["folder_n1"] / d["folder_n2"]
            dst_dir.mkdir(parents=True, exist_ok=True)
            target_name = d["new_name"]
            dst = dst_dir / target_name
            target_key = str(dst)
            name_dedup[target_key] += 1
            if name_dedup[target_key] > 1:
                dst = dst_dir / f"{Path(target_name).stem} ({name_dedup[target_key]}){Path(target_name).suffix}"
            try:
                shutil.copy2(src, dst)
                d["final_path"] = str(dst.relative_to(out_dir))
                d["copy_status"] = "OK"
                copied_main += 1
            except Exception as e:
                d["copy_status"] = f"ERROR: {e}"

    return {"out_dir": out_dir, "copied_main": copied_main, "copied_dup": copied_dup, "copied_ver": copied_ver}


# ============================================================================
# EXCEL MAPPING (multilingue)
# ============================================================================

def generate_excel(docs: list, dedup: dict, out_dir: Path, project: str, excel_template: dict):
    wb = Workbook()
    wb.remove(wb.active)

    headers_t = excel_template["excel_headers"]
    dedup_labels = excel_template["dedup_labels"]
    conf_labels = excel_template["confidence_labels"]

    def H(c):
        c.font = Font(name="Calibri", size=11, color=WHITE, bold=True)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Dashboard
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = f"DataPrep — {project}"
    ws["A1"].font = Font(name="Cardo", size=25, color=NAVY)
    ws.row_dimensions[1].height = 35
    ws.column_dimensions["A"].width = 50
    ws.column_dimensions["B"].width = 25
    row = 3
    kpis = [
        ("Documents en entrée", len(docs)),
        ("HIGH confidence", sum(1 for d in docs if d['confidence']=='HIGH')),
        ("LOW (à trier)", sum(1 for d in docs if d['confidence']=='LOW')),
        ("Doublons exacts archivés", dedup["total_exact_dups"]),
        ("Versions obsolètes archivées", dedup["total_obsolete"]),
        ("Dans dataroom principale", sum(1 for d in docs if d.get('status_dedup')=='UNIQUE')),
    ]
    for label, val in kpis:
        ws.cell(row=row, column=1, value=label).font = Font(name="Calibri", size=10, color="0E2841")
        c = ws.cell(row=row, column=2, value=val)
        c.font = Font(name="Calibri", size=11, color="0E2841", bold=True)
        c.alignment = Alignment(horizontal="right", indent=1)
        row += 1

    # Mapping
    ws = wb.create_sheet("Mapping")
    headers = [headers_t["number"], headers_t["original_name"], headers_t["source_path"],
               headers_t["new_name"], headers_t["folder_n1"], headers_t["folder_n2"],
               headers_t["confidence"], headers_t["dedup_status"], headers_t["final_path"]]
    for i, h in enumerate(headers):
        H(ws.cell(row=1, column=i + 1, value=h))
    ws.row_dimensions[1].height = 30

    for i, d in enumerate(docs, 2):
        ws.cell(row=i, column=1, value=i - 1)
        ws.cell(row=i, column=2, value=d["filename"])
        ws.cell(row=i, column=3, value=d["relpath"])
        ws.cell(row=i, column=4, value=d.get("new_name", ""))
        ws.cell(row=i, column=5, value=d.get("folder_n1", ""))
        ws.cell(row=i, column=6, value=d.get("folder_n2", "") or "—")
        cc = ws.cell(row=i, column=7, value=conf_labels.get(d["confidence"], d["confidence"]))
        if d["confidence"] == "HIGH":
            cc.fill = PatternFill("solid", fgColor=MINT)
        elif d["confidence"] == "MEDIUM":
            cc.fill = PatternFill("solid", fgColor=WARNING)
        else:
            cc.fill = PatternFill("solid", fgColor=ALERT)
            cc.font = Font(color=WHITE, bold=True)
        cd = ws.cell(row=i, column=8, value=dedup_labels.get(d.get("status_dedup", ""), ""))
        if d.get("status_dedup") == "DOUBLON_EXACT":
            cd.fill = PatternFill("solid", fgColor=ALERT)
            cd.font = Font(color=WHITE, bold=True)
        elif d.get("status_dedup") == "VERSION_OBSOLETE":
            cd.fill = PatternFill("solid", fgColor=WARNING)
        else:
            cd.fill = PatternFill("solid", fgColor=MINT)
        ws.cell(row=i, column=9, value=d.get("final_path", ""))
        for col in range(1, 10):
            c = ws.cell(row=i, column=col)
            c.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
            if not c.font.bold:
                c.font = Font(name="Calibri", size=9, color=c.font.color)
    widths = [5, 55, 55, 65, 22, 28, 12, 18, 70]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "C2"

    # Structure dataroom
    ws = wb.create_sheet("Structure")
    ws["A1"] = "Arborescence de la dataroom"
    ws["A1"].font = Font(name="Cardo", size=18, color=NAVY)
    ws.column_dimensions["A"].width = 60
    ws.column_dimensions["B"].width = 12
    row = 3
    H(ws.cell(row=row, column=1, value=headers_t["folder_n1"]))
    H(ws.cell(row=row, column=2, value="# docs"))
    row += 1
    by_n1 = defaultdict(int)
    by_n2 = defaultdict(lambda: defaultdict(int))
    for d in docs:
        if d.get("status_dedup") != "UNIQUE":
            continue
        n1 = d.get("folder_n1", "")
        n2 = d.get("folder_n2", "")
        by_n1[n1] += 1
        if n2:
            by_n2[n1][n2] += 1
    for n1 in sorted(by_n1):
        if by_n1[n1] == 0:
            continue
        ws.cell(row=row, column=1, value=n1).font = Font(name="Calibri", size=11, bold=True, color=NAVY)
        ws.cell(row=row, column=2, value=by_n1[n1]).font = Font(name="Calibri", size=11, bold=True)
        row += 1
        for n2 in sorted(by_n2[n1]):
            ws.cell(row=row, column=1, value=f"    └─ {n2}").font = Font(name="Calibri", size=10, color="0E2841")
            ws.cell(row=row, column=2, value=by_n2[n1][n2]).font = Font(name="Calibri", size=10)
            row += 1

    excel_path = out_dir / "_Rapport DataPrep.xlsx"
    wb.save(excel_path)
    return excel_path


# ============================================================================
# MAIN
# ============================================================================

def main():
    parser = argparse.ArgumentParser(description="DataPrep pipeline complet")
    parser.add_argument("--source", help="Chemin source des documents")
    parser.add_argument("--extracted-json", help="JSON déjà extrait (skip extract)")
    parser.add_argument("--destination", required=True, help="Chemin destination de la dataroom")
    parser.add_argument("--project-name", required=True, help="Nom du projet (préfixe fichiers)")
    parser.add_argument("--folder-lang", default="fr", help="Langue de la structure (fr/en/de/it/es)")
    parser.add_argument("--excel-lang", default="fr", help="Langue de l'Excel (fr/en/de/it/es)")
    parser.add_argument("--make-zip", action="store_true", help="Génère le ZIP final")
    args = parser.parse_args()

    destination = Path(args.destination)
    project = args.project_name
    destination.mkdir(parents=True, exist_ok=True)

    if not args.source and not args.extracted_json:
        print("ERREUR : --source ou --extracted-json requis")
        sys.exit(1)

    print(f"=== dataroom-prep pipeline ===")
    print(f"Destination: {destination}")
    print(f"Projet     : {project}")
    print(f"Lang dossiers: {args.folder_lang}  /  Lang Excel: {args.excel_lang}\n")

    # 1. Extract OR load
    if args.extracted_json:
        print(f"[1/5] Chargement depuis {args.extracted_json}...")
        docs = json.loads(Path(args.extracted_json).read_text(encoding="utf-8"))
        print(f"      {len(docs)} documents chargés")
    else:
        source = Path(args.source)
        if not source.exists():
            print(f"ERREUR : source introuvable {source}")
            sys.exit(1)
        print(f"[1/5] Extraction texte depuis {source}...")
        docs = extract_all(source)
        print(f"      {len(docs)} documents extraits")

    # 2. Classify
    print("\n[2/5] Classification multilingue...")
    folder_template = load_template(args.folder_lang)
    for d in docs:
        n1, n2, conf = classify(d, folder_template)
        d["folder_n1"] = n1
        d["folder_n2"] = sanitize(n2) if n2 else ""
        d["confidence"] = conf
        d["new_name"] = d["filename"] if conf == "LOW" else build_new_name(d, n2 or "", project)
    high = sum(1 for d in docs if d["confidence"] == "HIGH")
    low = sum(1 for d in docs if d["confidence"] == "LOW")
    print(f"      HIGH: {high}  /  LOW: {low}")

    # 3. Dedup
    print("\n[3/5] Détection doublons + versions...")
    dedup = detect_duplicates(docs)
    print(f"      Doublons exacts: {dedup['total_exact_dups']}  /  Versions obsolètes: {dedup['total_obsolete']}")

    # 4. Build
    print("\n[4/5] Construction dataroom N1/N2...")
    build_result = build_dataroom(docs, destination, project, folder_template)
    print(f"      Dataroom: {build_result['copied_main']} docs")
    print(f"      Archive doublons: {build_result['copied_dup']}")
    print(f"      Archive versions: {build_result['copied_ver']}")

    # 5. Excel
    print("\n[5/5] Génération Excel mapping...")
    excel_template = load_template(args.excel_lang)
    excel_path = generate_excel(docs, dedup, build_result["out_dir"], project, excel_template)
    print(f"      Excel: {excel_path.name}")

    # ZIP optionnel
    if args.make_zip:
        zip_path = destination / f"{project}.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
            for p in build_result["out_dir"].rglob("*"):
                if p.is_file():
                    zf.write(p, p.relative_to(build_result["out_dir"].parent))
        print(f"\n      ZIP: {zip_path.name} ({zip_path.stat().st_size/1024/1024:.1f} Mo)")

    print(f"\n✓ Pipeline complet terminé")
    print(f"  Dataroom : {build_result['out_dir']}")
    print(f"  Excel    : {excel_path}")


if __name__ == "__main__":
    main()
