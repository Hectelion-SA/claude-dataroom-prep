"""dataroom-prep — Pipeline principal (extract + classify + dedup + build + Excel).

CLI:
    python run_pipeline.py \
        --source "<chemin source>" \
        --destination "<chemin destination>" \
        --project-name "Project Medicaps" \
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

# Extraction enrichie : PyMuPDF (meilleure que pypdf) + OCR optionnel
try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

try:
    import pytesseract  # noqa: F401
    from PIL import Image  # noqa: F401
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

sys.stdout.reconfigure(encoding="utf-8")

# Dossier tessdata maison (indépendant de l'install système) : garantit que fra+eng
# sont toujours disponibles même si le paquet OS/winget n'embarque que eng.
_TESSDATA_DIR = Path(__file__).parent.parent / "tessdata"
_TESSDATA_URLS = {
    "eng": "https://github.com/tesseract-ocr/tessdata_fast/raw/main/eng.traineddata",
    "fra": "https://github.com/tesseract-ocr/tessdata_fast/raw/main/fra.traineddata",
}


def _ensure_tessdata():
    """Télécharge eng/fra.traineddata dans un dossier local si absents. Best-effort."""
    import urllib.request
    _TESSDATA_DIR.mkdir(parents=True, exist_ok=True)
    for lang, url in _TESSDATA_URLS.items():
        dest = _TESSDATA_DIR / f"{lang}.traineddata"
        if dest.exists() and dest.stat().st_size > 0:
            continue
        try:
            urllib.request.urlretrieve(url, dest)
        except Exception:
            try:
                dest.unlink(missing_ok=True)
            except Exception:
                pass
    if (_TESSDATA_DIR / "eng.traineddata").exists() and (_TESSDATA_DIR / "fra.traineddata").exists():
        os.environ["TESSDATA_PREFIX"] = str(_TESSDATA_DIR)


def ensure_ocr_stack():
    """Best-effort auto-install of the OCR stack (pytesseract, Pillow, tesseract binary,
    langue fra+eng) so scanned PDFs get OCR'd instead of silently falling into
    11_Document à trier. Silent and non-fatal: any failure here just leaves HAS_OCR as-is."""
    global HAS_OCR
    if HAS_OCR and shutil.which("tesseract"):
        _ensure_tessdata()
        return
    import subprocess
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "--quiet", "pytesseract", "pillow"],
                       check=False, capture_output=True, timeout=120)
    except Exception:
        pass
    if not shutil.which("tesseract"):
        try:
            if sys.platform == "win32":
                subprocess.run(["winget", "install", "--id", "UB-Mannheim.TesseractOCR", "-e",
                                 "--silent", "--accept-package-agreements", "--accept-source-agreements"],
                               check=False, capture_output=True, timeout=300)
            elif sys.platform == "darwin":
                subprocess.run(["brew", "install", "tesseract"], check=False, capture_output=True, timeout=300)
            else:
                subprocess.run(["sudo", "-n", "apt-get", "install", "-y", "tesseract-ocr"],
                               check=False, capture_output=True, timeout=300)
        except Exception:
            pass
    try:
        import pytesseract as _pt
        from PIL import Image as _img  # noqa: F401
        globals()["pytesseract"] = _pt
        globals()["Image"] = _img
        if sys.platform == "win32" and not shutil.which("tesseract"):
            guess = Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Tesseract-OCR"
            if (guess / "tesseract.exe").exists():
                os.environ["PATH"] += os.pathsep + str(guess)
        HAS_OCR = bool(shutil.which("tesseract"))
        if HAS_OCR:
            _ensure_tessdata()
    except ImportError:
        HAS_OCR = False

# ============================================================================
# CONSTANTS
# ============================================================================

SKILL_DIR = Path(__file__).parent.parent
TEMPLATES_DIR = SKILL_DIR / "templates"

# Noms réels à anonymiser (société cible + marques/entités) : renseigné via --scrub.
SCRUB_NAMES: list = []

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


def _ocr_page(page) -> str:
    """OCR d'une page PyMuPDF rendue en image. Vide si tesseract indisponible."""
    if not HAS_OCR:
        return ""
    try:
        import io
        pix = page.get_pixmap(dpi=200)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        return pytesseract.image_to_string(img, lang="fra+eng")
    except Exception:
        try:
            return pytesseract.image_to_string(img)
        except Exception:
            return ""


def _extract_pdf(path: Path) -> tuple[str, str]:
    # Voie 1 : PyMuPDF (couche texte native, plus fiable que pypdf)
    if HAS_FITZ:
        try:
            d = fitz.open(str(path))
            n = d.page_count
            text = ""
            for i in range(min(3, n)):
                text += d[i].get_text() + "\n"
                if len(text) > MAX_CHARS:
                    break
            # Voie 2 : PDF scanné (pas de couche texte) -> OCR fallback
            if len(text.strip()) < 30 and n:
                ocr = ""
                for i in range(min(2, n)):
                    ocr += _ocr_page(d[i]) + "\n"
                    if len(ocr) > MAX_CHARS:
                        break
                if ocr.strip():
                    return ocr[:MAX_CHARS], f"pdf:{n}p:OCR"
                return "", f"pdf:{n}p:SCANNED_NO_OCR"
            return text[:MAX_CHARS], f"pdf:{n}p"
        except Exception as e:
            return "", f"pdf:ERROR:{type(e).__name__}"
    # Voie 3 : fallback pypdf si PyMuPDF absent
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
            "marque_bodyo": "Marque Bodyo",
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
            "marque_bodyo": "Bodyo Trademark",
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
            "marque_bodyo": "Bodyo Marke",
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
            "marque_bodyo": "Marchio Bodyo",
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
            "marque_bodyo": "Marca Bodyo",
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
        # Clustering de versions UNIQUEMENT à extension identique :
        # un PDF et un XLSX du même intitulé ne sont pas deux versions à dédupliquer.
        ext = Path(d["filename"]).suffix.lower()
        key = (d["parent_folder"], ext, normalize_for_clustering(d["filename"]))
        by_cluster[key].append(d)

    version_clusters = []
    obs_set = {}
    for (parent, _ext, base), group in by_cluster.items():
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


def scrub_names(text: str, names: list) -> str:
    """Retire toute occurrence des noms interdits (société/marques réelles).

    ANONYMISATION ABSOLUE : le nom réel de la cible ne doit JAMAIS apparaître dans un
    nom de fichier ou de dossier généré — uniquement le nom de PROJET. Insensible à la
    casse, sur limites de mots, puis nettoyage des séparateurs orphelins.
    """
    if not text or not names:
        return text
    for n in names:
        n = (n or "").strip()
        if not n:
            continue
        text = re.sub(rf"\b{re.escape(n)}\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*[-–—]\s*[-–—]\s*", " - ", text)   # tirets devenus consécutifs
    text = re.sub(r"[\s_]+", " ", text).strip(" -–—_.")
    return text or "document"


def sanitize(name: str, max_len: int = 80) -> str:
    name = unicodedata.normalize("NFC", name)
    name = re.sub(r'[<>:"/\\|?*]', " - ", name)
    name = re.sub(r"\s+", " ", name).strip().rstrip(". ")
    return name[:max_len] if name else "Divers"


YYYYMM_RX = re.compile(r"(20\d{2})[\.\-_\s]?(0\d|1[0-2])")
YYYY_RX = re.compile(r"\b(20\d{2})\b")
DDMMYYYY_RX = re.compile(r"\b(0\d|[12]\d|3[01])[\.\-_/](0\d|1[0-2])[\.\-_/](20\d{2})\b")


def extract_date(fname: str, snippet: str) -> str:
    """Retourne une date au format yyyymmdd (ou yyyymm / yyyy si jour/mois absent)."""
    text = fname + " " + (snippet[:300] if snippet else "")
    m = DDMMYYYY_RX.search(text)
    if m:
        return f"{m.group(3)}{m.group(2)}{m.group(1)}"  # yyyymmdd
    m = YYYYMM_RX.search(text)
    if m:
        return f"{m.group(1)}{m.group(2)}"  # yyyymm
    m = YYYY_RX.search(text)
    if m:
        return m.group(1)  # yyyy
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


def build_new_name(doc: dict, project: str, title: str = None, date: str = None) -> str:
    """Convention : {Project} - {Intitulé du document} - {yyyymmdd}.ext

    - PAS de nom de thème/dossier dans le nom de fichier (déjà porté par l'arborescence).
    - PAS de reprise du vieux nom de fichier brut : on privilégie un intitulé propre
      fourni par l'analyse de contenu (title). À défaut, on nettoie le nom d'origine.
    """
    ext = Path(doc["filename"]).suffix.lower()
    if date == "__none__":          # date explicitement supprimée
        date = ""
    elif not date:                   # date absente -> heuristique sur nom + contenu
        date = extract_date(doc["filename"], doc.get("snippet", ""))
    subject = (title or clean_subject(doc["filename"])).strip()
    subject = scrub_names(subject, SCRUB_NAMES)   # anonymisation absolue
    parts = [project, subject, date if date and date != "undated" else ""]
    name = " - ".join(p for p in parts if p)
    return sanitize(name + ext, max_len=200)


def build_dataroom(docs: list, destination: Path, project: str, template: dict) -> dict:
    """Build dataroom structure N1/N2 + archives dedup."""
    out_dir = destination / project
    out_dir.mkdir(parents=True, exist_ok=True)
    # Vide le CONTENU sans supprimer le dossier racine : sur OneDrive / si le dossier
    # est ouvert dans l'Explorateur, rmdir du dossier racine échoue (WinError 32 lock).
    for child in list(out_dir.iterdir()):
        try:
            if child.is_dir():
                shutil.rmtree(child, onerror=_force_remove)
            else:
                os.chmod(child, stat.S_IWRITE)
                child.unlink()
        except Exception:
            pass

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
        ws.cell(row=i, column=2, value=scrub_names(d["filename"], SCRUB_NAMES))
        ws.cell(row=i, column=3, value=scrub_names(d["relpath"], SCRUB_NAMES))
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
# DOCUMENTS MANQUANTS (Top 50 par catégorie via listes DD France/Suisse)
# ============================================================================

DD_FILENAMES = {"FR": "Hectelion - Listes DD - FRANCE.xlsx",
                "CH": "Hectelion - Listes DD - SUISSE.xlsx"}
PRIO_RANK = {"haute": 0, "moyenne": 1, "basse": 2, "": 3}
_STOP = set("de des du la le les et ou un une au aux pour par sur dans en avec sans "
            "document documents information informations liste tous toutes leurs leur "
            "societe société entreprise relatifs relatives relative relatif ainsi que "
            "concernant chaque type types nom noms date dernier derniere dernière".split())


def _singular(w: str) -> str:
    # déspluralisation légère FR/EN : enlève un 's' final (livres->livre, comptes->compte)
    return w[:-1] if len(w) >= 5 and w.endswith("s") else w


def _norm_tokens(text: str) -> set:
    t = unicodedata.normalize("NFKD", str(text or "")).encode("ascii", "ignore").decode().lower()
    toks = re.findall(r"[a-z0-9]+", t)
    return {_singular(w) for w in toks if len(w) >= 4 and w not in _STOP}


def load_dd_checklist(country: str, sector: str, dd_dir: Path) -> list:
    """Parse la liste DD (pays + secteur) -> items {ref, element, type, periode, prio, base, section}."""
    fname = DD_FILENAMES.get(country.upper())
    path = dd_dir / fname
    wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
    if sector not in wb.sheetnames:
        # match tolérant (accents/casse)
        norm = {unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().lower(): s
                for s in wb.sheetnames}
        key = unicodedata.normalize("NFKD", sector).encode("ascii", "ignore").decode().lower()
        sector = norm.get(key, wb.sheetnames[0])
    ws = wb[sector]
    items, section = [], ""
    for r in ws.iter_rows(min_row=5, values_only=True):
        num = str(r[0] or "").strip()
        el = str(r[1] or "").strip()
        if re.match(r"^\d+\.$", num) and el:
            section = el
        elif re.match(r"^\d+\.\d+", num) and el:
            items.append({
                "ref": num, "element": el, "type": str(r[2] or ""),
                "periode": str(r[3] or ""), "prio": str(r[4] or "").strip(),
                "base": str(r[5] or ""), "section": section,
            })
    return items, sector


def _present_tokens(docs: list) -> list:
    sets = []
    for d in docs:
        if d.get("status_dedup") not in (None, "UNIQUE"):
            continue
        if d.get("confidence") == "LOW":
            continue
        toks = _norm_tokens(d.get("new_name", "")) | _norm_tokens(d.get("folder_n2", ""))
        if toks:
            sets.append(toks)
    return sets


# Tokens DD ultra-fréquents : partagés seuls ils ne prouvent rien (évite les faux positifs
# type "Rapport du commissaire aux comptes" matché par {rapport, compte}).
_GENERIC = {"compte", "rapport", "attestation", "conformite", "jour", "autre", "copie",
            "depot", "detail", "dispo", "derniere", "historique", "controle", "document",
            "societe", "management", "liste", "tableau", "suivi", "commentaire"}


def match_missing(items: list, docs: list) -> list:
    """Marque chaque item present/missing.

    Présent si un doc partage >= 2 tokens AVEC au moins un token distinctif (hors _GENERIC) :
    deux mots génériques en commun (ex. 'rapport' + 'compte') ne suffisent pas.
    """
    present_sets = _present_tokens(docs)
    for it in items:
        itoks = _norm_tokens(it["element"])
        present = False
        for ps in present_sets:
            shared = itoks & ps
            if len(shared) >= 2 and (shared - _GENERIC):
                present = True
                break
        it["present"] = present
    return items


def append_missing_sheet(excel_path: Path, items: list, country: str, sector: str,
                         excel_template: dict, top_n: int = 50):
    """Ajoute l'onglet 'Documents à demander' : Top N manquants prioritaires, groupés par catégorie."""
    def _refkey(ref):
        return tuple(int(x) if x.isdigit() else 0 for x in str(ref).strip(". ").split("."))

    missing = [it for it in items if not it["present"]]
    # Ordre des sections selon le template
    section_order, seen = [], set()
    for it in items:
        if it["section"] not in seen:
            seen.add(it["section"]); section_order.append(it["section"])

    # Répartition PROPORTIONNELLE des {top_n} slots entre catégories (au lieu de tout
    # concentrer sur les grosses sections type Finance) : apportionnement au plus fort
    # reste, pondéré par le nb de manquants de chaque section, avec un minimum de 1 par
    # catégorie représentée. Priorité Haute d'abord à l'intérieur de chaque catégorie.
    by_sec = {}
    for it in missing:
        by_sec.setdefault(it["section"], []).append(it)
    for s in by_sec:
        by_sec[s].sort(key=lambda it: (PRIO_RANK.get(it["prio"].lower(), 3), _refkey(it["ref"])))
    secs = [s for s in section_order if by_sec.get(s)]
    total_missing = sum(len(by_sec[s]) for s in secs)

    if total_missing <= top_n:
        quota = {s: len(by_sec[s]) for s in secs}
    else:
        raw = {s: top_n * len(by_sec[s]) / total_missing for s in secs}
        quota = {s: min(int(raw[s]), len(by_sec[s])) for s in secs}
        # min 1 par catégorie représentée (si assez de slots)
        if len(secs) <= top_n:
            for s in secs:
                if quota[s] == 0:
                    quota[s] = 1
        # ajoute les slots restants au plus fort reste (dans la limite de capacité)
        for s in sorted(secs, key=lambda s: raw[s] - int(raw[s]), reverse=True):
            if sum(quota.values()) >= top_n:
                break
            if quota[s] < len(by_sec[s]):
                quota[s] += 1
        # si le min-1 a fait dépasser top_n, retire des plus grosses quotas (>1)
        while sum(quota.values()) > top_n:
            s = max(secs, key=lambda s: quota[s])
            if quota[s] <= 1:
                break
            quota[s] -= 1

    grouped = {s: by_sec[s][:quota.get(s, 0)] for s in section_order if by_sec.get(s)}
    top = [it for s in section_order if s in grouped for it in grouped[s]]

    wb = openpyxl.load_workbook(str(excel_path))
    if "Documents à demander" in wb.sheetnames:
        del wb["Documents à demander"]
    ws = wb.create_sheet("Documents à demander")

    def H(c):
        c.font = Font(name="Calibri", size=11, color=WHITE, bold=True)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    pays = "France" if country.upper() == "FR" else "Suisse"
    ws["A1"] = f"Documents importants manquants — Top {top_n} ({pays} / {sector})"
    ws["A1"].font = Font(name="Cardo", size=18, color=NAVY)
    ws.row_dimensions[1].height = 26
    n_items = len(items); n_present = sum(1 for it in items if it["present"])
    n_haute = sum(1 for it in items if it["prio"].lower() == "haute")
    n_haute_miss = sum(1 for it in items if it["prio"].lower() == "haute" and not it["present"])
    ws["A2"] = (f"Référentiel : {n_items} exigences DD · présents (estimés) : {n_present} · "
                f"manquants : {n_items - n_present} · priorité Haute manquante : {n_haute_miss}/{n_haute}")
    ws["A2"].font = Font(name="Calibri", size=9, italic=True, color="0E2841")

    headers = ["Catégorie", "Réf.", "Document / information à demander", "Priorité", "Période", "Base légale"]
    hrow = 4
    for i, h in enumerate(headers, 1):
        H(ws.cell(row=hrow, column=i, value=h))
    ws.row_dimensions[hrow].height = 28
    row = hrow + 1
    for section in section_order:
        rows_s = grouped.get(section, [])
        if not rows_s:
            continue
        c = ws.cell(row=row, column=1, value=section)
        c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor="0E2841")
        for col in range(2, 7):
            ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor="0E2841")
        row += 1
        for it in rows_s:
            ws.cell(row=row, column=1, value="")
            ws.cell(row=row, column=2, value=it["ref"])
            ws.cell(row=row, column=3, value=it["element"])
            pc = ws.cell(row=row, column=4, value=it["prio"])
            pl = it["prio"].lower()
            pc.fill = PatternFill("solid", fgColor=ALERT if pl == "haute" else (WARNING if pl == "moyenne" else SOFT_SKY))
            if pl == "haute":
                pc.font = Font(name="Calibri", size=9, color=WHITE, bold=True)
            ws.cell(row=row, column=5, value=it["periode"])
            ws.cell(row=row, column=6, value=it["base"])
            for col in range(1, 7):
                cc = ws.cell(row=row, column=col)
                cc.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
                if not cc.font.bold:
                    cc.font = Font(name="Calibri", size=9)
            row += 1
    for col, w in zip("ABCDEF", [26, 7, 70, 11, 26, 34]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"
    wb.save(str(excel_path))
    return len(top), n_haute_miss, n_haute


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
    parser.add_argument("--extract-only", help="Extrait le texte vers ce JSON puis s'arrête (pour titrage IA)")
    parser.add_argument("--decisions-json", help="JSON de décisions IA {relpath: {n1, n2, title, date, confidence}}")
    parser.add_argument("--scrub", default="", help="Noms réels à anonymiser, séparés par des virgules (société + marques)")
    parser.add_argument("--scrub-file", help="Fichier (1 nom par ligne) listant les noms réels à anonymiser. Préférer à --scrub pour que le nom réel n'apparaisse pas dans la commande shell.")
    parser.add_argument("--country", choices=["FR", "CH"], help="Pays de la doc -> liste DD de référence (FR/CH)")
    parser.add_argument("--sector", help="Secteur (onglet de la liste DD : Tech, Sante, Industrie, ...)")
    parser.add_argument("--dd-dir", default=str(SKILL_DIR / "data_sources"),
                        help="Dossier contenant les listes DD France/Suisse")
    args = parser.parse_args()

    # Anonymisation absolue : aucun nom réel ne doit apparaître, seul le nom de projet.
    # Les noms réels ne sont JAMAIS imprimés (ni console, ni log) — on n'affiche qu'un compte.
    global SCRUB_NAMES
    SCRUB_NAMES = [n.strip() for n in args.scrub.split(",") if n.strip()]
    if args.scrub_file:
        try:
            with open(args.scrub_file, "r", encoding="utf-8") as f:
                SCRUB_NAMES += [ln.strip() for ln in f if ln.strip()]
        except OSError as e:
            print(f"[anonymisation] avertissement : fichier --scrub-file illisible ({e.strerror})")
    # dédoublonnage en conservant l'ordre
    SCRUB_NAMES = list(dict.fromkeys(SCRUB_NAMES))
    if SCRUB_NAMES:
        print(f"[anonymisation] {len(SCRUB_NAMES)} nom(s) scrubé(s) des fichiers/dossiers/Excel (non affichés).")

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
        if not (HAS_OCR and shutil.which("tesseract")):
            print("      OCR absent -> installation automatique en cours (pytesseract, Pillow, Tesseract)...")
            ensure_ocr_stack()
        ocr_flag = " (OCR actif)" if HAS_OCR else " (OCR indisponible : installation auto échouée, pytesseract/tesseract absent)"
        engine = "PyMuPDF" if HAS_FITZ else "pypdf"
        print(f"[1/5] Extraction texte depuis {source}... [moteur {engine}{ocr_flag}]")
        docs = extract_all(source)
        print(f"      {len(docs)} documents extraits")

    # Mode extraction seule : dump JSON et stop (permet titrage IA externe)
    if args.extract_only:
        Path(args.extract_only).write_text(
            json.dumps(docs, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\n✓ Extraction écrite : {args.extract_only}")
        print("  Étape suivante : générer les décisions IA puis relancer avec --decisions-json")
        return

    # 2. Classify (heuristiques) puis surcharge par décisions IA si fournies
    print("\n[2/5] Classification...")
    folder_template = load_template(args.folder_lang)
    folders_n1 = folder_template["folders_n1"]

    decisions = {}
    if args.decisions_json:
        raw = json.loads(Path(args.decisions_json).read_text(encoding="utf-8"))
        # Normalisation des clés (NFC + slashes) pour matcher quel que soit l'encodage
        # des accents (NFC vs NFD) ou le séparateur de chemin.
        def _norm_key(k):
            return unicodedata.normalize("NFC", k).replace("\\", "/")
        decisions = {_norm_key(k): v for k, v in raw.items()}
        print(f"      {len(decisions)} décisions IA chargées (titres + classement par contenu)")

    n_ai = 0
    for d in docs:
        rk = unicodedata.normalize("NFC", d["relpath"]).replace("\\", "/")
        dec = decisions.get(rk)
        if dec:
            n_ai += 1
            d["folder_n1"] = folders_n1.get(str(dec["n1"]), folders_n1["11"])
            d["folder_n2"] = sanitize(scrub_names(dec.get("n2", ""), SCRUB_NAMES)) if dec.get("n2") else ""
            d["confidence"] = dec.get("confidence", "HIGH")
            if d["confidence"] == "LOW":
                d["new_name"] = d["filename"]
            else:
                # Si "date" est fournie explicitement (même vide), on la respecte telle
                # quelle (vide = pas de suffixe date). Si absente, fallback heuristique.
                if "date" in dec:
                    dval = str(dec.get("date") or "").strip()
                    d["new_name"] = build_new_name(d, project, title=dec.get("title"), date=dval or "__none__")
                else:
                    d["new_name"] = build_new_name(d, project, title=dec.get("title"))
        else:
            n1, n2, conf = classify(d, folder_template)
            d["folder_n1"] = n1
            d["folder_n2"] = sanitize(scrub_names(n2, SCRUB_NAMES)) if n2 else ""
            d["confidence"] = conf
            d["new_name"] = d["filename"] if conf == "LOW" else build_new_name(d, project)
    high = sum(1 for d in docs if d["confidence"] == "HIGH")
    low = sum(1 for d in docs if d["confidence"] == "LOW")
    print(f"      HIGH: {high}  /  LOW: {low}  (dont {n_ai} titrés par analyse de contenu)")

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

    # 5bis. Documents manquants (Top 50 par catégorie via liste DD pays + secteur)
    if args.country and args.sector:
        print("\n[5bis] Documents manquants (liste DD)...")
        try:
            items, sector = load_dd_checklist(args.country, args.sector, Path(args.dd_dir))
            items = match_missing(items, docs)
            n_top, n_hm, n_h = append_missing_sheet(excel_path, items, args.country, sector, excel_template)
            print(f"      Référentiel {args.country}/{sector} : {len(items)} exigences")
            print(f"      Onglet 'Documents à demander' : Top {n_top} manquants  "
                  f"(priorité Haute manquante : {n_hm}/{n_h})")
        except FileNotFoundError:
            print(f"      ⚠ Liste DD introuvable dans {args.dd_dir} — étape ignorée")

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
