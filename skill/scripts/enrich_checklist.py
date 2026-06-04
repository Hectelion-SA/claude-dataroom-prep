"""dataroom-prep — Phase 2 : checklist exhaustive des documents à demander client.

Génère un Excel "Liste documents à demander.xlsx" enrichi via :
- Site web société (scrapé en amont par Claude via firecrawl_scrape, passé en --website-content)
- Localisation société (FR/CH/BE/LU/autre)
- Secteur d'activité (déclaré OU détecté via site web)
- Sources juridiques/fiscales (à consulter par Claude en amont, passées en --legal-context)

CLI:
    python enrich_checklist.py \
        --dataroom "<path Project Acme>" \
        --extracted-json "extracted.json" \
        --location "France" \
        --sector "MedTech" \
        --website-content "scraped_homepage.txt" \
        --legal-context "legal_research.txt" \
        --language fr
"""
import argparse
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

NAVY = "182E4E"
SOFT_SKY = "D3E7FF"
MINT = "6FCF9A"
WARNING = "FFC000"
ALERT = "C00000"
WHITE = "FFFFFF"


# Checklist de base M&A (générique, 60+ items) — extensible par localisation/secteur
BASE_CHECKLIST = {
    "01_Informations générales": [
        ("Présentation entreprise / pitch / corporate deck", "CRITICAL", "Pratique standard M&A — premier livrable"),
        ("Organigramme du groupe + filiales + participations", "CRITICAL", "Cartographie corporate complète"),
        ("Liste des produits / services", "HIGH", "Catalogue commercial à jour"),
        ("Catalogue / brochures commerciales", "MEDIUM", "Marketing collateral"),
        ("Communiqués de presse 3 dernières années", "MEDIUM", "Image publique de la société"),
    ],
    "02_Finance": [
        ("Bilans 3 derniers exercices (N, N-1, N-2)", "CRITICAL", "Pratique standard DD financière"),
        ("Comptes de résultat 3 derniers exercices", "CRITICAL", "Idem"),
        ("Flux de trésorerie 3 ans", "HIGH", "Analyse capacité d'autofinancement"),
        ("Annexes comptables aux bilans", "HIGH", "Détails normés (immo, créances, dettes)"),
        ("Business plan 3-5 ans + hypothèses", "CRITICAL", "Projections valorisation"),
        ("Budget de l'année en cours + suivi", "HIGH", "Tracking court terme"),
        ("État détaillé des emprunts + lignes de crédit", "HIGH", "Endettement net"),
        ("Cap table à jour + historique levées de fonds", "CRITICAL", "Structure capital"),
        ("Pacte d'actionnaires (si existe)", "CRITICAL", "Clauses préemption, drag/tag-along"),
        ("Plan d'investissement (CAPEX) historique + prévu", "MEDIUM", "Politique CAPEX"),
        ("BFR détaillé par poste (créances, stocks, dettes)", "HIGH", "Analyse cycle d'exploitation"),
        ("Vieillissement créances clients + impayés", "HIGH", "Risque crédit"),
        ("Encours auprès des principaux clients", "HIGH", "Concentration commerciale"),
    ],
    "03_Légal": [
        ("Statuts à jour (dernière version signée)", "CRITICAL", "Document fondateur"),
        ("Extrait Kbis / Registre du commerce récent (<3 mois)", "CRITICAL", "Preuve d'existence"),
        ("Registre des bénéficiaires effectifs", "CRITICAL", "Obligation Sapin II / 5e directive AML"),
        ("PV d'AG des 3 dernières années", "HIGH", "Décisions actionnariales"),
        ("PV de CA / conseil de surveillance", "HIGH", "Gouvernance"),
        ("Term sheets / LOI passées (vente, levée, partenariat)", "MEDIUM", "Historique négociations"),
        ("Liste des contrats clients top 10 (avec CA)", "CRITICAL", "Concentration et engagement"),
        ("Liste des contrats fournisseurs critiques", "HIGH", "Dépendances opérationnelles"),
        ("Contrats de licence / IP / royalties", "HIGH", "Engagements commerciaux"),
        ("Liste des brevets, marques, modèles", "HIGH", "Patrimoine immatériel"),
        ("Contentieux en cours + historique 5 ans", "CRITICAL", "Risques juridiques cachés"),
        ("Mises en demeure reçues / envoyées", "HIGH", "Litiges latents"),
        ("Polices d'assurance souscrites (RC, multirisque, etc.)", "HIGH", "Couverture risques"),
        ("NDA templates utilisés en interne", "MEDIUM", "Pratique confidentialité"),
    ],
    "04_Fiscal": [
        ("Liasse fiscale 3 derniers exercices", "CRITICAL", "Document fiscal officiel"),
        ("Avis d'imposition IS + cotisations sociales", "HIGH", "Conformité fiscale"),
        ("Déclarations TVA 12 derniers mois", "HIGH", "Conformité TVA"),
        ("Conventions intragroupe (si filiales)", "HIGH", "Risque transfer pricing"),
        ("Notifications de redressement fiscal 5 ans", "CRITICAL", "Passifs latents"),
        ("Crédit d'Impôt Recherche (CIR) si applicable", "MEDIUM", "Crédit d'impôt"),
        ("Quitus fiscal récent (URSSAF, impôts)", "HIGH", "Attestation conformité"),
    ],
    "05_IT": [
        ("Architecture SI (schéma technique)", "MEDIUM", "Compréhension stack"),
        ("Liste logiciels + licences en cours", "MEDIUM", "Risques licensing"),
        ("Politique RGPD / Registre des traitements", "HIGH", "Conformité RGPD"),
        ("DPA avec sous-traitants externes", "HIGH", "Sous-traitance RGPD"),
        ("Politique de sauvegarde / DRP", "MEDIUM", "Résilience IT"),
        ("Audit cybersécurité récent (si existe)", "MEDIUM", "Maturité cyber"),
        ("Incidents de sécurité signalés CNIL 5 ans", "HIGH", "Notifications obligatoires"),
    ],
    "06_Immobilier": [
        ("Baux commerciaux + locataires + échéances", "HIGH", "Engagements immobiliers"),
        ("Titres de propriété (si propriétaire)", "HIGH", "Patrimoine immobilier"),
        ("Diagnostics obligatoires (DPE, amiante, plomb, etc.)", "HIGH", "Conformité bâtiment"),
        ("Inventaire des actifs immobilisés", "MEDIUM", "Comptabilité actifs"),
    ],
    "07_Assurances": [
        ("Police RC professionnelle + plafonds", "HIGH", "Couverture professionnelle"),
        ("Police multirisque entreprise", "HIGH", "Couverture biens"),
        ("Police cyber-risque (si souscrite)", "MEDIUM", "Couverture digitale"),
        ("Historique sinistres 5 ans + indemnisations", "HIGH", "Sinistralité"),
        ("Polices spécifiques (transport, RC produits, etc.)", "MEDIUM", "Couvertures sectorielles"),
    ],
    "08_RH": [
        ("Liste salariés + ancienneté + rémunération brute", "CRITICAL", "Masse salariale détaillée"),
        ("Contrats de travail types (CDI, CDD, alternant)", "HIGH", "Engagements RH standards"),
        ("Convention collective applicable + accords d'entreprise", "MEDIUM", "Cadre conventionnel"),
        ("Plans d'épargne (PEE, PERCO, intéressement)", "MEDIUM", "Avantages sociaux"),
        ("Plan de stock-options / BSPCE / actions gratuites", "HIGH", "Engagements equity employés"),
        ("Liste des contentieux prud'hommes 5 ans", "HIGH", "Risques sociaux"),
        ("Audit social / DSN récente", "MEDIUM", "Conformité sociale"),
        ("Politique télétravail + RH digital", "LOW", "Pratiques modernes"),
    ],
    "09_Opérations": [
        ("Top 20 clients + CA + ancienneté + concentration", "CRITICAL", "Risque concentration"),
        ("Top 20 fournisseurs + dépendances", "HIGH", "Risque chaîne logistique"),
        ("Carnet de commandes (backlog) actuel", "HIGH", "Visibilité commerciale"),
        ("Pipeline commercial (opportunités qualifiées)", "MEDIUM", "Croissance future"),
        ("KPIs opérationnels mensuels (12 mois)", "MEDIUM", "Performance suivi"),
        ("Liste partenaires stratégiques / distributeurs", "MEDIUM", "Réseau commercial"),
    ],
    "10_Processus": [
        ("Procédures qualité ISO 9001 (si certifié)", "MEDIUM", "Management qualité"),
        ("Certifications produits sectorielles (CE, FDA, ISO...)", "HIGH", "Conformité sectorielle"),
        ("Manuel qualité / procédures opérationnelles", "MEDIUM", "Industrialisation"),
        ("Carte des processus métier", "LOW", "Maturité opérationnelle"),
    ],
}

# Compléments par localisation
LOCATION_ADDITIONS = {
    "France": {
        "04_Fiscal": [
            ("CFE / CVAE déclarations 3 ans", "MEDIUM", "Fiscalité locale française"),
            ("Crédit Impôt Innovation (CII) si applicable", "MEDIUM", "Aide R&D PME"),
        ],
        "08_RH": [
            ("DUERP (Document Unique d'Évaluation des Risques)", "HIGH", "Obligation Code du travail FR"),
            ("Index égalité Hommes/Femmes (si >50 salariés)", "HIGH", "Obligation loi Avenir Pro"),
            ("Procès-verbal du CSE 12 derniers mois", "HIGH", "Représentation salariés FR"),
        ],
        "03_Légal": [
            ("Attestation Sapin II (lutte anti-corruption)", "MEDIUM", "Obligation loi Sapin II"),
            ("Politique RGPD écrite + nomination DPO", "HIGH", "Obligation RGPD UE"),
        ],
    },
    "Suisse": {
        "04_Fiscal": [
            ("Décisions de taxation cantonale 3 ans", "CRITICAL", "Fiscalité cantonale CH"),
            ("Bordereaux AVS / LPP / SUVA", "HIGH", "Assurances sociales CH"),
            ("Décompte TVA Suisse trimestriels", "HIGH", "TVA fédérale CH"),
        ],
        "08_RH": [
            ("Permis de travail employés étrangers (B, L, C)", "HIGH", "Conformité OFM CH"),
            ("Certificats de salaire", "MEDIUM", "Pratique RH CH"),
        ],
        "03_Légal": [
            ("Extrait du Registre du commerce cantonal récent", "CRITICAL", "Document légal CH"),
            ("Statuts conformes au CO (Code des obligations)", "CRITICAL", "Droit suisse"),
        ],
        "07_Assurances": [
            ("Assurance LAA (accidents professionnels)", "HIGH", "Obligation LAA CH"),
        ],
    },
    "Belgique": {
        "04_Fiscal": [
            ("Précomptes professionnels (3 ans)", "HIGH", "Fiscalité salariale BE"),
            ("Bilan social annuel BNB", "MEDIUM", "Dépôt Banque Nationale BE"),
        ],
    },
    "Luxembourg": {
        "04_Fiscal": [
            ("Bilans déposés au Registre de Commerce LU", "CRITICAL", "RCS LU"),
            ("Avis de taxation IRC / ICC", "HIGH", "Fiscalité LU"),
        ],
    },
}

# Compléments par secteur
SECTOR_ADDITIONS = {
    "MedTech": {
        "10_Processus": [
            ("Marquage CE MDR (Règlement 2017/745)", "CRITICAL", "Obligation mise sur marché UE"),
            ("Certification ISO 13485 (qualité dispositifs médicaux)", "CRITICAL", "Norme sectorielle MedTech"),
            ("Dossier technique des produits (DMR)", "CRITICAL", "Documentation MDR"),
            ("Plan de surveillance post-marché (PMS)", "HIGH", "Obligation post-MDR"),
            ("Rapports d'incidents matériovigilance", "HIGH", "Veille sécurité"),
        ],
        "05_IT": [
            ("Conformité SaMD / HIPAA si app de santé US", "HIGH", "Si produit logiciel médical"),
        ],
    },
    "SaaS": {
        "05_IT": [
            ("DPA (Data Processing Agreement) avec clients", "HIGH", "Obligation RGPD sous-traitance"),
            ("Hébergement certifié HDS (si données santé)", "HIGH", "Hébergeur Données Santé FR"),
            ("Audit pentest récent (1-2 ans)", "MEDIUM", "Sécurité produit"),
            ("SLA + SOC 2 (si applicable)", "MEDIUM", "Engagement clients enterprise"),
        ],
        "09_Opérations": [
            ("MRR / ARR / Churn historique 24 mois", "CRITICAL", "Métriques SaaS standard"),
            ("Cohortes clients (rétention)", "HIGH", "Analyse cohortes"),
            ("CAC / LTV par segment", "HIGH", "Unit economics"),
        ],
    },
    "Immobilier": {
        "06_Immobilier": [
            ("Diagnostics obligatoires complets (DPE, amiante, plomb, gaz, élec, ERP)", "CRITICAL", "Vente immobilière FR"),
            ("Surface Carrez certifiée (logements en copropriété)", "HIGH", "Loi Carrez FR"),
            ("Carnet d'entretien immeuble (copropriété)", "MEDIUM", "Conformité loi ELAN"),
        ],
        "07_Assurances": [
            ("Police dommage-ouvrage (si construction <10 ans)", "HIGH", "Loi Spinetta FR"),
        ],
    },
    "Retail": {
        "10_Processus": [
            ("Conformité étiquetage produits (langue, allergènes)", "HIGH", "Réglementation consommateurs"),
            ("Politique retours / réclamations", "MEDIUM", "Service clients"),
        ],
    },
    "Industriel / Manufacturier": {
        "06_Immobilier": [
            ("Arrêté ICPE (Installation Classée Protection Environnement)", "CRITICAL", "Obligation si ICPE FR"),
            ("Études d'impact environnemental", "HIGH", "Conformité ICPE"),
        ],
        "07_Assurances": [
            ("Police pollution / environnement", "HIGH", "Risques industriels"),
        ],
    },
}


def detect_present(docs: list, keywords: list) -> bool:
    """Test si un item est présent dans le corpus (par mots-clés dans filename + theme)."""
    for d in docs:
        text = (d.get("filename", "") + " " + d.get("theme", "") + " " + d.get("folder_n2", "")).lower()
        if any(kw.lower() in text for kw in keywords):
            return True
    return False


def _extract_keywords(item_label: str) -> list:
    """Heuristique pour extraire des mots-clés à matcher."""
    # Nettoie ponctuation, garde les mots significatifs
    txt = re.sub(r"[^\w\s]", " ", item_label.lower())
    stops = {"de", "des", "du", "la", "le", "les", "et", "ou", "à", "au", "aux", "en", "ans", "année",
             "années", "n", "n-1", "n-2", "si", "applicable", "existe", "an", "12", "3", "5", "10",
             "derniers", "dernières", "récent", "obligatoires", "actuel"}
    words = [w for w in txt.split() if len(w) > 3 and w not in stops]
    # Garde les 3-4 premiers mots significatifs
    return words[:4]


def build_checklist(docs: list, location: str, sector: str, website_summary: str = "") -> list:
    """Construit la checklist enrichie."""
    checklist = []
    base = dict(BASE_CHECKLIST)

    # Compléments localisation
    if location in LOCATION_ADDITIONS:
        for section, items in LOCATION_ADDITIONS[location].items():
            if section in base:
                base[section] = base[section] + items
            else:
                base[section] = items

    # Compléments secteur
    if sector in SECTOR_ADDITIONS:
        for section, items in SECTOR_ADDITIONS[sector].items():
            if section in base:
                base[section] = base[section] + items
            else:
                base[section] = items

    for section, items in base.items():
        for item_label, criticality, rationale in items:
            keywords = _extract_keywords(item_label)
            present = detect_present(docs, keywords)
            checklist.append({
                "section": section,
                "item": item_label,
                "criticality": criticality,
                "rationale": rationale,
                "status": "Présent" if present else "À demander",
                "source": "Base" if (section in BASE_CHECKLIST and item_label in [i[0] for i in BASE_CHECKLIST.get(section, [])])
                          else (f"Localisation: {location}" if location in LOCATION_ADDITIONS and item_label in [i[0] for i in LOCATION_ADDITIONS[location].get(section, [])]
                          else f"Secteur: {sector}"),
            })

    if website_summary:
        # Ajoute une ligne "contexte société" en début
        checklist.insert(0, {
            "section": "00_Contexte société",
            "item": f"Site web analysé : {website_summary[:200]}",
            "criticality": "INFO",
            "rationale": "Analyse contextuelle automatique du site web",
            "status": "Info",
            "source": "Firecrawl",
        })

    return checklist


def write_checklist_excel(checklist: list, dataroom_dir: Path, project: str, location: str, sector: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Liste documents à demander"

    # Titre
    ws["A1"] = f"Liste documents à demander — {project}"
    ws["A1"].font = Font(name="Cardo", size=22, color=NAVY)
    ws["A2"] = f"Localisation: {location} | Secteur: {sector}"
    ws["A2"].font = Font(name="Cardo", size=14, color="0E2841", italic=True)
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 22

    # Headers
    headers = ["#", "Catégorie", "Document / Information demandé", "Criticité", "Statut", "Justification", "Source"]
    for i, h in enumerate(headers):
        c = ws.cell(row=4, column=i + 1, value=h)
        c.font = Font(name="Calibri", size=11, color=WHITE, bold=True)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[4].height = 28

    for i, item in enumerate(checklist, 5):
        ws.cell(row=i, column=1, value=i - 4)
        ws.cell(row=i, column=2, value=item["section"])
        ws.cell(row=i, column=3, value=item["item"])
        cc = ws.cell(row=i, column=4, value=item["criticality"])
        if item["criticality"] == "CRITICAL":
            cc.fill = PatternFill("solid", fgColor=ALERT)
            cc.font = Font(color=WHITE, bold=True)
        elif item["criticality"] == "HIGH":
            cc.fill = PatternFill("solid", fgColor=WARNING)
        elif item["criticality"] == "MEDIUM":
            cc.fill = PatternFill("solid", fgColor=SOFT_SKY)
        elif item["criticality"] == "LOW":
            cc.fill = PatternFill("solid", fgColor="F2F2F2")
        cc.alignment = Alignment(horizontal="center", vertical="center")

        cs = ws.cell(row=i, column=5, value=item["status"])
        if item["status"] == "Présent":
            cs.fill = PatternFill("solid", fgColor=MINT)
        elif item["status"] == "Info":
            cs.fill = PatternFill("solid", fgColor=SOFT_SKY)
        else:
            cs.fill = PatternFill("solid", fgColor=WARNING)
        cs.font = Font(bold=True)
        cs.alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(row=i, column=6, value=item["rationale"])
        ws.cell(row=i, column=7, value=item["source"])

        for col in range(1, 8):
            c = ws.cell(row=i, column=col)
            c.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
            if not c.font.bold and col not in (4, 5):
                c.font = Font(name="Calibri", size=10, color="0E2841")

    widths = [5, 30, 65, 12, 14, 50, 25]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A5"

    out_path = dataroom_dir / "_Liste documents à demander.xlsx"
    wb.save(out_path)
    return out_path


def main():
    parser = argparse.ArgumentParser(description="Phase 2 - Checklist enrichie")
    parser.add_argument("--dataroom", required=True, help="Chemin dataroom (Project X)")
    parser.add_argument("--extracted-json", required=True, help="JSON des docs déjà classifiés")
    parser.add_argument("--location", default="France", help="France/Suisse/Belgique/Luxembourg/autre")
    parser.add_argument("--sector", default="Generic", help="MedTech/SaaS/Immobilier/Retail/Industriel")
    parser.add_argument("--website-content", help="Chemin fichier contenu site web scrapé")
    parser.add_argument("--language", default="fr", help="Langue de sortie")
    args = parser.parse_args()

    dataroom = Path(args.dataroom)
    docs = json.loads(Path(args.extracted_json).read_text(encoding="utf-8"))

    print(f"=== Phase 2 : Checklist enrichie ===")
    print(f"Dataroom    : {dataroom}")
    print(f"Localisation: {args.location}")
    print(f"Secteur     : {args.sector}")

    website_summary = ""
    if args.website_content and Path(args.website_content).exists():
        content = Path(args.website_content).read_text(encoding="utf-8")
        website_summary = content[:500].replace("\n", " ")
        print(f"Site web    : {len(content)} caractères chargés")

    checklist = build_checklist(docs, args.location, args.sector, website_summary)
    print(f"\nChecklist générée : {len(checklist)} items")
    n_present = sum(1 for c in checklist if c["status"] == "Présent")
    n_missing = sum(1 for c in checklist if c["status"] == "À demander")
    print(f"  Présents     : {n_present}")
    print(f"  À demander   : {n_missing}")

    project = dataroom.name
    out_path = write_checklist_excel(checklist, dataroom, project, args.location, args.sector)
    print(f"\n✓ Checklist sauvegardée : {out_path}")


if __name__ == "__main__":
    main()
