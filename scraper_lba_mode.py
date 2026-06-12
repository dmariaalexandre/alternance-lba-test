"""
Scraper LBA (La Bonne Alternance) - Mode / Ile-de-France
=========================================================
API    : https://api.apprentissage.beta.gouv.fr/api/job/v1/search
Reponse: { jobs: [...], recruiters: [...], warnings: [...] }
Zone   : Paris + 30 km
Filtre : offres des 14 derniers jours
Output : lba_mode_YYYY-MM-DD.xlsx  +  lba_mode_YYYY-MM-DD.csv

Secret requis (GitHub Secrets) :
  LBA_API_TOKEN -> jeton cree sur https://api.apprentissage.beta.gouv.fr/fr/compte/profil

Mode debug : LBA_DEBUG=true -> affiche le JSON brut du premier appel
"""

import csv
import hashlib
import json
import os
import re
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import requests
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# CONFIG
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

LBA_API_TOKEN = os.getenv("LBA_API_TOKEN", "")
LBA_JOBS_URL  = "https://api.apprentissage.beta.gouv.fr/api/job/v1/search"

# 12 codes ROME mode (B1806 exclu - tapisserie != mode)
ROMES_MODE = [
    "B1801",  # Chapellerie / Modiste
    "B1803",  # Vetements sur mesure / petite serie
    "B1805",  # Stylisme
    "B1808",  # Confection, production en serie
    "B1809",  # Couture flou
    "B1813",  # Maroquinerie et gainerie
    "H1205",  # Modeliste industriel (matieres souples)
    "H2401",  # Assemblage-montage cuirs / peaux
    "H2402",  # Assemblage-montage vetements / textiles
    "H2411",  # Montage prototype cuir / matieres souples
    "H2412",  # Patronnage-graduation
    "D1214",  # Vente habillement et accessoires
]

PARIS_LAT = 48.8534
PARIS_LON = 2.3488
RADIUS_KM = 30
DAYS_BACK = 14

OUTPUT_DIR = Path(__file__).parent
TODAY      = date.today().isoformat()
CUTOFF     = datetime.now(tz=timezone.utc) - timedelta(days=DAYS_BACK)

# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# FORMAT EXCEL
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

# (label, largeur_col, cle_dict)
COLUMNS = [
    ("Statut",             14, "statut"),
    ("Titre du poste",     38, "titre"),
    ("Entreprise",         22, "entreprise"),
    ("Ville",              18, "ville"),
    ("Adresse",            35, "adresse"),
    ("Categorie",          16, "categorie"),
    ("Contrat",            14, "contrat"),
    ("Niveau diplome",     16, "niveau_diplome"),
    ("Date debut",         14, "date_debut"),
    ("Date expiration",    14, "date_expiration"),
    ("Date publication",   16, "date_publication"),
    ("Description",        70, "description"),
    ("Competences req.",   50, "competences_requises"),
    ("Competences acq.",   50, "competences_acquises"),
    ("Conditions acces",   40, "conditions_acces"),
    ("Contact tel",        16, "contact_tel"),
    ("Contact email",      30, "contact_email"),
    ("Site entreprise",    35, "site_entreprise"),
    ("Lien candidature",   45, "lien_candidature"),
    ("Source",             20, "source"),
    ("Notes",              30, "notes"),
    ("ID",                 18, "id"),
    ("Date scraping",      14, "date_scraping"),
    ("Code ROME",          14, "rome_debug"),
]

HEADER_FILL = PatternFill("solid", fgColor="1A1A2E")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)

SOURCE_COLORS = {
    "LBA - Direct":    "E8F5E9",   # vert clair
    "France Travail":  "E3F2FD",   # bleu clair
    "LBA - Spontanee": "FFF9E6",   # jaune clair
}

WRAP_COLS = {12, 13, 14, 15}
LINK_COLS = {18: "site_entreprise", 19: "lien_candidature"}


# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# API
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

def _headers() -> dict:
    h = {"Accept": "application/json"}
    if LBA_API_TOKEN:
        h["Authorization"] = f"Bearer {LBA_API_TOKEN}"
    return h


def fetch_jobs(rome: str, debug: bool = False) -> dict:
    """Appelle GET /api/job/v1/search pour un code ROME."""
    params = {
        "romes":     rome,
        "longitude": PARIS_LON,
        "latitude":  PARIS_LAT,
        "radius":    RADIUS_KM,
    }
    try:
        r = requests.get(LBA_JOBS_URL, headers=_headers(), params=params, timeout=30)
        r.raise_for_status()
        data = r.json()
        if debug:
            print(f"\n  DEBUG {rome} - cles reponse: {list(data.keys())}")
            for k, v in data.items():
                if isinstance(v, list):
                    print(f"    '{k}': {len(v)} items")
                    if v:
                        print(f"    Premier item de '{k}':")
                        print(json.dumps(v[0], indent=6, ensure_ascii=False)[:2000])
                else:
                    print(f"    '{k}': {type(v).__name__}")
            print()
        return data
    except requests.HTTPError as e:
        status = e.response.status_code
        body   = e.response.text[:300]
        print(f"\n  ERREUR HTTP {status}: {body}")
        if status == 401:
            print("  -> Token invalide ou absent.")
            print("  -> Creer un compte sur https://api.apprentissage.beta.gouv.fr")
        return {}
    except Exception as e:
        print(f"\n  Erreur reseau: {e}")
        return {}


# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# PARSING
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

def _parse_date(raw) -> datetime | None:
    if not raw:
        return None
    if isinstance(raw, datetime):
        return raw.replace(tzinfo=timezone.utc) if raw.tzinfo is None else raw
    raw = str(raw)
    for fmt in ("%Y-%m-%dT%H:%M:%S.%fZ", "%Y-%m-%dT%H:%M:%SZ",
                "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(raw[:26], fmt)
            return dt.replace(tzinfo=timezone.utc) if dt.tzinfo is None else dt
        except ValueError:
            continue
    return None


def is_recent(date_raw) -> bool:
    if not date_raw:
        return True
    dt = _parse_date(date_raw)
    return dt >= CUTOFF if dt else True


def _s(val) -> str:
    if val is None:
        return ""
    if isinstance(val, list):
        return " | ".join(str(v) for v in val if v)
    return str(val).strip()


def _extract_city(address_str: str) -> str:
    """Extrait la ville depuis une adresse string LBA."""
    if not address_str:
        return ""
    # Format "75001 Paris" ou "Paris 75001"
    m = re.search(r'\b\d{5}\s+([A-Za-z\xc0-\xff\s\-]+)', address_str)
    if m:
        return m.group(1).strip().title()
    # Derniere partie apres virgule contenant des lettres
    parts = [p.strip() for p in address_str.split(',')]
    for p in reversed(parts):
        if re.search(r'[A-Za-z\xc0-\xff]', p) and not re.match(r'^\d{5}$', p.strip()):
            return p.strip().title()
    return address_str.strip()


def _source_label(partner_label: str) -> str:
    """Convertit partner_label en libelle source lisible."""
    pl = (partner_label or "").lower()
    if pl == "offres_emploi_lba" or not pl:
        return "LBA - Direct"
    if "france travail" in pl or "pole emploi" in pl:
        return "France Travail"
    if pl == "recruteurs_lba":
        return "LBA - Spontanee"
    return f"Partenaire - {partner_label}"


def parse_job(raw: dict) -> dict | None:
    """
    Parse un item de la liste 'jobs' (vraies offres d'emploi).

    Structure API reelle (swagger.json confirme) :
      identifier   : id, partner_job_id, partner_label
      workplace    : name, brand, legal_name, website, location.address (string)
      offer        : title, description, desired_skills[], to_be_acquired_skills[],
                     access_conditions[], rome_codes[], publication.{creation,expiration},
                     target_diploma.{european,label}
      contract     : type, start, duration, remote
      apply        : url, phone
      is_delegated : bool
    """
    identifier  = raw.get("identifier",  {}) or {}
    workplace   = raw.get("workplace",   {}) or {}
    offer       = raw.get("offer",       {}) or {}
    contract    = raw.get("contract",    {}) or {}
    apply_info  = raw.get("apply",       {}) or {}
    location    = workplace.get("location", {}) or {}
    publication = offer.get("publication", {}) or {}

    # Date de publication
    date_creation = publication.get("creation")
    if not is_recent(date_creation):
        return None
    date_pub = _s(date_creation)[:10] if date_creation else ""

    # Titre (obligatoire)
    titre = _s(offer.get("title"))
    if not titre:
        return None

    # Entreprise
    entreprise = (
        _s(workplace.get("brand"))
        or _s(workplace.get("name"))
        or _s(workplace.get("legal_name"))
    )

    # Adresse / Ville
    adresse = _s(location.get("address", ""))
    ville   = _extract_city(adresse)

    # Contrat
    ct = contract.get("type", "Alternance")
    contrat = _s(ct) if ct else "Alternance"

    # Dates
    date_debut = _s(contract.get("start", ""))[:10] if contract.get("start") else ""
    date_exp   = _s(publication.get("expiration", ""))[:10] if publication.get("expiration") else ""

    # Diplome cible
    target_dip = offer.get("target_diploma") or {}
    if isinstance(target_dip, dict):
        niveau_diplome = _s(target_dip.get("label") or target_dip.get("european") or "")
    else:
        niveau_diplome = _s(target_dip)

    # Description + competences
    description          = _s(offer.get("description", ""))[:4000]
    competences_requises = _s(offer.get("desired_skills", []))[:2000]
    competences_acquises = _s(offer.get("to_be_acquired_skills", []))[:2000]
    conditions_acces     = _s(offer.get("access_conditions", []))[:500]

    # Contact
    contact_tel     = _s(apply_info.get("phone", ""))
    lien_cand       = _s(apply_info.get("url", ""))
    site_entreprise = _s(workplace.get("website", ""))

    # Source
    partner_label = _s(identifier.get("partner_label", ""))
    source = _source_label(partner_label)

    # ROME
    rome_debug = _s(offer.get("rome_codes", []))

    # ID unique
    uid_src = (
        _s(identifier.get("id"))
        or _s(identifier.get("partner_job_id"))
        or lien_cand
        or f"{titre}|{entreprise}"
    )
    offer_id = hashlib.sha256(uid_src.encode()).hexdigest()[:16]

    return {
        "statut":               "Nouveau",
        "titre":                titre,
        "entreprise":           entreprise,
        "ville":                ville,
        "adresse":              adresse,
        "categorie":            "",
        "contrat":              contrat,
        "niveau_diplome":       niveau_diplome,
        "date_debut":           date_debut,
        "date_expiration":      date_exp,
        "date_publication":     date_pub,
        "description":          description,
        "competences_requises": competences_requises,
        "competences_acquises": competences_acquises,
        "conditions_acces":     conditions_acces,
        "contact_tel":          contact_tel,
        "contact_email":        "",
        "site_entreprise":      site_entreprise,
        "lien_candidature":     lien_cand,
        "source":               source,
        "notes":                "",
        "id":                   offer_id,
        "date_scraping":        TODAY,
        "rome_debug":           rome_debug,
    }


def parse_recruiter(raw: dict) -> dict | None:
    """Parse un item de la liste 'recruiters' (candidatures spontanees)."""
    identifier = raw.get("identifier", {}) or {}
    workplace  = raw.get("workplace",  {}) or {}
    apply_info = raw.get("apply",      {}) or {}
    location   = workplace.get("location", {}) or {}

    entreprise = (
        _s(workplace.get("brand"))
        or _s(workplace.get("name"))
        or _s(workplace.get("legal_name"))
    )
    if not entreprise:
        return None

    adresse     = _s(location.get("address", ""))
    ville       = _extract_city(adresse)
    apply_url   = _s(apply_info.get("url", ""))
    contact_tel = _s(apply_info.get("phone", ""))
    site_web    = _s(workplace.get("website", ""))

    uid_src  = _s(identifier.get("id")) or f"{entreprise}|{ville}"
    offer_id = hashlib.sha256(uid_src.encode()).hexdigest()[:16]

    return {
        "statut":               "Candidature spontanee",
        "titre":                f"Candidature spontanee - {entreprise}",
        "entreprise":           entreprise,
        "ville":                ville,
        "adresse":              adresse,
        "categorie":            "",
        "contrat":              "Alternance",
        "niveau_diplome":       "",
        "date_debut":           "",
        "date_expiration":      "",
        "date_publication":     TODAY,
        "description":          "Entreprise susceptible d'embaucher en alternance (algorithme LBA)",
        "competences_requises": "",
        "competences_acquises": "",
        "conditions_acces":     "",
        "contact_tel":          contact_tel,
        "contact_email":        "",
        "site_entreprise":      site_web,
        "lien_candidature":     apply_url,
        "source":               "LBA - Spontanee",
        "notes":                "",
        "id":                   offer_id,
        "date_scraping":        TODAY,
        "rome_debug":           "",
    }


# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# SCRAPE
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

def scrape_lba(debug: bool = False) -> list[dict]:
    """
    Appelle l'API une fois par code ROME.
    Reponse reelle : { jobs: [...], recruiters: [...], warnings: [...] }
    """
    all_offers: list[dict] = []
    total_raw = 0
    first_call = True

    print(f"\nLBA Scraper - {len(ROMES_MODE)} codes ROME - Paris +{RADIUS_KM}km - {DAYS_BACK} derniers jours")
    print(f"   Token : {'PRESENT' if LBA_API_TOKEN else 'ABSENT'}")
    print(f"   URL   : {LBA_JOBS_URL}")

    for rome in ROMES_MODE:
        print(f"  -> {rome}...", end=" ", flush=True)

        do_debug = debug and first_call
        data = fetch_jobs(rome, debug=do_debug)
        first_call = False

        if not data:
            print("vide ou erreur")
            continue

        # Vraies cles de reponse confirmees par swagger.json
        jobs       = data.get("jobs",       []) or []
        recruiters = data.get("recruiters", []) or []

        page_raw = len(jobs) + len(recruiters)
        total_raw += page_raw
        kept = filtered = 0

        for raw in jobs:
            o = parse_job(raw)
            if o:
                all_offers.append(o)
                kept += 1
            else:
                filtered += 1

        for raw in recruiters:
            o = parse_recruiter(raw)
            if o:
                all_offers.append(o)
                kept += 1
            else:
                filtered += 1

        print(f"{kept} retenues / {page_raw} ({filtered} filtrees)")

    print(f"\nRecuperation terminee : {len(all_offers)} offres / {total_raw} analysees")
    return all_offers


# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# DEDUPLICATION
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

def deduplicate(offers: list[dict]) -> list[dict]:
    seen: set[str] = set()
    unique: list[dict] = []
    for o in offers:
        key = o["id"]
        if key not in seen:
            seen.add(key)
            unique.append(o)
    removed = len(offers) - len(unique)
    if removed:
        print(f"Deduplication : {removed} doublons supprimes -> {len(unique)} uniques")
    return unique


# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# EXPORT EXCEL
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

def export_excel(offers: list[dict], filepath: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "LBA Mode IDF"

    for col_idx, (label, width, _) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    for row_idx, offer in enumerate(offers, start=2):
        src_color = SOURCE_COLORS.get(offer.get("source", ""), "FFFFFF")
        row_fill  = PatternFill("solid", fgColor=src_color)

        for col_idx, (_, _, key) in enumerate(COLUMNS, start=1):
            val  = offer.get(key, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.fill = row_fill
            cell.alignment = Alignment(
                wrap_text=(col_idx in WRAP_COLS),
                vertical="top",
            )

        for col_idx, key in LINK_COLS.items():
            lnk = offer.get(key, "")
            if lnk and str(lnk).startswith("http"):
                lc = ws.cell(row=row_idx, column=col_idx)
                lc.hyperlink = lnk
                lc.font = Font(color="0563C1", underline="single")

    ws.auto_filter.ref = ws.dimensions
    wb.save(filepath)
    print(f"Excel : {filepath.name}  ({filepath.stat().st_size // 1024} ko)")


# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# EXPORT CSV
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

def export_csv(offers: list[dict], filepath: Path) -> None:
    if not offers:
        return
    fieldnames = [col[0] for col in COLUMNS]
    key_map    = {col[0]: col[2] for col in COLUMNS}

    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for o in offers:
            writer.writerow({col: o.get(key_map[col], "") for col in fieldnames})

    print(f"CSV : {filepath.name}  ({filepath.stat().st_size // 1024} ko)")


# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# MAIN
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

def main() -> None:
    debug = os.getenv("LBA_DEBUG", "").lower() in ("1", "true", "yes")

    print(f"LBA Mode Scraper - {TODAY}")
    print("=" * 60)

    if not LBA_API_TOKEN:
        print("ERREUR : LBA_API_TOKEN absent.")
        print("  1. Creer un compte gratuit sur https://api.apprentissage.beta.gouv.fr/fr/compte/profil")
        print("  2. Generer un jeton d'API")
        print("  3. Dans GitHub : Settings > Secrets > Actions > Mettre a jour LBA_API_TOKEN")
        return

    offers = scrape_lba(debug=debug)
    offers = deduplicate(offers)

    if not offers:
        print("\nAucune offre recuperee.")
        print("  -> Verifier que LBA_API_TOKEN est valide")
        print("  -> Relancer avec LBA_DEBUG=true pour voir la structure de reponse")
        return

    job_count   = sum(1 for o in offers if o.get("source") in ("LBA - Direct", "France Travail"))
    lba_count   = sum(1 for o in offers if o.get("source") == "LBA - Direct")
    ft_count    = sum(1 for o in offers if o.get("source") == "France Travail")
    spont_count = sum(1 for o in offers if o.get("source") == "LBA - Spontanee")
    with_phone  = sum(1 for o in offers if o.get("contact_tel"))
    with_desc   = sum(1 for o in offers if o.get("description") and len(o["description"]) > 20)

    print(f"\nResultats - {TODAY}")
    print(f"   Total             : {len(offers)}")
    print(f"   Vraies offres     : {job_count}")
    print(f"     LBA Direct      : {lba_count}")
    print(f"     France Travail  : {ft_count}")
    print(f"   Spontanees        : {spont_count}")
    print(f"   Avec telephone    : {with_phone}")
    print(f"   Avec description  : {with_desc}")

    xlsx_path = OUTPUT_DIR / f"lba_mode_{TODAY}.xlsx"
    csv_path  = OUTPUT_DIR / f"lba_mode_{TODAY}.csv"

    export_excel(offers, xlsx_path)
    export_csv(offers,   csv_path)

    print(f"\nDone - {len(offers)} offres - Paris +{RADIUS_KM}km - {DAYS_BACK}j - {len(ROMES_MODE)} codes ROME")


if __name__ == "__main__":
    main()
