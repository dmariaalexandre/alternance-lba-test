"""
Scraper LBA (La Bonne Alternance) - Mode / Ile-de-France
=========================================================
API  : https://api.apprentissage.beta.gouv.fr/api/job/v1/search  (nouvelle API 2025+)
Sources : offres_emploi_lba + offres_emploi_partenaires + recruteurs_lba
Zone    : Paris + 30 km
Filtre  : offres des 14 derniers jours
Output  : lba_mode_YYYY-MM-DD.xlsx  +  lba_mode_YYYY-MM-DD.csv

Secret requis (GitHub Secrets) :
  LBA_API_TOKEN -> jeton cree sur https://api.apprentissage.beta.gouv.fr/fr/compte/profil

Mode debug : LBA_DEBUG=true -> affiche le JSON brut du premier appel
"""

import csv
import hashlib
import json
import os
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import requests
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────────────────────────────────────

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
    "H2412",  # Patronnage-gradation
    "D1214",  # Vente habillement et accessoires
]

PARIS_LAT = 48.8534
PARIS_LON = 2.3488
RADIUS_KM = 30
DAYS_BACK = 14

OUTPUT_DIR = Path(__file__).parent
TODAY      = date.today().isoformat()
CUTOFF     = datetime.now(tz=timezone.utc) - timedelta(days=DAYS_BACK)

# ──────────────────────────────────────────────────────────────────────────────
# FORMAT EXCEL
# ──────────────────────────────────────────────────────────────────────────────

COLUMNS = [
    ("Statut",           14, "statut"),
    ("Titre du poste",   38, "titre"),
    ("Entreprise",       22, "entreprise"),
    ("Ville",            20, "ville"),
    ("Region",           18, "region"),
    ("Departement",      18, "departement"),
    ("Categorie",        16, "categorie"),
    ("Contrat",          14, "contrat"),
    ("Type emploi",      14, "type_emploi"),
    ("Experience",       16, "experience"),
    ("Date debut",       14, "date_debut"),
    ("Date publication", 16, "date_publication"),
    ("Description",      70, "description"),
    ("Profil recherche", 60, "profil_recherche"),
    ("Contact nom",      22, "contact_nom"),
    ("Contact tel",      16, "contact_tel"),
    ("Contact email",    30, "contact_email"),
    ("Site internet",    35, "site_internet"),
    ("Lien candidature", 45, "lien_candidature"),
    ("Source",           20, "source"),
    ("Lien offre",       50, "lien"),
    ("Notes",            30, "notes"),
    ("ID",               18, "id"),
    ("Date scraping",    14, "date_scraping"),
    ("Code ROME",        12, "rome_debug"),
]

HEADER_FILL = PatternFill("solid", fgColor="1A1A2E")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)

SOURCE_COLORS = {
    "LBA - Direct":               "E8F5E9",   # vert clair
    "LBA - France Travail":       "E3F2FD",   # bleu clair
    "LBA - Spontanee":            "FFF9E6",   # jaune clair
}

WRAP_COLS = {13, 14}
LINK_COLS = {17: "site_internet", 19: "lien_candidature", 21: "lien"}


# ──────────────────────────────────────────────────────────────────────────────
# API
# ──────────────────────────────────────────────────────────────────────────────

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
                        print(json.dumps(v[0], indent=6, ensure_ascii=False)[:1500])
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
            print("  -> Puis mettre a jour le secret GitHub LBA_API_TOKEN")
        return {}
    except Exception as e:
        print(f"\n  Erreur reseau: {e}")
        return {}


# ──────────────────────────────────────────────────────────────────────────────
# PARSING
# ──────────────────────────────────────────────────────────────────────────────

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
        return ", ".join(str(v) for v in val if v)
    return str(val).strip()


def _parse_address(location: dict) -> tuple[str, str, str]:
    """Extrait (ville, region, departement) depuis location."""
    address = location.get("address") or {}
    if isinstance(address, dict):
        ville       = _s(address.get("city")       or address.get("label") or "")
        region      = _s(address.get("region")      or "Ile-de-France")
        departement = _s(address.get("departement") or address.get("department") or "")
    else:
        ville, region, departement = _s(address), "Ile-de-France", ""
    return ville, region, departement


def parse_job_offer(raw: dict, source_override: str | None = None) -> dict | None:
    """
    Normalise une offre d'emploi depuis la nouvelle API.

    Nouvelle structure attendue :
      raw.identifier  -> id, partner_label, partner_job_id
      raw.contract    -> type (list), start, duration, remote
      raw.offer       -> title, description, rome_codes (list), target_diploma
      raw.workplace   -> name, brand, legal_name, siret, website
      raw.location    -> address (dict: city, region, departement), geopoint
      raw.publication -> creation (ISO date), expiration
      raw.apply       -> url, phone, recipient_id
    """
    identifier  = raw.get("identifier",  {}) or {}
    contract    = raw.get("contract",    {}) or {}
    offer       = raw.get("offer",       {}) or {}
    workplace   = raw.get("workplace",   {}) or {}
    location    = raw.get("location",    {}) or {}
    publication = raw.get("publication", {}) or {}
    apply_info  = raw.get("apply",       {}) or {}

    # Date de publication
    date_raw = publication.get("creation")
    if not is_recent(date_raw):
        return None
    date_pub = _s(date_raw)[:10] if date_raw else ""

    # Titre (obligatoire)
    titre = _s(offer.get("title"))
    if not titre:
        return None

    # Entreprise
    entreprise = (_s(workplace.get("brand"))
               or _s(workplace.get("name"))
               or _s(workplace.get("legal_name")))

    # Localisation
    ville, region, departement = _parse_address(location)

    # Contrat
    contract_types = contract.get("type", [])
    if isinstance(contract_types, list):
        contrat = ", ".join(contract_types) if contract_types else "Alternance"
    else:
        contrat = _s(contract_types) or "Alternance"

    # Date debut
    date_debut = _s(contract.get("start", ""))[:10] if contract.get("start") else ""

    # Description
    description   = _s(offer.get("description", ""))[:3000]
    profil        = _s(offer.get("access_conditions", ""))[:1000]

    # Contact
    contact_tel = _s(apply_info.get("phone", ""))
    apply_url   = _s(apply_info.get("url", ""))
    site_web    = _s(workplace.get("website", ""))

    # Source
    if source_override:
        source = source_override
    else:
        partner_label = _s(identifier.get("partner_label", ""))
        lbl_lower = partner_label.lower()
        if not partner_label or "bonne alternance" in lbl_lower or lbl_lower == "lba":
            source = "LBA - Direct"
        elif "france travail" in lbl_lower or "pole emploi" in lbl_lower:
            source = "LBA - France Travail"
        else:
            source = f"LBA - {partner_label}"

    # ROME codes
    rome_codes = offer.get("rome_codes", [])
    rome_debug = _s(rome_codes)

    # ID unique
    uid_src = (_s(identifier.get("id"))
            or _s(identifier.get("partner_job_id"))
            or apply_url
            or f"{titre}|{entreprise}")
    offer_id = hashlib.sha256(uid_src.encode()).hexdigest()[:16]

    return {
        "statut":           "Nouveau",
        "titre":            titre,
        "entreprise":       entreprise,
        "ville":            ville,
        "region":           region,
        "departement":      departement,
        "categorie":        "",
        "contrat":          contrat,
        "type_emploi":      "",
        "experience":       "",
        "date_debut":       date_debut,
        "date_publication": date_pub,
        "description":      description,
        "profil_recherche": profil,
        "contact_nom":      "",
        "contact_tel":      contact_tel,
        "contact_email":    "",
        "site_internet":    site_web,
        "lien_candidature": apply_url,
        "source":           source,
        "lien":             apply_url,
        "notes":            "",
        "id":               offer_id,
        "date_scraping":    TODAY,
        "rome_debug":       rome_debug,
    }


def parse_recruiter(raw: dict) -> dict | None:
    """Normalise un recruteur potentiel (candidature spontanee)."""
    identifier = raw.get("identifier", {}) or {}
    workplace  = raw.get("workplace",  {}) or {}
    location   = raw.get("location",   {}) or {}
    apply_info = raw.get("apply",      {}) or {}

    entreprise = (_s(workplace.get("brand"))
               or _s(workplace.get("name"))
               or _s(workplace.get("legal_name")))
    if not entreprise:
        return None

    ville, region, departement = _parse_address(location)
    apply_url   = _s(apply_info.get("url", ""))
    contact_tel = _s(apply_info.get("phone", ""))
    site_web    = _s(workplace.get("website", ""))

    uid_src  = _s(identifier.get("id")) or f"{entreprise}|{ville}"
    offer_id = hashlib.sha256(uid_src.encode()).hexdigest()[:16]

    return {
        "statut":           "Candidature spontanee",
        "titre":            f"Candidature spontanee - {entreprise}",
        "entreprise":       entreprise,
        "ville":            ville,
        "region":           region,
        "departement":      departement,
        "categorie":        "",
        "contrat":          "Alternance",
        "type_emploi":      "",
        "experience":       "",
        "date_debut":       "",
        "date_publication": TODAY,
        "description":      "Entreprise susceptible d'embaucher en alternance (algorithme LBA)",
        "profil_recherche": "",
        "contact_nom":      "",
        "contact_tel":      contact_tel,
        "contact_email":    "",
        "site_internet":    site_web,
        "lien_candidature": apply_url,
        "source":           "LBA - Spontanee",
        "lien":             apply_url,
        "notes":            "",
        "id":               offer_id,
        "date_scraping":    TODAY,
        "rome_debug":       "",
    }


# ──────────────────────────────────────────────────────────────────────────────
# SCRAPE
# ──────────────────────────────────────────────────────────────────────────────

def scrape_lba(debug: bool = False) -> list[dict]:
    """
    Appelle l'API une fois par code ROME pour maximiser les resultats.
    Nouvelle API : pas de pagination, max ~150 offres par source par appel.
    """
    all_offers: list[dict] = []
    total_raw = 0
    first_call = True

    print(f"\nLBA Scraper - {len(ROMES_MODE)} codes ROME - Paris +{RADIUS_KM}km - {DAYS_BACK} derniers jours")
    print(f"   Token : {'PRESENT' if LBA_API_TOKEN else 'ABSENT'}")
    print(f"   URL   : {LBA_JOBS_URL}")

    for rome in ROMES_MODE:
        print(f"  -> {rome}...", end=" ", flush=True)

        # Debug seulement sur le premier appel pour ne pas inonder les logs
        do_debug = debug and first_call
        data = fetch_jobs(rome, debug=do_debug)
        first_call = False

        if not data:
            print("vide ou erreur")
            continue

        # Nouvelle API : offres_emploi_lba, offres_emploi_partenaires, recruteurs_lba
        # Fallback vers les cles plus generiques si la structure change
        lba_direct   = (data.get("offres_emploi_lba",         None) or
                        data.get("lbaJobs",                   None) or [])
        ft_offers    = (data.get("offres_emploi_partenaires", None) or
                        data.get("peJobs",                    None) or [])
        recruiters   = (data.get("recruteurs_lba",            None) or
                        data.get("recruiters",                None) or [])

        # Si la cle est un dict avec sous-cle "results" (ancien format)
        if isinstance(lba_direct, dict):
            lba_direct = lba_direct.get("results", [])
        if isinstance(ft_offers, dict):
            ft_offers = ft_offers.get("results", [])

        page_raw = len(lba_direct) + len(ft_offers) + len(recruiters)
        total_raw += page_raw
        kept = filtered = 0

        for raw in (lba_direct or []):
            o = parse_job_offer(raw, source_override="LBA - Direct")
            if o:
                all_offers.append(o)
                kept += 1
            else:
                filtered += 1

        for raw in (ft_offers or []):
            o = parse_job_offer(raw, source_override="LBA - France Travail")
            if o:
                all_offers.append(o)
                kept += 1
            else:
                filtered += 1

        for raw in (recruiters or []):
            o = parse_recruiter(raw)
            if o:
                all_offers.append(o)
                kept += 1
            else:
                filtered += 1

        print(f"{kept} retenues / {page_raw} ({filtered} filtrees)")

    print(f"\nRecuperation terminee : {len(all_offers)} offres / {total_raw} analysees")
    return all_offers


# ──────────────────────────────────────────────────────────────────────────────
# DEDUPLICATION
# ──────────────────────────────────────────────────────────────────────────────

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


# ──────────────────────────────────────────────────────────────────────────────
# EXPORT EXCEL
# ──────────────────────────────────────────────────────────────────────────────

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


# ──────────────────────────────────────────────────────────────────────────────
# EXPORT CSV
# ──────────────────────────────────────────────────────────────────────────────

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


# ──────────────────────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────────────────────

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
        print("  -> Verifier que LBB_API_TOKEN est valide (nouveau jeton api.apprentissage.beta.gouv.fr)")
        print("  -> Relancer avec LBA_DEBUG=true pour voir la structure de reponse")
        return

    ft_count    = sum(1 for o in offers if "France Travail" in o.get("source", ""))
    lba_count   = sum(1 for o in offers if "Direct" in o.get("source", ""))
    spont_count = sum(1 for o in offers if "Spontanee" in o.get("source", ""))
    with_phone  = sum(1 for o in offers if o.get("contact_tel"))

    print(f"\nResultats - {TODAY}")
    print(f"   Total          : {len(offers)}")
    print(f"   LBA Direct     : {lba_count}")
    print(f"   France Travail : {ft_count}")
    print(f"   Spontanees     : {spont_count}")
    print(f"   Avec telephone : {with_phone}")

    xlsx_path = OUTPUT_DIR / f"lba_mode_{TODAY}.xlsx"
    csv_path  = OUTPUT_DIR / f"lba_mode_{TODAY}.csv"

    export_excel(offers, xlsx_path)
    export_csv(offers,   csv_path)

    print(f"\nDone - {len(offers)} offres - Paris +{RADIUS_KM}km - {DAYS_BACK}j - {len(ROMES_MODE)} codes ROME")


if __name__ == "__main__":
    main()
