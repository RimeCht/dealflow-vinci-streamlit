import json
import os
import random
import re
import sys
import threading
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openai import AzureOpenAI
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from tqdm import tqdm

import semantic_prefilter
import learning_memory
import dealflow_profiles
import excel_styling


# ============================================================
# CONFIGURATION
# ============================================================

load_dotenv()

def env_int(name: str, default: int) -> int:
    try:
        return int(os.getenv(name, default))
    except (TypeError, ValueError):
        return default


def env_float(name: str, default: float) -> float:
    try:
        return float(os.getenv(name, default))
    except (TypeError, ValueError):
        return default


def env_bool(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() not in {"0", "false", "non", "no", "off"}


MODEL_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-5-mini-2")
AZURE_OPENAI_API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION", "2024-12-01-preview")
ALGO1_API_MODE = os.getenv("ALGO1_API_MODE", "chat").strip().lower()

OUTPUT_SUFFIX = "_algo_1_prequalification.xlsx"
TEMP_SUFFIX = "_algo_1_prequalification_sauvegarde_temp.xlsx"
SAVE_EVERY = env_int("ALGO1_SAVE_EVERY", 20)
REQUEST_DELAY_SECONDS = env_float("ALGO1_REQUEST_DELAY_SECONDS", 0.1)
MAX_API_RETRIES = env_int("ALGO1_MAX_API_RETRIES", 3)
MAX_WORKERS = max(1, env_int("ALGO1_MAX_WORKERS", 3))
SCRAPE_TIMEOUT_SECONDS = env_int("ALGO1_SCRAPE_TIMEOUT_SECONDS", 8)
MAX_WEBSITE_TEXT_CHARS = env_int("ALGO1_MAX_WEBSITE_TEXT_CHARS", 4000)
LIMIT_ROWS = env_int("ALGO1_LIMIT_ROWS", 0)
RESUME_FROM_TEMP = env_bool("ALGO1_RESUME_FROM_TEMP", True)
RETRY_FAILED_TEMP_ROWS = env_bool("ALGO1_RETRY_FAILED_TEMP_ROWS", True)
ENABLE_WEBSITE_RESEARCH = env_bool("ALGO1_ENABLE_WEBSITE_RESEARCH", True)
RESEARCH_MAX_PAGES = env_int("ALGO1_RESEARCH_MAX_PAGES", 5)
RESEARCH_MAX_TEXT_CHARS = env_int("ALGO1_RESEARCH_MAX_TEXT_CHARS", 6000)
MAX_EXCEL_CONTEXT_CHARS = env_int("ALGO1_MAX_EXCEL_CONTEXT_CHARS", 14000)
MAX_EXCEL_CELL_CHARS = env_int("ALGO1_MAX_EXCEL_CELL_CHARS", 3200)
INCLUDE_EXCEL_COLUMN_LIST_IN_PROMPT = env_bool("ALGO1_INCLUDE_EXCEL_COLUMN_LIST_IN_PROMPT", False)
ENABLE_EXCEL_SHORTCUT = env_bool("ALGO1_ENABLE_EXCEL_SHORTCUT", True)
DISABLE_API = env_bool("ALGO1_DISABLE_API", False)
NON_RETRYABLE_API_STATUS_CODES = {400, 401, 403, 404}

EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}

TARGET_SECTORS = [
    "Construction",
    "Built World",
    "Infrastructure",
    "Real Estate",
    "Energy",
    "Mobility",
]

COMPETITORS = [
    "Bouygues",
    "Eiffage",
    "Spie",
    "Colas",
    "Saint-Gobain",
    "Suez",
    "Veolia",
    "Ferrovial",
    "ACS",
    "Skanska",
    "Hochtief",
    "Strabag",
    "Balfour Beatty",
    "edf",
    "total energies",
    "engie",
    "natran",
    "grt gaz",
    "Kajima",
    "suffolk construction",
    "Bouygues Construction",
    "Bouygues Energies & Services",
    "Bouygues Immobilier",
    "Eiffage Construction",
    "Eiffage Energie Systèmes",
    "Eiffage Énergie Systèmes",
    "Eiffage Génie Civil",
    "Eiffage Route",
    "Fayat",
    "Fayat Group",
    "NGE",
    "NGE Groupe",
    "Demathieu Bard",
    "Léon Grosse",
    "Leon Grosse",
    "GCC",
    "Spie Batignolles",
    "Razel-Bec",
    "Groupe Legendre",
    "Legendre Construction",
    "Rabot Dutilleul",
    "Baudin Chateauneuf",
    "Charier",
    "Groupe Angevin",
    "Groupe Cassous",
    "Colas Rail",
    "Colas France",
    "Keolis",
    "Transdev",
    "APRR",
    "AREA",
    "Sanef",
    "Abertis",
    "Atlantia",
    "Mundys",
    "Meridiam",
    "Macquarie Infrastructure",
    "Globalvia",
    "Acciona",
    "Sacyr",
    "OHL",
    "OHLA",
    "FCC",
    "Webuild",
    "Salini Impregilo",
    "Astaldi",
    "Porr",
    "Implenia",
    "BAM",
    "Royal BAM Group",
    "Besix",
    "Jan De Nul",
    "DEME",
    "Egis",
    "Bechtel",
    "Fluor",
    "Kiewit",
    "Turner Construction",
    "Whiting-Turner",
    "DPR Construction",
    "Mortenson",
    "Gilbane",
    "Clark Construction",
    "Lendlease",
    "Multiplex",
    "Laing O'Rourke",
    "Laing O Rourke",
    "Mace",
    "Kier",
    "Morgan Sindall",
    "Costain",
    "VolkerWessels",
    "Royal VolkerWessels",
    "China State Construction",
    "CSCEC",
    "China Communications Construction",
    "CCCC",
    "China Railway Construction",
    "CRCC",
    "China Railway Group",
    "CREC",
    "PowerChina",
    "Obayashi",
    "Shimizu",
    "Taisei",
    "Takenaka",
    "Penta-Ocean Construction",
    "Samsung C&T",
    "Hyundai Engineering",
    "Hyundai E&C",
    "GS Engineering & Construction",
    "Larsen & Toubro",
    "L&T Construction",
    "SPIE",
    "SPIE Group",
    "Equans",
    "Schneider Electric",
    "Siemens",
    "Siemens Energy",
    "ABB",
    "Hitachi Energy",
    "Alstom",
    "GE Vernova",
    "GE Grid Solutions",
    "RTE",
    "Enedis",
    "GRDF",
    "Terega",
    "Teréga",
    "Elia",
    "E.ON",
    "Iberdrola",
    "Enel",
    "Naturgy",
    "Repsol",
    "Shell",
    "BP",
    "TotalEnergies",
    "Total Energies",
    "SUEZ",
    "Veolia Environnement",
    "Paprec",
    "Séché Environnement",
    "Seche Environnement",
    "Derichebourg",
    "Urbaser",
    "FCC Environment",
    "Remondis",
    "Holcim",
    "Lafarge",
    "LafargeHolcim",
    "Heidelberg Materials",
    "Cemex",
    "CRH",
    "Vicat",
    "Etex",
    "Knauf",
    "Kingspan",
    "Rockwool",
    "ArcelorMittal",
    "Artelia",
    "Setec",
    "Systra",
    "WSP",
    "Arcadis",
    "AtkinsRéalis",
    "AtkinsRealis",
    "Ramboll",
    "Arup",
    "AECOM",
    "Jacobs",
    "Mott MacDonald",
    "Sweco",
    "Dar Group",
    "TYLin",
    "T.Y. Lin",
    "Antea Group",

]

VINCI_NAMES = [
    "VINCI",
    "VINCI Group",
    "Groupe VINCI",
    "Leonard",
    "Leonard VINCI",
    "Léonard VINCI",
    "VINCI Construction",
    "VINCI Energies",
    "VINCI Autoroutes",
    "VINCI Airports",
    "VINCI Highways",
    "Cobra IS",
    "Eurovia",
    "Cobra IS",
    "Eurovia",
    "VINCI Construction Grands Projets",
    "VINCI Construction France",
    "VINCI Construction GeoInfrastructure",
    "VINCI Construction Terrassement",
    "VINCI Construction Maritime et Fluvial",
    "VINCI Construction Dom-Tom",
    "VINCI Construction Outre-mer",
    "VINCI Building",
    "VINCI Facilities",
    "Soletanche Freyssinet",
    "Soletanche Bachy",
    "Freyssinet",
    "Menard",
    "Ménard",
    "Terre Armée",
    "Terre Armee",
    "Nuvia",
    "Sixense",
    "Geoquest",
    "Dodin Campenon Bernard",
    "Campenon Bernard",
    "Sogea",
    "Sogea-Satom",
    "Sogea Satom",
    "Chantiers Modernes Construction",
    "Chantiers Modernes",
    "Botte Fondations",
    "Spiecapag",
    "ETF",
    "Signature",
    "Cardem",
    "Equo Vivo",
    "Arbonis",
    "Seymour Whyte",
    "Hubbard Construction",
    "Blythe Construction",
    "Eurovia UK",
    "Ringway",
    "Actemium",
    "Axians",
    "Omexom",
    "Citeos",
    "Cegelec",
    "VINCI Facilities",
    "VINCI Energies Building Solutions",
    "VINCI Energies France",
    "VINCI Energies International",
    "Cobra",
    "Grupo Cobra",
    "Cobra Instalaciones",
    "Cobra Industrial Services",
    "Dragados Offshore",
    "Semi",
    "Moncobra",
    "Maetel",
    "ASF",
    "Autoroutes du Sud de la France",
    "Escota",
    "Cofiroute",
    "Arcour",
    "Arcos",
    "A19",
    "Duplex A86",
    "VINCI Airports",
    "ANA Aeroportos de Portugal",
    "ANA Airports",
    "Aéroports de Lyon",
    "Aeroports de Lyon",
    "Lyon Aéroport",
    "Lyon Aeroport",
    "London Gatwick",
    "Gatwick Airport",
    "Belfast International Airport",
    "Edinburgh Airport",
    "Santiago Airport",
    "Cambodia Airports",
    "Kansai Airports",
    "VINCI Highways",
    "ViaPlus",
    "Lima Expresa",
    "Gefyra",
    "Gefyra Litourgia",
    "Regina Bypass",
    "Texas SH 130",
    "TollPlus",
    "VINCI Immobilier",
    "Fondation VINCI",
    "Fondation VINCI pour la Cité",
    "VINCI Innovation",
    "VINCI Startup Tour",

    

]

B2B_KEYWORDS = [
    "b2b",
    "enterprise",
    "enterprises",
    "businesses",
    "companies",
    "corporate",
    "industrial",
    "industrials",
    "operators",
    "infrastructure operators",
    "public sector",
    "public authorities",
    "local authorities",
    "collectivités",
    "collectivites",
    "entreprises",
    "industriels",
    "opérateurs",
    "operateurs",
    "promoteurs",
    "constructeurs",
    "gestionnaires d'actifs",
    "acteurs publics",
    "professionnels",
    "asset managers",
    "developers",
    "contractors",
    "utilities",
    "integrateur", 
    "integrator",
]

B2C_KEYWORDS = [
    "b2c",
    "consumer",
    "consumers",
    "individuals",
    "households",
    "families",
    "homeowners",
    "particuliers",
    "ménages",
    "menages",
    "familles",
    "grand public",
    "citoyens",
    "parents",
    "pet owners",
    "tourists",
    "students",
    "b2c marketplace",
    "consumer app",
    "mobile app",
    "personal app",
    "users",
    "end users",
    "final users",
    "general public",
    "retail customers",
    "private customers",
    "private individuals",
    "residents",
    "inhabitants",
    "occupants",
    "tenants",
    "renters",
    "locataires",
    "résidents",
    "residents",
    "habitants",
    "occupants",
    "propriétaires",
    "proprietaires",
    "clients particuliers",
    "clients individuels",
    "utilisateurs finaux",
    "usagers",
    "riverains",
]

ARTISAN_KEYWORDS = [
    "artisan",
    "artisans",
    "craftsmen",
    "craftsperson",
    "freelancers",
    "independent workers",
    "auto entrepreneurs",
    "auto-entrepreneurs",
    "contractor",
    "contractors",
    "subcontractor",
    "subcontractors",
    "sous-traitant",
    "sous-traitants",
    "indépendant",
    "indépendants",
    "independant",
    "independants",
    "self-employed",
    "small business",
    "small businesses",
    "sme",
    "local businesses",
    "local professionals",

]

SECTOR_KEYWORDS = {
    "Construction": [
        "chantier",
        "travaux",
        "bim",
        "construction",
        "matériaux",
        "materiaux",
        "rénovation",
        "renovation",
        "gestion de chantier",
        "construction site",
        "building materials",
        "robotique",
        "robotics",
        "robot",
        "robots",
        "robot de chantier",
        "construction robot",
        "autonomous robot",
        "automation",
        "automatisation",
        "site automation",
        "robotic construction",
        "layout robot",
        "bricklaying robot",
        "painting robot",
        "demolition robot",
        "3d printing construction",
        "impression 3d béton",
        "impression 3d beton",
        "béton",
        "beton",
        "ciment",
        "cement",
        "gros œuvre",
        "gros oeuvre",
        "second œuvre",
        "second oeuvre",
        "maçonnerie",
        "maconnerie",
        "préfabriqué",
        "prefab",
        "modular construction",
        "offsite construction",
        "construction management",
        "site management",
        "project delivery",
        "quantity surveying",
        "cost estimation",
        "devis",
        "planification chantier",
        "sécurité chantier",
        "site safety",
        "quality control",
        "construction equipment",
        "engins de chantier",
        "robotique chantier",
        "construction robotics",
        "drone de chantier",
        "suivi de chantier",

    ],
    "Built World": [
        "built environment",
        "smart building",
        "facility",
        "asset management",
        "maintenance",
        "digital twin",
        "jumeau numérique",
        "jumeau numerique",
        "urban tech",
        "inspection robot",
        "robot d'inspection",
        "robot de maintenance",
        "maintenance robot",
        "autonomous inspection",
        "inspection autonome",
        "drone inspection",
        "drone",
        "uav",
        "computer vision",
        "vision par ordinateur",
         "built world",
        "smart city",
        "ville intelligente",
        "building operations",
        "building management",
        "facility services",
        "asset performance",
        "space management",
        "occupancy",
        "indoor air quality",
        "qualité air intérieur",
        "qualite air interieur",
        "building monitoring",
        "building automation",
        "predictive maintenance",
        "maintenance prédictive",
        "maintenance predictive",
        "gestion technique bâtiment",
        "gestion technique batiment",
        "gtb",
        "bms",
        "iot building",
        "connected building",
    ],
    "Infrastructure": [
        "infrastructure",
        "routes",
        "ponts",
        "bridges",
        "tunnels",
        "rail",
        "ouvrages",
        "réseaux",
        "reseaux",
        "génie civil",
        "genie civil",
        "utilities",
        "rail robot",
        "robot ferroviaire",
        "inspection infrastructure",
        "bridge inspection",
        "tunnel inspection",
        "robot inspection",
        "drones d'inspection",
        "drone inspection",
        "autonomous inspection",
        "inspection autonome",
        "robotique mobile",
        "mobile robotics",
        "civil engineering",
        "road",
        "highway",
        "autoroute",
        "railway",
        "ferroviaire",
        "metro",
        "tramway",
        "airport",
        "aéroport",
        "aeroport",
        "port",
        "water network",
        "réseau d'eau",
        "reseau d'eau",
        "wastewater",
        "eaux usées",
        "eaux usees",
        "pipeline",
        "canalisation",
        "grid infrastructure",
        "public works",
        "travaux publics",
        "tp",
        "earthworks",
        "terrassement",
        "geotechnical",
        "géotechnique",
        "geotechnique",
        "structural health monitoring",
        "inspection ouvrage",
        "bridge inspection",
        "tunnel inspection",
        "rail inspection",
        "drone inspection",
        "robot inspection",
    ],
    "Real Estate": [
        "real estate",
        "immobilier",
        "proptech",
        "property",
        "gestion d'actifs",
        "property management",
        "facility management",
        "exploitation bâtiment",
        "exploitation batiment",
        "asset valuation",
        "valuation",
        "transaction immobilière",
        "transaction immobiliere",
        "leasing",
        "location",
        "commercial real estate",
        "residential real estate",
        "housing",
        "logement",
        "workspace",
        "bureau",
        "offices",
        "occupancy management",
        "tenant experience",
        "gestion locative",
        "promotion immobilière",
        "promotion immobiliere",
        "real estate investment",
        "portfolio management",
        "property data",
        "building performance",
        "real estate analytics",
    ],
    "Energy": [
        "energy",
        "énergie",
        "energie",
        "renewable",
        "solar",
        "battery",
        "grid",
        "smart grid",
        "energy management",
        "decarbonation",
        "décarbonation",
        "efficacité énergétique",
        "efficacite energetique",
        "electricity",
        "électricité",
        "electricite",
        "photovoltaic",
        "pv",
        "wind",
        "éolien",
        "eolien",
        "hydrogen",
        "hydrogène",
        "hydrogene",
        "storage",
        "stockage énergie",
        "stockage energie",
        "thermal",
        "heat pump",
        "pompe à chaleur",
        "pompe a chaleur",
        "charging station",
        "borne de recharge",
        "microgrid",
        "carbon",
        "co2",
        "low carbon",
        "net zero",
        "energy efficiency",
        "energy consumption",
        "building energy",
        "energy monitoring",

    ],
    "Mobility": [
        "mobility",
        "mobilité",
        "mobilite",
        "transport",
        "fleet",
        "flotte",
        "ev charging",
        "recharge",
        "logistics",
        "logistique",
        "traffic",
        "smart mobility",
        "public transport",
        "transport public",
        "urban mobility",
        "mobilité urbaine",
        "mobilite urbaine",
        "shared mobility",
        "car sharing",
        "bike sharing",
        "micro mobility",
        "micromobility",
        "last mile",
        "dernier kilomètre",
        "dernier kilometre",
        "route optimization",
        "optimisation itinéraire",
        "optimisation itineraire",
        "vehicle",
        "véhicule",
        "vehicule",
        "electric vehicle",
        "véhicule électrique",
        "vehicule electrique",
        "parking",
        "road safety",
        "sécurité routière",
        "securite routiere",
        "traffic management",
        "mobility data",
        "transport infrastructure",
    ],
}

THREAD_LOCAL = threading.local()

COLUMN_ALIASES = {
    "startup": [
        "startup",
        "startups",
        "nom",
        "nom startup",
        "nom de la startup",
        "name",
        "company",
        "company name",
        "entreprise",
        "societe",
        "société",
        "organisation",
        "organization",
    ],
    "sector": [
        "secteur",
        "sector",
        "domaine",
        "industry",
        "categorie",
        "catégorie",
        "category",
        "vertical",
        "market",
        "marché",
        "segment",
    ],
    "description": [
        "description",
        "desc",
        "activité",
        "activite",
        "activity",
        "about",
        "resume",
        "résumé",
        "summary",
        "presentation",
        "pitch",
        "what they do",
        "business description",
        "short description",
    ],
    "site": [
        "site web",
        "website",
        "site",
        "url",
        "web",
        "official website",
        "website url",
        "url site",
        "site internet",
        "homepage",
        "home page",
        "lien",
        "link",
        "domain",
        "domaine web",
        "site alternatif",
        "alternative website",
        "alternative url",
    ],
}

EXCEL_CONTEXT_PRIORITY_TERMS = [
    "oneliner",
    "description",
    "technology description",
    "business model",
    "vertical",
    "subtopic",
    "target customer",
    "target project type",
    "fit in construction value chain",
    "specific use case",
    "success stories",
    "pain points",
    "primary target market",
    "startup development level",
    "purpose vision",
    "roadmap",
    "trl",
    "employee",
    "revenue",
    "arr",
    "funding",
    "patents",
    "worked with any player",
    "programs before",
    "awards",
    "additional information",
    "product demo",
    "country",
    "region",
    "inactive",
]

EXCEL_CONTEXT_EXCLUDED_EXACT = {
    "logo",
    "alternative logo",
    "entity id",
    "affinity id",
    "hubspot id",
    "pipedrive id",
    "pitchbook id",
    "creditsafe id",
    "coc number",
    "applicant email",
    "applicant name",
    "email",
    "phone",
    "registered contact title",
    "registered contact name",
    "accepted privacy policy and competition privacy note",
    "accepted t c",
    "ubo birth date",
    "owner birth date",
    "director birth date",
    "ubo name",
    "owner name",
    "director name",
    "ubo date",
    "owner date",
    "director date",
    "holding name",
    "holding date",
    "feeds",
}

EXCEL_CONTEXT_EXCLUDED_PREFIXES = (
    "algo1 ",
    "algo2 ",
    "algo3 ",
)

EXCEL_CONTEXT_EXCLUDED_OUTPUT_COLUMNS = {
    "decision",
    "decision finale",
    "confidence",
    "is b2b",
    "target sector",
    "sector relevance",
    "sector reason",
    "vinci reference detected",
    "competitor reference detected",
    "official website likely",
    "short description clean",
    "reason",
    "missing information",
    "regle finale",
}

RESULT_COLUMNS = [
    "dealflow_profile",
    "profile_exclusion_matches",
    "decision",
    "confidence",
    "is_b2b",
    "target_sector",
    "sector_relevance",
    "sector_reason",
    "vinci_reference_detected",
    "competitor_reference_detected",
    "official_website_likely",
    "short_description_clean",
    "reason",
    "missing_information",
    "algo1_scraped_sources",
    "algo1_website_research_quality",
    "algo1_b2b_signals_detected",
    "algo1_b2c_signals_detected",
    "algo1_artisan_only_risk_detected",
    "algo1_sector_signals_detected",
    "algo1_vinci_signals_detected",
    "algo1_competitor_signals_detected",
    "algo1_research_text_used",
    "algo1_excel_columns_used",
    "algo1_excel_columns_used_count",
    "algo1_excel_columns_available_count",
    "algo1_excel_context_truncated",
    "algo1_processing_mode",
    "algo1_api_called",
    "algo1_research_called",
    "algo1_semantic_backend",
    "algo1_semantic_summary",
    "algo1_semantic_sector",
    "algo1_semantic_sector_score",
    "algo1_semantic_b2b_score",
    "algo1_semantic_b2c_score",
    "algo1_learning_match_label",
    "algo1_learning_match_score",
    "algo1_learning_match_source",
    "decision_finale",
    "regle_finale",
]

FAILED_TEMP_REASON_PATTERNS = [
    "Erreur pendant la classification",
    "Erreur ligne pendant le traitement",
    "Erreur API inconnue",
    "APIConnectionError",
    "APITimeoutError",
    "RateLimitError",
    "BadRequestError",
    "InternalServerError",
]


# ============================================================
# OUTILS
# ============================================================

def log(message: str) -> None:
    print(f"[Algo 1] {message}")


def strip_accents(text: str) -> str:
    return "".join(
        char
        for char in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(char)
    )


def normalize_column_name(col: str) -> str:
    text = strip_accents(str(col)).strip().lower()
    text = re.sub(r"[_\-/]+", " ", text)
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_text(text: str) -> str:
    text = strip_accents(str(text)).strip().lower()
    text = re.sub(r"[_\-/]+", " ", text)
    text = re.sub(r"[^a-z0-9 +#.&'€$]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def column_score(normalized_col: str, normalized_alias: str) -> int:
    if normalized_col == normalized_alias:
        return 100

    col_tokens = set(normalized_col.split())
    alias_tokens = set(normalized_alias.split())

    if not col_tokens or not alias_tokens:
        return 0

    if alias_tokens.issubset(col_tokens):
        return 80 + len(alias_tokens)

    if normalized_alias in normalized_col or normalized_col in normalized_alias:
        return 60

    overlap = len(col_tokens & alias_tokens)
    if overlap:
        return 20 + overlap * 10

    return 0


def find_column(df: pd.DataFrame, possible_names: list[str]) -> str | None:
    normalized_columns = {
        original: normalize_column_name(original)
        for original in df.columns
    }
    normalized_aliases = [normalize_column_name(name) for name in possible_names]

    exact_candidates = []
    for alias_index, normalized_alias in enumerate(normalized_aliases):
        for original, normalized_col in normalized_columns.items():
            if normalized_col != normalized_alias:
                continue
            non_empty_count = int(
                df[original]
                .fillna("")
                .astype(str)
                .str.strip()
                .ne("")
                .sum()
            )
            exact_candidates.append((non_empty_count, -alias_index, original))

    if exact_candidates:
        return max(exact_candidates)[2]

    best_column = None
    best_score = 0

    for original, normalized_col in normalized_columns.items():
        for normalized_alias in normalized_aliases:
            score = column_score(normalized_col, normalized_alias)
            if score > best_score:
                best_column = original
                best_score = score

    return best_column if best_score >= 60 else None


def clean_text(text) -> str:
    if pd.isna(text):
        return ""

    text = str(text)
    text = ILLEGAL_CHARACTERS_RE.sub("", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def should_exclude_excel_context_column(column: str) -> bool:
    normalized = normalize_column_name(column)
    if not normalized:
        return True
    if normalized in EXCEL_CONTEXT_EXCLUDED_EXACT:
        return True
    if normalized in EXCEL_CONTEXT_EXCLUDED_OUTPUT_COLUMNS:
        return True
    if normalized.startswith(EXCEL_CONTEXT_EXCLUDED_PREFIXES):
        return True
    return False


def excel_context_priority(column: str) -> int:
    normalized = normalize_column_name(column)
    for index, term in enumerate(EXCEL_CONTEXT_PRIORITY_TERMS):
        if normalize_column_name(term) in normalized:
            return len(EXCEL_CONTEXT_PRIORITY_TERMS) - index
    return 0


def build_excel_row_context(row: pd.Series) -> dict:
    entries = []
    excluded_count = 0

    for original_index, (column, raw_value) in enumerate(row.items()):
        if should_exclude_excel_context_column(str(column)):
            excluded_count += 1
            continue

        value = clean_text(raw_value)
        if not value or normalize_text(value) in {"nan", "none", "null", "n a", "na"}:
            continue

        cell_truncated = len(value) > MAX_EXCEL_CELL_CHARS
        if cell_truncated:
            value = value[:MAX_EXCEL_CELL_CHARS].rstrip() + " [cellule tronquée]"

        entries.append({
            "column": str(column),
            "value": value,
            "priority": excel_context_priority(str(column)),
            "original_index": original_index,
            "cell_truncated": cell_truncated,
        })

    entries.sort(key=lambda item: (-item["priority"], item["original_index"]))

    lines = []
    used_columns = []
    total_chars = 0
    context_truncated = False

    for entry in entries:
        line = f"- {entry['column']}: {entry['value']}"
        separator_chars = 1 if lines else 0
        remaining = MAX_EXCEL_CONTEXT_CHARS - total_chars - separator_chars

        if remaining <= 0:
            context_truncated = True
            break

        if len(line) > remaining:
            if remaining >= 80:
                line = line[:remaining].rstrip() + " [contexte tronqué]"
                lines.append(line)
                used_columns.append(entry["column"])
            context_truncated = True
            break

        lines.append(line)
        used_columns.append(entry["column"])
        total_chars += len(line) + separator_chars
        context_truncated = context_truncated or entry["cell_truncated"]

    if len(used_columns) < len(entries):
        context_truncated = True

    return {
        "text": "\n".join(lines),
        "used_columns": used_columns,
        "used_count": len(used_columns),
        "available_count": len(entries),
        "excluded_count": excluded_count,
        "truncated": context_truncated,
    }


def sanitize_excel_value(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)

    if isinstance(value, (list, tuple, set)):
        return ", ".join(sanitize_excel_value(item) for item in value)

    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False)

    if pd.isna(value):
        return value

    return value


def sanitize_excel_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    return df.copy().map(sanitize_excel_value)


def is_valid_url(url: str) -> bool:
    if not isinstance(url, str) or not url.strip():
        return False

    url = url.strip()

    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    parsed = urlparse(url)
    return bool(parsed.netloc and "." in parsed.netloc)


def normalize_url(url: str) -> str:
    if not isinstance(url, str) or not url.strip():
        return ""

    url = url.strip()

    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    return url


def is_generated_excel(path: Path) -> bool:
    name = path.name.lower()
    return (
        name.startswith("~$")
        or "_algo_" in name
        or name.endswith(OUTPUT_SUFFIX.lower())
        or name.endswith(TEMP_SUFFIX.lower())
    )


def discover_excel_files(base_dir: Path, cli_args: list[str]) -> list[Path]:
    if cli_args:
        return [Path(arg).expanduser().resolve() for arg in cli_args]

    files = []
    for path in sorted(base_dir.iterdir()):
        if path.suffix.lower() in EXCEL_EXTENSIONS and not is_generated_excel(path):
            files.append(path)

    return files


def make_output_path(input_file: Path) -> Path:
    return input_file.with_name(f"{input_file.stem}{OUTPUT_SUFFIX}")


def make_temp_path(input_file: Path) -> Path:
    return input_file.with_name(f"{input_file.stem}{TEMP_SUFFIX}")


def build_azure_client() -> AzureOpenAI:
    endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
    api_key = os.getenv("AZURE_OPENAI_API_KEY")

    missing = []
    if not endpoint:
        missing.append("AZURE_OPENAI_ENDPOINT")
    if not api_key:
        missing.append("AZURE_OPENAI_API_KEY")

    if missing:
        raise RuntimeError(
            "Variables d'environnement manquantes : "
            + ", ".join(missing)
            + ". Ajoute-les dans ton fichier .env."
        )

    return AzureOpenAI(
        azure_endpoint=endpoint,
        api_key=api_key,
        api_version=AZURE_OPENAI_API_VERSION,
    )


def get_thread_azure_client() -> AzureOpenAI:
    client = getattr(THREAD_LOCAL, "azure_client", None)
    if client is None:
        client = build_azure_client()
        THREAD_LOCAL.azure_client = client
    return client


def get_thread_http_session() -> requests.Session:
    session = getattr(THREAD_LOCAL, "http_session", None)
    if session is None:
        session = requests.Session()
        THREAD_LOCAL.http_session = session
    return session


def same_domain(url: str, base_url: str) -> bool:
    parsed_url = urlparse(url)
    parsed_base = urlparse(base_url)
    url_host = parsed_url.netloc.lower().removeprefix("www.")
    base_host = parsed_base.netloc.lower().removeprefix("www.")
    return bool(url_host and url_host == base_host)


def fetch_page_text_and_links(url: str, timeout: int = SCRAPE_TIMEOUT_SECONDS) -> tuple[str, list[str]]:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (compatible; StartupClassifier/1.0; "
            "for research and startup prequalification)"
        )
    }

    try:
        response = get_thread_http_session().get(url, headers=headers, timeout=timeout)
        response.raise_for_status()
        response.encoding = response.apparent_encoding or response.encoding
    except requests.RequestException:
        return "", []

    content_type = response.headers.get("content-type", "")
    if "text/html" not in content_type and "application/xhtml" not in content_type:
        return "", []

    soup = BeautifulSoup(response.text, "html.parser")

    for tag in soup(["script", "style", "noscript", "svg", "footer", "nav"]):
        tag.decompose()

    title = soup.title.get_text(" ", strip=True) if soup.title else ""

    meta_description = ""
    meta = soup.find("meta", attrs={"name": "description"})
    if meta and meta.get("content"):
        meta_description = meta.get("content")

    headings = " ".join(
        h.get_text(" ", strip=True)
        for h in soup.find_all(["h1", "h2", "h3"])[:15]
    )

    body_text = " ".join(
        tag.get_text(" ", strip=True)
        for tag in soup.find_all(["p", "li"])[:70]
    )

    links = []
    for anchor in soup.find_all("a", href=True):
        href = anchor.get("href", "")
        absolute_url = urljoin(url, href.split("#")[0])
        if absolute_url.startswith(("http://", "https://")):
            links.append(absolute_url)

    text = clean_text(f"{title}. {meta_description}. {headings}. {body_text}")
    return text, links


def research_link_score(url: str) -> int:
    normalized = normalize_text(url)
    priority_terms = {
        "about": 45,
        "solution": 40,
        "solutions": 40,
        "product": 35,
        "platform": 25,
        "industries": 35,
        "industry": 35,
        "customers": 50,
        "clients": 50,
        "case studies": 50,
        "case study": 50,
        "use cases": 35,
        "partners": 25,
        "press": 25,
        "news": 20,
        "blog": 10,
        "construction": 25,
        "energy": 25,
        "mobility": 25,
        "infrastructure": 25,
        "real estate": 25,
        "built": 20,
    }

    score = 0
    for term, weight in priority_terms.items():
        if term in normalized:
            score += weight

    return score


def select_research_links(base_url: str, links: list[str]) -> list[str]:
    base_url = normalize_url(base_url).rstrip("/")
    selected = [base_url]
    seen = {base_url}
    candidates = []

    for link in links:
        if not same_domain(link, base_url):
            continue
        parsed = urlparse(link)
        clean_link = parsed._replace(query="", fragment="").geturl().rstrip("/")
        if clean_link in seen:
            continue
        score = research_link_score(clean_link)
        if score <= 0:
            continue
        candidates.append((score, clean_link))

    candidates.sort(reverse=True)

    for _, link in candidates:
        if len(selected) >= RESEARCH_MAX_PAGES:
            break
        if link not in seen:
            selected.append(link)
            seen.add(link)

    return selected


def detect_keyword_matches(text: str, keywords: list[str]) -> list[str]:
    normalized = normalize_text(text)
    matches = []
    for keyword in keywords:
        normalized_keyword = normalize_text(keyword)
        if not normalized_keyword:
            continue

        found_positive = False
        left_boundary = r"(?<![a-z0-9])" if normalized_keyword[0].isalnum() else ""
        right_boundary = r"(?![a-z0-9])" if normalized_keyword[-1].isalnum() else ""
        pattern = re.compile(left_boundary + re.escape(normalized_keyword) + right_boundary)

        for match in pattern.finditer(normalized):
            position = match.start()
            prefix = normalized[max(0, position - 25):position]
            if not re.search(r"\b(not|no|non|pas|sans|not for|not targeting)\s+$", prefix):
                found_positive = True
                break

        if found_positive and keyword not in matches:
            matches.append(keyword)
    return matches[:12]


def detect_sector_matches(text: str) -> dict[str, list[str]]:
    matches = {}
    for sector, keywords in SECTOR_KEYWORDS.items():
        sector_matches = detect_keyword_matches(text, keywords)
        if sector_matches:
            matches[sector] = sector_matches[:8]
    return matches


def format_matches(matches) -> str:
    if isinstance(matches, dict):
        if not matches:
            return ""
        return " ; ".join(f"{key}: {', '.join(values)}" for key, values in matches.items())
    return ", ".join(matches) if matches else ""


def detect_reference_signals(text: str) -> dict[str, list[str]]:
    return {
        "vinci": detect_keyword_matches(text, VINCI_NAMES),
        "competitors": detect_keyword_matches(text, COMPETITORS),
    }


def website_quality_from_research(sources: list[str], text: str) -> str:
    if not sources:
        return "inaccessible"
    if len(text) < 400:
        return "faible"
    if len(sources) >= 3 and len(text) >= 2000:
        return "forte"
    return "moyenne"


def research_website(url: str) -> dict:
    if not is_valid_url(url):
        return {
            "text": "",
            "sources": "",
            "quality": "inconnu",
        }

    base_url = normalize_url(url).rstrip("/")

    if not ENABLE_WEBSITE_RESEARCH:
        homepage_text = scrape_homepage_text(base_url)
        return {
            "text": homepage_text,
            "sources": base_url if homepage_text else "",
            "quality": "désactivée" if homepage_text else "inaccessible",
        }

    homepage_text, links = fetch_page_text_and_links(base_url)

    selected_links = select_research_links(base_url, links)
    texts = []
    sources = []

    if homepage_text:
        texts.append(f"Source: {base_url}\n{homepage_text}")
        sources.append(base_url)

    for link in selected_links:
        if sum(len(item) for item in texts) >= RESEARCH_MAX_TEXT_CHARS:
            break
        if link == base_url:
            continue
        page_text, _ = fetch_page_text_and_links(link)
        if not page_text:
            continue
        texts.append(f"Source: {link}\n{page_text}")
        sources.append(link)
        if len(sources) >= RESEARCH_MAX_PAGES:
            break

    text = clean_text("\n\n".join(texts))[:RESEARCH_MAX_TEXT_CHARS]
    return {
        "text": text,
        "sources": " | ".join(sources),
        "quality": website_quality_from_research(sources, text),
    }


def research_website_cached(url: str, cache: dict[str, dict], cache_lock: threading.Lock) -> dict:
    if not is_valid_url(url):
        return {
            "text": "",
            "sources": "",
            "quality": "inconnu",
        }

    cache_key = normalize_url(url).rstrip("/")

    with cache_lock:
        if cache_key in cache:
            return cache[cache_key]

    result = research_website(cache_key)

    with cache_lock:
        cache[cache_key] = result

    return result


def build_audit_signals(
    startup_name: str,
    sector: str,
    description: str,
    site: str,
    website_text: str,
    excel_context: str = "",
) -> dict:
    combined_text = clean_text(" ".join([
        startup_name,
        sector,
        description,
        site,
        excel_context,
        website_text,
    ]))
    references = detect_reference_signals(combined_text)

    return {
        "b2b": format_matches(detect_keyword_matches(combined_text, B2B_KEYWORDS)),
        "b2c": format_matches(detect_keyword_matches(combined_text, B2C_KEYWORDS)),
        "artisan_only_risk": format_matches(detect_keyword_matches(combined_text, ARTISAN_KEYWORDS)),
        "sectors": format_matches(detect_sector_matches(combined_text)),
        "vinci": format_matches(references["vinci"]),
        "competitors": format_matches(references["competitors"]),
    }


def build_excel_shortcut_result_algo1(
    startup_name: str,
    sector: str,
    description: str,
    site: str,
    excel_context: str,
    excel_context_info: dict,
) -> dict | None:
    if not ENABLE_EXCEL_SHORTCUT:
        return None

    text = clean_text(" ".join([startup_name, sector, description, site, excel_context]))
    if len(text) < 120:
        return None

    references = detect_reference_signals(text)
    b2b_matches = detect_keyword_matches(text, B2B_KEYWORDS)
    b2c_matches = detect_keyword_matches(text, B2C_KEYWORDS)
    artisan_matches = detect_keyword_matches(text, ARTISAN_KEYWORDS)
    sector_matches = detect_sector_matches(text)
    semantic_score = semantic_prefilter.score_algo1(text)
    learning_match = learning_memory.predict("algo1", text)

    selected_sector = next(iter(sector_matches.keys()), semantic_score.get("sector", "inconnu"))
    b2b_count = len(b2b_matches)
    b2c_count = len(b2c_matches)
    artisan_count = len(artisan_matches)
    has_sector = bool(sector_matches)
    semantic_b2b_strong = semantic_prefilter.is_strong(float(semantic_score.get("b2b_score", 0) or 0))
    semantic_b2c_strong = semantic_prefilter.is_strong(float(semantic_score.get("b2c_score", 0) or 0))
    semantic_sector_strong = semantic_prefilter.is_strong(float(semantic_score.get("sector_score", 0) or 0))

    result = None
    if learning_match.get("matched") and learning_match.get("label") in {"A_GARDER", "A_VERIFIER", "A_ECARTER"}:
        learned_label = learning_match["label"]
        result = {
            "decision": learned_label,
            "confidence": min(0.92, max(0.76, float(learning_match.get("score", 0) or 0))),
            "is_b2b": "oui" if learned_label == "A_GARDER" else "inconnu",
            "target_sector": selected_sector,
            "sector_relevance": "moyenne" if selected_sector != "inconnu" else "inconnue",
            "sector_reason": "Décision proposée par mémoire d'apprentissage à partir d'une correction similaire.",
            "vinci_reference_detected": "oui" if references["vinci"] else "non",
            "competitor_reference_detected": "oui" if references["competitors"] else "non",
            "official_website_likely": "inconnu",
            "short_description_clean": description[:600],
            "reason": (
                "Décision sans API : mémoire d'apprentissage proche "
                f"({learning_match.get('name') or learning_match.get('source')})."
            ),
            "missing_information": "Vérifier si la correction historique est bien comparable.",
        }
    elif references["vinci"]:
        result = {
            "decision": "A_GARDER",
            "confidence": 0.95,
            "is_b2b": "oui" if b2b_count else "inconnu",
            "target_sector": selected_sector,
            "sector_relevance": "forte" if has_sector else "moyenne",
            "sector_reason": "Référence VINCI explicite détectée dans les données Excel.",
            "vinci_reference_detected": "oui",
            "competitor_reference_detected": "non",
            "official_website_likely": "inconnu",
            "short_description_clean": description[:600],
            "reason": "Décision prise sans API : référence VINCI explicite dans l'Excel.",
            "missing_information": "",
        }
    elif references["competitors"] and has_sector:
        result = {
            "decision": "A_GARDER",
            "confidence": 0.9,
            "is_b2b": "oui" if b2b_count else "inconnu",
            "target_sector": selected_sector,
            "sector_relevance": "forte",
            "sector_reason": "Référence concurrent et secteur cible détectés dans les données Excel.",
            "vinci_reference_detected": "non",
            "competitor_reference_detected": "oui",
            "official_website_likely": "inconnu",
            "short_description_clean": description[:600],
            "reason": "Décision prise sans API : référence concurrent et secteur cible suffisamment explicites.",
            "missing_information": "",
        }
    elif (
        (b2b_count >= 2 or semantic_b2b_strong)
        and (has_sector or semantic_sector_strong)
        and b2c_count == 0
        and artisan_count == 0
        and not semantic_b2c_strong
    ):
        result = {
            "decision": "A_GARDER",
            "confidence": 0.82,
            "is_b2b": "oui",
            "target_sector": selected_sector,
            "sector_relevance": "forte" if sum(len(v) for v in sector_matches.values()) >= 3 or semantic_sector_strong else "moyenne",
            "sector_reason": "Signaux B2B et sectoriels suffisants dans les colonnes Excel et le score sémantique local.",
            "vinci_reference_detected": "non",
            "competitor_reference_detected": "non",
            "official_website_likely": "inconnu",
            "short_description_clean": description[:600],
            "reason": "Décision prise sans API : l'Excel et le score sémantique local contiennent déjà des preuves B2B et sectorielles suffisantes.",
            "missing_information": "",
        }
    elif (b2c_count >= 2 or semantic_b2c_strong) and not has_sector and not semantic_sector_strong and b2b_count == 0:
        result = {
            "decision": "A_ECARTER",
            "confidence": 0.86,
            "is_b2b": "non",
            "target_sector": "hors_scope",
            "sector_relevance": "faible",
            "sector_reason": "Signaux B2C forts et absence de secteur cible dans les données Excel.",
            "vinci_reference_detected": "non",
            "competitor_reference_detected": "non",
            "official_website_likely": "inconnu",
            "short_description_clean": description[:600],
            "reason": "Décision prise sans API : l'Excel indique une activité B2C/hors scope.",
            "missing_information": "",
        }

    if result is None:
        return None

    result["algo1_scraped_sources"] = ""
    result["algo1_website_research_quality"] = "non_requise_excel_suffisant"
    result["algo1_b2b_signals_detected"] = format_matches(b2b_matches)
    result["algo1_b2c_signals_detected"] = format_matches(b2c_matches)
    result["algo1_artisan_only_risk_detected"] = format_matches(artisan_matches)
    result["algo1_sector_signals_detected"] = format_matches(sector_matches)
    result["algo1_vinci_signals_detected"] = format_matches(references["vinci"])
    result["algo1_competitor_signals_detected"] = format_matches(references["competitors"])
    result["algo1_research_text_used"] = ""
    result["algo1_excel_columns_used"] = " | ".join(excel_context_info["used_columns"])
    result["algo1_excel_columns_used_count"] = excel_context_info["used_count"]
    result["algo1_excel_columns_available_count"] = excel_context_info["available_count"]
    result["algo1_excel_context_truncated"] = "oui" if excel_context_info["truncated"] else "non"
    result["algo1_processing_mode"] = "excel_shortcut_no_api_no_research"
    result["algo1_api_called"] = "non"
    result["algo1_research_called"] = "non"
    result["algo1_semantic_backend"] = semantic_score.get("backend", "")
    result["algo1_semantic_summary"] = semantic_prefilter.summarize(semantic_score)
    result["algo1_semantic_sector"] = semantic_score.get("sector", "")
    result["algo1_semantic_sector_score"] = semantic_score.get("sector_score", 0)
    result["algo1_semantic_b2b_score"] = semantic_score.get("b2b_score", 0)
    result["algo1_semantic_b2c_score"] = semantic_score.get("b2c_score", 0)
    result["algo1_learning_match_label"] = learning_match.get("label", "")
    result["algo1_learning_match_score"] = learning_match.get("score", 0)
    result["algo1_learning_match_source"] = learning_match.get("source", "")
    return apply_profile_rules(result, text)


# ============================================================
# SCRAPING SIMPLE DU SITE
# ============================================================

def scrape_homepage_text(url: str, timeout: int = SCRAPE_TIMEOUT_SECONDS) -> str:
    """
    Récupère un texte court depuis la page d'accueil du site.
    Si le site est inaccessible, retourne une chaîne vide.
    """

    if not is_valid_url(url):
        return ""

    url = normalize_url(url)

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (compatible; StartupClassifier/1.0; "
            "for research and startup prequalification)"
        )
    }

    try:
        response = get_thread_http_session().get(url, headers=headers, timeout=timeout)
        response.raise_for_status()
        response.encoding = response.apparent_encoding or response.encoding

        soup = BeautifulSoup(response.text, "html.parser")

        for tag in soup(["script", "style", "noscript", "svg", "footer", "nav"]):
            tag.decompose()

        title = soup.title.get_text(" ", strip=True) if soup.title else ""

        meta_description = ""
        meta = soup.find("meta", attrs={"name": "description"})
        if meta and meta.get("content"):
            meta_description = meta.get("content")

        headings = " ".join(
            h.get_text(" ", strip=True)
            for h in soup.find_all(["h1", "h2", "h3"])[:12]
        )

        paragraphs = " ".join(
            p.get_text(" ", strip=True)
            for p in soup.find_all("p")[:25]
        )

        text = f"{title}. {meta_description}. {headings}. {paragraphs}"
        text = clean_text(text)

        return text[:MAX_WEBSITE_TEXT_CHARS]

    except requests.RequestException:
        return ""


def scrape_homepage_text_cached(url: str, cache: dict[str, str], cache_lock: threading.Lock) -> str:
    if not is_valid_url(url):
        return ""

    cache_key = normalize_url(url)

    with cache_lock:
        if cache_key in cache:
            return cache[cache_key]

    text = scrape_homepage_text(cache_key)

    with cache_lock:
        cache[cache_key] = text

    return text


# ============================================================
# PROMPT
# ============================================================

def build_prompt(
    startup_name: str,
    sector: str,
    description: str,
    site: str,
    website_text: str,
    excel_context: str = "",
    excel_context_columns: str = "",
    excel_context_truncated: str = "non",
    audit_signals: dict | None = None,
    scraped_sources: str = "",
    website_research_quality: str = "",
) -> str:
    competitors = ", ".join(COMPETITORS)
    target_sectors = ", ".join(TARGET_SECTORS)
    audit_signals = audit_signals or {}
    profile_context = dealflow_profiles.profile_prompt_context()

    return f"""
Tu es un analyste innovation pour Leonard, la plateforme d'innovation et de prospective du groupe VINCI.

Ta mission est de préqualifier une startup pour savoir si elle est pertinente pour VINCI.

Profil de traitement :
{profile_context}

L'algorithme 1 doit classer chaque startup en :
- A_GARDER
- A_VERIFIER
- A_ECARTER

Critères principaux :

1. Critère B2B
La startup doit être principalement B2B.
Une startup B2B vend à des entreprises, collectivités, industriels, opérateurs, promoteurs,
constructeurs, gestionnaires d'actifs, acteurs publics ou professionnels.
Une startup principalement B2C vend surtout à des particuliers.
Sois strict : si le client/payeur principal n'est pas clairement professionnel, mets is_b2b = "inconnu"
ou "non", mais pas "oui".

2. Critère secteur
La startup doit appartenir clairement à au moins un secteur cible pertinent pour VINCI.
Les seuls secteurs cibles autorisés sont : {target_sectors}.
Built World est un secteur cible de la liste, au même niveau que Construction, Infrastructure,
Real Estate, Energy et Mobility. Ne le traite pas comme un macro-secteur.

3. Référence VINCI
Si la startup a déjà travaillé avec VINCI ou une filiale VINCI, c'est un signal très fort.
Mets vinci_reference_detected = "oui" uniquement si c'est explicitement indiqué dans les données
ou sur le site. N'invente jamais cette information.

4. Référence concurrent
Si la startup a déjà travaillé avec un concurrent ou acteur proche de VINCI, c'est un signal positif.
Exemples : {competitors}.
Mets competitor_reference_detected = "oui" uniquement si c'est explicitement indiqué.

Décision :

A_GARDER :
- startup clairement B2B ;
- et secteur clairement dans la liste cible ;
- ou référence VINCI clairement détectée ;
- ou référence concurrent clairement détectée.

A_VERIFIER :
- startup potentiellement pertinente mais informations insuffisantes ;
- secteur possiblement dans la liste mais pas certain ;
- B2B non confirmé ;
- description trop vague ;
- site web inaccessible, douteux ou probablement pas officiel ;
- activité intéressante mais besoin de vérification humaine.

A_ECARTER :
- startup clairement B2C ;
- secteur clairement hors de la liste cible ;
- aucun lien clair avec Construction, Built World, Infrastructure, Real Estate, Energy ou Mobility ;
- activité non pertinente pour VINCI.

Règles de prudence :
- Si l'information manque, écris "inconnu".
- Ne devine pas.
- Ne classe pas A_GARDER si le B2B ou le secteur cible n'est pas clair, sauf référence VINCI
  ou concurrent explicitement détectée.
- Si tu hésites entre A_GARDER et A_VERIFIER, choisis A_VERIFIER.
- Si le site ne semble pas être le site officiel de la startup, indique-le.
- Les signaux automatiques ci-dessous sont des indices, pas des preuves absolues.
- Si les signaux B2C sont forts ou si seuls des artisans/freelances sont détectés, sois prudent :
  ne conclus pas automatiquement à un B2B fort.
- Si la recherche site est faible ou inaccessible, ne classe A_GARDER que si l'Excel suffit clairement.
- Utilise toutes les informations Excel complémentaires ci-dessous. Les réponses directes de la startup
  sur ses clients, cas d'usage, technologie, verticale, chaîne de valeur, références et marché sont
  particulièrement importantes.
- Les données administratives, financières, TRL, financement ou taille d'équipe peuvent servir de contexte,
  mais elles ne remplacent jamais les preuves B2B et sectorielles nécessaires à A_GARDER.
- En cas de contradiction entre plusieurs colonnes, mentionne-la dans missing_information et choisis
  A_VERIFIER plutôt que d'inventer une interprétation.
- La langue de réponse doit suivre le profil de traitement ci-dessus.

Réponds uniquement avec un objet JSON valide contenant exactement ces champs :
- decision
- confidence
- is_b2b
- target_sector
- sector_relevance
- sector_reason
- vinci_reference_detected
- competitor_reference_detected
- official_website_likely
- short_description_clean
- reason
- missing_information

Données disponibles :

Nom de la startup :
{startup_name}

Secteur déjà présent dans l'Excel :
{sector}

Description déjà présente dans l'Excel :
{description}

Site web :
{site}

Toutes les informations Excel utiles disponibles pour cette startup :
{excel_context or "aucune information complémentaire"}

Colonnes Excel utilisées :
{excel_context_columns or "aucune"}

Contexte Excel tronqué par sécurité :
{excel_context_truncated}

Texte extrait du site web :
{website_text}

Sources consultées sur le site :
{scraped_sources}

Qualité de la recherche site :
{website_research_quality}

Signaux B2B détectés automatiquement :
{audit_signals.get("b2b", "") or "aucun"}

Signaux B2C détectés automatiquement :
{audit_signals.get("b2c", "") or "aucun"}

Risque artisans/freelances détecté :
{audit_signals.get("artisan_only_risk", "") or "aucun"}

Signaux sectoriels détectés automatiquement :
{audit_signals.get("sectors", "") or "aucun"}

Signaux VINCI détectés automatiquement :
{audit_signals.get("vinci", "") or "aucun"}

Signaux concurrents détectés automatiquement :
{audit_signals.get("competitors", "") or "aucun"}
""".strip()


def classification_schema() -> dict:
    return {
        "type": "object",
        "properties": {
            "decision": {
                "type": "string",
                "enum": ["A_GARDER", "A_VERIFIER", "A_ECARTER"],
            },
            "confidence": {
                "type": "number",
                "description": "Score de confiance entre 0 et 1.",
            },
            "is_b2b": {
                "type": "string",
                "enum": ["oui", "non", "inconnu"],
            },
            "target_sector": {
                "type": "string",
                "enum": [
                    "Construction",
                    "Built World",
                    "Infrastructure",
                    "Real Estate",
                    "Energy",
                    "Mobility",
                    "inconnu",
                    "hors_scope",
                ],
                "description": "Secteur cible principal identifié.",
            },
            "sector_relevance": {
                "type": "string",
                "enum": ["forte", "moyenne", "faible", "inconnue"],
                "description": "Pertinence du secteur de la startup par rapport aux secteurs cibles.",
            },
            "sector_reason": {
                "type": "string",
                "description": "Explication courte du rattachement sectoriel.",
            },
            "vinci_reference_detected": {
                "type": "string",
                "enum": ["oui", "non", "inconnu"],
            },
            "competitor_reference_detected": {
                "type": "string",
                "enum": ["oui", "non", "inconnu"],
            },
            "official_website_likely": {
                "type": "string",
                "enum": ["oui", "non", "inconnu"],
            },
            "short_description_clean": {
                "type": "string",
                "description": "Description courte et propre de l'activité de la startup en français.",
            },
            "reason": {
                "type": "string",
                "description": "Explication courte de la décision.",
            },
            "missing_information": {
                "type": "string",
                "description": "Informations manquantes à vérifier manuellement.",
            },
        },
        "required": [
            "decision",
            "confidence",
            "is_b2b",
            "target_sector",
            "sector_relevance",
            "sector_reason",
            "vinci_reference_detected",
            "competitor_reference_detected",
            "official_website_likely",
            "short_description_clean",
            "reason",
            "missing_information",
        ],
        "additionalProperties": False,
    }


# ============================================================
# CLASSIFICATION AVEC AZURE OPENAI
# ============================================================

def error_result(message: str) -> dict:
    return {
        "decision": "A_VERIFIER",
        "confidence": 0,
        "is_b2b": "inconnu",
        "target_sector": "inconnu",
        "sector_relevance": "inconnue",
        "sector_reason": "Erreur ou informations insuffisantes.",
        "vinci_reference_detected": "inconnu",
        "competitor_reference_detected": "inconnu",
        "official_website_likely": "inconnu",
        "short_description_clean": "",
        "reason": message,
        "missing_information": "Classification à refaire manuellement.",
    }


def compact_error_message(text: str, max_length: int = 700) -> str:
    text = re.sub(r"\s+", " ", str(text)).strip()
    if len(text) <= max_length:
        return text
    return text[:max_length] + "..."


def format_api_error(exc: Exception) -> str:
    status_code = getattr(exc, "status_code", None)
    code = getattr(exc, "code", None)
    body = getattr(exc, "body", None)

    details = [exc.__class__.__name__]

    if status_code:
        details.append(f"status={status_code}")

    if code:
        details.append(f"code={code}")

    message = ""
    if isinstance(body, dict):
        error = body.get("error")
        if isinstance(error, dict):
            message = error.get("message") or error.get("code") or ""
        else:
            message = body.get("message") or ""

    if not message:
        message = str(exc)

    if message:
        details.append(compact_error_message(message))

    return " | ".join(details)


def is_non_retryable_api_error(exc: Exception) -> bool:
    return getattr(exc, "status_code", None) in NON_RETRYABLE_API_STATUS_CODES


def get_retry_after_seconds(exc: Exception) -> float | None:
    response = getattr(exc, "response", None)
    headers = getattr(response, "headers", None)
    if not headers:
        return None

    retry_after = headers.get("retry-after") or headers.get("Retry-After")
    if not retry_after:
        return None

    try:
        return max(0.0, float(retry_after))
    except ValueError:
        return None


def parse_json_text(text: str) -> dict:
    text = clean_text(text)

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start != -1 and end != -1 and end > start:
            return json.loads(text[start:end + 1])
        raise


def build_messages(prompt: str) -> list[dict]:
    return [
        {
            "role": "system",
            "content": (
                "Tu es un analyste innovation rigoureux. "
                "Tu réponds uniquement avec un JSON valide, sans markdown."
            ),
        },
        {
            "role": "user",
            "content": prompt,
        },
    ]


def call_responses_json_schema(
    client: AzureOpenAI,
    messages: list[dict],
    schema: dict,
    schema_name: str,
) -> str:
    response = client.responses.create(
        model=MODEL_DEPLOYMENT,
        input=messages,
        text={
            "format": {
                "type": "json_schema",
                "name": schema_name,
                "schema": schema,
                "strict": True,
            }
        },
    )
    return response.output_text


def call_chat_json_schema(
    client: AzureOpenAI,
    messages: list[dict],
    schema: dict,
    schema_name: str,
) -> str:
    response = client.chat.completions.create(
        model=MODEL_DEPLOYMENT,
        messages=messages,
        response_format={
            "type": "json_schema",
            "json_schema": {
                "name": schema_name,
                "schema": schema,
                "strict": True,
            },
        },
    )
    return response.choices[0].message.content or ""


def call_chat_json_object(client: AzureOpenAI, messages: list[dict]) -> str:
    response = client.chat.completions.create(
        model=MODEL_DEPLOYMENT,
        messages=messages,
        response_format={"type": "json_object"},
    )
    return response.choices[0].message.content or ""


def call_model_json(
    client: AzureOpenAI,
    messages: list[dict],
    schema: dict,
    schema_name: str,
) -> dict:
    mode = ALGO1_API_MODE

    if mode not in {"chat", "responses", "auto", "json_object"}:
        mode = "chat"

    if mode == "responses":
        return parse_json_text(call_responses_json_schema(client, messages, schema, schema_name))

    if mode == "json_object":
        return parse_json_text(call_chat_json_object(client, messages))

    if mode == "auto":
        try:
            return parse_json_text(call_responses_json_schema(client, messages, schema, schema_name))
        except Exception as exc:
            if getattr(exc, "status_code", None) != 400:
                raise
            log(f"Responses API refusée par Azure, bascule en chat : {format_api_error(exc)}")

    try:
        return parse_json_text(call_chat_json_schema(client, messages, schema, schema_name))
    except Exception as exc:
        if getattr(exc, "status_code", None) != 400:
            raise
        log(f"JSON schema refusé par Azure, bascule en json_object : {format_api_error(exc)}")
        return parse_json_text(call_chat_json_object(client, messages))


def validate_azure_client() -> None:
    client = build_azure_client()

    schema = {
        "type": "object",
        "properties": {
            "ok": {
                "type": "string",
                "enum": ["oui"],
            }
        },
        "required": ["ok"],
        "additionalProperties": False,
    }

    try:
        call_model_json(
            client=client,
            messages=[
                {
                    "role": "system",
                    "content": "Tu réponds uniquement avec un JSON valide, sans markdown.",
                },
                {
                    "role": "user",
                    "content": 'Réponds avec ce JSON exact : {"ok":"oui"}',
                },
            ],
            schema=schema,
            schema_name="algo1_healthcheck",
        )
        log("Connexion Azure OpenAI OK.")

    except Exception as exc:
        raise RuntimeError(f"Test Azure OpenAI échoué : {format_api_error(exc)}")


def classify_startup(
    client: AzureOpenAI,
    startup_name: str,
    sector: str,
    description: str,
    site: str,
    website_text: str,
    excel_context: str = "",
    excel_context_columns: str = "",
    excel_context_truncated: str = "non",
    audit_signals: dict | None = None,
    scraped_sources: str = "",
    website_research_quality: str = "",
) -> dict:
    prompt = build_prompt(
        startup_name=startup_name,
        sector=sector,
        description=description,
        site=site,
        website_text=website_text,
        excel_context=excel_context,
        excel_context_columns=excel_context_columns if INCLUDE_EXCEL_COLUMN_LIST_IN_PROMPT else "",
        excel_context_truncated=excel_context_truncated,
        audit_signals=audit_signals,
        scraped_sources=scraped_sources,
        website_research_quality=website_research_quality,
    )
    schema = classification_schema()

    for attempt in range(1, MAX_API_RETRIES + 1):
        try:
            return call_model_json(
                client=client,
                messages=build_messages(prompt),
                schema=schema,
                schema_name="startup_prequalification",
            )

        except Exception as exc:
            error_details = format_api_error(exc)

            if is_non_retryable_api_error(exc):
                log(f"Erreur API non réessayée : {error_details}")
                return error_result(f"Erreur pendant la classification : {error_details}")

            if attempt == MAX_API_RETRIES:
                log(f"Erreur API finale : {error_details}")
                return error_result(f"Erreur pendant la classification : {error_details}")

            retry_after_seconds = get_retry_after_seconds(exc)
            wait_seconds = retry_after_seconds or min(30, (2 ** attempt) + random.uniform(0, 0.5))
            log(
                "Erreur API, nouvel essai dans "
                f"{wait_seconds:.1f}s ({attempt}/{MAX_API_RETRIES}) : {error_details}"
            )
            time.sleep(wait_seconds)

    return error_result("Erreur API inconnue.")


# ============================================================
# REGLES FINALES APRES MODELE
# ============================================================

def apply_final_rules(result: dict) -> dict:
    decision = result.get("decision", "A_VERIFIER")
    confidence = float(result.get("confidence", 0) or 0)
    is_b2b = result.get("is_b2b", "inconnu")
    target_sector = result.get("target_sector", "inconnu")
    sector_relevance = result.get("sector_relevance", "inconnue")
    vinci_ref = result.get("vinci_reference_detected", "inconnu")
    competitor_ref = result.get("competitor_reference_detected", "inconnu")
    website_official = result.get("official_website_likely", "inconnu")

    if vinci_ref == "oui":
        result["decision_finale"] = "A_GARDER"
        result["regle_finale"] = "Référence VINCI détectée."
        return result

    if competitor_ref == "oui":
        result["decision_finale"] = "A_GARDER"
        result["regle_finale"] = "Référence concurrent détectée."
        return result

    if decision == "A_GARDER" and confidence < 0.75:
        result["decision_finale"] = "A_VERIFIER"
        result["regle_finale"] = "Décision positive mais confiance insuffisante."
        return result

    if decision == "A_GARDER" and is_b2b != "oui":
        result["decision_finale"] = "A_VERIFIER"
        result["regle_finale"] = "B2B non confirmé."
        return result

    if decision == "A_GARDER" and target_sector in ["inconnu", "hors_scope"]:
        result["decision_finale"] = "A_VERIFIER"
        result["regle_finale"] = "Secteur cible non confirmé."
        return result

    if decision == "A_GARDER" and sector_relevance == "faible":
        result["decision_finale"] = "A_VERIFIER"
        result["regle_finale"] = "Pertinence sectorielle trop faible pour garder automatiquement."
        return result

    if decision == "A_GARDER" and website_official == "non":
        result["decision_finale"] = "A_VERIFIER"
        result["regle_finale"] = "Site web probablement non officiel."
        return result

    result["decision_finale"] = decision
    result["regle_finale"] = "Décision conservée après règles finales."
    return result


def apply_profile_rules(result: dict, text: str) -> dict:
    result = apply_final_rules(result)
    result["dealflow_profile"] = dealflow_profiles.profile_label()
    exclusions = dealflow_profiles.detect_excluded_scope(text)
    result["profile_exclusion_matches"] = dealflow_profiles.format_matches(exclusions)

    if exclusions:
        result["decision"] = "A_ECARTER"
        result["decision_finale"] = "A_ECARTER"
        result["target_sector"] = "hors_scope"
        result["sector_relevance"] = "faible"
        result["sector_reason"] = "LATAM excluded scope: " + dealflow_profiles.format_matches(exclusions)
        result["reason"] = (
            "LATAM profile exclusion: company appears mainly related to an excluded scope "
            f"({dealflow_profiles.format_matches(exclusions)})."
        )
        result["missing_information"] = ""
        result["regle_finale"] = "LATAM profile exclusion applied."

    return result


# ============================================================
# TRAITEMENT EXCEL
# ============================================================

def detect_columns(df: pd.DataFrame) -> dict[str, str]:
    columns = {
        key: find_column(df, aliases)
        for key, aliases in COLUMN_ALIASES.items()
    }

    if columns["startup"] is None:
        raise ValueError("Impossible de trouver la colonne du nom de la startup.")

    if columns["sector"] is None:
        df["secteur"] = ""
        columns["sector"] = "secteur"

    if columns["description"] is None:
        df["description"] = ""
        columns["description"] = "description"

    if columns["site"] is None:
        df["site web"] = ""
        columns["site"] = "site web"

    return columns


def empty_row_result() -> dict:
    return {
        "decision": "A_VERIFIER",
        "confidence": 0,
        "is_b2b": "inconnu",
        "target_sector": "inconnu",
        "sector_relevance": "inconnue",
        "sector_reason": "Aucune information exploitable.",
        "vinci_reference_detected": "inconnu",
        "competitor_reference_detected": "inconnu",
        "official_website_likely": "inconnu",
        "short_description_clean": "",
        "reason": "Ligne vide ou données insuffisantes.",
        "missing_information": "Nom, description et site web à compléter.",
        "decision_finale": "A_VERIFIER",
        "regle_finale": "Données insuffisantes.",
    }


def normalize_resume_value(value):
    if pd.isna(value):
        return ""
    return value


def normalize_resume_record(record: dict) -> dict:
    return {
        key: normalize_resume_value(value)
        for key, value in record.items()
    }


def saved_result_is_failed(record: dict) -> bool:
    decision = clean_text(record.get("decision", ""))
    decision_finale = clean_text(record.get("decision_finale", ""))
    reason = clean_text(record.get("reason", ""))

    if not decision or not decision_finale:
        return True

    reason_lower = reason.lower()
    return any(pattern.lower() in reason_lower for pattern in FAILED_TEMP_REASON_PATTERNS)


def load_resume_results(temp_file: Path, expected_rows: int) -> list[dict]:
    if not RESUME_FROM_TEMP:
        return []

    if not temp_file.exists():
        return []

    try:
        temp_df = pd.read_excel(temp_file, sheet_name="Toutes les startups")
    except Exception as exc:
        log(f"Reprise impossible depuis {temp_file.name} : {exc}")
        return []

    if temp_df.empty:
        return []

    saved_rows = min(len(temp_df), expected_rows)
    if saved_rows < len(temp_df):
        log(
            f"Sauvegarde temporaire plus longue que le fichier courant : "
            f"{saved_rows}/{len(temp_df)} lignes utilisables."
        )

    result_columns = [column for column in RESULT_COLUMNS if column in temp_df.columns]
    if not result_columns:
        log(f"Reprise impossible : aucune colonne résultat reconnue dans {temp_file.name}.")
        return []

    temp_records = [
        normalize_resume_record(record)
        for record in temp_df.iloc[:saved_rows][result_columns].to_dict("records")
    ]

    resume_records = temp_records
    if RETRY_FAILED_TEMP_ROWS:
        for index, record in enumerate(temp_records):
            if saved_result_is_failed(record):
                resume_records = temp_records[:index]
                reason = compact_error_message(record.get("reason", ""), 180)
                log(
                    f"Première ligne temporaire à refaire : ligne {index + 1} "
                    f"({reason or 'résultat incomplet'})."
                )
                break

    if resume_records:
        log(
            f"Reprise depuis {temp_file.name} : "
            f"{len(resume_records)} ligne(s) conservée(s), "
            f"{max(0, expected_rows - len(resume_records))} ligne(s) à traiter."
        )

    return resume_records


def save_results(df_original: pd.DataFrame, results: list[dict], output_file: Path) -> None:
    results_df = pd.DataFrame(results)

    partial_df = df_original.iloc[:len(results)].reset_index(drop=True)
    final_df = pd.concat([partial_df, results_df.reset_index(drop=True)], axis=1)
    final_df = sanitize_excel_dataframe(final_df)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="Toutes les startups", index=False)

        for decision, sheet_name in [
            ("A_GARDER", "A garder"),
            ("A_VERIFIER", "A verifier"),
            ("A_ECARTER", "A ecarter"),
        ]:
            if "decision_finale" in final_df.columns:
                sheet_df = final_df[final_df["decision_finale"] == decision]
            else:
                sheet_df = final_df.iloc[0:0]

            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

        if "sauvegarde_temp" not in output_file.stem.lower():
            excel_styling.style_workbook(writer, title="Résultats Algo 1 - Éligibilité")


def process_startup_row(
    row: pd.Series,
    columns: dict[str, str],
    website_cache: dict[str, dict],
    cache_lock: threading.Lock,
) -> dict:
    try:
        startup_name = clean_text(row.get(columns["startup"], ""))
        sector = clean_text(row.get(columns["sector"], ""))
        description = clean_text(row.get(columns["description"], ""))
        site = clean_text(row.get(columns["site"], ""))
        excel_context_info = build_excel_row_context(row)
        excel_context = excel_context_info["text"]
        excel_context_columns = " | ".join(excel_context_info["used_columns"])
        excel_context_truncated = "oui" if excel_context_info["truncated"] else "non"

        if not startup_name and not description and not site and not excel_context:
            return empty_row_result()

        shortcut_result = build_excel_shortcut_result_algo1(
            startup_name=startup_name,
            sector=sector,
            description=description,
            site=site,
            excel_context=excel_context,
            excel_context_info=excel_context_info,
        )
        if shortcut_result is not None:
            return shortcut_result

        if DISABLE_API:
            result = error_result(
                "API Algo 1 désactivée et Excel insuffisant pour une décision automatique fiable."
            )
            result["algo1_scraped_sources"] = ""
            result["algo1_website_research_quality"] = "non_requise_api_desactivee"
            result["algo1_b2b_signals_detected"] = ""
            result["algo1_b2c_signals_detected"] = ""
            result["algo1_artisan_only_risk_detected"] = ""
            result["algo1_sector_signals_detected"] = ""
            result["algo1_vinci_signals_detected"] = ""
            result["algo1_competitor_signals_detected"] = ""
            result["algo1_research_text_used"] = ""
            result["algo1_excel_columns_used"] = excel_context_columns
            result["algo1_excel_columns_used_count"] = excel_context_info["used_count"]
            result["algo1_excel_columns_available_count"] = excel_context_info["available_count"]
            result["algo1_excel_context_truncated"] = excel_context_truncated
            result["algo1_processing_mode"] = "api_disabled_excel_insufficient"
            result["algo1_api_called"] = "non"
            result["algo1_research_called"] = "non"
            semantic_score = semantic_prefilter.score_algo1(" ".join([startup_name, sector, description, site, excel_context]))
            result["algo1_semantic_backend"] = semantic_score.get("backend", "")
            result["algo1_semantic_summary"] = semantic_prefilter.summarize(semantic_score)
            result["algo1_semantic_sector"] = semantic_score.get("sector", "")
            result["algo1_semantic_sector_score"] = semantic_score.get("sector_score", 0)
            result["algo1_semantic_b2b_score"] = semantic_score.get("b2b_score", 0)
            result["algo1_semantic_b2c_score"] = semantic_score.get("b2c_score", 0)
            learning_match = learning_memory.predict("algo1", " ".join([startup_name, sector, description, site, excel_context]))
            result["algo1_learning_match_label"] = learning_match.get("label", "")
            result["algo1_learning_match_score"] = learning_match.get("score", 0)
            result["algo1_learning_match_source"] = learning_match.get("source", "")
            return apply_profile_rules(result, " ".join([startup_name, sector, description, site, excel_context]))

        website_research = research_website_cached(site, website_cache, cache_lock)
        website_text = website_research.get("text", "")
        scraped_sources = website_research.get("sources", "")
        website_research_quality = website_research.get("quality", "inconnu")
        audit_signals = build_audit_signals(
            startup_name=startup_name,
            sector=sector,
            description=description,
            site=site,
            website_text=website_text,
            excel_context=excel_context,
        )

        result = classify_startup(
            client=get_thread_azure_client(),
            startup_name=startup_name,
            sector=sector,
            description=description,
            site=site,
            website_text=website_text,
            excel_context=excel_context,
            excel_context_columns=excel_context_columns,
            excel_context_truncated=excel_context_truncated,
            audit_signals=audit_signals,
            scraped_sources=scraped_sources,
            website_research_quality=website_research_quality,
        )

        result["algo1_scraped_sources"] = scraped_sources
        result["algo1_website_research_quality"] = website_research_quality
        result["algo1_b2b_signals_detected"] = audit_signals.get("b2b", "")
        result["algo1_b2c_signals_detected"] = audit_signals.get("b2c", "")
        result["algo1_artisan_only_risk_detected"] = audit_signals.get("artisan_only_risk", "")
        result["algo1_sector_signals_detected"] = audit_signals.get("sectors", "")
        result["algo1_vinci_signals_detected"] = audit_signals.get("vinci", "")
        result["algo1_competitor_signals_detected"] = audit_signals.get("competitors", "")
        result["algo1_research_text_used"] = website_text
        result["algo1_excel_columns_used"] = excel_context_columns
        result["algo1_excel_columns_used_count"] = excel_context_info["used_count"]
        result["algo1_excel_columns_available_count"] = excel_context_info["available_count"]
        result["algo1_excel_context_truncated"] = excel_context_truncated
        result["algo1_processing_mode"] = "api_with_research"
        result["algo1_api_called"] = "oui"
        result["algo1_research_called"] = "oui" if scraped_sources else "non"
        semantic_score = semantic_prefilter.score_algo1(" ".join([startup_name, sector, description, site, excel_context]))
        result["algo1_semantic_backend"] = semantic_score.get("backend", "")
        result["algo1_semantic_summary"] = semantic_prefilter.summarize(semantic_score)
        result["algo1_semantic_sector"] = semantic_score.get("sector", "")
        result["algo1_semantic_sector_score"] = semantic_score.get("sector_score", 0)
        result["algo1_semantic_b2b_score"] = semantic_score.get("b2b_score", 0)
        result["algo1_semantic_b2c_score"] = semantic_score.get("b2c_score", 0)
        learning_match = learning_memory.predict("algo1", " ".join([startup_name, sector, description, site, excel_context]))
        result["algo1_learning_match_label"] = learning_match.get("label", "")
        result["algo1_learning_match_score"] = learning_match.get("score", 0)
        result["algo1_learning_match_source"] = learning_match.get("source", "")

        if REQUEST_DELAY_SECONDS > 0:
            time.sleep(REQUEST_DELAY_SECONDS)

        return apply_profile_rules(result, " ".join([startup_name, sector, description, site, excel_context, website_text]))

    except Exception as exc:
        return apply_profile_rules(error_result(f"Erreur ligne pendant le traitement : {exc}"), "")


def process_excel_file(input_file: Path) -> None:
    output_file = make_output_path(input_file)
    temp_file = make_temp_path(input_file)

    log(f"Chargement : {input_file.name}")
    df = pd.read_excel(input_file)

    if LIMIT_ROWS > 0:
        original_count = len(df)
        df = df.head(LIMIT_ROWS).copy()
        log(f"Mode test : {len(df)}/{original_count} lignes seront traitées.")

    columns = detect_columns(df)

    log("Colonnes détectées :")
    log(f"- Startup : {columns['startup']}")
    log(f"- Secteur : {columns['sector']}")
    log(f"- Description : {columns['description']}")
    log(f"- Site web : {columns['site']}")
    log(f"- Colonnes totales du fichier : {len(df.columns)}")

    results = load_resume_results(temp_file, len(df))
    website_cache = {}
    cache_lock = threading.Lock()
    resume_start = len(results)
    rows = [row for _, row in df.iloc[resume_start:].iterrows()]

    log(f"Traitement parallèle : {MAX_WORKERS} worker(s)")
    log(f"Texte site max : {MAX_WEBSITE_TEXT_CHARS} caractères")
    log(f"Timeout scraping : {SCRAPE_TIMEOUT_SECONDS}s")
    log(f"Recherche site enrichie : {'activée' if ENABLE_WEBSITE_RESEARCH else 'désactivée'}")
    log(f"Pages site max : {RESEARCH_MAX_PAGES}")
    log(f"Contexte Excel utile max : {MAX_EXCEL_CONTEXT_CHARS} caractères par startup")
    if resume_start > 0:
        log(f"Reprise active : démarrage du traitement à la ligne {resume_start + 1}.")
    elif RESUME_FROM_TEMP:
        log("Reprise active : aucune sauvegarde exploitable, traitement depuis le début.")
    else:
        log("Reprise désactivée : traitement depuis le début.")

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        mapped_results = executor.map(
            lambda row: process_startup_row(row, columns, website_cache, cache_lock),
            rows,
        )

        for result in tqdm(
            mapped_results,
            total=len(df),
            initial=resume_start,
            desc=f"Préqualification {input_file.name}",
        ):
            results.append(result)

            if SAVE_EVERY > 0 and len(results) % SAVE_EVERY == 0:
                save_results(df, results, temp_file)
                log(f"Sauvegarde temporaire : {temp_file.name}")

    save_results(df, results, output_file)
    log(f"Fichier généré : {output_file.name}")


# ============================================================
# SCRIPT PRINCIPAL
# ============================================================

def main() -> None:
    base_dir = Path.cwd()
    excel_files = discover_excel_files(base_dir, sys.argv[1:])

    if not excel_files:
        log("Aucun fichier Excel trouvé dans le dossier.")
        log("Tu peux aussi lancer : python Algo1.py chemin\\vers\\fichier.xlsx")
        return

    log(f"Déploiement Azure OpenAI utilisé : {MODEL_DEPLOYMENT}")
    log(f"Version API Azure OpenAI : {AZURE_OPENAI_API_VERSION}")
    log(f"Workers parallèles : {MAX_WORKERS}")
    log(f"{len(excel_files)} fichier(s) Excel à traiter.")

    validate_azure_client()

    for input_file in excel_files:
        try:
            process_excel_file(input_file)
        except Exception as exc:
            log(f"Erreur sur {input_file.name} : {exc}")

    log("Traitement terminé.")


if __name__ == "__main__":
    main()
