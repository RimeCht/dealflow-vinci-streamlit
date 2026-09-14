import json
import os
import random
import re
import sys
import threading
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from difflib import SequenceMatcher
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openai import AzureOpenAI
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from tqdm import tqdm

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
ALGO3_API_MODE = os.getenv(
    "ALGO3_API_MODE",
    os.getenv("ALGO2_API_MODE", os.getenv("ALGO1_API_MODE", "chat")),
).strip().lower()

OUTPUT_SUFFIX = "_algo_3_orientation_programme.xlsx"
TEMP_SUFFIX = "_algo_3_orientation_programme_sauvegarde_temp.xlsx"
SAVE_EVERY = env_int("ALGO3_SAVE_EVERY", 50)
REQUEST_DELAY_SECONDS = env_float("ALGO3_REQUEST_DELAY_SECONDS", 0.1)
MAX_API_RETRIES = env_int("ALGO3_MAX_API_RETRIES", 3)
MAX_WORKERS = max(1, env_int("ALGO3_MAX_WORKERS", 2))
LIMIT_ROWS = env_int("ALGO3_LIMIT_ROWS", 0)
RESUME_FROM_TEMP = env_bool("ALGO3_RESUME_FROM_TEMP", True)
ENABLE_RESEARCH = env_bool("ALGO3_ENABLE_RESEARCH", True)
TARGETED_MATURITY_RESEARCH = env_bool("ALGO3_TARGETED_MATURITY_RESEARCH", True)
RESEARCH_MAX_PAGES = env_int("ALGO3_RESEARCH_MAX_PAGES", 8)
RESEARCH_MAX_LINKS_TO_TRY = env_int("ALGO3_RESEARCH_MAX_LINKS_TO_TRY", max(RESEARCH_MAX_PAGES * 3, RESEARCH_MAX_PAGES))
RESEARCH_TIMEOUT_SECONDS = env_int("ALGO3_RESEARCH_TIMEOUT_SECONDS", 8)
RESEARCH_MAX_TEXT_CHARS = env_int("ALGO3_RESEARCH_MAX_TEXT_CHARS", 9000)
RESEARCH_TEXT_PER_PAGE_CHARS = env_int("ALGO3_RESEARCH_TEXT_PER_PAGE_CHARS", 1800)
MAX_EXCEL_CONTEXT_CHARS = env_int("ALGO3_MAX_EXCEL_CONTEXT_CHARS", 18000)
MAX_EXCEL_CELL_CHARS = env_int("ALGO3_MAX_EXCEL_CELL_CHARS", 3200)
INCLUDE_EXCEL_COLUMN_LIST_IN_PROMPT = env_bool("ALGO3_INCLUDE_EXCEL_COLUMN_LIST_IN_PROMPT", False)
PAPPERS_API_TOKEN = os.getenv("PAPPERS_API_TOKEN", "").strip()
PAPPERS_API_URL = "https://api.pappers.fr/v2/entreprise"
PAPPERS_SEARCH_API_URL = "https://api.pappers.fr/v2/recherche"
ENABLE_RECHERCHE_ENTREPRISES_API = env_bool("ALGO3_ENABLE_RECHERCHE_ENTREPRISES_API", True)
RECHERCHE_ENTREPRISES_API_URL = "https://recherche-entreprises.api.gouv.fr/search"
RECHERCHE_ENTREPRISES_DELAY_SECONDS = env_float("ALGO3_RECHERCHE_ENTREPRISES_DELAY_SECONDS", 0.2)
ENABLE_SOCIETE_SEARCH = env_bool("ALGO3_ENABLE_SOCIETE_SEARCH", True)
SOCIETE_NAME_SEARCH = env_bool("ALGO3_SOCIETE_NAME_SEARCH", True)
SOCIETE_REQUEST_DELAY_SECONDS = env_float("ALGO3_SOCIETE_REQUEST_DELAY_SECONDS", 0.4)
SOCIETE_BASE_URL = "https://www.societe.com"
SOCIETE_SEARCH_URL = "https://www.societe.com/cgi-bin/search"
ENABLE_LINKEDIN_EMPLOYEE_SEARCH = env_bool("ALGO3_ENABLE_LINKEDIN_EMPLOYEE_SEARCH", True)
LINKEDIN_REQUEST_DELAY_SECONDS = env_float("ALGO3_LINKEDIN_REQUEST_DELAY_SECONDS", 0.6)
ELIGIBLE_DECISIONS = {
    item.strip().upper()
    for item in os.getenv("ALGO3_ELIGIBLE_DECISIONS", "PREQUALIFIEE").split(",")
    if item.strip()
}
NON_RETRYABLE_API_STATUS_CODES = {400, 401, 403, 404}

ORIENTATIONS = ["Seed", "Catalyst", "Matériaux", "Autre", "A_VERIFIER"]
SECONDARY_ORIENTATIONS = ["Seed", "Catalyst", "Matériaux", "Autre", "Aucune", "inconnu"]
SOLUTION_TYPES = ["saas", "hardware", "hybride", "service", "inconnu"]
MATURITY_STAGES = ["early", "intermediate", "mature", "inconnu"]
TRL_VALUES = ["<7", "7-8", "9+", "inconnu"]
MRL_VALUES = ["<3", "3-6", "7+", "inconnu"]
FUNDING_STAGES = ["pre-seed", "seed", "before_series_a", "series_a_or_later", "inconnu"]
YES_NO_UNKNOWN = ["oui", "non", "inconnu"]
PRODUCT_VALUES = ["oui", "non", "proche", "inconnu"]
CAPACITY_VALUES = ["oui", "non", "non_obligatoire", "inconnu"]
REFERENCE_VALUES = ["oui", "non", "souhaitable", "inconnu"]
COMPANY_CONFIDENCE_VALUES = ["High", "Medium", "Low", ""]

TRUSTED_EXTERNAL_SOURCE_DOMAINS = {
    "linkedin.com": "LinkedIn",
    "crunchbase.com": "Crunchbase",
    "dealroom.co": "Dealroom",
    "pappers.fr": "Pappers",
    "societe.com": "Société.com",
    "infogreffe.fr": "Infogreffe",
}

RESULT_COLUMNS = [
    "dealflow_profile",
    "profile_exclusion_matches",
    "algo3_orientation",
    "algo3_confidence",
    "secondary_orientation",
    "solution_type",
    "maturity_stage",
    "trl_estimate",
    "mrl_estimate",
    "employees_estimate",
    "revenue_arr_estimate",
    "funding_stage",
    "product_commercialized",
    "production_capacity",
    "international_deployment_capacity",
    "field_reference",
    "materials_vertical",
    "orientation_reason",
    "missing_information",
    "algo3_solution_type_keyword",
    "algo3_materials_keyword",
    "algo3_trl_signal_estimate",
    "algo3_trl_signal_evidence",
    "algo3_research_sources",
    "algo3_research_quality",
    "algo3_research_signals",
    "algo3_research_targeted_sources",
    "algo3_research_targeted_categories",
    "algo3_research_targeted_summary",
    "algo3_evidence_employees",
    "algo3_evidence_funding",
    "algo3_evidence_customers",
    "algo3_evidence_product",
    "algo3_evidence_deployment",
    "algo3_evidence_international",
    "algo3_evidence_production",
    "algo3_evidence_materials",
    "algo3_research_text_used",
    "algo3_maturity_score",
    "algo3_seed_score",
    "algo3_catalyst_score",
    "algo3_materials_score",
    "nombre_employes",
    "source_nombre_employes",
    "statut_entreprise",
    "source_statut",
    "niveau_confiance",
    "commentaire_verification",
    "algo3_excel_columns_used",
    "algo3_excel_columns_used_count",
    "algo3_excel_columns_available_count",
    "algo3_excel_context_truncated",
    "algo3_orientation_finale",
    "algo3_regle_finale",
]

THREAD_LOCAL = threading.local()

MATURITY_RESEARCH_LINK_CATEGORIES = {
    "employees_team": {
        "label": "employees/team",
        "weight": 55,
        "terms": [
            "team",
            "about",
            "company",
            "careers",
            "jobs",
            "people",
            "employees",
            "collaborateurs",
            "equipe",
            "recrutement",
        ],
    },
    "funding_investors": {
        "label": "funding/investors",
        "weight": 65,
        "terms": [
            "funding",
            "investors",
            "investment",
            "raised",
            "series a",
            "seed round",
            "pre seed",
            "press",
            "news",
            "levee de fonds",
            "investisseurs",
            "financement",
        ],
    },
    "customers_references": {
        "label": "customers/references",
        "weight": 75,
        "terms": [
            "customers",
            "clients",
            "references",
            "trusted by",
            "customer story",
            "success story",
            "case study",
            "case studies",
            "cas client",
            "cas clients",
            "etude de cas",
            "etudes de cas",
        ],
    },
    "product_technology": {
        "label": "product/technology",
        "weight": 60,
        "terms": [
            "product",
            "products",
            "solution",
            "solutions",
            "platform",
            "technology",
            "technologie",
            "software",
            "hardware",
            "features",
            "prototype",
            "pilot",
            "commercially available",
        ],
    },
    "deployments_pilots": {
        "label": "deployments/pilots",
        "weight": 70,
        "terms": [
            "deployment",
            "deployments",
            "deployed",
            "pilot",
            "pilots",
            "field trial",
            "on site",
            "in production",
            "rollout",
            "implementation",
            "deploiement",
            "pilote",
            "terrain",
        ],
    },
    "production_industrial": {
        "label": "production/industrial",
        "weight": 65,
        "terms": [
            "manufacturing",
            "factory",
            "production",
            "industrialization",
            "industrialisation",
            "plant",
            "scale production",
            "supply chain",
            "capacity",
            "usine",
            "capacite de production",
        ],
    },
    "international_scale": {
        "label": "international/scale",
        "weight": 45,
        "terms": [
            "international",
            "global",
            "worldwide",
            "europe",
            "countries",
            "markets",
            "export",
            "pays",
            "monde",
            "deploiement international",
        ],
    },
    "materials_vertical": {
        "label": "materials",
        "weight": 80,
        "terms": [
            "materials",
            "material",
            "concrete",
            "cement",
            "steel",
            "wood",
            "composite",
            "polymer",
            "insulation",
            "coating",
            "matÃ©riaux",
            "materiaux",
            "beton",
            "ciment",
            "acier",
            "bois",
            "biosource",
            "bas carbone",
        ],
    },
}

COMMON_MATURITY_PATHS = [
    "/about",
    "/team",
    "/company",
    "/careers",
    "/jobs",
    "/customers",
    "/clients",
    "/references",
    "/case-studies",
    "/case-study",
    "/customer-stories",
    "/success-stories",
    "/product",
    "/products",
    "/technology",
    "/solution",
    "/solutions",
    "/platform",
    "/deployments",
    "/pilots",
    "/press",
    "/news",
    "/investors",
    "/funding",
    "/manufacturing",
    "/production",
    "/materials",
    "/material",
    "/fr/a-propos",
    "/fr/equipe",
    "/fr/recrutement",
    "/fr/carrieres",
    "/fr/clients",
    "/fr/references",
    "/fr/cas-clients",
    "/fr/etudes-de-cas",
    "/fr/produits",
    "/fr/technologie",
    "/fr/solutions",
    "/fr/deploiements",
    "/fr/presse",
    "/fr/actualites",
    "/fr/investisseurs",
    "/fr/production",
    "/fr/materiaux",
]

LOW_VALUE_RESEARCH_LINK_TERMS = [
    "privacy",
    "terms",
    "legal",
    "cookies",
    "login",
    "sign in",
    "signup",
    "contact",
    "demo",
    "support",
    "documentation",
    "docs",
    "events",
]


MATERIALS_KEYWORDS = [
    "material",
    "materials",
    "matériau",
    "materiaux",
    "matériaux",
    "béton",
    "beton",
    "concrete",
    "cement",
    "ciment",
    "acier",
    "steel",
    "bois",
    "wood",
    "composite",
    "polymer",
    "polymère",
    "polymere",
    "biosourcé",
    "biosource",
    "bio based",
    "low carbon material",
    "low carbon concrete",
    "bas carbone",
    "recycled material",
    "circular material",
    "recyclage matériaux",
    "recyclage materiaux",
    "construction material",
    "insulation",
    "isolation",
    "coating",
    "revêtement",
    "revetement",
]

HARDWARE_KEYWORDS = [
    "hardware",
    "robot",
    "robotics",
    "device",
    "sensor",
    "capteur",
    "drone",
    "machine",
    "equipment",
    "équipement",
    "equipement",
    "battery",
    "camera",
    "iot",
    "edge device",
    "industrial system",
]

SAAS_KEYWORDS = [
    "saas",
    "software",
    "platform",
    "plateforme",
    "app",
    "application",
    "dashboard",
    "cloud",
    "analytics",
    "ai platform",
    "data platform",
    "workflow",
    "api",
    "digital solution",
    "logiciel",
]

SERVICE_KEYWORDS = [
    "service",
    "consulting",
    "conseil",
    "marketplace",
    "bureau d'études",
    "bureau d'etudes",
    "engineering services",
]

TRL_SIGNAL_KEYWORDS = {
    "<7": [
        # Early research / concept
        "research project",
        "basic research",
        "academic research",
        "scientific research",
        "concept",
        "technology concept",
        "early-stage technology",
        "early stage technology",
        "emerging technology",
        "experimental technology",

        # Proof of concept
        "proof of concept",
        "proof-of-concept",
        "poc",
        "feasibility study",
        "feasibility phase",
        "technical feasibility",

        # Laboratory validation
        "laboratory",
        "lab",
        "lab validation",
        "validated in lab",
        "laboratory validation",
        "prototype in laboratory",
        "lab-scale prototype",
        "early prototype",

        # Pre-market
        "pre-commercial",
        "pre commercial",
        "pre-market",
        "not yet commercialized",
        "under development",
        "in development",
        "r&d project",
        "research and development",

        # French
        "recherche",
        "recherche fondamentale",
        "projet de recherche",
        "concept technologique",
        "technologie émergente",
        "preuve de concept",
        "poc",
        "étude de faisabilité",
        "etude de faisabilite",
        "validation en laboratoire",
        "prototype laboratoire",
        "prototype amont",
        "prototype précoce",
        "prototype precoce",
        "en développement",
        "en developpement",
        "non commercialisé",
        "non commercialise",
    ],

    "7-8": [
        # Prototype / demonstrator
        "prototype demonstration",
        "system prototype",
        "demonstrator",
        "technology demonstrator",
        "demonstration in relevant environment",
        "relevant environment",
        "operational environment",

        # Tests / pilots
        "field trial",
        "field test",
        "pilot project",
        "pilot",
        "industrial pilot",
        "pilot deployment",
        "beta",
        "mvp",
        "minimum viable product",

        # Real conditions
        "tested on site",
        "tested in real conditions",
        "tested in the field",
        "on-site test",
        "onsite test",
        "in-situ test",
        "real-world test",
        "real life conditions",
        "trial with customers",
        "pilot with customers",
        "first deployment",
        "initial deployment",

        # Pre-commercial but advanced
        "pre-series",
        "pre series",
        "limited deployment",
        "early adopters",
        "validation with partners",
        "validated with partners",

        # French
        "prototype démontré",
        "prototype demontre",
        "démonstrateur",
        "demonstrateur",
        "pilote terrain",
        "projet pilote",
        "pilote industriel",
        "déploiement pilote",
        "deploiement pilote",
        "test terrain",
        "testé sur site",
        "teste sur site",
        "test en conditions réelles",
        "test en conditions reelles",
        "conditions réelles",
        "conditions reelles",
        "environnement réel",
        "environnement reel",
        "validation terrain",
        "validation avec partenaires",
        "premier déploiement",
        "premier deploiement",
        "déploiement limité",
        "deploiement limite",
    ],

    "9+": [
        # Commercial maturity
        "commercially available",
        "available now",
        "available on the market",
        "market-ready",
        "market ready",
        "commercial product",
        "commercial solution",
        "production-ready",
        "production ready",
        "in production",
        "industrialized",
        "scalable solution",

        # Deployment
        "deployed",
        "deployed at scale",
        "large-scale deployment",
        "rolled out",
        "implemented",
        "installed",
        "installations",
        "operating in production",
        "used in production",
        "live deployment",

        # Customers / references
        "used by",
        "trusted by",
        "customers",
        "clients",
        "paying customers",
        "enterprise customers",
        "customer base",
        "case study",
        "case studies",
        "customer story",
        "success story",
        "reference customer",
        "references",

        # Business maturity
        "launched",
        "sold to",
        "sales",
        "revenue",
        "recurring revenue",
        "annual recurring revenue",
        "arr",
        "contracts",
        "commercial contracts",
        "partnership with",
        "supplier of",
        "vendor",
        "platform used by",

        # French
        "commercialisé",
        "commercialise",
        "commercialisée",
        "commercialisee",
        "disponible sur le marché",
        "disponible sur le marche",
        "solution commerciale",
        "produit commercial",
        "prêt pour le marché",
        "pret pour le marche",
        "déployé",
        "deploye",
        "déployée",
        "deployee",
        "déployé à grande échelle",
        "deploye a grande echelle",
        "en production",
        "industrialisé",
        "industrialise",
        "installé",
        "installe",
        "installations",
        "utilisé par",
        "utilise par",
        "clients",
        "clients payants",
        "références clients",
        "references clients",
        "cas client",
        "cas d’usage client",
        "cas d'usage client",
        "contrats commerciaux",
        "chiffre d’affaires",
        "chiffre d'affaires",
        "revenus récurrents",
        "revenus recurrents",
    ],
}

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
    ],
    "sector": [
        "secteur",
        "sector",
        "domaine",
        "industry",
        "category",
        "categorie",
        "catégorie",
        "target_sector",
        "target sector",
        "confirmed_sector",
        "confirmed sector",
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
        "short_description_clean",
        "short description clean",
        "description courte",
    ],
    "site": [
        "site web",
        "website",
        "site",
        "url",
        "web",
        "official website",
        "website url",
        "site internet",
        "lien",
        "link",
    ],
    "algo2_decision_finale": [
        "algo2_decision_finale",
        "algo2 decision finale",
        "algo 2 decision finale",
        "decision finale algo 2",
        "algo2_decision",
        "algo2 decision",
    ],
    "algo2_raw_decision": [
        "algo2_decision",
        "algo2 decision",
        "algo2_decision_finale",
    ],
    "algo1_decision": [
        "decision_finale",
        "decision finale",
        "decision",
        "algo1 decision",
        "algo 1 decision",
    ],
    "algo1_reason": [
        "algo1_reason",
        "algo1 reason",
        "algo 1 reason",
        "reason",
        "raison",
        "reason.1",
        "sector_reason",
        "sector reason",
        "regle_finale",
        "regle finale",
    ],
    "algo2_reason": [
        "algo2_reason",
        "algo2 reason",
        "algo 2 reason",
        "orientation_reason",
        "reason.1",
        "reason",
        "algo2_regle_finale",
        "algo2 regle finale",
        "issue_alignment_reason",
        "sector_alignment_reason",
        "b2b_reason",
    ],
    "b2b_strength": [
        "b2b_strength",
        "b2b strength",
        "is_b2b",
        "is b2b",
    ],
    "confirmed_sector": [
        "confirmed_sector",
        "confirmed sector",
        "target_sector",
        "target sector",
        "secteur",
        "sector",
    ],
    "main_issue_matrix": [
        "main_issue_matrix",
        "main issue matrix",
        "matched_issue_matrices",
        "matched issue matrices",
    ],
    "matched_issue_matrices": [
        "matched_issue_matrices",
        "matched issue matrices",
        "algo2_issue_keyword_categories",
        "algo2 issue keyword categories",
    ],
    "issue_alignment_reason": [
        "issue_alignment_reason",
        "issue alignment reason",
        "algo2_issue_keyword_matches",
        "algo2 issue keyword matches",
    ],
    "sector_alignment_reason": [
        "sector_alignment_reason",
        "sector alignment reason",
        "algo2_sector_keyword_matches",
        "algo2 sector keyword matches",
    ],
    "nombre_employes": [
        "nombre_employes",
        "nombre employés",
        "nombre employes",
        "employees",
        "employee count",
        "effectif",
        "effectifs",
    ],
    "source_nombre_employes": [
        "source_nombre_employes",
        "source nombre employés",
        "source nombre employes",
        "employee source",
        "source effectif",
    ],
    "statut_entreprise": [
        "statut_entreprise",
        "statut entreprise",
        "company status",
        "etat entreprise",
        "état entreprise",
    ],
    "source_statut": [
        "source_statut",
        "source statut",
        "status source",
        "source état",
        "source etat",
    ],
    "niveau_confiance": [
        "niveau_confiance",
        "niveau confiance",
        "confidence level",
        "confiance",
    ],
    "commentaire_verification": [
        "commentaire_verification",
        "commentaire verification",
        "commentaire vérification",
        "verification comment",
        "comment",
    ],
    "siren_siret": [
        "siren",
        "siret",
        "siren siret",
        "siren/siret",
        "numero siren",
        "numéro siren",
        "numero siret",
        "numéro siret",
        "identifiant entreprise",
        "company registration number",
        "registration number",
    ],
}

EXCEL_CONTEXT_PRIORITY_TERMS = [
    "algo2 decision finale",
    "algo2 decision",
    "b2b strength",
    "confirmed sector",
    "main issue matrix",
    "matched issue matrices",
    "issue alignment reason",
    "sector alignment reason",
    "decision finale",
    "is b2b",
    "target sector",
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
    "startup development level",
    "primary target market",
    "purpose vision",
    "roadmap",
    "trl",
    "mrl",
    "employee",
    "revenue",
    "arr",
    "funding",
    "patents",
    "worked with any player",
    "programs before",
    "awards",
    "additional information",
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
    "algo1 research text used",
    "algo2 research text used",
    "algo1 excel columns used",
    "algo1 excel columns used count",
    "algo1 excel columns available count",
    "algo1 excel context truncated",
    "algo2 excel columns used",
    "algo2 excel columns used count",
    "algo2 excel columns available count",
    "algo2 excel context truncated",
}

EXCEL_CONTEXT_EXCLUDED_PREFIXES = (
    "algo3 ",
)


# ============================================================
# OUTILS
# ============================================================

def log(message: str) -> None:
    print(f"[Algo 3] {message}")


def strip_accents(text: str) -> str:
    return "".join(
        char
        for char in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(char)
    )


def normalize_text(text: str) -> str:
    text = strip_accents(str(text)).lower()
    text = re.sub(r"[_\-/]+", " ", text)
    text = re.sub(r"[^a-z0-9 +#.&'<>=€$]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_column_name(col: str) -> str:
    return normalize_text(col)


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
                df[original].fillna("").astype(str).str.strip().ne("").sum()
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


def clean_text(value) -> str:
    if pd.isna(value):
        return ""

    value = str(value)
    value = ILLEGAL_CHARACTERS_RE.sub("", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def should_exclude_excel_context_column(column: str) -> bool:
    normalized = normalize_column_name(column)
    if not normalized:
        return True
    if normalized in EXCEL_CONTEXT_EXCLUDED_EXACT:
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
                lines.append(line[:remaining].rstrip() + " [contexte tronqué]")
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


def clean_decision(value) -> str:
    return normalize_text(value).replace(" ", "_").upper()


def row_value(row: pd.Series, column: str | None) -> str:
    if not column:
        return ""
    return clean_text(row.get(column, ""))


def first_row_value(row: pd.Series, columns: list[str | None]) -> str:
    for column in columns:
        value = row_value(row, column)
        if value:
            return value
    return ""


def compact_join(values: list[str]) -> str:
    cleaned = [clean_text(value) for value in values if clean_text(value)]
    return " | ".join(cleaned)


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


def contains_any(text: str, keywords: list[str]) -> bool:
    normalized = normalize_text(text)
    return any(normalize_text(keyword) in normalized for keyword in keywords)


def infer_solution_type(text: str) -> str:
    has_hardware = contains_any(text, HARDWARE_KEYWORDS)
    has_saas = contains_any(text, SAAS_KEYWORDS)
    has_service = contains_any(text, SERVICE_KEYWORDS)

    if has_hardware and has_saas:
        return "hybride"
    if has_hardware:
        return "hardware"
    if has_saas:
        return "saas"
    if has_service:
        return "service"
    return "inconnu"


def infer_materials_vertical(text: str) -> str:
    return "oui" if contains_any(text, MATERIALS_KEYWORDS) else "non"


def first_keyword_evidence(text: str, keywords: list[str], limit: int = 8) -> list[str]:
    normalized = normalize_text(text)
    evidence = []

    for keyword in keywords:
        normalized_keyword = normalize_text(keyword)
        if normalized_keyword and normalized_keyword in normalized and keyword not in evidence:
            evidence.append(keyword)
        if len(evidence) >= limit:
            break

    return evidence


def extract_explicit_trl_numbers(text: str) -> list[int]:
    numbers = []
    patterns = [
        r"\bTRL\s*[:=]?\s*([1-9])\+?\b",
        r"\btechnology readiness level\s*[:=]?\s*([1-9])\+?\b",
        r"\bniveau\s+de\s+maturit[ée]\s+technologique\s*[:=]?\s*([1-9])\+?\b",
    ]

    for pattern in patterns:
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            try:
                value = int(match.group(1))
            except (TypeError, ValueError):
                continue
            if 1 <= value <= 9 and value not in numbers:
                numbers.append(value)

    return numbers


def trl_bucket_from_number(value: int) -> str:
    if value >= 9:
        return "9+"
    if value >= 7:
        return "7-8"
    if value >= 1:
        return "<7"
    return "inconnu"


def infer_trl_from_text(text: str) -> dict[str, str]:
    text = clean_text(text)
    if not text:
        return {
            "estimate": "inconnu",
            "evidence": "Aucune preuve TRL détectée.",
        }

    explicit_numbers = extract_explicit_trl_numbers(text)
    if explicit_numbers:
        strongest = max(explicit_numbers)
        return {
            "estimate": trl_bucket_from_number(strongest),
            "evidence": "TRL explicite détecté : " + ", ".join(f"TRL {value}" for value in explicit_numbers),
        }

    evidence_by_level = {
        level: first_keyword_evidence(text, keywords)
        for level, keywords in TRL_SIGNAL_KEYWORDS.items()
    }

    if len(evidence_by_level["9+"]) >= 2 or any(
        normalize_text(term) in normalize_text(text)
        for term in [
            "commercially available",
            "in production",
            "deployed at scale",
            "commercialisé",
            "commercialisee",
            "en production",
        ]
    ):
        return {
            "estimate": "9+",
            "evidence": "Preuves TRL 9+ : " + ", ".join(evidence_by_level["9+"][:8]),
        }

    if evidence_by_level["7-8"]:
        return {
            "estimate": "7-8",
            "evidence": "Preuves TRL 7-8 : " + ", ".join(evidence_by_level["7-8"][:8]),
        }

    if evidence_by_level["<7"]:
        return {
            "estimate": "<7",
            "evidence": "Preuves TRL <7 : " + ", ".join(evidence_by_level["<7"][:8]),
        }

    return {
        "estimate": "inconnu",
        "evidence": "Aucune preuve TRL suffisamment claire détectée.",
    }


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


def same_domain(url: str, base_url: str) -> bool:
    parsed_url = urlparse(url)
    parsed_base = urlparse(base_url)
    url_host = parsed_url.netloc.lower().removeprefix("www.")
    base_host = parsed_base.netloc.lower().removeprefix("www.")
    return bool(url_host and url_host == base_host)


def trusted_external_source_name(url: str) -> str:
    host = urlparse(url).netloc.lower().removeprefix("www.")
    for domain, source_name in TRUSTED_EXTERNAL_SOURCE_DOMAINS.items():
        if host == domain or host.endswith(f".{domain}"):
            return source_name
    return ""


def link_url(link) -> str:
    if isinstance(link, dict):
        return clean_text(link.get("url", ""))
    return clean_text(link)


def link_text(link) -> str:
    if isinstance(link, dict):
        return clean_text(link.get("text", ""))
    return ""


def collect_trusted_external_links(links: list, limit: int = 12) -> list[str]:
    trusted_links = []
    seen = set()

    for link in links:
        link = link_url(link)
        source_name = trusted_external_source_name(link)
        if not source_name:
            continue

        parsed = urlparse(link)
        clean_link = parsed._replace(query="", fragment="").geturl().rstrip("/")
        key = normalize_text(clean_link)
        if not clean_link or key in seen:
            continue

        seen.add(key)
        trusted_links.append(clean_link)
        if len(trusted_links) >= limit:
            break

    return trusted_links


def get_thread_http_session() -> requests.Session:
    session = getattr(THREAD_LOCAL, "http_session", None)
    if session is None:
        session = requests.Session()
        THREAD_LOCAL.http_session = session
    return session


def fetch_page_text_and_links(url: str) -> tuple[str, list[dict[str, str]]]:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (compatible; LeonardAlgo3Research/1.0; "
            "startup orientation research)"
        )
    }

    try:
        response = get_thread_http_session().get(
            url,
            headers=headers,
            timeout=RESEARCH_TIMEOUT_SECONDS,
        )
        response.raise_for_status()
        response.encoding = response.apparent_encoding or response.encoding
    except requests.RequestException:
        return "", []

    content_type = response.headers.get("content-type", "")
    if "text/html" not in content_type and "application/xhtml" not in content_type:
        return "", []

    soup = BeautifulSoup(response.text, "html.parser")

    links = []
    for anchor in soup.find_all("a", href=True):
        href = anchor.get("href", "")
        absolute_url = urljoin(url, href.split("#")[0])
        if absolute_url.startswith(("http://", "https://")):
            links.append({
                "url": absolute_url,
                "text": clean_text(anchor.get_text(" ", strip=True)),
            })

    for tag in soup(["script", "style", "noscript", "svg", "footer", "nav"]):
        tag.decompose()

    title = soup.title.get_text(" ", strip=True) if soup.title else ""

    meta_description = ""
    meta = soup.find("meta", attrs={"name": "description"})
    if meta and meta.get("content"):
        meta_description = meta.get("content")

    headings = " ".join(
        heading.get_text(" ", strip=True)
        for heading in soup.find_all(["h1", "h2", "h3"])[:20]
    )

    paragraphs = " ".join(
        paragraph.get_text(" ", strip=True)
        for paragraph in soup.find_all(["p", "li"])[:80]
    )

    text = clean_text(f"{title}. {meta_description}. {headings}. {paragraphs}")

    return text, links


def targeted_maturity_categories_from_text(text: str) -> list[str]:
    normalized = normalize_text(text)
    categories = []
    for category, config in MATURITY_RESEARCH_LINK_CATEGORIES.items():
        for term in config["terms"]:
            if normalize_text(term) in normalized:
                categories.append(category)
                break
    return categories


def format_maturity_category_labels(categories: list[str]) -> str:
    labels = []
    for category in categories:
        if category == "homepage":
            label = "homepage"
        else:
            label = MATURITY_RESEARCH_LINK_CATEGORIES.get(category, {}).get("label", category)
        if label not in labels:
            labels.append(label)
    return ", ".join(labels) if labels else "non classe"


def unique_values(values: list[str]) -> list[str]:
    unique = []
    for value in values:
        value = clean_text(value)
        if value and value not in unique:
            unique.append(value)
    return unique


def research_link_score(url: str, text: str = "") -> int:
    normalized = normalize_text(f"{url} {text}")
    priority_terms = {
        "about": 45,
        "team": 35,
        "careers": 35,
        "jobs": 35,
        "customers": 45,
        "clients": 45,
        "case studies": 50,
        "case study": 50,
        "success stories": 50,
        "press": 30,
        "news": 25,
        "blog": 15,
        "product": 35,
        "technology": 35,
        "solution": 30,
        "pricing": 20,
        "funding": 40,
        "investors": 40,
        "materials": 45,
        "material": 45,
        "sustainability": 30,
        "decarbon": 30,
        "international": 30,
        "partners": 25,
    }

    score = 0
    if TARGETED_MATURITY_RESEARCH:
        for category, config in MATURITY_RESEARCH_LINK_CATEGORIES.items():
            if category in targeted_maturity_categories_from_text(normalized):
                score += int(config["weight"])

    for term, weight in priority_terms.items():
        if term in normalized:
            score += weight

    for term in LOW_VALUE_RESEARCH_LINK_TERMS:
        if normalize_text(term) in normalized:
            score -= 45

    return score


def make_maturity_link(url: str, text: str = "", source: str = "site") -> dict:
    clean_url = normalize_url(url).rstrip("/")
    categories = targeted_maturity_categories_from_text(f"{clean_url} {text}")
    score = research_link_score(clean_url, text)
    if source == "homepage_link":
        score += 30
    elif source == "common_path":
        score -= 10
    return {
        "url": clean_url,
        "text": clean_text(text),
        "source": source,
        "categories": categories,
        "score": score,
    }


def common_maturity_candidates(base_url: str) -> list[dict]:
    parsed = urlparse(normalize_url(base_url))
    if not parsed.scheme or not parsed.netloc:
        return []
    root_url = f"{parsed.scheme}://{parsed.netloc}"
    return [
        make_maturity_link(urljoin(root_url, path), path, source="common_path")
        for path in COMMON_MATURITY_PATHS
    ]


def normalized_maturity_link_record(link) -> dict | None:
    raw_url = link_url(link)
    if not raw_url:
        return None
    parsed = urlparse(normalize_url(raw_url))
    clean_url = parsed._replace(query="", fragment="").geturl().rstrip("/")
    if not clean_url:
        return None
    return make_maturity_link(clean_url, link_text(link), source="homepage_link")


def select_research_links(base_url: str, links: list) -> list[dict]:
    base_url = normalize_url(base_url).rstrip("/")
    selected = [{
        "url": base_url,
        "text": "homepage",
        "source": "homepage",
        "categories": ["homepage"],
        "score": 0,
    }]
    seen = {base_url}
    candidates_by_url = {}

    for link in links:
        record = normalized_maturity_link_record(link)
        if not record:
            continue
        clean_link = record["url"]
        if not same_domain(clean_link, base_url):
            continue
        if clean_link in seen:
            continue
        if record["score"] <= 0:
            continue
        current = candidates_by_url.get(clean_link)
        if current is None or record["score"] > current["score"]:
            candidates_by_url[clean_link] = record

    for record in common_maturity_candidates(base_url):
        clean_link = record["url"]
        if clean_link in seen or not same_domain(clean_link, base_url):
            continue
        if record["score"] <= 0:
            continue
        current = candidates_by_url.get(clean_link)
        if current is None or record["score"] > current["score"]:
            candidates_by_url[clean_link] = record

    candidates = list(candidates_by_url.values())
    candidates.sort(
        key=lambda item: (
            1 if item.get("source") == "homepage_link" else 0,
            item.get("score", 0),
            len([category for category in item.get("categories", []) if category != "homepage"]),
        ),
        reverse=True,
    )

    if TARGETED_MATURITY_RESEARCH:
        covered_categories = set()
        for candidate in candidates:
            if len(selected) >= RESEARCH_MAX_LINKS_TO_TRY:
                break
            new_categories = [
                category
                for category in candidate.get("categories", [])
                if category != "homepage" and category not in covered_categories
            ]
            if not new_categories:
                continue
            selected.append(candidate)
            seen.add(candidate["url"])
            covered_categories.update(new_categories)

    for candidate in candidates:
        if len(selected) >= RESEARCH_MAX_LINKS_TO_TRY:
            break
        if candidate["url"] not in seen:
            selected.append(candidate)
            seen.add(candidate["url"])

    return selected


def first_regex_matches(text: str, patterns: list[str], limit: int = 5) -> list[str]:
    matches = []
    for pattern in patterns:
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            value = clean_text(match.group(0))
            if value and value not in matches:
                matches.append(value)
            if len(matches) >= limit:
                return matches
    return matches


def detect_research_signals(text: str) -> dict[str, str]:
    normalized = normalize_text(text)
    trl_inference = infer_trl_from_text(text)

    employees = first_regex_matches(text, [
        r"\b\d{1,5}\+?\s+(?:employees|employés|employes|collaborateurs|people|team members)\b",
        r"\bteam of\s+\d{1,5}\b",
        r"\b(?:équipe|equipe)\s+de\s+\d{1,5}\b",
    ])

    trl = first_regex_matches(text, [
        r"\bTRL\s*[0-9]\+?\b",
        r"\btechnology readiness level\s*[0-9]\+?\b",
    ])

    mrl = first_regex_matches(text, [
        r"\bMRL\s*[0-9]\+?\b",
        r"\bmanufacturing readiness level\s*[0-9]\+?\b",
    ])

    revenue = first_regex_matches(text, [
        r"\bARR\s*(?:of|de|:)?\s*[€$]?\s*\d+(?:[.,]\d+)?\s*(?:k|m|million|millions|m€|m\$)?\b",
        r"\b(?:revenue|chiffre d'affaires|ca)\s*(?:of|de|:)?\s*[€$]?\s*\d+(?:[.,]\d+)?\s*(?:k|m|million|millions|m€|m\$)?\b",
        r"\b[€$]\s*\d+(?:[.,]\d+)?\s*(?:k|m|million|millions)\s+(?:arr|revenue)\b",
    ])

    funding = first_regex_matches(text, [
        r"\bpre[- ]seed\b",
        r"\bseed round\b",
        r"\bseries\s+a\b",
        r"\bserie\s+a\b",
        r"\bsérie\s+a\b",
        r"\bseries\s+b\b",
        r"\braised\s+[€$]?\s*\d+(?:[.,]\d+)?\s*(?:k|m|million|millions|m€|m\$)?\b",
        r"\blevée\s+de\s+fonds\s+de\s+\d+(?:[.,]\d+)?\s*(?:k|m|million|millions|m€)?\b",
    ])

    commercial = []
    if any(term in normalized for term in [
        "customers",
        "clients",
        "trusted by",
        "case study",
        "case studies",
        "deployed",
        "commercially available",
        "available now",
        "in production",
        "client references",
    ]):
        commercial.append("Signaux produit commercialisé / clients détectés")

    field_reference = []
    if any(term in normalized for term in [
        "case study",
        "case studies",
        "pilot project",
        "field trial",
        "deployed on site",
        "customer story",
        "success story",
        "trusted by",
        "référence",
        "reference terrain",
    ]):
        field_reference.append("Signaux référence terrain / cas client détectés")

    production = []
    if any(term in normalized for term in [
        "manufacturing",
        "factory",
        "production capacity",
        "scale production",
        "industrialization",
        "industrialisation",
        "capacity to produce",
        "plant",
        "usine",
    ]):
        production.append("Signaux capacité de production détectés")

    international = []
    if any(term in normalized for term in [
        "international",
        "global",
        "worldwide",
        "across europe",
        "in europe",
        "countries",
        "pays",
        "déploiement international",
        "deploiement international",
    ]):
        international.append("Signaux déploiement international détectés")

    materials = []
    if infer_materials_vertical(text) == "oui":
        materials.append("Signaux verticale Matériaux détectés")

    signals = {
        "employees": "; ".join(employees) or "inconnu",
        "trl": "; ".join(trl) or "inconnu",
        "trl_inferred": trl_inference["estimate"],
        "trl_evidence": trl_inference["evidence"],
        "mrl": "; ".join(mrl) or "inconnu",
        "revenue_arr": "; ".join(revenue) or "inconnu",
        "funding": "; ".join(funding) or "inconnu",
        "commercialization": "; ".join(commercial) or "inconnu",
        "field_reference": "; ".join(field_reference) or "inconnu",
        "production_capacity": "; ".join(production) or "inconnu",
        "international_deployment": "; ".join(international) or "inconnu",
        "materials": "; ".join(materials) or "inconnu",
    }

    return signals


def format_research_signals(signals: dict[str, str]) -> str:
    return " | ".join(f"{key}: {value}" for key, value in signals.items())


def truncate_research_page_text(text: str) -> str:
    text = clean_text(text)
    if RESEARCH_TEXT_PER_PAGE_CHARS <= 0:
        return text
    if len(text) <= RESEARCH_TEXT_PER_PAGE_CHARS:
        return text
    return text[:RESEARCH_TEXT_PER_PAGE_CHARS].rstrip() + " [page tronquee]"


def split_research_sentences(text: str) -> list[str]:
    text = clean_text(text)
    if not text:
        return []
    parts = re.split(r"(?<=[.!?])\s+|\s+\|\s+", text)
    if len(parts) <= 2 and len(text) > 500:
        parts = [text[index:index + 260] for index in range(0, len(text), 260)]
    return [clean_text(part) for part in parts if clean_text(part)]


def terms_for_maturity_categories(categories: list[str]) -> list[str]:
    terms = []
    for category in categories:
        config = MATURITY_RESEARCH_LINK_CATEGORIES.get(category)
        if config:
            terms.extend(config["terms"])
    return unique_values(terms)


def extract_evidence_snippets(text: str, terms: list[str], limit: int = 4) -> list[str]:
    normalized_terms = [normalize_text(term) for term in terms if normalize_text(term)]
    if not normalized_terms:
        return []

    snippets = []
    for sentence in split_research_sentences(text):
        normalized_sentence = normalize_text(sentence)
        if not any(term in normalized_sentence for term in normalized_terms):
            continue
        snippet = sentence[:280].rstrip()
        if len(sentence) > 280:
            snippet += "..."
        if snippet not in snippets:
            snippets.append(snippet)
        if len(snippets) >= limit:
            break

    return snippets


def format_maturity_page_records(page_records: list[dict]) -> str:
    parts = []
    for record in page_records:
        categories = [category for category in record.get("categories", []) if category != "homepage"]
        if not categories:
            continue
        parts.append(f"{format_maturity_category_labels(categories)}: {record.get('url', '')}")
    return " | ".join(unique_values(parts))


def build_maturity_research_evidence(page_records: list[dict]) -> dict:
    groups = {
        "employees": ["employees_team"],
        "funding": ["funding_investors"],
        "customers": ["customers_references"],
        "product": ["product_technology"],
        "deployment": ["deployments_pilots", "customers_references"],
        "international": ["international_scale"],
        "production": ["production_industrial"],
        "materials": ["materials_vertical"],
    }
    evidence = {key: [] for key in groups}

    for record in page_records:
        text = record.get("text", "")
        categories = record.get("categories", [])
        url = record.get("url", "")
        for group_name, group_categories in groups.items():
            if not any(category in categories for category in group_categories):
                continue
            terms = terms_for_maturity_categories(group_categories)
            snippets = extract_evidence_snippets(text, terms, limit=2)
            if snippets:
                for snippet in snippets:
                    evidence[group_name].append(f"{url}: {snippet}")
            else:
                evidence[group_name].append(f"{url}: page ciblee detectee, contenu a verifier")

    return {
        key: " | ".join(unique_values(values)[:6])
        for key, values in evidence.items()
    }


def build_maturity_research_summary(page_records: list[dict], evidence: dict) -> str:
    categories = []
    for record in page_records:
        categories.extend([category for category in record.get("categories", []) if category != "homepage"])
    parts = [
        f"Pages ciblees maturite: {format_maturity_page_records(page_records) or 'aucune page ciblee accessible'}",
        f"Categories ciblees: {format_maturity_category_labels(unique_values(categories)) if categories else 'aucune'}",
        f"Preuves effectif/team: {evidence.get('employees') or 'aucune'}",
        f"Preuves funding/investors: {evidence.get('funding') or 'aucune'}",
        f"Preuves clients/references: {evidence.get('customers') or 'aucune'}",
        f"Preuves produit/technology: {evidence.get('product') or 'aucune'}",
        f"Preuves deployment/pilots: {evidence.get('deployment') or 'aucune'}",
        f"Preuves international/scale: {evidence.get('international') or 'aucune'}",
        f"Preuves production/industrial: {evidence.get('production') or 'aucune'}",
        f"Preuves materials: {evidence.get('materials') or 'aucune'}",
    ]
    return " || ".join(parts)


def research_quality_from_sources(sources: list[str], text: str, categories: list[str]) -> str:
    if not sources:
        return "inaccessible"
    unique_categories = {category for category in categories if category and category != "homepage"}
    if len(unique_categories) >= 4 and len(text) >= 3000:
        return "forte"
    if len(unique_categories) >= 2 and len(text) >= 1500:
        return "moyenne"
    if len(text) < 700:
        return "faible"
    return "moyenne"


def empty_company_info() -> dict:
    return {
        "nombre_employes": "",
        "source_nombre_employes": "",
        "statut_entreprise": "",
        "source_statut": "",
        "niveau_confiance": "",
        "commentaire_verification": "",
    }


def confidence_rank(value: str) -> int:
    normalized = normalize_text(value)
    if normalized == "high":
        return 3
    if normalized == "medium":
        return 2
    if normalized == "low":
        return 1
    return 0


def normalize_company_confidence(value: str) -> str:
    normalized = normalize_text(value)
    if normalized == "high":
        return "High"
    if normalized == "medium":
        return "Medium"
    if normalized == "low":
        return "Low"
    return ""


def is_uncertain_status(value: str) -> bool:
    normalized = normalize_text(value)
    return not normalized or normalized in {"a verifier", "a vérifier", "inconnu", "unknown"}


def is_decisive_status(value: str) -> bool:
    normalized = normalize_text(value)
    return normalized in {"active", "inactive radiee", "inactive radie", "inactive / radiee", "inactive / radie"}


def company_info_from_row(row: pd.Series, columns: dict[str, str | None]) -> dict:
    info = empty_company_info()
    for key in info:
        info[key] = row_value(row, columns.get(key))
    info["niveau_confiance"] = normalize_company_confidence(info["niveau_confiance"])
    return info


def merge_company_info(existing: dict, candidate: dict) -> dict:
    result = empty_company_info()
    result.update({key: clean_text(existing.get(key, "")) for key in result})
    result["niveau_confiance"] = normalize_company_confidence(result["niveau_confiance"])

    candidate_clean = empty_company_info()
    candidate_clean.update({key: clean_text(candidate.get(key, "")) for key in candidate_clean})
    candidate_clean["niveau_confiance"] = normalize_company_confidence(candidate_clean["niveau_confiance"])

    existing_rank = confidence_rank(result["niveau_confiance"])
    candidate_rank = confidence_rank(candidate_clean["niveau_confiance"])

    replace_all = candidate_rank > existing_rank
    fill_only = candidate_rank <= existing_rank

    for key in [
        "nombre_employes",
        "source_nombre_employes",
        "statut_entreprise",
        "source_statut",
        "commentaire_verification",
    ]:
        candidate_value = candidate_clean.get(key, "")
        if not candidate_value:
            continue

        if replace_all:
            if key == "statut_entreprise" or not result.get(key):
                result[key] = candidate_value
            elif key == "commentaire_verification":
                result[key] = compact_join([result.get(key, ""), candidate_value])
            else:
                result[key] = candidate_value
            continue

        if fill_only and not result.get(key):
            result[key] = candidate_value

    if (
        candidate_clean.get("statut_entreprise")
        and is_uncertain_status(result.get("statut_entreprise", ""))
        and is_decisive_status(candidate_clean["statut_entreprise"])
        and candidate_rank >= existing_rank
    ):
        result["statut_entreprise"] = candidate_clean["statut_entreprise"]
        result["source_statut"] = candidate_clean.get("source_statut", result.get("source_statut", ""))

    if candidate_rank > existing_rank:
        result["niveau_confiance"] = candidate_clean["niveau_confiance"]
    elif not result["niveau_confiance"] and candidate_rank:
        result["niveau_confiance"] = candidate_clean["niveau_confiance"]

    if not result["niveau_confiance"] and (
        result["nombre_employes"] or result["statut_entreprise"] or result["commentaire_verification"]
    ):
        result["niveau_confiance"] = "Low"

    return result


def extract_siren_siret(text: str) -> dict[str, str]:
    text = clean_text(text)
    if not text:
        return {"siren": "", "siret": ""}

    candidates = []
    labeled_patterns = [
        r"\b(?:siret|siren|num[eé]ro\s+siret|num[eé]ro\s+siren|n[°o]\s*siret|n[°o]\s*siren)\b\s*[:\-]?\s*([0-9][0-9 .-]{8,18})",
        r"\bpappers\.fr/entreprise/[^/\s]*?([0-9]{9})(?:\b|/)",
        r"\bsociete\.com/societe/[^/\s]*?([0-9]{9})(?:\b|/)",
    ]

    for pattern in labeled_patterns:
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            digits = re.sub(r"\D", "", match.group(1))
            if len(digits) in {9, 14}:
                candidates.append(digits)

    if not candidates:
        return {"siren": "", "siret": ""}

    siret = next((value for value in candidates if len(value) == 14), "")
    siren = siret[:9] if siret else next((value for value in candidates if len(value) == 9), "")

    return {"siren": siren, "siret": siret}


def normalize_legal_name(value: str) -> str:
    normalized = normalize_text(value)
    legal_terms = {
        "sas",
        "sasu",
        "sa",
        "sarl",
        "eurl",
        "ltd",
        "limited",
        "inc",
        "corp",
        "corporation",
        "gmbh",
        "bv",
        "sl",
        "plc",
    }
    tokens = [token for token in normalized.split() if token not in legal_terms]
    return " ".join(tokens)


def find_first_nested_value(data, keys: set[str]):
    if isinstance(data, dict):
        for key, value in data.items():
            if normalize_text(key) in keys and value not in (None, ""):
                return value
        for value in data.values():
            nested = find_first_nested_value(value, keys)
            if nested not in (None, ""):
                return nested
    elif isinstance(data, list):
        for item in data:
            nested = find_first_nested_value(item, keys)
            if nested not in (None, ""):
                return nested
    return None


def pappers_effectif_label(value) -> str:
    if value in (None, ""):
        return ""

    value_text = clean_text(value)
    code = value_text.upper()
    effectif_codes = {
        "NN": "",
        "00": "0 salarié",
        "01": "1-2 salariés",
        "02": "3-5 salariés",
        "03": "6-9 salariés",
        "11": "10-19 salariés",
        "12": "20-49 salariés",
        "21": "50-99 salariés",
        "22": "100-199 salariés",
        "31": "200-249 salariés",
        "32": "250-499 salariés",
        "41": "500-999 salariés",
        "42": "1 000-1 999 salariés",
        "51": "2 000-4 999 salariés",
        "52": "5 000-9 999 salariés",
        "53": "10 000 salariés et plus",
    }
    return effectif_codes.get(code, value_text)


def call_pappers_api_by_siren(siren: str) -> dict:
    if not PAPPERS_API_TOKEN or not siren:
        return {}

    try:
        response = get_thread_http_session().get(
            PAPPERS_API_URL,
            params={"api_token": PAPPERS_API_TOKEN, "siren": siren},
            timeout=RESEARCH_TIMEOUT_SECONDS,
        )
        response.raise_for_status()
        return response.json()
    except (requests.RequestException, ValueError):
        return {}


def call_pappers_search_by_name(startup: str) -> dict:
    if not PAPPERS_API_TOKEN or not env_bool("ALGO3_PAPPERS_NAME_SEARCH", False) or not startup:
        return {}

    try:
        response = get_thread_http_session().get(
            PAPPERS_SEARCH_API_URL,
            params={"api_token": PAPPERS_API_TOKEN, "q": startup, "par_page": 3},
            timeout=RESEARCH_TIMEOUT_SECONDS,
        )
        response.raise_for_status()
        payload = response.json()
    except (requests.RequestException, ValueError):
        return {}

    results = payload.get("resultats") if isinstance(payload, dict) else None
    if not isinstance(results, list):
        return {}

    expected = normalize_legal_name(startup)
    exact_matches = []
    for item in results:
        if not isinstance(item, dict):
            continue
        names = [
            item.get("nom_entreprise"),
            item.get("denomination"),
            item.get("nom_complet"),
            item.get("enseigne"),
        ]
        if any(normalize_legal_name(name or "") == expected for name in names):
            exact_matches.append(item)

    if len(exact_matches) != 1:
        return {}

    siren = clean_text(exact_matches[0].get("siren", ""))
    return call_pappers_api_by_siren(siren) if siren else {}


def parse_pappers_company_info(data: dict) -> dict:
    if not isinstance(data, dict) or not data:
        return empty_company_info()

    info = empty_company_info()

    effectif_keys = {
        "tranche effectif",
        "tranche effectif salarie",
        "tranche effectif salarie entreprise",
        "effectif",
        "effectif salarie",
    }
    min_keys = {"effectif min", "effectif minimum"}
    max_keys = {"effectif max", "effectif maximum"}

    effectif_value = find_first_nested_value(data, effectif_keys)
    effectif_min = find_first_nested_value(data, min_keys)
    effectif_max = find_first_nested_value(data, max_keys)

    if effectif_min not in (None, "") and effectif_max not in (None, ""):
        info["nombre_employes"] = f"{clean_text(effectif_min)}-{clean_text(effectif_max)} salariés"
    else:
        info["nombre_employes"] = pappers_effectif_label(effectif_value)

    if info["nombre_employes"]:
        info["source_nombre_employes"] = "Pappers"

    entreprise_cessee = data.get("entreprise_cessee")
    date_cessation = find_first_nested_value(data, {
        "date cessation",
        "date cessation entreprise",
        "date cessation unite legale",
    })
    etat_value = find_first_nested_value(data, {
        "etat administratif",
        "etat administratif unite legale",
        "statut rcs",
        "statut diffusion",
    })
    normalized_status = normalize_text(etat_value or "")

    inactive_markers = [
        "cesse",
        "cessee",
        "radi",
        "inactive",
        "liquidation",
        "ferme",
        "closed",
        "dissolved",
    ]

    if entreprise_cessee is True or date_cessation or any(marker in normalized_status for marker in inactive_markers):
        info["statut_entreprise"] = "Inactive / radiée"
        info["source_statut"] = "Pappers"
    elif entreprise_cessee is False or normalized_status in {"a", "actif", "active"} or "actif" in normalized_status:
        info["statut_entreprise"] = "Active"
        info["source_statut"] = "Pappers"
    else:
        info["statut_entreprise"] = "À vérifier"
        info["source_statut"] = "Pappers (statut non conclusif)"

    info["niveau_confiance"] = "High" if info["nombre_employes"] or info["statut_entreprise"] != "À vérifier" else "Low"
    info["commentaire_verification"] = "Données issues de Pappers via SIREN/SIRET ou correspondance exacte."
    return info


def call_recherche_entreprises_api(query: str, per_page: int = 5) -> list[dict]:
    query = clean_text(query)
    if not ENABLE_RECHERCHE_ENTREPRISES_API or not query:
        return []

    if RECHERCHE_ENTREPRISES_DELAY_SECONDS > 0:
        time.sleep(RECHERCHE_ENTREPRISES_DELAY_SECONDS)

    try:
        response = get_thread_http_session().get(
            RECHERCHE_ENTREPRISES_API_URL,
            params={"q": query, "per_page": per_page},
            headers={
                "User-Agent": "LeonardAlgo3CompanyVerification/1.0",
                "Accept": "application/json",
            },
            timeout=RESEARCH_TIMEOUT_SECONDS,
        )
        response.raise_for_status()
        payload = response.json()
    except (requests.RequestException, ValueError):
        return []

    results = payload.get("results") if isinstance(payload, dict) else None
    return results if isinstance(results, list) else []


def gouv_candidate_names(candidate: dict) -> list[str]:
    return [
        clean_text(candidate.get("nom_raison_sociale", "")),
        clean_text(candidate.get("nom_complet", "")),
        clean_text(candidate.get("sigle", "")),
        clean_text((candidate.get("siege") or {}).get("nom_commercial", "")),
    ]


def gouv_candidate_name_score(startup: str, candidate: dict) -> float:
    expected = normalize_legal_name(startup)
    if not expected:
        return 0

    scores = []
    for name in gouv_candidate_names(candidate):
        normalized_name = normalize_legal_name(name)
        if not normalized_name:
            continue
        if normalized_name == expected:
            scores.append(1.0)
            continue
        expected_tokens = set(expected.split())
        name_tokens = set(normalized_name.split())
        token_overlap = len(expected_tokens & name_tokens) / max(len(expected_tokens), 1)
        sequence_ratio = SequenceMatcher(None, expected, normalized_name).ratio()
        scores.append(max(token_overlap, sequence_ratio))

    return max(scores or [0])


def select_recherche_entreprises_candidate(results: list[dict], startup: str, siren: str = "") -> dict:
    if not results:
        return {}

    if siren:
        exact_siren = [item for item in results if clean_text(item.get("siren", "")) == siren]
        return exact_siren[0] if exact_siren else {}

    expected = normalize_legal_name(startup)
    exact_matches = []
    for item in results:
        if any(normalize_legal_name(name) == expected for name in gouv_candidate_names(item)):
            exact_matches.append(item)

    if len(exact_matches) == 1:
        return exact_matches[0]
    if len(exact_matches) > 1:
        return {}

    scored = sorted(
        [(gouv_candidate_name_score(startup, item), item) for item in results],
        key=lambda item: item[0],
        reverse=True,
    )
    best_score, best_candidate = scored[0]
    second_score = scored[1][0] if len(scored) > 1 else 0

    if len(scored) == 1 and best_score >= 0.94:
        return best_candidate
    if best_score >= 0.97 and second_score <= 0.80:
        return best_candidate
    return {}


def parse_recherche_entreprises_info(candidate: dict) -> dict:
    info = empty_company_info()
    if not isinstance(candidate, dict) or not candidate:
        return info

    siege = candidate.get("siege") or {}
    siren = clean_text(candidate.get("siren", ""))
    siret = clean_text(siege.get("siret", ""))
    source = "API Recherche d’Entreprises / INSEE"
    if siren:
        source = f"{source} - SIREN {siren}"
    if siret:
        source = f"{source} / SIRET {siret}"

    effectif = (
        pappers_effectif_label(candidate.get("tranche_effectif_salarie"))
        or pappers_effectif_label(siege.get("tranche_effectif_salarie"))
    )
    if effectif:
        info["nombre_employes"] = effectif
        info["source_nombre_employes"] = source

    etat = normalize_text(candidate.get("etat_administratif", "") or siege.get("etat_administratif", ""))
    date_fermeture = candidate.get("date_fermeture") or siege.get("date_fermeture")

    if etat == "a" and not date_fermeture:
        info["statut_entreprise"] = "Active"
        info["source_statut"] = source
    elif etat == "c" or date_fermeture:
        info["statut_entreprise"] = "Inactive / radiée"
        info["source_statut"] = source
    else:
        info["statut_entreprise"] = "À vérifier"
        info["source_statut"] = source

    if info["nombre_employes"] or info["statut_entreprise"] in {"Active", "Inactive / radiée"}:
        info["niveau_confiance"] = "High"
        info["commentaire_verification"] = "Données issues de l’API officielle Recherche d’Entreprises."
    else:
        info["niveau_confiance"] = "Low"
        info["commentaire_verification"] = "API officielle consultée, mais statut/effectif non conclusif."

    return info


def build_recherche_entreprises_company_info(row_data: dict, identifiers: dict[str, str]) -> dict:
    if not ENABLE_RECHERCHE_ENTREPRISES_API:
        return empty_company_info()

    query = identifiers.get("siren") or row_data.get("startup", "")
    if not query:
        return empty_company_info()

    results = call_recherche_entreprises_api(query, per_page=5)
    selected = select_recherche_entreprises_candidate(
        results,
        row_data.get("startup", ""),
        identifiers.get("siren", ""),
    )
    return parse_recherche_entreprises_info(selected)


def extract_societe_siren_from_url(url: str) -> str:
    match = re.search(r"-(\d{9})\.html(?:$|[?#])", clean_text(url))
    return match.group(1) if match else ""


def is_societe_company_url(url: str) -> bool:
    parsed = urlparse(url)
    host = parsed.netloc.lower().removeprefix("www.")
    return host == "societe.com" and "/societe/" in parsed.path and bool(extract_societe_siren_from_url(url))


def extract_societe_company_urls(text: str) -> list[str]:
    urls = []
    seen = set()
    for match in re.finditer(
        r"https?://(?:www\.)?societe\.com/societe/[^\s\"'<>)]*?-\d{9}\.html",
        clean_text(text),
        flags=re.IGNORECASE,
    ):
        url = match.group(0).rstrip(".,;")
        if url not in seen:
            seen.add(url)
            urls.append(url)
    return urls


def fetch_societe_html(url: str, params: dict | None = None) -> tuple[str, str]:
    if not ENABLE_SOCIETE_SEARCH:
        return "", ""

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (compatible; LeonardAlgo3Societe/1.0; "
            "startup verification research)"
        ),
        "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.7",
    }

    if SOCIETE_REQUEST_DELAY_SECONDS > 0:
        time.sleep(SOCIETE_REQUEST_DELAY_SECONDS)

    try:
        response = get_thread_http_session().get(
            url,
            params=params,
            headers=headers,
            timeout=RESEARCH_TIMEOUT_SECONDS,
        )
        response.raise_for_status()
        response.encoding = response.apparent_encoding or response.encoding
        return response.text, response.url
    except requests.RequestException:
        return "", ""


def societe_search_candidates(query: str) -> list[dict]:
    query = clean_text(query)
    if not ENABLE_SOCIETE_SEARCH or not query:
        return []

    html, final_url = fetch_societe_html(SOCIETE_SEARCH_URL, params={"q": query})
    if not html:
        return []

    candidates = []
    seen_sirens = set()

    if final_url and is_societe_company_url(final_url):
        siren = extract_societe_siren_from_url(final_url)
        candidates.append({
            "name": "",
            "siren": siren,
            "url": final_url,
            "snippet": "",
        })
        seen_sirens.add(siren)

    soup = BeautifulSoup(html, "html.parser")
    for anchor in soup.find_all("a", href=True):
        absolute_url = urljoin(SOCIETE_BASE_URL, anchor.get("href", ""))
        if not is_societe_company_url(absolute_url):
            continue

        siren = extract_societe_siren_from_url(absolute_url)
        if not siren or siren in seen_sirens:
            continue

        parent_text = ""
        parent = anchor.find_parent(["li", "article", "div", "tr"])
        if parent:
            parent_text = clean_text(parent.get_text(" ", strip=True))

        name = clean_text(anchor.get_text(" ", strip=True))
        if not name and parent_text:
            name = parent_text.split(" SIREN")[0].strip()

        candidates.append({
            "name": name,
            "siren": siren,
            "url": absolute_url,
            "snippet": parent_text,
        })
        seen_sirens.add(siren)

        if len(candidates) >= 8:
            break

    return candidates


def candidate_name_score(startup: str, candidate: dict) -> float:
    expected = normalize_legal_name(startup)
    candidate_text = normalize_legal_name(compact_join([
        candidate.get("name", ""),
        candidate.get("snippet", ""),
    ]))

    if not expected or not candidate_text:
        return 0

    if expected == candidate_text:
        return 1.0

    expected_tokens = set(expected.split())
    candidate_tokens = set(candidate_text.split())
    if not expected_tokens or not candidate_tokens:
        return 0

    token_overlap = len(expected_tokens & candidate_tokens) / max(len(expected_tokens), 1)
    sequence_ratio = SequenceMatcher(None, expected, candidate_text).ratio()
    return max(token_overlap, sequence_ratio)


def select_societe_candidate(candidates: list[dict], startup: str, siren: str = "") -> dict:
    if not candidates:
        return {}

    if siren:
        siren_matches = [candidate for candidate in candidates if candidate.get("siren") == siren]
        if siren_matches:
            return siren_matches[0]
        return {}

    if not SOCIETE_NAME_SEARCH:
        return {}

    scored = sorted(
        [(candidate_name_score(startup, candidate), candidate) for candidate in candidates],
        key=lambda item: item[0],
        reverse=True,
    )
    best_score, best_candidate = scored[0]
    second_score = scored[1][0] if len(scored) > 1 else 0

    if best_score >= 0.98:
        return best_candidate

    if len(scored) == 1 and best_score >= 0.90:
        return best_candidate

    if best_score >= 0.93 and second_score <= 0.75:
        return best_candidate

    return {}


def normalize_employee_range(value: str) -> str:
    value = clean_text(value)
    if not value:
        return ""

    value = re.sub(r"(?<=\d)\s*(?:à|a|to|–)\s*(?=\d)", "-", value, flags=re.IGNORECASE)
    value = re.sub(r"(?<=\d),(?=\d{3}\b)", " ", value)
    value = re.sub(r"\s+", " ", value)
    value = value.replace("salarie", "salarié")
    value = value.replace("salaries", "salariés")
    if re.fullmatch(r"\d{1,3}(?: \d{3})?(?:-\d{1,3}(?: \d{3})?)?\+?", value):
        value = f"{value} employés"
    return value.strip()


def parse_societe_detail_info(html: str, url: str) -> dict:
    info = empty_company_info()
    if not html:
        return info

    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg"]):
        tag.decompose()

    page_text = clean_text(soup.get_text(" ", strip=True))
    sample = page_text[:3500]
    normalized_sample = normalize_text(sample)

    inactive_pattern = (
        r"\b(?:radiee|radie|radiation|cessation|cessee|cesse|inactive|"
        r"liquidation|dissolution|fermee|ferme)\b"
    )

    if re.search(inactive_pattern, normalized_sample):
        info["statut_entreprise"] = "Inactive / radiée"
        info["source_statut"] = url
    elif re.search(r"\bActive\b", sample):
        info["statut_entreprise"] = "Active"
        info["source_statut"] = url
    else:
        info["statut_entreprise"] = "À vérifier"
        info["source_statut"] = url

    employee_patterns = [
        r"\b(?:Active|Radi[ée]e|Radiée|Inactive|Cess[ée]e)\s+((?:\d{1,5}\s*(?:à|a|-|–|to)\s*\d{1,5}|\d{1,5}\+?)\s+salari[ée]s?)\b",
        r"\b(?:effectif|nombre de salari[ée]s|salari[ée]s)\s*[:\-]?\s*((?:\d{1,5}\s*(?:à|a|-|–|to)\s*\d{1,5}|\d{1,5}\+?)\s*(?:salari[ée]s?)?)\b",
        r"\bson effectif est compris entre\s+([^.]{1,80})",
    ]

    for pattern in employee_patterns:
        match = re.search(pattern, sample, flags=re.IGNORECASE)
        if match:
            employees = normalize_employee_range(match.group(1))
            if employees:
                info["nombre_employes"] = employees
                info["source_nombre_employes"] = url
                break

    if info["statut_entreprise"] in {"Active", "Inactive / radiée"} or info["nombre_employes"]:
        info["niveau_confiance"] = "High"
        info["commentaire_verification"] = "Données extraites de la fiche Société.com."
    else:
        info["niveau_confiance"] = "Low"
        info["commentaire_verification"] = "Fiche Société.com trouvée mais statut/effectif non conclusif."

    return info


def build_societe_company_info(row_data: dict, combined_text: str, identifiers: dict[str, str]) -> dict:
    if not ENABLE_SOCIETE_SEARCH:
        return empty_company_info()

    direct_urls = extract_societe_company_urls(combined_text)
    if identifiers.get("siren"):
        direct_urls = [
            url for url in direct_urls
            if extract_societe_siren_from_url(url) == identifiers["siren"]
        ] or direct_urls

    for url in direct_urls[:2]:
        html, final_url = fetch_societe_html(url)
        info = parse_societe_detail_info(html, final_url or url)
        if info.get("nombre_employes") or is_decisive_status(info.get("statut_entreprise", "")):
            return info

    query = identifiers.get("siren") or row_data.get("startup", "")
    if not query:
        return empty_company_info()

    candidates = societe_search_candidates(query)
    selected = select_societe_candidate(
        candidates,
        row_data.get("startup", ""),
        identifiers.get("siren", ""),
    )
    if not selected:
        return empty_company_info()

    html, final_url = fetch_societe_html(selected["url"])
    info = parse_societe_detail_info(html, final_url or selected["url"])
    if info.get("commentaire_verification"):
        info["commentaire_verification"] = compact_join([
            info["commentaire_verification"],
            f"Correspondance Société.com retenue : SIREN {selected.get('siren', '')}".strip(),
        ])
    return info


def extract_linkedin_company_urls(text: str) -> list[str]:
    urls = []
    seen = set()

    for match in re.finditer(
        r"https?://(?:[\w.-]+\.)?linkedin\.com/(?:company|showcase)/[^\s\"'<>)]*",
        clean_text(text),
        flags=re.IGNORECASE,
    ):
        url = match.group(0).rstrip(".,;")
        parsed = urlparse(url)
        clean_url = parsed._replace(query="", fragment="").geturl().rstrip("/")
        if clean_url and clean_url not in seen:
            seen.add(clean_url)
            urls.append(clean_url)

    return urls


def fetch_linkedin_public_text(url: str) -> tuple[str, str]:
    if not ENABLE_LINKEDIN_EMPLOYEE_SEARCH or not is_valid_url(url):
        return "", ""

    parsed = urlparse(normalize_url(url))
    host = parsed.netloc.lower()
    if "linkedin.com" not in host or not re.search(r"/(?:company|showcase)/", parsed.path):
        return "", ""

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0 Safari/537.36"
        ),
        "Accept-Language": "en-US,en;q=0.9,fr-FR;q=0.8,fr;q=0.7",
    }

    if LINKEDIN_REQUEST_DELAY_SECONDS > 0:
        time.sleep(LINKEDIN_REQUEST_DELAY_SECONDS)

    try:
        response = get_thread_http_session().get(
            normalize_url(url),
            headers=headers,
            timeout=RESEARCH_TIMEOUT_SECONDS,
        )
        if response.status_code in {401, 403, 429, 999}:
            return "", response.url
        response.raise_for_status()
        response.encoding = response.apparent_encoding or response.encoding
    except requests.RequestException:
        return "", ""

    soup = BeautifulSoup(response.text, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg"]):
        tag.decompose()

    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    meta_description = ""
    for meta in soup.find_all("meta"):
        name = normalize_text(meta.get("name", "") or meta.get("property", ""))
        if name in {"description", "og description", "twitter description"} and meta.get("content"):
            meta_description = compact_join([meta_description, meta.get("content", "")])

    visible_text = clean_text(soup.get_text(" ", strip=True))
    return clean_text(compact_join([title, meta_description, visible_text]))[:12000], response.url


def detect_linkedin_employee_range(text: str) -> str:
    text = clean_text(text)
    if not text:
        return ""

    patterns = [
        r"\bcompany size\s*[:\-]?\s*((?:\d{1,3}(?:,\d{3})?|\d{1,5})\s*(?:-|–|to|à|a)\s*(?:\d{1,3}(?:,\d{3})?|\d{1,5})\+?)\s*(?:employees|employés|employes)?\b",
        r"\btaille de l['’ ]entreprise\s*[:\-]?\s*((?:\d{1,3}(?:,\d{3})?|\d{1,5})\s*(?:-|–|to|à|a)\s*(?:\d{1,3}(?:,\d{3})?|\d{1,5})\+?)\s*(?:employees|employés|employes|personnes)?\b",
        r"\b((?:\d{1,3}(?:,\d{3})?|\d{1,5})\s*(?:-|–|to|à|a)\s*(?:\d{1,3}(?:,\d{3})?|\d{1,5})\+?)\s+(?:employees|employés|employes)\b",
        r"\b((?:\d{1,3}(?:,\d{3})?|\d{1,5})\+?)\s+(?:employees|employés|employes)\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if not match:
            continue
        employees = normalize_employee_range(match.group(1))
        if employees:
            return employees

    return ""


def build_linkedin_company_info(row: pd.Series, row_data: dict, combined_text: str) -> dict:
    info = empty_company_info()
    if not ENABLE_LINKEDIN_EMPLOYEE_SEARCH:
        return info

    row_text = compact_join([clean_text(value) for value in row.to_dict().values()])
    linkedin_urls = extract_linkedin_company_urls(compact_join([
        combined_text,
        row_text,
        row_data.get("public_external_sources", ""),
    ]))

    available_text = compact_join([
        combined_text,
        row_text,
    ])
    employees = detect_linkedin_employee_range(available_text)
    if employees and linkedin_urls:
        info["nombre_employes"] = employees
        info["source_nombre_employes"] = f"LinkedIn - {linkedin_urls[0]}"
        info["niveau_confiance"] = "Medium"
        info["commentaire_verification"] = "Taille d'entreprise LinkedIn trouvée dans les données disponibles."
        return info

    for linkedin_url in linkedin_urls[:2]:
        linkedin_text, final_url = fetch_linkedin_public_text(linkedin_url)
        employees = detect_linkedin_employee_range(linkedin_text)
        if employees:
            info["nombre_employes"] = employees
            info["source_nombre_employes"] = f"LinkedIn - {final_url or linkedin_url}"
            info["niveau_confiance"] = "Medium"
            info["commentaire_verification"] = "Taille d'entreprise extraite d'une page LinkedIn publique accessible."
            return info

    return info


def detect_employees_from_text(text: str) -> str:
    text = clean_text(text)
    if not text:
        return ""

    patterns = [
        r"\b(?:company size|taille de l[' ]entreprise|effectif|employees|employés|employes|collaborateurs)\s*[:\-]?\s*((?:\d{1,3}(?:,\d{3})?|\d{1,5})\s*(?:-|–|to|à|a)\s*(?:\d{1,3}(?:,\d{3})?|\d{1,5}))\b",
        r"\b((?:\d{1,3}(?:,\d{3})?|\d{1,5})\s*(?:-|–|to|à|a)\s*(?:\d{1,3}(?:,\d{3})?|\d{1,5}))\s+(?:employees|employés|employes|collaborateurs|people|personnes)\b",
        r"(?<![-,\d])\b(\d{1,5}\+?)\s+(?:employees|employés|employes|collaborateurs|people|team members|personnes)\b",
        r"\bteam of\s+(\d{1,5}\+?)\b",
        r"\b(?:équipe|equipe)\s+de\s+(\d{1,5}\+?)\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if not match:
            continue
        value = clean_text(match.group(1))
        return normalize_employee_range(value)

    return ""


def has_recent_activity_signal(text: str, external_sources: str) -> bool:
    normalized = normalize_text(text)
    external_normalized = normalize_text(external_sources)

    recent_year = re.search(r"\b202[4-6]\b", normalized) is not None
    active_terms = [
        "we are hiring",
        "we re hiring",
        "hiring",
        "careers",
        "jobs",
        "recrutement",
        "recrute",
        "news",
        "press",
        "blog",
        "case study",
        "case studies",
        "customers",
        "clients",
        "trusted by",
        "raised",
        "funding",
        "levee de fonds",
        "partnership",
        "partners",
    ]

    return recent_year or any(term in normalized for term in active_terms) or "linkedin" in external_normalized


def detect_status_from_public_research(research: dict, website: str) -> dict:
    info = empty_company_info()
    text = clean_text(research.get("text", ""))
    sources = clean_text(research.get("sources", ""))
    external_sources = clean_text(research.get("external_sources", ""))
    combined = compact_join([text, external_sources])
    normalized = normalize_text(combined)

    if not is_valid_url(website):
        info["statut_entreprise"] = "À vérifier"
        info["niveau_confiance"] = "Low"
        info["commentaire_verification"] = "Aucun site officiel exploitable pour vérifier l'activité."
        return info

    inactive_terms = [
        "company closed",
        "business closed",
        "no longer operating",
        "shut down",
        "shutdown",
        "dissolved",
        "liquidation",
        "cessation d activite",
        "cessé son activité",
        "cesse son activite",
        "radiation",
        "radiee",
        "radiée",
    ]

    if any(normalize_text(term) in normalized for term in inactive_terms):
        info["statut_entreprise"] = "Inactive / radiée"
        info["source_statut"] = sources or website
        info["niveau_confiance"] = "Medium"
        info["commentaire_verification"] = "Signal public d'inactivité ou de cessation détecté."
        return info

    if text and has_recent_activity_signal(text, external_sources):
        info["statut_entreprise"] = "Active"
        source_parts = [sources or normalize_url(website)]
        if external_sources:
            source_parts.append(external_sources)
        info["source_statut"] = compact_join(source_parts)
        info["niveau_confiance"] = "Medium"
        info["commentaire_verification"] = "Site officiel accessible avec signaux publics d'activité récente ou lien social/profil société."
        return info

    if text:
        info["statut_entreprise"] = "À vérifier"
        info["source_statut"] = sources or normalize_url(website)
        info["niveau_confiance"] = "Low"
        info["commentaire_verification"] = "Site officiel accessible, mais activité récente insuffisamment prouvée."
        return info

    info["statut_entreprise"] = "À vérifier"
    info["niveau_confiance"] = "Low"
    info["commentaire_verification"] = "Site inaccessible, timeout, DNS ou contenu non exploitable."
    return info


def build_public_company_info(row_data: dict, research: dict) -> dict:
    info = detect_status_from_public_research(research, row_data.get("website", ""))

    employees = detect_employees_from_text(compact_join([
        research.get("text", ""),
        research.get("signals", ""),
        row_data.get("description", ""),
    ]))
    if employees:
        info["nombre_employes"] = employees
        info["source_nombre_employes"] = research.get("sources", "") or row_data.get("website", "")
        if confidence_rank(info.get("niveau_confiance", "")) < 2:
            info["niveau_confiance"] = "Medium"

    return info


def enrich_company_info(row: pd.Series, columns: dict[str, str | None], row_data: dict, research: dict) -> dict:
    existing = company_info_from_row(row, columns)
    combined_text = compact_join([
        row_data.get("startup", ""),
        row_data.get("description", ""),
        row_data.get("website", ""),
        row_data.get("siren_siret", ""),
        row_data.get("excel_context", ""),
        row_data.get("public_research_sources", ""),
        row_data.get("public_external_sources", ""),
        row_data.get("public_research_text", ""),
    ])

    identifiers = extract_siren_siret(combined_text)
    pappers_info = empty_company_info()

    if identifiers["siren"]:
        pappers_info = parse_pappers_company_info(call_pappers_api_by_siren(identifiers["siren"]))
    elif PAPPERS_API_TOKEN:
        pappers_info = parse_pappers_company_info(call_pappers_search_by_name(row_data.get("startup", "")))

    gouv_info = build_recherche_entreprises_company_info(row_data, identifiers)
    societe_info = build_societe_company_info(row_data, combined_text, identifiers)
    linkedin_info = build_linkedin_company_info(row, row_data, combined_text)
    public_info = build_public_company_info(row_data, research)
    merged = merge_company_info(existing, pappers_info)
    merged = merge_company_info(merged, gouv_info)
    merged = merge_company_info(merged, societe_info)
    merged = merge_company_info(merged, linkedin_info)
    merged = merge_company_info(merged, public_info)

    if not merged["statut_entreprise"] and not merged["nombre_employes"]:
        merged["statut_entreprise"] = "À vérifier"
        merged["niveau_confiance"] = "Low"
        merged["commentaire_verification"] = "Aucune information fiable trouvée pour compléter l'effectif ou l'état de l'entreprise."

    return merged


def research_startup_website(site: str, cache: dict[str, dict], cache_lock: threading.Lock) -> dict:
    if not ENABLE_RESEARCH or not is_valid_url(site):
        return {
            "sources": "",
            "external_sources": "",
            "quality": "desactivee" if not ENABLE_RESEARCH else "inconnu",
            "text": "",
            "signals": format_research_signals(detect_research_signals("")),
            "targeted_sources": "",
            "targeted_categories": "",
            "targeted_summary": "",
            "evidence_employees": "",
            "evidence_funding": "",
            "evidence_customers": "",
            "evidence_product": "",
            "evidence_deployment": "",
            "evidence_international": "",
            "evidence_production": "",
            "evidence_materials": "",
        }

    base_url = normalize_url(site).rstrip("/")

    with cache_lock:
        cached = cache.get(base_url)
        if cached:
            return cached

    homepage_text, links = fetch_page_text_and_links(base_url)
    selected_links = select_research_links(base_url, links)

    page_texts = []
    sources = []
    page_records = []
    targeted_categories = []
    trusted_external_links = collect_trusted_external_links(links)

    for link_record in selected_links:
        if sum(len(item) for item in page_texts) >= RESEARCH_MAX_TEXT_CHARS:
            break
        link = link_record.get("url", "")
        if not link:
            continue
        if link == base_url:
            text = homepage_text
            page_links = links
        else:
            text, page_links = fetch_page_text_and_links(link)
            trusted_external_links.extend(collect_trusted_external_links(page_links))
        if not text:
            continue

        categories = unique_values(
            list(link_record.get("categories", []))
            + targeted_maturity_categories_from_text(f"{link} {link_record.get('text', '')} {text[:1200]}")
        )
        if link == base_url and not categories:
            categories = ["homepage"]

        text = truncate_research_page_text(text)
        category_labels = format_maturity_category_labels(categories)
        page_texts.append(f"Source: {link}\nType de page: {category_labels}\n{text}")
        sources.append(link)
        page_records.append({
            "url": link,
            "categories": categories,
            "text": text,
        })
        targeted_categories.extend([category for category in categories if category != "homepage"])
        if len(sources) >= RESEARCH_MAX_PAGES:
            break

    research_text = clean_text("\n\n".join(page_texts))[:RESEARCH_MAX_TEXT_CHARS]
    signals = detect_research_signals(research_text)
    evidence = build_maturity_research_evidence(page_records)
    targeted_summary = build_maturity_research_summary(page_records, evidence)
    targeted_sources = format_maturity_page_records(page_records)

    result = {
        "sources": " | ".join(sources),
        "external_sources": " | ".join(dict.fromkeys(trusted_external_links)),
        "quality": research_quality_from_sources(sources, research_text, targeted_categories),
        "text": research_text,
        "signals": format_research_signals(signals),
        "targeted_sources": targeted_sources,
        "targeted_categories": format_maturity_category_labels(unique_values(targeted_categories)),
        "targeted_summary": targeted_summary,
        "evidence_employees": evidence.get("employees", ""),
        "evidence_funding": evidence.get("funding", ""),
        "evidence_customers": evidence.get("customers", ""),
        "evidence_product": evidence.get("product", ""),
        "evidence_deployment": evidence.get("deployment", ""),
        "evidence_international": evidence.get("international", ""),
        "evidence_production": evidence.get("production", ""),
        "evidence_materials": evidence.get("materials", ""),
    }

    with cache_lock:
        cache[base_url] = result

    return result


def strip_algo2_suffix(stem: str) -> str:
    for suffix in [
        "_algo_2_affinage_strategique_sauvegarde_temp",
        "_algo_2_affinage_strategique",
    ]:
        if stem.endswith(suffix):
            return stem[: -len(suffix)]
    return stem


def discover_input_files(base_dir: Path, cli_args: list[str]) -> list[Path]:
    if cli_args:
        return [Path(arg).expanduser().resolve() for arg in cli_args]

    candidates = []
    for path in sorted(base_dir.iterdir()):
        if not path.is_file() or path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
            continue
        if path.name.startswith("~$"):
            continue
        if OUTPUT_SUFFIX in path.name or TEMP_SUFFIX in path.name:
            continue
        if "_algo_2_affinage_strategique" in path.stem:
            candidates.append(path)

    selected_by_base = {}
    for path in candidates:
        base = strip_algo2_suffix(path.stem)
        current = selected_by_base.get(base)
        if current is None:
            selected_by_base[base] = path
            continue
        if current.name.endswith("_sauvegarde_temp.xlsx") and not path.name.endswith("_sauvegarde_temp.xlsx"):
            selected_by_base[base] = path

    return list(selected_by_base.values())


def make_output_path(input_file: Path) -> Path:
    base = strip_algo2_suffix(input_file.stem)
    return input_file.with_name(f"{base}{OUTPUT_SUFFIX}")


def make_temp_path(input_file: Path) -> Path:
    base = strip_algo2_suffix(input_file.stem)
    return input_file.with_name(f"{base}{TEMP_SUFFIX}")


# ============================================================
# AZURE OPENAI
# ============================================================

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
                "Tu es un analyste innovation rigoureux pour Leonard / VINCI. "
                "Tu réponds uniquement avec un objet JSON valide, sans markdown."
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


def call_chat_plain_json(client: AzureOpenAI, messages: list[dict]) -> str:
    response = client.chat.completions.create(
        model=MODEL_DEPLOYMENT,
        messages=messages,
    )
    return response.choices[0].message.content or ""


def call_model_json(
    client: AzureOpenAI,
    messages: list[dict],
    schema: dict,
    schema_name: str,
) -> dict:
    mode = ALGO3_API_MODE

    if mode not in {"chat", "responses", "auto", "json_object", "plain"}:
        mode = "chat"

    if mode == "responses":
        return parse_json_text(call_responses_json_schema(client, messages, schema, schema_name))

    if mode == "json_object":
        try:
            return parse_json_text(call_chat_json_object(client, messages))
        except Exception as exc:
            if getattr(exc, "status_code", None) != 400:
                raise
            log(f"json_object refusé par Azure, bascule en JSON libre : {format_api_error(exc)}")
            return parse_json_text(call_chat_plain_json(client, messages))

    if mode == "plain":
        return parse_json_text(call_chat_plain_json(client, messages))

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

    try:
        return parse_json_text(call_chat_json_object(client, messages))
    except Exception as exc:
        if getattr(exc, "status_code", None) != 400:
            raise
        log(f"json_object refusé par Azure, bascule en JSON libre : {format_api_error(exc)}")
        return parse_json_text(call_chat_plain_json(client, messages))


def validate_azure_client() -> None:
    client = build_azure_client()

    schema = {
        "type": "object",
        "properties": {
            "ok": {
                "type": "string",
                "enum": ["oui"],
            },
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
            schema_name="algo3_healthcheck",
        )
        log("Connexion Azure OpenAI OK.")

    except Exception as exc:
        raise RuntimeError(f"Test Azure OpenAI échoué : {format_api_error(exc)}")


# ============================================================
# PROMPT ET SCHEMA
# ============================================================

def algo3_schema() -> dict:
    return {
        "type": "object",
        "properties": {
            "algo3_orientation": {
                "type": "string",
                "enum": ORIENTATIONS,
            },
            "algo3_confidence": {
                "type": "number",
                "description": "Score de confiance entre 0 et 1.",
            },
            "secondary_orientation": {
                "type": "string",
                "enum": SECONDARY_ORIENTATIONS,
            },
            "solution_type": {
                "type": "string",
                "enum": SOLUTION_TYPES,
            },
            "maturity_stage": {
                "type": "string",
                "enum": MATURITY_STAGES,
            },
            "trl_estimate": {
                "type": "string",
                "enum": TRL_VALUES,
            },
            "mrl_estimate": {
                "type": "string",
                "enum": MRL_VALUES,
            },
            "employees_estimate": {
                "type": "string",
            },
            "revenue_arr_estimate": {
                "type": "string",
            },
            "funding_stage": {
                "type": "string",
                "enum": FUNDING_STAGES,
            },
            "product_commercialized": {
                "type": "string",
                "enum": PRODUCT_VALUES,
            },
            "production_capacity": {
                "type": "string",
                "enum": CAPACITY_VALUES,
            },
            "international_deployment_capacity": {
                "type": "string",
                "enum": CAPACITY_VALUES,
            },
            "field_reference": {
                "type": "string",
                "enum": REFERENCE_VALUES,
            },
            "materials_vertical": {
                "type": "string",
                "enum": YES_NO_UNKNOWN,
            },
            "orientation_reason": {
                "type": "string",
            },
            "missing_information": {
                "type": "string",
            },
        },
        "required": [
            "algo3_orientation",
            "algo3_confidence",
            "secondary_orientation",
            "solution_type",
            "maturity_stage",
            "trl_estimate",
            "mrl_estimate",
            "employees_estimate",
            "revenue_arr_estimate",
            "funding_stage",
            "product_commercialized",
            "production_capacity",
            "international_deployment_capacity",
            "field_reference",
            "materials_vertical",
            "orientation_reason",
            "missing_information",
        ],
        "additionalProperties": False,
    }


def build_prompt(row_data: dict) -> str:
    prompt_row_data = dict(row_data)
    if not INCLUDE_EXCEL_COLUMN_LIST_IN_PROMPT:
        prompt_row_data.pop("excel_context_columns", None)
        prompt_row_data.pop("excel_context_truncated", None)
    profile_context = dealflow_profiles.profile_prompt_context()
    seed_arr_label = "500k EUR" if dealflow_profiles.is_latam() else "1M EUR"
    catalyst_arr_label = "500k EUR" if dealflow_profiles.is_latam() else "1M EUR"

    return f"""
Tu es un analyste innovation pour Leonard / VINCI.

Tu appliques l'Algo 3 : orientation programme.
L'Algo 3 n'est pas un algorithme de préqualification. Il intervient uniquement après Algo 2.
Son rôle est d'orienter une startup déjà préqualifiée ou à vérifier vers le bon programme ou la bonne catégorie.

Profil de traitement :
{profile_context}

Orientations possibles :
- Seed
- Catalyst
- Matériaux
- Autre
- A_VERIFIER

1. Orientation Seed
Une startup est plutôt Seed si elle est à un stade jeune ou intermédiaire :
- TRL > 7 ;
- MRL autour de 3 ;
- équipe fondatrice ou petite équipe ;
- SaaS < 5 employés ;
- Hardware < 15 employés ;
- chiffre d'affaires ou ARR < {seed_arr_label} ;
- levée de fonds avant Série A ;
- produit pas obligatoirement commercialisé, mais proche de la commercialisation ;
- capacité à produire pas obligatoire ;
- capacité internationale pas obligatoire ;
- référence terrain pas obligatoire mais souhaitable.

2. Orientation Catalyst
Une startup est plutôt Catalyst si elle est mature, structurée et déployable :
- TRL > 9 ;
- MRL autour de 7 ;
- SaaS > 5 employés ;
- Hardware > 15 employés ;
- chiffre d'affaires ou ARR > {catalyst_arr_label} ;
- levée de fonds à partir de la Série A ;
- produit déjà commercialisé obligatoire ;
- capacité à produire obligatoire ;
- capacité de déploiement à l'international obligatoire ;
- au moins une référence terrain obligatoire.

3. Orientation Matériaux
Priorise Matériaux si la startup est fortement liée à :
- matériaux de construction ;
- béton, ciment, acier, bois, composites, polymères ;
- matériaux bas carbone ;
- matériaux biosourcés ;
- recyclage ou circularité des matériaux ;
- procédés matériaux ;
- revêtements, isolation, performance matériau ;
- innovation matière directement exploitable dans construction, infrastructure, énergie ou real estate.

Important :
La catégorie Matériaux peut être choisie même si la startup semble aussi compatible Seed ou Catalyst.
Dans ce cas, indique une secondary_orientation Seed ou Catalyst si possible.

4. Orientation Autre
Choisis Autre si la startup est préqualifiée mais ne correspond pas clairement à Seed, Catalyst ou Matériaux.

5. Orientation A_VERIFIER
Choisis A_VERIFIER si les données sont insuffisantes pour orienter de manière fiable.

Règles de prudence :
- Ne pas inventer les chiffres.
- Si employés, chiffre d'affaires, levée de fonds, TRL ou MRL sont inconnus, écris "inconnu".
- Utilise les extraits de recherche publique uniquement comme indices : s'ils ne prouvent pas clairement
  un critère, laisse ce critère à "inconnu".
- Grille TRL adaptée à l'analyse startup/VINCI :
  - TRL <7 : concept, recherche, preuve de concept, validation laboratoire, prototype amont.
  - TRL 7-8 : prototype ou démonstrateur testé en environnement réel/proche réel, pilote terrain,
    test sur site, MVP/beta avec premiers retours.
  - TRL 9+ : produit complet commercialisé, en production, déployé chez des clients, références
    terrain/cas clients, installations ou usage en conditions réelles.
- Ne considère pas une simple mention "innovant", "IA", "plateforme" ou "prototype" comme TRL 7-8
  sans preuve de test terrain ou de pilote.
- Ne considère pas une startup comme TRL 9+ sans preuve de commercialisation, déploiement,
  clients ou usage opérationnel réel.
- Catalyst exige des preuves plus fortes que Seed.
- Si la startup semble mature mais sans preuve de produit commercialisé, référence terrain,
  capacité de production ou capacité internationale, ne l'oriente pas automatiquement vers Catalyst.
- Si plusieurs critères sont manquants, mets A_VERIFIER.
- Seed peut être proposé avec moins de preuves que Catalyst, mais il faut au moins des indices
  de maturité jeune ou intermédiaire.
- Matériaux doit être priorisé si la verticale matériau est clairement détectée.
- Utilise toutes les informations du champ excel_context : réponses de candidature, clients cibles,
  cas d'usage, technologie, TRL, financement, effectif, revenus, roadmap et références.
- Utilise aussi les résultats Algo 1 et Algo 2, mais confronte-les toujours aux données originales.
- Si les colonnes se contredisent sur la maturité, les clients, le TRL ou la commercialisation,
  conserve "inconnu" pour le critère concerné et oriente vers A_VERIFIER si nécessaire.
- Réponds uniquement avec un objet JSON valide.

Données disponibles :
{json.dumps(prompt_row_data, ensure_ascii=False, indent=2)}
""".strip()


# ============================================================
# CLASSIFICATION
# ============================================================

def default_result(message: str = "Données insuffisantes.") -> dict:
    return {
        "algo3_orientation": "A_VERIFIER",
        "algo3_confidence": 0,
        "secondary_orientation": "inconnu",
        "solution_type": "inconnu",
        "maturity_stage": "inconnu",
        "trl_estimate": "inconnu",
        "mrl_estimate": "inconnu",
        "employees_estimate": "inconnu",
        "revenue_arr_estimate": "inconnu",
        "funding_stage": "inconnu",
        "product_commercialized": "inconnu",
        "production_capacity": "inconnu",
        "international_deployment_capacity": "inconnu",
        "field_reference": "inconnu",
        "materials_vertical": "inconnu",
        "orientation_reason": message,
        "missing_information": "Orientation à refaire manuellement.",
    }


def normalize_enum(value: str, allowed: list[str], default: str) -> str:
    normalized = normalize_text(value)
    for item in allowed:
        if normalized == normalize_text(item):
            return item
    return default


def ensure_result_shape(result: dict) -> dict:
    base = default_result()
    if not isinstance(result, dict):
        return base

    base.update(result)

    try:
        confidence = float(base.get("algo3_confidence", 0) or 0)
    except (TypeError, ValueError):
        confidence = 0
    base["algo3_confidence"] = max(0, min(1, confidence))

    base["algo3_orientation"] = normalize_enum(base.get("algo3_orientation", ""), ORIENTATIONS, "A_VERIFIER")
    base["secondary_orientation"] = normalize_enum(
        base.get("secondary_orientation", ""),
        SECONDARY_ORIENTATIONS,
        "inconnu",
    )
    base["solution_type"] = normalize_enum(base.get("solution_type", ""), SOLUTION_TYPES, "inconnu")
    base["maturity_stage"] = normalize_enum(base.get("maturity_stage", ""), MATURITY_STAGES, "inconnu")
    base["trl_estimate"] = normalize_enum(base.get("trl_estimate", ""), TRL_VALUES, "inconnu")
    base["mrl_estimate"] = normalize_enum(base.get("mrl_estimate", ""), MRL_VALUES, "inconnu")
    base["funding_stage"] = normalize_enum(base.get("funding_stage", ""), FUNDING_STAGES, "inconnu")
    base["product_commercialized"] = normalize_enum(base.get("product_commercialized", ""), PRODUCT_VALUES, "inconnu")
    base["production_capacity"] = normalize_enum(base.get("production_capacity", ""), CAPACITY_VALUES, "inconnu")
    base["international_deployment_capacity"] = normalize_enum(
        base.get("international_deployment_capacity", ""),
        CAPACITY_VALUES,
        "inconnu",
    )
    base["field_reference"] = normalize_enum(base.get("field_reference", ""), REFERENCE_VALUES, "inconnu")
    base["materials_vertical"] = normalize_enum(base.get("materials_vertical", ""), YES_NO_UNKNOWN, "inconnu")

    for key in [
        "employees_estimate",
        "revenue_arr_estimate",
        "orientation_reason",
        "missing_information",
    ]:
        base[key] = clean_text(base.get(key, "")) or "inconnu"

    return base


def classify_algo3(row_data: dict) -> dict:
    prompt = build_prompt(row_data)
    schema = algo3_schema()

    for attempt in range(1, MAX_API_RETRIES + 1):
        try:
            result = call_model_json(
                client=get_thread_azure_client(),
                messages=build_messages(prompt),
                schema=schema,
                schema_name="algo3_orientation_programme",
            )
            return ensure_result_shape(result)

        except Exception as exc:
            error_details = format_api_error(exc)

            if is_non_retryable_api_error(exc):
                log(f"Erreur API non réessayée : {error_details}")
                return default_result(f"Erreur pendant l'orientation Algo 3 : {error_details}")

            if attempt == MAX_API_RETRIES:
                log(f"Erreur API finale : {error_details}")
                return default_result(f"Erreur pendant l'orientation Algo 3 : {error_details}")

            retry_after_seconds = get_retry_after_seconds(exc)
            wait_seconds = retry_after_seconds or min(30, (2 ** attempt) + random.uniform(0, 0.5))
            log(
                "Erreur API, nouvel essai dans "
                f"{wait_seconds:.1f}s ({attempt}/{MAX_API_RETRIES}) : {error_details}"
            )
            time.sleep(wait_seconds)

    return default_result("Erreur API inconnue.")


def apply_final_rules(result: dict) -> dict:
    result = ensure_result_shape(result)

    orientation = result.get("algo3_orientation", "A_VERIFIER")
    confidence = float(result.get("algo3_confidence", 0) or 0)
    materials_vertical = result.get("materials_vertical", "inconnu")
    product_commercialized = result.get("product_commercialized", "inconnu")
    production_capacity = result.get("production_capacity", "inconnu")
    international_capacity = result.get("international_deployment_capacity", "inconnu")
    field_reference = result.get("field_reference", "inconnu")
    funding_stage = result.get("funding_stage", "inconnu")
    trl_estimate = result.get("trl_estimate", "inconnu")

    if orientation == "Matériaux" and materials_vertical == "oui" and confidence >= 0.65:
        result["algo3_orientation_finale"] = "Matériaux"
        result["algo3_regle_finale"] = "Orientation Matériaux confirmée."
        return result

    if confidence < 0.65:
        result["algo3_orientation_finale"] = "A_VERIFIER"
        result["algo3_regle_finale"] = "Confiance insuffisante pour orienter automatiquement."
        return result

    if orientation == "Catalyst":
        blockers = []
        if trl_estimate != "9+":
            blockers.append("TRL 9+ non confirmé")
        if product_commercialized != "oui":
            blockers.append("produit commercialisé non confirmé")
        if production_capacity != "oui":
            blockers.append("capacité à produire non confirmée")
        if international_capacity != "oui":
            blockers.append("capacité internationale non confirmée")
        if field_reference != "oui":
            blockers.append("référence terrain non confirmée")

        if blockers:
            result["algo3_orientation_finale"] = "A_VERIFIER"
            result["algo3_regle_finale"] = "Catalyst non confirmé : " + ", ".join(blockers)
            return result

        result["algo3_orientation_finale"] = "Catalyst"
        result["algo3_regle_finale"] = "Critères Catalyst suffisamment confirmés."
        return result

    if orientation == "Seed":
        if trl_estimate == "<7":
            result["algo3_orientation_finale"] = "A_VERIFIER"
            result["algo3_regle_finale"] = "TRL trop amont pour orienter automatiquement vers Seed."
            return result

        if funding_stage == "series_a_or_later":
            result["algo3_orientation_finale"] = "A_VERIFIER"
            result["algo3_regle_finale"] = "Startup potentiellement trop mature pour Seed."
            return result

        result["algo3_orientation_finale"] = "Seed"
        result["algo3_regle_finale"] = "Orientation Seed conservée."
        return result

    if orientation == "Autre":
        result["algo3_orientation_finale"] = "Autre"
        result["algo3_regle_finale"] = "Startup préqualifiée mais hors Seed/Catalyst/Matériaux."
        return result

    result["algo3_orientation_finale"] = "A_VERIFIER"
    result["algo3_regle_finale"] = "Orientation à vérifier manuellement."
    return result


def has_useful_evidence(value: str) -> bool:
    normalized = normalize_text(value)
    if not normalized:
        return False
    if normalized in {"aucune", "inconnu", "none", "nan"}:
        return False
    if "page ciblee detectee" in normalized and "contenu a verifier" in normalized:
        return False
    return True


def apply_evidence_hints(result: dict, row_data: dict) -> dict:
    result = ensure_result_shape(result)

    evidence_customers = row_data.get("maturity_evidence_customers", "")
    evidence_deployment = row_data.get("maturity_evidence_deployment", "")
    evidence_product = row_data.get("maturity_evidence_product", "")
    evidence_production = row_data.get("maturity_evidence_production", "")
    evidence_international = row_data.get("maturity_evidence_international", "")
    evidence_materials = row_data.get("maturity_evidence_materials", "")

    if result.get("product_commercialized") == "inconnu":
        if has_useful_evidence(evidence_customers) or has_useful_evidence(evidence_deployment):
            result["product_commercialized"] = "oui"
        elif has_useful_evidence(evidence_product):
            result["product_commercialized"] = "proche"

    if result.get("field_reference") == "inconnu" and (
        has_useful_evidence(evidence_customers) or has_useful_evidence(evidence_deployment)
    ):
        result["field_reference"] = "oui"

    if result.get("production_capacity") == "inconnu" and has_useful_evidence(evidence_production):
        result["production_capacity"] = "oui"

    if result.get("international_deployment_capacity") == "inconnu" and has_useful_evidence(evidence_international):
        result["international_deployment_capacity"] = "oui"

    if result.get("materials_vertical") != "oui" and (
        has_useful_evidence(evidence_materials)
        or row_data.get("inferred_materials_vertical_by_keywords") == "oui"
    ):
        result["materials_vertical"] = "oui"

    return ensure_result_shape(result)


def employee_maturity_signal(value: str) -> str:
    normalized = normalize_text(value)
    if not normalized or normalized == "inconnu":
        return "inconnu"

    numbers = []
    for match in re.finditer(r"\d{1,5}", normalized.replace(",", "")):
        try:
            numbers.append(int(match.group(0)))
        except ValueError:
            pass
    if not numbers:
        return "inconnu"

    maximum = max(numbers)
    if maximum <= 10:
        return "small"
    if maximum <= 50:
        return "medium"
    return "large"


def calculate_orientation_scores(result: dict, row_data: dict) -> dict[str, float]:
    result = ensure_result_shape(result)

    has_customers = has_useful_evidence(row_data.get("maturity_evidence_customers", ""))
    has_deployment = has_useful_evidence(row_data.get("maturity_evidence_deployment", ""))
    has_product = has_useful_evidence(row_data.get("maturity_evidence_product", ""))
    has_production = has_useful_evidence(row_data.get("maturity_evidence_production", ""))
    has_international = has_useful_evidence(row_data.get("maturity_evidence_international", ""))
    has_materials = has_useful_evidence(row_data.get("maturity_evidence_materials", ""))
    has_funding = has_useful_evidence(row_data.get("maturity_evidence_funding", ""))
    employees_signal = employee_maturity_signal(result.get("employees_estimate", ""))

    catalyst_score = 0.0
    catalyst_score += 0.18 if result.get("trl_estimate") == "9+" else 0
    catalyst_score += 0.16 if result.get("mrl_estimate") == "7+" else 0
    catalyst_score += 0.16 if result.get("product_commercialized") == "oui" or has_product else 0
    catalyst_score += 0.16 if result.get("field_reference") == "oui" or has_customers or has_deployment else 0
    catalyst_score += 0.12 if result.get("production_capacity") == "oui" or has_production else 0
    catalyst_score += 0.10 if result.get("international_deployment_capacity") == "oui" or has_international else 0
    catalyst_score += 0.08 if result.get("funding_stage") == "series_a_or_later" or has_funding else 0
    catalyst_score += 0.04 if employees_signal in {"medium", "large"} else 0

    seed_score = 0.0
    seed_score += 0.22 if result.get("trl_estimate") in {"7-8", "9+"} else 0
    seed_score += 0.16 if result.get("mrl_estimate") in {"3-6", "inconnu"} else 0
    seed_score += 0.16 if result.get("funding_stage") in {"pre-seed", "seed", "before_series_a", "inconnu"} else 0
    seed_score += 0.14 if result.get("product_commercialized") in {"proche", "oui", "inconnu"} else 0
    seed_score += 0.12 if employees_signal in {"small", "medium", "inconnu"} else 0
    seed_score += 0.10 if has_deployment or has_product else 0
    seed_score += 0.10 if result.get("field_reference") in {"souhaitable", "oui", "inconnu"} else 0
    if result.get("funding_stage") == "series_a_or_later":
        seed_score = max(0.0, seed_score - 0.25)

    materials_score = 0.0
    materials_score += 0.55 if result.get("materials_vertical") == "oui" else 0
    materials_score += 0.25 if row_data.get("inferred_materials_vertical_by_keywords") == "oui" else 0
    materials_score += 0.20 if has_materials else 0

    maturity_score = max(seed_score, catalyst_score)

    return {
        "algo3_maturity_score": round(min(1.0, maturity_score), 2),
        "algo3_seed_score": round(min(1.0, seed_score), 2),
        "algo3_catalyst_score": round(min(1.0, catalyst_score), 2),
        "algo3_materials_score": round(min(1.0, materials_score), 2),
    }


def apply_profile_rules(result: dict, text: str) -> dict:
    result = apply_final_rules(result)
    result["dealflow_profile"] = dealflow_profiles.profile_label()
    exclusions = dealflow_profiles.detect_excluded_scope(text)
    result["profile_exclusion_matches"] = dealflow_profiles.format_matches(exclusions)

    if exclusions:
        result["algo3_orientation"] = "A_VERIFIER"
        result["algo3_orientation_finale"] = "A_VERIFIER"
        result["orientation_reason"] = (
            "LATAM profile exclusion guardrail: this company appears related to an excluded scope "
            f"({dealflow_profiles.format_matches(exclusions)}). It should not be oriented automatically."
        )
        result["missing_information"] = ""
        result["algo3_regle_finale"] = "LATAM profile exclusion guardrail applied."

    return result


# ============================================================
# TRAITEMENT EXCEL
# ============================================================

def detect_columns(df: pd.DataFrame) -> dict[str, str | None]:
    columns = {
        key: find_column(df, aliases)
        for key, aliases in COLUMN_ALIASES.items()
    }

    if columns["startup"] is None:
        raise ValueError("Impossible de trouver la colonne du nom de la startup.")

    if columns["algo2_decision_finale"] is None:
        raise ValueError("Impossible de trouver la colonne algo2_decision_finale de l'Algo 2.")

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


def filter_eligible_rows(
    df: pd.DataFrame,
    decision_column: str,
    fallback_decision_column: str | None = None,
) -> pd.DataFrame:
    decisions = df[decision_column].apply(clean_decision)
    eligible_mask = decisions.isin(ELIGIBLE_DECISIONS)

    if fallback_decision_column and fallback_decision_column != decision_column:
        fallback_decisions = df[fallback_decision_column].apply(clean_decision)
        fallback_mask = fallback_decisions.isin(ELIGIBLE_DECISIONS)
        if fallback_mask.any():
            empty_primary = df[decision_column].apply(clean_text).eq("")
            eligible_mask = eligible_mask | (empty_primary & fallback_mask)
            log(f"Fallback activé : {fallback_decision_column} utilisé quand {decision_column} est vide.")

    filtered_df = df[eligible_mask].copy().reset_index(drop=True)

    log(f"Startups éligibles Algo 3 : {len(filtered_df)}/{len(df)}")
    log(f"Filtre appliqué : algo2_decision_finale dans {sorted(ELIGIBLE_DECISIONS)}")

    return filtered_df


def build_row_data(row: pd.Series, columns: dict[str, str | None]) -> dict:
    startup = row_value(row, columns["startup"])
    sector = row_value(row, columns["sector"])
    description = row_value(row, columns["description"])
    website = row_value(row, columns["site"])

    algo2_decision = first_row_value(row, [
        columns["algo2_decision_finale"],
        columns.get("algo2_raw_decision"),
    ])

    full_text = compact_join([
        startup,
        sector,
        description,
        website,
        row_value(row, columns.get("siren_siret")),
        row_value(row, columns.get("algo1_reason")),
        row_value(row, columns.get("algo2_reason")),
        row_value(row, columns.get("b2b_strength")),
        row_value(row, columns.get("confirmed_sector")),
        row_value(row, columns.get("main_issue_matrix")),
        row_value(row, columns.get("matched_issue_matrices")),
        row_value(row, columns.get("issue_alignment_reason")),
        row_value(row, columns.get("sector_alignment_reason")),
    ])

    return {
        "startup": startup,
        "sector": sector,
        "description": description,
        "website": website,
        "siren_siret": row_value(row, columns.get("siren_siret")),
        "algo1_decision_finale": row_value(row, columns.get("algo1_decision")),
        "algo1_reason": row_value(row, columns.get("algo1_reason")),
        "algo2_decision_finale": algo2_decision,
        "algo2_reason": row_value(row, columns.get("algo2_reason")),
        "b2b_strength": row_value(row, columns.get("b2b_strength")),
        "confirmed_sector": row_value(row, columns.get("confirmed_sector")),
        "main_issue_matrix": row_value(row, columns.get("main_issue_matrix")),
        "matched_issue_matrices": row_value(row, columns.get("matched_issue_matrices")),
        "issue_alignment_reason": row_value(row, columns.get("issue_alignment_reason")),
        "sector_alignment_reason": row_value(row, columns.get("sector_alignment_reason")),
        "nombre_employes_existing": row_value(row, columns.get("nombre_employes")),
        "source_nombre_employes_existing": row_value(row, columns.get("source_nombre_employes")),
        "statut_entreprise_existing": row_value(row, columns.get("statut_entreprise")),
        "source_statut_existing": row_value(row, columns.get("source_statut")),
        "niveau_confiance_existing": row_value(row, columns.get("niveau_confiance")),
        "commentaire_verification_existing": row_value(row, columns.get("commentaire_verification")),
        "inferred_solution_type_by_keywords": infer_solution_type(full_text),
        "inferred_materials_vertical_by_keywords": infer_materials_vertical(full_text),
    }


def process_startup_row(
    row: pd.Series,
    columns: dict[str, str | None],
    research_cache: dict[str, dict],
    research_cache_lock: threading.Lock,
) -> dict:
    try:
        row_data = build_row_data(row, columns)
        excel_context_info = build_excel_row_context(row)
        excel_context = excel_context_info["text"]
        excel_context_columns = " | ".join(excel_context_info["used_columns"])
        excel_context_truncated = "oui" if excel_context_info["truncated"] else "non"
        row_data["excel_context"] = excel_context
        row_data["excel_context_columns"] = excel_context_columns
        row_data["excel_context_truncated"] = excel_context_truncated

        inference_text = compact_join([
            row_data.get("startup", ""),
            row_data.get("sector", ""),
            row_data.get("description", ""),
            row_data.get("algo1_reason", ""),
            row_data.get("algo2_reason", ""),
            excel_context,
        ])
        row_data["inferred_solution_type_by_keywords"] = infer_solution_type(inference_text)
        row_data["inferred_materials_vertical_by_keywords"] = infer_materials_vertical(inference_text)

        if not row_data["startup"] and not row_data["description"] and not row_data["website"] and not excel_context:
            return apply_profile_rules(default_result("Ligne vide ou données insuffisantes."), inference_text)

        research = research_startup_website(
            row_data["website"],
            research_cache,
            research_cache_lock,
        )
        row_data["public_research_sources"] = research.get("sources", "")
        row_data["public_external_sources"] = research.get("external_sources", "")
        row_data["public_research_quality"] = research.get("quality", "inconnu")
        row_data["public_research_signals"] = research.get("signals", "")
        row_data["maturity_targeted_sources"] = research.get("targeted_sources", "")
        row_data["maturity_targeted_categories"] = research.get("targeted_categories", "")
        row_data["maturity_targeted_summary"] = research.get("targeted_summary", "")
        row_data["maturity_evidence_employees"] = research.get("evidence_employees", "")
        row_data["maturity_evidence_funding"] = research.get("evidence_funding", "")
        row_data["maturity_evidence_customers"] = research.get("evidence_customers", "")
        row_data["maturity_evidence_product"] = research.get("evidence_product", "")
        row_data["maturity_evidence_deployment"] = research.get("evidence_deployment", "")
        row_data["maturity_evidence_international"] = research.get("evidence_international", "")
        row_data["maturity_evidence_production"] = research.get("evidence_production", "")
        row_data["maturity_evidence_materials"] = research.get("evidence_materials", "")
        row_data["public_research_text"] = research.get("text", "")

        trl_signal = infer_trl_from_text(compact_join([
            row_data.get("startup", ""),
            row_data.get("sector", ""),
            row_data.get("description", ""),
            row_data.get("algo1_reason", ""),
            row_data.get("algo2_reason", ""),
            excel_context,
            row_data.get("public_research_signals", ""),
            row_data.get("maturity_targeted_summary", ""),
            row_data.get("public_research_text", ""),
        ]))
        row_data["trl_signal_estimate"] = trl_signal["estimate"]
        row_data["trl_signal_evidence"] = trl_signal["evidence"]

        company_info = enrich_company_info(row, columns, row_data, research)
        row_data.update(company_info)

        result = classify_algo3(row_data)
        result = apply_evidence_hints(result, row_data)
        if result.get("trl_estimate", "inconnu") == "inconnu" and trl_signal["estimate"] != "inconnu":
            result["trl_estimate"] = trl_signal["estimate"]
        if result.get("employees_estimate", "inconnu") == "inconnu" and company_info.get("nombre_employes"):
            result["employees_estimate"] = company_info["nombre_employes"]
        result = apply_evidence_hints(result, row_data)
        result.update(calculate_orientation_scores(result, row_data))
        result["algo3_solution_type_keyword"] = row_data["inferred_solution_type_by_keywords"]
        result["algo3_materials_keyword"] = row_data["inferred_materials_vertical_by_keywords"]
        result["algo3_trl_signal_estimate"] = trl_signal["estimate"]
        result["algo3_trl_signal_evidence"] = trl_signal["evidence"]
        result["algo3_research_sources"] = compact_join([
            row_data.get("public_research_sources", ""),
            row_data.get("public_external_sources", ""),
        ])
        result["algo3_research_quality"] = row_data.get("public_research_quality", "")
        result["algo3_research_signals"] = row_data.get("public_research_signals", "")
        result["algo3_research_targeted_sources"] = row_data.get("maturity_targeted_sources", "")
        result["algo3_research_targeted_categories"] = row_data.get("maturity_targeted_categories", "")
        result["algo3_research_targeted_summary"] = row_data.get("maturity_targeted_summary", "")
        result["algo3_evidence_employees"] = row_data.get("maturity_evidence_employees", "")
        result["algo3_evidence_funding"] = row_data.get("maturity_evidence_funding", "")
        result["algo3_evidence_customers"] = row_data.get("maturity_evidence_customers", "")
        result["algo3_evidence_product"] = row_data.get("maturity_evidence_product", "")
        result["algo3_evidence_deployment"] = row_data.get("maturity_evidence_deployment", "")
        result["algo3_evidence_international"] = row_data.get("maturity_evidence_international", "")
        result["algo3_evidence_production"] = row_data.get("maturity_evidence_production", "")
        result["algo3_evidence_materials"] = row_data.get("maturity_evidence_materials", "")
        result["algo3_research_text_used"] = row_data.get("public_research_text", "")
        result["algo3_excel_columns_used"] = excel_context_columns
        result["algo3_excel_columns_used_count"] = excel_context_info["used_count"]
        result["algo3_excel_columns_available_count"] = excel_context_info["available_count"]
        result["algo3_excel_context_truncated"] = excel_context_truncated
        result.update(company_info)

        if REQUEST_DELAY_SECONDS > 0:
            time.sleep(REQUEST_DELAY_SECONDS)

        profile_text = compact_join([
            inference_text,
            row_data.get("public_research_signals", ""),
            row_data.get("maturity_targeted_summary", ""),
            row_data.get("public_research_text", ""),
        ])
        return apply_profile_rules(result, profile_text)

    except Exception as exc:
        return apply_profile_rules(default_result(f"Erreur ligne pendant le traitement : {exc}"), "")


def save_results(df_original: pd.DataFrame, results: list[dict], output_file: Path) -> None:
    results_df = pd.DataFrame(results)
    for column in RESULT_COLUMNS:
        if column not in results_df.columns:
            results_df[column] = ""
    if not results_df.empty:
        results_df = results_df[RESULT_COLUMNS]
    partial_df = df_original.iloc[:len(results)].reset_index(drop=True)
    final_df = pd.concat([partial_df, results_df.reset_index(drop=True)], axis=1)
    final_df = sanitize_excel_dataframe(final_df)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="Toutes les startups", index=False)

        for orientation, sheet_name in [
            ("Seed", "Seed"),
            ("Catalyst", "Catalyst"),
            ("Matériaux", "Materiaux"),
            ("Autre", "Autre"),
            ("A_VERIFIER", "A verifier"),
        ]:
            if "algo3_orientation_finale" in final_df.columns:
                sheet_df = final_df[final_df["algo3_orientation_finale"] == orientation]
            else:
                sheet_df = final_df.iloc[0:0]

            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

        if "sauvegarde_temp" not in output_file.stem.lower():
            excel_styling.style_workbook(writer, title="Résultats Algo 3 - Orientation")


def get_saved_result_value(row: pd.Series, result_column: str, original_columns: set[str]):
    if result_column in original_columns:
        possible_columns = [f"{result_column}.{index}" for index in range(1, 6)]
        possible_columns.append(result_column)
    else:
        possible_columns = [result_column]

    for possible_column in possible_columns:
        if possible_column in row.index:
            value = row.get(possible_column)
            if not pd.isna(value):
                return value

    return ""


def load_temp_results(temp_file: Path, df: pd.DataFrame) -> list[dict]:
    if not RESUME_FROM_TEMP or not temp_file.exists():
        return []

    try:
        temp_df = pd.read_excel(temp_file, sheet_name="Toutes les startups")
    except Exception as exc:
        log(f"Impossible de relire la sauvegarde temporaire : {exc}")
        return []

    if temp_df.empty:
        return []

    resume_count = min(len(temp_df), len(df))
    original_columns = set(df.columns)
    temp_df = temp_df.iloc[:resume_count].reset_index(drop=True)

    results = []
    for _, row in temp_df.iterrows():
        result = {
            column: get_saved_result_value(row, column, original_columns)
            for column in RESULT_COLUMNS
        }
        results.append(apply_final_rules(result))

    log(f"Reprise depuis la sauvegarde temporaire : {len(results)} ligne(s) déjà traitée(s).")
    return results


def process_excel_file(input_file: Path) -> None:
    output_file = make_output_path(input_file)
    temp_file = make_temp_path(input_file)

    log(f"Chargement : {input_file.name}")
    df = pd.read_excel(input_file)
    columns = detect_columns(df)

    log("Colonnes détectées :")
    log(f"- Startup : {columns['startup']}")
    log(f"- Secteur : {columns['sector']}")
    log(f"- Description : {columns['description']}")
    log(f"- Site web : {columns['site']}")
    log(f"- Décision Algo 2 : {columns['algo2_decision_finale']}")
    log(f"- Décision Algo 2 fallback : {columns.get('algo2_raw_decision') or 'non trouvée'}")
    log(f"- Colonnes totales du fichier : {len(df.columns)}")

    df = filter_eligible_rows(df, columns["algo2_decision_finale"], columns.get("algo2_raw_decision"))

    if LIMIT_ROWS > 0:
        original_count = len(df)
        df = df.head(LIMIT_ROWS).copy()
        log(f"Mode test : {len(df)}/{original_count} lignes seront traitées.")

    results = load_temp_results(temp_file, df)
    start_index = len(results)
    rows = [row for _, row in df.iloc[start_index:].iterrows()]
    research_cache = {}
    research_cache_lock = threading.Lock()

    if start_index >= len(df):
        log("Toutes les lignes semblent déjà présentes dans la sauvegarde temporaire.")
        save_results(df, results, output_file)
        log(f"Fichier généré : {output_file.name}")
        return

    log(f"Traitement parallèle : {MAX_WORKERS} worker(s)")
    log(f"Contexte Excel utile max : {MAX_EXCEL_CONTEXT_CHARS} caractères par startup")
    if start_index:
        log(f"Reprise à la ligne {start_index + 1}/{len(df)}.")

    executor = ThreadPoolExecutor(max_workers=MAX_WORKERS)
    try:
        mapped_results = executor.map(
            lambda row: process_startup_row(row, columns, research_cache, research_cache_lock),
            rows,
        )

        for result in tqdm(
            mapped_results,
            total=len(df),
            initial=start_index,
            desc=f"Algo 3 {input_file.name}",
        ):
            results.append(result)

            if SAVE_EVERY > 0 and len(results) % SAVE_EVERY == 0:
                save_results(df, results, temp_file)
                log(f"Sauvegarde temporaire : {temp_file.name}")
    except KeyboardInterrupt:
        executor.shutdown(wait=False, cancel_futures=True)
        if results:
            save_results(df, results, temp_file)
            log(
                f"Interruption clavier : sauvegarde temporaire écrite avec "
                f"{len(results)}/{len(df)} ligne(s)."
            )
        else:
            log("Interruption clavier : aucune ligne terminée, donc aucune sauvegarde utile à écrire.")
        raise
    except Exception:
        executor.shutdown(wait=False, cancel_futures=True)
        if results:
            save_results(df, results, temp_file)
            log(
                f"Erreur pendant le traitement : sauvegarde temporaire écrite avec "
                f"{len(results)}/{len(df)} ligne(s)."
            )
        raise
    else:
        executor.shutdown(wait=True)

    save_results(df, results, output_file)
    log(f"Fichier généré : {output_file.name}")


# ============================================================
# SCRIPT PRINCIPAL
# ============================================================

def main() -> None:
    base_dir = Path.cwd()
    input_files = discover_input_files(base_dir, sys.argv[1:])

    if not input_files:
        log("Aucun fichier Excel issu de l'Algo 2 trouvé dans le dossier.")
        log("Tu peux aussi lancer : python Algo3.py chemin\\vers\\fichier_algo_2_affinage_strategique.xlsx")
        return

    log(f"Déploiement Azure OpenAI utilisé : {MODEL_DEPLOYMENT}")
    log(f"Version API Azure OpenAI : {AZURE_OPENAI_API_VERSION}")
    log(f"Mode API : {ALGO3_API_MODE}")
    log(f"Workers parallèles : {MAX_WORKERS}")
    log(f"Décisions Algo 2 traitées : {sorted(ELIGIBLE_DECISIONS)}")
    log(f"Recherche Pappers API : {'activée' if PAPPERS_API_TOKEN else 'désactivée'}")
    if PAPPERS_API_TOKEN and env_bool("ALGO3_PAPPERS_NAME_SEARCH", False):
        log("Recherche Pappers par nom : activée uniquement sur correspondance exacte.")
    log(
        "API officielle Recherche d’Entreprises : "
        f"{'activée' if ENABLE_RECHERCHE_ENTREPRISES_API else 'désactivée'}"
    )
    log(f"Recherche Société.com : {'activée' if ENABLE_SOCIETE_SEARCH else 'désactivée'}")
    if ENABLE_SOCIETE_SEARCH:
        log("Société.com utilisé surtout avec SIREN/lien direct ; la recherche par nom HTML est limitée.")
    log(
        "Recherche effectif LinkedIn public : "
        f"{'activée' if ENABLE_LINKEDIN_EMPLOYEE_SEARCH else 'désactivée'}"
    )
    log(f"Recherche maturite ciblee Algo3 : {'activee' if TARGETED_MATURITY_RESEARCH else 'desactivee'}")
    log(f"Pages recherche max Algo3 : {RESEARCH_MAX_PAGES}")
    log(f"Liens cibles essayes max Algo3 : {RESEARCH_MAX_LINKS_TO_TRY}")
    log(f"{len(input_files)} fichier(s) Algo 2 à traiter.")

    validate_azure_client()

    for input_file in input_files:
        try:
            process_excel_file(input_file)
        except Exception as exc:
            log(f"Erreur sur {input_file.name} : {exc}")

    log("Traitement terminé.")


if __name__ == "__main__":
    main()
