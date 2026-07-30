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
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openai import AzureOpenAI
from tqdm import tqdm

import semantic_prefilter
import learning_memory
import dealflow_profiles


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
ALGO2_API_MODE = os.getenv("ALGO2_API_MODE", os.getenv("ALGO1_API_MODE", "chat")).strip().lower()

OUTPUT_SUFFIX = "_algo_2_affinage_strategique.xlsx"
TEMP_SUFFIX = "_algo_2_affinage_strategique_sauvegarde_temp.xlsx"
SAVE_EVERY = env_int("ALGO2_SAVE_EVERY", 50)
REQUEST_DELAY_SECONDS = env_float("ALGO2_REQUEST_DELAY_SECONDS", 0.1)
MAX_API_RETRIES = env_int("ALGO2_MAX_API_RETRIES", 3)
MAX_WORKERS = max(1, env_int("ALGO2_MAX_WORKERS", 2))
LIMIT_ROWS = env_int("ALGO2_LIMIT_ROWS", 0)
RESUME_FROM_TEMP = env_bool("ALGO2_RESUME_FROM_TEMP", True)
ENABLE_RESEARCH = env_bool("ALGO2_ENABLE_RESEARCH", True)
TARGETED_RESEARCH = env_bool("ALGO2_TARGETED_RESEARCH", True)
RESEARCH_MAX_PAGES = env_int("ALGO2_RESEARCH_MAX_PAGES", 8)
RESEARCH_MAX_LINKS_TO_TRY = env_int("ALGO2_RESEARCH_MAX_LINKS_TO_TRY", max(RESEARCH_MAX_PAGES * 3, RESEARCH_MAX_PAGES))
RESEARCH_TIMEOUT_SECONDS = env_int("ALGO2_RESEARCH_TIMEOUT_SECONDS", 8)
RESEARCH_MAX_TEXT_CHARS = env_int("ALGO2_RESEARCH_MAX_TEXT_CHARS", 10000)
RESEARCH_TEXT_PER_PAGE_CHARS = env_int("ALGO2_RESEARCH_TEXT_PER_PAGE_CHARS", 1800)
MAX_EXCEL_CONTEXT_CHARS = env_int("ALGO2_MAX_EXCEL_CONTEXT_CHARS", 16000)
MAX_EXCEL_CELL_CHARS = env_int("ALGO2_MAX_EXCEL_CELL_CHARS", 3200)
INCLUDE_EXCEL_COLUMN_LIST_IN_PROMPT = env_bool("ALGO2_INCLUDE_EXCEL_COLUMN_LIST_IN_PROMPT", False)
ENABLE_EXCEL_SHORTCUT = env_bool("ALGO2_ENABLE_EXCEL_SHORTCUT", True)
DISABLE_API = env_bool("ALGO2_DISABLE_API", False)
ELIGIBLE_DECISIONS = {
    item.strip().upper()
    for item in os.getenv("ALGO2_ELIGIBLE_DECISIONS", "A_GARDER").split(",")
    if item.strip()
}
NON_RETRYABLE_API_STATUS_CODES = {400, 401, 403, 404}

ALGO1_INPUT_SUFFIXES = [
    "_algo_1_prequalification_sauvegarde_temp.xlsx",
    "_algo_1_prequalification.xlsx",
]

TARGET_SECTORS = [
    "Construction",
    "Built World",
    "Infrastructure",
    "Real Estate",
    "Energy",
    "Mobility",
]

ISSUE_MATRICES = [
    "Productivity",
    "Competitiveness",
    "Environment",
    "Climate",
    "Safety/Security",
]

EDDA_STATUSES = [
    "Préqualifiée — à orienter",
    "Préqualifiée — en attente de classification",
    "À vérifier manuellement",
    "Non préqualifiée",
]

STRENGTH_VALUES = ["fort", "moyen", "faible", "inconnu"]
ALGO2_DECISIONS = ["PREQUALIFIEE", "A_VERIFIER", "NON_PREQUALIFIEE"]
SECTOR_VALUES = TARGET_SECTORS + ["inconnu", "hors_scope"]
RESULT_COLUMNS = [
    "dealflow_profile",
    "profile_exclusion_matches",
    "algo2_decision",
    "algo2_confidence",
    "b2b_strength",
    "b2b_reason",
    "confirmed_sector",
    "sector_alignment_strength",
    "sector_alignment_reason",
    "matched_issue_matrices",
    "main_issue_matrix",
    "issue_alignment_strength",
    "issue_alignment_reason",
    "edda_status",
    "reason",
    "missing_information",
    "algo2_sector_keyword_matches",
    "algo2_issue_keyword_matches",
    "algo2_sector_keyword_categories",
    "algo2_issue_keyword_categories",
    "algo2_research_sources",
    "algo2_research_quality",
    "algo2_research_signals",
    "algo2_research_targeted_sources",
    "algo2_research_targeted_categories",
    "algo2_research_targeted_summary",
    "algo2_research_evidence_clients",
    "algo2_research_evidence_use_cases",
    "algo2_research_evidence_industries",
    "algo2_research_evidence_products",
    "algo2_research_text_used",
    "algo2_b2b_signal_matches",
    "algo2_b2c_risk_matches",
    "algo2_artisan_risk_matches",
    "algo2_exclusion_signal_matches",
    "algo2_vinci_fit_score",
    "algo2_deterministic_review",
    "algo2_excel_columns_used",
    "algo2_excel_columns_used_count",
    "algo2_excel_columns_available_count",
    "algo2_excel_context_truncated",
    "algo2_processing_mode",
    "algo2_api_called",
    "algo2_research_called",
    "algo2_semantic_backend",
    "algo2_semantic_summary",
    "algo2_semantic_sector",
    "algo2_semantic_sector_score",
    "algo2_semantic_issue",
    "algo2_semantic_issue_score",
    "algo2_semantic_b2b_score",
    "algo2_semantic_b2c_score",
    "algo2_learning_match_label",
    "algo2_learning_match_score",
    "algo2_learning_match_source",
    "algo2_decision_finale",
    "algo2_regle_finale",
    "edda_status_final",
]

THREAD_LOCAL = threading.local()

TARGETED_RESEARCH_LINK_CATEGORIES = {
    "clients_customers": {
        "label": "clients/customers",
        "weight": 70,
        "terms": [
            "clients",
            "customers",
            "customer",
            "references",
            "testimonials",
            "trusted by",
            "who we serve",
            "logos clients",
            "customer stories",
            "stories clients",
        ],
    },
    "case_studies": {
        "label": "case studies",
        "weight": 80,
        "terms": [
            "case study",
            "case studies",
            "etude de cas",
            "etudes de cas",
            "cas client",
            "cas clients",
            "customer story",
            "success story",
            "success stories",
            "deployment",
            "deployments",
            "pilot",
            "pilots",
            "proof of concept",
            "poc",
        ],
    },
    "use_cases": {
        "label": "use cases",
        "weight": 65,
        "terms": [
            "use case",
            "use cases",
            "cas d'usage",
            "cas dusage",
            "cas usage",
            "applications",
            "application",
            "workflows",
            "jobs to be done",
            "for operations",
        ],
    },
    "industries_served": {
        "label": "industries served",
        "weight": 60,
        "terms": [
            "industries",
            "industry",
            "sectors",
            "secteurs",
            "markets",
            "verticals",
            "for construction",
            "for energy",
            "for infrastructure",
            "for real estate",
            "for mobility",
        ],
    },
    "solutions_products": {
        "label": "solutions/products",
        "weight": 55,
        "terms": [
            "solution",
            "solutions",
            "product",
            "products",
            "platform",
            "software",
            "saas",
            "technology",
            "technologie",
            "features",
            "fonctionnalites",
            "offering",
        ],
    },
}

COMMON_TARGETED_PATHS = [
    "/solutions",
    "/solution",
    "/products",
    "/product",
    "/platform",
    "/technology",
    "/industries",
    "/industry",
    "/sectors",
    "/markets",
    "/use-cases",
    "/usecases",
    "/applications",
    "/customers",
    "/clients",
    "/references",
    "/case-studies",
    "/case-study",
    "/customer-stories",
    "/success-stories",
    "/fr/solutions",
    "/fr/produits",
    "/fr/industries",
    "/fr/secteurs",
    "/fr/marches",
    "/fr/cas-dusage",
    "/fr/applications",
    "/fr/cas-clients",
    "/fr/etudes-de-cas",
    "/fr/references",
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
    "careers",
    "jobs",
    "investors",
    "press kit",
    "media kit",
    "blog",
    "events",
]


SECTOR_KEYWORDS = {
    "Construction": [
        "construction",
        "chantier",
        "travaux",
        "suivi de travaux",
        "bim",
        "matériaux",
        "materiaux",
        "rénovation",
        "renovation",
        "sécurité terrain",
        "securite terrain",
        "gestion de chantier",
        "pilotage opérationnel",
        "pilotage operationnel",
        "construction site",
        "construction management",
        "worksite",
        "building materials",
    ],
    "Built World": [
        "built world",
        "built environment",
        "smart building",
        "facility",
        "asset management",
        "maintenance",
        "digital twin",
        "jumeau numérique",
        "jumeau numerique",
        "urban tech",
        "building operations",
        "gestion technique bâtiment",
        "gestion technique batiment",
        "gtb",
    ],
    "Infrastructure": [
        "infrastructure",
        "infrastructures",
        "routes",
        "route",
        "ponts",
        "pont",
        "tunnels",
        "tunnel",
        "rail",
        "ouvrages",
        "réseaux",
        "reseaux",
        "génie civil",
        "genie civil",
        "inspection",
        "maintenance infrastructure",
        "infrastructure maintenance",
        "utilities",
        "water infrastructure",
    ],
    "Real Estate": [
        "real estate",
        "immobilier",
        "proptech",
        "property",
        "real estate",
        "gestion d'actifs",
        "gestion actifs",
        "property management",
        "exploitation bâtiment",
        "exploitation batiment",
        "facility management",
        "rénovation énergétique",
        "renovation energetique",
    ],
    "Energy": [
        "energy",
        "énergie",
        "energie",
        "efficacité énergétique",
        "efficacite energetique",
        "renewable energy",
        "solar",
        "battery",
        "grid",
        "smart grid",
        "energy management",
        "décarbonation",
        "decarbonation",
        "photovoltaic",
        "electricity",
        "heat pump",
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
        "trafic",
        "smart mobility",
        "public transport",
        "véhicule électrique",
        "vehicule electrique",
        "last mile",
    ],
}

ISSUE_KEYWORDS = {
    "Productivity": [
        "gain de temps",
        "automatisation",
        "automation",
        "automated",
        "simplification",
        "optimisation",
        "optimization",
        "workflow",
        "processus",
        "productivity",
        "efficiency",
        "plus vite",
        "faster",
        "robot",
        "ai",
        "ia",
    ],
    "Competitiveness": [
        "performance économique",
        "performance economique",
        "roi",
        "coût",
        "cout",
        "cost reduction",
        "quality of service",
        "qualité de service",
        "qualite de service",
        "avantage concurrentiel",
        "différenciation",
        "differenciation",
        "appels d'offres",
        "appel d'offres",
        "business risk",
        "risques commerciaux",
        "margin",
        "marge",
        "tender",
        "bid",
    ],
    "Environment": [
        "impact environnemental",
        "déchets",
        "dechets",
        "waste",
        "consommation",
        "pollution",
        "emissions",
        "émissions",
        "co2",
        "gaz à effet de serre",
        "gaz a effet de serre",
        "carbon footprint",
        "circular",
        "recycling",
        "biodiversity",
        "water consumption",
        "pfas",
        "water management",
    ],
    "Climate": [
        "adaptation climatique",
        "résilience",
        "resilience",
        "décarbonation",
        "decarbonation",
        "climate risk",
        "risques climatiques",
        "net zero",
        "carbon reduction",
        "low carbon",
    ],
    "Safety/Security": [
        "sécurité",
        "securite",
        "security",
        "safety",
        "cyber",
        "données",
        "donnees",
        "data protection",
        "risque opérationnel",
        "risque operationnel",
        "accident",
        "worker safety",
        "site safety",
        "surveillance",
        "compliance",
    ],
}

SECTOR_KEYWORDS["Construction"].extend([
    "contech",
    "construction tech",
    "site monitoring",
    "progress tracking",
    "suivi d'avancement",
    "planning chantier",
    "ordonnancement",
    "qse",
    "qhse",
    "hse",
    "lean construction",
    "supply chain construction",
    "subcontractor management",
    "sous-traitants",
    "engins de chantier",
    "heavy equipment",
    "earthworks",
    "terrassement",
    "excavation",
    "demolition",
    "déconstruction",
    "prefabrication",
    "préfabrication",
    "modular construction",
    "offsite construction",
    "béton bas carbone",
    "low carbon concrete",
    "ciment bas carbone",
    "construction waste",
    "déchets de chantier",
    "quality control construction",
    "inspection chantier",
    "rebar",
    "formwork",
])

SECTOR_KEYWORDS["Built World"].extend([
    "building management system",
    "bms",
    "cmms",
    "iwms",
    "gmao",
    "maintenance prédictive",
    "predictive maintenance",
    "hvac",
    "cvc",
    "occupancy",
    "space management",
    "workplace management",
    "building data",
    "iot building",
    "building sensors",
    "smart facilities",
    "asset operations",
    "building performance",
    "technical building management",
    "exploitation technique",
    "maintenance bâtiment",
    "maintenance batiment",
])

SECTOR_KEYWORDS["Infrastructure"].extend([
    "road infrastructure",
    "pavement",
    "asphalt",
    "chaussée",
    "chaussee",
    "bridge inspection",
    "inspection pont",
    "tunnel inspection",
    "structural health monitoring",
    "geotechnical",
    "géotechnique",
    "water network",
    "wastewater",
    "stormwater",
    "réseau d'eau",
    "reseau d'eau",
    "utilities network",
    "fiber network",
    "telecom infrastructure",
    "airport infrastructure",
    "port infrastructure",
    "rail infrastructure",
    "asset inspection",
    "drone inspection",
    "maintenance ouvrage",
    "ouvrages d'art",
])

SECTOR_KEYWORDS["Real Estate"].extend([
    "commercial real estate",
    "real estate portfolio",
    "asset valuation",
    "property operations",
    "tenant experience",
    "lease management",
    "building compliance",
    "esg real estate",
    "rénovation logement",
    "renovation logement",
    "energy retrofit",
    "retrofit énergétique",
    "retrofit energetique",
    "smart home b2b",
    "property maintenance",
    "bailleur",
    "foncière",
    "fonciere",
    "gestion locative",
])

SECTOR_KEYWORDS["Energy"].extend([
    "energy storage",
    "stockage énergie",
    "stockage energie",
    "grid flexibility",
    "flexibilité réseau",
    "flexibilite reseau",
    "demand response",
    "effacement",
    "microgrid",
    "wind energy",
    "éolien",
    "eolien",
    "hydrogen",
    "hydrogène",
    "hydrogene",
    "heat network",
    "district heating",
    "réseau de chaleur",
    "reseau de chaleur",
    "electrification",
    "energy procurement",
    "energy analytics",
    "smart meter",
    "building energy",
    "scope 1",
    "scope 2",
    "scope 3",
])

SECTOR_KEYWORDS["Mobility"].extend([
    "fleet management",
    "gestion de flotte",
    "charging infrastructure",
    "borne de recharge",
    "route optimization",
    "optimisation de tournée",
    "optimisation de tournee",
    "transport planning",
    "traffic management",
    "parking",
    "tolling",
    "péage",
    "peage",
    "logistics optimization",
    "supply chain transport",
    "warehouse logistics",
    "public transit",
    "rail mobility",
    "shared mobility",
    "autonomous mobility",
    "low emission mobility",
])

ISSUE_KEYWORDS["Productivity"].extend([
    "time saving",
    "gain opérationnel",
    "gain operationnel",
    "réduction des tâches manuelles",
    "reduction des taches manuelles",
    "manual work reduction",
    "process automation",
    "workflow automation",
    "ai assistant",
    "copilot",
    "scheduling",
    "planification",
    "dispatch",
    "resource planning",
    "productivité terrain",
    "productivite terrain",
])

ISSUE_KEYWORDS["Competitiveness"].extend([
    "cost savings",
    "opex",
    "capex",
    "profitability",
    "rentabilité",
    "rentabilite",
    "risk reduction",
    "contract risk",
    "claims management",
    "non-conformity",
    "non conformité",
    "non conformite",
    "quality control",
    "service level",
    "sla",
    "bid management",
    "pricing optimization",
])

ISSUE_KEYWORDS["Environment"].extend([
    "life cycle assessment",
    "lca",
    "analyse cycle de vie",
    "acv",
    "embodied carbon",
    "carbone incorporé",
    "carbone incorpore",
    "material reuse",
    "réemploi",
    "reemploi",
    "resource efficiency",
    "water efficiency",
    "soil pollution",
    "air quality",
    "biodiversité",
    "biodiversite",
    "circular economy",
])

ISSUE_KEYWORDS["Climate"].extend([
    "climate adaptation",
    "climate resilience",
    "heat stress",
    "urban heat island",
    "îlot de chaleur",
    "ilot de chaleur",
    "flood risk",
    "risque inondation",
    "drought",
    "sécheresse",
    "secheresse",
    "carbon removal",
    "carbon capture",
    "carbon accounting",
    "trajectoire carbone",
])

ISSUE_KEYWORDS["Safety/Security"].extend([
    "risk management",
    "risk monitoring",
    "hazard detection",
    "détection danger",
    "detection danger",
    "ppe",
    "epi",
    "worker protection",
    "access control",
    "intrusion",
    "critical infrastructure security",
    "data security",
    "operational safety",
    "cybersecurity",
    "conformité réglementaire",
    "conformite reglementaire",
])

B2B_STRONG_KEYWORDS = [
    "enterprise",
    "enterprises",
    "businesses",
    "companies",
    "corporates",
    "large companies",
    "grands comptes",
    "entreprises",
    "industriels",
    "industrial companies",
    "operators",
    "opérateurs",
    "operateurs",
    "infrastructure operators",
    "asset owners",
    "asset managers",
    "public sector",
    "collectivités",
    "collectivites",
    "municipalities",
    "utilities",
    "contractors",
    "general contractors",
    "construction companies",
    "promoteurs",
    "developers",
    "project developers",
    "epc",
    "facility managers",
    "property managers",
    "energy providers",
    "fleet operators",
    "logistics providers",
    "b2b",
    "saas b2b",
    "platform for companies",
    "solution for professionals",
    "solution pour entreprises",
]

B2C_RISK_KEYWORDS = [
    "consumer",
    "consumers",
    "b2c",
    "grand public",
    "particuliers",
    "homeowners",
    "ménages",
    "menages",
    "families",
    "citizens",
    "citoyens",
    "students",
    "tourists",
    "travelers",
    "app mobile grand public",
    "personal finance",
    "fitness",
    "wellness",
    "beauty",
    "fashion",
    "gaming",
    "dating",
    "food delivery",
    "restaurant booking",
    "e-commerce consumers",
]

ARTISAN_ONLY_RISK_KEYWORDS = [
    "artisans",
    "craftsmen",
    "plombiers",
    "electriciens",
    "électriciens",
    "indépendants",
    "independants",
    "auto entrepreneurs",
    "auto-entrepreneurs",
    "micro entrepreneurs",
    "small trades",
    "tradespeople",
    "home services",
    "services à domicile",
    "services a domicile",
]

EXCLUSION_KEYWORDS = [
    "media",
    "advertising",
    "marketing agency",
    "social media",
    "influencer",
    "hr platform",
    "recruitment only",
    "edtech children",
    "consumer fintech",
    "neobank",
    "crypto trading",
    "insurance broker",
    "foodtech b2c",
    "restaurant",
    "hotel booking",
    "tourism",
    "gaming",
    "entertainment",
    "fashion",
    "beauty",
    "pet care",
    "medical app consumer",
    "pure consulting",
    "agency services",
]

VINCI_FIT_KEYWORDS = [
    "construction",
    "infrastructure",
    "mobility",
    "energy",
    "real estate",
    "built environment",
    "building",
    "civil engineering",
    "génie civil",
    "genie civil",
    "roads",
    "routes",
    "rail",
    "airport",
    "autoroute",
    "concession",
    "facility management",
    "maintenance",
    "asset management",
    "decarbonation",
    "safety",
    "productivity",
    "public works",
    "travaux publics",
    "smart city",
    "urban infrastructure",
    "construction",
    "contech",
    "construction tech",
    "chantier",
    "travaux",
    "suivi de travaux",
    "site monitoring",
    "progress tracking",
    "suivi d'avancement",
    "gestion de chantier",
    "construction management",
    "worksite",
    "construction site",
    "public works",
    "travaux publics",
    "tp",
    "civil works",
    "building site",
    "planning chantier",
    "planification chantier",
    "ordonnancement",
    "pilotage opérationnel",
    "pilotage operationnel",
    "lean construction",
    "qse",
    "qhse",
    "hse",
    "quality control",
    "quality control construction",
    "inspection chantier",
    "site safety",
    "sécurité chantier",
    "securite chantier",
    "sécurité terrain",
    "securite terrain",
    "building materials",
    "matériaux",
    "materiaux",
    "béton",
    "beton",
    "concrete",
    "cement",
    "ciment",
    "béton bas carbone",
    "beton bas carbone",
    "low carbon concrete",
    "ciment bas carbone",
    "formwork",
    "coffrage",
    "rebar",
    "armatures",
    "prefabrication",
    "préfabrication",
    "prefabrication",
    "modular construction",
    "offsite construction",
    "industrialized construction",
    "engins de chantier",
    "heavy equipment",
    "construction equipment",
    "earthworks",
    "terrassement",
    "excavation",
    "demolition",
    "déconstruction",
    "deconstruction",
    "construction waste",
    "déchets de chantier",
    "dechets de chantier",
    "infrastructure",
    "infrastructures",
    "urban infrastructure",
    "civil engineering",
    "génie civil",
    "genie civil",
    "ouvrages",
    "ouvrages d'art",
    "ouvrage d'art",
    "structural engineering",
    "structural health monitoring",
    "asset inspection",
    "inspection ouvrage",
    "maintenance ouvrage",
    "roads",
    "road",
    "routes",
    "route",
    "highway",
    "autoroute",
    "motorway",
    "road infrastructure",
    "pavement",
    "asphalt",
    "chaussée",
    "chaussee",
    "traffic infrastructure",
    "bridge",
    "bridges",
    "pont",
    "ponts",
    "bridge inspection",
    "inspection pont",
    "tunnel",
    "tunnels",
    "tunnel inspection",
    "rail",
    "railway",
    "rail infrastructure",
    "ferroviaire",
    "metro",
    "tramway",
    "station",
    "gare",
    "airport",
    "airports",
    "airport infrastructure",
    "aéroport",
    "aeroport",
    "port",
    "port infrastructure",
    "utilities",
    "networks",
    "réseaux",
    "reseaux",
    "utilities network",
    "water infrastructure",
    "water network",
    "réseau d'eau",
    "reseau d'eau",
    "wastewater",
    "stormwater",
    "assainissement",
    "pipeline",
    "canalisation",
    "fiber network",
    "telecom infrastructure",
    "geotechnical",
    "géotechnique",
    "geotechnique",
    "soil",
    "ground engineering",
    "fondations",
    "foundations",
    "built environment",
    "built world",
    "building",
    "buildings",
    "smart building",
    "connected building",
    "iot building",
    "building operations",
    "building management",
    "building management system",
    "bms",
    "gtb",
    "gestion technique bâtiment",
    "gestion technique batiment",
    "technical building management",
    "facility management",
    "facility",
    "facilities",
    "smart facilities",
    "maintenance",
    "building maintenance",
    "maintenance bâtiment",
    "maintenance batiment",
    "maintenance prédictive",
    "maintenance predictive",
    "predictive maintenance",
    "cmms",
    "gmao",
    "iwms",
    "building performance",
    "building data",
    "building sensors",
    "hvac",
    "cvc",
    "occupancy",
    "space management",
    "workplace management",
    "asset operations",
    "asset performance",
    "asset management",
    "bim",
    "digital twin",
    "jumeau numérique",
    "jumeau numerique",
    "3d model",
    "scan 3d",
    "ifc",
    "maquette numérique",
    "maquette numerique",
    "real estate",
    "immobilier",
    "proptech",
    "property",
    "property management",
    "property operations",
    "real estate portfolio",
    "asset valuation",
    "gestion d'actifs",
    "gestion actifs",
    "gestion locative",
    "lease management",
    "tenant experience",
    "commercial real estate",
    "residential real estate",
    "housing",
    "logement",
    "bailleur",
    "foncière",
    "fonciere",
    "energy",
    "énergie",
    "energie",
    "energy management",
    "energy efficiency",
    "efficacité énergétique",
    "efficacite energetique",
    "energy analytics",
    "building energy",
    "energy monitoring",
    "energy procurement",
    "electricity",
    "électricité",
    "electricite",
    "grid",
    "smart grid",
    "microgrid",
    "grid flexibility",
    "flexibilité réseau",
    "flexibilite reseau",
    "demand response",
    "effacement",
    "smart meter",
    "electrification",
    "renewable energy",
    "renewable",
    "solar",
    "photovoltaic",
    "pv",
    "wind energy",
    "éolien",
    "eolien",
    "battery",
    "energy storage",
    "stockage énergie",
    "stockage energie",
    "hydrogen",
    "hydrogène",
    "hydrogene",
    "heat pump",
    "pompe à chaleur",
    "pompe a chaleur",
    "heat network",
    "district heating",
    "réseau de chaleur",
    "reseau de chaleur",
     "decarbonation",
    "décarbonation",
    "carbon",
    "co2",
    "low carbon",
    "net zero",
    "carbon reduction",
    "carbon footprint",
    "carbon accounting",
    "scope 1",
    "scope 2",
    "scope 3",
     "mobility",
    "mobilité",
    "mobilite",
    "smart mobility",
    "transport",
    "public transport",
    "public transit",
    "transport planning",
    "urban mobility",
    "mobilité urbaine",
    "mobilite urbaine",
    "fleet",
    "flotte",
    "fleet management",
    "gestion de flotte",
    "vehicle",
    "véhicule",
    "vehicule",
    "electric vehicle",
    "véhicule électrique",
    "vehicule electrique",
    "ev charging",
    "charging infrastructure",
    "borne de recharge",
    "recharge",
    "traffic",
    "trafic",
    "traffic management",
    "route optimization",
    "optimisation de tournée",
    "optimisation de tournee",
    "parking",
    "tolling",
    "péage",
    "peage",
    "logistics",
    "logistique",
    "warehouse logistics",
    "supply chain transport",
    "logistics optimization",
    "last mile",
    "dernier kilomètre",
    "dernier kilometre",
    "shared mobility",
    "autonomous mobility",
    "low emission mobility",
    "rail mobility",
    "environment",
    "environnement",
    "environmental impact",
    "impact environnemental",
    "circular economy",
    "économie circulaire",
    "economie circulaire",
    "recycling",
    "recyclage",
    "reuse",
    "réemploi",
    "reemploi",
    "material reuse",
    "waste",
    "déchets",
    "dechets",
    "construction waste",
    "resource efficiency",
    "water management",
    "water consumption",
    "water efficiency",
    "soil pollution",
    "air quality",
    "biodiversity",
    "biodiversité",
    "biodiversite",
    "nature-based solutions",
    "solutions fondées sur la nature",
    "solutions fondees sur la nature",
    "life cycle assessment",
    "lca",
    "analyse cycle de vie",
    "acv",
    "embodied carbon",
    "carbone incorporé",
    "carbone incorpore",
    "climate",
    "climat",
    "climate risk",
    "risques climatiques",
    "adaptation climatique",
    "climate adaptation",
    "resilience",
    "résilience",
    "climate resilience",
    "heat stress",
    "urban heat island",
    "îlot de chaleur",
    "ilot de chaleur",
    "flood risk",
    "risque inondation",
    "drought",
    "sécheresse",
    "secheresse",
    "productivity",
    "productivité",
    "productivite",
    "efficiency",
    "efficacité",
    "efficacite",
    "optimization",
    "optimisation",
    "automation",
    "automatisation",
    "workflow",
    "process automation",
    "workflow automation",
    "time saving",
    "gain de temps",
    "gain opérationnel",
    "gain operationnel",
    "cost reduction",
    "cost savings",
    "réduction des coûts",
    "reduction des couts",
    "roi",
    "opex",
    "capex",
    "performance économique",
    "performance economique",
    "safety",
    "sécurité",
    "securite",
    "site safety",
    "worker safety",
    "worker protection",
    "ppe",
    "epi",
    "hazard detection",
    "détection danger",
    "detection danger",
    "risk management",
    "risk monitoring",
    "operational safety",
    "risque opérationnel",
    "risque operationnel",
    "accident",
    "compliance",
    "conformité",
    "conformite",
    "conformité réglementaire",
    "conformite reglementaire",
    "security",
    "cybersecurity",
    "cyber",
    "data protection",
    "data security",
    "critical infrastructure security",
    "access control",
    "intrusion",
    "surveillance",
    "digital solution",
    "software",
    "saas",
    "platform",
    "plateforme",
    "dashboard",
    "tableau de bord",
    "data",
    "analytics",
    "ai",
    "ia",
    "artificial intelligence",
    "intelligence artificielle",
    "machine learning",
    "computer vision",
    "vision par ordinateur",
    "iot",
    "sensors",
    "capteurs",
    "monitoring",
    "real-time monitoring",
    "robot",
    "robotics",
    "robotique",
    "drone",
    "drone inspection",

]

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
        "target sector",
        "target_sector",
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
    ],
    "algo1_decision": [
        "decision finale",
        "decision_finale",
        "décision finale",
        "algo1 decision",
        "algo 1 decision",
        "decision algo 1",
        "decision",
    ],
    "algo1_raw_decision": [
        "decision",
        "décision",
        "decision modele",
        "decision modèle",
        "model decision",
    ],
    "algo1_reason": [
        "reason",
        "raison",
        "algo1 reason",
        "algo 1 reason",
        "raison algo 1",
        "regle finale",
        "regle_finale",
        "règle finale",
        "sector reason",
        "sector_reason",
        "missing information",
        "missing_information",
    ],
    "algo1_secondary_reason": [
        "sector reason",
        "sector_reason",
        "regle finale",
        "regle_finale",
        "règle finale",
        "missing information",
        "missing_information",
    ],
    "algo1_short_description": [
        "short description clean",
        "short_description_clean",
        "description courte",
        "description propre",
        "clean description",
    ],
    "algo1_b2b": [
        "is b2b",
        "is_b2b",
        "b2b",
        "b2b algo 1",
    ],
    "algo1_sector_relevance": [
        "sector relevance",
        "sector_relevance",
        "pertinence secteur",
        "pertinence sectorielle",
    ],
}

EXCEL_CONTEXT_PRIORITY_TERMS = [
    "decision finale",
    "is b2b",
    "target sector",
    "sector relevance",
    "sector reason",
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
    "algo1 excel columns used",
    "algo1 excel columns used count",
    "algo1 excel columns available count",
    "algo1 excel context truncated",
}

EXCEL_CONTEXT_EXCLUDED_PREFIXES = (
    "algo2 ",
    "algo3 ",
)


# ============================================================
# OUTILS
# ============================================================

def log(message: str) -> None:
    print(f"[Algo 2] {message}")


def strip_accents(text: str) -> str:
    return "".join(
        char
        for char in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(char)
    )


def normalize_text(text: str) -> str:
    text = strip_accents(str(text)).lower()
    text = re.sub(r"[_\-/]+", " ", text)
    text = re.sub(r"[^a-z0-9 +#.&']+", " ", text)
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
    clean_df = df.copy()
    return clean_df.map(sanitize_excel_value)


def clean_decision(text) -> str:
    return normalize_text(text).replace(" ", "_").upper()


def compact_join(values: list[str]) -> str:
    cleaned = [clean_text(value) for value in values if clean_text(value)]
    return " | ".join(cleaned)


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
        if "_algo_1_prequalification" in path.stem:
            candidates.append(path)

    selected_by_base = {}
    for path in candidates:
        base = strip_algo1_suffix(path.stem)
        current = selected_by_base.get(base)
        if current is None:
            selected_by_base[base] = path
            continue
        if current.name.endswith("_sauvegarde_temp.xlsx") and not path.name.endswith("_sauvegarde_temp.xlsx"):
            selected_by_base[base] = path

    return list(selected_by_base.values())


def strip_algo1_suffix(stem: str) -> str:
    for suffix in [
        "_algo_1_prequalification_sauvegarde_temp",
        "_algo_1_prequalification",
    ]:
        if stem.endswith(suffix):
            return stem[: -len(suffix)]
    return stem


def make_output_path(input_file: Path) -> Path:
    base = strip_algo1_suffix(input_file.stem)
    return input_file.with_name(f"{base}{OUTPUT_SUFFIX}")


def make_temp_path(input_file: Path) -> Path:
    base = strip_algo1_suffix(input_file.stem)
    return input_file.with_name(f"{base}{TEMP_SUFFIX}")


def detect_keyword_matches(text: str, keyword_map: dict[str, list[str]]) -> dict[str, list[str]]:
    normalized_text = normalize_text(text)
    matches = {}

    for category, keywords in keyword_map.items():
        matched_keywords = []
        for keyword in keywords:
            normalized_keyword = normalize_text(keyword)
            if normalized_keyword and normalized_keyword in normalized_text:
                matched_keywords.append(keyword)

        if matched_keywords:
            matches[category] = matched_keywords[:8]

    return matches


def detect_flat_keyword_matches(text: str, keywords: list[str], limit: int = 25) -> list[str]:
    normalized_text = normalize_text(text)
    matches = []

    for keyword in keywords:
        normalized_keyword = normalize_text(keyword)
        if normalized_keyword and normalized_keyword in normalized_text and keyword not in matches:
            matches.append(keyword)
        if len(matches) >= limit:
            break

    return matches


def count_nested_keyword_matches(matches: dict[str, list[str]]) -> int:
    return sum(len(values) for values in matches.values())


def score_from_count(count: int, threshold: int) -> float:
    if threshold <= 0:
        return 0
    return round(min(1.0, count / threshold), 2)


def format_keyword_matches(matches: dict[str, list[str]]) -> str:
    if not matches:
        return "Aucun mot-clé détecté automatiquement."

    parts = []
    for category, keywords in matches.items():
        parts.append(f"{category}: {', '.join(keywords)}")
    return " ; ".join(parts)


def matched_categories(matches: dict[str, list[str]]) -> list[str]:
    return list(matches.keys())


def format_flat_matches(matches: list[str]) -> str:
    return ", ".join(matches) if matches else "Aucun signal détecté automatiquement."


def build_deterministic_audit(
    keyword_text: str,
    sector_matches: dict[str, list[str]],
    issue_matches: dict[str, list[str]],
    research_quality: str,
) -> dict:
    b2b_matches = detect_flat_keyword_matches(keyword_text, B2B_STRONG_KEYWORDS)
    b2c_matches = detect_flat_keyword_matches(keyword_text, B2C_RISK_KEYWORDS)
    artisan_matches = detect_flat_keyword_matches(keyword_text, ARTISAN_ONLY_RISK_KEYWORDS)
    exclusion_matches = detect_flat_keyword_matches(keyword_text, EXCLUSION_KEYWORDS)
    vinci_matches = detect_flat_keyword_matches(keyword_text, VINCI_FIT_KEYWORDS)

    sector_category_count = len(sector_matches)
    issue_category_count = len(issue_matches)
    sector_keyword_count = count_nested_keyword_matches(sector_matches)
    issue_keyword_count = count_nested_keyword_matches(issue_matches)

    b2b_score = score_from_count(len(b2b_matches), 5)
    sector_score = min(1.0, round((sector_category_count * 0.25) + (sector_keyword_count * 0.04), 2))
    issue_score = min(1.0, round((issue_category_count * 0.25) + (issue_keyword_count * 0.04), 2))
    vinci_context_score = score_from_count(len(vinci_matches), 5)
    negative_score = min(
        1.0,
        round((len(b2c_matches) * 0.18) + (len(artisan_matches) * 0.16) + (len(exclusion_matches) * 0.22), 2),
    )

    quality_bonus = {
        "forte": 0.12,
        "moyenne": 0.07,
        "faible": 0.02,
        "inaccessible": 0,
        "inconnu": 0,
        "désactivée": 0,
    }.get(research_quality, 0)

    vinci_fit_score = round(
        max(
            0.0,
            min(
                1.0,
                (b2b_score * 0.25)
                + (sector_score * 0.30)
                + (issue_score * 0.25)
                + (vinci_context_score * 0.15)
                + quality_bonus
                - (negative_score * 0.25),
            ),
        ),
        2,
    )

    evidence_gaps = []
    if b2b_score < 0.35:
        evidence_gaps.append("B2B fort peu prouvé")
    if sector_category_count == 0:
        evidence_gaps.append("aucun secteur cible détecté")
    if issue_category_count == 0:
        evidence_gaps.append("aucune matrice d'enjeux détectée")
    if negative_score >= 0.35:
        evidence_gaps.append("signaux B2C/artisan/exclusion à vérifier")
    if research_quality in {"inaccessible", "faible", "désactivée"}:
        evidence_gaps.append("recherche publique insuffisante")

    review = (
        f"score_vinci={vinci_fit_score}; b2b={b2b_score}; secteur={sector_score}; "
        f"enjeux={issue_score}; contexte_vinci={vinci_context_score}; risque_negatif={negative_score}; "
        f"gaps={', '.join(evidence_gaps) if evidence_gaps else 'aucun gap majeur'}"
    )

    return {
        "b2b_matches": b2b_matches,
        "b2c_matches": b2c_matches,
        "artisan_matches": artisan_matches,
        "exclusion_matches": exclusion_matches,
        "vinci_matches": vinci_matches,
        "b2b_score": b2b_score,
        "sector_score": sector_score,
        "issue_score": issue_score,
        "vinci_context_score": vinci_context_score,
        "negative_score": negative_score,
        "vinci_fit_score": vinci_fit_score,
        "evidence_gaps": evidence_gaps,
        "review": review,
    }


def build_excel_shortcut_result_algo2(
    startup_name: str,
    sector: str,
    description: str,
    site: str,
    algo1_decision: str,
    algo1_reason: str,
    algo1_b2b: str,
    algo1_sector_relevance: str,
    excel_context: str,
    excel_context_info: dict,
) -> dict | None:
    if not ENABLE_EXCEL_SHORTCUT:
        return None

    keyword_text = compact_join([
        startup_name,
        sector,
        description,
        site,
        algo1_decision,
        algo1_reason,
        algo1_b2b,
        algo1_sector_relevance,
        excel_context,
    ])
    if len(keyword_text) < 180:
        return None

    sector_matches = detect_keyword_matches(keyword_text, SECTOR_KEYWORDS)
    issue_matches = detect_keyword_matches(keyword_text, ISSUE_KEYWORDS)
    semantic_score = semantic_prefilter.score_algo2(keyword_text)
    deterministic_audit = build_deterministic_audit(
        keyword_text=keyword_text,
        sector_matches=sector_matches,
        issue_matches=issue_matches,
        research_quality="non_requise_excel_suffisant",
    )

    b2b_score = deterministic_audit["b2b_score"]
    sector_score = deterministic_audit["sector_score"]
    issue_score = deterministic_audit["issue_score"]
    negative_score = deterministic_audit["negative_score"]
    vinci_fit_score = deterministic_audit["vinci_fit_score"]
    has_sector = bool(sector_matches)
    has_issue = bool(issue_matches)
    has_exclusion = bool(deterministic_audit["exclusion_matches"])
    semantic_b2b_strong = semantic_prefilter.is_strong(float(semantic_score.get("b2b_score", 0) or 0))
    semantic_b2c_strong = semantic_prefilter.is_strong(float(semantic_score.get("b2c_score", 0) or 0))
    semantic_sector_strong = semantic_prefilter.is_strong(float(semantic_score.get("sector_score", 0) or 0))
    semantic_issue_strong = semantic_prefilter.is_strong(float(semantic_score.get("issue_score", 0) or 0))

    selected_sector = "inconnu"
    if sector_matches:
        selected_sector = max(sector_matches.items(), key=lambda item: len(item[1]))[0]
    elif semantic_sector_strong:
        selected_sector = semantic_score.get("sector", "inconnu")
    matched_issues = matched_categories(issue_matches)
    if not matched_issues and semantic_issue_strong and semantic_score.get("issue"):
        matched_issues = [semantic_score["issue"]]
    main_issue = matched_issues[0] if matched_issues else "inconnu"
    learning_match = learning_memory.predict("algo2", keyword_text)

    result = None
    if learning_match.get("matched") and learning_match.get("label") in {
        "PREQUALIFIEE",
        "A_VERIFIER",
        "NON_PREQUALIFIEE",
    }:
        learned_label = learning_match["label"]
        if learned_label == "PREQUALIFIEE" and (selected_sector == "inconnu" or main_issue == "inconnu"):
            learned_label = "A_VERIFIER"
        result = {
            "algo2_decision": learned_label,
            "algo2_confidence": min(0.92, max(0.76, float(learning_match.get("score", 0) or 0))),
            "b2b_strength": "fort" if learned_label == "PREQUALIFIEE" else "inconnu",
            "b2b_reason": "Décision proposée par mémoire d'apprentissage à partir d'une correction similaire.",
            "confirmed_sector": selected_sector,
            "sector_alignment_strength": "moyen" if selected_sector != "inconnu" else "inconnu",
            "sector_alignment_reason": "Secteur repris depuis l'analyse locale et la mémoire d'apprentissage.",
            "matched_issue_matrices": matched_issues,
            "main_issue_matrix": main_issue,
            "issue_alignment_strength": "moyen" if main_issue != "inconnu" else "inconnu",
            "issue_alignment_reason": "Enjeu repris depuis l'analyse locale et la mémoire d'apprentissage.",
            "edda_status": "Préqualifiée — à orienter" if learned_label == "PREQUALIFIEE" else "À vérifier manuellement",
            "reason": (
                "Décision sans API : mémoire d'apprentissage proche "
                f"({learning_match.get('name') or learning_match.get('source')})."
            ),
            "missing_information": "Vérifier si la correction historique est bien comparable.",
        }
    elif (
        (b2b_score >= 0.4 or semantic_b2b_strong)
        and (has_sector or semantic_sector_strong)
        and (has_issue or semantic_issue_strong)
        and (sector_score >= 0.25 or semantic_sector_strong)
        and (issue_score >= 0.25 or semantic_issue_strong)
        and (vinci_fit_score >= 0.45 or (semantic_b2b_strong and semantic_sector_strong and semantic_issue_strong))
        and negative_score < 0.3
        and not semantic_b2c_strong
        and not has_exclusion
    ):
        result = {
            "algo2_decision": "PREQUALIFIEE",
            "algo2_confidence": 0.82,
            "b2b_strength": "fort",
            "b2b_reason": "Décision sans API : signaux B2B forts déjà présents dans l'Excel et le score sémantique local.",
            "confirmed_sector": selected_sector,
            "sector_alignment_strength": "fort" if sector_score >= 0.45 else "moyen",
            "sector_alignment_reason": "Décision sans API : secteur cible détecté dans les colonnes Excel ou par score sémantique local.",
            "matched_issue_matrices": matched_issues,
            "main_issue_matrix": main_issue,
            "issue_alignment_strength": "fort" if issue_score >= 0.45 else "moyen",
            "issue_alignment_reason": "Décision sans API : au moins une matrice d'enjeux est clairement détectée dans l'Excel ou par score sémantique local.",
            "edda_status": "Préqualifiée — à orienter",
            "reason": "Préqualification sans API : Excel et score sémantique local suffisants sur B2B, secteur et enjeu stratégique.",
            "missing_information": "",
        }
    elif negative_score >= 0.6 and not has_sector and not has_issue and b2b_score < 0.35:
        result = {
            "algo2_decision": "NON_PREQUALIFIEE",
            "algo2_confidence": 0.82,
            "b2b_strength": "faible",
            "b2b_reason": "Décision sans API : signaux négatifs forts et B2B fort non prouvé.",
            "confirmed_sector": "hors_scope",
            "sector_alignment_strength": "faible",
            "sector_alignment_reason": "Aucun secteur cible détecté dans l'Excel.",
            "matched_issue_matrices": [],
            "main_issue_matrix": "inconnu",
            "issue_alignment_strength": "faible",
            "issue_alignment_reason": "Aucune matrice d'enjeux claire détectée dans l'Excel.",
            "edda_status": "Non préqualifiée",
            "reason": "Non-préqualification sans API : Excel suffisamment négatif/hors scope.",
            "missing_information": "",
        }

    if result is None:
        return None

    result["algo2_sector_keyword_matches"] = format_keyword_matches(sector_matches)
    result["algo2_issue_keyword_matches"] = format_keyword_matches(issue_matches)
    result["algo2_sector_keyword_categories"] = ", ".join(matched_categories(sector_matches))
    result["algo2_issue_keyword_categories"] = ", ".join(matched_categories(issue_matches))
    result["algo2_research_sources"] = ""
    result["algo2_research_quality"] = "non_requise_excel_suffisant"
    result["algo2_research_signals"] = ""
    result["algo2_research_targeted_sources"] = ""
    result["algo2_research_targeted_categories"] = ""
    result["algo2_research_targeted_summary"] = ""
    result["algo2_research_evidence_clients"] = ""
    result["algo2_research_evidence_use_cases"] = ""
    result["algo2_research_evidence_industries"] = ""
    result["algo2_research_evidence_products"] = ""
    result["algo2_research_text_used"] = ""
    result["algo2_b2b_signal_matches"] = format_flat_matches(deterministic_audit["b2b_matches"])
    result["algo2_b2c_risk_matches"] = format_flat_matches(deterministic_audit["b2c_matches"])
    result["algo2_artisan_risk_matches"] = format_flat_matches(deterministic_audit["artisan_matches"])
    result["algo2_exclusion_signal_matches"] = format_flat_matches(deterministic_audit["exclusion_matches"])
    result["algo2_vinci_fit_score"] = deterministic_audit["vinci_fit_score"]
    result["algo2_deterministic_review"] = deterministic_audit["review"]
    result["algo2_excel_columns_used"] = " | ".join(excel_context_info["used_columns"])
    result["algo2_excel_columns_used_count"] = excel_context_info["used_count"]
    result["algo2_excel_columns_available_count"] = excel_context_info["available_count"]
    result["algo2_excel_context_truncated"] = "oui" if excel_context_info["truncated"] else "non"
    result["algo2_processing_mode"] = "excel_shortcut_no_api_no_research"
    result["algo2_api_called"] = "non"
    result["algo2_research_called"] = "non"
    result["algo2_semantic_backend"] = semantic_score.get("backend", "")
    result["algo2_semantic_summary"] = semantic_prefilter.summarize(semantic_score)
    result["algo2_semantic_sector"] = semantic_score.get("sector", "")
    result["algo2_semantic_sector_score"] = semantic_score.get("sector_score", 0)
    result["algo2_semantic_issue"] = semantic_score.get("issue", "")
    result["algo2_semantic_issue_score"] = semantic_score.get("issue_score", 0)
    result["algo2_semantic_b2b_score"] = semantic_score.get("b2b_score", 0)
    result["algo2_semantic_b2c_score"] = semantic_score.get("b2c_score", 0)
    result["algo2_learning_match_label"] = learning_match.get("label", "")
    result["algo2_learning_match_score"] = learning_match.get("score", 0)
    result["algo2_learning_match_source"] = learning_match.get("source", "")
    return apply_profile_rules(result, keyword_text)


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


def get_thread_http_session() -> requests.Session:
    session = getattr(THREAD_LOCAL, "http_session", None)
    if session is None:
        session = requests.Session()
        THREAD_LOCAL.http_session = session
    return session


def fetch_page_text_and_links(url: str) -> tuple[str, list[dict[str, str]]]:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (compatible; LeonardAlgo2Research/1.0; "
            "strategic startup filtering)"
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

    body_text = " ".join(
        tag.get_text(" ", strip=True)
        for tag in soup.find_all(["p", "li"])[:100]
    )

    text = clean_text(f"{title}. {meta_description}. {headings}. {body_text}")
    return text, links


def targeted_categories_from_text(text: str) -> list[str]:
    normalized = normalize_text(text)
    categories = []
    for category, config in TARGETED_RESEARCH_LINK_CATEGORIES.items():
        for term in config["terms"]:
            if normalize_text(term) in normalized:
                categories.append(category)
                break
    return categories


def format_targeted_category_labels(categories: list[str]) -> str:
    labels = []
    for category in categories:
        if category == "homepage":
            label = "homepage"
        else:
            label = TARGETED_RESEARCH_LINK_CATEGORIES.get(category, {}).get("label", category)
        if label not in labels:
            labels.append(label)
    return ", ".join(labels) if labels else "non classe"


def make_research_link(url: str, link_text: str = "", source: str = "site") -> dict:
    clean_url = normalize_url(url).rstrip("/")
    combined = f"{clean_url} {link_text}"
    categories = targeted_categories_from_text(combined)
    score = research_link_score(clean_url, link_text)
    if source == "homepage_link":
        score += 30
    elif source == "common_path":
        score -= 10
    return {
        "url": clean_url,
        "text": clean_text(link_text),
        "source": source,
        "categories": categories,
        "score": score,
    }


def research_link_score(url: str, link_text: str = "") -> int:
    normalized = normalize_text(f"{url} {link_text}")
    priority_terms = {
        "about": 45,
        "company": 25,
        "team": 25,
        "solution": 45,
        "solutions": 45,
        "product": 40,
        "platform": 30,
        "technology": 35,
        "industries": 40,
        "industry": 40,
        "customers": 55,
        "clients": 55,
        "case studies": 60,
        "case study": 60,
        "success stories": 55,
        "use cases": 45,
        "partners": 30,
        "press": 25,
        "news": 20,
        "blog": 10,
        "security": 30,
        "safety": 30,
        "sustainability": 35,
        "climate": 35,
        "decarbon": 35,
        "energy": 30,
        "construction": 30,
        "infrastructure": 30,
        "mobility": 30,
        "real estate": 30,
        "building": 25,
    }

    score = 0
    if TARGETED_RESEARCH:
        for category, config in TARGETED_RESEARCH_LINK_CATEGORIES.items():
            if category in targeted_categories_from_text(normalized):
                score += int(config["weight"])

    for term, weight in priority_terms.items():
        if term in normalized:
            score += weight

    for term in LOW_VALUE_RESEARCH_LINK_TERMS:
        if normalize_text(term) in normalized:
            score -= 45

    return score


def common_targeted_candidates(base_url: str) -> list[dict]:
    parsed = urlparse(normalize_url(base_url))
    if not parsed.scheme or not parsed.netloc:
        return []

    root_url = f"{parsed.scheme}://{parsed.netloc}"
    candidates = []
    for path in COMMON_TARGETED_PATHS:
        candidates.append(make_research_link(urljoin(root_url, path), path, source="common_path"))
    return candidates


def normalized_link_record(link) -> dict | None:
    if isinstance(link, dict):
        raw_url = link.get("url", "")
        link_text = link.get("text", "")
    else:
        raw_url = str(link)
        link_text = ""

    if not raw_url:
        return None

    parsed = urlparse(normalize_url(raw_url))
    clean_url = parsed._replace(query="", fragment="").geturl().rstrip("/")
    if not clean_url:
        return None

    return make_research_link(clean_url, link_text, source="homepage_link")


def select_research_links(base_url: str, links: list[dict[str, str]]) -> list[dict]:
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
        record = normalized_link_record(link)
        if not record:
            continue
        clean_link = record["url"]
        if not same_domain(clean_link, base_url):
            continue
        if clean_link in seen:
            continue
        score = record["score"]
        if score <= 0:
            continue
        current = candidates_by_url.get(clean_link)
        if current is None or score > current["score"]:
            candidates_by_url[clean_link] = record

    for record in common_targeted_candidates(base_url):
        clean_link = record["url"]
        if clean_link in seen or not same_domain(clean_link, base_url):
            continue
        score = record["score"]
        if score <= 0:
            continue
        current = candidates_by_url.get(clean_link)
        if current is None or score > current["score"]:
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

    if TARGETED_RESEARCH:
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


def first_regex_matches(text: str, patterns: list[str], limit: int = 8) -> list[str]:
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
    sector_matches = detect_keyword_matches(text, SECTOR_KEYWORDS)
    issue_matches = detect_keyword_matches(text, ISSUE_KEYWORDS)
    b2b_signal_matches = detect_flat_keyword_matches(text, B2B_STRONG_KEYWORDS)
    b2c_risk_matches = detect_flat_keyword_matches(text, B2C_RISK_KEYWORDS)
    artisan_risk_matches = detect_flat_keyword_matches(text, ARTISAN_ONLY_RISK_KEYWORDS)
    exclusion_matches = detect_flat_keyword_matches(text, EXCLUSION_KEYWORDS)
    vinci_context_matches = detect_flat_keyword_matches(text, VINCI_FIT_KEYWORDS)

    b2b_terms = first_regex_matches(text, [
        r"\b(?:enterprise|businesses|companies|industrials|operators|contractors|developers|asset managers|public sector|local authorities|utilities)\b",
        r"\b(?:entreprises|industriels|opérateurs|operateurs|constructeurs|promoteurs|collectivités|collectivites|acteurs publics|professionnels)\b",
    ])
    b2c_terms = first_regex_matches(text, [
        r"\b(?:consumers|individuals|homeowners|households|families|tourists|students)\b",
        r"\b(?:particuliers|ménages|menages|familles|grand public|citoyens)\b",
    ])
    customer_terms = first_regex_matches(text, [
        r"\b(?:customers|clients|trusted by|case studies|case study|customer story|success story|deployed with|used by)\b",
        r"\b(?:clients|références|references|cas client|cas d'usage|déployé chez|deploiement chez)\b",
    ])
    safety_terms = first_regex_matches(text, [
        r"\b(?:safety|security|cybersecurity|worker safety|site safety|risk prevention|compliance)\b",
        r"\b(?:sécurité|securite|cybersécurité|cybersecurite|prévention des risques|prevention des risques)\b",
    ])
    environment_terms = first_regex_matches(text, [
        r"\b(?:co2|carbon|decarbonization|decarbonisation|emissions|waste|recycling|circular economy|energy efficiency)\b",
        r"\b(?:décarbonation|decarbonation|émissions|emissions|déchets|dechets|recyclage|économie circulaire|economie circulaire)\b",
    ])

    explicit_negative = []
    if any(term in normalized for term in ["not for businesses", "not b2b", "consumers only", "particuliers uniquement"]):
        explicit_negative.append("Signal négatif B2B explicite")

    signals = {
        "b2b_terms": ", ".join(b2b_terms) or "inconnu",
        "b2c_terms": ", ".join(b2c_terms) or "inconnu",
        "customer_reference_terms": ", ".join(customer_terms) or "inconnu",
        "b2b_signal_matches": format_flat_matches(b2b_signal_matches),
        "b2c_risk_matches": format_flat_matches(b2c_risk_matches),
        "artisan_only_risk_matches": format_flat_matches(artisan_risk_matches),
        "exclusion_signal_matches": format_flat_matches(exclusion_matches),
        "vinci_context_matches": format_flat_matches(vinci_context_matches),
        "sector_matches": format_keyword_matches(sector_matches),
        "issue_matches": format_keyword_matches(issue_matches),
        "safety_security_terms": ", ".join(safety_terms) or "inconnu",
        "environment_climate_terms": ", ".join(environment_terms) or "inconnu",
        "explicit_negative_signals": ", ".join(explicit_negative) or "inconnu",
    }

    return signals


def format_research_signals(signals: dict[str, str]) -> str:
    return " | ".join(f"{key}: {value}" for key, value in signals.items())


def unique_values(values: list[str]) -> list[str]:
    unique = []
    for value in values:
        value = clean_text(value)
        if value and value not in unique:
            unique.append(value)
    return unique


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


def terms_for_targeted_categories(categories: list[str]) -> list[str]:
    terms = []
    for category in categories:
        config = TARGETED_RESEARCH_LINK_CATEGORIES.get(category)
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
        snippet = sentence[:260].rstrip()
        if len(sentence) > 260:
            snippet += "..."
        if snippet not in snippets:
            snippets.append(snippet)
        if len(snippets) >= limit:
            break

    return snippets


def format_targeted_page_records(page_records: list[dict]) -> str:
    parts = []
    for record in page_records:
        categories = [category for category in record.get("categories", []) if category != "homepage"]
        if not categories:
            continue
        parts.append(f"{format_targeted_category_labels(categories)}: {record.get('url', '')}")
    return " | ".join(unique_values(parts))


def build_targeted_research_evidence(page_records: list[dict]) -> dict:
    groups = {
        "clients": ["clients_customers", "case_studies"],
        "use_cases": ["use_cases", "case_studies"],
        "industries": ["industries_served"],
        "products": ["solutions_products"],
    }
    evidence = {key: [] for key in groups}

    for record in page_records:
        text = record.get("text", "")
        categories = record.get("categories", [])
        url = record.get("url", "")
        for group_name, group_categories in groups.items():
            if not any(category in categories for category in group_categories):
                continue
            terms = terms_for_targeted_categories(group_categories)
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


def build_targeted_research_summary(page_records: list[dict], evidence: dict) -> str:
    targeted_pages = format_targeted_page_records(page_records)
    categories = []
    for record in page_records:
        categories.extend([category for category in record.get("categories", []) if category != "homepage"])
    category_labels = format_targeted_category_labels(unique_values(categories))
    parts = [
        f"Pages ciblees: {targeted_pages or 'aucune page ciblee accessible'}",
        f"Categories ciblees: {category_labels if categories else 'aucune'}",
        f"Preuves clients/customers: {evidence.get('clients') or 'aucune'}",
        f"Preuves use cases/case studies: {evidence.get('use_cases') or 'aucune'}",
        f"Preuves industries served: {evidence.get('industries') or 'aucune'}",
        f"Preuves solutions/products: {evidence.get('products') or 'aucune'}",
    ]
    return " || ".join(parts)


def research_quality_from_sources(sources: list[str], text: str, targeted_categories: list[str] | None = None) -> str:
    if not sources:
        return "inaccessible"
    targeted_categories = targeted_categories or []
    unique_targeted_categories = {
        category
        for category in targeted_categories
        if category and category != "homepage"
    }
    if len(unique_targeted_categories) >= 3 and len(text) >= 2500:
        return "forte"
    if len(unique_targeted_categories) >= 2 and len(text) >= 1500:
        return "moyenne"
    if len(text) < 700:
        return "faible"
    if len(sources) >= 4 and len(text) >= 3500:
        return "forte"
    return "moyenne"


def research_startup_website(site: str, cache: dict[str, dict], cache_lock: threading.Lock) -> dict:
    if not ENABLE_RESEARCH or not is_valid_url(site):
        empty_signals = detect_research_signals("")
        return {
            "sources": "",
            "quality": "désactivée" if ENABLE_RESEARCH is False else "inconnu",
            "text": "",
            "signals": format_research_signals(empty_signals),
            "targeted_sources": "",
            "targeted_categories": "",
            "targeted_summary": "",
            "evidence_clients": "",
            "evidence_use_cases": "",
            "evidence_industries": "",
            "evidence_products": "",
        }

    base_url = normalize_url(site).rstrip("/")

    with cache_lock:
        cached = cache.get(base_url)
        if cached:
            return cached

    homepage_text, links = fetch_page_text_and_links(base_url)
    selected_links = select_research_links(base_url, links)

    texts = []
    sources = []
    page_records = []
    targeted_categories = []

    for link_record in selected_links:
        if sum(len(item) for item in texts) >= RESEARCH_MAX_TEXT_CHARS:
            break
        link = link_record.get("url", "")
        if not link:
            continue
        if link == base_url:
            page_text = homepage_text
        else:
            page_text, _ = fetch_page_text_and_links(link)
        if not page_text:
            continue

        categories = unique_values(
            list(link_record.get("categories", []))
            + targeted_categories_from_text(f"{link} {link_record.get('text', '')} {page_text[:1200]}")
        )
        if link == base_url and not categories:
            categories = ["homepage"]

        page_text = truncate_research_page_text(page_text)
        category_labels = format_targeted_category_labels(categories)
        texts.append(f"Source: {link}\nType de page: {category_labels}\n{page_text}")
        sources.append(link)
        page_records.append({
            "url": link,
            "categories": categories,
            "text": page_text,
        })
        targeted_categories.extend([category for category in categories if category != "homepage"])
        if len(sources) >= RESEARCH_MAX_PAGES:
            break

    research_text = clean_text("\n\n".join(texts))[:RESEARCH_MAX_TEXT_CHARS]
    signals = detect_research_signals(research_text)
    evidence = build_targeted_research_evidence(page_records)
    targeted_summary = build_targeted_research_summary(page_records, evidence)
    targeted_sources = format_targeted_page_records(page_records)
    result = {
        "sources": " | ".join(sources),
        "quality": research_quality_from_sources(sources, research_text, targeted_categories),
        "text": research_text,
        "signals": format_research_signals(signals),
        "targeted_sources": targeted_sources,
        "targeted_categories": format_targeted_category_labels(unique_values(targeted_categories)),
        "targeted_summary": targeted_summary,
        "evidence_clients": evidence.get("clients", ""),
        "evidence_use_cases": evidence.get("use_cases", ""),
        "evidence_industries": evidence.get("industries", ""),
        "evidence_products": evidence.get("products", ""),
    }

    with cache_lock:
        cache[base_url] = result

    return result


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
                "Tu es un analyste innovation rigoureux pour VINCI. "
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
    mode = ALGO2_API_MODE

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
            schema_name="algo2_healthcheck",
        )
        log("Connexion Azure OpenAI OK.")

    except Exception as exc:
        raise RuntimeError(f"Test Azure OpenAI échoué : {format_api_error(exc)}")


# ============================================================
# PROMPT ET SCHEMA
# ============================================================

def algo2_schema() -> dict:
    return {
        "type": "object",
        "properties": {
            "algo2_decision": {
                "type": "string",
                "enum": ALGO2_DECISIONS,
            },
            "algo2_confidence": {
                "type": "number",
                "description": "Score de confiance entre 0 et 1.",
            },
            "b2b_strength": {
                "type": "string",
                "enum": STRENGTH_VALUES,
            },
            "b2b_reason": {
                "type": "string",
            },
            "confirmed_sector": {
                "type": "string",
                "enum": SECTOR_VALUES,
            },
            "sector_alignment_strength": {
                "type": "string",
                "enum": STRENGTH_VALUES,
            },
            "sector_alignment_reason": {
                "type": "string",
            },
            "matched_issue_matrices": {
                "type": "array",
                "items": {
                    "type": "string",
                    "enum": ISSUE_MATRICES,
                },
            },
            "main_issue_matrix": {
                "type": "string",
                "enum": ISSUE_MATRICES + ["inconnu"],
            },
            "issue_alignment_strength": {
                "type": "string",
                "enum": STRENGTH_VALUES,
            },
            "issue_alignment_reason": {
                "type": "string",
            },
            "edda_status": {
                "type": "string",
                "enum": EDDA_STATUSES,
            },
            "reason": {
                "type": "string",
            },
            "missing_information": {
                "type": "string",
            },
        },
        "required": [
            "algo2_decision",
            "algo2_confidence",
            "b2b_strength",
            "b2b_reason",
            "confirmed_sector",
            "sector_alignment_strength",
            "sector_alignment_reason",
            "matched_issue_matrices",
            "main_issue_matrix",
            "issue_alignment_strength",
            "issue_alignment_reason",
            "edda_status",
            "reason",
            "missing_information",
        ],
        "additionalProperties": False,
    }


def build_prompt(
    startup_name: str,
    sector: str,
    description: str,
    site: str,
    algo1_decision: str,
    algo1_reason: str,
    algo1_b2b: str,
    algo1_sector_relevance: str,
    sector_keyword_matches: dict[str, list[str]],
    issue_keyword_matches: dict[str, list[str]],
    research_sources: str = "",
    research_quality: str = "",
    research_signals: str = "",
    research_targeted_summary: str = "",
    research_text: str = "",
    deterministic_review: str = "",
    b2b_signal_matches: str = "",
    b2c_risk_matches: str = "",
    artisan_risk_matches: str = "",
    exclusion_signal_matches: str = "",
    vinci_fit_score: str = "",
    excel_context: str = "",
    excel_context_columns: str = "",
    excel_context_truncated: str = "non",
) -> str:
    target_sectors = ", ".join(TARGET_SECTORS)
    issue_matrices = ", ".join(ISSUE_MATRICES)
    profile_context = dealflow_profiles.profile_prompt_context()

    return f"""
Tu es un analyste innovation pour Leonard, la plateforme d'innovation et de prospective du groupe VINCI.

Tu construis l'Algo 2, un filtre d'affinage stratégique après l'Algo 1.
Important : l'Algo 2 n'est pas un enrichissement de données. Il vérifie si la startup est suffisamment
alignée avec les enjeux métiers et stratégiques de VINCI pour mériter une qualification approfondie
et une intégration dans Edda avec un statut intermédiaire.

Profil de traitement :
{profile_context}

Sorties possibles :
- PREQUALIFIEE
- A_VERIFIER
- NON_PREQUALIFIEE

Statuts Edda possibles :
- Préqualifiée — à orienter
- Préqualifiée — en attente de classification
- À vérifier manuellement
- Non préqualifiée

Analyse à réaliser :

1. B2B fort
La startup doit s'adresser principalement à des entreprises, grands comptes, industriels,
opérateurs d'infrastructures, acteurs publics ou parapublics, ou professionnels du bâtiment,
de l'énergie, de la mobilité ou de l'immobilier.
Important : les artisans ne sont pas un signal suffisant de B2B fort.
Une startup principalement B2C ou orientée particuliers doit être NON_PREQUALIFIEE ou A_VERIFIER
selon l'incertitude.
Un B2B fort doit être prouvé par au moins un des éléments suivants :
- clients entreprises, grands comptes, industriels, collectivités, opérateurs, asset owners ;
- cas clients, références, pilotes terrain ou déploiements B2B ;
- solution vendue à une fonction métier professionnelle : opérations, maintenance, chantier,
  énergie, exploitation, sécurité, finance projet, immobilier, mobilité, infrastructure ;
- vocabulaire de vente B2B : enterprise, platform, SaaS B2B, API, procurement, operations,
  asset management, facility management, fleet management.
Attention : "professionnels", "artisans", "indépendants" ou "pros" seuls ne suffisent pas.

2. Alignement sectoriel
Les seuls secteurs cibles sont : {target_sectors}.
Tu dois analyser le sens réel de l'activité, pas seulement les mots-clés exacts.
Built World est un secteur cible au même niveau que les autres.
Tu dois vérifier les sous-cas métiers :
- Construction : chantier, travaux, planning, QHSE, engins, sous-traitants, BIM, matériaux,
  déchets de chantier, préfabrication, béton/ciment, contrôle qualité.
- Built World : exploitation bâtiment, facility, maintenance, GTB/BMS, capteurs, jumeau numérique,
  performance bâtiment, asset operations.
- Infrastructure : routes, ponts, tunnels, rail, réseaux, eau, ouvrage d'art, inspection,
  maintenance, génie civil, aéroport, port, concessions.
- Real Estate : proptech, property management, asset management immobilier, bailleurs,
  rénovation énergétique, ESG immobilier, exploitation de parc.
- Energy : efficacité énergétique, stockage, grid, flexibilité, solaire, éolien, hydrogène,
  réseau de chaleur, électrification, décarbonation.
- Mobility : flotte, recharge EV, trafic, logistique, transport public, péage, parking,
  optimisation de tournées, mobilité bas carbone.

3. Alignement avec au moins une matrice d'enjeux
Matrices autorisées : {issue_matrices}.
- Productivity : gain de temps, automatisation, simplification, optimisation d'un processus.
- Competitiveness : performance économique, qualité de service, avantage concurrentiel,
  différenciation, capacité à répondre à plus d'appels d'offres, limitation des risques business.
- Environment : réduction de l'impact environnemental, déchets, consommations, pollution,
  émissions CO2 ou gaz à effet de serre.
- Climate : adaptation climatique, résilience, décarbonation, réponse aux risques climatiques.
- Safety/Security : sécurité des personnes, des sites, des infrastructures, des données
  ou des opérations.
L'enjeu doit être concret et exploitable pour VINCI. Un discours marketing vague sur
"innovation", "impact", "IA" ou "plateforme" ne suffit pas.

4. Signaux négatifs à surveiller
Tu dois dégrader vers A_VERIFIER ou NON_PREQUALIFIEE si l'activité est principalement :
- application grand public, bien-être, mode, food, tourisme, gaming, réseau social ;
- fintech/assurance/HR/marketing générique sans lien clair avec les métiers VINCI ;
- conseil pur ou agence sans produit scalable ;
- marketplace principalement B2C ;
- solution seulement pour artisans ou indépendants sans preuve de grands comptes ou opérateurs.

Règles de décision :

PREQUALIFIEE :
- B2B fort ;
- liée clairement à au moins un secteur cible ;
- alignée clairement avec au moins une matrice d'enjeux.

A_VERIFIER :
- B2B probable mais pas certain ;
- secteur cible possible mais pas clair ;
- enjeu stratégique possible mais pas explicite ;
- description trop vague ;
- données insuffisantes.

NON_PREQUALIFIEE :
- pas de B2B fort ;
- ou pas de lien clair avec les secteurs cibles ;
- ou aucune matrice d'enjeux claire.

Règles de prudence :
- Ne devine pas.
- Si une information manque, indique "inconnu".
- Si tu hésites entre PREQUALIFIEE et A_VERIFIER, choisis A_VERIFIER.
- Utilise la recherche publique ci-dessous comme un ensemble de preuves, mais ne transforme pas
  un simple mot-clé isolé en preuve forte.
- Pour les startups issues de Algo 1 = A_VERIFIER, tu peux les préqualifier uniquement si la recherche
  apporte des preuves claires sur B2B fort, secteur cible et matrice d'enjeux.
- Si la recherche site est inaccessible ou faible, reste prudent.
- Si le score déterministe VINCI est faible mais que tu penses préqualifier, explique précisément
  quelles preuves compensent ce score.
- Si plusieurs signaux négatifs sont détectés, ne mets PREQUALIFIEE que si les preuves positives
  sont très fortes.
- Utilise toutes les informations Excel complémentaires, notamment clients cibles, cas d'usage,
  chaîne de valeur, technologie, pain points, références, verticale, TRL et réponses de candidature.
- Les résultats d'Algo 1 restent des indices : confronte-les aux données originales et à la recherche.
- En cas de contradiction entre colonnes, choisis A_VERIFIER et explique la contradiction.
- matched_issue_matrices doit être une liste, vide si aucune matrice n'est confirmée.
- main_issue_matrix doit être "inconnu" si aucune matrice n'est confirmée.
- Réponds uniquement avec un objet JSON valide.

Données issues du fichier Algo 1 :

Nom startup :
{startup_name}

Secteur disponible :
{sector}

Description disponible :
{description}

Site web :
{site}

Décision finale Algo 1 :
{algo1_decision}

Raison Algo 1 :
{algo1_reason}

B2B Algo 1 :
{algo1_b2b}

Pertinence sectorielle Algo 1 :
{algo1_sector_relevance}

Toutes les informations Excel utiles disponibles :
{excel_context or "aucune information complémentaire"}

Colonnes Excel utilisées :
{excel_context_columns or "aucune"}

Contexte Excel tronqué par sécurité :
{excel_context_truncated}

Mots-clés sectoriels détectés automatiquement :
{format_keyword_matches(sector_keyword_matches)}

Mots-clés matrices d'enjeux détectés automatiquement :
{format_keyword_matches(issue_keyword_matches)}

Pré-analyse déterministe détaillée :
{deterministic_review or "aucune"}

Score déterministe d'alignement VINCI :
{vinci_fit_score or "inconnu"}

Signaux B2B forts détectés :
{b2b_signal_matches or "aucun"}

Risques B2C détectés :
{b2c_risk_matches or "aucun"}

Risques artisan-only détectés :
{artisan_risk_matches or "aucun"}

Signaux d'exclusion ou hors-scope détectés :
{exclusion_signal_matches or "aucun"}

Sources consultées pendant la recherche :
{research_sources or "aucune"}

Qualité de la recherche :
{research_quality or "inconnu"}

Signaux extraits de la recherche :
{research_signals or "aucun"}

Recherche web ciblee clients / use cases / industries / produit :
{research_targeted_summary or "aucune preuve ciblee"}

Texte public extrait du site et pages clés :
{research_text or "aucun"}
""".strip()


# ============================================================
# CLASSIFICATION
# ============================================================

def default_result(message: str = "Données insuffisantes.") -> dict:
    return {
        "algo2_decision": "A_VERIFIER",
        "algo2_confidence": 0,
        "b2b_strength": "inconnu",
        "b2b_reason": message,
        "confirmed_sector": "inconnu",
        "sector_alignment_strength": "inconnu",
        "sector_alignment_reason": message,
        "matched_issue_matrices": [],
        "main_issue_matrix": "inconnu",
        "issue_alignment_strength": "inconnu",
        "issue_alignment_reason": message,
        "edda_status": "À vérifier manuellement",
        "reason": message,
        "missing_information": "Informations à vérifier manuellement.",
    }


def normalize_strength(value: str) -> str:
    value = normalize_text(value)
    for allowed in STRENGTH_VALUES:
        if value == normalize_text(allowed):
            return allowed
    return "inconnu"


def normalize_sector(value: str) -> str:
    normalized = normalize_text(value)
    for allowed in SECTOR_VALUES:
        if normalized == normalize_text(allowed):
            return allowed
    return "inconnu"


def normalize_edda_status(value: str, decision: str) -> str:
    normalized = normalize_text(value)
    for allowed in EDDA_STATUSES:
        if normalized == normalize_text(allowed):
            return allowed

    if decision == "PREQUALIFIEE":
        return "Préqualifiée — à orienter"
    if decision == "NON_PREQUALIFIEE":
        return "Non préqualifiée"
    return "À vérifier manuellement"


def normalize_issue_matrices(value) -> list[str]:
    if value is None:
        return []

    if isinstance(value, str):
        raw_values = [part.strip() for part in re.split(r"[,;/|]+", value) if part.strip()]
    elif isinstance(value, list):
        raw_values = value
    else:
        raw_values = []

    normalized_values = []
    for raw_value in raw_values:
        normalized = normalize_text(raw_value)
        for allowed in ISSUE_MATRICES:
            if normalized == normalize_text(allowed) and allowed not in normalized_values:
                normalized_values.append(allowed)

    return normalized_values


def ensure_result_shape(result: dict) -> dict:
    base = default_result()
    if not isinstance(result, dict):
        return base

    base.update(result)

    decision = clean_text(base.get("algo2_decision", "A_VERIFIER")).upper()
    if decision not in ALGO2_DECISIONS:
        decision = "A_VERIFIER"

    try:
        confidence = float(base.get("algo2_confidence", 0) or 0)
    except (TypeError, ValueError):
        confidence = 0
    confidence = max(0, min(1, confidence))

    matrices = normalize_issue_matrices(base.get("matched_issue_matrices"))
    main_matrix = clean_text(base.get("main_issue_matrix", "inconnu"))
    if main_matrix not in ISSUE_MATRICES:
        main_matrix = matrices[0] if matrices else "inconnu"

    base["algo2_decision"] = decision
    base["algo2_confidence"] = confidence
    base["b2b_strength"] = normalize_strength(base.get("b2b_strength", "inconnu"))
    base["confirmed_sector"] = normalize_sector(base.get("confirmed_sector", "inconnu"))
    base["sector_alignment_strength"] = normalize_strength(base.get("sector_alignment_strength", "inconnu"))
    base["matched_issue_matrices"] = matrices
    base["main_issue_matrix"] = main_matrix
    base["issue_alignment_strength"] = normalize_strength(base.get("issue_alignment_strength", "inconnu"))
    base["edda_status"] = normalize_edda_status(base.get("edda_status", ""), decision)

    for key in [
        "b2b_reason",
        "sector_alignment_reason",
        "issue_alignment_reason",
        "reason",
        "missing_information",
    ]:
        base[key] = clean_text(base.get(key, ""))

    return base


def classify_startup(
    startup_name: str,
    sector: str,
    description: str,
    site: str,
    algo1_decision: str,
    algo1_reason: str,
    algo1_b2b: str,
    algo1_sector_relevance: str,
    sector_keyword_matches: dict[str, list[str]],
    issue_keyword_matches: dict[str, list[str]],
    research_sources: str = "",
    research_quality: str = "",
    research_signals: str = "",
    research_targeted_summary: str = "",
    research_text: str = "",
    deterministic_review: str = "",
    b2b_signal_matches: str = "",
    b2c_risk_matches: str = "",
    artisan_risk_matches: str = "",
    exclusion_signal_matches: str = "",
    vinci_fit_score: str = "",
    excel_context: str = "",
    excel_context_columns: str = "",
    excel_context_truncated: str = "non",
) -> dict:
    prompt = build_prompt(
        startup_name=startup_name,
        sector=sector,
        description=description,
        site=site,
        algo1_decision=algo1_decision,
        algo1_reason=algo1_reason,
        algo1_b2b=algo1_b2b,
        algo1_sector_relevance=algo1_sector_relevance,
        sector_keyword_matches=sector_keyword_matches,
        issue_keyword_matches=issue_keyword_matches,
        research_sources=research_sources,
        research_quality=research_quality,
        research_signals=research_signals,
        research_targeted_summary=research_targeted_summary,
        research_text=research_text,
        deterministic_review=deterministic_review,
        b2b_signal_matches=b2b_signal_matches,
        b2c_risk_matches=b2c_risk_matches,
        artisan_risk_matches=artisan_risk_matches,
        exclusion_signal_matches=exclusion_signal_matches,
        vinci_fit_score=vinci_fit_score,
        excel_context=excel_context,
        excel_context_columns=excel_context_columns if INCLUDE_EXCEL_COLUMN_LIST_IN_PROMPT else "",
        excel_context_truncated=excel_context_truncated,
    )
    schema = algo2_schema()

    for attempt in range(1, MAX_API_RETRIES + 1):
        try:
            result = call_model_json(
                client=get_thread_azure_client(),
                messages=build_messages(prompt),
                schema=schema,
                schema_name="algo2_strategic_filter",
            )
            return ensure_result_shape(result)

        except Exception as exc:
            error_details = format_api_error(exc)

            if is_non_retryable_api_error(exc):
                log(f"Erreur API non réessayée : {error_details}")
                return default_result(f"Erreur pendant la classification : {error_details}")

            if attempt == MAX_API_RETRIES:
                log(f"Erreur API finale : {error_details}")
                return default_result(f"Erreur pendant la classification : {error_details}")

            retry_after_seconds = get_retry_after_seconds(exc)
            wait_seconds = retry_after_seconds or min(30, (2 ** attempt) + random.uniform(0, 0.5))
            log(
                "Erreur API, nouvel essai dans "
                f"{wait_seconds:.1f}s ({attempt}/{MAX_API_RETRIES}) : {error_details}"
            )
            time.sleep(wait_seconds)

    return default_result("Erreur API inconnue.")


def edda_status_from_final_decision(result: dict) -> str:
    decision = result.get("algo2_decision_finale", result.get("algo2_decision", "A_VERIFIER"))

    if decision == "PREQUALIFIEE":
        status = result.get("edda_status", "")
        if status in [
            "Préqualifiée — à orienter",
            "Préqualifiée — en attente de classification",
        ]:
            return status
        return "Préqualifiée — à orienter"

    if decision == "NON_PREQUALIFIEE":
        return "Non préqualifiée"

    return "À vérifier manuellement"


def has_detected_signal(value: str) -> bool:
    normalized = normalize_text(value)
    return bool(normalized and not normalized.startswith("aucun signal") and not normalized.startswith("aucun mot cle"))


def apply_final_rules(result: dict) -> dict:
    result = ensure_result_shape(result)

    decision = result.get("algo2_decision", "A_VERIFIER")
    confidence = float(result.get("algo2_confidence", 0) or 0)
    b2b_strength = result.get("b2b_strength", "inconnu")
    confirmed_sector = result.get("confirmed_sector", "inconnu")
    sector_alignment_strength = result.get("sector_alignment_strength", "inconnu")
    matched_matrices = result.get("matched_issue_matrices", [])
    issue_alignment_strength = result.get("issue_alignment_strength", "inconnu")
    try:
        vinci_fit_score = float(result.get("algo2_vinci_fit_score", 0) or 0)
    except (TypeError, ValueError):
        vinci_fit_score = 0

    has_b2c_risk = has_detected_signal(result.get("algo2_b2c_risk_matches", ""))
    has_artisan_risk = has_detected_signal(result.get("algo2_artisan_risk_matches", ""))
    has_exclusion_risk = has_detected_signal(result.get("algo2_exclusion_signal_matches", ""))

    if decision == "PREQUALIFIEE" and confidence < 0.75:
        result["algo2_decision_finale"] = "A_VERIFIER"
        result["algo2_regle_finale"] = "Décision positive mais confiance insuffisante."
    elif decision == "PREQUALIFIEE" and b2b_strength != "fort":
        result["algo2_decision_finale"] = "A_VERIFIER"
        result["algo2_regle_finale"] = "B2B fort non confirmé."
    elif decision == "PREQUALIFIEE" and confirmed_sector in ["inconnu", "hors_scope"]:
        result["algo2_decision_finale"] = "A_VERIFIER"
        result["algo2_regle_finale"] = "Secteur cible non confirmé."
    elif decision == "PREQUALIFIEE" and sector_alignment_strength not in ["fort", "moyen"]:
        result["algo2_decision_finale"] = "A_VERIFIER"
        result["algo2_regle_finale"] = "Alignement sectoriel insuffisant."
    elif decision == "PREQUALIFIEE" and not matched_matrices:
        result["algo2_decision_finale"] = "A_VERIFIER"
        result["algo2_regle_finale"] = "Aucune matrice d'enjeux confirmée."
    elif decision == "PREQUALIFIEE" and issue_alignment_strength not in ["fort", "moyen"]:
        result["algo2_decision_finale"] = "A_VERIFIER"
        result["algo2_regle_finale"] = "Alignement enjeu insuffisant."
    elif decision == "PREQUALIFIEE" and vinci_fit_score and vinci_fit_score < 0.35:
        result["algo2_decision_finale"] = "A_VERIFIER"
        result["algo2_regle_finale"] = "Score déterministe VINCI trop faible pour préqualifier automatiquement."
    elif decision == "PREQUALIFIEE" and has_exclusion_risk and confidence < 0.9:
        result["algo2_decision_finale"] = "A_VERIFIER"
        result["algo2_regle_finale"] = "Signaux d'exclusion détectés : vérification manuelle requise."
    elif decision == "PREQUALIFIEE" and (has_b2c_risk or has_artisan_risk) and b2b_strength != "fort":
        result["algo2_decision_finale"] = "A_VERIFIER"
        result["algo2_regle_finale"] = "Risque B2C/artisan détecté et B2B fort insuffisamment prouvé."
    else:
        result["algo2_decision_finale"] = decision
        result["algo2_regle_finale"] = "Décision conservée après règles finales."

    result["edda_status_final"] = edda_status_from_final_decision(result)
    result["matched_issue_matrices"] = ", ".join(result.get("matched_issue_matrices", []))
    return result


def apply_profile_rules(result: dict, text: str) -> dict:
    result = apply_final_rules(result)
    result["dealflow_profile"] = dealflow_profiles.profile_label()
    exclusions = dealflow_profiles.detect_excluded_scope(text)
    result["profile_exclusion_matches"] = dealflow_profiles.format_matches(exclusions)

    if exclusions:
        result["algo2_decision"] = "NON_PREQUALIFIEE"
        result["algo2_decision_finale"] = "NON_PREQUALIFIEE"
        result["algo2_confidence"] = max(float(result.get("algo2_confidence", 0) or 0), 0.86)
        result["confirmed_sector"] = "hors_scope"
        result["sector_alignment_strength"] = "faible"
        result["sector_alignment_reason"] = "LATAM excluded scope: " + dealflow_profiles.format_matches(exclusions)
        result["issue_alignment_strength"] = "faible"
        result["edda_status"] = "Non préqualifiée"
        result["edda_status_final"] = "Non préqualifiée"
        result["reason"] = (
            "LATAM profile exclusion: company appears mainly related to an excluded scope "
            f"({dealflow_profiles.format_matches(exclusions)})."
        )
        result["missing_information"] = ""
        result["algo2_regle_finale"] = "LATAM profile exclusion applied."

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

    if columns["algo1_decision"] is None:
        raise ValueError("Impossible de trouver la colonne decision_finale de l'Algo 1.")

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

    log(f"Startups éligibles Algo 2 : {len(filtered_df)}/{len(df)}")
    log(f"Filtre appliqué : décision Algo 1 dans {sorted(ELIGIBLE_DECISIONS)}")

    return filtered_df


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


def process_startup_row(
    row: pd.Series,
    columns: dict[str, str | None],
    research_cache: dict[str, dict],
    research_cache_lock: threading.Lock,
) -> dict:
    try:
        startup_name = row_value(row, columns["startup"])
        sector = row_value(row, columns["sector"])
        description = compact_join([
            row_value(row, columns["description"]),
            row_value(row, columns.get("algo1_short_description")),
        ])
        site = row_value(row, columns["site"])
        algo1_decision = first_row_value(row, [
            columns["algo1_decision"],
            columns.get("algo1_raw_decision"),
        ])
        algo1_reason = first_row_value(row, [
            columns.get("algo1_reason"),
            columns.get("algo1_secondary_reason"),
        ])
        algo1_b2b = row_value(row, columns.get("algo1_b2b"))
        algo1_sector_relevance = row_value(row, columns.get("algo1_sector_relevance"))
        excel_context_info = build_excel_row_context(row)
        excel_context = excel_context_info["text"]
        excel_context_columns = " | ".join(excel_context_info["used_columns"])
        excel_context_truncated = "oui" if excel_context_info["truncated"] else "non"
        local_keyword_text = compact_join([
            startup_name,
            sector,
            description,
            site,
            algo1_decision,
            algo1_reason,
            algo1_b2b,
            algo1_sector_relevance,
            excel_context,
        ])

        if not startup_name and not description and not site and not excel_context:
            return apply_profile_rules(default_result("Ligne vide ou données insuffisantes."), local_keyword_text)

        shortcut_result = build_excel_shortcut_result_algo2(
            startup_name=startup_name,
            sector=sector,
            description=description,
            site=site,
            algo1_decision=algo1_decision,
            algo1_reason=algo1_reason,
            algo1_b2b=algo1_b2b,
            algo1_sector_relevance=algo1_sector_relevance,
            excel_context=excel_context,
            excel_context_info=excel_context_info,
        )
        if shortcut_result is not None:
            return shortcut_result

        if DISABLE_API:
            result = default_result(
                "API Algo 2 désactivée et Excel insuffisant pour une décision automatique fiable."
            )
            result["algo2_sector_keyword_matches"] = ""
            result["algo2_issue_keyword_matches"] = ""
            result["algo2_sector_keyword_categories"] = ""
            result["algo2_issue_keyword_categories"] = ""
            result["algo2_research_sources"] = ""
            result["algo2_research_quality"] = "non_requise_api_desactivee"
            result["algo2_research_signals"] = ""
            result["algo2_research_targeted_sources"] = ""
            result["algo2_research_targeted_categories"] = ""
            result["algo2_research_targeted_summary"] = ""
            result["algo2_research_evidence_clients"] = ""
            result["algo2_research_evidence_use_cases"] = ""
            result["algo2_research_evidence_industries"] = ""
            result["algo2_research_evidence_products"] = ""
            result["algo2_research_text_used"] = ""
            result["algo2_b2b_signal_matches"] = ""
            result["algo2_b2c_risk_matches"] = ""
            result["algo2_artisan_risk_matches"] = ""
            result["algo2_exclusion_signal_matches"] = ""
            result["algo2_vinci_fit_score"] = 0
            result["algo2_deterministic_review"] = "API désactivée et données Excel insuffisantes."
            result["algo2_excel_columns_used"] = excel_context_columns
            result["algo2_excel_columns_used_count"] = excel_context_info["used_count"]
            result["algo2_excel_columns_available_count"] = excel_context_info["available_count"]
            result["algo2_excel_context_truncated"] = excel_context_truncated
            result["algo2_processing_mode"] = "api_disabled_excel_insufficient"
            result["algo2_api_called"] = "non"
            result["algo2_research_called"] = "non"
            semantic_score = semantic_prefilter.score_algo2(local_keyword_text)
            result["algo2_semantic_backend"] = semantic_score.get("backend", "")
            result["algo2_semantic_summary"] = semantic_prefilter.summarize(semantic_score)
            result["algo2_semantic_sector"] = semantic_score.get("sector", "")
            result["algo2_semantic_sector_score"] = semantic_score.get("sector_score", 0)
            result["algo2_semantic_issue"] = semantic_score.get("issue", "")
            result["algo2_semantic_issue_score"] = semantic_score.get("issue_score", 0)
            result["algo2_semantic_b2b_score"] = semantic_score.get("b2b_score", 0)
            result["algo2_semantic_b2c_score"] = semantic_score.get("b2c_score", 0)
            learning_match = learning_memory.predict("algo2", local_keyword_text)
            result["algo2_learning_match_label"] = learning_match.get("label", "")
            result["algo2_learning_match_score"] = learning_match.get("score", 0)
            result["algo2_learning_match_source"] = learning_match.get("source", "")
            return apply_profile_rules(result, local_keyword_text)

        research = research_startup_website(site, research_cache, research_cache_lock)
        research_sources = research.get("sources", "")
        research_quality = research.get("quality", "inconnu")
        research_signals = research.get("signals", "")
        research_targeted_sources = research.get("targeted_sources", "")
        research_targeted_categories = research.get("targeted_categories", "")
        research_targeted_summary = research.get("targeted_summary", "")
        research_evidence_clients = research.get("evidence_clients", "")
        research_evidence_use_cases = research.get("evidence_use_cases", "")
        research_evidence_industries = research.get("evidence_industries", "")
        research_evidence_products = research.get("evidence_products", "")
        research_text = research.get("text", "")

        keyword_text = compact_join([
            startup_name,
            sector,
            description,
            site,
            algo1_decision,
            algo1_reason,
            algo1_b2b,
            algo1_sector_relevance,
            excel_context,
            research_signals,
            research_targeted_summary,
            research_evidence_clients,
            research_evidence_use_cases,
            research_evidence_industries,
            research_evidence_products,
            research_text,
        ])
        sector_matches = detect_keyword_matches(keyword_text, SECTOR_KEYWORDS)
        issue_matches = detect_keyword_matches(keyword_text, ISSUE_KEYWORDS)
        deterministic_audit = build_deterministic_audit(
            keyword_text=keyword_text,
            sector_matches=sector_matches,
            issue_matches=issue_matches,
            research_quality=research_quality,
        )

        result = classify_startup(
            startup_name=startup_name,
            sector=sector,
            description=description,
            site=site,
            algo1_decision=algo1_decision,
            algo1_reason=algo1_reason,
            algo1_b2b=algo1_b2b,
            algo1_sector_relevance=algo1_sector_relevance,
            sector_keyword_matches=sector_matches,
            issue_keyword_matches=issue_matches,
            research_sources=research_sources,
            research_quality=research_quality,
            research_signals=research_signals,
            research_targeted_summary=research_targeted_summary,
            research_text=research_text,
            deterministic_review=deterministic_audit["review"],
            b2b_signal_matches=format_flat_matches(deterministic_audit["b2b_matches"]),
            b2c_risk_matches=format_flat_matches(deterministic_audit["b2c_matches"]),
            artisan_risk_matches=format_flat_matches(deterministic_audit["artisan_matches"]),
            exclusion_signal_matches=format_flat_matches(deterministic_audit["exclusion_matches"]),
            vinci_fit_score=str(deterministic_audit["vinci_fit_score"]),
            excel_context=excel_context,
            excel_context_columns=excel_context_columns,
            excel_context_truncated=excel_context_truncated,
        )

        result["algo2_sector_keyword_matches"] = format_keyword_matches(sector_matches)
        result["algo2_issue_keyword_matches"] = format_keyword_matches(issue_matches)
        result["algo2_sector_keyword_categories"] = ", ".join(matched_categories(sector_matches))
        result["algo2_issue_keyword_categories"] = ", ".join(matched_categories(issue_matches))
        result["algo2_research_sources"] = research_sources
        result["algo2_research_quality"] = research_quality
        result["algo2_research_signals"] = research_signals
        result["algo2_research_targeted_sources"] = research_targeted_sources
        result["algo2_research_targeted_categories"] = research_targeted_categories
        result["algo2_research_targeted_summary"] = research_targeted_summary
        result["algo2_research_evidence_clients"] = research_evidence_clients
        result["algo2_research_evidence_use_cases"] = research_evidence_use_cases
        result["algo2_research_evidence_industries"] = research_evidence_industries
        result["algo2_research_evidence_products"] = research_evidence_products
        result["algo2_research_text_used"] = research_text
        result["algo2_b2b_signal_matches"] = format_flat_matches(deterministic_audit["b2b_matches"])
        result["algo2_b2c_risk_matches"] = format_flat_matches(deterministic_audit["b2c_matches"])
        result["algo2_artisan_risk_matches"] = format_flat_matches(deterministic_audit["artisan_matches"])
        result["algo2_exclusion_signal_matches"] = format_flat_matches(deterministic_audit["exclusion_matches"])
        result["algo2_vinci_fit_score"] = deterministic_audit["vinci_fit_score"]
        result["algo2_deterministic_review"] = deterministic_audit["review"]
        result["algo2_excel_columns_used"] = excel_context_columns
        result["algo2_excel_columns_used_count"] = excel_context_info["used_count"]
        result["algo2_excel_columns_available_count"] = excel_context_info["available_count"]
        result["algo2_excel_context_truncated"] = excel_context_truncated
        result["algo2_processing_mode"] = "api_with_research"
        result["algo2_api_called"] = "oui"
        result["algo2_research_called"] = "oui" if research_sources else "non"
        semantic_score = semantic_prefilter.score_algo2(keyword_text)
        result["algo2_semantic_backend"] = semantic_score.get("backend", "")
        result["algo2_semantic_summary"] = semantic_prefilter.summarize(semantic_score)
        result["algo2_semantic_sector"] = semantic_score.get("sector", "")
        result["algo2_semantic_sector_score"] = semantic_score.get("sector_score", 0)
        result["algo2_semantic_issue"] = semantic_score.get("issue", "")
        result["algo2_semantic_issue_score"] = semantic_score.get("issue_score", 0)
        result["algo2_semantic_b2b_score"] = semantic_score.get("b2b_score", 0)
        result["algo2_semantic_b2c_score"] = semantic_score.get("b2c_score", 0)
        learning_match = learning_memory.predict("algo2", keyword_text)
        result["algo2_learning_match_label"] = learning_match.get("label", "")
        result["algo2_learning_match_score"] = learning_match.get("score", 0)
        result["algo2_learning_match_source"] = learning_match.get("source", "")

        if REQUEST_DELAY_SECONDS > 0:
            time.sleep(REQUEST_DELAY_SECONDS)

        return apply_profile_rules(result, keyword_text)

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

        for decision, sheet_name in [
            ("PREQUALIFIEE", "Prequalifiees"),
            ("A_VERIFIER", "A verifier"),
            ("NON_PREQUALIFIEE", "Non prequalifiees"),
        ]:
            if "algo2_decision_finale" in final_df.columns:
                sheet_df = final_df[final_df["algo2_decision_finale"] == decision]
            else:
                sheet_df = final_df.iloc[0:0]

            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)


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
    log(f"- Décision Algo 1 : {columns['algo1_decision']}")
    log(f"- Décision Algo 1 fallback : {columns.get('algo1_raw_decision') or 'non trouvée'}")
    log(f"- Colonnes totales du fichier : {len(df.columns)}")
    log(f"- Raison Algo 1 : {columns.get('algo1_reason') or 'non trouvée'}")
    log(f"- Raison Algo 1 fallback : {columns.get('algo1_secondary_reason') or 'non trouvée'}")

    df = filter_eligible_rows(df, columns["algo1_decision"], columns.get("algo1_raw_decision"))

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
    log(f"Recherche approfondie : {'activée' if ENABLE_RESEARCH else 'désactivée'}")
    log(f"Recherche ciblée Algo2 : {'activée' if TARGETED_RESEARCH else 'désactivée'}")
    log(f"Pages recherche max : {RESEARCH_MAX_PAGES}")
    log(f"Liens ciblés essayés max : {RESEARCH_MAX_LINKS_TO_TRY}")
    if start_index:
        log(f"Reprise à la ligne {start_index + 1}/{len(df)}.")

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        mapped_results = executor.map(
            lambda row: process_startup_row(row, columns, research_cache, research_cache_lock),
            rows,
        )

        for result in tqdm(mapped_results, total=len(rows), desc=f"Algo 2 {input_file.name}"):
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
    input_files = discover_input_files(base_dir, sys.argv[1:])

    if not input_files:
        log("Aucun fichier Excel issu de l'Algo 1 trouvé dans le dossier.")
        log("Tu peux aussi lancer : python Algo2.py chemin\\vers\\fichier_algo_1_prequalification.xlsx")
        return

    log(f"Déploiement Azure OpenAI utilisé : {MODEL_DEPLOYMENT}")
    log(f"Version API Azure OpenAI : {AZURE_OPENAI_API_VERSION}")
    log(f"Mode API : {ALGO2_API_MODE}")
    log(f"Workers parallèles : {MAX_WORKERS}")
    log(f"Décisions Algo 1 traitées : {sorted(ELIGIBLE_DECISIONS)}")
    log(f"{len(input_files)} fichier(s) Algo 1 à traiter.")

    validate_azure_client()

    for input_file in input_files:
        try:
            process_excel_file(input_file)
        except Exception as exc:
            log(f"Erreur sur {input_file.name} : {exc}")

    log("Traitement terminé.")


if __name__ == "__main__":
    main()
