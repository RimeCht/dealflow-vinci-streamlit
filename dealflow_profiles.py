import os
import re
import unicodedata


EUROPE = "EUROPE"
LATAM = "LATAM"

AVAILABLE_PROFILES = [EUROPE, LATAM]

LATAM_SEED_ARR_THRESHOLD_EUR = 500_000
LATAM_CATALYST_ARR_THRESHOLD_EUR = 500_000

LATAM_EXCLUDED_SCOPE_TERMS = [
    "building construction",
    "construction of buildings",
    "construction de batiments",
    "construction de bâtiments",
    "residential construction",
    "commercial building",
    "residential building",
    "building site",
    "real estate",
    "immobilier",
    "proptech",
    "property management",
    "asset management immobilier",
    "real estate asset management",
    "building facility management",
    "facility management buildings",
    "housing",
    "home renovation",
    "building renovation",
]


def strip_accents(text: str) -> str:
    return "".join(
        char
        for char in unicodedata.normalize("NFKD", str(text))
        if not unicodedata.combining(char)
    )


def normalize_text(text: str) -> str:
    text = strip_accents(str(text)).lower()
    text = re.sub(r"[_\-/]+", " ", text)
    text = re.sub(r"[^a-z0-9 +#.&']+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def current_profile() -> str:
    value = os.getenv("DEALFLOW_PROFILE", EUROPE).strip().upper()
    return LATAM if value == LATAM else EUROPE


def is_latam(profile: str | None = None) -> bool:
    return (profile or current_profile()).strip().upper() == LATAM


def output_language(profile: str | None = None) -> str:
    return "English" if is_latam(profile) else "French"


def _term_in_text(normalized_text: str, term: str) -> bool:
    normalized_term = normalize_text(term)
    if not normalized_term:
        return False
    if " " in normalized_term:
        return f" {normalized_term} " in f" {normalized_text} "
    return re.search(rf"\b{re.escape(normalized_term)}\b", normalized_text) is not None


def detect_excluded_scope(text: str, profile: str | None = None) -> list[str]:
    if not is_latam(profile):
        return []

    normalized = normalize_text(text)
    matches = []
    for term in LATAM_EXCLUDED_SCOPE_TERMS:
        if _term_in_text(normalized, term):
            clean_term = term.strip()
            if clean_term not in matches:
                matches.append(clean_term)
    return matches


def format_matches(matches: list[str]) -> str:
    return ", ".join(matches)


def profile_prompt_context() -> str:
    if not is_latam():
        return (
            "Profile: EUROPE.\n"
            "Use the existing Leonard / VINCI Europe criteria and answer in French."
        )

    return f"""
Profile: LATAM.
Use English for all model explanations, reasons, missing information and short descriptions.
Keep the same JSON field names and allowed enum values.

LATAM scope is narrower than Europe, but it still includes construction-related topics.
Exclude only companies mainly focused on:
- construction of buildings;
- residential/commercial building development;
- real estate, proptech, property management, asset management real estate, housing;
- building-only renovation or building facility management.

Do include other construction topics when relevant to VINCI, especially:
- infrastructure construction;
- civil engineering and public works;
- roads, bridges, tunnels, highways, rail and railway infrastructure;
- construction materials, building materials, concrete, cement, steel, asphalt, pavement and structural materials.

For LATAM, prioritize B2B opportunities linked to construction/infrastructure/materials, energy,
industrial decarbonization, climate/environment, operational productivity, safety/security and mobility.

ARR thresholds for LATAM orientation:
- Seed: ARR < {LATAM_SEED_ARR_THRESHOLD_EUR // 1000}k EUR.
- Catalyst: ARR > {LATAM_CATALYST_ARR_THRESHOLD_EUR // 1000}k EUR.
If ARR is unknown, write "inconnu" and do not invent it.
""".strip()


def profile_label() -> str:
    return current_profile()
