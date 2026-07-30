import hashlib
import math
import os
import re
import unicodedata
from functools import lru_cache


def env_bool(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() not in {"0", "false", "non", "no", "off"}


ENABLED = env_bool("LOCAL_SEMANTIC_PREFILTER", True)
USE_TRANSFORMER = env_bool("LOCAL_SEMANTIC_USE_TRANSFORMER", False)
MODEL_NAME = os.getenv(
    "LOCAL_SEMANTIC_MODEL",
    "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2",
)
MODEL_LOCAL_ONLY = env_bool("LOCAL_SEMANTIC_MODEL_LOCAL_ONLY", True)
HASH_DIMS = int(os.getenv("LOCAL_SEMANTIC_HASH_DIMS", "768"))

HASH_STRONG_THRESHOLD = float(os.getenv("LOCAL_SEMANTIC_HASH_STRONG_THRESHOLD", "0.18"))
HASH_MEDIUM_THRESHOLD = float(os.getenv("LOCAL_SEMANTIC_HASH_MEDIUM_THRESHOLD", "0.13"))
TRANSFORMER_STRONG_THRESHOLD = float(os.getenv("LOCAL_SEMANTIC_TRANSFORMER_STRONG_THRESHOLD", "0.42"))
TRANSFORMER_MEDIUM_THRESHOLD = float(os.getenv("LOCAL_SEMANTIC_TRANSFORMER_MEDIUM_THRESHOLD", "0.34"))

_MODEL = None
_MODEL_FAILED = False


PHRASES = {
    "b2b": {
        "B2B": [
            "solution for enterprises",
            "platform for construction companies",
            "software for industrial operators",
            "tool for infrastructure operators",
            "SaaS for large companies",
            "enterprise asset management",
            "facility management for buildings",
            "project management for contractors",
            "fleet management for professional operators",
            "energy management for businesses",
            "public sector and municipalities",
            "professional users and corporate clients",
        ],
        "B2C": [
            "consumer mobile app",
            "marketplace for individuals",
            "service for homeowners",
            "wellness app for consumers",
            "tourism platform for travelers",
            "retail product for consumers",
            "social network for users",
            "personal finance application",
            "fashion or food consumer brand",
        ],
        "Artisans": [
            "tool for artisans",
            "solution for freelancers",
            "application for independent workers",
            "service for small tradespeople only",
        ],
    },
    "sectors": {
        "Construction": [
            "construction site management",
            "building construction workflow",
            "BIM and construction planning",
            "construction safety and quality",
            "materials and renovation for buildings",
            "prefabrication for construction",
        ],
        "Built World": [
            "smart building operations",
            "facility management and maintenance",
            "digital twin for buildings",
            "asset management for built environment",
            "building performance monitoring",
        ],
        "Infrastructure": [
            "roads bridges tunnels rail infrastructure",
            "civil engineering asset inspection",
            "infrastructure maintenance",
            "water networks and utility infrastructure",
            "airport port highway operations",
        ],
        "Real Estate": [
            "property management",
            "real estate asset management",
            "proptech for building owners",
            "building renovation and ESG real estate",
            "facility services for real estate portfolios",
        ],
        "Energy": [
            "energy efficiency management",
            "renewable energy projects",
            "smart grid and flexibility",
            "battery storage and solar energy",
            "decarbonization of buildings and industry",
        ],
        "Mobility": [
            "fleet management and mobility",
            "electric vehicle charging",
            "traffic and transport optimization",
            "public transport and logistics",
            "parking tolling and smart mobility",
        ],
    },
    "issues": {
        "Productivity": [
            "automation and time savings",
            "process optimization for operations",
            "manual task reduction",
            "workflow simplification",
            "faster than human work",
        ],
        "Competitiveness": [
            "cost savings and profitability",
            "quality of service improvement",
            "competitive advantage",
            "risk reduction for business",
            "better tender response capability",
        ],
        "Environment": [
            "reduced environmental impact",
            "waste reduction and circular economy",
            "lower energy consumption",
            "pollution and CO2 reduction",
            "resource efficiency",
        ],
        "Climate": [
            "climate resilience",
            "climate adaptation",
            "decarbonization pathway",
            "carbon reduction",
            "climate risk response",
        ],
        "Safety/Security": [
            "worker safety",
            "site security",
            "infrastructure safety monitoring",
            "cybersecurity and data security",
            "operational risk prevention",
        ],
    },
}


def normalize_text(text) -> str:
    if text is None:
        return ""
    text = str(text)
    text = "".join(
        char for char in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(char)
    )
    text = text.lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _terms(text: str) -> list[str]:
    normalized = normalize_text(text)
    if not normalized:
        return []

    words = normalized.split()
    terms = []
    terms.extend(words)
    terms.extend(" ".join(words[index:index + 2]) for index in range(max(0, len(words) - 1)))
    terms.extend(" ".join(words[index:index + 3]) for index in range(max(0, len(words) - 2)))

    compact = normalized.replace(" ", "_")
    if len(compact) >= 4:
        terms.extend(compact[index:index + 4] for index in range(len(compact) - 3))

    return terms


@lru_cache(maxsize=8192)
def hashed_embedding(text: str) -> tuple[tuple[int, float], ...]:
    weights = {}
    for term in _terms(text):
        digest = hashlib.blake2b(term.encode("utf-8"), digest_size=8).digest()
        integer = int.from_bytes(digest, "big", signed=False)
        index = integer % HASH_DIMS
        sign = 1.0 if (integer >> 63) == 0 else -1.0
        weight = 1.0 + min(2.0, len(term) / 18.0)
        weights[index] = weights.get(index, 0.0) + sign * weight

    norm = math.sqrt(sum(value * value for value in weights.values()))
    if norm == 0:
        return tuple()
    return tuple(sorted((index, value / norm) for index, value in weights.items()))


def dot_sparse(left: tuple[tuple[int, float], ...], right: tuple[tuple[int, float], ...]) -> float:
    left_dict = dict(left)
    return sum(left_dict.get(index, 0.0) * value for index, value in right)


def _get_transformer_model():
    global _MODEL, _MODEL_FAILED
    if not USE_TRANSFORMER or _MODEL_FAILED:
        return None
    if _MODEL is not None:
        return _MODEL
    try:
        from sentence_transformers import SentenceTransformer

        try:
            _MODEL = SentenceTransformer(MODEL_NAME, local_files_only=MODEL_LOCAL_ONLY)
        except TypeError:
            _MODEL = SentenceTransformer(MODEL_NAME)
        return _MODEL
    except Exception:
        _MODEL_FAILED = True
        return None


def backend_name() -> str:
    if _get_transformer_model() is not None:
        return f"sentence_transformers:{MODEL_NAME}"
    if USE_TRANSFORMER:
        return "hashing_fallback_transformer_unavailable"
    return "hashing"


def _hash_scores(text: str, phrases_by_label: dict[str, list[str]]) -> dict[str, dict]:
    text_embedding = hashed_embedding(text[:12000])
    output = {}
    for label, phrases in phrases_by_label.items():
        scored = [
            (dot_sparse(text_embedding, hashed_embedding(phrase)), phrase)
            for phrase in phrases
        ]
        best_score, best_phrase = max(scored, default=(0.0, ""))
        output[label] = {
            "score": round(max(0.0, best_score), 4),
            "phrase": best_phrase,
        }
    return output


def _transformer_scores(text: str, phrases_by_label: dict[str, list[str]]) -> dict[str, dict] | None:
    model = _get_transformer_model()
    if model is None:
        return None

    try:
        labels = list(phrases_by_label.keys())
        phrases = [phrase for label in labels for phrase in phrases_by_label[label]]
        phrase_to_label = [label for label in labels for _ in phrases_by_label[label]]
        vectors = model.encode([text[:12000], *phrases], normalize_embeddings=True)
        text_vector = vectors[0]
        phrase_vectors = vectors[1:]

        output = {
            label: {"score": 0.0, "phrase": ""}
            for label in labels
        }
        for label, phrase, vector in zip(phrase_to_label, phrases, phrase_vectors):
            score = float(text_vector @ vector)
            if score > output[label]["score"]:
                output[label] = {"score": round(score, 4), "phrase": phrase}
        return output
    except Exception:
        return None


def score_labels(text: str, phrases_by_label: dict[str, list[str]]) -> dict[str, dict]:
    if not ENABLED:
        return {
            label: {"score": 0.0, "phrase": ""}
            for label in phrases_by_label
        }

    transformer_scores = _transformer_scores(text, phrases_by_label)
    if transformer_scores is not None:
        return transformer_scores
    return _hash_scores(text, phrases_by_label)


def best_label(scores: dict[str, dict]) -> tuple[str, float, str]:
    if not scores:
        return "inconnu", 0.0, ""
    label, data = max(scores.items(), key=lambda item: item[1]["score"])
    return label, float(data["score"]), data.get("phrase", "")


def threshold(kind: str = "strong") -> float:
    backend = backend_name()
    if backend.startswith("sentence_transformers"):
        return TRANSFORMER_STRONG_THRESHOLD if kind == "strong" else TRANSFORMER_MEDIUM_THRESHOLD
    return HASH_STRONG_THRESHOLD if kind == "strong" else HASH_MEDIUM_THRESHOLD


def is_strong(score: float) -> bool:
    return score >= threshold("strong")


def is_medium(score: float) -> bool:
    return score >= threshold("medium")


def score_algo1(text: str) -> dict:
    b2b_scores = score_labels(text, PHRASES["b2b"])
    sector_scores = score_labels(text, PHRASES["sectors"])
    sector, sector_score, sector_phrase = best_label(sector_scores)

    return {
        "enabled": ENABLED,
        "backend": backend_name(),
        "b2b_score": b2b_scores.get("B2B", {}).get("score", 0.0),
        "b2b_phrase": b2b_scores.get("B2B", {}).get("phrase", ""),
        "b2c_score": b2b_scores.get("B2C", {}).get("score", 0.0),
        "b2c_phrase": b2b_scores.get("B2C", {}).get("phrase", ""),
        "artisan_score": b2b_scores.get("Artisans", {}).get("score", 0.0),
        "artisan_phrase": b2b_scores.get("Artisans", {}).get("phrase", ""),
        "sector": sector,
        "sector_score": sector_score,
        "sector_phrase": sector_phrase,
    }


def score_algo2(text: str) -> dict:
    base = score_algo1(text)
    issue_scores = score_labels(text, PHRASES["issues"])
    issue, issue_score, issue_phrase = best_label(issue_scores)
    base.update({
        "issue": issue,
        "issue_score": issue_score,
        "issue_phrase": issue_phrase,
    })
    return base


def summarize(score: dict) -> str:
    if not score:
        return "semantic_prefilter=disabled"
    parts = [
        f"backend={score.get('backend', 'unknown')}",
        f"b2b={score.get('b2b_score', 0)}",
        f"b2c={score.get('b2c_score', 0)}",
        f"artisan={score.get('artisan_score', 0)}",
        f"sector={score.get('sector', 'inconnu')}:{score.get('sector_score', 0)}",
    ]
    if "issue" in score:
        parts.append(f"issue={score.get('issue', 'inconnu')}:{score.get('issue_score', 0)}")
    return "; ".join(parts)
