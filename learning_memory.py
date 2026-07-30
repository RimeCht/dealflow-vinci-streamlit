import json
import os
import time
from pathlib import Path

import pandas as pd

import semantic_prefilter


MEMORY_PATH = Path(os.getenv("LEARNING_MEMORY_PATH", "learning_memory.json"))
ENABLED = os.getenv("LEARNING_ENABLED", "1").strip().lower() not in {
    "0",
    "false",
    "non",
    "no",
    "off",
}
MATCH_THRESHOLD = float(os.getenv("LEARNING_MATCH_THRESHOLD", "0.28"))
MAX_EXAMPLES = int(os.getenv("LEARNING_MAX_EXAMPLES", "5000"))
MAX_TEXT_CHARS = int(os.getenv("LEARNING_MAX_TEXT_CHARS", "8000"))
_MEMORY_CACHE = None
_MEMORY_CACHE_MTIME = None

MANUAL_ALGO1_COLUMNS = [
    "manual_algo1_decision",
    "correction_algo1_decision",
    "algo1_manual_decision",
    "decision_algo1_corrigee",
]
MANUAL_ALGO2_COLUMNS = [
    "manual_algo2_decision",
    "correction_algo2_decision",
    "algo2_manual_decision",
    "decision_algo2_corrigee",
]
MANUAL_ALGO3_COLUMNS = [
    "manual_algo3_orientation",
    "correction_algo3_orientation",
    "algo3_manual_orientation",
    "orientation_algo3_corrigee",
]

ALGO1_LABELS = {"A_GARDER", "A_VERIFIER", "A_ECARTER"}
ALGO2_LABELS = {"PREQUALIFIEE", "A_VERIFIER", "NON_PREQUALIFIEE"}
ALGO3_LABELS = {"Seed", "Catalyst", "Matériaux", "Materiaux", "Autre", "A_VERIFIER"}

TEXT_PRIORITY_TERMS = [
    "name",
    "startup",
    "company",
    "description",
    "oneliner",
    "vertical",
    "sector",
    "subtopic",
    "target customer",
    "business model",
    "technology",
    "specific use case",
    "fit in construction value chain",
    "pain points",
    "success stories",
    "trl",
    "mrl",
    "employee",
    "funding",
    "revenue",
    "algo1_reason",
    "algo2_reason",
    "reason",
]

EXCLUDED_TERMS = [
    "email",
    "phone",
    "birth",
    "logo",
    "privacy",
    "owner",
    "director",
    "ubo",
    "affinity id",
    "entity id",
]


def clean_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def normalize_label(value: str, allowed: set[str]) -> str:
    raw = clean_text(value)
    if not raw:
        return ""

    normalized = raw.strip().replace(" ", "_").upper()
    for label in allowed:
        if normalized == label.replace(" ", "_").upper():
            return label
    if normalized == "MATERIAUX":
        return "Matériaux"
    return ""


def normalize_column(column: str) -> str:
    return str(column).strip().lower().replace("_", " ")


def find_first_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    normalized = {normalize_column(column): column for column in df.columns}
    for candidate in candidates:
        column = normalized.get(normalize_column(candidate))
        if column is not None:
            return column
    return None


def useful_row_text(row: pd.Series) -> str:
    entries = []
    for column, value in row.items():
        normalized = normalize_column(column)
        if any(term in normalized for term in EXCLUDED_TERMS):
            continue
        text = clean_text(value)
        if not text:
            continue
        priority = 0 if any(term in normalized for term in TEXT_PRIORITY_TERMS) else 1
        entries.append((priority, normalized, f"{column}: {text[:1200]}"))

    entries.sort(key=lambda item: (item[0], item[1]))
    parts = []
    total = 0
    for _, _, line in entries:
        if total + len(line) > MAX_TEXT_CHARS:
            break
        parts.append(line)
        total += len(line) + 1
    return "\n".join(parts)


def empty_memory() -> dict:
    return {
        "version": 1,
        "updated_at": "",
        "examples": [],
    }


def load_memory(path: Path = MEMORY_PATH) -> dict:
    global _MEMORY_CACHE, _MEMORY_CACHE_MTIME
    if not ENABLED:
        return empty_memory()
    if not path.exists():
        return empty_memory()
    try:
        mtime = path.stat().st_mtime
        if _MEMORY_CACHE is not None and _MEMORY_CACHE_MTIME == mtime:
            return _MEMORY_CACHE
    except Exception:
        mtime = None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return empty_memory()
    if not isinstance(data, dict):
        return empty_memory()
    if "examples" not in data or not isinstance(data["examples"], list):
        data["examples"] = []
    _MEMORY_CACHE = data
    _MEMORY_CACHE_MTIME = mtime
    return data


def save_memory(memory: dict, path: Path = MEMORY_PATH) -> None:
    global _MEMORY_CACHE, _MEMORY_CACHE_MTIME
    memory["version"] = 1
    memory["updated_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
    examples = memory.get("examples", [])
    memory["examples"] = examples[-MAX_EXAMPLES:]
    path.write_text(json.dumps(memory, ensure_ascii=False, indent=2), encoding="utf-8")
    _MEMORY_CACHE = memory
    try:
        _MEMORY_CACHE_MTIME = path.stat().st_mtime
    except Exception:
        _MEMORY_CACHE_MTIME = None


def example_key(stage: str, label: str, text: str) -> str:
    normalized_text = semantic_prefilter.normalize_text(text)
    return f"{stage}|{label}|{normalized_text[:500]}"


def add_example(memory: dict, stage: str, label: str, text: str, source: str, name: str = "", comment: str = "") -> bool:
    if not text or not label:
        return False
    examples = memory.setdefault("examples", [])
    key = example_key(stage, label, text)
    for example in examples:
        if example.get("key") == key:
            return False
    examples.append({
        "key": key,
        "stage": stage,
        "label": label,
        "text": text[:MAX_TEXT_CHARS],
        "name": name,
        "source": source,
        "comment": comment,
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    })
    return True


def learn_from_dataframe(df: pd.DataFrame, source: str, memory: dict | None = None) -> tuple[dict, dict]:
    memory = memory or load_memory()
    stats = {"algo1": 0, "algo2": 0, "algo3": 0}

    algo1_column = find_first_column(df, MANUAL_ALGO1_COLUMNS)
    algo2_column = find_first_column(df, MANUAL_ALGO2_COLUMNS)
    algo3_column = find_first_column(df, MANUAL_ALGO3_COLUMNS)
    comment_column = find_first_column(df, ["manual_comment", "commentaire_correction", "commentaire"])
    name_column = find_first_column(df, ["Name", "startup", "company", "nom", "entreprise"])

    for _, row in df.iterrows():
        text = useful_row_text(row)
        if not text:
            continue
        comment = clean_text(row.get(comment_column, "")) if comment_column else ""
        name = clean_text(row.get(name_column, "")) if name_column else ""

        if algo1_column:
            label = normalize_label(row.get(algo1_column, ""), ALGO1_LABELS)
            if label and add_example(memory, "algo1", label, text, source, name, comment):
                stats["algo1"] += 1
        if algo2_column:
            label = normalize_label(row.get(algo2_column, ""), ALGO2_LABELS)
            if label and add_example(memory, "algo2", label, text, source, name, comment):
                stats["algo2"] += 1
        if algo3_column:
            label = normalize_label(row.get(algo3_column, ""), ALGO3_LABELS)
            if label and add_example(memory, "algo3", label, text, source, name, comment):
                stats["algo3"] += 1

    return memory, stats


def learn_from_excel(path: Path, memory: dict | None = None) -> tuple[dict, dict]:
    memory = memory or load_memory()
    stats = {"algo1": 0, "algo2": 0, "algo3": 0}
    workbook = pd.ExcelFile(path)
    for sheet_name in workbook.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet_name)
        memory, sheet_stats = learn_from_dataframe(df, f"{path.name}::{sheet_name}", memory)
        for key, value in sheet_stats.items():
            stats[key] += value
    return memory, stats


def predict(stage: str, text: str, threshold: float = MATCH_THRESHOLD) -> dict:
    if not ENABLED or not text:
        return {"matched": False, "label": "", "score": 0.0, "source": "", "name": ""}

    examples = [
        example for example in load_memory().get("examples", [])
        if example.get("stage") == stage and example.get("text")
    ]
    if not examples:
        return {"matched": False, "label": "", "score": 0.0, "source": "", "name": ""}

    text_embedding = semantic_prefilter.hashed_embedding(text[:MAX_TEXT_CHARS])
    best = None
    for example in examples:
        score = semantic_prefilter.dot_sparse(
            text_embedding,
            semantic_prefilter.hashed_embedding(example["text"][:MAX_TEXT_CHARS]),
        )
        if best is None or score > best["score"]:
            best = {
                "matched": score >= threshold,
                "label": example.get("label", ""),
                "score": round(max(0.0, score), 4),
                "source": example.get("source", ""),
                "name": example.get("name", ""),
                "comment": example.get("comment", ""),
            }

    return best or {"matched": False, "label": "", "score": 0.0, "source": "", "name": ""}
