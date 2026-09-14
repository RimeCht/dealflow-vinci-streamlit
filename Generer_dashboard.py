import base64
import json
import math
import re
import sys
import unicodedata
from datetime import datetime
from html import escape
from pathlib import Path

import pandas as pd


OUTPUT_SUFFIX = "_dashboard.html"
PIPELINE_SUFFIX = "_pipeline_algo_1_2_3"
TEMP_MARKER = "_sauvegarde_temp"
ASSETS_DIR = Path(__file__).resolve().parent / "assets"


def asset_data_uri(
    filename: str,
    mime_type: str,
    *,
    fix_svg: bool = False,
    replacements: dict[str, str] | None = None,
) -> str:
    path = ASSETS_DIR / filename
    try:
        data = path.read_bytes()
    except OSError:
        return ""
    if fix_svg or replacements:
        text = data.decode("utf-8")
        if fix_svg:
            text = text.replace("viewbox=", "viewBox=")
        for source, target in (replacements or {}).items():
            text = text.replace(source, target)
        data = text.encode("utf-8")
    return f"data:{mime_type};base64,{base64.b64encode(data).decode('ascii')}"


def dashboard_font_css() -> str:
    fonts = [
        ("Vinci Sans", "Vinci-Sans-Regular.woff2", 400),
        ("Vinci Sans", "Vinci-Sans-Medium.woff2", 500),
        ("Vinci Sans", "Vinci-Sans-Bold.woff2", 700),
        ("Vinci Serif", "Vinci-Serif-Regular.woff2", 400),
    ]
    rules = []
    for family, filename, weight in fonts:
        uri = asset_data_uri(filename, "font/woff2")
        if uri:
            rules.append(
                "@font-face {"
                f"font-family:'{family}';"
                f"src:url('{uri}') format('woff2');"
                f"font-weight:{weight};font-style:normal;font-display:swap;"
                "}"
            )
    return "".join(rules)


def log(message: str) -> None:
    print(f"[Dashboard] {message}")


def clean_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def strip_accents(value: str) -> str:
    return "".join(
        char
        for char in unicodedata.normalize("NFKD", clean_text(value))
        if not unicodedata.combining(char)
    )


def normalize_key(value: str) -> str:
    value = strip_accents(value).lower()
    value = re.sub(r"[_\-/]+", " ", value)
    value = re.sub(r"[^a-z0-9 +#.&']+", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def clean_decision(value: str) -> str:
    return normalize_key(value).replace(" ", "_").upper()


def discover_input_files(base_dir: Path, cli_args: list[str]) -> list[Path]:
    if cli_args:
        return [Path(arg).expanduser().resolve() for arg in cli_args]

    return [
        path
        for path in sorted(base_dir.glob(f"*{PIPELINE_SUFFIX}.xlsx"))
        if path.is_file()
        and not path.name.startswith("~$")
        and TEMP_MARKER not in path.stem
    ]


def make_output_path(input_file: Path) -> Path:
    return input_file.with_name(f"{input_file.stem}{OUTPUT_SUFFIX}")


def normalize_header(value: str) -> str:
    return normalize_key(value).replace("_", " ")


def first_matching_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    if df.empty:
        return None

    columns = [str(column) for column in df.columns]
    exact = {column: column for column in columns}
    for candidate in candidates:
        if candidate in exact:
            return exact[candidate]

    normalized_columns = {
        normalize_header(column): column
        for column in columns
    }
    for candidate in candidates:
        normalized_candidate = normalize_header(candidate)
        if normalized_candidate in normalized_columns:
            return normalized_columns[normalized_candidate]

    for candidate in candidates:
        normalized_candidate = normalize_header(candidate)
        for normalized_column, original_column in normalized_columns.items():
            if normalized_candidate and normalized_candidate in normalized_column:
                return original_column

    return None


def value_from_row(row: pd.Series, column_map: dict[str, str | None], key: str) -> str:
    column = column_map.get(key)
    if not column:
        return ""
    return clean_text(row.get(column, ""))


FIELD_ALIASES = {
    "pipeline_index": ["pipeline_index", "index", "Index"],
    "startup": [
        "startup",
        "Startup",
        "Company Name",
        "Nom startup",
        "Nom de la startup",
        "Company",
        "Company name",
        "Startup name",
        "name",
        "Name",
    ],
    "site_web": [
        "site_web",
        "site web",
        "Official Website",
        "website",
        "Website",
        "url",
        "URL",
        "company website",
    ],
    "secteur_source": [
        "secteur_source",
        "sector",
        "secteur",
        "Industry / Market",
        "industry",
        "vertical",
        "category",
        "categorie",
    ],
    "dealflow_profile": [
        "dealflow_profile",
        "algo1_dealflow_profile",
        "algo2_dealflow_profile",
        "algo3_dealflow_profile",
    ],
    "profile_exclusion_matches": [
        "profile_exclusion_matches",
        "algo1_profile_exclusion_matches",
        "algo2_profile_exclusion_matches",
        "algo3_profile_exclusion_matches",
    ],
    "resultat_global": ["resultat_global"],
    "etape_finale": ["etape_finale", "pipeline_stage"],
    "statut_pipeline": ["statut_pipeline", "pipeline_status"],
    "raison_arret_ou_passage": ["raison_arret_ou_passage", "pipeline_stop_reason"],
    "prochaine_action": ["prochaine_action", "pipeline_next_action"],
    "algo1_execute": ["algo1_execute", "pipeline_algo1_executed"],
    "algo1_decision_finale": ["algo1_decision_finale", "decision_finale"],
    "algo1_passe_vers_algo2": ["algo1_passe_vers_algo2"],
    "algo1_raison": ["algo1_raison", "algo1_reason", "reason"],
    "algo1_regle_finale": ["algo1_regle_finale", "regle_finale"],
    "algo1_infos_manquantes": ["algo1_infos_manquantes", "algo1_missing_information", "missing_information"],
    "algo2_execute": ["algo2_execute", "pipeline_algo2_executed"],
    "algo2_decision_finale": ["algo2_decision_finale"],
    "algo2_passe_vers_algo3": ["algo2_passe_vers_algo3"],
    "algo2_raison": ["algo2_raison", "algo2_reason"],
    "algo2_regle_finale": ["algo2_regle_finale"],
    "algo2_b2b": ["algo2_b2b", "b2b_strength", "algo2_b2b_strength"],
    "algo2_secteur_confirme": ["algo2_secteur_confirme", "confirmed_sector", "algo2_confirmed_sector"],
    "algo2_matrice_enjeu": ["algo2_matrice_enjeu", "main_issue_matrix", "algo2_main_issue_matrix"],
    "algo2_infos_manquantes": ["algo2_infos_manquantes", "algo2_missing_information"],
    "edda_status_final": ["edda_status_final", "algo2_edda_status_final"],
    "algo3_execute": ["algo3_execute", "pipeline_algo3_executed"],
    "algo3_orientation_finale": ["algo3_orientation_finale"],
    "algo3_raison": ["algo3_raison", "algo3_orientation_reason", "orientation_reason"],
    "algo3_regle_finale": ["algo3_regle_finale"],
    "algo3_infos_manquantes": ["algo3_infos_manquantes", "algo3_missing_information"],
    "algo3_solution_type": ["algo3_solution_type", "solution_type"],
    "algo3_maturite": ["algo3_maturite", "maturity_stage", "algo3_maturity_stage"],
    "algo3_trl": ["algo3_trl", "trl_estimate", "algo3_trl_estimate", "TRL (1-9)", "TRL (1–9)"],
    "algo3_mrl": ["algo3_mrl", "mrl_estimate", "algo3_mrl_estimate", "MRL (1-9)", "MRL (1–9)"],
    "algo3_employes": [
        "algo3_employes",
        "employees_estimate",
        "nombre_employes",
        "Estimated Number of Employees",
    ],
    "algo3_seed_score": ["algo3_seed_score"],
    "algo3_catalyst_score": ["algo3_catalyst_score"],
    "algo3_materials_score": ["algo3_materials_score"],
    "commentaire_verification": ["commentaire_verification"],
}


def read_pipeline_dataframe(input_file: Path) -> tuple[pd.DataFrame, str]:
    xls = pd.ExcelFile(input_file)
    preferred_sheets = ["Synthese equipe", "Synthèse équipe", "Toutes les startups"]
    for sheet_name in preferred_sheets:
        if sheet_name in xls.sheet_names:
            return pd.read_excel(input_file, sheet_name=sheet_name), sheet_name

    return pd.read_excel(input_file, sheet_name=xls.sheet_names[0]), xls.sheet_names[0]


def build_global_result(status: str, algo3_orientation: str) -> str:
    status = clean_text(status)
    status_key = clean_decision(status)
    orientation = clean_text(algo3_orientation)
    if status_key == "ARRETEE_ALGO1_A_ECARTER":
        return "Ecartée par Algo1"
    if status_key == "ARRETEE_ALGO1_A_VERIFIER":
        return "À vérifier après Algo1"
    if status_key == "ARRETEE_ALGO2_NON_PREQUALIFIEE":
        return "Non préqualifiée par Algo2"
    if status_key == "ARRETEE_ALGO2_A_VERIFIER":
        return "À vérifier après Algo2"
    if status_key == "TERMINE_ALGO3":
        return f"Orientée Algo3 - {orientation or 'orientation inconnue'}"
    if status_key == "ERREUR_PIPELINE":
        return "Erreur pipeline"
    return status or "En cours / inconnu"


def pass_flag(decision: str, pass_values: set[str]) -> str:
    decision = clean_decision(decision)
    if not decision:
        return ""
    return "oui" if decision in pass_values else "non"


def record_from_row(row: pd.Series, column_map: dict[str, str | None], row_index: int) -> dict:
    record = {
        key: value_from_row(row, column_map, key)
        for key in FIELD_ALIASES
    }
    record["pipeline_index"] = record["pipeline_index"] or str(row_index + 1)
    record["algo1_passe_vers_algo2"] = record["algo1_passe_vers_algo2"] or pass_flag(
        record["algo1_decision_finale"],
        {"A_GARDER"},
    )
    record["algo2_passe_vers_algo3"] = record["algo2_passe_vers_algo3"] or pass_flag(
        record["algo2_decision_finale"],
        {"PREQUALIFIEE"},
    )
    record["resultat_global"] = record["resultat_global"] or build_global_result(
        record["statut_pipeline"],
        record["algo3_orientation_finale"],
    )
    record["search_blob"] = normalize_key(" ".join(clean_text(value) for value in record.values()))
    return record


def build_records(df: pd.DataFrame) -> list[dict]:
    column_map = {
        key: first_matching_column(df, aliases)
        for key, aliases in FIELD_ALIASES.items()
    }
    return [
        record_from_row(row, column_map, index)
        for index, row in df.iterrows()
    ]


def clean_label(value: str, fallback: str = "Non renseigné") -> str:
    value = clean_text(value)
    return value if value else fallback


def count_values(records: list[dict], field: str, fallback: str = "Non renseigné") -> list[dict]:
    counts: dict[str, int] = {}
    for record in records:
        value = clean_label(record.get(field, ""), fallback)
        counts[value] = counts.get(value, 0) + 1
    return [
        {"label": label, "value": value}
        for label, value in sorted(counts.items(), key=lambda item: (-item[1], item[0]))
    ]


def count_matching(records: list[dict], field: str, expected: str) -> int:
    expected = clean_decision(expected)
    return sum(1 for record in records if clean_decision(record.get(field, "")) == expected)


def contains_any_decision(record: dict, values: set[str]) -> bool:
    values = {clean_decision(value) for value in values}
    return any(
        clean_decision(record.get(field, "")) in values
        for field in [
            "algo1_decision_finale",
            "algo2_decision_finale",
            "algo3_orientation_finale",
            "statut_pipeline",
        ]
    )


def pct(part: int, total: int) -> int:
    if total <= 0:
        return 0
    return int(round((part / total) * 100))


def compact_reason(*values: str, limit: int = 220) -> str:
    for value in values:
        text = clean_text(value)
        if text:
            return text[:limit].rstrip() + ("..." if len(text) > limit else "")
    return ""


def score_value(value: str) -> float:
    try:
        parsed = float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return 0.0
    if math.isnan(parsed):
        return 0.0
    return max(0.0, min(1.0, parsed))


def priority_level(record: dict) -> tuple[int, str]:
    status = clean_decision(record.get("statut_pipeline", ""))
    algo2 = clean_decision(record.get("algo2_decision_finale", ""))
    algo3 = clean_decision(record.get("algo3_orientation_finale", ""))
    seed = score_value(record.get("algo3_seed_score", ""))
    catalyst = score_value(record.get("algo3_catalyst_score", ""))
    materials = score_value(record.get("algo3_materials_score", ""))
    sector = clean_text(record.get("algo2_secteur_confirme", ""))
    issue = clean_text(record.get("algo2_matrice_enjeu", ""))

    if status == "ERREUR_PIPELINE":
        return 1, "Erreur à reprendre"
    if algo2 == "PREQUALIFIEE" and algo3 == "A_VERIFIER":
        return 1, "Préqualifiée mais orientation Algo3 à trancher"
    if status == "ARRETEE_ALGO2_A_VERIFIER" and (sector or issue):
        return 1, "Pertinence possible, preuves Algo2 à compléter"
    if max(seed, catalyst, materials) >= 0.45 and algo3 == "A_VERIFIER":
        return 1, "Score programme intéressant mais preuve manquante"
    if clean_decision(record.get("algo1_decision_finale", "")) == "A_GARDER" and algo2 == "A_VERIFIER":
        return 2, "Gardée par Algo1, blocage Algo2"
    if status == "ARRETEE_ALGO1_A_VERIFIER":
        return 3, "Blocage dès Algo1"
    return 9, ""


def build_priority_records(records: list[dict], limit: int = 20) -> list[dict]:
    prioritized = []
    for record in records:
        rank, reason = priority_level(record)
        if rank >= 9:
            continue
        prioritized.append({
            "rank": rank,
            "startup": record.get("startup", ""),
            "sector": record.get("algo2_secteur_confirme") or record.get("secteur_source", ""),
            "status": record.get("resultat_global", ""),
            "reason": reason,
            "detail": compact_reason(
                record.get("raison_arret_ou_passage", ""),
                record.get("algo2_raison", ""),
                record.get("algo3_raison", ""),
                record.get("algo1_raison", ""),
            ),
            "next_action": record.get("prochaine_action", ""),
        })
    prioritized.sort(key=lambda item: (item["rank"], item["startup"]))
    return prioritized[:limit]


def orientation_sort_key(record: dict) -> tuple[int, float, str]:
    orientation = normalize_key(record.get("algo3_orientation_finale", ""))
    order = {
        "catalyst": 1,
        "materiaux": 2,
        "matériaux": 2,
        "seed": 3,
        "autre": 4,
    }
    score = max(
        score_value(record.get("algo3_seed_score", "")),
        score_value(record.get("algo3_catalyst_score", "")),
        score_value(record.get("algo3_materials_score", "")),
    )
    return order.get(orientation, 9), -score, clean_text(record.get("startup", ""))


def build_opportunities(records: list[dict], limit: int = 12) -> list[dict]:
    selected = [
        record
        for record in records
        if clean_decision(record.get("algo2_decision_finale", "")) == "PREQUALIFIEE"
        and clean_decision(record.get("algo3_orientation_finale", "")) not in {"", "A_VERIFIER"}
    ]
    selected.sort(key=orientation_sort_key)
    return [
        {
            "startup": record.get("startup", ""),
            "sector": record.get("algo2_secteur_confirme") or record.get("secteur_source", ""),
            "issue": record.get("algo2_matrice_enjeu", ""),
            "orientation": record.get("algo3_orientation_finale", ""),
            "reason": compact_reason(record.get("algo3_raison", ""), record.get("algo2_raison", "")),
        }
        for record in selected[:limit]
    ]


def build_funnel(records: list[dict]) -> list[dict]:
    total = len(records)
    algo1_garder = count_matching(records, "algo1_decision_finale", "A_GARDER")
    algo2_prequal = count_matching(records, "algo2_decision_finale", "PREQUALIFIEE")
    algo3_done = sum(
        1
        for record in records
        if clean_decision(record.get("algo3_execute", "")) == "OUI"
        or clean_text(record.get("algo3_orientation_finale", ""))
    )
    oriented = sum(
        1
        for record in records
        if clean_decision(record.get("algo3_orientation_finale", "")) in {"SEED", "CATALYST", "MATERIAUX", "AUTRE"}
    )
    return [
        {"label": "Startups", "value": total, "sub": "fichier source"},
        {"label": "Algo1 A_GARDER", "value": algo1_garder, "sub": f"{pct(algo1_garder, total)}% du total"},
        {"label": "Algo2 PREQUALIFIEE", "value": algo2_prequal, "sub": f"{pct(algo2_prequal, total)}% du total"},
        {"label": "Algo3 traité", "value": algo3_done, "sub": f"{pct(algo3_done, total)}% du total"},
        {"label": "Orientées", "value": oriented, "sub": f"{pct(oriented, total)}% du total"},
    ]


def build_summary(records: list[dict], input_file: Path, sheet_name: str) -> dict:
    total = len(records)
    algo1_garder = count_matching(records, "algo1_decision_finale", "A_GARDER")
    algo2_prequal = count_matching(records, "algo2_decision_finale", "PREQUALIFIEE")
    seed = count_matching(records, "algo3_orientation_finale", "Seed")
    catalyst = count_matching(records, "algo3_orientation_finale", "Catalyst")
    materials = count_matching(records, "algo3_orientation_finale", "Matériaux")
    to_verify = sum(
        1
        for record in records
        if contains_any_decision(record, {"A_VERIFIER", "ARRETEE_ALGO1_A_VERIFIER", "ARRETEE_ALGO2_A_VERIFIER"})
    )
    profile_exclusions = sum(
        1
        for record in records
        if clean_text(record.get("profile_exclusion_matches", ""))
    )

    return {
        "source_file": input_file.name,
        "source_sheet": sheet_name,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "kpis": [
            {"label": "Startups analysées", "value": total, "hint": "lignes pipeline"},
            {"label": "Gardées Algo1", "value": algo1_garder, "hint": f"{pct(algo1_garder, total)}% du total"},
            {"label": "Préqualifiées Algo2", "value": algo2_prequal, "hint": f"{pct(algo2_prequal, total)}% du total"},
            {"label": "Seed", "value": seed, "hint": "orientation Algo3"},
            {"label": "Catalyst", "value": catalyst, "hint": "orientation Algo3"},
            {"label": "Matériaux", "value": materials, "hint": "orientation Algo3"},
            {"label": "À vérifier", "value": to_verify, "hint": "toutes étapes confondues"},
            {"label": "Exclusions LATAM", "value": profile_exclusions, "hint": "scope hors profil"},
        ],
        "funnel": build_funnel(records),
        "charts": {
            "algo1": count_values(records, "algo1_decision_finale"),
            "algo2": count_values(records, "algo2_decision_finale"),
            "algo3": count_values(records, "algo3_orientation_finale"),
            "sectors": count_values(records, "algo2_secteur_confirme")[:8],
            "issues": count_values(records, "algo2_matrice_enjeu")[:8],
            "status": count_values(records, "resultat_global")[:10],
        },
        "priority": build_priority_records(records),
        "opportunities": build_opportunities(records),
    }


def build_dashboard_html(data: dict) -> str:
    data_json = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    title = escape(f"Dashboard dealflow - {data['summary']['source_file']}")
    logo_uri = asset_data_uri(
        "leonard-vinci-logo.svg",
        "image/svg+xml",
        fix_svg=True,
        replacements={
            "#c41d6d": "#ffffff",
            "#00b4ff": "#ffffff",
            "#004489": "#ffffff",
            "#ff005a": "#ffffff",
        },
    )
    orbit_uri = asset_data_uri("constellation-orbit.svg", "image/svg+xml", fix_svg=True)
    orbit_blue_uri = asset_data_uri(
        "constellation-orbit.svg",
        "image/svg+xml",
        fix_svg=True,
        replacements={"#fff": "#4A7CC9"},
    )
    template = r"""<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>__TITLE__</title>
  <style>
    __VINCI_FONT_CSS__
    :root {
      --bg: #f3f6f8;
      --panel: #ffffff;
      --ink: #102b45;
      --muted: #617487;
      --line: #d7e2ea;
      --green: #18a566;
      --teal: #00857c;
      --blue: #004489;
      --cyan: #00b4ff;
      --pink: #ff005a;
      --amber: #b7791f;
      --red: #b23b3b;
      --violet: #6950a1;
      --constellation-field: #62636f;
      --constellation-blue: #4a7cc9;
      --constellation-violet: #6b4c9a;
      --constellation-rose: #d64c7e;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--bg);
      color: var(--ink);
      font-family: "Vinci Sans", Arial, sans-serif;
      line-height: 1.45;
      letter-spacing: 0;
      position: relative;
      isolation: isolate;
      overflow-x: hidden;
    }
    body::before {
      content: "";
      position: fixed;
      right: -260px;
      bottom: -360px;
      width: min(980px, 80vw);
      aspect-ratio: 1;
      background: url("__ORBIT_BLUE_URI__") center / contain no-repeat;
      opacity: 0.08;
      pointer-events: none;
      z-index: -1;
    }
    .shell { max-width: 1400px; margin: 0 auto; padding: 16px 20px 32px; }
    header {
      position: relative;
      display: flex;
      justify-content: space-between;
      gap: 20px;
      align-items: center;
      margin-bottom: 14px;
      padding: 16px 18px;
      min-height: 108px;
      overflow: hidden;
      background: var(--constellation-field);
      border-top: 2px solid var(--cyan);
      border-bottom: 2px solid var(--pink);
      border-radius: 6px;
    }
    .header-orbit {
      position: absolute;
      width: 420px;
      height: 420px;
      right: -90px;
      top: -220px;
      opacity: 0.22;
      pointer-events: none;
    }
    .brand-lockup { position: relative; z-index: 1; display: flex; align-items: center; gap: 16px; min-width: 0; }
    .brand-lockup img:not(.header-orbit) { display: block; width: 146px; height: auto; flex-shrink: 0; }
    .brand-copy { min-width: 0; padding-left: 16px; border-left: 1px solid rgba(255,255,255,0.35); }
    .brand-kicker { color: var(--cyan); font-size: 10px; font-weight: 700; text-transform: uppercase; margin-bottom: 2px; }
    h1, .kpi-value, .step-value { font-family: "Vinci Serif", Georgia, serif; }
    h1 { margin: 0; color: #ffffff; font-size: 24px; font-weight: 400; letter-spacing: 0; }
    h2 { margin: 0 0 12px; font-size: 15px; font-weight: 700; letter-spacing: 0; color: var(--blue); }
    p { margin: 0; }
    .meta { color: var(--muted); font-size: 12px; margin-top: 4px; }
    header .meta { color: rgba(255,255,255,0.72); }
    .toolbar {
      position: relative;
      z-index: 1;
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      justify-content: flex-end;
    }
    button, input, select {
      border: 1px solid var(--line);
      background: var(--panel);
      color: var(--ink);
      min-height: 38px;
      border-radius: 5px;
      font: inherit;
      font-size: 13px;
    }
    button { padding: 0 12px; cursor: pointer; transition: border-color 150ms ease, color 150ms ease, background 150ms ease; }
    button:hover { border-color: var(--pink); color: var(--pink); }
    #exportCsv { color: #ffffff; background: var(--pink); border-color: var(--pink); }
    input, select { padding: 0 10px; min-width: 190px; }
    .grid { display: grid; gap: 12px; }
    .kpis { grid-template-columns: repeat(4, minmax(0, 1fr)); margin-bottom: 12px; }
    .card {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 6px;
      box-shadow: 0 1px 3px rgba(16,43,69,0.05);
    }
    .kpi { padding: 12px 14px 12px 17px; min-height: 92px; border-left: 3px solid var(--cyan); }
    .kpi:nth-child(3n+2) { border-left-color: var(--pink); }
    .kpi:nth-child(3n) { border-left-color: var(--blue); }
    .kpi-value { font-size: 27px; font-weight: 400; margin-top: 5px; }
    .kpi-label { font-size: 12px; color: var(--muted); text-transform: uppercase; }
    .kpi-hint { color: var(--muted); font-size: 12px; margin-top: 2px; }
    .main { grid-template-columns: 1.25fr 0.75fr; align-items: start; }
    .panel { padding: 16px 17px; }
    .funnel {
      display: grid;
      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: 10px;
    }
    .step {
      border: 0;
      border-left: 4px solid var(--step-color, var(--green));
      border-radius: 0;
      padding: 10px 12px;
      min-height: 118px;
      position: relative;
      overflow: hidden;
      background: transparent;
    }
    .step::before {
      display: none;
    }
    .step-label { color: var(--muted); font-size: 12px; }
    .step-value { font-weight: 820; font-size: 28px; margin: 8px 0 2px; }
    .step-sub { color: var(--muted); font-size: 12px; }
    .charts { grid-template-columns: repeat(3, minmax(0, 1fr)); margin-top: 12px; }
    .bar-row {
      display: grid;
      grid-template-columns: minmax(110px, 1fr) 4fr 40px;
      gap: 10px;
      align-items: center;
      margin: 10px 0;
      font-size: 13px;
    }
    .bar-track { height: 11px; border-radius: 2px; background: #e8eef3; overflow: hidden; }
    .bar-fill { height: 100%; border-radius: 2px; background: var(--bar-color, var(--blue)); min-width: 2px; }
    .label { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: var(--ink); }
    .value { text-align: right; color: var(--muted); font-variant-numeric: tabular-nums; }
    .split { grid-template-columns: 1fr 1fr; margin-top: 12px; align-items: start; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th {
      text-align: left;
      color: var(--muted);
      font-weight: 700;
      border-bottom: 1px solid var(--line);
      padding: 10px 8px;
      white-space: nowrap;
    }
    td { border-bottom: 1px solid #e8eef3; padding: 10px 8px; vertical-align: top; }
    tr:hover td { background: #f5fbfe; }
    .pill {
      display: inline-flex;
      align-items: center;
      min-height: 24px;
      border-radius: 5px;
      padding: 2px 9px;
      font-size: 12px;
      border: 1px solid var(--line);
      background: #f3f6f8;
      white-space: nowrap;
    }
    .pill.green { color: #0f6a48; background: #edf8f2; border-color: #c8ead8; }
    .pill.amber { color: #875409; background: #fff8e8; border-color: #f0dfb5; }
    .pill.red { color: #963333; background: #fff1f1; border-color: #efcaca; }
    .pill.blue { color: #244f9c; background: #eff5ff; border-color: #cbdcf8; }
    .pill.violet { color: #643d95; background: #f5f0ff; border-color: #ddcff4; }
    .muted { color: var(--muted); }
    .controls {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      align-items: center;
      margin-bottom: 12px;
    }
    .wide-table {
      overflow: auto;
      max-height: 620px;
      border: 1px solid var(--line);
      border-radius: 5px;
    }
    .wide-table table { min-width: 1360px; }
    .wide-table thead th {
      position: sticky;
      top: 0;
      background: #f5f9fc;
      z-index: 1;
    }
    .footer { color: var(--muted); font-size: 12px; margin-top: 16px; }
    @media (max-width: 1100px) {
      .kpis { grid-template-columns: repeat(3, minmax(0, 1fr)); }
      .main, .split { grid-template-columns: 1fr; }
      .charts { grid-template-columns: 1fr; }
      .funnel { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      header { flex-direction: column; }
      .brand-lockup { align-items: flex-start; flex-direction: column; }
      .brand-copy { padding-left: 0; border-left: 0; }
      .toolbar { justify-content: flex-start; }
    }
    @media (max-width: 640px) {
      .shell { padding: 14px; }
      .kpis { grid-template-columns: 1fr; }
      .funnel { grid-template-columns: 1fr; }
      input, select { min-width: 100%; }
      .controls { display: grid; grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <div class="shell">
    <header>
      <img class="header-orbit" src="__ORBIT_URI__" alt="">
      <div class="brand-lockup">
        <img src="__LOGO_URI__" alt="Leonard, powered by VINCI">
        <div class="brand-copy">
          <div class="brand-kicker">Dealflow constellation</div>
          <h1>Dashboard dealflow startups</h1>
          <p class="meta"><span id="sourceFile"></span> · généré le <span id="generatedAt"></span> · sheet source <span id="sourceSheet"></span></p>
        </div>
      </div>
      <div class="toolbar">
        <button id="showPriority">À vérifier</button>
        <button id="resetFilters">Réinitialiser</button>
        <button id="exportCsv">Exporter CSV filtré</button>
      </div>
    </header>

    <section id="kpis" class="grid kpis"></section>

    <section class="grid main">
      <div class="card panel">
        <h2>Funnel de qualification</h2>
        <div id="funnel" class="funnel"></div>
      </div>
      <div class="card panel">
        <h2>Résultat global</h2>
        <div id="statusChart"></div>
      </div>
    </section>

    <section class="grid charts">
      <div class="card panel"><h2>Algo1</h2><div id="algo1Chart"></div></div>
      <div class="card panel"><h2>Algo2</h2><div id="algo2Chart"></div></div>
      <div class="card panel"><h2>Algo3</h2><div id="algo3Chart"></div></div>
    </section>

    <section class="grid split">
      <div class="card panel">
        <h2>À revoir en priorité</h2>
        <div id="priorityTable"></div>
      </div>
      <div class="card panel">
        <h2>Top opportunités</h2>
        <div id="opportunityTable"></div>
      </div>
    </section>

    <section class="grid split">
      <div class="card panel"><h2>Secteurs confirmés Algo2</h2><div id="sectorChart"></div></div>
      <div class="card panel"><h2>Matrices d’enjeux</h2><div id="issueChart"></div></div>
    </section>

    <section class="card panel" style="margin-top: 14px;">
      <h2>Vue équipe filtrable</h2>
      <div class="controls">
        <input id="searchInput" type="search" placeholder="Rechercher startup, secteur, profil, raison...">
        <select id="statusFilter"><option value="">Tous les résultats</option></select>
        <select id="sectorFilter"><option value="">Tous les secteurs</option></select>
        <select id="orientationFilter"><option value="">Toutes les orientations</option></select>
      </div>
      <p class="meta" id="filteredCount"></p>
      <div class="wide-table"><table id="recordsTable"></table></div>
    </section>

    <p class="footer">Dashboard HTML autonome généré depuis le fichier pipeline. Les décisions restent celles des algorithmes et doivent être validées humainement avant action engageante.</p>
  </div>

  <script>
    const DATA = __DASHBOARD_DATA__;
    const COLORS = ["#4a7cc9", "#6b4c9a", "#d64c7e", "#00b4ff", "#ff005a", "#00857c", "#b7791f", "#617487"];
    const records = DATA.records || [];

    function esc(value) {
      return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
    }

    function norm(value) {
      return String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
    }

    function pillClass(value) {
      const v = norm(value);
      if (v.includes("prequal") || v.includes("garder") || v === "seed") return "green";
      if (v.includes("verifier") || v.includes("pending")) return "amber";
      if (v.includes("ecart") || v.includes("non")) return "red";
      if (v.includes("catalyst")) return "blue";
      if (v.includes("mater")) return "violet";
      return "";
    }

    function pill(value) {
      if (!String(value ?? "").trim()) return "";
      return `<span class="pill ${pillClass(value)}">${esc(value)}</span>`;
    }

    function renderKpis() {
      document.getElementById("kpis").innerHTML = (DATA.summary.kpis || []).map(kpi => `
        <article class="card kpi">
          <div class="kpi-label">${esc(kpi.label)}</div>
          <div class="kpi-value">${esc(kpi.value)}</div>
          <div class="kpi-hint">${esc(kpi.hint)}</div>
        </article>
      `).join("");
    }

    function renderFunnel() {
      const colors = ["#4a7cc9", "#6b4c9a", "#d64c7e", "#00b4ff", "#ff005a"];
      document.getElementById("funnel").innerHTML = (DATA.summary.funnel || []).map((step, index) => `
        <article class="step" style="--step-color:${colors[index % colors.length]}">
          <div class="step-label">${esc(step.label)}</div>
          <div class="step-value">${esc(step.value)}</div>
          <div class="step-sub">${esc(step.sub)}</div>
        </article>
      `).join("");
    }

    function renderBarChart(id, rows) {
      const max = Math.max(1, ...rows.map(row => Number(row.value) || 0));
      document.getElementById(id).innerHTML = rows.length ? rows.map((row, index) => {
        const width = Math.max(4, Math.round(((Number(row.value) || 0) / max) * 100));
        return `
          <div class="bar-row" title="${esc(row.label)}">
            <div class="label">${esc(row.label)}</div>
            <div class="bar-track"><div class="bar-fill" style="width:${width}%;--bar-color:${COLORS[index % COLORS.length]}"></div></div>
            <div class="value">${esc(row.value)}</div>
          </div>
        `;
      }).join("") : `<p class="muted">Aucune donnée.</p>`;
    }

    function renderSmallTable(id, rows, columns) {
      const table = document.getElementById(id);
      if (!rows.length) {
        table.innerHTML = `<p class="muted">Aucune ligne à afficher.</p>`;
        return;
      }
      table.innerHTML = `
        <table>
          <thead><tr>${columns.map(col => `<th>${esc(col.label)}</th>`).join("")}</tr></thead>
          <tbody>
            ${rows.map(row => `<tr>${columns.map(col => `<td>${col.render ? col.render(row[col.key], row) : esc(row[col.key] || "")}</td>`).join("")}</tr>`).join("")}
          </tbody>
        </table>
      `;
    }

    function uniqueValues(field) {
      return [...new Set(records.map(row => String(row[field] || "").trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
    }

    function fillSelect(id, values) {
      const select = document.getElementById(id);
      const first = select.options[0].outerHTML;
      select.innerHTML = first + values.map(value => `<option value="${esc(value)}">${esc(value)}</option>`).join("");
    }

    function filteredRecords() {
      const q = norm(document.getElementById("searchInput").value);
      const status = document.getElementById("statusFilter").value;
      const sector = document.getElementById("sectorFilter").value;
      const orientation = document.getElementById("orientationFilter").value;
      return records.filter(row => {
        if (q && !norm(row.search_blob).includes(q)) return false;
        if (status && row.resultat_global !== status) return false;
        const rowSector = row.algo2_secteur_confirme || row.secteur_source || "";
        if (sector && rowSector !== sector) return false;
        if (orientation && row.algo3_orientation_finale !== orientation) return false;
        return true;
      });
    }

    function renderRecords() {
      const rows = filteredRecords();
      document.getElementById("filteredCount").textContent = `${rows.length} startup(s) affichée(s) sur ${records.length}`;
      const columns = [
        ["startup", "Startup"],
        ["dealflow_profile", "Profil"],
        ["resultat_global", "Résultat"],
        ["profile_exclusion_matches", "Exclusions LATAM"],
        ["algo1_decision_finale", "Algo1"],
        ["algo1_raison", "Raison Algo1"],
        ["algo2_decision_finale", "Algo2"],
        ["algo2_raison", "Raison Algo2"],
        ["algo3_orientation_finale", "Algo3"],
        ["algo3_raison", "Raison Algo3"],
        ["prochaine_action", "Action"]
      ];
      const html = `
        <thead><tr>${columns.map(([_, label]) => `<th>${esc(label)}</th>`).join("")}</tr></thead>
        <tbody>${rows.map(row => `
          <tr>
            ${columns.map(([key]) => {
              const value = row[key] || "";
              const isDecision = key.includes("decision") || key.includes("orientation") || key === "resultat_global";
              return `<td>${isDecision ? pill(value) : esc(value)}</td>`;
            }).join("")}
          </tr>
        `).join("")}</tbody>
      `;
      document.getElementById("recordsTable").innerHTML = html;
    }

    function exportCsv() {
      const rows = filteredRecords();
      const columns = ["startup", "dealflow_profile", "resultat_global", "profile_exclusion_matches", "algo1_decision_finale", "algo1_raison", "algo2_decision_finale", "algo2_raison", "algo3_orientation_finale", "algo3_raison", "prochaine_action"];
      const csv = [
        columns.join(";"),
        ...rows.map(row => columns.map(col => `"${String(row[col] || "").replaceAll('"', '""')}"`).join(";"))
      ].join("\n");
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "dashboard_filtre.csv";
      a.click();
      URL.revokeObjectURL(url);
    }

    function init() {
      document.getElementById("sourceFile").textContent = DATA.summary.source_file;
      document.getElementById("generatedAt").textContent = DATA.summary.generated_at;
      document.getElementById("sourceSheet").textContent = DATA.summary.source_sheet;
      renderKpis();
      renderFunnel();
      renderBarChart("statusChart", DATA.summary.charts.status || []);
      renderBarChart("algo1Chart", DATA.summary.charts.algo1 || []);
      renderBarChart("algo2Chart", DATA.summary.charts.algo2 || []);
      renderBarChart("algo3Chart", DATA.summary.charts.algo3 || []);
      renderBarChart("sectorChart", DATA.summary.charts.sectors || []);
      renderBarChart("issueChart", DATA.summary.charts.issues || []);
      renderSmallTable("priorityTable", DATA.summary.priority || [], [
        { key: "startup", label: "Startup" },
        { key: "status", label: "Statut", render: pill },
        { key: "reason", label: "Pourquoi" },
        { key: "detail", label: "Détail" }
      ]);
      renderSmallTable("opportunityTable", DATA.summary.opportunities || [], [
        { key: "startup", label: "Startup" },
        { key: "orientation", label: "Orientation", render: pill },
        { key: "sector", label: "Secteur" },
        { key: "reason", label: "Raison" }
      ]);
      fillSelect("statusFilter", uniqueValues("resultat_global"));
      fillSelect("sectorFilter", [...new Set(records.map(row => row.algo2_secteur_confirme || row.secteur_source || "").filter(Boolean))].sort((a, b) => a.localeCompare(b)));
      fillSelect("orientationFilter", uniqueValues("algo3_orientation_finale"));
      ["searchInput", "statusFilter", "sectorFilter", "orientationFilter"].forEach(id => {
        document.getElementById(id).addEventListener("input", renderRecords);
        document.getElementById(id).addEventListener("change", renderRecords);
      });
      document.getElementById("resetFilters").addEventListener("click", () => {
        document.getElementById("searchInput").value = "";
        document.getElementById("statusFilter").value = "";
        document.getElementById("sectorFilter").value = "";
        document.getElementById("orientationFilter").value = "";
        renderRecords();
      });
      document.getElementById("showPriority").addEventListener("click", () => {
        document.getElementById("searchInput").value = "verifier";
        document.getElementById("statusFilter").value = "";
        document.getElementById("sectorFilter").value = "";
        document.getElementById("orientationFilter").value = "";
        renderRecords();
      });
      document.getElementById("exportCsv").addEventListener("click", exportCsv);
      renderRecords();
    }

    init();
  </script>
</body>
</html>
"""
    return (
        template
        .replace("__TITLE__", title)
        .replace("__VINCI_FONT_CSS__", dashboard_font_css())
        .replace("__LOGO_URI__", logo_uri)
        .replace("__ORBIT_URI__", orbit_uri)
        .replace("__ORBIT_BLUE_URI__", orbit_blue_uri)
        .replace("__DASHBOARD_DATA__", data_json)
    )


def generate_dashboard(input_file: Path) -> Path:
    df, sheet_name = read_pipeline_dataframe(input_file)
    records = build_records(df)
    summary = build_summary(records, input_file, sheet_name)
    data = {
        "summary": summary,
        "records": records,
    }

    output_file = make_output_path(input_file)
    output_file.write_text(build_dashboard_html(data), encoding="utf-8")
    return output_file


def main() -> None:
    input_files = discover_input_files(Path.cwd(), sys.argv[1:])
    if not input_files:
        log("Aucun fichier pipeline trouvé.")
        log('Exemple : python -B Generer_dashboard.py "mon_fichier_pipeline_algo_1_2_3.xlsx"')
        return

    for input_file in input_files:
        try:
            log(f"Lecture : {input_file.name}")
            output_file = generate_dashboard(input_file)
            log(f"Dashboard généré : {output_file.name}")
        except Exception as exc:
            log(f"Erreur sur {input_file.name} : {exc}")


if __name__ == "__main__":
    main()
