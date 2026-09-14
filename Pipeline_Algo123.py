import os
import json
import sys
import threading
import time
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from tqdm import tqdm

import Algo1
import Algo2
import Algo3
import dealflow_profiles
import Generer_dashboard
import excel_styling


load_dotenv()


OUTPUT_SUFFIX = "_pipeline_algo_1_2_3.xlsx"
TEMP_SUFFIX = "_pipeline_algo_1_2_3_sauvegarde_temp.xlsx"
SAVE_EVERY = int(os.getenv("PIPELINE_SAVE_EVERY", "5"))
LIMIT_ROWS = int(os.getenv("PIPELINE_LIMIT_ROWS", "0"))
GENERATE_DASHBOARD = os.getenv("PIPELINE_GENERATE_DASHBOARD", "1").strip().lower() not in {
    "0",
    "false",
    "non",
    "no",
    "off",
}
RESUME_FROM_TEMP = os.getenv("PIPELINE_RESUME_FROM_TEMP", "1").strip().lower() not in {
    "0",
    "false",
    "non",
    "no",
    "off",
}
STREAMLIT_PROGRESS = os.getenv("PIPELINE_STREAMLIT_PROGRESS", "0").strip().lower() in {
    "1",
    "true",
    "oui",
    "yes",
    "on",
}

ALGO1_PASS_DECISIONS = {
    item.strip().upper()
    for item in os.getenv("PIPELINE_ALGO1_PASS_DECISIONS", "A_GARDER").split(",")
    if item.strip()
}
ALGO2_PASS_DECISIONS = {
    item.strip().upper()
    for item in os.getenv("PIPELINE_ALGO2_PASS_DECISIONS", "PREQUALIFIEE").split(",")
    if item.strip()
}

EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}

PIPELINE_COLUMNS = [
    "pipeline_index",
    "pipeline_stage",
    "pipeline_status",
    "pipeline_stop_reason",
    "pipeline_next_action",
    "pipeline_algo1_executed",
    "pipeline_algo2_executed",
    "pipeline_algo3_executed",
    "pipeline_started_at",
    "pipeline_finished_at",
    "pipeline_elapsed_seconds",
    "manual_algo1_decision",
    "manual_algo2_decision",
    "manual_algo3_orientation",
    "manual_comment",
    "manual_validated_by",
    "manual_validated_at",
]


def log(message: str) -> None:
    print(f"[Pipeline Algo123] {message}")


def emit_progress(
    done: int,
    total: int,
    started: float,
    stage: str = "Analyse",
    baseline_done: int = 0,
) -> None:
    if not STREAMLIT_PROGRESS:
        return
    elapsed = max(0.0, time.time() - started)
    eta = None
    active_done = max(0, done - baseline_done)
    remaining = max(0, total - done)
    if active_done > 0 and remaining > 0:
        eta = (elapsed / active_done) * remaining
    payload = {
        "done": done,
        "total": total,
        "percent": round((done / total) * 100, 1) if total else 0,
        "elapsed_seconds": round(elapsed, 1),
        "eta_seconds": round(eta, 1) if eta is not None else None,
        "stage": stage,
    }
    print("__PIPELINE_PROGRESS__ " + json.dumps(payload, ensure_ascii=False), flush=True)


def clean_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def clean_decision(value) -> str:
    return clean_text(value).upper()


def now_text() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


def is_generated_excel(path: Path) -> bool:
    name = path.name.lower()
    return (
        name.startswith("~$")
        or "_algo_" in name
        or "_pipeline_algo_1_2_3" in name
    )


def discover_input_files(base_dir: Path, cli_args: list[str]) -> list[Path]:
    if cli_args:
        return [Path(arg).expanduser().resolve() for arg in cli_args]

    return [
        path
        for path in sorted(base_dir.iterdir())
        if path.is_file()
        and path.suffix.lower() in EXCEL_EXTENSIONS
        and not is_generated_excel(path)
    ]


def make_output_path(input_file: Path) -> Path:
    return input_file.with_name(f"{input_file.stem}{OUTPUT_SUFFIX}")


def make_temp_path(input_file: Path) -> Path:
    return input_file.with_name(f"{input_file.stem}{TEMP_SUFFIX}")


def normalize_value(value):
    if pd.isna(value):
        return ""
    return value


def row_to_dict(row: pd.Series) -> dict:
    return {str(key): normalize_value(value) for key, value in row.to_dict().items()}


def add_missing_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    prepared = df.copy()
    for column in columns:
        if column not in prepared.columns:
            prepared[column] = ""
    return prepared


def prepare_detection_frame(df: pd.DataFrame) -> pd.DataFrame:
    prepared = df.copy()
    prepared = add_missing_columns(prepared, PIPELINE_COLUMNS)
    prepared = add_missing_columns(prepared, Algo1.RESULT_COLUMNS)
    prepared = add_missing_columns(prepared, prefixed_columns("algo1", Algo1.RESULT_COLUMNS))
    prepared = add_missing_columns(prepared, Algo2.RESULT_COLUMNS)
    prepared = add_missing_columns(prepared, prefixed_columns("algo2", Algo2.RESULT_COLUMNS))
    prepared = add_missing_columns(prepared, Algo3.RESULT_COLUMNS)
    return prepared


def prefixed_columns(prefix: str, columns: list[str]) -> list[str]:
    return [column if column.startswith(f"{prefix}_") else f"{prefix}_{column}" for column in columns]


def with_prefixed_result(prefix: str, result: dict) -> dict:
    prefixed = {}
    for key, value in result.items():
        prefixed_key = key if key.startswith(f"{prefix}_") else f"{prefix}_{key}"
        prefixed[prefixed_key] = value
    return prefixed


def update_with_algo1(row_data: dict, result: dict) -> None:
    row_data.update(result)
    row_data.update(with_prefixed_result("algo1", result))
    row_data["algo1_decision"] = result.get("decision", "")
    row_data["algo1_decision_finale"] = result.get("decision_finale", "")
    row_data["algo1_reason"] = result.get("reason", "")
    row_data["algo1_missing_information"] = result.get("missing_information", "")
    row_data["algo1_regle_finale"] = result.get("regle_finale", "")


def update_with_algo2(row_data: dict, result: dict) -> None:
    row_data.update(result)
    row_data.update(with_prefixed_result("algo2", result))
    row_data["algo2_decision"] = result.get("algo2_decision", "")
    row_data["algo2_decision_finale"] = result.get("algo2_decision_finale", "")
    row_data["algo2_reason"] = result.get("reason", "")
    row_data["algo2_missing_information"] = result.get("missing_information", "")
    row_data["algo2_regle_finale"] = result.get("algo2_regle_finale", "")


def update_with_algo3(row_data: dict, result: dict) -> None:
    row_data.update(result)
    row_data.update(with_prefixed_result("algo3", result))
    row_data["algo3_orientation_finale"] = result.get("algo3_orientation_finale", "")
    row_data["algo3_orientation_reason"] = result.get("orientation_reason", "")
    row_data["algo3_missing_information"] = result.get("missing_information", "")
    row_data["algo3_regle_finale"] = result.get("algo3_regle_finale", "")


def make_series(row_data: dict) -> pd.Series:
    return pd.Series(row_data)


def stop_pipeline(
    row_data: dict,
    stage: str,
    status: str,
    reason: str,
    next_action: str,
) -> dict:
    row_data["pipeline_stage"] = stage
    row_data["pipeline_status"] = status
    row_data["pipeline_stop_reason"] = reason
    row_data["pipeline_next_action"] = next_action
    return row_data


def process_pipeline_row(
    source_row: pd.Series,
    index: int,
    columns_algo1: dict,
    columns_algo2: dict,
    columns_algo3: dict,
    website_cache: dict,
    website_cache_lock: threading.Lock,
    algo2_research_cache: dict,
    algo2_research_cache_lock: threading.Lock,
    algo3_research_cache: dict,
    algo3_research_cache_lock: threading.Lock,
) -> dict:
    started = time.time()
    row_data = row_to_dict(source_row)
    row_data.update({
        "pipeline_index": index + 1,
        "pipeline_stage": "ALGO1",
        "pipeline_status": "EN_COURS",
        "pipeline_stop_reason": "",
        "pipeline_next_action": "",
        "pipeline_algo1_executed": "non",
        "pipeline_algo2_executed": "non",
        "pipeline_algo3_executed": "non",
        "pipeline_started_at": now_text(),
        "pipeline_finished_at": "",
        "pipeline_elapsed_seconds": "",
    })
    for column in PIPELINE_COLUMNS:
        row_data.setdefault(column, "")

    try:
        algo1_result = Algo1.process_startup_row(
            make_series(row_data),
            columns_algo1,
            website_cache,
            website_cache_lock,
        )
        row_data["pipeline_algo1_executed"] = "oui"
        update_with_algo1(row_data, algo1_result)

        algo1_decision = clean_decision(algo1_result.get("decision_finale"))
        if algo1_decision not in ALGO1_PASS_DECISIONS:
            if algo1_decision == "A_ECARTER":
                status = "ARRETEE_ALGO1_A_ECARTER"
                next_action = "Non transmise a Algo2."
            else:
                status = "ARRETEE_ALGO1_A_VERIFIER"
                next_action = "Verification manuelle avant Algo2."
            return finalize_row(
                stop_pipeline(
                    row_data,
                    "ALGO1",
                    status,
                    algo1_result.get("reason") or algo1_result.get("regle_finale", ""),
                    next_action,
                ),
                started,
            )

        row_data["pipeline_stage"] = "ALGO2"
        algo2_result = Algo2.process_startup_row(
            make_series(row_data),
            columns_algo2,
            algo2_research_cache,
            algo2_research_cache_lock,
        )
        row_data["pipeline_algo2_executed"] = "oui"
        update_with_algo2(row_data, algo2_result)

        algo2_decision = clean_decision(algo2_result.get("algo2_decision_finale"))
        if algo2_decision not in ALGO2_PASS_DECISIONS:
            if algo2_decision == "NON_PREQUALIFIEE":
                status = "ARRETEE_ALGO2_NON_PREQUALIFIEE"
                next_action = "Non transmise a Algo3."
            else:
                status = "ARRETEE_ALGO2_A_VERIFIER"
                next_action = "Verification manuelle avant orientation Algo3."
            return finalize_row(
                stop_pipeline(
                    row_data,
                    "ALGO2",
                    status,
                    algo2_result.get("reason") or algo2_result.get("algo2_regle_finale", ""),
                    next_action,
                ),
                started,
            )

        row_data["pipeline_stage"] = "ALGO3"
        algo3_result = Algo3.process_startup_row(
            make_series(row_data),
            columns_algo3,
            algo3_research_cache,
            algo3_research_cache_lock,
        )
        row_data["pipeline_algo3_executed"] = "oui"
        update_with_algo3(row_data, algo3_result)

        return finalize_row(
            stop_pipeline(
                row_data,
                "ALGO3",
                "TERMINE_ALGO3",
                algo3_result.get("algo3_regle_finale", ""),
                "Orientation programme disponible.",
            ),
            started,
        )

    except Exception as exc:
        return finalize_row(
            stop_pipeline(
                row_data,
                row_data.get("pipeline_stage", "INCONNU"),
                "ERREUR_PIPELINE",
                f"{exc.__class__.__name__}: {exc}",
                "Relancer avec reprise ou verifier la ligne manuellement.",
            ),
            started,
        )


def finalize_row(row_data: dict, started: float) -> dict:
    row_data["pipeline_finished_at"] = now_text()
    row_data["pipeline_elapsed_seconds"] = round(time.time() - started, 2)
    return row_data


def load_resume_results(temp_file: Path) -> list[dict]:
    if not RESUME_FROM_TEMP or not temp_file.exists():
        return []

    try:
        temp_df = pd.read_excel(temp_file, sheet_name="Toutes les startups")
    except Exception as exc:
        log(f"Reprise impossible depuis {temp_file.name} : {exc}")
        return []

    if temp_df.empty:
        return []

    records = [row_to_dict(row) for _, row in temp_df.iterrows()]
    log(f"Reprise depuis {temp_file.name} : {len(records)} ligne(s) deja traitees.")
    return records


def sheet_safe(df: pd.DataFrame, mask) -> pd.DataFrame:
    try:
        return df[mask].copy()
    except Exception:
        return df.iloc[0:0].copy()


def normalize_header(value: str) -> str:
    return clean_text(value).lower().replace("_", " ").replace("-", " ")


def first_matching_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    if df.empty:
        return None

    exact = {str(column): str(column) for column in df.columns}
    for candidate in candidates:
        if candidate in exact:
            return exact[candidate]

    normalized_columns = {
        normalize_header(str(column)): str(column)
        for column in df.columns
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


def summary_series(df: pd.DataFrame, candidates: list[str]) -> pd.Series:
    column = first_matching_column(df, candidates)
    if column and column in df.columns:
        return df[column].fillna("").astype(str)
    return pd.Series([""] * len(df), index=df.index, dtype="object")


def yes_no_from_execution(df: pd.DataFrame, executed_column: str, positive_column: str = "") -> pd.Series:
    executed = df.get(executed_column, pd.Series([""] * len(df), index=df.index)).fillna("").astype(str).str.lower()
    if not positive_column:
        return executed.map(lambda value: "oui" if value == "oui" else "non")
    positive = df.get(positive_column, pd.Series([""] * len(df), index=df.index)).fillna("").astype(str)
    return [
        "oui" if exec_value.lower() == "oui" and clean_text(pos_value) else ("non" if exec_value.lower() == "oui" else "")
        for exec_value, pos_value in zip(executed, positive)
    ]


def build_global_result(status: str, algo3_orientation: str) -> str:
    status = clean_text(status)
    if status == "ARRETEE_ALGO1_A_ECARTER":
        return "Ecartee par Algo1"
    if status == "ARRETEE_ALGO1_A_VERIFIER":
        return "A verifier apres Algo1"
    if status == "ARRETEE_ALGO2_NON_PREQUALIFIEE":
        return "Non prequalifiee par Algo2"
    if status == "ARRETEE_ALGO2_A_VERIFIER":
        return "A verifier apres Algo2"
    if status == "TERMINE_ALGO3":
        return f"Orientee Algo3 - {clean_text(algo3_orientation) or 'orientation inconnue'}"
    if status == "ERREUR_PIPELINE":
        return "Erreur pipeline"
    return status or "En cours / inconnu"


def build_team_summary(final_df: pd.DataFrame) -> pd.DataFrame:
    summary = pd.DataFrame(index=final_df.index)

    summary["pipeline_index"] = final_df.get(
        "pipeline_index",
        pd.Series(range(1, len(final_df) + 1), index=final_df.index),
    )
    summary["startup"] = summary_series(final_df, [
        "startup",
        "Startup",
        "Nom startup",
        "Nom de la startup",
        "company",
        "Company",
        "Company name",
        "Startup name",
        "name",
        "Name",
    ])
    summary["site_web"] = summary_series(final_df, [
        "site",
        "site web",
        "website",
        "Website",
        "url",
        "URL",
        "company website",
    ])
    summary["secteur_source"] = summary_series(final_df, [
        "sector",
        "secteur",
        "industry",
        "vertical",
        "category",
        "categorie",
    ])
    summary["dealflow_profile"] = summary_series(final_df, [
        "dealflow_profile",
        "algo1_dealflow_profile",
        "algo2_dealflow_profile",
        "algo3_dealflow_profile",
    ])
    summary["profile_exclusion_matches"] = summary_series(final_df, [
        "profile_exclusion_matches",
        "algo1_profile_exclusion_matches",
        "algo2_profile_exclusion_matches",
        "algo3_profile_exclusion_matches",
    ])

    status = final_df.get("pipeline_status", pd.Series([""] * len(final_df), index=final_df.index)).fillna("").astype(str)
    algo3_orientation = final_df.get(
        "algo3_orientation_finale",
        pd.Series([""] * len(final_df), index=final_df.index),
    ).fillna("").astype(str)

    summary["resultat_global"] = [
        build_global_result(row_status, row_orientation)
        for row_status, row_orientation in zip(status, algo3_orientation)
    ]
    summary["etape_finale"] = final_df.get("pipeline_stage", "")
    summary["statut_pipeline"] = status
    summary["raison_arret_ou_passage"] = final_df.get("pipeline_stop_reason", "")
    summary["prochaine_action"] = final_df.get("pipeline_next_action", "")

    summary["algo1_execute"] = final_df.get("pipeline_algo1_executed", "")
    summary["algo1_decision_finale"] = summary_series(final_df, ["algo1_decision_finale", "decision_finale"])
    summary["algo1_passe_vers_algo2"] = [
        "oui" if clean_decision(value) in ALGO1_PASS_DECISIONS else ("non" if clean_text(value) else "")
        for value in summary["algo1_decision_finale"]
    ]
    summary["algo1_raison"] = summary_series(final_df, ["algo1_reason", "reason"])
    summary["algo1_regle_finale"] = summary_series(final_df, ["algo1_regle_finale", "regle_finale"])
    summary["algo1_infos_manquantes"] = summary_series(final_df, ["algo1_missing_information", "missing_information"])

    summary["algo2_execute"] = final_df.get("pipeline_algo2_executed", "")
    summary["algo2_decision_finale"] = summary_series(final_df, ["algo2_decision_finale"])
    summary["algo2_passe_vers_algo3"] = [
        "oui" if clean_decision(value) in ALGO2_PASS_DECISIONS else ("non" if clean_text(value) else "")
        for value in summary["algo2_decision_finale"]
    ]
    summary["algo2_raison"] = summary_series(final_df, ["algo2_reason", "algo2_reason.1"])
    summary["algo2_regle_finale"] = summary_series(final_df, ["algo2_regle_finale"])
    summary["algo2_b2b"] = summary_series(final_df, ["b2b_strength", "algo2_b2b_strength"])
    summary["algo2_secteur_confirme"] = summary_series(final_df, ["confirmed_sector", "algo2_confirmed_sector"])
    summary["algo2_matrice_enjeu"] = summary_series(final_df, ["main_issue_matrix", "algo2_main_issue_matrix"])
    summary["algo2_infos_manquantes"] = summary_series(final_df, ["algo2_missing_information"])
    summary["edda_status_final"] = summary_series(final_df, ["edda_status_final", "algo2_edda_status_final"])

    summary["algo3_execute"] = final_df.get("pipeline_algo3_executed", "")
    summary["algo3_orientation_finale"] = summary_series(final_df, ["algo3_orientation_finale"])
    summary["algo3_raison"] = summary_series(final_df, ["algo3_orientation_reason", "orientation_reason"])
    summary["algo3_regle_finale"] = summary_series(final_df, ["algo3_regle_finale"])
    summary["algo3_infos_manquantes"] = summary_series(final_df, ["algo3_missing_information"])
    summary["algo3_solution_type"] = summary_series(final_df, ["solution_type", "algo3_solution_type"])
    summary["algo3_maturite"] = summary_series(final_df, ["maturity_stage", "algo3_maturity_stage"])
    summary["algo3_trl"] = summary_series(final_df, ["trl_estimate", "algo3_trl_estimate"])
    summary["algo3_mrl"] = summary_series(final_df, ["mrl_estimate", "algo3_mrl_estimate"])
    summary["algo3_employes"] = summary_series(final_df, ["employees_estimate", "nombre_employes"])
    summary["algo3_seed_score"] = summary_series(final_df, ["algo3_seed_score"])
    summary["algo3_catalyst_score"] = summary_series(final_df, ["algo3_catalyst_score"])
    summary["algo3_materials_score"] = summary_series(final_df, ["algo3_materials_score"])
    summary["commentaire_verification"] = summary_series(final_df, ["commentaire_verification"])

    return summary.reset_index(drop=True)


def save_results(results: list[dict], output_file: Path) -> None:
    final_df = pd.DataFrame(results)
    if final_df.empty:
        final_df = pd.DataFrame(columns=PIPELINE_COLUMNS)

    final_df = Algo3.sanitize_excel_dataframe(final_df)

    status = final_df.get("pipeline_status", pd.Series([""] * len(final_df))).fillna("").astype(str)
    algo1_decision = final_df.get("decision_finale", pd.Series([""] * len(final_df))).fillna("").astype(str)
    algo2_decision = final_df.get("algo2_decision_finale", pd.Series([""] * len(final_df))).fillna("").astype(str)
    algo3_orientation = final_df.get("algo3_orientation_finale", pd.Series([""] * len(final_df))).fillna("").astype(str)
    team_summary_df = build_team_summary(final_df)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        team_summary_df.to_excel(writer, sheet_name="Synthese equipe", index=False)
        final_df.to_excel(writer, sheet_name="Toutes les startups", index=False)
        sheet_safe(final_df, status.eq("ARRETEE_ALGO1_A_VERIFIER")).to_excel(
            writer,
            sheet_name="Stop Algo1 A verifier",
            index=False,
        )
        sheet_safe(final_df, status.eq("ARRETEE_ALGO1_A_ECARTER")).to_excel(
            writer,
            sheet_name="Stop Algo1 A ecarter",
            index=False,
        )
        sheet_safe(final_df, status.eq("ARRETEE_ALGO2_A_VERIFIER")).to_excel(
            writer,
            sheet_name="Stop Algo2 A verifier",
            index=False,
        )
        sheet_safe(final_df, status.eq("ARRETEE_ALGO2_NON_PREQUALIFIEE")).to_excel(
            writer,
            sheet_name="Stop Algo2 Non prequal",
            index=False,
        )
        sheet_safe(final_df, algo1_decision.eq("A_GARDER")).to_excel(
            writer,
            sheet_name="Algo1 A garder",
            index=False,
        )
        sheet_safe(final_df, algo2_decision.eq("PREQUALIFIEE")).to_excel(
            writer,
            sheet_name="Algo2 Prequalifiees",
            index=False,
        )
        for orientation, sheet_name in [
            ("Seed", "Algo3 Seed"),
            ("Catalyst", "Algo3 Catalyst"),
            ("Materiaux", "Algo3 Materiaux"),
            ("Matériaux", "Algo3 Materiaux accent"),
            ("Autre", "Algo3 Autre"),
            ("A_VERIFIER", "Algo3 A verifier"),
        ]:
            if sheet_name == "Algo3 Materiaux accent":
                continue
            if orientation == "Materiaux":
                mask = algo3_orientation.isin(["Materiaux", "Matériaux"])
            else:
                mask = algo3_orientation.eq(orientation)
            sheet_safe(final_df, mask).to_excel(writer, sheet_name=sheet_name, index=False)

        if "sauvegarde_temp" not in output_file.stem.lower():
            excel_styling.style_workbook(
                writer,
                include_guide=True,
                title="Résultats du pipeline Algo 1 → Algo 2 → Algo 3",
            )


def generate_dashboard_if_enabled(output_file: Path) -> None:
    if not GENERATE_DASHBOARD:
        return
    try:
        dashboard_file = Generer_dashboard.generate_dashboard(output_file)
        log(f"Dashboard genere : {dashboard_file.name}")
    except Exception as exc:
        log(f"Dashboard non genere : {exc}")


def process_excel_file(input_file: Path) -> None:
    output_file = make_output_path(input_file)
    temp_file = make_temp_path(input_file)
    file_started = time.time()

    log(f"Chargement : {input_file.name}")
    df = pd.read_excel(input_file)
    if LIMIT_ROWS > 0:
        original_count = len(df)
        df = df.head(LIMIT_ROWS).copy()
        log(f"Mode test : {len(df)}/{original_count} lignes.")

    detection_df = prepare_detection_frame(df)
    columns_algo1 = Algo1.detect_columns(detection_df.copy())
    columns_algo2 = Algo2.detect_columns(detection_df.copy())
    columns_algo3 = Algo3.detect_columns(detection_df.copy())

    results = load_resume_results(temp_file)
    start_index = min(len(results), len(df))
    if start_index >= len(df):
        log("Toutes les lignes sont deja presentes dans la sauvegarde temporaire.")
        emit_progress(len(df), len(df), file_started, "Déjà terminé", start_index)
        save_results(results[:len(df)], output_file)
        generate_dashboard_if_enabled(output_file)
        log(f"Fichier genere : {output_file.name}")
        return

    website_cache = {}
    website_cache_lock = threading.Lock()
    algo2_research_cache = {}
    algo2_research_cache_lock = threading.Lock()
    algo3_research_cache = {}
    algo3_research_cache_lock = threading.Lock()

    log(f"Lignes a traiter : {len(df) - start_index}/{len(df)}")
    log(f"Profil dealflow : {dealflow_profiles.profile_label()}")
    log(f"Passage Algo1 -> Algo2 : {sorted(ALGO1_PASS_DECISIONS)}")
    log(f"Passage Algo2 -> Algo3 : {sorted(ALGO2_PASS_DECISIONS)}")
    emit_progress(start_index, len(df), file_started, "Préparation", start_index)

    try:
        iterator = df.iloc[start_index:].iterrows()
        for row_index, row in tqdm(
            iterator,
            total=len(df),
            initial=start_index,
            desc=f"Pipeline Algo123 {input_file.name}",
            disable=STREAMLIT_PROGRESS,
        ):
            result = process_pipeline_row(
                row,
                row_index,
                columns_algo1,
                columns_algo2,
                columns_algo3,
                website_cache,
                website_cache_lock,
                algo2_research_cache,
                algo2_research_cache_lock,
                algo3_research_cache,
                algo3_research_cache_lock,
            )
            results.append(result)
            emit_progress(len(results), len(df), file_started, "Analyse", start_index)

            if SAVE_EVERY > 0 and len(results) % SAVE_EVERY == 0:
                save_results(results, temp_file)
                log(f"Sauvegarde temporaire : {temp_file.name}")
    except KeyboardInterrupt:
        if results:
            save_results(results, temp_file)
            log(f"Interruption : sauvegarde temporaire ecrite avec {len(results)}/{len(df)} ligne(s).")
        raise
    except Exception:
        if results:
            save_results(results, temp_file)
            log(f"Erreur : sauvegarde temporaire ecrite avec {len(results)}/{len(df)} ligne(s).")
        raise

    save_results(results, output_file)
    emit_progress(len(results), len(df), file_started, "Génération du dashboard", start_index)
    generate_dashboard_if_enabled(output_file)
    emit_progress(len(results), len(df), file_started, "Terminé", start_index)
    log(f"Fichier genere : {output_file.name}")


def main() -> None:
    base_dir = Path.cwd()
    input_files = discover_input_files(base_dir, sys.argv[1:])

    if not input_files:
        log("Aucun fichier Excel source trouve.")
        log('Exemple : python -B Pipeline_Algo123.py "mon_fichier.xlsx"')
        return

    log(f"Fichiers a traiter : {len(input_files)}")
    for input_file in input_files:
        try:
            process_excel_file(input_file)
        except Exception as exc:
            log(f"Erreur sur {input_file.name} : {exc}")

    log("Traitement termine.")


if __name__ == "__main__":
    main()
