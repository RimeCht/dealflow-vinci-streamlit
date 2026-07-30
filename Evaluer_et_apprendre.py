import sys
from pathlib import Path

import pandas as pd

import learning_memory


OUTPUT_SUFFIX = "_rapport_evaluation_apprentissage.xlsx"

ALGO_CONFIG = {
    "algo1": {
        "manual_columns": learning_memory.MANUAL_ALGO1_COLUMNS,
        "prediction_columns": ["decision_finale", "algo1_decision_finale", "algo1_decision"],
        "allowed": learning_memory.ALGO1_LABELS,
        "positive_labels": {"A_GARDER"},
    },
    "algo2": {
        "manual_columns": learning_memory.MANUAL_ALGO2_COLUMNS,
        "prediction_columns": ["algo2_decision_finale", "algo2_decision"],
        "allowed": learning_memory.ALGO2_LABELS,
        "positive_labels": {"PREQUALIFIEE"},
    },
    "algo3": {
        "manual_columns": learning_memory.MANUAL_ALGO3_COLUMNS,
        "prediction_columns": ["algo3_orientation_finale", "algo3_orientation"],
        "allowed": learning_memory.ALGO3_LABELS,
        "positive_labels": {"Seed", "Catalyst", "Matériaux", "Materiaux", "Autre"},
    },
}

PRIORITY_COLUMNS = [
    "pipeline_index",
    "Name",
    "startup",
    "company",
    "Website",
    "website",
    "pipeline_status",
    "pipeline_stop_reason",
    "decision_finale",
    "algo2_decision_finale",
    "algo3_orientation_finale",
    "confidence",
    "algo2_confidence",
    "algo3_confidence",
    "algo1_api_called",
    "algo2_api_called",
    "algo1_learning_match_score",
    "algo2_learning_match_score",
    "manual_algo1_decision",
    "manual_algo2_decision",
    "manual_algo3_orientation",
    "manual_comment",
]


def log(message: str) -> None:
    print(f"[Evaluation] {message}")


def clean_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def normalize_column(column: str) -> str:
    return str(column).strip().lower().replace("_", " ")


def find_first_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    normalized_columns = {normalize_column(column): column for column in df.columns}
    for candidate in candidates:
        column = normalized_columns.get(normalize_column(candidate))
        if column is not None:
            return column
    return None


def normalize_label(value, allowed: set[str]) -> str:
    return learning_memory.normalize_label(value, allowed)


def discover_files(args: list[str]) -> list[Path]:
    if args:
        return [Path(arg).expanduser().resolve() for arg in args]

    return [
        path
        for path in sorted(Path.cwd().iterdir())
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}
        and "_pipeline_algo_1_2_3" in path.name.lower()
        and "sauvegarde_temp" not in path.name.lower()
        and not path.name.startswith("~$")
    ]


def read_main_sheet(path: Path) -> pd.DataFrame:
    workbook = pd.ExcelFile(path)
    sheet_name = "Toutes les startups" if "Toutes les startups" in workbook.sheet_names else workbook.sheet_names[0]
    return pd.read_excel(path, sheet_name=sheet_name)


def preferred_columns(df: pd.DataFrame) -> list[str]:
    selected = [column for column in PRIORITY_COLUMNS if column in df.columns]
    remaining = [column for column in df.columns if column not in selected]
    return selected + remaining


def evaluate_stage(df: pd.DataFrame, stage: str) -> tuple[pd.DataFrame, dict]:
    config = ALGO_CONFIG[stage]
    manual_column = find_first_column(df, config["manual_columns"])
    prediction_column = find_first_column(df, config["prediction_columns"])

    if not manual_column or not prediction_column:
        empty = df.iloc[0:0].copy()
        return empty, {
            "stage": stage,
            "manual_column": manual_column or "",
            "prediction_column": prediction_column or "",
            "corrected_rows": 0,
            "correct": 0,
            "incorrect": 0,
            "accuracy": "",
            "false_positive": 0,
            "false_negative": 0,
        }

    rows = []
    for index, row in df.iterrows():
        expected = normalize_label(row.get(manual_column, ""), config["allowed"])
        if not expected:
            continue
        predicted = normalize_label(row.get(prediction_column, ""), config["allowed"])
        is_correct = expected == predicted
        false_positive = predicted in config["positive_labels"] and expected not in config["positive_labels"]
        false_negative = expected in config["positive_labels"] and predicted not in config["positive_labels"]
        record = row.to_dict()
        record.update({
            "evaluation_stage": stage,
            "evaluation_row": index + 2,
            "manual_column_used": manual_column,
            "prediction_column_used": prediction_column,
            "expected_label": expected,
            "predicted_label": predicted,
            "is_correct": "oui" if is_correct else "non",
            "error_type": (
                "false_positive" if false_positive
                else "false_negative" if false_negative
                else "" if is_correct
                else "wrong_label"
            ),
        })
        rows.append(record)

    result_df = pd.DataFrame(rows)
    corrected_rows = len(result_df)
    correct = int(result_df["is_correct"].eq("oui").sum()) if corrected_rows else 0
    incorrect = corrected_rows - correct
    false_positive_count = int(result_df["error_type"].eq("false_positive").sum()) if corrected_rows else 0
    false_negative_count = int(result_df["error_type"].eq("false_negative").sum()) if corrected_rows else 0

    summary = {
        "stage": stage,
        "manual_column": manual_column,
        "prediction_column": prediction_column,
        "corrected_rows": corrected_rows,
        "correct": correct,
        "incorrect": incorrect,
        "accuracy": round(correct / corrected_rows, 4) if corrected_rows else "",
        "false_positive": false_positive_count,
        "false_negative": false_negative_count,
    }
    return result_df, summary


def build_priority_review(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for index, row in df.iterrows():
        reasons = []
        score = 0

        status = clean_text(row.get("pipeline_status", ""))
        if "A_VERIFIER" in status:
            reasons.append("pipeline a verifier")
            score += 3
        if clean_text(row.get("decision_finale", "")) == "A_VERIFIER":
            reasons.append("algo1 a verifier")
            score += 2
        if clean_text(row.get("algo2_decision_finale", "")) == "A_VERIFIER":
            reasons.append("algo2 a verifier")
            score += 3
        if clean_text(row.get("algo3_orientation_finale", "")) == "A_VERIFIER":
            reasons.append("algo3 a verifier")
            score += 2

        for column in ["confidence", "algo2_confidence", "algo3_confidence"]:
            try:
                confidence = float(row.get(column, 1) or 1)
            except (TypeError, ValueError):
                confidence = 1
            if 0 < confidence < 0.75:
                reasons.append(f"{column} faible")
                score += 2

        if clean_text(row.get("algo1_api_called", "")).lower() == "oui":
            reasons.append("algo1 a appele API")
            score += 1
        if clean_text(row.get("algo2_api_called", "")).lower() == "oui":
            reasons.append("algo2 a appele API")
            score += 1

        for column in ["algo1_learning_match_score", "algo2_learning_match_score"]:
            try:
                match_score = float(row.get(column, 0) or 0)
            except (TypeError, ValueError):
                match_score = 0
            if 0.18 <= match_score < learning_memory.MATCH_THRESHOLD:
                reasons.append(f"{column} proche du seuil")
                score += 2

        if score <= 0:
            continue

        record = row.to_dict()
        record.update({
            "review_priority_score": score,
            "review_reasons": " ; ".join(reasons),
            "evaluation_row": index + 2,
        })
        rows.append(record)

    if not rows:
        return df.iloc[0:0].copy()

    priority_df = pd.DataFrame(rows)
    priority_df = priority_df.sort_values(
        by=["review_priority_score"],
        ascending=False,
    )
    return priority_df


def label_distribution(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    columns = [
        "decision_finale",
        "algo2_decision_finale",
        "algo3_orientation_finale",
        "pipeline_status",
        "algo1_api_called",
        "algo2_api_called",
    ]
    for column in columns:
        if column not in df.columns:
            continue
        counts = df[column].fillna("").astype(str).replace("", "(vide)").value_counts(dropna=False)
        for label, count in counts.items():
            rows.append({"column": column, "value": label, "count": int(count)})
    return pd.DataFrame(rows)


def sanitize_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    try:
        import Algo3

        return Algo3.sanitize_excel_dataframe(df)
    except Exception:
        return df


def save_report(
    output_path: Path,
    summary_df: pd.DataFrame,
    evaluated: dict[str, pd.DataFrame],
    priority_df: pd.DataFrame,
    distribution_df: pd.DataFrame,
    learned_stats: dict,
) -> None:
    learned_df = pd.DataFrame([learned_stats])

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        sanitize_for_excel(summary_df).to_excel(writer, sheet_name="Synthese", index=False)
        sanitize_for_excel(learned_df).to_excel(writer, sheet_name="Apprentissage", index=False)
        sanitize_for_excel(distribution_df).to_excel(writer, sheet_name="Distributions", index=False)
        sanitize_for_excel(priority_df[preferred_columns(priority_df)] if not priority_df.empty else priority_df).to_excel(
            writer,
            sheet_name="A corriger priorite",
            index=False,
        )

        for stage, stage_df in evaluated.items():
            errors_df = stage_df[stage_df["is_correct"].eq("non")].copy() if not stage_df.empty else stage_df
            sheet_name = f"Erreurs {stage.capitalize()}"
            if not errors_df.empty:
                errors_df = errors_df[preferred_columns(errors_df)]
            sanitize_for_excel(errors_df).to_excel(writer, sheet_name=sheet_name[:31], index=False)


def process_file(path: Path) -> None:
    log(f"Lecture : {path.name}")
    df = read_main_sheet(path)

    evaluated = {}
    summaries = []
    for stage in ["algo1", "algo2", "algo3"]:
        stage_df, summary = evaluate_stage(df, stage)
        evaluated[stage] = stage_df
        summaries.append(summary)

    summary_df = pd.DataFrame(summaries)
    priority_df = build_priority_review(df)
    distribution_df = label_distribution(df)

    memory = learning_memory.load_memory()
    memory, learned_stats = learning_memory.learn_from_dataframe(df, f"{path.name}::Toutes les startups", memory)
    learning_memory.save_memory(memory)

    output_path = path.with_name(f"{path.stem}{OUTPUT_SUFFIX}")
    save_report(output_path, summary_df, evaluated, priority_df, distribution_df, learned_stats)

    log(f"Rapport genere : {output_path.name}")
    log(
        "Apprentissage ajoute : "
        f"algo1={learned_stats['algo1']}, "
        f"algo2={learned_stats['algo2']}, "
        f"algo3={learned_stats['algo3']}"
    )


def main() -> None:
    files = discover_files(sys.argv[1:])
    if not files:
        log("Aucun fichier pipeline trouve.")
        log('Exemple : python -B Evaluer_et_apprendre.py "mon_fichier_pipeline_corrige.xlsx"')
        return

    for path in files:
        try:
            process_file(path)
        except Exception as exc:
            log(f"Erreur sur {path.name} : {exc}")


if __name__ == "__main__":
    main()
