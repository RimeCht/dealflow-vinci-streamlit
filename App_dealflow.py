import os
import json
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from dotenv import load_dotenv

import dealflow_profiles
import Generer_dashboard


BASE_DIR = Path(__file__).resolve().parent
RUNS_DIR = BASE_DIR / "_runs_equipe"
PIPELINE_SCRIPT = BASE_DIR / "Pipeline_Algo123.py"


load_dotenv(BASE_DIR / ".env")


def load_streamlit_secrets_to_env() -> None:
    try:
        secrets = st.secrets
        for key, value in secrets.items():
            if isinstance(value, (dict, list, tuple)):
                continue
            os.environ[str(key)] = str(value)
    except Exception:
        return


load_streamlit_secrets_to_env()


PROGRESS_PREFIX = "__PIPELINE_PROGRESS__ "


def safe_filename(name: str) -> str:
    name = Path(name).name
    name = re.sub(r"[^\w .()\-]+", "_", name, flags=re.UNICODE)
    name = re.sub(r"\s+", " ", name).strip()
    return name or "fichier.xlsx"


def run_dir_name(uploaded_name: str) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = Path(safe_filename(uploaded_name)).stem[:70]
    return f"{timestamp}_{stem}"


def save_uploaded_file(uploaded_file, target_dir: Path) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    input_path = target_dir / safe_filename(uploaded_file.name)
    input_path.write_bytes(uploaded_file.getbuffer())
    return input_path


def pipeline_output_paths(input_path: Path) -> tuple[Path, Path, Path]:
    excel_path = input_path.with_name(f"{input_path.stem}_pipeline_algo_1_2_3.xlsx")
    dashboard_path = excel_path.with_name(f"{excel_path.stem}_dashboard.html")
    temp_path = input_path.with_name(f"{input_path.stem}_pipeline_algo_1_2_3_sauvegarde_temp.xlsx")
    return excel_path, dashboard_path, temp_path


def download_button(label: str, path: Path, mime: str, key: str) -> None:
    if not path.exists():
        st.caption(f"{label} non disponible.")
        return
    st.download_button(
        label=label,
        data=path.read_bytes(),
        file_name=path.name,
        mime=mime,
        key=key,
        use_container_width=True,
    )


def format_duration(seconds: float | int | None) -> str:
    if seconds is None:
        return "calcul en cours"
    seconds = max(0, int(seconds))
    hours, remainder = divmod(seconds, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours}h {minutes:02d}min"
    if minutes:
        return f"{minutes}min {secs:02d}s"
    return f"{secs}s"


def parse_progress_line(line: str) -> dict | None:
    if not line.startswith(PROGRESS_PREFIX):
        return None
    try:
        return json.loads(line[len(PROGRESS_PREFIX):])
    except json.JSONDecodeError:
        return None


def render_css() -> None:
    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 2rem;
            padding-bottom: 3rem;
            max-width: 1440px;
        }
        .main-title {
            font-size: 2.1rem;
            font-weight: 780;
            letter-spacing: 0;
            margin-bottom: 0.15rem;
        }
        .subtitle {
            color: #65736c;
            font-size: 0.95rem;
            margin-bottom: 1.2rem;
        }
        .metric-card {
            border: 1px solid #dbe2de;
            background: #ffffff;
            border-radius: 8px;
            padding: 1rem;
            min-height: 118px;
            box-shadow: 0 10px 26px rgba(23, 33, 29, 0.07);
        }
        .metric-label {
            color: #65736c;
            font-size: 0.78rem;
            text-transform: uppercase;
            font-weight: 700;
        }
        .metric-value {
            font-size: 1.85rem;
            font-weight: 820;
            margin-top: 0.35rem;
            color: #17211d;
        }
        .metric-hint {
            color: #65736c;
            font-size: 0.78rem;
            margin-top: 0.15rem;
        }
        .status-box {
            border: 1px solid #dbe2de;
            border-radius: 8px;
            padding: 0.85rem 1rem;
            background: #f9fbfa;
        }
        .soft-panel {
            border: 1px solid #dbe2de;
            background: #ffffff;
            border-radius: 8px;
            padding: 1rem;
        }
        .small-muted {
            color: #65736c;
            font-size: 0.84rem;
        }
        div[data-testid="stMetricValue"] {
            font-weight: 780;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header() -> None:
    st.markdown('<div class="main-title">Dealflow startups - Leonard / VINCI</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="subtitle">Interface interne pour lancer Algo1 -> Algo2 -> Algo3, récupérer l’Excel et partager un dashboard visuel.</div>',
        unsafe_allow_html=True,
    )


def render_config_status() -> str:
    deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT", "non configure")
    endpoint = os.getenv("AZURE_OPENAI_ENDPOINT", "")
    api_version = os.getenv("AZURE_OPENAI_API_VERSION", "non configure")
    api_key_present = bool(os.getenv("AZURE_OPENAI_API_KEY", "").strip())

    default_profile = os.getenv("DEALFLOW_PROFILE", dealflow_profiles.EUROPE).strip().upper()
    if default_profile not in dealflow_profiles.AVAILABLE_PROFILES:
        default_profile = dealflow_profiles.EUROPE

    profile_labels = {
        dealflow_profiles.EUROPE: "Europe - criteres actuels",
        dealflow_profiles.LATAM: "LATAM - anglais, ARR 500k",
    }
    st.sidebar.markdown("### Profil")
    selected_profile = st.sidebar.selectbox(
        "Zone / criteres",
        options=dealflow_profiles.AVAILABLE_PROFILES,
        index=dealflow_profiles.AVAILABLE_PROFILES.index(default_profile),
        format_func=lambda value: profile_labels.get(value, value),
    )
    os.environ["DEALFLOW_PROFILE"] = selected_profile
    if selected_profile == dealflow_profiles.LATAM:
        st.sidebar.caption("LATAM exclut surtout construction of buildings et real estate. Infra, civil engineering, roads/bridges/tunnels/rail et materials restent inclus.")
    else:
        st.sidebar.caption("Europe conserve les criteres deja valides.")
    st.sidebar.divider()

    st.sidebar.markdown("### Configuration")
    st.sidebar.write(f"Modèle Azure : `{deployment}`")
    st.sidebar.write(f"API version : `{api_version}`")
    st.sidebar.write("Endpoint : " + ("configuré" if endpoint else "manquant"))
    st.sidebar.write("Clé API : " + ("configurée" if api_key_present else "manquante"))
    st.sidebar.divider()
    st.sidebar.write(f"Dossier runs : `{RUNS_DIR.name}`")
    st.sidebar.caption("La clé API n'est jamais affichée dans l'interface.")
    return selected_profile


def run_pipeline(input_path: Path, profile: str) -> tuple[int, list[str]]:
    command = [sys.executable, "-B", str(PIPELINE_SCRIPT), str(input_path)]
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUNBUFFERED"] = "1"
    env["PIPELINE_GENERATE_DASHBOARD"] = env.get("PIPELINE_GENERATE_DASHBOARD", "1")
    env["DEALFLOW_PROFILE"] = profile
    env["PIPELINE_STREAMLIT_PROGRESS"] = "1"

    process = subprocess.Popen(
        command,
        cwd=str(BASE_DIR),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
    )

    logs: list[str] = []
    progress_bar = st.progress(0, text="Préparation de l'analyse...")
    progress_detail = st.empty()
    status_box = st.empty()
    with st.expander("Logs techniques", expanded=False):
        log_box = st.empty()
    started = time.time()
    last_progress: dict | None = None

    assert process.stdout is not None
    for line in process.stdout:
        line = line.rstrip()
        progress = parse_progress_line(line)
        if progress:
            last_progress = progress
            done = int(progress.get("done") or 0)
            total = int(progress.get("total") or 0)
            percent = float(progress.get("percent") or 0)
            ratio = min(max(percent / 100, 0), 1)
            stage = str(progress.get("stage") or "Analyse")
            elapsed = format_duration(progress.get("elapsed_seconds"))
            eta = format_duration(progress.get("eta_seconds"))
            progress_bar.progress(ratio, text=f"{stage} - {percent:.1f}%")
            progress_detail.markdown(
                f"**{done}/{total} startups traitées** · temps écoulé : `{elapsed}` · temps restant estimé : `{eta}`"
            )
            status_box.info("Analyse en cours")
            continue

        if line:
            logs.append(line)
        elapsed = format_duration(time.time() - started)
        if not last_progress:
            progress_detail.markdown(
                f"Analyse en cours · temps écoulé : `{elapsed}` · estimation disponible après les premières lignes."
            )
        status_box.info("Analyse en cours")
        log_box.code("\n".join(logs[-60:]) or "Démarrage...", language="text")

    return_code = process.wait()
    elapsed = format_duration(time.time() - started)
    if return_code == 0:
        progress_bar.progress(1.0, text="Analyse terminée - 100%")
        status_box.success(f"Analyse terminée en {elapsed}")
    else:
        status_box.error(f"Analyse arrêtée avec le code {return_code}")
    log_box.code("\n".join(logs[-120:]) or "Aucun log.", language="text")
    return return_code, logs


def read_dashboard_summary(excel_path: Path) -> dict | None:
    try:
        df, sheet = Generer_dashboard.read_pipeline_dataframe(excel_path)
        records = Generer_dashboard.build_records(df)
        return Generer_dashboard.build_summary(records, excel_path, sheet)
    except Exception:
        return None


def render_kpi_cards(summary: dict) -> None:
    kpis = summary.get("kpis", [])
    if not kpis:
        return
    columns = st.columns(min(len(kpis), 7))
    for index, kpi in enumerate(kpis):
        with columns[index % len(columns)]:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-label">{kpi.get('label', '')}</div>
                    <div class="metric-value">{kpi.get('value', '')}</div>
                    <div class="metric-hint">{kpi.get('hint', '')}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def render_dashboard_preview(dashboard_path: Path) -> None:
    if not dashboard_path.exists():
        return
    with st.expander("Aperçu du dashboard HTML", expanded=True):
        components.html(dashboard_path.read_text(encoding="utf-8"), height=900, scrolling=True)


def render_outputs(excel_path: Path, dashboard_path: Path, temp_path: Path) -> None:
    st.markdown("### Résultats")
    summary = read_dashboard_summary(excel_path) if excel_path.exists() else None
    if summary:
        render_kpi_cards(summary)

    col1, col2, col3 = st.columns(3)
    with col1:
        download_button(
            "Télécharger Excel résultat",
            excel_path,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            f"xlsx-{excel_path}",
        )
    with col2:
        download_button(
            "Télécharger dashboard HTML",
            dashboard_path,
            "text/html",
            f"html-{dashboard_path}",
        )
    with col3:
        download_button(
            "Télécharger sauvegarde temporaire",
            temp_path,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            f"temp-{temp_path}",
        )

    render_dashboard_preview(dashboard_path)


def page_run_pipeline(profile: str) -> None:
    st.markdown("### Nouvelle analyse")
    st.markdown(
        '<div class="status-box">Dépose un Excel source. L’app lance le pipeline complet, puis génère l’Excel final et le dashboard HTML.</div>',
        unsafe_allow_html=True,
    )
    st.caption(f"Profil actif : {profile}")
    uploaded_file = st.file_uploader(
        "Fichier Excel source",
        type=["xlsx", "xlsm", "xls"],
        accept_multiple_files=False,
    )

    col1, col2 = st.columns([1, 2])
    with col1:
        start = st.button("Lancer l'analyse complète", type="primary", use_container_width=True)
    with col2:
        st.caption("Pour les gros fichiers, laisse la fenêtre ouverte. Le pipeline garde aussi une sauvegarde temporaire.")

    if not uploaded_file:
        return

    run_dir = RUNS_DIR / run_dir_name(uploaded_file.name)
    input_path = save_uploaded_file(uploaded_file, run_dir)
    excel_path, dashboard_path, temp_path = pipeline_output_paths(input_path)
    st.session_state["last_run_paths"] = [str(excel_path), str(dashboard_path), str(temp_path)]

    st.write(f"Fichier enregistré : `{input_path.name}`")
    if start:
        return_code, _ = run_pipeline(input_path, profile)
        if return_code == 0:
            render_outputs(excel_path, dashboard_path, temp_path)
        else:
            st.warning("Le pipeline n'a pas terminé correctement. Regarde les logs ci-dessus et la sauvegarde temporaire si elle existe.")
            render_outputs(excel_path, dashboard_path, temp_path)


def page_dashboard_only() -> None:
    st.markdown("### Générer seulement le dashboard")
    st.caption("À utiliser si tu as déjà un fichier `*_pipeline_algo_1_2_3.xlsx`.")
    uploaded_file = st.file_uploader(
        "Fichier Excel pipeline",
        type=["xlsx", "xlsm", "xls"],
        accept_multiple_files=False,
        key="dashboard-only-upload",
    )
    if not uploaded_file:
        return

    run_dir = RUNS_DIR / run_dir_name(uploaded_file.name)
    excel_path = save_uploaded_file(uploaded_file, run_dir)
    dashboard_path = excel_path.with_name(f"{excel_path.stem}_dashboard.html")

    if st.button("Générer dashboard", type="primary", use_container_width=True):
        try:
            dashboard_path = Generer_dashboard.generate_dashboard(excel_path)
            st.success(f"Dashboard généré : {dashboard_path.name}")
            render_outputs(excel_path, dashboard_path, excel_path.with_suffix(".tmp.xlsx"))
        except Exception as exc:
            st.error(f"Impossible de générer le dashboard : {exc}")


def iter_run_outputs() -> list[tuple[Path, Path | None]]:
    if not RUNS_DIR.exists():
        return []
    outputs = []
    for excel_path in sorted(RUNS_DIR.glob("**/*_pipeline_algo_1_2_3.xlsx"), reverse=True):
        if "sauvegarde_temp" in excel_path.stem:
            continue
        dashboard_path = excel_path.with_name(f"{excel_path.stem}_dashboard.html")
        outputs.append((excel_path, dashboard_path if dashboard_path.exists() else None))
    return outputs[:30]


def page_history() -> None:
    st.markdown("### Derniers runs")
    outputs = iter_run_outputs()
    if not outputs:
        st.info("Aucun run disponible pour l'instant.")
        return

    for index, (excel_path, dashboard_path) in enumerate(outputs):
        with st.container(border=True):
            stat = excel_path.stat()
            st.write(f"**{excel_path.name}**")
            st.caption(f"{excel_path.parent.name} - modifié le {datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')}")
            col1, col2, col3 = st.columns(3)
            with col1:
                download_button(
                    "Excel",
                    excel_path,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    f"hist-xlsx-{index}",
                )
            with col2:
                if dashboard_path:
                    download_button("Dashboard", dashboard_path, "text/html", f"hist-html-{index}")
                else:
                    st.caption("Dashboard absent.")
            with col3:
                if st.button("Prévisualiser", key=f"preview-{index}", use_container_width=True):
                    if dashboard_path:
                        render_dashboard_preview(dashboard_path)
                    else:
                        st.warning("Dashboard absent pour ce run.")


def main() -> None:
    st.set_page_config(
        page_title="Dealflow VINCI",
        page_icon=None,
        layout="wide",
        initial_sidebar_state="expanded",
    )
    render_css()
    render_header()
    profile = render_config_status()

    tab_run, tab_dashboard, tab_history = st.tabs([
        "Lancer une analyse",
        "Dashboard seul",
        "Historique",
    ])
    with tab_run:
        page_run_pipeline(profile)
    with tab_dashboard:
        page_dashboard_only()
    with tab_history:
        page_history()


if __name__ == "__main__":
    main()
