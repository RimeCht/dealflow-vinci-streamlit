import os
import json
import re
import signal
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
JOB_META_FILE = "run_status.json"
JOB_LOG_FILE = "pipeline.log"


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


def read_json_file(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def write_json_file(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def read_log_tail(path: Path, limit: int = 30000) -> str:
    if not path.exists():
        return ""
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except Exception:
        return ""
    return text[-limit:]


def last_progress_from_log(log_path: Path) -> dict | None:
    for line in reversed(read_log_tail(log_path).splitlines()):
        progress = parse_progress_line(line.strip())
        if progress:
            return progress
    return None


def pid_is_running(pid: int | str | None) -> bool:
    try:
        pid_int = int(pid or 0)
    except (TypeError, ValueError):
        return False
    if pid_int <= 0:
        return False

    if os.name == "nt":
        try:
            result = subprocess.run(
                ["tasklist", "/FI", f"PID eq {pid_int}"],
                capture_output=True,
                text=True,
                timeout=3,
            )
            return str(pid_int) in result.stdout
        except Exception:
            return False

    try:
        os.kill(pid_int, 0)
        return True
    except OSError:
        return False


def append_log_line(path: Path, message: str) -> None:
    try:
        with path.open("a", encoding="utf-8", errors="replace") as log_file:
            log_file.write(message.rstrip() + "\n")
    except Exception:
        return


def stop_background_job(meta: dict) -> bool:
    try:
        pid_int = int(meta.get("pid") or 0)
    except (TypeError, ValueError):
        return False
    if pid_int <= 0:
        return False

    stopped_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    updated_meta = dict(meta)
    updated_meta["stop_requested_at"] = stopped_at
    meta_path_raw = str(updated_meta.get("meta_path", "")).strip()
    log_path_raw = str(updated_meta.get("log_path", "")).strip()
    meta_path = Path(meta_path_raw) if meta_path_raw else None
    log_path = Path(log_path_raw) if log_path_raw else Path()
    if meta_path is not None:
        write_json_file(meta_path, updated_meta)
    append_log_line(log_path, f"[App] Stop requested at {stopped_at}.")

    try:
        if os.name == "nt":
            ctrl_break = getattr(signal, "CTRL_BREAK_EVENT", None)
            if ctrl_break is not None:
                os.kill(pid_int, ctrl_break)
                time.sleep(2)
            if pid_is_running(pid_int):
                subprocess.run(
                    ["taskkill", "/PID", str(pid_int), "/T", "/F"],
                    capture_output=True,
                    text=True,
                    timeout=10,
                )
        else:
            try:
                os.killpg(pid_int, signal.SIGINT)
            except ProcessLookupError:
                return True
            time.sleep(2)
            if pid_is_running(pid_int):
                os.killpg(pid_int, signal.SIGTERM)
        return True
    except Exception as exc:
        append_log_line(log_path, f"[App] Stop failed: {exc}")
        return False


def job_state(meta: dict) -> str:
    excel_raw = str(meta.get("excel_path", "")).strip()
    if excel_raw and Path(excel_raw).exists():
        return "done"
    if pid_is_running(meta.get("pid")):
        if meta.get("stop_requested_at"):
            return "stopping"
        return "running"
    log_tail = read_log_tail(Path(meta.get("log_path", "")))
    if "Erreur sur" in log_tail or "Traceback" in log_tail:
        return "error"
    return "stopped"


def start_background_pipeline(input_path: Path, profile: str) -> dict:
    excel_path, dashboard_path, temp_path = pipeline_output_paths(input_path)
    log_path = input_path.parent / JOB_LOG_FILE
    meta_path = input_path.parent / JOB_META_FILE

    command = [sys.executable, "-B", str(PIPELINE_SCRIPT), str(input_path)]
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUNBUFFERED"] = "1"
    env["PIPELINE_GENERATE_DASHBOARD"] = env.get("PIPELINE_GENERATE_DASHBOARD", "1")
    env["DEALFLOW_PROFILE"] = profile
    env["PIPELINE_STREAMLIT_PROGRESS"] = "1"

    log_file = log_path.open("w", encoding="utf-8", errors="replace")
    popen_kwargs = {
        "cwd": str(BASE_DIR),
        "stdout": log_file,
        "stderr": subprocess.STDOUT,
        "text": True,
        "encoding": "utf-8",
        "errors": "replace",
        "env": env,
    }
    if os.name == "nt":
        popen_kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP
    else:
        popen_kwargs["start_new_session"] = True

    process = subprocess.Popen(command, **popen_kwargs)
    log_file.close()

    meta = {
        "pid": process.pid,
        "profile": profile,
        "started_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "input_path": str(input_path),
        "excel_path": str(excel_path),
        "dashboard_path": str(dashboard_path),
        "temp_path": str(temp_path),
        "log_path": str(log_path),
        "meta_path": str(meta_path),
    }
    write_json_file(meta_path, meta)
    return meta


def iter_job_metas(limit: int = 20) -> list[dict]:
    if not RUNS_DIR.exists():
        return []

    metas = []
    for meta_path in sorted(RUNS_DIR.glob(f"**/{JOB_META_FILE}"), reverse=True):
        meta = read_json_file(meta_path)
        if not meta:
            continue
        meta.setdefault("meta_path", str(meta_path))
        metas.append(meta)
    return metas[:limit]


def job_display_name(meta: dict) -> str:
    input_path = Path(str(meta.get("input_path", "")))
    if input_path.name:
        return input_path.name
    return Path(str(meta.get("meta_path", ""))).parent.name or "run"


def render_job_monitor(meta: dict, key_prefix: str = "job") -> None:
    state = job_state(meta)
    excel_path = Path(meta.get("excel_path", ""))
    dashboard_path = Path(meta.get("dashboard_path", ""))
    temp_path = Path(meta.get("temp_path", ""))
    log_path = Path(meta.get("log_path", ""))
    progress = last_progress_from_log(log_path)

    if progress:
        done = int(progress.get("done") or 0)
        total = int(progress.get("total") or 0)
        percent = float(progress.get("percent") or 0)
        ratio = min(max(percent / 100, 0), 1)
        stage = str(progress.get("stage") or "Analyse")
        elapsed = format_duration(progress.get("elapsed_seconds"))
        eta = format_duration(progress.get("eta_seconds"))
    else:
        done = 0
        total = 0
        percent = 0.0
        ratio = 0.0
        stage = "Préparation"
        elapsed = "calcul en cours"
        eta = "calcul en cours"

    if state == "done":
        ratio = 1.0
        percent = 100.0
        st.progress(ratio, text="Analyse terminée - 100%")
        st.success("Analyse terminée. Les fichiers sont disponibles ci-dessous.")
        render_outputs(excel_path, dashboard_path, temp_path)
    elif state in {"running", "stopping"}:
        st.progress(ratio, text=f"{stage} - {percent:.1f}%")
        st.markdown(
            f"**{done}/{total} startups traitées** · temps écoulé : `{elapsed}` · temps restant estimé : `{eta}`"
        )
        if state == "stopping":
            st.warning("Arrêt demandé. Le job est en train de se fermer.")
        else:
            st.info("Analyse en cours en arrière-plan. Tu peux fermer cet onglet, puis revenir dans l'app pour reprendre le suivi.")
            if st.button("Arrêter l'analyse", key=f"{key_prefix}-stop", type="secondary", use_container_width=True):
                if stop_background_job(meta):
                    refreshed = read_json_file(Path(str(meta.get("meta_path", ""))))
                    st.session_state["active_job_meta"] = refreshed or meta
                    st.warning("Arrêt demandé. La dernière sauvegarde temporaire disponible sera conservée.")
                else:
                    st.error("Impossible d'arrêter le job automatiquement. Vérifie le serveur ou les logs.")
                time.sleep(1)
                st.rerun()
        time.sleep(3)
        st.rerun()
    elif state == "error":
        st.error("Le job semble arrêté avec une erreur. La sauvegarde temporaire est disponible si elle a été créée.")
        render_outputs(excel_path, dashboard_path, temp_path)
    else:
        st.warning("Le job n'est plus actif et le fichier final n'est pas encore disponible.")
        render_outputs(excel_path, dashboard_path, temp_path)

    with st.expander("Logs techniques", expanded=False):
        st.code(read_log_tail(log_path, 50000) or "Aucun log disponible.", language="text")


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

    active_meta = st.session_state.get("active_job_meta")
    if active_meta:
        render_job_monitor(active_meta, "active")
        if job_state(active_meta) != "running":
            if st.button("Lancer une nouvelle analyse", use_container_width=True):
                st.session_state.pop("active_job_meta", None)
                st.rerun()
        return

    running_jobs = [
        meta
        for meta in iter_job_metas(limit=10)
        if job_state(meta) in {"running", "stopping"}
    ]
    if running_jobs:
        st.markdown("### Analyses en cours")
        for index, meta in enumerate(running_jobs[:3]):
            with st.container(border=True):
                st.write(f"**{job_display_name(meta)}**")
                st.caption(f"Profil : {meta.get('profile', '')} · démarré le {meta.get('started_at', '')}")
                if st.button("Suivre ce run", key=f"follow-running-{index}", use_container_width=True):
                    st.session_state["active_job_meta"] = meta
                    st.rerun()

    uploaded_file = st.file_uploader(
        "Fichier Excel source",
        type=["xlsx", "xlsm", "xls"],
        accept_multiple_files=False,
    )

    col1, col2 = st.columns([1, 2])
    with col1:
        start = st.button("Lancer l'analyse en arrière-plan", type="primary", use_container_width=True)
    with col2:
        st.caption("Tu peux fermer l'onglet pendant l'analyse. Le serveur doit rester allumé et le pipeline garde une sauvegarde temporaire.")

    if not uploaded_file:
        return

    st.write(f"Fichier prêt : `{uploaded_file.name}`")
    if start:
        run_dir = RUNS_DIR / run_dir_name(uploaded_file.name)
        input_path = save_uploaded_file(uploaded_file, run_dir)
        meta = start_background_pipeline(input_path, profile)
        st.session_state["active_job_meta"] = meta
        st.session_state["last_run_paths"] = [
            meta["excel_path"],
            meta["dashboard_path"],
            meta["temp_path"],
        ]
        st.rerun()


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
