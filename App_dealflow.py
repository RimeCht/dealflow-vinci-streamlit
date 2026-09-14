import os
import base64
import html
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
ASSETS_DIR = BASE_DIR / "assets"
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


def asset_data_uri(
    filename: str,
    mime_type: str,
    replacements: dict[str, str] | None = None,
) -> str:
    path = ASSETS_DIR / filename
    try:
        data = path.read_bytes()
    except OSError:
        return ""
    if replacements:
        text = data.decode("utf-8")
        for source, target in replacements.items():
            text = text.replace(source, target)
        data = text.encode("utf-8")
    encoded = base64.b64encode(data).decode("ascii")
    return f"data:{mime_type};base64,{encoded}"


def inline_svg(
    filename: str,
    class_name: str,
    replacements: dict[str, str] | None = None,
) -> str:
    path = ASSETS_DIR / filename
    try:
        svg = path.read_text(encoding="utf-8")
    except OSError:
        return ""
    svg = svg.replace("viewbox=", "viewBox=")
    for source, target in (replacements or {}).items():
        svg = svg.replace(source, target)
    if 'class="logo-colored"' in svg:
        return svg.replace('class="logo-colored"', f'class="{class_name}"', 1)
    return svg.replace("<svg", f'<svg class="{class_name}"', 1)


def vinci_font_faces() -> str:
    fonts = [
        ("Vinci Sans", "Vinci-Sans-Regular.woff2", 400),
        ("Vinci Sans", "Vinci-Sans-Medium.woff2", 500),
        ("Vinci Sans", "Vinci-Sans-Bold.woff2", 700),
        ("Vinci Serif", "Vinci-Serif-Regular.woff2", 400),
    ]
    rules = []
    for family, filename, weight in fonts:
        uri = asset_data_uri(filename, "font/woff2")
        if not uri:
            continue
        rules.append(
            "@font-face {"
            f"font-family: '{family}';"
            f"src: url('{uri}') format('woff2');"
            f"font-weight: {weight};"
            "font-style: normal;"
            "font-display: swap;"
            "}"
        )
    return "\n".join(rules)


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


def can_resume_job(meta: dict) -> bool:
    input_path = Path(str(meta.get("input_path", "")))
    temp_path = Path(str(meta.get("temp_path", "")))
    return (
        job_state(meta) in {"error", "stopped"}
        and input_path.is_file()
        and temp_path.is_file()
    )


def resume_background_job(meta: dict) -> dict:
    input_path = Path(str(meta.get("input_path", "")))
    profile = str(meta.get("profile") or os.getenv("DEALFLOW_PROFILE", "EUROPE"))
    return start_background_pipeline(input_path, profile)


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

    if can_resume_job(meta):
        if st.button(
            "Reprendre depuis la sauvegarde",
            key=f"{key_prefix}-resume",
            type="primary",
            use_container_width=True,
        ):
            resumed_meta = resume_background_job(meta)
            st.session_state["active_job_meta"] = resumed_meta
            st.success("Reprise lancée depuis la dernière ligne sauvegardée.")
            time.sleep(1)
            st.rerun()

    with st.expander("Logs techniques", expanded=False):
        st.code(read_log_tail(log_path, 50000) or "Aucun log disponible.", language="text")


def render_css() -> None:
    orbit_uri = asset_data_uri(
        "constellation-orbit.svg",
        "image/svg+xml",
        {"#fff": "#4A7CC9"},
    )
    css = vinci_font_faces() + """
    :root {
        --leonard-navy: #004489;
        --leonard-blue: #00b4ff;
        --leonard-pink: #ff005a;
        --leonard-ink: #102b45;
        --leonard-muted: #617487;
        --leonard-canvas: #eef1f4;
        --leonard-line: #d5dbe2;
        --leonard-white: #ffffff;
        --leonard-soft-blue: #e9f7fd;
        --constellation-field: #62636f;
        --constellation-blue: #4a7cc9;
        --constellation-violet: #6b4c9a;
        --constellation-rose: #d64c7e;
    }

    html, body, .stApp,
    .stApp p, .stApp label, .stApp button,
    .stApp input, .stApp textarea,
    .stApp [role="tab"] {
        font-family: "Vinci Sans", Arial, sans-serif;
        letter-spacing: 0 !important;
    }
    .stApp,
    [data-testid="stAppViewContainer"] {
        background: var(--leonard-canvas);
        color: var(--leonard-ink);
    }
    [data-testid="stAppViewContainer"] {
        position: relative;
        isolation: isolate;
        overflow: hidden;
    }
    [data-testid="stAppViewContainer"]::before {
        content: "";
        position: fixed;
        right: -18vw;
        bottom: -38vw;
        width: min(980px, 82vw);
        aspect-ratio: 1;
        background: url("__CONSTELLATION_ORBIT__") center / contain no-repeat;
        opacity: 0.08;
        pointer-events: none;
        z-index: -1;
    }
    [data-testid="stHeader"] {
        background: transparent;
    }
    [data-testid="stToolbar"] {
        right: 0.75rem;
    }
    .block-container {
        max-width: 1360px;
        padding-top: 0.8rem;
        padding-bottom: 3rem;
    }
    h1, h2,
    .brand-title,
    .metric-value {
        font-family: "Vinci Serif", Georgia, serif !important;
        letter-spacing: 0 !important;
    }
    h1, h2, h3, p {
        color: var(--leonard-ink);
    }
    h2 {
        font-size: 1.65rem !important;
        font-weight: 400 !important;
        margin: 0 0 0.35rem !important;
    }
    h3 {
        font-size: 1.05rem !important;
        font-weight: 700 !important;
        margin-top: 1.4rem !important;
    }

    [data-testid="stSidebar"] {
        background: var(--leonard-white);
        border-right: 1px solid var(--leonard-line);
    }
    [data-testid="stSidebar"] > div:first-child {
        padding-top: 1.2rem;
    }
    [data-testid="stSidebar"]::before {
        content: "";
        display: block;
        position: absolute;
        inset: 0 auto 0 0;
        width: 4px;
        background: var(--leonard-pink);
    }
    .sidebar-brand {
        padding: 0.35rem 0 1.15rem;
        border-bottom: 1px solid var(--leonard-line);
        margin-bottom: 1.25rem;
    }
    .sidebar-brand .sidebar-logo {
        display: block;
        width: 188px;
        max-width: 100%;
        height: auto;
        margin-bottom: 0.95rem;
    }
    .sidebar-product {
        color: var(--leonard-navy);
        font-family: "Vinci Serif", Georgia, serif;
        font-size: 1.22rem;
    }
    .sidebar-caption,
    .sidebar-label {
        color: var(--leonard-muted);
        font-size: 0.74rem;
        font-weight: 700;
        text-transform: uppercase;
        margin-top: 0.2rem;
    }
    .sidebar-label {
        color: var(--leonard-navy);
        margin: 0 0 0.45rem;
    }
    .system-status {
        border-top: 1px solid var(--leonard-line);
        border-bottom: 1px solid var(--leonard-line);
        padding: 0.85rem 0;
        margin: 1.25rem 0 0.75rem;
        color: var(--leonard-muted);
        font-size: 0.82rem;
    }
    .system-status strong {
        color: var(--leonard-ink);
        font-weight: 700;
    }
    .status-dot {
        display: inline-block;
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: #18a566;
        margin-right: 0.45rem;
    }
    .status-dot.warning {
        background: #e5a100;
    }

    .brand-header {
        position: relative;
        background: var(--constellation-field);
        border-bottom: 4px solid var(--leonard-pink);
        padding: 1.45rem 1.7rem 1.25rem;
        margin-bottom: 0.85rem;
        overflow: hidden;
        min-height: 250px;
    }
    .brand-header::before {
        content: "";
        position: absolute;
        inset: 0 auto auto 0;
        width: 34%;
        height: 3px;
        background: var(--leonard-blue);
        z-index: 2;
    }
    .constellation-orbit {
        position: absolute;
        width: 660px;
        height: 660px;
        right: -185px;
        top: -270px;
        opacity: 0.32;
        pointer-events: none;
    }
    .brand-content {
        position: relative;
        z-index: 1;
    }
    .brand-top,
    .brand-main {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 1.25rem;
    }
    .brand-logo {
        width: 215px;
        max-width: 42%;
        height: auto;
        display: block;
    }
    .brand-logo-light {
        filter: none;
    }
    .internal-badge,
    .profile-badge {
        display: inline-flex;
        align-items: center;
        min-height: 28px;
        padding: 0.25rem 0.65rem;
        border: 1px solid rgba(255, 255, 255, 0.42);
        color: #ffffff;
        background: rgba(255, 255, 255, 0.08);
        border-radius: 3px;
        font-size: 0.72rem;
        font-weight: 700;
        text-transform: uppercase;
        white-space: nowrap;
    }
    .brand-main {
        align-items: flex-end;
        margin-top: 1.35rem;
        margin-bottom: 1.3rem;
    }
    .brand-kicker,
    .section-kicker {
        color: var(--leonard-pink);
        font-size: 0.74rem;
        font-weight: 700;
        text-transform: uppercase;
        margin: 0 0 0.25rem;
    }
    .brand-title {
        color: #ffffff;
        font-size: 2.25rem;
        font-weight: 400;
        line-height: 1.05;
        margin: 0;
    }
    h1.brand-title {
        color: #ffffff !important;
    }
    .brand-subtitle {
        color: rgba(255, 255, 255, 0.78);
        font-size: 0.9rem;
        margin: 0.4rem 0 0;
    }
    .profile-badge {
        border-color: rgba(0, 180, 255, 0.8);
        background: rgba(0, 180, 255, 0.14);
    }
    .header-index {
        display: grid;
        justify-items: end;
        gap: 0.25rem;
        color: rgba(255, 255, 255, 0.72);
        font-size: 0.7rem;
        text-transform: uppercase;
    }
    .header-index strong {
        color: #ffffff;
        font-family: "Vinci Serif", Georgia, serif;
        font-size: 1.55rem;
        font-weight: 400;
        line-height: 1;
    }
    .workflow-strip {
        position: relative;
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 2rem;
        padding-top: 0.85rem;
    }
    .workflow-strip::before {
        content: "";
        position: absolute;
        left: 17px;
        right: calc(33.333% - 17px);
        top: 23px;
        height: 1px;
        background: rgba(255, 255, 255, 0.32);
    }
    .workflow-step {
        position: relative;
        display: grid;
        grid-template-columns: 34px 1fr;
        align-items: center;
        gap: 0.7rem;
        min-height: 54px;
        padding: 0.45rem 0;
    }
    .workflow-step + .workflow-step {
        padding-left: 0;
    }
    .workflow-number {
        position: relative;
        z-index: 1;
        display: grid;
        place-items: center;
        width: 34px;
        height: 34px;
        border: 1px solid rgba(255, 255, 255, 0.65);
        background: var(--constellation-field);
        color: #ffffff;
        font-size: 0.72rem;
        line-height: 1;
        transform: rotate(45deg);
    }
    .workflow-number span {
        transform: rotate(-45deg);
    }
    .workflow-step strong {
        display: block;
        color: #ffffff;
        font-size: 0.82rem;
    }
    .workflow-step span:last-child {
        display: block;
        color: rgba(255, 255, 255, 0.67);
        font-size: 0.72rem;
        margin-top: 0.08rem;
    }

    .section-heading {
        margin: 1.4rem 0 0.9rem;
    }
    .section-heading h2 {
        margin: 0 !important;
    }
    .status-box {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 1rem;
        border-left: 4px solid var(--leonard-blue);
        border-top: 1px solid var(--leonard-line);
        border-right: 1px solid var(--leonard-line);
        border-bottom: 1px solid var(--leonard-line);
        border-radius: 3px;
        padding: 0.78rem 0.95rem;
        background: var(--leonard-white);
        color: var(--leonard-ink);
        font-size: 0.88rem;
    }
    .status-box strong {
        color: var(--leonard-navy);
    }
    .status-box-meta {
        color: var(--leonard-muted);
        font-size: 0.78rem;
        white-space: nowrap;
    }
    .workspace-title {
        margin-bottom: 0.85rem;
    }
    .workspace-title strong {
        display: block;
        color: var(--leonard-ink);
        font-size: 1rem;
    }
    .workspace-title span {
        color: var(--leonard-muted);
        font-size: 0.78rem;
    }
    .run-context {
        min-height: 248px;
        border-top: 3px solid var(--constellation-violet);
        background: var(--leonard-white);
        padding: 1rem 1.05rem;
    }
    .run-context-title {
        color: var(--leonard-ink);
        font-family: "Vinci Serif", Georgia, serif;
        font-size: 1.05rem;
        margin-bottom: 0.8rem;
    }
    .context-row {
        display: grid;
        grid-template-columns: 26px minmax(0, 1fr);
        gap: 0.7rem;
        align-items: center;
        padding: 0.65rem 0;
        border-top: 1px solid var(--leonard-line);
    }
    .context-row:first-of-type {
        border-top: 0;
    }
    .context-key {
        display: grid;
        place-items: center;
        width: 24px;
        height: 24px;
        background: var(--constellation-field);
        color: #ffffff;
        font-size: 0.68rem;
        font-weight: 700;
    }
    .context-row strong {
        display: block;
        color: var(--leonard-ink);
        font-size: 0.78rem;
    }
    .context-row span:last-child {
        display: block;
        color: var(--leonard-muted);
        font-size: 0.72rem;
    }
    .file-ready {
        border-left: 3px solid var(--leonard-blue);
        background: var(--leonard-soft-blue);
        color: var(--leonard-ink);
        font-size: 0.78rem;
        margin: 0.7rem 0;
        padding: 0.65rem 0.75rem;
        overflow-wrap: anywhere;
    }

    .metric-card {
        position: relative;
        border: 1px solid var(--leonard-line);
        border-radius: 3px;
        background: var(--leonard-white);
        padding: 0.95rem 1rem;
        min-height: 112px;
        margin-bottom: 0.8rem;
    }
    .metric-card::before {
        content: "";
        position: absolute;
        inset: 0 auto 0 0;
        width: 4px;
        background: var(--leonard-blue);
    }
    .metric-card.metric-1::before {
        background: var(--leonard-pink);
    }
    .metric-card.metric-2::before {
        background: var(--leonard-navy);
    }
    .metric-label {
        color: var(--leonard-muted);
        font-size: 0.72rem;
        text-transform: uppercase;
        font-weight: 700;
    }
    .metric-value {
        color: var(--leonard-ink);
        font-size: 1.75rem;
        font-weight: 400;
        margin-top: 0.28rem;
    }
    .metric-hint {
        color: var(--leonard-muted);
        font-size: 0.76rem;
        margin-top: 0.15rem;
    }

    .stButton > button,
    .stDownloadButton > button {
        min-height: 42px;
        border-radius: 3px !important;
        border: 1px solid var(--leonard-navy);
        background: var(--leonard-white);
        color: var(--leonard-navy);
        font-family: "Vinci Sans", Arial, sans-serif;
        font-weight: 700;
        box-shadow: none !important;
    }
    .stButton > button:hover,
    .stDownloadButton > button:hover {
        border-color: var(--leonard-pink);
        color: var(--leonard-pink);
    }
    .stButton > button[kind="primary"] {
        background: var(--leonard-pink);
        border-color: var(--leonard-pink);
        color: var(--leonard-white);
    }
    .stButton > button[kind="primary"]:hover {
        background: #d9004d;
        border-color: #d9004d;
        color: var(--leonard-white);
    }
    .stButton > button[kind="primary"]:disabled {
        background: #d8dee5;
        border-color: #d8dee5;
        color: #718090;
        opacity: 1;
        cursor: not-allowed;
    }
    [data-testid="stFileUploaderDropzone"] {
        min-height: 150px;
        border: 1px dashed var(--leonard-blue);
        border-radius: 3px;
        background: var(--leonard-white);
    }
    [data-testid="stFileUploaderDropzone"] button {
        border-radius: 3px !important;
        border-color: var(--leonard-navy);
        color: var(--leonard-navy);
    }
    [data-testid="stProgress"] > div > div > div > div {
        background: var(--leonard-pink);
    }
    [data-testid="stAlert"] {
        border-radius: 3px;
        border: 1px solid var(--leonard-line);
    }
    [data-testid="stExpander"] {
        border-color: var(--leonard-line);
        border-radius: 3px;
        background: var(--leonard-white);
    }
    [data-testid="stVerticalBlockBorderWrapper"] {
        border-color: var(--leonard-line) !important;
        border-radius: 3px !important;
        background: var(--leonard-white);
    }
    [data-baseweb="select"] > div,
    .stTextInput input {
        border-radius: 3px !important;
        border-color: var(--leonard-line) !important;
    }

    .stTabs [data-baseweb="tab-list"] {
        width: fit-content;
        gap: 4px;
        padding: 4px;
        border: 1px solid var(--leonard-line);
        background: var(--leonard-white);
        margin-bottom: 0.75rem;
    }
    .stTabs [data-baseweb="tab"] {
        height: 40px;
        padding: 0 1.2rem;
        border-radius: 2px !important;
        color: var(--leonard-muted);
        font-weight: 700;
        background: transparent;
    }
    .stTabs [aria-selected="true"] {
        color: #ffffff !important;
        background: var(--leonard-navy) !important;
        box-shadow: none;
    }
    .stTabs [data-baseweb="tab-highlight"] {
        display: none;
    }
    code {
        color: var(--leonard-navy) !important;
        background: var(--leonard-soft-blue) !important;
        border-radius: 2px !important;
    }
    [data-testid="stIconMaterial"] {
        font-family: "Material Symbols Rounded" !important;
        font-weight: normal !important;
        font-style: normal !important;
        line-height: 1 !important;
        letter-spacing: normal !important;
        text-transform: none !important;
        white-space: nowrap !important;
        word-wrap: normal !important;
        direction: ltr !important;
        -webkit-font-feature-settings: "liga" !important;
        -webkit-font-smoothing: antialiased !important;
    }

    @media (max-width: 760px) {
        .block-container {
            padding: 0.65rem 0.85rem 2rem;
        }
        .brand-header {
            min-height: 0;
            padding: 1rem 1rem 1.15rem;
        }
        .brand-top,
        .brand-main {
            align-items: flex-start;
        }
        .brand-main {
            flex-direction: column;
            gap: 0.75rem;
        }
        .brand-logo {
            max-width: 65%;
        }
        .brand-title {
            font-size: 1.65rem;
        }
        .constellation-orbit {
            width: 470px;
            height: 470px;
            right: -260px;
            top: -170px;
            opacity: 0.24;
        }
        .header-index {
            justify-items: start;
        }
        .workflow-strip {
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 0.45rem;
            padding-top: 0.45rem;
        }
        .workflow-strip::before {
            left: 14px;
            right: calc(33.333% - 14px);
            top: 19px;
            bottom: auto;
            width: auto;
            height: 1px;
        }
        .workflow-step,
        .workflow-step + .workflow-step {
            grid-template-columns: 28px minmax(0, 1fr);
            align-items: start;
            gap: 0.5rem;
            min-height: 62px;
            padding: 0.3rem 0;
            border-right: 0;
            border-bottom: 0;
        }
        .workflow-number {
            width: 28px;
            height: 28px;
            font-size: 0.64rem;
        }
        .workflow-step strong {
            font-size: 0.72rem;
        }
        .workflow-step > span:last-child > span:last-child {
            display: none;
        }
        .status-box {
            align-items: flex-start;
            flex-direction: column;
        }
        .status-box-meta {
            white-space: normal;
        }
        .stTabs [data-baseweb="tab"] {
            padding: 0 0.7rem;
        }
        .stTabs [data-baseweb="tab-list"] {
            width: 100%;
        }
    }
    """.replace("__CONSTELLATION_ORBIT__", orbit_uri)
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


def render_header(profile: str) -> None:
    logo_svg = inline_svg(
        "leonard-vinci-logo.svg",
        "brand-logo brand-logo-light",
        {
            "#c41d6d": "#ffffff",
            "#00b4ff": "#ffffff",
            "#004489": "#ffffff",
            "#ff005a": "#ffffff",
        },
    )
    orbit_svg = inline_svg("constellation-orbit.svg", "constellation-orbit")
    profile_label = "Europe" if profile == dealflow_profiles.EUROPE else "LATAM"
    logo_markup = logo_svg or '<div class="brand-title">Leonard / VINCI</div>'
    st.markdown(
        f"""
        <header class="brand-header">
            {orbit_svg}
            <div class="brand-content">
                <div class="brand-top">
                    {logo_markup}
                    <span class="internal-badge">Usage interne</span>
                </div>
                <div class="brand-main">
                    <div>
                        <div class="brand-kicker">Dealflow constellation</div>
                        <h1 class="brand-title">Startup Dealflow</h1>
                        <p class="brand-subtitle">Qualification stratégique pour les métiers et programmes Leonard.</p>
                    </div>
                    <div class="header-index">
                        <span class="profile-badge">Profil {profile_label}</span>
                        <span><strong>3</strong> niveaux d'analyse</span>
                    </div>
                </div>
                <div class="workflow-strip" aria-label="Parcours de qualification">
                    <div class="workflow-step">
                        <span class="workflow-number"><span>01</span></span>
                        <span><strong>Éligibilité</strong><span>Premier filtre</span></span>
                    </div>
                    <div class="workflow-step">
                        <span class="workflow-number"><span>02</span></span>
                        <span><strong>Alignement</strong><span>Enjeux stratégiques</span></span>
                    </div>
                    <div class="workflow-step">
                        <span class="workflow-number"><span>03</span></span>
                        <span><strong>Orientation</strong><span>Seed · Catalyst · Matériaux</span></span>
                    </div>
                </div>
            </div>
        </header>
        """,
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
        dealflow_profiles.EUROPE: "Europe",
        dealflow_profiles.LATAM: "LATAM",
    }
    logo_markup = inline_svg("leonard-vinci-logo.svg", "sidebar-logo")
    st.sidebar.markdown(
        f"""
        <div class="sidebar-brand">
            {logo_markup}
            <div class="sidebar-product">Dealflow Studio</div>
            <div class="sidebar-caption">Qualification & orientation</div>
        </div>
        <div class="sidebar-label">Profil d'analyse</div>
        """,
        unsafe_allow_html=True,
    )
    selected_profile = st.sidebar.selectbox(
        "Zone d'analyse",
        options=dealflow_profiles.AVAILABLE_PROFILES,
        index=dealflow_profiles.AVAILABLE_PROFILES.index(default_profile),
        format_func=lambda value: profile_labels.get(value, value),
        label_visibility="collapsed",
    )
    os.environ["DEALFLOW_PROFILE"] = selected_profile
    if selected_profile == dealflow_profiles.LATAM:
        st.sidebar.caption("Critères LATAM · analyse en anglais · seuil ARR adapté")
    else:
        st.sidebar.caption("Critères Europe · référentiel Leonard actuel")

    configuration_ready = bool(endpoint and api_key_present and deployment)
    status_class = "" if configuration_ready else " warning"
    status_label = "Système prêt" if configuration_ready else "Configuration incomplète"
    st.sidebar.markdown(
        f"""
        <div class="system-status">
            <div><span class="status-dot{status_class}"></span><strong>{status_label}</strong></div>
            <div>{html.escape(deployment)} · Azure OpenAI</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.sidebar.expander("Configuration technique", expanded=False):
        st.write(f"Modèle : `{deployment}`")
        st.write(f"API : `{api_version}`")
        st.write("Endpoint : " + ("configuré" if endpoint else "manquant"))
        st.write("Clé : " + ("configurée" if api_key_present else "manquante"))
        st.write(f"Runs : `{RUNS_DIR.name}`")
        st.caption("La clé API n'est jamais affichée.")
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
    columns = st.columns(min(len(kpis), 4))
    for index, kpi in enumerate(kpis):
        with columns[index % len(columns)]:
            label = html.escape(str(kpi.get("label", "")))
            value = html.escape(str(kpi.get("value", "")))
            hint = html.escape(str(kpi.get("hint", "")))
            st.markdown(
                f"""
                <div class="metric-card metric-{index % 3}">
                    <div class="metric-label">{label}</div>
                    <div class="metric-value">{value}</div>
                    <div class="metric-hint">{hint}</div>
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
    st.markdown(
        '<div class="section-heading"><div class="section-kicker">Livrables</div><h2>Résultats</h2></div>',
        unsafe_allow_html=True,
    )
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
    profile_label = "Europe" if profile == dealflow_profiles.EUROPE else "LATAM"
    st.markdown(
        f"""
        <div class="section-heading">
            <div class="section-kicker">Nouveau dossier</div>
            <h2>Lancer une qualification</h2>
        </div>
        <div class="status-box">
            <span><strong>Profil {profile_label}</strong> · Pipeline Algo 1 → Algo 2 → Algo 3</span>
            <span class="status-box-meta">Excel source · .xlsx, .xlsm ou .xls</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    active_meta = st.session_state.get("active_job_meta")
    if active_meta:
        render_job_monitor(active_meta, "active")
        if job_state(active_meta) not in {"running", "stopping"}:
            if st.button("Lancer une nouvelle analyse", use_container_width=True):
                st.session_state.pop("active_job_meta", None)
                st.rerun()
        return

    recent_jobs = iter_job_metas(limit=20)
    running_jobs = [
        meta
        for meta in recent_jobs
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

    resumable_jobs = [meta for meta in recent_jobs if can_resume_job(meta)]
    if resumable_jobs:
        st.markdown("### Analyses interrompues")
        for index, meta in enumerate(resumable_jobs[:5]):
            with st.container(border=True):
                st.write(f"**{job_display_name(meta)}**")
                st.caption(f"Profil : {meta.get('profile', '')} · sauvegarde disponible")
                if st.button(
                    "Reprendre ce run",
                    key=f"resume-stopped-{index}",
                    use_container_width=True,
                ):
                    resumed_meta = resume_background_job(meta)
                    st.session_state["active_job_meta"] = resumed_meta
                    st.rerun()

    workspace, context = st.columns([1.65, 0.75], gap="large")
    with workspace:
        with st.container(border=True):
            st.markdown(
                """
                <div class="workspace-title">
                    <strong>Source à analyser</strong>
                    <span>Portefeuille Excel conservé avec toutes ses colonnes.</span>
                </div>
                """,
                unsafe_allow_html=True,
            )
            uploaded_file = st.file_uploader(
                "Portefeuille de startups",
                type=["xlsx", "xlsm", "xls"],
                accept_multiple_files=False,
                label_visibility="collapsed",
            )
            if uploaded_file:
                st.markdown(
                    f'<div class="file-ready"><strong>Fichier prêt</strong><br>{html.escape(uploaded_file.name)}</div>',
                    unsafe_allow_html=True,
                )
            start = st.button(
                "Lancer la qualification",
                type="primary",
                use_container_width=True,
                disabled=uploaded_file is None,
            )
            st.caption("Exécution en arrière-plan avec sauvegarde progressive.")

    with context:
        st.markdown(
            f"""
            <aside class="run-context">
                <div class="run-context-title">Contexte du run</div>
                <div class="context-row">
                    <span class="context-key">01</span>
                    <span><strong>Profil {profile_label}</strong><span>Référentiel régional actif</span></span>
                </div>
                <div class="context-row">
                    <span class="context-key">02</span>
                    <span><strong>{len(running_jobs)} analyse(s) active(s)</strong><span>Suivi des processus actifs</span></span>
                </div>
                <div class="context-row">
                    <span class="context-key">03</span>
                    <span><strong>{len(resumable_jobs)} reprise(s) disponible(s)</strong><span>Sauvegardes détectées</span></span>
                </div>
            </aside>
            """,
            unsafe_allow_html=True,
        )

    if not uploaded_file:
        return

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
    st.markdown(
        """
        <div class="section-heading">
            <div class="section-kicker">Visualisation</div>
            <h2>Générer un dashboard</h2>
        </div>
        """,
        unsafe_allow_html=True,
    )
    uploaded_file = st.file_uploader(
        "Résultat du pipeline",
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
    st.markdown(
        """
        <div class="section-heading">
            <div class="section-kicker">Archives</div>
            <h2>Historique des analyses</h2>
        </div>
        """,
        unsafe_allow_html=True,
    )
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
        page_title="Startup Dealflow · Leonard",
        page_icon=str(ASSETS_DIR / "leonard-mark.png"),
        layout="wide",
        initial_sidebar_state="auto",
    )
    render_css()
    profile = render_config_status()
    render_header(profile)

    tab_run, tab_dashboard, tab_history = st.tabs([
        "Analyse",
        "Dashboard",
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
