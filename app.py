"""
app.py — Web (Streamlit) del Cambridge Writing Corrector
=========================================================
Corrige las redacciones de un Excel que está en Google Drive y escribe cada
corrección DIRECTAMENTE en ese mismo archivo (no crea copias ni descargas).

Ejecutar en local:   streamlit run app.py
Publicar:            Streamlit Community Cloud (ver GUIA_STREAMLIT.md)

Secrets necesarios (.streamlit/secrets.toml en local, o «Secrets» en Streamlit Cloud):
  APP_PASSWORD      = "contraseña para entrar en la web"
  GOOGLE_API_KEY    = "clave de la API de Gemini"
  DRIVE_FOLDER_ID   = "ID de la carpeta compartida de Drive"
  [gcp_service_account]   ← contenido del JSON de la cuenta de servicio
"""

import hmac
import io
import os

import streamlit as st
from openpyxl import load_workbook

from corrector_cambridge_v4_b1_b2_c1_c2 import (
    CambridgeCorrector,
    process_workbook,
    LEVEL_NAMES,
)
from drive_io import DriveClient


# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Cambridge Writing Corrector",
    page_icon="🎓",
    layout="centered",
)

st.markdown("""
<style>
div[data-testid="stButton"] > button[kind="primary"] {
    background-color: #2e75b6; color: white; font-weight: bold;
    border-radius: 6px; height: 3em; width: 100%;
}
</style>
""", unsafe_allow_html=True)


def get_secret(name: str, default=None):
    """Lee un secret de Streamlit; si no existe, de las variables de entorno."""
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass   # no hay secrets.toml
    return os.environ.get(name, default)


# ─────────────────────────────────────────────────────────────
# CONTRASEÑA
# ─────────────────────────────────────────────────────────────

def check_password() -> bool:
    if st.session_state.get("authed"):
        return True

    expected = get_secret("APP_PASSWORD")
    st.title("🎓 Cambridge Writing Corrector")
    if not expected:
        st.error("Falta configurar APP_PASSWORD en los secrets. La web no se abre sin contraseña.")
        return False

    with st.form("login"):
        pwd = st.text_input("Contraseña", type="password")
        ok = st.form_submit_button("Entrar", type="primary")
    if ok:
        if hmac.compare_digest(pwd.encode(), str(expected).encode()):
            st.session_state.authed = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    return False


if not check_password():
    st.stop()


# ─────────────────────────────────────────────────────────────
# CONEXIÓN CON DRIVE
# ─────────────────────────────────────────────────────────────

@st.cache_resource(show_spinner=False)
def get_drive() -> DriveClient:
    sa_info = get_secret("gcp_service_account")
    folder  = get_secret("DRIVE_FOLDER_ID")
    if not sa_info or not folder:
        raise RuntimeError("Faltan los secrets 'gcp_service_account' o 'DRIVE_FOLDER_ID'.")
    return DriveClient(dict(sa_info), folder)


@st.cache_data(ttl=120, show_spinner=False)
def load_listings(_drive: DriveClient, refresh_token: int):
    """Listados de Drive (se guardan 2 minutos; el botón 🔄 los recarga)."""
    root_name = _drive.check_root()
    levels = {lvl: _drive.get_level_files(lvl) for lvl in LEVEL_NAMES}
    return {
        "root_name":  root_name,
        "excels":     _drive.list_excels(),
        "enunciados": _drive.list_enunciados(),
        "levels":     levels,
    }


# ─────────────────────────────────────────────────────────────
# ESTADO DE SESIÓN
# ─────────────────────────────────────────────────────────────

defaults = {
    "running": False,       # hay una corrección en marcha
    "job": None,            # parámetros de la corrección en marcha
    "results": [],          # [{"Alumno", "Estado", "Nota"}]
    "stats": None,          # resultado final
    "stopped_by_user": False,
    "refresh": 0,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ─────────────────────────────────────────────────────────────
# CABECERA
# ─────────────────────────────────────────────────────────────

st.title("🎓 Cambridge Writing Corrector")
st.caption("Corrige las redacciones del Excel de Drive y guarda cada corrección en el mismo archivo.")

try:
    drive = get_drive()
    with st.spinner("Leyendo la carpeta de Drive…"):
        data = load_listings(drive, st.session_state.refresh)
except Exception as e:
    st.error(f"❌ No se pudo acceder a Google Drive: {e}")
    st.info("Comprueba que la carpeta está compartida como **Editor** con el email de la "
            "cuenta de servicio y que DRIVE_FOLDER_ID es correcto (ver GUIA_STREAMLIT.md).")
    st.stop()

top_l, top_r = st.columns([4, 1])
top_l.markdown(f"📁 Carpeta de Drive: **{data['root_name']}**")
if top_r.button("🔄 Actualizar", disabled=st.session_state.running,
                help="Vuelve a leer los archivos de Drive"):
    st.session_state.refresh += 1
    st.rerun()

st.markdown("---")


# ─────────────────────────────────────────────────────────────
# FORMULARIO
# ─────────────────────────────────────────────────────────────

running = st.session_state.running

col1, col2 = st.columns(2)
with col1:
    level = st.selectbox(
        "🎯 Nivel",
        options=list(LEVEL_NAMES),
        format_func=lambda l: LEVEL_NAMES[l],
        disabled=running,
    )
level_files = data["levels"][level]
type_options = sorted(level_files["types"])
with col2:
    essay_type = st.selectbox(
        "📝 Tipo de redacción",
        options=type_options,
        format_func=str.capitalize,
        disabled=running or not type_options,
        help=f"Se leen de redacciones/{level}/ ({level.lower()}_<tipo>.md)",
    )

excels = data["excels"]
excel_file = st.selectbox(
    "📊 Excel de estudiantes",
    options=excels,
    format_func=lambda f: f["name"],
    index=None,
    placeholder="Elige un Excel de la carpeta estudiantes/",
    disabled=running,
)

NO_ENUNCIADO = {"id": None, "name": "(Sin enunciado)"}
enunciado_file = st.selectbox(
    "📄 Enunciado (opcional)",
    options=[NO_ENUNCIADO] + data["enunciados"],
    format_func=lambda f: f["name"],
    disabled=running,
)

# ── Validaciones ──────────────────────────────────────────────
problems = []
if not level_files["rubric"]:
    problems.append(f"Falta la rúbrica `redacciones/{level}/{level.lower()}_rubrica.md` en Drive.")
if not type_options:
    problems.append(f"No hay ejemplos de corrección en `redacciones/{level}/` "
                    f"(archivos tipo `{level.lower()}_essay.md`).")
if not excels:
    problems.append("No hay archivos .xlsx en la carpeta `estudiantes/` de Drive.")
for p in problems:
    st.warning(p)

ready = not problems and excel_file is not None


# ─────────────────────────────────────────────────────────────
# BOTONES
# ─────────────────────────────────────────────────────────────

st.markdown("---")
b1, b2 = st.columns(2)

with b1:
    if st.button("🚀 Corregir redacciones", type="primary", disabled=not ready or running):
        st.session_state.running = True
        st.session_state.stopped_by_user = False
        st.session_state.results = []
        st.session_state.stats = None
        st.session_state.job = {
            "level": level,
            "essay_type": essay_type,
            "excel": excel_file,
            "rubric": level_files["rubric"],
            "example": level_files["types"][essay_type],
            "enunciado": enunciado_file if enunciado_file["id"] else None,
        }
        st.rerun()

with b2:
    if running:
        # Pulsarlo interrumpe la ejecución en curso. No se pierde nada:
        # cada alumno se guarda en Drive en cuanto se corrige.
        if st.button("⏹️ Detener corrección"):
            st.session_state.running = False
            st.session_state.stopped_by_user = True
            st.rerun()


# ─────────────────────────────────────────────────────────────
# TABLA DE RESULTADOS
# ─────────────────────────────────────────────────────────────

def render_results(placeholder):
    rows = st.session_state.results
    if rows:
        placeholder.dataframe(rows, hide_index=True, width="stretch")


# ─────────────────────────────────────────────────────────────
# EJECUCIÓN
# ─────────────────────────────────────────────────────────────

if st.session_state.running:
    job = st.session_state.job
    st.markdown("### Progreso")
    st.caption("Cada corrección se guarda en Drive al momento. Si se interrumpe, "
               "vuelve a pulsar «Corregir» y seguirá por donde iba.")
    progress_bar = st.progress(0.0, text="Preparando…")
    table = st.empty()

    try:
        # 1) Archivos de referencia y corrector
        with st.spinner("Cargando rúbrica, ejemplo y enunciado desde Drive…"):
            rubric_text  = drive.read_text(job["rubric"])
            example_text = drive.read_text(job["example"])
            enunciado_text = drive.read_text(job["enunciado"]) if job["enunciado"] else ""
            corrector = CambridgeCorrector(
                job["level"], job["essay_type"], rubric_text, example_text,
                enunciado_text, api_key=get_secret("GOOGLE_API_KEY"),
            )

        # 2) Excel original
        with st.spinner("Descargando el Excel…"):
            file_id = job["excel"]["id"]
            drive.keep_current_version(file_id)   # copia de seguridad en el historial
            wb = load_workbook(io.BytesIO(drive.download_bytes(file_id)))

        def save_to_drive(workbook):
            buffer = io.BytesIO()
            workbook.save(buffer)
            drive.overwrite_xlsx(file_id, buffer.getvalue())

        def on_progress(current, total, name, status):
            pct = current / total if total else 1.0
            if status == "correcting":
                progress_bar.progress((current - 1) / total,
                                      text=f"Corrigiendo {current}/{total}: {name}…")
                return
            if status == "summary":
                progress_bar.progress(1.0, text="Actualizando la hoja Resumen…")
                return
            if status == "skipped":
                st.session_state.results.append(
                    {"Alumno": name, "Estado": "⏭️ Ya corregido", "Nota": ""})
            elif status.startswith("done:"):
                score = status.split(":", 1)[1]
                st.session_state.results.append(
                    {"Alumno": name, "Estado": "✅ Corregido",
                     "Nota": f"{score}/20" if score else "—"})
            elif status.startswith("error:"):
                st.session_state.results.append(
                    {"Alumno": name, "Estado": f"❌ {status.split(':', 1)[1][:120]}",
                     "Nota": ""})
            progress_bar.progress(pct, text=f"Procesados {current}/{total}")
            render_results(table)

        # 3) Corrección (guarda en Drive tras cada alumno)
        stats = process_workbook(wb, corrector, save_fn=save_to_drive,
                                 progress_callback=on_progress)
        stats["file_id"] = file_id
        stats["file_name"] = job["excel"]["name"]
        st.session_state.stats = stats

    except Exception as e:
        st.session_state.stats = {"error": str(e)}

    st.session_state.running = False
    st.rerun()


# ─────────────────────────────────────────────────────────────
# RESULTADO FINAL
# ─────────────────────────────────────────────────────────────

stats = st.session_state.stats

if st.session_state.stopped_by_user:
    st.info("⏹️ Corrección detenida. Lo corregido hasta ahora ya está guardado en el Excel de Drive. "
            "Para continuar, vuelve a pulsar «Corregir».")

if stats and "error" in stats:
    st.error(f"❌ La corrección se interrumpió: {stats['error']}")
    st.info("Lo corregido antes del error ya está guardado. Vuelve a pulsar «Corregir» para continuar.")
elif stats:
    st.markdown("### ✅ Corrección terminada")
    m1, m2, m3 = st.columns(3)
    m1.metric("Corregidas", stats["success"])
    m2.metric("Ya estaban corregidas", stats["skipped"])
    m3.metric("Con error", len(stats["failed"]))
    if stats["failed"]:
        st.warning("No se pudieron corregir: " + ", ".join(stats["failed"])
                   + ". Vuelve a pulsar «Corregir» para reintentarlo.")
    st.link_button("📂 Abrir el Excel en Drive",
                   f"https://drive.google.com/file/d/{stats['file_id']}/view")

if st.session_state.results and not st.session_state.running:
    render_results(st.empty())


st.markdown("---")
st.caption("Cambridge Writing Corrector · Las correcciones se guardan solo en tu Excel de Drive.")
