"""
Cambridge Writing Corrector — versión Excel
============================================
Lee las redacciones de la columna E del Excel (formato Google Forms),
corrige cada una con Gemini y escribe el resultado desglosado en F-K.

Estructura del Excel esperada:
  A: Marca temporal
  B: Email del alumno
  C: Puntuación (Google Forms)
  D: Nombre del alumno
  E: Redacción original
  F: Grade Summary
  G: Nota /20
  H: Error Analysis
  I: Examiner Comments
  J: Corrected Essay
  K: Corrección completa (todo junto)

Uso:
  1. Instalación de dependencias (ejecutar una sola vez):
    pip install google-genai
    pip install openpyxl
    pip install python-docx
    pip install python-dotenv

    Librerías de la librería estándar de Python (no requieren instalación):
    os, sys, re, time, shutil, pathlib — vienen incluidas con Python.

  2. Ejecuta: python corrector_cambridge_v4_b1_b2_c1_c2.py
     (o usa la web: streamlit run app.py — ver GUIA_STREAMLIT.md)

Estructura de carpetas esperada:
  tu_carpeta_proyecto/
  │
  ├── corrector_cambridge_v4.py     ← este script
  │── .env                          ← contiene la API_KEY (API_KEY="tu_api_key")
  │
  ├── redacciones/                  ← archivos de referencia para la IA, separados por nivel
  │   ├── B1/
  │   │   ├── b1_rubrica.md         ← rúbrica oficial de Cambridge para B1 (en Markdown)
  │   │   ├── b1_essay.md           ← ejemplo de corrección de Essay (B1)
  │   │   └── b1_report.md          ← ejemplo de corrección de Report (B1)
  │   ├── B2/
  │   │   ├── b2_rubrica.md
  │   │   ├── b2_essay.md
  │   │   └── b2_proposal.md
  │   ├── C1/
  │   │   ├── c1_rubrica.md
  │   │   ├── c1_essay.md
  │   │   └── c1_proposal.md
  │   └── C2/
  │       ├── c2_rubrica.md
  │       └── c2_letter.md
  │
  ├── estudiantes/                  ← archivos Excel con las redacciones
  │   └── Writing_4B.xlsx           ← el Excel que sube el usuario
  │
  ├── estudiantes_backup/           ← copias de seguridad (crear antes de usar)
  │   └── Writing_4B_BACKUP.xlsx    ← generado automáticamente por el programa
  │
  └── enunciados/                   ← (opcional) enunciados de los ejercicios
      └── Proposal_enunciado.docx   ← archivo .docx con el task prompt

Notas:
  - Las carpetas de nivel dentro de redacciones/ van en mayúsculas (B1, B2, C1, C2).
  - Los archivos .md dentro de cada carpeta de nivel van en minúsculas y con el
    prefijo del nivel: <nivel>_rubrica.md (p.ej. b2_rubrica.md) y
    <nivel>_<tipo>.md (p.ej. c1_proposal.md).
  - El Excel debe tener las redacciones en la columna E (formato Google Forms).
"""

import os
import sys
import re
import time
import shutil
from pathlib import Path

# ── Librería nueva (google-genai, reemplaza google-generativeai) ───────────────
# Instala con: pip install google-genai
from google import genai
from google.genai import types

from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill, Font

from dotenv import load_dotenv
load_dotenv(dotenv_path=Path(__file__).parent / ".env")

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────

MODEL_NAME = "gemma-4-31b-it"
TEMPERATURE = 0.1
TOP_P = 0.9
TOP_K = 40

DELAY_BETWEEN_CALLS = 0

# Columnas del Excel (índice 1-based de openpyxl)
COL_NOMBRE     = 4   # D
COL_REDACCION  = 5   # E
COL_GRADE      = 6   # F — Grade Summary
COL_NOTA       = 7   # G — Nota /20
COL_ERRORS     = 8   # H — Error Analysis
COL_COMMENTS   = 9   # I — Examiner Comments
COL_CORRECTED  = 10  # J — Corrected Essay
COL_FULL       = 11  # K — Corrección completa
COL_FULL_CLEAN = 12  # L — Corrección completa limpia (sin ##)


# ─────────────────────────────────────────────────────────────
# INICIALIZACIÓN DE LA API
# ─────────────────────────────────────────────────────────────

# El cliente se crea al construir el corrector (no al importar el módulo),
# para que la web pueda pasar su propia clave desde los secrets de Streamlit.

def make_client(api_key: str = None):
    """Crea el cliente de Gemini con la clave dada o la de GOOGLE_API_KEY."""
    key = api_key or os.environ.get("GOOGLE_API_KEY", "")
    if not key:
        raise RuntimeError(
            "No se encontró GOOGLE_API_KEY. Ponla en el archivo .env "
            "(GOOGLE_API_KEY=tu_clave) o en los secrets de Streamlit."
        )
    return genai.Client(api_key=key)


# ─────────────────────────────────────────────────────────────
# CLASE PRINCIPAL
# ─────────────────────────────────────────────────────────────

# Nombre real del examen Cambridge asociado a cada nivel CEFR
LEVEL_NAMES = {
    "B1": "B1 Preliminary",
    "B2": "B2 First",
    "C1": "C1 Advanced",
    "C2": "C2 Proficiency",
}


def build_system_instruction(level: str) -> str:
    """Construye el SYSTEM_INSTRUCTION para el nivel indicado (B1/B2/C1/C2)."""
    exam_name = LEVEL_NAMES.get(level, level)
    return (
    f"You are an expert Cambridge Examiner for {exam_name} level. "
    "Your ONLY references for grading are the official rubric and correction example provided. "
    "Apply the rubric criteria consistently and literally across every essay you correct.\n\n"

    "For each essay you receive, produce the following sections exactly:\n\n"

    "## GRADE SUMMARY\n"
    "Score each criterion on the 0-5 scale defined in the rubric, followed by a one-sentence justification. Format:\n"
    "  Content: X/5 — [justification]\n"
    "  Communicative Achievement: X/5 — [justification]\n"
    "  Organization: X/5 — [justification]\n"
    "  Language: X/5 — [justification]\n"
    "  TOTAL: X/20\n\n"

    "## ERROR ANALYSIS\n"
    "List every significant error found. Use EXACTLY this format for each error:\n"
    '  - "original fragment" → Error type → Corrected form → One-sentence explanation\n'
    'Example: - "could be resolve" → Grammar → could be resolved → The passive voice requires the past participle.\n\n'

    "## EXAMINER COMMENTS\n"
    "2 to 4 sentences of global feedback: strengths, main weaknesses, and the most important area to improve.\n\n"

    "## CORRECTED ESSAY\n"
    "A full rewrite of the essay correcting ALL mistakes while keeping the student's original ideas, structure, and approximate length. "
    "Do not use any markdown formatting (no asterisks, no # symbols). Write all section headings in UPPERCASE PLAIN TEXT (e.g. INTRODUCTION, CONCLUSION). "
    "This section must be self-contained — do not reference the original text.\n\n"

    "CONSISTENCY RULES (critical):\n"
    "- A score of 3/5 in Language means the same thing in every essay.\n"
    "- Do not reward or penalise based on essay length alone.\n"
    "- If two essays have equivalent quality, they must receive identical scores.\n"
    "- Errors must always be classified in the same way regardless of the student."
    )


class CambridgeCorrector:

    def __init__(self, level: str, essay_type: str, rubric_text: str, example_text: str,
                 enunciado_text: str = "", api_key: str = None):
        """
        rubric_text / example_text: contenido (texto) de la rúbrica y del ejemplo.
        La versión de consola los lee de disco; la web los lee de Google Drive.
        """
        self.level = level
        self.essay_type = essay_type
        self.enunciado_text = enunciado_text
        self.system_instruction = build_system_instruction(level)
        self.rubric_text = rubric_text
        self.example_text = example_text
        self.client = make_client(api_key)

    # ── Llamada a la API con reintentos ────────────────────────

    def _call_api(self, essay_text: str, student_name: str, retries: int = 3) -> str:
        """
        Llama a la API y devuelve el texto de respuesta.
        Lanza excepción si todos los reintentos fallan. Nunca devuelve None.
        """
        enunciado_block = (
            f"TASK PROMPT (what the student was asked to write):\n\"\"\"\n{self.enunciado_text}\n\"\"\"\n\n"
            if self.enunciado_text else ""
        )
        prompt_text = (
            f"Please correct the following student essay for the '{self.essay_type}' task. "
            "Use ONLY the attached rubric and correction example as your grading reference. "
            f"Student name: {student_name}\n\n"
            f"{enunciado_block}"
            f"STUDENT ESSAY:\n\"\"\"\n{essay_text}\n\"\"\""
        )

        contents = [
            types.Part.from_text(text=(
                f"OFFICIAL RUBRIC:\n\"\"\"\n{self.rubric_text}\n\"\"\"\n\n"
                f"CORRECTION EXAMPLE:\n\"\"\"\n{self.example_text}\n\"\"\"\n\n"
                + prompt_text
            )),
        ]

        config = types.GenerateContentConfig(
            system_instruction=self.system_instruction,
            temperature=TEMPERATURE,
            top_p=TOP_P,
            top_k=TOP_K,
        )

        last_error = None
        for attempt in range(retries):
            try:
                response = self.client.models.generate_content(
                    model=MODEL_NAME,
                    contents=contents,
                    config=config,
                )
                text = response.text
                if text is None or not text.strip():
                    raise ValueError("La API devolvió una respuesta vacía.")
                return text.strip()

            except Exception as e:
                last_error = e
                err = str(e).lower()
                if "quota" in err or "rate" in err or "429" in err or "resource_exhausted" in err:
                    wait = 60 * (attempt + 1)   # 60s → 120s → 180s
                    print(f"    Rate limit alcanzado. Esperando {wait}s...")
                    time.sleep(wait)
                elif attempt < retries - 1:
                    print(f"    Error en API (intento {attempt+1}): {e}. Reintentando en 15s...")
                    time.sleep(15)
                else:
                    break

        raise RuntimeError(f"API falló tras {retries} intentos: {last_error}")

    # ── Corrección de una redacción ────────────────────────────

    def correct_essay(self, student_name: str, essay_text: str) -> str:
        return self._call_api(essay_text, student_name)


# ─────────────────────────────────────────────────────────────
# PARSING DE SECCIONES
# ─────────────────────────────────────────────────────────────

def parse_sections(full_text: str) -> dict:
    """Divide el texto completo en sus cuatro secciones."""
    sections = {"grade": "", "errors": "", "comments": "", "corrected": ""}
    if not full_text:
        return sections

    blocks = re.split(r"(?=##\s)", full_text)
    for block in blocks:
        b = block.strip()
        if re.match(r"##\s*GRADE SUMMARY", b, re.IGNORECASE):
            sections["grade"] = re.sub(r"^##\s*GRADE SUMMARY\s*", "", b, flags=re.IGNORECASE).strip()
        elif re.match(r"##\s*ERROR ANALYSIS", b, re.IGNORECASE):
            sections["errors"] = re.sub(r"^##\s*ERROR ANALYSIS\s*", "", b, flags=re.IGNORECASE).strip()
        elif re.match(r"##\s*EXAMINER COMMENTS", b, re.IGNORECASE):
            sections["comments"] = re.sub(r"^##\s*EXAMINER COMMENTS\s*", "", b, flags=re.IGNORECASE).strip()
        elif re.match(r"##\s*CORRECTED ESSAY", b, re.IGNORECASE):
            sections["corrected"] = re.sub(r"^##\s*CORRECTED ESSAY\s*", "", b, flags=re.IGNORECASE).strip()
    return sections


def extract_total_score(correction_text: str):
    """Extrae el número de 'TOTAL: X/20'. Devuelve int o ''."""
    if not correction_text:
        return ""
    match = re.search(r"TOTAL\s*:\s*(\d{1,2})\s*/\s*20", correction_text, re.IGNORECASE)
    return int(match.group(1)) if match else ""


def read_enunciado(path: str) -> str:
    """Lee .md, .txt o .docx y devuelve el texto como string limpio."""
    ext = Path(path).suffix.lower()
    if ext == ".docx":
        from docx import Document
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    else:  # .md, .txt
        with open(path, "r", encoding="utf-8") as f:
            return f.read()


# ─────────────────────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────────────────────

STYLE_GRADE     = {"fill": PatternFill("solid", start_color="D6E4F0"), "font": Font(name="Calibri", size=11), "align": Alignment(wrap_text=True, vertical="top")}
STYLE_NOTA      = {"fill": PatternFill("solid", start_color="FFF2CC"), "font": Font(name="Calibri", size=12, bold=True), "align": Alignment(horizontal="center", vertical="center")}
STYLE_ERRORS    = {"fill": PatternFill("solid", start_color="FCE4D6"), "font": Font(name="Calibri", size=11), "align": Alignment(wrap_text=True, vertical="top")}
STYLE_COMMENTS  = {"fill": PatternFill("solid", start_color="EDE7F6"), "font": Font(name="Calibri", size=11), "align": Alignment(wrap_text=True, vertical="top")}
STYLE_CORRECTED = {"fill": PatternFill("solid", start_color="E8F5E9"), "font": Font(name="Calibri", size=11), "align": Alignment(wrap_text=True, vertical="top")}
STYLE_FULL      = {"fill": PatternFill("solid", start_color="E3F2FD"), "font": Font(name="Calibri", size=11), "align": Alignment(wrap_text=True, vertical="top")}
STYLE_FULL_CLEAN= {"fill": PatternFill("solid", start_color="F8F9FA"), "font": Font(name="Calibri", size=11), "align": Alignment(wrap_text=True, vertical="top")}


def _apply(cell, style: dict):
    cell.fill      = style["fill"]
    cell.font      = style["font"]
    cell.alignment = style["align"]


# ─────────────────────────────────────────────────────────────
# HOJA DE RESUMEN
# ─────────────────────────────────────────────────────────────

def build_summary_sheet(wb, ws_data, essay_type: str):
    SHEET_NAME = "Resumen"
    if SHEET_NAME in wb.sheetnames:
        del wb[SHEET_NAME]
    ws = wb.create_sheet(SHEET_NAME)

    header_font  = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill  = PatternFill("solid", start_color="2E75B6")
    center_align = Alignment(horizontal="center", vertical="center")
    left_align   = Alignment(horizontal="left",   vertical="center")
    score_fill   = PatternFill("solid", start_color="FFF2CC")
    score_font   = Font(name="Calibri", size=11, bold=True)
    alt_fill     = PatternFill("solid", start_color="F2F7FB")

    # Solo el título en A1, SIN fusionar celdas (la fusión daba error al
    # volver a ejecutar el programa sobre el mismo Excel).
    title_cell       = ws["A1"]
    title_cell.value = f"Resumen de correcciones — {essay_type.capitalize()}"
    title_cell.font  = Font(name="Calibri", size=13, bold=True, color="2E75B6")
    title_cell.alignment = left_align
    ws.row_dimensions[1].height = 24

    headers    = ["Alumno", "TOTAL /20", "Content /5", "Comm. Ach. /5", "Organization /5", "Language /5"]
    col_widths = [28, 13, 13, 16, 16, 13]
    for col_idx, (header, width) in enumerate(zip(headers, col_widths), start=1):
        cell           = ws.cell(row=2, column=col_idx, value=header)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = center_align
        ws.column_dimensions[cell.column_letter].width = width
    ws.row_dimensions[2].height = 18

    criteria_patterns = {
        "content":  re.compile(r"Content\s*:\s*(\d)/5",                     re.IGNORECASE),
        "comm":     re.compile(r"Communicative\s+Achievement\s*:\s*(\d)/5", re.IGNORECASE),
        "org":      re.compile(r"Organi[sz]ation\s*:\s*(\d)/5",             re.IGNORECASE),
        "language": re.compile(r"Language\s*:\s*(\d)/5",                    re.IGNORECASE),
    }

    data_row_idx   = 3
    scores_for_avg = []

    for row in ws_data.iter_rows(min_row=2):
        name_cell = row[COL_NOMBRE - 1]
        full_cell = row[COL_FULL - 1]
        name      = str(name_cell.value).strip() if name_cell.value else ""
        full_text = str(full_cell.value).strip() if full_cell.value else ""
        if not name and not full_text:
            continue

        total = extract_total_score(full_text)
        sub   = {k: (int(p.search(full_text).group(1)) if p.search(full_text) else "") for k, p in criteria_patterns.items()}

        is_alt = (data_row_idx % 2 == 0)
        for col_idx, value in enumerate([name, total, sub["content"], sub["comm"], sub["org"], sub["language"]], start=1):
            cell           = ws.cell(row=data_row_idx, column=col_idx, value=value)
            cell.font      = Font(name="Calibri", size=11, bold=(col_idx == 2))
            cell.alignment = left_align if col_idx == 1 else center_align
            if col_idx == 2:
                cell.fill = score_fill
                cell.font = score_font
            elif is_alt:
                cell.fill = alt_fill
        ws.row_dimensions[data_row_idx].height = 16

        if isinstance(total, int):
            scores_for_avg.append(total)
        data_row_idx += 1

    if scores_for_avg:
        avg_row       = data_row_idx + 1
        lbl           = ws.cell(row=avg_row, column=1, value="MEDIA DE CLASE")
        lbl.font      = Font(name="Calibri", size=11, bold=True)
        lbl.alignment = left_align
        val           = ws.cell(row=avg_row, column=2, value=round(sum(scores_for_avg) / len(scores_for_avg), 2))
        val.font      = Font(name="Calibri", size=12, bold=True, color="C00000")
        val.fill      = PatternFill("solid", start_color="FCE4D6")
        val.alignment = center_align

    wb.move_sheet(SHEET_NAME, offset=-len(wb.sheetnames) + 1)
    # Dejar activa la hoja de respuestas (no la de Resumen) al guardar
    wb.active = wb.sheetnames.index(ws_data.title)
    for sheet in wb.worksheets:
        sheet.sheet_view.tabSelected = (sheet.title == ws_data.title)
    print(f"  ✓ Hoja 'Resumen' creada/actualizada ({data_row_idx - 3} alumnos).")


# ─────────────────────────────────────────────────────────────
# LÓGICA EXCEL
# ─────────────────────────────────────────────────────────────

OUTPUT_COLS = [
    (COL_GRADE,     "Grade Summary",       45, STYLE_GRADE),
    (COL_NOTA,      "Nota /20",            10, STYLE_NOTA),
    (COL_ERRORS,    "Error Analysis",      55, STYLE_ERRORS),
    (COL_COMMENTS,  "Examiner Comments",   45, STYLE_COMMENTS),
    (COL_CORRECTED, "Corrected Essay",     70, STYLE_CORRECTED),
    (COL_FULL,      "Corrección completa", 80, STYLE_FULL),
    (COL_FULL_CLEAN, "Corrección limpia",  80, STYLE_FULL_CLEAN),
]

def _col_letter(n: int) -> str:
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def setup_headers(ws):
    for col_idx, header, width, style in OUTPUT_COLS:
        cell = ws.cell(row=1, column=col_idx)
        if cell.value is None:
            cell.value     = header
            cell.font      = Font(bold=True, name="Calibri", size=11)
            cell.fill      = style["fill"]
            cell.alignment = Alignment(horizontal="center")
            ws.column_dimensions[_col_letter(col_idx)].width = width


def write_row(ws, row_num: int, full_text: str):
    """Escribe el texto completo desglosado en F-L. Devuelve la nota."""
    sections = parse_sections(full_text)
    score    = extract_total_score(full_text)
    full_text_clean = (
        f"GRADE SUMMARY\n{sections['grade']}\n\n"
        f"ERROR ANALYSIS\n{sections['errors']}\n\n"
        f"EXAMINER COMMENTS\n{sections['comments']}\n\n"
        f"CORRECTED ESSAY\n{sections['corrected']}"
    )

    pairs = [
        (COL_GRADE,      sections["grade"],     STYLE_GRADE),
        (COL_NOTA,       score,                 STYLE_NOTA),
        (COL_ERRORS,     sections["errors"],    STYLE_ERRORS),
        (COL_COMMENTS,   sections["comments"],  STYLE_COMMENTS),
        (COL_CORRECTED,  sections["corrected"], STYLE_CORRECTED),
        (COL_FULL,       full_text,             STYLE_FULL),
        (COL_FULL_CLEAN, full_text_clean,       STYLE_FULL_CLEAN),
    ]
    for col, value, style in pairs:
        c = ws.cell(row=row_num, column=col, value=value)
        _apply(c, style)

    return score


SUMMARY_SHEET_NAME = "Resumen"


def get_data_sheet(wb):
    """
    Devuelve la hoja con las respuestas del formulario: la primera que no sea
    'Resumen'. (Antes se usaba wb.active, que podía ser la hoja Resumen.)
    """
    for sheet in wb.worksheets:
        if sheet.title != SUMMARY_SHEET_NAME:
            return sheet
    raise ValueError("El Excel no tiene ninguna hoja de respuestas.")


def process_workbook(wb, corrector: CambridgeCorrector, save_fn,
                     progress_callback=None, should_stop=None) -> dict:
    """
    Corrige todas las redacciones pendientes de un Workbook ya abierto.

    save_fn(wb)                 → guarda el libro (en disco o en Drive). Se llama
                                  tras cada alumno corregido, así nunca se pierde
                                  el trabajo hecho si algo falla a mitad.
    progress_callback(i, total, nombre, estado)
                                  estado: "correcting" | "skipped" |
                                  "done:<nota>" | "error:<mensaje>" | "summary"
    should_stop()               → si devuelve True, se para antes del siguiente alumno.

    Devuelve {"success", "skipped", "failed", "total", "stopped"}.
    """
    def notify(*args):
        if progress_callback:
            progress_callback(*args)

    ws = get_data_sheet(wb)
    setup_headers(ws)

    data_rows = []
    for row in ws.iter_rows(min_row=2):
        name_cell  = row[COL_NOMBRE - 1]
        essay_cell = row[COL_REDACCION - 1]
        full_cell  = row[COL_FULL - 1]
        if essay_cell.value and str(essay_cell.value).strip():
            data_rows.append((essay_cell.row, name_cell, essay_cell, full_cell))

    total = len(data_rows)
    print(f"Encontradas {total} redacciones para corregir.\n")

    success = 0
    skipped = 0
    failed  = []
    stopped = False

    for idx, (row_num, name_cell, essay_cell, full_cell) in enumerate(data_rows, start=1):
        student_name = str(name_cell.value).strip() if name_cell.value else f"Alumno_fila{row_num}"
        essay_text   = str(essay_cell.value).strip()

        # Saltar si K ya tiene corrección completa
        if full_cell.value and str(full_cell.value).strip():
            print(f"[{idx}/{total}] OMITIDO (ya corregido): {student_name}")
            # Rellenar F-J si faltan (reanudación parcial), SIN llamar a la API
            grade_cell = ws.cell(row=row_num, column=COL_GRADE)
            if not grade_cell.value:
                write_row(ws, row_num, str(full_cell.value))
                save_fn(wb)
            skipped += 1
            notify(idx, total, student_name, "skipped")
            continue

        if should_stop and should_stop():
            stopped = True
            print("\nCorrección detenida por el usuario.")
            break

        print(f"[{idx}/{total}] Corrigiendo: {student_name}...")
        notify(idx, total, student_name, "correcting")

        try:
            full_text = corrector.correct_essay(student_name, essay_text)
            score     = write_row(ws, row_num, full_text)
            save_fn(wb)
            success += 1
            print(f"       ✓ Guardado (F-L)  |  Nota: {score}/20")
            notify(idx, total, student_name, f"done:{score}")

        except Exception as e:
            print(f"       ✗ Error: {e}")
            failed.append(student_name)
            notify(idx, total, student_name, f"error:{e}")

        if idx < total and DELAY_BETWEEN_CALLS:
            print(f"       Esperando {DELAY_BETWEEN_CALLS}s...")
            time.sleep(DELAY_BETWEEN_CALLS)

    print("\nGenerando hoja de resumen...")
    notify(total, total, "", "summary")
    build_summary_sheet(wb, ws, corrector.essay_type)
    save_fn(wb)

    return {"success": success, "skipped": skipped, "failed": failed,
            "total": total, "stopped": stopped}


def process_excel(excel_path: str, corrector: CambridgeCorrector):
    """Versión de consola: trabaja sobre un .xlsx en disco, como siempre."""
    backup_dir  = Path(excel_path).parent.parent / "estudiantes_backup"
    backup_dir.mkdir(exist_ok=True)
    backup_path = backup_dir / Path(excel_path).name.replace(".xlsx", "_BACKUP.xlsx")
    shutil.copy2(excel_path, str(backup_path))
    print(f"✓ Copia de seguridad creada: {backup_path}\n")

    wb = load_workbook(excel_path)
    stats = process_workbook(wb, corrector, save_fn=lambda w: w.save(excel_path))

    print(f"\n{'='*55}")
    print(f"  Proceso finalizado.")
    print(f"  ✓ Corregidas:  {stats['success']}/{stats['total']}")
    if stats["skipped"]:
        print(f"  Omitidas (ya tenían corrección): {stats['skipped']}")
    if stats["failed"]:
        print(f"  ✗ Fallidas:   {', '.join(stats['failed'])}")
    print(f"  Archivo guardado: {excel_path}")
    print(f"{'='*55}")


# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────

def main():
    print("=" * 55)
    print("  Cambridge Writing Corrector  (Excel edition)")
    print("=" * 55)

    script_dir = Path(__file__).parent.absolute()

    NIVELES = {"1": "B1", "2": "B2", "3": "C1", "4": "C2"}
    print("\nNivel de las redacciones a corregir:")
    for num, lvl in NIVELES.items():
        print(f"  {num}) {lvl}")
    level = None
    while level is None:
        opcion = input("Elige un número (1-4): ").strip()
        if opcion in NIVELES:
            level = NIVELES[opcion]
        else:
            print("  Opción no válida. Introduce un número del 1 al 4.")

    essay_type = input("\nTipo de redacción (ej: essay, report, proposal, letter): ").strip().lower()
    if not essay_type:
        print("Error: debes indicar el tipo de redacción.")
        sys.exit(1)

    base_dir     = script_dir / "redacciones" / level
    rubric_path  = base_dir / f"{level.lower()}_rubrica.md"
    example_path = base_dir / f"{level.lower()}_{essay_type}.md"

    missing = []
    if not rubric_path.exists():  missing.append(str(rubric_path))
    if not example_path.exists(): missing.append(str(example_path))
    if missing:
        print("\nError: no se encontraron los siguientes archivos:")
        for m in missing:
            print(f"  ✗ {m}")
        sys.exit(1)

    excel_nombre = input("\nNombre del archivo Excel: ").strip()
    if not excel_nombre.lower().endswith(".xlsx"):
        excel_nombre += ".xlsx"
    excel_path = script_dir / "estudiantes" / excel_nombre
    if not excel_path.exists():
        print(f"\nError: No se encontró '{excel_nombre}' en la carpeta 'estudiantes'.")
        print(f"Ruta buscada: {excel_path}")
        sys.exit(1)
    
    # ── Enunciado (opcional) ──────────────────────────────────
    enunciado_nombre = input("\nNombre del .docx con el enunciado o Enter para omitir: ").strip()
    enunciado_text = ""
    if enunciado_nombre:
        if not enunciado_nombre.lower().endswith(".docx"):
            enunciado_nombre += ".docx"
        enunciado_path = script_dir / "enunciados" / enunciado_nombre
        if not enunciado_path.exists():
            print(f"\nError: No se encontró '{enunciado_nombre}' en la carpeta 'enunciados'.")
            print(f"Ruta buscada: {enunciado_path}")
            sys.exit(1)
        enunciado_text = read_enunciado(str(enunciado_path))
        print(f"✓ Enunciado cargado ({len(enunciado_text.split())} palabras).\n")
    else:
        print("  (Sin enunciado — evaluación sin referencia al task prompt)\n")

    print(f"\nCargando archivos de referencia para '{essay_type.capitalize()}' (nivel {level})...")
    rubric_text  = rubric_path.read_text(encoding="utf-8")
    example_text = example_path.read_text(encoding="utf-8")
    print("✓ Archivos de referencia cargados.\n")

    try:
        corrector = CambridgeCorrector(level, essay_type, rubric_text, example_text, enunciado_text)
    except RuntimeError as e:
        print(f"ERROR: {e}")
        sys.exit(1)
    process_excel(str(excel_path), corrector)


if __name__ == "__main__":
    main()