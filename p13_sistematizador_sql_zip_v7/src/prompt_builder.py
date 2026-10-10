from __future__ import annotations
from pathlib import Path
from datetime import datetime
import json

from config import (
    GUIA_PDF, TRANSCRIPCION_PDF, OBJETIVO_TXT,
    PROMPT_TXT_OUT, PROMPT_META_OUT, PROMP_EXTRA_DIR,
    DECLARACION_PARTICIPANTES,
    USE_SQL_PREGUNTAS, NO_DOCUMENTACION_PDF_INCLUIDA, EVALUACION_ID, DB_CONFIG_FILE, SQL_DRIVER,
    DATA_PILOTO_DIR
)
from src.utils.pdf_utils import read_pdf_text
from src.db_questions import get_connection, load_df_preguntas

def _read_txt(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="ignore").strip()

def build_prompt() -> tuple[str, dict]:
    missing_files = []
    objetivo = ""
    trans = ""

    if OBJETIVO_TXT.exists():
        objetivo = _read_txt(OBJETIVO_TXT)
    else:
        missing_files.append(str(OBJETIVO_TXT))
        objetivo = "[MISSING OBJETIVO_FINAL.txt] (Coloca OBJETIVO_FINAL.txt en DATA_PILOTO_DIR)"

    if TRANSCRIPCION_TXT.exists():
        trans = _read_txt(TRANSCRIPCION_TXT)
    elif TRANSCRIPCION_PDF.exists():
        # fallback histórico
        trans = read_pdf_text(TRANSCRIPCION_PDF)
    else:
        missing_files.append(str(TRANSCRIPCION_TXT))
        trans = "[MISSING Transcripcion.txt] (Coloca Transcripcion.txt en DATA_PILOTO_DIR)"

    participantes = DECLARACION_PARTICIPANTES.strip()
    if participantes:
        objetivo = objetivo + "\n\nDECLARACIÓN DE PARTICIPANTES (config.py)\n" + participantes

    guia_section_title = ""
    guia_section_body = ""
    guia_count = 0

    if USE_SQL_PREGUNTAS:
        try:
            with get_connection(DB_CONFIG_FILE, SQL_DRIVER) as conn:
                df_preg = load_df_preguntas(conn, EVALUACION_ID)
            guia_count = len(df_preg)
            guia_section_title = "### GUIA (desde SQL: VW_PREGUNTA_POR_EVALUACION)"
            lines = []
            for r in df_preg.itertuples(index=False):
                lines.append(
                    f"{r.PREGUNTA_CODIGO}\tPREGUNTA: {r.PREGUNTA_DESCRIPCION}\tSITUACION_DESEADA: {r.PREGUNTA_SITUACION_DESEADA}"
                )
            guia_section_body = "\n".join(lines)
        except Exception as e:
            guia_section_title = "### GUIA (SQL) [ERROR DE CONEXION]"
            guia_section_body = f"[NO SE PUDO LEER SQL: {e}]"
    else:
        if GUIA_PDF.exists():
            guia_section_title = "### GUIA COMPLETA (PREGUNTA_CODIGO.pdf)"
            guia_section_body = read_pdf_text(GUIA_PDF)
        else:
            missing_files.append(str(GUIA_PDF))
            guia_section_title = "### GUIA COMPLETA (PREGUNTA_CODIGO.pdf)"
            guia_section_body = "[MISSING PREGUNTA_CODIGO.pdf]"

    marcador = "NO_DOCUMENTACION_PDF_INCLUIDA" if (NO_DOCUMENTACION_PDF_INCLUIDA and USE_SQL_PREGUNTAS) else "DOCUMENTACION_PDF_INCLUIDA"
    if missing_files:
        marcador = marcador + " | INSUMOS_FALTANTES"

    prompt = f"""### BANDERA
{marcador}

### DATA_PILOTO_DIR
{DATA_PILOTO_DIR}

### PARAMETROS PILOTO
- EVALUACION_ID: {EVALUACION_ID}
- GUIA_FUENTE: {'SQL' if USE_SQL_PREGUNTAS else 'PDF'}
- N_PREGUNTAS_SQL: {guia_count}

### INSTRUCCIONES (OBJETIVO FINAL)
{objetivo}

{guia_section_title}
(Se entrega en formato TSV: PREGUNTA_CODIGO \t PREGUNTA \t SITUACION_DESEADA)
{guia_section_body}

### TRANSCRIPCION COMPLETA (Transcripcion.pdf)
{trans}

### REGLAS CRITICAS
- Tu fuente de preguntas es la sección de GUIA. NO inventes preguntas fuera de esa guía.
- Para cada PREGUNTA_CODIGO:
  - CONVERSACION_ENTREVISTADOR_1: pega solo lo dicho por entrevistadores.
  - CONVERSACION_INFORMANTE_1: pega solo lo dicho por informantes.
- Si no hay evidencia, deja la celda vacía.
- Mantén literalidad. Si recortas, no cambies el sentido.
- Saltos de línea dentro de celda: usa "\\n".

### FORMATO DE SALIDA OBLIGATORIO (TSV)
Devuelve SOLO un bloque TSV (tab-separated) dentro de un codeblock ```tsv ... ``` con encabezados EXACTOS:
PREGUNTA_CODIGO\tCONVERSACION_ENTREVISTADOR_1\tCONVERSACION_INFORMANTE_1\tPREGUNTA_DESCRIPCION\tPREGUNTA_SITUACION_DESEADA\tANÁLISIS_DE_COBERTURA\tCONCLUCIÓN_FINAL(LEIDO)\tJUSTIFICACIÓN\tRESPUESTA_COVERTURA(LEIDO)
"""

    meta = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "mode": "sql" if USE_SQL_PREGUNTAS else "pdf",
        "evaluacion_id": EVALUACION_ID,
        "data_piloto_dir": str(DATA_PILOTO_DIR),
        "missing_files": missing_files,
        "n_preguntas_sql": guia_count,
        "files": {
            "objetivo": str(OBJETIVO_TXT),
            "guia_pdf": str(GUIA_PDF),
            "transcripcion_txt": str(TRANSCRIPCION_TXT),
            "transcripcion_pdf": str(TRANSCRIPCION_PDF),
            "db_config": str(DB_CONFIG_FILE),
        },
        "output_prompt": str(PROMPT_TXT_OUT),
    }
    return prompt, meta

def write_prompt_files() -> Path:
    prompt, meta = build_prompt()
    PROMPT_TXT_OUT.write_text(prompt, encoding="utf-8")
    PROMPT_META_OUT.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    # Copia adicional (no bloquea si no existe la unidad/ruta)
    try:
        if PROMP_EXTRA_DIR and str(PROMP_EXTRA_DIR).strip():
            PROMP_EXTRA_DIR.mkdir(parents=True, exist_ok=True)
            (PROMP_EXTRA_DIR / PROMPT_TXT_OUT.name).write_text(prompt, encoding="utf-8")
            (PROMP_EXTRA_DIR / PROMPT_META_OUT.name).write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    if meta.get("missing_files"):
        print("[WARN] Faltan insumos; el prompt se generó con placeholders. Revisa prompt_metadata.json -> missing_files.")
    return PROMPT_TXT_OUT
