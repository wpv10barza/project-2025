# -*- coding: utf-8 -*-
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
import shutil

from config import (
    DATA_PILOTO_DIR, OBJETIVO_TXT,
    TRANSCRIPCION_TXT, TRANSCRIPCION_PDF, TRANSCRIPCIONES_DIR,
    PROMP_DIR, RESULT_DIR,
    PROMPT_TXT_OUT, PROMPT_META_OUT,
    PROMP_COPY_DIR,
    PROMPT_TRABAJO_OUT, PROMPT_TRABAJO_META_OUT,
    DB_CONFIG_FILE, DB_CONFIG_FILE_FALLBACK,
)
from src.db import read_db_config, connect
from src.db_questions import fetch_questions
from src.io_utils import read_text_file, read_many_transcriptions, read_pdf_text

DEFAULT_OBJETIVO = """INSTRUCCIONES (OBJETIVO FINAL)
Construye una Matriz de Sistematización en TSV.
Reglas:
- SOLO evidencia literal (citas textuales) extraída de la(s) transcripción(es).
- No inventes información ni completes con supuestos.
- Si no hay evidencia, deja la celda vacía o usa: "No se encontró evidencia literal ...".
- Mantén trazabilidad: incluye [Fuente: <archivo>, p.X] si existe paginado o referencia.
"""

def _ensure_file(path: Path, default_text: str) -> None:
    if not path.exists():
        path.write_text(default_text, encoding="utf-8")

def _gather_transcription_paths(data_dir: Path, override: Path | None) -> list[Path]:
    if override is not None:
        return [override]
    paths: list[Path] = []
    # carpeta con varios archivos
    tdir = TRANSCRIPCIONES_DIR if TRANSCRIPCIONES_DIR.is_absolute() else data_dir / TRANSCRIPCIONES_DIR.name
    if tdir.exists() and tdir.is_dir():
        for p in sorted(tdir.glob("*.txt")) + sorted(tdir.glob("*.pdf")):
            paths.append(p)
    # single file default
    t_txt = TRANSCRIPCION_TXT if TRANSCRIPCION_TXT.is_absolute() else data_dir / TRANSCRIPCION_TXT.name
    t_pdf = TRANSCRIPCION_PDF if TRANSCRIPCION_PDF.is_absolute() else data_dir / TRANSCRIPCION_PDF.name
    if t_txt.exists():
        paths.append(t_txt)
    elif t_pdf.exists():
        paths.append(t_pdf)
    return paths

def build_prompt(evaluacion_id: int, data_piloto_dir: Path, objetivo_path: Path | None, transcripcion_override: Path | None, db_config_path: Path | None) -> tuple[str, dict]:
    # Objetivo
    obj_path = objetivo_path or (data_piloto_dir / OBJETIVO_TXT.name)
    _ensure_file(obj_path, DEFAULT_OBJETIVO)
    objetivo = read_text_file(obj_path).strip()

    # DB config
    cfg_path = Path(db_config_path) if db_config_path else (DB_CONFIG_FILE if DB_CONFIG_FILE.exists() else DB_CONFIG_FILE_FALLBACK)
    cfg = read_db_config(cfg_path)

    # Preguntas desde SQL
    with connect(cfg) as conn:
        df_q = fetch_questions(conn, evaluacion_id)
    # Formato guía en TSV dentro del prompt (sin PDF)
    guia_lines = ["PREGUNTA_CODIGO\tPREGUNTA_DESCRIPCION\tPREGUNTA_SITUACION_DESEADA"]
    for _, r in df_q.iterrows():
        cod = str(r.get("PREGUNTA_CODIGO","")).strip()
        desc = str(r.get("PREGUNTA_DESCRIPCION","")).strip().replace("\t"," ").replace("\n"," ")
        sit  = str(r.get("PREGUNTA_SITUACION_DESEADA","")).strip().replace("\t"," ").replace("\n"," ")
        guia_lines.append(f"{cod}\t{desc}\t{sit}")
    guia_tsv = "\n".join(guia_lines)

    # Transcripción(es)
    trans_paths = _gather_transcription_paths(data_piloto_dir, transcripcion_override)
    trans = read_many_transcriptions(trans_paths).strip()
    if not trans:
        trans = "NO_TRANSCRIPCION_INCLUIDA"

    # Prompt final
    prompt = f"""### INSTRUCCIONES (OBJETIVO FINAL)
{objetivo}

### TRANSCRIPCIONES (texto completo)
{trans}

### GUIA (desde SQL: VW_PREGUNTA_POR_EVALUACION, EvaluacionId={evaluacion_id})
(Se entrega como TSV: PREGUNTA_CODIGO, PREGUNTA_DESCRIPCION, PREGUNTA_SITUACION_DESEADA)
{guia_tsv}

### FORMATO DE SALIDA OBLIGATORIO (TSV) copiable
Devuelve SOLO un bloque TSV con encabezados EXACTOS:
PREGUNTA_CODIGO\tCONVERSACION_ENTREVISTADOR_1\tCONVERSACION_INFORMANTE_1\tPREGUNTA_DESCRIPCION\tPREGUNTA_SITUACION_DESEADA\tANÁLISIS_DE_COBERTURA\tCONCLUCIÓN_FINAL(LEIDO)\tJUSTIFICACIÓN\tRESPUESTA_COVERTURA(LEIDO)

Reglas de formato:
- No agregues texto fuera del bloque TSV.
- No uses comillas dobles salvo que sea estrictamente necesario.
- Si necesitas saltos de línea dentro de una celda, usa "\\n" dentro de la celda (dos caracteres: barra invertida + n).
"""
    meta = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "evaluacion_id": evaluacion_id,
        "data_piloto_dir": str(data_piloto_dir),
        "files": {
            "objetivo": str(obj_path),
            "transcripciones": [str(p) for p in trans_paths],
            "db_config": str(cfg_path),
        },
        "outputs": {
            "prompt_txt": str(PROMPT_TXT_OUT),
            "prompt_meta": str(PROMPT_META_OUT),
        },
    }
    return prompt, meta

def write_prompt_files(evaluacion_id: int = 16, data_piloto_dir: str | None = None, objetivo_path: str | None = None, transcripcion_path: str | None = None, db_config_path: str | None = None) -> Path:
    dp = Path(data_piloto_dir) if data_piloto_dir else DATA_PILOTO_DIR
    dp.mkdir(parents=True, exist_ok=True)
    PROMP_DIR.mkdir(parents=True, exist_ok=True)
    RESULT_DIR.mkdir(parents=True, exist_ok=True)

    obj = Path(objetivo_path) if objetivo_path else None
    tr  = Path(transcripcion_path) if transcripcion_path else None

    prompt, meta = build_prompt(evaluacion_id=evaluacion_id, data_piloto_dir=dp, objetivo_path=obj, transcripcion_override=tr, db_config_path=db_config_path)
    PROMPT_TXT_OUT.write_text(prompt, encoding="utf-8")
    PROMPT_META_OUT.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    # ✅ Ajuste fijo solicitado: además de generar el prompt_completo.txt en el DATA PILOTO,
    # guarda una copia lista para el automatizador Selenium (prompt_trabajo.txt).
    try:
        PROMPT_TRABAJO_OUT.parent.mkdir(parents=True, exist_ok=True)
        PROMPT_TRABAJO_OUT.write_text(prompt, encoding="utf-8")
        PROMPT_TRABAJO_META_OUT.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

    # Copia a G:\Mi unidad\PROMP
    try:
        PROMP_COPY_DIR.mkdir(parents=True, exist_ok=True)
        shutil.copy2(PROMPT_TXT_OUT, PROMP_COPY_DIR / PROMPT_TXT_OUT.name)
        shutil.copy2(PROMPT_META_OUT, PROMP_COPY_DIR / PROMPT_META_OUT.name)
    except Exception:
        # No reviente si no existe unidad G: en la máquina
        pass

    return PROMPT_TXT_OUT
