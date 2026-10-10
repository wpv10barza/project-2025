# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from datetime import datetime
import json

from config import (
    DEFAULT_OBJETIVO_TXT, DEFAULT_TRANSCRIPCION_TXT, DEFAULT_TRANSCRIPCION_PDF,
    PROMPT_TXT_OUT, PROMPT_META_OUT, DEFAULT_PROMPT_TRABAJO,
    DEFAULT_DB_CONFIG
)
from src.repo import resolve_data_piloto_dir
from src.transcription import read_transcription
from src.db import read_db_config, connect
from src.db_questions import load_questions_df

def _read_txt(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="ignore").strip()

def build_prompt(
    evaluacion_id: int,
    data_piloto_dir: Path,
    objetivo_path: Path,
    transcripcion_path: Path,
    db_config_path: Path,
) -> tuple[str, dict]:
    if not objetivo_path.exists():
        raise FileNotFoundError(
            f"Falta OBJETIVO_FINAL.txt en DATA PILOTO:\n - {objetivo_path}\n"
            "Solución: copia OBJETIVO_FINAL.txt al DATA PILOTO."
        )
    if not transcripcion_path.exists():
        raise FileNotFoundError(
            f"Falta Transcripcion (txt/pdf/docx) en DATA PILOTO:\n - {transcripcion_path}\n"
            "Solución: copia Transcripcion.* al DATA PILOTO."
        )

    objetivo = _read_txt(objetivo_path)
    trans = read_transcription(transcripcion_path)

    cfg = read_db_config(db_config_path)
    with connect(cfg) as conn:
        df_q = load_questions_df(conn, evaluacion_id=int(evaluacion_id))

    preguntas_lines = []
    for r in df_q.itertuples(index=False):
        codigo = str(r.PREGUNTA_CODIGO).strip()
        desc = str(getattr(r, "PREGUNTA_DESCRIPCION", "") or "").strip()
        sit  = str(getattr(r, "PREGUNTA_SITUACION_DESEADA", "") or "").strip()
        preguntas_lines.append(f"[{codigo}] {desc}\nSITUACION_DESEADA: {sit}".strip())

    preguntas_txt = "\n\n".join(preguntas_lines)

    prompt = f"""### INSTRUCCIONES (OBJETIVO FINAL)
{objetivo}

### PREGUNTAS DESDE SQL (VW_PREGUNTA_POR_EVALUACION / EvaluacionId={evaluacion_id})
{preguntas_txt}

### TRANSCRIPCION COMPLETA (Transcripcion)
{trans}

### FORMATO DE SALIDA OBLIGATORIO (TSV) copiable
Devuelve SOLO un bloque TSV con encabezados EXACTOS:
PREGUNTA_CODIGO\tCONVERSACION_ENTREVISTADOR_1\tCONVERSACION_INFORMANTE_1\tPREGUNTA_DESCRIPCION\tPREGUNTA_SITUACION_DESEADA\tANÁLISIS_DE_COBERTURA\tCONCLUCIÓN_FINAL(LEIDO)\tJUSTIFICACIÓN\tRESPUESTA_COVERTURA(LEIDO)

REGLAS:
- El TSV debe ser pegable directo a Excel/Sheets.
- NO agregues texto antes ni después del bloque TSV.
- Si necesitas saltos de línea dentro de una celda, usa "\\n" (dos caracteres) dentro de la celda.
"""
    meta = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "evaluacion_id": int(evaluacion_id),
        "data_piloto_dir": str(data_piloto_dir),
        "files": {
            "objetivo": str(objetivo_path),
            "transcripcion": str(transcripcion_path),
            "db_config": str(db_config_path),
        },
        "output_prompt": str(PROMPT_TXT_OUT),
    }
    return prompt, meta

def write_prompt_files(
    evaluacion_id: int,
    data_piloto_dir: str | None = None,
    objetivo_path: str | None = None,
    transcripcion_override: str | None = None,
    db_config_path: str | None = None,
    prompt_trabajo_path: str | None = None,
) -> Path:
    dp = resolve_data_piloto_dir(data_piloto_dir, fallback=DEFAULT_OBJETIVO_TXT.parent)

    obj = Path(objetivo_path).expanduser().resolve() if objetivo_path else (dp / DEFAULT_OBJETIVO_TXT.name)

    if transcripcion_override:
        tr = Path(transcripcion_override).expanduser().resolve()
    else:
        tr = (dp / DEFAULT_TRANSCRIPCION_TXT.name) if (dp / DEFAULT_TRANSCRIPCION_TXT.name).exists() else (dp / DEFAULT_TRANSCRIPCION_PDF.name)

    cfg_path = Path(db_config_path).expanduser().resolve() if db_config_path else (dp / DEFAULT_DB_CONFIG.name)

    prompt, meta = build_prompt(
        evaluacion_id=evaluacion_id,
        data_piloto_dir=dp,
        objetivo_path=obj,
        transcripcion_path=tr,
        db_config_path=cfg_path,
    )

    PROMPT_TXT_OUT.write_text(prompt, encoding="utf-8")
    PROMPT_META_OUT.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    pt = Path(prompt_trabajo_path).expanduser().resolve() if prompt_trabajo_path else DEFAULT_PROMPT_TRABAJO.resolve()
    pt.parent.mkdir(parents=True, exist_ok=True)
    pt.write_text(prompt, encoding="utf-8")

    return PROMPT_TXT_OUT
