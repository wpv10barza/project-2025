from __future__ import annotations
from pathlib import Path
import json
from datetime import datetime
from typing import Tuple, Dict, Optional
from docx import Document
from pypdf import PdfReader

from config import DEFAULT_DATA_PILOTO_DIR, DEFAULT_PROMPT_TRABAJO_PATH
from .db_questions import load_questions

def _read_text_file(p: Path) -> str:
    return p.read_text(encoding="utf-8", errors="ignore").strip()

def _read_docx(p: Path) -> str:
    doc = Document(str(p))
    parts = [para.text for para in doc.paragraphs if para.text]
    return "\n".join(parts).strip()

def _read_pdf(p: Path) -> str:
    reader = PdfReader(str(p))
    texts = []
    for page in reader.pages:
        t = page.extract_text() or ""
        texts.append(t)
    return "\n".join(texts).strip()

def read_transcription(p: Path) -> str:
    if p.suffix.lower() == ".txt":
        return _read_text_file(p)
    if p.suffix.lower() == ".docx":
        return _read_docx(p)
    if p.suffix.lower() == ".pdf":
        return _read_pdf(p)
    raise ValueError(f"Extensión no soportada para transcripción: {p.suffix}")

def build_prompt(evaluacion_id: int, transcripcion_path: Path, objetivo_path: Path, db_config_path: Path) -> Tuple[str, Dict]:
    objetivo = _read_text_file(objetivo_path)
    transcripcion = read_transcription(transcripcion_path)

    df = load_questions(evaluacion_id=evaluacion_id, db_config_path=db_config_path)

    preguntas_lines = []
    for _, r in df.iterrows():
        preguntas_lines.append(f"{r['PREGUNTA_CODIGO']}\t{r.get('PREGUNTA_DESCRIPCION','')}\t{r.get('PREGUNTA_SITUACION_DESEADA','')}".strip())

    preguntas_block = "\n".join(preguntas_lines)

    prompt = f"""OBJETIVO FINAL
{objetivo}

### PREGUNTAS (desde SQL, EvaluacionId={evaluacion_id})
PREGUNTA_CODIGO\\tPREGUNTA_DESCRIPCION\\tPREGUNTA_SITUACION_DESEADA
{preguntas_block}

### TRANSCRIPCIÓN (fuente de evidencia literal)
{transcripcion}

### FORMATO DE SALIDA OBLIGATORIO (TSV) copiable
Devuelve SOLO un bloque TSV con encabezados EXACTOS:
PREGUNTA_CODIGO\\tCONVERSACION_ENTREVISTADOR_1\\tCONVERSACION_INFORMANTE_1\\tPREGUNTA_DESCRIPCION\\tPREGUNTA_SITUACION_DESEADA\\tANÁLISIS_DE_COBERTURA\\tCONCLUCIÓN_FINAL(LEIDO)\\tJUSTIFICACIÓN\\tRESPUESTA_COVERTURA(LEIDO)
"""
    meta = {
        "evaluacion_id": evaluacion_id,
        "transcripcion": str(transcripcion_path),
        "objetivo": str(objetivo_path),
        "preguntas": int(len(df)),
        "generated_at": datetime.now().isoformat(timespec="seconds"),
    }
    return prompt, meta

def write_prompt_files(
    evaluacion_id: int,
    db_config_path: Path,
    data_piloto_dir: Path = DEFAULT_DATA_PILOTO_DIR,
    objetivo_path: Optional[Path] = None,
    transcripcion_override: Optional[Path] = None,
    prompt_trabajo_path: Path = DEFAULT_PROMPT_TRABAJO_PATH,
) -> Dict[str, str]:
    dp = Path(data_piloto_dir)
    objetivo = objetivo_path or (dp / "OBJETIVO_FINAL.txt")

    if transcripcion_override:
        trans = Path(transcripcion_override)
    else:
        # orden preferido: TXT > DOCX > PDF
        for name in ["Transcripcion.txt", "Transcripcion.docx", "Transcripcion.pdf"]:
            cand = dp / name
            if cand.exists():
                trans = cand
                break
        else:
            raise FileNotFoundError(f"No se encontró transcripción en {dp}. Usa --transcripcion.")

    if not objetivo.exists():
        raise FileNotFoundError(f"Falta OBJETIVO_FINAL.txt en: {objetivo}")

    out_prompt_dir = dp / "01.02 Promp"
    out_prompt_dir.mkdir(parents=True, exist_ok=True)

    prompt, meta = build_prompt(
        evaluacion_id=evaluacion_id,
        transcripcion_path=trans,
        objetivo_path=objetivo,
        db_config_path=db_config_path,
    )

    prompt_path = out_prompt_dir / "prompt_completo.txt"
    meta_path = out_prompt_dir / "prompt_metadata.json"

    prompt_path.write_text(prompt, encoding="utf-8")
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    # copia adicional solicitada
    try:
        prompt_trabajo_path.parent.mkdir(parents=True, exist_ok=True)
        prompt_trabajo_path.write_text(prompt, encoding="utf-8")
    except Exception:
        pass

    print(f"[OK] Prompt generado: {prompt_path}")
    print(f"[OK] Metadata: {meta_path}")
    if prompt_trabajo_path.exists():
        print(f"[OK] Copia prompt_trabajo: {prompt_trabajo_path}")

    return {
        "prompt_path": str(prompt_path),
        "meta_path": str(meta_path),
        "prompt_trabajo": str(prompt_trabajo_path),
    }
