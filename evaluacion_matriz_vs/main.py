from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from tqdm import tqdm

import config_local as cfg
from utils.pdf_utils import read_pdf_text, find_first_pdf_by_keywords
from utils.parse_transcript import parse_turns, turns_to_structured_tsv
from utils.parse_questions import parse_questions, merge_question_lists
from utils.prompt_builder import build_prompt
from utils.api_client import get_client, OpenAIConfig, chat_completion_text
from utils.matrix_builder import MatrixSpec, parse_tsv_row, rows_to_dataframe, save_outputs, save_rows_tsv


def _list_pdfs(folder: Path) -> str:
    pdfs = sorted(folder.glob("*.pdf"))
    if not pdfs:
        return "(no hay PDFs en la carpeta)"
    return "\n".join([f"- {p.name}" for p in pdfs])


def _load_objective_final(path: Path) -> str:
    if not path.exists():
        # Creamos un archivo plantilla para que el usuario pegue el texto
        path.write_text(
            "PEGA AQUÍ TU TEXTO COMPLETO DEL OBJETIVO FINAL (sin resumir)\n",
            encoding="utf-8",
        )
        raise RuntimeError(
            f"No existe {path}. Se creó una plantilla.\n"
            "Abre ese archivo, pega tu OBJETIVO FINAL completo (sin resumir) y vuelve a ejecutar."
        )

    text = path.read_text(encoding="utf-8").strip()
    if not text or "PEGA AQUÍ" in text.upper():
        raise RuntimeError(
            f"El archivo {path} está vacío o tiene placeholder.\n"
            "Pega tu OBJETIVO FINAL completo (sin resumir) y vuelve a ejecutar."
        )
    return text


def _pick_files() -> tuple[Path, list[Path]]:
    base = cfg.BASE_INPUT_DIR
    if not base.exists():
        raise RuntimeError(f"BASE_INPUT_DIR no existe: {base}")

    transcript_pdf = find_first_pdf_by_keywords(base, cfg.TRANSCRIPT_KEYWORDS)
    if not transcript_pdf:
        raise RuntimeError(
            "No se encontró el PDF de transcripción.\n"
            f"Busqué keywords: {cfg.TRANSCRIPT_KEYWORDS}\n"
            f"Carpeta: {base}\n\nPDFs disponibles:\n{_list_pdfs(base)}"
        )

    guide_pdfs: list[Path] = []
    g1 = find_first_pdf_by_keywords(base, cfg.GUIDE_KEYWORDS_PRIMARY)
    if g1:
        guide_pdfs.append(g1)
    g2 = find_first_pdf_by_keywords(base, cfg.GUIDE_KEYWORDS_SECONDARY)
    if g2 and g2 not in guide_pdfs:
        guide_pdfs.append(g2)

    if not guide_pdfs:
        raise RuntimeError(
            "No se encontró ningún PDF de guía (preguntas).\n"
            f"Busqué keywords primary: {cfg.GUIDE_KEYWORDS_PRIMARY}\n"
            f"Busqué keywords secondary: {cfg.GUIDE_KEYWORDS_SECONDARY}\n"
            f"Carpeta: {base}\n\nPDFs disponibles:\n{_list_pdfs(base)}"
        )

    return transcript_pdf, guide_pdfs


def _sanitize_single_tsv_line(s: str) -> str:
    # Asegura UNA línea (por compatibilidad con TSV/CSV).
    # Reemplaza saltos de línea por separador.
    s = s.strip()
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    if "\n" in s:
        s = " | ".join([x.strip() for x in s.split("\n") if x.strip()])
    return s


def main():
    load_dotenv()  # si usas .env local

    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=0, help="Procesa solo N preguntas (0=todo)")
    parser.add_argument("--codes", type=str, default="", help="Lista de códigos separados por coma (ej: 1.1.1,1.2.1)")
    parser.add_argument("--dry-run", action="store_true", help="No llama a la API; solo genera prompts")
    args = parser.parse_args()

    run_id = datetime.now().strftime("run_%Y%m%d_%H%M%S")
    run_out = cfg.OUTPUT_DIR / run_id
    prompts_out = run_out / "prompts"
    run_out.mkdir(parents=True, exist_ok=True)
    prompts_out.mkdir(parents=True, exist_ok=True)

    transcript_pdf, guide_pdfs = _pick_files()
    print(f"Transcripción: {transcript_pdf}")
    print("Guías:")
    for p in guide_pdfs:
        print(f"- {p}")

    # 1) Leer PDFs
    transcript_text = read_pdf_text(transcript_pdf)
    guide_texts = [read_pdf_text(p) for p in guide_pdfs]

    # 2) Parsear transcripción
    turns = parse_turns(transcript_text)
    if not turns:
        preview = transcript_text[:2000]
        raise RuntimeError(
            "No se detectaron turnos en la transcripción.\n"
            "Revisa si tu PDF tiene texto seleccionable o si el formato L### E1 ... está presente.\n\n"
            f"Preview (2000 chars):\n{preview}"
        )

    structured_tsv = turns_to_structured_tsv(turns, cfg.SPEAKER_ROLES)
    if cfg.MAX_TRANSCRIPT_CHARS and len(structured_tsv) > cfg.MAX_TRANSCRIPT_CHARS:
        structured_tsv = structured_tsv[: cfg.MAX_TRANSCRIPT_CHARS]
        print(f"[WARN] Transcripción truncada a {cfg.MAX_TRANSCRIPT_CHARS} caracteres para el prompt.")

    # 3) Cargar OBJETIVO FINAL
    objective_final = _load_objective_final(cfg.OBJECTIVE_FINAL_PATH)

    # 4) Parsear preguntas
    parsed_lists = [parse_questions(t) for t in guide_texts]
    questions = merge_question_lists(parsed_lists)

    print(f"✅ Preguntas extraídas: {len(questions)}")
    if len(questions) == 0:
        raise RuntimeError(
            "No se extrajeron preguntas.\n"
            "Sugerencia: revisa el texto extraído del PDF guía; quizá es escaneado y necesita OCR.\n"
            f"Preview guía (2000 chars):\n{guide_texts[0][:2000]}"
        )

    # Filtrar por códigos si aplica
    if args.codes.strip():
        wanted = {c.strip() for c in args.codes.split(",") if c.strip()}
        questions = [q for q in questions if q.get("PREGUNTA_CODIGO") in wanted]
        print(f"Filtrado por códigos: {len(questions)}")

    if args.limit and args.limit > 0:
        questions = questions[: args.limit]
        print(f"Limit: procesando {len(questions)}")

    # 5) Preparar MatrixSpec
    spec = MatrixSpec(
        interviewer_speakers=cfg.INTERVIEWER_SPEAKERS,
        informant_speakers=cfg.INFORMANT_SPEAKERS,
    )
    columns = spec.columns

    # 6) Cliente OpenAI
    client = None
    if not args.dry_run:
        client = get_client()
        oaicfg = OpenAIConfig(
            model_name=cfg.MODEL_NAME,
            temperature=cfg.TEMPERATURE,
            max_tokens=cfg.MAX_TOKENS,
            max_retries=cfg.MAX_RETRIES,
            backoff_base_seconds=cfg.RETRY_BACKOFF_BASE_SECONDS,
        )

    # 7) Ejecutar por pregunta
    rows_dicts = []
    rows_lines = []

    for q in tqdm(questions, desc="Procesando preguntas"):
        code = q.get("PREGUNTA_CODIGO", "").strip()
        desc = q.get("PREGUNTA_DESCRIPCION", "").strip()
        desired = q.get("PREGUNTA_SITUACION_DESEADA", "").strip()

        prompt = build_prompt(
            question_code=code,
            question_desc=desc,
            question_desired=desired,
            objective_final_text=objective_final,
            structured_transcript_tsv=structured_tsv,
            interviewer_speakers=cfg.INTERVIEWER_SPEAKERS,
            informant_speakers=cfg.INFORMANT_SPEAKERS,
            speaker_roles=cfg.SPEAKER_ROLES,
        )

        if cfg.SAVE_PROMPTS:
            (prompts_out / f"{code}.txt").write_text(prompt, encoding="utf-8")

        if args.dry_run:
            continue

        assert client is not None
        raw = chat_completion_text(client, prompt, oaicfg)
        raw = _sanitize_single_tsv_line(raw)

        rows_lines.append(raw)
        rows_dicts.append(parse_tsv_row(raw, columns))

    # 8) Guardar outputs
    if args.dry_run:
        print("Dry-run completado: prompts generados (sin llamar a la API)." )
        print(f"Carpeta prompts: {prompts_out}")
        return

    # TSV por filas (sin header, como pediste)
    rows_tsv_path = run_out / "resultado_rows.tsv"
    save_rows_tsv(rows_lines, rows_tsv_path)

    # Matriz consolidada
    df = rows_to_dataframe(rows_dicts, columns)
    csv_path, xlsx_path = save_outputs(df, run_out, base_name="systematization_matrix")

    print("\n✅ Listo")
    print(f"- TSV (filas): {rows_tsv_path}")
    print(f"- CSV matriz: {csv_path}")
    print(f"- XLSX matriz: {xlsx_path}")
    print(f"- Prompts: {prompts_out}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n[ERROR] {e}")
        sys.exit(1)
