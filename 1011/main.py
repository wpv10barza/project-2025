import argparse

import config
from src.io_utils import ensure_dir, find_required_inputs, read_pdf_text, read_txt, write_text
from src.parse_guide import parse_questions_from_guide_text
from src.parse_transcript import parse_transcript_lines
from src.prompt_builder import build_full_prompt
from src.gpt_runner import run_batches_and_save_outputs


def cmd_preview() -> None:
    ensure_dir(config.OUTPUT_DIR)

    guide_path, transcript_path, objective_path = find_required_inputs(
        base_dir=config.BASE_DIR,
        guide_name=config.GUIDE_PDF_NAME,
        transcript_name=config.TRANSCRIPT_PDF_NAME,
        objective_name=config.OBJECTIVE_TXT_NAME,
    )

    guide_text = read_pdf_text(guide_path)
    transcript_text = read_pdf_text(transcript_path)
    objective_text = read_txt(objective_path)

    questions = parse_questions_from_guide_text(guide_text)
    transcript_lines = parse_transcript_lines(transcript_text)

    prompt = build_full_prompt(
        objective_text=objective_text,
        questions=questions,
        transcript_lines=transcript_lines,
        columns=config.COLUMNS,
    )

    out_path = config.OUTPUT_DIR / "prompt_completo.txt"
    write_text(out_path, prompt)
    print(f"OK: Prompt completo guardado en:\n  {out_path}")


def cmd_run(batch_size: int) -> None:
    ensure_dir(config.OUTPUT_DIR)

    guide_path, transcript_path, objective_path = find_required_inputs(
        base_dir=config.BASE_DIR,
        guide_name=config.GUIDE_PDF_NAME,
        transcript_name=config.TRANSCRIPT_PDF_NAME,
        objective_name=config.OBJECTIVE_TXT_NAME,
    )

    guide_text = read_pdf_text(guide_path)
    transcript_text = read_pdf_text(transcript_path)
    objective_text = read_txt(objective_path)

    questions = parse_questions_from_guide_text(guide_text)
    if not questions:
        raise SystemExit("No se extrajeron preguntas desde la guía. Revisa PREGUNTA_CODIGO.pdf.")

    transcript_lines = parse_transcript_lines(transcript_text)
    if not transcript_lines:
        raise SystemExit("No se pudo parsear la transcripción. Revisa Transcripcion.pdf.")

    run_batches_and_save_outputs(
        model=config.MODEL,
        output_dir=config.OUTPUT_DIR,
        objective_text=objective_text,
        questions=questions,
        transcript_lines=transcript_lines,
        columns=config.COLUMNS,
        batch_size=batch_size,
    )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--preview", action="store_true", help="Genera prompt_completo.txt (sin llamar a GPT).")
    parser.add_argument("--run", action="store_true", help="Llama a GPT y genera matriz (tsv/csv/xlsx).")
    parser.add_argument("--batch-size", type=int, default=config.BATCH_SIZE, help="Cantidad de preguntas por llamada.")
    args = parser.parse_args()

    if args.preview:
        cmd_preview()
        return
    if args.run:
        cmd_run(batch_size=args.batch_size)
        return

    parser.print_help()


if __name__ == "__main__":
    main()
