from __future__ import annotations
import argparse
from pathlib import Path

from src.prompt_builder import write_prompt_files
from src.selenium_runner import run_chatgpt_via_edge
from src.loaders import load_excel_to_sql
from config import DEFAULT_DATA_PILOTO_DIR, DEFAULT_TABLE

def main():
    parser = argparse.ArgumentParser(prog="p13_sistematizador")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_prev = sub.add_parser("preview")
    p_prev.add_argument("--evaluacion-id", type=int, required=True)
    p_prev.add_argument("--data-piloto-dir", type=str, default=str(DEFAULT_DATA_PILOTO_DIR))
    p_prev.add_argument("--objetivo", type=str, default=None)
    p_prev.add_argument("--transcripcion", type=str, default=None)
    p_prev.add_argument("--db-config", type=str, required=True)

    p_run = sub.add_parser("run")
    p_run.add_argument("--wait-seconds", type=int, default=None)
    p_run.add_argument("--sheets-url", type=str, default=None)
    p_run.add_argument("--headless", action="store_true")

    p_load = sub.add_parser("load")
    p_load.add_argument("--evaluacion-id", type=int, required=True)
    p_load.add_argument("--informante-id", type=int, required=True)
    p_load.add_argument("--excel", type=str, required=True)
    p_load.add_argument("--db-config", type=str, required=True)
    p_load.add_argument("--table", type=str, default=DEFAULT_TABLE)

    args = parser.parse_args()

    if args.cmd == "preview":
        out = write_prompt_files(
            evaluacion_id=args.evaluacion_id,
            db_config_path=Path(args.db_config),
            data_piloto_dir=Path(args.data_piloto_dir),
            objetivo_path=Path(args.objetivo) if args.objetivo else None,
            transcripcion_override=Path(args.transcripcion) if args.transcripcion else None,
        )
        return out

    if args.cmd == "run":
        # Por defecto toma el prompt generado en DATA PILOTO / 01.02 Promp
        prompt_path = Path(DEFAULT_DATA_PILOTO_DIR) / "01.02 Promp" / "prompt_completo.txt"
        wait_seconds = args.wait_seconds if args.wait_seconds is not None else 420
        run_chatgpt_via_edge(prompt_path=prompt_path, wait_seconds=wait_seconds, sheets_url=args.sheets_url, headless=args.headless)
        return

    if args.cmd == "load":
        return load_excel_to_sql(
            evaluacion_id=args.evaluacion_id,
            informante_id=args.informante_id,
            excel_path=Path(args.excel),
            db_config_path=Path(args.db_config),
            table=args.table,
        )

if __name__ == "__main__":
    main()
