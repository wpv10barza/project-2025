# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
from pathlib import Path

from src.prompt_builder import write_prompt_files
from src.selenium_runner import run_chatgpt_via_edge
from src.sql_loader import load_to_sql

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="p13_sistematizador",
        description="P13 Sistematizador: preview (C8) -> run (Edge/Selenium) -> load (SQL)"
    )
    sub = p.add_subparsers(dest="cmd", required=False)

    # preview
    p_prev = sub.add_parser("preview", help="Genera prompt_completo.txt (C8) desde SQL + transcripción")
    p_prev.add_argument("--evaluacion-id", type=int, required=False, default=16)
    p_prev.add_argument("--data-piloto-dir", type=str, required=False, default=None,
                        help="Carpeta donde están OBJETIVO_FINAL/Transcripcion y donde se guardan 01.02/01.03")
    p_prev.add_argument("--objetivo", type=str, required=False, default=None, help="Ruta a OBJETIVO_FINAL.txt")
    p_prev.add_argument("--transcripcion", type=str, required=False, default=None, help="Ruta a Transcripcion.txt o .pdf")
    p_prev.add_argument("--db-config", type=str, required=False, default=None, help="Ruta a db_config.json")

    # run
    p_run = sub.add_parser("run", help="Abre ChatGPT en Edge y genera TSV desde el prompt (sin API)")
    p_run.add_argument("--wait-seconds", type=int, default=420, help="Espera después de enviar el prompt")
    p_run.add_argument("--sheets-url", type=str, default=None, help="Opcional: URL de Google Sheets para pegar el TSV")
    p_run.add_argument("--headless", action="store_true", help="Modo headless (no recomendado para chatgpt.com)")

    # load
    p_load = sub.add_parser("load", help="Carga resultado XLSX/TSV a SQL (INSERT/MERGE)")
    p_load.add_argument("--evaluacion-id", type=int, required=False, default=16)
    p_load.add_argument("--informante-id", type=int, required=True)
    p_load.add_argument("--excel", type=str, required=False, default=None, help="Ruta a Matriz de Sistematización.xlsx")
    p_load.add_argument("--db-config", type=str, required=False, default=None, help="Ruta a db_config.json")

    return p

def main():
    parser = build_parser()
    args = parser.parse_args()

    cmd = (args.cmd or "preview").lower().strip()

    if cmd == "preview":
        out = write_prompt_files(
            evaluacion_id=args.evaluacion_id,
            data_piloto_dir=args.data_piloto_dir,
            objetivo_path=args.objetivo,
            transcripcion_path=args.transcripcion,
            db_config_path=args.db_config,
        )
        print(f"[OK] Prompt generado para revisar: {out}")
        return

    if cmd == "run":
        paths = run_chatgpt_via_edge(
            wait_seconds=args.wait_seconds,
            sheets_url=args.sheets_url,
            headless=args.headless,
        )
        # paths contiene los outputs guardados localmente
        print("[OK] Resultado guardado:")
        for k, v in paths.items():
            print(f"  - {k}: {v}")
        return

    if cmd == "load":
        stats = load_to_sql(
            evaluacion_id=args.evaluacion_id,
            informante_id=args.informante_id,
            excel_path=args.excel,
            db_config_path=args.db_config,
        )
        print("[OK] Carga a SQL completada:")
        for k, v in stats.items():
            print(f"  - {k}: {v}")
        return

    parser.print_help()

if __name__ == "__main__":
    main()
