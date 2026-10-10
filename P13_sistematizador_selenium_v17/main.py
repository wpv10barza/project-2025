from __future__ import annotations
import argparse
from pathlib import Path

from config import DEFAULT_WAIT_SECONDS
from src.prompt_builder import write_prompt_files
from src.selenium_runner import run_chatgpt_via_edge
from src.loaders import load_excel_to_sql


def _p(path_str: str) -> Path:
    return Path(path_str).expanduser().resolve()


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="p13_sistematizador", description="Preview (C8) + Run (Selenium) + Load (SQL).")
    sub = p.add_subparsers(dest="cmd", required=True)

    pv = sub.add_parser("preview", help="Genera prompt_completo.txt desde SQL + Transcripcion.*")
    pv.add_argument("--evaluacion-id", type=int, default=16)
    pv.add_argument("--data-piloto-dir", type=_p, default=None,
                    help="Carpeta donde están OBJETIVO_FINAL/Transcripcion y donde se guardan 01.02/01.03")
    pv.add_argument("--objetivo", type=_p, default=None, help="Ruta a OBJETIVO_FINAL.txt (override)")
    pv.add_argument("--transcripcion", type=_p, default=None, help="Ruta a Transcripcion.txt/.pdf/.docx (override)")
    pv.add_argument("--db-config", type=_p, default=None, help="Ruta a db_config.json (override)")

    rn = sub.add_parser("run", help="Usa Selenium (Edge) para pegar el prompt en ChatGPT web.")
    rn.add_argument("--data-piloto-dir", type=_p, default=None,
                    help="Para ubicar el prompt generado.")
    rn.add_argument("--prompt-file", type=_p, default=None, help="Ruta a prompt_completo.txt (override)")
    rn.add_argument("--wait-seconds", type=int, default=DEFAULT_WAIT_SECONDS, help="Espera después de enviar el prompt")
    rn.add_argument("--sheets-url", type=str, default="", help="Opcional: URL de Google Sheets para pegar el TSV")
    rn.add_argument("--headless", action="store_true", help="Modo headless (no recomendado para chatgpt.com)")

    ld = sub.add_parser("load", help="Carga Matriz.xlsx a SQL (merge/insert).")
    ld.add_argument("--evaluacion-id", type=int, default=16)
    ld.add_argument("--informante-id", type=int, required=True)
    ld.add_argument("--excel", type=_p, default=None, help="Ruta a Matriz de Sistematización.xlsx")
    ld.add_argument("--db-config", type=_p, default=None, help="Ruta a db_config.json (override)")
    ld.add_argument("--table", type=str, default="dbo.Respuesta_Prueba_Plaud", help="Tabla destino")

    return p


def main() -> None:
    args = build_parser().parse_args()

    if args.cmd == "preview":
        out = write_prompt_files(
            evaluacion_id=args.evaluacion_id,
            data_piloto_dir=args.data_piloto_dir,
            objetivo_path=args.objetivo,
            transcripcion_override=args.transcripcion,
            db_config_path=args.db_config,
        )
        print(f"[OK] Prompt generado para revisar: {out}")
        return

    if args.cmd == "run":
        run_chatgpt_via_edge(
            data_piloto_dir=args.data_piloto_dir,
            prompt_file=args.prompt_file,
            wait_seconds=args.wait_seconds,
            sheets_url=args.sheets_url.strip(),
            headless=args.headless,
        )
        return

    if args.cmd == "load":
        load_excel_to_sql(
            evaluacion_id=args.evaluacion_id,
            informante_id=args.informante_id,
            excel_path=args.excel,
            db_config_path=args.db_config,
            table=args.table,
        )
        return


if __name__ == "__main__":
    main()
