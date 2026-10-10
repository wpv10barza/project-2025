# -*- coding: utf-8 -*-
import argparse
from pathlib import Path

from config import DEFAULT_DB_CONFIG
from src.prompt_builder import write_prompt_files
from src.selenium_runner import run_chatgpt_via_edge
from src.sql_loader import load_stub

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="p13_sistematizador")
    sub = p.add_subparsers(dest="cmd", required=True)

    p_prev = sub.add_parser("preview", help="Genera prompt_completo.txt desde SQL + Transcripcion")
    p_prev.add_argument("--evaluacion-id", type=int, required=True)
    p_prev.add_argument("--data-piloto-dir", default=None)
    p_prev.add_argument("--objetivo", default=None)
    p_prev.add_argument("--transcripcion", default=None, help="Ruta a Transcripcion.txt/.pdf/.docx")
    p_prev.add_argument("--db-config", default=str(DEFAULT_DB_CONFIG), help="Ruta a db_config.json")
    p_prev.add_argument("--prompt-trabajo", default=None, help="Ruta extra para copiar el prompt (prompt_trabajo.txt)")

    p_run = sub.add_parser("run", help="Ejecuta Selenium/Edge: pega prompt_trabajo.txt en chatgpt.com")
    p_run.add_argument("--wait-seconds", type=int, default=420)
    p_run.add_argument("--sheets-url", default=None)
    p_run.add_argument("--headless", action="store_true")
    p_run.add_argument("--prompt", default=None, help="Ruta a prompt_trabajo.txt (si no, usa el default)")
    p_run.add_argument("--no-kill-edge", action="store_true", help="No matar msedge/msedgedriver antes de abrir")
    p_run.add_argument("--edge-user-data-dir", default=None, help=r"Ej: C:\Users\ASUS\AppData\Local\Microsoft\Edge\User Data")
    p_run.add_argument("--edge-profile-dir", default=None, help="Ej: Default")
    p_run.add_argument("--edge-binary", default=None, help=r"Ruta a msedge.exe")

    p_load = sub.add_parser("load", help="(stub) Verifica conexión y obtiene AplicacionId (eval+informante)")
    p_load.add_argument("--evaluacion-id", type=int, required=True)
    p_load.add_argument("--informante-id", type=int, required=True)
    p_load.add_argument("--db-config", default=str(DEFAULT_DB_CONFIG))

    return p

def main():
    args = _build_parser().parse_args()

    if args.cmd == "preview":
        out = write_prompt_files(
            evaluacion_id=args.evaluacion_id,
            data_piloto_dir=args.data_piloto_dir,
            objetivo_path=args.objetivo,
            transcripcion_override=args.transcripcion,
            db_config_path=args.db_config,
            prompt_trabajo_path=args.prompt_trabajo,
        )
        print(f"[OK] Prompt generado para revisar: {out}")
        return

    if args.cmd == "run":
        paths = run_chatgpt_via_edge(
            wait_seconds=args.wait_seconds,
            sheets_url=args.sheets_url,
            headless=args.headless,
            prompt_path=args.prompt,
            kill_edge=(not args.no_kill_edge),
            edge_user_data_dir=args.edge_user_data_dir,
            edge_profile_dir=args.edge_profile_dir,
            edge_binary=args.edge_binary,
        )
        print("[OK] Selenium terminó. Prompt usado:", paths["prompt_used"])
        return

    if args.cmd == "load":
        app_id = load_stub(Path(args.db_config), args.evaluacion_id, args.informante_id)
        print(f"[OK] Conexión OK. AplicacionId={app_id}")
        return

if __name__ == "__main__":
    main()
