# -*- coding: utf-8 -*-
"""
Monitorea prompt_completo.txt y cuando cambia, ejecuta:
  python main.py run

Uso:
  python scripts/loop_sistematizador.py
"""
from __future__ import annotations
import os, time, subprocess
from pathlib import Path
from config import PROMPT_TXT_OUT

def main():
    prompt = Path(PROMPT_TXT_OUT)
    if not prompt.exists():
        print(f"[WARN] No existe: {prompt}")
        return

    last = prompt.stat().st_mtime
    print("[OK] Monitor activo. Observando:", prompt)
    while True:
        try:
            now = prompt.stat().st_mtime
            if now != last:
                print("[OK] Cambio detectado. Ejecutando run...")
                subprocess.call(["python", "main.py", "run"])
                last = prompt.stat().st_mtime
            time.sleep(3)
        except KeyboardInterrupt:
            break
        except Exception:
            time.sleep(3)

if __name__ == "__main__":
    main()
