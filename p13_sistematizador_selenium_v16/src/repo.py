# -*- coding: utf-8 -*-
from pathlib import Path
import os

def resolve_data_piloto_dir(arg_dir: str | None, fallback: Path) -> Path:
    if arg_dir and str(arg_dir).strip():
        return Path(arg_dir).expanduser().resolve()
    env = os.environ.get("DATA_PILOTO_DIR", "")
    if env.strip():
        return Path(env).expanduser().resolve()
    return fallback.resolve()
