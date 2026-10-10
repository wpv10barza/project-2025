from __future__ import annotations
from pathlib import Path

# Repo root (donde está main.py)
REPO_DIR = Path(__file__).resolve().parent

# DATA PILOTO por defecto:
# - si existe carpeta "03.00 DATA PILOTO" junto al repo, úsala
# - si no, usa una carpeta "03.00 DATA PILOTO" dentro del repo
DEFAULT_DATA_PILOTO_DIR = (REPO_DIR.parent / "03.00 DATA PILOTO") if (REPO_DIR.parent / "03.00 DATA PILOTO").exists() else (REPO_DIR / "03.00 DATA PILOTO")

# Copia adicional del prompt (prompt_trabajo.txt).
# Por defecto se guarda en una subcarpeta del repo para evitar rutas rotas.
DEFAULT_PROMPT_TRABAJO_PATH = REPO_DIR / "prompt_trabajo.txt"

# Selenium / Edge
EDGE_PROFILE_DIR = Path(r"C:\Users\ASUS\AppData\Local\Microsoft\Edge\User Data")
EDGE_PROFILE_NAME = "Default"
CHATGPT_URL = "https://chatgpt.com/"
DEFAULT_WAIT_SECONDS = 420  # 7 min

# SQL defaults
ODBC_DRIVER = "ODBC Driver 18 for SQL Server"
DRIVER = ODBC_DRIVER  # alias para compatibilidad
DB_CONFIG_FILE = REPO_DIR / "db_config.json"  # alias para compatibilidad

DEFAULT_TABLE = "dbo.Respuesta"

# UsuarioRegistroId por defecto (ajusta según tu BD)
DEFAULT_USUARIO_REGISTRO_ID = 1
