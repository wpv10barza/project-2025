from pathlib import Path

# Carpeta DATA PILOTO por defecto (puedes sobreescribir con --data-piloto-dir)
DEFAULT_DATA_PILOTO_DIR = Path(r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 03. Proyectos\P13. Sistema Evaluación v2\03.00 EJECUCION\03.00 DATA PILOTO")

# Copia adicional del prompt (pedido)
DEFAULT_PROMPT_TRABAJO_PATH = Path(r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 03. Proyectos\P13. Sistema Evaluación v2\03.00 EJECUCION\03.02 PRUEBA CON REGISTROS\01.02 Promp\Sistematizador\prompt_trabajo.txt")

# Selenium / Edge
EDGE_PROFILE_DIR = Path(r"C:\Users\ASUS\AppData\Local\Microsoft\Edge\User Data")
EDGE_PROFILE_NAME = "Default"
CHATGPT_URL = "https://chatgpt.com/"
DEFAULT_WAIT_SECONDS = 420  # 7 min

# SQL defaults
ODBC_DRIVER = "ODBC Driver 18 for SQL Server"
DEFAULT_TABLE = "dbo.Respuesta"

# UsuarioRegistroId por defecto (ajusta según tu BD)
DEFAULT_USUARIO_REGISTRO_ID = 1
