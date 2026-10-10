# Ejecutar desde: 02.00 DATA PILOTO\01.01 Código de sistematizador
$ErrorActionPreference = "Stop"

py -m venv .venv
. .\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py preview
