# Ejecuta C8 (preview) con instalación segura
Set-Location -Path $PSScriptRoot
if (!(Test-Path ".\.venv\Scripts\Activate.ps1")) {
  python -m venv .venv
}
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip
python -m pip install -r requirements.txt
python .\main.py preview
