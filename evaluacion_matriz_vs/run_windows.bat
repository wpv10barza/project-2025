@echo off
REM ====== Activar venv (si existe) ======
IF EXIST ".venv\Scripts\activate.bat" (
  call ".venv\Scripts\activate.bat"
)

REM ====== Ejecutar ======
python main.py

pause
