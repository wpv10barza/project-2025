# Evaluación Matriz - C8 (FIX) Prompt Builder

Este proyecto genera y guarda en **TXT** el PROMPT completo (Objetivo Final + Guía + Transcripción)
para que puedas revisarlo antes de llamar a GPT.

## Archivos esperados (en tu carpeta DATA PILOTO)
- OBJETIVO_FINAL.txt
- PREGUNTA_CODIGO.pdf
- Transcripcion.pdf

## Uso (PowerShell)
```powershell
cd "C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 03. Proyectos\P13. Sistema Evaluación v2\02.00 DATA PILOTO"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

# (1) Crear/actualizar roles.json automáticamente (si no existe)
python main.py --init-roles

# (2) Edita roles.json y define roles (E1/E2/I1...) y qué speakers son entrevistadores/informantes.

# (3) Genera el prompt TXT para un código específico
python main.py --codigo 1.1.1

# (4) O genera prompts para todos los códigos detectados
python main.py --all
```

## Salida
Guarda archivos TXT en:
`...\02.00 DATA PILOTO\PRUEBA\prompt_C8_<codigo>.txt`


## Inventario de migración desde Google Drive

- Nombre: evaluacion_prompt_c8_vs.
- Propósito documentado: Evaluación Matriz - C8 (FIX) Prompt Builder.
- Periodo de los archivos: 2025-12-17 a 2025-12-17.
- Tecnologías y formatos presentes: .md, .py, .txt.
- Estructura original: `README.md`, `main.py`, `requirements.txt`, `src`.
- Módulos principales: `main.py`, `src/prompt_builder.py`, `src/config.py`.
- Origen: Google Drive, Proyectos_Exportados; ID de carpeta `1kEIroHTbl6LIrToYeL_zQu5ptw4p8Pzv`.
- Archivos originales seleccionados: 5.
- Observación: se conserva la estructura y el contenido original de los archivos; la copia se verifica por hash antes de eliminar el origen.
