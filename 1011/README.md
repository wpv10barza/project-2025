# Evaluación – Matriz de Sistematización (GPT‑4o) – VS/PowerShell

Este proyecto hace 2 cosas:

1) **PREVIEW**: lee **PREGUNTA_CODIGO.pdf**, **Transcripcion.pdf** y **OBJETIVO_FINAL.txt** (o variantes de nombre),
   arma el **PROMPT COMPLETO** (guía + objetivo + transcripción) y lo guarda como:
   `Resultado sistematizado\prompt_completo.txt`  (para revisión previa).

2) **RUN**: (opcional) llama a **GPT‑4o** en lotes para construir la matriz y guarda:
   - `matriz.tsv`
   - `matriz.csv`
   - `matriz.xlsx`
   dentro de `Resultado sistematizado\`

## 0) Ubica los archivos de entrada
En tu carpeta DATA PILOTO (por ejemplo):
`C:\Users\ASUS\...\02.00 DATA PILOTO\`

Asegúrate de tener:
- `PREGUNTA_CODIGO.pdf`
- `Transcripcion.pdf`
- `OBJETIVO_FINAL.txt`  (o `OBJETIVO FINAL.txt` / `objetivo final.txt`)

## 1) Crear venv e instalar dependencias
En PowerShell (en la carpeta del proyecto):

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 2) Configurar tu API key
En PowerShell:

```powershell
setx OPENAI_API_KEY "TU_API_KEY_AQUI"
```

Cierra y abre una nueva consola (para que tome la variable).

## 3) Ajusta rutas
Edita `config.py` y ajusta `BASE_DIR` para que apunte a tu DATA PILOTO.
También puedes dejarlo en `"."` si copias el proyecto dentro de DATA PILOTO.

## 4) Generar el preview del prompt (C8)
```powershell
python main.py --preview
```

## 5) Ejecutar GPT‑4o y generar matriz
```powershell
python main.py --run
```


## Inventario de migración desde Google Drive

- Nombre: 1011.
- Propósito documentado: Evaluación – Matriz de Sistematización (GPT‑4o) – VS/PowerShell.
- Periodo de los archivos: 2025-12-17 a 2025-12-17.
- Tecnologías y formatos presentes: .md, .py, .txt.
- Estructura original: `README.md`, `config.py`, `main.py`, `requirements.txt`.
- Módulos principales: `main.py`, `config.py`.
- Origen: Google Drive, Proyectos_Exportados; ID de carpeta `1AIW3B4woeBHaRS1XXAtZFqIik4d4qV9Z`.
- Archivos originales seleccionados: 4.
- Observación: se conserva la estructura y el contenido original de los archivos; la copia se verifica por hash antes de eliminar el origen.
