# P13 Sistematizador (sin API) — Preview → Selenium (Edge) → Load SQL

## 0) Preparar carpeta de piloto
1. Descomprime este ZIP en una carpeta, por ejemplo:
   `...\03.02 PRUEBA CON REGISTROS\p13_sistematizador_selenium_v12`

2. En la carpeta **padre** (DATA_PILOTO) coloca:
- `OBJETIVO_FINAL.txt` (si no existe, se creará uno base automáticamente)
- `Transcripcion.txt` **o** `Transcripcion.pdf`
- (Opcional) carpeta `Transcripciones\` con varios `.txt/.pdf`
- `db_config.json` (credenciales de BD, no se incluye en el ZIP)

Ejemplo `db_config.json`:
```json
{
  "server": "gaeval.database.windows.net",
  "database": "gaeval",
  "username": "gaevaluser",
  "password": "PON_TU_PASSWORD"
}
```

## 1) Crear venv + instalar
En PowerShell, dentro de la carpeta del ZIP:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r .\requirements.txt
```

## 2) Preview (C8)
Genera el `prompt_completo.txt` en `01.02 Promp` **y además** intenta copiarlo a `G:\Mi unidad\PROMP` si existe.

```powershell
python .\main.py preview --evaluacion-id 17
```

Opcional:
```powershell
python .\main.py preview --evaluacion-id 17 --transcripcion "C:\...\Transcripcion.txt"
python .\main.py preview --evaluacion-id 17 --objetivo "C:\...\OBJETIVO_FINAL.txt"
python .\main.py preview --evaluacion-id 17 --db-config "C:\...\db_config.json"
```

## 3) Run (sin API) — Edge + Selenium
Abre `chatgpt.com` usando tu perfil de Edge, pega el prompt y guarda el TSV/CSV/XLSX en `01.03 Resultado ChatGPT`
y también en `G:\Mi unidad\TGA_PROYECTO\Matriz de sistematización` si existe.

```powershell
python .\main.py run
```

Si tu matriz es grande:
```powershell
python .\main.py run --wait-seconds 900
```

Opcional pegar en Google Sheets:
```powershell
python .\main.py run --sheets-url "https://docs.google.com/spreadsheets/d/...."
```

> **Nota**: Debes estar logueado en ChatGPT dentro de Edge. Si el botón “Copiar código” no aparece,
el sistema intentará leer el último bloque de código directamente.

## 4) Load a SQL
Carga el XLSX resultante (por defecto `01.03 Resultado ChatGPT\resultado.xlsx`) hacia `dbo.Respuesta` usando MERGE.

```powershell
python .\main.py load --evaluacion-id 17 --informante-id 97
```

O indicar un Excel específico:
```powershell
python .\main.py load --evaluacion-id 17 --informante-id 97 --excel "G:\Mi unidad\TGA_PROYECTO\Matriz de sistematización\resultado.xlsx"
```

## 5) Monitor (opcional)
Cuando cambie el prompt, ejecuta automáticamente `run`:
```powershell
python .\scripts\loop_sistematizador.py
```


## Inventario de migración desde Google Drive

- Nombre: p13_sistematizador_selenium_v12.
- Propósito documentado: P13 Sistematizador (sin API) — Preview → Selenium (Edge) → Load SQL.
- Periodo de los archivos: 2025-12-28 a 2025-12-29.
- Tecnologías y formatos presentes: .json, .md, .py, .txt.
- Estructura original: `README.md`, `config.py`, `db_config.example.json`, `main.py`, `requirements.txt`, `scripts`, `src`.
- Módulos principales: `src/prompt_builder.py`, `src/selenium_runner.py`, `src/db_questions.py`, `src/sql_loader.py`, `src/db.py`, `src/diagnostics.py`, `src/output_utils.py`, `scripts/loop_sistematizador.py`, `src/io_utils.py`, `config.py`, `main.py`.
- Origen: Google Drive, Proyectos_Exportados; ID de carpeta `17Otk1lluev5U6YB-MUWCm1VHyqi16RtM`.
- Archivos originales seleccionados: 14.
- Observación: se conserva la estructura y el contenido original de los archivos; la copia se verifica por hash antes de eliminar el origen.
