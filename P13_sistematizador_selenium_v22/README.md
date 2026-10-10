# P13 Sistematizador (v22)

Cambios clave:
- Lee preguntas con `SELECT *` desde `dbo.VW_PREGUNTA_POR_EVALUACION` y **normaliza** nombres de columnas (evita error `Invalid column name`).
- LOAD hace MERGE a `dbo.Respuesta` usando **nombres reales**: `SituacionActual`, `Notas`, `RespuestaLiteral`.
- En los 3 campos aplica reemplazo de saltos reales de línea -> `\\n` literal antes de insertar/actualizar.

## Comandos

### 1) Preview (genera prompt_completo.txt)
```powershell
python .\main.py preview --evaluacion-id 17 --db-config ".\db_config.json" --transcripcion "C:\ruta\Transcripcion.txt"
```

### 2) Run (Selenium)
> Requiere `msedgedriver.exe` en la raíz del repo y estar logueado en Edge con el perfil Default.
```powershell
python .\main.py run --wait-seconds 420 --sheets-url "https://docs.google.com/..."
```

### 3) Load (sube Excel a SQL)
```powershell
python .\main.py load --evaluacion-id 17 --informante-id 108 --db-config ".\db_config.json" --excel "C:\ruta\MATRIZ_sistematización.xlsx"
```


## Inventario de migración desde Google Drive

- Nombre: P13_sistematizador_selenium_v22.
- Propósito documentado: P13 Sistematizador (v22).
- Periodo de los archivos: 2025-12-30 a 2025-12-30.
- Tecnologías y formatos presentes: .md, .py, .txt.
- Estructura original: `README.md`, `config.py`, `main.py`, `requirements.txt`, `src`.
- Módulos principales: `src/selenium_runner.py`, `src/prompt_builder.py`, `src/loaders.py`, `src/excel_reader.py`, `src/mapping.py`, `src/db_questions.py`, `src/sql.py`, `src/db.py`, `src/repo.py`, `main.py`, `config.py`.
- Origen: Google Drive, Proyectos_Exportados; ID de carpeta `1nkpA36ReD2GzrH_m9z6bA9whzRuggtRz`.
- Archivos originales seleccionados: 13.
- Observación: se conserva la estructura y el contenido original de los archivos; la copia se verifica por hash antes de eliminar el origen.
