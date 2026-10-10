# P13 Sistematizador (v17)

Flujo:
1) preview: genera prompt desde SQL (EvaluacionId) + Transcripcion (txt/docx/pdf)
2) run: Selenium/Edge pega el prompt en chatgpt.com (NO usa API)
3) load: toma el Excel final (descargado de Sheets) y hace MERGE a SQL

## Archivos de salida del preview
En DATA PILOTO (carpeta que pasas con --data-piloto-dir):
- 01.02 Promp/prompt_completo.txt
- 01.02 Promp/prompt_metadata.json
- 01.02 Promp/Sistematizador/prompt_trabajo.txt

## Ejemplos
python .\main.py preview --evaluacion-id 17 --data-piloto-dir "C:\...\03.03 PILOTO JA" --db-config ".\db_config.json"
python .\main.py run --data-piloto-dir "C:\...\03.03 PILOTO JA" --wait-seconds 420
python .\main.py load --evaluacion-id 17 --informante-id 108 --excel "C:\...\Matriz de Sistematización.xlsx" --db-config ".\db_config.json"



## Inventario de migración desde Google Drive

- Nombre: P13_sistematizador_selenium_v19.
- Propósito documentado: P13 Sistematizador (v17).
- Periodo de los archivos: 2025-12-30 a 2025-12-30.
- Tecnologías y formatos presentes: .md, .py, .txt.
- Estructura original: `README.md`, `config.py`, `main.py`, `requirements.txt`.
- Módulos principales: `main.py`, `config.py`.
- Origen: Google Drive, Proyectos_Exportados; ID de carpeta `1lGWOfrVANlruGhweBNws9tK1srfE7YOd`.
- Archivos originales seleccionados: 4.
- Observación: hay archivos excluidos por contener configuración sensible o requerir revisión; se conserva el proyecto original en Drive.
