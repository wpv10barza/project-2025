# Sistema de Sistematización — Visual Studio / PowerShell

## Dónde se instala
Copia/extrae esta carpeta dentro de:
`02.00 DATA PILOTO\01.01 Código de sistematizador`

## Insumos (en `02.00 DATA PILOTO`)
Nombres exactos:
- `PREGUNTA_CODIGO.pdf`
- `Transcripcion.pdf`
- `OBJETIVO_FINAL.txt`

## Salidas
- Prompt para revisión (C8 completo):
  `02.00 DATA PILOTO\01.02 Promp\prompt_completo.txt`
- Resultados:
  `02.00 DATA PILOTO\01.03 Resultado ChatGPT\resultado_raw.txt`
  `02.00 DATA PILOTO\01.03 Resultado ChatGPT\resultado.tsv|csv|xlsx`

## PowerShell
```powershell
cd "C:\Users\ASUS\...\02.00 DATA PILOTO\01.01 Código de sistematizador"
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt

# 1) Solo generar prompt (C8)
python main.py preview

# 2) Ejecutar GPT-4o (requiere API key)
$env:OPENAI_API_KEY="TU_API_KEY"
python main.py run
```

Si PowerShell bloquea scripts:
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```


## Inventario de migración desde Google Drive

- Nombre: 01.01 Código de sistematizador.
- Propósito documentado: Sistema de Sistematización — Visual Studio / PowerShell.
- Periodo de los archivos: 2025-12-17 a 2025-12-17.
- Tecnologías y formatos presentes: .md, .ps1, .py, .txt.
- Estructura original: `README.md`, `config.py`, `main.py`, `requirements.txt`, `run_gpt.ps1`, `run_preview.ps1`, `src`.
- Módulos principales: `main.py`, `config.py`, `src/utils/pdf_utils.py`.
- Origen: Google Drive, Proyectos_Exportados; ID de carpeta `1NBGPgnjWMLGcHO6gofUSO0L0n2-Jm4h1`.
- Archivos originales seleccionados: 7.
- Observación: se conserva la estructura y el contenido original de los archivos; la copia se verifica por hash antes de eliminar el origen.
