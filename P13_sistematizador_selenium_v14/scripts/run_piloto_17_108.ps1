\
    # P13 - Piloto (PowerShell) - v14
    # Ajusta estas rutas:
    $repoPath = "C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 03. Proyectos\P13. Sistema Evaluación v2\03.00 EJECUCION\03.02 PRUEBA CON REGISTROS\P13_sistematizador_selenium_v13"
    $docxPath = "C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 03. Proyectos\P13. Sistema Evaluación v2\03.00 EJECUCION\03.03 PILOTO JA\LIBRERIA.Transcripción.docx"

    $dataPilotoDir = Split-Path $docxPath -Parent
    $evaluacionId = 17
    $informanteId = 108

    Set-Location $repoPath

    if (!(Test-Path ".\.venv\Scripts\Activate.ps1")) { python -m venv .venv }
    .\.venv\Scripts\Activate.ps1

    python -m pip install --upgrade pip
    python -m pip install -r .\requirements.txt
    python -m pip install python-docx

    # Export DOCX -> Transcripcion.txt
    $outTxt = Join-Path $dataPilotoDir "Transcripcion.txt"
    python .\scripts\convert_docx_to_txt.py --docx "$docxPath" --out "$outTxt"

    # DATA_PILOTO_DIR para que el sistema lea insumos/salidas del piloto
    $env:DATA_PILOTO_DIR = $dataPilotoDir

    # Preview / Run / Load
    python .\main.py preview --evaluacion-id $evaluacionId
    python .\main.py run --evaluacion-id $evaluacionId
    python .\main.py load --evaluacion-id $evaluacionId --informante-id $informanteId
