import os
import sys

PRT_path_base = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\05. RPA EXEC"
Carpeta_RPA = "09. MAF-DL"
PRT_CarpetaAProcesar = "REQ-PRUEBA"


sys.path.append(os.path.dirname(__file__))

import script_1_descarga_IA10_MAF
import HDR_MAF_result
import RPA_MAF_DL

def main():
    # Configuración compartida
    config = {
        "PRT_path_base": PRT_path_base,
        "Carpeta_RPA": Carpeta_RPA,
        "PRT_CarpetaAProcesar": PRT_CarpetaAProcesar,
    }

    # Ejecución en orden
    print("▶ Ejecutando Script 1: Descarga IA10_MAF")
    script_1_descarga_IA10_MAF.run(config)

    print("▶ Ejecutando Script 2: Formateo HDR + MAF")
    HDR_MAF_result.run(config)

    print("▶ Ejecutando Script 3: Descarga PRT CV04N")
    RPA_MAF_DL.run(config)

if __name__ == "__main__":
    main()