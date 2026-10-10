import os
import shutil
import pandas as pd


# Rutas
ruta_base = r"C:\temp\PRTs"   # Carpeta base donde están todas las carpetas de ubicación técnica
ruta_tareas = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\04. RQ TI Dev\00. Data\ORIEN\Orien_Sap.xlsx" # Archivo maestro de tareas
ruta_tag_mapping = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\04. RQ TI Dev\00. Data\SAP\01 SAP C601 2025.06.02\IH06.xlsx" # Archivo con Tag y Ubicación Técnica
output_base = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\04. Información para Consultores\05. Equipos Criticos" # Carpeta donde se crearán los nuevos directorios



df_tags = pd.read_excel(ruta_tag_mapping, sheet_name='Hoja3')
df_tareas = pd.read_excel(ruta_tareas)


for tag in df_tags['Campo clasif.'].unique():
    tag_folder = os.path.join(output_base, str(tag))
    os.makedirs(tag_folder, exist_ok=True)
    
    ubicaciones_all = df_tags[df_tags['Campo clasif.'] == tag]['Ubicac.técnica'].unique()
    ubicaciones = [u for u in ubicaciones_all if os.path.exists(os.path.join(ruta_base, str(u)))]
    for ubic in ubicaciones:
        ubic_folder = os.path.join(tag_folder, str(ubic))
        sap_folder = os.path.join(ubic_folder, "SAP")
        orien_folder = os.path.join(ubic_folder, "ORIEN")
        doc_maf_folder = os.path.join(sap_folder, "Doc MAF")
        reportes_folder = os.path.join(sap_folder, "Reportes")
        

        os.makedirs(doc_maf_folder, exist_ok=True)
        os.makedirs(reportes_folder, exist_ok=True)
        os.makedirs(orien_folder, exist_ok=True)
        

        ubic_orig_folder = os.path.join(ruta_base, str(ubic))
        
        if os.path.exists(ubic_orig_folder):
            for file in os.listdir(ubic_orig_folder):
                if file.lower().endswith(".pdf"):
                    shutil.copy2(os.path.join(ubic_orig_folder, file), doc_maf_folder)
            

            xlsx_path = os.path.join(ubic_orig_folder, "PRT_data.xlsx")
            if os.path.exists(xlsx_path):
                new_filename = f"{tag}_PRT_data.xlsx"
                dest_path = os.path.join(reportes_folder, new_filename)
                shutil.copy2(xlsx_path, dest_path)
                    

        tareas_ubic = df_tareas[df_tareas['Ubicación Técnica'] == ubic]
        output_tareas = os.path.join(orien_folder, f"{ubic}_Tareas_Orien_Sap.xlsx")
        tareas_ubic.to_excel(output_tareas, index=False)






