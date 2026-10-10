import pandas as pd
from sentence_transformers import SentenceTransformer, util
import torch

path_tareas_ORIEN       = r'C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\04. RQ TI Dev\00. Data\ORIEN\Actividades Totales Orien.xlsx'
path_tareas_SAP         = r"C:\temp\PRTs\Tareas_Totales.xlsx"
path_tareas_COMPARACION = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\04. RQ TI Dev\00. Data\ORIEN\Orien_Sap_2.xlsx"

rows = []
model = SentenceTransformer('all-MiniLM-L6-v2')
ORIEN = pd.read_excel(path_tareas_ORIEN, sheet_name='Actividades',skiprows=9) 
# SAP   = pd.read_excel(r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\04. Información para Consultores\03. Area 3000\PRTs\Tareas_Totales.xlsx")
SAP   = pd.read_excel(path_tareas_SAP)


SAP = SAP[SAP['Tarea'].notnull() & (SAP['Tarea'].str.strip() != '')].reset_index(drop=True)

ORIEN['Equipo']   = ORIEN['Descripción del componente'].apply(lambda x: x.split()[0].strip() if pd.notnull(x) and '/' in x else "No se encontró") 
ORIEN['ItemMant'] = ORIEN['Descripción del componente'].apply(lambda x: x.split("/", 1)[1].strip() if pd.notnull(x) and '/' in x else "No se encontró")
SAP['Equipo']     = SAP['Ubicación Técnica'].astype(str) 
SAP['ItemMant']   = SAP['Ítem mantenible'].astype(str)  
ORIEN['Origen']   = "ORIEN"
SAP['Origen']     = "SAP"

# Procesar por equipo
equipos = SAP['Equipo'].unique()
ORIEN = ORIEN[ORIEN['Equipo'].isin(equipos)].reset_index(drop=True)

for eq in equipos:
    orien_df = ORIEN[ORIEN['Equipo'] == eq]
    orien_tasks = ORIEN[ORIEN['Equipo'] == eq]['Descripción de la actividad'].dropna().unique().tolist()
    sap_df = SAP[SAP['Equipo'] == eq][['Tarea','Límites aceptables', 'Comentario condicional','ItemMant']].dropna(subset=['Tarea'])
    sap_tasks = sap_df['Tarea'].unique().tolist()

    if len(orien_tasks) == 0:
        # Si no existe en ORIEN, registrar todas las tareas SAP como nuevas
        for task_sap in sap_tasks:
            task_info = sap_df[sap_df['Tarea'] == task_sap].iloc[0]
            rows.append({
                'Equipo': eq,
                'Item_SAP': task_info['ItemMant'],
                'Item_ORIEN':None,
                'Tarea_ORIEN': None,
                'Tarea_SAP': task_sap,
                'Similitud': None,
                'Origen_Final': 'SAP',
                'Limites_Aceptables': task_info['Límites aceptables'],
        'Comentarios_Condicionales': task_info['Comentario condicional']
            })
        continue    


    # Tareas iguales exactas
    common_tasks = set(orien_tasks).intersection(set(sap_tasks))

    # Registrar coincidencias exactas
    for task in common_tasks:
        task_info = sap_df[sap_df['Tarea'] == task].iloc[0]
        item_orien = orien_df['ItemMant'].iloc[0] if not orien_df.empty else None
        rows.append({
            'Equipo': eq,
            'Item_SAP': task_info['ItemMant'],
            'Item_ORIEN':item_orien,
            'Tarea_ORIEN': task,
            'Tarea_SAP': task,
            'Similitud': 1.0,
            'Origen_Final': 'SAP_ORIEN',
            'Limites_Aceptables': task_info['Límites aceptables'],
        'Comentarios_Condicionales': task_info['Comentario condicional']
        })

    # Filtrar tareas no exactas en SAP
    sap_remaining = list(set(sap_tasks) - common_tasks)

    if len(sap_remaining) > 0 and len(orien_tasks) > 0:
        # Embeddings para ORIEN y SAP restante
        orien_embeddings = model.encode(orien_tasks, convert_to_tensor=True)
        sap_embeddings = model.encode(sap_remaining, convert_to_tensor=True)
        
        # Calcular similitudes
        cosine_scores = util.pytorch_cos_sim(sap_embeddings, orien_embeddings)
        # print(cosine_scores)
        for i, task_sap in enumerate(sap_remaining):
            max_idx = torch.argmax(cosine_scores[i]).item()
            max_sim = cosine_scores[i][max_idx].item()
            task_orien_similar = orien_tasks[max_idx]
            task_info = sap_df[sap_df['Tarea'] == task_sap].iloc[0]
            if max_sim >= 0.80:
                origen_final = 'SAP_ORIEN' if max_sim == 1.0 else 'SAP'
                rows.append({
                    'Equipo': eq,
                    'Item_SAP': task_info['ItemMant'],
            'Item_ORIEN':item_orien,
                    'Tarea_ORIEN': task_orien_similar,
                    'Tarea_SAP': task_sap,
                    'Similitud': round(max_sim, 3),
                    'Origen_Final': origen_final,
        'Limites_Aceptables': task_info['Límites aceptables'],
        'Comentarios_Condicionales': task_info['Comentario condicional']
                })
            else:
                rows.append({
                    'Equipo': eq,
                    'Item_SAP': task_info['ItemMant'],
                    'Item_ORIEN':item_orien,
                    'Tarea_ORIEN': task_orien_similar,
                    'Tarea_SAP': task_sap,
                    'Similitud': round(max_sim, 3),
                    'Origen_Final': 'SAP',
        'Limites_Aceptables': task_info['Límites aceptables'],
        'Comentarios_Condicionales': task_info['Comentario condicional']
                })
    
    # Agregar tareas de ORIEN no registradas aún (no comunes)
    for task_orien in set(orien_tasks) - common_tasks:
        rows.append({
            'Equipo': eq,
            'Item_SAP': None,
            'Item_ORIEN': item_orien,
            'Tarea_ORIEN': task_orien,
            'Tarea_SAP': None,
            'Similitud': None,
            'Origen_Final': 'ORIEN',
    'Límites_Aceptables': None,
    'Comentarios_Condicionales': None
        })


final_result = pd.DataFrame(rows)



final_result.to_excel(path_tareas_COMPARACION, index = False)




