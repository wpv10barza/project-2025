import pandas as pd
import os
from datetime import date 
from rapidfuzz import process, fuzz
import re
def norm_keys_to_int(df, keys):
    for k in keys:
        if k in df.columns:
            df[k] = pd.to_numeric(df[k], errors='coerce').astype('Int64')

fecha_hoy  = date.today().strftime("%Y.%m.%d")
ruta_sap   = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION SEMANAL\04. RQ TI Des\00. Data\SAP\08. SAP"
ruta_hdr   = r"IA09_250826.xlsx"
ruta_pos   = r"IP18_250827.xlsx"
ruta_op    = r"IA10_OP_Format.xlsx"
ruta_maf   = r"IA10_MAF_Format.xlsx"
ruta_maf_d = r"CV04N.xlsx"
ruta_ubi_t = r"IH06.xlsx"



HDR   = pd.read_excel(os.path.join(ruta_sap, ruta_hdr))
OP    = pd.read_excel(os.path.join(ruta_sap, ruta_op))
POS   = pd.read_excel(os.path.join(ruta_sap, ruta_pos))
MAF   = pd.read_excel(os.path.join(ruta_sap, ruta_maf))
MAF_D = pd.read_excel(os.path.join(ruta_sap, ruta_maf_d))
UT    = pd.read_excel(os.path.join(ruta_sap, ruta_ubi_t))




HDR = HDR.rename(columns={'Descripción': 'Cabecera HDR'}).drop(
                                                                columns=["Fecha modificación", "Centro", 
                                                                        "Creado el", "Creado por",  "Indicador de borrado",
                                                                        "Modificado por", "Petición de borrado",  "Inicio validez",
                                                                        "Ubicación técnica", "Ctro.pto.trab.", "HRuta jerárquica", "Elem.referencia PM/PS"], 
                                                                errors='ignore')


OP  = OP.rename(columns={'Grupo hojas ruta': 'Grupo hojas ruta',
                         'Cont.grupo HRuta': 'Cont.grupo HRuta'})

POS  = POS.rename(columns={'Grupo hojas ruta': 'Grupo hojas ruta',
                         'Cont.grupo HRuta': 'Cont.grupo HRuta',
                         'Ubicación técnica':'Ubicac.técnica',
                         'Txt.pos.mantenim.':'TxtPosicManten',})




key_cols = ['Grupo hojas ruta', 'Cont.grupo HRuta']

for df in (HDR, POS, OP):
    norm_keys_to_int(df, key_cols)


HDR_AD_HOC = HDR[HDR['Estrategia mantenim.'].isna()]
UT         = UT.drop_duplicates(subset=['Campo clasif.'], keep='first')
HDR_AD_HOC['TAG'] = HDR['Cabecera HDR'].apply(lambda x: str(x).split(' ')[-1])
HDR_AD_HOC['TAG'] = HDR_AD_HOC['TAG'].apply(lambda x: x[:4]+"-"+x[4:6]+"-"+x[6:])
HDR_AD_HOC = pd.merge(HDR_AD_HOC, UT[['Ubicac.técnica','Campo clasif.']], left_on='TAG', right_on='Campo clasif.', how='left')

HDR_OP = pd.merge(HDR, OP, on = ["Grupo hojas ruta", "Cont.grupo HRuta"], how ='inner')
HDR_OP = pd.merge(HDR_OP, POS[["Grupo hojas ruta", "Cont.grupo HRuta", "Ubicac.técnica", "TxtPosicManten"]], on = ["Grupo hojas ruta", "Cont.grupo HRuta"], how ='left')


HDR_OP_MAF = pd.merge(HDR_OP, MAF, left_on = ["Grupo hojas ruta", "Cont.grupo HRuta"], right_on = ["GRUPO", "CONTADOR"], how ='left')
HDR_OP_MAF["DocumentoMAF"] = HDR_OP_MAF["DocumentoMAF"].astype('Int64')


HDR_OP_MAF = pd.merge(HDR_OP_MAF, MAF_D[["Documento","Document Long Text"]], left_on = ["DocumentoMAF"], right_on = ["Documento"], how ='left')
HDR_OP_MAF = HDR_OP_MAF[['Cabecera HDR', 'Grupo hojas ruta', 'Cont.grupo HRuta', 'Estrategia mantenim.', 
       'Puesto de trabajo', 'Grupo planif.', 'Estado global', 'Conjunto', 'Utilización', 
       'Descripción', 'Operación', 'Duración normal', 'Estado instalación_y', 'Pto.tbjo.operación',
       'Trabajo', 'Cantidad', 'Ubicac.técnica', 'TxtPosicManten', 'DocumentoMAF',
       'TieneDocumentoMAF', 'Documento', 'Document Long Text']].drop_duplicates()

HDR_OP_MAF = HDR_OP_MAF.set_index(['Grupo hojas ruta', 'Cont.grupo HRuta'])
HDR_AD_HOC = HDR_AD_HOC.set_index(['Grupo hojas ruta', 'Cont.grupo HRuta'])

HDR_OP_MAF.update(HDR_AD_HOC)
HDR_OP_MAF = HDR_OP_MAF.reset_index()

HDR_OP_MAF.to_excel(os.path.join(ruta_sap, f"HDR_OP_MAF_{fecha_hoy}.xlsx"), index = False )

######################################################################################

# ruta_prts_entregables = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION SEMANAL\03. Entregables\REQ_7"
# lista_equipos = os.listdir(ruta_prts_entregables)


ruta_data_input  = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION SEMANAL\04. RQ TI Des\00. Data\ORIEN\250901"
ruta_data_output = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.01 EJECUCION TI\05. RPA EXEC\- ENCARGOS TRANSVERSALES\20251908 EQ Criticos" 

OPERACION_INPUT = "2025829553_Quellaveco_ExportacindeOperacionesdeAltoNivel_PEH_oWzZ-.xlsx"
TAREAS_INPUT    = "20258291616_Quellaveco_ExportacindeListadeTareasdeAltoNivel__cdL5PiOO.xlsx"
EQUIPOS_INPUT   = "lista_equipos.xlsx"

df_OPERACION_INPUT = pd.read_excel(os.path.join(ruta_data_input, OPERACION_INPUT ), sheet_name="Operaciones de Alto Nivel",header=9)
df_TAREAS_INPUT    = pd.read_excel(os.path.join(ruta_data_input, TAREAS_INPUT ), sheet_name="Lista de Tareas de Alto Nivel",header=9)
df_EQUIPOS_INPUT   = pd.read_excel(os.path.join(ruta_data_input, EQUIPOS_INPUT ))

col_ubi = "Código de Ubicación"
col_op  = "Descripción de Operación"


df_OPERACION_INPUT = df_OPERACION_INPUT[[col_ubi, col_op, "Mano de Obra","No. de Personas Requeridas", "Trabajo"]]
df_TAREAS_INPUT    = df_TAREAS_INPUT[["Código de Ubicación de la Lista de Tareas", "Nombre de Lista de Tareas", "Estrategia de Mantenimiento"]]


df_ORIEN = pd.merge(df_OPERACION_INPUT, df_TAREAS_INPUT, left_on=col_ubi, right_on="Código de Ubicación de la Lista de Tareas",how="left" )
df_ORIEN = df_ORIEN[df_ORIEN[col_ubi].isin(df_EQUIPOS_INPUT["Equipos"])]


equipos = df_ORIEN[col_ubi].dropna().unique()
resultado = []  # filas por operación clasificada
umbral  = 92

cols_candidatas = ["Descripción", "Cabecera HDR", "TxtPosicManten", "Document Long Text"]

def _payload_row(row, payload_cols):
    # devuelve TODAS las columnas de la fila SAP con prefijo
    if row is None:
        return {f"SAP_{c}": None for c in payload_cols}


    return {f"SAP_{c}": row.get(c) for c in payload_cols}

def _norm(s):
    if pd.isna(s): return None
    s = re.sub(r"\s+", " ", str(s)).strip()
    return s.upper()  # ignora may/min
                  

def _best_score(a, b):
    if a is None or b is None: return -1
    return max(
        fuzz.token_sort_ratio(a, b),
        fuzz.token_sort_ratio(_norm(a), _norm(b)),
        fuzz.ratio(a, b),
    )
def comparar_ORIEN_vs_SAP(ops_orien, df_sap_full, cols_match, umbral=92, payload_cols=None):
    if payload_cols is None:
        payload_cols = list(df_sap_full.columns)

    ops_raw  = [None if pd.isna(x) else str(x) for x in ops_orien]
    ops_norm = [_norm(x) for x in ops_raw]

    usados_sap_rows, usados_sap_vals, usados_orien = set(), set(), set()
    res = []

    # EXACTOS (incluye exacto tras neutralizar periodicidad)
    for op_raw, op_norm in zip(ops_raw, ops_norm):
        if op_norm in usados_orien or op_norm is None:
            continue
        for j, row in df_sap_full.loc[~df_sap_full.index.isin(usados_sap_rows)].iterrows():
            for col in cols_match:
                val_raw  = row[col]
                val_norm = _norm(val_raw)
                if val_norm is None or val_norm in usados_sap_vals:
                    continue
                if op_norm == val_norm or _norm(op_norm) == _norm(val_norm):
                    estado = "EXISTE: MODIFICAR RECURSOS" if op_norm == val_norm else "actualizada"
                    out = {"OP_ORIEN": op_raw, "OP_SAP": str(val_raw), "score": 100, "estado": estado, "SAP_col_match": col}
                    out.update(_payload_row(row, payload_cols))
                    res.append(out)
                    usados_sap_rows.add(j); usados_sap_vals.add(val_norm); usados_orien.add(op_norm)
                    break
            if op_norm in usados_orien:
                break

    # FUZZY (mejor score ≥ umbral, 1–a–1)
    for op_raw, op_norm in zip(ops_raw, ops_norm):
        if op_norm in usados_orien or op_norm is None:
            continue
        best = (-1, None, None, None)  # score, j, col, val_norm
        for j, row in df_sap_full.loc[~df_sap_full.index.isin(usados_sap_rows)].iterrows():
            for col in cols_match:
                val_raw  = row[col]
                val_norm = _norm(val_raw)
                if val_norm is None or val_norm in usados_sap_vals:
                    continue
                sc = _best_score(op_norm, val_norm)
                if sc > best[0]:
                    best = (sc, j, col, val_norm)
        if best[1] is not None and best[0] >= umbral:
            sc, j, col, val_norm = best
            row = df_sap_full.loc[j]
            out = {"OP_ORIEN": op_raw, "OP_SAP": str(row[col]), "score": int(sc), "estado": "ACTUALIZAR", "SAP_col_match": col}
            out.update(_payload_row(row, payload_cols))
            res.append(out)
            usados_sap_rows.add(j); usados_sap_vals.add(val_norm); usados_orien.add(op_norm)

    # NUEVAS
    for op_raw, op_norm in zip(ops_raw, ops_norm):
        if op_norm in usados_orien:
            continue
        out = {"OP_ORIEN": op_raw, "OP_SAP": None, "score": 0, "estado": "CREAR", "SAP_col_match": None}
        out.update(_payload_row(None, payload_cols))
        res.append(out)
        usados_orien.add(op_norm)

    # ELIMINADAS (SAP sin pareja)
    for j, row in df_sap_full.iterrows():
        if j not in usados_sap_rows:
            mostrado = next((row[c] for c in ["Descripción","Cabecera HDR","TxtPosicManten","Document Long Text"]
                             if c in df_sap_full.columns and pd.notna(row.get(c))), None)
            out = {"OP_ORIEN": None, "OP_SAP": (None if mostrado is None else str(mostrado)),
                   "score": 100, "estado": "OPERACIÓN o HDR eliminada en ORIEN", "SAP_col_match": None}
            out.update(_payload_row(row, payload_cols))
            res.append(out)

    return res

for eq in equipos:
    ops_new = df_ORIEN.loc[df_ORIEN[col_ubi] == eq, col_op].astype(str).tolist()
    
    df_old_eq = HDR_OP_MAF[HDR_OP_MAF["Ubicac.técnica"] == eq]

#     print(df_old_eq.head(5))
    res_eq = comparar_ORIEN_vs_SAP(ops_new, df_old_eq, cols_candidatas, umbral=92)
    for r in res_eq:
        r[col_ubi] = eq
        resultado.append(r)
    # resultado.append(res_eq)



df_resultado = pd.DataFrame(resultado)
df_resultado.to_excel(os.path.join(ruta_data_output, f"ORIEN_SAP_{fecha_hoy}_v2.xlsx"), index = False )

# print(df_resultado[df_resultado[col_ubi] == "CE4000051"])
