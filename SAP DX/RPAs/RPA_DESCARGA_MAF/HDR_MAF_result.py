import pandas as pd
import os 
import Descarga_IA10_MAF

# Datos configurables
PRT_path_base        = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\05. RPA EXEC"
Carpeta_RPA          = "09. MAF-DL"
PRT_CarpetaAProcesar = "REQ-PRUEBA"
Nombre_Planes_Mant   = "IP18.xlsx"
Nombre_Equipos       = "Lista_Equipos.xlsx"
Nombre_HDR_MAF       = "IA10_MAF.xlsx"


PATH_RPA_BASE        = os.path.join(PRT_path_base, Carpeta_RPA)
PATH_DESCARGA        = os.path.join(PATH_RPA_BASE, PRT_CarpetaAProcesar)
IP18                 = os.path.join(PATH_DESCARGA, Nombre_Planes_Mant)
IA10_MAF             = os.path.join(PATH_DESCARGA, Nombre_HDR_MAF)
EQUT                 = os.path.join(PATH_DESCARGA, Nombre_Equipos)

df_PM        = pd.read_excel(IP18) # Planes
df_EQ_CRIT   = pd.read_excel(EQUT) # Lista de equipos como Ubicaciones técnicas

def run(config):
    PRT_path_base = config["PRT_path_base"]
    Carpeta_RPA = config["Carpeta_RPA"]
    PRT_CarpetaAProcesar = config["PRT_CarpetaAProcesar"]

    PATH_RPA_BASE = os.path.join(PRT_path_base, Carpeta_RPA)
    PATH_DESCARGA = os.path.join(PATH_RPA_BASE, PRT_CarpetaAProcesar)

def format_IA10_MAF(IA10_MAF, df_HDR):
    df_MAF = pd.read_excel(IA10_MAF)


    # Formatear la IA10_MAF
    resultados = []
    for i in range(len(df_MAF)):
        fila_actual = df_MAF.iloc[i]
        if i+2 < len(df_MAF):
            fila_siguiente = df_MAF.iloc[i + 2]
            grupo = fila_actual['Unnamed: 1'] if pd.notna(fila_actual['Unnamed: 1']) else ""
            contador = fila_actual['Unnamed: 3'] if pd.notna(fila_actual['Unnamed: 3']) else ""
            documento_maf = fila_siguiente['Unnamed: 5'] if pd.notna(fila_siguiente['Unnamed: 2']) else ""
        else:
            grupo = fila_actual['Unnamed: 1'] if pd.notna(fila_actual['Unnamed: 1']) else ""
            contador = fila_actual['Unnamed: 3'] if pd.notna(fila_actual['Unnamed: 3']) else "" 
            documento_maf = ""
        # Verificar si tiene documento MAF asignado
        tiene_documento = pd.notna(fila_siguiente['Unnamed: 2']) and pd.notna(fila_siguiente['Unnamed: 4'])
        
        if grupo != "" and contador != "" or documento_maf != "":
            resultados.append({
                'GRUPO': grupo,
                'CONTADOR': contador,
                'DocumentoMAF': documento_maf,
                'TieneDocumentoMAF': tiene_documento
            })
    IA10_MAF_format = pd.DataFrame(resultados)

    IA10_MAF_format[['GRUPO','CONTADOR','DocumentoMAF']] = IA10_MAF_format[['GRUPO','CONTADOR','DocumentoMAF']].apply(
                                                                                                            pd.to_numeric, errors='coerce').astype('Int64').astype(str).replace('<NA>','')


    df_result = pd.merge(df_HDR, IA10_MAF_format, on =['GRUPO','CONTADOR'], how = "left")
    df_result = df_result[df_result['TieneDocumentoMAF']==True]

    df_result_group = df_result.groupby('Ubicación técnica', as_index=False)['DocumentoMAF'].count()
    df_result_group = df_result_group.rename(columns={'DocumentoMAF': '#MAF en SAP'})
 
    with pd.ExcelWriter(os.path.join(PATH_DESCARGA,'HDR con MAF.xlsx')) as writer:
        df_result.to_excel(writer, sheet_name="Resultados", index=False)
        df_result_group.to_excel(writer, sheet_name="MAF por EQ", index=False)


def filtro(df_EQ_CRIT, df_PM):
    # Renombrando filas
    df_EQ_CRIT = df_EQ_CRIT.rename(columns={'UT' : 'Ubicación técnica'})
    df_PM      = df_PM.rename(columns={'Ubicac.técnica':'Ubicación técnica',
                                        'Gpo.hojas ruta': 'GRUPO',
                                        'Cont.gpo.HRuta': 'CONTADOR'})
    # Filtrando UT
    df_EQ_CRIT = df_EQ_CRIT['Ubicación técnica'].dropna().unique()
    df_EQ_CRIT = pd.DataFrame(df_EQ_CRIT, columns=['Ubicación técnica'])

    df_HDR = pd.merge(df_EQ_CRIT, df_PM, on='Ubicación técnica', how='inner')
    df_HDR[['GRUPO','CONTADOR']] = df_HDR[['GRUPO','CONTADOR']].astype('Int64').astype(str)
    df_HDR = df_HDR[['Ubicación técnica','GRUPO','CONTADOR']]

    with pd.ExcelWriter(os.path.join(PATH_DESCARGA,'Hojas de Ruta con UT.xlsx')) as writer:
        df_HDR.to_excel(writer, sheet_name="HRD_DOC_MAF",index=False)
    df_HDR['GRUPO'].to_clipboard(index=False)

    return df_HDR 

if __name__ == '__main__':
    df_HDR = filtro(df_EQ_CRIT, df_PM)
    Descarga_IA10_MAF.main()
    format_IA10_MAF(IA10_MAF, df_HDR)
