import pandas as pd 
import os

lista_clases = {
                'PM01':"Trabajo ad-hoc", # OT Correctivo
                'PM03':"Mantenimiento planificado" # OT Preventivo
                }

lista_status_usu = {'PLAN': "Planificado",
                    'PROV': "Provisionado",
                    'PROG': "Programado",
                    'EXEC': "En ejecución"
                    }

lista_status_sis = {'CTEC': "Cerrada técnicamente",
                    'JBFI': "Job finalizado",
                    'KKMP': "Sin componentes de material",
                    'MOVM': "Mov.mercancías ejecutado",
                    'NLIQ': "Norma de liquidación entrada",
                    'PREC': "Precálculo del coste",
                    'CERR': "Cerrada",
                    'NOTP': "Parcialmente confirmados",
                    'IMPR': "Impreso",
                    'MACO': "Material confirmado",
                    'NOTI': "Notificado",
                    'IMPP': "Impreso parcialmente",
                    'EDET': "Error al determinar costes",
                    'NEJE': "No se ha completado",
                    'LIB.': "Liberada",
                    'FMAT': "Escasez de material",
                    'FENA': "Fechas no actuales",
                    'ABIE': "Abiertos",
                    'LIBR': "Liberación rechazada",

                    'INBO': "In. de borrado, archivable",
                    'OTPS': "Oper.tbjo.ext.parcialm.sumin.",
                    'PWF': "<DATO-NO-ENCONTRADO>"
                    }

# Obtener datos
path_archivos  = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\04. RQ TI Dev\00. Data"
path_resultado = r""

IP16 = path_archivos + "IP16.xlsx"
IP18 = path_archivos + "IP18.xlsx"
IA09 = path_archivos + "IA09.xlsx"
IW39 = path_archivos + "IW39_Filtrada.xlsx"
CR05 = path_archivos + "CR05.xlsx"
IW37 = path_archivos + "IW37N_Filtrada.xlsx"
HROP = path_archivos + "IA10_OP.xlsx"
MAF  = path_archivos + "IA10_MAF.xlsx"

Lista_Ubicaciones = [
    "BU4000046", "BU4000047", "CE4000051", "CE4000052", "CE4000053", "CE4000054", "CE4000055", "CE4000056", "CE4000057", "CE4000058", "CE4000059", "CE4000060", "CE4000061", "CE4000062", "CE4000063", "CE4000064",
    "CE4000065", "CE4000066", "CE4000067", "CE4000068", "CE4000069", "CE4000070", "CE4000071", "CE4000072", "CE4000073", "CE4000074", "CE4000075", "CE4000076", "CH4000045", "CH4000045", "CH4000046", "CH4000047",
    "CH4000048", "CH4000049", "CH4000051", "CH4000053", "CH4000054", "CH4000059", "CH4000061", "CH4000063", "CH4000065", "CH4000066", "CN4000015", "CV4000003", "CV4000004", "CV4000007", "CV4000008", "CV4000013",
    "EN4000001", "FE4000003", "FE4000004", "FE4000005", "FE4000006", "FE4000007", "FE4000008", "FE4000009", "FE4000010", "FL4000002", "FL4000003", "FL4000004", "HP4000005", "HP4000006", "HP4000007", "HP4000008",
    "HP4000009", "HP4000010", "HP4000011", "HP4000012", "HP4000014", "LA4000004", "LA4000005", "LA4000012", "LA4000013", "LA4000014", "MA4000003", "MA4000004", "MA4000005", "MI4000001", "MI4000002", "MI4000003",
    "MI4000004", "PD4000010", "PD4000011", "PD4000013", "PD4000014", "PD4000015", "SA4000100", "SM4000011", "SM4000012", "SM4000018", "SM4000024", "SM4000200", "SP4000204", "SR4000001", "SR4000002", "SU4000014",
    "SU4000020", "SU4000021", "SU4000022", "SU4000023", "TA4000048", "TH4000002", "TH4000003", "TH4000004", "TH4000005", "TH4000111", "TH4000201", "TH4000203", "TH4000204", "TH4000219", "TH4000226", "TH4000227",
    "TH4000231", "WS4000002", "WS4000003", "CH4000032", "CH4000033", "CH4000034", "CH4000057", "CH4000098", "CH4000068", "CH4000083", "CH4000084", "CH4000085", "CH4000086", "CH4000087", "CH4000088", "WS4000004",
    "WS4000005", "WS4000006", "40020103.35.10.01", "40020103.35.10.03", "40020103.34.10.01", "40020103.31.40.02", "40020103.32.10.01", "40020103.32.10.02", "40020103.32.20.01", "40020103.36.30.01", "40020103.31.40.01",
    "40020103.32.10.09", "40020103.33.10.14", "40020103.37.10.01", "40020103.32.20.06", "40020103.33.20.04", "40020103.34.10.04", "40020103.37.40.01", "40020103.33.10.12", "40020103.32.10.08", "40020103.33.10.15",
    "40020103.37.20.01", "40020103.36.20.01", "40020103.32.10.05","CH4000020", "CH4000021", "CH4000022", "CH4000023", "CH4000024", "CH4000025",
    "CH4000026", "CH4000027", "CH4000037", "CH4000038", "CH4000039", "CH4000040",
    "CH4000041", "CH4000042", "CH4000043", "CH4000044"
]





def formatear_hoja_excel(
    writer,
    df,
    nombre_hoja,
    nombre_tabla,
    agrupar_columnas=[],
    estilo_tabla='Table Style Medium 9',
    aplicar_agrupado=True,
    columnas_colorear=None
):
    df.to_excel(writer, sheet_name=nombre_hoja, index=False)
    workbook = writer.book
    worksheet = writer.sheets[nombre_hoja]

    formato_cabecera = workbook.add_format({
        'bold': True,
        'bg_color': '#18814a',
        'border': 1,
        'align': 'center'
    })
    # Ajustar ancho de columnas
    for i, col in enumerate(df.columns):
        max_len = max(df[col].astype(str).map(len).max(), len(col))
        worksheet.set_column(i, i, max_len + 2)
        # Colorear cabecera si está en la lista
        if columnas_colorear and col in columnas_colorear:
            worksheet.write(0, i, col, columnas_colorear)

    # Agrupamiento
    if aplicar_agrupado:
        worksheet.outline_settings(True, False, False, False)
        worksheet.set_row(0, None, None, {'level': 1})
        for rango in agrupar_columnas:
            worksheet.set_column(rango, None, None, {'level': 1})

    # Tabla
    max_row, max_col = df.shape
    worksheet.add_table(0, 0, max_row, max_col - 1, {
        'columns': [{'header': col} for col in df.columns],
        'name': nombre_tabla,
        'style': estilo_tabla
    })

def acotar_df_por_fecha(df,columna,inicio,fin):
  # df_OTOP Falta filtrar la "Fecha Referencia" < 19/03/2025
  fecha_inicio = pd.Timestamp(inicio)
  fecha_fin = pd.Timestamp(fin)

  df[columna] = pd.to_datetime(df[columna], format='%d.%m.%Y')
  return df[(df[columna] >= fecha_inicio) & (df[columna] <= fecha_fin)]



df_PM        = pd.read_excel(IP16)
df_PPM       = pd.read_excel(IP18)
df_HR        = pd.read_excel(IA09)
df_PRT       = pd.read_excel(MAF)
df_HROP      = pd.read_excel(HROP)
df_OTOP      = pd.read_excel(IW37)
df_OT        = pd.read_excel(IW39)
df_CR05      = pd.read_excel(CR05)

df_UT_Filter = pd.DataFrame({'Ubicación técnica':Lista_Ubicaciones})


df_PPM  = df_PPM.rename (columns={'Gpo.hojas ruta':"GRUPO",
                                  'Cont.gpo.HRuta':"CONTADOR",
                                  'Ubicac.técnica':'Ubicación técnica',
                                  'Plan mantenim.':'Plan mant.preventivo',
                                  'Pos.mantenim.':'Posición mantenim.'})

df_PM  = df_PM.rename (columns={'Plan mantenim.':'Plan mant.preventivo',
                                'Pos.mantenim.':'Posición mantenim.'})

df_HR   = df_HR.rename  (columns={"Gpo.hojas ruta":"GRUPO",
                                  "Cont.gpo.HRuta":"CONTADOR"})

df_HROP = df_HROP.rename(columns={"Grupo hojas ruta":"GRUPO",
                                  "Cont.grupo HRuta":"CONTADOR",
                                  "Operación":"OPERACIÓN",
                                  "Pto.tbjo.operación":"PuestoTrabajo",
                                  "Duración normal":"Duración_OP",
                                  "Cantidad": "Cantidad_OP",
                                  'Trabajo': "Trabajo_OP",
                                  'Estado instalación':'EstadoInstalación'})

df_CR05 = df_CR05.rename(columns={"PtoTbjo":"PuestoTrabajo",
                                  "Descripción":"PuestoTrabajoDescripcion"})

df_OTOP = df_OTOP.rename(columns={"Gpo.hojas ruta":"GRUPO",
                                  "Cont.gpo.HRuta":"CONTADOR",
                                  "Operación":"OPERACIÓN",
                                  "Plan mantenim.":"Plan mant.preventivo",
                                  "Pos.mantenim.":"Posición mantenim.",
                                  'Cantidad':'Cantidad_Plan',
                                  'Trabajo':'Trabajo_Plan',
                                  'Duración normal':'Duración normal',
                                  'Duración mínima':'Duración mínima',
                                  'Cantidad real':'Cantidad real',
                                  'Trabajo real':'Trabajo real',
                                  'Cantidad requerida':'Cantidad requerida',
                                  'Descripción': 'Descripción de OTOP'})

df_OT   = df_OT.rename  (columns={"Gpo.hojas ruta":"GRUPO",
                                  "Cont.gpo.HRuta":"CONTADOR",
                                  "Plan mantenim.":"Plan mant.preventivo",
                                  "Pos.mantenim.":"Posición mantenim."})

df_UT_Filter.name = "Ubicación técnica"


df_HROP = pd.merge(
         df_PRT,df_HROP,
         on=['GRUPO', 'CONTADOR'],#, 'OPERACIÓN'],
         how='inner') # IA10 + IA39

df_HROP= df_HROP[['GRUPO', 'CONTADOR', 'OPERACIÓN',
                  'DocumentoMAF','Cantidad_OP',	'Duración_OP',
                  'Trabajo_OP',	'EstadoInstalación',
                  'PuestoTrabajo',	'Descripción']]
df_HROP = df_HROP.drop_duplicates()

df_PPM = df_PPM.merge(
    df_UT_Filter,
    on = ['Ubicación técnica'],
    how = 'inner') # Posiciones con las UT
df_PPM = df_PPM.drop_duplicates()

df_PM = df_PM.merge(
    df_PPM['Plan mant.preventivo'].drop_duplicates().reset_index(drop=True),
    on = ['Plan mant.preventivo'],
    how = 'inner') # IP16 + IP18
df_PM = df_PM.drop_duplicates()

df_HR = df_HR.merge(
    df_PPM[['GRUPO','CONTADOR']].drop_duplicates().reset_index(drop=True),
    on = ['GRUPO','CONTADOR'],
    how = 'inner') # IA09 + IP18

df_HROP = df_HROP.merge(
    df_HR[['GRUPO','CONTADOR']],
    on = ['GRUPO','CONTADOR'],
    how = 'inner') # IA09 + IA39

df_HROP = df_HROP.merge(
    df_CR05[['PuestoTrabajo','PuestoTrabajoDescripcion']],
    on = ['PuestoTrabajo'],
    how = 'inner')
df_HROP = df_HROP.drop_duplicates()

df_OT = df_OT.merge(
    df_PPM[['GRUPO','CONTADOR','Plan mant.preventivo','Posición mantenim.']],
    on = ['GRUPO','CONTADOR','Plan mant.preventivo','Posición mantenim.'],
    how = 'inner')
df_OT = df_OT.drop_duplicates()

df_OTOP = df_OTOP.merge(
    df_PPM[['GRUPO','CONTADOR','Plan mant.preventivo','Posición mantenim.']],
    on = ['GRUPO','CONTADOR','Plan mant.preventivo','Posición mantenim.'],
    how = 'inner')
df_OTOP = df_OTOP.drop_duplicates()

df_OT   = acotar_df_por_fecha(df_OT  ,'Fin programado','2023-01-01','2025-06-06')
df_OTOP = acotar_df_por_fecha(df_OTOP,'Fin programado','2023-01-01','2025-06-06')

df_OT  ["Status_Sis"]    = df_OT  ["Status sistema"].str.split(" ").str[0]
df_OTOP["Status_Sis"]    = df_OTOP["Status sistema"].str.split(" ").str[0]
df_OTOP["Status_Sis_Op"] = df_OTOP["Stat.sist.op."] .str.split(" ").str[0]

df_cuenta_ots = df_OTOP.groupby(['GRUPO','CONTADOR','OPERACIÓN','Plan mant.preventivo','Posición mantenim.']).agg(
    USO            = ('Orden', 'count'),
    Fech_Últi_uso  = ('Fin programado', 'max'),
    ).reset_index()
df_filtrado    = df_OTOP[df_OTOP['Trabajo real'] != 0]

df_metrica_fil = df_filtrado.groupby(['GRUPO','CONTADOR','OPERACIÓN','Plan mant.preventivo','Posición mantenim.']).agg(

    Max_Cantidad     = ('Cantidad_Plan', 'max'),
    Min_Cantidad     = ('Cantidad_Plan', 'min'),
    Avg_Cantidad     = ('Cantidad_Plan', 'mean'),

    Max_Duracion     = ('Duración normal', 'max'),
    Min_Duracion     = ('Duración normal', 'min'),
    Avg_Duracion     = ('Duración normal', 'mean'),

    Max_Trabajo      = ('Trabajo_Plan', 'max'),
    Min_Trabajo      = ('Trabajo_Plan', 'min'),
    Avg_Trabajo      = ('Trabajo_Plan', 'mean'),

    Min_HH_Real = ('Trabajo real', 'min'),
    Max_HH_Real = ('Trabajo real', 'max'),
    Prom_HH_Real = ('Trabajo real', 'mean'),

).reset_index()
df_OTOP_grouped = df_cuenta_ots.reset_index().merge(
                                                    df_metrica_fil.reset_index(),
                                                    on=['GRUPO', 'CONTADOR', 'OPERACIÓN','Plan mant.preventivo','Posición mantenim.'],
                                                    how='left'
                                                    ).drop_duplicates()

cols_to_round = [
    'Avg_Cantidad', 'Avg_Duracion', 'Avg_Trabajo', 'Prom_HH_Real'
]
df_OTOP_grouped[cols_to_round]  = df_OTOP_grouped[cols_to_round].round(2)

df_OTOP_grouped['Fech_Últi_uso'] = df_OTOP_grouped['Fech_Últi_uso'].dt.date

df_HR['Fe.modificación']        = pd.to_datetime(df_HR['Fe.modificación']).dt.date
df_HR['Creado el']              = pd.to_datetime(df_HR['Creado el']).dt.date

df_HROP['DocumentoMAF']         = df_HROP['DocumentoMAF'].astype(str)


# Generando reporte

# Posiciones
df_result = df_PPM[['GRUPO','CONTADOR','Plan mant.preventivo',
                    'Posición mantenim.','Ubicación técnica',
                    'TxtPosicManten',
                    'Estrategia', 'Campo clasif.']].rename(columns={'Campo clasif.':'TAG'
                                                                       })

# + Planes
df_result = df_result.merge(
    df_PM[['Plan mant.preventivo','Txt plan MT']],
    on = ['Plan mant.preventivo'],
    how = 'inner') # Posiciones + Planes ID: Plan de mant

# + Hojas de ruta
df_result = df_result.merge(
    df_HR[['GRUPO','CONTADOR','Estado',
           'Descripción','Fe.modificación','Creado el']].rename(columns={'Estado':'EstadoInstalacion_HR',
                                                                       'Descripción':'Descripcion_HR',
                                                                       'Fe.modificación':'Fecha_Modif_HR' ,
                                                                       'Creado el':'Fecha_Creada_HR'
                                                                       }),
    on = ['GRUPO','CONTADOR'],
    how = 'inner')

# + Operaciones de hojas de ruta
df_result = df_result.merge(
    df_HROP[['GRUPO', 'CONTADOR', 'OPERACIÓN', 'DocumentoMAF', 'Cantidad_OP','Duración_OP', 'Trabajo_OP', 'EstadoInstalación', 'PuestoTrabajo',
           'Descripción', 'PuestoTrabajoDescripcion']].rename(columns={'Descripción':'Descripcion_HROP'}),
    on = ['GRUPO', 'CONTADOR'],
    how = 'inner')
df_result = df_result.drop_duplicates()

# + OTs a nivel de operación
df_result_1 = df_result.merge(
    df_OTOP,
    on = ['GRUPO', 'CONTADOR', 'OPERACIÓN', 'Plan mant.preventivo', 'Posición mantenim.'],
    how = 'left')


# + Cuenta de OTS a nivel de operación
df_result_2 = df_result.merge(
    df_OTOP_grouped,
    on = ['GRUPO', 'CONTADOR', 'OPERACIÓN', 'Plan mant.preventivo', 'Posición mantenim.'],
    how = 'left')


df_result_1 = df_result_1[ [
    'Plan mant.preventivo', 'Posición mantenim.', 'GRUPO', 'CONTADOR', 'Descripcion_HR', 'TAG',
    'Ubicación técnica', 'Equipo', 'OPERACIÓN', 'TxtPosicManten', 'Estrategia',
    'PuestoTrabajo', 'PuestoTrabajoDescripcion', 'Fecha_Creada_HR', 'Fecha_Modif_HR',
    'DocumentoMAF', 'Descripcion_HROP', 'Pto.tbjo.resp.', 'Pto.tbjo.op.', 'Cantidad_Plan',
    'Trabajo_Plan', 'Duración normal', 'Cantidad real',
    'Trabajo real', 'Descripción de OTOP', 'Texto breve', 'Txt.brv.oper.', 'Duración mínima',
    'Txt plan MT', 'EstadoInstalacion_HR', 'Status_Sis', 'Status_Sis_Op'
]]

df_result_2 = df_result_2[ [
    'TAG','Ubicación técnica','Descripcion_HR','OPERACIÓN','Txt plan MT','TxtPosicManten','PuestoTrabajo', 'PuestoTrabajoDescripcion','Cantidad_OP',
    'Duración_OP', 'Trabajo_OP','USO', 'Min_HH_Real' , 'Max_HH_Real' ,'Prom_HH_Real' , 'Fech_Últi_uso','DocumentoMAF', 'Estrategia','EstadoInstalacion_HR',
    'Plan mant.preventivo', 'Posición mantenim.', 'GRUPO', 'CONTADOR',

]]


output_path = os.path.join(path_resultado, "Frecuencia_de_OTs.xlsx")
with pd.ExcelWriter(output_path, engine = 'xlsxwriter') as writer:
    formatear_hoja_excel( 
    writer=writer,
    df=df_result_1,
    nombre_hoja="Total de OTs",
    nombre_tabla="Tabla_Total_OTs",

)
    formatear_hoja_excel(
    writer=writer,
    df=df_result_2,
    nombre_hoja="Estadísticas",
    nombre_tabla="Estadísticas_OTs",

)
