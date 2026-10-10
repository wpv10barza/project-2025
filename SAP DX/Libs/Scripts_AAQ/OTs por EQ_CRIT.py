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
path_archivos  = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION SEMANAL\04. RQ TI Dev\00. Data\SAP\07. SAP 134_EQ_CRIT 2025.07.30"
path_resultado = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION SEMANAL\04. RQ TI Dev\00. Data\SAP\07. SAP 134_EQ_CRIT 2025.07.30"

IP16 = os.path.join(path_archivos, "IP16.xlsx")
IP18 = os.path.join(path_archivos,  "IP18.xlsx")
IA09 = os.path.join(path_archivos,  "IA09.xlsx")
IW39 = os.path.join(path_archivos,  "IW39.xlsx")
CR05 = os.path.join(path_archivos, "CR05.xlsx")

Lista_Equipos = [
    "2610-EL-001", "3140-CV-002-CH1", "3140-CV-002-CH2", "3140-CV-002-CH3", "3210-BX-001", "3210-CH-039", "3210-CH-045", "3210-CH-047",
    "3210-MI-001-CH1", "3210-SN-001", "3210-SN-003", "3210-CH-012", "3210-CH-011", "3210-CH-040", "3210-CH-046", "3210-CH-048",
    "3210-HO-003", "3210-MI-004-CH1", "3210-MI-CS2", "3210-SN-002", "3210-EL-001", "3210-ER-001", "3220-CV-001", "3220-CV-003",
    "3220-CV-001-CH1", "3220-CV-003-CH1", "3220-CV-009-CH1", "3220-CV-009-CH2", "3220-CV-010-CH1", "3220-CV-010-CH2",
    "3310-BX-030", "3310-LN-010", "3310-LN-021", "3310-ER-001", "3410-BG-002", "3410-FT-001", "3410-FT-002", "3410-FT-003",
    "3410-FT-004", "3410-FT-005", "3410-FT-006", "3410-FT-007", "3410-FT-008", "3410-FT-009", "3410-FT-010", "3410-FT-011",
    "3410-FT-012", "3410-FT-013", "3410-FT-014", "3410-FT-015", "3410-FT-016", "3410-FT-017", "3410-FT-018", "3410-FT-023",
    "3410-FT-024", "3410-FT-025", "3410-FT-026", "3410-FT-027", "3410-FT-028", "3410-FT-029", "3410-FT-030", "3420-ER-001",
    "3710-LN-005", "3710-TH-001", "3710-TH-002", "5810-ER-001", "5810-ER-002", "5840-EE-001", "5840-SL-001", "4110-LN-002",
    "4330-ER-001", "2630-TF-003", "2720-TF-002", "2730-TF-001", "2730-TF-002", "4310-TF-020", "5110-ER-002", "5110-HA-001",
    "5110-HA-002", "5110-HA-003", "3210-TF-001", "3210-TF-002", "3210-TF-011", "3210-TF-012", "3210-TF-113", "3210-TF-114",
    "3210-TF-121", "3210-TF-122", "3210-TF-123", "3210-TF-024", "3210-TF-131", "3210-TF-132", "3210-TF-133", "3210-TF-134",
    "3210-TF-141", "3210-TF-142", "3210-TF-143", "3210-TF-144", "3220-TF-001", "3220-TF-003", "3320-TF-003", "3740-TF-001",
    "3740-TF-002"
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
df_OT        = pd.read_excel(IW39)
df_CR05      = pd.read_excel(CR05)


###############################################
df_UT_Filter = pd.DataFrame({'Campo clasif.':Lista_Equipos})

df_PPM  = df_PPM.rename (columns={'Gpo.hojas ruta':"GRUPO",
                                  'Cont.gpo.HRuta':"CONTADOR",
                                  'Ubicac.técnica':'Ubicación técnica',
                                  'Plan mantenim.':'Plan mant.preventivo',
                                  'Pos.mantenim.':'Posición mantenim.'})

df_PM  = df_PM.rename (columns={'Plan mantenim.':'Plan mant.preventivo',
                                'Pos.mantenim.':'Posición mantenim.'})

df_HR   = df_HR.rename  (columns={"Gpo.hojas ruta":"GRUPO",
                                  "Cont.gpo.HRuta":"CONTADOR",
                                  "Puesto trabajo":"PuestoTrabajo"})


df_CR05 = df_CR05.rename(columns={"PtoTbjo":"PuestoTrabajo",
                                  "Descripción":"PuestoTrabajoDescripcion"})


df_OT   = df_OT.rename  (columns={"Gpo.hojas ruta":"GRUPO",
                                  "Cont.gpo.HRuta":"CONTADOR",
                                  "Plan mantenim.":"Plan mant.preventivo",
                                  "Pos.mantenim.":"Posición mantenim.",
                                  "Texto breve": "Descripción OT"})

df_UT_Filter.name = "Campo clasif."



df_PPM = df_PPM.merge(
    df_UT_Filter,
    on = ['Campo clasif.'],
    how = 'inner') 


df_PM = df_PM.merge(
    df_PPM['Plan mant.preventivo'],
    on = ['Plan mant.preventivo'],
    how = 'inner') 


df_HR = df_HR.merge(
    df_PPM[['GRUPO','CONTADOR']],
    on = ['GRUPO','CONTADOR'],
    how = 'inner')


df_HR = df_HR.merge(
    df_CR05[['PuestoTrabajo','PuestoTrabajoDescripcion']],
    on = ['PuestoTrabajo'],
    how = 'inner')
df_HR = df_HR.drop_duplicates()

df_OT = df_OT.merge(
    df_PPM[['GRUPO','CONTADOR','Plan mant.preventivo','Posición mantenim.']],
    on = ['GRUPO','CONTADOR','Plan mant.preventivo','Posición mantenim.'],
    how = 'inner')
df_OT = df_OT.drop_duplicates()


# df_OT   = acotar_df_por_fecha(df_OT  ,'Fin programado','2023-01-01','2025-06-06')
df_OT  ["Status_Sis"]    = df_OT  ["Status sistema"].str.split(" ").str[0]





# Generando reporte

# Posiciones
df_result = df_PPM[['GRUPO','CONTADOR','Plan mant.preventivo',
                    'Posición mantenim.','Ubicación técnica',
                    'TxtPosicManten',
                    'Estrategia', 'Campo clasif.']].rename(columns={'Campo clasif.':'TAG'
                                                                       })

# + Planes
df_result = df_result.merge(
    df_PM[['Plan mant.preventivo','Txt plan MT',"Status sistema","Horizonte-apert",
           "Estrategia","Ctd.posiciones"]].rename(columns={#'Estado':'EstadoInstalacion_HR',
                                                                       'Status sistema':'Status Sis Plan',
                                                                       'Estrategia':'Estrategia_Plan' ,
                                                                       'Ctd.posiciones':'Cantidad de posiciones'
                                                                       }),
    on = ['Plan mant.preventivo'],
    how = 'inner') # Posiciones + Planes ID: Plan de mant

# + Hojas de ruta
df_result = df_result.merge(
    df_HR[['GRUPO','CONTADOR',#'Estado',
           'Descripción','Fe.modificación','Creado el',
           "PuestoTrabajoDescripcion",'PuestoTrabajo']].rename(columns={#'Estado':'EstadoInstalacion_HR',
                                                                       'Descripción':'Descripcion_HR',
                                                                       'Fe.modificación':'Fecha_Modif_HR' ,
                                                                       'Creado el':'Fecha_Creada_HR'
                                                                       }),
    on = ['GRUPO','CONTADOR'],
    how = 'inner')

# + OTs a nivel de operación
df_result = df_result.merge(
    df_OT,
    on = ['GRUPO', 'CONTADOR', 'Plan mant.preventivo', 'Posición mantenim.'],
    how = 'inner')

print(df_result.columns.to_list)
df_result = df_result.drop_duplicates()

df_result = df_result[[

'TAG','Txt plan MT','Plan mant.preventivo','Status Sis Plan','Estrategia_Plan' ,'Cantidad de posiciones',
'TxtPosicManten','Posición mantenim.','Descripcion_HR', "PuestoTrabajoDescripcion",'PuestoTrabajo',"Grupo planif.", 'GRUPO', 'CONTADOR',
    "Orden","Descripción OT", 'Estado instal.', 'Res./SolPed', 'Fe.inicio real', 'Hora in.real',
       'Liberación real', 'Fin programado', 'Fin (hora)', 'Fecha fin real',
       'Fin real (hora)','Cl.actividad PM', 'TotalCostesPlan', 'Cst.tot.reales',
       'IndPlanifOrden', 'Parada', 'Pet.borrado', 'Status_Sis'

#     # 'Plan mant.preventivo', ,  ,"Grupo planif."
#     # 'Ubicación técnica', 'Equipo',   'Estrategia',
#     # , 'PuestoTrabajoDescripcion', 'Fecha_Creada_HR', 'Fecha_Modif_HR',
#     # 'DocumentoMAF', 'Pto.tbjo.resp.', 'Pto.tbjo.op.', 'Cantidad_Plan',
#     # 'Trabajo_Plan' , 'Texto breve', 
#     #  'EstadoInstalacion_HR', 'Status_Sis', 'Status_Sis_Op'
]]


# for col in ['TAG','Txt plan MT','Estrategia_Plan','TxtPosicManten','Descripcion_HR',"PuestoTrabajoDescripcion",'PuestoTrabajo',"Grupo planif." ]:
#     df_result[col] = df_result[col].where(~df_result[col].duplicated(), '')



output_path = os.path.join(path_resultado, "Frecuencia_de_OTs_V3.xlsx")
with pd.ExcelWriter(output_path, engine = 'xlsxwriter') as writer:
    formatear_hoja_excel( 
    writer=writer,
    df=df_result,
    nombre_hoja="Total de OTs",
    nombre_tabla="Tabla_Total_OTs",
)

