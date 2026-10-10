import pandas as pd
from Formateo_IA10 import operacion
import os 
ruta_path = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.02 EJECUCION FAST TRACK\250819 REQ FT 4\2. Entregables"

IP16     = os.path.join(ruta_path,"IP16.xlsx")
IP18     = os.path.join(ruta_path,"IP18.xlsx")
IA09     = os.path.join(ruta_path,"IA09.xlsx")
IA10_OP  = os.path.join(ruta_path,"IA10_OP.xlsx")
IA10_PAQ = os.path.join(ruta_path,"IA10_PAQ.xlsx")
IA10_MAF = os.path.join(ruta_path,"IA10_MAF.xlsx")

dtype_common = {'Plan mantenimiento': 'string'}
df_IP16     = pd.read_excel(IP16, dtype=dtype_common)
df_IP18     = pd.read_excel(IP18, dtype=dtype_common)
df_IA09     = pd.read_excel(IA09)
df_IA10_OP  = pd.read_excel(IA10_OP)
df_IA10_PAQ = pd.read_excel(IA10_PAQ)
df_IA10_MAF = pd.read_excel(IA10_MAF)

# ----- CAMPOS ----- #
Columnas_IP16 = [
                    "Plan mantenimiento",
                    "Txt plan mantenim.",
                    "Estrategia mantenim.",

                    # "Pos.mantenim.",
                    
                    "Fact.dec.conclusión retr.",
                    "Tolerancia (+)",
                    "Fact.dec.concl.anticipada",
                    "Tolerancia (-)",
                    "Calendario fábrica",
                    "Horizonte apertura",
                    "Intervalo de toma",
                    "Inicio de ciclo"
                ]
 
Columnas_IP18 = [
                    "Posición mantenim.",
                    "Plan mantenimiento",
                    "Txt.pos.mantenim.",
                    "Estrategia mantenim.",
                    "Ubicación técnica",
                    "Centro planificación",
                    "Clase de orden",
                    "Pto.tbjo.responsable",
                    "Prioridad",
                    "Grupo planificación",
                    "Clase actividad PM",
                    #"Tipo HojaRuta",
                    "Grupo hojas ruta",
                    "Cont.grupo HRuta",
                    "Impacto"
                ]

Columnas_HDR = [
                # HDR Cabecera
                    "Descripción",
                    "Grupo hojas ruta",
                    "Cont.grupo HRuta",
                    "Ctro.pto.trab.",
                    "Puesto de trabajo",
                    "Utilización",
                    "Grupo planif.",
                    "Estado global",
                    "Estado instalación",
                    "Estrategia mantenim.",
                    "Conjunto"
]

Columnas_HDR_OP = [
                # HDR Operación
                    "Operación",
                    "Descripción_OP",
                    "Estado instalación",
                    "Pto.tbjo.operación",
                    "Cantidad",
                    "Duración",
                    "Trabajo",

                # HDR Frecuencia 
                    "Paquete",
                
                # HDR MAF
                    "MAF"
                ] 

# ----- FORMATEO ----- #
# df_IA10_MAF = operacion(df_IA10_MAF)
df_IA10_OP = operacion(df_IA10_OP)

# ----- REPORTE ----- #

PM_PPM = pd.merge(df_IP16[Columnas_IP16], df_IP18[Columnas_IP18].drop(columns=["Estrategia mantenim."]), on = "Plan mantenimiento", how = 'inner')
PM_PPM["Tipo HojaRuta"] = "A"
PM  = PM_PPM[Columnas_IP16 + [ "Posición mantenim."]]
PPM = PM_PPM[Columnas_IP18]

HDR_PPM = pd.merge(df_IP18[["Grupo hojas ruta","Cont.grupo HRuta"]], df_IA09[Columnas_HDR], on = ["Grupo hojas ruta","Cont.grupo HRuta"], how = 'inner')
HDR_PPM = HDR_PPM[Columnas_HDR]

HDR_OP  = pd.merge(HDR_PPM, df_IA10_OP, on = ["Grupo hojas ruta","Cont.grupo HRuta"], how = 'inner')
print(HDR_OP[HDR_OP["Descripción"]=="2W/O PREV INST SENSOR NIVEL 4110LN001"])
HDR_OP  = pd.merge(HDR_OP, df_IA10_PAQ, on = ["Grupo hojas ruta","Cont.grupo HRuta","Operación"], how = 'left') #falta formatear IA10 PAQ
HDR_OP  = pd.merge(HDR_OP, df_IA10_MAF, on = ["Grupo hojas ruta","Cont.grupo HRuta"], how = 'left')

HDR_OP  = HDR_OP[Columnas_HDR + Columnas_HDR_OP]



HDR_OP_UT = pd.merge(HDR_OP, PPM[["Grupo hojas ruta","Cont.grupo HRuta","Ubicación técnica"]], on = ["Grupo hojas ruta","Cont.grupo HRuta"], how = "left")
PM_UT     = pd.merge(PM, PPM[["Posición mantenim.","Ubicación técnica"]], on = "Posición mantenim.", how = "left")


# # Conteos en cada dataframe
# conteo_HDR_OP = HDR_OP_UT["Ubicación técnica"].value_counts().rename("HDR_OP_count")
# conteo_HDR = (
#     HDR_OP_UT.drop_duplicates(subset=["Ubicación técnica", "Descripción"])
#              .groupby("Ubicación técnica")["Descripción"]
#              .nunique()
#              .rename("HDR_count")
# )
# conteo_PM  = PM_UT["Ubicación técnica"].value_counts().rename("PM_count")
# conteo_PPM = PPM["Ubicación técnica"].value_counts().rename("PPM_count")

# # Unir todo en un DataFrame con todas las ubicaciones únicas
# Resumen = (
#     PPM["Ubicación técnica"].drop_duplicates().reset_index(drop=True).to_frame()
#     .merge(conteo_HDR, left_on="Ubicación técnica", right_index=True, how="left")
#     .merge(conteo_HDR_OP, left_on="Ubicación técnica", right_index=True, how="left")
#     .merge(conteo_PM,  left_on="Ubicación técnica", right_index=True, how="left")
#     .merge(conteo_PPM, left_on="Ubicación técnica", right_index=True, how="left")
#     .fillna(0)
#     .astype(int)
# )

with pd.ExcelWriter(os.path.join(ruta_path,"250820_ReporteSAP_Solicitud_FastTrack_004_v2.xlsx")) as writer:
    # Resumen.to_excel(writer, sheet_name="Resumen",index=False)
    HDR_OP_UT.to_excel(writer, sheet_name="HDR",index=False)
    PPM.to_excel(writer, sheet_name="PPM",index=False)
    PM_UT.to_excel(writer, sheet_name="PM",index=False)




