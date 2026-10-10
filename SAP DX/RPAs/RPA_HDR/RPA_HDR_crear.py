from dotenv import load_dotenv
from ventana import pantalla_final
import threading
import os, time, subprocess
import pandas as pd
import win32com.client
import numpy as np
import pygetwindow as gw
import pyautogui
import datetime
#Carga variables de .env
load_dotenv()


sap_logon_path      = r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"
PRT_path_base       = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.01 EJECUCION TI\05. RPA EXEC\04. HDR"

# Variables constantes
NOMBRE_CONEXION_SAP = "Copper SAP ECC Producción(PS8)"
RSYST_MANDT         = "300"
RSYST_BNAME         = os.getenv("SAP_USER")
RSYST_BCODE         = os.getenv("SAP_PASS")
RSYST_LANGU         = "ES"
TRANSACCION_IA05    = "/NIA05"
TRANSACCION_IA06    = "/NIA06"
FILTRO_IA06_PERFIL  = "0000008"
MENSAJE_ERROR       = ""
# Variables de configuración
PRT_CarpetaAProcesar     = "REQ-009 Solicitud 04_20250820 (CREACION)"
PRT_ArchivoExcelNombre   = "DATOS.xlsx"
ArchivoHojaNombre        = "HDR"
IA10_RutaNombre          = "IA10.xlsx"
VENTANA_BORRAR_MAF       = "Borrar asignaciones MAF"
PATH_EXCEL_DATOS         = os.path.join(PRT_path_base, PRT_CarpetaAProcesar, PRT_ArchivoExcelNombre)
DESCARGA_PATH_REQUERI    = os.path.join(PRT_path_base, PRT_CarpetaAProcesar)


COLUMNA_GRUPO            = "GRUPO"

COLUMNAS_DATA_CREACION       = [
                          "DESCRIPCIÓN HDR",
                          "CENTRO",
                          "PtoTjoResponsable",
                          "Ctro.pto.trab.",
                          "Utilización",
                          "Grupo planif.",
                          "Estado global",
                          "Estado instal.",
                          "CODIGO ESTRATEGIA",
                          "Conjunto",

                          "PtoTrabajo",
                          "CANTIDAD",
                          "DURACIÓN",

                          "PAQUETE",

                          "DESCRIPCIÓN OPERACIÓN"
                      ]



def transaccion_IA05(GRUPO, CREACION_DATOS):
  # Filtro de transacción 
  session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_IA05
  session.findById("wnd[0]").sendVKey(0)

  if pd.isna(GRUPO):  
    session.findById("wnd[0]/usr/ctxt[0]").text = ""
    session.findById("wnd[0]/usr/ctxt[1]").text = FILTRO_IA06_PERFIL
    session.findById("wnd[0]").sendVKey(0)

  else:
    session.findById("wnd[0]/usr/ctxt[0]").text = GRUPO
    session.findById("wnd[0]/usr/ctxt[1]").text = FILTRO_IA06_PERFIL
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[6]").press() 
  print(CREACION_DATOS[COLUMNAS_DATA_CREACION[0]].iloc[0])
  contador = session.findById("wnd[0]/usr/txt[2]").text
  session.findById("wnd[0]/usr/txt[3]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[0]].iloc[0]#Descripción
  session.findById("wnd[0]/usr/ctxt[1]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[1]].iloc[0] # Centro
  session.findById("wnd[0]/usr/ctxt[2]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[2]].iloc[0]# Puesto de trabajo
  session.findById("wnd[0]/usr/ctxt[3]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[3]].iloc[0] # Centro psto trabajo
  session.findById("wnd[0]/usr/ctxt[4]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[4]].iloc[0] # Estado instalación
  session.findById("wnd[0]/usr/ctxt[5]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[5]].iloc[0] # Grupo de plan
  session.findById("wnd[0]/usr/ctxt[6]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[6]].iloc[0] #Estado 
  session.findById("wnd[0]/usr/ctxt[7]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[7]].iloc[0] # Estado instalación
  session.findById("wnd[0]/usr/ctxt[8]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[8]].iloc[0] # Estrategia
  session.findById("wnd[0]/usr/ctxt[9]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[9]].iloc[0] # Conjunto
  session.findById("wnd[0]/tbar[1]/btn[16]").press()

  session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3400/ctxt[2,0]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[10]].iloc[0]# Psto Trabajo
  session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3400/txt[5,0]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[14]].iloc[0]# Descripción
  session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3400/txt[9,0]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[11]].iloc[0] # Cantidad 
  session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3400/txt[10,0]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[12]].iloc[0] # Duración
  session.findById("wnd[0]").sendVKey(0)
  session.findById("wnd[0]").sendVKey(0)

  session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3400").getAbsoluteRow(0).selected = True
  session.findById("wnd[0]/usr/btn[0]").press()
  session.findById("wnd[0]/tbar[1]/btn[26]").press()
  session.findById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txt[0,0]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[13]].iloc[0]
  session.findById("wnd[0]").sendVKey(0)
#CTRL + S 
  session.findById("wnd[0]/tbar[0]/btn[11]").press()

  # session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_IA06
  grupo = session.findById("wnd[0]/usr/ctxt[0]").text
  return grupo, contador

def SAP_start():
  if not os.path.exists(sap_logon_path):
          raise FileNotFoundError("Ruta de SAP inválida")
  subprocess.Popen(sap_logon_path)
  time.sleep(3)
  
  try:
          SapGuiAuto  = win32com.client.GetObject('SAPGUI')
          application = SapGuiAuto.GetScriptingEngine
  except Exception as e:
          raise RuntimeError("[ERROR] Error al cargar SAP GUI scripting") from e
  
  # Inicialización conexión
  try:
      connection = application.OpenConnection(NOMBRE_CONEXION_SAP, True)
      connection = application.Children(0)
      session = connection.Children(0)
      return session
  except Exception as e:
      print(f"[ERROR] Error al abrir conexión SAP: {e}")
      exit()

def eliminar_espacio(df):
  for columna in df.columns:
    df[columna] = df[columna].apply(lambda x: x.strip() if isinstance(x,str) else x)
    # df[columna] = df[columna].apply(lambda x: x.replace("  "," ") if isinstance(x,str) else x)
  return df 
  
if __name__=="__main__":
  print("Hora de inicio: ",datetime.datetime.now())
  
  # Lectura de Excel con requerimientos
  if not os.path.exists(PATH_EXCEL_DATOS):
    raise FileNotFoundError(f"Ruta de Excel no encontrada: {PATH_EXCEL_DATOS}")
  
  try:
    df_req  =  pd.read_excel(PATH_EXCEL_DATOS, sheet_name = ArchivoHojaNombre)
    df_req  =  df_req.replace(r'^\s*$',np.nan,regex=True).infer_objects(copy=False)
    COLUMNAS_DE_CAMBIO = df_req.columns.to_list() 
  except Exception as e:
    print("No encuentra el nombre de la hoja de Excel")


  DF_ANALIZAR = df_req
  DF_ANALIZAR = eliminar_espacio(DF_ANALIZAR)

# ------------------- Inicialización de SAP -----------------
  session = SAP_start()
      
  # Registro de cuenta SAP
  session.findById("wnd[0]/usr/txtRSYST-MANDT").text = RSYST_MANDT
  session.findById("wnd[0]/usr/txtRSYST-BNAME").text = RSYST_BNAME
  session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = RSYST_BCODE
  session.findById("wnd[0]/usr/txtRSYST-LANGU").text = RSYST_LANGU
  session.findById("wnd[0]").sendVKey(0)

  
  for index in range(len(DF_ANALIZAR)):
    
    if DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] == "ERROR_NO_MATCH":
      continue
    
    MENSAJE_ERROR = ""
    CAMBIAR_DATOS = DF_ANALIZAR.iloc[[index]][COLUMNAS_DE_CAMBIO]
    GRUPO         = DF_ANALIZAR.loc[index, COLUMNA_GRUPO]
    try:
      grupo, contador = transaccion_IA05( GRUPO, CAMBIAR_DATOS)  
    except Exception as e:
      print(e)
      contador = ""
      grupo = ""
    DF_ANALIZAR.loc[index, "CONTADOR"] = contador 
    DF_ANALIZAR.loc[index, COLUMNA_GRUPO] = grupo 

    DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] = "OK" if MENSAJE_ERROR == "" else MENSAJE_ERROR 
  DF_ANALIZAR = DF_ANALIZAR[[
                            "DESCRIPCIÓN HDR",
                            "CONTADOR",
                            "GRUPO",
                            "ESTADO_EJECUCION_RPA",
                            ]]
  with pd.ExcelWriter(os.path.join(DESCARGA_PATH_REQUERI,'Resultados.xlsx')) as writer:
    DF_ANALIZAR.to_excel(writer, sheet_name="HojaRuta_ANALIZADOS",index=False)
    
  print(f"Hora de fin: {datetime.datetime.now()}")
  pantalla_final()