from dotenv import load_dotenv
# from ventana import pantalla_final
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
PRT_path_base       = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.01 EJECUCION TI\05. RPA EXEC\06. PPM"

# Variables constantes
NOMBRE_CONEXION_SAP = "Copper SAP ECC Producción(PS8)"
RSYST_MANDT         = "300"
RSYST_BNAME         = os.getenv("SAP_USER")
RSYST_BCODE         = os.getenv("SAP_PASS")
RSYST_LANGU         = "ES"
TRANSACCION_IP04    = "/NIP04"
TRANSACCION_IP05    = "/NIP05"
TP_PLAN_MANTTO      = "PM"
MENSAJE_ERROR       = ""
IMPACTO_AAQ         = "3"

# Variables de configuración
PRT_CarpetaAProcesar     = "REQ-011 Solicitud 4_20250820 (CREACION)"
PRT_ArchivoExcelNombre   = "DATOS.xlsx"
ArchivoHojaNombre        = "PPM"

VENTANA_BORRAR_MAF       = "Borrar asignaciones MAF"
PATH_EXCEL_DATOS         = os.path.join(PRT_path_base, PRT_CarpetaAProcesar, PRT_ArchivoExcelNombre)
DESCARGA_PATH_REQUERI    = os.path.join(PRT_path_base, PRT_CarpetaAProcesar)


COLUMNA_ESTRATEGIA       = "Maintenance Strategy"

COLUMNAS_DATA_CREACION       = [
                          "TxtPosicManten",
                          "Ubicacion",
                          "Centr planificacion",
                          "Clase orden",
                          "Main Work Center",
                          "Centro",
                          "Prioridad",

                          "Grupo planificador",
                          "Clase actividad",
                          "Tipo HojaRuta",
                          "Grupo",
                          "Contador",
                          "Impacto"

                      ]

def transaccion_IP04(ESTRATEGIA, CREACION_DATOS):
  # Filtro de transacción 
  session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_IP04
  session.findById("wnd[0]").sendVKey(0)

     
  session.findById("wnd[0]/usr/cmb").key = TP_PLAN_MANTTO
  session.findById("wnd[0]/usr/ctxt").text = ESTRATEGIA
  session.findById("wnd[0]").sendVKey(0)

  session.findById("wnd[0]/usr/sub/1/txt[1]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[0]].iloc[0]
  print(CREACION_DATOS[COLUMNAS_DATA_CREACION[0]].iloc[0])
  session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/2/ctxt[0]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[1]].iloc[0]
  session.findById("wnd[0]").sendVKey(0)
  session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/4/ctxt[0]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[2]].iloc[0]
  session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/4/ctxt[2]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[3]].iloc[0]
  session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/4/ctxt[4]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[4]].iloc[0]
  session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/4/ctxt[5]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[5]].iloc[0]
  session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/4/ctxt[1]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[7]].iloc[0]
  session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/4/ctxt[3]").text = "00"+str(CREACION_DATOS[COLUMNAS_DATA_CREACION[8]].iloc[0])
  session.findById("wnd[0]").sendVKey(0)
  session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/4/cmb").key = CREACION_DATOS[COLUMNAS_DATA_CREACION[6]].iloc[0]
  session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/4/txt[8]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[9]].iloc[0]
  session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/4/txt[9]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[10]].iloc[0]
  session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/4/txt[10]").text = CREACION_DATOS[COLUMNAS_DATA_CREACION[11]].iloc[0]
  session.findById("wnd[0]").sendVKey(0)

  session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\17").select()
  session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\17/ssub/3/1/2/sub/3/1/2/1/ctxt[1]").text = IMPACTO_AAQ
  
#CTRL + S 
  session.findById("wnd[0]/tbar[0]/btn[11]").press()

  session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_IP05
  session.findById("wnd[0]").sendVKey(0)
  ulti_posicion = session.findById("wnd[0]/usr/ctxt").text
  return ulti_posicion

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
    ESTRATEGIA    = DF_ANALIZAR.loc[index, COLUMNA_ESTRATEGIA]

    try:
      ulti_pos = transaccion_IP04( ESTRATEGIA, CAMBIAR_DATOS)  
    except:
      ulti_pos = ""
    DF_ANALIZAR.loc[index,"Posición"] = ulti_pos
    print(DF_ANALIZAR.loc[index,"Posición"])
    DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] = "OK" if MENSAJE_ERROR == "" else MENSAJE_ERROR 
  DF_ANALIZAR = DF_ANALIZAR[[
                            "DESCRIPCIÓN DE LA OPERACIÓN",
                            "Grupo",
                            "Contador",
                            "Posición",
                            "ESTADO_EJECUCION_RPA",
                            # "cantidad_match",
                            # "COLUMNA_MATCH",
                            ]]
  with pd.ExcelWriter(os.path.join(DESCARGA_PATH_REQUERI,'Resultados.xlsx')) as writer:
    DF_ANALIZAR.to_excel(writer, sheet_name="DocumentosMAF_ANALIZADOS",index=False)
    
  print(f"Hora de fin: {datetime.datetime.now()}")
  # pantalla_final()