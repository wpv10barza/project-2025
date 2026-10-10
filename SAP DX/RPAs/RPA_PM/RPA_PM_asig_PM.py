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
PRT_path_base       = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION SEMANAL\05. RPA EXEC\05. PM"

# Variables constantes
NOMBRE_CONEXION_SAP = "Copper SAP ECC Producción(PS8)"
RSYST_MANDT         = "300"
RSYST_BNAME         = os.getenv("SAP_USER")
RSYST_BCODE         = os.getenv("SAP_PASS")
RSYST_LANGU         = "ES"
TRANSACCION_IP02    = "/NIP02"
TP_PLAN_MANTTO      = "PM"
MENSAJE_ERROR       = ""

# Variables de configuración
PRT_CarpetaAProcesar     = "REQ-007 Solicitud 03_20250731 (ASIGNACION)"
PRT_ArchivoExcelNombre   = "DATOS.xlsx"
ArchivoHojaNombre        = "PM"

VENTANA_BORRAR_MAF       = "Borrar asignaciones MAF"
PATH_EXCEL_DATOS         = os.path.join(PRT_path_base, PRT_CarpetaAProcesar, PRT_ArchivoExcelNombre)
DESCARGA_PATH_REQUERI    = os.path.join(PRT_path_base, PRT_CarpetaAProcesar)


COLUMNA_POS_ASIG       = "Posicion"
COLUMNA_PLAN           = "Plan"

def transaccion_IP02(PLAN, POS):
  global MENSAJE_ERROR 
  # Filtro de transacción 
  session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_IP02
  session.findById("wnd[0]").sendVKey(0)

  session.findById("wnd[0]/usr/ctxt").text = PLAN
  session.findById("wnd[0]").sendVKey(0)
  print(PLAN)
  try:
    session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\10/ssub/3/1/2/sub/3/1/2/1/btn[1]").press()
  except:
     try:
        session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/1/btn[2]").press()
     except Exception as e:
        MENSAJE_ERROR="Error al asignar posición"

  session.findById("wnd[0]/usr/txt[5]").text = POS
  session.findById("wnd[0]/tbar[1]/btn[8]").press()
  
  try:
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
    session.findById("wnd[0]/tbar[1]/btn[42]").press()
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
  except Exception as e:
    MENSAJE_ERROR = "Error con posición"
#CTRL + S 

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
  DF_ANALIZAR["ESTADO_EJECUCION_RPA"] = ""
# ------------------- Inicialización de SAP -----------------
  session = SAP_start()
      
  # Registro de cuenta SAP
  session.findById("wnd[0]/usr/txtRSYST-MANDT").text = RSYST_MANDT
  session.findById("wnd[0]/usr/txtRSYST-BNAME").text = RSYST_BNAME
  session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = RSYST_BCODE
  session.findById("wnd[0]/usr/txtRSYST-LANGU").text = RSYST_LANGU
  session.findById("wnd[0]").sendVKey(0)
  
  # CREACION_DATOS = DF_ANALIZAR.iloc[[1]][COLUMNAS_DATA_CREACION]
  # ESTRATEGIA     = DF_ANALIZAR.loc[1, COLUMNA_ESTRATEGIA]

  # transaccion_IP04(ESTRATEGIA, CREACION_DATOS)  
  
  for index in range(len(DF_ANALIZAR)):
    
    if DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] == "ERROR_NO_MATCH":
      continue
    
    MENSAJE_ERROR = ""
    POS_ASIG      = DF_ANALIZAR.loc[index, COLUMNA_POS_ASIG]
    PLAN          = DF_ANALIZAR.loc[index, COLUMNA_PLAN]

    transaccion_IP02( PLAN,POS_ASIG)  

    
    DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] = "OK" if MENSAJE_ERROR == "" else MENSAJE_ERROR 
  DF_ANALIZAR = DF_ANALIZAR[[
                            COLUMNA_POS_ASIG,
                            COLUMNA_PLAN,
                            "Plan de Mantenimiento",
                            "Posición de Mantenimiento",
                            "ESTADO_EJECUCION_RPA",

                            ]]
  with pd.ExcelWriter(os.path.join(DESCARGA_PATH_REQUERI,'Resultados.xlsx')) as writer:
    DF_ANALIZAR.to_excel(writer, sheet_name="DocumentosMAF_ANALIZADOS",index=False)
    
  print(f"Hora de fin: {datetime.datetime.now()}")
  pantalla_final()