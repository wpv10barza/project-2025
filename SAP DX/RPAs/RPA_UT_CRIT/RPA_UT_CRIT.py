from dotenv import load_dotenv
import os, sys
#Carga variables de .env
path_RPAs = os.path.abspath(os.path.join(os.path.dirname(__file__),".."))
doten_path =os.path.join(path_RPAs,".env")
load_dotenv(doten_path)
sys.path.append(path_RPAs)

from ventana import pantalla_final
import time, subprocess
import pandas as pd
import win32com.client
import numpy as np
import datetime
import pygetwindow as gw
import pyautogui


sap_logon_path      = r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"
PRT_path_base       = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\05. RPA EXEC"

# Variables constantes
NOMBRE_CONEXION_SAP = "Copper SAP ECC Producción(PS8)"
RSYST_MANDT         = "300"
RSYST_BNAME         = os.getenv("SAP_USER")
RSYST_BCODE         = os.getenv("SAP_PASS")
RSYST_LANGU         = "ES"
TRANSACCION_IL02    = "/NIL02"
MENSAJE_ERROR       = ""

# Variables de configuración
CarpetaAProcesar         = "REQ-001"
ArchivoExcelNombre       = "DATOS.xlsx"
Carpeta_RPA              = "10. UT-CRIT"
PRT_ArchivoHojaNombre    = "UT_CRIT"
PATH_RPA_BASE            = os.path.join(PRT_path_base, Carpeta_RPA)
PATH_EXCEL_DATOS         = os.path.join(PATH_RPA_BASE, CarpetaAProcesar, ArchivoExcelNombre)
DESCARGA_PATH_REQUERI    = os.path.join(PATH_RPA_BASE, CarpetaAProcesar)

COLUMNA_UBICACIONES      = "UBICACIONES TÉCNICAS"
COLUMNA_ESTADO           = "ESTADO_EJECUCION_RPA"
COLUMNA_CRITICIDAD       = "CRITICIDAD"


def transaccion_IL02(ubicaciones, criticidad):
  global MENSAJE_ERROR

  # Ubicaciones técnicas IL02
  session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_IL02
  session.findById("wnd[0]").sendVKey(0) 
  
  # Seleccionar Ubicaciones técnicas (equipos)
  session.findById("wnd[0]/usr/ctxt[0]").text = ubicaciones
  session.findById("wnd[0]").sendVKey(0) 

  session.findById(r"wnd[0]/usr/tabsTABSTRIP/tabpT\02").select()
  session.findById(r"wnd[0]/usr/tabsTABSTRIP/tabpT\02/ssub/1/2/sub/1/2/1/ctxt[4]").text = criticidad

  #Guardar
  session.findById("wnd[0]/tbar[0]/btn[11]").press()

      
def eliminar_espacio(df):
  for columna in df.columns:
    df[columna] = df[columna].apply(lambda x: x.strip() if isinstance(x,str) else x)
    # df[columna] = df[columna].apply(lambda x: x.replace("  "," ") if isinstance(x,str) else x)
  return df    
       
      

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
    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = RSYST_MANDT
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = RSYST_BNAME
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = RSYST_BCODE
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = RSYST_LANGU
    session.findById("wnd[0]").sendVKey(0)
    return session
  except Exception as e:
    print(f"[ERROR] Error al abrir conexión SAP: {e}")
    exit()



if __name__=="__main__":
  print("Hora de inicio: ",datetime.datetime.now())

  # Lectura de Excel con requerimientos
  if not os.path.exists(PATH_EXCEL_DATOS):
    raise FileNotFoundError(f"Ruta de Excel no encontrada: {PATH_EXCEL_DATOS}")
   
  try:
    df_req  = pd.read_excel(PATH_EXCEL_DATOS, sheet_name = PRT_ArchivoHojaNombre)
    df_req  = df_req.replace(r'^\s*$',np.nan,regex=True).infer_objects(copy=False)
    # COLUMNAS_DE_CAMBIO = df_req.columns.to_list()  
    DF_ANALIZAR = eliminar_espacio(df_req)
    DF_ANALIZAR[COLUMNA_ESTADO]=""
  except Exception as e:
    print("No encuentra el nombre de la hoja de Excel", e)
  
# ------------------- Inicialización de SAP -----------------
  session = SAP_start()


  for index in range(len(DF_ANALIZAR)):
    
    if (DF_ANALIZAR.loc[index, COLUMNA_ESTADO ] == "OK"):
      continue
    
    MENSAJE_ERROR = ""
    UBICACIONES   = DF_ANALIZAR.loc[index, COLUMNA_UBICACIONES]
    CRITICIDAD    = DF_ANALIZAR.loc[index, COLUMNA_CRITICIDAD]

    try:
      transaccion_IL02(UBICACIONES, CRITICIDAD)
      print(UBICACIONES)  
    except Exception as e:
      DF_ANALIZAR.loc[index, COLUMNA_ESTADO ] = "ERROR_EJECUCIÓN_RPA"
      print(e)
      
    DF_ANALIZAR.loc[index, COLUMNA_ESTADO ] = "OK" if MENSAJE_ERROR == "" else MENSAJE_ERROR 
  
  DF_ANALIZAR = DF_ANALIZAR[[
                            COLUMNA_UBICACIONES,
                            COLUMNA_CRITICIDAD,
                            COLUMNA_ESTADO 
                            ]]
  with pd.ExcelWriter(os.path.join(DESCARGA_PATH_REQUERI,'Resultados.xlsx')) as writer:
    DF_ANALIZAR.to_excel(writer, sheet_name="DocumentosMAF_ANALIZADOS",index=False)
    
  print(f"Hora de fin: {datetime.datetime.now()}")
  pantalla_final()