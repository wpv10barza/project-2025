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
PRT_path_base       = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.01 EJECUCION TI\05. RPA EXEC\03. OP-OPE"

# Variables constantes
NOMBRE_CONEXION_SAP = "Copper SAP ECC Producción(PS8)"
RSYST_MANDT         = "300"
RSYST_BNAME         = os.getenv("SAP_USER")
RSYST_BCODE         = os.getenv("SAP_PASS")
RSYST_LANGU         = "ES"
TRANSACCION_IA06    = "/NIA06"
FILTRO_IA06_PERFIL  = "0000008"
MENSAJE_ERROR       = ""

# Variables de configuración
PRT_CarpetaAProcesar     = "REQ-009 Solicitud 04_20250820 (CREACION)"
ArchivoExcelNombre       = "DATOS.xlsx"
PRT_ArchivoHojaNombre    = "OP-OPE"
PATH_EXCEL_DATOS         = os.path.join(PRT_path_base, PRT_CarpetaAProcesar, ArchivoExcelNombre)
DESCARGA_PATH_REQUERI    = os.path.join(PRT_path_base, PRT_CarpetaAProcesar)

COLUMNA_GRUPO            = "GRUPO"
COLUMNA_CONTADOR         = "CONTADOR"
COLUMNAS_DE_DATOS        = [
                          "Op",
                          "Descripción OP",
                          "PtoTrabajo",
                        
                          #"TRABAJO",
                          "CANTIDAD",
                          "DURACIÓN"
                      ]



def transaccion_IA06(GRUPO_HDR, CONTADOR_HDR, df_OPERACIONES):
  # Filtro de transacción 
  session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_IA06
  session.findById("wnd[0]").sendVKey(0) 
  session.findById("wnd[0]/usr/ctxt[0]").text = GRUPO_HDR
  session.findById("wnd[0]/usr/ctxt[1]").text = FILTRO_IA06_PERFIL 
  session.findById("wnd[0]/usr/ctxt[2]").text = ""
  session.findById("wnd[0]/usr/txt[1]").text = CONTADOR_HDR
  session.findById("wnd[0]").sendVKey(0)
  
  seleccion_ope(df_OPERACIONES)
  
def seleccion_ope(df_OPERACIONES):
  global MENSAJE_ERROR
  scroll = 3
  for index in range(len(df_OPERACIONES)):
    OPERACION      = df_OPERACIONES["Op"].iloc[index]
    DESCRIPCION_OP = df_OPERACIONES["Descripción OP"].iloc[index] 
    PST_TRABAJO    = df_OPERACIONES["PtoTrabajo"].iloc[index]
    CANTIDAD       = df_OPERACIONES["CANTIDAD"].iloc[index]
    DURACION       = df_OPERACIONES["DURACIÓN"].iloc[index]  
                                
    # Seleccionar operacion
    try:
        OPERACION = (OPERACION // 10)-1
    except TypeError as e:
        MENSAJE_ERROR = "[ERROR]: No se puede realizar la operación matemática de un string."   
        print(f"{MENSAJE_ERROR}:  {e}")
        return
    try:
        session.findById(f"wnd[0]/usr/tblSAPLCPDITCTRL_3400/ctxt[2,{OPERACION}]").text = PST_TRABAJO
        session.findById(f"wnd[0]/usr/tblSAPLCPDITCTRL_3400/txt[5,{OPERACION}]").text = DESCRIPCION_OP
        session.findById(f"wnd[0]/usr/tblSAPLCPDITCTRL_3400/txt[9,{OPERACION}]").text = CANTIDAD 
        session.findById(f"wnd[0]/usr/tblSAPLCPDITCTRL_3400/txt[10,{OPERACION}]").text = DURACION 
    except:
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3400").verticalScrollbar.position = scroll
        session.findById(f"wnd[0]/usr/tblSAPLCPDITCTRL_3400/ctxt[2,{OPERACION - scroll}]").text = PST_TRABAJO
        session.findById(f"wnd[0]/usr/tblSAPLCPDITCTRL_3400/txt[5,{OPERACION - scroll}]").text = DESCRIPCION_OP
        session.findById(f"wnd[0]/usr/tblSAPLCPDITCTRL_3400/txt[9,{OPERACION - scroll}]").text = CANTIDAD 
        session.findById(f"wnd[0]/usr/tblSAPLCPDITCTRL_3400/txt[10,{OPERACION - scroll}]").text = DURACION 
        scroll += 1
    # Guardar
  session.findById("wnd[0]").sendVKey(0)
  session.findById("wnd[0]/tbar[0]/btn[11]").press() ######################
  try:
    session.findById("wnd[1]/tbar[0]/btn[0]").press()   
    session.findById("wnd[1]").close()
  except:
     print("No existe advertencia")
 
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
    df_req  =  pd.read_excel(PATH_EXCEL_DATOS, sheet_name = PRT_ArchivoHojaNombre)
    df_req  =  df_req.replace(r'^\s*$',np.nan,regex=True).infer_objects(copy=False)
    COLUMNAS_DE_CAMBIO = df_req.columns.to_list() 
  except Exception as e:
    print("No encuentra el nombre de la hoja de Excel")
      
  DF_ANALIZAR = eliminar_espacio(df_req)
  
# ------------------- Inicialización de SAP -----------------

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
      session.findById("wnd[0]").maximize
  except Exception as e:
      print(f"[ERROR] Error al abrir conexión SAP: {e}")
      exit()
      
  # Registro de cuenta SAP
  session.findById("wnd[0]/usr/txtRSYST-MANDT").text = RSYST_MANDT
  session.findById("wnd[0]/usr/txtRSYST-BNAME").text = RSYST_BNAME
  session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = RSYST_BCODE
  session.findById("wnd[0]/usr/txtRSYST-LANGU").text = RSYST_LANGU
  session.findById("wnd[0]").sendVKey(0)
  
  
  #Crear operaciones considerando que existe la hoja de ruta correspondiente
  DF_ANALIZAR = DF_ANALIZAR.fillna(method = 'ffill')
  HDRs = DF_ANALIZAR[["GRUPO","CONTADOR"]].drop_duplicates().reset_index(drop=True)

  for index in range(len(HDRs)):
    
    if DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] == "ERROR_NO_MATCH":
      continue
    
    MENSAJE_ERROR = ""
    HDR    = HDRs.loc[[index]]
    HDR_OP = pd.merge(HDR, DF_ANALIZAR, on=["GRUPO", "CONTADOR"], how='inner')



    GRUPO         = HDR.loc[index, COLUMNA_GRUPO]
    CONTADOR      = HDR.loc[index, COLUMNA_CONTADOR]

    DATOS_OPERACION = HDR_OP[COLUMNAS_DE_DATOS]
    
    transaccion_IA06(GRUPO, CONTADOR, DATOS_OPERACION)  

    
    HDR.loc[index, "ESTADO_EJECUCION_RPA"] = "OK" if MENSAJE_ERROR == "" else MENSAJE_ERROR 
  HDR = HDR[[
                            "GRUPO",
                            "CONTADOR",
                            "ESTADO_EJECUCION_RPA",

                            ]]
  with pd.ExcelWriter(os.path.join(DESCARGA_PATH_REQUERI,'Resultados.xlsx')) as writer:
    DF_ANALIZAR.to_excel(writer, sheet_name="DocumentosMAF_ANALIZADOS",index=False)

    
  print(f"Hora de fin: {datetime.datetime.now()}")
  pantalla_final()