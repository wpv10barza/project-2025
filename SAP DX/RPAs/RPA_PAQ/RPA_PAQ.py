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


sap_logon_path      = r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"
PRT_path_base       = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.01 EJECUCION TI\05. RPA EXEC\08. PAQ_FREQ"

# Variables constantes
NOMBRE_CONEXION_SAP = "Copper SAP ECC Producción(PS8)"
RSYST_MANDT         = "300"
RSYST_BNAME         = os.getenv("SAP_USER")
RSYST_BCODE         = os.getenv("SAP_PASS")
RSYST_LANGU         = "ES"
TRANSACCION_IA06    = "/NIA06"
PMA_FIELD           = "PMA" # Clase de documento
DOC_PARCIAL         = "000" # Documento parcial
CATEGORIA_ARCHIVO   = "CF_IMP_AS" # Categoría de archivo DMS_C1_ST para versión SAP/4HANA (Versión 770)
FILTRO_IA06_PERFIL  = "0000008"
MENSAJE_ERROR       = ""
# Variables de configuración
PRT_CarpetaAProcesar     = "REQ-009 Solicitud 04_20250820 (ACTUALIZACIÓN)"
ArchivoExcelNombre       = "DATOS.xlsx"
PRT_ArchivoHojaNombre    = "PAQ_FREQ"
IA10_RutaNombre          = "IA10.xlsx"
VENTANA_BORRAR_MAF       = "Borrar asignaciones MAF"
PATH_EXCEL_DATOS         = os.path.join(PRT_path_base, PRT_CarpetaAProcesar, ArchivoExcelNombre)
PATH_EXCEL_IA10          = os.path.join(PRT_path_base, PRT_CarpetaAProcesar, IA10_RutaNombre)
DESCARGA_PATH_REQUERI    = os.path.join(PRT_path_base, PRT_CarpetaAProcesar)

COLUMNA_GRUPO            = "GRUPO"
COLUMNA_CONTADOR         = "CONTADOR"
COLUMNA_OPERACION        = "OPERACIÓN"
COLUMNA_FREQ             = "PAQ_FREQ"
COLUMNA_ESTADO           = "ESTADO_EJECUCION_RPA"


def transaccion_IA06(GRUPO_HDR, CONTADOR_HDR, OPERACION_HDR, paquete_freq):
  # Filtro de transacción 
  session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_IA06
  session.findById("wnd[0]").sendVKey(0) 
  session.findById("wnd[0]/usr/ctxt[0]").text = GRUPO_HDR
  session.findById("wnd[0]/usr/ctxt[1]").text = FILTRO_IA06_PERFIL 
  session.findById("wnd[0]/usr/ctxt[2]").text = ""
  session.findById("wnd[0]/usr/txt[1]").text = CONTADOR_HDR
  session.findById("wnd[0]").sendVKey(0)
  seleccion_ope(OPERACION_HDR,paquete_freq)
  
  
def seleccion_ope(operacion, paquete_freq):
  global MENSAJE_ERROR
  cont = 0                          
  # Seleccionar operacion
  try:
    operacion = (operacion // 10) - 1
  except TypeError as e:
    MENSAJE_ERROR = "[ERROR]: No se puede realizar la operación matemática de un string."   
    print(f"{MENSAJE_ERROR}:  {e}")
    return

  try:
    session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3400").getAbsoluteRow(operacion).selected = True
    
    session.findById("wnd[0]/usr/btn[0]").press() # Selección de paquetes de mantenimiento
    session.findById("wnd[0]/tbar[1]/btn[26]").press() # Selección de paquete mantenimiento preventivo 
  except Exception as e:
    MENSAJE_ERROR ="[ERROR]: No existe la fila de operación"
    print(f"{MENSAJE_ERROR} : {e}")
    return
  
  for i in range(0,9): 
    if session.findById(f"wnd[0]/usr/tblSAPLCIDITCTRL_3000/txt[0,{i}]").text != "": # and cont > 1:
      cont += 1
    elif cont>1:
      MENSAJE_ERROR = "Existe más de un paquete asociado."
      return  
  try:      
    
    session.findById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txt[0,0]").text = paquete_freq
  except:
    session.findById("wnd[0]/usr/tblSAPLCIDITCTRL_3000").getAbsoluteRow(0).selected = True
    session.findById("wnd[0]/tbar[1]/btn[14]").press()
    session.findById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txt[0,0]").text = paquete_freq # Texto para llenar
  session.findById("wnd[0]").sendVKey(0)
  #GUARDAR
  session.findById("wnd[0]/tbar[0]/btn[11]").press()
  try:
    session.findById("wnd[1]/tbar[0]/btn[0]").press()   
  except:
     print("No existe advertencia")


def busqueda_grupo(df_req,df_ia10):
  """
  Solo busca los grupos, operaciones y contadores de las Descripciones 
  que se encuentran una única vez en IA10
  """
  df_req["COLUMNA_MATCH"] = ""
  
  conteo_match = df_ia10["DESCRIPCIÓN"].value_counts()
  
  df_req['cantidad_match'] = df_req["DESCRIPCIÓN DE LA OPERACIÓN"].map(conteo_match).fillna(0).astype(int)
  df_req.loc[df_req['cantidad_match']> 1 , 'COLUMNA_MATCH'] = "TIENE MÁS DE UN MATCH"
  df_req.loc[df_req['cantidad_match'] == 0 , 'COLUMNA_MATCH'] = "NO TIENE NINGÚN MATCH"
  df_req.loc[df_req['cantidad_match'] == 1 , 'COLUMNA_MATCH'] = "TIENE MATCH"

  
  df_sin_grupo = df_req[(df_req[COLUMNA_GRUPO].isnull()) & (df_req['COLUMNA_MATCH']=="TIENE MATCH")]
  df_sin_grupo = pd.merge(
                          df_sin_grupo.drop(columns=[COLUMNA_CONTADOR,COLUMNA_GRUPO, COLUMNA_OPERACION]),
                          df_ia10[[COLUMNA_GRUPO, COLUMNA_CONTADOR,COLUMNA_OPERACION,"DESCRIPCIÓN"]],
                          left_on = "DESCRIPCIÓN DE LA OPERACIÓN",right_on = "DESCRIPCIÓN",
                          how="left")

  
  df_req[COLUMNA_GRUPO] = df_req[COLUMNA_GRUPO].fillna(
                                                df_req["DESCRIPCIÓN DE LA OPERACIÓN"].map(df_sin_grupo.set_index("DESCRIPCIÓN")[COLUMNA_GRUPO])) 
  df_req[COLUMNA_CONTADOR] = df_req[COLUMNA_CONTADOR].fillna(
                                                df_req["DESCRIPCIÓN DE LA OPERACIÓN"].map(df_sin_grupo.set_index("DESCRIPCIÓN")[COLUMNA_CONTADOR])) 
  df_req[COLUMNA_OPERACION] = df_req[COLUMNA_OPERACION].fillna(
                                                df_req["DESCRIPCIÓN DE LA OPERACIÓN"].map(df_sin_grupo.set_index("DESCRIPCIÓN")[COLUMNA_OPERACION])) 

  
  df_req.loc[pd.isna(df_req[COLUMNA_CONTADOR]),COLUMNA_ESTADO ] = "ERROR_NO_MATCH"

  return df_req 
 
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
      
  if df_req[COLUMNA_GRUPO].isnull().any():  
    if not os.path.exists(PATH_EXCEL_IA10):
      raise FileNotFoundError(f"Documento de Operaciones no encontrada. Revisar ruta: {PATH_EXCEL_IA10}")
    df_ia10 =  pd.read_excel(PATH_EXCEL_IA10, sheet_name="DATA")
    DF_ANALIZAR = busqueda_grupo(df_req,df_ia10)
    DF_ANALIZAR = eliminar_espacio(DF_ANALIZAR)
  else:
    DF_ANALIZAR = eliminar_espacio(df_req)
    DF_ANALIZAR[COLUMNA_ESTADO]=""
  
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
  

  for index in range(len(DF_ANALIZAR)):
    
    if DF_ANALIZAR.loc[index, COLUMNA_ESTADO ] == "ERROR_NO_MATCH":
      continue
    
    MENSAJE_ERROR = ""
    GRUPO         = DF_ANALIZAR.loc[index, COLUMNA_GRUPO]
    CONTADOR      = DF_ANALIZAR.loc[index, COLUMNA_CONTADOR]
    OPERACION     = DF_ANALIZAR.loc[index, COLUMNA_OPERACION]
    FREQ          = DF_ANALIZAR.loc[index, COLUMNA_FREQ]
  
    try:
      transaccion_IA06(GRUPO, CONTADOR, OPERACION, FREQ)  
    except:
      DF_ANALIZAR.loc[index, COLUMNA_ESTADO ] = "ERROR_EJECUCIÓN_RPA"
       
    DF_ANALIZAR.loc[index, COLUMNA_ESTADO ] = "OK" if MENSAJE_ERROR == "" else MENSAJE_ERROR

  DF_ANALIZAR = DF_ANALIZAR[[
                            # "DESCRIPCIÓN DE LA OPERACIÓN",
                            # "DOCUMENTO MAF",
                            COLUMNA_GRUPO,
                            COLUMNA_CONTADOR,
                            COLUMNA_OPERACION,
                            COLUMNA_ESTADO 
                            # "cantidad_match",
                            # "COLUMNA_MATCH"
                            ]]
  with pd.ExcelWriter(os.path.join(DESCARGA_PATH_REQUERI,'Resultados.xlsx')) as writer:
    DF_ANALIZAR.to_excel(writer, sheet_name="DocumentosMAF_ANALIZADOS",index=False)

    
  print(f"Hora de fin: {datetime.datetime.now()}")
  pantalla_final()