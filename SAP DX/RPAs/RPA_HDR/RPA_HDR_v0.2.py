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
TRANSACCION_IA06    = "/NIA06"
PMA_FIELD           = "PMA" # Clase de documento
DOC_PARCIAL         = "000" # Documento parcial
CATEGORIA_ARCHIVO   = "CF_IMP_AS" # Categoría de archivo DMS_C1_ST para versión SAP/4HANA (Versión 770)
FILTRO_IA06_PERFIL  = "0000008"
MENSAJE_ERROR       = ""
# Variables de configuración
PRT_CarpetaAProcesar     = "REQ-008 Solicitud 04_20250820 (ACTUALIZACIÓN)"
PRT_ArchivoExcelNombre   = "DATOS.xlsx"
ArchivoHojaNombre        = "HDR"
IA10_RutaNombre          = "IA10.xlsx"
VENTANA_BORRAR_MAF       = "Borrar asignaciones MAF"
PATH_EXCEL_DATOS         = os.path.join(PRT_path_base, PRT_CarpetaAProcesar, PRT_ArchivoExcelNombre)
PATH_EXCEL_IA10          = os.path.join(PRT_path_base, PRT_CarpetaAProcesar, IA10_RutaNombre)
DESCARGA_PATH_REQUERI    = os.path.join(PRT_path_base, PRT_CarpetaAProcesar)

COLUMNA_GRUPO            = "GRUPO"
COLUMNA_CONTADOR         = "CONTADOR"
COLUMNA_OPERACION        = "OPERACION"
COLUMNAS_DE_CAMBIO       = [
                          "DESCRIPCIÓN HDR",
                          "IDENTIFICADOR",
                          "PUESTO DE TRABAJO",
                          "CENTRO",
                          "UTILIZACIÓN",
                          "GRUPO PLANF",
                          "ESTADO GLOBAL",
                          "ESTADO DE INSTALACIÓN",
                          "CONJUNTO",
                          "Elemento refer.PM/PS",
                          "PETICIÓN DE BORRADO",
                          "PUNTOS DE INSPECCIÓN",
                          "NUMERACIÓN EXTERNA"
                      ]

COLUMNAS_ID={
            "DESCRIPCIÓN HDR": "wnd[0]/usr/txt[3]",
            "IDENTIFICADOR": "wnd[0]/usr/txt[4]",
            "PUESTO DE TRABAJO": "wnd[0]/usr/ctxt[2]",
            "CENTRO": "wnd[0]/usr/ctxt[3]",
            "UTILIZACIÓN": "wnd[0]/usr/ctxt[4]",
            "GRUPO PLANF": "wnd[0]/usr/ctxt[5]",
            "ESTADO GLOBAL": "wnd[0]/usr/ctxt[6]",
            "ESTADO DE INSTALACIÓN": "wnd[0]/usr/ctxt[7]",
            "CONJUNTO": "wnd[0]/usr/ctxt[9]",
            "Elemento refer.PM/PS": "wnd[0]/usr/ctxt[10]",
            "PETICIÓN DE BORRADO": "wnd[0]/usr/chk[1]",
            "PUNTOS DE INSPECCIÓN": "wnd[0]/usr/ctxt[11]",
            "NUMERACIÓN EXTERNA": "wnd[0]/usr/cmb"     
            }

NUM_EXT_CHECK_LIST = {
                "Cualq.numeración externa posible":     "0",
                "Ninguna numeración externa":           "1",
                "Numeración externa requerida":         "2",
                "Numeración externa unívoca posible":   "4",
                "Requerida numeración externa unívoca": "3"
                      }


def transaccion_IA06(GRUPO_HDR, CONTADOR_HDR, OPERACION_HDR, df_CAMBIAR_DATOS):
  # Filtro de transacción 
  session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_IA06
  session.findById("wnd[0]").sendVKey(0) 
  session.findById("wnd[0]/usr/ctxt[0]").text = GRUPO_HDR
  session.findById("wnd[0]/usr/ctxt[1]").text = FILTRO_IA06_PERFIL 
  session.findById("wnd[0]/usr/ctxt[2]").text = ""
  session.findById("wnd[0]/usr/txt[1]").text = CONTADOR_HDR
  session.findById("wnd[0]").sendVKey(0)
  
  seleccion_ope(OPERACION_HDR,df_CAMBIAR_DATOS)
  
def seleccion_ope(operacion,df_CAMBIAR_DATOS):
  global MENSAJE_ERROR

  # Seleccionar operacion
  try:
    operacion = (operacion // 10) - 1
  except TypeError as e:
    MENSAJE_ERROR = "[ERROR]: No se puede realizar la operación matemática de un string."   
    print(f"{MENSAJE_ERROR}:  {e}")
    return

  try:
    session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3400").getAbsoluteRow(operacion).selected = True
    session.findById("wnd[0]/tbar[1]/btn[32]").press() # Selección de cabecera
  except Exception as e:
    MENSAJE_ERROR ="[ERROR]: No existe la fila de operación"
    print(f"{MENSAJE_ERROR} : {e}")
    return
  
  
     
  for columna_cambio in COLUMNAS_DE_CAMBIO:
    if columna_cambio in COLUMNAS_ID:
      try:
        val =  df_CAMBIAR_DATOS[columna_cambio].iloc[0]
        
        if COLUMNAS_ID[columna_cambio] =="wnd[0]/usr/chk[1]" :
          if pd.notna(val):
            session.findById(COLUMNAS_ID[columna_cambio]).selected = True 
          else:
            session.findById(COLUMNAS_ID[columna_cambio]).selected = False # Clave modelo
        elif COLUMNAS_ID[columna_cambio] =="wnd[0]/usr/cmb" :
          print(f"valor: {val}")
          if val in NUM_EXT_CHECK_LIST: 
            session.findById(COLUMNAS_ID[columna_cambio]).key = NUM_EXT_CHECK_LIST[val]
        else:
            if pd.notna(val):
                session.findById(COLUMNAS_ID[columna_cambio]).text = val if val != "DEJAR EN BLANCO" else "" 
                  
      except:
        MENSAJE_ERROR =f"[ERROR]: Valor no aceptado para modificar. Revisar {val}"
        print(f"{MENSAJE_ERROR} : {e}")
  # Guardar
  session.findById("wnd[0]/tbar[0]/btn[11]").press() 
  
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

  return df_req 
 
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
      
      
  if df_req[COLUMNA_GRUPO].isnull().any():  
    if not os.path.exists(PATH_EXCEL_IA10):
      raise FileNotFoundError(f"Documento de Operaciones no encontrada. Revisar ruta: {PATH_EXCEL_IA10}")
    df_ia10  =  pd.read_excel(PATH_EXCEL_IA10, sheet_name="DATA")
    DF_ANALIZAR = busqueda_grupo(df_req,df_ia10)
    DF_ANALIZAR = eliminar_espacio(DF_ANALIZAR)
    
  else:
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
  
  
  # CAMBIAR_DATOS = DF_ANALIZAR.iloc[[1]][COLUMNAS_DE_CAMBIO]
  # GRUPO         = DF_ANALIZAR.loc[1, COLUMNA_GRUPO]
  # CONTADOR      = DF_ANALIZAR.loc[1, COLUMNA_CONTADOR]
  # OPERACION     = DF_ANALIZAR.loc[1, COLUMNA_OPERACION]

  # transaccion_IA06(GRUPO, CONTADOR, OPERACION, CAMBIAR_DATOS)  
  
  for index in range(len(DF_ANALIZAR)):
    
    if DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] == "ERROR_NO_MATCH":
      continue
    
    MENSAJE_ERROR = ""
    CAMBIAR_DATOS = DF_ANALIZAR.iloc[[index]][COLUMNAS_DE_CAMBIO]
    GRUPO         = DF_ANALIZAR.loc[index, COLUMNA_GRUPO]
    CONTADOR      = DF_ANALIZAR.loc[index, COLUMNA_CONTADOR]
    OPERACION     = DF_ANALIZAR.loc[index, COLUMNA_OPERACION]
    
    transaccion_IA06(GRUPO, CONTADOR, OPERACION, CAMBIAR_DATOS)  

    
    DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] = "OK" if MENSAJE_ERROR == "" else MENSAJE_ERROR 
  DF_ANALIZAR = DF_ANALIZAR[[
                            "DESCRIPCIÓN DE LA OPERACIÓN",
                            # "DOCUMENTO MAF",
                            # "NUEVA_VERSION",
                            "GRUPO",
                            "CONTADOR",
                            "OPERACION",
                            "ESTADO_EJECUCION_RPA"
                            ]]
  with pd.ExcelWriter(os.path.join(DESCARGA_PATH_REQUERI,'Resultados.xlsx')) as writer:
    DF_ANALIZAR.to_excel(writer, sheet_name="DocumentosMAF_ANALIZADOS",index=False)
    
  print(f"Hora de fin: {datetime.datetime.now()}")
  pantalla_final()