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


sap_logon_path      = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
PRT_path_base       = r"D:\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\05. RPA EXEC\03. OP-OPE"

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
PRT_CarpetaAProcesar     = "REQ-001"
PRT_ArchivoExcelNombre   = "DATOS.xlsx"
PRT_ArchivoHojaNombre    = "OP-OPE"
IA39_RutaNombre          = "IA39.xlsx"
VENTANA_BORRAR_MAF       = "Borrar asignaciones MAF"
PATH_EXCEL_DATOS         = os.path.join(PRT_path_base, PRT_CarpetaAProcesar, PRT_ArchivoExcelNombre)
PATH_EXCEL_IA39          = os.path.join(PRT_path_base, PRT_CarpetaAProcesar, IA39_RutaNombre)
DESCARGA_PATH_REQUERI    = os.path.join(PRT_path_base, PRT_CarpetaAProcesar)

COLUMNA_GRUPO            = "GRUPO"
COLUMNA_CONTADOR         = "CONTADOR"
COLUMNA_OPERACION        = "OPERACION"
COLUMNAS_DE_CAMBIO       = [
                          "CLAVE DE MODELO",
                          "TEXTO BREVE DE OPERACIÓN",
                          "PUESTO DE TRABAJO",
                          "CENTRO",
                          "CONJUNTO",
                          "ESTADO DE INSTALACIÓN",
                          "UBICACIÓN TÉCNICA",
                          "EQUIPO",
                          "FACTOR DE EJECUCIÓN",
                          "CLASE DE ACTIVIDAD",
                          "DISTRIBUCIÓN DE TRABAJO INTERNO",
                          
                          "TRABAJO",
                          "CANTIDAD",
                          "PORCENTAJE",
                          "DURACIÓN NORMAL",
                          
                      ]

COLUMNAS_ID=[
               "wnd[0]/usr/sub/1/ctxt[0]", # CLAVE
               "wnd[0]/usr/sub/1/txt[4]",  # TEXTO BREVE DE OPERACIÓN
               "wnd[0]/usr/sub/1/ctxt[1]", # PUESTO DE TRABAJO
               "wnd[0]/usr/sub/1/ctxt[2]", # CENTRO
               "wnd[0]/usr/sub/1/ctxt[4]", # CONJUNTO
               "wnd[0]/usr/sub/1/ctxt[5]", # ESTADO INSTALACIÓN
               "wnd[0]/usr/sub/1/ctxt[6]", # UBICACIÓN TÉCNICA
               "wnd[0]/usr/sub/1/ctxt[7]", # EQUIPO
               "wnd[0]/usr/sub/1/txt[11]", # FACTOR DE EJECUCIÓN
               "wnd[0]/usr/ctxt[1]",       # CLASE DE ACTIVIDAD
               "wnd[0]/usr/ctxt[4]"        # DISTRIBUCIÓN DE TRABAJO INTERNO
               
               "wnd[0]/usr/txt[0]",        # TRABAJO
               "wnd[0]/usr/txt[1]",        # CANTIDAD
               "wnd[0]/usr/txt[2]",        # PORCENTAJE
               "wnd[0]/usr/txt[3]",        # DURACIÓN NORMAL
               
               ]


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
  
  TRABAJO    = df_CAMBIAR_DATOS[COLUMNAS_DE_CAMBIO[11]].iloc[0]
  CANTIDAD   = df_CAMBIAR_DATOS[COLUMNAS_DE_CAMBIO[12]].iloc[0]
  PORCENTAJE = df_CAMBIAR_DATOS[COLUMNAS_DE_CAMBIO[13]].iloc[0]
  DURACION   = df_CAMBIAR_DATOS[COLUMNAS_DE_CAMBIO[14]].iloc[0]  
  
  # Seleccionar operacion
  try:
    operacion = (operacion // 10) - 1
  except TypeError as e:
    MENSAJE_ERROR = "[ERROR]: No se puede realizar la operación matemática de un string."   
    print(f"{MENSAJE_ERROR}:  {e}")
    return

  try:
    session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3400").getAbsoluteRow(operacion).selected = True
    session.findById("wnd[0]/tbar[1]/btn[17]").press() # Selección de operación interna
  except Exception as e:
    MENSAJE_ERROR ="[ERROR]: No existe la fila de operación"
    print(f"{MENSAJE_ERROR} : {e}")
    return
    
  for columna_cambio, columna_id in zip(COLUMNAS_DE_CAMBIO[:-4],COLUMNAS_ID[:-4]):
    try:
      val =  df_CAMBIAR_DATOS[columna_cambio].iloc[0]
      if pd.notna(val):
        session.findById(columna_id).text = val if val != "DEJAR EN BLANCO" else "" # Clave modelo
    except:
      MENSAJE_ERROR =f"[ERROR]: Valor no aceptado para modificar. Revisar {val}"
      print(f"{MENSAJE_ERROR} : {e}")
  
 
  if session.findById("wnd[0]/usr/ctxt[3]").text == "0": # Actualizar manual  
    session.findById("wnd[0]/usr/txt[0]").text = TRABAJO if pd.notna(TRABAJO) else session.findById("wnd[0]/usr/txt[0]").text# Trabajo ***CORROBORAR
    session.findById("wnd[0]/usr/txt[1]").text = CANTIDAD if pd.notna(CANTIDAD) else session.findById("wnd[0]/usr/txt[1]").text # Cantidad
    session.findById("wnd[0]/usr/txt[2]").text = PORCENTAJE if pd.notna(PORCENTAJE) else session.findById("wnd[0]/usr/txt[2]").text # Porcentaje
    session.findById("wnd[0]/usr/txt[3]").text = DURACION if pd.notna(DURACION) else session.findById("wnd[0]/usr/txt[3]").text# Duración normal       
  
  elif session.findById("wnd[0]/usr/ctxt[3]").text == "1": # Calcular duración
    session.findById("wnd[0]/usr/txt[0]").text = TRABAJO if pd.notna(TRABAJO) else session.findById("wnd[0]/usr/txt[0]").text# Trabajo ***CORROBORAR
    session.findById("wnd[0]/usr/txt[1]").text = CANTIDAD if pd.notna(CANTIDAD) else session.findById("wnd[0]/usr/txt[1]").text# Cantidad
    session.findById("wnd[0]/usr/txt[2]").text = PORCENTAJE if pd.notna(PORCENTAJE) else session.findById("wnd[0]/usr/txt[2]").text# Porcentaje
  
  elif session.findById("wnd[0]/usr/ctxt[3]").text == "2": # Calcular trabajo
    session.findById("wnd[0]/usr/txt[1]").text = CANTIDAD if pd.notna(CANTIDAD) else session.findById("wnd[0]/usr/txt[1]").text# Cantidad
    session.findById("wnd[0]/usr/txt[2]").text = PORCENTAJE if pd.notna(PORCENTAJE) else session.findById("wnd[0]/usr/txt[2]").text# Porcentaje
    session.findById("wnd[0]/usr/txt[3]").text = DURACION if pd.notna(DURACION) else session.findById("wnd[0]/usr/txt[3]").text# Duración normal       
  
  # elif session.findById("wnd[0]/usr/ctxt[3]").text == "3": # Calcular capacidad de visualización
  #   # session.findById("wnd[0]/usr/txt[0]").text = "2" # Trabajo ***CORROBORAR
  #   session.findById("wnd[0]/usr/txt[1]").text = "2" # Cantidad
  #   session.findById("wnd[0]/usr/txt[2]").text = "50" # Porcentaje
  #   session.findById("wnd[0]/usr/txt[3]").text = "2" # Duración normal       
  
  # Guardar
  # session.findById("wnd[0]").sendVKey(11) ######################  
  
def busqueda_grupo(df_req,df_ia39):
  """
  Solo busca los grupos, operaciones y contadores de las Descripciones 
  que se encuentran una única vez en IA39
  """
  df_req["COLUMNA_MATCH"] = ""
  
  conteo_match = df_ia39["Txt.brv.operación"].value_counts()
  
  df_req['cantidad_match'] = df_req["DESCRIPCIÓN DE LA OPERACIÓN"].map(conteo_match).fillna(0).astype(int)
  df_req.loc[df_req['cantidad_match']> 1 , 'COLUMNA_MATCH'] = "TIENE MÁS DE UN MATCH"
  df_req.loc[df_req['cantidad_match'] == 0 , 'COLUMNA_MATCH'] = "NO TIENE NINGÚN MATCH"
  df_req.loc[df_req['cantidad_match'] == 1 , 'COLUMNA_MATCH'] = "TIENE MATCH"

  
  df_sin_grupo = df_req[(df_req[COLUMNA_GRUPO].isnull()) & (df_req['COLUMNA_MATCH']=="TIENE MATCH")]
  df_sin_grupo = pd.merge(
                          df_sin_grupo.drop(columns=[COLUMNA_CONTADOR,COLUMNA_GRUPO, COLUMNA_OPERACION]),
                          df_ia39[["Grupo hojas ruta", "Cont.grupo HRuta","Operación","Txt.brv.operación"]],
                          left_on = "DESCRIPCIÓN DE LA OPERACIÓN",right_on = "Txt.brv.operación",
                          how="left")

  
  df_req[COLUMNA_GRUPO] = df_req[COLUMNA_GRUPO].fillna(
                                                df_req["DESCRIPCIÓN DE LA OPERACIÓN"].map(df_sin_grupo.set_index("Txt.brv.operación")["Grupo hojas ruta"])) 
  df_req[COLUMNA_CONTADOR] = df_req[COLUMNA_CONTADOR].fillna(
                                                df_req["DESCRIPCIÓN DE LA OPERACIÓN"].map(df_sin_grupo.set_index("Txt.brv.operación")["Cont.grupo HRuta"])) 
  df_req[COLUMNA_OPERACION] = df_req[COLUMNA_OPERACION].fillna(
                                                df_req["DESCRIPCIÓN DE LA OPERACIÓN"].map(df_sin_grupo.set_index("Txt.brv.operación")["Operación"])) 

  return df_req 
 
  
if __name__=="__main__":
  print("Hora de inicio: ",datetime.datetime.now())

  
  # Lectura de Excel con requerimientos
  if not os.path.exists(PATH_EXCEL_DATOS):
    raise FileNotFoundError(f"Ruta de Excel no encontrada: {PATH_EXCEL_DATOS}")
  
  try:
    df_req  =  pd.read_excel(PATH_EXCEL_DATOS, sheet_name = PRT_ArchivoHojaNombre)
  except Exception as e:
    print("No encuentra el nombre de la hoja de Excel")
      
  if df_req[COLUMNA_GRUPO].isnull().any():  
    if not os.path.exists(PATH_EXCEL_IA39):
      raise FileNotFoundError(f"Documento de Operaciones no encontrada. Revisar ruta: {PATH_EXCEL_IA39}")
    df_ia39 =  pd.read_excel(PATH_EXCEL_IA39)
    DF_ANALIZAR = busqueda_grupo(df_req,df_ia39)
  else:
    DF_ANALIZAR = df_req
  
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
  
  # CAMBIAR_DATOS = DF_ANALIZAR.iloc[[5]][COLUMNAS_DE_CAMBIO]
  # GRUPO         = DF_ANALIZAR.loc[5, COLUMNA_GRUPO]
  # CONTADOR      = DF_ANALIZAR.loc[5, COLUMNA_CONTADOR]
  # OPERACION     = DF_ANALIZAR.loc[5, COLUMNA_OPERACION]
  # print(CAMBIAR_DATOS)
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
                            "DOCUMENTO MAF",
                            "NUEVA_VERSION",
                            "GRUPO",
                            "CONTADOR",
                            "OPERACION",
                            "ESTADO_EJECUCION_RPA",
                            "cantidad_match",
                            "COLUMNA_MATCH",
                            ]]
  with pd.ExcelWriter(os.path.join(DESCARGA_PATH_REQUERI,'Resultados.xlsx')) as writer:
    DF_ANALIZAR.to_excel(writer, sheet_name="DocumentosMAF_ANALIZADOS",index=False)

    
  print(f"Hora de fin: {datetime.datetime.now()}")
  pantalla_final()