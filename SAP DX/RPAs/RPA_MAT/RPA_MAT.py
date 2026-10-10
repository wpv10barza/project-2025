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
PRT_path_base       = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\05. RPA EXEC"

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
Carpeta_RPA              = "07. MAT"
CarpetaAProcesar         = "REQ-001"
ArchivoExcelNombre       = "DATOS.xlsx"
ArchivoHojaNombre        = "MAT"
PATH_RPA_BASE            = os.path.join(PRT_path_base, Carpeta_RPA)
PATH_EXCEL_DATOS         = os.path.join(PATH_RPA_BASE, CarpetaAProcesar, ArchivoExcelNombre)
DESCARGA_PATH_REQUERI    = os.path.join(PATH_RPA_BASE, CarpetaAProcesar)

COLUMNA_GRUPO            = "GRUPO"
COLUMNA_CONTADOR         = "CONTADOR"
COLUMNA_OPERACION        = "OPERACIÓN"
COLUMNA_MATERIALES       = "MATERIALES"
COLUMNA_CANTIDAD         = "CANTIDAD"
COLUMNA_ESTADO           = "ESTADO_EJECUCION_RPA"
 


def transaccion_IA06(GRUPO_HDR, CONTADOR_HDR, OPERACION_HDR, MATERIALES_HDR):
  # Filtro de transacción 
  session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_IA06
  session.findById("wnd[0]").sendVKey(0) 
  session.findById("wnd[0]/usr/ctxt[0]").text = GRUPO_HDR
  session.findById("wnd[0]/usr/ctxt[1]").text = FILTRO_IA06_PERFIL 
  session.findById("wnd[0]/usr/ctxt[2]").text = ""
  session.findById("wnd[0]/usr/txt[1]").text = CONTADOR_HDR
  session.findById("wnd[0]").sendVKey(0)
  seleccion_ope(OPERACION_HDR, MATERIALES_HDR)
  
  
def seleccion_ope(operacion, MATERIALES_HDR):
  global MENSAJE_ERROR
  ult_posicion = 0   
  flag = 0                          
  # Seleccionar operacion
  try:
    operacion = (operacion // 10) - 1
  except TypeError as e:
    MENSAJE_ERROR = "[ERROR]: No se puede realizar la operación matemática de un string."   
    print(f"{MENSAJE_ERROR}:  {e}")
    return


  try:
    session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3400").getAbsoluteRow(operacion).selected = True
    session.findById("wnd[0]/usr/btn[1]").press()
  except Exception as e:
    MENSAJE_ERROR ="[ERROR]: No existe la fila de operación"
    print(f"{MENSAJE_ERROR} : {e}")
    return
  
  materiales_existentes = obtener_materiales_act()
  materiales_existentes = dict(materiales_existentes)
  # print(materiales_existentes)
  materiales_nuevos     = dict(MATERIALES_HDR)
  # print(materiales_nuevos)
  
  if materiales_existentes:
    ult_posicion = len(materiales_existentes)

    eliminar_materiales = [ (i,materiales_existentes[i]) for i in materiales_existentes.keys() - materiales_nuevos.keys()]
    
    print(eliminar_materiales)
    # Eliminar
    for mat_e, pos_e in eliminar_materiales:
      if flag==0:
        session.findById("wnd[0]/usr/tblSAPLCMDITCTRL_3500").getAbsoluteRow(pos_e).selected = True 
      else:
        session.findById("wnd[0]/usr/tblSAPLCMDITCTRL_3500").getAbsoluteRow(pos_e-1).selected = True
      flag += 1 
      session.findById("wnd[0]/tbar[1]/btn[14]").press()
      
    nueva_pos = ult_posicion - flag 
    if  nueva_pos < 0:   
      ult_posicion = 0 
    else:
      ult_posicion -= flag 
      
  for (material, cantidad) in (MATERIALES_HDR):
    # Actualizar
    if material in materiales_existentes:
      pos = materiales_existentes[material]  
      session.findById(f"wnd[0]/usr/tblSAPLCMDITCTRL_3500/txt[1,{pos}]").text = str(cantidad)
    
    # Asignar
    else:
      session.findById(f"wnd[0]/usr/tblSAPLCMDITCTRL_3500/ctxt[0,{ult_posicion}]").text = str(material) # Material
      session.findById(f"wnd[0]/usr/tblSAPLCMDITCTRL_3500/txt[1,{ult_posicion}]").text = str(cantidad) # Cantidad    
      ult_posicion+=1


  # session.findById("wnd[0]").sendVKey(0)
  # Falta Guardar    
  
def obtener_materiales_act():
  i = 0
  materiales_actuales = []
  posicion = []
  
  print(session.findById("wnd[0]/usr/tblSAPLCMDITCTRL_3500/ctxt[0,0]").text)
  print("")
  while session.findById(f"wnd[0]/usr/tblSAPLCMDITCTRL_3500/ctxt[0,{i}]").text !="":
    materiales_actuales.append(session.findById(f"wnd[0]/usr/tblSAPLCMDITCTRL_3500/ctxt[0,{i}]").text)
    posicion.append(i)
    i+=1
  return list(zip(materiales_actuales,posicion))
 
def SAP_start():
  if not os.path.exists(sap_logon_path):
          raise FileNotFoundError("Ruta de SAP inválida")
  subprocess.Popen(sap_logon_path)
  time.sleep(5)
  
  try:
          SapGuiAuto  = win32com.client.GetObject('SAPGUI')
          application = SapGuiAuto.GetScriptingEngine
  except Exception:
          raise RuntimeError("[ERROR] Error al cargar SAP GUI scripting")
  
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
      
  DF_ANALIZAR = eliminar_espacio(df_req)

  ## Agrupar Material, Cantidad
  DF_ANALIZAR = DF_ANALIZAR.groupby([COLUMNA_GRUPO, COLUMNA_CONTADOR, COLUMNA_OPERACION]).apply( 
                                                                                          lambda x: list(zip(x[COLUMNA_MATERIALES],x[COLUMNA_CANTIDAD]))).reset_index(name='Materiales') 
  DF_ANALIZAR[COLUMNA_ESTADO]=""
  
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
    MATERIALES    = DF_ANALIZAR.loc[index,'Materiales']
    GRUPO         = DF_ANALIZAR.loc[index, COLUMNA_GRUPO]
    CONTADOR      = DF_ANALIZAR.loc[index, COLUMNA_CONTADOR]
    OPERACION     = DF_ANALIZAR.loc[index, COLUMNA_OPERACION]
    try:
      transaccion_IA06(GRUPO, CONTADOR, OPERACION, MATERIALES)    
    except Exception as e:
      DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] = "ERROR_EJECUCIÓN_RPA"
      print(e)
    DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] = "OK" if MENSAJE_ERROR == "" else MENSAJE_ERROR 
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