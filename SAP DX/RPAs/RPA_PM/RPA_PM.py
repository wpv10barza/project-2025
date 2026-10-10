from dotenv import load_dotenv
import os, sys
#Carga variables de .env
path_RPAs = os.path.abspath(os.path.join(os.path.dirname(__file__),".."))
doten_path =os.path.join(path_RPAs,".env")
load_dotenv(doten_path)
sys.path.append(path_RPAs)

from ventana import pantalla_final
import threading
import time, subprocess
import pandas as pd
from itertools import islice
import win32com.client
import numpy as np
import pygetwindow as gw
import pyautogui
import datetime



sap_logon_path      = r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"
PRT_path_base       = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.01 EJECUCION TI\05. RPA EXEC\05. PM"

# Variables constantes
NOMBRE_CONEXION_SAP = "Copper SAP ECC Producción(PS8)"
RSYST_MANDT         = "300"
RSYST_BNAME         = os.getenv("SAP_USER")
RSYST_BCODE         = os.getenv("SAP_PASS")
RSYST_LANGU         = "ES"
TRANSACCION_IP02    = "/NIP02"
MENSAJE_ERROR       = ""
# Variables de configuración
PRT_CarpetaAProcesar     = "REQ-008 Solicitud 04_20250820 (ACTUALIZACION)"
PRT_ArchivoExcelNombre   = "DATOS.xlsx"
PRT_ArchivoHojaNombre    = "PM"
NOMBRE_VENTANA           = "Modif.plan de mantenimiento preventivo: Plan estrategia"

PATH_EXCEL_DATOS         = os.path.join(PRT_path_base, PRT_CarpetaAProcesar, PRT_ArchivoExcelNombre)
DESCARGA_PATH_REQUERI    = os.path.join(PRT_path_base, PRT_CarpetaAProcesar)

COLUMNA_PLAN             = "PLANES DE MANTT"
COLUMNA_ESTADO           = "ESTADO_EJECUCION_RPA"
COLUMNAS_DE_CAMBIO       = [
                          "TEXTO DESCRIPTIVO DE PLAN DE MANTENIMIENTO",
                          "CONTADOR",
                          
                          "FAC.DEC.CONCLUSIÓN RETR.",
                          "TOLERANCIA (+)",
                          "FACT.DEC.CONCL.ANTICIPADA",
                          "TOLERANCIA (-)",
                          "FACT. DE DILATACIÓN",
                          "HORIZONTE APERTURA",
                          "INTERVALO TOMA",
                          "INTERVALO DE TOMA_UNIDAD",
                          "SUJETO A CONCLUSIÓN",
                          "CALENDARIO DE FÁBRICA",
                          
                          "CAMPO CLASIFICACIÓN",
                          "GRUPO AUTORIZACIONES",
                          "FECHA CLAVE"
                      ]

COLUMNAS_ID={
            "FAC.DEC.CONCLUSIÓN RETR.":  r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02/ssub/2/1/2/sub/2/1/2/1/txt[0]", # Fac.dec.conclusión retr.
            "TOLERANCIA (+)":            r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02/ssub/2/1/2/sub/2/1/2/1/txt[1]", # Tolerancia (+)
            "FACT.DEC.CONCL.ANTICIPADA": r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02/ssub/2/1/2/sub/2/1/2/1/txt[2]", # Fact.dec.concl.anticipada
            "TOLERANCIA (-)":            r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02/ssub/2/1/2/sub/2/1/2/1/txt[3]", # Tolerancia (-)
            "FACT. DE DILATACIÓN":       r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02/ssub/2/1/2/sub/2/1/2/1/txt[4]", # Facto de dilatación
            "CALENDARIO DE FÁBRICA":     r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02/ssub/2/1/2/sub/2/1/2/1/ctxt[0]", # Calendario de fábrica
            
            "HORIZONTE APERTURA":        r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02/ssub/2/1/2/sub/2/1/2/1/txt[5]", # Horizonte apertura
            "INTERVALO DE TOMA":         r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02/ssub/2/1/2/sub/2/1/2/1/txt[7]", # Intervalo toma    
            "INTERVALO DE TOMA_UNIDAD":  r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02/ssub/2/1/2/sub/2/1/2/1/ctxt[1]", # Unidad de intervalo de toma
            "SUJETO A CONCLUSIÓN":       r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02/ssub/2/1/2/sub/2/1/2/1/chk",
          
            # "INDICADOR DE PROGRAMACIÓN": r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02/ssub/2/1/2/sub/2/1/2/1/rad[0]").select
            "CAMPO CLASIFICACIÓN":       r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\03/ssub/2/1/2/sub/2/1/2/1/cmb[0]",
            "GRUPO AUTORIZACIONES":      r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\03/ssub/2/1/2/sub/2/1/2/1/ctxt[0]",
            "FECHA CLAVE":               r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\03/ssub/2/1/2/sub/2/1/2/1/ctxt[1]"   
          }


radio_ids = [
  r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02/ssub/2/1/2/sub/2/1/2/1/rad[0]",
  r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02/ssub/2/1/2/sub/2/1/2/1/rad[1]",
  r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02/ssub/2/1/2/sub/2/1/2/1/rad[2]"
]

def llenar_campos(texto):
  ventanas_act = [v for v in gw.getAllTitles() if NOMBRE_VENTANA in v]
  ventanas     = gw.getWindowsWithTitle(ventanas_act[0])
  print(ventanas)
  if ventanas:
      ventana = ventanas[0]
      try:
          # ventana.restore()
          ventana.activate()
          time.sleep(0.1)
          pyautogui.press('tab',presses=4,interval=0.01)
          time.sleep(0.1)
          pyautogui.hotkey('ctrl','a')
          # time.sleep(0.1)
          # pyautogui.press('delete')
          time.sleep(0.1)
          pyautogui.write(texto, interval=0.01)
      except Exception as e:
          print(f"Error al manejar la ventana: {e}")
          
      time.sleep(0.3) 
  time.sleep(0.3)
  
  
  
    
  

def transaccion_IP02(PLAN_MANTTO, DF_CAMBIAR):
  global MENSAJE_ERROR
  DES_PLAN_MANT = DF_CAMBIAR["TEXTO DESCRIPTIVO DE PLAN DE MANTENIMIENTO"].iloc[0]
  contador = False
  

  # Plan de mantenmiento
  session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_IP02
  session.findById("wnd[0]").sendVKey(0) 
  
  session.findById("wnd[0]/usr/ctxt").text = PLAN_MANTTO #Plan de mantenimiento
  session.findById("wnd[0]").sendVKey(0) 
  
  #------------------------ CAMPO de Texto breve de Planes --------------------------
  session.findById("wnd[0]/usr/sub/1/txt[1]").text = "" if (pd.notna(DES_PLAN_MANT)) and (DES_PLAN_MANT=="DEJAR EN BLANCO")     else(DES_PLAN_MANT if pd.notna(DES_PLAN_MANT) else session.findById("wnd[0]/usr/sub/1/txt[1]").text) # Texto descriptivo de plan de mantenimiento
  
  
  #-------------------Parámetros de programación de planes de mantenimiento------------------------------- 
  try:
    session.findById(r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\01/ssub/2/1/2/sub/2/1/2/1/ctxt")
    contador = True 
  except:
    contador = False
    
  
  session.findById(r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02").select()
  if contador:  
    for col_cam in COLUMNAS_DE_CAMBIO:
      if col_cam not in dict(islice(COLUMNAS_ID.items(),0,10)):
        continue
      
      control = session.findById(COLUMNAS_ID[col_cam])
      if  col_cam == "SUJETO A CONCLUSIÓN":
        if DF_CAMBIAR[col_cam].iloc[0]=="SI":
          control.selected = True   
        else:
          control.selected = False   
      
      elif col_cam in COLUMNAS_DE_CAMBIO[:-3]: ###        
        val = DF_CAMBIAR[col_cam].iloc[0]
        if pd.notna(val) and col_cam == "INTERVALO DE TOMA_UNIDAD":
          session.findById(r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\02/ssub/2/1/2/sub/2/1/2/1/ctxt").text = val
        elif pd.notna(val):
          control.text = val 
  
  else:    
    for i, radio_id in enumerate(radio_ids):
      if session.findById(radio_id).selected:
        if i==1:
          MENSAJE_ERROR = "No es posible la programación exacta por día fijado"
          return       
        
        for col_cam in COLUMNAS_DE_CAMBIO:
          if col_cam not in dict(islice(COLUMNAS_ID.items(),0,9)):
            continue
            
          control = session.findById(COLUMNAS_ID[col_cam])
          if  col_cam == "SUJETO A CONCLUSIÓN":
            if DF_CAMBIAR[col_cam].iloc[0]=="SI":
              control.selected = True   
            else:
              control.selected = False   
          else: 
            if i ==0 and col_cam == "FACT. DE DILATACIÓN":
              continue
            else: 
              val = DF_CAMBIAR[col_cam].iloc[0]
              if pd.notna(val):
                control.text = val

  if any(col in COLUMNAS_DE_CAMBIO for col in ["CAMPO CLASIFICACIÓN", "GRUPO AUTORIZACIONES", "FECHA CLAVE"]):       
    session.findById(r"wnd[0]/usr/sub/2/tabsTABSTRIP_HEAD/tabpT\03").select()
    time.sleep(1)
    for llave, val in COLUMNAS_ID.items():
      if llave in ["CAMPO CLASIFICACIÓN", "GRUPO AUTORIZACIONES", "FECHA CLAVE"]:
        if llave == "CAMPO CLASIFICACIÓN":
          CAMPO_CLAS = DF_CAMBIAR["CAMPO CLASIFICACIÓN"].iloc[0]
          if pd.isna(CAMPO_CLAS) :
            continue
          llenar_campos(CAMPO_CLAS)
        
        elif llave == "GRUPO AUTORIZACIONES":
          GRUPO_AUTO = DF_CAMBIAR["GRUPO AUTORIZACIONES"].iloc[0]
          if pd.isna(GRUPO_AUTO) :
            continue
          session.findById(val).text = "" if (pd.notna(GRUPO_AUTO)) and (GRUPO_AUTO=="DEJAR EN BLANCO") else(GRUPO_AUTO if pd.notna(GRUPO_AUTO) else session.findById(val).text)

        elif llave =="FECHA CLAVE":
          FECHA_CLAVE = DF_CAMBIAR["FECHA CLAVE"].iloc[0]
          if pd.isna(FECHA_CLAVE) :
            continue
          session.findById(val).text = "" if (pd.notna(FECHA_CLAVE)) and (FECHA_CLAVE=="DEJAR EN BLANCO") else(FECHA_CLAVE if pd.notna(FECHA_CLAVE) else session.findById(val).text)

    # print("breakpoint")

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
    df_req  =  pd.read_excel(PATH_EXCEL_DATOS, sheet_name = PRT_ArchivoHojaNombre)
    df_req  =  df_req.replace(r'^\s*$',np.nan,regex=True).infer_objects(copy=False)
    COLUMNAS_DE_CAMBIO = df_req.columns.to_list()  
    DF_ANALIZAR = eliminar_espacio(df_req)
    DF_ANALIZAR[COLUMNA_ESTADO]=""
    
    
  except Exception as e:
    print("No encuentra el nombre de la hoja de Excel")
  
# ------------------- Inicialización de SAP -----------------
  session = SAP_start()
      
  # Registro de cuenta SAP
  session.findById("wnd[0]/usr/txtRSYST-MANDT").text = RSYST_MANDT
  session.findById("wnd[0]/usr/txtRSYST-BNAME").text = RSYST_BNAME
  session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = RSYST_BCODE
  session.findById("wnd[0]/usr/txtRSYST-LANGU").text = RSYST_LANGU
  session.findById("wnd[0]").sendVKey(0)

  for index in range(len(DF_ANALIZAR)):
    
    if (DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] == "OK"):
      continue
    
    MENSAJE_ERROR = ""
    CAMBIAR_DATOS = DF_ANALIZAR.iloc[[index]][COLUMNAS_DE_CAMBIO]
    PLAN          = DF_ANALIZAR.loc[index, COLUMNA_PLAN]
    transaccion_IP02(PLAN, CAMBIAR_DATOS)
    # try:
    #   transaccion_IP02(PLAN, CAMBIAR_DATOS)  
    # except Exception as e:
    #   DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] = "ERROR_EJECUCIÓN_RPA"
    #   print(e)
          
    DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] = "OK" if MENSAJE_ERROR == "" else MENSAJE_ERROR 
  
  DF_ANALIZAR = DF_ANALIZAR[[
                            "PLANES DE MANTT",
                            "ESTADO_EJECUCION_RPA",
                            ]]
  with pd.ExcelWriter(os.path.join(DESCARGA_PATH_REQUERI,'Resultados.xlsx')) as writer:
    DF_ANALIZAR.to_excel(writer, sheet_name="DocumentosMAF_ANALIZADOS",index=False)
    
  print(f"Hora de fin: {datetime.datetime.now()}")
  pantalla_final()