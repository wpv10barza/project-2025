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
PRT_path_base       = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.01 EJECUCION TI\05. RPA EXEC"

# Variables constantes
NOMBRE_CONEXION_SAP = "Copper SAP ECC Producción(PS8)"
RSYST_MANDT         = "300"
RSYST_BNAME         = os.getenv("SAP_USER")
RSYST_BCODE         = os.getenv("SAP_PASS")
RSYST_LANGU         = "ES"
TRANSACCION_IP02    = "/NIP02"
MENSAJE_ERROR       = ""

# Variables de configuración
CarpetaAProcesar         = "REQ-010 Solicitud 4_20250820 (ACTUALIZACION)"
ArchivoExcelNombre       = "DATOS.xlsx"
ArchivoExcelNombre_IP118 = "IP18.xlsx"
Carpeta_RPA              = "06. PPM"
PRT_ArchivoHojaNombre    = "PPM"
PATH_RPA_BASE            = os.path.join(PRT_path_base, Carpeta_RPA)
PATH_EXCEL_DATOS         = os.path.join(PATH_RPA_BASE, CarpetaAProcesar, ArchivoExcelNombre)
PATH_EXCEL_IP18          = os.path.join(PATH_RPA_BASE, CarpetaAProcesar, ArchivoExcelNombre_IP118)
DESCARGA_PATH_REQUERI    = os.path.join(PATH_RPA_BASE, CarpetaAProcesar)

COLUMNA_PLAN             = "PLANES DE MANTT"
COLUMNA_POSICION         = "POSICION"
COLUMNA_ESTADO           = "ESTADO_EJECUCION_RPA"
COLUMNA_CANTIDAD_POS     = "CANTIDAD DE POSICIONES"
COLUMNAS_DE_CAMBIO       = [
                          "NO LIBERAR INMEDIATAMENTE",
                          "UBICACIÓN TÉCNICA",
                          "EQUIPO",
                          "CONJUNTO",
                          "TEXTO DE MANT",
                          "CENTRO DE PLANIFICACIÓN",
                          "GRUPO DE PLANIFICACIÓN",
                          "CLASE DE ORDEN",
                          "CLASE ACTIVIDAD PM",
                          "PUESTO DE TRABAJO",
                          "CENTRO DE MANTENIMIENTO AUXILIAR",
                          "DIVISIÓN",
                          "PRIORIDAD",
                          "TIPO",
                          "GRUPO HDR",
                          "CONTADOR HDR",
                          
                          "NORMA ESTANDAR",
                          "IMPACTO",
                          "ALTERNATIVE TOLERANCE (DAYS)"

                      ]

COLUMNAS_ID={
            "NO LIBERAR INMEDIATAMENTE":        "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/4/chk",
            "UBICACIÓN TÉCNICA":                "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/2/ctxt[0]",
            "EQUIPO":                           "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/2/ctxt[1]",
            "CONJUNTO":                         "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/2/ctxt[2]",
            "TEXTO DE MANT":                    "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/1/txt[1]",
            "CENTRO DE PLANIFICACIÓN":          "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/4/ctxt[0]",
            "GRUPO DE PLANIFICACIÓN":           "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/4/ctxt[1]",
            "CLASE DE ORDEN":                   "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/4/ctxt[2]",
            "CLASE ACTIVIDAD PM":               "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/4/ctxt[3]",
            "PUESTO DE TRABAJO":                "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/4/ctxt[4]",
            "CENTRO DE MANTENIMIENTO AUXILIAR": "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/4/ctxt[5]",
            "DIVISIÓN":                         "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/4/ctxt[6]",
            "PRIORIDAD":                        "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/4/cmb",
            "TIPO":                             "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/4/txt[8]",
            "GRUPO HDR":                        "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/4/txt[9]",
            "CONTADOR HDR":                     "wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\\11/ssub/3/1/2/sub/3/1/2/4/txt[10]", 
          }

# Campos relativos a la ventana de extensión
COLUMNA_CAMBIOS_EXT = [
                          "NORMA ESTANDAR",
                          "IMPACTO",
                          "ALTERNATIVE TOLERANCE (DAYS)"
                      ]

COLUMNAS_ID_EXT ={
                "NORMA ESTANDAR" :r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\17/ssub/3/1/2/sub/3/1/2/1/ctxt[0]",
                "IMPACTO" : r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\17/ssub/3/1/2/sub/3/1/2/1/ctxt[1]",
                "ALTERNATIVE TOLERANCE (DAYS)" : r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\17/ssub/3/1/2/sub/3/1/2/1/txt[2]"
                }


# Depende de los campos que manejen 
COLUMNA_PRIORIDAD = {
                  "Urgente. Resp Inmedi":"1",
                  "N/A"                 :"2",
                  "Urgente.Resp o Conte":"3",
                  "Urg. Resp ant fe req":"A",
                  "Urge.Res o Cont fech":"B",
                  "No urg Prg ant fe rq":"C"
                      }

# LIMITE_COLUMNAS = {
#                   "NO LIBERAR INMEDIATAMENTE" :,
#                   "UBICACIÓN TÉCNICA"        :,
#                   "EQUIPO" :,
#                   "CONJUNTO":,
#                   "TEXTO DE MANT":,
#                   "CENTRO DE PLANIFICACIÓN":,
#                   "GRUPO DE PLANIFICACIÓN":,
#                   "CLASE DE ORDEN":,
#                   "CLASE ACTIVIDAD PM":,
#                   "PUESTO DE TRABAJO":,
#                   "CENTRO DE MANTENIMIENTO AUXILIAR":,
#                   "DIVISIÓN":,
#                   "PRIORIDAD":,
#                   "TIPO":,
#                   "GRUPO HDR",
#                   "CONTADOR HDR"
#                   }


def transaccion_IP02(PLAN_MANTTO, DF_CAMBIAR, POSICION, CANTIDAD_POS):
  global MENSAJE_ERROR
  pos_ant = ""
  
  def llena_campos_pos():
    global MENSAJE_ERROR
    
    for col_cam in COLUMNAS_DE_CAMBIO:
        if col_cam not in COLUMNAS_ID:
          continue
        
        try:    
          control = session.findById(COLUMNAS_ID[col_cam])
          val = DF_CAMBIAR[col_cam].iloc[0]
          
          if  col_cam == "NO LIBERAR INMEDIATAMENTE": 
            if DF_CAMBIAR[col_cam].iloc[0] == "SI":
              control.selected = True   
            else:
              control.selected = False   
          
          elif col_cam == "PRIORIDAD":
            if val in COLUMNA_PRIORIDAD:
              control.key = COLUMNA_PRIORIDAD[val]
            elif pd.isna(val):
              control.key = "2"
          
          elif pd.notna(val):  
            control.text = val 
            
        except:
          MENSAJE_ERROR = f"No se pudo actualizar el campo {col_cam}"
          return
  
  # Plan de mantenmiento
  session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_IP02
  session.findById("wnd[0]").sendVKey(0) 
  
  session.findById("wnd[0]/usr/ctxt").text = PLAN_MANTTO #Plan de mantenimiento
  session.findById("wnd[0]").sendVKey(0) 
  
  print(PLAN_MANTTO,  CANTIDAD_POS) 

  #-------------------Actualización de campos de posición ------------------------------- 
  if CANTIDAD_POS == 1:
    llena_campos_pos()
  # Seleccionar primera posición en caso exista más de una posición
  elif CANTIDAD_POS > 1:
    session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\10/ssub/3/1/2/sub/3/1/2/1/tblSAPLIWP3TCTRL_0220").getAbsoluteRow(0).selected = True
    session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11").select()
    

    while ( session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/1/txt[0]").text != str(POSICION)) :
      pos_ant = session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/1/txt[0]").text
      session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/1/btn[6]").press() #siguiente posición
      
      if session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\11/ssub/3/1/2/sub/3/1/2/1/txt[0]").text == pos_ant:
        MENSAJE_ERROR = f"No se encontró ninguna Posición {POSICION} dentro del Plan {PLAN_MANTTO}"
        return 

    # Campos de Planes de mantenimiento
    llena_campos_pos()
   
    
  # Campo de extensión de cliente
  col_ext = [col_cam for col_cam in COLUMNAS_DE_CAMBIO if col_cam in COLUMNAS_ID_EXT]
  if not col_ext:
    return
  try:
    session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\17").select()
    try:
      session.findById(r"wnd[0]/usr/sub/3/tabsTABSTRIP_ITEM/tabpT\17/ssub/3/1/2/sub/3/1/2/1/ctxt[0]")
      for col_cam in col_ext:
        val = DF_CAMBIAR[col_cam].iloc[0]
        if pd.notna(val):
          session.findById(COLUMNAS_ID_EXT[col_cam]).text = val
    except:
      MENSAJE_ERROR = "No existe Ubicación técnica, no se puede modificar los campos de Extensión de cliente"
      return
  except:
    MENSAJE_ERROR = "No existe ventana de 'Extensión de cliente' "
    return
  session.findById("wnd[0]/tbar[0]/btn[11]").press()
      
      

      
def eliminar_espacio(df):
  for columna in df.columns:
    df[columna] = df[columna].apply(lambda x: x.strip() if isinstance(x,str) else x)
    # df[columna] = df[columna].apply(lambda x: x.replace("  "," ") if isinstance(x,str) else x)
  return df    
       

def cantidad_pos(DF_ANALIZAR, df_ip18):
  conteo_posiciones = df_ip18["Plan mantenim."].value_counts()  
  DF_ANALIZAR[COLUMNA_CANTIDAD_POS] = DF_ANALIZAR[COLUMNA_PLAN].map(conteo_posiciones).fillna(0).astype(int)
  print(DF_ANALIZAR[COLUMNA_CANTIDAD_POS])
  return DF_ANALIZAR 
      
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
  
  # Lectura de IP18 
  if not os.path.exists(PATH_EXCEL_IP18):
    raise FileNotFoundError(f"No se encontró reporte de Planes IP18: {PATH_EXCEL_IP18}")
  
  try:
    df_req  = pd.read_excel(PATH_EXCEL_DATOS, sheet_name = PRT_ArchivoHojaNombre)
    df_req  = df_req.replace(r'^\s*$',np.nan,regex=True).infer_objects(copy=False)
    df_ip18 = pd.read_excel(PATH_EXCEL_IP18)
    COLUMNAS_DE_CAMBIO = df_req.columns.to_list()  
    DF_ANALIZAR = eliminar_espacio(df_req)
    DF_ANALIZAR[COLUMNA_ESTADO]=""
    
    DF_ANALIZAR = cantidad_pos(DF_ANALIZAR, df_ip18[["Plan mantenim.", "Pos.mantenim."]])
  except Exception as e:
    print("No encuentra el nombre de la hoja de Excel", e)
  
# ------------------- Inicialización de SAP -----------------
  session = SAP_start()
      
  # Registro de cuenta SAP
  session.findById("wnd[0]/usr/txtRSYST-MANDT").text = RSYST_MANDT
  session.findById("wnd[0]/usr/txtRSYST-BNAME").text = RSYST_BNAME
  session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = RSYST_BCODE
  session.findById("wnd[0]/usr/txtRSYST-LANGU").text = RSYST_LANGU
  session.findById("wnd[0]").sendVKey(0)
  
  for index in range(len(DF_ANALIZAR)):
    
    if (DF_ANALIZAR.loc[index, COLUMNA_ESTADO ] == "OK"):
      continue
    
    MENSAJE_ERROR = ""
    CAMBIAR_DATOS = DF_ANALIZAR.iloc[[index]][COLUMNAS_DE_CAMBIO]
    PLAN          = DF_ANALIZAR.loc[index, COLUMNA_PLAN]
    POSICION      = DF_ANALIZAR.loc[index, COLUMNA_POSICION]
    CANTIDAD_POS  = DF_ANALIZAR.loc[index, COLUMNA_CANTIDAD_POS]
    try:
      transaccion_IP02(PLAN, CAMBIAR_DATOS, POSICION, CANTIDAD_POS)  
    except Exception as e:
      DF_ANALIZAR.loc[index, COLUMNA_ESTADO ] = "ERROR_EJECUCIÓN_RPA"
      print(e)
      
    DF_ANALIZAR.loc[index, COLUMNA_ESTADO ] = "OK" if MENSAJE_ERROR == "" else MENSAJE_ERROR 
  
  DF_ANALIZAR = DF_ANALIZAR[[
                            COLUMNA_PLAN,
                            COLUMNA_POSICION ,
                            COLUMNA_ESTADO 
                            ]]
  with pd.ExcelWriter(os.path.join(DESCARGA_PATH_REQUERI,'Resultados.xlsx')) as writer:
    DF_ANALIZAR.to_excel(writer, sheet_name="DocumentosMAF_ANALIZADOS",index=False)
    
  print(f"Hora de fin: {datetime.datetime.now()}")
  pantalla_final()