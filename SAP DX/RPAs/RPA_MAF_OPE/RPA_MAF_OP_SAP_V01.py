from dotenv import load_dotenv
from ventana import pantalla_final
import threading
import os, time, subprocess
import pandas as pd
import win32com.client
import pygetwindow as gw
import pyautogui
import datetime
#Carga variables de .env
load_dotenv()


sap_logon_path      = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
PRT_path_base       = r"D:\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\05. RPA EXEC\02. MAF-OP"

# Variables constantes
NOMBRE_CONEXION_SAP = "Copper SAP ECC Producción(PS8)"
RSYST_MANDT         = "300"
RSYST_BNAME         = os.getenv("SAP_USER")
RSYST_BCODE         = os.getenv("SAP_PASS")
RSYST_LANGU         = "ES"
TRANSACCION_IA06    = "/NIA06"
TRANSACCION_CV04N   = "/NCV04N"
PMA_FIELD           = "PMA" # Clase de documento
DOC_PARCIAL         = "000" # Documento parcial
CATEGORIA_ARCHIVO   = "CF_IMP_AS" # Categoría de archivo DMS_C1_ST para versión SAP/4HANA (Versión 770)
FILTRO_IA06_PERFIL  = "0000008"
MENSAJE_ERROR       = ""
# Variables de configuración
PRT_CarpetaAProcesar     = "REQ-001 Rociadores"
PRT_ArchivoExcelNombre   = "DATOS.xlsx"
PRT_ArchivoHojaNombre    = "MAF-OP"
IA39_RutaNombre          = "IA39.xlsx"
NOMBRE_CVN40N_DESCARGA   = "CV04N.xlsx"
NOMBRE_VENTANA_EMERGENTE = "Seguridad SAP GUI"
VENTANA_BORRAR_MAF       = "Borrar asignaciones MAF"
PATH_EXCEL_DATOS         = os.path.join(PRT_path_base, PRT_CarpetaAProcesar, PRT_ArchivoExcelNombre)
PATH_EXCEL_IA39          = os.path.join(PRT_path_base, PRT_CarpetaAProcesar, IA39_RutaNombre)
DESCARGA_PATH_REQUERI    = os.path.join(PRT_path_base, PRT_CarpetaAProcesar)

COLUMNA_GRUPO            = "GRUPO"
COLUMNA_CONTADOR         = "CONTADOR"
COLUMNA_OPERACION        = "OPERACION"
COLUMNA_DOC_MAF          = "DOCUMENTO MAF"
COLUMNA_ULT_VERSION      = "ULT_VERSION"


def manejar_ventana_borrar_MAF(stop_event):
    while not stop_event.is_set():
        ventanas = gw.getWindowsWithTitle(VENTANA_BORRAR_MAF)
        if ventanas:
            ventana = ventanas[0]
            try:
                ventana.activate()
                time.sleep(0.1)
                pyautogui.press('tab',presses=2, interval=0.1)
                pyautogui.press('enter')
            except Exception as e:
                print(f"Error al manejar la ventana: {e}")
            time.sleep(0.3) 
        time.sleep(0.3)

def manejar_ventana_seguridad(stop_event):
    while not stop_event.is_set():
        ventanas = gw.getWindowsWithTitle(NOMBRE_VENTANA_EMERGENTE)
        if ventanas:
            ventana = ventanas[0]
            try:
                ventana.activate()
                time.sleep(0.1)
                pyautogui.press('left')
                pyautogui.press('enter')
                #print("Ventana 'Seguridad SAP GUI' gestionada.")
            except Exception as e:
                print(f"Error al manejar la ventana: {e}")
            time.sleep(0.3) 
        time.sleep(0.3)

def match_operaciones(df_prts, df_ia39):
  
  df_prts["ESTADO_EJECUCION_RPA"]=""
  df_prts["COLUMNA_MATCH"] = ""
  
  conteo_matches          = df_ia39["Txt.brv.operación"].value_counts()
  
  df_prts['cantidad_match'] = df_prts["DESCRIPCIÓN DE LA OPERACIÓN"].map(conteo_matches).fillna(0).astype(int)
  df_prts.loc[df_prts['cantidad_match']> 1 , 'COLUMNA_MATCH'] = "TIENE MÁS DE UN MATCH"
  df_prts.loc[df_prts['cantidad_match'] == 0 , 'COLUMNA_MATCH'] = "NO TIENE NINGÚN MATCH"
  df_prts.loc[df_prts['cantidad_match'] == 1 , 'COLUMNA_MATCH'] = "TIENE MATCH"
  
  conteo_matches          = pd.DataFrame(conteo_matches)
  conteo_matches          = conteo_matches[conteo_matches["count"] == 1]


  conteo_matches     = pd.merge(df_ia39,conteo_matches,on="Txt.brv.operación",how="left")
  conteo_matches     = conteo_matches[conteo_matches["count"] == 1]
  df_prts            = pd.merge(df_prts,conteo_matches, left_on = "DESCRIPCIÓN DE LA OPERACIÓN",right_on = "Txt.brv.operación", how="left") 
  df_prts            = df_prts[[
                                          #"ACCIÓN",
                                          "DESCRIPCIÓN DE LA OPERACIÓN",
                                          "DOCUMENTO MAF",
                                          "ULT_VERSION", 
                                          "Grupo hojas ruta",
                                          "Cont.grupo HRuta",
                                          "Operación",
                                          'COLUMNA_MATCH',
                                          'cantidad_match'
                                      ]].rename(columns={
                                                         "Grupo hojas ruta":COLUMNA_GRUPO,
                                                         "Cont.grupo HRuta":COLUMNA_CONTADOR,
                                                         "Operación": COLUMNA_OPERACION
                                                        })
  df_prts.loc[pd.isna(df_prts[COLUMNA_CONTADOR]),"ESTADO_EJECUCION_RPA"] = "ERROR_NO_MATCH"
                           
  df_prts["ULT_VERSION"]    = df_prts["ULT_VERSION"].astype(str)                                      
  df_prts["DOCUMENTO MAF"]    = df_prts["DOCUMENTO MAF"].astype(str) 
  return df_prts

def descarga_ult_version(df_documentosMAF):
    # Cantidad de valores
    valores = len(df_documentosMAF) * 3
    print(valores)
    # Copiar código de documentos MAF en el portapapeles
    codigos_MAF = df_documentosMAF.dropna().astype(int).unique().tolist()
    pd.Series(codigos_MAF).to_clipboard(index=False)

    # Inicio de transacción
    session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_CV04N #Transacción
    session.findById("wnd[0]").sendVKey(0)
    
    # Filtro por código documentos MAF
    session.findById("wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssub/1/2/sub/1/2/1/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Aumentar número de coincidencias
    session.findById("wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssub/1/2/sub/1/2/1/txt[5]").text = valores
    
    # Última versión
    session.findById("wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssub/1/2/sub/1/2/1/rad[2]").select()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    

    # Descarga
    session.findById("wnd[0]").sendVKey(16)
    session.findById("wnd[1]/tbar[0]/btn[20]").press()
    session.findById("wnd[1]/usr/ctxt[0]").text = DESCARGA_PATH_REQUERI
    session.findById("wnd[1]/usr/ctxt[1]").text = NOMBRE_CVN40N_DESCARGA
    session.findById("wnd[1]/usr/ctxt[1]").caretPosition = 5
    session.findById("wnd[1]/tbar[0]/btn[0]").press()  
  
def transaccion_IA06(GRUPO_HDR, CONTADOR_HDR, OPERACION_HDR, documentoMAF, ult_version):
  # Filtro de transacción 
  session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_IA06
  session.findById("wnd[0]").sendVKey(0) 
  session.findById("wnd[0]/usr/ctxt[0]").text = GRUPO_HDR
  session.findById("wnd[0]/usr/ctxt[1]").text = FILTRO_IA06_PERFIL 
  session.findById("wnd[0]/usr/ctxt[2]").text = ""
  session.findById("wnd[0]/usr/txt[1]").text = CONTADOR_HDR
  session.findById("wnd[0]").sendVKey(0)
  
  seleccion_docMAF(documentoMAF, ult_version, OPERACION_HDR)
    
def seleccion_docMAF(documentoMAF,ult_version, operacion):
  global MENSAJE_ERROR
  
  try:
    operacion = (operacion // 10) - 1
  except TypeError as e:
    MENSAJE_ERROR = "[ERROR]: No se puede realizar la operación matemática de un string."   
    print(f"{MENSAJE_ERROR}:  {e}")
    return
  VERSION = "0" + ult_version
  # Selección de la operación                                        
  try:
    session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3400").getAbsoluteRow(operacion).selected = True
    session.findById("wnd[0]/usr/btn[3]").press() # Selección de documento MAF
  except Exception as e:
    MENSAJE_ERROR ="[ERROR]: No existe la fila de operación"
    print(f"{MENSAJE_ERROR} : {e}")
    return
    
  try:
    ventana = session.findById("wnd[1]")
  except:
    ventana = None

  if ventana: #and session.findById("wnd[1]").Text == "Modificar instrucción: nuevas entradas MAF":
   escribir_documento_MAF(documentoMAF,VERSION)
    # Selección de documento MAF anterior
  else:
    try:
      session.findById("wnd[0]/usr/tblSAPLCFDITCTRL_0102").getAbsoluteRow(0).selected = True
      session.findById("wnd[0]/tbar[1]/btn[14]").press() # Borrar documento MAF 
      time.sleep(2)
      session.findById("wnd[0]").sendVKey(3)
      session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3400").getAbsoluteRow(operacion).selected = True
      session.findById("wnd[0]/usr/btn[3]").press() # Selección de documento MAF
      time.sleep(1)    
      escribir_documento_MAF(documentoMAF,VERSION)
    except Exception as e:
      MENSAJE_ERROR = "[ERROR]: No existe un documento MAF asociado a la operación; no se pudo borrar ningún elemento" 
      print(f"{MENSAJE_ERROR} : {e}")
      return

def escribir_documento_MAF(documentoMAF,VERSION):
  global MENSAJE_ERROR
  try:
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[1]/usr/ctxt[0]").text = documentoMAF
    session.findById("wnd[1]/usr/ctxt[1]").text = PMA_FIELD 
    session.findById("wnd[1]/usr/txt[3]").text = DOC_PARCIAL
    session.findById("wnd[1]/usr/txt[4]").text = VERSION[-2:]
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(11) # Guardar cambios totales
  except Exception as e:
    MENSAJE_ERROR = f"[ERROR]: No se pudo concretar la operación para el documento {documentoMAF}" 
    print(f"{MENSAJE_ERROR}: {e}")


if __name__=="__main__":
  print("Hora de inicio: ",datetime.datetime.now())
  # Inicializar hilos
  stop_event = threading.Event()
  handler_thread = threading.Thread(target=manejar_ventana_borrar_MAF, args=(stop_event,))
  handler_thread2 = threading.Thread(target=manejar_ventana_seguridad, args=(stop_event,))
  
  handler_thread.daemon = True 
  handler_thread2.daemon = True 
  
  handler_thread.start()
  handler_thread2.start()
  
  # Lectura de Excel con requerimientos
  if not os.path.exists(PATH_EXCEL_DATOS):
    raise FileNotFoundError(f"Ruta de Excel no encontrada: {PATH_EXCEL_DATOS}")
  try:
    df_prts  =  pd.read_excel(PATH_EXCEL_DATOS, sheet_name = PRT_ArchivoHojaNombre)
  except Exception as e:
    print("No encuentra el nombre de la hoja de Excel")
    
  if not os.path.exists(PATH_EXCEL_IA39):
    raise FileNotFoundError(f"Documento de Operaciones no encontrada. Revisar ruta: {PATH_EXCEL_IA39}")
  df_ia39 =  pd.read_excel(PATH_EXCEL_IA39)
  
  
  
  # Inicialización de SAP
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
  
  if os.path.exists(os.path.join(os.path.join(DESCARGA_PATH_REQUERI, NOMBRE_CVN40N_DESCARGA))):
    os.remove(os.path.join(DESCARGA_PATH_REQUERI, NOMBRE_CVN40N_DESCARGA))
  descarga_ult_version(df_prts[COLUMNA_DOC_MAF])
  
  df_CV04N = pd.read_excel(os.path.join(DESCARGA_PATH_REQUERI, NOMBRE_CVN40N_DESCARGA)) 
  df_prts[COLUMNA_ULT_VERSION]  = df_prts[COLUMNA_DOC_MAF].map(df_CV04N.set_index("Documento")["Versión documento"]).astype(int)

  DF_ANALIZAR = match_operaciones(df_prts, df_ia39)
  
  for index in range(len(DF_ANALIZAR)):
    
    if DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] == "ERROR_NO_MATCH":
      continue
    
    MENSAJE_ERROR = ""
    
    GRUPO       = DF_ANALIZAR.loc[index, COLUMNA_GRUPO]
    CONTADOR    = DF_ANALIZAR.loc[index, COLUMNA_CONTADOR]
    OPERACION   = DF_ANALIZAR.loc[index, COLUMNA_OPERACION]
    DOC_MAF     = DF_ANALIZAR.loc[index, COLUMNA_DOC_MAF]
    ULT_VERSION = DF_ANALIZAR.loc[index, COLUMNA_ULT_VERSION]
    
    # transaccion_IA06(GRUPO, CONTADOR, OPERACION, DOC_MAF, ULT_VERSION)  

    
    DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] = "OK" if MENSAJE_ERROR == "" else MENSAJE_ERROR 

  with pd.ExcelWriter(os.path.join(DESCARGA_PATH_REQUERI,'Resultados.xlsx')) as writer:
    DF_ANALIZAR.to_excel(writer, sheet_name="DocumentosMAF_ANALIZADOS",index=False)

    
  print(f"Hora de fin: {datetime.datetime.now()}")
  pantalla_final()