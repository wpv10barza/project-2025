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
import datetime
import pygetwindow as gw
import pyautogui
import threading
import pyperclip
import re
import shutil
sap_logon_path      = r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"
PRT_path_base       = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\05. RPA EXEC"

# Variables constantes
NOMBRE_CONEXION_SAP = "Copper SAP ECC Producción(PS8)"
RSYST_MANDT         = "300"
RSYST_BNAME         = os.getenv("SAP_USER")
RSYST_BCODE         = os.getenv("SAP_PASS")
RSYST_LANGU         = "ES"
TRANSACCION_CV04N   = "/NCV04N"
PMA_FIELD           = "PMA" # Clase de documento
DOC_PARCIAL         = "000" # Documento parcial
PATH_DESCARGA_PRT   = r'C:\temp'

# Variables de configuración
coord_x = 694
coord_y = 679
Carpeta_RPA              = "09. MAF-DL"
PRT_CarpetaAProcesar     = "REQ-PRUEBA"
PRT_ArchivoExcelNombre   = "HDR con MAF.xlsx"
PRT_ArchivoHojaNombre    = "Resultados"
PRT_ArchivoHojaMAF       = "MAF por EQ"
NOMBRE_CV40N_DESCARGA    = "CV04N.xlsx"
NOMBRE_VENTANA_EMERGENTE = 'Seguridad SAP GUI'
NOMBRE_VENTANA_GUARDADO  = 'Esta página desea guardar'
NOMBRE_GUARDAR_COMO      = 'Confirmar Guardar como'
NOMBRE_CVN40N_DESCARGA   = "CV04N.xlsx"
PATH_RPA_BASE            = os.path.join(PRT_path_base, Carpeta_RPA)
PATH_EXCEL_DATOS         = os.path.join(PATH_RPA_BASE, PRT_CarpetaAProcesar, PRT_ArchivoExcelNombre)
DESCARGA_PATH_REQUERI    = os.path.join(PATH_RPA_BASE, PRT_CarpetaAProcesar)
COLUMNA_DOC_MAF          = "DocumentoMAF"
COLUMNA_UT_EQUI          = "Ubicación técnica"

def manejar_venta_reemplazo(stop_event):
  while not stop_event.is_set():
    ventanas = gw.getWindowsWithTitle(NOMBRE_GUARDAR_COMO)
    if ventanas:
        ventana = ventanas[0]
        try:
            ventana.activate()
            time.sleep(0.1)
            pyautogui.press('left')
            time.sleep(0.1)
            pyautogui.press('enter')
            #print("Ventana 'Seguridad SAP GUI' gestionada.")
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
              time.sleep(0.1)
              pyautogui.press('enter')
              #print("Ventana 'Seguridad SAP GUI' gestionada.")
          except Exception as e:
              print(f"Error al manejar la ventana: {e}")
          time.sleep(0.3) 
      time.sleep(0.3)

def limpiar_texto(x):
    x = re.sub(r'[\\\/:\*\?"<>\|-]', '', x)
    x = re.sub(r'[.]', ' ', x)
    x = re.sub(r'\s+', ' ', x).strip()
    return x

def descarga_PRT_CV04N(df_documentosMAF, path_descarga):
    df_result = []
    global MENSAJE_ERROR
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
    

    # Descarga de PRT
    for fila in range(len(df_documentosMAF)):
      MENSAJE_ERROR = ""
      try:
        if fila != 0:
          session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").currentCellRow = fila
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").selectedRows = str(fila)
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").doubleClickCurrentCell()
        Texto_Doc = session.findById("wnd[0]/usr/ctxt[0]").text
        Descr_Doc = session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/txt[0]").text
        try:
          session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell[1]").selectNode("          1")
          session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell[1]").doubleClickNode("          1")
        except Exception as e:
          print("No existe nodo 1. Probando nodo 2 ...")
          try:
              session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell[1]").selectNode("          2")
              session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell[1]").doubleClickNode("          2")                
          except Exception as e:
              print("No existe nodo 2 o no está disponible", e)

        time.sleep(10)
        Descr_Doc = limpiar_texto(Descr_Doc)
        pyautogui.moveTo(coord_x,coord_y)
        pyautogui.click()
        pyautogui.moveTo(coord_x-32,coord_y+202)
        pyautogui.click()
        pyautogui.hotkey('ctrl', 's')
        time.sleep(5)
        ventanas = [v for v in gw.getAllWindows() if 'desea guardar' in v.title.lower()]
        if ventanas:
          # pyautogui.write(Descr_Doc)
          pyperclip.copy(Descr_Doc)
          pyautogui.hotkey('ctrl', 'v')
          pyautogui.press("tab",presses=7, interval=0.1)
          pyautogui.press("enter")
          pyperclip.copy(path_descarga)
          pyautogui.hotkey('ctrl', 'v')
          pyautogui.press("enter")
          pyautogui.press("tab",presses=10, interval=0.1)
          pyautogui.press("enter")
        pyautogui.hotkey('ctrl', 'w')
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

      except:
        MENSAJE_ERROR = "[Error] No se descargó corrrectamente"
      finally:
        df_result.append({
            'Documento MAF'    : Texto_Doc,
            'Descripción'      : Descr_Doc,
            'Resultado de descarga'    : MENSAJE_ERROR 
        })
    df_result = pd.DataFrame(df_result)
    
    archivos_pdf = [os.path.splitext(file)[0] for file in os.listdir(path_descarga) if file.endswith('.pdf')]
    df_result['Resultado de descarga'] = df_result['Descripción'].isin(archivos_pdf)
    return df_result


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
 

def carpeta_por_UT(df_HDR, df_DL_MAF):
  df_HDR['DocumentoMAF'] = df_HDR['DocumentoMAF'].astype('Int64').astype(str)
  df_DL_MAF['Documento MAF'] = df_DL_MAF['Documento MAF'].astype(str)

  df_descargados = pd.merge(df_HDR, df_DL_MAF, left_on='DocumentoMAF', right_on='Documento MAF', how='right')
  df_descargados = df_descargados.groupby('Ubicación técnica').apply(lambda x: list(x['Descripción'])).reset_index(name='Descripciones por UT')
  
  for _, row in df_descargados.iterrows(): 
    carpeta = os.path.join(PATH_DESCARGA_PRT, row['Ubicación técnica'] )
    if not os.path.exists(carpeta):
     os.mkdir(carpeta)
      
  
  archivos_pdf   = [os.path.splitext(file)[0] for file in os.listdir(PATH_DESCARGA_PRT) if file.endswith('.pdf')]

  for _, row in df_descargados.iterrows():
    for descripcion in row['Descripciones por UT']:
      if descripcion in archivos_pdf:
        archivo_origen = os.path.join(PATH_DESCARGA_PRT, f"{descripcion}.pdf")
        carpeta_destino = os.path.join(PATH_DESCARGA_PRT, row['Ubicación técnica'], f"{descripcion}.pdf")
        try:
          shutil.move(archivo_origen, carpeta_destino)
        except FileNotFoundError:
          pass 

if __name__=="__main__":
  print("Hora de inicio: ",datetime.datetime.now())
  df_result_total = pd.DataFrame()
  stop_event_1 = threading.Event()
  stop_event_2 = threading.Event()
  handler_thread_1 = threading.Thread(target=manejar_ventana_seguridad, args=(stop_event_1,))
  handler_thread_2 = threading.Thread(target=manejar_venta_reemplazo, args=(stop_event_2,))
  handler_thread_1.daemon = True
  handler_thread_2.daemon = True
  handler_thread_1.start()
  handler_thread_2.start()

  # Lectura de Excel con requerimientos
  if not os.path.exists(PATH_EXCEL_DATOS):
    raise FileNotFoundError(f"Ruta de Excel no encontrada: {PATH_EXCEL_DATOS}")
  
  try:
    df_prts     =  pd.read_excel(PATH_EXCEL_DATOS, sheet_name = PRT_ArchivoHojaNombre)
    df_maf_equ  =  pd.read_excel(PATH_EXCEL_DATOS, sheet_name = PRT_ArchivoHojaMAF )

  except Exception as e:
    print("No encuentra el nombre de la hoja de Excel")

  # Inicialización de SAP
  session = SAP_start()
  

  # Descarga de PRT
  equipos     = df_prts[COLUMNA_UT_EQUI].unique()
  for equi in equipos:

    carpeta_base = os.path.join(PATH_DESCARGA_PRT, equi)
    os.makedirs(carpeta_base, exist_ok=True)

    os.makedirs(os.path.join(carpeta_base, "ORIEN"), exist_ok=True)
    carpeta_sap = os.path.join(carpeta_base, "SAP")
    os.makedirs(carpeta_sap, exist_ok=True)

    carpeta_MAF = os.path.join(carpeta_sap, "Doc MAF")
    os.makedirs(os.path.join(carpeta_sap, "Reportes"), exist_ok=True)
    os.makedirs(carpeta_MAF, exist_ok=True)

    try:
      df_result = descarga_PRT_CV04N( df_prts[df_prts[COLUMNA_UT_EQUI] == equi][COLUMNA_DOC_MAF], carpeta_MAF)
      df_result['Ubicación técnica'] = equi            

    # df_result = pd.merge(df_result, )

      df_result_total = pd.concat([df_result_total, df_result], ignore_index=True)
    
      pdfs      = [f for f in os.listdir(carpeta_MAF) if f.endswith('.pdf')]
      df_maf_equ.loc[df_maf_equ[COLUMNA_UT_EQUI] == equi, '#MAF descargados'] = len(pdfs)
      with pd.ExcelWriter(os.path.join(DESCARGA_PATH_REQUERI,'Resultados de RPA.xlsx')) as writer:
        df_result_total.to_excel(writer, sheet_name="Resultados",index=False)
        df_maf_equ.to_excel(writer, sheet_name="CantidadMAF_Descargados",index=False)
    except:
      print(f"Error al descargar PRT del equipo {equi}")


  


  # Crear carpetas por Ub.Tec
  # carpeta_por_UT(df_prts , df_result)


  print(f"Hora de fin: {datetime.datetime.now()}")
  pantalla_final()