import os, sys
from dotenv import load_dotenv
path_RPAs = os.path.abspath(os.path.join(os.path.dirname(__file__),".."))
doten_path =os.path.join(path_RPAs,".env")
load_dotenv(doten_path)
sys.path.append(path_RPAs)

from ventana import pantalla_final
import time, subprocess
import pandas as pd
import win32com.client
import pygetwindow as gw
import pyautogui
import threading

sap_logon_path      = r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"
PRT_path_base       = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\05. RPA EXEC"

# Variables constantes
NOMBRE_CONEXION_SAP = "Copper SAP ECC Producción(PS8)"
RSYST_MANDT         = "300"
RSYST_BNAME         = os.getenv("SAP_USER")
RSYST_BCODE         = os.getenv("SAP_PASS")
RSYST_LANGU         = "ES"
TRANSACCION_IA10    = "/NIA10"
PMA_FIELD           = "PMA" # Clase de documento
DOC_PARCIAL         = "000" # Documento parcial
MENSAJE_ERROR       = ""

# Variables de configuración
Carpeta_RPA              = "09. MAF-DL"
PRT_CarpetaAProcesar     = "REQ-PRUEBA"
DESCARGA_ARCH_IA10       = "IA10_MAF.xlsx"
NOMBRE_VENTANA_EMERGENTE = 'Seguridad SAP GUI'
PATH_RPA_BASE            = os.path.join(PRT_path_base, Carpeta_RPA)
DESCARGA_PATH_REQUERI    = os.path.join(PATH_RPA_BASE, PRT_CarpetaAProcesar)



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


def descarga_IA10_MAF(session):
    session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_IA10
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/usr/btn[2]").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/btn[11]").press() # Desmarcar todo
    session.findById("wnd[0]/usr/chk[0]").selected = True # Selección Hoja de Ruta
    session.findById("wnd[0]/usr/chk[5]").selected = True # Selección MAF
    session.findById("wnd[0]/tbar[1]/btn[8]").press()


    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[2,0]").select()
    session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[2,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ssub/1/txt").text = DESCARGA_ARCH_IA10
    session.findById("wnd[1]/tbar[0]/btn[20]").press()
    session.findById("wnd[1]/usr/ctxt[0]").text = DESCARGA_PATH_REQUERI 
    session.findById("wnd[1]/tbar[0]/btn[0]").press()


def main():
    stop_event = threading.Event()
    handler_thread = threading.Thread(target=manejar_ventana_seguridad, args=(stop_event,))
    handler_thread.daemon = True
    handler_thread.start()

    session = SAP_start()
    try: 
      descarga_IA10_MAF(session)
      # pantalla_final()
    except:
      print("Error en la descarga de la IA10_MAF")
    finally:
      stop_event.set()
      handler_thread.join()


if __name__ == "__main__":
    main()
    