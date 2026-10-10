import win32com.client
import os
import time
import subprocess
from dotenv import load_dotenv
import pygetwindow as gw
import pyautogui
import pandas as pd
load_dotenv()

sap_logon_path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
excel_file_path = r"D:\RecordsSAP"

# Variables Constantes
NOMBRE_CONEXION_SAP  = "Copper SAP ECC Producción(PS8)"
RSYST_MANDT          = "300"
RSYST_BNAME          = os.getenv("SAP_USER")
RSYST_BCODE          = os.getenv("SAP_PASS")
RSYST_LANGU          = "ES"
PMA_FIELD            = "PMA" # Clase de documento
DOC_PARCIAL          = "000" # Documento parcial
CATEGORIA_ARCHIVO    = "CF_IMP_AS" # Categoría de archivo DMS_C1_ST para versión SAP/4HANA (Versión 770)
CENTRO_EMPLAZAMIENTO = "C601"

# Variables de Configuración
NOMBRE_VENTANA_EMERGENTE   = "Seguridad SAP GUI"
NOMBRE_VENTANA_TRANSACCION = "Lista de puestos de trabajo" # Nombre de la ventana que requiere atajo de teclados para descargar información
TRANSACCIONES              = ["IW39", "IW37N"]#, "IW29", "CR05", "IP18", "IP16", "IA09","IA10" ]  # IW23,"IA10","IA17"
# IW39
IW39_FECHA_INICIO          = "01.01.2023"
IW39_FECHA_FIN             = "08.04.2025"
# IW37N
IW37N_FECHA_INICIO         = "01.01.2023"
IW37N_FECHA_FIN            = "08.04.2025"

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
                print("Ventana 'Seguridad SAP GUI' gestionada.")
            except Exception as e:
                print(f"Error al manejar la ventana: {e}")
            time.sleep(0.3) 
        time.sleep(0.3)

# Ejecutar transacción con filtros
def transaccion(codigo): 
    
    session.findById("wnd[0]/tbar[0]/okcd").text = f"/N{codigo}"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)
    
    if codigo == "IP18":
        # Filtro por Ubicaciones Técnicas
        # session.findById("wnd[0]/usr/ctxt[6]").text = "CR4000002" #Ubicación técnica
        session.findById("wnd[0]/usr/ctxt[38]").text = CENTRO_EMPLAZAMIENTO
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(1)
        
    elif codigo == "IA09":
        # Filtro por Centro de Emplazamiento y HDR instructivas
        session.findById("wnd[0]/usr/chk[2]").selected = True
        session.findById("wnd[0]/usr/ctxt[9]").text = CENTRO_EMPLAZAMIENTO  # Centro
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(1)
        
    elif codigo == "IA10":
        session.findById("wnd[0]/usr/btn[2]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(1)
        
    elif codigo == "IP16":
        session.findById("wnd[0]/usr/btn[0]").press()
        session.findById("wnd[0]").sendVKey(24)
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(1)
        
    elif codigo == "IW39":
        session.findById("wnd[0]/usr/chk[2]").selected = True  # Concluido
        session.findById("wnd[0]/usr/chk[3]").selected = True  # Histórico
        # Centro emplazamiento
        session.findById("wnd[0]/usr/ctxt[85]").text = CENTRO_EMPLAZAMIENTO

        session.findById("wnd[0]/usr/ctxt[15]").text = IW39_FECHA_INICIO
        session.findById("wnd[0]/usr/ctxt[16]").text = IW39_FECHA_FIN
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(1)
        
    elif codigo == "IW37N":
        session.findById("wnd[0]/usr/chk[2]").selected = True  # Cerrado
        session.findById("wnd[0]/usr/chk[3]").selected = True  # Historia
        # session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub/1/2/ctxt[4]").text = "CF400002" # Ubicación técnica
        session.findById(
            "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub/1/2/ctxt[18]").text = IW37N_FECHA_INICIO # Fecha inicio
        session.findById(
            "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub/1/2/ctxt[19]").text = IW37N_FECHA_FIN # Fecha fin
        session.findById(
            "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").select()
        session.findById(
            "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub/1/2/ctxt[0]").text = CENTRO_EMPLAZAMIENTO
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(1)
        
    elif codigo == "IW29":
        session.findById("wnd[0]/usr/ctxt[106]").text = CENTRO_EMPLAZAMIENTO
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(1)
        
    elif codigo == "CR05":
        session.findById("wnd[0]/usr/ctxt[0]").text = CENTRO_EMPLAZAMIENTO  
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(1)


def seleccion_columnas():
    session.findById("wnd[0]").sendVKey(32)
    time.sleep(1)
    session.findById(
        "wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssub/1/2/cntlCONTAINER1_LAYO/shellcont/shell").setCurrentCell(-1, "")
    session.findById(
        "wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssub/1/2/cntlCONTAINER1_LAYO/shellcont/shell").selectAll()
    session.findById(
        "wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssub/1/2/btn[1]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    # session.findById("wnd[1]").sendVKey(0)
    # print("llegó hasta aquí")


def cerrar_ventanas_secundarias():
    # Intenta cerrar todas las ventanas excepto wnd[0]
    while True:
        try:
            session.findById("wnd[1]").sendVKey(0)
            time.sleep(1)
        except:
            break


def exportar_excel(path, nombre_archivo):

    ruta = os.path.join(path, nombre_archivo)
    print("Ruta de guardado:  ", ruta)
    
    if os.path.exists(ruta):
        os.remove(ruta)
    cerrar_ventanas_secundarias()
    print("Sí se cerró la segunda ventana. Cargando las columnas ...")
    
    session.findById("wnd[0]").sendVKey(16)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[0,0]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    # time.sleep(5)#Espera a que carguen los archivos
    try:
        excel_app = win32com.client.GetActiveObject("Excel.Application")
    except:
        excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = True
    wb = excel_app.ActiveWorkbook

    wb.SaveAs(ruta)
    wb.Close(SaveChanges=True)
    excel_app.Quit()
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except Exception as e:
        print("Se cerró la ventana")


def exportar_excel_cr05(path, nombre_archivo):
    
    ventanas_sap = gw.getWindowsWithTitle(NOMBRE_VENTANA_TRANSACCION )
    ventanas_sap[0].activate()
    ventanas_sap[0].maximize()
    pyautogui.hotkey('alt', 'l')
    time.sleep(0.5)
    pyautogui.press('g')
    time.sleep(0.5)
    pyautogui.press('f')
    
    session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[2,0]").select()
    session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[2,0]").setFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[20]").press()
    session.findById("wnd[1]/usr/ctxt[0]").text = path
    session.findById("wnd[1]/usr/ctxt[1]").text = nombre_archivo
    session.findById("wnd[1]/usr/ctxt[1]").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

def exportar_excel_ia10(path, nombre_archivo):
    session.findById("wnd[0]").sendVKey(45)
    session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[2,0]").select()
    session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[2,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    
    # ------------Falta seleccionar campos para la descarga. VER--------- 
    
    session.findById("wnd[1]/usr/ssub/1/txt").text = nombre_excel.replace("",".xlsx")
    session.findById("wnd[1]/usr/ssub/1/txt").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[20]").press()
    session.findById("wnd[1]/usr/ctxt[0]").text = path
    session.findById("wnd[1]/usr/ctxt[0]").setFocus()
    session.findById("wnd[1]/usr/ctxt[0]").caretPosition = 13
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    

if __name__=="__main__":
    
    # Inicialización de SAP
    if not os.path.exists(sap_logon_path):
            raise FileNotFoundError("Ruta de SAP inválida")
    subprocess.Popen(sap_logon_path)
    time.sleep(3)
    try:
            SapGuiAuto  = win32com.client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
    except Exception as e:
            raise RuntimeError("Error al cargar SAP GUI scripting") from e
    
    # Inicialización conexión
    connection = application.OpenConnection(NOMBRE_CONEXION_SAP, True)
    connection = application.Children(0)
    session = connection.Children(0)
    session.findById("wnd[0]").maximize

    # Registro de cuenta SAP
    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = RSYST_MANDT
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = RSYST_BNAME
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = RSYST_BCODE
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = RSYST_LANGU
    session.findById("wnd[0]").sendVKey(0)

    
    # #Extracción de datos
    for codigo in TRANSACCIONES:
        
        nombre_excel = codigo+".xlsx"
        print("codigo: ", codigo )
        if codigo == "IP16":
            df = pd.read_excel("D:/RecordsSAP/IP18.xlsx")
            planes = df["Plan mantenim."].dropna().astype(int).unique().tolist()
            pd.Series(planes).to_clipboard(index=False)
            transaccion(codigo)
            seleccion_columnas()
            exportar_excel(excel_file_path, nombre_excel)
            
        elif codigo == "CR05":
            transaccion(codigo)
            exportar_excel_cr05(excel_file_path, nombre_excel)
        
        elif codigo == "IA10":
            df = pd.read_excel("D:/RecordsSAP/IA09.xlsx")
            gruposHDR = df["Gpo.hojas ruta"].dropna().astype(int).unique().tolist()
            pd.Series(gruposHDR).to_clipboard(index=False)
            transaccion(codigo)
            exportar_excel_ia10(excel_file_path,nombre_excel)
            
        else:
            transaccion(codigo)
            seleccion_columnas()
            exportar_excel(excel_file_path, nombre_excel)
