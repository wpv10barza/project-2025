from dotenv import load_dotenv
from ventana import pantalla_final
import threading
import os, time, subprocess
import pandas as pd
import win32com.client
import pygetwindow as gw
import pyautogui
import datetime
import numpy as np
import traceback
#Carga variables de .env
load_dotenv()

# Constantes
sap_logon_path      = r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"
PRT_path_base       = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.01 EJECUCION TI\05. RPA EXEC\01. MAF C-U"
NOMBRE_CONEXION_SAP = "Copper SAP ECC Producción(PS8)"
RSYST_MANDT         = "300"
RSYST_BNAME         = os.getenv("SAP_USER")
RSYST_BCODE         = os.getenv("SAP_PASS")
RSYST_LANGU         = "ES"
TRANSACCION         = "/NCV01N"
TRANSACCION_CV04N   = "/NCV04N"
PMA_FIELD           = "PMA" # Clase de documento
DOC_PARCIAL         = "000" # Documento parcial
CATEGORIA_ARCHIVO   = "CF_IMP_AS" # Categoría de archivo DMS_C1_ST para versión SAP/4HANA (Versión 770)


# Variables de configuración
PRT_CarpetaAProcesar     = "REQ-015 Solicitud_04_20250820 (CREACION)"
PRT_ArchivoExcelNombre   = "DATOS.xlsx"
PRT_ArchivoHojaNombre    = "MAF-CU"
PRT_CarpetaDeArchivos    = "Files"
COLUMNA_ACCION           = "ACCIÓN"
COLUMNA_DESCRIPCION      = "DESCRIPCIÓN DE LA OPERACIÓN"	
COLUMNA_DOCUMENTO_MAF    = "DOCUMENTO MAF"
COLUMNA_VERSION          = "VERSIÓN"
COLUMNA_ESTADO           = "ESTADO"
MAXIMO_ERRORES           = 150
NOMBRE_VENTANA_EMERGENTE = "Seguridad SAP GUI"
NOMBRE_VENTANA_COPIA     = "Copiar originales no archivados"
DESCARGA_PATH_REQUERI    =  os.path.join(PRT_path_base, PRT_CarpetaAProcesar)
NOMBRE_CVN40N_DESCARGA   = "CV04N.xlsx"

# Actualizando variables
PRT_path        = os.path.join(PRT_path_base, PRT_CarpetaAProcesar, PRT_CarpetaDeArchivos)
path_prt_subir  = os.path.join(PRT_path_base, PRT_CarpetaAProcesar, PRT_ArchivoExcelNombre)

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
        
def manejar_ventana_copia(stop_event):
    while not stop_event.is_set():
        ventanas = gw.getWindowsWithTitle(NOMBRE_VENTANA_COPIA) 
        if ventanas:
            ventana = ventanas[0]
            try:
                ventana.activate()
                time.sleep(0.1)
                pyautogui.press('esc')
                #print("Ventana 'Seguridad SAP GUI' gestionada.")
            except Exception as e:
                print(f"Error al manejar la ventana: {e}")
            time.sleep(0.3) 
        time.sleep(0.3)

def descarga_transaccion_cv04n(df_documentosMAF):
    valores = len(df_documentosMAF) * 3
    #Copiar código de documentos MAF en el portapapeles
    codigos_MAF = df_documentosMAF.dropna().astype(int).unique().tolist()
    pd.Series(codigos_MAF).to_clipboard(index=False)

    #Inicio de transacción
    session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION_CV04N #Transacción
    session.findById("wnd[0]").sendVKey(0)
    
    # Aumentar número de coincidencias
    session.findById("wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssub/1/2/sub/1/2/1/txt[5]").text = valores    

    # Filtro por código documentos MAF
    session.findById("wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssub/1/2/sub/1/2/1/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[32]").press()
    

    # Selección de columnas
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssub/1/2/cntlCONTAINER1_LAYO/shellcont/shell").setCurrentCell(-1,"")
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssub/1/2/cntlCONTAINER1_LAYO/shellcont/shell").selectAll()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssub/1/2/btn[1]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(2)
    # session.findById("wnd[1]/tbar[0]/btn[0]").press()
    
    # Descarga
    session.findById("wnd[0]").sendVKey(16)
    session.findById("wnd[1]/tbar[0]/btn[20]").press()
    session.findById("wnd[1]/usr/ctxt[0]").text = DESCARGA_PATH_REQUERI
    session.findById("wnd[1]/usr/ctxt[1]").text = NOMBRE_CVN40N_DESCARGA
    session.findById("wnd[1]/usr/ctxt[1]").caretPosition = 5
    session.findById("wnd[1]/tbar[0]/btn[0]").press()


# Liberar documento
def transaccion_cv02n(version, documentoMAF):
    
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NCV02N" #Transacción
    session.findById("wnd[0]").sendVKey(0)
    
    session.findById("wnd[0]/usr/ctxt[0]").text = documentoMAF # Código de PRT
    session.findById("wnd[0]/usr/ctxt[1]").text = PMA_FIELD 
    session.findById("wnd[0]/usr/ctxt[2]").text = DOC_PARCIAL
    session.findById("wnd[0]/usr/ctxt[3]").text = version # Versión
    session.findById("wnd[0]").sendVKey(0)
    
    # Estado de documento: Liberado
    session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/ctxt[0]").setFocus()
    session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/ctxt[0]").caretPosition = 2
    session.findById("wnd[0]").sendVKey (4)
    session.findById("wnd[1]/usr/lbl[6,3]").setFocus()
    session.findById("wnd[1]/usr/lbl[6,3]").caretPosition = 5
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    
    # Guardar
    session.findById("wnd[0]").sendVKey(11)
    

# Creación de PRT y nueva versión. CV01N
# Precondición: documento debe estar previamente liberado
def transaccion_cv01n(tipo, descripcion_prt, nombre_archivo, versionVigente, estado, documentoMAF=""):
    session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION #Transacción
    session.findById("wnd[0]").sendVKey(0)
    
    
    
    if tipo == "CREAR": # *************REVISAR
        session.findById("wnd[0]/usr/ctxt[0]").text = "" # Código de PRT
        session.findById("wnd[0]/usr/ctxt[1]").text = PMA_FIELD 
        session.findById("wnd[0]/usr/ctxt[2]").text = DOC_PARCIAL
        session.findById("wnd[0]/usr/ctxt[3]").text = "" # Versión
        session.findById("wnd[0]").sendVKey(0)
             
        subir_prt(descripcion_prt,nombre_archivo)  
      
    elif tipo == "MODIFICAR":
        if versionVigente is None: # Identifica que exista el documentoMAF
            print(f"No existe ninguna versión del documento MAF. Revisar el código: {documentoMAF}") 
            return
        # if estado == "en creación":
        #     print(f"El documento MAF {documentoMAF} no está liberado")
        #     return
        session.findById("wnd[0]/usr/ctxt[0]").text = documentoMAF # Código de PRT
        session.findById("wnd[0]/usr/ctxt[1]").text = PMA_FIELD 
        session.findById("wnd[0]/usr/ctxt[2]").text = DOC_PARCIAL
        session.findById("wnd[0]/usr/ctxt[3]").text = "" # Versión
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        session.findById("wnd[1]/usr/btn[0]").press()
        # if versionVigente == 0:
        #         # Ventana emergente
        #         # session.findById("wnd[1]/usr/btn[0]").press()
        #         session.findById("wnd[1]/tbar[0]/btn[12]").press()
        #         subir_prt(descripcion_prt, nombre_archivo)
        # else:
                # Eliminar archivo
        time.sleep(3)
        try:
            session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell[1]").selectNode("          1")
            session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/btn[8]").press()
            session.findById("wnd[1]/usr/btn[0]").press()
            session.findById("wnd[0]/usr/ctxt[2]").setFocus()
            session.findById("wnd[0]/usr/ctxt[2]").caretPosition = 0
        except Exception as e:
            print("Acción de eliminar no ejecutada")
        subir_prt(descripcion_prt, nombre_archivo)


def subir_prt(descripcion_prt, nombre_archivo):
         # Subir documento MAF
        # print(descripcion_prt)
        session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/txt[0]").text = descripcion_prt
        session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/btn[4]").press()
        session.findById("wnd[1]/usr/ctxt[0]").text = "PDF" # Tipo de documento
        session.findById("wnd[1]/usr/txt[1]").text = descripcion_prt # Descripción de PRT
        session.findById("wnd[1]/usr/ctxt[2]").setFocus()
        session.findById("wnd[1]/usr/ctxt[2]").caretPosition = 0
        session.findById("wnd[1]").sendVKey(4)
        session.findById("wnd[2]/usr/ctxt[0]").text = PRT_path # Dirección de documento
        session.findById("wnd[2]/usr/ctxt[1]").text = nombre_archivo # Nombre de Documento
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        #print("Duerme")
        time.sleep(3)

        # Categoría de Archivo: DMS_C1_ST
        try:
            session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell[1]").selectNode("          1")
            session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell[1]").nodeContextMenu("          1")
        except Exception as e:
            print("No existe nodo 1. Probando nodo 2 ...")
            try:
                session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell[1]").selectNode("          2")
                session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell[1]").nodeContextMenu("          2")                
            except Exception as e:
                print("No existe nodo 2 o no está disponible", e)

        session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssub/2/3/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell[1]").selectContextMenuItem(CATEGORIA_ARCHIVO)
        session.findById("wnd[1]/usr/tblSAPLCV117TBL_SCAT").getAbsoluteRow(2).selected = True
        session.findById("wnd[1]/usr/tblSAPLCV117TBL_SCAT/txt[0,2]").setFocus()
        session.findById("wnd[1]/usr/tblSAPLCV117TBL_SCAT/txt[0,2]").caretPosition = 0
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # Ctrl + S GUARDAR
        session.findById("wnd[0]").sendVKey(11)


# MAIN
if __name__ == '__main__':
    print("Hora de inicio: ",datetime.datetime.now())
    CONTADOR_ERROR = 0

    stop_event = threading.Event()
    handler_thread = threading.Thread(target=manejar_ventana_seguridad, args=(stop_event,))
    handler_thread_2 = threading.Thread(target=manejar_ventana_copia, args=(stop_event,))
    
    handler_thread.daemon = True
    handler_thread_2.daemon = True 
    handler_thread.start()
    handler_thread_2.start()
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

    # Lectura de Excel con requerimientos
    if not os.path.exists(path_prt_subir):
          raise FileNotFoundError(f"Ruta de Excel no encontrada: {path_prt_subir}")
    
    df_prts  =  pd.read_excel(path_prt_subir, sheet_name = PRT_ArchivoHojaNombre)
    ruta_prt = os.path.join(DESCARGA_PATH_REQUERI, NOMBRE_CVN40N_DESCARGA)
     
    if os.path.isfile(ruta_prt):
        os.remove(ruta_prt)

    DF_ANALIZAR = df_prts
    
      
    for index in range(len(DF_ANALIZAR)):
        
        NUEVA_VERSION = -1
        MENSAJE_ERROR = ""
        
        try:
            OPERACION       = DF_ANALIZAR.loc[index, COLUMNA_ACCION] 
            DESCRIPCION     = DF_ANALIZAR.loc[index, COLUMNA_DESCRIPCION] 
            DOCUMENTO_MAF   = DF_ANALIZAR.loc[index, COLUMNA_DOCUMENTO_MAF]
            VERSION         = DF_ANALIZAR.loc[index, COLUMNA_VERSION]
            ESTADO          = DF_ANALIZAR.loc[index, COLUMNA_ESTADO]
        except Exception as e:
            print(f"[ERROR] Problemas con la fila {index}: {e}")
            break            
        
        
        descrip_base = DESCRIPCION.replace("/","").replace("  "," ")
        
        variaciones_pdf = [
            descrip_base + ".pdf",
            descrip_base.replace(".","") + ".pdf",
            descrip_base.replace("-","") + ".pdf",
            descrip_base.replace(".","").replace("-","") + ".pdf"
            
        ]
            
        for nombre in variaciones_pdf:
            # print(nombre)
            PATH_PRT = os.path.join(PRT_path, nombre)
            if os.path.isfile( PATH_PRT):
                NOMBRE_PRT_PDF    = nombre
                break

        else:
            # print(f"[INFO] No se ha encontrado el nombre del archivo {DESCRIPCION}") #
            DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] = "No se encontró el nombre de archivo" 
            CONTADOR_ERROR += 1
            if CONTADOR_ERROR >= MAXIMO_ERRORES:
                print(f"[ERROR] No se encontraron {MAXIMO_ERRORES} documentos MAF. Revisar nombre de documentos MAF")
                break
            continue
        
        if ESTADO == "en creación":
            try:
                transaccion_cv02n( 
                    VERSION,
                    DOCUMENTO_MAF)
            except Exception as e:
                print(f"[ERROR] Ocurrió un error en la transacción CV02N para {DOCUMENTO_MAF}: {e}")
                MENSAJE_ERROR = "Error en CV02N"
                
        if MENSAJE_ERROR == "":
            try:
                transaccion_cv01n(
                        OPERACION, 
                        DESCRIPCION,
                        NOMBRE_PRT_PDF,
                        VERSION,
                        ESTADO,
                        DOCUMENTO_MAF)
                # print(type(DOCUMENTO_MAF), DOCUMENTO_MAF)
                NUEVA_VERSION = VERSION + 1
                # if np.isnan(DOCUMENTO_MAF=='<NA>'):
                if DOCUMENTO_MAF=='<NA>' or np.isnan(DOCUMENTO_MAF) :   
                    DOCUMENTO_MAF = session.findById("wnd[0]/usr/ctxt[0]").text
                    NUEVA_VERSION = "00"
                    DF_ANALIZAR.loc[index,COLUMNA_DOCUMENTO_MAF] = DOCUMENTO_MAF  
                    
                transaccion_cv02n( 
                        NUEVA_VERSION,
                        DOCUMENTO_MAF)                
                print(f"Proceso exitoso para el documento {DOCUMENTO_MAF}")
            except Exception as e:
                print(f"[ERROR] Ocurrió un error en la transacción CV01N para {DOCUMENTO_MAF}: {e}")
                traceback.print_exc()
                MENSAJE_ERROR = "Error en CV01N"  

        DF_ANALIZAR.loc[index, "ESTADO_EJECUCION_RPA"] = "OK" if MENSAJE_ERROR == "" else MENSAJE_ERROR 
        DF_ANALIZAR.loc[index, "NUEVA_VERSION"] = NUEVA_VERSION   
                    
    DF_ANALIZAR[COLUMNA_DOCUMENTO_MAF] = DF_ANALIZAR[COLUMNA_DOCUMENTO_MAF].astype(str)
    DF_ANALIZAR = DF_ANALIZAR[[ COLUMNA_DESCRIPCION,
                                COLUMNA_DOCUMENTO_MAF,
                                "ESTADO_EJECUCION_RPA",
                                "NUEVA_VERSION"
                               ]]
    DF_ANALIZAR.to_excel(os.path.join(DESCARGA_PATH_REQUERI,"Resultados.xlsx"), index = False)  
    print("Hora de fin: ",datetime.datetime.now())
    pantalla_final()            