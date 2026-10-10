import pandas as pd
import os
import pdfplumber
import sys
import threading
import itertools
import time 
import copy
from openpyxl import load_workbook
import re
pd.set_option('future.no_silent_downcasting', True)

# Variables configurables:
path_prts         = r"C:\Users\jesus\GESTIONA ACTIVOS S.A.C\TGA_AAQ - 2402_PMO MEJORA ESTRATEGIAS - 2402_PMO MEJORA ESTRATEGIAS\03.00 EJECUCION\04. Información para Consultores"
carpeta_area      = r"03. Area 3000\PRTs"
path_prts         = r"C:\Temp"
carpeta_area      = r"PRTs"
carpeta_sap       = r"SAP\Doc MAF"
carpeta_reportes  = r"SAP\Reportes"

lista_equipos     = ["40020103.31.40.01", "40020103.31.40.02", "40020103.32.10.01", "40020103.32.10.02", "40020103.32.10.08", "40020103.32.20.01", "CH4000032", "CH4000033", "CH4000034", "CH4000045", "CH4000046", "CH4000047", "CH4000048", "CH4000049", "CH4000051", "CH4000053", "CH4000054", "CH4000057", "CH4000059", "CH4000061", "CH4000063", "CH4000065", "CH4000066", "CH4000068", "CH4000083", "CH4000084", "CH4000085", "CH4000086", "CH4000087", "CH4000088", "CH4000098", "CV4000003", "CV4000004", "CV4000007", "CV4000008", "FE4000005", "FE4000007", "FE4000008", "FE4000009", "FE4000010", "HP4000014", "SR4000001", "SR4000002", "TH4000201", "WS4000002", "WS4000003", "WS4000004", "WS4000005", "WS4000006"]
lista_equipos     = ['CE4000051', 'CE4000052', 'CE4000053', 'CE4000054', 'CE4000055', 'CE4000056', 'CE4000057', 'CE4000058', 'CE4000059', 'CE4000060', 'CE4000061', 'CE4000062', 'CE4000063', 'CE4000064', 'CE4000065', 'CE4000066', 'CE4000067', 'CE4000068', 'CE4000069', 'CE4000070', 'CE4000071', 'CE4000072', 'CE4000073', 'CE4000074', 'CE4000075', 'CE4000076', 'CH4000002', 'CH4000003', 'CH4000004', 'CH4000005', 'CH4000032', 'CH4000033', 'CH4000034', 'CH4000045', 'CH4000046', 'CH4000047', 'CH4000048', 'CH4000049', 'CH4000051', 'CH4000053', 'CH4000054', 'CH4000057', 'CH4000059', 'CH4000060', 'CH4000061', 'CH4000063', 'CH4000065', 'CH4000066', 'CH4000067', 'CH4000068', 'CH4000083', 'CH4000084', 'CH4000085', 'CH4000086', 'CH4000087', 'CH4000088', 'CN4000022', 'CV4000001', 'CV4000002', 'CV4000003', 'CV4000004', 'CV4000007', 'CV4000008', 'EN4000001', 'FE4000004', 'FE4000005', 'FE4000008', 'FE4000009', 'HP4000014', 'LA4000012', 'LA4000013', 'LA4000014', 'LA4000016', 'LI4000001', 'LI4000002', 'MA4000001', 'MA4000002', 'MI4000001', 'MI4000002', 'MI4000003', 'MI4000004', 'PD4000008', 'PD4000009', 'PD4000028', 'PD4000042', 'PD4000043', 'PD4000044', 'PD4000055', 'PD4000069', 'PD4000070', 'PD4000077', 'PD4000078', 'PD4000079', 'PD4000080', 'PD4000081', 'PD4000082', 'PD4000083', 'PD4000084', 'PD4000085', 'PD4000086', 'PD4000087', 'PD4000088', 'PD4000089', 'PD4000090', 'PD4000091', 'PD4000093', 'PD4000099', 'PD4000105', 'PD4000106', 'SP4000501', 'SR4000001', 'SR4000002', 'SU4000001', 'SU4000001-SSSY', 'SU4000014', 'SU4000014-SSSY', 'SU4000020', 'SU4000020-SSSY', 'SU4000022', 'SU4000022-SSSY', 'SU4000029', 'SU4000029-SSSY', 'SU4000037', 'SU4000038', 'SU4000039', 'SU4000040', 'SU4000041', 'SU4000042', 'TH4000004', 'TH4000005', 'TH4000201', 'TH4000219', 'WS4000001', 'WS4000002', 'WS4000003', 'WS4000004', 'WS4000005', 'WS4000006']
# lista_equipos     = ['SR4000002',       'SU4000037', 'SU4000038', 'SU4000039', 'SU4000040', 'SU4000041', 'SU4000042', 'TH4000004', 'TH4000005', 'TH4000201', 'TH4000219', 'WS4000001', 'WS4000002', 'WS4000003', 'WS4000004', 'WS4000005', 'WS4000006']
# lista_equipos     = ['SU4000001','SU4000001-SSSY','SU4000014','SU4000014-SSSY','SU4000020','SU4000020-SSSY', 'SU4000022', 'SU4000022-SSSY',  'SU4000029']

class SimulatedDatabase:
    def __init__(self):
        self.prt = pd.DataFrame(columns=[
            'PRTId', 'Nombre', 'GeneradoPor', 'FechaRevision'
        ])
        self.recursos = pd.DataFrame(columns=[
            'RecursoId', 'PRTId', 'Tipo', 'Cantidad', 'HH', 'Especialidad',
            'Descripcion', 'MAT_CodigoRepuesto', 'MAT_CodigoFabricante'
        ])
        self.tareas = pd.DataFrame(columns=[
            'TareaId', 'PRTId', 'Equipo', 'Ítem mantenible', 'Paso',
            'Tarea', 'Comentario condicional', 'Límites aceptables'
        ])
        self.next_prt_id = 1
        self.next_recurso_id = 1
        self.next_tarea_id = 1

    def insert_prt(self, nombre, generado_por, fecha_revision, pdf_nombre):
        prt_id = self.next_prt_id
        self.next_prt_id += 1
        new_row = {
            'PRTId': prt_id,
            'Nombre': nombre,
            'GeneradoPor': generado_por,
            'FechaRevision': fecha_revision,
            'Nombre_PDF': pdf_nombre.split("PRT")[1]
        }
        self.prt = pd.concat(
            [self.prt, pd.DataFrame([new_row])], ignore_index=True)
        return prt_id

    def insert_recurso(self, prt_id, tipo, cantidad, hh, especialidad, descripcion, mat_codigo_repuesto, mat_codigo_fabricante):
        if prt_id not in self.prt['PRTId'].values:
            raise ValueError("El PRTId no existe")
        recurso_id = self.next_recurso_id
        self.next_recurso_id += 1
        new_row = {
            'RecursoId': recurso_id,
            'PRTId': prt_id,
            'Tipo': tipo,
            'Cantidad': cantidad,
            'HH': hh,
            'Especialidad': especialidad,
            'Descripcion': descripcion,
            'MAT_CodigoRepuesto': mat_codigo_repuesto,
            'MAT_CodigoFabricante': mat_codigo_fabricante
        }
        self.recursos = pd.concat(
            [self.recursos, pd.DataFrame([new_row])], ignore_index=True)
        return recurso_id

    def insert_tarea(self, prt_id, equipo, item_mantenible, paso, tarea, comentario_condicional, limites_aceptables):
        if prt_id not in self.prt['PRTId'].values:
            raise ValueError("El PRTId no existe")
        tarea_id = self.next_tarea_id
        self.next_tarea_id += 1
        new_row = {
            'TareaId': tarea_id,
            'PRTId': prt_id,
            'Equipo': equipo,
            'Ítem mantenible': item_mantenible,
            'Paso': paso,
            'Tarea': tarea,
            'Comentario condicional': comentario_condicional,
            'Límites aceptables': limites_aceptables
        }
        self.tareas = pd.concat(
            [self.tareas, pd.DataFrame([new_row])], ignore_index=True)
        return tarea_id

def convierte_valor(val):
    """Convierte a número si es posible, o devuelve el texto limpio;
    si no hay valor, retorna None."""

    if val is None:
        return None

    val = val.strip()
    if not val:
        return None

    try:
        return int(val)
    except ValueError:
        try:
            return float(val)
        except ValueError:
            return val


def obtener_filas_datos(tabla, tabla_prefij):
    """
    Para cada columna se crea una clave del tipo <table_prefix>_<Encabezado>
    Si no existe la celda de datos, se asigna None.
    """
    result = {}
    encabezados = tabla[1]

    # Inicializar diccionario
    for encabezado in encabezados:
        key = f"{tabla_prefij}_{encabezado.strip()}"
        result[key] = []

    if len(tabla) <= 2:
        for key in result:
            result[key].append(None)
        return result

    for tabla_data in tabla[2:]:
        for i, encabezado in enumerate(encabezados):
            key = f"{tabla_prefij}_{encabezado.strip()}"
            data = tabla_data[i]
            result[key].append(convierte_valor(data))
    return result


def pre_proc_tex(texto):
    if "\nComentarios Condicionales:" in texto:
        # if texto.endswith("\n"):
        new_table = texto.split("\nComentarios Condicionales:", 1)
        tarea = new_table[0]
        # comentarios_condicionales = new_table[1].split(":", 1)[
        #     1].replace("\n", " ")
        comentarios_condicionales = new_table[1].replace("\n", " ")
    else:
        tarea = texto
        comentarios_condicionales = ""
    return tarea, comentarios_condicionales

def insert_recurso(tipo, prt_id, output):

    lista_data = list()
    for key in output.keys():
        if key.startswith(tipo):
            lista_data.append(output[key])

    if tipo == "Personas":

        for cantidad, horas, especialidad in zip(*lista_data):
            descripcion = ""
            mat_codigo_repuesto = ""
            mat_codigo_fabricante = ""
            db.insert_recurso(prt_id, tipo, cantidad, horas, especialidad,
                              descripcion, mat_codigo_repuesto, mat_codigo_fabricante)

    elif tipo == "Materiales":

        for cantidad, descripcion, mat_codigo_repuesto, mat_codigo_fabricante in zip(*lista_data):
            horas = ""
            especialidad = ""
            db.insert_recurso(prt_id, tipo, cantidad, horas, especialidad,
                              descripcion, mat_codigo_repuesto, mat_codigo_fabricante)

    elif tipo == "Herramienta":
        for cantidad, horas, descripcion in zip(*lista_data):
            especialidad = ""
            mat_codigo_repuesto = ""
            mat_codigo_fabricante = ""
            db.insert_recurso(prt_id, tipo, cantidad, horas, especialidad,
                              descripcion, mat_codigo_repuesto, mat_codigo_fabricante)

    elif tipo == "Maquinaria":

        for cantidad, horas, descripcion in zip(*lista_data):
            especialidad = ""
            mat_codigo_repuesto = ""
            mat_codigo_fabricante = ""
            db.insert_recurso(prt_id, tipo, cantidad, horas, especialidad,
                              descripcion, mat_codigo_repuesto, mat_codigo_fabricante)

def spinner_func(stop_event):
    for c in itertools.cycle(['/', '—','\\', '|']):
        if stop_event.is_set():
            break
        sys.stdout.write(f'\rProcesando... {c}')
        sys.stdout.flush()
        time.sleep(0.1)

def autofiltro_excel(output_path):
    wb = load_workbook(output_path)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # Ajustar ancho de columnas (AutoFit)
        for column_cells in ws.columns:
            length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
            adjusted_width = length  # margen extra
            col_letter = column_cells[0].column_letter
            ws.column_dimensions[col_letter].width = adjusted_width

        # Aplicar autofiltro a toda la tabla
        ws.auto_filter.ref = ws.dimensions

    # Guardar cambios en el archivo
    wb.save(output_path)

if __name__ == "__main__":
    
    stop_event = threading.Event()
    spinner_thread = threading.Thread(target=spinner_func, args=(stop_event,))
    spinner_thread.start()

    # -------------------
    # Ejecución de código principal
    # -------------------
    try:    
        df_tareas_total = pd.DataFrame()
        for p, equipo_ in enumerate(lista_equipos):

            pdf_folder    = os.path.join(path_prts, carpeta_area, equipo_, carpeta_sap) 
            pdf_files     = [os.path.join(pdf_folder, file) for file in os.listdir(pdf_folder) if file.lower().endswith('.pdf')]
            output_path = os.path.join(os.path.join(path_prts, carpeta_area, equipo_, carpeta_reportes) , f"{equipo_}_PRT_data.xlsx")

            db = SimulatedDatabase()
            
            for pdf_path in pdf_files:
                try:
                    with pdfplumber.open(pdf_path) as pdf:
                        # considerar más de una página en caso los materiales se extiendan en más de una página
                        
                        pagina_eval = pdf.pages[0]
                        if pagina_eval.extract_text().startswith("COMOS"):
                            pagina_uno = pdf.pages[1]
                            numero_pag = 1
                            pagina_dos_adelante = pdf.pages[2:]
                        else:
                            pagina_uno = pdf.pages[0]
                            numero_pag = 0
                            pagina_dos_adelante = pdf.pages[1:]

                        tablas = pagina_uno.extract_tables()
                        if len(tablas) < 5:
                            pag_rest = 0
                            tablas_totales = []
                            encontrado = False
                            for pag in pdf.pages[numero_pag:]:
                                tablas = pag.extract_tables()   
                                for matriz in tablas:
                                    for fila in matriz:
                                        if 'Paso N°' in fila or 'Aceptable' in fila:
                                            encontrado = True
                                            break
                                    if encontrado:
                                        break
                                if encontrado:
                                    break
                                else:
                                    pag_rest += 1
                                    tablas_totales += tablas 
                            
                            agrupadas = dict()

                            for tabla in tablas_totales:
                                key = (tuple(tabla[0]), tuple(tabla[1]))  
                                if key in agrupadas:
                                    # Añade el resto de filas (excepto las dos primeras) a la lista ya existente
                                    agrupadas[key].extend(tabla[2:])
                                else:
                                    agrupadas[key] = copy.deepcopy(tabla)

                            # --- Resultado final ---
                            tablas = list(agrupadas.values())
                            numero_pag = numero_pag + pag_rest
                            pagina_dos_adelante = pdf.pages[numero_pag:]
                        # ---------- TABLA 0: Encabezado --------

                        encabezado = tablas[0]

                        descripcion = encabezado[0][1]
                        generado_por = encabezado[1][1]
                        fecha_revision = encabezado[1][2]
                        if "\n" in generado_por:
                            generado_por = generado_por.split("\n")[1]
                        else:
                            generado_por = generado_por.split(":")[1]
                        fecha_revision = fecha_revision.split(":")[-1].strip() if fecha_revision else ""

                        output = {
                            "Descripción": descripcion,
                            "Generado por": generado_por,
                            "Fecha Revisión": fecha_revision,
                        }
                        # print(pdf_path)
                        # print("------------------------------------------------------------")
                        # print(tablas)
                        # print(tablas)
                        # ---------- TABLA 1: Personas --------
                        personas_data = obtener_filas_datos(tablas[1], "Personas")
                        output.update(personas_data)

                        # ---------- TABLA 2: Materiales --------
                        materiales_data = obtener_filas_datos(tablas[2], "Materiales")
                        output.update(materiales_data)

                        # ---------- TABLA 3: Herramientas --------
                        herramientas_data = obtener_filas_datos(tablas[3], "Herramienta")
                        output.update(herramientas_data)

                        # ---------- TABLA 4: Maquinaria ---------
                        if len(tablas)>=5:
                            maquinaria_data = obtener_filas_datos(tablas[4], "Maquinaria")
                            output.update(maquinaria_data)
                        else:
                            tablas.append([['Maquinaria', None, None], ['Cantidad', 'Horas', 'Descripción'], ['[Error] revisar datos de PRT', '[Error] revisar datos de PRT', '[Error] revisar datos de PRT']])
                            maquinaria_data = obtener_filas_datos(tablas[4], "Maquinaria")
                            output.update(maquinaria_data)

                        # ------------------------------------------------------
                        output_mant = []
                        tablas_recuperadas = list()
                        tablas_totales = list()
                        # Procesa páginas 2 en adelante (índices 1 en adelante)
                        flag = False
                        for pagina in pagina_dos_adelante:

                            tables = pagina.extract_tables()
                            if flag:
                                break
                            for i, table in enumerate(tables):
                                # Deja de evaluar cuando llega a la tabla de comentarios adicionales
                                if any(fila == ['Paso N°', 'Comentarios adicionales'] for fila in table):
                                    flag = True
                                    break
                                # No considerar tablas que cuentan con la estructura [['Firma/'],['Iniciales']]. Tablas parásito
                                if table[0] == ['Firma/'] and table[1] == ['Iniciales']:
                                    continue

                                # Filtro para evitar imágenes
                                filtered_table = [
                                    fila for fila in table
                                    if not (fila and isinstance(fila[0], str) and fila[0].strip() == '' and fila.count(None) > 1)
                                ]

                                tablas_totales = tablas_totales + filtered_table

                    output_mant = []
                    equipo = None

                    for tabla in tablas_totales:
                        if tabla[0] is not None and all(cell is None for cell in tabla[1:]):
                            equipo = {'Equipo': tabla[0], 'Data': []}
                            output_mant.append(equipo)
                        else:
                            if equipo is not None:
                                equipo['Data'].append(tabla)

                    # -----------------------PÁG 1: INSERT_PRT_INFO-------------------------------------#

                    nombre_prt = output["Descripción"]
                    generadoPor = output["Generado por"]
                    fecha_revision = output["Fecha Revisión"]
                    prt_id = db.insert_prt(nombre_prt, generadoPor, fecha_revision, pdf_path)
                    
                    

                    # -----------------------PÁG 1: INSERT_RECURSOS-------------------------------------#
                    for tipo in ["Personas", "Materiales", "Herramienta", "Maquinaria"]:
                        insert_recurso(tipo, prt_id, output)

                    # -----------------------PÁG 2-: INSERT_TAREAS-------------------------------------#
                    for equipo_mant in output_mant:
                        equipo = equipo_mant["Equipo"]
                        data = equipo_mant["Data"]

                        for pasos_tareas in data:
                            if pasos_tareas[0] == "Paso N°":
                                item_mantenible = pasos_tareas[1]
                                continue
                            if pasos_tareas[0] == '':
                                last_index = db.tareas.index[-1]
                                columnas = ['Tarea', 'Comentario condicional',
                                            'Límites aceptables']
                                if '\nComentarios Condicionales:' in pasos_tareas[1]:
                                    datos_div = pasos_tareas[1].split(
                                        "\nComentarios Condicionales:", 1)
                                    pasos_tareas[0] = datos_div[0]
                                    pasos_tareas[1] = datos_div[1]

                                elif 'Comentarios Condicionales:' in pasos_tareas[1]:
                                    pasos_tareas[1] = pasos_tareas[1].split(":", 1)[1]

                                for col, nuevo_valor in zip(columnas, pasos_tareas[0:3]):
                                    valor_actual = db.tareas.loc[last_index, col]
                                    if pd.isna(valor_actual):
                                        valor_actual = ""
                                    db.tareas.loc[last_index, col] = valor_actual + nuevo_valor
                                continue

                            paso = pasos_tareas[0]
                            texto = pasos_tareas[1]

                            tarea, comentario_condicional = pre_proc_tex(texto) if texto else ('','') 
                            # print(pdf_path)
                            if len(pasos_tareas)>4:
                                limites_aceptables = pasos_tareas[2].replace("\n", " ") if pasos_tareas[2] else ''
                                db.insert_tarea(prt_id, equipo, item_mantenible, paso,
                                            tarea, comentario_condicional, limites_aceptables)
                            else: 
                                print(pdf_path)
                except:
                    print(pdf_path)

            df_tareas   =  db.tareas
            df_info     =  db.prt
            df_tareas = pd.merge(df_tareas, df_info[['PRTId','Nombre']], on = 'PRTId', how = 'left')
            df_tareas['Nombre'] = df_tareas['Nombre'].str.replace(r'\n', ' ', regex=True).str.replace('\(_\)$', '', regex=True).str.strip()
            df_tareas["Frecuencia"]           = df_tareas['Nombre'].str.extract(r'(\d)')
            df_tareas["Unidad"]               = df_tareas['Nombre'].str.extract(r'\d([A-Za-z])[^/]*\/')
            df_tareas["Constraint"]           = df_tareas['Nombre'].str.extract(r'/\s*([A-Za-z])')
            df_tareas["Tipo de Tarea Equiv."] = df_tareas['Nombre'].str.split().str[1]
            df_tareas["Especialidad"]         = df_tareas['Nombre'].str.split().str[2]

            df_tareas["Constraint"]  = df_tareas["Constraint"].replace({'D': 'Detenido no energizado - bloqueado', 'O': 'Operativo'})
            df_tareas["Unidad"]      = df_tareas['Unidad'].replace({'D': 'Calendar Día',
                                                                'W': 'Calendar Semana',
                                                                'M': 'Calendar Mes',
                                                                'Y': 'Calendar Año'})

            df_tareas["Unidad_dias"] = df_tareas['Unidad'].replace({'Calendar Día': 1,
                                                                'Calendar Semana': 7,
                                                                'Calendar Mes': 30,
                                                                'Calendar Año': 360}).infer_objects(copy=False)
            df_tareas["frecuencia_dias"] = df_tareas['Unidad_dias'] * pd.to_numeric(df_tareas["Frecuencia"], errors="coerce")
            df_tareas = df_tareas[df_tareas['Constraint']=='Detenido no energizado - bloqueado']
            df_tareas = ( df_tareas
                            .sort_values(by=['Tipo de Tarea Equiv.', 'Especialidad', 'Tarea', 'frecuencia_dias'])
                            .groupby(['Tipo de Tarea Equiv.', 'Especialidad', 'Tarea'], group_keys=False)
                            .first()
                            .reset_index()
                        )
            df_tareas = df_tareas[['Nombre', 'Equipo','Ítem mantenible', 'Frecuencia', 'Unidad', 'Constraint', 'Tarea','Límites aceptables', 'Comentario condicional', 'Tipo de Tarea Equiv.', 'Especialidad', 'frecuencia_dias']].sort_values(by=['Tipo de Tarea Equiv.', 'Especialidad', 'frecuencia_dias'],
                ascending=False)
            df_tareas_total = pd.concat([df_tareas_total, df_tareas], ignore_index=True)

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                db.prt.to_excel(writer, sheet_name="Info", index= False)
                db.recursos.to_excel(writer, sheet_name="Recursos", index= False)
                db.tareas.to_excel(writer, sheet_name="Tareas", index= False)
                df_tareas.to_excel(writer, sheet_name="Análisis", index= False)

                
            print(f"\n Equipo {equipo_} completado")
            autofiltro_excel(output_path)

        df_tareas_total['Ubicación Técnica'] = df_tareas_total['Equipo'].apply( lambda x: re.search(r'\((.*?)\)', x).group(1) if pd.notnull(x) and re.search(r'\((.*?)\)', x) else "No se encontró UT")
        df_tareas_total = df_tareas_total[['Ubicación Técnica', 'Nombre', 'Equipo','Ítem mantenible', 'Frecuencia', 'Unidad', 'Constraint', 'Tipo de Tarea Equiv.', 'Especialidad', 'Tarea','Límites aceptables', 'Comentario condicional']]

        output_total_path = os.path.join(path_prts, carpeta_area, "Tareas_Totales.xlsx", )
        with pd.ExcelWriter(output_total_path, engine='openpyxl') as writer:
            df_tareas_total.to_excel(writer, sheet_name="Tareas Total", index=False)
        autofiltro_excel(output_total_path)
    
    finally:
        stop_event.set()
        spinner_thread.join()
        print("\nProceso completado.")