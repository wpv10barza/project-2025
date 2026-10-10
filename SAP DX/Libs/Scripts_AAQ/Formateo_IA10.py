import numpy as np
import pandas as pd
def operacion(df_HROP):
    # Formatear la IA10_OP
    resultados = []

    for i in range(len(df_HROP)):
        fila_actual = df_HROP.iloc[i]

        if i+3 < len(df_HROP):
            fila_primera = df_HROP.iloc[i + 2]
            fila_segunda = df_HROP.iloc[i + 3]

            grupo              = fila_actual['Unnamed: 1'] if pd.notna(fila_actual['Unnamed: 1']) else  np.nan
            contador           = fila_actual['Unnamed: 3'] if pd.notna(fila_actual['Unnamed: 3']) else  np.nan

            operacion          = fila_primera['Unnamed: 4'] if pd.notna(fila_primera['Unnamed: 2']) and pd.notna(fila_primera['Unnamed: 4']) else  np.nan
            descripcion        = fila_primera['Unnamed: 5'] if pd.notna(fila_primera['Unnamed: 4']) and pd.notna(fila_primera['Unnamed: 5']) else  np.nan
            duracion           = fila_primera['Unnamed: 6'] if pd.notna(fila_primera['Unnamed: 4']) and pd.notna(fila_primera['Unnamed: 6']) else  np.nan
            estado_instalacion = fila_primera['Unnamed: 7'] if pd.notna(fila_primera['Unnamed: 4']) and pd.notna(fila_primera['Unnamed: 7']) else  np.nan
            puesto_trabajo     = fila_primera['Unnamed: 8'] if pd.notna(fila_primera['Unnamed: 4']) and pd.notna(fila_primera['Unnamed: 8']) else  np.nan
            trabajo            = fila_primera['Unnamed: 9'] if pd.notna(fila_primera['Unnamed: 4']) and pd.notna(fila_primera['Unnamed: 9']) else  np.nan
            cantidad           = fila_primera['Unnamed: 10'] if pd.notna(fila_primera['Unnamed: 4']) and pd.notna(fila_primera['Unnamed: 10']) else  np.nan
            # unidad_minima      = fila_primera['Unnamed: 11'] if pd.notna(fila_primera['Unnamed: 4']) and pd.notna(fila_primera['Unnamed: 11']) else  np.nan
            # unidad_real        = fila_primera['Unnamed: 12'] if pd.notna(fila_primera['Unnamed: 4']) and pd.notna(fila_primera['Unnamed: 12']) else  np.nan

        else:
            grupo              = fila_actual['Unnamed: 1'] if pd.notna(fila_actual['Unnamed: 1']) else  np.nan
            contador           = fila_actual['Unnamed: 3'] if pd.notna(fila_actual['Unnamed: 3']) else  np.nan
            operacion          =  np.nan
            descripcion        =  np.nan
            duracion           =  np.nan
            estado_instalacion =  np.nan
            trabajo            =  np.nan
            cantidad           =  np.nan
            puesto_trabajo     =  np.nan
            
        if pd.notna(grupo) and pd.notna(contador) or pd.notna(operacion):
            resultados.append({
                'Grupo hojas ruta': grupo,
                'Cont.grupo HRuta': contador, 
                'Operación': operacion,
                'Descripción_OP': descripcion,
                'Duración': duracion,
                'Estado instalación_OP': estado_instalacion,
                'Trabajo': trabajo,
                'Pto.tbjo.operación':puesto_trabajo,
                # 'Unidad mínima': unidad_minima,
                'Cantidad': cantidad,
            })
    IA10_OP_format = pd.DataFrame(resultados)
    IA10_OP_format['Grupo hojas ruta'] = IA10_OP_format['Grupo hojas ruta'].ffill()
    IA10_OP_format['Cont.grupo HRuta'] = IA10_OP_format['Cont.grupo HRuta'].ffill()
    return IA10_OP_format