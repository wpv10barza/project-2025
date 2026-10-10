import pandas as pd


def ordenar_df_por_columna(df1, df2, columna_orden):
    '''Función que ordena un Dataframe de acuerdo a una columna de otro
       Dataframe '''

    # Obtener la lista de orden desde df1
    order_list = df1[columna_orden].tolist()

    # Aplicar el orden a df2
    df2[columna_orden] = pd.Categorical(df2[columna_orden],
                                        categories=order_list,
                                        ordered=True)

    return df2.sort_values(columna_orden).reset_index(drop=True)


def comparar_tablas_celda_celdad(df1, df2, columna_orden):
    '''Función que compara celda a celda entre dos tablas (dataframes).

       Primero, ordena la segunda Tabla (df2) de acuerdo a una columna
       de la primera Tabla (df1)
       Luego evalúa la comparación entre celda y celda. El resultado es
       un dataframe de valores booleanos, donde False significa valores
       idénticos, mientras que True significan valores diferentes.
      '''
    if len(df1) != len(df2):
        raise ValueError(
            "La cantidad de filas entre ambos Dataframes son diferentes")

    df2_ = ordenar_df_por_columna(df1, df2, columna_orden)

    # Asegurar que ambos DataFrames tengan las mismas columnas
    columnas_comunes = df1.columns.intersection(df2_.columns)
    df1 = df1[columnas_comunes].reset_index(drop=True)
    df2_ = df2_[columnas_comunes].reset_index(drop=True)

    # Comparar valores de cada celda
    diferencias = (df1 != df2_) & ~(df1.isna() & df2_.isna())

    # Contar la cantidad de diferencias por columna
    diferencias_count = diferencias.sum()

    # Crear una fila con los totales de diferencias
    total_diferencias_fila = pd.DataFrame(
        [diferencias_count], columns=df1.columns)
    total_diferencias_fila.index = ["Total_de_diferencias"]

    # Agregar la fila con los totales al final del DataFrame original
    df_con_diferencias = pd.concat(
        [df1, total_diferencias_fila], ignore_index=False)

    # Agregar la fila con los totales al final del Dataframe de comparación
    df_valores_booleanos = pd.concat(
        [diferencias, total_diferencias_fila], ignore_index=False)

    # Mostrar solo las filas donde hay diferencias
    if diferencias_count.sum() == 0:
        print("No hay diferencias entre los DataFrames")
        return df_con_diferencias
    else:
        return df_valores_booleanos
        # return df2.head()


def celdas_diferentes_columna(df1, df2, columna_evaluada):
    '''Función que compara celda a celda entre dos DataFrames con las mismas
       columnas.
        '''
    if columna_evaluada not in set(df1.columns) or columna_evaluada not in set(df2.columna_evaluada):
        raise ValueError(
            "La columna evaluada no existe en al menos un Dataframe")

    if len(df1) != len(df2):
        raise ValueError("Los DataFrames deben tener el mismo número de filas")

    mascara = df1[columna_evaluada] != df2[columna_evaluada]

    return pd.Dataframe({
        'df1': df1.loc[mascara, columna_evaluada],
        'df2': df2.loc[mascara, columna_evaluada]})


def filas_unicas(df, columna):
    '''Función que devuelve las filas únicas de acuerdo a los valores únicos
       de una columna específica*

       *Nota: Solo se toman los valores de la primera fila única que coincide,
       las demás filas las elimina'''

    columns_repetidas = df[columna].value_counts()

    df_unicos = df.drop_duplicates(columna)

    return pd.merge(df_unicos, columns_repetidas, on=columna, how='left')


def agrupar_por_valor_unico(df, columna):
    """
    Agrupa las filas del DataFrame por los valores únicos de columna
    y conserva la información de las demás columnas agregándola en listas.
    """
    return df.groupby(columna).agg(lambda x: list(x)).reset_index()


def filas_duplicadas(df, columna):
    '''Función que entrega las filas que tienen elementos duplicados
       en la columna consultada '''

    conteo_valores = df[columna].value_counts()

    # Genera una lista de los elementos que aparecen más de una vez
    valores_duplicados = conteo_valores[conteo_valores > 1].index

    return df[df[columna].isin(valores_duplicados)]


def cantidad_filas(df1, df2):
    '''Validación de cantidades de filas'''

    if (len(df1) == len(df2)):
        print(f"Son únicos. Val únicos: {len(df1)}   Total filas:{len(df2)}")

    else:
        print(
            f"No son únicos. Val únicos: {len(df1)}   Total filas:{len(df2)}")


def grupo_celdas_unicas(df, columnas):
    '''Verifica que el conjunto de dos elementos de dos columnas
     no se repitan '''

    contador = df[columnas].value_counts()
    duplicados = contador[contador != 1]

    if duplicados.empty:
        print("El grupo de elementos son únicos")
        return None

    else:
        return duplicados
