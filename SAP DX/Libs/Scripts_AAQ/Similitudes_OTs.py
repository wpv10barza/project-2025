import pandas as pd
import re
import nltk
import os
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from sklearn.cluster import AgglomerativeClustering, OPTICS
from sentence_transformers import SentenceTransformer

nltk.download('stopwords')
nltk.download('punkt')
nltk.download('punkt_tab')
# Cargar stopwords en español
stop_words = set(stopwords.words('spanish'))


IW39         = r"C:\Users\jesus\Desktop\Nueva carpeta\IW39.xlsx"
path_output  = r"C:\Users\jesus\Desktop\Nueva carpeta"
output_excel = os.path.join(path_output,'sim_subagr.xlsx')
df_OT   = pd.read_excel(IW39, sheet_name='iw39_')

lista_equipos    = [
    "2410-UF-3001"
]


def preprocess_text(text):
    text = text.lower()
    text = re.sub(r'[^a-záéíóúüñ0-9 ]', ' ', text)
    tokens = nltk.word_tokenize(text)
    tokens = [word for word in tokens if word not in stop_words]
    return ' '.join(tokens)

def formatear_hoja_excel(
    writer,
    df,
    nombre_hoja,
    nombre_tabla,
    agrupar_columnas,
    estilo_tabla='Table Style Medium 9',
    aplicar_agrupado=True,
    columnas_colorear=None):

    df.to_excel(writer, sheet_name=nombre_hoja, index=False)
    workbook = writer.book
    worksheet = writer.sheets[nombre_hoja]

    formato_cabecera = workbook.add_format({
        'bold': True,
        'bg_color': '#18814a',
        'border': 1,
        'align': 'center'
    })
    # Ajustar ancho de columnas
    for i, col in enumerate(df.columns):
        max_len = max(df[col].astype(str).map(len).max(), len(col))
        worksheet.set_column(i, i, max_len + 2)
        # Colorear cabecera si está en la lista
        if columnas_colorear and col in columnas_colorear:
            worksheet.write(0, i, col, columnas_colorear)

    # Agrupamiento
    if aplicar_agrupado:
        worksheet.outline_settings(True, False, False, False)
        worksheet.set_row(0, None, None, {'level': 1})
        for rango in agrupar_columnas:
            worksheet.set_column(rango, None, None, {'level': 1})

    # Tabla
    max_row, max_col = df.shape
    worksheet.add_table(0, 0, max_row, max_col - 1, {
        'columns': [{'header': col} for col in df.columns],
        'name': nombre_tabla,
        'style': estilo_tabla
    })



# df_OT["Status_Sis"] = df_OT["Status sistema"].str.split(" ").str[0]
# df_OT['Texto breve'] = df_OT['Texto breve'].fillna('')
# df_OT['texto_procesado'] = df_OT['Texto breve'].apply(preprocess_text)
df_OT['texto_procesado'] = df_OT['Description'].apply(preprocess_text)
# df_OT = df_OT[['Pos.mantenim.','Gpo.hojas ruta','Cont.gpo.HRuta','Orden','Campo clasif.','Ubicac.técnica','Equipo','Texto breve','Status_Sis','texto_procesado']]#'Plan mant.prev.',

df_OT = df_OT[['Maintenance Item','Group','Group Counter','Order','Sort Field','Functional Location','Equipment','Description','texto_procesado']]#'Plan mant.prev.',


df_BERT = df_OT.copy()
df_BERT['Similitud'] = ""
df_BERT['Subgrupo'] = ""
df_BERT['Cluster_OPTICS'] = ""



model = SentenceTransformer('multi-qa-MiniLM-L6-cos-v1')

for equipo in lista_equipos:
    # df_ml = df_BERT[df_BERT['Sort Field'] == equipo].copy()
    df_ml = df_BERT[df_BERT['Sort Field'] == '2410-UF-3001'].copy()

    
    if len(df_ml) < 2:
        continue

    textos = df_ml['texto_procesado'].to_list()
    embeddings = model.encode(textos)

    # Agrupación principal (Agglomerative)
    clustering_main = AgglomerativeClustering(
        n_clusters=None,
        distance_threshold=0.5,
        metric='cosine',
        linkage='average'
    )
    etiquetas_main = clustering_main.fit_predict(embeddings)
    df_BERT.loc[df_ml.index, 'Similitud'] = etiquetas_main

    # Subagrupación dentro de cada grupo principal
    for grupo in set(etiquetas_main):
        indices_grupo = df_ml.index[etiquetas_main == grupo]
        if len(indices_grupo) < 2:
            continue
        embeddings_sub = model.encode(df_ml.loc[indices_grupo, 'texto_procesado'].to_list())

        clustering_sub = AgglomerativeClustering(
            n_clusters=None,
            distance_threshold=0.3,
            metric='cosine',
            linkage='average'
        )
        etiquetas_sub = clustering_sub.fit_predict(embeddings_sub)
        df_BERT.loc[indices_grupo, 'Subgrupo'] = etiquetas_sub

    # Agrupación general con OPTICS
    etiquetas_optics = OPTICS(min_samples=2, metric='cosine').fit_predict(embeddings)
    df_BERT.loc[df_ml.index, 'Cluster_OPTICS'] = etiquetas_optics


df_BERT.rename(columns={'Similitud': 'Grupo_Similitud 1',
                        'Subgrupo': 'Grupo_Similitud 2',
                        'Cluster_OPTICS': 'Grupo_Similitud 3' }, inplace=True)


df_BERT.to_excel(output_excel, index=False)

























































