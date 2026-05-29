import pandas as pd

import pandas as pd

def de_Moodle_a_actas(excelMoodle: str, excelActas: str, nombre_columna_con_las_calificaciones: str, ColumnaPresentados: str = 'Total Presentado (Real)'):
    """Función que cumplimenta una copia de las actas descargadas del portal de
    gestión académica con las notas calculadas en Moodle y almacenadas en un
    archivo Excel descargado desde Moodle.

    Si el archivo de Moodle contiene una columna con el nombre indicado en
    ColumnaPresentados, solo se copian las calificaciones de los alumnos que
    tienen un 1 en dicha columna.

    excelMoodle: nombre (incluido el path) del archivo descargado desde Moodle
    excelActas:  nombre (incluido el path) del archivo de actas descargado desde GEA
    nombre_columna_con_las_calificaciones: nombre de la columna del archivo excelMoodle donde están las notas finales
    ColumnaPresentados: nombre de la columna que indica si el alumno está presentado (por defecto 'Total Presentado (Real)').
                        Solo se copian las calificaciones de los alumnos con valor 1 en dicha columna.
                        Si la columna no existe en el archivo de Moodle, se copian todas las calificaciones.
    """

    # Cargar los datos
    df_moodle = pd.read_excel(excelMoodle)  # Excel descargado de Moodle
    df_gea = pd.read_excel(excelActas)      # Excel de las actas (vacías) descargado de GEA

    # Si existe la columna indicada en ColumnaPresentados, filtrar solo los alumnos con valor 1
    if ColumnaPresentados in df_moodle.columns:
        df_moodle = df_moodle[df_moodle[ColumnaPresentados] == 1].copy()

    # Limpia el DNI en el df_moodle para que no contenga la letra final
    df_moodle['DNI_sin_letra'] = df_moodle['Número de ID'].str[:-1]

    # Mantiene los ceros a la izquierda en el df_moodle
    df_moodle['DNI_sin_letra'] = df_moodle['DNI_sin_letra'].str.zfill(8)

    # Mantiene los ceros a la izquierda en el df_gea
    df_gea['ALUMNO_DNI'] = df_gea['ALUMNO_DNI'].astype(str).str.zfill(8)

    # Realiza el merge (unión) entre los dos DataFrames
    merged_df = pd.merge(df_gea,
                         df_moodle[['DNI_sin_letra', nombre_columna_con_las_calificaciones]],
                         left_on='ALUMNO_DNI',
                         right_on='DNI_sin_letra',
                         how='left')

    # Rellena la columna 'CALIFICACION_NUMERICA' con las notas de la columna correspondiente en df_moodle
    # Los alumnos no presentados (o no encontrados) quedan con cadena vacía
    df_gea['CALIFICACION_NUMERICA'] = merged_df[nombre_columna_con_las_calificaciones].replace('-', '').fillna('')

    # Asegura que los NaN en 'CODIGO_CALIFICACION' del DataFrame de GEA se mantengan como vacíos
    df_gea['CODIGO_CALIFICACION'] = df_gea['CODIGO_CALIFICACION'].where(df_gea['CODIGO_CALIFICACION'].notna(), '')

    # Devuelve el DataFrame
    return df_gea
