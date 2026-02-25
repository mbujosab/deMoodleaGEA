import pandas as pd

def de_Moodle_a_actas(excelMoodle: str, excelActas: str, nombre_columna_con_las_calificaciones: str):
    """Función que cumplimenta una copia de las actas descargadas del portal de
    gestión académica con las notas calculadas en Moodle y almacenadas en un
    archivo Excel descargado desde Moodle.

    excelMoodle: nombre (incluido el path) del archivo descargado desde Moodle
    excelActas:  nombre (incluido el path) del archivo de actas descargado desde GEA
    nombre_columna_con_las_calificaciones: nombre de la columna del archivo excelMoodle donde están las notas finales
    """
    
    # Cargar los datos
    df_moodle = pd.read_excel(excelMoodle)  # Excel descargado de Moodle
    df_gea = pd.read_excel(excelActas)      # Excel de las actas (vacías) descargado de la página de calificación de actas (con extensión .xls).

    # Limpia el DNI en el df_moodle para que no contenga la letra final
    df_moodle['DNI_sin_letra'] = df_moodle['Número de ID'].str[:-1]

    # Mantiene los ceros a la izquierda en el df_moodle
    df_moodle['DNI_sin_letra'] = df_moodle['DNI_sin_letra'].str.zfill(8)

    # Mantiene los ceros a la izquierda en el df_gea
    df_gea['Doc. de identidad'] = df_gea['Doc. de identidad'].astype(str).str.zfill(8)

    # Realiza el merge (unión) entre los dos DataFrames
    merged_df = pd.merge(df_gea, 
                         df_moodle[['DNI_sin_letra', nombre_columna_con_las_calificaciones]],
                         left_on='Doc. de identidad', 
                         right_on='DNI_sin_letra', 
                         how='left')

    # Rellena la columna 'Nota num.' con las notas de la columna correspondiente en df_moodle y reemplaza '-' por vacíos
    df_gea['Nota num.'] = merged_df[nombre_columna_con_las_calificaciones].replace('-', '')

    # Asegura que los NaN en 'Calificación' del DataFrame de GEA se mantengan como vacíos
    df_gea['Calificación'] = df_gea['Calificación'].where(df_gea['Calificación'].notna(), '')

    # Devuelve el DataFrame
    return df_gea
