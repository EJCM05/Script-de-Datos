import os
import pandas as pd
from docx import Document
from docx.shared import Pt
from datetime import datetime

# Función para procesar una hoja y agregar su contenido al documento
def procesar_hoja(sheet_name, df, doc):
    """
    Procesa los datos de una hoja específica y los agrega al documento Word.
    
    Args:
        sheet_name (str): El nombre de la hoja de Excel.
        df (pd.DataFrame): El DataFrame de pandas con los datos de la hoja.
        doc (docx.Document): El objeto del documento de Word al que se agregará el contenido.
    """
    # Normalizar los nombres de las columnas
    df.columns = [col.lower().strip() for col in df.columns]

    # Ignorar columnas específicas
    columnas_a_ignorar = ["material utilizado", "status", "comisiones"]
    df = df[[col for col in df.columns if col not in columnas_a_ignorar]]

    # Verificar si quedan columnas para procesar
    if df.empty:
        print(f"No hay datos útiles en la hoja: {sheet_name}. Se omitirá.")
        return

    # Limpiar espacios en blanco en los datos y reemplazar celdas vacías por NaN
    # Usamos .copy() para evitar SettingWithCopyWarning
    df = df.replace(r"^\s*$", pd.NA, regex=True).copy()

    # Agregar el nombre de la hoja como título
    titulo = doc.add_heading(f'Datos de la hoja: {sheet_name}', level=2)
    titulo_run = titulo.runs[0]
    titulo_run.font.size = Pt(13.5)
    doc.add_paragraph() # Agrega un salto de línea después del título

    # Verificar si la columna "tipo de procedimiento" existe
    if "tipo de procedimiento" not in df.columns:
        print(f"La columna 'tipo de procedimiento' no se encontró en la hoja: {sheet_name}. Se omitirá el análisis detallado.")
        return

    # Conteo por "Tipo de Procedimiento"
    conteo_tipos = df["tipo de procedimiento"].value_counts(dropna=True)

    # Agregar al documento el conteo
    doc.add_heading("Conteo por Tipo de Procedimiento", level=3)
    for tipo, cantidad in conteo_tipos.items():
        doc.add_paragraph(f"{tipo}: {cantidad} Procedimientos")

    # Separación de datos por "Tipo de Procedimiento"
    doc.add_heading("Datos separados por Tipo de Procedimiento", level=3)
    for tipo in df["tipo de procedimiento"].dropna().unique():
        # Agregamos un subtítulo para cada tipo de procedimiento
        tipo_subtitulo = doc.add_heading(f'Tipo de Procedimiento: {tipo}', level=4)
        tipo_subtitulo.runs[0].bold = True

        # Filtrar las filas correspondientes al tipo actual
        df_tipo = df[df["tipo de procedimiento"] == tipo]

        # Verificar columnas con datos para este tipo de procedimiento
        columnas_con_datos = [
            col for col in df_tipo.columns if df_tipo[col].notna().any()
        ]
        df_tipo = df_tipo[columnas_con_datos]

        # Exportar filas al documento
        for _, item in df_tipo.iterrows():
            data_line = []
            for col in columnas_con_datos:
                valor = item.get(col, None)
                if pd.notna(valor):  # Agregar solo si el valor no es NaN
                    if isinstance(valor, str):
                        data_line.append(f"{col.replace('_', ' ').title()}: {valor.strip()}")
                    else:
                        data_line.append(f"{col.replace('_', ' ').title()}: {valor}")
            # Agregar solo si hay datos válidos en esta fila
            if data_line:
                doc.add_paragraph(', '.join(data_line))
    

# Directorio donde se encuentran los archivos .xlsx y el script
directorio = './por_procesar/junio/'

# Buscar todos los archivos con extensión .xlsx
archivos_xlsx = [archivo for archivo in os.listdir(directorio) if archivo.endswith('.xlsx')]

# --- Inicio de la lógica principal ---
# Verificar si se encontraron archivos .xlsx
if archivos_xlsx:
    print(f"Archivos Excel encontrados: {archivos_xlsx}")
    
    # Crear un nuevo documento de Word que contendrá todos los análisis
    doc = Document()
    print("Creando documento de Word...")

    # Procesar cada archivo encontrado
    for archivo_xlsx in archivos_xlsx:
        print(f"\n--- Procesando archivo: {archivo_xlsx} ---")
        
        # Agregar un título para el archivo actual en el documento
        doc.add_heading(f"Análisis del archivo: {archivo_xlsx}", level=1)
        
        ruta_completa = os.path.join(directorio, archivo_xlsx)
        try:
            xls = pd.ExcelFile(ruta_completa, engine='openpyxl')
        except Exception as e:
            print(f"Error al leer el archivo Excel: {e}")
            continue # Continuar con el siguiente archivo si hay un error

        # Procesar cada hoja del archivo actual
        for sheet_name in xls.sheet_names:
            print(f"  Procesando hoja: {sheet_name}")
            df = pd.read_excel(ruta_completa, sheet_name=sheet_name)
            procesar_hoja(sheet_name, df, doc)
        
        # Agrega un salto de página después de cada archivo para mantener la individualidad
        doc.add_page_break()

    # --- Lógica de guardado después de procesar todos los archivos ---
    try:
        # Generar un nombre de archivo único y descriptivo para el documento final
        fecha_actual = datetime.now().strftime('%Y%m%d')
        nombre_salida = f'analisis_consolidado_{fecha_actual}.docx'
        
        doc.save(nombre_salida)
        print(f"\nProceso completado. Datos de todos los archivos exportados a '{nombre_salida}' correctamente.")
    except Exception as e:
        print(f"Error al exportar a Word: {e}")
else:
    print("No se encontró ningún archivo .xlsx en el directorio.")

print("Directorio actual:", os.getcwd())
