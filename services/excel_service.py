import pandas as pd
from config import ruta_sistem
from utils.helpers import limpiar_serial



def cargar_sistem():
    """
    Lee todas las hojas históricas de SISTEM.xlsx.
    Toda la limpieza de datos ocurre aquí una sola vez.
    """
    sheets = pd.read_excel(
        ruta_sistem,
        sheet_name=['ENTRADAS', 'DEVOLUCIONES', 'SALIDAS', 'ENTREGAS', 'ENVIOS'],
        dtype=str          # leer todo como texto evita que Excel convierta seriales a float
    )

    doc_entradas     = sheets['ENTRADAS']
    doc_devoluciones = sheets['DEVOLUCIONES']
    doc_salidas      = sheets['SALIDAS']
    doc_entregas     = sheets['ENTREGAS']
    doc_envios       = sheets['ENVIOS']

    # Normalizar nombres de columna
    for df in [doc_entradas, doc_devoluciones, doc_salidas, doc_entregas, doc_envios]:
        df.columns = df.columns.str.strip()

    # Limpiar seriales con función estándar
    for df in [doc_entradas, doc_devoluciones, doc_salidas, doc_entregas]:
        if 'Serial' in df.columns:
            df['Serial'] = limpiar_serial(df['Serial'])

    doc_envios['NºSerieFab'] = limpiar_serial(doc_envios['NºSerieFab'])

    # Parsear fechas solo en columnas que las tienen
    fecha_map = {
        'doc_entradas':     'Fecha Ingreso',
        'doc_devoluciones': 'FECHA SISTEMA.',
        'doc_salidas':      'Fecha Salida',
        'doc_entregas':     'Fecha Sistema',
    }
    for df, col in [
        (doc_entradas, 'Fecha Ingreso'),
        (doc_devoluciones, 'FECHA SISTEMA.'),
        (doc_salidas, 'Fecha Salida'),
        (doc_entregas, 'Fecha Sistema')
    ]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    return doc_entradas, doc_devoluciones, doc_salidas, doc_entregas, doc_envios