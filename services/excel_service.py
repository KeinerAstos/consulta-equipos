import pandas as pd
from config import ruta_sistem
from utils.helpers import limpiar_serial


def cargar_sistem():
    # 1. Leer el Excel
    sheets = pd.read_excel(
        ruta_sistem,
        sheet_name=['ENTRADAS', 'DEVOLUCIONES', 'SALIDAS', 'ENTREGAS', 'ENVIOS'],
        dtype=str
    )

    # 2. Extraer CADA hoja explícitamente (Asegúrate de que los nombres coincidan)
    # Si sheets['ENTRADAS'] no existe por un espacio en blanco, esto fallará aquí y no en el tablero.
    try:
        return (
            sheets['ENTRADAS'], 
            sheets['DEVOLUCIONES'], 
            sheets['SALIDAS'], 
            sheets['ENTREGAS'], 
            sheets['ENVIOS']
        )
    except KeyError as e:
        print(f"Error: No se encontró la hoja {e} en el archivo Excel")
        # Retornar DataFrames vacíos para que no rompa el código posterior
        empty = pd.DataFrame()
        return empty, empty, empty, empty, empty