import pandas as pd
from functools import lru_cache
from config import ruta_sistem

HOJAS = ['ENTRADAS','DEVOLUCIONES','SALIDAS','ENTREGAS','ENVIOS']

COLUMNAS = {
    "ENTRADAS": ['Serial','Codigo SAP','Descripción','Fecha Ingreso','Observación'],
    "DEVOLUCIONES": ['Serial','Codigo SAP','Descripción','FECHA SISTEMA.'],
    "SALIDAS": ['Serial','Codigo SAP','Descripción','Fecha Salida'],
    "ENTREGAS": ['Serial','Codigo SAP','Descripción SAP','Fecha Sistema'],
    "ENVIOS": ['NºSerieFab','Material','Texto breve de material','OTP','OTH',
               'COD CENTRO','COD ALM','Destino','CLIENTE','PRC/SOLPED','LOTE']
}

@lru_cache(maxsize=1)
def cargar_sistem():
    try:
        resultado = []

        for hoja in HOJAS:
            df = pd.read_excel(
                ruta_sistem,
                sheet_name=hoja,
                dtype=str,
                usecols=COLUMNAS[hoja],
                engine="openpyxl"
            )
            resultado.append(df)

        return tuple(resultado)

    except Exception as e:
        print(f"Error leyendo excel: {e}")
        empty = pd.DataFrame()
        return (empty,) * 5