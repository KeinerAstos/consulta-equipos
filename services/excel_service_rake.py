import pandas as pd
from config import ruta_sistem
from utils.helpers import limpiar_serial


def tablas_limpias():
    doc_entregas = pd.read_excel(ruta_sistem, sheet_name='ENTREGAS')
    doc_devoluciones = pd.read_excel(ruta_sistem, sheet_name='DEVOLUCIONES')
    doc_salidas = pd.read_excel(ruta_sistem, sheet_name='SALIDAS')
    doc_entradas = pd.read_excel(ruta_sistem, sheet_name="ENTRADAS")
    doc_stock = pd.read_excel(ruta_sistem, sheet_name="STOCK")
    doc_envios = pd.read_excel(ruta_sistem, sheet_name="ENVIOS")
    doc_seriales_terreno = pd.read_excel(ruta_sistem,sheet_name="TERRENO_SERIALES")

    for df in [doc_entregas, doc_devoluciones, doc_salidas, doc_entradas,doc_stock]:
            if "Serial" in df.columns:
                df["Serial"] = df["Serial"].astype(str).str.strip()
                
    doc_envios["NºSerieFab"] = doc_envios["NºSerieFab"].astype(str).str.strip()

        # Conversión segura de columnas tipo fecha
    for df in [doc_entregas, doc_devoluciones, doc_salidas, doc_entradas]:
        for col in df.columns:
            col_str = str(col).lower()
            if "fecha" in col_str:
                df[col] = pd.to_datetime(df[col], errors="coerce")

    return doc_devoluciones,doc_entregas,doc_salidas,doc_entradas,doc_stock,doc_envios,doc_seriales_terreno