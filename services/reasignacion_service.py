from config import ruta_re
import pandas as pd

def aplicar_reasignaciones(tabla):

    doc_reasignado = pd.read_excel(ruta_re)
    doc_reasignado.columns = doc_reasignado.columns.str.strip()
    doc_reasignado['SERIAL'] = doc_reasignado['SERIAL'].astype(str).str.strip()

    tabla['SERIAL'] = tabla['SERIAL'].astype(str).str.strip()

    tabla = tabla.merge(
        doc_reasignado[['SERIAL', 'OTP_NUEVA', 'NUEVO_CLIENTE']],
        on='SERIAL',
        how='left'
    )

    mask = tabla['OTP_NUEVA'].notna()

    tabla.loc[mask, 'OTP']        = tabla.loc[mask, 'OTP_NUEVA']
    tabla.loc[mask, 'CLIENTE']    = tabla.loc[mask, 'NUEVO_CLIENTE']
    tabla.loc[mask, 'Estatus']    = 'REASIGNADO'
    tabla.loc[mask, 'ALMACEN']    = 'Q500'
    tabla.loc[mask, 'Tipo_de_OT'] = 'INSTALACIONES'

    return tabla