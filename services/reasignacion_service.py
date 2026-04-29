from config import ruta_re
import pandas as pd

def aplicar_reasignaciones(tabla):
    """
    Aplica las reasignaciones desde el archivo de reasignaciones.
    Versión corregida que maneja columnas categóricas.
    """
    
    # Cargar documento de reasignaciones
    doc_reasignado = pd.read_excel(ruta_re)
    doc_reasignado.columns = doc_reasignado.columns.str.strip()
    doc_reasignado['SERIAL'] = doc_reasignado['SERIAL'].astype(str).str.strip()

    tabla['SERIAL'] = tabla['SERIAL'].astype(str).str.strip()

    # Guardar el tipo original de las columnas que vamos a modificar
    columnas_a_modificar = ['Estatus', 'ALMACEN', 'Tipo_de_OT', 'OTP', 'CLIENTE']
    tipos_originales = {}
    
    for col in columnas_a_modificar:
        if col in tabla.columns:
            tipos_originales[col] = tabla[col].dtype
            # Si es categórica, convertir temporalmente a string
            if hasattr(tabla[col], 'cat'):
                tabla[col] = tabla[col].astype(str)

    # Merge con reasignaciones
    tabla = tabla.merge(
        doc_reasignado[['SERIAL', 'OTP_NUEVA', 'NUEVO_CLIENTE']],
        on='SERIAL',
        how='left'
    )

    # Máscara para equipos reasignados
    mask = tabla['OTP_NUEVA'].notna()

    # Aplicar cambios
    tabla.loc[mask, 'OTP']        = tabla.loc[mask, 'OTP_NUEVA']
    tabla.loc[mask, 'CLIENTE']    = tabla.loc[mask, 'NUEVO_CLIENTE']
    tabla.loc[mask, 'Estatus']    = 'REASIGNADO'
    tabla.loc[mask, 'ALMACEN']    = 'Q500'
    tabla.loc[mask, 'Tipo_de_OT'] = 'INSTALACIONES'

    # También actualizar Estado_Actual para los reasignados
    if 'Estado_Actual' in tabla.columns:
        if hasattr(tabla['Estado_Actual'], 'cat'):
            tabla['Estado_Actual'] = tabla['Estado_Actual'].astype(str)
        tabla.loc[mask, 'Estado_Actual'] = 'REASIGNADO'

    # Eliminar columnas temporales del merge
    tabla = tabla.drop(['OTP_NUEVA', 'NUEVO_CLIENTE'], axis=1)

    # Opcional: Restaurar tipos categóricos para mejor rendimiento
    # Agregar las nuevas categorías antes de convertir
    if 'Estatus' in tabla.columns:
        categorias_actuales = tabla['Estatus'].unique().tolist()
        tabla['Estatus'] = pd.Categorical(
            tabla['Estatus'], 
            categories=categorias_actuales
        )
    
    if 'Estado_Actual' in tabla.columns:
        categorias_actuales = tabla['Estado_Actual'].unique().tolist()
        tabla['Estado_Actual'] = pd.Categorical(
            tabla['Estado_Actual'], 
            categories=categorias_actuales
        )
    
    if 'Tipo_de_OT' in tabla.columns:
        categorias_actuales = tabla['Tipo_de_OT'].unique().tolist()
        tabla['Tipo_de_OT'] = pd.Categorical(
            tabla['Tipo_de_OT'], 
            categories=categorias_actuales
        )

    return tabla