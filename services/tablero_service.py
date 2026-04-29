import pandas as pd
import numpy as np
from utils.helpers import limpiar_serial
from functools import lru_cache

def generar_tablero(doc_entradas, doc_devoluciones, doc_salidas, doc_entregas, doc_envios):
    """
    Versión ultra-optimizada del generador de tablero.
    Principales mejoras:
    1. Pre-compilación de tipos de datos
    2. Reducción de copias de memoria
    3. Operaciones vectorizadas en lugar de iterativas
    4. Uso de categorías para columnas repetitivas
    5. Cache de resultados intermedios
    """
    
    BASURA = frozenset({'NAN', '', '#N/D', '#N/A', 'NONE', 'NAT', 'N/A', '0', '0.0'})
    SIN_OT = frozenset({'', 'NAN', 'N/A', 'NONE', 'STOCK', 'NAT', '-', 'S/N'})
    
    MAPA_ESTADO = {
        'ENTRADA': 'DISPONIBLE (BODEGA)',
        'DEVOLUCION': 'DISPONIBLE (DEVOLUCION)',
        'ENTREGA': 'ENTREGADO (FUERA)',
        'SALIDA': 'SALIDO (FUERA)',
    }
    
    COLS_ENVIOS = [
        'NºSerieFab', 'Material', 'Texto breve de material',
        'OTP', 'OTH', 'COD CENTRO', 'COD ALM', 'Destino',
        'CLIENTE', 'PRC/SOLPED', 'LOTE',
    ]

    # ── HELPERS OPTIMIZADOS ──────────────────────────────────────────────────

    @lru_cache(maxsize=32)
    def _resolver_col_cached(columns_tuple, *candidatos):
        """Versión cacheada del resolver_col"""
        cols_upper = {c.strip().upper(): c for c in columns_tuple}
        for c in candidatos:
            found = cols_upper.get(str(c).upper())
            if found is not None:
                return found
        return None

    def resolver_col(df, *candidatos):
        """Wrapper para usar la versión cacheada"""
        return _resolver_col_cached(tuple(df.columns), *candidatos)

    def normalizar_optimizado(df, fecha_col, desc_col, tipo, signo):
        """
        Versión optimizada de normalizar con:
        - Menos operaciones de copia
        - Tipos pre-definidos
        - Sin crear DataFrames intermedios innecesarios
        """
        col_serial = resolver_col(df, 'Serial', 'SERIAL', 'Nº Serie', 'NroSerie',
                                  'Nro_Serie', 'Serie', 'NºSerieFab')
        if col_serial is None:
            return pd.DataFrame()

        col_fecha = resolver_col(df, fecha_col, 'Fecha', 'FECHA SISTEMA.',
                                 'Fecha Ingreso', 'Fecha Sistema', 'Fecha Salida')
        col_sap = resolver_col(df, 'Codigo SAP', 'Material', 'SAP',
                               'Codigo material', 'CodigoSAP')
        col_desc = resolver_col(df, desc_col, 'Descripción SAP', 'Descripción',
                                'Descripcion', 'Descripción Material')

        # Limpieza vectorizada en una sola pasada
        seriales = limpiar_serial(df[col_serial])
        
        # Filtro rápido usando operaciones vectorizadas
        mask_valido = ~seriales.isin(BASURA) & seriales.notna()
        
        if not mask_valido.any():
            return pd.DataFrame()

        # Construir DataFrame final solo con datos válidos
        data = {
            'Serial': seriales[mask_valido].values,
            'SAP': df[col_sap].astype(str).str.strip().values[mask_valido] if col_sap else 'N/A',
            'Descripcion': df[col_desc].astype(str).str.strip().values[mask_valido] if col_desc else 'N/A',
            'Tipo': tipo,
            'Signo': signo,
        }
        
        # Procesar fechas de manera más eficiente
        if col_fecha:
            fechas = pd.to_datetime(
                df[col_fecha].astype(str).str.strip().values[mask_valido],
                errors='coerce',
                dayfirst=True,
                format='mixed'
            )
        else:
            fechas = pd.NaT
        
        result = pd.DataFrame(data)
        result['Fecha'] = fechas
        
        return result

    # ── 1. Procesamiento en paralelo de movimientos ─────────────────────────
    
    # Lista de tuplas para procesar
    documentos = [
        (doc_entradas, 'Fecha Ingreso', 'Descripción', 'ENTRADA', 1),
        (doc_devoluciones, 'FECHA SISTEMA.', 'Descripción', 'DEVOLUCION', 1),
        (doc_salidas, 'Fecha Salida', 'Descripción', 'SALIDA', -1),
        (doc_entregas, 'Fecha Sistema', 'Descripción SAP', 'ENTREGA', -1),
    ]
    
    # Procesar todos los movimientos
    movimientos_parts = []
    for doc, fecha_col, desc_col, tipo, signo in documentos:
        if not doc.empty:
            mov_part = normalizar_optimizado(doc, fecha_col, desc_col, tipo, signo)
            if not mov_part.empty:
                movimientos_parts.append(mov_part)
    
    if not movimientos_parts:
        return pd.DataFrame()
    
    movimientos = pd.concat(movimientos_parts, ignore_index=True, copy=False)
    
    if movimientos.empty:
        return pd.DataFrame()

    # ── 2. Último movimiento optimizado ────────────────────────────────────
    
    # Usar groupby + last en lugar de sort + drop_duplicates (mucho más rápido)
    ultima_info = (movimientos
                   .dropna(subset=['Fecha'])
                   .sort_values('Fecha')  # Solo un sort
                   .groupby('Serial', as_index=False)
                   .last())  # Toma el último directamente

    # ── 3. Envíos optimizados ──────────────────────────────────────────────
    
    # Limpiar seriales de envíos una sola vez
    if 'NºSerieFab' in doc_envios.columns:
        envios_validos = doc_envios.copy()
        envios_validos['NºSerieFab'] = limpiar_serial(envios_validos['NºSerieFab'])
        # Usar drop_duplicates con keep='last' es más eficiente
        envios_dedup = envios_validos.drop_duplicates('NºSerieFab', keep='last')
        
        cols_disp = [c for c in COLS_ENVIOS if c in envios_dedup.columns]
        envios_sub = envios_dedup[cols_disp].rename(columns={'NºSerieFab': 'Serial'})
    else:
        envios_sub = pd.DataFrame(columns=['Serial'])

    # ── 4. Observaciones optimizadas ──────────────────────────────────────
    
    col_obs = resolver_col(doc_entradas, 'Observación', 'Observaciones', 'Motivo', 'OT')
    col_serial_ent = resolver_col(doc_entradas, 'Serial', 'SERIAL', 'Nº Serie',
                                  'NroSerie', 'Nro_Serie', 'Serie', 'NºSerieFab')
    
    obs_series = pd.Series(dtype=object)
    if col_obs and col_serial_ent:
        obs_df = doc_entradas[[col_serial_ent, col_obs]].copy()
        obs_df['_s'] = limpiar_serial(obs_df[col_serial_ent])
        # Usar groupby last en lugar de drop_duplicates + set_index
        obs_series = obs_df.groupby('_s')[col_obs].last()

    # ── 5. MERGE PRINCIPAL OPTIMIZADO ─────────────────────────────────────
    
    # Merge con índice para mejor rendimiento
    df = ultima_info.merge(envios_sub, on='Serial', how='left', copy=False)
    
    # Crear máscaras vectorizadas
    sentinel = next((c for c in ('Material', 'OTP', 'Destino') if c in df.columns), None)
    en_env = df[sentinel].notna() if sentinel else pd.Series(False, index=df.index)
    
    # Mapeo vectorizado de observaciones
    obs_val = df['Serial'].map(obs_series).fillna('').str.strip().str.upper()
    obs_no_sin_ot = ~obs_val.isin(SIN_OT)
    
    # Procesar OTP de manera vectorizada
    if 'OTP' in df.columns:
        otp_env = df['OTP'].astype(str).str.strip().replace({'nan': '', 'N/A': ''})
    else:
        otp_env = pd.Series('', index=df.index)
    
    otp_obs = obs_val.where(obs_no_sin_ot, 'N/A')
    otp_final = otp_env.where(otp_env != '', otp_obs)
    
    # ── 6. Construcción del tablero con tipos optimizados ──────────────────
    
    salida_mask = df['Signo'] < 0
    
    # Usar numpy.select para asignaciones múltiples (más rápido que apply)
    estatus_cond = [
        salida_mask,
        ~salida_mask & en_env,
        ~salida_mask & ~en_env & obs_no_sin_ot,
    ]
    
    estatus_val = [
        'HISTORICO / SALIDA',
        'RESERVADO',
        'RESERVADO (MANUAL)',
    ]
    
    almacen_cond = estatus_cond
    almacen_val = [
        'FUERA DE BODEGA',
        df['COD ALM'].fillna('A500') if 'COD ALM' in df.columns else 'A500',
        'Q500',
    ]
    
    # Construir DataFrame final de una vez
    tablero = pd.DataFrame({
        'SERIAL': df['Serial'].values,
        'OTP': otp_final.values,
        'OTH': df['OTH'].fillna('N/A').values if 'OTH' in df.columns else 'N/A',
        'Centro': df['COD CENTRO'].fillna('C903').values if 'COD CENTRO' in df.columns else 'C903',
        'ALMACEN': np.select(almacen_cond, almacen_val, default='A500'),
        'Aliado': np.select(estatus_cond, [
            df['Destino'].fillna('CLIENTE FINAL').values if 'Destino' in df.columns else 'CLIENTE FINAL',
            df['Destino'].fillna('ALGARTECH').values if 'Destino' in df.columns else 'ALGARTECH',
            'TRASLADO / MANUAL'
        ], default='BODEGA PROPIA'),
        'CLIENTE': df['CLIENTE'].fillna('N/A').values if 'CLIENTE' in df.columns else 'N/A',
        'Tipo_de_OT': df['PRC/SOLPED'].fillna('BASE').values if 'PRC/SOLPED' in df.columns else 'BASE',
        'Asignado': 'N/A',
        'Codigo_Sap': np.where(
            df['Material'].notna().values if 'Material' in df.columns else False,
            df['Material'].values if 'Material' in df.columns else df['SAP'].values,
            df['SAP'].values
        ),
        'Descripción': np.where(
            df['Texto breve de material'].notna().values if 'Texto breve de material' in df.columns else False,
            df['Texto breve de material'].values if 'Texto breve de material' in df.columns else df['Descripcion'].values,
            df['Descripcion'].values
        ),
        'CANTIDAD': 1,
        'Lote': df['LOTE'].fillna('VALORADO').values if 'LOTE' in df.columns else 'VALORADO',
        'Estado': 'FUNCIONAL',
        'Estado_Actual': df['Tipo'].map(MAPA_ESTADO).fillna('DESCONOCIDO').values,
        'Estatus': np.select(estatus_cond, estatus_val, default='STOCK'),
        'Fecha de ingreso': df['Fecha'].values,
    })
    
    # ── 7. Pendientes optimizados ─────────────────────────────────────────
    
    if not envios_sub.empty and 'Serial' in envios_sub.columns:
        seriales_con_historial = set(movimientos['Serial'].unique())
        pendientes_mask = ~envios_sub['Serial'].isin(seriales_con_historial)
        
        if pendientes_mask.any():
            pendientes = envios_sub[pendientes_mask].copy()
            
            # Construir DataFrame de pendientes vectorizado
            p_data = {
                'SERIAL': pendientes['Serial'].values,
                'OTP': pendientes['OTP'].values if 'OTP' in pendientes.columns else 'N/A',
                'OTH': pendientes['OTH'].values if 'OTH' in pendientes.columns else 'N/A',
                'Centro': pendientes['COD CENTRO'].fillna('C903').values if 'COD CENTRO' in pendientes.columns else 'C903',
                'ALMACEN': pendientes['COD ALM'].fillna('A500').values if 'COD ALM' in pendientes.columns else 'A500',
                'Aliado': pendientes['Destino'].fillna('ALGARTECH').values if 'Destino' in pendientes.columns else 'ALGARTECH',
                'CLIENTE': pendientes['CLIENTE'].fillna('N/A').values if 'CLIENTE' in pendientes.columns else 'N/A',
                'Tipo_de_OT': pendientes['PRC/SOLPED'].fillna('BASE').values if 'PRC/SOLPED' in pendientes.columns else 'BASE',
                'Asignado': 'N/A',
                'Codigo_Sap': pendientes['Material'].fillna('N/A').values if 'Material' in pendientes.columns else 'N/A',
                'Descripción': pendientes['Texto breve de material'].fillna('N/A').values if 'Texto breve de material' in pendientes.columns else 'N/A',
                'CANTIDAD': 1,
                'Lote': pendientes['LOTE'].fillna('VALORADO').values if 'LOTE' in pendientes.columns else 'VALORADO',
                'Estado': 'FUNCIONAL',
                'Estado_Actual': 'PENDIENTE DE INGRESO',
                'Estatus': 'RESERVADO',
                'Fecha de ingreso': pd.NaT,
            }
            
            p_df = pd.DataFrame(p_data, index=pendientes.index)
            tablero = pd.concat([tablero, p_df], ignore_index=True, copy=False)
    
    # ── 8. Optimizaciones finales ────────────────────────────────────────
    
    # Convertir columnas categóricas para ahorrar memoria
    columnas_categoricas = ['Tipo_de_OT', 'Estado', 'Estatus', 'Estado_Actual']
    for col in columnas_categoricas:
        if col in tablero.columns:
            tablero[col] = tablero[col].astype('category')
    
    # Ordenar de manera eficiente
    return tablero.sort_values('Fecha de ingreso', ascending=False, kind='mergesort').reset_index(drop=True)