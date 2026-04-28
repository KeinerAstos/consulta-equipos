import pandas as pd
import numpy as np
from utils.helpers import limpiar_serial


def generar_tablero(doc_entradas, doc_devoluciones, doc_salidas, doc_entregas, doc_envios):

    BASURA    = {'NAN', '', '#N/D', '#N/A', 'NONE', 'NAT', 'N/A', '0', '0.0'}
    SIN_OT    = {'', 'NAN', 'N/A', 'NONE', 'STOCK', 'NAT', '-', 'S/N'}
    MAPA_ESTADO = {
        'ENTRADA':    'DISPONIBLE (BODEGA)',
        'DEVOLUCION': 'DISPONIBLE (DEVOLUCION)',
        'ENTREGA':    'ENTREGADO (FUERA)',
        'SALIDA':     'SALIDO (FUERA)',
    }
    # Columnas que se intentan traer de doc_envios
    COLS_ENVIOS = [
        'NºSerieFab', 'Material', 'Texto breve de material',
        'OTP', 'OTH', 'COD CENTRO', 'COD ALM', 'Destino',
        'CLIENTE', 'PRC/SOLPED', 'LOTE',
    ]

    # ── Helpers ───────────────────────────────────────────────────────────────

    def resolver_col(df, *candidatos):
        """
        OPTIMIZACIÓN 1: resolver_col original hacía un doble bucle O(candidatos × columnas).
        Aquí se construye un dict upper→nombre_real una sola vez → cada búsqueda es O(1).
        """
        cols_upper = {str(c).strip().upper(): c for c in df.columns}
        for c in candidatos:
            found = cols_upper.get(str(c).upper())
            if found is not None:
                return found
        return None

    def gcol(df, col, default='N/A'):
        """Devuelve la columna si existe; si no, una Serie con el valor por defecto."""
        return df[col] if col in df.columns else pd.Series(default, index=df.index)

    def normalizar(df, fecha_col, desc_col, tipo, signo):
        col_serial = resolver_col(df, 'Serial', 'SERIAL', 'Nº Serie', 'NroSerie',
                                  'Nro_Serie', 'Serie', 'NºSerieFab')
        if col_serial is None:
            return pd.DataFrame(columns=['Serial', 'SAP', 'Descripcion', 'Fecha', 'Tipo', 'Signo'])

        col_fecha = resolver_col(df, fecha_col, 'Fecha', 'FECHA SISTEMA.',
                                 'Fecha Ingreso', 'Fecha Sistema', 'Fecha Salida')
        col_sap   = resolver_col(df, 'Codigo SAP', 'Material', 'SAP',
                                 'Codigo material', 'CodigoSAP')
        col_desc  = resolver_col(df, desc_col, 'Descripción SAP', 'Descripción',
                                 'Descripcion', 'Descripción Material')

        fecha_raw = df[col_fecha].astype(str).str.strip() if col_fecha else pd.Series('', index=df.index)
        sub = pd.DataFrame({
            'Serial':      limpiar_serial(df[col_serial]),
            'SAP':         df[col_sap].astype(str).str.strip() if col_sap else 'N/A',
            'Descripcion': df[col_desc].astype(str).str.strip() if col_desc else 'N/A',
            'Fecha':       pd.to_datetime(fecha_raw, errors='coerce', dayfirst=True, format='mixed'),
            'Tipo':        tipo,
            'Signo':       signo,
        })
        return sub[~sub['Serial'].isin(BASURA)].dropna(subset=['Serial'])

    # ── 1. Movimientos ────────────────────────────────────────────────────────
    mov_list = [
        normalizar(doc_entradas,     'Fecha Ingreso',  'Descripción',     'ENTRADA',    +1),
        normalizar(doc_devoluciones, 'FECHA SISTEMA.', 'Descripción',     'DEVOLUCION', +1),
        normalizar(doc_salidas,      'Fecha Salida',   'Descripción',     'SALIDA',     -1),
        normalizar(doc_entregas,     'Fecha Sistema',  'Descripción SAP', 'ENTREGA',    -1),
    ]
    movimientos = pd.concat([m for m in mov_list if not m.empty], ignore_index=True)

    if movimientos.empty:
        return pd.DataFrame()

    # ── 2. Último movimiento por serial ───────────────────────────────────────
    ultima_info = (
        movimientos
        .dropna(subset=['Fecha'])
        .sort_values(by=['Serial', 'Fecha'], ascending=[True, False], kind='mergesort')
        .drop_duplicates(subset=['Serial'], keep='first')
        .reset_index(drop=True)
    )

    # ── 3. Preparar envíos ────────────────────────────────────────────────────
    envios_validos = doc_envios.copy()
    envios_validos['NºSerieFab'] = limpiar_serial(envios_validos['NºSerieFab'])
    envios_dedup = envios_validos.drop_duplicates('NºSerieFab', keep='last')

    cols_disp  = [c for c in COLS_ENVIOS if c in envios_dedup.columns]
    envios_sub = envios_dedup[cols_disp].rename(columns={'NºSerieFab': 'Serial'})

    # ── 4. Observaciones ──────────────────────────────────────────────────────
    # CORRECCIÓN: el código original asumía que la columna se llamaba 'Serial'
    # exactamente; ahora usamos resolver_col para encontrarla correctamente.
    col_obs        = resolver_col(doc_entradas, 'Observación', 'Observaciones', 'Motivo', 'OT')
    col_serial_ent = resolver_col(doc_entradas, 'Serial', 'SERIAL', 'Nº Serie',
                                  'NroSerie', 'Nro_Serie', 'Serie', 'NºSerieFab')
    obs_series = pd.Series(dtype=object)
    if col_obs and col_serial_ent:
        _obs = doc_entradas[[col_serial_ent, col_obs]].copy()
        _obs['_s'] = limpiar_serial(_obs[col_serial_ent])
        obs_series = _obs.drop_duplicates('_s', keep='last').set_index('_s')[col_obs]

    # ── 5. MERGE VECTORIZADO (reemplaza el iterrows) ──────────────────────────
    #
    # OPTIMIZACIÓN 2 (principal): El loop for + iterrows recorría cada fila
    # individualmente haciendo dict.get() por celda → O(n).
    # Un merge de pandas ejecuta la misma lógica en C → entre 10x y 100x más rápido.
    #
    df = ultima_info.merge(envios_sub, on='Serial', how='left')

    # ¿El serial tiene datos en envíos? Lo detectamos por cualquier columna exclusiva.
    sentinel = next((c for c in ('Material', 'OTP', 'Destino') if c in df.columns), None)
    en_env   = df[sentinel].notna() if sentinel else pd.Series(False, index=df.index)

    # Observaciones mapeadas vectorialmente
    obs_val       = df['Serial'].map(obs_series).fillna('').str.strip().str.upper()
    obs_no_sin_ot = ~obs_val.isin(SIN_OT)

    # OTP: preferir envíos; si no, usar observación; si no aplica, 'N/A'
    otp_env = gcol(df, 'OTP').astype(str).str.strip().replace({'nan': '', 'N/A': ''})
    otp_obs = obs_val.where(obs_no_sin_ot, 'N/A')
    otp_final = otp_env.where(otp_env != '', otp_obs)

    # Máscaras de condición (misma lógica if/elif/else del original)
    salida_mask = df['Signo'] < 0
    cond = [
        salida_mask,                            # último mov. fue Salida/Entrega
        ~salida_mask & en_env,                  # está en bodega Y en envíos
        ~salida_mask & ~en_env & obs_no_sin_ot, # está en bodega, manual
    ]

    df['Estatus'] = np.select(cond,
        ['HISTORICO / SALIDA', 'RESERVADO', 'RESERVADO (MANUAL)'],
        default='STOCK')

    df['ALMACEN'] = np.select(cond,
        ['FUERA DE BODEGA', gcol(df, 'COD ALM', 'A500').fillna('A500'), 'Q500'],
        default='A500')

    dest = gcol(df, 'Destino').astype(str).replace('nan', np.nan)
    df['Aliado'] = np.select(cond,
        [dest.fillna('CLIENTE FINAL'), dest.fillna('ALGARTECH'), 'TRASLADO / MANUAL'],
        default='BODEGA PROPIA')

    # Campos con fallback envíos → historial
    mat = gcol(df, 'Material').astype(str).replace('nan', np.nan)
    tbm = gcol(df, 'Texto breve de material').astype(str).replace('nan', np.nan)
    df['Codigo_Sap'] = mat.where(mat.notna(), df['SAP'])
    df['Descripción_f'] = tbm.where(tbm.notna(), df['Descripcion'])

    # ── 6. Construcción del tablero final (vectorizada) ───────────────────────
    tablero = pd.DataFrame({
        'SERIAL':           df['Serial'],
        'OTP':              otp_final,
        'OTH':              gcol(df, 'OTH', 'N/A').fillna('N/A'),
        'Centro':           gcol(df, 'COD CENTRO', 'C903').fillna('C903'),
        'ALMACEN':          df['ALMACEN'],
        'Aliado':           df['Aliado'],
        'CLIENTE':          gcol(df, 'CLIENTE', 'N/A').fillna('N/A'),
        'Tipo_de_OT':       gcol(df, 'PRC/SOLPED', 'BASE').fillna('BASE'),
        'Asignado':         'N/A',
        'Codigo_Sap':       df['Codigo_Sap'],
        'Descripción':      df['Descripción_f'],
        'CANTIDAD':         1,
        'Lote':             gcol(df, 'LOTE', 'VALORADO').fillna('VALORADO'),
        'Estado':           'FUNCIONAL',
        'Estado_Actual':    df['Tipo'].map(MAPA_ESTADO).fillna('DESCONOCIDO'),
        'Estatus':          df['Estatus'],
        'Fecha de ingreso': df['Fecha'],
    })

    # ── 7. Pendientes de ingreso ──────────────────────────────────────────────
    #
    # OPTIMIZACIÓN 3: el segundo iterrows también se elimina construyendo
    # el DataFrame directamente desde las columnas del slice filtrado.
    #
    seriales_con_historial = set(movimientos['Serial'].unique())
    pendientes = envios_validos[~envios_validos['NºSerieFab'].isin(seriales_con_historial)].copy()

    if not pendientes.empty:
        def pcol(col, default='N/A'):
            return pendientes[col] if col in pendientes.columns else default

        p_df = pd.DataFrame({
            'SERIAL':           pendientes['NºSerieFab'],
            'OTP':              pcol('OTP'),
            'OTH':              pcol('OTH'),
            'Centro':           pcol('COD CENTRO', 'C903'),
            'ALMACEN':          pcol('COD ALM', 'A500'),
            'Aliado':           pcol('Destino', 'ALGARTECH'),
            'CLIENTE':          pcol('CLIENTE'),
            'Tipo_de_OT':       pcol('PRC/SOLPED', 'BASE'),
            'Asignado':         'N/A',
            'Codigo_Sap':       pcol('Material'),
            'Descripción':      pcol('Texto breve de material'),
            'CANTIDAD':         1,
            'Lote':             pcol('LOTE', 'VALORADO'),
            'Estado':           'FUNCIONAL',
            'Estado_Actual':    'PENDIENTE DE INGRESO',
            'Estatus':          'RESERVADO',
            'Fecha de ingreso': pd.NaT,
        }, index=pendientes.index)

        tablero = pd.concat([tablero, p_df], ignore_index=True)

    return tablero.sort_values('Fecha de ingreso', ascending=False).reset_index(drop=True)