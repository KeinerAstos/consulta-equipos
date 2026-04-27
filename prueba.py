from flask import Flask, render_template, request, jsonify, redirect, url_for, session
import pandas as pd
import os
import psycopg2
import openpyxl
from collections import defaultdict
from psycopg2 import errors

app = Flask(__name__, template_folder='frontend')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ==========================================
# CARGA DE ARCHIVOS — SISTEM.xlsx
# ==========================================

ruta_sistem = os.path.join(BASE_DIR, 'datos', 'SISTEM.xlsx')
ruta_re     = os.path.join(BASE_DIR, 'datos', 'reasignacion.xlsx')

def limpiar_serial(col):
    """Limpieza estándar de seriales — aplicar siempre antes de cualquier comparación."""
    return (
        col.astype(str)
           .str.upper()
           .str.strip()
           .str.replace(r'\s+', '', regex=True)
           .str.replace(r'\.0$', '', regex=True)
    )

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

# Carga inicial
doc_entradas, doc_devoluciones, doc_salidas, doc_entregas, doc_envios = cargar_sistem()

resultado_cache = None


# ==========================================
# FUNCIÓN PRINCIPAL QUE GENERA EL TABLERO
# ==========================================
import pandas as pd
import numpy as np
def generar_tablero():
    """
    Versión de Trazabilidad Total: Muestra todos los equipos que han pasado por la bodega.
    Los equipos que salen no se eliminan, se marcan con estado 'SALIDA' o 'ENTREGA'.
    """
    BASURA = {'NAN', '', '#N/D', '#N/A', 'NONE', 'NAT', 'N/A', '0', '0.0'}

    def resolver_col(df, *candidatos):
        for c in candidatos:
            for col_real in df.columns:
                if str(col_real).strip().upper() == str(c).upper():
                    return col_real
        return None

    def normalizar(df, fecha_col, desc_col, tipo, signo):
        # 1. Búsqueda flexible de columnas
        col_serial = resolver_col(df, 'Serial', 'SERIAL', 'Nº Serie', 'NroSerie', 'Nro_Serie', 'Serie', 'NºSerieFab')
        if col_serial is None:
            return pd.DataFrame(columns=['Serial','SAP','Descripcion','Fecha','Tipo','Signo'])

        col_fecha = resolver_col(df, fecha_col, 'Fecha', 'FECHA SISTEMA.', 'Fecha Ingreso', 'Fecha Sistema', 'Fecha Salida')
        col_sap   = resolver_col(df, 'Codigo SAP', 'Material', 'SAP', 'Codigo material', 'CodigoSAP')
        col_desc  = resolver_col(df, desc_col, 'Descripción SAP', 'Descripción', 'Descripcion', 'Descripción Material')

        # 2. Creación del DataFrame normalizado
        sub = pd.DataFrame()
        sub['Serial']      = limpiar_serial(df[col_serial])
        sub['SAP']         = df[col_sap].astype(str).str.strip() if col_sap else 'N/A'
        sub['Descripcion'] = df[col_desc].astype(str).str.strip() if col_desc else 'N/A'
        sub['Fecha'] = pd.to_datetime(
            df[col_fecha].astype(str).str.strip(),
            errors='coerce',
            dayfirst=True,
            format='mixed'
        )
        sub['Tipo']        = tipo
        sub['Signo']       = signo
        
        # Limpieza de nulos y basura
        sub = sub[~sub['Serial'].isin(BASURA)].dropna(subset=['Serial'])
        return sub

    # --- 1. CARGA DE MOVIMIENTOS (Historial completo) ---
    mov_list = [
        normalizar(doc_entradas,     'Fecha Ingreso',  'Descripción',     'ENTRADA',    +1),
        normalizar(doc_devoluciones, 'FECHA SISTEMA.', 'Descripción',     'DEVOLUCION', +1),
        normalizar(doc_salidas,      'Fecha Salida',   'Descripción',     'SALIDA',     -1),
        normalizar(doc_entregas,     'Fecha Sistema',  'Descripción SAP', 'ENTREGA',    -1),
    ]
    movimientos = pd.concat([m for m in mov_list if not m.empty], ignore_index=True)

    if movimientos.empty:
        return pd.DataFrame()

    # --- 2. DETERMINAR EL ÚLTIMO MOVIMIENTO (Estado Actual) ---
    # El estado actual es simplemente el movimiento con la fecha/hora más reciente.
    # Si el Excel incluye hora en la columna de fecha, el desempate es automático y exacto.
    ultima_info = (
        movimientos
        .dropna(subset=['Fecha'])
        .sort_values(
            by=['Serial', 'Fecha'],
            ascending=[True, False],
            kind='mergesort'
        )
        .drop_duplicates(subset=['Serial'], keep='first')
        .copy()
    )

    # --- 3. DICCIONARIOS DE APOYO (ENVIOS Y OBSERVACIONES) ---
    envios_validos = doc_envios.copy()
    envios_validos['NºSerieFab'] = limpiar_serial(envios_validos['NºSerieFab'])
    envios_dict = envios_validos.drop_duplicates('NºSerieFab', keep='last').set_index('NºSerieFab').to_dict('index')

    col_obs = resolver_col(doc_entradas, 'Observación', 'Observaciones', 'Motivo', 'OT')
    obs_dict = {}
    if col_obs:
        obs_df = doc_entradas[['Serial', col_obs]].copy()
        obs_df['Serial'] = limpiar_serial(obs_df['Serial'])
        obs_dict = obs_df.drop_duplicates('Serial', keep='last').set_index('Serial')[col_obs].to_dict()

    # --- 4. CONSTRUCCIÓN DEL TABLERO FINAL ---
    MAPA_ESTADO = {
        'ENTRADA':    'DISPONIBLE (BODEGA)', 
        'DEVOLUCION': 'DISPONIBLE (DEVOLUCION)', 
        'ENTREGA':    'ENTREGADO (FUERA)', 
        'SALIDA':     'SALIDO (FUERA)'
    }
    SIN_OT = {'', 'NAN', 'N/A', 'NONE', 'STOCK', 'NAT', '-', 'S/N'}

    final_rows = []
    for _, row in ultima_info.iterrows():
        s = row['Serial']
        en_env = s in envios_dict
        env_data = envios_dict.get(s, {})

        # Si no hay datos en envíos, usamos los datos capturados del historial (ENTRADAS/DEVOLUCIONES)
        sap_final  = env_data.get('Material', row['SAP'])
        desc_final = env_data.get('Texto breve de material', row['Descripcion'])
        
        # Lógica de Estatus y Almacén basada en el signo del último movimiento
        obs_val = str(obs_dict.get(s, '')).strip().upper()
        
        if row['Signo'] < 0: # Si el último movimiento fue Salida o Entrega
            estatus = 'HISTORICO / SALIDA'
            almacen = 'FUERA DE BODEGA'
            aliado  = env_data.get('Destino', 'CLIENTE FINAL')
        else: # Si el último movimiento fue Entrada o Devolución (está en bodega)
            if en_env:
                estatus = 'RESERVADO'
                almacen = env_data.get('COD ALM', 'A500')
                aliado  = env_data.get('Destino', 'ALGARTECH')
            elif obs_val not in SIN_OT:
                estatus = 'RESERVADO (MANUAL)'
                almacen = 'Q500'
                aliado  = 'TRASLADO / MANUAL'
            else:
                estatus = 'STOCK'
                almacen = 'A500'
                aliado  = 'BODEGA PROPIA'

        final_rows.append({
            'SERIAL':           s,
            'OTP':              env_data.get('OTP', obs_val if obs_val not in SIN_OT else 'N/A'),
            'OTH':              env_data.get('OTH', 'N/A'),
            'Centro':           env_data.get('COD CENTRO', 'C903'),
            'ALMACEN':          almacen,
            'Aliado':           aliado,
            'CLIENTE':          env_data.get('CLIENTE', 'N/A'),
            'Tipo_de_OT':       env_data.get('PRC/SOLPED', 'BASE'),
            'Asignado':         'N/A',
            'Codigo_Sap':       sap_final,
            'Descripción':      desc_final,
            'CANTIDAD':         1,
            'Lote':             env_data.get('LOTE', 'VALORADO'),
            'Estado':           'FUNCIONAL',
            'Estado_Actual':    MAPA_ESTADO.get(row['Tipo'], 'DESCONOCIDO'),
            'Estatus':          estatus,
            'Fecha de ingreso': row['Fecha']
        })

    tablero = pd.DataFrame(final_rows)

    # --- 5. AGREGAR PENDIENTES DE INGRESO ---
    seriales_con_historial = set(movimientos['Serial'].unique())
    pendientes = envios_validos[~envios_validos['NºSerieFab'].isin(seriales_con_historial)].copy()

    if not pendientes.empty:
        p_rows = []
        for _, p in pendientes.iterrows():
            p_rows.append({
                'SERIAL': p['NºSerieFab'], 'OTP': p.get('OTP','N/A'), 'OTH': p.get('OTH','N/A'),
                'Centro': p.get('COD CENTRO','C903'), 'ALMACEN': p.get('COD ALM','A500'),
                'Aliado': p.get('Destino','ALGARTECH'), 'CLIENTE': p.get('CLIENTE','N/A'),
                'Tipo_de_OT': p.get('PRC/SOLPED','BASE'), 'Asignado': 'N/A',
                'Codigo_Sap': p.get('Material','N/A'), 'Descripción': p.get('Texto breve de material','N/A'),
                'CANTIDAD': 1, 'Lote': p.get('LOTE','VALORADO'), 'Estado': 'FUNCIONAL',
                'Estado_Actual': 'PENDIENTE DE INGRESO', 'Estatus': 'RESERVADO',
                'Fecha de ingreso': pd.NaT
            })
        tablero = pd.concat([tablero, pd.DataFrame(p_rows)], ignore_index=True)

    return tablero.sort_values('Fecha de ingreso', ascending=False).reset_index(drop=True)
# ==========================================
# FUNCIÓN REASIGNACIONES
# ==========================================

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


# ==========================================
# CACHÉ
# ==========================================

def obtener_resultado():
    global resultado_cache
    if resultado_cache is None:
        tabla = generar_tablero()
        tabla = aplicar_reasignaciones(tabla)
        resultado_cache = tabla
    return resultado_cache


def refrescar_cache():
    """Recarga SISTEM.xlsx desde disco y regenera el tablero completo."""
    global resultado_cache, doc_entradas, doc_devoluciones, doc_salidas, doc_entregas, doc_envios

    doc_entradas, doc_devoluciones, doc_salidas, doc_entregas, doc_envios = cargar_sistem()

    tabla = generar_tablero()
    tabla = aplicar_reasignaciones(tabla)
    resultado_cache = tabla


# ==========================================
# RUTAS
# ==========================================

@app.route("/buscar_seriales_sap", methods=["POST"])
def buscar_seriales_sap():

    resultado = obtener_resultado()
    sap = str(request.json.get("sap", "")).strip()

    resultado["Codigo_Sap"] = resultado["Codigo_Sap"].astype(str).str.strip()
    resultado["SERIAL"]     = resultado["SERIAL"].astype(str).str.strip()

    filtrado = resultado[resultado["Codigo_Sap"] == sap]
    seriales = filtrado["SERIAL"].drop_duplicates().tolist()

    return jsonify({"seriales": seriales[:20]})


@app.route("/buscar_serial", methods=["POST"])
def buscar_serial():

    resultado_final = obtener_resultado()
    serial = str(request.json.get("serial")).strip()

    fila = resultado_final[
        resultado_final["SERIAL"].astype(str).str.strip() == serial
    ]

    if not fila.empty:
        fila = fila.iloc[0]
        return jsonify({
            "existe":     True,
            "sap":        fila["Codigo_Sap"],
            "otp_actual": fila["OTP"],
            "cliente":    fila["CLIENTE"],
            "estatus":    fila["Estatus"]
        })

    return jsonify({"existe": False})


@app.route("/buscar_sap", methods=["POST"])
def buscar_sap():

    resultado = obtener_resultado()
    texto = str(request.json.get("sap", "")).strip()

    if texto == "":
        return jsonify({"saps": []})

    resultado["Codigo_Sap"] = resultado["Codigo_Sap"].astype(str).str.strip()

    filtrado = resultado[
        resultado["Codigo_Sap"].str.contains(texto, case=False, na=False)
    ]
    saps = filtrado["Codigo_Sap"].drop_duplicates().tolist()

    return jsonify({"saps": saps[:20]})


@app.route("/guardar_reasignacion", methods=["POST"])
def guardar_reasignacion():

    bodega        = request.form.get("bodega")
    centro        = request.form.get("centro")
    almacen       = request.form.get("almacen")
    otp_nueva     = request.form.get("otp_nueva")
    oth_nueva     = request.form.get("oth_nueva")
    nuevo_cliente = request.form.get("nuevo_cliente")
    tipo_ot       = request.form.get("tipo_ot")
    fecha_cambio  = request.form.get("fecha_cambio")

    saps         = request.form.getlist("sap[]")
    seriales     = request.form.getlist("serial[]")
    otp_actuales = request.form.getlist("otp_actual[]")

    workbook = openpyxl.load_workbook(ruta_re)
    sheet    = workbook.active

    for i in range(len(seriales)):
        nuevo_id   = sheet.max_row
        nueva_fila = [
            nuevo_id, bodega, centro, almacen,
            saps[i], seriales[i], otp_actuales[i],
            otp_nueva, oth_nueva, nuevo_cliente,
            tipo_ot, fecha_cambio
        ]
        sheet.append(nueva_fila)

    workbook.save(ruta_re)
    refrescar_cache()
    return redirect(url_for("home"))


@app.route("/tablero_stock")
def tablero_stock():
    resultado_final = obtener_resultado()
    datos = resultado_final.to_dict(orient="records")
    return render_template("tablero_stock.html", datos=datos)


@app.route("/")
def home():
    return render_template("reasignacion.html")


if __name__ == "__main__":
    app.run(debug=True, port=5000)