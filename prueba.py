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

def cargar_sistem():
    """
    Lee todas las hojas históricas de SISTEM.xlsx y devuelve los DataFrames limpios.
    Se llama al arrancar y cada vez que se refresca el caché.
    """
    doc_entradas     = pd.read_excel(ruta_sistem, sheet_name='ENTRADAS')
    doc_devoluciones = pd.read_excel(ruta_sistem, sheet_name='DEVOLUCIONES')
    doc_salidas      = pd.read_excel(ruta_sistem, sheet_name='SALIDAS')
    doc_entregas     = pd.read_excel(ruta_sistem, sheet_name='ENTREGAS')
    doc_envios       = pd.read_excel(ruta_sistem, sheet_name='ENVIOS')

    # Normalizar nombres de columna (quitar espacios)
    for df in [doc_entradas, doc_devoluciones, doc_salidas, doc_entregas, doc_envios]:
        df.columns = df.columns.str.strip()

    # Limpiar serial en todas las hojas
    for df in [doc_entradas, doc_devoluciones, doc_salidas, doc_entregas]:
        df['Serial'] = df['Serial'].astype(str).str.strip()

    doc_envios['NºSerieFab'] = doc_envios['NºSerieFab'].astype(str).str.strip()

    return doc_entradas, doc_devoluciones, doc_salidas, doc_entregas, doc_envios


# Carga inicial
doc_entradas, doc_devoluciones, doc_salidas, doc_entregas, doc_envios = cargar_sistem()

resultado_cache = None


# ==========================================
# FUNCIÓN PRINCIPAL QUE GENERA EL TABLERO
# ==========================================

def generar_tablero():
    """
    Reconstruye el stock actual a partir del historial completo de SISTEM.xlsx:

        ENTRADAS     → signo +1  (equipos que ingresaron a bodega)
        DEVOLUCIONES → signo +1  (equipos devueltos = vuelven a bodega)
        SALIDAS      → signo -1  (equipos que salieron de bodega)
        ENTREGAS     → signo -1  (equipos entregados a técnico = salen de bodega)

    Un serial está EN BODEGA si la suma de signos > 0.
    """

    # --------------------------------------------------
    # 1. NORMALIZAR LAS 4 HOJAS AL MISMO ESQUEMA
    #    Columnas resultado: Serial | SAP | Descripcion | Fecha | Signo
    # --------------------------------------------------

    def normalizar(df, fecha_col, descripcion_col, signo, tipo):
        """Extrae solo las columnas necesarias, asigna el signo y el tipo de movimiento."""
        sub = df[['Serial', 'Codigo SAP', descripcion_col, fecha_col]].copy()
        sub.columns = ['Serial', 'SAP', 'Descripcion', 'Fecha']
        sub['Fecha']  = pd.to_datetime(sub['Fecha'], errors='coerce')
        sub['Signo']  = signo
        sub['Tipo']   = tipo          # <-- etiqueta del movimiento
        sub['Serial'] = sub['Serial'].astype(str).str.strip()
        sub['SAP']    = sub['SAP'].astype(str).str.strip()
        return sub

    # Ajusta los nombres de columna si difieren en tu archivo
    movimientos = pd.concat([
        normalizar(doc_entradas,     fecha_col='Fecha Ingreso',   descripcion_col='Descripción',     signo=+1, tipo='ENTRADA'),
        normalizar(doc_devoluciones, fecha_col='FECHA SISTEMA.',  descripcion_col='Descripción',     signo=+1, tipo='DEVOLUCION'),
        normalizar(doc_salidas,      fecha_col='Fecha Salida',    descripcion_col='Descripción',     signo=-1, tipo='SALIDA'),
        normalizar(doc_entregas,     fecha_col='Fecha Sistema',   descripcion_col='Descripción SAP', signo=-1, tipo='ENTREGA'),
    ], ignore_index=True)

    # Descartar filas sin serial válido
    movimientos = movimientos[
        ~movimientos['Serial'].isin(['nan', 'NaN', '', '#N/D', '#N/A'])
    ]

    # --------------------------------------------------
    # 2. CALCULAR STOCK ACTUAL POR SERIAL
    # --------------------------------------------------

    stock_signos = (
        movimientos
        .groupby(['Serial', 'SAP', 'Descripcion'], as_index=False)['Signo']
        .sum()
    )

    # Solo los que siguen en bodega (suma positiva)
    en_bodega = stock_signos[stock_signos['Signo'] > 0].copy()

    # Último movimiento por serial: fecha Y tipo
    idx_ultima = movimientos.dropna(subset=['Fecha']).groupby('Serial')['Fecha'].idxmax()
    ultimo_mov = (
        movimientos.loc[idx_ultima, ['Serial', 'Fecha', 'Tipo']]
        .rename(columns={'Fecha': 'UltimaFecha', 'Tipo': 'UltimoTipo'})
    )

    en_bodega = en_bodega.merge(ultimo_mov, on='Serial', how='left')

    # Mapear tipo de movimiento → etiqueta de estado legible
    mapa_estado = {
        'ENTRADA':    'DISPONIBLE',
        'DEVOLUCION': 'DISPONIBLE',
        'ENTREGA':    'EN ENTREGA',
        'SALIDA':     'EN SALIDA',
    }
    en_bodega['Estado_Actual'] = en_bodega['UltimoTipo'].map(mapa_estado).fillna('DISPONIBLE')

    # --------------------------------------------------
    # 3. ARMAR TABLA FINAL
    # --------------------------------------------------

    filas = []

    # Seriales ya procesados desde el historial de movimientos
    seriales_con_historial = set(en_bodega['Serial'].tolist())

    # ── PASO 1: equipos con historial de movimientos ──────────────────────────
    for _, row in en_bodega.iterrows():

        serial        = row['Serial']
        sap           = row['SAP']
        descripcion   = row['Descripcion']
        fecha         = row['UltimaFecha']
        cantidad      = int(row['Signo'])
        estado_actual = row['Estado_Actual']

        if serial in doc_envios['NºSerieFab'].values:
            envio = doc_envios[doc_envios['NºSerieFab'] == serial].iloc[0]
            filas.append({
                "OTP":              envio.get('OTP', 'N/A'),
                "OTH":              envio.get('OTH', 'N/A'),
                "Centro":           envio.get('COD CENTRO', 'C903'),
                "ALMACEN":          envio.get('COD ALM', 'A500'),
                "Aliado":           envio.get('Destino', 'N/A'),
                "CLIENTE":          envio.get('CLIENTE', 'N/A'),
                "Tipo_de_OT":       envio.get('PRC/SOLPED', 'BASE'),
                "Asignado":         "N/A",
                "Codigo_Sap":       envio.get('Material', sap),
                "Descripción":      envio.get('Texto breve de material', descripcion),
                "CANTIDAD":         1,
                "SERIAL":           serial,
                "Lote":             envio.get('LOTE', 'VALORADO'),
                "Estado":           "FUNCIONAL",
                "Estado_Actual":    estado_actual,
                "Estatus":          "RESERVADO",
                "Fecha de ingreso": envio.get('Fecha', fecha)
            })
        else:
            filas.append({
                "OTP":              "N/A",
                "OTH":              "N/A",
                "Centro":           "C903",
                "ALMACEN":          "A500",
                "Aliado":           "ALGARTECH",
                "CLIENTE":          "N/A",
                "Tipo_de_OT":       "BASE",
                "Asignado":         "N/A",
                "Codigo_Sap":       sap,
                "Descripción":      descripcion,
                "CANTIDAD":         cantidad,
                "SERIAL":           serial,
                "Lote":             "VALORADO",
                "Estado":           "FUNCIONAL",
                "Estado_Actual":    estado_actual,
                "Estatus":          "STOCK",
                "Fecha de ingreso": fecha
            })

    # ── PASO 2: equipos en ENVIOS sin ningún movimiento registrado ────────────
    #    Nunca entraron físicamente → RESERVADO + PENDIENTE DE INGRESO
    envios_sin_historial = doc_envios[
        ~doc_envios['NºSerieFab'].isin(seriales_con_historial)
    ]

    for _, envio in envios_sin_historial.iterrows():

        serial = str(envio.get('NºSerieFab', 'N/A')).strip()

        if serial in ('nan', 'NaN', '', '#N/D', '#N/A'):
            continue

        filas.append({
            "OTP":              envio.get('OTP', 'N/A'),
            "OTH":              envio.get('OTH', 'N/A'),
            "Centro":           envio.get('COD CENTRO', 'C903'),
            "ALMACEN":          envio.get('COD ALM', 'A500'),
            "Aliado":           envio.get('Destino', 'N/A'),
            "CLIENTE":          envio.get('CLIENTE', 'N/A'),
            "Tipo_de_OT":       envio.get('PRC/SOLPED', 'BASE'),
            "Asignado":         "N/A",
            "Codigo_Sap":       envio.get('Material', 'N/A'),
            "Descripción":      envio.get('Texto breve de material', 'N/A'),
            "CANTIDAD":         1,
            "SERIAL":           serial,
            "Lote":             envio.get('LOTE', 'VALORADO'),
            "Estado":           "FUNCIONAL",
            "Estado_Actual":    "PENDIENTE DE INGRESO",
            "Estatus":          "RESERVADO",
            "Fecha de ingreso": envio.get('Fecha', pd.NaT)
        })

    return pd.DataFrame(filas)


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