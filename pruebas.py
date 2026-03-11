from flask import Flask, render_template, request, jsonify, redirect, url_for, session
import pandas as pd
import os
import psycopg2
import openpyxl
from collections import defaultdict
from psycopg2 import errors

app = Flask(__name__, template_folder='frontend')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ruta = os.path.join(BASE_DIR, 'datos', 'Inventario_(19).xlsx')
ruta_envio = os.path.join(BASE_DIR,'datos','ENVIO(7).xlsx')
ruta_re = os.path.join(BASE_DIR,'datos','reasignacion.xlsx')
ruta_movi = os.path.join(BASE_DIR,'datos','movimientos.xlsx')

resultado_cache = None
doc_inventario = pd.read_excel(ruta, sheet_name='Inventario')
doc_aliados = pd.read_excel(ruta_envio, sheet_name="DESPACHO ALIADOS 2026")
doc_movimiento = pd.read_excel(ruta_movi, sheet_name="Movimientos")

doc_aliados['NºSerieFab'] = doc_aliados['NºSerieFab'].astype(str).str.strip()
doc_inventario['Serial1'] = doc_inventario['Serial1'].astype(str).str.strip()
doc_movimiento['Serial1'] = doc_movimiento['Serial1'].astype(str).str.strip()

doc_movimiento.columns = doc_movimiento.columns.str.strip()
# ==========================================
# FUNCION PRINCIPAL QUE GENERA EL TABLERO
# ==========================================

def generar_tablero():  

    doc_movimiento['Serial1'] = (
        doc_movimiento['Serial1']
        .replace(['nan','NaN','','#N/D','#N/A'],'N/A')
    )

    doc_movimiento['TipoMovimiento'] = (
        doc_movimiento['TipoMovimiento']
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # -------------------------
    # FILTRAR MOVIMIENTOS
    # -------------------------

    movimientos_validos = doc_movimiento[
        doc_movimiento['TipoMovimiento'].isin(['INGRESO','SALIDA'])
    ].copy()

    movimientos_validos['Signo'] = movimientos_validos['TipoMovimiento'].map({
        'INGRESO':1,
        'SALIDA':-1
    })

    # -------------------------
    # CALCULAR STOCK
    # -------------------------

    stock_actual = (
        movimientos_validos
        .groupby(['NumeroParte','NombreParte','Serial1'], as_index=False)['Signo']
        .sum()
    )

    en_bodega = stock_actual[stock_actual['Signo'] > 0].copy()

    ultimas_fechas = (
        movimientos_validos
        .groupby(['NumeroParte','NombreParte','Serial1'], as_index=False)['FechaMovimiento']
        .max()
    )

    en_bodega = en_bodega.merge(
        ultimas_fechas,
        on=['NumeroParte','NombreParte','Serial1'],
        how='left'
    )

    resultado = en_bodega.rename(columns={
        'NumeroParte':'SAP',
        'NombreParte':'DESCRIPCION',
        'Serial1':'SERIAL',
        'Signo':'CANTIDAD',
        'FechaMovimiento':'FECHA'
    })

    resultado = resultado[['SERIAL','SAP','DESCRIPCION','CANTIDAD','FECHA']]

    # -----------------------------------
    # ARMAR TABLA FINAL
    # -----------------------------------

    filas = []

    for i in range(len(resultado)):

        serial = resultado['SERIAL'].iloc[i]

        if serial in doc_aliados['NºSerieFab'].values:

            fila = doc_aliados[doc_aliados['NºSerieFab'] == serial].iloc[0]

            filas.append({
                "OTP": fila['OTP'],
                "OTH": fila['OTH'],
                "Centro": fila['COD CENTRO'],
                "ALMACEN": fila['COD ALM'],
                "Aliado": fila['Destino'],
                "CLIENTE": fila['CLIENTE'],
                "Tipo_de_OT": fila['PRC/SOLPED'],
                "Asignado": "N/A",
                "Codigo_Sap": fila['Material'],
                "Descripción": fila['Texto breve de material'],
                "CANTIDAD": 1,
                "SERIAL": serial,
                "Lote": fila['LOTE'],
                "Estado": "FUNCIONAL",
                "Estatus": "RESERVADO",
                "Fecha de ingreso": fila['Fecha']
            })

        else:

            filas.append({
                "OTP": "N/A",
                "OTH": "N/A",
                "Centro": "C903",
                "ALMACEN": "A500",
                "Aliado": "ALGARTECH",
                "CLIENTE": "N/A",
                "Tipo_de_OT": "BASE",
                "Asignado": "N/A",
                "Codigo_Sap": resultado['SAP'].iloc[i],
                "Descripción": resultado['DESCRIPCION'].iloc[i],
                "CANTIDAD": resultado['CANTIDAD'].iloc[i],
                "SERIAL": serial,
                "Lote": "VALORADO",
                "Estado": "FUNCIONAL",
                "Estatus": "STOCK",
                "Fecha de ingreso": resultado['FECHA'].iloc[i]
            })

    resultado_final = pd.DataFrame(filas)

    return resultado_final


# ==========================================
# FUNCION REASIGNACIONES
# ==========================================

def aplicar_reasignaciones(tabla):

    doc_reasignado = pd.read_excel(ruta_re)

    doc_reasignado.columns = doc_reasignado.columns.str.strip()

    doc_reasignado['SERIAL'] = doc_reasignado['SERIAL'].astype(str).str.strip()

    tabla['SERIAL'] = tabla['SERIAL'].astype(str).str.strip()

    tabla = tabla.merge(
        doc_reasignado[['SERIAL','OTP_NUEVA','NUEVO_CLIENTE']],
        on='SERIAL',
        how='left'
    )

    mask = tabla['OTP_NUEVA'].notna()

    tabla.loc[mask,'OTP'] = tabla.loc[mask,'OTP_NUEVA']
    tabla.loc[mask,'CLIENTE'] = tabla.loc[mask,'NUEVO_CLIENTE']
    tabla.loc[mask,'Estatus'] = 'REASIGNADO'
    tabla.loc[mask,'ALMACEN'] = 'Q500'
    tabla.loc[mask,'Tipo_de_OT'] = 'INSTALACIONES'

    return tabla

def obtener_resultado():

    global resultado_cache

    if resultado_cache is None:

        tabla = generar_tablero()

        tabla = aplicar_reasignaciones(tabla)

        resultado_cache = tabla

    return resultado_cache

def refrescar_cache():

    global resultado_cache

    tabla = generar_tablero()
    tabla = aplicar_reasignaciones(tabla)

    resultado_cache = tabla

@app.route("/buscar_seriales_sap", methods=["POST"])
def buscar_seriales_sap():

    resultado = obtener_resultado()

    sap = str(request.json.get("sap", "")).strip()

    resultado["Codigo_Sap"] = (
        resultado["Codigo_Sap"]
        .astype(str)
        .str.strip()
    )

    resultado["SERIAL"] = (
        resultado["SERIAL"]
        .astype(str)
        .str.strip()
    )

    filtrado = resultado[
        resultado["Codigo_Sap"] == sap
    ]

    seriales = (
        filtrado["SERIAL"]
        .drop_duplicates()
        .tolist()
    )

    return jsonify({
        "seriales": seriales[:20]
    })


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
            "existe": True,
            "sap": fila["Codigo_Sap"],
            "otp_actual": fila["OTP"],
            "cliente": fila["CLIENTE"],
            "estatus": fila["Estatus"]
        })

    return jsonify({"existe": False})

@app.route("/buscar_sap", methods=["POST"])
def buscar_sap():

    resultado = obtener_resultado()

    texto = str(request.json.get("sap", "")).strip()

    if texto == "":
        return jsonify({"saps": []})

    resultado["Codigo_Sap"] = (
        resultado["Codigo_Sap"]
        .astype(str)
        .str.strip()
    )

    filtrado = resultado[
        resultado["Codigo_Sap"].str.contains(texto, case=False, na=False)
    ]

    saps = (
        filtrado["Codigo_Sap"]
        .drop_duplicates()
        .tolist()
    )

    return jsonify({"saps": saps[:20]})

@app.route("/guardar_reasignacion", methods=["POST"])
def guardar_reasignacion():

    bodega = request.form.get("bodega")
    centro = request.form.get("centro")
    almacen = request.form.get("almacen")
    otp_nueva = request.form.get("otp_nueva")
    oth_nueva = request.form.get("oth_nueva")
    nuevo_cliente = request.form.get("nuevo_cliente")
    tipo_ot = request.form.get("tipo_ot")
    fecha_cambio = request.form.get("fecha_cambio")

    saps = request.form.getlist("sap[]")
    seriales = request.form.getlist("serial[]")
    otp_actuales = request.form.getlist("otp_actual[]")

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    ruta_re = os.path.join(BASE_DIR, 'datos', 'reasignacion.xlsx')

    workbook = openpyxl.load_workbook(ruta_re)
    sheet = workbook.active

    for i in range(len(seriales)):

        nuevo_id = sheet.max_row

        nueva_fila = [
            nuevo_id,
            bodega,
            centro,
            almacen,
            saps[i],
            seriales[i],
            otp_actuales[i],
            otp_nueva,
            oth_nueva,
            nuevo_cliente,
            tipo_ot,
            fecha_cambio
        ]

        sheet.append(nueva_fila)

    workbook.save(ruta_re)
    refrescar_cache()
    return redirect(url_for("home"))



@app.route("/tablero_stock")
def tablero_stock():

    resultado_final = obtener_resultado()  # recalcula todo

    datos = resultado_final.to_dict(orient="records")

    return render_template(
        "tablero_stock.html",
        datos=datos
    )

@app.route("/")
def home():
    return render_template("reasignacion.html")

if __name__ == "__main__":
    app.run(debug=True, port=5000)