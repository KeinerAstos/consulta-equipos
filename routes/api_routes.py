from flask import Blueprint, render_template, request, jsonify, redirect, url_for
import openpyxl

from services.tablero_service import generar_tablero
from services.reasignacion_service import aplicar_reasignaciones
from services.excel_service import cargar_sistem
from config import ruta_re
api = Blueprint("api", __name__)
resultado_cache = None



def obtener_resultado():
    global resultado_cache
    if resultado_cache is None:
        docs_cache = cargar_sistem()
        tabla = generar_tablero(*docs_cache)
        tabla = aplicar_reasignaciones(tabla)
        resultado_cache = tabla
    return resultado_cache


def refrescar_cache():
    global resultado_cache

    docs_cache = cargar_sistem()
    tabla = generar_tablero(*docs_cache)
    tabla = aplicar_reasignaciones(tabla)

    resultado_cache = tabla



@api.route("/buscar_seriales_sap", methods=["POST"])
def buscar_seriales_sap():

    resultado = obtener_resultado()
    sap = str(request.json.get("sap", "")).strip()

    resultado["Codigo_Sap"] = resultado["Codigo_Sap"].astype(str).str.strip()
    resultado["SERIAL"]     = resultado["SERIAL"].astype(str).str.strip()

    filtrado = resultado[resultado["Codigo_Sap"] == sap]
    seriales = filtrado["SERIAL"].drop_duplicates().tolist()

    return jsonify({"seriales": seriales[:20]})


@api.route("/buscar_serial", methods=["POST"])
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


@api.route("/buscar_sap", methods=["POST"])
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


@api.route("/guardar_reasignacion", methods=["POST"])
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
    return redirect(url_for("api.home"))


@api.route("/tablero_stock")
def tablero_stock():
    resultado_final = obtener_resultado()
    datos = resultado_final.to_dict(orient="records")
    return render_template("tablero_stock.html", datos=datos)



@api.route("/")
def home():
    return render_template("reasignacion.html")

