from flask import Blueprint, render_template, request, jsonify, redirect, url_for
import openpyxl
import pandas as pd

from services.tablero_service import generar_tablero
from services.reasignacion_service import aplicar_reasignaciones
from services.excel_service import cargar_sistem
from config import ruta_re

api = Blueprint("api", __name__)

resultado_cache = None

def obtener_resultado():
    global resultado_cache
    if resultado_cache is None:
        # 1. Lectura rápida
        docs_cache = cargar_sistem() 
        
        # 2. Generar tablero
        tabla = generar_tablero(*docs_cache)
        
        # 3. IMPORTANTE: Asegurar que las columnas existan y estén en formato string
        for col in ["Codigo_Sap", "SERIAL", "Estado_Actual", "Estatus"]:
            if col in tabla.columns:
                tabla[col] = tabla[col].astype(str).str.strip()
        
        # 4. Aplicar reasignaciones
        tabla = aplicar_reasignaciones(tabla)
        
        # 5. Volver a asegurar formato string después de reasignaciones
        for col in ["Codigo_Sap", "SERIAL", "Estado_Actual", "Estatus"]:
            if col in tabla.columns:
                tabla[col] = tabla[col].astype(str).str.strip()
        
        # 6. Resetear índice para búsquedas consistentes
        tabla = tabla.reset_index(drop=True)
        
        resultado_cache = tabla
            
    return resultado_cache

def refrescar_cache():
    global resultado_cache
    resultado_cache = None
    return obtener_resultado()

# Estados disponibles:
# CLAVE   = valor real en la columna Estado_Actual del tablero (viene de MAPA_ESTADO en generar_tablero)
# VALOR   = texto descriptivo para mostrar al usuario
ESTADOS_DISPONIBLES = {
    "DISPONIBLE (BODEGA)":      "DISPONIBLE (BODEGA)",
    "DISPONIBLE (DEVOLUCION)":  "DISPONIBLE (DEVOLUCION)",
}

# ==========================================================
# SERIALes por SAP (solo disponibles)
# ==========================================================
@api.route("/buscar_seriales_sap", methods=["POST"])
def buscar_seriales_sap():
    resultado = obtener_resultado()
    sap = str(request.json.get("sap", "")).strip()
    
    print(f"Buscando seriales para SAP: '{sap}'")
    
    if not sap or resultado is None or resultado.empty:
        return jsonify({"seriales": []})
    
    # Debug: ver qué valores hay en Estado_Actual
    print(f"Estados únicos en datos: {resultado['Estado_Actual'].unique().tolist()}")
    print(f"SAPs únicos (primeros 5): {resultado['Codigo_Sap'].unique()[:5].tolist()}")
    
    # FIX: .values() en lugar de .keys() — Estado_Actual contiene los valores mapeados
    # ej: "DISPONIBLE (BODEGA)", no "ENTRADA"
    mask_sap    = resultado["Codigo_Sap"].str.strip() == sap
    mask_estado = resultado["Estado_Actual"].str.strip().isin(ESTADOS_DISPONIBLES.values())
    mask = mask_sap & mask_estado
    
    filtrado = resultado.loc[mask]
    
    print(f"Encontrados {len(filtrado)} registros para SAP {sap}")
    
    if filtrado.empty:
        return jsonify({"seriales": []})
    
    seriales = filtrado["SERIAL"].dropna().str.strip().unique().tolist()
    print(f"Seriales encontrados: {seriales}")
    
    return jsonify({"seriales": seriales[:50]})

# ==========================================================
# Buscar serial exacto (solo disponible)
# ==========================================================
@api.route("/buscar_serial", methods=["POST"])
def buscar_serial():
    resultado = obtener_resultado()
    serial = str(request.json.get("serial", "")).strip()

    if not serial:
        return jsonify({
            "existe": False,
            "error": "Falta serial"
        })

    # Buscar serial
    mask = resultado["SERIAL"].astype(str).str.strip() == serial
    fila = resultado[mask]

    if fila.empty:
        return jsonify({"existe": False, "error": "Serial no encontrado en base de datos"})

    fila = fila.iloc[0]

    estado_actual = str(fila.get("Estado_Actual", "")).strip()

    # FIX: .values() en lugar de comparar con .keys()
    if estado_actual not in ESTADOS_DISPONIBLES.values():
        return jsonify({
            "existe": False,
            "error": f"Equipo no disponible (Estado: {estado_actual})",
            "estado_actual": estado_actual
        })

    return jsonify({
        "existe": True,
        "sap": str(fila.get("Codigo_Sap", "")).strip(),
        "otp_actual": str(fila.get("OTP", "")).strip(),
        "cliente": str(fila.get("CLIENTE", "")).strip(),
        "estatus": str(fila.get("Estatus", "")).strip(),
        "estado_actual": estado_actual,
        "estado_texto": ESTADOS_DISPONIBLES.get(estado_actual, "DESCONOCIDO")
    })

@api.route("/buscar_sap", methods=["POST"])
def buscar_sap():
    resultado = obtener_resultado()
    texto = str(request.json.get("sap", "")).strip()

    if not texto or resultado is None or resultado.empty:
        return jsonify({"saps": []})

    # FIX: .values() en lugar de .keys()
    mask_disponible = resultado["Estado_Actual"].str.strip().isin(ESTADOS_DISPONIBLES.values())
    mask_texto      = resultado["Codigo_Sap"].astype(str).str.contains(texto, case=False, na=False)
    mask = mask_disponible & mask_texto
    
    if not mask.any():
        # Si no hay disponibles, buscar en todos (para debug)
        mask = resultado["Codigo_Sap"].astype(str).str.contains(texto, case=False, na=False)
    
    saps = resultado.loc[mask, "Codigo_Sap"].unique().tolist()
    
    # Para cada SAP, incluir cuántos seriales disponibles tiene
    saps_con_info = []
    for sap in saps[:10]:
        count = len(resultado[
            (resultado["Codigo_Sap"] == sap) &
            # FIX: .values() en lugar de .keys()
            (resultado["Estado_Actual"].isin(ESTADOS_DISPONIBLES.values()))
        ])
        saps_con_info.append({
            "sap": sap,
            "disponibles": count
        })
    
    return jsonify({"saps": saps_con_info})

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

    if not seriales:
        return jsonify({"error": "No se proporcionaron seriales"}), 400

    workbook = openpyxl.load_workbook(ruta_re)
    sheet    = workbook.active

    for i in range(len(seriales)):
        if seriales[i].strip():  # Solo guardar si hay serial
            nuevo_id   = sheet.max_row + 1
            nueva_fila = [
                nuevo_id, bodega, centro, almacen,
                saps[i] if i < len(saps) else "",
                seriales[i],
                otp_actuales[i] if i < len(otp_actuales) else "",
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

@api.route("/debug_cache")
def debug_cache():
    resultado = obtener_resultado()
    return jsonify({
        "total_registros": len(resultado),
        "columnas": resultado.columns.tolist(),
        "ejemplo_saps": resultado["Codigo_Sap"].unique()[:10].tolist(),
        "ejemplo_seriales": resultado["SERIAL"].head(10).tolist(),
        "estados_unicos": resultado["Estado_Actual"].unique().tolist(),
        "estados_disponibles_count": len(resultado[resultado["Estado_Actual"].isin(ESTADOS_DISPONIBLES.values())])
    })


@api.route("/debug_saps_disponibles")
def debug_saps_disponibles():
    resultado = obtener_resultado()
    # FIX: .values() en lugar de .keys()
    disponibles = resultado[resultado["Estado_Actual"].isin(ESTADOS_DISPONIBLES.values())]
    
    resumen = disponibles.groupby("Codigo_Sap").agg({
        "SERIAL": "count",
        "Estado_Actual": lambda x: x.mode().iloc[0] if not x.empty else "N/A"
    }).rename(columns={"SERIAL": "cantidad"})
    
    return jsonify({
        "total_saps": len(resumen),
        "total_equipos": len(disponibles),
        "resumen": resumen.head(20).to_dict(orient="index")
    })

@api.route("/")
def home():
    return render_template("reasignacion.html")