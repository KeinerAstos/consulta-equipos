from flask import Flask, render_template, request, jsonify, redirect, url_for, session
import pandas as pd
import os
import openpyxl
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ruta = os.path.join(BASE_DIR, 'datos', 'Inventario_(19).xlsx')
ruta_envio = os.path.join(BASE_DIR,'datos','ENVIO(7).xlsx')
ruta_re = os.path.join(BASE_DIR,'datos','reasignacion.xlsx')
ruta_movi = os.path.join(BASE_DIR,'datos','movimientos.xlsx')

doc_inventario = pd.read_excel(ruta, sheet_name='Inventario')
doc_aliados = pd.read_excel(ruta_envio, sheet_name="DESPACHO ALIADOS 2026")
doc_reasignado = pd.read_excel(ruta_re, sheet_name="Hoja1")
doc_movimiento = pd.read_excel(ruta_movi,sheet_name="Movimientos")


doc_aliados['NºSerieFab'] = doc_aliados['NºSerieFab'].astype(str).str.strip()
doc_inventario['Serial1'] = doc_inventario['Serial1'].astype(str).str.strip()
doc_movimiento['Serial1'] = doc_movimiento['Serial1'].astype(str).str.strip()


# Ordenar por fecha (muy importante)
doc_movimiento = doc_movimiento.sort_values(by='FechaMovimiento')

# ===============================
# LIMPIEZA GENERAL
# ===============================

# Limpiar nombres de columnas (quita espacios invisibles)
doc_movimiento.columns = doc_movimiento.columns.str.strip()

doc_movimiento['Serial1'] = (
    doc_movimiento['Serial1']
    .replace(['nan', 'NaN', '', '#N/D', '#N/A'], 'N/A')
)

doc_inventario['Serial1'] = (
    doc_inventario['Serial1']
    .replace(['nan', 'NaN', '', '#N/D', '#N/A'], 'N/A')
)


# Normalizar tipo movimiento
doc_movimiento['TipoMovimiento'] = (
    doc_movimiento['TipoMovimiento']
    .astype(str)
    .str.strip()
    .str.upper()
)

# Ordenar por fecha
doc_movimiento = doc_movimiento.sort_values(by='FechaMovimiento')


# ===============================
# FILTRAR SOLO INGRESO Y SALIDA
# ===============================

movimientos_validos = doc_movimiento[
    doc_movimiento['TipoMovimiento'].isin(['INGRESO', 'SALIDA'])
].copy()

# Crear signo matemático
movimientos_validos['Signo'] = movimientos_validos['TipoMovimiento'].map({
    'INGRESO': 1,
    'SALIDA': -1
})


# ===============================
# CALCULAR STOCK REAL
# ===============================

stock_actual = (
    movimientos_validos
    .groupby(['NumeroParte', 'NombreParte', 'Serial1'], as_index=False)
    ['Signo']
    .sum()
)

# Solo lo que realmente existe en bodega
en_bodega = stock_actual[stock_actual['Signo'] > 0].copy()


# ===============================
# OBTENER ULTIMA FECHA REAL
# ===============================

ultimas_fechas = (
    movimientos_validos
    .groupby(['NumeroParte', 'NombreParte', 'Serial1'], as_index=False)
    ['FechaMovimiento']
    .max()
)

# Unir fecha al stock actual
en_bodega = en_bodega.merge(
    ultimas_fechas,
    on=['NumeroParte', 'NombreParte', 'Serial1'],
    how='left'
)


# ===============================
# ARMAR RESULTADO FINAL
# ===============================

resultado = en_bodega.rename(columns={
    'NumeroParte': 'SAP',
    'NombreParte': 'DESCRIPCION',
    'Serial1': 'SERIAL',
    'Signo': 'CANTIDAD',
    'FechaMovimiento': 'FECHA'
})

# Selección final (ahora sí existe FECHA)
resultado = resultado[['SERIAL', 'SAP', 'DESCRIPCION', 'CANTIDAD', 'FECHA']]

OTP = []
OTH = []
ESTADO_OTP_CRM = [] 
ESTADO_OTH_CRM = []
Centro = []
ALMACEN = []
Aliado = []
Nombre_Almacén = []
CLIENTE = []
Tipo_de_OT = []
Asignado = []
Codigo_Sap = []
Descripción = []
CANTIDAD = []
SERIAL = []
lote = []
Estado = []
Estatus = []
Fecha_cambio_de_estatus = []
Fecha_de_reporte = []
Fecha_de_ingreso = []
Días_en_Reserva = []

numero = 0

doc_aliados["Fecha"] = pd.to_datetime(
    doc_aliados["Fecha"],
    errors="coerce"
).dt.strftime("%d/%m/%Y")

doc_inventario['FechaOrden'] = pd.to_datetime(
    doc_inventario['FechaOrden'],
    errors="coerce"
).dt.strftime("%d/%m/%Y")


for i in range(len(resultado['SERIAL'])):
    serial = resultado['SERIAL'].iloc[i]

    if serial in doc_aliados['NºSerieFab'].values:
        
        for x in range(len(doc_aliados['NºSerieFab'])):
            if serial == doc_aliados['NºSerieFab'].iloc[x]:
                OTP.append(doc_aliados['OTP'].iloc[x])
                OTH.append(doc_aliados['OTH'].iloc[x])
                Centro.append(doc_aliados['COD CENTRO'].iloc[x])
                ALMACEN.append(doc_aliados['COD ALM'].iloc[x])
                Aliado.append(doc_aliados['Destino'].iloc[x])
                CLIENTE.append(doc_aliados['CLIENTE'].iloc[x])
                valor_prc = str(doc_aliados['PRC/SOLPED'].iloc[x]).strip()
                if valor_prc.upper().startswith("PRC"):
                    Tipo_de_OT.append(valor_prc)
                else:
                    Tipo_de_OT.append("INSTALACIONES")
                Codigo_Sap.append(doc_aliados['Material'].iloc[x])
                Descripción.append(doc_aliados['Texto breve de material'].iloc[x])
                CANTIDAD.append('1')
                SERIAL.append(serial)
                lote.append(doc_aliados['LOTE'].iloc[x])
                Estado.append('FUNCIONAL')
                Estatus.append('RESERVADO')
                Fecha_de_ingreso.append(doc_aliados['Fecha'].iloc[x])
                Fecha_cambio_de_estatus.append(doc_aliados['Fecha'].iloc[x])


                numero +=1
                break   # 🔥 importante para que no duplique
                
    else:
        OTP.append('N/A')
        OTH.append('N/A')
        Centro.append('C903')
        ALMACEN.append('A500')
        Aliado.append('ALGARTECH')
        CLIENTE.append('N/A') 
        Tipo_de_OT.append('BASE')
        Codigo_Sap.append(resultado['SAP'].iloc[i])
        Descripción.append(resultado['DESCRIPCION'].iloc[i])
        CANTIDAD.append(resultado['CANTIDAD'].iloc[i])
        SERIAL.append(serial)
        lote.append('VALORADO')
        Estado.append('FUNCIONAL')
        Estatus.append('STOCK')
        Fecha_de_ingreso.append(resultado['FECHA'].iloc[i])
        Fecha_cambio_de_estatus.append(resultado['FECHA'].iloc[i])

    ESTADO_OTP_CRM.append('N/A')
    ESTADO_OTH_CRM.append('N/A')
    Asignado.append('N/A')
    Fecha_de_reporte.append('17/02/2026')

resultado_final = pd.DataFrame({
    "OTP": OTP,
    "OTH": OTH,
    "ESTADO_OTP_CRM": ESTADO_OTP_CRM,
    "ESTADO_OTH_CRM": ESTADO_OTH_CRM,
    "Centro": Centro,
    "ALMACEN": ALMACEN,
    "Aliado": Aliado,
    "CLIENTE": CLIENTE,
    "Tipo_de_OT": Tipo_de_OT,
    "Asignado": Asignado,
    "Codigo_Sap": Codigo_Sap,
    "Descripción": Descripción,
    "CANTIDAD": CANTIDAD,
    "SERIAL": SERIAL,
    "Lote": lote,
    "Estado": Estado,
    "Estatus": Estatus,
    "Fecha cambio de estatus": Fecha_cambio_de_estatus,
    "Fecha de reporte": Fecha_de_reporte,
    "Fecha de ingreso": Fecha_de_ingreso
})

doc_reasignado['FECHA_CAMBIO'] = pd.to_datetime(
    doc_reasignado['FECHA_CAMBIO'],
    errors="coerce"
).dt.strftime("%d/%m/%Y")

resultado_final = resultado_final.merge(
    doc_reasignado,
    on='SERIAL',
    how='left'
)

# Sobrescribir si está reasignado
mask = resultado_final['OTP NUEVA'].notna()

# Normalizar columnas clave
resultado_final["SERIAL"] = resultado_final["SERIAL"].astype(str).str.strip()
resultado_final["Codigo_Sap"] = resultado_final["Codigo_Sap"].astype(str).str.strip()
resultado_final["Estatus"] = resultado_final["Estatus"].astype(str).str.strip()
resultado_final.loc[mask, 'OTP'] = resultado_final['OTP NUEVA']
resultado_final.loc[mask, 'CLIENTE'] = resultado_final['NUEVO_CLIENTE']
resultado_final.loc[mask, 'Estatus'] = 'REASIGNADO'
resultado_final.loc[mask, 'ALMACEN'] = 'Q500'
resultado_final.loc[mask, 'Tipo_de_OT'] = 'INSTALACIONES'

resultado_final.to_excel("resultado_final.xlsx", index=False)


app = Flask(__name__, template_folder='frontend')


@app.route("/buscar_seriales_sap", methods=["POST"])
def buscar_seriales_sap():

    sap = str(request.json.get("sap")).strip()

    filtrado = resultado_final[
        (resultado_final["Codigo_Sap"].astype(str).str.strip() == sap) &
        (resultado_final["Estatus"] == "STOCK")
    ]

    seriales = filtrado["SERIAL"].head(10).tolist()

    return jsonify({
        "existe": len(seriales) > 0,
        "seriales": seriales
    })


@app.route("/buscar_serial", methods=["POST"])
def buscar_serial():

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
    
@app.route("/guardar_reasignacion", methods=["POST"])
def guardar_reasignacion():

    # 🔹 Campos generales
    bodega = request.form.get("bodega")
    centro = request.form.get("centro")
    almacen = request.form.get("almacen")
    otp_nueva = request.form.get("otp_nueva")
    oth_nueva = request.form.get("oth_nueva")
    nuevo_cliente = request.form.get("nuevo_cliente")
    tipo_ot = request.form.get("tipo_ot")
    fecha_cambio = request.form.get("fecha_cambio")

    # 🔹 Campos tipo lista (tabla)
    saps = request.form.getlist("sap[]")
    seriales = request.form.getlist("serial[]")
    otp_actuales = request.form.getlist("otp_actual[]")

    print("Bodega:", bodega)
    print("Centro:", centro)
    print("Seriales:", seriales)

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    ruta_re = os.path.join(BASE_DIR, 'datos', 'reasignacion.xlsx')

    workbook = openpyxl.load_workbook(ruta_re)
    sheet = workbook.active

    # 🔥 Insertar una fila por cada equipo
    for i in range(len(seriales)):

        nuevo_id = sheet.max_row  # ID automático

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

    return redirect(url_for("home"))

@app.route("/")
def home():
    return render_template("reasignacion.html")

if __name__ == "__main__":
    app.run(debug=True)