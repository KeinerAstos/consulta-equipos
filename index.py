from flask import Flask, render_template, request, jsonify
import pandas as pd

app = Flask(__name__, template_folder='frontend')

# Cargar los archivos una sola vez (al iniciar la aplicación)
ruta = 'datos/SISTEM.xlsx'
doc_entregas = pd.read_excel(ruta, sheet_name='ENTREGAS')
doc_devoluciones = pd.read_excel(ruta, sheet_name='DEVOLUCIONES')
doc_salidas = pd.read_excel(ruta, sheet_name='SALIDAS')
doc_entradas = pd.read_excel(ruta, sheet_name="ENTRADAS")
doc_envios = pd.read_excel(ruta, sheet_name="ENVIOS")

# Limpieza y conversión
for df in [doc_entregas, doc_devoluciones, doc_salidas, doc_entradas]:
    if "Serial" in df.columns:
        df["Serial"] = df["Serial"].astype(str).str.strip()
doc_envios["NºSerieFab"] = doc_envios["NºSerieFab"].astype(str).str.strip()

for df in [doc_entregas, doc_devoluciones, doc_salidas, doc_entradas]:
    for col in df.columns:
        if "Fecha" in col or "fecha" in col:
            df[col] = pd.to_datetime(df[col], errors="coerce")


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/buscar', methods=['POST'])
def buscar():
    try:
        OT = int(request.form['ot'])
    except:
        return render_template('index.html', resultado="⚠️ Ingresa un número de OT válido")

    entrega_envio = doc_envios[doc_envios["OTP"] == OT]
    if entrega_envio.empty:
        return render_template('index.html', resultado=f"⚠️ No se encontraron registros con OT {OT}")

    variable = entrega_envio['NºSerieFab'].tolist()
    sap_envio = entrega_envio["Material"].tolist()
    descrip_envio = entrega_envio["Texto breve de material"].tolist()
    cantidad_envio = entrega_envio["Ctd.en UM entrada"].tolist()

    resultados = []
    casca = 0

    for serial in variable:
        print(f"\n📦 Procesando serial: {serial}")

        entrega = doc_entregas[doc_entregas["Serial"] == serial]
        devolucion = doc_devoluciones[doc_devoluciones["Serial"] == serial]
        salida = doc_salidas[doc_salidas["Serial"] == serial]
        entrada = doc_entradas[doc_entradas["Serial"] == serial]

        movimientos = []
        keiner_pri = []  # guardará toda la info adicional

        for df, tipo, fecha_col, col_sap, col_descrip in [
            (entrega, "Entrega", "Fecha Sistema", "Codigo SAP", "Descripción SAP"),
            (entrada, "Entrada", "Fecha Ingreso", "Codigo SAP", "Descripción"),
            (devolucion, "Devolución", "FECHA SISTEMA.", "Codigo SAP", "Descripción"),
            (salida, "Salida", "Fecha Salida", "Codigo SAP", "Descripción")
        ]:
            if fecha_col in df.columns and not df.empty:
                for _, fila in df.iterrows():
                    fecha = fila.get(fecha_col)
                    sap = fila.get(col_sap)
                    descrip = fila.get(col_descrip)
                    if pd.notna(fecha):
                        movimientos.append((tipo, fecha, sap, descrip))

                        # Información general del movimiento
                        detalle_item = {
                            "tipo": tipo,
                            "fecha": str(fecha),
                            "sap": sap,
                            "descripcion": descrip,
                            "cedula": "N/A",
                            "tecnico": "N/A",
                            "observaciones": "N/A",
                            "consecutivo": "N/A"
                        }

                        # Si es una entrega, trae cédula, técnico y observaciones
                        if tipo == "Entrega":
                            detalle_item["cedula"] = fila.get("Cedula", "N/A")
                            detalle_item["tecnico"] = fila.get("Técnico", "N/A")
                            detalle_item["observaciones"] = fila.get("Observaciones", "N/A")

                        # Si es una salida, trae observaciones y consecutivo contratista
                        if tipo == "Salida":
                            detalle_item["observaciones"] = fila.get("Observación", "N/A")
                            detalle_item["consecutivo"] = fila.get("Consecutivo Contratista", "N/A")

                        keiner_pri.append(detalle_item)



        if not movimientos:
            resultados.append({
                "serial": f"cantidad: {cantidad_envio[casca]}",
                "tipo": "Sin movimientos",
                "fecha": "-",
                "estado": "⚠️ No hay registros",
                "SAP": sap_envio[casca],
                "descrip": descrip_envio[casca],
                "detalle": []  # vacía para ver más
            })
        else:
            movimientos_df = pd.DataFrame(movimientos, columns=["Tipo", "Fecha", "SAP", "descrip"])
            ultimo = movimientos_df.loc[movimientos_df["Fecha"].idxmax()]
            tipo = ultimo["Tipo"]
            fecha = ultimo["Fecha"].strftime("%Y-%m-%d")
            sap = ultimo["SAP"]
            descrip = ultimo["descrip"]

            estado = "📦 ENTREGADO" if tipo == "Entrega" else \
                     "📦 SALIDA" if tipo == "Salida" else "🏠 DISPONIBLE"

            resultados.append({
                "serial": serial,
                "tipo": tipo,
                "fecha": fecha,
                "estado": estado,
                "SAP": sap,
                "descrip": descrip,
                "detalle": keiner_pri  # para el botón "Ver más"
            })

        casca += 1

    return render_template('index.html', resultados=resultados)


# Ruta para obtener detalles (AJAX)
@app.route('/detalle/<serial>', methods=['GET'])
def detalle(serial):
    # Aquí podrías buscar en resultados globales, o en el Excel, si necesitas hacerlo dinámico
    return jsonify({"mensaje": f"Detalles de {serial} (por implementar)"})


if __name__ == '__main__':
    app.run(debug=True)
