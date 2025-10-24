    from flask import Flask, render_template, request, jsonify
    import pandas as pd
    import os

    # === CONFIGURACIÓN DEL SERVIDOR ===
    app = Flask(__name__, template_folder='frontend')

    # === CARGAR ARCHIVOS DE EXCEL ===
    # Usamos rutas absolutas seguras (Render las necesita)
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    ruta = os.path.join(BASE_DIR, 'datos', 'SISTEM.xlsx')

    # Leer las hojas del Excel
    doc_entregas = pd.read_excel(ruta, sheet_name='ENTREGAS')
    doc_devoluciones = pd.read_excel(ruta, sheet_name='DEVOLUCIONES')
    doc_salidas = pd.read_excel(ruta, sheet_name='SALIDAS')
    doc_entradas = pd.read_excel(ruta, sheet_name="ENTRADAS")
    doc_envios = pd.read_excel(ruta, sheet_name="ENVIOS")

    # === LIMPIEZA DE DATOS ===
    for df in [doc_entregas, doc_devoluciones, doc_salidas, doc_entradas]:
        if "Serial" in df.columns:
            df["Serial"] = df["Serial"].astype(str).str.strip()
    doc_envios["NºSerieFab"] = doc_envios["NºSerieFab"].astype(str).str.strip()

    for df in [doc_entregas, doc_devoluciones, doc_salidas, doc_entradas]:
        for col in df.columns:
            if "Fecha" in col or "fecha" in col:
                df[col] = pd.to_datetime(df[col], errors="coerce")


    # === RUTA PRINCIPAL ===
    @app.route('/')
    def index():
        return render_template('index.html')


    # === BÚSQUEDA DE OT ===
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
            entrega = doc_entregas[doc_entregas["Serial"] == serial]
            devolucion = doc_devoluciones[doc_devoluciones["Serial"] == serial]
            salida = doc_salidas[doc_salidas["Serial"] == serial]
            entrada = doc_entradas[doc_entradas["Serial"] == serial]

            movimientos = []
            detalle_info = []

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

                            if tipo == "Entrega":
                                detalle_item["cedula"] = fila.get("Cedula", "N/A")
                                detalle_item["tecnico"] = fila.get("Técnico", "N/A")
                                detalle_item["observaciones"] = fila.get("Observaciones", "N/A")

                            if tipo == "Salida":
                                detalle_item["observaciones"] = fila.get("Observación", "N/A")
                                detalle_item["consecutivo"] = fila.get("Consecutivo Contratista", "N/A")

                            detalle_info.append(detalle_item)

            if not movimientos:
                resultados.append({
                    "serial": f"cantidad: {cantidad_envio[casca]}",
                    "tipo": "Sin movimientos",
                    "fecha": "-",
                    "estado": "⚠️ No hay registros",
                    "SAP": sap_envio[casca],
                    "descrip": descrip_envio[casca],
                    "detalle": []
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

                # 🔹 Mostrar solo el último movimiento en detalle
                ultimo_detalle = sorted(detalle_info, key=lambda x: x["fecha"], reverse=True)[:1]

                resultados.append({
                    "serial": serial,
                    "tipo": tipo,
                    "fecha": fecha,
                    "estado": estado,
                    "SAP": sap,
                    "descrip": descrip,
                    "detalle": ultimo_detalle
                })

            casca += 1

        return render_template('index.html', resultados=resultados)


    # === EJECUCIÓN (para entorno local) ===
    if __name__ == '__main__':
        port = int(os.environ.get("PORT", 5000))  # usa el puerto que Render le pasa
        app.run(host='0.0.0.0', port=port, debug=True)