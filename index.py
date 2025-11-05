from flask import Flask, render_template, request, jsonify, redirect, url_for, session
import pandas as pd
import os
import psycopg2
from psycopg2 import errors

# === CONFIGURACIÓN GENERAL ===
app = Flask(__name__, template_folder='frontend')
app.secret_key = "clave_super_secreta"  # Necesaria para manejar sesiones

# === CONEXIÓN A LA BASE DE DATOS (Neon PostgreSQL) ===
def get_connection():
    return psycopg2.connect(
        dbname="neondb",
        user="neondb_owner",
        password="npg_B0ZyzNDGFb3k",
        host="ep-cool-snow-ad0pqcmu-pooler.c-2.us-east-1.aws.neon.tech",
        port="5432",
        sslmode="require"
    )

# === CARGA DE ARCHIVOS DE EXCEL ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ruta = os.path.join(BASE_DIR, 'datos', 'SISTEM.xlsx')

doc_entregas = pd.read_excel(ruta, sheet_name='ENTREGAS')
doc_devoluciones = pd.read_excel(ruta, sheet_name='DEVOLUCIONES')
doc_salidas = pd.read_excel(ruta, sheet_name='SALIDAS')
doc_entradas = pd.read_excel(ruta, sheet_name="ENTRADAS")
doc_stock = pd.read_excel(ruta, sheet_name="STOCK")
doc_envios = pd.read_excel(ruta, sheet_name="ENVIOS")

# Limpieza básica
for df in [doc_entregas, doc_devoluciones, doc_salidas, doc_entradas,doc_stock]:
    if "Serial" in df.columns:
        df["Serial"] = df["Serial"].astype(str).str.strip()
doc_envios["NºSerieFab"] = doc_envios["NºSerieFab"].astype(str).str.strip()

for df in [doc_entregas, doc_devoluciones, doc_salidas, doc_entradas]:
    for col in df.columns:
        if "Fecha" in col or "fecha" in col:
            df[col] = pd.to_datetime(df[col], errors="coerce")


@app.route('/table')
def tabla_usuario():
    conn = get_connection()
    cur = conn.cursor()

    # Consulta: obtener usuarios y sus seriales asignados
    cur.execute("""
        SELECT u.nombre as nombre, a.serial_f as serial, e.sap as sap, e.descripcion as descripcion
        FROM asignacion a
        JOIN usuario u ON a.cedula_f = u.cedula
        JOIN equipo e ON e.serial = a.serial_f
        ORDER BY u.nombre;
    """)
    resultados = cur.fetchall()

    cur.close()
    conn.close()

    # Agrupar seriales por usuario
    data = {}
    for nombre, serial, sap, descripcion in resultados:
        if nombre not in data:
            data[nombre] = []
        if serial is not None:
            data[nombre].append({'serial': serial, 'sap': sap, 'descripcion': descripcion})

    # Pasamos la estructura al template
    return render_template('table.html', data=data)


@app.route('/buscar_usuarios', methods=['GET'])
def buscar_usuarios():
    nombre = request.args.get('nombre', '').lower()
    if not nombre:
        return jsonify([])

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT nombre, cedula 
        FROM usuario
        WHERE LOWER(nombre) LIKE %s
        LIMIT 10
    """, (f"%{nombre}%",))
    
    resultados = cur.fetchall()
    cur.close()
    conn.close()

    usuarios = [{'nombre': r[0], 'cedula': r[1]} for r in resultados]
    return jsonify(usuarios)



# === LOGIN ===
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        correo = request.form['correo']
        contrasena = request.form['contrasena']

        conn = get_connection()
        cur = conn.cursor()
        cur.execute("SELECT id, nombre, contrasena, rol FROM usuarios WHERE correo = %s", (correo,))
        usuario = cur.fetchone()
        cur.close()
        conn.close()

        # Verificación de credenciales
        if usuario and usuario[2] == contrasena:
            session['usuario_id'] = usuario[0]
            session['usuario_nombre'] = usuario[1]
            session['usuario_rol'] = usuario[3]
            return redirect(url_for('dashboard'))
        else:
            return render_template('login.html', error="❌ Credenciales incorrectas")

    return render_template('login.html')

# === LOGOUT ===
@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

# === DASHBOARD ===
@app.route('/dashboard')
def dashboard():
    if 'usuario_id' not in session:
        return redirect(url_for('login'))

    return render_template('dashboard.html', 
                           nombre=session['usuario_nombre'], 
                           rol=session['usuario_rol'])


@app.route('/insertar_asignacion', methods=['POST'])
def insertar_asignacion():
    serial = request.form['serial']
    cedula = request.form['cedula']
    fecha = request.form['fecha']

    # Confirmar que el serial exista en doc_stock
    confirmacion = doc_stock['Serial'].tolist()
    if serial in confirmacion:
        dato = doc_stock[doc_stock["Serial"] == serial].iloc[0]  # Tomamos la primera fila
        sap = dato["Codigo material"]
        descripcion = dato["Descripción"]

        try:
            conn = get_connection()
            cur = conn.cursor()

            # INSERT en tabla equipo
            cur.execute(
                "INSERT INTO equipo (serial, sap, descripcion) VALUES (%s, %s, %s)",
                (serial, sap, descripcion)
            )

            # INSERT en tabla asignacion
            cur.execute(
                "INSERT INTO asignacion (serial_f, cedula_f, fecha_asig) VALUES (%s, %s, %s)",
                (serial, cedula, fecha)
            )

            conn.commit()
            cur.close()
            conn.close()

            return jsonify({'status': 'ok', 'message': '✅ Asignación registrada correctamente.'})

        except errors.UniqueViolation:
            conn.rollback()
            return jsonify({'status': 'error', 'message': '⚠️ Este equipo ya está asignado.'}), 400

        except psycopg2.Error as e:
            conn.rollback()
            return jsonify({'status': 'error', 'message': f'❌ Error en la base de datos: {e.pgerror}'}), 500

        finally:
            if conn:
                conn.close()
    else:
        return jsonify({'status': 'error', 'message': '⚠️ El serial no existe en siscos.'}), 400


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