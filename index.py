from flask import Flask, render_template, request, jsonify, redirect, url_for, session
import pandas as pd
import os
import psycopg2
from collections import defaultdict
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
doc_seriales_terreno = pd.read_excel(ruta,sheet_name="TERRENO_SERIALES")

# Limpieza básica
for df in [doc_entregas, doc_devoluciones, doc_salidas, doc_entradas,doc_stock]:
    if "Serial" in df.columns:
        df["Serial"] = df["Serial"].astype(str).str.strip()
doc_envios["NºSerieFab"] = doc_envios["NºSerieFab"].astype(str).str.strip()

for df in [doc_entregas, doc_devoluciones, doc_salidas, doc_entradas]:
    for col in df.columns:
        if "Fecha" in col or "fecha" in col:
            df[col] = pd.to_datetime(df[col], errors="coerce")
# === FUNCIÓN AUXILIAR PARA CONSULTAR DATOS ===
def obtener_asignaciones():
    conn = get_connection()
    cur = conn.cursor()
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
    data = {}
    for nombre, serial, sap, descripcion in resultados:
        if nombre not in data:
            data[nombre] = []
        if serial:
            data[nombre].append({
                'serial': serial,
                'sap': sap,
                'descripcion': descripcion
            })
    return data


# === RUTAS REUTILIZANDO LA FUNCIÓN ===
@app.route('/dash_table')
def dash_table():
    data = obtener_asignaciones()
    return render_template('dash_table.html', data=data)


@app.route('/table')
def tabla_usuario():
    data = obtener_asignaciones()
    return render_template('table.html', data=data)

# === DASHBOARD ===
@app.route('/consultas_grf')
def consultar_grf():
    # Verificar sesión
    if 'usuario_id' not in session:
        return redirect(url_for('login'))
    data = obtener_asignaciones()
    # Enviar datos al template
    return render_template('consultas_grf.html',
        nombre=session.get('usuario_nombre', 'Invitado'),
        rol=session.get('usuario_rol', 'Sin rol'),
        data=data)


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

# Buscar equipos por serial
@app.route('/buscar_serial', methods=['GET'])
def buscar_serial():
    texto = request.args.get('serial', '').strip().lower()
    if not texto:
        return jsonify([])

    # Filtrar desde el Excel cargado
    resultados = doc_stock[
        doc_stock['Serial'].astype(str).str.lower().str.contains(texto)
    ].head(10)

    data = []
    for _, row in resultados.iterrows():
        data.append({
            "serial": str(row["Serial"]),
            "descripcion": str(row.get("Descripción", "Sin descripción")),
            "sap": str(row.get("Codigo SAP", "Sin SAP"))
        })

    return jsonify(data)



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
    # Verificar sesión
    if 'usuario_id' not in session:
        return redirect(url_for('login'))
    
    conn = get_connection()
    cur = conn.cursor()

    # Consulta: obtener equipos asignados por usuario
    cur.execute("""
        SELECT 
            u.nombre AS nombre, 
            a.serial_f AS serial, 
            e.sap AS sap, 
            e.descripcion AS descripcion
        FROM asignacion a
        JOIN usuario u ON a.cedula_f = u.cedula
        JOIN equipo e ON e.serial = a.serial_f
        ORDER BY u.nombre;
    """)
    
    resultados = cur.fetchall()

    cur.close()
    conn.close()

    # Agrupar equipos por usuario
    data = defaultdict(list)
    for nombre, serial, sap, descripcion in resultados:
        if serial:
            data[nombre].append({
                'nombre': nombre,
                'serial': serial,
                'sap': sap,
                'descripcion': descripcion
            })

    # Enviar datos al template
    return render_template(
        'dashboard.html',
        nombre=session.get('usuario_nombre', 'Invitado'),
        rol=session.get('usuario_rol', 'Sin rol'),
        data=data
    )
@app.route('/terreno_seriales')
def terreno_seriales():
    seriales = doc_seriales_terreno['serial'].tolist()
    sap = doc_seriales_terreno['codigo_sap'].tolist()
    descripcion = doc_seriales_terreno['descripcion_sap'].tolist()
    nombre_tecnico = doc_seriales_terreno['nombre_tecnico'].tolist()

    data = defaultdict(list)

    for i in range(len(seriales)):
        tecnico = nombre_tecnico[i]
        data[tecnico].append({
            'nombre': tecnico,
            'serial': seriales[i],
            'sap': sap[i],
            'descripcion': descripcion[i]
        })

    return render_template(
        "table_pim.html",  # ← tu archivo HTML
        data=data
    )



@app.route('/eliminar_asignacion/<serial>', methods=['DELETE'])
def eliminar_asignacion(serial):
    """
    Mueve una asignación de 'asignacion' a 'historial_movimiento' 
    registrando quién la eliminó y la fecha exacta.
    """
    try:
        # Verificar sesión activa
        if 'usuario_nombre' not in session:
            return jsonify({'status': 'error', 'message': 'Sesión expirada. Inicie sesión nuevamente.'}), 401

        usuario_actual = session['usuario_nombre']

        conn = get_connection()
        cur = conn.cursor()

        # Buscar la asignación
        cur.execute("""
            SELECT a.serial_f, a.cedula_f, u.nombre, e.sap, e.descripcion
            FROM asignacion a
            JOIN usuario u ON a.cedula_f = u.cedula
            JOIN equipo e ON e.serial = a.serial_f
            WHERE a.serial_f = %s;
        """, (serial,))
        asignacion = cur.fetchone()

        if not asignacion:
            cur.close()
            conn.close()
            return jsonify({'status': 'error', 'message': f'El serial {serial} no está asignado.'}), 404

        serial_f, cedula_f, nombre_usuario, sap, descripcion = asignacion

        # Insertar en historial_movimiento
        cur.execute("""
            INSERT INTO historial_movimiento (serial, cedula_usuario, nombre_usuario, sap, descripcion, movimiento_realizado_por, fecha_movimiento)
            VALUES (%s, %s, %s, %s, %s, %s, NOW());
        """, (serial_f, cedula_f, nombre_usuario, sap, descripcion, usuario_actual))

        # Eliminar de asignacion
        cur.execute("DELETE FROM asignacion WHERE serial_f = %s;", (serial_f,))
        conn.commit()

        cur.close()
        conn.close()

        return jsonify({'status': 'ok', 'message': f'Asignación del equipo {serial} movida al historial correctamente.'})

    except Exception as e:
        print("❌ Error al mover asignación al historial:", e)
        return jsonify({'status': 'error', 'message': 'Error al registrar el movimiento.'}), 500


from flask import Flask, request, send_file, jsonify
import grf # Asegúrate de importar tu módulo

# ... (resto de la configuración de Flask)

@app.route('/usuario_grf', methods=['POST'])
def usuario_grf():
    import sys
    
    usuario = request.form['usuario']
    contra = request.form['contraseña']
    archivo = request.files['archivo']
    
    # Leer el archivo para verificar tamaño y dar feedback
    try:
        df_temp = pd.read_excel(archivo)
        num_seriales = len(df_temp)
        
        print(f"📊 Iniciando proceso con {num_seriales} seriales...")
        print(f"⏱️ Tiempo estimado: {num_seriales * 15} segundos (~{num_seriales * 15 / 60:.1f} minutos)")
        sys.stdout.flush()  # Fuerza que se imprima inmediatamente
        
        # Resetear el puntero del archivo después de leerlo
        archivo.seek(0)
        
    except Exception as e:
        print(f"⚠️ No se pudo leer el archivo: {e}")
        return jsonify({
            'status': 'error',
            'message': '❌ Error al leer el archivo. Verifica que sea un Excel válido.'
        }), 400
    
    # Llamar a la función de scraping
    excel_buffer = grf.consultar_en_grf(usuario, contra, archivo)

    # Verificar si se generó el archivo correctamente
    if excel_buffer is None or excel_buffer.getbuffer().nbytes == 0:
        print(f"❌ No se pudo generar el archivo para el usuario {usuario}")
        return jsonify({
            'status': 'error', 
            'message': '❌ No se encontraron datos o hubo un error en el proceso.'
        }), 404

    # Si todo salió bien, enviar el archivo
    print(f"✅ Archivo generado exitosamente para {usuario}")
    sys.stdout.flush()
    
    return send_file(
        excel_buffer,
        as_attachment=True,
        download_name=f"resultados_grf_{usuario}.xlsx", 
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

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

@app.route('/dash_sap')
def dash_sap():
    return render_template('dash_sap.html')


#------------ RUTA PRINCIPAL
@app.route('/')
def index():
    return render_template('index.html')


def buscar_ot_data(OT):
    """Devuelve los resultados de la búsqueda de una OT."""
    entrega_envio = doc_envios[doc_envios["OTP"] == OT]
    if entrega_envio.empty:
        return {"resultado": f"⚠️ No se encontraron registros con OT {OT}", "resultados": []}

    variable = entrega_envio['NºSerieFab'].tolist()
    sap_envio = entrega_envio["Material"].tolist()
    descrip_envio = entrega_envio["Texto breve de material"].tolist()
    cantidad_envio = entrega_envio["Ctd.en UM entrada"].tolist()

    resultados = []
    casca = 0
    # al momento en el que recibe la ot, lee el documento de envios y entra en un ciclo for pasando serial por serial
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
        # si no se encuentra en movimientos va a entrar al if
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
        # si se encuentra en movimientos continua con el proceso mas a fondo
        else:
            movimientos_df = pd.DataFrame(movimientos, columns=["Tipo", "Fecha", "SAP", "descrip"])
            ultimo = movimientos_df.loc[movimientos_df["Fecha"].idxmax()]
            tipo = ultimo["Tipo"]
            fecha = ultimo["Fecha"].strftime("%Y-%m-%d")
            sap = ultimo["SAP"]
            descrip = ultimo["descrip"]

            estado = "📦 ENTREGADO" if tipo == "Entrega" else \
                     "📦 SALIDA" if tipo == "Salida" else "🏠 DISPONIBLE"

            # Mostrar el ultimo dato de registro o movimiento
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
    # retornar valores
    return {"resultado": None, "resultados": resultados}

@app.route('/buscar', methods=['POST'])
def buscar_ot():
    try:
        OT = int(request.form['ot'])
    except:
        return render_template('index.html', resultado="⚠️ Ingresa un número de OT válido")

    data = buscar_ot_data(OT)
    return render_template('index.html', **data)


@app.route('/buscar1', methods=['POST'])
def buscar_ot1():
    try:
        OT = int(request.form['ot'])
    except:
        return render_template('dash_sap.html', resultado="⚠️ Ingresa un número de OT válido")

    data = buscar_ot_data(OT)
    return render_template('dash_sap.html', **data)


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)  # debug=False en producción