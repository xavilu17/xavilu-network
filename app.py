# ============================================================
# XAVILU NETWORK · PLANTA
# Backend completo (Flask + SQLite)
# ============================================================

import os
import sqlite3
import qrcode
from io import BytesIO
from datetime import datetime
from flask import (Flask, render_template, request, jsonify,
                   send_file, redirect, url_for, flash, abort)

app = Flask(__name__)
app.secret_key = 'xavilu-planta-secret-2026'
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'xavilu_planta.db')


# ------------------------------------------------------------
# BASE DE DATOS
# ------------------------------------------------------------
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_db()
    c = conn.cursor()

    c.execute("""CREATE TABLE IF NOT EXISTS tipos_vehiculo (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL UNIQUE,
        activo INTEGER DEFAULT 1
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS tipos_merma (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL UNIQUE,
        activo INTEGER DEFAULT 1
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS empresas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL,
        cif TEXT,
        direccion TEXT,
        telefono TEXT,
        email TEXT,
        activo INTEGER DEFAULT 1,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT DEFAULT CURRENT_TIMESTAMP
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS vehiculos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        empresa_id INTEGER NOT NULL,
        matricula TEXT NOT NULL UNIQUE,
        tipo_id INTEGER,
        tara REAL DEFAULT 0,
        activo INTEGER DEFAULT 1,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (empresa_id) REFERENCES empresas(id),
        FOREIGN KEY (tipo_id) REFERENCES tipos_vehiculo(id)
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS choferes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        empresa_id INTEGER NOT NULL,
        nombre TEXT NOT NULL,
        dni TEXT NOT NULL,
        activo INTEGER DEFAULT 1,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (empresa_id) REFERENCES empresas(id)
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS ordenes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        numero_orden TEXT NOT NULL UNIQUE,
        empresa_id INTEGER,
        vehiculo_id INTEGER,
        chofer_id INTEGER,
        matricula_snap TEXT,
        chofer_nombre_snap TEXT,
        dni_snap TEXT,
        tara_snap REAL,
        codigo_ler TEXT,
        tipo_merma_id INTEGER,
        estado TEXT NOT NULL DEFAULT 'creada',
        area_actual TEXT NOT NULL DEFAULT 'porteria',
        fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP,
        fecha_pesaje TEXT,
        fecha_albaran TEXT,
        fecha_salida TEXT,
        motivo_anulacion TEXT,
        observaciones TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (empresa_id) REFERENCES empresas(id),
        FOREIGN KEY (vehiculo_id) REFERENCES vehiculos(id),
        FOREIGN KEY (chofer_id) REFERENCES choferes(id),
        FOREIGN KEY (tipo_merma_id) REFERENCES tipos_merma(id)
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS pesajes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        orden_id INTEGER NOT NULL UNIQUE,
        peso_bruto REAL NOT NULL,
        merma REAL DEFAULT 0,
        peso_neto REAL NOT NULL,
        fecha_pesaje TEXT DEFAULT CURRENT_TIMESTAMP,
        observaciones TEXT,
        FOREIGN KEY (orden_id) REFERENCES ordenes(id)
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS albaranes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        orden_id INTEGER NOT NULL UNIQUE,
        numero_albaran TEXT NOT NULL UNIQUE,
        fecha_emision TEXT DEFAULT CURRENT_TIMESTAMP,
        usuario TEXT,
        observaciones TEXT,
        FOREIGN KEY (orden_id) REFERENCES ordenes(id)
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS eventos_orden (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        orden_id INTEGER NOT NULL,
        estado_anterior TEXT,
        estado_nuevo TEXT,
        area TEXT,
        usuario TEXT,
        observaciones TEXT,
        timestamp TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (orden_id) REFERENCES ordenes(id)
    )""")

    for tv in ('Camión', 'Furgoneta', 'Cisterna', 'Tractora'):
        c.execute("INSERT OR IGNORE INTO tipos_vehiculo (nombre) VALUES (?)", (tv,))
    for tm in ('Ninguna', 'Humedad', 'Impurezas', 'Rotura'):
        c.execute("INSERT OR IGNORE INTO tipos_merma (nombre) VALUES (?)", (tm,))

    conn.commit()
    conn.close()


# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
def generar_numero(conn, tabla, campo, prefix):
    year = datetime.now().year
    full_prefix = f"{prefix}-{year}-"
    row = conn.execute(
        f"SELECT {campo} FROM {tabla} WHERE {campo} LIKE ? ORDER BY id DESC LIMIT 1",
        (full_prefix + '%',)
    ).fetchone()
    if row:
        try:
            seq = int(row[campo].split('-')[-1]) + 1
        except (ValueError, IndexError):
            seq = 1
    else:
        seq = 1
    return f"{full_prefix}{seq:04d}"


def registrar_evento(conn, orden_id, estado_ant, estado_nuevo, area, usuario=None, obs=None):
    conn.execute("""INSERT INTO eventos_orden
        (orden_id, estado_anterior, estado_nuevo, area, usuario, observaciones)
        VALUES (?, ?, ?, ?, ?, ?)""",
        (orden_id, estado_ant, estado_nuevo, area, usuario, obs))


def buscar_orden(conn, numero):
    if not numero:
        return None
    return conn.execute(
        "SELECT * FROM ordenes WHERE numero_orden = ?", (numero.strip().upper(),)
    ).fetchone()


# ------------------------------------------------------------
# CONTEXT PROCESSOR (para mostrar siempre empresa por defecto en forms)
# ------------------------------------------------------------
@app.context_processor
def inject_globals():
    return {'now': datetime.now()}


# ------------------------------------------------------------
# HOME = PANEL EN PLANTA
# ------------------------------------------------------------
@app.route('/')
@app.route('/panel')
def panel():
    return render_template('panel.html', active='panel')


@app.route('/api/en-planta')
def api_en_planta():
    conn = get_db()
    rows = conn.execute("""
        SELECT o.*, e.nombre AS empresa_nombre,
               (SELECT numero_albaran FROM albaranes WHERE orden_id=o.id) AS numero_albaran
        FROM ordenes o
        LEFT JOIN empresas e ON e.id = o.empresa_id
        WHERE o.estado NOT IN ('salida', 'anulada')
        ORDER BY o.fecha_creacion DESC
    """).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


# ------------------------------------------------------------
# PORTERÍA · CREAR ORDEN
# ------------------------------------------------------------
@app.route('/porteria')
def porteria():
    conn = get_db()
    empresas = conn.execute("SELECT * FROM empresas WHERE activo=1 ORDER BY nombre").fetchall()
    tipos_veh = conn.execute("SELECT * FROM tipos_vehiculo WHERE activo=1").fetchall()
    tipos_merma = conn.execute("SELECT * FROM tipos_merma WHERE activo=1").fetchall()
    conn.close()
    return render_template('porteria.html', active='porteria',
                           empresas=empresas, tipos_veh=tipos_veh,
                           tipos_merma=tipos_merma)


@app.route('/api/vehiculos/<int:empresa_id>')
def api_vehiculos(empresa_id):
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM vehiculos WHERE empresa_id=? AND activo=1 ORDER BY matricula",
        (empresa_id,)
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/choferes/<int:empresa_id>')
def api_choferes(empresa_id):
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM choferes WHERE empresa_id=? AND activo=1 ORDER BY nombre",
        (empresa_id,)
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/maestro/empresa', methods=['POST'])
def crear_empresa():
    data = request.get_json()
    conn = get_db()
    c = conn.execute(
        "INSERT INTO empresas (nombre, cif, telefono, email) VALUES (?,?,?,?)",
        (data['nombre'], data.get('cif'), data.get('telefono'), data.get('email'))
    )
    conn.commit()
    new_id = c.lastrowid
    conn.close()
    return jsonify({'id': new_id, 'nombre': data['nombre']})


@app.route('/maestro/vehiculo', methods=['POST'])
def crear_vehiculo():
    data = request.get_json()
    conn = get_db()
    try:
        c = conn.execute(
            "INSERT INTO vehiculos (empresa_id, matricula, tipo_id, tara) VALUES (?,?,?,?)",
            (data['empresa_id'], data['matricula'].upper(),
             data.get('tipo_id'), float(data.get('tara') or 0))
        )
        conn.commit()
        new_id = c.lastrowid
        conn.close()
        return jsonify({'id': new_id, 'matricula': data['matricula'].upper(),
                        'tara': float(data.get('tara') or 0)})
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({'error': 'Matrícula duplicada'}), 400


@app.route('/maestro/chofer', methods=['POST'])
def crear_chofer():
    data = request.get_json()
    conn = get_db()
    c = conn.execute(
        "INSERT INTO choferes (empresa_id, nombre, dni) VALUES (?,?,?)",
        (data['empresa_id'], data['nombre'], data['dni'].upper())
    )
    conn.commit()
    new_id = c.lastrowid
    conn.close()
    return jsonify({'id': new_id, 'nombre': data['nombre'], 'dni': data['dni'].upper()})


@app.route('/porteria/nueva', methods=['POST'])
def nueva_orden():
    f = request.form
    conn = get_db()

    empresa_id = f.get('empresa_id')
    vehiculo_id = f.get('vehiculo_id')
    chofer_id = f.get('chofer_id')

    veh = conn.execute("SELECT * FROM vehiculos WHERE id=?", (vehiculo_id,)).fetchone()
    cho = conn.execute("SELECT * FROM choferes WHERE id=?", (chofer_id,)).fetchone()

    if not veh or not cho:
        conn.close()
        flash('Vehículo o chofer no válido', 'error')
        return redirect(url_for('porteria'))

    numero = generar_numero(conn, 'ordenes', 'numero_orden', 'P')

    cur = conn.execute("""INSERT INTO ordenes
        (numero_orden, empresa_id, vehiculo_id, chofer_id,
         matricula_snap, chofer_nombre_snap, dni_snap, tara_snap,
         codigo_ler, tipo_merma_id, estado, area_actual, observaciones)
        VALUES (?,?,?,?,?,?,?,?,?,?, 'creada', 'porteria', ?)""",
        (numero, empresa_id, vehiculo_id, chofer_id,
         veh['matricula'], cho['nombre'], cho['dni'], veh['tara'],
         f.get('codigo_ler'), f.get('tipo_merma_id') or None,
         f.get('observaciones')))
    orden_id = cur.lastrowid
    registrar_evento(conn, orden_id, None, 'creada', 'porteria')
    conn.commit()
    conn.close()
    return redirect(url_for('ver_orden', numero=numero))


@app.route('/porteria/orden/<numero>')
def ver_orden(numero):
    conn = get_db()
    orden = buscar_orden(conn, numero)
    if not orden:
        conn.close()
        abort(404)
    empresa = conn.execute("SELECT * FROM empresas WHERE id=?", (orden['empresa_id'],)).fetchone()
    conn.close()
    return render_template('orden_creada.html', active='porteria',
                           orden=orden, empresa=empresa)


@app.route('/qr/<numero>')
def qr_image(numero):
    img = qrcode.make(numero)
    buf = BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return send_file(buf, mimetype='image/png')


# ------------------------------------------------------------
# BÁSCULA
# ------------------------------------------------------------
@app.route('/bascula')
def bascula():
    return render_template('bascula.html', active='bascula', orden=None)


@app.route('/bascula/cargar', methods=['POST'])
def bascula_cargar():
    numero = request.form.get('numero_orden', '').strip().upper()
    return redirect(url_for('bascula_orden', numero=numero))


@app.route('/bascula/<numero>')
def bascula_orden(numero):
    conn = get_db()
    orden = buscar_orden(conn, numero)
    if not orden:
        conn.close()
        flash(f'Orden {numero} no encontrada', 'error')
        return redirect(url_for('bascula'))
    empresa = conn.execute("SELECT * FROM empresas WHERE id=?", (orden['empresa_id'],)).fetchone()
    conn.close()
    if orden['estado'] in ('anulada',):
        flash(f'Orden {numero} está anulada', 'error')
        return redirect(url_for('bascula'))
    return render_template('bascula.html', active='bascula',
                           orden=orden, empresa=empresa)


@app.route('/bascula/<numero>/pesar', methods=['POST'])
def bascula_pesar(numero):
    conn = get_db()
    orden = buscar_orden(conn, numero)
    if not orden:
        conn.close()
        flash('Orden no encontrada', 'error')
        return redirect(url_for('bascula'))

    if orden['estado'] not in ('creada', 'en_bascula'):
        conn.close()
        flash(f'Orden ya en estado "{orden["estado"]}"', 'error')
        return redirect(url_for('bascula'))

    try:
        bruto = float(request.form['peso_bruto'])
        merma = float(request.form.get('merma') or 0)
    except ValueError:
        conn.close()
        flash('Peso inválido', 'error')
        return redirect(url_for('bascula_orden', numero=numero))

    tara = orden['tara_snap'] or 0
    neto = round(bruto - tara - merma, 2)

    conn.execute("""INSERT INTO pesajes (orden_id, peso_bruto, merma, peso_neto, observaciones)
                    VALUES (?,?,?,?,?)""",
                 (orden['id'], bruto, merma, neto, request.form.get('observaciones')))
    conn.execute("""UPDATE ordenes SET estado='pendiente_albaran', area_actual='recepcion',
                    fecha_pesaje=CURRENT_TIMESTAMP, updated_at=CURRENT_TIMESTAMP WHERE id=?""",
                 (orden['id'],))
    registrar_evento(conn, orden['id'], orden['estado'], 'pendiente_albaran', 'bascula',
                     obs=f'Bruto={bruto} Merma={merma} Neto={neto}')
    conn.commit()
    conn.close()
    flash(f'Pesaje guardado · Neto {neto} kg', 'ok')
    return redirect(url_for('bascula'))


# ------------------------------------------------------------
# RECEPCIÓN
# ------------------------------------------------------------
@app.route('/recepcion')
def recepcion():
    return render_template('recepcion.html', active='recepcion')


@app.route('/api/pendientes')
def api_pendientes():
    conn = get_db()
    rows = conn.execute("""
        SELECT o.*, e.nombre AS empresa_nombre, p.peso_neto
        FROM ordenes o
        LEFT JOIN empresas e ON e.id = o.empresa_id
        LEFT JOIN pesajes p ON p.orden_id = o.id
        WHERE o.estado = 'pendiente_albaran'
        ORDER BY o.fecha_pesaje ASC
    """).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/recepcion/<numero>')
def recepcion_albaran(numero):
    conn = get_db()
    orden = buscar_orden(conn, numero)
    if not orden:
        conn.close()
        abort(404)
    empresa = conn.execute("SELECT * FROM empresas WHERE id=?", (orden['empresa_id'],)).fetchone()
    pesaje = conn.execute("SELECT * FROM pesajes WHERE orden_id=?", (orden['id'],)).fetchone()
    conn.close()
    return render_template('recepcion_albaran.html', active='recepcion',
                           orden=orden, empresa=empresa, pesaje=pesaje)


@app.route('/recepcion/<numero>/emitir', methods=['POST'])
def recepcion_emitir(numero):
    conn = get_db()
    orden = buscar_orden(conn, numero)
    if not orden or orden['estado'] != 'pendiente_albaran':
        conn.close()
        flash('Estado no válido para emitir albarán', 'error')
        return redirect(url_for('recepcion'))

    num_alb = generar_numero(conn, 'albaranes', 'numero_albaran', 'ALB')
    conn.execute("""INSERT INTO albaranes (orden_id, numero_albaran, usuario, observaciones)
                    VALUES (?,?,?,?)""",
                 (orden['id'], num_alb, request.form.get('usuario', 'recepción'),
                  request.form.get('observaciones')))
    conn.execute("""UPDATE ordenes SET estado='albaranada', area_actual='salida',
                    fecha_albaran=CURRENT_TIMESTAMP, updated_at=CURRENT_TIMESTAMP WHERE id=?""",
                 (orden['id'],))
    registrar_evento(conn, orden['id'], 'pendiente_albaran', 'albaranada', 'recepcion',
                     obs=f'Albarán {num_alb}')
    conn.commit()
    conn.close()
    flash(f'Albarán {num_alb} emitido', 'ok')
    return redirect(url_for('recepcion'))


# ------------------------------------------------------------
# SALIDA
# ------------------------------------------------------------
@app.route('/salida')
def salida():
    return render_template('salida.html', active='salida')


@app.route('/salida/<numero>/confirmar', methods=['POST'])
def salida_confirmar(numero):
    conn = get_db()
    orden = buscar_orden(conn, numero)
    if not orden:
        conn.close()
        flash('Orden no encontrada', 'error')
        return redirect(url_for('salida'))
    if orden['estado'] != 'albaranada':
        conn.close()
        flash(f'La orden está en estado "{orden["estado"]}", no se puede cerrar', 'error')
        return redirect(url_for('salida'))
    conn.execute("""UPDATE ordenes SET estado='salida', area_actual='salida',
                    fecha_salida=CURRENT_TIMESTAMP, updated_at=CURRENT_TIMESTAMP WHERE id=?""",
                 (orden['id'],))
    registrar_evento(conn, orden['id'], 'albaranada', 'salida', 'salida')
    conn.commit()
    conn.close()
    flash('Salida confirmada', 'ok')
    return redirect(url_for('salida'))


# ------------------------------------------------------------
# ANULAR
# ------------------------------------------------------------
@app.route('/orden/<numero>/anular', methods=['POST'])
def anular_orden(numero):
    conn = get_db()
    orden = buscar_orden(conn, numero)
    if not orden:
        conn.close()
        abort(404)
    if orden['estado'] in ('salida', 'anulada'):
        conn.close()
        flash('No se puede anular', 'error')
        return redirect(url_for('panel'))
    motivo = request.form.get('motivo', 'Sin motivo')
    conn.execute("""UPDATE ordenes SET estado='anulada', area_actual='anulada',
                    motivo_anulacion=?, updated_at=CURRENT_TIMESTAMP WHERE id=?""",
                 (motivo, orden['id']))
    registrar_evento(conn, orden['id'], orden['estado'], 'anulada', 'admin', obs=motivo)
    conn.commit()
    conn.close()
    flash('Orden anulada', 'ok')
    return redirect(url_for('panel'))


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------
if __name__ == '__main__':
    if os.path.exists(DB_PATH):
        pass
    init_db()
    app.run(host='0.0.0.0', port=5000, debug=True)
