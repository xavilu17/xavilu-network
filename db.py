import sqlite3
import pandas as pd

DB_NAME = "taller_olivares.db"

def crear_conexion():
    return sqlite3.connect(DB_NAME, check_same_thread=False)

def inicializar_bd():
    with crear_conexion() as conn:
        cursor = conn.cursor()
        # Tabla de Clientes
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                telefono TEXT,
                matricula TEXT,
                moto TEXT
            )
        """)
        # Tabla de Historial
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS historial (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                id_cliente INTEGER,
                fecha TEXT,
                trabajo TEXT,
                FOREIGN KEY(id_cliente) REFERENCES clientes(id)
            )
        """)
        # Tabla de Avisos
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS avisos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                id_cliente INTEGER,
                fecha TEXT,
                motivo TEXT,
                FOREIGN KEY(id_cliente) REFERENCES clientes(id)
            )
        """)
        conn.commit()

# --- FUNCIONES PARA CLIENTES ---

def añadir_cliente(nombre, telefono, matricula, moto):
    with crear_conexion() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO clientes (nombre, telefono, matricula, moto) 
            VALUES (?, ?, ?, ?)
        """, (nombre, telefono, matricula, moto))
        conn.commit()
        return cursor.lastrowid

def obtener_clientes():
    with crear_conexion() as conn:
        df = pd.read_sql_query("SELECT * FROM clientes", conn)
        return df

def actualizar_cliente(id_cliente, nombre, telefono, matricula, moto):
    with crear_conexion() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE clientes 
            SET nombre=?, telefono=?, matricula=?, moto=? 
            WHERE id=?
        """, (nombre, telefono, matricula, moto, id_cliente))
        conn.commit()

# --- FUNCIONES PARA EL HISTORIAL ---

def guardar_historial(id_cliente, fecha, trabajo):
    with crear_conexion() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO historial (id_cliente, fecha, trabajo) 
            VALUES (?, ?, ?)
        """, (id_cliente, fecha, trabajo))
        conn.commit()

def obtener_historial(id_cliente):
    with crear_conexion() as conn:
        query = "SELECT fecha, trabajo FROM historial WHERE id_cliente = ? ORDER BY id DESC"
        df = pd.read_sql_query(query, conn, params=(id_cliente,))
        return df

# --- FUNCIONES PARA AVISOS ---

def añadir_aviso(id_cliente, fecha, motivo):
    with crear_conexion() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO avisos (id_cliente, fecha, motivo) 
            VALUES (?, ?, ?)
        """, (id_cliente, fecha, motivo))
        conn.commit()

def obtener_avisos():
    with crear_conexion() as conn:
        query = """
            SELECT a.id, c.nombre, c.matricula, a.fecha, a.motivo 
            FROM avisos a
            JOIN clientes c ON a.id_cliente = c.id
            ORDER BY a.fecha ASC
        """
        df = pd.read_sql_query(query, conn)
        return df

# NUEVA FUNCIÓN: Obtener avisos solo de un cliente
def obtener_avisos_cliente(id_cliente):
    with crear_conexion() as conn:
        query = "SELECT id, fecha, motivo FROM avisos WHERE id_cliente = ?"
        df = pd.read_sql_query(query, conn, params=(id_cliente,))
        return df

def eliminar_aviso(id_aviso):
    with crear_conexion() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM avisos WHERE id = ?", (id_aviso,))
        conn.commit()

# Inicializamos la base de datos al importar este archivo
inicializar_bd()
