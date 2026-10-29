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
        # Tabla de Historial (Nueva)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS historial (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                id_cliente INTEGER,
                fecha TEXT,
                trabajo TEXT,
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
        # Buscamos los trabajos de ese cliente, ordenados del más reciente al más antiguo
        query = "SELECT fecha, trabajo FROM historial WHERE id_cliente = ? ORDER BY id DESC"
        df = pd.read_sql_query(query, conn, params=(id_cliente,))
        return df

# Inicializamos la base de datos al importar este archivo
inicializar_bd()
