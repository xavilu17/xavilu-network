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
        # Usamos pandas para que Streamlit lo muestre en tabla muy fácil
        df = pd.read_sql_query("SELECT * FROM clientes", conn)
        return df

# Inicializamos la base de datos al importar este archivo
inicializar_bd()
