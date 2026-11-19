"""Local maintenance helpers. No Streamlit imports or external services."""
from contextlib import contextmanager, closing
from datetime import datetime
from pathlib import Path
from logging.handlers import RotatingFileHandler
import hashlib
import json
import logging
import re
import sqlite3
import tempfile
import threading
import zipfile

_backup_lock = threading.Lock()
_log_lock = threading.Lock()


def registrar_error(raiz, operacion, error):
    """Log only operation and exception class/code, never keys or document text."""
    carpeta = Path(raiz) / 'registros'
    carpeta.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger('xavilu.' + hashlib.sha256(str(carpeta.resolve()).encode()).hexdigest()[:12])
    with _log_lock:
        handler = RotatingFileHandler(carpeta / 'xavilu.log', maxBytes=1_000_000, backupCount=3, encoding='utf-8')
        handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
        logger.addHandler(handler)
        logger.setLevel(logging.ERROR)
        logger.propagate = False
        try:
            logger.error('%s: %s errno=%s', operacion, type(error).__name__, getattr(error, 'errno', None))
        finally:
            logger.removeHandler(handler)
            handler.close()



@contextmanager
def conexion_sqlite(*args, **kwargs):
    """Commit/rollback and always close the connection, including on Windows."""
    conexion = sqlite3.connect(*args, **kwargs)
    try:
        with conexion:
            yield conexion
    finally:
        conexion.close()


@contextmanager
def temporales_media(datos, sufijo_entrada, sufijo_salida):
    """Own both paths in one temporary directory, including failed preparations."""
    for sufijo in (sufijo_entrada, sufijo_salida):
        if not re.fullmatch(r'\.[a-zA-Z0-9]{1,10}', str(sufijo)):
            raise ValueError('Extensión de archivo no válida.')
    with tempfile.TemporaryDirectory(prefix='xavilu_media_') as carpeta:
        entrada = Path(carpeta) / ('entrada' + sufijo_entrada)
        salida = Path(carpeta) / ('salida' + sufijo_salida)
        entrada.write_bytes(datos)
        yield str(entrada), str(salida)


def crear_respaldo(raiz, forzar=False):
    """Daily backup while the app is used; online SQLite snapshot and data files.

    No automatic deletions. Secrets, virtual environments and caches are excluded.
    A manual forced backup creates a new version, preserving earlier copies.
    """
    raiz = Path(raiz).resolve()
    carpeta = raiz / 'Respaldos'
    with _backup_lock:
        carpeta.mkdir(exist_ok=True)
        fecha = datetime.now().strftime('%Y%m%d')
        diarios = sorted(carpeta.glob(f'xavilu_{fecha}_*.zip'))
        if diarios and not forzar:
            return diarios[-1]
        with tempfile.TemporaryDirectory(prefix='xavilu_respaldo_') as temporal:
            snapshot = Path(temporal) / 'xavilu_memoria.db'
            db = raiz / 'xavilu_memoria.db'
            if db.exists():
                with closing(sqlite3.connect(db, timeout=15)) as origen, closing(sqlite3.connect(snapshot)) as destino:
                    origen.backup(destino)
                    if destino.execute('PRAGMA quick_check').fetchone()[0] != 'ok':
                        raise RuntimeError('La verificación de la copia SQLite ha fallado.')
            archivos = [(snapshot, 'xavilu_memoria.db')] if snapshot.exists() else []
            for nombre in ('app.py', 'iniciar_xavilu.py', 'requirements.txt', 'requirements.lock.txt', 'xavilu_mantenimiento.py', 'xavilu_agentes_estudio.py', 'xavilu_aprendizaje.py', 'xavilu_archivo.py', 'xavilu_busqueda.py', 'xavilu_estado.py', 'xavilu_estudio_auto.py', 'xavilu_estudio_flujo.py', 'xavilu_imagenes.py', 'xavilu_mapa_pdf.py', 'xavilu_mapas.py', 'xavilu_material.py', 'xavilu_normativa_prl.py', 'xavilu_office.py', 'xavilu_pdf_editorial.py', 'xavilu_temporales.py', 'xavilu_trabajos.py', 'xavilu_video.py', 'xavimusic_historial.json'):
                p = raiz / nombre
                if p.is_file() and not p.is_symlink():
                    archivos.append((p, nombre))
            for nombre in ('ARCHIVO_XAVILU', 'Descargas', 'documentos_comunidad', 'Papelera_Xavilu'):
                base = raiz / nombre
                if base.exists():
                    for p in sorted(base.rglob('*')):
                        if p.is_file() and not p.is_symlink() and p.resolve().is_relative_to(raiz):
                            archivos.append((p, p.relative_to(raiz).as_posix()))
            preparado = Path(temporal) / 'respaldo.zip'
            manifest = []
            with zipfile.ZipFile(preparado, 'w', zipfile.ZIP_DEFLATED) as z:
                for archivo, relativo in archivos:
                    previo = archivo.stat()
                    digest = hashlib.sha256()
                    with archivo.open('rb') as entrada, z.open(relativo, 'w', force_zip64=True) as salida:
                        for bloque in iter(lambda: entrada.read(1024 * 1024), b''):
                            digest.update(bloque)
                            salida.write(bloque)
                    posterior = archivo.stat()
                    if (previo.st_size, previo.st_mtime_ns) != (posterior.st_size, posterior.st_mtime_ns):
                        raise RuntimeError('Un archivo cambió durante el respaldo. Vuelve a intentarlo.')
                    manifest.append({'ruta': relativo, 'sha256': digest.hexdigest(), 'bytes': previo.st_size})
                z.writestr('manifest.json', json.dumps(manifest, ensure_ascii=False, indent=2))
            with zipfile.ZipFile(preparado) as z:
                if z.testzip() is not None:
                    raise RuntimeError('No se pudo verificar el respaldo ZIP.')
            nombre = 'xavilu_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f') + '.zip'
            destino_zip = carpeta / nombre
            # Cross-volume safe: publish only when the complete ZIP has copied.
            parcial = carpeta / (nombre + '.partial')
            try:
                with preparado.open('rb') as entrada, parcial.open('xb') as salida:
                    for bloque in iter(lambda: entrada.read(1024 * 1024), b''):
                        salida.write(bloque)
                parcial.replace(destino_zip)
            finally:
                parcial.unlink(missing_ok=True)
            return destino_zip
