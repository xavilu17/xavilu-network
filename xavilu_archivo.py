"""Organización y papelera local. No IA, red ni borrado definitivo."""
from datetime import datetime, timezone
from pathlib import Path
import os
import sqlite3
import uuid


class ArchivoLocal:
    MAX_ARCHIVOS_LISTADO = 20000
    BASES = ('Descargas', 'ARCHIVO_XAVILU')

    def __init__(self, raiz):
        self.raiz = Path(raiz).resolve()
        self.listado_limitado = False
        with self._db() as db:
            db.executescript('''
                CREATE TABLE IF NOT EXISTS archivo_organizacion (
                    ruta TEXT PRIMARY KEY, categoria TEXT NOT NULL DEFAULT '',
                    favorito INTEGER NOT NULL DEFAULT 0);
                CREATE TABLE IF NOT EXISTS archivo_papelera (
                    id TEXT PRIMARY KEY, original TEXT NOT NULL, fecha TEXT NOT NULL,
                    estado TEXT NOT NULL DEFAULT 'pendiente');
                CREATE TABLE IF NOT EXISTS archivo_actividad (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, ruta TEXT NOT NULL,
                    fecha TEXT NOT NULL, accion TEXT NOT NULL, detalle TEXT NOT NULL DEFAULT '');
            ''')

    def _db(self):
        # sqlite3's own context does not close handles on Windows.
        from xavilu_mantenimiento import conexion_sqlite
        return conexion_sqlite(self.raiz / 'xavilu_memoria.db', timeout=15)

    def _ruta(self, ruta, papelera=False):
        p = Path(ruta)
        if not p.is_absolute():
            p = self.raiz / p
        relativo = p.relative_to(self.raiz)
        bases = ('Papelera_Xavilu',) if papelera else self.BASES
        if len(relativo.parts) < 2 or relativo.parts[0] not in bases:
            raise ValueError('Archivo fuera de la biblioteca de Xavilu.')
        for parte in (p, *p.parents):
            if parte == self.raiz:
                break
            if parte.is_symlink() or getattr(parte, 'is_junction', lambda: False)():
                raise ValueError('No se permiten enlaces ni accesos a otras carpetas.')
        real = p.resolve()
        base = self.raiz / relativo.parts[0]
        if not real.is_relative_to(base) or '..' in relativo.parts:
            raise ValueError('Ruta no permitida.')
        return real

    def _clave(self, ruta):
        return self._ruta(ruta).relative_to(self.raiz).as_posix()

    @staticmethod
    def _fecha():
        return datetime.now(timezone.utc).isoformat(timespec='seconds')

    def _evento(self, db, ruta, accion, detalle=''):
        db.execute('INSERT INTO archivo_actividad(ruta,fecha,accion,detalle) VALUES(?,?,?,?)',
                   (ruta, self._fecha(), accion, detalle))

    def listar(self):
        with self._db() as db:
            meta = {r[0]: {'categoria': r[1], 'favorito': bool(r[2])}
                    for r in db.execute('SELECT ruta,categoria,favorito FROM archivo_organizacion')}
        archivos = []
        self.listado_limitado = False
        for nombre in self.BASES:
            base = self.raiz / nombre
            if not base.exists():
                continue
            self._ruta(base / '_validacion')
            for p in base.rglob('*'):
                if not p.is_file():
                    continue
                if len(archivos) >= self.MAX_ARCHIVOS_LISTADO:
                    self.listado_limitado = True
                    break
                try:
                    p = self._ruta(p)
                    stat = p.stat()
                except (FileNotFoundError, ValueError):
                    continue
                relativo = p.relative_to(self.raiz).as_posix()
                archivos.append({'ruta': relativo, 'nombre': p.name,
                    'origen': p.parent.relative_to(self.raiz).as_posix(),
                    'bytes': stat.st_size, 'modificado': stat.st_mtime,
                    'categoria': meta.get(relativo, {}).get('categoria', ''),
                    'favorito': meta.get(relativo, {}).get('favorito', False)})
            if self.listado_limitado:
                break
        return sorted(archivos, key=lambda x: x['modificado'], reverse=True)

    def resolver_archivo(self, ruta):
        """Return a verified file path inside the Xavilu library."""
        return self._ruta(ruta)

    def organizar(self, ruta, categoria, favorito):
        ruta = self._clave(ruta)
        if not (self.raiz / ruta).is_file():
            raise FileNotFoundError('El archivo ya no está disponible.')
        categoria = str(categoria).strip()
        if len(categoria) > 80 or any(ord(c) < 32 for c in categoria):
            raise ValueError('Usa una categoría de hasta 80 caracteres, en una sola línea.')
        if not isinstance(favorito, bool):
            raise ValueError('Favorito debe ser verdadero o falso.')
        with self._db() as db:
            previo = db.execute('SELECT categoria,favorito FROM archivo_organizacion WHERE ruta=?', (ruta,)).fetchone()
            if previo == (categoria, int(favorito)):
                return
            db.execute('INSERT INTO archivo_organizacion VALUES(?,?,?) ON CONFLICT(ruta) DO UPDATE SET categoria=excluded.categoria,favorito=excluded.favorito',
                       (ruta, categoria, int(favorito)))
            self._evento(db, ruta, 'Organización actualizada', categoria or 'Sin categoría')

    @staticmethod
    def _mover(origen, destino):
        # Never overwrite a file that appeared since the UI was rendered.
        if os.name == 'nt':
            os.rename(origen, destino)  # Windows raises if destination exists.
        else:
            os.link(origen, destino)
            origen.unlink()

    def enviar_papelera(self, ruta):
        original = self._ruta(ruta)
        if not original.is_file():
            raise FileNotFoundError('El archivo ya no está disponible.')
        clave = self._clave(original)
        identificador = uuid.uuid4().hex
        destino = self._ruta(Path('Papelera_Xavilu') / identificador, papelera=True)
        destino.parent.mkdir(exist_ok=True)
        # Persist the intent before moving. A crash still leaves a recoverable entry.
        with self._db() as db:
            db.execute('INSERT INTO archivo_papelera(id,original,fecha) VALUES(?,?,?)',
                       (identificador, clave, self._fecha()))
        self._mover(original, destino)
        with self._db() as db:
            db.execute("UPDATE archivo_papelera SET estado='en papelera' WHERE id=?", (identificador,))
            self._evento(db, clave, 'Enviado a papelera')
        return identificador

    def papelera(self):
        with self._db() as db:
            filas = db.execute("SELECT id,original,fecha FROM archivo_papelera WHERE estado IN ('pendiente','en papelera') ORDER BY fecha DESC").fetchall()
        return [{'id': i, 'original': o, 'fecha': f} for i, o, f in filas
                if self._ruta(Path('Papelera_Xavilu') / i, papelera=True).is_file()]

    def restaurar(self, identificador):
        if not isinstance(identificador, str) or len(identificador) != 32 or any(c not in '0123456789abcdef' for c in identificador):
            raise ValueError('Identificador de papelera no válido.')
        with self._db() as db:
            fila = db.execute('SELECT original FROM archivo_papelera WHERE id=?', (identificador,)).fetchone()
        if not fila:
            raise ValueError('El archivo no figura en la papelera.')
        original = self._ruta(fila[0])
        origen = self._ruta(Path('Papelera_Xavilu') / identificador, papelera=True)
        if not origen.is_file():
            raise FileNotFoundError('El archivo ya no está en la papelera.')
        destino = original
        numero = 1
        while destino.exists():
            destino = original.with_name(f'{original.stem}_restaurado_{numero}{original.suffix}')
            numero += 1
        destino.parent.mkdir(parents=True, exist_ok=True)
        self._mover(origen, destino)
        clave = self._clave(destino)
        with self._db() as db:
            db.execute("UPDATE archivo_papelera SET estado='restaurado' WHERE id=?", (identificador,))
            if clave != fila[0]:
                db.execute('INSERT OR IGNORE INTO archivo_organizacion SELECT ?,categoria,favorito FROM archivo_organizacion WHERE ruta=?', (clave, fila[0]))
            self._evento(db, clave, 'Restaurado de papelera', fila[0])
        return destino

    def actividad(self, ruta):
        clave = self._clave(ruta)
        with self._db() as db:
            return [{'fecha': f, 'acción': a, 'detalle': d} for f,a,d in db.execute(
                'SELECT fecha,accion,detalle FROM archivo_actividad WHERE ruta=? ORDER BY id DESC LIMIT 50', (clave,))]
