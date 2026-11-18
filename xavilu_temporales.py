"""Gestión de resultados temporales de una sesión de Xavilu."""
from pathlib import Path
from tempfile import TemporaryDirectory
import re


class ResultadosTemporales:
    def __init__(self):
        self._temporal = TemporaryDirectory(prefix='xavilu_resultados_')
        self.raiz = Path(self._temporal.name)
        self.archivos = {}

    def guardar(self, nombre, contenido, seccion='General'):
        base = Path(str(nombre or 'resultado')).name
        base = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', base).strip(' .') or 'resultado'
        ruta = self.raiz / base
        indice = 2
        while ruta.exists():
            ruta = self.raiz / f'{Path(base).stem}_{indice}{Path(base).suffix}'
            indice += 1
        datos = contenido.encode('utf-8') if isinstance(contenido, str) else bytes(contenido)
        ruta.write_bytes(datos)
        self.archivos[str(ruta)] = {'ruta': str(ruta), 'nombre': ruta.name, 'seccion': str(seccion or 'General')}
        return str(ruta)

    def descartar(self, ruta):
        clave = str(ruta)
        self.archivos.pop(clave, None)
        try:
            Path(clave).unlink(missing_ok=True)
        except OSError:
            pass

    def limpiar(self):
        self.archivos.clear()
        self._temporal.cleanup()

