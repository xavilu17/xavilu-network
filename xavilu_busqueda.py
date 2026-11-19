"""Búsqueda local de texto recuperable. Sin IA, red, OCR ni activación de contexto."""
from pathlib import Path
import re
import unicodedata
import zipfile
from xavilu_archivo import ArchivoLocal
from xavilu_mantenimiento import conexion_sqlite

VERSION = 1
EXTENSIONES = {'.pdf', '.docx', '.xlsx', '.pptx', '.txt', '.csv', '.md'}
MAX_ARCHIVO = 50 * 1024 * 1024
MAX_TEXTO = 2_000_000
MAX_PAGINAS = 1000


class LecturaLimitada(Exception):
    pass


def _normalizar(texto):
    return ''.join(c for c in unicodedata.normalize('NFKD', texto.casefold()) if not unicodedata.combining(c))


def consulta_segura(consulta):
    if not isinstance(consulta, str) or len(consulta) > 200:
        raise ValueError('Escribe una búsqueda de hasta 200 caracteres.')
    omitidas = {'a','al','de','del','el','la','los','las','en','un','una','unos','unas','y','o','por','para','con','que','se'}
    palabras = list(dict.fromkeys(t for t in re.findall(r'[^\W_]+', _normalizar(consulta)) if t not in omitidas))
    if not palabras:
        raise ValueError('Escribe alguna palabra del contenido que quieras encontrar.')
    if len(palabras) > 12:
        raise ValueError('Usa hasta 12 palabras para concretar la búsqueda.')
    # Tokens are data, never SQL or FTS operators. Prefixes also find plural endings.
    return ' AND '.join('"' + t.replace('"', '""') + '"*' for t in palabras)


def _trocear(texto, tamano=6000):
    texto = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', ' ', str(texto)).strip()
    if not texto:
        return
    for inicio in range(0, len(texto), tamano - 250):
        yield texto[inicio:inicio + tamano]


def extraer_bloques(ruta, informe):
    """Yield real page/slide/paragraph/row locations; no invented pagination."""
    extension = ruta.suffix.lower()
    if extension in {'.docx', '.xlsx', '.pptx'}:
        with zipfile.ZipFile(ruta) as z:
            if sum(i.file_size for i in z.infolist()) > 100 * 1024 * 1024:
                raise LecturaLimitada('El contenido descomprimido supera el límite de lectura local (100 MB).')
    if extension == '.pdf':
        from pypdf import PdfReader
        with ruta.open('rb') as source:
            doc = PdfReader(source, strict=False)
            if doc.is_encrypted and not doc.decrypt(''):
                raise LecturaLimitada('PDF protegido: necesita contraseña para recuperar su texto.')
            total = len(doc.pages)
            if total > MAX_PAGINAS:
                informe.append(f'Lectura parcial: primeras {MAX_PAGINAS} de {total} páginas.')
            sin_texto = 0
            for numero in range(min(total, MAX_PAGINAS)):
                texto = doc.pages[numero].extract_text() or ''
                if not texto.strip():
                    sin_texto += 1
                yield f'Página {numero + 1}', texto
            if sin_texto:
                informe.append(f'{sin_texto} página(s) sin texto recuperable; esta búsqueda no aplica OCR.')
    elif extension == '.docx':
        from docx import Document
        from docx.text.paragraph import Paragraph
        from docx.table import Table
        doc = Document(ruta)
        parrafo = tabla = 0
        for item in doc.element.body:
            if item.tag.endswith('}p'):
                parrafo += 1
                yield f'Párrafo {parrafo}', Paragraph(item, doc).text
            elif item.tag.endswith('}tbl'):
                tabla += 1
                yield f'Tabla {tabla}', '\n'.join(' | '.join(c.text for c in row.cells) for row in Table(item, doc).rows)
        for numero, section in enumerate(doc.sections, 1):
            for label, container in [('Cabecera', section.header), ('Pie', section.footer)]:
                yield f'{label} · sección {numero}', '\n'.join(p.text for p in container.paragraphs)
    elif extension == '.xlsx':
        from openpyxl import load_workbook
        from openpyxl.utils import get_column_letter
        libro = load_workbook(ruta, read_only=True, data_only=False, keep_links=False)
        try:
            for hoja in libro.worksheets:
                max_filas = min(hoja.max_row or 20000, 20000)
                max_columnas = min(hoja.max_column or 200, 200)
                if (hoja.max_row or 0) > 20000 or (hoja.max_column or 0) > 200:
                    informe.append(f'Lectura parcial de {hoja.title}: límite de 20.000 filas y 200 columnas.')
                lineas = [];inicio = 1
                for n, row in enumerate(hoja.iter_rows(max_row=max_filas, max_col=max_columnas, values_only=True), 1):
                    valores = [f'{get_column_letter(i)}: {v}' for i,v in enumerate(row, 1) if v is not None]
                    if valores:
                        lineas.append(f'Fila {n}: ' + ' | '.join(valores))
                    if n % 50 == 0:
                        yield f'Hoja {hoja.title} · filas {inicio}–{n}', '\n'.join(lineas)
                        lineas = [];inicio = n + 1
                if lineas:
                    yield f'Hoja {hoja.title} · filas {inicio}–{max_filas}', '\n'.join(lineas)
        finally:
            libro.close()
    elif extension == '.pptx':
        from pptx import Presentation
        def textos(shapes):
            for shape in shapes:
                if getattr(shape, 'has_text_frame', False):
                    yield shape.text_frame.text
                if getattr(shape, 'has_table', False):
                    for row in shape.table.rows:
                        yield ' | '.join(c.text for c in row.cells)
                if hasattr(shape, 'shapes'):
                    yield from textos(shape.shapes)
        for numero, slide in enumerate(Presentation(ruta).slides, 1):
            yield f'Diapositiva {numero}', '\n'.join(textos(slide.shapes))
            if slide.has_notes_slide:
                notas = slide.notes_slide.notes_text_frame
                if notas is not None:
                    yield f'Notas · diapositiva {numero}', notas.text
    else:
        datos = ruta.read_bytes()
        try:
            texto = datos.decode('utf-8-sig')
        except UnicodeDecodeError:
            texto = datos.decode('cp1252', errors='replace')
        lineas = texto.splitlines()
        for inicio in range(0, len(lineas), 60):
            yield f'Líneas {inicio + 1}–{min(inicio + 60, len(lineas))}', '\n'.join(lineas[inicio:inicio + 60])


class IndiceArchivo:
    def __init__(self, raiz):
        self.archivo = ArchivoLocal(raiz)
        self.raiz = self.archivo.raiz
        with self._db() as db:
            db.executescript('''
                CREATE TABLE IF NOT EXISTS archivo_indice_documentos (
                    ruta TEXT PRIMARY KEY, tamano INTEGER NOT NULL, mtime_ns INTEGER NOT NULL,
                    version INTEGER NOT NULL, estado TEXT NOT NULL, detalle TEXT NOT NULL,
                    caracteres INTEGER NOT NULL);
                CREATE VIRTUAL TABLE IF NOT EXISTS archivo_busqueda_fts USING fts5(
                    texto, ruta UNINDEXED, ubicacion UNINDEXED,
                    tokenize='unicode61 remove_diacritics 2');
            ''')

    def _db(self):
        return conexion_sqlite(self.raiz / 'xavilu_memoria.db', timeout=15)

    def _actuales(self):
        actuales = {}
        fichas = self.archivo.listar()
        if self.archivo.listado_limitado:
            raise RuntimeError(f'La biblioteca supera {self.archivo.MAX_ARCHIVOS_LISTADO:,} archivos; reduce o separa el contenido antes de indexarla.')
        for ficha in fichas:
            if Path(ficha['ruta']).suffix.lower() not in EXTENSIONES:
                continue
            try:
                ruta = self.archivo.resolver_archivo(ficha['ruta'])
                stat = ruta.stat()
            except (FileNotFoundError, ValueError):
                continue
            actuales[ficha['ruta']] = (stat.st_size, stat.st_mtime_ns)
        return actuales

    def estado(self):
        actuales = self._actuales()
        with self._db() as db:
            guardados = {r[0]: r[1:] for r in db.execute('SELECT ruta,tamano,mtime_ns,version,estado,detalle,caracteres FROM archivo_indice_documentos')}
        resumen = {'compatibles': len(actuales), 'preparados': 0, 'pendientes': 0, 'incidencias': []}
        for ruta, firma in actuales.items():
            guardado = guardados.get(ruta)
            if not guardado or guardado[:3] != (*firma, VERSION):
                resumen['pendientes'] += 1
            elif guardado[3] in ('ok', 'parcial'):
                resumen['preparados'] += 1
            if guardado and guardado[:3] == (*firma, VERSION) and guardado[3] != 'ok':
                resumen['incidencias'].append({'archivo': ruta, 'estado': guardado[3], 'detalle': guardado[4]})
        return resumen

    def actualizar(self, progreso=None):
        actuales = self._actuales()
        with self._db() as db:
            previos = {r[0]: r[1:] for r in db.execute('SELECT ruta,tamano,mtime_ns,version,estado FROM archivo_indice_documentos')}
        pendientes = [(r,f) for r,f in actuales.items() if r not in previos or previos[r][:3] != (*f,VERSION) or previos[r][3] == 'error']
        leidos = 0
        for numero, (clave, firma) in enumerate(pendientes, 1):
            ruta = self.archivo.resolver_archivo(clave)
            bloques = [];detalle = [];caracteres = 0;estado = 'ok'
            try:
                if firma[0] > MAX_ARCHIVO:
                    raise LecturaLimitada('Supera el límite de lectura de 50 MB.')
                generador = extraer_bloques(ruta, detalle)
                try:
                    for ubicacion, texto in generador:
                        restante = MAX_TEXTO - caracteres
                        if len(texto) > restante:
                            texto = texto[:restante]
                            detalle.append('Lectura parcial: límite de 2 millones de caracteres por archivo.')
                        for trozo in _trocear(texto):
                            bloques.append((trozo, clave, ubicacion))
                        caracteres += len(texto)
                        if caracteres >= MAX_TEXTO:
                            if not any('2 millones' in aviso for aviso in detalle):
                                detalle.append('Lectura parcial: límite de 2 millones de caracteres por archivo.')
                            break
                finally:
                    generador.close()
                despues = ruta.stat()
                if (despues.st_size, despues.st_mtime_ns) != firma:
                    raise LecturaLimitada('El archivo cambió durante la lectura. Actualiza de nuevo.')
                estado = 'parcial' if detalle else 'ok'
                if not bloques:
                    estado = 'sin texto'
                    detalle.append('Sin texto recuperable. Un PDF escaneado puede necesitar OCR.')
            except LecturaLimitada as exc:
                bloques = [];caracteres = 0;estado = 'no leído';detalle = [str(exc)]
            except Exception as exc:
                bloques = [];caracteres = 0;estado = 'error';detalle = ['No se pudo leer el archivo (' + type(exc).__name__ + ').']
            with self._db() as db:
                db.execute('DELETE FROM archivo_busqueda_fts WHERE ruta=?', (clave,))
                db.executemany('INSERT INTO archivo_busqueda_fts(texto,ruta,ubicacion) VALUES(?,?,?)', bloques)
                db.execute('INSERT INTO archivo_indice_documentos VALUES(?,?,?,?,?,?,?) ON CONFLICT(ruta) DO UPDATE SET tamano=excluded.tamano,mtime_ns=excluded.mtime_ns,version=excluded.version,estado=excluded.estado,detalle=excluded.detalle,caracteres=excluded.caracteres',
                           (clave, *firma, VERSION, estado, ' '.join(detalle), caracteres))
            leidos += 1
            if progreso:
                progreso(numero, len(pendientes), clave)
        # Drop index entries for files removed from the library or sent to trash.
        with self._db() as db:
            for clave in set(previos) - set(actuales):
                db.execute('DELETE FROM archivo_busqueda_fts WHERE ruta=?', (clave,))
                db.execute('DELETE FROM archivo_indice_documentos WHERE ruta=?', (clave,))
        return {'leidos': leidos, **self.estado()}

    def buscar(self, consulta, limite=20):
        expresion = consulta_segura(consulta)
        actuales = self._actuales()
        with self._db() as db:
            db.execute('CREATE TEMP TABLE actuales(ruta TEXT PRIMARY KEY,tamano INTEGER,mtime_ns INTEGER)')
            db.executemany('INSERT INTO actuales VALUES(?,?,?)', [(k,*v) for k,v in actuales.items()])
            cursor = db.execute('''
                SELECT archivo_busqueda_fts.ruta,archivo_busqueda_fts.ubicacion,
                    snippet(archivo_busqueda_fts,0,'','', ' … ',36),
                    bm25(archivo_busqueda_fts),d.estado,d.detalle
                FROM archivo_busqueda_fts
                JOIN archivo_indice_documentos d ON d.ruta=archivo_busqueda_fts.ruta
                JOIN actuales a ON a.ruta=d.ruta AND a.tamano=d.tamano AND a.mtime_ns=d.mtime_ns
                WHERE archivo_busqueda_fts MATCH ? AND d.version=? AND d.estado IN ('ok','parcial')
                ORDER BY bm25(archivo_busqueda_fts)
            ''', (expresion, VERSION))
            encontrados = {}
            for ruta,ubicacion,fragmento,puntuacion,estado,detalle in cursor:
                if ruta not in encontrados:
                    if len(encontrados) >= max(1, min(int(limite), 50)):
                        break
                    encontrados[ruta] = {'ruta': ruta, 'estado': estado, 'detalle': detalle, 'coincidencias': []}
                coincidencias = encontrados[ruta]['coincidencias']
                if len(coincidencias) < 2:
                    coincidencias.append({'ubicacion': ubicacion, 'fragmento': fragmento})
            return list(encontrados.values())
