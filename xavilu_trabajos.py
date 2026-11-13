"""Último borrador por herramienta y perfil. No guarda adjuntos ni credenciales."""
from datetime import datetime, timezone
from pathlib import Path
import hashlib,json
from xavilu_mantenimiento import conexion_sqlite

EXTRA_PERMITIDO={'formato','titulo','documento','entrada','voces','modulo','voz','estilo'}


class Trabajos:
    def __init__(self,raiz):
        self.raiz=Path(raiz)
        with self._db() as db:
            db.execute('''CREATE TABLE IF NOT EXISTS trabajos_recientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT, perfil TEXT NOT NULL,
                clave TEXT NOT NULL, titulo TEXT NOT NULL, tipo TEXT NOT NULL,
                texto TEXT NOT NULL, destino TEXT NOT NULL, extra TEXT NOT NULL,
                huella TEXT NOT NULL, actualizado TEXT NOT NULL,
                UNIQUE(perfil,clave))''')

    def _db(self):return conexion_sqlite(self.raiz/'xavilu_memoria.db',timeout=15)

    @staticmethod
    def _perfil(perfil):
        if perfil not in {'admin','ninos'}:raise ValueError('Perfil no válido.')
        return perfil

    def guardar(self,perfil,clave,titulo,tipo,texto,destino,extra=None):
        perfil=self._perfil(perfil)
        if not isinstance(texto,str) or not texto.strip() or len(texto)>2_000_000:
            raise ValueError('El borrador está vacío o supera el límite de recuperación.')
        if not isinstance(clave,str) or not 1<=len(clave)<=240:raise ValueError('Clave de trabajo no válida.')
        if not isinstance(destino,dict) or set(destino)-{'macro','sub','accion'}:raise ValueError('Destino no válido.')
        macro=destino.get('macro')
        if not (type(macro)is int and 0<=macro<=7) and macro!='audio':raise ValueError('Sección no válida.')
        if perfil=='ninos' and macro==7:raise ValueError('Comunidad requiere acceso de administrador.')
        if type(destino.get('sub',0))is not int or not 0<=destino.get('sub',0)<=12:raise ValueError('Herramienta no válida.')
        if len(str(destino.get('accion','')))>100:raise ValueError('Acción no válida.')
        extra=extra or {}
        if not isinstance(extra,dict) or set(extra)-EXTRA_PERMITIDO:raise ValueError('Solo se guardan datos del borrador, no configuración privada.')
        if any(not isinstance(v,(str,int)) or len(str(v))>4000 for v in extra.values()):raise ValueError('Dato de trabajo no válido.')
        d=json.dumps(destino,sort_keys=True);e=json.dumps(extra,sort_keys=True,ensure_ascii=False)
        titulo=' '.join(str(titulo).split())[:160] or 'Trabajo guardado';tipo=str(tipo)[:60]
        huella=hashlib.sha256((titulo+'\0'+tipo+'\0'+texto+'\0'+d+'\0'+e).encode()).hexdigest()
        with self._db() as db:
            anterior=db.execute('SELECT id,huella FROM trabajos_recientes WHERE perfil=? AND clave=?',(perfil,clave)).fetchone()
            if anterior and anterior[1]==huella:return anterior[0]
            fecha=datetime.now(timezone.utc).isoformat(timespec='microseconds')
            db.execute('''INSERT INTO trabajos_recientes(perfil,clave,titulo,tipo,texto,destino,extra,huella,actualizado)
                VALUES(?,?,?,?,?,?,?,?,?) ON CONFLICT(perfil,clave) DO UPDATE SET titulo=excluded.titulo,
                tipo=excluded.tipo,texto=excluded.texto,destino=excluded.destino,extra=excluded.extra,
                huella=excluded.huella,actualizado=excluded.actualizado''',(perfil,clave,titulo,tipo,texto,d,e,huella,fecha))
            return db.execute('SELECT id FROM trabajos_recientes WHERE perfil=? AND clave=?',(perfil,clave)).fetchone()[0]

    def listar(self,perfil,limite=8):
        perfil=self._perfil(perfil)
        with self._db() as db:
            return [{'id':i,'titulo':t,'tipo':k,'actualizado':f}for i,t,k,f in db.execute(
                'SELECT id,titulo,tipo,actualizado FROM trabajos_recientes WHERE perfil=? ORDER BY actualizado DESC,id DESC LIMIT ?',
                (perfil,max(1,min(int(limite),50))))]

    def abrir(self,perfil,ident):
        perfil=self._perfil(perfil)
        if type(ident)is not int or ident<1:raise ValueError('Trabajo no válido.')
        with self._db() as db:
            fila=db.execute('SELECT id,clave,titulo,tipo,texto,destino,extra,actualizado FROM trabajos_recientes WHERE id=? AND perfil=?',(ident,perfil)).fetchone()
        if not fila:raise ValueError('El trabajo no está disponible para este perfil.')
        ident,clave,titulo,tipo,texto,destino,extra,fecha=fila
        return {'id':ident,'clave':clave,'titulo':titulo,'tipo':tipo,'texto':texto,'destino':json.loads(destino),'extra':json.loads(extra),'actualizado':fecha}
