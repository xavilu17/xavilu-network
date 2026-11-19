"""Repaso espaciado, simulacros y cuaderno de errores para Xavilu."""
from __future__ import annotations
import hashlib
from datetime import datetime, timedelta
from contextlib import contextmanager

INTERVALOS_DIAS = (1, 3, 7, 15, 30)


@contextmanager
def _db(path):
    from xavilu_mantenimiento import conexion_sqlite
    with conexion_sqlite(str(path), timeout=15) as db:
        db.executescript('''
        CREATE TABLE IF NOT EXISTS aprendizaje_tarjetas (
            id TEXT PRIMARY KEY, perfil TEXT NOT NULL, curso TEXT NOT NULL, tema TEXT NOT NULL,
            pregunta TEXT NOT NULL, respuesta TEXT NOT NULL, referencia TEXT NOT NULL DEFAULT '',
            fecha_revision TEXT, paso INTEGER NOT NULL DEFAULT 0, facilidad REAL NOT NULL DEFAULT 2.5,
            repeticiones INTEGER NOT NULL DEFAULT 0, fallos INTEGER NOT NULL DEFAULT 0,
            origen TEXT NOT NULL DEFAULT 'tarjeta', actualizada TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_aprendizaje_vencimiento ON aprendizaje_tarjetas(perfil,fecha_revision);
        CREATE TABLE IF NOT EXISTS aprendizaje_intentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT, perfil TEXT NOT NULL, curso TEXT NOT NULL,
            tema TEXT NOT NULL, total INTEGER NOT NULL, aciertos INTEGER NOT NULL,
            fallos INTEGER NOT NULL, puntuacion REAL NOT NULL, aprobado INTEGER NOT NULL,
            examen INTEGER NOT NULL DEFAULT 0, fecha TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_aprendizaje_intentos_fecha ON aprendizaje_intentos(perfil,fecha);
        CREATE TABLE IF NOT EXISTS aprendizaje_eventos (id INTEGER PRIMARY KEY AUTOINCREMENT, perfil TEXT NOT NULL, tipo TEXT NOT NULL, fecha TEXT NOT NULL);
        CREATE INDEX IF NOT EXISTS idx_aprendizaje_eventos_fecha ON aprendizaje_eventos(perfil,fecha);
        CREATE TABLE IF NOT EXISTS aprendizaje_errores (
            id TEXT PRIMARY KEY, perfil TEXT NOT NULL, curso TEXT NOT NULL, tema TEXT NOT NULL,
            pregunta TEXT NOT NULL, elegida TEXT NOT NULL, correcta TEXT NOT NULL,
            microexplicacion TEXT NOT NULL, referencia TEXT NOT NULL DEFAULT '',
            fecha TEXT NOT NULL, repasado INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS aprendizaje_objetivos (perfil TEXT PRIMARY KEY, tarjetas_dia INTEGER NOT NULL DEFAULT 10, minutos_dia INTEGER NOT NULL DEFAULT 30, actualizada TEXT NOT NULL);
        ''')
        yield db


def _now(): return datetime.now().replace(microsecond=0)
def _iso(dt): return dt.isoformat(timespec='seconds')


def registrar_tarjetas(path, perfil, curso, tema, tarjetas):
    ahora=_now(); cuenta=0
    with _db(path) as db:
        for t in tarjetas or []:
            pregunta=str(t.get('pregunta','')).strip();respuesta=str(t.get('respuesta','')).strip()
            if not pregunta or not respuesta:continue
            ref=str(t.get('referencia',''))
            ident=hashlib.sha256((str(perfil)+'\0'+str(curso)+'\0'+str(tema)+'\0'+pregunta+'\0'+respuesta).encode()).hexdigest()
            cursor=db.execute('''INSERT OR IGNORE INTO aprendizaje_tarjetas
                (id,perfil,curso,tema,pregunta,respuesta,referencia,fecha_revision,actualizada,origen)
                VALUES(?,?,?,?,?,?,?,?,?,?)''',(ident,str(perfil),str(curso),str(tema),pregunta,respuesta,ref,_iso(ahora),_iso(ahora),'tarjeta'))
            cuenta+=max(0,cursor.rowcount)
    return cuenta


def registrar_intento(path, perfil, curso, tema, preguntas, respuestas, examen=False, penalizacion=0, corte=70):
    preguntas=list(preguntas or []);respuestas=list(respuestas or []);total=len(preguntas)
    correctas=fallos=0;errores=[];fecha=_now()
    for i,(p,r) in enumerate(zip(preguntas,respuestas)):
        correcta=str(p.get('correcta',p.get('respuesta',''))).upper()
        if str(r).upper()==correcta:correctas+=1
        elif r not in ('','Sin responder',None):
            fallos+=1;op=p.get('opciones',{}) or {};elegida=str(r).upper()
            micro=(f"Elegiste {elegida}) {op.get(elegida,'')}. La correcta es {correcta}) {op.get(correcta,'')}. "
                    f"Diferencia y motivo: {p.get('explicacion','Repasa la referencia indicada y compara ambas opciones.')}")
            errores.append((i,p,elegida,correcta,micro))
    divisor=max(0,int(penalizacion or 0));descontados=(fallos//divisor) if divisor else 0
    puntuacion=round(max(0,correctas-descontados)*100/max(1,total),1);aprobado=puntuacion>=float(corte)
    with _db(path) as db:
        db.execute('INSERT INTO aprendizaje_intentos(perfil,curso,tema,total,aciertos,fallos,puntuacion,aprobado,examen,fecha) VALUES(?,?,?,?,?,?,?,?,?,?)',
                   (str(perfil),str(curso),str(tema),total,correctas,fallos,puntuacion,int(aprobado),int(examen),_iso(fecha)))
        db.execute('INSERT INTO aprendizaje_eventos(perfil,tipo,fecha) VALUES(?,?,?)',(str(perfil),'examen' if examen else 'test',_iso(fecha)))
        for i,p,elegida,correcta,micro in errores[:10]:
            ident=hashlib.sha256((str(perfil)+'\0'+str(curso)+'\0'+str(tema)+'\0'+str(p.get('pregunta',p.get('enunciado','')))).encode()).hexdigest()
            db.execute('''INSERT INTO aprendizaje_errores(id,perfil,curso,tema,pregunta,elegida,correcta,microexplicacion,referencia,fecha,repasado)
                         VALUES(?,?,?,?,?,?,?,?,?,?,0) ON CONFLICT(id) DO UPDATE SET
                         elegida=excluded.elegida,correcta=excluded.correcta,microexplicacion=excluded.microexplicacion,fecha=excluded.fecha,repasado=0''',
                       (ident,str(perfil),str(curso),str(tema),str(p.get('pregunta',p.get('enunciado',''))),elegida,correcta,micro,str(p.get('referencia',p.get('cita',''))),_iso(fecha)))
            tarjeta_id='error_'+ident
            db.execute('''INSERT INTO aprendizaje_tarjetas(id,perfil,curso,tema,pregunta,respuesta,referencia,fecha_revision,actualizada,origen)
                          VALUES(?,?,?,?,?,?,?,?,?,?) ON CONFLICT(id) DO UPDATE SET
                          respuesta=excluded.respuesta,fecha_revision=excluded.fecha_revision,actualizada=excluded.actualizada''',
                       (tarjeta_id,str(perfil),str(curso),str(tema),'Repaso de error: '+str(p.get('pregunta',p.get('enunciado',''))),micro,str(p.get('referencia',p.get('cita',''))),_iso(fecha),_iso(fecha),'error'))
    return {'total':total,'aciertos':correctas,'fallos':fallos,'puntuacion':puntuacion,'aprobado':aprobado,'errores':errores,'descontados':descontados}


def tarjetas_pendientes(path, perfil, limite=20):
    with _db(path) as db:
        rows=db.execute('''SELECT id,curso,tema,pregunta,respuesta,referencia,origen,fecha_revision,paso
            FROM aprendizaje_tarjetas WHERE perfil=? AND fecha_revision<=? ORDER BY fecha_revision,tema LIMIT ?''',
            (str(perfil),_iso(_now()),max(1,min(int(limite),100)))).fetchall()
    keys=('id','curso','tema','pregunta','respuesta','referencia','origen','fecha_revision','paso')
    return [dict(zip(keys,r)) for r in rows]


def calificar_tarjeta(path, perfil, ident, calidad):
    if isinstance(calidad, bool) or not str(calidad).strip().lstrip('-').isdigit():
        raise ValueError('La calidad debe ser un número entero entre 0 y 5.')
    calidad=int(calidad)
    if not 0 <= calidad <= 5:
        raise ValueError('La calidad debe estar entre 0 y 5.')
    perfil=str(perfil)
    ahora=_now()
    with _db(path) as db:
        fila=db.execute('SELECT paso,facilidad,repeticiones,fallos,origen FROM aprendizaje_tarjetas WHERE id=? AND perfil=?',(str(ident),perfil)).fetchone()
        if not fila:return False
        paso,facilidad,repeticiones,fallos,origen=fila
        if calidad<3:
            paso=0;repeticiones=0;fallos+=1;facilidad=max(1.3,facilidad-.2);dias=1
        else:
            paso=0 if repeticiones==0 else min(len(INTERVALOS_DIAS)-1,paso+(2 if calidad>=5 else 1))
            repeticiones+=1;facilidad=max(1.3,facilidad+(0.1 if calidad>=5 else 0.0));dias=INTERVALOS_DIAS[paso]
        db.execute('UPDATE aprendizaje_tarjetas SET paso=?,facilidad=?,repeticiones=?,fallos=?,fecha_revision=?,actualizada=? WHERE id=? AND perfil=?',
                   (paso,facilidad,repeticiones,fallos,_iso(ahora+timedelta(days=dias)),_iso(ahora),str(ident),perfil))
        if origen=='error' and calidad>=3:
            db.execute('UPDATE aprendizaje_errores SET repasado=1 WHERE id=? AND perfil=?',(str(ident).removeprefix('error_'),perfil))
            db.execute('DELETE FROM aprendizaje_tarjetas WHERE id=? AND perfil=?',(str(ident),perfil))
        db.execute('INSERT INTO aprendizaje_eventos(perfil,tipo,fecha) VALUES(?,?,?)',(perfil,'repaso',_iso(ahora)))
    return True


def objetivo_diario(path, perfil):
    with _db(path) as db:fila=db.execute('SELECT tarjetas_dia,minutos_dia FROM aprendizaje_objetivos WHERE perfil=?',(str(perfil),)).fetchone()
    return {'tarjetas':fila[0],'minutos':fila[1]} if fila else {'tarjetas':10,'minutos':30}


def guardar_objetivo_diario(path, perfil, tarjetas, minutos):
    tarjetas=max(1,min(100,int(tarjetas)));minutos=max(5,min(240,int(minutos)))
    with _db(path) as db:
        db.execute('''INSERT INTO aprendizaje_objetivos(perfil,tarjetas_dia,minutos_dia,actualizada) VALUES(?,?,?,?)
            ON CONFLICT(perfil) DO UPDATE SET tarjetas_dia=excluded.tarjetas_dia,minutos_dia=excluded.minutos_dia,actualizada=excluded.actualizada''',
            (str(perfil),tarjetas,minutos,_iso(_now())))
    return {'tarjetas':tarjetas,'minutos':minutos}


def resumen(path, perfil):
    ahora=_now();hoy=ahora.date().isoformat()
    with _db(path) as db:
        debidas=db.execute('SELECT COUNT(*) FROM aprendizaje_tarjetas WHERE perfil=? AND fecha_revision<=?',(str(perfil),_iso(ahora))).fetchone()[0]
        total=db.execute('SELECT COUNT(*) FROM aprendizaje_tarjetas WHERE perfil=?',(str(perfil),)).fetchone()[0]
        hoy_vistas=db.execute("SELECT COUNT(*) FROM aprendizaje_eventos WHERE perfil=? AND tipo='repaso' AND substr(fecha,1,10)=?",(str(perfil),hoy)).fetchone()[0]
        dias=[r[0] for r in db.execute('SELECT DISTINCT substr(fecha,1,10) FROM aprendizaje_eventos WHERE perfil=? ORDER BY 1 DESC',(str(perfil),)).fetchall()]
        streak=0;fecha=ahora.date()
        for day in dias:
            d=datetime.fromisoformat(day).date()
            if d==fecha or d==fecha-timedelta(days=1):streak+=1;fecha=d-timedelta(days=1)
            else:break
        flojos=db.execute('''SELECT tema,COUNT(*) n,AVG(puntuacion) media FROM aprendizaje_intentos WHERE perfil=? GROUP BY tema HAVING n>0 ORDER BY media ASC,n DESC LIMIT 3''',(str(perfil),)).fetchall()
        recientes=db.execute('SELECT COUNT(*),SUM(aprobado) FROM aprendizaje_intentos WHERE perfil=?',(str(perfil),)).fetchone()
        errores=db.execute('SELECT COUNT(*) FROM aprendizaje_errores WHERE perfil=? AND repasado=0',(str(perfil),)).fetchone()[0]
    return {'debidas':debidas,'total_tarjetas':total,'vistas_hoy':hoy_vistas,'racha':streak,'objetivo':objetivo_diario(path,perfil),
            'temas_flojos':[{'tema':r[0],'intentos':r[1],'media':round(r[2] or 0,1)} for r in flojos],
            'intentos':recientes[0] or 0,'aprobados':recientes[1] or 0,'errores_pendientes':errores}


def errores_pendientes(path, perfil, limite=10):
    with _db(path) as db:
        rows=db.execute('SELECT curso,tema,pregunta,elegida,correcta,microexplicacion,referencia,fecha FROM aprendizaje_errores WHERE perfil=? AND repasado=0 ORDER BY fecha DESC LIMIT ?',
                        (str(perfil),max(1,min(int(limite),100)))).fetchall()
    return [{'curso':r[0],'tema':r[1],'pregunta':r[2],'elegida':r[3],'correcta':r[4],'explicacion':r[5],'referencia':r[6],'fecha':r[7]} for r in rows]
