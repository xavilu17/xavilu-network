"""PRL source selection and grounded citations; independent of the UI."""
import json
import re
import unicodedata


def _normalizar(texto):
    return ' '.join(str(texto).casefold().split())


def ids_normativa_prl(consulta):
    q = ''.join(c for c in unicodedata.normalize('NFKD', consulta.casefold()) if not unicodedata.combining(c))
    ids = []
    reglas = [
        (r'carretill|maquin|equipo de trabajo|trabajo en altura|escalera|andam', 'BOE-A-1997-17824'),
        (r'\bepi\b|\bepis\b|proteccion individual|casco|guante|mascarilla', 'BOE-A-1997-12735'),
        (r'manipulacion|carga|levantamiento|niosh', 'BOE-A-1997-8670'),
        (r'lugar de trabajo|iluminacion|temperatura|locales|aseo', 'BOE-A-1997-8669'),
        (r'ruido|acustic|decibel', 'BOE-A-2006-4414'),
        (r'obra|construccion|coordinador|plan de seguridad', 'BOE-A-1997-22614'),
        (r'servicio de prevencion|organizacion preventiva|modalidad preventiva', 'BOE-A-1997-1853'),
    ]
    for patron, identificador in reglas:
        if re.search(patron, q) and identificador not in ids:
            ids.append(identificador)
    return ids[:3]


def contexto_normativo_prl(fuentes):
    bloques = []
    presupuesto = 40000
    for fuente in fuentes:
        for bloque in fuente.get('fragmentos', [])[:6]:
            registro = {'norma': fuente['id'], 'titulo_norma': fuente['titulo'],
                        'bloque': bloque['id'], 'encabezado': bloque['titulo'],
                        'texto': bloque['texto'][:4500]}
            tamano = len(json.dumps(registro, ensure_ascii=False))
            if tamano <= presupuesto:
                bloques.append(registro)
                presupuesto -= tamano
    return json.dumps(bloques, ensure_ascii=False)


def validar_respuesta_normativa_prl(respuesta, fuentes):
    bruto = re.sub(r'^```(?:json)?\s*|\s*```$', '', str(respuesta).strip(), flags=re.I)
    datos = json.loads(bruto)
    if not isinstance(datos, dict) or not isinstance(datos.get('respuesta'), str) or not datos['respuesta'].strip():
        raise ValueError('La respuesta normativa no tiene un texto válido.')
    citas = datos.get('citas')
    if not isinstance(citas, list) or not citas:
        raise ValueError('No hay citas verificables que respalden la explicación.')
    fuentes_por_id = {f['id']: f for f in fuentes}
    verificadas = []
    entregados = {(b['norma'], b['bloque']): b['texto'] for b in json.loads(contexto_normativo_prl(fuentes))}
    for cita in citas:
        if not isinstance(cita, dict):
            raise ValueError('Cita normativa no válida.')
        fuente = fuentes_por_id.get(cita.get('norma'))
        bloque = next((b for b in (fuente or {}).get('fragmentos', [])[:6] if b.get('id') == cita.get('bloque')), None)
        extracto = cita.get('extracto')
        if bloque is None or not isinstance(extracto, str) or len(extracto.strip()) < 15 or len(extracto) > 1200:
            raise ValueError('La IA citó una norma o un bloque no recuperado, o una cita insuficiente.')
        if _normalizar(extracto) not in _normalizar(entregados.get((fuente['id'], bloque['id']), '')):
            raise ValueError('La cita no coincide con el texto del BOE recuperado.')
        verificadas.append({'norma': fuente['id'], 'titulo': fuente['titulo'], 'articulo': bloque['titulo'],
                            'extracto': extracto, 'url': f"https://www.boe.es/buscar/act.php?id={fuente['id']}#{bloque['id']}"})
    # Prevent the common failure of inventing a periodic renewal requirement.
    periodos = re.findall(r'\b(\d+)\s*(años?|mes(?:es)?|d[ií]as?|horas?)\b', datos['respuesta'], re.I)
    evidencia = ' '.join(c['extracto'] for c in verificadas).casefold()
    for numero, unidad in periodos:
        if not re.search(r'\b' + re.escape(numero) + r'\s*' + re.escape(unidad) + r'\b', evidencia, re.I):
            raise ValueError('Se indicó un plazo numérico que no está respaldado por las citas.')
    return {'respuesta': datos['respuesta'].strip(), 'citas': verificadas,
            'limites': str(datos.get('limites') or '').strip()}
