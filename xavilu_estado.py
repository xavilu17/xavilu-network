"""Comprobaciones locales bajo petición. Sin pruebas de IA ni lecturas de secretos."""
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
from contextlib import closing
import importlib.metadata,importlib.util,re,shutil,sqlite3,subprocess

DEPENDENCIAS=[('PDF editorial','reportlab','reportlab'),('Lectura y edición PDF','pymupdf','PyMuPDF'),
    ('PDF compatible','pypdf','pypdf'),('Word','docx','python-docx'),('Excel','openpyxl','openpyxl'),
    ('PowerPoint','pptx','python-pptx'),('Edición de imágenes','PIL','pillow'),('Voces','edge_tts','edge-tts'),
    ('OCR local','rapidocr','rapidocr'),('Transcripción por API','openai','openai')]


def revisar_estado(raiz,ffmpeg=None):
    raiz=Path(raiz).resolve();resultado={'fecha':datetime.now(ZoneInfo('Europe/Madrid')).isoformat(timespec='seconds')}
    resultado['disco_libre']=shutil.disk_usage(raiz).free
    db=raiz/'xavilu_memoria.db'
    if db.exists():
        try:
            with closing(sqlite3.connect(db.as_uri()+'?mode=ro',uri=True,timeout=5)) as c:
                resultado['base_datos']='Correcta' if c.execute('PRAGMA quick_check').fetchone()[0]=='ok' else 'Necesita revisión'
        except sqlite3.Error:resultado['base_datos']='No se pudo comprobar; puede estar ocupada o necesitar revisión'
    else:resultado['base_datos']='Aún no se ha creado'
    copias=list((raiz/'Respaldos').glob('xavilu_*.zip'))
    ultima=max(copias,key=lambda p:p.stat().st_mtime) if copias else None
    resultado['respaldo']={'nombre':ultima.name,'fecha':datetime.fromtimestamp(ultima.stat().st_mtime,ZoneInfo('Europe/Madrid')).isoformat(timespec='seconds'),'bytes':ultima.stat().st_size} if ultima else None
    filas=[]
    for nombre,modulo,paquete in DEPENDENCIAS:
        disponible=importlib.util.find_spec(modulo) is not None
        try:version=importlib.metadata.version(paquete) if disponible else ''
        except importlib.metadata.PackageNotFoundError:version=''
        filas.append({'Herramienta':nombre,'Estado':'Disponible' if disponible else 'No instalada','Versión':version})
    ruta_ffmpeg=ffmpeg or shutil.which('ffmpeg')
    ejecutable=shutil.which(str(ruta_ffmpeg)) if ruta_ffmpeg else None
    if ruta_ffmpeg and Path(str(ruta_ffmpeg)).is_file():ejecutable=str(ruta_ffmpeg)
    ff_ok=False
    if ejecutable:
        try:
            proceso=subprocess.run([ejecutable,'-version'],capture_output=True,timeout=5,check=True)
            ff_ok=bool(proceso.stdout)
        except (OSError,subprocess.SubprocessError):pass
    filas.append({'Herramienta':'Audio y vídeo · FFmpeg','Estado':'Disponible' if ff_ok else 'No se pudo comprobar','Versión':''})
    candidatos=[shutil.which('soffice'),shutil.which('libreoffice'),r'C:\Program Files\LibreOffice\program\soffice.exe',r'C:\Program Files (x86)\LibreOffice\program\soffice.exe']
    libre=any(p and Path(p).is_file()for p in candidatos)
    filas.append({'Herramienta':'Office a PDF · LibreOffice','Estado':'Disponible' if libre else 'No encontrado; necesario para esa conversión','Versión':''})
    resultado['herramientas']=filas
    totales=[]
    for nombre in ['ARCHIVO_XAVILU','Descargas','Papelera_Xavilu']:
        cantidad=tamano=0;parcial=False
        for p in (raiz/nombre).rglob('*'):
            if p.is_symlink() or not p.is_file() or not p.resolve().is_relative_to(raiz):continue
            try:tamano+=p.stat().st_size;cantidad+=1
            except OSError:parcial=True
            if cantidad>=20000:parcial=True;break
        totales.append({'Carpeta':nombre,'Archivos':cantidad,'MB':round(tamano/1024**2,2),'Lectura':'Parcial' if parcial else 'Completa'})
    resultado['almacenamiento']=totales
    resultado['errores']=[]
    log=raiz/'registros/xavilu.log'
    if log.is_file():
        for linea in log.read_text(encoding='utf-8',errors='replace').splitlines()[-15:]:
            m=re.match(r'^(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d)[,\d]* ERROR ([^:]{1,160}): ([A-Za-z_][A-Za-z_0-9]*) errno=',linea)
            if m:resultado['errores'].append({'Fecha':m[1],'Operación':m[2],'Tipo':m[3]})
    return resultado
