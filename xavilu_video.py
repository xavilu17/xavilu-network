"""Descarga acotada en un proceso independiente; publica solo archivos completos."""
from pathlib import Path
import ipaddress
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
from urllib.parse import urlsplit

LIMITE = 500 * 1024 * 1024

def validar_url(url):
    try:
        p = urlsplit(url)
        if p.scheme not in ('https', 'http') or not p.hostname or p.username or p.password or p.port not in (None, 80, 443):
            raise ValueError()
        direcciones = socket.getaddrinfo(p.hostname, p.port or 443, type=socket.SOCK_STREAM)
        if not direcciones or not all(ipaddress.ip_address(x[4][0].split('%')[0]).is_global for x in direcciones):
            raise ValueError()
    except (ValueError, OSError):
        raise ValueError('Introduce una dirección HTTP/HTTPS pública válida.') from None

def publicar(origen, carpeta, nombre):
    origen, carpeta = Path(origen), Path(carpeta)
    if origen.suffix not in ('.mp4', '.mp3') or not 0 < origen.stat().st_size <= LIMITE:
        raise ValueError('El resultado está vacío, supera 500 MB o tiene un formato inesperado.')
    nombre = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', nombre).strip(' .')[:120] or 'video'
    if nombre.split('.')[0].upper() in {'CON', 'PRN', 'AUX', 'NUL', *('COM'+str(i) for i in range(1,10)), *('LPT'+str(i) for i in range(1,10))}:
        nombre = '_' + nombre
    carpeta.mkdir(parents=True, exist_ok=True)
    # A complete temporary file is linked atomically; concurrent saves never overwrite.
    with tempfile.TemporaryDirectory(prefix='.descarga_', dir=carpeta) as tmp:
        preparado = Path(tmp) / ('completo' + origen.suffix)
        shutil.copyfile(origen, preparado)
        for n in range(10000):
            destino = carpeta / (nombre + (f' ({n})' if n else '') + origen.suffix)
            try:
                if os.name == 'nt':
                    os.rename(preparado, destino)  # Windows rename refuses existing targets.
                else:
                    os.link(preparado, destino)
                return destino
            except FileExistsError:
                continue
    raise RuntimeError('No se pudo asignar un nombre al archivo.')

def comprimir(origen, raiz, ffmpeg, nivel='Equilibrado', nombre=''):
    perfiles = {'Buena calidad':(23,1080), 'Equilibrado':(28,720), 'Archivo pequeño':(32,480), '360p':(32,360), '240p':(32,240)}
    if nivel not in perfiles:
        raise ValueError('Nivel de compresión no válido.')
    origen=Path(origen)
    if not origen.is_file() or not 0 < origen.stat().st_size <= LIMITE:
        raise ValueError('El vídeo debe ocupar entre 1 byte y 500 MB.')
    crf,alto=perfiles[nivel]
    with tempfile.TemporaryDirectory(prefix='xavilu_comprimir_') as tmp:
        salida=Path(tmp)/'comprimido.mp4'
        comando=[str(ffmpeg),'-nostdin','-y','-i',str(origen),'-map','0:v:0','-map','0:a:0?',
                 '-vf',f'scale=-2:trunc(min(ih\\,{alto})/2)*2','-c:v','libx264','-preset','fast',
                 '-crf',str(crf),'-pix_fmt','yuv420p','-c:a','aac','-b:a','96k','-movflags','+faststart',
                 '-fs',str(LIMITE),str(salida)]
        try:
            subprocess.run(comando,check=True,stdout=subprocess.DEVNULL,stderr=subprocess.PIPE,timeout=900,
                           creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
        except subprocess.TimeoutExpired:
            raise RuntimeError('La compresión ha superado 15 minutos. Prueba un vídeo más corto.') from None
        except subprocess.CalledProcessError:
            raise RuntimeError('No se pudo comprimir este vídeo. Comprueba que se reproduce correctamente.') from None
        if salida.stat().st_size >= LIMITE:
            raise ValueError('El resultado alcanza el límite de tamaño y no se guardará incompleto.')
        if salida.stat().st_size >= origen.stat().st_size:
            return {'reducido':False,'antes':origen.stat().st_size,'despues':salida.stat().st_size}
        destino=publicar(salida,Path(raiz)/'Descargas'/'Videos comprimidos',(nombre or origen.stem)+'_comprimido')
        return {'reducido':True,'ruta':str(destino),'antes':origen.stat().st_size,'despues':destino.stat().st_size}

def descargar(url, raiz, ffmpeg, modo='MP4', calidad=720, nombre=''):
    try:
        modo = str(modo).strip().upper()
        calidad = int(str(calidad).strip().lower().removesuffix('p'))
    except (ValueError, TypeError):
        raise ValueError('Elige una calidad: 1080p, 720p, 480p, 360p o 240p.') from None
    validar_url(url)
    if modo not in ('MP4','MP3') or calidad not in (240,360,480,720,1080):
        raise ValueError('Formato o calidad no válidos.')
    raiz = Path(raiz)
    carpeta = raiz / 'Descargas' / ('Videos descargados' if modo == 'MP4' else 'Audio descargado')
    carpeta.mkdir(parents=True, exist_ok=True)
    if shutil.disk_usage(carpeta).free < 3 * LIMITE:
        raise ValueError('Se necesitan al menos 1,5 GB libres para preparar la descarga.')
    with tempfile.TemporaryDirectory(prefix='xavilu_video_') as tmp:
        datos = dict(url=url, carpeta=tmp, ffmpeg=str(ffmpeg), modo=modo, calidad=calidad)
        proceso = subprocess.Popen([sys.executable, str(Path(__file__).resolve()), '--worker'],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True,
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        try:
            salida, _ = proceso.communicate(json.dumps(datos), timeout=900)
        except subprocess.TimeoutExpired:
            if os.name == 'nt':
                subprocess.run(['taskkill','/PID',str(proceso.pid),'/T','/F'], capture_output=True,
                               creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            else:
                proceso.kill()
            proceso.communicate()
            raise RuntimeError('La descarga ha superado 15 minutos. Prueba un vídeo más corto.') from None
        try:
            resultado = json.loads(salida)
        except (ValueError, TypeError):
            raise RuntimeError('El descargador no pudo completar el trabajo.') from None
        if proceso.returncode or not resultado.get('ok'):
            raise RuntimeError(resultado.get('error', 'No se pudo completar la descarga.'))
        origen = Path(resultado['archivo']).resolve()
        if not origen.is_relative_to(Path(tmp).resolve()) or origen.is_symlink():
            raise RuntimeError('Ruta de resultado no válida.')
        destino = publicar(origen, carpeta, nombre.strip() or resultado['titulo'])
        return {'ruta':str(destino), 'titulo':resultado['titulo'], 'bytes':destino.stat().st_size}

def worker(datos):
    # Apply network restrictions only in this disposable process, never in Streamlit.
    resolver = socket.getaddrinfo
    def resolver_publico(*args, **kwargs):
        filas = resolver(*args, **kwargs)
        if not filas or any(not ipaddress.ip_address(x[4][0].split('%')[0]).is_global for x in filas):
            raise OSError('Destino de red no público.')
        return filas
    socket.getaddrinfo = resolver_publico
    conectar = socket.socket.connect
    def conectar_publico(sock, destino):
        if isinstance(destino, tuple) and not ipaddress.ip_address(destino[0].split('%')[0]).is_global:
            raise OSError('Destino de red no público.')
        return conectar(sock, destino)
    socket.socket.connect = conectar_publico
    from yt_dlp import YoutubeDL
    from yt_dlp.downloader.external import FFmpegFD
    # Network downloads must use guarded Python sockets; FFmpeg is for local merging only.
    def bloquear_red_externa(*args, **kwargs):
        raise ValueError('Este flujo requiere un descargador externo no habilitado.')
    FFmpegFD.real_download = bloquear_red_externa
    class Descargador(YoutubeDL):
        def urlopen(self, req):
            validar_url(req if isinstance(req,str) else (getattr(req,'url',None) or req.full_url))
            return super().urlopen(req)
    class Silencio:
        def debug(self,*a): pass
        def warning(self,*a): pass
        def error(self,*a): pass
    carpeta = Path(datos['carpeta'])
    herramientas = carpeta / 'tools'
    herramientas.mkdir()
    ffmpeg = Path(datos['ffmpeg'])
    if not ffmpeg.is_file():
        ffmpeg = Path(shutil.which(str(ffmpeg)) or '')
    if not ffmpeg.is_file():
        raise RuntimeError('No se encuentra FFmpeg para preparar el archivo.')
    for origen, nombre in ((str(ffmpeg), 'ffmpeg'), (shutil.which('ffprobe'), 'ffprobe')):
        if origen:
            target = herramientas / (nombre + ('.exe' if os.name == 'nt' else ''))
            try: os.link(origen,target)
            except OSError: shutil.copyfile(origen,target)
    def limite(evento):
        total = sum(p.stat().st_size for p in carpeta.iterdir() if p.is_file())
        if total > LIMITE or evento.get('downloaded_bytes',0) > LIMITE:
            raise ValueError('La descarga supera el límite de 500 MB.')
    def filtrar(info, *, incomplete=False):
        if info.get('is_live') or info.get('live_status') == 'is_live':
            return 'No se descargan emisiones en directo.'
        if info.get('duration',0) and info['duration'] > 8*3600:
            return 'El vídeo supera ocho horas.'
    mp3 = datos['modo'] == 'MP3'
    altura = datos['calidad']
    formato = 'bestaudio/best' if mp3 else f'bv[ext=mp4][height<={altura}]+ba[ext=m4a]/b[ext=mp4][height<={altura}]'
    node = shutil.which('node')
    if not node and os.name == 'nt':
        try:
            candidatos = sorted((Path(os.environ.get('ProgramFiles','C:/Program Files')) / 'WindowsApps').glob('OpenAI.CodexPrimaryRuntime.*/dependencies/node/bin/node.exe'), reverse=True)
            node = str(candidatos[0]) if candidatos else None
        except OSError:
            pass
    opciones = dict(outtmpl=str(carpeta/'entrada.%(ext)s'), format=formato, noplaylist=True,
        playlistend=1, max_filesize=LIMITE, socket_timeout=20, retries=2,
        fragment_retries=2, extractor_retries=1, concurrent_fragment_downloads=1, hls_prefer_native=True,
        proxy='', quiet=True, no_warnings=True, logger=Silencio(), progress_hooks=[limite],
        match_filter=filtrar, ffmpeg_location=str(herramientas), merge_output_format='mp4',
        js_runtimes={'node':{'path':node}} if node else {'deno':{}},
        postprocessors=[{'key':'FFmpegExtractAudio','preferredcodec':'mp3','preferredquality':'192'}] if mp3 else [])
    with Descargador(opciones) as ydl:
        info = ydl.extract_info(datos['url'], download=False)
        formatos = info.get('requested_formats') or [info]
        estimado = sum(f.get('filesize') or f.get('filesize_approx') or 0 for f in formatos)
        if estimado > LIMITE:
            raise ValueError('El vídeo supera 500 MB en esta calidad. Elige 360p o 240p, o guarda solo el audio en MP3.')
        ydl.process_info(info)
    if not info or info.get('_type') in ('playlist','multi_video'):
        raise ValueError('Pega el enlace de un único vídeo.')
    destino = carpeta / ('entrada.mp3' if mp3 else 'entrada.mp4')
    if not destino.is_file():
        raise ValueError('La página no ofrece un archivo en el formato elegido.')
    return {'ok':True,'archivo':str(destino),'titulo':info.get('title') or 'video'}

if __name__ == '__main__':
    try:
        print(json.dumps(worker(json.loads(sys.stdin.read())), ensure_ascii=True))
    except Exception as error:
        mensaje = str(error) if type(error) is ValueError else 'No se pudo descargar: comprueba el enlace. Puede requerir acceso, estar restringido o no ofrecer el formato elegido. Prueba una resolución menor o un vídeo diferente.'
        print(json.dumps({'ok':False,'error':mensaje}))
        sys.exit(1)

