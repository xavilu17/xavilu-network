"""Selecciones explícitas del material, sin modificar el original ni usar IA."""
from io import BytesIO
from pathlib import Path
import re
import subprocess
import tempfile
import shutil
import json

def paginas(expresion,total):
    salida=[]
    if not str(expresion).strip():raise ValueError('Indica páginas: por ejemplo, 3-8, 12, 15-20.')
    for parte in str(expresion).replace('–','-').replace('—','-').split(','):
        m=re.fullmatch(r'\s*(\d+)\s*(?:-\s*(\d+))?\s*',parte)
        if not m:raise ValueError('Usa números y rangos separados por comas: 3-8, 12.')
        a=int(m[1]);b=int(m[2] or a)
        if not 1<=a<=b<=total:raise ValueError(f'El rango {parte.strip()} no es válido: el documento tiene {total} páginas.')
        for n in range(a,b+1):
            if n not in salida:salida.append(n)
    return salida

def texto_local(datos,nombre):
    ext=Path(nombre).suffix.lower()
    if ext=='.pdf':
        from pypdf import PdfReader
        r=PdfReader(BytesIO(datos))
        textos=[p.extract_text() or '' for p in r.pages]
        texto='\n\n'.join(f'[Página original {i}]\n{t}' for i,t in enumerate(textos,1))
        if not any(t.strip() for t in textos):raise ValueError('Este PDF no contiene texto seleccionable. Usa páginas concretas o realiza OCR primero.')
        return texto
    if ext=='.docx':
        from docx import Document
        doc=Document(BytesIO(datos))
        return '\n'.join([p.text for p in doc.paragraphs]+[' | '.join(c.text for c in row.cells) for t in doc.tables for row in t.rows])
    if ext=='.pptx':
        from pptx import Presentation
        return '\n\n'.join(f'[Diapositiva {i}]\n'+'\n'.join(s.text for s in slide.shapes if s.has_text_frame) for i,slide in enumerate(Presentation(BytesIO(datos)).slides,1))
    if ext in ('.txt','.csv','.md'):
        try:return datos.decode('utf-8-sig')
        except UnicodeDecodeError:return datos.decode('cp1252')
    raise ValueError('Para seleccionar texto en este formato, conviértelo primero a PDF, DOCX o TXT. El archivo completo sigue disponible.')

def _coincidencia(texto,frase):
    if not frase.strip():raise ValueError('Escribe el texto o título que delimita la selección.')
    patron=r'\s+'.join(re.escape(x) for x in frase.split())
    coincidencias=list(re.finditer(patron,texto,re.IGNORECASE))
    if len(coincidencias)!=1:raise ValueError('El fragmento no aparece o aparece varias veces. Copia una frase más concreta del documento.')
    return coincidencias[0]

def recortar(datos,nombre,modo,**op):
    ext=Path(nombre).suffix.lower();base=Path(nombre).stem
    if modo=='Diapositivas':
        from pptx import Presentation
        doc=Presentation(BytesIO(datos));seleccion=paginas(op.get('paginas',''),len(doc.slides))
        ids=list(doc.slides._sldIdLst)
        for i,ident in enumerate(ids,1):
            if i not in seleccion:doc.part.drop_rel(ident.rId)
            doc.slides._sldIdLst.remove(ident)
        for n in seleccion:doc.slides._sldIdLst.append(ids[n-1])
        out=BytesIO();doc.save(out)
        return {'data':out.getvalue(),'name':base+' [diapositivas '+op['paginas']+'].pptx','type':'application/vnd.openxmlformats-officedocument.presentationml.presentation','detalle':f'{len(seleccion)} diapositivas · originales {op["paginas"]}'}
    if modo=='Hoja y rango de celdas':
        from openpyxl import load_workbook
        from openpyxl.utils.cell import range_boundaries,get_column_letter
        rango=op.get('rango','').upper().strip()
        if not re.fullmatch(r'[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?',rango):raise ValueError('Indica un rango de celdas, por ejemplo A1:F20.')
        a,b,c,d=range_boundaries(rango)
        if c<a or d<b or c>16384 or d>1048576 or (c-a+1)*(d-b+1)>10000:raise ValueError('El rango debe contener como máximo 10.000 celdas y ser válido.')
        doc=load_workbook(BytesIO(datos),read_only=True,data_only=False)
        try:
            if op.get('hoja') not in doc.sheetnames:raise ValueError('Elige una hoja del archivo.')
            ws=doc[op['hoja']]
            if a>ws.max_column or b>ws.max_row:raise ValueError('El rango queda fuera de las celdas con datos.')
            filas=[]
            for i,row in enumerate(ws.iter_rows(min_row=b,max_row=d,min_col=a,max_col=c),b):
                filas.append(' | '.join(f'{get_column_letter(j)}{i}: {cell.value if cell.value is not None else ""}' for j,cell in enumerate(row,a)))
            texto=f'Fuente: {nombre}\nHoja: {op["hoja"]} · Rango original: {rango}\nLas fórmulas se muestran literalmente; no se han ejecutado ni calculado.\n\n'+'\n'.join(filas)
        finally:doc.close()
        return {'data':texto.encode('utf-8'),'name':base+' [celdas '+rango.replace(':','-')+'].txt','type':'text/plain','detalle':f'Hoja {op["hoja"]} · {rango} · texto de las celdas, sin ejecutar fórmulas'}
    if modo=='Páginas del PDF':
        from pypdf import PdfReader,PdfWriter
        reader=PdfReader(BytesIO(datos));seleccion=paginas(op.get('paginas',''),len(reader.pages));writer=PdfWriter()
        for n in seleccion:writer.add_page(reader.pages[n-1])
        out=BytesIO();writer.write(out)
        return {'data':out.getvalue(),'name':f'{base} [paginas originales {op["paginas"]}].pdf','type':'application/pdf','detalle':f'{len(seleccion)} de {len(reader.pages)} páginas · originales {op["paginas"]}'}
    if modo in ('Fragmento de texto','Entre dos frases o títulos'):
        texto=texto_local(datos,nombre)
        if modo=='Fragmento de texto':
            m=_coincidencia(texto,op.get('fragmento',''));parte=texto[m.start():m.end()]
        else:
            a=_coincidencia(texto,op.get('inicio',''));b=_coincidencia(texto,op.get('fin',''))
            if b.start()<a.start():raise ValueError('El final debe estar después del inicio.')
            parte=texto[a.start():b.end()]
        return {'data':(f'Fuente: {nombre}\nSelección: {modo}\n\n'+parte).encode('utf-8'),'name':base+' [fragmento].txt','type':'text/plain','detalle':f'Fragmento verificado · {len(parte)} caracteres'}
    if modo=='Zona de la imagen':
        from PIL import Image,ImageOps
        if len(datos)>25*1024*1024:raise ValueError('La imagen supera el límite de 25 MB.')
        with Image.open(BytesIO(datos)) as im:
            if im.width*im.height>40_000_000:raise ValueError('La imagen tiene más de 40 megapíxeles; reduce su tamaño antes de recortarla.')
            im=ImageOps.exif_transpose(im);x,y,w,h=[int(op[k]) for k in ('x','y','ancho','alto')]
            if min(x,y)<0 or min(w,h)<=0 or x+w>im.width or y+h>im.height:raise ValueError('El recuadro debe quedar dentro de la imagen.')
            out=BytesIO();im.crop((x,y,x+w,y+h)).save(out,format='PNG')
        return {'data':out.getvalue(),'name':base+' [zona].png','type':'image/png','detalle':f'Zona {w} × {h} píxeles'}
    if modo=='Tramo de audio o vídeo':
        inicio=float(op['desde']);fin=float(op['hasta'])
        if inicio<0 or fin<=inicio:raise ValueError('El final debe ser mayor que el inicio, en segundos.')
        video=ext in ('.mp4','.mov','.avi','.mkv','.webm');sufijo='.mp4' if video else '.mp3'
        with tempfile.TemporaryDirectory(prefix='xavilu_tramo_') as carpeta:
            entrada=Path(carpeta)/('original'+ext);salida=Path(carpeta)/('tramo'+sufijo);entrada.write_bytes(datos)
            if shutil.which('ffprobe'):
                info=subprocess.run([shutil.which('ffprobe'),'-v','error','-show_entries','format=duration','-of','json',str(entrada)],check=True,capture_output=True,text=True,timeout=20)
                duracion=float(json.loads(info.stdout)['format']['duration'])
                if inicio>=duracion or fin>duracion+0.05:raise ValueError(f'El archivo dura {duracion:.1f} segundos. Elige un tramo dentro de esa duración.')
            ffmpeg=str(op.get('ffmpeg') or '')
            ejecutable=Path(ffmpeg)
            if not ejecutable.is_file():
                localizado=shutil.which(ffmpeg)
                if not localizado:raise ValueError('No se encontró FFmpeg. Comprueba la ruta en la configuración local.')
                ejecutable=Path(localizado)
            cmd=[str(ejecutable),'-nostdin','-y','-i',str(entrada),'-ss',str(inicio),'-t',str(fin-inicio)]
            cmd+=['-map','0:v:0','-map','0:a:0?','-c:v','libx264','-crf','23','-c:a','aac'] if video else ['-vn','-c:a','libmp3lame','-b:a','192k']
            try:subprocess.run(cmd+[str(salida)],check=True,stdout=subprocess.DEVNULL,stderr=subprocess.PIPE,timeout=300)
            except subprocess.TimeoutExpired:raise ValueError('Preparar el tramo ha superado cinco minutos. Elige uno más corto.') from None
            except subprocess.CalledProcessError:raise ValueError('No se pudo extraer ese tramo del archivo.') from None
            if not salida.is_file() or salida.stat().st_size<100:raise ValueError('No hay contenido en ese tramo. Comprueba la duración del archivo.')
            resultado=salida.read_bytes()
        return {'data':resultado,'name':f'{base} [segundos {inicio:g}-{fin:g}]'+sufijo,'type':'video/mp4' if video else 'audio/mpeg','detalle':f'Tramo {inicio:g}–{fin:g} segundos'}
    raise ValueError('Selección de material no válida.')
