"""Generación explícita de imágenes con proveedores actuales; una petición, sin reintentos de pago."""
from io import BytesIO
from urllib.parse import quote
import base64,json
from PIL import Image,ImageOps

MODELOS_GOOGLE={'Equilibrada · Nano Banana 2.1':'gemini-nano-banana-2.1','Rápida · Nano Banana 2 Lite':'gemini-3.1-flash-lite-image','Detallada · Nano Banana Pro':'gemini-3-pro-image'}
RATIOS={'1:1','16:9','9:16','4:3','3:4','3:2','2:3'}


class ErrorImagen(Exception):pass


def normalizar_imagen(datos,limite=25*1024*1024):
    if not isinstance(datos,(bytes,bytearray)) or not datos or len(datos)>limite:
        raise ErrorImagen('La imagen está vacía o supera el límite de tamaño.')
    try:
        with Image.open(BytesIO(datos)) as image:
            if image.width*image.height>36_000_000:raise ErrorImagen('La imagen supera el límite de resolución.')
            image=ImageOps.exif_transpose(image)
            image=image.convert('RGBA' if 'A'in image.getbands() else 'RGB')
            salida=BytesIO();image.save(salida,format='PNG')
            return salida.getvalue(),image.size
    except ErrorImagen:raise
    except Exception:raise ErrorImagen('El contenido recibido no es una imagen compatible.') from None


def _leer_respuesta(response):
    if response.status_code in (401,403):raise ErrorImagen('El proveedor no ha autorizado la generación. Comprueba la clave, el acceso al modelo y la facturación de tu API.')
    if response.status_code==429:raise ErrorImagen('Se ha alcanzado una cuota o límite del proveedor. Revisa tu cuota antes de volver a intentarlo.')
    if response.status_code!=200:raise ErrorImagen('El proveedor no pudo generar la imagen (HTTP '+str(response.status_code)+'). No se ha intentado otro modelo automáticamente.')
    data=bytearray()
    for chunk in response.iter_content(65536):
        data.extend(chunk)
        if len(data)>40*1024*1024:raise ErrorImagen('La respuesta del proveedor supera el tamaño permitido.')
    return bytes(data)


def generar_imagen_google(prompt,clave,modelo,ratio='1:1',referencia=None,http=None):
    if not str(clave).strip():raise ErrorImagen('Configura tu clave de Gemini para generar imágenes con Google.')
    if modelo not in MODELOS_GOOGLE.values() or ratio not in RATIOS:raise ErrorImagen('Modelo o formato de imagen no válido.')
    if not isinstance(prompt,str) or not 1<=len(prompt.strip())<=10000:raise ErrorImagen('Escribe una descripción de hasta 10.000 caracteres.')
    parts=[{'text':prompt.strip()}]
    if referencia:
        png,size=normalizar_imagen(referencia,10*1024*1024)
        parts.append({'inlineData':{'mimeType':'image/png','data':base64.b64encode(png).decode('ascii')}})
    payload={'contents':[{'role':'user','parts':parts}],'generationConfig':{'responseModalities':['TEXT','IMAGE'],'imageConfig':{'aspectRatio':ratio}}}
    if http is None:
        import requests
        http=requests
    try:
        with http.post('https://generativelanguage.googleapis.com/v1beta/models/'+modelo+':generateContent',headers={'x-goog-api-key':clave.strip(),'Content-Type':'application/json'},json=payload,timeout=(10,120),allow_redirects=False,stream=True) as response:
            raw=json.loads(_leer_respuesta(response))
        for candidate in raw.get('candidates',[]):
            for part in candidate.get('content',{}).get('parts',[]):
                inline=part.get('inlineData')or part.get('inline_data')
                if inline and str(inline.get('mimeType',inline.get('mime_type',''))).startswith('image/'):
                    try:data=base64.b64decode(inline['data'],validate=True)
                    except (KeyError,ValueError):raise ErrorImagen('El proveedor devolvió una imagen no válida.') from None
                    png,size=normalizar_imagen(data)
                    return {'data':png,'ancho':size[0],'alto':size[1],'proveedor':'Gemini','modelo':modelo}
        raise ErrorImagen('Google no devolvió una imagen. Puede haber rechazado la petición o generado solo texto; modifica la descripción antes de volver a intentarlo.')
    except ErrorImagen:raise
    except Exception:raise ErrorImagen('No se pudo completar la generación con Google. La petición no se ha repetido automáticamente.') from None


def generar_imagen_pollinations(prompt,clave,ratio='1:1',http=None):
    if not str(clave).strip():raise ErrorImagen('La API actual de Pollinations requiere una clave propia.')
    if ratio not in RATIOS or not isinstance(prompt,str) or not 1<=len(prompt.strip())<=10000:raise ErrorImagen('Descripción o formato de imagen no válido.')
    w,h=map(int,ratio.split(':'));ancho=1024 if w>=h else round(1024*w/h);alto=1024 if h>=w else round(1024*h/w)
    if http is None:
        import requests
        http=requests
    try:
        with http.get('https://gen.pollinations.ai/image/'+quote(prompt.strip(),safe=''),headers={'Authorization':'Bearer '+clave.strip()},params={'width':ancho,'height':alto},timeout=(10,120),allow_redirects=False,stream=True) as response:
            data=_leer_respuesta(response)
        png,size=normalizar_imagen(data)
        return {'data':png,'ancho':size[0],'alto':size[1],'proveedor':'Pollinations','modelo':'Predeterminado del proveedor'}
    except ErrorImagen:raise
    except Exception:raise ErrorImagen('No se pudo completar la generación con Pollinations. La petición no se ha repetido automáticamente.') from None
