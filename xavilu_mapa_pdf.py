"""Recibe el SVG del visor y lo guarda como PDF vectorial A4, sin servicios externos."""
from io import BytesIO
import math
import re
import xml.etree.ElementTree as ET
import hashlib
import streamlit as st
API_VERSION=3

_JS=r'''
const instancias=new WeakMap();
export default function(component){
  const {parentElement,data,setStateValue}=component;
  const iframe=parentElement.querySelector('iframe');if(!iframe||!data)return;
  const anterior=instancias.get(parentElement);
  if(anterior?.token===data.token)return;
  if(anterior?.handler)window.removeEventListener('message',anterior.handler);
  const handler=(evento)=>{
    if(evento.source!==iframe.contentWindow||evento.data?.token!==data.token)return;
    if(evento.data.tipo==='xavilu-svg'&&typeof evento.data.svg==='string'&&evento.data.svg.length<2500000){
      setStateValue('exportacion',{token:data.token,svg:evento.data.svg});
    }
    if(evento.data.tipo==='xavilu-error')setStateValue('exportacion',{token:data.token,error:'No se pudo preparar el dibujo. Comprueba la conexión o revisa el mapa.'});
  };
  window.addEventListener('message',handler);instancias.set(parentElement,{token:data.token,handler});
  iframe.allow='fullscreen';iframe.allowFullscreen=true;
  iframe.srcdoc=data.html;
}
'''
_VISOR=st.components.v2.component('xavilu_visormapa_pdf',html='<iframe title="Mapa interactivo con exportación PDF" style="width:100%;height:730px;border:0" sandbox="allow-scripts allow-downloads"></iframe>',js=_JS)

def visor_pdf(codigo,direccion,clave):
    import importlib,xavilu_mapas
    html_mapa=importlib.reload(xavilu_mapas).html_mapa
    html=html_mapa(codigo,direccion)
    insertar=r'''
const copia=svg.cloneNode(true),originales=[svg,...svg.querySelectorAll('*')],copias=[copia,...copia.querySelectorAll('*')];
for(let i=0;i<originales.length;i++){
  const estilo=getComputedStyle(originales[i]);
  for(const prop of ['fill','fill-opacity','stroke','stroke-width','stroke-opacity','font-size','font-family','font-weight','text-anchor','dominant-baseline','opacity'])copias[i].style.setProperty(prop,estilo.getPropertyValue(prop));
}
// Flatten nested SVG text into explicit baselines for faithful PDF rendering.
for(let i=0;i<originales.length;i++){
  const origen=originales[i],destino=copias[i];
  if(origen.localName==='text'){
    const grupo=document.createElementNS('http://www.w3.org/2000/svg','g');
    if(origen.hasAttribute('transform'))grupo.setAttribute('transform',origen.getAttribute('transform'));
    const lineas=[...origen.children].filter(e=>e.localName==='tspan');let posicion=0;
    for(const linea of (lineas.length?lineas:[origen])){
      const texto=linea.textContent||'';if(!texto)continue;
      const punto=origen.getStartPositionOfChar(posicion),estilo=getComputedStyle(linea);
      const t=document.createElementNS('http://www.w3.org/2000/svg','text');
      t.setAttribute('x',punto.x);t.setAttribute('y',punto.y);t.setAttribute('text-anchor','start');
      t.setAttribute('font-family','Arial');t.setAttribute('font-size',parseFloat(estilo.fontSize)||18);t.setAttribute('font-weight',estilo.fontWeight);t.setAttribute('fill',estilo.fill);t.setAttribute('stroke','none');t.setAttribute('xml:space','preserve');t.textContent=texto;
      grupo.appendChild(t);posicion+=texto.length;
    }
    destino.replaceWith(grupo);
  }
  if(origen.localName==='path'){
    const largo=origen.getTotalLength();
    for(const extremo of ['end','start']){
      const marca=origen.getAttribute('marker-'+extremo)||'';
      if(!marca||!marca.toLowerCase().includes('point'))continue;
      const p=origen.getPointAtLength(extremo==='end'?largo:0),q=origen.getPointAtLength(extremo==='end'?Math.max(0,largo-4):Math.min(4,largo));
      const a=Math.atan2(p.y-q.y,p.x-q.x),ux=Math.cos(a),uy=Math.sin(a);
      const triangulo=document.createElementNS('http://www.w3.org/2000/svg','polygon');
      triangulo.setAttribute('points',`${p.x},${p.y} ${p.x-9*ux+4*uy},${p.y-9*uy-4*ux} ${p.x-9*ux-4*uy},${p.y-9*uy+4*ux}`);
      triangulo.setAttribute('fill',getComputedStyle(origen).stroke||'#617a90');
      if(origen.hasAttribute('transform'))triangulo.setAttribute('transform',origen.getAttribute('transform'));
      destino.parentNode.appendChild(triangulo);destino.removeAttribute('marker-'+extremo);
    }
  }
}
copia.setAttribute('xmlns','http://www.w3.org/2000/svg');copia.setAttribute('width',ancho);copia.setAttribute('height',alto);copia.style.width=ancho+'px';copia.style.height=alto+'px';
window.parent.postMessage({tipo:'xavilu-svg',token:__TOKEN__,svg:new XMLSerializer().serializeToString(copia)},'*');
'''
    import json
    token=hashlib.sha256((html+insertar).encode()).hexdigest()
    insertar=insertar.replace('__TOKEN__',json.dumps(token))
    marker="}catch(e){estado.textContent='No se pudo dibujar el mapa."
    assert html.count(marker)==1
    html=html.replace(marker,insertar+marker,1)
    html=html.replace("lienzo.textContent='El contenido original se conserva en el editor del mapa.'", "window.parent.postMessage({tipo:'xavilu-error',token:"+json.dumps(token)+"},'*');lienzo.textContent='El contenido original se conserva en el editor del mapa.'",1)
    resultado=_VISOR(data={'token':token,'html':html},key='visor_pdf_'+clave,height=750,on_exportacion_change=lambda:None)
    exportacion=resultado.exportacion
    return exportacion if isinstance(exportacion,dict) and exportacion.get('token')==token else None

def _svg_seguro(svg):
    if not isinstance(svg,str) or len(svg)>2500000 or re.search(r'<!DOCTYPE|<!ENTITY',svg,re.I):raise ValueError('El dibujo no es un SVG válido para exportar.')
    root=ET.fromstring(svg)
    if root.tag.split('}')[-1]!='svg':raise ValueError('Se esperaba un dibujo SVG.')
    for parent in root.iter():
        for child in list(parent):
            if child.tag.split('}')[-1]=='style':parent.remove(child)
        tag=parent.tag.split('}')[-1]
        if tag in ('script','foreignObject','image','iframe','object','embed'):raise ValueError('El dibujo contiene elementos que no se pueden exportar.')
        for k,v in parent.attrib.items():
            local=k.split('}')[-1].lower()
            if local.startswith('on') or (local in ('href','src') and not v.startswith('#')):raise ValueError('El dibujo contiene referencias externas no permitidas.')
            for referencia in re.findall(r'url\((.*?)\)',v,re.I):
                if not re.fullmatch(r'[\s\'"]*#[\w.:-]+[\s\'"]*',referencia):raise ValueError('No se permiten recursos externos en el dibujo.')
    ET.register_namespace('','http://www.w3.org/2000/svg')
    return ET.tostring(root,encoding='utf-8')

def generar_pdf_mapa(svg,titulo='Mapa',orientacion='A4 horizontal'):
    from svglib.svglib import svg2rlg
    from reportlab.graphics import renderPDF
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4,landscape
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from pathlib import Path
    import reportlab
    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        candidatos=[Path('C:/Windows/Fonts/arial.ttf'),Path(reportlab.__file__).parent/'fonts'/'Vera.ttf']
        fuente=next((p for p in candidatos if p.is_file()),None)
        if fuente:pdfmetrics.registerFont(TTFont('Arial',str(fuente)))
    dibujo=svg2rlg(BytesIO(_svg_seguro(svg)))
    if dibujo is None or not 0<dibujo.width<50000 or not 0<dibujo.height<50000:raise ValueError('No se pudo preparar el dibujo del mapa.')
    pagina=landscape(A4) if orientacion=='A4 horizontal' else A4
    ancho,alto=pagina;margen=28;w=ancho-2*margen;h=alto-90
    escala=min(w/dibujo.width,h/dibujo.height)
    detalle=escala<0.65
    escala_detalle=0.75;pasox=w-20;pasoy=h-20
    columnas=max(1,math.ceil((dibujo.width*escala_detalle-20)/pasox));filas=max(1,math.ceil((dibujo.height*escala_detalle-20)/pasoy))
    if detalle and columnas*filas>30:raise ValueError('El mapa es demasiado extenso para una exportación legible. Simplifícalo o divídelo en partes.')
    out=BytesIO();c=canvas.Canvas(out,pagesize=pagina);c.setTitle(titulo);c.setAuthor('Xavilu Network')
    total=1+(columnas*filas if detalle else 0)
    def cabecera(subtitulo,numero):
        c.setFillColorRGB(.07,.16,.24);c.setFont('Helvetica-Bold',15);c.drawString(margen,alto-27,str(titulo)[:95])
        c.setFont('Helvetica',9);c.setFillColorRGB(.30,.39,.46);c.drawString(margen,alto-43,subtitulo)
        c.drawString(margen,16,'Xavilu Network · Mapa visual A4');c.drawRightString(ancho-margen,16,f'{numero} / {total}')
    cabecera('Vista completa · las páginas siguientes ofrecen detalle legible.' if detalle else 'Vista completa del mapa',1)
    c.saveState();c.translate(margen+(w-dibujo.width*escala)/2,30+(h-dibujo.height*escala)/2);c.scale(escala,escala);renderPDF.draw(dibujo,c,0,0);c.restoreState();c.showPage()
    if detalle:
        numero=2
        for fila in range(filas):
            for columna in range(columnas):
                cabecera(f'Detalle · fila {fila+1} de {filas} · columna {columna+1} de {columnas} · vista con solape entre páginas',numero)
                c.saveState();p=c.beginPath();p.rect(margen,30,w,h);c.clipPath(p,stroke=0)
                c.translate(margen-columna*pasox,30+h-dibujo.height*escala_detalle+fila*pasoy);c.scale(escala_detalle,escala_detalle);renderPDF.draw(dibujo,c,0,0);c.restoreState();c.showPage();numero+=1
    c.save();return out.getvalue()
