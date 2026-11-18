"""Visor compartido de mapas, sin ejecutar instrucciones del contenido."""
import json
import re

def limpiar_mermaid(codigo):
    codigo=str(codigo or '').strip()
    codigo=re.sub(r'^```(?:mermaid)?\s*|\s*```$', '', codigo, flags=re.IGNORECASE).strip()
    codigo=re.sub(r'%%\{.*?\}%%','',codigo,flags=re.DOTALL).strip()
    if not re.match(r'^(?:flowchart|graph)\s+(?:TD|TB|LR|RL|BT)\b',codigo):
        raise ValueError('El resultado no contiene un mapa conceptual Mermaid válido. Se conserva la versión anterior.')
    if len(codigo)>50000:raise ValueError('El mapa es demasiado extenso para el visor. Simplifícalo o divídelo en varios mapas.')
    return codigo

def mental_a_mermaid(markdown):
    lineas=str(markdown or '').expandtabs(2).splitlines();nodos=[];aristas=[];pila=[];nivel_titulo=0
    def etiqueta(t):
        t=re.sub(r'\[([^]]+)\]\([^)]*\)',r'\1',t)
        t=re.sub(r'<[^>]*>','',t).replace('**','').replace('`','').strip()
        return t.replace('&','&amp;').replace('"','&quot;').replace('<','&lt;').replace('>','&gt;')
    for linea in lineas:
        if not linea.strip() or linea.strip().startswith('```'):continue
        titulo=re.match(r'^\s*(#{1,6})\s+(.+)',linea)
        lista=re.match(r'^(\s*)(?:[-*+] |\d+[.)]\s+)(.+)',linea)
        if titulo:nivel=len(titulo[1])-1;nivel_titulo=nivel;texto=titulo[2]
        elif lista:nivel=nivel_titulo+1+len(lista[1])//2;texto=lista[2]
        else:nivel=nivel_titulo+1 if nodos else 0;texto=linea.strip()
        if not nodos:nivel=0
        elif nivel==0:nivel=1
        texto=etiqueta(texto)
        if not texto:continue
        ident='M'+str(len(nodos));nodos.append(f'{ident}["{texto}"]')
        while pila and pila[-1][0]>=nivel:pila.pop()
        if pila:aristas.append(pila[-1][1]+' --> '+ident)
        pila.append((nivel,ident))
        if len(nodos)>180:raise ValueError('El mapa tiene más de 180 conceptos. Usa Simplificar para crear una vista legible; el texto completo se conserva.')
    if len(nodos)<2:raise ValueError('Añade al menos dos conceptos mediante títulos y listas para mostrar el mapa.')
    return 'flowchart LR\n'+'\n'.join(nodos+aristas)

def html_mapa(codigo,direccion='LR',altura=620):
    codigo=limpiar_mermaid(codigo)
    codigo=re.sub(r'^(flowchart|graph)\s+(TD|TB|LR|RL|BT)\b',r'\1 '+('TD' if direccion=='TD' else 'LR'),codigo,count=1)
    seguro=json.dumps(codigo,ensure_ascii=True).replace('<','\\u003c').replace('>','\\u003e').replace('&','\\u0026')
    return _HTML.replace('__CODIGO__',seguro).replace('__ALTURA__',str(int(altura)))

_HTML=r'''<!doctype html><html lang="es"><head><meta charset="utf-8"><style>
*{box-sizing:border-box}body{margin:0;font-family:Arial,sans-serif;color:#17283d;background:#f8fafc}
#visor{border:1px solid #b9ccd8;border-radius:14px;overflow:hidden;background:#fff}
header{display:flex;align-items:center;gap:8px;padding:10px 12px;flex-wrap:wrap;background:#edf5fa;border-bottom:1px solid #cbdbe6}
button{font:600 14px Arial;padding:8px 12px;border:1px solid #9db7cb;border-radius:8px;background:white;color:#17344a;cursor:pointer}button:hover{background:#dcecf9}button:focus-visible{outline:3px solid #2196f3}
#estado{font-size:13px;color:#415b70;margin-left:8px}#zona{height:__ALTURA__px;overflow:auto;cursor:grab;background:radial-gradient(#d9e4ed .7px,transparent .7px);background-size:20px 20px;padding:24px;touch-action:pan-x pan-y}
#zona.arrastrando{cursor:grabbing}#lienzo{width:max-content;min-width:100%;min-height:100%}svg{display:block;max-width:none!important;overflow:visible}
.node rect,.node circle,.node polygon,.node path{fill:#e0f2fe!important;stroke:#4b83aa!important;stroke-width:1.5px!important}
.node:nth-child(3n) rect,.node:nth-child(3n) polygon{fill:#e7e8fc!important;stroke:#7779ad!important}
.node:nth-child(4n) rect,.node:nth-child(4n) polygon{fill:#dcf4e8!important;stroke:#53876d!important}
.nodeLabel,.label,.node text,.node tspan{color:#12283b!important;fill:#12283b!important;font-family:Arial,sans-serif!important;font-size:18px!important}
.edgeLabel,.edgeLabel text,.edgeLabel tspan{font-family:Arial,sans-serif!important;font-size:15px!important;color:#334a60!important;fill:#334a60!important}
.flowchart-link{stroke:#617a90!important;stroke-width:1.8px!important}.edgeLabel rect{fill:#fff!important}
#visor:fullscreen{background:white;width:100vw;height:100vh}#visor:fullscreen #zona{height:calc(100vh - 65px)}
</style></head><body><section id="visor" aria-label="Mapa interactivo">
<header><button id="menos" aria-label="Alejar mapa">−</button><button id="mas" aria-label="Acercar mapa">+</button><button id="leer">Tamaño de lectura</button><button id="ajustar">Ver completo</button><button id="pantalla">Pantalla completa</button><span id="estado" role="status">Preparando mapa…</span></header>
<div id="zona" tabindex="0" aria-label="Mapa; usa las barras o arrastra para desplazarte"><div id="lienzo"></div></div></section>
<script type="module">
const codigo=__CODIGO__;const zona=document.getElementById('zona'),lienzo=document.getElementById('lienzo'),estado=document.getElementById('estado');let svg,escala=1,ancho=1000,alto=600;
function zoom(n){if(!svg)return;const anterior=escala,cx=(zona.scrollLeft+zona.clientWidth/2)/anterior,cy=(zona.scrollTop+zona.clientHeight/2)/anterior;escala=Math.max(.2,Math.min(4,n));svg.style.width=(ancho*escala)+'px';svg.style.height=(alto*escala)+'px';estado.textContent=Math.round(escala*100)+'% · Arrastra para recorrer el mapa';zona.scrollLeft=cx*escala-zona.clientWidth/2;zona.scrollTop=cy*escala-zona.clientHeight/2}
document.getElementById('mas').onclick=()=>zoom(escala*1.25);document.getElementById('menos').onclick=()=>zoom(escala/1.25);
document.getElementById('leer').onclick=()=>{zoom(1);zona.scrollTop=Math.max(0,(alto-zona.clientHeight)/2);zona.scrollLeft=0};
document.getElementById('ajustar').onclick=()=>{zoom(Math.min((zona.clientWidth-48)/ancho,(zona.clientHeight-48)/alto));zona.scrollTop=0;zona.scrollLeft=0};
document.getElementById('pantalla').onclick=async()=>{try{if(document.fullscreenElement)await document.exitFullscreen();else await document.getElementById('visor').requestFullscreen()}catch(e){estado.textContent='Usa el zoom o amplía la ventana del navegador.'}};
if(typeof document.getElementById('visor').requestFullscreen!=='function')document.getElementById('pantalla').hidden=true;
zona.addEventListener('wheel',e=>{if(e.ctrlKey){e.preventDefault();zoom(escala*(e.deltaY<0?1.1:1/1.1))}},{passive:false});
let arrastre=null;zona.addEventListener('pointerdown',e=>{if(e.button!==0||e.pointerType==='touch')return;arrastre={x:e.clientX,y:e.clientY,l:zona.scrollLeft,t:zona.scrollTop};zona.setPointerCapture(e.pointerId);zona.classList.add('arrastrando')});
zona.addEventListener('pointermove',e=>{if(arrastre){zona.scrollLeft=arrastre.l-(e.clientX-arrastre.x);zona.scrollTop=arrastre.t-(e.clientY-arrastre.y)}});zona.addEventListener('pointerup',()=>{arrastre=null;zona.classList.remove('arrastrando')});zona.addEventListener('pointercancel',()=>{arrastre=null;zona.classList.remove('arrastrando')});
try{const {default:mermaid}=await import('https://cdn.jsdelivr.net/npm/mermaid@11.4.1/dist/mermaid.esm.min.mjs');mermaid.initialize({startOnLoad:false,securityLevel:'strict',htmlLabels:false,theme:'base',fontFamily:'Arial',themeVariables:{fontSize:'18px',primaryTextColor:'#12283b',lineColor:'#617a90'},flowchart:{htmlLabels:false,useMaxWidth:false,nodeSpacing:32,rankSpacing:75,wrappingWidth:220,curve:'basis'}});const resultado=await mermaid.render('mapa-svg',codigo);lienzo.innerHTML=resultado.svg;svg=lienzo.querySelector('svg');svg.setAttribute('role','img');svg.setAttribute('aria-label','Mapa de conceptos y relaciones');const v=svg.viewBox.baseVal;ancho=v.width||1000;alto=v.height||600;zoom(1);zona.scrollTop=Math.max(0,(alto-zona.clientHeight)/2);zona.scrollLeft=0}catch(e){estado.textContent='No se pudo dibujar el mapa. Revisa el código o usa Simplificar; comprueba también la conexión a Internet.';lienzo.textContent='El contenido original se conserva en el editor del mapa.'}
</script></body></html>'''

