"""Mesa de estudio para continuar trabajando con documentos y resultados."""
import json, re

API_VERSION=2
ACCIONES={
 'explicar':'Explicar','resumir':'Resumir','mapa':'Mapa o esquema','extraer':'Extraer preguntas',
 'crear':'Crear preguntas','resolver':'Resolver preguntas','tarjetas':'Tarjetas','guion':'Crear guion dialogado',
 'test':'Test interactivo','ampliar':'Ampliar explicación','tabla':'Crear tabla','glosario':'Glosario',
 'cronologia':'Cronología','audio':'Audio o podcast','editar':'Editar resultado'}


def iniciar(texto,titulo,tipo='texto'):
 return {'versiones':[{'texto':str(texto or ''),'titulo':str(titulo or 'Material'),'tipo':tipo}],'indice':0}


def actual(estado):
 versiones=estado.get('versiones') or [{}]
 indice=max(0,min(int(estado.get('indice',len(versiones)-1)),len(versiones)-1))
 return versiones[indice]


def prompt_estudio(accion,material,notas='',cantidad=8,dificultad='Media',modo='Nivel medio'):
 tareas={
  'explicar':'explica el contenido con ejemplos claros','resumir':'resume las ideas esenciales y conserva los matices importantes',
  'mapa':'crea un mapa jerárquico en Markdown limpio, con un tema central y ramas breves',
  'extraer':'extrae las preguntas existentes en la fuente; no inventes respuestas',
  'crear':'crea preguntas nuevas de opción múltiple A-D con respuesta y explicación',
  'resolver':'resuelve las preguntas de la fuente y justifica cada respuesta',
  'tarjetas':'crea tarjetas de estudio de pregunta y respuesta',
  'guion':'crea un guion dialogado entre Locutor 1 y Locutor 2, natural y riguroso',
  'test':'crea un test de opción múltiple autocontenido',
  'ampliar':'amplía y aclara el contenido con ejemplos y relaciones entre conceptos',
  'tabla':'organiza la información en una tabla Markdown', 'glosario':'crea un glosario de términos y definiciones',
  'cronologia':'ordena los hechos, procesos y fechas cronológicamente'}
 formato=''
 if accion=='test':formato='Devuelve solo JSON válido: {"preguntas":[{"pregunta":"...","opciones":{"A":"...","B":"...","C":"...","D":"..."},"correcta":"A","explicacion":"...","referencia":"..."}]}'
 if accion=='tarjetas':formato='Devuelve JSON válido: {"tarjetas":[{"pregunta":"...","respuesta":"...","referencia":"..."}]}'
 return f'''Trabaja solo con el material aportado. No inventes hechos, citas ni páginas. Escribe en español claro, nivel {modo}, y explica tecnicismos con palabras sencillas. Tarea: {tareas.get(accion,'trabaja el material con rigor')}. Dificultad: {dificultad}. Cantidad orientativa: {cantidad}. Indicaciones del usuario: {notas or 'ninguna'}. {formato}
MATERIAL:
{material}'''


def _objeto_json(texto):
 bruto=re.sub(r'^```(?:json)?\s*|\s*```$','',str(texto or '').strip(),flags=re.I)
 try:return json.loads(bruto)
 except json.JSONDecodeError:
  ini=bruto.find('{');fin=bruto.rfind('}')
  if ini<0 or fin<=ini:raise ValueError('La IA no devolvió JSON legible.')
  return json.loads(bruto[ini:fin+1])


def validar_test(texto,material='',cantidad=10):
 datos=_objeto_json(texto);preguntas=datos.get('preguntas',[]) if isinstance(datos,dict) else []
 if not isinstance(preguntas,list) or not preguntas:raise ValueError('No se encontraron preguntas en el test.')
 salida=[];vistos=set()
 for q in preguntas[:max(1,int(cantidad))]:
  if not isinstance(q,dict):continue
  enunciado=str(q.get('pregunta',q.get('enunciado',''))).strip();op=q.get('opciones') or {}
  correcta=str(q.get('correcta',q.get('respuesta',''))).strip().upper()
  if not enunciado or enunciado.casefold() in vistos or not isinstance(op,dict):continue
  op={k:str(op.get(k,'')).strip() for k in 'ABCD'}
  if correcta not in op or not all(op.values()):continue
  vistos.add(enunciado.casefold())
  salida.append({'pregunta':enunciado,'opciones':op,'correcta':correcta,'explicacion':str(q.get('explicacion','')).strip(),'referencia':str(q.get('referencia',q.get('cita',''))).strip()})
 if not salida:raise ValueError('El test no tiene preguntas completas con cuatro opciones y respuesta correcta.')
 return salida


def texto_test(preguntas):
 partes=[]
 for i,q in enumerate(preguntas or [],1):
  opciones=q.get('opciones') or {}
  partes.append(f"{i}. {q.get('pregunta',q.get('enunciado',''))}\n"+'\n'.join(f'{k}) {v}' for k,v in opciones.items())+f"\nRespuesta: {q.get('correcta',q.get('respuesta',''))}\n{q.get('explicacion','')}\nFuente: {q.get('referencia',q.get('cita',''))}")
 return '\n\n'.join(partes)


def _mostrar(st,nodo,clave,mapa):
 st.markdown('### '+str(nodo.get('titulo','Resultado')))
 if nodo.get('tipo')=='audio' and nodo.get('audio'):st.audio(nodo['audio'])
 elif nodo.get('tipo')=='mapa':mapa(str(nodo.get('texto','')),clave+'_mapa_'+str(len(nodo.get('texto',''))))
 else:st.markdown(str(nodo.get('texto','')))
 if nodo.get('preguntas'):
  with st.expander('Preguntas incluidas en este resultado'):
   for i,q in enumerate(nodo['preguntas'],1):
    st.markdown(f"**{i}. {q.get('pregunta','')}**")
    st.write(' · '.join(f'{k}) {v}' for k,v in q.get('opciones',{}).items()))
    st.caption('Respuesta: '+str(q.get('correcta',''))+' · '+str(q.get('explicacion','')))


def dibujar(estado,clave,ejecutar,leer,guardar,mapa,hay_archivos=False):
 import hashlib
 import streamlit as st
 versiones=estado.get('versiones',[])
 if not versiones:return
 etiquetas=[f"{i+1}. {v.get('titulo','Resultado')}" for i,v in enumerate(versiones)]
 elegido=st.selectbox('Historial de trabajo',range(len(versiones)),index=max(0,min(estado.get('indice',len(versiones)-1),len(versiones)-1)),format_func=lambda i:etiquetas[i],key=clave+'_historial')
 estado['indice']=elegido;nodo=actual(estado)
 with st.container(border=True):_mostrar(st,nodo,clave,mapa)
 subida=st.file_uploader('Cambiar o añadir documento para seguir trabajando',type=['pdf','txt','md','doc','docx','xls','xlsx','ppt','pptx','png','jpg','jpeg','webp','mp3','wav','m4a','mp4','mov','avi','mkv'],key=clave+'_nuevo_documento')
 if subida:
  firma=hashlib.sha256(subida.name.encode()+subida.getvalue()).hexdigest()
  if st.session_state.get(clave+'_documento_procesado')!=firma:
   try:
    with st.spinner('Leyendo el nuevo documento…'):nuevo=str(leer(subida) or '')
    if not nuevo.strip():raise ValueError('No se pudo extraer contenido de ese archivo.')
    versiones.append({'texto':nuevo,'titulo':subida.name,'tipo':'texto'});estado['indice']=len(versiones)-1
    st.session_state[clave+'_documento_procesado']=firma;st.rerun()
   except Exception as exc:st.warning('No se pudo usar el archivo: '+str(exc))
 st.markdown('#### ¿Qué quieres hacer ahora?')
 columnas=st.columns(3)
 for i,(codigo,label) in enumerate(ACCIONES.items()):
  with columnas[i%3]:
   if st.button(label,key=f'{clave}_accion_{codigo}',width='stretch'):st.session_state[clave+'_accion']=codigo
 accion=st.session_state.get(clave+'_accion','resumir')
 if accion not in ACCIONES:accion='resumir'
 with st.form(clave+'_seguir_form'):
  fuente=st.radio('Material de trabajo',['Resultado actual','Documento original'],horizontal=True,key=clave+'_fuente_tipo')
  material=str(nodo.get('texto','')) if fuente=='Resultado actual' else str(versiones[0].get('texto',''))
  notas=st.text_area('Indicación para esta tarea',placeholder='Indica la parte, páginas, extensión, dificultad o enfoque.',key=clave+'_notas')
  c1,c2,c3=st.columns(3)
  cantidad=c1.number_input('Número de elementos',min_value=1,max_value=50,value=8,key=clave+'_cantidad')
  dificultad=c2.selectbox('Dificultad',['Fácil','Media','Alta'],index=1,key=clave+'_dificultad')
  modo=c3.selectbox('Nivel de explicación',['Nivel medio','Sencillo','Técnico'],key=clave+'_modo')
  editado=st.text_area('Contenido editable',value=material,key=clave+'_editado',height=100)
  audio_modo=st.selectbox('Qué hacer con el audio',['Crear guion dialogado','Producir podcast del guion','Leer en voz alta'],key=clave+'_audio_modo') if accion=='audio' else 'Crear guion dialogado'
  ejecutar_paso=st.form_submit_button('▶ Ejecutar: '+ACCIONES[accion],type='primary',width='stretch')
 if ejecutar_paso:
  base=editado if accion=='editar' else material
  if not base.strip():st.warning('El material está vacío; selecciona otro documento o resultado.')
  else:
   try:
    with st.spinner('Trabajando con el material…'):salida=ejecutar(accion,base,notas,int(cantidad),dificultad,modo,audio_modo,editado)
    if not isinstance(salida,dict) or not str(salida.get('texto','')).strip():raise ValueError('La tarea no devolvió un resultado con contenido.')
    versiones.append(salida);estado['indice']=len(versiones)-1;st.session_state[clave+'_accion']='resumir';st.rerun()
   except Exception as exc:st.error('No se pudo completar esta acción: '+str(exc))
 if len(versiones)>1 and st.button('↩️ Volver al resultado anterior',key=clave+'_anterior'):
  estado['indice']=max(0,estado['indice']-1);st.rerun()
 st.markdown('#### Guardar si quieres conservarlo')
 guardar(nodo,clave+'_guardar')

