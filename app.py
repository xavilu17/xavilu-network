                        if doble_revision_actas else "Leído; una pasada más síntesis conjunta"
                        else:
                            estado_archivo = "Leído; revisión parcial o páginas pendientes"
                        detalle_archivo = "\n\n---\n\n".join(hallazgos_bloques) if hallazgos_bloques else "No se pudo recuperar contenido legible."
                        detalles_por_archivo.append({"nombre": arch.name, "estado": estado_archivo, "detalle": detalle_archivo})
                        progreso.progress((i + 1) / len(archs_acta))

                    inventario_estado = "\n".join(f"- {x['nombre']}: RECIBIDO; {x['estado']}" for x in detalles_por_archivo)
                    presupuesto_por_archivo = max(500, 35000 // max(1, len(detalles_por_archivo)))
                    datos_para_sintesis = []
                    for detalle in detalles_por_archivo:
                        contenido_sintesis, omitidos = recortar_texto_balanceado(detalle["detalle"], presupuesto_por_archivo)
                        datos_para_sintesis.append(f"\n===== ARCHIVO RECIBIDO: {detalle['nombre']} · ESTADO: {detalle['estado']} =====\n{contenido_sintesis}")
                        if omitidos:
                            datos_para_sintesis.append(f"AVISO: se omitieron {omitidos} caracteres del detalle de {detalle['nombre']} por límite de contexto; no declarar cobertura completa.")
                    prompt_sintesis_actas = f"""Prepara una auditoría conjunta, profesional y exhaustiva de todos los archivos que la aplicación confirmó como recibidos. El inventario y estados siguientes proceden de la aplicación. Es imposible afirmar que un archivo del inventario no se adjuntó; si su análisis falló, di que se recibió pero no pudo leerse.

INSTRUCCIÓN DEL USUARIO: {peticion_actas}

ARCHIVOS Y ESTADO REAL:
{inventario_estado}

REGLAS: Integra solo documentos realmente leídos; cita archivo y página en cada hallazgo. Compara las dos actas entre sí y con los estatutos: acuerdos, fechas, importes, responsables, plazos, reglas de convocatoria y mayorías. Distingue lo que se propone, se vota, se aprueba o queda pendiente. Detecta diferencias, contradicciones, acuerdos no reflejados, referencias a anexos no aportados, errores de suma, información que falta para ejecutar un acuerdo y cargas económicas. Si una lectura no aporta evidencia, no la des por hecho. Separa defectos de redacción/documentación de posibles cuestiones jurídicas. No inventes artículos ni dictámenes: para LPH cita únicamente los números que se verifican en los encabezados del texto BOE proporcionado; explica el contenido y la condición aplicable. Si la fuente oficial no se recuperó, declara esa limitación. No publiques nombres personales innecesarios.

ESTRUCTURA DEL INFORME:
1. Confirmación del inventario y cobertura real de cada archivo.
2. Resumen ejecutivo con los hallazgos principales.
3. Ficha de cada acta: fecha, tipo, orden del día y acuerdos registrados.
4. Tabla de acuerdos: archivo/página, acuerdo y estado, votación/mayoría/coeficientes si constan, importe/derrama, responsable y plazo.
5. Contraste acta contra acta y actas contra estatutos: coincidencias, diferencias y contradicciones concretas.
6. Revisión orientativa frente a LPH: asunto, dato documental, artículo verificable, condición que puede cambiar el análisis y límite.
7. Errores, omisiones y datos que los vecinos deberían conocer antes de ejecutar o votar; distingue error probado de dato no localizado.
8. Preguntas concretas y documentos que pedir, solo derivados de hallazgos.
9. Lista priorizada de próximos pasos y límites de la revisión.

TEXTO OFICIAL CONSOLIDADO RECUPERADO:
Artículos cuyos encabezados se verificaron: {', '.join(articulos_lph_actas) if articulos_lph_actas else 'ninguno; no citar números de artículo'}
{contexto_lph_actas[:10000]}

LECTURAS DOCUMENTALES POR ARCHIVO:
{''.join(datos_para_sintesis)}"""
                    estado_texto.text("Cruzando acuerdos entre documentos y redactando el informe final...")
                    sintesis_actas = llamar_ia(modelo_seleccionado, prompt_sintesis_actas, _get_key("gemini"), incluir_contexto_global=False)
                    if "[ERROR" in sintesis_actas:
                        sintesis_actas = "No se pudo generar la síntesis conjunta. A continuación se conservan las lecturas por archivo para no perder el análisis.\n\n" + "\n\n".join(f"## {x['nombre']}\n{x['detalle']}" for x in detalles_por_archivo)
                    st.session_state["acta_analisis_detalles"] = detalles_por_archivo
                    st.session_state["acta_analisis_res"] = f"INVENTARIO VALIDADO POR LA APLICACIÓN\n{inventario_estado}\n\n{str(sintesis_actas).strip()}"
                    st.session_state["acta_analisis_errores"] = errores_documentos
                    progreso.empty()
                    estado_texto.empty()

            for nombre_error, motivo_error in st.session_state.get("acta_analisis_errores", []):
                st.warning(f"No se pudo analizar {nombre_error}. Se omitió del informe. Motivo: {motivo_error}")

            if st.session_state.get("acta_analisis_res"):
                st.markdown(st.session_state["acta_analisis_res"])
                paquete_actas = st.session_state["acta_analisis_res"]
                for detalle in st.session_state.get("acta_analisis_detalles", []):
                    with st.expander(f"📄 Evidencia y lectura por archivo: {detalle['nombre']} ({detalle['estado']})"):
                        st.markdown(detalle["detalle"])
                    paquete_actas += f"\n\n===== LECTURA DETALLADA: {detalle['nombre']} =====\n{detalle['detalle']}"
                botones_descarga(paquete_actas, "Auditoria_Actas_Estatutos", formatos_extra=True)

    elif sub_idx == 4:
        st.markdown("### 📝 Citaciones y Notificaciones")
        tab_cit = st.tabs(["📝 Crear Nueva", "🔍 Analizar / Preguntar"])

        with tab_cit[0]:
            motivo = st.text_input("Motivo de la citación:", placeholder="Ej: Reunión extraordinaria por derrama...")
            fecha_hora = st.text_input("Fecha y Hora:", placeholder="Ej: 15 de Noviembre a las 19:00h")
            lugar = st.text_input("Lugar:", value="Portal del edificio o Sala de juntas")
            if st.button("📝 Redactar Citación", type="primary", key="btn_redactar_cit") and motivo:
                with st.spinner("Redactando..."):
                    prompt = f"Actúa como Administrador de Fincas. Redacta una notificación formal para los vecinos. Motivo: {motivo}. Fecha y hora: {fecha_hora}. Lugar: {lugar}. Incluye un lenguaje formal, la obligación de asistencia o delegación de voto, y espacio para la firma."
                    res = llamar_ia(modelo_seleccionado, prompt, _get_key("gemini"))
                    if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                    else: st.session_state["citacion_crear_res"] = res
            
            if st.session_state.get("citacion_crear_res"):
                st.text_area("Citación:", value=st.session_state["citacion_crear_res"], height=300)
                botones_descarga(st.session_state["citacion_crear_res"], "Citacion_Vecinos")

        with tab_cit[1]:
            st.markdown("Sube una citación o notificación (foto o PDF) para que la IA la revise en busca de defectos de forma o errores legales.")
            arch_citacion = st.file_uploader("Sube el documento:", type=["pdf", "jpg", "jpeg", "png"], key="up_citacion")
            pregunta_cit = st.text_area("¿Qué quieres saber sobre esta citación?", placeholder="Ej: ¿Falta el orden del día?, ¿Es válida legalmente?, redáctala de nuevo corrigiendo errores...", height=100)

            if arch_citacion and st.button("🔍 Revisar Documento", type="primary", key="btn_analizar_cit"):
                with st.spinner("Analizando el documento..."):
                    prompt_revision = f"""Actúa como un Administrador de Fincas y Abogado experto en propiedad horizontal en España.
                    Tu objetivo es revisar la siguiente citación.

                    INSTRUCCIÓN DEL USUARIO: {pregunta_cit if pregunta_cit.strip() else 'Analiza los defectos de forma y propón una versión corregida.'}

                    REGLAS ESTRICTAS:
                    1. NO TE INVENTES FALLOS. Si el documento tiene fecha, hora, lugar y orden del día, indica que es correcto.
                    2. Si el usuario te ha pedido algo concreto, responde a eso.
                    3. Si propones una versión corregida, usa exactamente los datos que aparecen en el documento.

                    Estructura tu respuesta así:
                    VEREDICTO: (Es válido o tiene fallos reales)
                    ANÁLISIS DE FORMA: (Puntos clave que sí tiene o que faltan de verdad)
                    PROPUESTA DE MEJORA: (Si es necesario, redacta la versión final con los datos del documento)
                    """

                    if arch_citacion.name.lower().endswith(".pdf"):
                        res = analizar_archivo_con_gemini(arch_citacion.getvalue(), "application/pdf", prompt_revision, _get_key("gemini"), modelo_local=modelo_seleccionado)
                    else:
                        mime = "image/png" if arch_citacion.name.lower().endswith(".png") else "image/jpeg"
                        res = analizar_archivo_con_gemini(arch_citacion.getvalue(), mime, prompt_revision, _get_key("gemini"), modelo_local=modelo_seleccionado)

                    if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                    else: st.session_state["citacion_res"] = res

            if st.session_state.get("citacion_res"):
                st.markdown(st.session_state["citacion_res"])
                botones_descarga(st.session_state["citacion_res"], "Revision_Citacion", formatos_extra=True)
