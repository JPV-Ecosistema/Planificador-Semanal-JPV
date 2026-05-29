import streamlit as st
import pandas as pd
import os
import json
import uuid
from datetime import datetime

# ---------------------------------------------------------
# BLOQUE 0: CONFIGURACIÓN GENERAL Y PERSISTENCIA
# ---------------------------------------------------------
import streamlit as st
import pandas as pd
import os
import json
import uuid
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="JPV - OpsControl", layout="wide")

PERSISTENCE_DIR = "persistence"

def init_system():
    os.makedirs(PERSISTENCE_DIR, exist_ok=True)

def get_week_identifier():
    return datetime.now().strftime("%Y_W%W")

def apply_custom_styles():
    st.markdown("""
        <style>
        .main { background-color: #f5f7f9; }
        .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #004a99; color: white; font-weight: bold;}
        .btn-guardar>button { background-color: #217346; color: white; width: 100%; height: 3em; margin-top: 20px;}
        .stDataFrame { border: 1px solid #c4ced4; }
        .marco-caso { background-color: white; padding: 15px; border-radius: 5px; border-left: 5px solid #217346; margin-bottom: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        .marco-gestion { background-color: white; padding: 15px; border-radius: 5px; border-left: 5px solid #004a99; margin-bottom: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        .tarea-marco { background-color: white; padding: 15px; border-radius: 5px; border-left: 5px solid #004a99; margin-bottom: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        .tarea-realizada { border-left: 5px solid #217346; opacity: 0.8; }
        h1, h2, h3, h4 { color: #003366 !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        </style>
    """, unsafe_allow_html=True)

init_system()
apply_custom_styles()

# ---------------------------------------------------------
# BLOQUE 1: FUNCIONES DE MEMORIA Y BASE DE DATOS LOCAL/NUBE
# ---------------------------------------------------------
def load_master_base():
    st.sidebar.header("📁 Base Maestra")
    uploaded_file = st.sidebar.file_uploader("Cargar 'Reporte de Acciones'", type=["xlsx", "csv"])
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.xlsx'):
                return pd.read_excel(uploaded_file, skiprows=5)
            else:
                return pd.read_csv(uploaded_file, skiprows=5)
        except Exception as e:
            st.sidebar.error(f"Error técnico: {e}")
    return None

def limpiar_monto_mcl(valor):
    if pd.isna(valor) or str(valor).strip() == "": return 0.0
    if isinstance(valor, (int, float)): return float(valor)
    
    # Limpieza agresiva de textos financieros
    v_str = str(valor).strip().replace('$', '').replace(' ', '')
    
    # Manejo robusto de formato chileno vs formato americano
    if '.' in v_str and ',' in v_str:
        if v_str.rfind(',') > v_str.rfind('.'): # Ej: 1.500.000,50
            v_str = v_str.replace('.', '').replace(',', '.')
        else: # Ej: 1,500,000.50
            v_str = v_str.replace(',', '')
    elif ',' in v_str: # Ej: 1500,50
        v_str = v_str.replace(',', '.')
        
    try:
        return float(v_str)
    except:
        return 0.0

def calcular_tramo_mcl(fila):
    valor = 0.0
    divisa = str(fila.get('Divisa', '')).upper()
    
    # Lectura estricta y exclusiva de la columna BI ('Perdida bruta (en moneda del caso)')
    col_perdida = 'Perdida bruta (en moneda del caso)'
    
    if col_perdida in fila and pd.notna(fila[col_perdida]):
        valor = limpiar_monto_mcl(fila[col_perdida])

    is_mcl = False
    tramo_str = "<= 1000 UF"
    
    # Filtro lógico para Major and Complex Losses
    if 'USD' in divisa or 'US$' in divisa or 'DÓLAR' in divisa or 'DOLAR' in divisa:
        if valor > 200000:
            is_mcl = True
            tramo_str = "> 200.000 USD (MCL)"
        else:
            tramo_str = "<= 200.000 USD"
    else: 
        if valor <= 1000: tramo_str = "<= 1000 UF"
        elif valor <= 5000: tramo_str = "> 1000 Y <= 5000 UF"
        else:
            is_mcl = True
            tramo_str = "> 5000 UF (MCL)"
            
    return tramo_str, is_mcl

def get_google_sheet():
    try:
        if "gcp_service_account" in st.secrets and "google_sheet_url" in st.secrets:
            scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
            client = gspread.authorize(creds)
            return client.open_by_url(st.secrets["google_sheet_url"]).sheet1
    except Exception:
        pass
    return None

def load_plan_semanal(ajustador):
    week_id = get_week_identifier()
    filename = f"plan_{ajustador.replace(' ', '_')}_{week_id}.json"
    filepath = os.path.join(PERSISTENCE_DIR, filename)
    if os.path.exists(filepath):
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f), filepath
    return [], filepath 

def save_plan_actualizado(filepath, data):
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
        
    sheet = get_google_sheet()
    if sheet:
        try:
            pass # Sincronización Google en Fase 2
        except: pass

# ---------------------------------------------------------
# BLOQUE 2: VISTA - PLANIFICADOR (SEMANAL Y MENSUAL MCL)
# ---------------------------------------------------------
def vista_planificador(modo="Semanal"):
    if modo == "Semanal":
        st.title("🗓️ Planificador Semanal")
        st.markdown("Seleccione los casos de la Base Maestra que proyecta gestionar durante la semana en curso.")
    else:
        st.title("🏆 Planificador Mensual MCL")
        st.markdown("Gestión estratégica de casos complejos (Major and Complex Losses > 5000 UF / > 200k USD).")
    
    CATALOGO_ACCIONES = {
        "En Ajuste": ["Revisión de cobertura", "Revisión de antecedentes", "Otro / Manual"],
        "Inspección": ["Presencial", "Remota", "Otro / Manual"],
        "Correos": ["Solicitud de Antecedentes", "Reiteracion 1", "Reiteracion 2", "Reiteracion 3", "Ultimatum", "Cierre por falta de interés", "Otro / Manual"],
        "Reunión": ["Presencial", "Presentación pptx", "On line", "Presentacion on line", "Otro / Manual"],
        "Preparar Informe": ["Preliminar Extendido", "Preliminar Corto", "Carta de Análisis de Pérdidas", "Carta de Cobertura (Rechazo)", "Informe Intermedio 1", "Informe Intermedio 2", "Informe Intermedio 3", "Informe Intermedio 4", "Informe Intermedio 5", "Informe Intermedio", "Informe Final de Liquidación", "Respuesta a Impugnación", "Ademdum", "Otro / Manual"],
        "Otra Acción (Manual)": ["Describir manualmente"]
    }
    
    df_maestro = load_master_base()
    
    if df_maestro is not None:
        col_ajustador = 'Ajustador senior' if 'Ajustador senior' in df_maestro.columns else df_maestro.columns[9]
        ajustadores_validos = sorted(df_maestro[col_ajustador].dropna().unique())
        
        ajustador_seleccionado = st.selectbox("Identificación de Ajustador:", [""] + ajustadores_validos)
        
        if ajustador_seleccionado:
            casos_vigentes = df_maestro[(df_maestro[col_ajustador] == ajustador_seleccionado) & (df_maestro['Estado'] != 'Cerrado')].copy()
            
            if modo == "Mensual":
                casos_vigentes = casos_vigentes[casos_vigentes.apply(lambda x: calcular_tramo_mcl(x)[1], axis=1)]
                st.warning("📊 Filtro MCL Activo: Mostrando exclusivamente siniestros complejos.")
            else:
                casos_vigentes = casos_vigentes[~casos_vigentes.apply(lambda x: calcular_tramo_mcl(x)[1], axis=1)]
                
            estados_maestros = sorted([str(x) for x in df_maestro['Estado'].dropna().unique() if str(x).strip()]) if 'Estado' in df_maestro.columns else ["Ajuste", "IFL", "Liquidación"]
            subestados_maestros = sorted([str(x) for x in df_maestro['Sub estado'].dropna().unique() if str(x).strip()]) if 'Sub estado' in df_maestro.columns else ["En Proceso", "Informe Preliminar", "Revisión Jefatura"]

            st.markdown("---")
            st.header("1. Selección de Casos Operativos")
            st.info(f"Inventario Vigente: {len(casos_vigentes)} casos disponibles bajo este filtro.")
            
            def formato_caso_nickname(x):
                fila = casos_vigentes.loc[x]
                base_str = f"Caso {fila['Número de caso']} - {fila['Asegurado']}"
                if 'Nickname' in fila and pd.notna(fila['Nickname']) and str(fila['Nickname']).strip() != "":
                    return f"{base_str} - [{fila['Nickname']}]"
                return base_str

            selected_indices = st.multiselect(
                "Seleccione los casos que intervendrá:",
                options=casos_vigentes.index.tolist(),
                format_func=formato_caso_nickname
            )
            
            plan_transaccional = []
            
            if selected_indices:
                st.markdown("---")
                st.header("2. Detalle de Acciones Operativas y Proyección de Estado")
                st.info("💡 Valide o modifique el Estado/Sub-estado proyectado, defina la cantidad de actividades y establezca las fechas de compromiso.")
                
                for idx in selected_indices:
                    fila = casos_vigentes.loc[idx]
                    caso_num = fila['Número de caso']
                    asegurado = fila['Asegurado']
                    estado_actual = str(fila['Estado']) if 'Estado' in fila and pd.notna(fila['Estado']) else "N/D"
                    subestado_actual = str(fila['Sub estado']) if 'Sub estado' in fila and pd.notna(fila['Sub estado']) else "N/D"
                    tramo, is_mcl = calcular_tramo_mcl(fila)
                    
                    nickname = ""
                    if 'Nickname' in fila and pd.notna(fila['Nickname']) and str(fila['Nickname']).strip() != "":
                        nickname = f" <span style='color:#004a99;'>[{fila['Nickname']}]</span>"
                    
                    with st.container():
                        st.markdown(f"""
                        <div class="marco-caso">
                            <h4>[{caso_num}] {asegurado}{nickname}</h4>
                            <p style="color:gray; font-size: 0.9em; margin-bottom: 5px;">
                                <b>Clasificación:</b> {tramo} | <b>Estado en Excel:</b> {estado_actual} | <b>Sub-estado en Excel:</b> {subestado_actual}
                            </p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        col_est, col_sub = st.columns(2)
                        with col_est:
                            opts_est = estados_maestros.copy()
                            if estado_actual not in opts_est and estado_actual != "N/D":
                                opts_est.append(estado_actual)
                            opts_est = sorted(list(set(opts_est)))
                            default_est_idx = opts_est.index(estado_actual) if estado_actual in opts_est else 0
                            estado_proyectado = st.selectbox(f"Proyectar Estado Final:", opts_est, index=default_est_idx, key=f"est_proj_{idx}")
                            
                        with col_sub:
                            opts_sub = subestados_maestros.copy()
                            if subestado_actual not in opts_sub and subestado_actual != "N/D":
                                opts_sub.append(subestado_actual)
                            opts_sub = sorted(list(set(opts_sub)))
                            default_sub_idx = opts_sub.index(subestado_actual) if subestado_actual in opts_sub else 0
                            subestado_proyectado = st.selectbox(f"Proyectar Sub-estado Final:", opts_sub, index=default_sub_idx, key=f"sub_proj_{idx}")

                        num_actividades = st.number_input(f"Cantidad de actividades para el caso {caso_num}:", min_value=1, max_value=15, value=3, key=f"num_act_{idx}")
                        
                        for i in range(1, int(num_actividades) + 1):
                            colA, colB, colC = st.columns([2, 2, 1])
                            with colA:
                                cat_accion = st.selectbox(f"Categoría Acción {i}:", [""] + list(CATALOGO_ACCIONES.keys()), key=f"cat_{idx}_{i}")
                            with colB:
                                accion_final = ""
                                if cat_accion:
                                    if cat_accion == "Otra Acción (Manual)":
                                        accion_final = st.text_input(f"Describa la acción {i}:", key=f"man_{idx}_{i}")
                                    else:
                                        sub_accion = st.selectbox(f"Detalle Acción {i}:", [""] + CATALOGO_ACCIONES[cat_accion], key=f"sub_{idx}_{i}")
                                        if sub_accion == "Otro / Manual":
                                            texto_manual = st.text_input(f"Especifique el detalle {i}:", key=f"man_{idx}_{i}")
                                            if texto_manual:
                                                accion_final = f"{cat_accion} - {texto_manual}"
                                        elif sub_accion:
                                            accion_final = f"{cat_accion} - {sub_accion}"
                            with colC:
                                fecha_compromiso = st.date_input(f"Fecha compromiso {i}:", key=f"fecha_{idx}_{i}")
                            
                            if accion_final.strip():
                                plan_transaccional.append({
                                    "id_transaccion": str(uuid.uuid4()),
                                    "tipo_plan": modo,
                                    "tipo_actividad": "Programada",
                                    "categoria": "Operativa",
                                    "numero_caso": str(caso_num),
                                    "asegurado": str(asegurado),
                                    "tramo_uf": tramo,
                                    "estado_proyectado": estado_proyectado,
                                    "subestado_proyectado": subestado_proyectado,
                                    "accion": accion_final,
                                    "fecha_compromiso": fecha_compromiso.strftime("%Y-%m-%d"),
                                    "estado_cumplimiento": "Pendiente",
                                    "fecha_planificacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                })
                                
            st.markdown("---")
            st.header("3. Acciones de Gestión")
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown('<div class="marco-gestion"><h4>🤝 Gestión Comercial</h4></div>', unsafe_allow_html=True)
                comercial_raw = st.text_area("Reuniones, visitas a corredoras, etc. (Una por línea):", key="txt_comercial", height=150)
                fecha_comercial = st.date_input("Fecha para Comercial:", key="fecha_com")
            
            with col2:
                st.markdown('<div class="marco-gestion"><h4>⚙️ Gestión Administrativa</h4></div>', unsafe_allow_html=True)
                admin_raw = st.text_area("Capacitaciones, comités, trámites, etc. (Una por línea):", key="txt_admin", height=150)
                fecha_admin = st.date_input("Fecha para Administrativa:", key="fecha_adm")
            
            if comercial_raw.strip():
                for accion in [linea.strip() for linea in comercial_raw.split('\n') if linea.strip()]:
                    plan_transaccional.append({
                        "id_transaccion": str(uuid.uuid4()), "tipo_plan": modo, "tipo_actividad": "Programada", "categoria": "Gestión Comercial", "numero_caso": "N/A", "asegurado": "N/A", "tramo_uf": "N/A", "estado_proyectado": "N/A", "subestado_proyectado": "N/A", "accion": accion, "fecha_compromiso": fecha_comercial.strftime("%Y-%m-%d"), "estado_cumplimiento": "Pendiente", "fecha_planificacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    })
                    
            if admin_raw.strip():
                for accion in [linea.strip() for linea in admin_raw.split('\n') if linea.strip()]:
                    plan_transaccional.append({
                        "id_transaccion": str(uuid.uuid4()), "tipo_plan": modo, "tipo_actividad": "Programada", "categoria": "Gestión Administrativa", "numero_caso": "N/A", "asegurado": "N/A", "tramo_uf": "N/A", "estado_proyectado": "N/A", "subestado_proyectado": "N/A", "accion": accion, "fecha_compromiso": fecha_admin.strftime("%Y-%m-%d"), "estado_cumplimiento": "Pendiente", "fecha_planificacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    })

            st.markdown("---")
            if len(plan_transaccional) > 0:
                st.info(f"Se han registrado **{len(plan_transaccional)} acciones**.")
                if st.button(f"💾 COMPROMETER PLAN {modo.upper()}"):
                    try:
                        week_id = get_week_identifier()
                        filename = f"plan_{ajustador_seleccionado.replace(' ', '_')}_{week_id}.json"
                        filepath = os.path.join(PERSISTENCE_DIR, filename)
                        
                        plan_existente = []
                        if os.path.exists(filepath):
                            with open(filepath, 'r', encoding='utf-8') as f:
                                plan_existente = json.load(f)
                        
                        plan_final = plan_existente + plan_transaccional
                        
                        save_plan_actualizado(filepath, plan_final)
                        st.success(f"Plan {modo} guardado exitosamente.")
                    except Exception as e:
                        st.error(f"Error: {e}")
            elif selected_indices:
                st.warning("Debe seleccionar al menos una acción válida para guardar el plan.")
    else:
        st.info("Módulo en espera: Suba el archivo 'Reporte de acciones' en el panel izquierdo.")

# ---------------------------------------------------------
# BLOQUE 3: VISTA - PROGRAMA DIARIO (EJECUCIÓN Y NO PROGRAMADOS)
# ---------------------------------------------------------
def vista_diario():
    st.title("☀️ Ejecución y Cumplimiento")
    hoy_str = datetime.now().strftime("%Y-%m-%d")
    st.markdown(f"**Fecha actual:** {datetime.now().strftime('%A, %d de %B de %Y')}")
    
    ajustador_input = st.text_input("Ingrese su nombre de Ajustador (Exacto al Plan Semanal/Mensual):", placeholder="Ej: Francisco Silva Ghisolfo")
    
    if ajustador_input:
        plan_data, filepath = load_plan_semanal(ajustador_input)
        
        # Módulo de Actividades No Programadas
        st.markdown("---")
        with st.expander("➕ REGISTRAR ACTIVIDAD NO PROGRAMADA (Urgencias / Fuera de Plan)", expanded=False):
            st.info("Utilice este módulo para reportar gestiones inmediatas que no estaban en su planificación original (Impactan positivamente en su métrica de cumplimiento).")
            colNP1, colNP2 = st.columns(2)
            with colNP1:
                np_caso = st.text_input("Número de Caso (o Referencia):", key="np_caso")
                np_asegurado = st.text_input("Asegurado:", key="np_aseg")
            with colNP2:
                np_accion = st.text_input("Acción Ejecutada:", key="np_acc")
                np_fecha = st.date_input("Fecha de Ejecución:", key="np_fec")
            
            if st.button("Guardar Actividad No Programada"):
                if np_caso and np_accion:
                    nueva_actividad = {
                        "id_transaccion": str(uuid.uuid4()),
                        "tipo_plan": "Diario",
                        "tipo_actividad": "No Programada",
                        "categoria": "Operativa",
                        "numero_caso": str(np_caso),
                        "asegurado": str(np_asegurado),
                        "tramo_uf": "N/D",
                        "estado_proyectado": "N/D",
                        "subestado_proyectado": "N/D",
                        "accion": np_accion,
                        "fecha_compromiso": np_fecha.strftime("%Y-%m-%d"),
                        "estado_cumplimiento": "Realizado",
                        "fecha_ejecucion": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "fecha_planificacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    plan_data.append(nueva_actividad)
                    save_plan_actualizado(filepath, plan_data)
                    st.success("Actividad no programada incorporada al reporte general exitosamente.")
                    st.rerun()
                else:
                    st.warning("⚠️ Debe ingresar al menos el Número de Caso y la Acción Ejecutada.")

        if not plan_data:
            st.warning(f"⚠️ No se encontraron compromisos agendados para **{ajustador_input}** en la semana en curso.")
        else:
            st.success("Plan Operativo sincronizado correctamente.")
            st.markdown("---")
            
            tareas_hoy = []
            tareas_resto = []
            
            for idx, tarea in enumerate(plan_data):
                tarea_con_indice = tarea.copy()
                tarea_con_indice['_posicion_original'] = idx
                if tarea.get("fecha_compromiso") == hoy_str:
                    tareas_hoy.append(tarea_con_indice)
                else:
                    tareas_resto.append(tarea_con_indice)
            
            total_tareas = len(plan_data)
            tareas_completadas = sum(1 for t in plan_data if t.get("estado_cumplimiento") == "Realizado")
            cambios_realizados = False
            
            with st.form(key="form_cumplimiento_estructurado"):
                if tareas_hoy:
                    st.subheader("🔥 Prioridad para Hoy (Compromisos del Día)")
                    for t in tareas_hoy:
                        pos = t['_posicion_original']
                        estado_actual = t.get("estado_cumplimiento", "Pendiente")
                        es_realizado = (estado_actual == "Realizado")
                        clase_css = "tarea-marco tarea-realizada" if es_realizado else "tarea-marco"
                        icono = "✅" if es_realizado else "⚡"
                        tipo_act = f" [{t.get('tipo_actividad', 'Programada').upper()}]"
                        
                        st.markdown(f'<div class="{clase_css}">', unsafe_allow_html=True)
                        if t["categoria"] == "Operativa":
                            est_p = t.get('estado_proyectado', 'N/D')
                            sub_p = t.get('subestado_proyectado', 'N/D')
                            st.markdown(f"**{icono} CASO [{t['numero_caso']}]** - {t['asegurado']} | *Tramo: {t['tramo_uf']}* | *Proyectado: {est_p} ({sub_p})*{tipo_act}")
                        else:
                            st.markdown(f"**{icono} {t['categoria'].upper()}**{tipo_act}")
                        st.markdown(f"**Entregable:** {t['accion']}")
                        
                        nuevo_estado = st.checkbox(f"Marcar como ejecutado", value=es_realizado, key=f"chk_hoy_{t['id_transaccion']}")
                        nuevo_texto_estado = "Realizado" if nuevo_estado else "Pendiente"
                        if nuevo_texto_estado != estado_actual:
                            plan_data[pos]["estado_cumplimiento"] = nuevo_texto_estado
                            plan_data[pos]["fecha_ejecucion"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if nuevo_estado else ""
                            cambios_realizados = True
                        st.markdown('</div>', unsafe_allow_html=True)
                else:
                    st.info("💡 No tienes actividades agendadas específicamente para la fecha de hoy. Abajo se despliega tu planificación extendida.")

                if tareas_resto:
                    st.subheader("📅 Resto de la Planificación (Semanal y Mensual)")
                    for t in tareas_resto:
                        pos = t['_posicion_original']
                        estado_actual = t.get("estado_cumplimiento", "Pendiente")
                        es_realizado = (estado_actual == "Realizado")
                        clase_css = "tarea-marco tarea-realizada" if es_realizado else "tarea-marco"
                        icono = "✅" if es_realizado else "⏳"
                        tipo_act = f" [{t.get('tipo_actividad', 'Programada').upper()}]"
                        
                        st.markdown(f'<div class="{clase_css}">', unsafe_allow_html=True)
                        if t["categoria"] == "Operativa":
                            est_p = t.get('estado_proyectado', 'N/D')
                            st.markdown(f"**{icono} CASO [{t['numero_caso']}]** - {t['asegurado']} | *Compromiso: {t['fecha_compromiso']}* | *Proyectado: {est_p}*{tipo_act}")
                        else:
                            st.markdown(f"**{icono} {t['categoria'].upper()}** | *Compromiso: {t['fecha_compromiso']}*{tipo_act}")
                        st.markdown(f"**Entregable:** {t['accion']}")
                        
                        nuevo_estado = st.checkbox(f"Marcar como ejecutado", value=es_realizado, key=f"chk_rest_{t['id_transaccion']}")
                        nuevo_texto_estado = "Realizado" if nuevo_estado else "Pendiente"
                        if nuevo_texto_estado != estado_actual:
                            plan_data[pos]["estado_cumplimiento"] = nuevo_texto_estado
                            plan_data[pos]["fecha_ejecucion"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if nuevo_estado else ""
                            cambios_realizados = True
                        st.markdown('</div>', unsafe_allow_html=True)
                
                st.markdown('<div class="btn-guardar">', unsafe_allow_html=True)
                submit_button = st.form_submit_button(label="💾 ACTUALIZAR CUMPLIMIENTO DIARIO")
                st.markdown('</div>', unsafe_allow_html=True)
            
            if submit_button:
                if cambios_realizados:
                    try:
                        save_plan_actualizado(filepath, plan_data)
                        st.success("¡Base de datos local actualizada!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error al escribir en el disco: {e}")
                else:
                    st.info("No se detectaron cambios.")
            
            st.markdown("---")
            progreso = int((tareas_completadas / total_tareas) * 100) if total_tareas > 0 else 0
            st.progress(progreso)
            st.caption(f"Avance de cumplimiento global: {tareas_completadas} de {total_tareas} tareas realizadas ({progreso}%).")

# ---------------------------------------------------------
# BLOQUE 4: VISTA - REPORTE DE JEFATURA (CARTA GANTT EXCEL/WORD)
# ---------------------------------------------------------
def vista_reportes():
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from docx import Document
    from docx.shared import RGBColor
    from docx.oxml.ns import nsdecls
    from docx.oxml import parse_xml
    
    st.title("📊 Carta Gantt de Planificación Semanal")
    st.markdown("Visión gerencial estructurada por Casos, Estados Proyectados y Entregables en la línea de tiempo.")
    
    archivos_json = [f for f in os.listdir(PERSISTENCE_DIR) if f.endswith('.json')]
    
    if not archivos_json:
        st.warning("No hay planes registrados en el servidor en este momento.")
        return
        
    datos_consolidados = []
    for archivo in archivos_json:
        partes_nombre = archivo.replace('plan_', '').replace('.json', '').split('_20')
        nombre_ajustador = partes_nombre[0].replace('_', ' ')
        
        filepath = os.path.join(PERSISTENCE_DIR, archivo)
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                plan = json.load(f)
                for tarea in plan:
                    tarea['Ajustador'] = nombre_ajustador
                    datos_consolidados.append(tarea)
        except Exception:
            pass

    if datos_consolidados:
        df_raw = pd.DataFrame(datos_consolidados)
        df_operativa = df_raw[df_raw['categoria'] == 'Operativa'].copy()
        
        if not df_operativa.empty:
            df_operativa['fecha_compromiso'] = pd.to_datetime(df_operativa['fecha_compromiso']).dt.date
            
            if 'estado_proyectado' not in df_operativa.columns:
                df_operativa['estado_proyectado'] = 'N/D'
            if 'subestado_proyectado' not in df_operativa.columns:
                df_operativa['subestado_proyectado'] = 'N/D'

            df_gantt_visual = df_operativa.pivot_table(
                index=['Ajustador', 'numero_caso', 'asegurado', 'estado_proyectado', 'subestado_proyectado'], 
                columns='fecha_compromiso', 
                values='accion', 
                aggfunc=lambda x: ' | '.join(x)
            ).fillna('')
            
            st.subheader("🛠️ Gantt Operativo (Vista Previa)")
            st.dataframe(df_gantt_visual, use_container_width=True)

            # --- GENERADOR DE EXCEL NATIVO CORPORATIVO ---
            wb = Workbook()
            ws = wb.active
            ws.title = "Plan Semanal Gantt"
            
            fechas_unicas = sorted(df_operativa['fecha_compromiso'].unique())
            headers_excel = ["Ajustador", "Caso", "Asegurado", "Acción y Entregable"] + [f.strftime('%A %d-%m') for f in fechas_unicas]
            ws.append(headers_excel)
            
            grouped = df_operativa.groupby(['Ajustador', 'numero_caso', 'asegurado', 'accion'])
            for name, group in grouped:
                ajustador, caso, asegurado, accion = name
                row = [ajustador, caso, asegurado, accion]
                for f in fechas_unicas:
                    if f in group['fecha_compromiso'].values:
                        row.append("X")
                    else:
                        row.append("")
                ws.append(row)
                
            header_fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            thin_border = Border(left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'), 
                                 top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3'))
            center_alignment = Alignment(horizontal="center", vertical="center")
            
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_alignment
                cell.border = thin_border
                
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in row:
                    cell.border = thin_border
                    if cell.column > 4:
                        cell.alignment = center_alignment
                        if cell.value == "X":
                            cell.fill = PatternFill(start_color="217346", end_color="217346", fill_type="solid")
                            cell.font = Font(color="FFFFFF", bold=True)
            
            for i, width in enumerate([25, 12, 35, 30] + [15]*len(fechas_unicas), 1):
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width
                
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            
            # --- GENERADOR DE WORD NATIVO CORPORATIVO ---
            doc = Document()
            doc.add_heading('Reporte Consolidado de Planificación Semanal - Gantt', 0)
            doc.add_paragraph('Visión gerencial estructurada por Casos, Tareas y Entregables en la línea de tiempo.')
            
            headers_word = ["Ajustador", "Caso", "Asegurado", "Acción/Entregable"] + [f.strftime('%d-%m') for f in fechas_unicas]
            table = doc.add_table(rows=1, cols=len(headers_word))
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            
            for i, title in enumerate(headers_word):
                hdr_cells[i].text = title
                hdr_cells[i].paragraphs[0].runs[0].font.bold = True
                shading_elm = parse_xml(r'<w:shd {} w:fill="003366"/>'.format(nsdecls('w')))
                hdr_cells[i]._tc.get_or_add_tcPr().append(shading_elm)
                hdr_cells[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
                
            for name, group in grouped:
                ajustador, caso, asegurado, accion = name
                row_cells = table.add_row().cells
                row_cells[0].text = str(ajustador)
                row_cells[1].text = str(caso)
                row_cells[2].text = str(asegurado)
                row_cells[3].text = str(accion)
                for i, f in enumerate(fechas_unicas):
                    col_idx = 4 + i
                    if f in group['fecha_compromiso'].values:
                        shading_elm = parse_xml(r'<w:shd {} w:fill="217346"/>'.format(nsdecls('w')))
                        row_cells[col_idx]._tc.get_or_add_tcPr().append(shading_elm)

            word_buffer = io.BytesIO()
            doc.save(word_buffer)

            # --- BOTONES DE DESCARGA ---
            st.markdown("---")
            st.markdown("### Opciones de Exportación Corporativa")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.download_button(
                    label="📥 DESCARGAR GANTT (EXCEL)",
                    data=excel_buffer.getvalue(),
                    file_name=f"Gantt_Planificacion_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with col2:
                st.download_button(
                    label="📥 DESCARGAR REPORTE (WORD)",
                    data=word_buffer.getvalue(),
                    file_name=f"Reporte_Planificacion_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            with col3:
                csv_raw = df_raw.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 DESCARGAR DATA BRUTA (CSV)",
                    data=csv_raw,
                    file_name=f"Data_Bruta_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                )
        else:
            st.info("No hay tareas operativas (casos) planificadas aún.")
# ---------------------------------------------------------
# BLOQUE PRINCIPAL: ENRUTADOR DE NAVEGACIÓN
# ---------------------------------------------------------
def main():
    st.sidebar.image("https://img.icons8.com/color/96/000000/engineering.png", width=60)
    st.sidebar.title("Navegación OpsControl")
    
    opcion = st.sidebar.radio(
        "Ir a:",
        ["Planificador Semanal", "Planificador Mensual MCL", "Programa Diario", "Reportes Jefatura"]
    )
    
    st.sidebar.markdown("---")
    
    if opcion == "Planificador Semanal":
        vista_planificador("Semanal")
    elif opcion == "Planificador Mensual MCL":
        vista_planificador("Mensual")
    elif opcion == "Programa Diario":
        vista_diario()
    elif opcion == "Reportes Jefatura":
        vista_reportes()

if __name__ == "__main__":
    main()
