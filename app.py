# ---------------------------------------------------------
# IMPORTACIONES DEL SISTEMA
# ---------------------------------------------------------
import streamlit as st
import pandas as pd
import os
import json
import uuid
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta

# ---------------------------------------------------------
# BLOQUE 0: CONFIGURACIÓN GENERAL Y PERSISTENCIA
# ---------------------------------------------------------
import streamlit as st
import pandas as pd
import os
import json
import uuid
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="JPV - OpsControl", layout="wide")

PERSISTENCE_DIR = "persistence"

def init_system():
    os.makedirs(PERSISTENCE_DIR, exist_ok=True)

def get_week_identifier(offset_weeks=0):
    target_date = datetime.now() + timedelta(weeks=offset_weeks)
    return target_date.strftime("%Y_W%W")

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
# VERSIÓN: 2.1.4 (Guardián de Sesión contra Bucles 429)
# ---------------------------------------------------------
def get_google_client():
    try:
        if "gcp_service_account" in st.secrets and "google_sheet_url" in st.secrets:
            scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
            return gspread.authorize(creds)
    except Exception as e:
        st.error(f"⚠️ Error crítico de conexión a Google Cloud: {e}")
    return None

def get_google_sheet():
    client = get_google_client()
    if client:
        try:
            return client.open_by_url(st.secrets["google_sheet_url"]).sheet1
        except Exception as e:
            st.error(f"⚠️ Error al abrir la hoja de cálculo: {e}")
    return None

def load_master_base():
    st.sidebar.header("📁 Base Maestra")
    filepath = os.path.join(PERSISTENCE_DIR, "BASE_MAESTRA.json")
    
    df_local = None
    fecha_actualizacion = None
    
    # 1. Recuperación en silencio desde Google Sheets si no hay base local
    if not os.path.exists(filepath):
        client = get_google_client()
        if client:
            try:
                doc = client.open_by_url(st.secrets["google_sheet_url"])
                ws = doc.worksheet("Base_Maestra")
                metadata = ws.row_values(1)
                
                if len(metadata) >= 2 and metadata[0] == "FECHA_ACTUALIZACION":
                    fecha_str = metadata[1]
                    datos_crud = ws.get_all_records(head=2) 
                    df_local = pd.DataFrame(datos_crud)
                    
                    datos_guardar = {"fecha": fecha_str, "data": df_local.to_dict(orient="records")}
                    with open(filepath, 'w', encoding='utf-8') as f:
                        json.dump(datos_guardar, f, ensure_ascii=False)
            except Exception:
                pass
                
    # 2. Lectura de caché local
    if os.path.exists(filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                datos = json.load(f)
                df_local = pd.DataFrame(datos['data'])
                fecha_actualizacion = datetime.strptime(datos['fecha'], "%Y-%m-%d %H:%M:%S")
        except Exception:
            pass

    # 3. Lógica de Semáforo y Caducidad (7 Días)
    necesita_actualizacion = True
    if fecha_actualizacion:
        dias = (datetime.now() - fecha_actualizacion).days
        if dias <= 7:
            st.sidebar.success(f"✅ Base actualizada hace {dias} días ({fecha_actualizacion.strftime('%d/%m/%Y')})")
            necesita_actualizacion = False
        else:
            st.sidebar.error(f"🚨 Base caducada (hace {dias} días). Requiere actualización urgente.")
    else:
        st.sidebar.warning("⚠️ No hay base de datos en el sistema.")

    # 4. Motor de Carga, Sanitización y Respaldo Nube
    with st.sidebar.expander("📥 Subir / Actualizar Base Maestra", expanded=necesita_actualizacion):
        uploaded_file = st.file_uploader("Cargar 'Reporte de Acciones'", type=["xlsx", "csv"])
        if uploaded_file is not None:
            # Creamos una huella digital única para el archivo subido
            file_signature = f"{uploaded_file.name}_{uploaded_file.size}"
            
            # El Guardián solo permite procesar si es un archivo nuevo o no ha sido registrado en esta sesión
            if st.session_state.get("ultima_base_procesada") != file_signature:
                try:
                    if uploaded_file.name.endswith('.xlsx'):
                        df_nuevo = pd.read_excel(uploaded_file, skiprows=5)
                    else:
                        df_nuevo = pd.read_csv(uploaded_file, skiprows=5)
                    
                    df_nuevo = df_nuevo.fillna("") 
                    for col in df_nuevo.columns:
                        df_nuevo[col] = df_nuevo[col].astype(str)
                        df_nuevo[col] = df_nuevo[col].apply(
                            lambda x: "" if str(x).strip().lower() in ["nan", "nat", "none", "<na>", "inf", "-inf"] else x
                        )

                    fecha_hoy = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    # Guardado Local Inmediato (Asegura operatividad interna)
                    datos_guardar = {"fecha": fecha_hoy, "data": df_nuevo.to_dict(orient="records")}
                    with open(filepath, 'w', encoding='utf-8') as f:
                        json.dump(datos_guardar, f, ensure_ascii=False)
                        
                    # Intento de Sincronización en la Nube corporativa
                    try:
                        client = get_google_client()
                        if client:
                            doc = client.open_by_url(st.secrets["google_sheet_url"])
                            try:
                                ws = doc.worksheet("Base_Maestra")
                            except:
                                ws = doc.add_worksheet(title="Base_Maestra", rows="100", cols="100")
                            
                            ws.clear()
                            matriz = [["FECHA_ACTUALIZACION", fecha_hoy]]
                            matriz.append(df_nuevo.columns.astype(str).tolist())
                            matriz.extend(df_nuevo.values.tolist())
                            ws.update("A1", matriz)

                        st.success("¡Base Maestra procesada y asegurada en la nube corporativa!")
                    except Exception as cloud_error:
                        if "429" in str(cloud_error):
                            st.warning("⚠️ Base guardada localmente con éxito. Google Cloud está en pausa (Límite 429 de peticiones). Puedes operar tu planificador normalmente.")
                        else:
                            st.warning(f"⚠️ Base guardada localmente. Hubo un detalle con el respaldo en nube: {cloud_error}")
                    
                    # Registramos el archivo como procesado para congelar ejecuciones repetidas
                    st.session_state["ultima_base_procesada"] = file_signature
                    st.rerun() 
                except Exception as e:
                    st.sidebar.error(f"Error crítico al procesar el Excel: {e}")
                
    return df_local

def limpiar_monto_mcl(valor):
    if pd.isna(valor) or str(valor).strip() in ["", "nan", "NaN", "NaT"]: return 0.0
    if isinstance(valor, (int, float)): return float(valor)
    
    v_str = str(valor).strip().replace('$', '').replace(' ', '')
    if '.' in v_str and ',' in v_str:
        if v_str.rfind(',') > v_str.rfind('.'):
            v_str = v_str.replace('.', '').replace(',', '.')
        else:
            v_str = v_str.replace(',', '')
    elif ',' in v_str:
        v_str = v_str.replace(',', '.')
        
    try:
        return float(v_str)
    except:
        return 0.0

def calcular_tramo_mcl(fila):
    valor = 0.0
    divisa = str(fila.get('Divisa', '')).upper()
    col_perdida = 'Perdida bruta (en moneda del caso)'
    
    if col_perdida in fila and pd.notna(fila[col_perdida]):
        valor = limpiar_monto_mcl(fila[col_perdida])

    is_mcl = False
    tramo_str = "<= 1000 UF"
    
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

def get_month_identifier(offset_months=0):
    now = datetime.now()
    year = now.year
    month = now.month + offset_months
    while month > 12:
        month -= 12
        year += 1
    return f"{year}_{month:02d}"

def sync_from_cloud(filename, filepath):
    sheet = get_google_sheet()
    if sheet:
        try:
            registros = sheet.get_all_records()
            for fila in registros:
                if str(fila.get('Archivo', '')) == filename:
                    datos_json = json.loads(fila.get('JSON_Data', '[]'))
                    with open(filepath, 'w', encoding='utf-8') as f:
                        json.dump(datos_json, f, ensure_ascii=False, indent=4)
                    return datos_json
        except Exception as e:
            # Evitamos alertas invasivas si la cuota de lectura también está retenida temporalmente
            pass
    return []

def load_plan_semanal(ajustador, offset_weeks=0):
    week_id = get_week_identifier(offset_weeks)
    filename = f"plan_{ajustador.replace(' ', '_')}_{week_id}.json"
    filepath = os.path.join(PERSISTENCE_DIR, filename)
    if os.path.exists(filepath):
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f), filepath
    else:
        data = sync_from_cloud(filename, filepath)
        return data, filepath 

def load_plan_mensual(ajustador, offset_months=0, explicit_month_id=None):
    month_id = explicit_month_id if explicit_month_id else get_month_identifier(offset_months)
    filename = f"plan_mensual_mcl_{ajustador.replace(' ', '_')}_{month_id}.json"
    filepath = os.path.join(PERSISTENCE_DIR, filename)
    if os.path.exists(filepath):
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f), filepath
    else:
        data = sync_from_cloud(filename, filepath)
        return data, filepath

def save_plan_actualizado(filepath, data):
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
        
    sheet = get_google_sheet()
    if sheet:
        try:
            filename = os.path.basename(filepath)
            json_str = json.dumps(data, ensure_ascii=False)
            fecha_upd = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            try:
                headers = sheet.row_values(1)
            except:
                headers = []
                
            if not headers or 'Archivo' not in headers:
                sheet.clear()
                sheet.append_row(['Archivo', 'JSON_Data', 'Ultima_Actualizacion'])
                sheet.append_row([filename, json_str, fecha_upd])
                return

            col_archivo_idx = headers.index('Archivo') + 1
            archivos_en_nube = sheet.col_values(col_archivo_idx)
            
            if filename in archivos_en_nube:
                fila_idx = archivos_en_nube.index(filename) + 1
                rango = f'A{fila_idx}:C{fila_idx}'
                sheet.update(rango, [[filename, json_str, fecha_upd]])
            else:
                sheet.append_row([filename, json_str, fecha_upd])
        except Exception as e:
            st.error(f"❌ Error al escribir filas en Google Sheets: {e}")

# ---------------------------------------------------------
# BLOQUE 2: VISTA - PLANIFICADOR (SEMANAL Y MENSUAL MCL)
# VERSIÓN: 2.2 (Autogestión de Reseteo Corporativo)
# ---------------------------------------------------------
def vista_planificador(modo="Semanal"):
    col_t1, col_t2 = st.columns([2, 1])
    with col_t1:
        if modo == "Semanal":
            st.title("🗓️ Planificador Semanal")
            st.markdown("Seleccione los casos de la Base Maestra que proyecta gestionar.")
        else:
            st.title("🏆 Planificador Mensual MCL")
            st.markdown("Gestión estratégica de casos complejos (Major and Complex Losses > 5000 UF / > 200k USD).")
            
    with col_t2:
        st.markdown("<br>", unsafe_allow_html=True)
        if modo == "Semanal":
            semana_opcion = st.radio("¿Qué semana estás planificando?", ["Semana Actual", "Próxima Semana"], horizontal=True)
            offset_weeks = 0 if semana_opcion == "Semana Actual" else 1
            offset_months = 0
        else:
            mes_opcion = st.radio("¿Qué mes estás planificando?", ["Mes Actual", "Próximo Mes"], horizontal=True)
            offset_months = 0 if mes_opcion == "Mes Actual" else 1
            offset_weeks = 0
            
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
                st.warning(f"📊 Filtro MCL Activo ({mes_opcion}): Mostrando exclusivamente siniestros complejos.")
            else:
                casos_vigentes = casos_vigentes[~casos_vigentes.apply(lambda x: calcular_tramo_mcl(x)[1], axis=1)]
                
            estados_maestros = sorted([str(x) for x in df_maestro['Estado'].dropna().unique() if str(x).strip()]) if 'Estado' in df_maestro.columns else ["Ajuste", "IFL", "Liquidación"]
            subestados_maestros = sorted([str(x) for x in df_maestro['Sub estado'].dropna().unique() if str(x).strip()]) if 'Sub estado' in df_maestro.columns else ["En Proceso", "Informe Preliminar", "Revisión Jefatura"]

            # ---------------------------------------------------------
            # MOTOR DE RESETEO DE PLANIFICACIÓN (ZONA DE CONTROL)
            # ---------------------------------------------------------
            if modo == "Mensual":
                _, path_boveda = load_plan_mensual(ajustador_seleccionado, offset_months=offset_months)
            else:
                t_week_id = get_week_identifier(offset_weeks)
                path_boveda = os.path.join(PERSISTENCE_DIR, f"plan_{ajustador_seleccionado.replace(' ', '_')}_{t_week_id}.json")
            
            plan_historico = []
            if os.path.exists(path_boveda):
                try:
                    with open(path_boveda, 'r', encoding='utf-8') as f:
                        plan_historico = json.load(f)
                except:
                    pass
            
            if plan_historico:
                with st.expander("🚨 Zona de Control: Modificar / Resetear Período Activo", expanded=False):
                    st.warning(f"Atención: Ya tienes un plan comprometido para este período con {len(plan_historico)} acciones registradas en el sistema.")
                    st.markdown("Si cometiste un error en las fechas, tramos o asignaciones, puedes vaciar el plan actual para volver a formularlo de manera correcta.")
                    if st.button("🗑️ ANULAR PLAN ACTUAL Y EMPEZAR DE CERO", key="btn_pánico_reset"):
                        try:
                            # Sobreescritura atómica con lista vacía en local y nube corporativa
                            save_plan_actualizado(path_boveda, [])
                            st.success("¡Planificación anulada exitosamente! La pizarra está limpia.")
                            st.rerun()
                        except Exception as reset_err:
                            st.error(f"Error al ejecutar el reseteo: {reset_err}")
            # ---------------------------------------------------------

            plan_transaccional = []
            
            # --- LÓGICA DE HERENCIA MCL EN CASCADA POR PERÍODO MENSUAL ---
            if modo == "Semanal":
                target_date = datetime.now() + timedelta(weeks=offset_weeks)
                target_month_id = target_date.strftime("%Y_%m")
                mcl_data, mcl_path = load_plan_mensual(ajustador_seleccionado, explicit_month_id=target_month_id)
                target_week_id = get_week_identifier(offset_weeks)
                
                mcl_pendientes = []
                for t in mcl_data:
                    try:
                        fec_obj = datetime.strptime(t['fecha_compromiso'], "%Y-%m-%d")
                        if fec_obj.strftime("%Y_W%W") == target_week_id and not t.get("agendado_semana"):
                            mcl_pendientes.append(t)
                    except: pass
                
                if mcl_pendientes:
                    st.markdown("---")
                    st.markdown('<div class="marco-gestion" style="border-left: 5px solid #d9534f;"><h4>🚨 Hitos MCL Heredados (Obligatorio asignar día)</h4></div>', unsafe_allow_html=True)
                    st.info(f"Estos compromisos provienen de tu Planificador Mensual de {target_date.strftime('%B %Y')}. Asígnales un día específico para incorporarlos a tu agenda semanal.")
                    
                    for idx, mcl_task in enumerate(mcl_pendientes):
                        c1, c2 = st.columns([3, 1])
                        with c1:
                            st.write(f"**Caso:** [{mcl_task['numero_caso']}] {mcl_task['asegurado']}")
                            st.write(f"**Entregable Estratégico:** {mcl_task['accion']}")
                        with c2:
                            fec_obj = datetime.strptime(mcl_task['fecha_compromiso'], "%Y-%m-%d")
                            nueva_fecha_mcl = st.date_input(f"Día de ejecución:", value=fec_obj, key=f"mcl_fec_{idx}")
                            
                        task_to_add = mcl_task.copy()
                        task_to_add['fecha_compromiso'] = nueva_fecha_mcl.strftime("%Y-%m-%d")
                        task_to_add['id_mcl_origen'] = mcl_task['id_transaccion'] 
                        plan_transaccional.append(task_to_add)

            st.markdown("---")
            st.header("1. Selección de Casos Operativos Regulares")
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
                    
                    honorarios_estimados = 0.0
                    try:
                        if len(casos_vigentes.columns) >= 67:
                            valor_bo = fila.iloc[66] 
                            honorarios_estimados = limpiar_monto_mcl(valor_bo)
                    except Exception:
                        pass
                    
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
                                    "honorarios_estimados": honorarios_estimados,
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
                num_comercial = st.number_input("Cantidad de gestiones comerciales:", min_value=0, max_value=15, value=1, key="num_comercial")
                
                for i in range(1, int(num_comercial) + 1):
                    c_acc, c_fec = st.columns([2, 1])
                    with c_acc:
                        acc_com = st.text_input(f"Detalle gestión {i}:", placeholder="Reuniones, visitas a corredoras...", key=f"txt_com_{i}")
                    with c_fec:
                        fec_com = st.date_input(f"Fecha {i}:", key=f"fec_com_{i}")
                    
                    if acc_com.strip():
                        plan_transaccional.append({
                            "id_transaccion": str(uuid.uuid4()), "tipo_plan": modo, "tipo_actividad": "Programada", 
                            "categoria": "Gestión Comercial", "numero_caso": "N/A", "asegurado": "N/A", "tramo_uf": "N/A", 
                            "honorarios_estimados": 0.0,
                            "estado_proyectado": "N/A", "subestado_proyectado": "N/A", "accion": acc_com.strip(), 
                            "fecha_compromiso": fec_com.strftime("%Y-%m-%d"), "estado_cumplimiento": "Pendiente", 
                            "fecha_planificacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        })
            
            with col2:
                st.markdown('<div class="marco-gestion"><h4>⚙️ Gestión Administrativa</h4></div>', unsafe_allow_html=True)
                num_admin = st.number_input("Cantidad de gestiones administrativas:", min_value=0, max_value=15, value=1, key="num_admin")
                
                for i in range(1, int(num_admin) + 1):
                    c_acc, c_fec = st.columns([2, 1])
                    with c_acc:
                        acc_adm = st.text_input(f"Detalle gestión {i}:", placeholder="Capacitaciones, comités...", key=f"txt_adm_{i}")
                    with c_fec:
                        fec_adm = st.date_input(f"Fecha {i}:", key=f"fec_adm_{i}")
                    
                    if acc_adm.strip():
                        plan_transaccional.append({
                            "id_transaccion": str(uuid.uuid4()), "tipo_plan": modo, "tipo_actividad": "Programada", 
                            "categoria": "Gestión Administrativa", "numero_caso": "N/A", "asegurado": "N/A", "tramo_uf": "N/A", 
                            "honorarios_estimados": 0.0,
                            "estado_proyectado": "N/A", "subestado_proyectado": "N/A", "accion": acc_adm.strip(), 
                            "fecha_compromiso": fec_adm.strftime("%Y-%m-%d"), "estado_cumplimiento": "Pendiente", 
                            "fecha_planificacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        })

            st.markdown("---")
            if len(plan_transaccional) > 0:
                st.info(f"Se han consolidado **{len(plan_transaccional)} acciones** en total para guardar.")
                if st.button(f"💾 COMPROMETER PLAN {modo.upper()}"):
                    try:
                        if modo == "Mensual":
                            _, filepath = load_plan_mensual(ajustador_seleccionado, offset_months=offset_months)
                            plan_existente = []
                            if os.path.exists(filepath):
                                with open(filepath, 'r', encoding='utf-8') as f:
                                    plan_existente = json.load(f)
                            save_plan_actualizado(filepath, plan_existente + plan_transaccional)
                            st.success(f"Plan Mensual MCL ({mes_opcion}) guardado exitosamente en la bóveda y respaldado en la nube.")
                        else:
                            target_week_id = get_week_identifier(offset_weeks)
                            filename = f"plan_{ajustador_seleccionado.replace(' ', '_')}_{target_week_id}.json"
                            filepath = os.path.join(PERSISTENCE_DIR, filename)
                            
                            plan_existente = []
                            if os.path.exists(filepath):
                                with open(filepath, 'r', encoding='utf-8') as f:
                                    plan_existente = json.load(f)
                                    
                            if 'mcl_data' in locals() and mcl_data:
                                mcl_ids_agendados = [t['id_mcl_origen'] for t in plan_transaccional if 'id_mcl_origen' in t]
                                for t in mcl_data:
                                    if t['id_transaccion'] in mcl_ids_agendados:
                                        t['agendado_semana'] = True
                                save_plan_actualizado(mcl_path, mcl_data) 
                                
                            save_plan_actualizado(filepath, plan_existente + plan_transaccional)
                            st.success(f"Plan Semanal guardado exitosamente para la {semana_opcion} y respaldado en la nube.")
                    except Exception as e:
                        st.error(f"Error al guardar: {e}")
            elif selected_indices or int(num_comercial) > 0 or int(num_admin) > 0:
                st.warning("Complete el detalle de las acciones o seleccione casos válidos para guardar el plan.")
    else:
        st.info("Módulo en espera: Suba el archivo 'Reporte de acciones' en el panel izquierdo.")

# ---------------------------------------------------------
# BLOQUE 3: VISTA - PROGRAMA DIARIO (EJECUCIÓN Y NO PROGRAMADOS)
# VERSIÓN: 2.4 (Sincronización Automática e Ingreso Múltiple No Programadas)
# ---------------------------------------------------------
def vista_diario():
    st.title("☀️ Ejecución y Cumplimiento")
    hoy_str = datetime.now().strftime("%Y-%m-%d")
    st.markdown(f"**Fecha actual:** {datetime.now().strftime('%A, %d de %B de %Y')}")
    
    week_id = get_week_identifier()
    
    # --- MOTOR DE SINCRONIZACIÓN AUTOMÁTICA (Rescate de Nube) ---
    archivos_locales = [f for f in os.listdir(PERSISTENCE_DIR) if f.startswith('plan_') and f.endswith('.json')]
    if not archivos_locales:
        with st.spinner("Sincronizando planes desde la nube corporativa..."):
            sheet = get_google_sheet()
            if sheet:
                try:
                    registros = sheet.get_all_records()
                    for fila in registros:
                        fname = str(fila.get('Archivo', ''))
                        if fname.endswith('.json') and fname != "BASE_MAESTRA.json":
                            fpath = os.path.join(PERSISTENCE_DIR, fname)
                            try:
                                datos_json = json.loads(fila.get('JSON_Data', '[]'))
                                with open(fpath, 'w', encoding='utf-8') as f:
                                    json.dump(datos_json, f, ensure_ascii=False, indent=4)
                            except Exception:
                                pass
                except Exception:
                    pass

    # --- MOTOR DE BÚSQUEDA AUTOMÁTICA DE AJUSTADORES ---
    try:
        archivos = [f for f in os.listdir(PERSISTENCE_DIR) if f.startswith('plan_') and f.endswith(f'_{week_id}.json')]
        ajustadores_con_plan = []
        for archivo in archivos:
            nombre = archivo.replace('plan_', '').replace(f'_{week_id}.json', '').replace('_', ' ')
            ajustadores_con_plan.append(nombre)
        ajustadores_con_plan = sorted(list(set(ajustadores_con_plan)))
    except Exception:
        ajustadores_con_plan = []

    if not ajustadores_con_plan:
        st.info("⚠️ Ningún ajustador ha comprometido su Plan Operativo para esta semana en la base del sistema.")
        return

    ajustador_input = st.selectbox("Seleccione su nombre de Ajustador:", [""] + ajustadores_con_plan)
    
    if ajustador_input:
        plan_data, filepath = load_plan_semanal(ajustador_input)
        
        # --- MÓDULO DE ACTIVIDADES NO PROGRAMADAS (INGRESO MÚLTIPLE) ---
        st.markdown("---")
        with st.expander("➕ REGISTRAR ACTIVIDADES NO PROGRAMADAS (Urgencias / Fuera de Plan)", expanded=False):
            st.info("Utilice este módulo para reportar gestiones inmediatas que no estaban en su planificación original. Estas impactan positivamente en su métrica de cumplimiento global.")
            
            num_np = st.number_input("Cantidad de urgencias a registrar ahora:", min_value=1, max_value=15, value=1, key="num_np_diario")
            
            nuevas_actividades = []
            
            for i in range(1, int(num_np) + 1):
                st.markdown(f"**Urgencia {i}**")
                colNP1, colNP2 = st.columns(2)
                with colNP1:
                    np_caso = st.text_input(f"Número de Caso (o Ref) {i}:", key=f"np_caso_{i}")
                    np_asegurado = st.text_input(f"Asegurado {i}:", key=f"np_aseg_{i}")
                with colNP2:
                    np_accion = st.text_input(f"Acción Ejecutada {i}:", key=f"np_acc_{i}")
                    # Por defecto sugiere la fecha de hoy, pero permite cambiarla si la urgencia fue ayer
                    np_fecha = st.date_input(f"Fecha de Ejecución {i}:", value=datetime.now(), key=f"np_fec_{i}")
                
                # Solo preparamos el registro si llenaron los datos mínimos clave
                if np_caso.strip() and np_accion.strip():
                    nuevas_actividades.append({
                        "id_transaccion": str(uuid.uuid4()),
                        "tipo_plan": "Diario",
                        "tipo_actividad": "No Programada",
                        "categoria": "Operativa",
                        "numero_caso": str(np_caso),
                        "asegurado": str(np_asegurado),
                        "tramo_uf": "N/D",
                        "honorarios_estimados": 0.0, 
                        "estado_proyectado": "N/D",
                        "subestado_proyectado": "N/D",
                        "accion": np_accion,
                        "fecha_compromiso": np_fecha.strftime("%Y-%m-%d"),
                        "estado_cumplimiento": "Realizado", 
                        "fecha_ejecucion": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "fecha_planificacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    })
                st.markdown("<hr style='margin-top: 5px; margin-bottom: 15px;'>", unsafe_allow_html=True)
            
            if st.button("💾 Guardar Actividades No Programadas"):
                if len(nuevas_actividades) > 0:
                    plan_data.extend(nuevas_actividades)
                    save_plan_actualizado(filepath, plan_data)
                    st.success(f"¡{len(nuevas_actividades)} actividades no programadas incorporadas al reporte general exitosamente!")
                    st.rerun()
                else:
                    st.warning("⚠️ Debe ingresar al menos el Número de Caso y la Acción Ejecutada en alguna de las filas para poder guardar.")

        if not plan_data:
            st.warning(f"⚠️ No se encontraron compromisos agendados para **{ajustador_input}** en la semana en curso.")
        else:
            st.success("Plan Operativo sincronizado correctamente.")
            st.markdown("---")
            
            # --- MOTOR DE CÁLCULO DE HONORARIOS Y TAREAS ---
            tareas_hoy = []
            tareas_resto = []
            uf_proyectadas_hoy = 0.0
            uf_ejecutadas_hoy = 0.0
            uf_proyectadas_semana = 0.0
            uf_ejecutadas_semana = 0.0
            
            for idx, tarea in enumerate(plan_data):
                tarea_con_indice = tarea.copy()
                tarea_con_indice['_posicion_original'] = idx
                
                try:
                    uf_tarea = float(tarea.get("honorarios_estimados", 0.0))
                except:
                    uf_tarea = 0.0
                    
                es_realizado = (tarea.get("estado_cumplimiento") == "Realizado")
                
                uf_proyectadas_semana += uf_tarea
                if es_realizado:
                    uf_ejecutadas_semana += uf_tarea
                
                if tarea.get("fecha_compromiso") == hoy_str:
                    tareas_hoy.append(tarea_con_indice)
                    uf_proyectadas_hoy += uf_tarea
                    if es_realizado:
                        uf_ejecutadas_hoy += uf_tarea
                else:
                    tareas_resto.append(tarea_con_indice)
            
            total_tareas = len(plan_data)
            tareas_completadas = sum(1 for t in plan_data if t.get("estado_cumplimiento") == "Realizado")
            cambios_realizados = False
            
            # --- TABLERO VISUAL DE HONORARIOS ---
            st.subheader("📊 Rendimiento Financiero del Plan")
            col_met1, col_met2, col_met3, col_met4 = st.columns(4)
            col_met1.metric("UF Proyectadas Hoy", f"{uf_proyectadas_hoy:,.2f}")
            col_met2.metric("UF Ejecutadas Hoy", f"{uf_ejecutadas_hoy:,.2f}")
            col_met3.metric("UF Proyectadas Semana", f"{uf_proyectadas_semana:,.2f}")
            col_met4.metric("UF Ejecutadas Semana", f"{uf_ejecutadas_semana:,.2f}")
            st.markdown("---")
            
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
                        uf_txt = f" | 💰 {float(t.get('honorarios_estimados', 0.0)):,.2f} UF"
                        
                        st.markdown(f'<div class="{clase_css}">', unsafe_allow_html=True)
                        if t["categoria"] == "Operativa":
                            est_p = t.get('estado_proyectado', 'N/D')
                            sub_p = t.get('subestado_proyectado', 'N/D')
                            st.markdown(f"**{icono} CASO [{t['numero_caso']}]** - {t['asegurado']} | *Tramo: {t['tramo_uf']}* | *Proyectado: {est_p} ({sub_p})*{tipo_act}{uf_txt}")
                        else:
                            st.markdown(f"**{icono} {t['categoria'].upper()}**{tipo_act}{uf_txt}")
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
                        uf_txt = f" | 💰 {float(t.get('honorarios_estimados', 0.0)):,.2f} UF"
                        
                        st.markdown(f'<div class="{clase_css}">', unsafe_allow_html=True)
                        if t["categoria"] == "Operativa":
                            est_p = t.get('estado_proyectado', 'N/D')
                            st.markdown(f"**{icono} CASO [{t['numero_caso']}]** - {t['asegurado']} | *Compromiso: {t['fecha_compromiso']}* | *Proyectado: {est_p}*{tipo_act}{uf_txt}")
                        else:
                            st.markdown(f"**{icono} {t['categoria'].upper()}** | *Compromiso: {t['fecha_compromiso']}*{tipo_act}{uf_txt}")
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
                        st.success("¡Base de datos local actualizada y respaldada en la nube corporativa!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error al escribir en el disco/nube: {e}")
                else:
                    st.info("No se detectaron cambios.")
            
            st.markdown("---")
            progreso = int((tareas_completadas / total_tareas) * 100) if total_tareas > 0 else 0
            st.progress(progreso)
            st.caption(f"Avance de cumplimiento global: {tareas_completadas} de {total_tareas} tareas realizadas ({progreso}%).")

# ---------------------------------------------------------
# BLOQUE 4: VISTA - REPORTE DE JEFATURA (GANTT, DASHBOARD Y OPERACIONAL)
# VERSIÓN: 4.2 (Word Gráfico, Tab Operacional, MCL y Control de Decimales)
# ---------------------------------------------------------
def vista_reportes():
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from docx import Document
    from docx.shared import RGBColor, Cm, Pt
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import nsdecls
    from docx.oxml import parse_xml
    from datetime import timedelta
    import numpy as np
    
    st.title("📊 Tablero de Control y Planificación")
    st.markdown("Visión gerencial del rendimiento financiero, cumplimiento operativo y línea de tiempo de la división.")
    
    col_radio, col_btn = st.columns([2, 1])
    with col_radio:
        week_id_obj = st.radio("Seleccione la semana a reportar:", ["Semana Actual", "Próxima Semana"], horizontal=True)
    with col_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        forzar_sync = st.button("🔄 Sincronizar Nube ahora", type="primary", use_container_width=True)

    offset = 0 if week_id_obj == "Semana Actual" else 1
    
    # --- CÁLCULO DE FECHAS ---
    hoy = datetime.now()
    target_date = hoy + timedelta(weeks=offset)
    lunes = target_date - timedelta(days=target_date.weekday())
    dias_semana_target = [(lunes + timedelta(days=i)).date() for i in range(7)]
    target_week_id = get_week_identifier(offset)
    
    # --- MOTOR DE SINCRONIZACIÓN ---
    archivos_existentes = [f for f in os.listdir(PERSISTENCE_DIR) if f.startswith('plan_') and f.endswith('.json')]
    
    if forzar_sync or len(archivos_existentes) == 0:
        with st.spinner("Descargando planes de todos los ajustadores desde Google Sheets..."):
            sheet = get_google_sheet()
            if sheet:
                try:
                    registros = sheet.get_all_records()
                    for fila in registros:
                        fname = str(fila.get('Archivo', ''))
                        if fname.endswith('.json') and fname != "BASE_MAESTRA.json":
                            fpath = os.path.join(PERSISTENCE_DIR, fname)
                            try:
                                datos_json = json.loads(fila.get('JSON_Data', '[]'))
                                with open(fpath, 'w', encoding='utf-8') as f:
                                    json.dump(datos_json, f, ensure_ascii=False, indent=4)
                            except Exception:
                                pass
                    if forzar_sync:
                        st.success("¡Sincronización completada! Todos los planes están actualizados.")
                except Exception as e:
                    st.warning(f"Error al sincronizar con la nube: {e}")
    
    # --- LECTURA GLOBAL Y CRUCE MAESTRO ---
    archivos_json = [f for f in os.listdir(PERSISTENCE_DIR) if f.endswith('.json') and f != "BASE_MAESTRA.json"]
    
    df_maestro = load_master_base()
    diccionario_nicknames = {}
    ajustadores_validos = []
    
    if df_maestro is not None:
        if 'Número de caso' in df_maestro.columns and 'Nickname' in df_maestro.columns:
            for _, row in df_maestro.iterrows():
                if pd.notna(row['Número de caso']) and str(row['Número de caso']).strip() != "":
                    diccionario_nicknames[str(row['Número de caso'])] = str(row['Nickname']) if pd.notna(row['Nickname']) else ""
        
        col_ajustador = 'Ajustador senior' if 'Ajustador senior' in df_maestro.columns else df_maestro.columns[9]
        ajustadores_validos = sorted(df_maestro[col_ajustador].dropna().unique())

    datos_consolidados = []
    for archivo in archivos_json:
        partes_nombre = archivo.replace('plan_mensual_mcl_', '').replace('plan_', '').replace('.json', '').split('_20')
        nombre_ajustador = partes_nombre[0].replace('_', ' ')
        
        filepath = os.path.join(PERSISTENCE_DIR, archivo)
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                plan = json.load(f)
                for tarea in plan:
                    tarea['Ajustador'] = nombre_ajustador
                    tarea['Nick Name'] = diccionario_nicknames.get(str(tarea.get('numero_caso', '')), '')
                    datos_consolidados.append(tarea)
        except Exception:
            pass

    df_week = pd.DataFrame()
    if datos_consolidados:
        df_raw = pd.DataFrame(datos_consolidados)
        df_raw['fecha_filtro'] = pd.to_datetime(df_raw['fecha_compromiso'], errors='coerce').dt.date
        df_week = df_raw[df_raw['fecha_filtro'].isin(dias_semana_target)].copy()
        
        if not df_week.empty:
            if 'honorarios_estimados' not in df_week.columns:
                df_week['honorarios_estimados'] = 0.0
            df_week['honorarios_estimados'] = pd.to_numeric(df_week['honorarios_estimados'], errors='coerce').fillna(0.0)
            if 'tramo_uf' not in df_week.columns:
                df_week['tramo_uf'] = 'N/D'

    # ---------------------------------------------------------
    # ARQUITECTURA DE PESTAÑAS (3 TABS)
    # ---------------------------------------------------------
    tab_dashboard, tab_operacional, tab_gantt = st.tabs(["📈 Dashboard Ejecutivo (BI)", "📋 Reporte Operacional de Equipo", "📊 Carta Gantt Operativa"])
    
    # =========================================================
    # PESTAÑA 1: DASHBOARD EJECUTIVO
    # =========================================================
    with tab_dashboard:
        if df_week.empty:
            st.info("No hay datos cargados para generar el Dashboard en esta semana.")
        else:
            df_realizados = df_week[df_week['estado_cumplimiento'] == 'Realizado'].copy()
            
            # --- 1. RESUMEN EJECUTIVO FINANCIERO ---
            st.markdown('<div class="marco-gestion" style="border-left: 5px solid #003366;"><h4>💰 Cuadrante 1: Valorización de Cartera y Facturación Proyectada</h4></div>', unsafe_allow_html=True)
            cond_facturable = df_realizados['accion'].str.contains('Informe Final de Liquidación|Carta de Cobertura \(Rechazo\)', case=False, na=False)
            uf_facturables_caja = df_realizados[cond_facturable]['honorarios_estimados'].sum()
            uf_traccionadas_wip = df_realizados[(df_realizados['categoria'] == 'Operativa') & (~cond_facturable)]['honorarios_estimados'].sum()
            
            c_fin1, c_fin2 = st.columns(2)
            c_fin1.metric("Ingreso Efectivo Facturable (Cierres/Rechazos)", f"{uf_facturables_caja:,.2f} UF", delta="Directo a Caja", delta_color="normal")
            c_fin2.metric("Valor Potencial Traccionado (WIP / Proceso)", f"{uf_traccionadas_wip:,.2f} UF", delta="Esfuerzo Operativo", delta_color="off")
            st.markdown("---")
            
            # --- 2. KPIS DE CUMPLIMIENTO ---
            st.markdown('<div class="marco-gestion" style="border-left: 5px solid #17a2b8;"><h4>⏱️ Cuadrante 2: Métricas de Cumplimiento y Carga Operativa</h4></div>', unsafe_allow_html=True)
            t_planificadas = len(df_week[df_week['tipo_actividad'] == 'Programada'])
            t_planificadas_hechas = len(df_realizados[df_realizados['tipo_actividad'] == 'Programada'])
            t_urgencias = len(df_realizados[df_realizados['tipo_actividad'] == 'No Programada'])
            
            adherencia = (t_planificadas_hechas / t_planificadas * 100) if t_planificadas > 0 else 0
            total_esfuerzo = t_planificadas_hechas + t_urgencias
            ratio_plan = (t_planificadas_hechas / total_esfuerzo * 100) if total_esfuerzo > 0 else 0
            ratio_urg = (t_urgencias / total_esfuerzo * 100) if total_esfuerzo > 0 else 0
            
            c_kpi1, c_kpi2, c_kpi3 = st.columns(3)
            c_kpi1.metric("Adherencia al Plan", f"{adherencia:.1f}%", f"{t_planificadas_hechas} de {t_planificadas} Tareas Originales")
            c_kpi2.metric("Ratio de Esfuerzo Planificado", f"{ratio_plan:.1f}%", "Trabajo Proactivo", delta_color="normal")
            c_kpi3.metric("Fricción (No Programado/Urgencias)", f"{ratio_urg:.1f}%", f"{t_urgencias} Actividades Reactivas", delta_color="inverse")
            st.markdown("---")
            
            # --- SECCIÓN ESPECIAL MCL ---
            st.markdown('<div class="marco-gestion" style="border-left: 5px solid #d9534f;"><h4>🏆 Radar de Casos MCL (Major and Complex Losses)</h4></div>', unsafe_allow_html=True)
            cond_mcl = df_realizados['tramo_uf'].str.contains('MCL|> 5', case=False, na=False)
            df_mcl = df_realizados[cond_mcl][['Ajustador', 'numero_caso', 'asegurado', 'accion', 'honorarios_estimados']].copy()
            if not df_mcl.empty:
                df_mcl['honorarios_estimados'] = df_mcl['honorarios_estimados'].apply(lambda x: f"{x:,.2f}")
                st.dataframe(df_mcl.rename(columns={'numero_caso': 'Caso', 'asegurado': 'Asegurado', 'accion': 'Gestión MCL', 'honorarios_estimados': 'Hon UF'}), use_container_width=True, hide_index=True)
            else:
                st.info("Sin movimientos reportados en casos de tramo MCL esta semana.")
            st.markdown("---")

            # --- 3. INVENTARIO DE PRODUCCIÓN ---
            st.markdown('<div class="marco-gestion" style="border-left: 5px solid #28a745;"><h4>🏭 Cuadrante 3: Inventario de Producción (Entregables de Valor)</h4></div>', unsafe_allow_html=True)
            cond_entregables = df_realizados['accion'].str.contains('Preparar Informe|Presentación pptx|Presentacion on line', case=False, na=False)
            df_entregables = df_realizados[cond_entregables][['Ajustador', 'numero_caso', 'Nick Name', 'asegurado', 'accion', 'honorarios_estimados']].copy()
            
            if not df_entregables.empty:
                df_entregables['accion'] = df_entregables['accion'].str.replace('Preparar Informe - ', '', regex=False)
                df_entregables['accion'] = df_entregables['accion'].str.replace('Reunión - ', '', regex=False)
                df_entregables['honorarios_estimados'] = df_entregables['honorarios_estimados'].apply(lambda x: f"{x:,.2f}")
                df_mostrar = df_entregables.rename(columns={'numero_caso': 'Caso', 'asegurado': 'Asegurado', 'accion': 'Tipo de Entregable', 'honorarios_estimados': 'Hon UF'})
                st.dataframe(df_mostrar, use_container_width=True, hide_index=True)
            else:
                st.info("Aún no se han ejecutado entregables de valor (Informes o Presentaciones) esta semana.")
            st.markdown("---")
            
            # --- 4. GESTIÓN ESTRATÉGICA ---
            st.markdown('<div class="marco-gestion" style="border-left: 5px solid #ffc107;"><h4>🤝 Cuadrante 4: Gestión Estratégica Transversal</h4></div>', unsafe_allow_html=True)
            col_com, col_adm = st.columns(2)
            with col_com:
                st.markdown("**Gestiones Comerciales:**")
                df_com = df_realizados[df_realizados['categoria'] == 'Gestión Comercial'][['Ajustador', 'accion']]
                if not df_com.empty:
                    st.dataframe(df_com.rename(columns={'accion': 'Detalle'}), use_container_width=True, hide_index=True)
                else:
                    st.caption("Sin gestiones comerciales cerradas.")
            with col_adm:
                st.markdown("**Gestiones Administrativas:**")
                df_adm = df_realizados[df_realizados['categoria'] == 'Gestión Administrativa'][['Ajustador', 'accion']]
                if not df_adm.empty:
                    st.dataframe(df_adm.rename(columns={'accion': 'Detalle'}), use_container_width=True, hide_index=True)
                else:
                    st.caption("Sin gestiones administrativas cerradas.")
                    
            # --- GENERADORES EXPORTACIÓN DASHBOARD (WORD GRÁFICO) ---
            dash_wb = Workbook()
            ws_resumen = dash_wb.active
            ws_resumen.title = "Resumen Ejecutivo"
            ws_resumen.append(["Métrica Financiera", "UF"])
            ws_resumen.append(["Ingreso Efectivo Facturable (Cierres/Rechazos)", round(uf_facturables_caja, 2)])
            ws_resumen.append(["Valor Potencial Traccionado (WIP / Proceso)", round(uf_traccionadas_wip, 2)])
            ws_resumen.append([])
            ws_resumen.append(["KPI Operativo", "Valor"])
            ws_resumen.append(["Adherencia al Plan (%)", f"{adherencia:.1f}%"])
            ws_resumen.append(["Ratio de Esfuerzo Planificado (%)", f"{ratio_plan:.1f}%"])
            ws_resumen.append(["Fricción (Urgencias) (%)", f"{ratio_urg:.1f}%"])
            excel_dash_buffer = io.BytesIO()
            dash_wb.save(excel_dash_buffer)
            
            # Funciones de estilo para Word
            def formatear_cabecera_tabla(table, bg_color="003366"):
                hdr_cells = table.rows[0].cells
                for cell in hdr_cells:
                    run = cell.paragraphs[0].runs[0] if cell.paragraphs[0].runs else cell.paragraphs[0].add_run(cell.text)
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), bg_color))
                    cell._tc.get_or_add_tcPr().append(shading_elm)

            dash_doc = Document()
            dash_doc.add_heading(f'Dashboard Ejecutivo y Cumplimiento - {week_id_obj}', 0)
            
            # Tabla Financiera
            dash_doc.add_heading('1. Valorización de Cartera y Facturación Proyectada', level=2)
            t_fin = dash_doc.add_table(rows=2, cols=2)
            t_fin.style = 'Table Grid'
            t_fin.rows[0].cells[0].text, t_fin.rows[0].cells[1].text = "Métrica Financiera", "Valor (UF)"
            formatear_cabecera_tabla(t_fin)
            t_fin.rows[1].cells[0].text = "Ingreso Efectivo Facturable (Cierres/Rechazos)"
            t_fin.rows[1].cells[1].text = f"{uf_facturables_caja:,.2f}"
            t_fin.add_row().cells[0].text = "Valor Potencial Traccionado (WIP / Proceso)"
            t_fin.rows[2].cells[1].text = f"{uf_traccionadas_wip:,.2f}"
            dash_doc.add_paragraph("")
            
            # Tabla KPIs
            dash_doc.add_heading('2. Métricas de Cumplimiento y Carga Operativa', level=2)
            t_kpi = dash_doc.add_table(rows=2, cols=3)
            t_kpi.style = 'Table Grid'
            t_kpi.rows[0].cells[0].text, t_kpi.rows[0].cells[1].text, t_kpi.rows[0].cells[2].text = "Adherencia al Plan", "Ratio Planificado", "Fricción (Urgencias)"
            formatear_cabecera_tabla(t_kpi, "17A2B8")
            t_kpi.rows[1].cells[0].text = f"{adherencia:.1f}%"
            t_kpi.rows[1].cells[1].text = f"{ratio_plan:.1f}%"
            t_kpi.rows[1].cells[2].text = f"{ratio_urg:.1f}%"
            dash_doc.add_paragraph("")

            # Tabla MCL
            if not df_mcl.empty:
                dash_doc.add_heading('Radar de Casos MCL', level=2)
                t_mcl = dash_doc.add_table(rows=1, cols=len(df_mcl.columns))
                t_mcl.style = 'Table Grid'
                for i, col_name in enumerate(["Ajustador", "Caso", "Asegurado", "Gestión MCL", "Hon UF"]):
                    t_mcl.rows[0].cells[i].text = col_name
                formatear_cabecera_tabla(t_mcl, "D9534F")
                for row_val in df_mcl.values.tolist():
                    row_cells = t_mcl.add_row().cells
                    for i, val in enumerate(row_val):
                        row_cells[i].text = str(val)
                dash_doc.add_paragraph("")
            
            # Tabla Entregables
            if not df_entregables.empty:
                dash_doc.add_heading('3. Inventario de Producción (Entregables)', level=2)
                t_ent = dash_doc.add_table(rows=1, cols=len(df_mostrar.columns))
                t_ent.style = 'Table Grid'
                for i, col_name in enumerate(df_mostrar.columns):
                    t_ent.rows[0].cells[i].text = str(col_name)
                formatear_cabecera_tabla(t_ent, "28A745")
                for row_val in df_mostrar.values.tolist():
                    row_cells = t_ent.add_row().cells
                    for i, val in enumerate(row_val):
                        row_cells[i].text = str(val)
                dash_doc.add_paragraph("")
                
            # Tablas Comerciales y Admin
            dash_doc.add_heading('4. Gestión Estratégica Transversal', level=2)
            if not df_com.empty:
                dash_doc.add_paragraph("Gestiones Comerciales:", style='List Bullet')
                for _, row in df_com.iterrows():
                    dash_doc.add_paragraph(f"{row['Ajustador']}: {row['accion']}")
            if not df_adm.empty:
                dash_doc.add_paragraph("Gestiones Administrativas:", style='List Bullet')
                for _, row in df_adm.iterrows():
                    dash_doc.add_paragraph(f"{row['Ajustador']}: {row['accion']}")

            word_dash_buffer = io.BytesIO()
            dash_doc.save(word_dash_buffer)
            
            st.markdown("### Exportar Resumen Ejecutivo")
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.download_button("📥 DESCARGAR DASHBOARD (EXCEL)", data=excel_dash_buffer.getvalue(), file_name=f"Dashboard_Ejecutivo_{target_week_id}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col_d2:
                st.download_button("📥 DESCARGAR DASHBOARD (WORD)", data=word_dash_buffer.getvalue(), file_name=f"Resumen_Ejecutivo_{target_week_id}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    # =========================================================
    # PESTAÑA 2: REPORTE OPERACIONAL DE EQUIPO (DETALLE INTERNO)
    # =========================================================
    with tab_operacional:
        st.markdown("### 📋 Radiografía Operacional por Ajustador")
        st.markdown("Análisis individual para reuniones de seguimiento y evaluación de carga de trabajo.")
        
        if not ajustadores_validos:
            st.warning("No se pudo cargar la lista de ajustadores desde la Base Maestra.")
        else:
            for ajustador in ajustadores_validos:
                if df_week.empty:
                    st.markdown(f"#### {ajustador}")
                    st.markdown("<h5 style='color:#d9534f;'>🚨 AJUSTADOR SIN PLAN</h5>", unsafe_allow_html=True)
                    st.markdown("---")
                    continue
                
                df_aj = df_week[df_week['Ajustador'] == ajustador].copy()
                st.markdown(f"#### 👤 {ajustador}")
                
                if df_aj.empty:
                    st.markdown("<h5 style='color:#d9534f; background-color:#f9dede; padding:10px; border-radius:5px;'>🚨 AJUSTADOR SIN PLAN ESTA SEMANA</h5>", unsafe_allow_html=True)
                else:
                    df_aj_realizado = df_aj[df_aj['estado_cumplimiento'] == 'Realizado']
                    t_prog = len(df_aj[df_aj['tipo_actividad'] == 'Programada'])
                    t_prog_hechas = len(df_aj_realizado[df_aj_realizado['tipo_actividad'] == 'Programada'])
                    t_np = len(df_aj_realizado[df_aj_realizado['tipo_actividad'] == 'No Programada'])
                    
                    adh_aj = (t_prog_hechas / t_prog * 100) if t_prog > 0 else 0
                    total_aj = t_prog_hechas + t_np
                    rat_prog = (t_prog_hechas / total_aj * 100) if total_aj > 0 else 0
                    rat_np = (t_np / total_aj * 100) if total_aj > 0 else 0
                    
                    c_op1, c_op2, c_op3 = st.columns(3)
                    c_op1.metric("Cumplimiento del Plan", f"{adh_aj:.0f}%", f"{t_prog_hechas}/{t_prog} tareas")
                    c_op2.metric("Trabajo Programado", f"{rat_prog:.0f}%", "Proactivo")
                    c_op3.metric("Trabajo No Programado", f"{rat_np:.0f}%", "Urgencias/Reactivas", delta_color="inverse")
                    
                    with st.expander(f"Ver desglose detallado de {ajustador}"):
                        if t_np > 0:
                            st.markdown("**🔴 Detalle de Carga No Programada (Urgencias):**")
                            df_np_view = df_aj_realizado[df_aj_realizado['tipo_actividad'] == 'No Programada'][['numero_caso', 'asegurado', 'accion', 'fecha_ejecucion']]
                            st.dataframe(df_np_view, use_container_width=True, hide_index=True)
                        else:
                            st.caption("No registró actividades fuera de programa.")
                            
                        # Admin y Comercial
                        cond_estr = df_aj_realizado['categoria'].isin(['Gestión Comercial', 'Gestión Administrativa'])
                        if cond_estr.any():
                            st.markdown("**🔵 Gestiones Comerciales y Administrativas:**")
                            st.dataframe(df_aj_realizado[cond_estr][['categoria', 'accion']], use_container_width=True, hide_index=True)
                st.markdown("---")

    # =========================================================
    # PESTAÑA 3: CARTA GANTT Y EXPORTACIÓN CORPORATIVA
    # =========================================================
    with tab_gantt:
        if df_week.empty:
            st.info("No hay tareas operativas planificadas aún para esta semana.")
        else:
            df_operativa = df_week[df_week['categoria'] == 'Operativa'].copy() if 'categoria' in df_week.columns else df_week.copy()
            
            if not df_operativa.empty:
                df_operativa['fecha_compromiso'] = pd.to_datetime(df_operativa['fecha_compromiso'], errors='coerce').dt.date
                if 'estado_proyectado' not in df_operativa.columns:
                    df_operativa['estado_proyectado'] = 'N/D'

                df_operativa = df_operativa.sort_values(by=['Ajustador', 'fecha_compromiso'])
                df_gantt_visual = df_operativa.pivot_table(
                    index=['Ajustador', 'numero_caso', 'Nick Name', 'asegurado', 'estado_proyectado'], 
                    columns='fecha_compromiso', values='accion', aggfunc=lambda x: ' | '.join(x)
                ).fillna('')
                
                st.subheader("🛠️ Gantt Operativo (Línea de tiempo de la división)")
                st.dataframe(df_gantt_visual, use_container_width=True)

                # Exportación Excel Gantt
                wb = Workbook()
                ws = wb.active
                ws.title = "Plan Semanal Gantt"
                fechas_unicas = dias_semana_target 
                ws.append(["Ajustador", "Caso", "Nick Name", "Asegurado", "Acción y Entregable"] + [f.strftime('%A %d-%m') for f in fechas_unicas] + ["Hon UF"])
                
                grouped = df_operativa.groupby(['Ajustador', 'numero_caso', 'Nick Name', 'asegurado', 'accion'])
                for name, group in grouped:
                    row = list(name)
                    for f in fechas_unicas:
                        row.append("X" if f in group['fecha_compromiso'].values else "")
                    row.append(round(group['honorarios_estimados'].sum(), 2))
                    ws.append(row)
                    
                header_fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment = header_fill, Font(color="FFFFFF", bold=True), Alignment(horizontal="center", vertical="center")
                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                
                # Exportación Word Gantt
                doc = Document()
                section = doc.sections[-1]
                section.orientation, section.page_width, section.page_height = WD_ORIENT.LANDSCAPE, section.page_height, section.page_width
                section.top_margin, section.bottom_margin, section.left_margin, section.right_margin = Cm(1.27), Cm(1.27), Cm(1.27), Cm(1.27)

                doc.add_heading(f'Reporte Consolidado de Planificación Semanal - {week_id_obj}', 0)
                headers_word = ["Ajustador", "Caso", "Nick Name", "Asegurado", "Acción/Entregable", "L", "M", "X", "J", "V", "S", "D", "Hon UF"]
                table = doc.add_table(rows=1, cols=len(headers_word))
                table.style = 'Table Grid'
                
                for i, title in enumerate(headers_word):
                    table.rows[0].cells[i].text = title
                    table.rows[0].cells[i].paragraphs[0].runs[0].font.bold = True
                    table.rows[0].cells[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
                    table.rows[0].cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    table.rows[0].cells[i]._tc.get_or_add_tcPr().append(parse_xml(r'<w:shd {} w:fill="003366"/>'.format(nsdecls('w'))))
                    
                ajustador_previo = ""
                for _, row_data in df_operativa.iterrows():
                    row_cells = table.add_row().cells
                    ajustador_actual = str(row_data['Ajustador'])
                    row_cells[0].text = "" if ajustador_actual == ajustador_previo else ajustador_actual
                    ajustador_previo = ajustador_actual
                    row_cells[1].text, row_cells[2].text, row_cells[3].text, row_cells[4].text = str(row_data['numero_caso']), str(row_data['Nick Name']), str(row_data['asegurado']), str(row_data['accion'])
                    
                    try:
                        col_idx = 5 + pd.to_datetime(row_data['fecha_compromiso']).weekday()
                        row_cells[col_idx]._tc.get_or_add_tcPr().append(parse_xml(r'<w:shd {} w:fill="217346"/>'.format(nsdecls('w'))))
                    except: pass
                    
                    try:
                        row_cells[12].text = f"{float(row_data['honorarios_estimados']):,.2f}" if float(row_data['honorarios_estimados']) > 0 else "-"
                    except: row_cells[12].text = "-"

                    for i, cell in enumerate(row_cells):
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs: run.font.size = Pt(7.5)
                            if i >= 5: paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                word_buffer = io.BytesIO()
                doc.save(word_buffer)

                st.markdown("---")
                st.markdown("### Opciones de Exportación Gantt Corporativa")
                col1, col2, col3 = st.columns(3)
                with col1: st.download_button("📥 DESCARGAR GANTT (EXCEL)", data=excel_buffer.getvalue(), file_name=f"Gantt_Planificacion_{target_week_id}.xlsx")
                with col2: st.download_button("📥 DESCARGAR REPORTE (WORD)", data=word_buffer.getvalue(), file_name=f"Reporte_Planificacion_{target_week_id}.docx")
                with col3: st.download_button("📥 DESCARGAR DATA BRUTA (CSV)", data=df_raw.to_csv(index=False).encode('utf-8-sig'), file_name=f"Data_Bruta_{target_week_id}.csv")
            else:
                st.info("No hay tareas operativas (casos) planificadas aún para esta semana.")
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
