import streamlit as st
import pandas as pd
import os
import json
import uuid
from datetime import datetime

# ---------------------------------------------------------
# BLOQUE 0: CONFIGURACIÓN GENERAL Y PERSISTENCIA
# ---------------------------------------------------------
st.set_page_config(page_title="JPV - OpsControl", layout="wide")

PERSISTENCE_DIR = "persistence"

def init_system():
    if not os.path.exists(PERSISTENCE_DIR):
        os.makedirs(PERSISTENCE_DIR)

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
        h1, h2, h3, h4 { color: #003366; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        </style>
    """, unsafe_allow_html=True)

init_system()
apply_custom_styles()

# ---------------------------------------------------------
# BLOQUE 1: FUNCIONES DE MEMORIA Y BASE DE DATOS LOCAL
# ---------------------------------------------------------
def load_master_base():
    st.sidebar.header("📁 Base Maestra")
    uploaded_file = st.sidebar.file_uploader("Cargar 'Reporte de Acciones'", type=["xlsx", "csv"])
    if uploaded_file is not None:
        try:
            # Se aplica skiprows=5 para saltar el membrete corporativo y leer encabezados reales
            if uploaded_file.name.endswith('.xlsx'):
                return pd.read_excel(uploaded_file, skiprows=5)
            else:
                return pd.read_csv(uploaded_file, skiprows=5)
        except Exception as e:
            st.sidebar.error(f"Error técnico: {e}")
    return None

def calcular_tramo_uf(fila):
    valor = 0.0
    if 'Honorarios (UF)' in fila and pd.notna(fila['Honorarios (UF)']):
        try:
            valor = float(fila['Honorarios (UF)'])
        except:
            pass
    elif 'Perdida bruta (en moneda del caso)' in fila and pd.notna(fila['Perdida bruta (en moneda del caso)']):
        try:
             valor = float(fila['Perdida bruta (en moneda del caso)'])
        except:
             pass

    if valor <= 1000:
        return "<= 1000 UF"
    elif valor <= 5000:
        return "> 1000 Y <= 5000 UF"
    else:
        return "> 5000 UF"

def load_plan_semanal(ajustador):
    week_id = get_week_identifier()
    filename = f"plan_{ajustador.replace(' ', '_')}_{week_id}.json"
    filepath = os.path.join(PERSISTENCE_DIR, filename)
    if os.path.exists(filepath):
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f), filepath
    return None, None

def save_plan_actualizado(filepath, data):
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
# ---------------------------------------------------------
# BLOQUE 2: VISTA - PLANIFICADOR SEMANAL (LUNES)
# ---------------------------------------------------------
def vista_planificador():
    st.title("🗓️ Planificador Semanal")
    st.markdown("Seleccione los casos de la Base Maestra que proyecta gestionar durante la semana en curso.")
    
    # Jerarquía extraída de los parámetros operativos
    CATALOGO_ACCIONES = {
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
            
            st.markdown("---")
            st.header("1. Selección de Casos Operativos")
            st.info(f"Inventario Vigente: {len(casos_vigentes)} casos disponibles en la Base Maestra.")
            
            selected_indices = st.multiselect(
                "Seleccione los casos que intervendrá esta semana:",
                options=casos_vigentes.index.tolist(),
                format_func=lambda x: f"Caso {casos_vigentes.loc[x, 'Número de caso']} - {casos_vigentes.loc[x, 'Asegurado']}"
            )
            
            plan_transaccional = []
            
            if selected_indices:
                st.markdown("---")
                st.header("2. Detalle de Acciones Operativas")
                st.info("💡 Seleccione la categoría y el detalle de la acción. Puede registrar hasta 3 acciones por caso.")
                
                for idx in selected_indices:
                    fila = casos_vigentes.loc[idx]
                    caso_num = fila['Número de caso']
                    asegurado = fila['Asegurado']
                    estado = fila['Estado'] if 'Estado' in fila else 'N/D'
                    subestado = fila['Sub estado'] if 'Sub estado' in fila else 'N/D'
                    tramo = calcular_tramo_uf(fila)
                    
                    with st.container():
                        st.markdown(f"""
                        <div class="marco-caso">
                            <h4>[{caso_num}] {asegurado}</h4>
                            <p style="color:gray; font-size: 0.9em; margin-bottom: 5px;"><b>Estado:</b> {estado} | <b>Sub-estado:</b> {subestado} | <b>Clasificación:</b> {tramo}</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # Permitir hasta 3 acciones jerárquicas por caso
                        for i in range(1, 4):
                            colA, colB = st.columns(2)
                            with colA:
                                cat_accion = st.selectbox(f"Categoría Acción {i} (Caso {caso_num}):", [""] + list(CATALOGO_ACCIONES.keys()), key=f"cat_{idx}_{i}")
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
                            
                            if accion_final.strip():
                                plan_transaccional.append({
                                    "id_transaccion": str(uuid.uuid4()),
                                    "categoria": "Operativa",
                                    "numero_caso": str(caso_num),
                                    "asegurado": str(asegurado),
                                    "tramo_uf": tramo,
                                    "accion": accion_final,
                                    "estado_cumplimiento": "Pendiente",
                                    "fecha_planificacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                })
                                
            st.markdown("---")
            st.header("3. Acciones de Gestión")
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown('<div class="marco-gestion"><h4>🤝 Gestión Comercial</h4></div>', unsafe_allow_html=True)
                comercial_raw = st.text_area("Reuniones, visitas a corredoras, etc. (Una por línea):", key="txt_comercial", height=150)
            
            with col2:
                st.markdown('<div class="marco-gestion"><h4>⚙️ Gestión Administrativa</h4></div>', unsafe_allow_html=True)
                admin_raw = st.text_area("Capacitaciones, comités, trámites, etc. (Una por línea):", key="txt_admin", height=150)
            
            if comercial_raw.strip():
                for accion in [linea.strip() for linea in comercial_raw.split('\n') if linea.strip()]:
                    plan_transaccional.append({
                        "id_transaccion": str(uuid.uuid4()), "categoria": "Gestión Comercial", "numero_caso": "N/A", "asegurado": "N/A", "tramo_uf": "N/A", "accion": accion, "estado_cumplimiento": "Pendiente", "fecha_planificacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    })
                    
            if admin_raw.strip():
                for accion in [linea.strip() for linea in admin_raw.split('\n') if linea.strip()]:
                    plan_transaccional.append({
                        "id_transaccion": str(uuid.uuid4()), "categoria": "Gestión Administrativa", "numero_caso": "N/A", "asegurado": "N/A", "tramo_uf": "N/A", "accion": accion, "estado_cumplimiento": "Pendiente", "fecha_planificacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    })

            st.markdown("---")
            if len(plan_transaccional) > 0:
                st.info(f"Se han registrado **{len(plan_transaccional)} acciones**.")
                if st.button("💾 COMPROMETER PLAN SEMANAL"):
                    try:
                        week_id = get_week_identifier()
                        filename = f"plan_{ajustador_seleccionado.replace(' ', '_')}_{week_id}.json"
                        filepath = os.path.join(PERSISTENCE_DIR, filename)
                        with open(filepath, 'w', encoding='utf-8') as f:
                            json.dump(plan_transaccional, f, ensure_ascii=False, indent=4)
                        st.success("Plan guardado exitosamente.")
                    except Exception as e:
                        st.error(f"Error: {e}")
            elif selected_indices:
                st.warning("Debe seleccionar al menos una acción válida para guardar el plan.")
    else:
        st.info("Módulo en espera: Suba el archivo 'Reporte de acciones' en el panel izquierdo.")

# ---------------------------------------------------------
# BLOQUE 3: VISTA - PROGRAMA DIARIO (MAR a VIE)
# ---------------------------------------------------------
def vista_diario():
    st.title("☀️ Ejecución Diaria")
    st.markdown(f"**Fecha actual:** {datetime.now().strftime('%A, %d de %B de %Y')}")
    
    ajustador_input = st.text_input("Ingrese su nombre de Ajustador (Exacto al Plan Semanal):", placeholder="Ej: Francisco Silva Ghisolfo")
    
    if ajustador_input:
        plan_data, filepath = load_plan_semanal(ajustador_input)
        
        if plan_data is None:
            st.warning(f"⚠️ No se encontró un Plan Semanal activo para **{ajustador_input}** en la semana en curso.")
        else:
            st.success("Plan Semanal sincronizado correctamente.")
            st.markdown("---")
            st.header("📋 Panel de Cumplimiento (Control PM)")
            
            total_tareas = len(plan_data)
            tareas_completadas = 0
            cambios_realizados = False
            
            with st.form(key="form_cumplimiento"):
                for idx, tarea in enumerate(plan_data):
                    estado_actual = tarea.get("estado_cumplimiento", "Pendiente")
                    es_realizado = (estado_actual == "Realizado")
                    
                    if es_realizado:
                        tareas_completadas += 1
                        clase_css = "tarea-marco tarea-realizada"
                        icono = "✅"
                    else:
                        clase_css = "tarea-marco"
                        icono = "⏳"
                    
                    st.markdown(f'<div class="{clase_css}">', unsafe_allow_html=True)
                    
                    if tarea["categoria"] == "Operativa":
                        st.markdown(f"**{icono} CASO [{tarea['numero_caso']}]** - {tarea['asegurado']}")
                        st.markdown(f"<span style='color:gray; font-size:0.9em;'>Tramo: {tarea['tramo_uf']}</span>", unsafe_allow_html=True)
                    else:
                        st.markdown(f"**{icono} {tarea['categoria'].upper()}**")
                    
                    st.markdown(f"**Acción:** {tarea['accion']}")
                    nuevo_estado = st.checkbox(f"Marcar como Realizado", value=es_realizado, key=f"chk_{tarea['id_transaccion']}")
                    
                    nuevo_texto_estado = "Realizado" if nuevo_estado else "Pendiente"
                    if nuevo_texto_estado != estado_actual:
                        plan_data[idx]["estado_cumplimiento"] = nuevo_texto_estado
                        plan_data[idx]["fecha_ejecucion"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if nuevo_estado else ""
                        cambios_realizados = True
                        
                    st.markdown('</div>', unsafe_allow_html=True)

                st.markdown('<div class="btn-guardar">', unsafe_allow_html=True)
                submit_button = st.form_submit_button(label="💾 ACTUALIZAR CUMPLIMIENTO DIARIO")
                st.markdown('</div>', unsafe_allow_html=True)
            
            if submit_button:
                if cambios_realizados:
                    try:
                        save_plan_actualizado(filepath, plan_data)
                        st.success("¡Cumplimiento actualizado!")
                        st.experimental_rerun()
                    except Exception as e:
                        st.error(f"Error al escribir en el disco: {e}")
                else:
                    st.info("No se detectaron cambios.")
            
            st.markdown("---")
            progreso = int((tareas_completadas / total_tareas) * 100) if total_tareas > 0 else 0
            st.progress(progreso)
            st.caption(f"Avance de la semana: {tareas_completadas} de {total_tareas} tareas realizadas ({progreso}%).")

# ---------------------------------------------------------
# BLOQUE PRINCIPAL: ENRUTADOR DE NAVEGACIÓN
# ---------------------------------------------------------
def main():
    st.sidebar.image("https://img.icons8.com/color/96/000000/engineering.png", width=60) # Icono genérico de ingeniería
    st.sidebar.title("Navegación OpsControl")
    
    # Selector de pantalla
    opcion = st.sidebar.radio(
        "Ir a:",
        ["Planificador Semanal", "Programa Diario"]
    )
    
    st.sidebar.markdown("---")
    
    # Enrutamiento según la opción seleccionada
    if opcion == "Planificador Semanal":
        vista_planificador()
    elif opcion == "Programa Diario":
        vista_diario()

if __name__ == "__main__":
    main()
