import streamlit as st
import pandas as pd
import os
import json
import uuid
from datetime import datetime

# ---------------------------------------------------------
# BLOQUE 0: CONFIGURACIÓN Y PERSISTENCIA (SIN CAMBIOS)
# ---------------------------------------------------------
st.set_page_config(page_title="JPV - OpsControl Semanal", layout="wide")

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
        .stDataFrame { border: 1px solid #c4ced4; }
        .marco-caso { background-color: white; padding: 15px; border-radius: 5px; border-left: 5px solid #217346; margin-bottom: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        .marco-gestion { background-color: white; padding: 15px; border-radius: 5px; border-left: 5px solid #004a99; margin-bottom: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        h1, h2, h3, h4 { color: #003366; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        </style>
    """, unsafe_allow_html=True)

init_system()
apply_custom_styles()

# ---------------------------------------------------------
# BLOQUE 1: EXTRACCIÓN Y LÓGICA DE TRAMOS (MOTOR SILENCIOSO)
# ---------------------------------------------------------
def load_master_base():
    st.sidebar.header("📁 Carga de Reporte de Acciones")
    uploaded_file = st.sidebar.file_uploader("Seleccione el archivo (Excel/CSV)", type=["xlsx", "csv"])
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.xlsx'):
                df = pd.read_excel(uploaded_file)
            else:
                df = pd.read_csv(uploaded_file)
            st.sidebar.success(f"Archivo cargado: {uploaded_file.name}")
            return df
        except Exception as e:
            st.sidebar.error(f"Error técnico al leer el archivo: {e}")
            return None
    return None

def calcular_tramo_uf(fila):
    """
    CRITERIO 2: Clasificación silenciosa de tramos.
    Busca 'Honorarios (UF)' o 'Perdida bruta' y clasifica.
    """
    valor = 0.0
    if 'Honorarios (UF)' in fila and pd.notna(fila['Honorarios (UF)']):
        try:
            valor = float(fila['Honorarios (UF)'])
        except:
            pass
    elif 'Perdida bruta (en moneda del caso)' in fila and pd.notna(fila['Perdida bruta (en moneda del caso)']):
        # Fallback si solo hay pérdida bruta (asumiendo valor referencial si es necesario)
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

# ---------------------------------------------------------
# BLOQUE 2: INTERFAZ Y CONSTRUCCIÓN TRANSACCIONAL
# ---------------------------------------------------------
def main_planificador():
    st.title("🗓️ JPV OpsControl - Planificador Semanal")
    
    df_maestro = load_master_base()
    
    if df_maestro is not None:
        col_ajustador = 'Ajustador senior' if 'Ajustador senior' in df_maestro.columns else df_maestro.columns[9]
        ajustadores_validos = sorted(df_maestro[col_ajustador].dropna().unique())
        
        ajustador_seleccionado = st.selectbox("Identificación de Ajustador:", [""] + ajustadores_validos)
        
        if ajustador_seleccionado:
            # CRITERIO 2: Filtro estricto por ajustador asignado
            casos_vigentes = df_maestro[df_maestro[col_ajustador] == ajustador_seleccionado].copy()
            
            st.markdown("---")
            st.header("1. Selección de Casos Operativos")
            
            selected_indices = st.multiselect(
                "Seleccione los casos que intervendrá esta semana:",
                options=casos_vigentes.index.tolist(),
                format_func=lambda x: f"Caso {casos_vigentes.loc[x, 'Número de caso']} - {casos_vigentes.loc[x, 'Asegurado']}"
            )
            
            # Contenedores para almacenar la data transaccional que construiremos
            plan_transaccional = []
            
            if selected_indices:
                st.markdown("---")
                st.header("2. Detalle de Acciones Operativas")
                st.info("💡 Escriba las acciones a realizar. Si hay múltiples acciones para un mismo caso, escríbalas en líneas separadas (Enter).")
                
                for idx in selected_indices:
                    fila = casos_vigentes.loc[idx]
                    caso_num = fila['Número de caso']
                    asegurado = fila['Asegurado']
                    estado = fila['Estado'] if 'Estado' in fila else 'N/D'
                    subestado = fila['Sub estado'] if 'Sub estado' in fila else 'N/D'
                    tramo = calcular_tramo_uf(fila)
                    
                    # CRITERIO 5: Mostrar estado para contexto sin bloquear nada
                    with st.container():
                        st.markdown(f"""
                        <div class="marco-caso">
                            <h4>[{caso_num}] {asegurado}</h4>
                            <p style="color:gray; font-size: 0.9em; margin-bottom: 5px;"><b>Estado:</b> {estado} | <b>Sub-estado:</b> {subestado} | <b>Clasificación:</b> {tramo}</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # CRITERIO 1 y 4: Texto libre, múltiples líneas = múltiples acciones
                        acciones_raw = st.text_area(f"Acciones para el caso {caso_num}:", key=f"txt_{idx}", height=100, 
                                                    placeholder="Ej:\nInspección a la planta\nRedacción de Informe Inicial")
                        
                        if acciones_raw.strip():
                            lineas_acciones = [linea.strip() for linea in acciones_raw.split('\n') if linea.strip()]
                            for accion in lineas_acciones:
                                plan_transaccional.append({
                                    "id_transaccion": str(uuid.uuid4()),
                                    "categoria": "Operativa",
                                    "numero_caso": str(caso_num),
                                    "asegurado": str(asegurado),
                                    "tramo_uf": tramo,
                                    "accion": accion,
                                    "estado_cumplimiento": "Pendiente",
                                    "fecha_planificacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                })

            # CRITERIO 3: Divisiones para Gestión Comercial y Administrativa
            st.markdown("---")
            st.header("3. Acciones de Gestión (No asociadas a Siniestros)")
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown('<div class="marco-gestion"><h4>🤝 Gestión Comercial</h4></div>', unsafe_allow_html=True)
                comercial_raw = st.text_area("Reuniones, visitas a corredoras, etc. (Una por línea):", key="txt_comercial", height=150)
            
            with col2:
                st.markdown('<div class="marco-gestion"><h4>⚙️ Gestión Administrativa</h4></div>', unsafe_allow_html=True)
                admin_raw = st.text_area("Capacitaciones, comités, trámites, etc. (Una por línea):", key="txt_admin", height=150)
            
            # Procesar Gestión Comercial
            if comercial_raw.strip():
                for accion in [linea.strip() for linea in comercial_raw.split('\n') if linea.strip()]:
                    plan_transaccional.append({
                        "id_transaccion": str(uuid.uuid4()),
                        "categoria": "Gestión Comercial",
                        "numero_caso": "N/A",
                        "asegurado": "N/A",
                        "tramo_uf": "N/A",
                        "accion": accion,
                        "estado_cumplimiento": "Pendiente",
                        "fecha_planificacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    })
                    
            # Procesar Gestión Administrativa
            if admin_raw.strip():
                for accion in [linea.strip() for linea in admin_raw.split('\n') if linea.strip()]:
                    plan_transaccional.append({
                        "id_transaccion": str(uuid.uuid4()),
                        "categoria": "Gestión Administrativa",
                        "numero_caso": "N/A",
                        "asegurado": "N/A",
                        "tramo_uf": "N/A",
                        "accion": accion,
                        "estado_cumplimiento": "Pendiente",
                        "fecha_planificacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    })

            # Botón de Guardado Final
            st.markdown("---")
            if len(plan_transaccional) > 0:
                st.info(f"Se han registrado **{len(plan_transaccional)} acciones en total** para esta semana.")
                if st.button("💾 COMPROMETER PLAN Y GENERAR BASE DE DATOS LOCAL"):
                    try:
                        week_id = get_week_identifier()
                        filename = f"plan_{ajustador_seleccionado.replace(' ', '_')}_{week_id}.json"
                        filepath = os.path.join(PERSISTENCE_DIR, filename)
                        
                        with open(filepath, 'w', encoding='utf-8') as f:
                            json.dump(plan_transaccional, f, ensure_ascii=False, indent=4)
                            
                        st.success(f"¡Plan guardado exitosamente!")
                        st.code(f"Ruta: {filepath}\nEstructura lista para ser consumida por el Programa Diario.", language="text")
                    except Exception as e:
                        st.error(f"Error al escribir en disco: {e}")
            elif selected_indices:
                st.warning("Debe escribir al menos una acción en los casos seleccionados o en los cuadros de gestión para guardar el plan.")
        else:
            st.warning("Selección requerida para iniciar el filtro.")
    else:
        st.info("Módulo en espera: Suba el archivo 'Reporte de acciones' en el panel izquierdo.")

if __name__ == "__main__":
    main_planificador()