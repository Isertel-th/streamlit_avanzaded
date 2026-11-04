import streamlit as st 
import pandas as pd 
import os 
import plotly.express as px 
import numpy as np
from datetime import datetime, timedelta 
import io 

# --- FUNCIÓN DE COMPACIDAD Y CONFIGURACIÓN --- 
def set_page_config_and_style(): 
# 1. Configurar layout en modo ancho ("wide") y título 
    st.set_page_config(layout="wide", page_title="Estadístico Isertel")

# 2. Custom CSS para máxima compacidad y minimalismo
    st.markdown(""" 
    <style> 
    /* Ahorro vertical general */ 
    .block-container { 
        padding-top: 4rem !important; 
        padding-bottom: 0rem !important; 
        padding-left: 1rem !important; 
        padding-right: 1rem !important; 
    }
    div[data-testid="stHorizontalBlock"] { gap: 0.75rem !important; }
    div[data-testid="stVerticalBlock"] > div:has(div[data-testid="stContainer"]) > div[data-testid="stContainer"] { 
        padding: 0.5rem !important; 
    }
    h3, h4, h5 { margin-top: 0.1rem !important; margin-bottom: 0.1rem !important; }
    hr { margin-top: 0.1rem !important; margin-bottom: 0.1rem !important; }
    .stSelectbox, .stMultiSelect, .stDateInput, div[data-testid="stForm"] { margin-bottom: 0.0rem !important; }
    div[data-testid="stMetric"] { padding: 0.2rem 0 !important; }
    div[data-testid="stMetricLabel"] { font-size: 1rem; }

    /* ESTILOS DE MÉTRICAS COMPACTAS */
    .metric-compact-container div[data-testid="stMetricValue"] { font-size: 1.8rem; color: #B71C1C; } 
    .metric-compact-container-total div[data-testid="stMetricValue"] { font-size: 1.8rem; color: #0D47A1; }
    div[data-testid="stMetricDelta"] { visibility: hidden; height: 0; }

    /* Header Delgado */ 
    div[data-testid="stSuccess"] { 
        padding: 0.5rem 1rem !important; margin-bottom: 0px; display: flex; 
        justify-content: flex-end; align-items: center; height: 100%; 
    } 
    .stButton>button { height: 30px; padding-top: 5px !important; padding-bottom: 5px !important; }

    /* Tablas Compactas */ 
    .stDataFrame { margin-top: 0.5rem; } 
    .stDataFrame .css-1dp5fcv { padding: 0.2rem 0.5rem; } 
    .stDataFrame .css-1dp5fcv button { padding: 0.1rem 0.4rem; font-size: 0.8rem; }
    </style> 
    """, unsafe_allow_html=True)

# Llama a la función al inicio de tu script 
set_page_config_and_style()

# --- CONFIGURACIÓN DE ARCHIVOS Y CARPETAS --- 
MASTER_EXCEL = "datos.xlsx" 
USUARIOS_EXCEL = "usuarios.xlsx" 
UPLOAD_FOLDER = "ExcelUploads" 
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 1. DEFINICIÓN FINAL DEL MAPEO 
MAPEO_COLUMNAS = { 
    'FECHA': 'A', 
    'UBICACIÓN': 'B', 
    'TÉCNICO': 'C', 
    'CONTRATO': 'D', 
    'CUENTA': 'E', 
    'MOTIVO DE INGRESO': 'F', 
    'CLIENTE': 'G', 
    'TECNOLOGÍA': 'H', 
    'TAREA': 'I', 
    'ESTADO TAREA': 'J', 
    'SOLUCIÓN': 'K', 
    'TIPO DE ORDEN': 'L', 
    'TIPO TAREA MANUAL': 'M'
}
COLUMNAS_SELECCIONADAS = list(MAPEO_COLUMNAS.values()) 
ENCABEZADOS_ESPERADOS = list(MAPEO_COLUMNAS.keys())
FINAL_RENAMING_MAP = {v: k for k, v in MAPEO_COLUMNAS.items()}

# CLAVES DE COLUMNA
COL_FECHA_KEY = 'A' 
COL_CIUDAD_KEY = 'B' 
COL_TECNICO_KEY = 'C' 
COL_CONTRATO_KEY = 'D'
COL_CUENTA_KEY = 'E'
COL_MOTIVO_KEY = 'F'
COL_CLIENTE_KEY = 'G'
COL_TECNOLOGIA_KEY = 'H'
COL_TAREA_KEY = 'I'
COL_ESTADO_KEY = 'J' 
COL_SOLUCION_KEY = 'K'
COL_TIPO_ORDEN_KEY = 'L'
COL_TIPO_MANUAL_KEY = 'M'

# DESCRIPCIONES DE COLUMNAS
COL_FECHA_DESCRIPTIVA = FINAL_RENAMING_MAP[COL_FECHA_KEY] 
COL_TEMP_DATETIME = '_DATETIME_' + COL_FECHA_KEY

COL_TECNICO_DESCRIPTIVA = FINAL_RENAMING_MAP.get(COL_TECNICO_KEY, 'TÉCNICO') 
COL_CIUDAD_DESCRIPTIVA = FINAL_RENAMING_MAP.get(COL_CIUDAD_KEY, 'UBICACIÓN') 
COL_CUENTA_DESCRIPTIVA = FINAL_RENAMING_MAP.get(COL_CUENTA_KEY, 'CUENTA')
COL_MOTIVO_DESCRIPTIVA = FINAL_RENAMING_MAP.get(COL_MOTIVO_KEY, 'MOTIVO DE INGRESO')
COL_TECNOLOGIA_DESCRIPTIVA = FINAL_RENAMING_MAP.get(COL_TECNOLOGIA_KEY, 'TECNOLOGÍA')
COL_ESTADO_DESCRIPTIVA = FINAL_RENAMING_MAP.get(COL_ESTADO_KEY, 'ESTADO TAREA')
COL_SOLUCION_DESCRIPTIVA = FINAL_RENAMING_MAP.get(COL_SOLUCION_KEY, 'SOLUCIÓN')
COL_TIPO_ORDEN_DESCRIPTIVA = FINAL_RENAMING_MAP.get(COL_TIPO_ORDEN_KEY, 'TIPO DE ORDEN')
COL_TIPO_MANUAL_DESCRIPTIVA = FINAL_RENAMING_MAP.get(COL_TIPO_MANUAL_KEY, 'TIPO TAREA MANUAL')

# Columnas temporales para filtros
COL_FILTRO_TECNICO = '_Filtro_Tecnico_' 
COL_FILTRO_CIUDAD = '_Filtro_Ubicacion_'
COL_FILTRO_ESTADO = '_Filtro_Estado_' 
COL_FILTRO_TIPO_ORDEN = '_Filtro_TipoOrden_'
COL_FILTRO_TECNOLOGIA = '_Filtro_Tecnologia_'
COL_FILTRO_TIPO_MANUAL = '_Filtro_TipoManual_'

# Columnas para Gráficos
COL_TIPO_INST = '_ES_INSTALACION_' 
COL_TIPO_VISITA = '_ES_VISITA_'
COL_TIPO_MIGRACION = '_ES_MIGRACION_'
COL_TIPO_MANUAL = '_ES_TAREA_MANUAL_'
COL_TIPO_CAMBIO_DIR = '_ES_CAMBIO_DIRECCION_'
COL_TIPO_MIGRATEC = '_ES_MIGRATEC_'


# --- FUNCIONES DE LIMPIEZA --- 
@st.cache_data 
def clean_tecnico(tecnico): 
    s = str(tecnico).strip()
    if '|' in s: s = s.split('|', 1)[1].strip() 
    suffix = ' (tecnico)'
    if s.lower().endswith(suffix): s = s[:-len(suffix)]
    return s.strip()

@st.cache_data 
def clean_ciudad(ciudad): 
    if isinstance(ciudad, str) and ',' in ciudad: return ciudad.split(',', 1)[0].strip() 
    return str(ciudad).strip()

# --- FUNCIONES DE COMPARACIÓN Y GRÁFICOS (MODO ESTANDAR) --- 
@st.cache_data 
def prepare_comparison_data(df): 
    # Agrupación por CIUDAD y TÉCNICO
    if df.empty: return pd.DataFrame()
    df_temp = df.copy()
    
    # Flags de tipos
    if COL_TIPO_ORDEN_KEY in df_temp.columns: 
        tipo_orden = df_temp[COL_TIPO_ORDEN_KEY].astype(str)
        df_temp[COL_TIPO_INST] = tipo_orden.str.contains('INSTALACION', case=False, na=False).astype(int) 
        df_temp[COL_TIPO_VISITA] = tipo_orden.str.contains('VISITA TECNICA', case=False, na=False).astype(int)
        mask_mig_orden = tipo_orden.str.contains(r'MIGRACI[ÓO]N', case=False, na=False, regex=True)
        mask_mig_manual = False
        if COL_TIPO_MANUAL_KEY in df_temp.columns:
            mask_mig_manual = df_temp[COL_TIPO_MANUAL_KEY].astype(str).str.contains(r'MIGRACI[ÓO]N', case=False, na=False, regex=True)
        df_temp[COL_TIPO_MIGRACION] = (mask_mig_orden | mask_mig_manual).astype(int)
        df_temp[COL_TIPO_MANUAL] = tipo_orden.str.contains('TAREA MANUAL', case=False, na=False).astype(int)
        df_temp[COL_TIPO_CAMBIO_DIR] = tipo_orden.str.contains(r'CAMBIO DE DIRECCI[ÓO]N', case=False, na=False, regex=True).astype(int)
        df_temp[COL_TIPO_MIGRATEC] = tipo_orden.str.contains('MIGRATEC', case=False, na=False).astype(int)
    else: 
        df_temp[COL_TIPO_INST] = 0; df_temp[COL_TIPO_VISITA] = 0; df_temp[COL_TIPO_MIGRACION] = 0
        df_temp[COL_TIPO_MANUAL] = 0; df_temp[COL_TIPO_CAMBIO_DIR] = 0; df_temp[COL_TIPO_MIGRATEC] = 0

    if COL_FILTRO_TECNICO not in df_temp.columns or COL_FILTRO_CIUDAD not in df_temp.columns: return pd.DataFrame()

    df_grouped = df_temp.groupby([COL_FILTRO_CIUDAD, COL_FILTRO_TECNICO]).agg( 
        Total_Instalaciones=(COL_TIPO_INST, 'sum'), 
        Total_Visitas=(COL_TIPO_VISITA, 'sum'),
        Total_Migracion=(COL_TIPO_MIGRACION, 'sum'),
        Total_TareaManual=(COL_TIPO_MANUAL, 'sum'),
        Total_CambioDireccion=(COL_TIPO_CAMBIO_DIR, 'sum'),
        Total_MigraTec=(COL_TIPO_MIGRATEC, 'sum'),
        Total_Tareas=(COL_TIPO_INST, 'count')
    ).reset_index()

    # Convertir a int
    cols_int = ['Total_Instalaciones','Total_Visitas','Total_Migracion','Total_TareaManual','Total_CambioDireccion','Total_MigraTec','Total_Tareas']
    for c in cols_int: df_grouped[c] = df_grouped[c].astype(int)

    return df_grouped.sort_values(by=COL_FILTRO_TECNICO)

@st.cache_data 
def prepare_city_comparison_data(df): 
    if df.empty: return pd.DataFrame()
    df_temp = df.copy()
    if COL_TIPO_ORDEN_KEY in df_temp.columns: 
        tipo_orden = df_temp[COL_TIPO_ORDEN_KEY].astype(str)
        df_temp[COL_TIPO_INST] = tipo_orden.str.contains('INSTALACION', case=False, na=False).astype(int) 
        df_temp[COL_TIPO_VISITA] = tipo_orden.str.contains('VISITA TECNICA', case=False, na=False).astype(int)
        mask_mig_orden = tipo_orden.str.contains(r'MIGRACI[ÓO]N', case=False, na=False, regex=True)
        mask_mig_manual = False
        if COL_TIPO_MANUAL_KEY in df_temp.columns:
            mask_mig_manual = df_temp[COL_TIPO_MANUAL_KEY].astype(str).str.contains(r'MIGRACI[ÓO]N', case=False, na=False, regex=True)
        df_temp[COL_TIPO_MIGRACION] = (mask_mig_orden | mask_mig_manual).astype(int)
        df_temp[COL_TIPO_MANUAL] = tipo_orden.str.contains('TAREA MANUAL', case=False, na=False).astype(int)
        df_temp[COL_TIPO_CAMBIO_DIR] = tipo_orden.str.contains(r'CAMBIO DE DIRECCI[ÓO]N', case=False, na=False, regex=True).astype(int)
        df_temp[COL_TIPO_MIGRATEC] = tipo_orden.str.contains('MIGRATEC', case=False, na=False).astype(int)
    else: 
        df_temp[COL_TIPO_INST] = 0; df_temp[COL_TIPO_VISITA] = 0; df_temp[COL_TIPO_MIGRACION] = 0
        df_temp[COL_TIPO_MANUAL] = 0; df_temp[COL_TIPO_CAMBIO_DIR] = 0; df_temp[COL_TIPO_MIGRATEC] = 0

    if COL_FILTRO_CIUDAD not in df_temp.columns: return pd.DataFrame()

    df_grouped = df_temp.groupby([COL_FILTRO_CIUDAD]).agg( 
        Total_Instalaciones=(COL_TIPO_INST, 'sum'), 
        Total_Visitas=(COL_TIPO_VISITA, 'sum'),
        Total_Migracion=(COL_TIPO_MIGRACION, 'sum'),
        Total_TareaManual=(COL_TIPO_MANUAL, 'sum'),
        Total_CambioDireccion=(COL_TIPO_CAMBIO_DIR, 'sum'),
        Total_MigraTec=(COL_TIPO_MIGRATEC, 'sum'),
    ).reset_index()

    cols_int = ['Total_Instalaciones','Total_Visitas','Total_Migracion','Total_TareaManual','Total_CambioDireccion','Total_MigraTec']
    for c in cols_int: df_grouped[c] = df_grouped[c].astype(int)

    return df_grouped.sort_values(by=COL_FILTRO_CIUDAD)

@st.cache_data
def prepare_technician_comparison_data(df):
    if df.empty: return pd.DataFrame()
    df_temp = df.copy()
    if COL_TIPO_ORDEN_KEY in df_temp.columns: 
        tipo_orden = df_temp[COL_TIPO_ORDEN_KEY].astype(str)
        df_temp[COL_TIPO_INST] = tipo_orden.str.contains('INSTALACION', case=False, na=False).astype(int) 
        df_temp[COL_TIPO_VISITA] = tipo_orden.str.contains('VISITA TECNICA', case=False, na=False).astype(int)
        mask_mig_orden = tipo_orden.str.contains(r'MIGRACI[ÓO]N', case=False, na=False, regex=True)
        mask_mig_manual = False
        if COL_TIPO_MANUAL_KEY in df_temp.columns:
            mask_mig_manual = df_temp[COL_TIPO_MANUAL_KEY].astype(str).str.contains(r'MIGRACI[ÓO]N', case=False, na=False, regex=True)
        df_temp[COL_TIPO_MIGRACION] = (mask_mig_orden | mask_mig_manual).astype(int)
        df_temp[COL_TIPO_MANUAL] = tipo_orden.str.contains('TAREA MANUAL', case=False, na=False).astype(int)
        df_temp[COL_TIPO_CAMBIO_DIR] = tipo_orden.str.contains(r'CAMBIO DE DIRECCI[ÓO]N', case=False, na=False, regex=True).astype(int)
        df_temp[COL_TIPO_MIGRATEC] = tipo_orden.str.contains('MIGRATEC', case=False, na=False).astype(int)
    else: 
        df_temp[COL_TIPO_INST] = 0; df_temp[COL_TIPO_VISITA] = 0; df_temp[COL_TIPO_MIGRACION] = 0
        df_temp[COL_TIPO_MANUAL] = 0; df_temp[COL_TIPO_CAMBIO_DIR] = 0; df_temp[COL_TIPO_MIGRATEC] = 0

    if COL_FILTRO_TECNICO not in df_temp.columns: return pd.DataFrame()

    df_grouped = df_temp.groupby([COL_FILTRO_TECNICO]).agg( 
        Total_Instalaciones=(COL_TIPO_INST, 'sum'), 
        Total_Visitas=(COL_TIPO_VISITA, 'sum'),
        Total_Migracion=(COL_TIPO_MIGRACION, 'sum'),
        Total_TareaManual=(COL_TIPO_MANUAL, 'sum'),
        Total_CambioDireccion=(COL_TIPO_CAMBIO_DIR, 'sum'),
        Total_MigraTec=(COL_TIPO_MIGRATEC, 'sum'),
    ).reset_index()

    cols_int = ['Total_Instalaciones','Total_Visitas','Total_Migracion','Total_TareaManual','Total_CambioDireccion','Total_MigraTec']
    for c in cols_int: df_grouped[c] = df_grouped[c].astype(int)

    return df_grouped.sort_values(by=COL_FILTRO_TECNICO)

# --- FUNCIÓN CORREGIDA (FIX PARA SESSION_STATE) ---
def st_multiselect_with_all_technicians(col, label, options, key):
    ALL_OPTION = "✨ Seleccionar Todos"
    SUP_OPTION = "👷 Seleccionar Supervisores"
    
    # 1. Verificar existencia de supervisores en las opciones FILTRADAS
    hay_supervisores = any(str(opt).strip().upper().startswith("SUP. ") for opt in options)
    
    # 2. Construir lista de visualización
    display_options = [ALL_OPTION]
    if hay_supervisores:
        display_options.append(SUP_OPTION)
    display_options += options

    # 3. Callback interno para manejar la lógica ANTES de que el widget termine de renderizarse
    def on_change_handler():
        current_selection = st.session_state[key]
        if SUP_OPTION in current_selection:
            # Calcular INTERSECCIÓN: Solo los supervisores que ESTÁN en las opciones disponibles
            valid_supervisors = [opt for opt in options if str(opt).strip().upper().startswith("SUP. ")]
            st.session_state[key] = valid_supervisors
        elif ALL_OPTION in current_selection:
            st.session_state[key] = options

    with col:
        if not options:
            st.markdown(f"**{label}**")
            st.info("No hay técnicos disponibles.", icon="🧑‍🔧")
            return []
        
        # 4. Renderizar widget con callback
        st.multiselect(
            label=label, 
            options=display_options, 
            key=key, 
            on_change=on_change_handler # <--- AQUÍ ESTÁ LA SOLUCIÓN
        )

    # 5. Devolver selección limpia para el resto del script
    # Se obtienen los datos directamente del session state ya actualizado por el callback
    final_selection = st.session_state.get(key, [])
    return [s for s in final_selection if s not in (ALL_OPTION, SUP_OPTION)]

@st.cache_data
def prepare_date_comparison_data(df):
    if df.empty or COL_TEMP_DATETIME not in df.columns: return pd.DataFrame()
    df_temp = df.copy()
    COL_FECHA_DIA_AGRUPACION = '_FECHA_DIA_'
    df_temp[COL_FECHA_DIA_AGRUPACION] = df_temp[COL_TEMP_DATETIME].dt.date

    if COL_TIPO_ORDEN_KEY in df_temp.columns: 
        tipo_orden = df_temp[COL_TIPO_ORDEN_KEY].astype(str)
        df_temp[COL_TIPO_INST] = tipo_orden.str.contains('INSTALACION', case=False, na=False).astype(int) 
        df_temp[COL_TIPO_VISITA] = tipo_orden.str.contains('VISITA TECNICA', case=False, na=False).astype(int)
        mask_mig_orden = tipo_orden.str.contains(r'MIGRACI[ÓO]N', case=False, na=False, regex=True)
        mask_mig_manual = False
        if COL_TIPO_MANUAL_KEY in df_temp.columns:
             mask_mig_manual = df_temp[COL_TIPO_MANUAL_KEY].astype(str).str.contains(r'MIGRACI[ÓO]N', case=False, na=False, regex=True)
        df_temp[COL_TIPO_MIGRACION] = (mask_mig_orden | mask_mig_manual).astype(int)
        df_temp[COL_TIPO_MANUAL] = tipo_orden.str.contains('TAREA MANUAL', case=False, na=False).astype(int)
        df_temp[COL_TIPO_CAMBIO_DIR] = tipo_orden.str.contains(r'CAMBIO DE DIRECCI[ÓO]N', case=False, na=False, regex=True).astype(int)
        df_temp[COL_TIPO_MIGRATEC] = tipo_orden.str.contains('MIGRATEC', case=False, na=False).astype(int)
    else: 
        df_temp[COL_TIPO_INST] = 0; df_temp[COL_TIPO_VISITA] = 0; df_temp[COL_TIPO_MIGRACION] = 0 
        df_temp[COL_TIPO_MANUAL] = 0; df_temp[COL_TIPO_CAMBIO_DIR] = 0; df_temp[COL_TIPO_MIGRATEC] = 0

    df_grouped = df_temp.groupby([COL_FECHA_DIA_AGRUPACION]).agg( 
        Total_Instalaciones=(COL_TIPO_INST, 'sum'), 
        Total_Visitas=(COL_TIPO_VISITA, 'sum'),
        Total_Migracion=(COL_TIPO_MIGRACION, 'sum'),
        Total_TareaManual=(COL_TIPO_MANUAL, 'sum'),
        Total_CambioDireccion=(COL_TIPO_CAMBIO_DIR, 'sum'),
        Total_MigraTec=(COL_TIPO_MIGRATEC, 'sum'),
    ).reset_index()

    cols_int = ['Total_Instalaciones','Total_Visitas','Total_Migracion','Total_TareaManual','Total_CambioDireccion','Total_MigraTec']
    for c in cols_int: df_grouped[c] = df_grouped[c].astype(int)
    return df_grouped.sort_values(by=COL_FECHA_DIA_AGRUPACION)

def render_comparison_charts_vertical(df_comparacion, x_col, title_prefix, is_city_view=False):
    chart_configs = [
        {'col_name': 'Total_Instalaciones', 'title': 'Instalaciones', 'color': '#4CAF50'},
        {'col_name': 'Total_Visitas', 'title': 'Visitas', 'color': '#FF9800'},
        {'col_name': 'Total_Migracion', 'title': 'Migración', 'color': '#2196F3'},
        {'col_name': 'Total_TareaManual', 'title': 'Tarea Manual', 'color': '#9C27B0'},
        {'col_name': 'Total_CambioDireccion', 'title': 'Cambio de Dirección', 'color': '#F44336'},
        {'col_name': 'Total_MigraTec', 'title': 'MigraTec', 'color': '#00BCD4'}
    ]

    st.markdown(f"#### Rendimiento {title_prefix} (Base Dinámica)")
    bottom_margin = 60
    CHART_HEIGHT = 200 
    xaxis_config = { 'tickangle': -45, 'tickfont': {'size': 9 if not is_city_view else 10} }
    grid_config = { 'showgrid': True, 'gridcolor': '#cccccc', 'griddash': 'dot' }

    for config in chart_configs:
        with st.container(border=True):
            st.markdown(f"##### {config['title']}")
            fig = px.line(
                df_comparacion, 
                x=x_col, 
                y=config['col_name'], 
                markers=True, 
                text=config['col_name'], 
                height=CHART_HEIGHT,
                color_discrete_sequence=[config['color']]
            ) 
            fig.update_traces(textposition='top center') 
            fig.update_layout(
                xaxis_title=None, 
                yaxis_title='Total', 
                margin=dict(t=20,b=bottom_margin,l=10,r=10), 
                xaxis=xaxis_config 
            )
            fig.update_xaxes(**grid_config)
            fig.update_yaxes(showgrid=False) 
            st.plotly_chart(fig, use_container_width=True)

# --- LECTURA DE USUARIOS ---
try: 
    usuarios_df = pd.read_excel(USUARIOS_EXCEL) 
    usuarios_df['Usuario'] = usuarios_df['Usuario'].astype(str).str.strip() 
    usuarios_df['Contraseña'] = usuarios_df['Contraseña'].astype(str).str.strip() 
    usuarios_df['Rol'] = usuarios_df['Rol'].astype(str).str.strip() 
except FileNotFoundError: 
    usuarios_data = { 'Usuario': ['admin', 'user'], 'Contraseña': ['12345', 'password'], 'Rol': ['admin', 'analyst'] } 
    usuarios_df = pd.DataFrame(usuarios_data) 

# --- SESSION STATE --- 
if 'login' not in st.session_state: st.session_state.login = False 
if 'rol' not in st.session_state: st.session_state.rol = None 
if 'usuario' not in st.session_state: st.session_state.usuario = None

# --- LOGIN / INTERFAZ PRINCIPAL --- 
if not st.session_state.login: 
    col_img_login, col_title_login, col_spacer_login = st.columns([0.8, 3.8, 6.2]) 
    with col_img_login:
        IMAGE_PATH = "logge.png" 
        if os.path.exists(IMAGE_PATH): st.image(IMAGE_PATH, width=100) 
        else: st.markdown("&nbsp;") 
    with col_title_login:
        st.markdown("<h2 style='margin-top:0.5rem; margin-left: -0.5rem;'>📊 Estadístico Isertel</h2>", unsafe_allow_html=True) 

    st.subheader("Inicia sesión para acceder")
    col_login_spacer_l, col_login_box, col_login_spacer_r = st.columns([1, 2, 1])
    with col_login_box: 
        usuario_input = st.text_input("Usuario") 
        contrasena_input = st.text_input("Contraseña", type="password")
        if st.button("Iniciar sesión", use_container_width=True): 
            user_row = usuarios_df[
                (usuarios_df["Usuario"].str.lower() == usuario_input.strip().lower()) & 
                (usuarios_df["Contraseña"] == contrasena_input.strip()) 
            ] 
            if not user_row.empty: 
                st.session_state.login = True 
                st.session_state.rol = user_row.iloc[0]["Rol"] 
                st.session_state.usuario = usuario_input.strip() 
                st.rerun() 
            else: st.error("Usuario o contraseña incorrectos")
else: 
    col_img, col_title, col_spacer, col_welcome, col_logout = st.columns([0.8, 3.8, 3, 2, 1]) 
    with col_img:
        IMAGE_PATH = "logge.png" 
        if os.path.exists(IMAGE_PATH): st.image(IMAGE_PATH, width=100) 
        else: st.markdown("&nbsp;") 
    with col_title:
        st.markdown("<h2 style='margin-top:0.5rem; margin-left: -0.5rem;'>📊 Estadístico Isertel</h2>", unsafe_allow_html=True) 
    with col_welcome: 
        st.success(f"Bienvenido {st.session_state.usuario} ({st.session_state.rol})", icon=None) 
    with col_logout: 
        st.button("Cerrar sesión", on_click=lambda: st.session_state.update({"login": False, "rol": None, "usuario": None}), key="logout_btn", use_container_width=True)

    # --- CARGA Y COMBINACIÓN DE DATOS --- 
    archivos_para_combinar_nombres = [f for f in os.listdir(UPLOAD_FOLDER) if f.lower().endswith(('.xlsx', '.xls', '.csv'))] 
    num_archivos_cargados = len(archivos_para_combinar_nombres) 
    datos = None 
    df_list = []

    if archivos_para_combinar_nombres: 
        st.info(f"💾 **{num_archivos_cargados}** archivo(s) cargado(s) y combinado(s).") 
        archivos_completos = [os.path.join(UPLOAD_FOLDER, f) for f in archivos_para_combinar_nombres]
        try: 
            total_columnas_mapeadas = 0 
            for f in archivos_completos: 
                try: df = pd.read_csv(f, encoding='latin1') if f.lower().endswith('.csv') else pd.read_excel(f) 
                except UnicodeDecodeError: 
                    try: df = pd.read_csv(f, encoding='utf-8') 
                    except Exception as csv_err: st.warning(f"Error leyendo {f}: {csv_err}"); continue 
                except Exception as e: st.warning(f"Error leyendo {f}: {e}"); continue

                cleaned_names = [] 
                name_counts = {} 
                for name in df.columns: 
                    cleaned_name = str(name).upper().strip() 
                    name_counts[cleaned_name] = name_counts.get(cleaned_name, 0) + 1 
                    if name_counts[cleaned_name] > 1: cleaned_name = f"{cleaned_name}_{name_counts[cleaned_name]}" 
                    cleaned_names.append(cleaned_name) 
                df.columns = cleaned_names

                df_temp = pd.DataFrame() 
                columnas_encontradas_en_archivo = 0 
                for encabezado_excel, columna_final in MAPEO_COLUMNAS.items(): 
                    if encabezado_excel in df.columns: 
                        columna_data = df[encabezado_excel]
                        if isinstance(columna_data, pd.DataFrame): columna_data = columna_data.iloc[:, 0]
                        if encabezado_excel in df.columns: 
                            df_temp[columna_final] = columna_data 
                            columnas_encontradas_en_archivo += 1
                if not df_temp.empty: 
                    df_temp = df_temp.reindex(columns=COLUMNAS_SELECCIONADAS, fill_value=None) 
                    df_list.append(df_temp) 
                    total_columnas_mapeadas += columnas_encontradas_en_archivo
            if df_list: datos = pd.concat(df_list, ignore_index=True)
            if datos is None or datos.empty or total_columnas_mapeadas == 0: 
                st.warning("No se encontraron columnas mapeables."); datos = None
        except Exception as e: st.error(f"Error al combinar: {e}"); datos = None

    if datos is None: 
        try: 
            datos = pd.read_excel(MASTER_EXCEL) 
            columnas_existentes = [col for col in COLUMNAS_SELECCIONADAS if col in datos.columns] 
            datos = datos[columnas_existentes] 
        except: 
# 💥 DATOS DE PRUEBA 💥
            data = { 
                'FECHA': pd.to_datetime([f'2025-10-{d:02d}' for d in range(1, 11)] * 10),
                'UBICACION': ['Bogotá, 123', 'Bogotá, 456', 'Cali, 123', 'Cali, 456', 'Bogotá, 789', 'Medellín, 123', 'Medellín, 456', 'Medellín, 789', 'Cali, 789', 'Bogotá, 123'] * 10,
                'TECNICO': ['T|Juan Pérez (tecnico)', 'T|Juan Pérez (tecnico)', 'T|Pedro López (tecnico)', 'T|Pedro López', 'T|Ana Gómez (tecnico)', 'T|Ana Gómez', 'T|Juan Pérez (tecnico)', 'T|Juan Pérez', 'T|Pedro López (tecnico)', 'T|Ana Gómez (tecnico)'] * 10,
                'CONTRATO': ['C1']*100,
                'CUENTA': [f'CTA-{i:03d}' for i in range(100)],
                'MOTIVO': ['Falla línea', 'Lentitud', 'Sin tono', 'Instalación', 'Traslado', 'Mantenimiento', 'Corte', 'Revisión', 'Configuración', 'Otro'] * 10,
                'CLIENTE': ['Cliente A']*100,
                'TECNOLOGIA_COL': ['ADSL', 'ADSL', 'HFC', 'HFC', 'GPON', 'GPON', 'ADSL', 'HFC', 'GPON', 'ADSL'] * 10,
                'ID_TAREA': [101, 102, 103, 104, 105, 106, 107, 108, 109, 110] * 10,
                'ESTADO': ['SATISFACTORIA', 'Pendiente', 'INSATISFACTORIA', 'SATISFACTORIA', 'Pendiente', 'INSATISFACTORIA', 'SATISFACTORIA', 'Pendiente', 'INSATISFACTORIA', 'SATISFACTORIA'] * 10,
                'SOLUCION': ['Reinicio', 'Cambio conector', 'Actualización perfil', 'Cableado', 'Reconfiguración', 'Sin acción', 'Reparación', 'Cambio módem', 'Pruebas OK', 'Sustitución'] * 10,
                'TIPO_ORDEN': ['INSTALACION', 'VISITA TECNICA', 'MIGRACIÓN', 'TAREA MANUAL', 'CAMBIO DE DIRECCIÓN', 'MIGRATEC', 'INSTALACION', 'VISITA TECNICA', 'MIGRACIÓN', 'TAREA MANUAL'] * 10,
                'TIPO_TAREA_MANUAL': ['N/A', 'N/A', 'N/A', 'Auditoría', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'Retorno'] * 10
            } 
            datos = pd.DataFrame(data) 
            RENAME_DUMMY = {
                'FECHA': 'A', 
                'UBICACION': 'B', 
                'TECNICO': 'C', 
                'CONTRATO': 'D', 
                'CUENTA': 'E', 
                'MOTIVO': 'F', 
                'CLIENTE': 'G', 
                'TECNOLOGIA_COL': 'H', 
                'ID_TAREA': 'I', 
                'ESTADO': 'J', 
                'SOLUCION': 'K', 
                'TIPO_ORDEN': 'L', 
                'TIPO_TAREA_MANUAL': 'M'
            }
            datos = datos.rename(columns=RENAME_DUMMY)
            datos.columns = COLUMNAS_SELECCIONADAS
    if not archivos_para_combinar_nombres: st.warning("Usando **Datos de Prueba**.")

    # --- TABS --- 
    tabs = ["📊 Dashboard", "⚙️ Administración de Datos"] 
    if st.session_state.rol.lower() == "admin": tab_dashboard, tab_admin = st.tabs(tabs) 
    else: tab_dashboard = st.tabs(["📊 Dashboard"])[0]; tab_admin = None

    if st.session_state.rol.lower() == "admin" and tab_admin: 
        with tab_admin: 
            st.header("⚙️ Administración de Archivos Fuente") 
            st.metric(label="Documentos Cargados", value=f"{num_archivos_cargados} archivos") 
            st.markdown("---")
            col_upload, col_delete = st.columns(2)
            with col_upload: 
                st.subheader("Subir y Añadir") 
                nuevos_archivos = st.file_uploader("Subir archivos", type=["xlsx", "xls", "csv"], accept_multiple_files=True) 
                if st.button("📤 Guardar archivos"): 
                    if nuevos_archivos: 
                        for f in nuevos_archivos: 
                            with open(os.path.join(UPLOAD_FOLDER, f.name), "wb") as file: file.write(f.getbuffer()) 
                            st.success(f"Archivo '{f.name}' guardado.") 
                    st.info("Recargando..."); st.rerun()
            with col_delete: 
                st.subheader("Eliminar") 
                archivos_actuales = os.listdir(UPLOAD_FOLDER)
                eliminar = st.multiselect("Selecciona a eliminar", archivos_actuales) 
                if st.button("🗑️ Eliminar seleccionados"): 
                    if eliminar: 
                        for f in eliminar: os.remove(os.path.join(UPLOAD_FOLDER, f)) 
                        st.success("Eliminados. Recargando..."); st.rerun()
                if archivos_actuales and st.button("🔴 Eliminar TODOS", type="primary"): 
                    for f in archivos_actuales: os.remove(os.path.join(UPLOAD_FOLDER, f)) 
                    if os.path.exists(MASTER_EXCEL): os.remove(MASTER_EXCEL) 
                    st.success("Todos eliminados."); st.rerun()
            st.markdown("---")

# ---------------------------------------------------------------------- 
    # --- PESTAÑA DEL DASHBOARD --- 
    # ---------------------------------------------------------------------- 
    with tab_dashboard: 
        if datos is None or datos.empty: 
            st.warning("No hay datos para mostrar.") 
        else:
            # 1. PREPARACIÓN INICIAL DE DATOS (LIMPIEZA GENERAL)
            datos_base_limpia = datos.copy() 
            datos_base_limpia[COL_TEMP_DATETIME] = pd.to_datetime(datos_base_limpia[COL_FECHA_KEY], errors='coerce') 
            datos_base_limpia.dropna(subset=[COL_TEMP_DATETIME], inplace=True)
            
            # Pre-procesamiento de columnas filtro (Se hace UNA VEZ sobre la base)
            if COL_TECNICO_KEY in datos_base_limpia.columns: 
                datos_base_limpia[COL_FILTRO_TECNICO] = datos_base_limpia[COL_TECNICO_KEY].astype(str).apply(clean_tecnico) 
            if COL_CIUDAD_KEY in datos_base_limpia.columns: 
                datos_base_limpia[COL_FILTRO_CIUDAD] = datos_base_limpia[COL_CIUDAD_KEY].astype(str).apply(clean_ciudad)
            if COL_ESTADO_KEY in datos_base_limpia.columns:
                datos_base_limpia[COL_FILTRO_ESTADO] = datos_base_limpia[COL_ESTADO_KEY].astype(str).str.upper().str.strip()
                datos_base_limpia[COL_FILTRO_ESTADO].fillna("SIN ESTADO", inplace=True) 
            if COL_TIPO_ORDEN_KEY in datos_base_limpia.columns:
                datos_base_limpia[COL_FILTRO_TIPO_ORDEN] = datos_base_limpia[COL_TIPO_ORDEN_KEY].astype(str).str.upper().str.strip()
                datos_base_limpia[COL_FILTRO_TIPO_ORDEN].fillna("SIN TIPO", inplace=True) 
            if COL_TECNOLOGIA_KEY in datos_base_limpia.columns:
                datos_base_limpia[COL_FILTRO_TECNOLOGIA] = datos_base_limpia[COL_TECNOLOGIA_KEY].astype(str).str.upper().str.strip()
                datos_base_limpia[COL_FILTRO_TECNOLOGIA].fillna("SIN TECNOLOGIA", inplace=True) 
            if COL_TIPO_MANUAL_KEY in datos_base_limpia.columns:
                datos_base_limpia[COL_FILTRO_TIPO_MANUAL] = datos_base_limpia[COL_TIPO_MANUAL_KEY].astype(str).str.upper().str.strip()
                datos_base_limpia[COL_FILTRO_TIPO_MANUAL] = datos_base_limpia[COL_FILTRO_TIPO_MANUAL].replace(['NAN', 'NONE'], 'SIN TIPO MANUAL')

            if datos_base_limpia.empty: 
                st.warning("No hay registros con fechas válidas para mostrar.") 
            else:
                @st.cache_data 
                def get_multiselect_options(df, col_key_filtro): 
                    if col_key_filtro not in df.columns: return [] 
                    opciones = sorted([v for v in df[col_key_filtro].astype(str).unique() if pd.notna(v) and str(v).strip() not in ('nan', 'none', '')]) 
                    return opciones

                @st.cache_data 
                def apply_filter(df, col_key_filtro, selected_options): 
                    if not selected_options: return df
                    if col_key_filtro not in df.columns: return df 
                    return df[df[col_key_filtro].astype(str).isin(selected_options)]
                    
                # -----------------------------------------------------------------------------
                # --- PANEL DE CONTROL: FILTROS (Lógica de Filtro Cruzado / Cross-Filtering) --- 
                # -----------------------------------------------------------------------------
                with st.container(border=True):
                    # --- HEADER FILTROS Y BOTÓN LIMPIAR ---
                    col_header_filtros, col_btn_limpiar = st.columns([6, 1])
                    with col_header_filtros:
                        st.markdown("#### ⚙️ Filtros de Segmentación") 
                    with col_btn_limpiar:
                         if st.button("🧹 Limpiar Filtros", use_container_width=True):
                            keys_to_clear = ['multiselect_ubicacion', 'filter_tecnico', 'multiselect_estado', 
                                             'multiselect_tipo_orden', 'multiselect_tecnologia', 'multiselect_tipo_manual']
                            for k in keys_to_clear:
                                if k in st.session_state:
                                    st.session_state[k] = []
                            st.rerun()

                    col_desde, col_hasta, col_ciu, col_tec, col_est, col_tipo_orden, col_tecnologia, col_tipo_manual = st.columns(
                        [1.0, 1.0, 1.3, 1.3, 1.3, 1.3, 1.3, 1.3] 
                    )

                    with col_desde: 
                        min_date_global = datos_base_limpia[COL_TEMP_DATETIME].min().replace(hour=0, minute=0, second=0, microsecond=0) 
                        max_date_global = datos_base_limpia[COL_TEMP_DATETIME].max().replace(hour=0, minute=0, second=0, microsecond=0) 
                        date_from = st.date_input("Desde:", value=min_date_global, min_value=min_date_global, max_value=max_date_global, key='filter_date_from')
                    
                    with col_hasta: 
                        date_to = st.date_input("Hasta:", value=max_date_global, min_value=min_date_global, max_value=max_date_global, key='filter_date_to')
                    
                    if date_from > date_to: 
                        st.error("⚠️ Fecha 'Desde' mayor que 'Hasta'."); st.stop()
                    
                    # Filtro inicial de fecha para calcular el dominio de opciones
                    filtro_inicio = pd.to_datetime(date_from) 
                    filtro_fin = pd.to_datetime(date_to) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)
                    df_base_fecha = datos_base_limpia[
                        (datos_base_limpia[COL_TEMP_DATETIME] >= filtro_inicio) & 
                        (datos_base_limpia[COL_TEMP_DATETIME] <= filtro_fin)
                    ].copy()

                    # --- LÓGICA DE FILTRO CRUZADO ---
                    # Para que los filtros se influyan entre sí, las opciones de cada filtro se calculan 
                    # aplicando todos los filtros EXCEPTO el filtro que se está calculando.

                    # Capturamos selecciones actuales de session_state (si existen)
                    s_ciu = st.session_state.get('multiselect_ubicacion', [])
                    s_tec = st.session_state.get('filter_tecnico', [])
                    s_est = st.session_state.get('multiselect_estado', [])
                    s_tip = st.session_state.get('multiselect_tipo_orden', [])
                    s_tcn = st.session_state.get('multiselect_tecnologia', [])
                    s_man = st.session_state.get('multiselect_tipo_manual', [])

                    # 1. OPCIONES CIUDAD (Filtradas por: Técnico, Estado, Tipo Orden, Tecnología, Manual)
                    df_c = apply_filter(df_base_fecha, COL_FILTRO_TECNICO, s_tec)
                    df_c = apply_filter(df_c, COL_FILTRO_ESTADO, s_est)
                    df_c = apply_filter(df_c, COL_FILTRO_TIPO_ORDEN, s_tip)
                    df_c = apply_filter(df_c, COL_FILTRO_TECNOLOGIA, s_tcn)
                    df_c = apply_filter(df_c, COL_FILTRO_TIPO_MANUAL, s_man)
                    opciones_ciudad = get_multiselect_options(df_c, COL_FILTRO_CIUDAD)

                    # 2. OPCIONES TÉCNICO (Filtradas por: Ciudad, Estado, Tipo Orden, Tecnología, Manual)
                    df_t = apply_filter(df_base_fecha, COL_FILTRO_CIUDAD, s_ciu)
                    df_t = apply_filter(df_t, COL_FILTRO_ESTADO, s_est)
                    df_t = apply_filter(df_t, COL_FILTRO_TIPO_ORDEN, s_tip)
                    df_t = apply_filter(df_t, COL_FILTRO_TECNOLOGIA, s_tcn)
                    df_t = apply_filter(df_t, COL_FILTRO_TIPO_MANUAL, s_man)
                    opciones_tecnico = get_multiselect_options(df_t, COL_FILTRO_TECNICO)

                    # 3. OPCIONES ESTADO (Filtradas por: Ciudad, Técnico, Tipo Orden, Tecnología, Manual)
                    df_e = apply_filter(df_base_fecha, COL_FILTRO_CIUDAD, s_ciu)
                    df_e = apply_filter(df_e, COL_FILTRO_TECNICO, s_tec)
                    df_e = apply_filter(df_e, COL_FILTRO_TIPO_ORDEN, s_tip)
                    df_e = apply_filter(df_e, COL_FILTRO_TECNOLOGIA, s_tcn)
                    df_e = apply_filter(df_e, COL_FILTRO_TIPO_MANUAL, s_man)
                    opciones_estado = get_multiselect_options(df_e, COL_FILTRO_ESTADO)

                    # 4. OPCIONES TIPO ORDEN (Filtradas por: Ciudad, Técnico, Estado, Tecnología, Manual)
                    df_o = apply_filter(df_base_fecha, COL_FILTRO_CIUDAD, s_ciu)
                    df_o = apply_filter(df_o, COL_FILTRO_TECNICO, s_tec)
                    df_o = apply_filter(df_o, COL_FILTRO_ESTADO, s_est)
                    df_o = apply_filter(df_o, COL_FILTRO_TECNOLOGIA, s_tcn)
                    df_o = apply_filter(df_o, COL_FILTRO_TIPO_MANUAL, s_man)
                    opciones_tipo_orden = get_multiselect_options(df_o, COL_FILTRO_TIPO_ORDEN)

                    # 5. OPCIONES TECNOLOGÍA (Filtradas por: Ciudad, Técnico, Estado, Tipo Orden, Manual)
                    df_te = apply_filter(df_base_fecha, COL_FILTRO_CIUDAD, s_ciu)
                    df_te = apply_filter(df_te, COL_FILTRO_TECNICO, s_tec)
                    df_te = apply_filter(df_te, COL_FILTRO_ESTADO, s_est)
                    df_te = apply_filter(df_te, COL_FILTRO_TIPO_ORDEN, s_tip)
                    df_te = apply_filter(df_te, COL_FILTRO_TIPO_MANUAL, s_man)
                    opciones_tecnologia = get_multiselect_options(df_te, COL_FILTRO_TECNOLOGIA)

                    # 6. OPCIONES TIPO MANUAL (Filtradas por: Ciudad, Técnico, Estado, Tipo Orden, Tecnología)
                    df_m = apply_filter(df_base_fecha, COL_FILTRO_CIUDAD, s_ciu)
                    df_m = apply_filter(df_m, COL_FILTRO_TECNICO, s_tec)
                    df_m = apply_filter(df_m, COL_FILTRO_ESTADO, s_est)
                    df_m = apply_filter(df_m, COL_FILTRO_TIPO_ORDEN, s_tip)
                    df_m = apply_filter(df_m, COL_FILTRO_TECNOLOGIA, s_tcn)
                    opciones_tipo_manual = get_multiselect_options(df_m, COL_FILTRO_TIPO_MANUAL)

                    # --- WIDGETS ---
                    # NOTA: Para preservar los filtros al cambiar las fechas, primero calculamos la intersección
                    # entre lo seleccionado anteriormente y las nuevas opciones disponibles, y luego
                    # FORZAMOS esa selección en st.session_state antes de renderizar el widget.
                    
                    with col_ciu:
                        valid_ciu = [v for v in s_ciu if v in opciones_ciudad]
                        if 'multiselect_ubicacion' not in st.session_state: st.session_state['multiselect_ubicacion'] = []
                        st.session_state['multiselect_ubicacion'] = valid_ciu
                        filtro_ciudad = st.multiselect(f"**{COL_CIUDAD_DESCRIPTIVA}**:", options=opciones_ciudad, key='multiselect_ubicacion', placeholder="Ciudad")
                    
                    with col_tec:
                        # Para técnicos, la función custom maneja sus propias opciones, pero nos aseguramos de limpiar selección inválida
                        valid_tec = [v for v in s_tec if v in opciones_tecnico or v in ["✨ Seleccionar Todos", "👷 Seleccionar Supervisores"]]
                        if 'filter_tecnico' not in st.session_state: st.session_state['filter_tecnico'] = []
                        st.session_state['filter_tecnico'] = valid_tec
                        filtro_tecnico = st_multiselect_with_all_technicians(col_tec, f"**{COL_TECNICO_DESCRIPTIVA}**", options=opciones_tecnico, key='filter_tecnico')
                    
                    with col_est:
                        valid_est = [v for v in s_est if v in opciones_estado]
                        if 'multiselect_estado' not in st.session_state: st.session_state['multiselect_estado'] = []
                        st.session_state['multiselect_estado'] = valid_est
                        filtro_estado = st.multiselect(f"**{COL_ESTADO_DESCRIPTIVA}**:", options=opciones_estado, key='multiselect_estado', placeholder="Estado")
                    
                    with col_tipo_orden:
                        valid_tip = [v for v in s_tip if v in opciones_tipo_orden]
                        if 'multiselect_tipo_orden' not in st.session_state: st.session_state['multiselect_tipo_orden'] = []
                        st.session_state['multiselect_tipo_orden'] = valid_tip
                        filtro_tipo_orden = st.multiselect(f"**{COL_TIPO_ORDEN_DESCRIPTIVA}**:", options=opciones_tipo_orden, key='multiselect_tipo_orden', placeholder="Tipo Orden")
                    
                    with col_tecnologia:
                        valid_tcn = [v for v in s_tcn if v in opciones_tecnologia]
                        if 'multiselect_tecnologia' not in st.session_state: st.session_state['multiselect_tecnologia'] = []
                        st.session_state['multiselect_tecnologia'] = valid_tcn
                        filtro_tecnologia = st.multiselect(f"**{COL_TECNOLOGIA_DESCRIPTIVA}**:", options=opciones_tecnologia, key='multiselect_tecnologia', placeholder="Tecnología")
                    
                    with col_tipo_manual:
                        if 'TAREA MANUAL' in filtro_tipo_orden:
                            valid_man = [v for v in s_man if v in opciones_tipo_manual]
                            if 'multiselect_tipo_manual' not in st.session_state: st.session_state['multiselect_tipo_manual'] = []
                            st.session_state['multiselect_tipo_manual'] = valid_man
                            filtro_tipo_manual = st.multiselect(f"**{COL_TIPO_MANUAL_DESCRIPTIVA}**:", options=opciones_tipo_manual, key='multiselect_tipo_manual', placeholder="Sub-tipo Manual")
                        else:
                            filtro_tipo_manual = [] 
                            st.markdown(f"<p style='margin-top:2.2rem; font-size: 0.9rem; color: #a0a0a0;'>{COL_TIPO_MANUAL_DESCRIPTIVA}</p>", unsafe_allow_html=True)

                    # --- APLICACIÓN FINAL DE FILTROS A LOS DATOS ---
                    df_final = df_base_fecha.copy()
                    df_final = apply_filter(df_final, COL_FILTRO_CIUDAD, filtro_ciudad) 
                    df_final = apply_filter(df_final, COL_FILTRO_TECNICO, filtro_tecnico) 
                    df_final = apply_filter(df_final, COL_FILTRO_ESTADO, filtro_estado) 
                    df_final = apply_filter(df_final, COL_FILTRO_TIPO_ORDEN, filtro_tipo_orden) 
                    df_final = apply_filter(df_final, COL_FILTRO_TECNOLOGIA, filtro_tecnologia) 
                    if filtro_tipo_manual: 
                        df_final = apply_filter(df_final, COL_FILTRO_TIPO_MANUAL, filtro_tipo_manual) 
                    
                    datos_filtrados = df_final 

                # -----------------------------------------------------------------------------
                # --- MÉTRICAS --- 
                # -----------------------------------------------------------------------------
                with st.container(border=True):
                    st.markdown("#### 🎯 Métricas Clave (KPIs)") 
                    col_m_sat_abs, col_m_inst_abs, col_m_vis_abs, col_m_mig_abs, col_m_man_abs, col_m_cd_abs, col_m_migratec_abs = st.columns([1,1,1,1,1,1,1])

                    if len(filtro_estado) == 1:
                        estado_base = filtro_estado[0]
                        datos_base_metricas = datos_filtrados[datos_filtrados[COL_FILTRO_ESTADO] == estado_base].copy()
                        etiqueta_estado = f" ({estado_base.title().replace(' ','')[:3]}.)"
                        etiqueta_total_base = f"Total ({estado_base.title().replace(' ','')[:3]}.)"
                    else:
                        estado_tarea = datos_filtrados[COL_ESTADO_KEY].astype(str)
                        es_satisfactoria = estado_tarea.str.contains('SATISFACTORIA', case=False, na=False)
                        es_insatisfactoria = estado_tarea.str.contains('INSATISFACTORIA', case=False, na=False)
                        datos_base_metricas = datos_filtrados[es_satisfactoria & ~es_insatisfactoria].copy()
                        estado_base = "SATISFACTORIA"
                        etiqueta_estado = " (Sat.)"; etiqueta_total_base = "Total Sat."

                    total_base = len(datos_base_metricas)
                    
                    if COL_TIPO_ORDEN_KEY in datos_base_metricas.columns: 
                        tipo_orden_base = datos_base_metricas[COL_TIPO_ORDEN_KEY].astype(str)
                        total_instalaciones = len(datos_base_metricas[tipo_orden_base.str.contains('INSTALACION', case=False, na=False)]) 
                        total_visitas_tecnicas = len(datos_base_metricas[tipo_orden_base.str.contains('VISITA TECNICA', case=False, na=False)])
                        mask_migracion_orden = tipo_orden_base.str.contains(r'MIGRACI[ÓO]N', case=False, na=False, regex=True)
                        mask_migracion_manual = False
                        if COL_TIPO_MANUAL_KEY in datos_base_metricas.columns:
                            mask_migracion_manual = datos_base_metricas[COL_TIPO_MANUAL_KEY].astype(str).str.contains(r'MIGRACI[ÓO]N', case=False, na=False, regex=True)
                        total_migracion = len(datos_base_metricas[mask_migracion_orden | mask_migracion_manual])
                        total_tarea_manual = len(datos_base_metricas[tipo_orden_base.str.contains('TAREA MANUAL', case=False, na=False)])
                        total_cambio_direccion = len(datos_base_metricas[tipo_orden_base.str.contains(r'CAMBIO DE DIRECCI[ÓO]N', case=False, na=False, regex=True)])
                        total_migratec = len(datos_base_metricas[tipo_orden_base.str.contains('MIGRATEC', case=False, na=False)])
                    else: 
                        total_instalaciones = 0; total_visitas_tecnicas = 0; total_migracion = 0 
                        total_tarea_manual = 0; total_cambio_direccion = 0; total_migratec = 0
                    
                    def metric_card(col, label, val, is_total=False):
                        css_class = "metric-compact-container-total" if is_total else "metric-compact-container"
                        with col:
                            st.markdown(f'<div class="{css_class}">', unsafe_allow_html=True) 
                            st.metric(label=label, value=f"{val:,}") 
                            st.markdown('</div>', unsafe_allow_html=True)

                    metric_card(col_m_sat_abs, etiqueta_total_base, total_base, True)
                    metric_card(col_m_inst_abs, f"Instal.{etiqueta_estado}", total_instalaciones)
                    metric_card(col_m_vis_abs, f"Visitas{etiqueta_estado}", total_visitas_tecnicas)
                    metric_card(col_m_mig_abs, f"Migra.{etiqueta_estado}", total_migracion)
                    metric_card(col_m_man_abs, f"Manual{etiqueta_estado}", total_tarea_manual)
                    metric_card(col_m_cd_abs, f"Cam.Dir.{etiqueta_estado}", total_cambio_direccion)
                    metric_card(col_m_migratec_abs, f"MigraTec{etiqueta_estado}", total_migratec)
                        
                st.markdown("---")
                datos_filtrados = datos_base_metricas.copy() 
                
                # ------------------------------------------------------------------------------------- 
                # --- LAYOUT PRINCIPAL --- 
                # -------------------------------------------------------------------------------------
                col_raw, col_graphs_group = st.columns([5, 15]) 

                # --- COLUMNA 1: TABLA RAW --- 
                with col_raw:
                    st.markdown(f"#### 📑 Datos ({len(datos_filtrados)})")
                    datos_filtrados_ordenados = datos_filtrados.sort_values(by=COL_TEMP_DATETIME, ascending=True).copy()
                    datos_vista = datos_filtrados_ordenados.rename(columns=FINAL_RENAMING_MAP) 
                    columnas_finales = [col for col in FINAL_RENAMING_MAP.values() if col in datos_vista.columns] 
                    if COL_FILTRO_TECNICO in datos_filtrados_ordenados.columns and FINAL_RENAMING_MAP['C'] in datos_vista.columns:
                         datos_vista[FINAL_RENAMING_MAP['C']] = datos_filtrados_ordenados[COL_FILTRO_TECNICO]
                    datos_vista = datos_vista[columnas_finales]

                    all_cols = datos_vista.columns.tolist() 
                    default_cols_raw = [FINAL_RENAMING_MAP['A'], FINAL_RENAMING_MAP['B'], FINAL_RENAMING_MAP['C'], FINAL_RENAMING_MAP['G']]
                    default_cols = [c for c in default_cols_raw if c in all_cols]

                    cols_to_show = st.multiselect("**Columnas**:", options=all_cols, default=default_cols, key='raw_table_col_select')
                    df_to_display = datos_vista[cols_to_show] if cols_to_show else datos_vista

                    st.markdown('<div style="overflow-x: auto;">', unsafe_allow_html=True) 
                    st.data_editor(df_to_display, use_container_width=True, hide_index=True, key='editable_raw', num_rows="fixed") 
                    st.markdown('</div>', unsafe_allow_html=True)

                    # Export
                    excel_buffer = io.BytesIO()
                    datos_filtrados.rename(columns=FINAL_RENAMING_MAP).to_excel(excel_buffer, index=False)
                    excel_buffer.seek(0)
                    st.download_button(label="⬇️ Excel Filtrado", data=excel_buffer, file_name='data.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)

                # --- COLUMNA 2: GRÁFICOS --- 
                with col_graphs_group: 
                    col_graphs_izq, col_graphs_der = st.columns([8, 7])
                    COL_AGRUPACION_KEY = COL_TECNOLOGIA_KEY 
                    COL_AGRUPACION_DESCRIPTIVA = COL_TECNOLOGIA_DESCRIPTIVA 

                    with col_graphs_izq:
                        with st.container(border=True):
                            st.markdown(f"#### Por Tecnología (Base: {estado_base.title()})") 
                            if len(datos_filtrados) > 0 and COL_AGRUPACION_KEY in datos_filtrados.columns:
                                conteo_tecnologia = datos_filtrados[COL_AGRUPACION_KEY].value_counts().reset_index()
                                conteo_tecnologia.columns = [COL_AGRUPACION_DESCRIPTIVA, 'Total_Tareas']
                                fig = px.bar(conteo_tecnologia, x=COL_AGRUPACION_DESCRIPTIVA, y='Total_Tareas', text='Total_Tareas', color=COL_AGRUPACION_DESCRIPTIVA, color_discrete_sequence=['#4CAF50', '#2196F3', '#FF9800'])
                                fig.update_layout(xaxis_title=None, yaxis_title=None, margin=dict(t=20, b=10, l=10, r=10), height=200)
                                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
                            else: st.info("Sin datos.")

                    with col_graphs_der: 
                        with st.container(border=True): 
                            is_single_city = len(filtro_ciudad) == 1
                            is_single_tech = len(filtro_tecnico) == 1
                            
                            if is_single_tech:
                                st.markdown(f"#### Tareas de **{filtro_tecnico[0]}**")
                                group_col = COL_FILTRO_CIUDAD
                            elif is_single_city:
                                st.markdown(f"#### Top Técnicos en **{filtro_ciudad[0]}**")
                                group_col = COL_FILTRO_TECNICO
                            else:
                                st.markdown(f"#### Distribución Ubicación") 
                                group_col = COL_FILTRO_CIUDAD

                            if group_col in datos_filtrados.columns and len(datos_filtrados) > 0: 
                                conteo = datos_filtrados[group_col].value_counts().reset_index() 
                                conteo.columns = ['Label', 'Total']
                                if is_single_city: conteo = conteo.head(5)
                                fig_pie = px.pie(conteo, values='Total', names='Label', hole=.4, color_discrete_sequence=px.colors.qualitative.Pastel) 
                                fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                                fig_pie.update_layout(showlegend=True, margin=dict(l=0, r=0, t=20, b=0), height=200)
                                st.plotly_chart(fig_pie, use_container_width=True)
                            else: st.info("Sin datos.")

                    # *** RENDIMIENTO DINÁMICO ***
                    st.markdown("---") 
                    
                    # Detectar si hay filtros específicos que requieran una vista de detalle único
                    filtros_especificos_activos = (len(filtro_tipo_orden) > 0) or (len(filtro_tecnologia) > 0) or (len(filtro_tipo_manual) > 0)

                    if filtros_especificos_activos:
                        # --- MODO DETALLE: GRÁFICO ÚNICO ---
                        # Construir título dinámico
                        partes_titulo = []
                        if filtro_tipo_orden: partes_titulo.append(f"Orden: {', '.join(filtro_tipo_orden)}")
                        if filtro_tecnologia: partes_titulo.append(f"Tech: {', '.join(filtro_tecnologia)}")
                        if filtro_tipo_manual: partes_titulo.append(f"Manual: {', '.join(filtro_tipo_manual)}")
                        titulo_grafico = " + ".join(partes_titulo) if partes_titulo else "Filtros Seleccionados"
                        
                        st.markdown(f"### 📈 Rendimiento Filtrado: **{titulo_grafico}**")
                        
                        with st.container(border=True):
                            if datos_filtrados.empty:
                                st.info("No hay datos para los filtros seleccionados.")
                            else:
                                # Determinar eje X (Agrupación)
                                if len(filtro_tecnico) == 1:
                                    # 1 Técnico -> Ver evolución por FECHA
                                    group_col = '_FECHA_DIA_'
                                    label_x = "Fecha"
                                    # Asegurar columna fecha día en datos filtrados para agrupar
                                    datos_filtrados['_FECHA_DIA_'] = datos_filtrados[COL_TEMP_DATETIME].dt.date
                                    es_temporal = True
                                elif len(filtro_tecnico) > 1:
                                    # Varios Técnicos (o Todos) -> Comparar TÉCNICOS
                                    group_col = COL_FILTRO_TECNICO
                                    label_x = "Técnico"
                                    es_temporal = False
                                else:
                                    # Ningún técnico seleccionado (por defecto todos, pero suele caer en Ciudad)
                                    # Si hay varias ciudades, agrupar por ciudad, sino por técnico
                                    if len(filtro_ciudad) > 1:
                                        group_col = COL_FILTRO_CIUDAD
                                        label_x = "Ubicación"
                                    else:
                                        group_col = COL_FILTRO_TECNICO
                                        label_x = "Técnico"
                                    es_temporal = False

                                # Agrupar y Contar
                                df_unico = datos_filtrados.groupby(group_col).size().reset_index(name='Total_Tareas')
                                
                                # Ordenar
                                if not es_temporal:
                                    df_unico = df_unico.sort_values(by='Total_Tareas', ascending=False)
                                else:
                                    df_unico = df_unico.sort_values(by=group_col, ascending=True)

                                # Renderizar Gráfico Único
                                height_u = 300
                                color_u = '#2196F3' # Azul estándar para el gráfico único
                                
                                if es_temporal:
                                    fig_u = px.line(df_unico, x=group_col, y='Total_Tareas', markers=True, text='Total_Tareas', height=height_u, color_discrete_sequence=[color_u])
                                else:
                                    # Usar Barras para comparación entre técnicos/ciudades (es más claro para volúmenes)
                                    fig_u = px.bar(df_unico, x=group_col, y='Total_Tareas', text='Total_Tareas', height=height_u, color_discrete_sequence=[color_u])
                                
                                fig_u.update_traces(textposition='outside' if not es_temporal else 'top center')
                                fig_u.update_layout(
                                    title=f"Total de trabajos ({titulo_grafico}) por {label_x}",
                                    xaxis_title=None,
                                    yaxis_title='Total',
                                    margin=dict(t=40, b=60, l=20, r=20),
                                    xaxis={'tickangle': -45}
                                )
                                fig_u.update_yaxes(showgrid=True, gridcolor='#eeeeee')
                                st.plotly_chart(fig_u, use_container_width=True)

                    else:
                        # --- MODO ESTÁNDAR: 6 GRÁFICOS (Instalación, Visita, etc.) ---
                        st.markdown(f"### 📈 Rendimiento Detallado (Base: {estado_base.title()})")
                        with st.container(border=True): 
                            if len(filtro_tecnico) == 1:
                                df_comparacion_view = prepare_date_comparison_data(datos_filtrados) 
                                x_col, title, is_city_view = '_FECHA_DIA_', f"por Día: **{filtro_tecnico[0]}**", False
                            elif len(filtro_tecnico) > 1:
                                df_comparacion_view = prepare_technician_comparison_data(datos_filtrados) 
                                x_col, title, is_city_view = COL_FILTRO_TECNICO, "por Técnico", False 
                            else:
                                df_comparacion_view = prepare_city_comparison_data(datos_filtrados) 
                                x_col, title, is_city_view = COL_FILTRO_CIUDAD, "por Ubicación", True
                            
                            if not df_comparacion_view.empty: 
                                render_comparison_charts_vertical(df_comparacion_view, x_col, title, is_city_view) 
                            else: st.info("No hay datos de rendimiento.")

                    # =====================================================================
                    # --- NUEVA FUNCIÓN: GRÁFICO "Rango Fecha-Tipo de Orden" ---
                    # =====================================================================
                    
                    # 1. Comprobar si el usuario eligió un rango personalizado (distinto al min/max predeterminado)
                    rango_modificado = (date_from != min_date_global.date()) or (date_to != max_date_global.date())
                    # 2. Comprobar si eligió al menos un "Tipo de Orden"
                    tiene_tipo_orden = len(filtro_tipo_orden) > 0

                    if rango_modificado and tiene_tipo_orden:
                        st.markdown("### 📈 Rango Fecha-Tipo de Orden")
                        with st.container(border=True):
                            df_rango = datos_filtrados.copy()
                            if not df_rango.empty:
                                # Asegurar la columna de fecha por día
                                df_rango['_FECHA_DIA_'] = df_rango[COL_TEMP_DATETIME].dt.date
                                # Agrupar cantidades por día
                                conteo_rango = df_rango.groupby('_FECHA_DIA_').size().reset_index(name='Total')
                                conteo_rango = conteo_rango.sort_values(by='_FECHA_DIA_')
                                
                                # Renderizar gráfico de barras
                                fig_rango = px.bar(
                                    conteo_rango, 
                                    x='_FECHA_DIA_', 
                                    y='Total', 
                                    text='Total',
                                    color='_FECHA_DIA_', # Cada día tendrá un color distinto para fácil lectura
                                    color_discrete_sequence=px.colors.qualitative.Pastel
                                )
                                fig_rango.update_traces(textposition='outside')
                                fig_rango.update_layout(
                                    showlegend=False,
                                    xaxis_title="Días (Rango Seleccionado)",
                                    yaxis_title="Cantidad",
                                    margin=dict(t=20, b=40, l=20, r=20),
                                    xaxis={'tickangle': -45, 'type': 'category'} # Obliga a que sean barras categóricas día a día
                                )
                                st.plotly_chart(fig_rango, use_container_width=True)
                            else:
                                st.info("No hay datos para esta combinación.")