import streamlit as st 
import pandas as pd 
import os 
import plotly.express as px 
import numpy as np
from datetime import datetime, timedelta 
import io 

# --- FUNCIÓN DE COMPACIDAD Y CONFIGURACIÓN --- 
def set_page_config_and_style(): 
    st.set_page_config(layout="wide", page_title="Estadístico Isertel - Nuevo Formato")

    st.markdown(""" 
    <style> 
    .block-container { 
        padding-top: 3rem !important; 
        padding-bottom: 0rem !important; 
        padding-left: 1rem !important; 
        padding-right: 1rem !important; 
    }
    div[data-testid="stHorizontalBlock"] { gap: 0.75rem !important; }
    h3, h4, h5 { margin-top: 0.1rem !important; margin-bottom: 0.1rem !important; }
    hr { margin-top: 0.1rem !important; margin-bottom: 0.1rem !important; }
    .stSelectbox, .stMultiSelect, .stDateInput, div[data-testid="stForm"] { margin-bottom: 0.0rem !important; }
    div[data-testid="stMetric"] { padding: 0.2rem 0 !important; }
    div[data-testid="stMetricLabel"] { font-size: 0.95rem; }

    .metric-compact-container div[data-testid="stMetricValue"] { font-size: 1.6rem; color: #B71C1C; } 
    .metric-compact-container-total div[data-testid="stMetricValue"] { font-size: 1.6rem; color: #0D47A1; }
    div[data-testid="stMetricDelta"] { visibility: hidden; height: 0; }

    div[data-testid="stSuccess"] { 
        padding: 0.5rem 1rem !important; margin-bottom: 0px; display: flex; 
        justify-content: flex-end; align-items: center; height: 100%; 
    } 
    .stButton>button { height: 32px; padding-top: 4px !important; padding-bottom: 4px !important; }
    .stDataFrame { margin-top: 0.5rem; } 
    
    /* ESTILOS PARA TARJETAS VERSUS */
    .vs-card-item {
        background-color: #1E3A8A10;
        border: 2px solid #1E3A8A;
        border-radius: 10px;
        padding: 15px;
        text-align: center;
        margin-bottom: 10px;
    }
    </style> 
    """, unsafe_allow_html=True)

set_page_config_and_style()

# --- CONFIGURACIÓN DE ARCHIVOS Y CARPETAS --- 
MASTER_EXCEL = "datos.xlsx" 
USUARIOS_EXCEL = "usuarios.xlsx" 
UPLOAD_FOLDER = "ExcelUploads" 
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 1. MAPEO DE COLUMNAS (Sin Columna H)
MAPEO_COLUMNAS = { 
    'FECHA': 'B', 
    'CIUDAD': 'C', 
    'TÉCNICO': 'D', 
    'EMPRESA': 'E', 
    'MOTIVO': 'I', 
    'GPON': 'M', 
    'TIPO ORDEN PRESENTACIÓN': 'N', 
    'CONTROL': 'O'
}
COLUMNAS_SELECCIONADAS = list(MAPEO_COLUMNAS.values()) 
ENCABEZADOS_ESPERADOS = list(MAPEO_COLUMNAS.keys())
FINAL_RENAMING_MAP = {v: k for k, v in MAPEO_COLUMNAS.items()}

# CLAVES DE COLUMNA
COL_FECHA_KEY = 'B' 
COL_CIUDAD_KEY = 'C' 
COL_TECNICO_KEY = 'D' 
COL_EMPRESA_KEY = 'E'
COL_MOTIVO_KEY = 'I'
COL_GPON_KEY = 'M'
COL_TIPO_PRESENTACION_KEY = 'N'
COL_CONTROL_KEY = 'O'

COL_TEMP_DATETIME = '_DATETIME_' + COL_FECHA_KEY

# DESCRIPCIONES DE COLUMNAS
COL_FECHA_DESCRIPTIVA = FINAL_RENAMING_MAP[COL_FECHA_KEY] 
COL_CIUDAD_DESCRIPTIVA = FINAL_RENAMING_MAP[COL_CIUDAD_KEY] 
COL_TECNICO_DESCRIPTIVA = FINAL_RENAMING_MAP[COL_TECNICO_KEY] 
COL_EMPRESA_DESCRIPTIVA = FINAL_RENAMING_MAP[COL_EMPRESA_KEY]
COL_MOTIVO_DESCRIPTIVA = FINAL_RENAMING_MAP[COL_MOTIVO_KEY]
COL_GPON_DESCRIPTIVA = FINAL_RENAMING_MAP[COL_GPON_KEY]
COL_TIPO_PRESENTACION_DESCRIPTIVA = FINAL_RENAMING_MAP[COL_TIPO_PRESENTACION_KEY]
COL_CONTROL_DESCRIPTIVA = FINAL_RENAMING_MAP[COL_CONTROL_KEY]

# Columnas para Filtros
COL_FILTRO_TECNICO = '_Filtro_Tecnico_' 
COL_FILTRO_CIUDAD = '_Filtro_Ubicacion_'
COL_FILTRO_EMPRESA = '_Filtro_Empresa_'
COL_FILTRO_MOTIVO = '_Filtro_Motivo_'
COL_FILTRO_GPON = '_Filtro_Gpon_'
COL_FILTRO_PRESENTACION = '_Filtro_Presentacion_'
COL_FILTRO_CONTROL = '_Filtro_Control_'

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

@st.cache_data
def clean_text(val):
    if pd.isna(val): return "SIN REGISTRO"
    s = str(val).strip().upper()
    if s in ('', 'NAN', 'NONE', '-'): return "SIN REGISTRO"
    return s

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
        st.markdown("<h2 style='margin-top:0.5rem;'>📊 Estadístico Isertel</h2>", unsafe_allow_html=True) 

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
        st.markdown("<h2 style='margin-top:0.5rem;'>📊 Estadístico Isertel</h2>", unsafe_allow_html=True) 
    with col_welcome: 
        st.success(f"Bienvenido {st.session_state.usuario} ({st.session_state.rol})", icon=None) 
    with col_logout: 
        st.button("Cerrar sesión", on_click=lambda: st.session_state.update({"login": False, "rol": None, "usuario": None}), key="logout_btn", use_container_width=True)

    # --- CARGA Y COMBINACIÓN DE DATOS CON SOPORTE MULTI-DOCUMENTO --- 
    archivos_para_combinar_nombres = [f for f in os.listdir(UPLOAD_FOLDER) if f.lower().endswith(('.xlsx', '.xls', '.csv'))] 
    num_archivos_cargados = len(archivos_para_combinar_nombres) 
    datos = None 
    df_list = []
    file_dfs_dict = {}

    if archivos_para_combinar_nombres: 
        archivos_completos = [os.path.join(UPLOAD_FOLDER, f) for f in archivos_para_combinar_nombres]
        try: 
            total_columnas_mapeadas = 0 
            for f_path, f_name in zip(archivos_completos, archivos_para_combinar_nombres): 
                try: 
                    df_raw = pd.read_csv(f_path, encoding='latin1') if f_path.lower().endswith('.csv') else pd.read_excel(f_path) 
                except UnicodeDecodeError: 
                    try: df_raw = pd.read_csv(f_path, encoding='utf-8') 
                    except Exception as csv_err: st.warning(f"Error leyendo {f_name}: {csv_err}"); continue 
                except Exception as e: st.warning(f"Error leyendo {f_name}: {e}"); continue

                df_temp = pd.DataFrame()
                if df_raw.shape[1] >= 15:
                    df_temp[COL_FECHA_KEY] = df_raw.iloc[:, 1]
                    df_temp[COL_CIUDAD_KEY] = df_raw.iloc[:, 2]
                    df_temp[COL_TECNICO_KEY] = df_raw.iloc[:, 3]
                    df_temp[COL_EMPRESA_KEY] = df_raw.iloc[:, 4]
                    # Se omite índice 7 (Columna H)
                    df_temp[COL_MOTIVO_KEY] = df_raw.iloc[:, 8]
                    df_temp[COL_GPON_KEY] = df_raw.iloc[:, 12]
                    df_temp[COL_TIPO_PRESENTACION_KEY] = df_raw.iloc[:, 13]
                    df_temp[COL_CONTROL_KEY] = df_raw.iloc[:, 14]
                else:
                    for k_name, col_key in MAPEO_COLUMNAS.items():
                        matched = [c for c in df_raw.columns if k_name.lower() in str(c).lower()]
                        if matched:
                            df_temp[col_key] = df_raw[matched[0]]
                        else:
                            df_temp[col_key] = "SIN REGISTRO"

                if not df_temp.empty: 
                    df_temp = df_temp.reindex(columns=COLUMNAS_SELECCIONADAS)
                    file_dfs_dict[f_name] = df_temp
                    df_list.append(df_temp) 
                    total_columnas_mapeadas += 1

            if df_list: 
                datos = pd.concat(df_list, ignore_index=True)
            if datos is None or datos.empty or total_columnas_mapeadas == 0: 
                st.warning("No se encontraron columnas mapeables en los archivos.") 
                datos = None
        except Exception as e: st.error(f"Error al combinar archivos: {e}"); datos = None

    if datos is None: 
        try: 
            datos = pd.read_excel(MASTER_EXCEL) 
            columnas_existentes = [col for col in COLUMNAS_SELECCIONADAS if col in datos.columns] 
            datos = datos[columnas_existentes] 
            file_dfs_dict["datos_maestros.xlsx"] = datos
        except: 
            np.random.seed(42)
            n_dummy_jun = 150
            n_dummy_jul = 170
            n_dummy_aug = 160
            n_dummy_sep = 210
            n_dummy_oct = 190
            
            fechas_jun = pd.date_range(start='2026-06-01', periods=n_dummy_jun, freq='D')
            fechas_jul = pd.date_range(start='2026-07-01', periods=n_dummy_jul, freq='D')
            fechas_aug = pd.date_range(start='2026-08-01', periods=n_dummy_aug, freq='D')
            fechas_sep = pd.date_range(start='2026-09-01', periods=n_dummy_sep, freq='D')
            fechas_oct = pd.date_range(start='2026-10-01', periods=n_dummy_oct, freq='D')
            
            fechas = list(fechas_jun) + list(fechas_jul) + list(fechas_aug) + list(fechas_sep) + list(fechas_oct)
            n_total = len(fechas)
            
            ciudades = np.random.choice(['AMBATO', 'QUITO', 'GUAYAQUIL', 'CUENCA', 'LATACUNGA'], n_total)
            tecnicos = np.random.choice(['JUAN PEREZ', 'PEDRO LOPEZ', 'ANA GOMEZ', 'LUIS TORRES', 'CARLOS MERA'], n_total)
            empresas = ['ISERTEL'] * n_total
            motivos = np.random.choice(['OPE/SATISFACTORIO', 'PROBLEMA DE NAVEGACION', 'CAMBIO DE MATERIAL', 'SIN TONO'], n_total)
            gpon_vals = np.random.choice(['SI', 'NO'], n_total, p=[0.85, 0.15])
            presentaciones = np.random.choice(['Instalación con Materiales', 'Visita con Materiales', 'Migración con Materiales'], n_total)
            control_vals = np.random.choice(['DENTRO DEL SLA', 'FUERA DEL SLA', 'SIN REGISTRO'], n_total, p=[0.7, 0.2, 0.1])

            datos = pd.DataFrame({
                COL_FECHA_KEY: fechas,
                COL_CIUDAD_KEY: ciudades,
                COL_TECNICO_KEY: tecnicos,
                COL_EMPRESA_KEY: empresas,
                COL_MOTIVO_KEY: motivos,
                COL_GPON_KEY: gpon_vals,
                COL_TIPO_PRESENTACION_KEY: presentaciones,
                COL_CONTROL_KEY: control_vals
            })
            file_dfs_dict["DETALLE LIQUIDADO ISERTEL JUN-OCT2026.xlsx"] = datos.copy()

    # --- TABS --- 
    tabs = ["📊 Dashboard", "🆚 Comparativa (Versus - Multimes)", "📁 Paginación por Documento", "⚙️ Administración de Datos"] 
    if st.session_state.rol.lower() == "admin": 
        tab_dashboard, tab_versus, tab_docs, tab_admin = st.tabs(tabs) 
    else: 
        tab_dashboard, tab_versus, tab_docs = st.tabs(["📊 Dashboard", "🆚 Comparativa (Versus - Multimes)", "📁 Paginación por Documento"])
        tab_admin = None

    # Pre-procesamiento de base limpia general para todas las pestañas
    datos_base_limpia = None
    if datos is not None and not datos.empty:
        datos_base_limpia = datos.copy() 
        datos_base_limpia[COL_TEMP_DATETIME] = pd.to_datetime(datos_base_limpia[COL_FECHA_KEY], errors='coerce') 
        datos_base_limpia.dropna(subset=[COL_TEMP_DATETIME], inplace=True)
        
        datos_base_limpia[COL_FILTRO_TECNICO] = datos_base_limpia[COL_TECNICO_KEY].apply(clean_tecnico)
        datos_base_limpia[COL_FILTRO_CIUDAD] = datos_base_limpia[COL_CIUDAD_KEY].apply(clean_ciudad)
        datos_base_limpia[COL_FILTRO_EMPRESA] = datos_base_limpia[COL_EMPRESA_KEY].apply(clean_text)
        datos_base_limpia[COL_FILTRO_MOTIVO] = datos_base_limpia[COL_MOTIVO_KEY].apply(clean_text)
        datos_base_limpia[COL_FILTRO_GPON] = datos_base_limpia[COL_GPON_KEY].apply(clean_text)
        datos_base_limpia[COL_FILTRO_PRESENTACION] = datos_base_limpia[COL_TIPO_PRESENTACION_KEY].apply(clean_text)
        datos_base_limpia[COL_FILTRO_CONTROL] = datos_base_limpia[COL_CONTROL_KEY].apply(clean_text)
        
        datos_base_limpia['PERIODO_MES'] = datos_base_limpia[COL_TEMP_DATETIME].dt.to_period('M')
        datos_base_limpia['MES_ANO'] = datos_base_limpia[COL_TEMP_DATETIME].dt.strftime('%Y-%m (%B)')

    # ======================================================================
    # --- PESTAÑA: 🆚 COMPARATIVA (VERSUS - MULTIMES) ---
    # ======================================================================
    with tab_versus:
        st.header("🆚 Módulo Comparativo Multimes (Mes vs Mes vs Mes...)")
        st.markdown("Compara múltiples meses de forma individual y simultánea, evaluando el total de cada mes sin duplicaciones.")
        
        if datos_base_limpia is None or datos_base_limpia.empty:
            st.warning("No hay datos cargados para realizar comparaciones.")
        else:
            with st.container(border=True):
                st.markdown("### 🎛️ Selección de Meses y Filtros Compartidos")
                
                df_unicos_periodos = datos_base_limpia[['PERIODO_MES', 'MES_ANO']].drop_duplicates().sort_values(by='PERIODO_MES')
                meses_unicos = df_unicos_periodos['MES_ANO'].unique().tolist()
                
                if not meses_unicos:
                    st.warning("No se encontraron meses válidos en los datos.")
                else:
                    col_sel_modo, col_sel_meses = st.columns([1, 2])
                    with col_sel_modo:
                        modo_sel_vs = st.radio("Método de Selección:", ["Rango de Meses (Desde - Hasta)", "Selección Múltiple de Meses"], key="vs_modo_sel")
                    
                    with col_sel_meses:
                        if modo_sel_vs == "Selección Múltiple de Meses":
                            default_meses = meses_unicos[:5] if len(meses_unicos) >= 5 else meses_unicos
                            meses_seleccionados = st.multiselect("Seleccione los meses a comparar:", options=meses_unicos, default=default_meses, key="vs_multi_meses")
                        else:
                            rc1, rc2 = st.columns(2)
                            with rc1:
                                mes_ini_m = st.selectbox("Desde Mes:", meses_unicos, index=0, key="vs_m_ini")
                            with rc2:
                                mes_fin_m = st.selectbox("Hasta Mes:", meses_unicos, index=len(meses_unicos)-1, key="vs_m_fin")
                            
                            idx_i = meses_unicos.index(mes_ini_m)
                            idx_f = meses_unicos.index(mes_fin_m)
                            if idx_i > idx_f: idx_i, idx_f = idx_f, idx_i
                            meses_seleccionados = meses_unicos[idx_i:idx_f+1]
                    
                    st.markdown("---")
                    st.markdown("#### 🎯 Segmentación Compartida")
                    
                    # 5 Columnas para incluir el filtro de Tipo de Presentación
                    vs_c1, vs_c2, vs_c3, vs_c4, vs_c5 = st.columns(5)
                    
                    op_vs_ciu = sorted(datos_base_limpia[COL_FILTRO_CIUDAD].unique().tolist())
                    op_vs_tec = sorted(datos_base_limpia[COL_FILTRO_TECNICO].unique().tolist())
                    op_vs_gpon = sorted(datos_base_limpia[COL_FILTRO_GPON].unique().tolist())
                    op_vs_ctrl = sorted(datos_base_limpia[COL_FILTRO_CONTROL].unique().tolist())
                    op_vs_pres = sorted(datos_base_limpia[COL_FILTRO_PRESENTACION].unique().tolist())

                    with vs_c1:
                        vs_f_ciu = st.multiselect("Ciudad (VS):", options=op_vs_ciu, key="vs_m_ciu", placeholder="Todas")
                    with vs_c2:
                        vs_f_tec = st.multiselect("Técnico (VS):", options=op_vs_tec, key="vs_m_tec", placeholder="Todos")
                    with vs_c3:
                        vs_f_gpon = st.multiselect("GPON (VS):", options=op_vs_gpon, key="vs_m_gpon", placeholder="Todos")
                    with vs_c4:
                        vs_f_ctrl = st.multiselect("Control SLA (VS):", options=op_vs_ctrl, key="vs_m_ctrl", placeholder="Todos")
                    with vs_c5:
                        vs_f_pres = st.multiselect("Presentación (VS):", options=op_vs_pres, key="vs_m_pres", placeholder="Todas")

            if not meses_seleccionados:
                st.warning("Por favor seleccione al menos un mes para realizar la comparativa.")
            else:
                dict_dfs_meses = {}
                for m in meses_seleccionados:
                    sub = datos_base_limpia[datos_base_limpia['MES_ANO'] == m].copy()
                    if vs_f_ciu: sub = sub[sub[COL_FILTRO_CIUDAD].isin(vs_f_ciu)]
                    if vs_f_tec: sub = sub[sub[COL_FILTRO_TECNICO].isin(vs_f_tec)]
                    if vs_f_gpon: sub = sub[sub[COL_FILTRO_GPON].isin(vs_f_gpon)]
                    if vs_f_ctrl: sub = sub[sub[COL_FILTRO_CONTROL].isin(vs_f_ctrl)]
                    if vs_f_pres: sub = sub[sub[COL_FILTRO_PRESENTACION].isin(vs_f_pres)]
                    dict_dfs_meses[m] = sub

                st.markdown("### 📊 Resultados del Versus Multimes")
                num_meses = len(meses_seleccionados)
                st.info(f"Comparando simultáneamente **{num_meses} meses**: " + " vs ".join([f"**{m}**" for m in meses_seleccionados]))
                
                cols_cards = st.columns(min(num_meses, 6))
                for idx, m in enumerate(meses_seleccionados):
                    col_target = cols_cards[idx % len(cols_cards)]
                    tot_m = len(dict_dfs_meses[m])
                    with col_target:
                        st.markdown(f"""
                        <div class="vs-card-item">
                            <h4 style="margin:0; font-size: 0.95rem; color: #1E3A8A;">{m}</h4>
                            <h2 style="margin: 5px 0 0 0; font-size: 1.4rem; color: #0D47A1;">{tot_m:,}</h2>
                            <p style="margin:0; font-size: 0.8rem; color: #555;">registros</p>
                        </div>
                        """, unsafe_allow_html=True)

                st.markdown("---")
                st.markdown("#### 📋 Tabla Comparativa con Tendencia")
                
                def get_indicators_dict(df_sub):
                    tot = len(df_sub)
                    dentro = len(df_sub[df_sub[COL_FILTRO_CONTROL].str.contains('DENTRO', case=False, na=False)])
                    fuera = len(df_sub[df_sub[COL_FILTRO_CONTROL].str.contains('FUERA', case=False, na=False)])
                    return {
                        "Total Registros": tot,
                        "Instalaciones GPON (Sí)": len(df_sub[df_sub[COL_FILTRO_GPON] == 'SI']),
                        "Instalaciones GPON (No)": len(df_sub[df_sub[COL_FILTRO_GPON] == 'NO']),
                        "Dentro del SLA (%)": f"{(dentro/tot*100):.1f}%" if tot > 0 else "0.0%",
                        "Fuera del SLA (%)": f"{(fuera/tot*100):.1f}%" if tot > 0 else "0.0%"
                    }

                indicadores_nombres = [
                    "Total Registros", 
                    "Instalaciones GPON (Sí)", 
                    "Instalaciones GPON (No)", 
                    "Dentro del SLA (%)", 
                    "Fuera del SLA (%)"
                ]

                tabla_vs_data = []
                m_primero = meses_seleccionados[0]
                m_ultimo = meses_seleccionados[-1]

                for ind in indicadores_nombres:
                    row_dict = {"Indicador": ind}
                    val_primero = 0
                    val_ultimo = 0
                    
                    for idx_m, m in enumerate(meses_seleccionados):
                        ind_vals = get_indicators_dict(dict_dfs_meses[m])
                        val_curr = ind_vals.get(ind, 0)
                        row_dict[m] = f"{val_curr:,}" if isinstance(val_curr, (int, float)) else val_curr
                        
                        if idx_m == 0 and isinstance(val_curr, (int, float)): val_primero = val_curr
                        if idx_m == len(meses_seleccionados) - 1 and isinstance(val_curr, (int, float)): val_ultimo = val_curr

                    if isinstance(val_primero, (int, float)) and isinstance(val_ultimo, (int, float)):
                        diff = val_ultimo - val_primero
                        if diff > 0: estado_tendencia = "🟢 Aumentó"
                        elif diff < 0: estado_tendencia = "🔴 Disminuyó"
                        else: estado_tendencia = "⚪ Se Mantuvo Igual"
                    else:
                        estado_tendencia = "➖ N/A"

                    row_dict["Tendencia Global"] = estado_tendencia
                    tabla_vs_data.append(row_dict)

                df_multimes_summary = pd.DataFrame(tabla_vs_data)
                st.dataframe(df_multimes_summary, use_container_width=True, hide_index=True)

    # ======================================================================
    # --- PESTAÑA: PAGINACIÓN POR DOCUMENTO ---
    # ======================================================================
    with tab_docs:
        st.header("📁 Paginación y Visualización por Documento Cargado")
        st.markdown("Navega página por página a través de cada documento o mes cargado en el sistema.")
        
        if file_dfs_dict:
            nombres_docs = list(file_dfs_dict.keys())
            selected_doc = st.selectbox("Seleccione el Documento (Página de Registro):", options=nombres_docs, key="doc_pagination_select")
            
            if selected_doc in file_dfs_dict:
                df_doc_actual = file_dfs_dict[selected_doc].copy()
                st.info(f"📄 Mostrando documento: **{selected_doc}** | Total de registros: **{len(df_doc_actual):,}**")
                
                df_doc_vista = df_doc_actual.rename(columns=FINAL_RENAMING_MAP)
                
                if COL_TECNICO_KEY in df_doc_actual.columns and FINAL_RENAMING_MAP[COL_TECNICO_KEY] in df_doc_vista.columns:
                    df_doc_vista[FINAL_RENAMING_MAP[COL_TECNICO_KEY]] = df_doc_actual[COL_TECNICO_KEY].apply(clean_tecnico)

                st.dataframe(df_doc_vista, use_container_width=True, hide_index=True)

                buf_doc = io.BytesIO()
                df_doc_vista.to_excel(buf_doc, index=False)
                buf_doc.seek(0)
                st.download_button(
                    label=f"⬇️ Descargar '{selected_doc}' procesado",
                    data=buf_doc,
                    file_name=f"procesado_{selected_doc}.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
        else:
            st.warning("No hay documentos cargados en el sistema.")

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

    # ======================================================================
    # --- PESTAÑA DEL DASHBOARD --- 
    # ======================================================================
    with tab_dashboard: 
        if datos is None or datos.empty: 
            st.warning("No hay datos para mostrar.") 
        else:
            if datos_base_limpia.empty: 
                st.warning("No hay registros con fechas válidas para mostrar.") 
            else:
                @st.cache_data 
                def get_multiselect_options(df, col_key_filtro): 
                    if col_key_filtro not in df.columns or df.empty: return [] 
                    opciones = sorted([v for v in df[col_key_filtro].astype(str).unique() if pd.notna(v) and str(v).strip() not in ('nan', 'none', '')]) 
                    return opciones

                @st.cache_data 
                def apply_filter(df, col_key_filtro, selected_options): 
                    if not selected_options: return df
                    if col_key_filtro not in df.columns: return df 
                    return df[df[col_key_filtro].astype(str).isin(selected_options)]
                    
                # -----------------------------------------------------------------------------
                # --- PANEL DE CONTROL: FILTROS DE SEGMENTACIÓN --- 
                # -----------------------------------------------------------------------------
                with st.container(border=True):
                    col_header_filtros, col_btn_limpiar = st.columns([6, 1])
                    with col_header_filtros:
                        st.markdown("#### ⚙️ Filtros de Segmentación Avanzada") 
                    with col_btn_limpiar:
                         if st.button("🧹 Limpiar Filtros", use_container_width=True):
                            keys_to_clear = ['f_desde', 'f_hasta', 'f_ciu', 'f_tec', 'f_emp', 'f_mot', 'f_gpon', 'f_pres', 'f_ctrl']
                            for k in keys_to_clear:
                                if k in st.session_state:
                                    st.session_state[k] = [] if 'desde' not in k and 'hasta' not in k else st.session_state[k]
                            st.rerun()

                    c1, c2, c3, c4, c5 = st.columns([1.1, 1.1, 1.2, 1.2, 1.2])
                    with c1:
                        min_date_global = datos_base_limpia[COL_TEMP_DATETIME].min().replace(hour=0, minute=0, second=0, microsecond=0) 
                        max_date_global = datos_base_limpia[COL_TEMP_DATETIME].max().replace(hour=0, minute=0, second=0, microsecond=0) 
                        date_from = st.date_input("Desde:", value=min_date_global, min_value=min_date_global, max_value=max_date_global, key='f_desde')
                    with c2:
                        date_to = st.date_input("Hasta:", value=max_date_global, min_value=min_date_global, max_value=max_date_global, key='f_hasta')
                    
                    if date_from > date_to: 
                        st.error("⚠️ Fecha 'Desde' mayor que 'Hasta'."); st.stop()

                    filtro_inicio = pd.to_datetime(date_from) 
                    filtro_fin = pd.to_datetime(date_to) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)
                    df_base_fecha = datos_base_limpia[
                        (datos_base_limpia[COL_TEMP_DATETIME] >= filtro_inicio) & 
                        (datos_base_limpia[COL_TEMP_DATETIME] <= filtro_fin)
                    ].copy()

                    s_ciu = st.session_state.get('f_ciu', [])
                    s_tec = st.session_state.get('f_tec', [])
                    s_emp = st.session_state.get('f_emp', [])
                    s_mot = st.session_state.get('f_mot', [])
                    s_gpon = st.session_state.get('f_gpon', [])
                    s_pres = st.session_state.get('f_pres', [])
                    s_ctrl = st.session_state.get('f_ctrl', [])

                    # LÓGICA DE FILTRADO CRUZADO COHERENTE
                    df_c = apply_filter(df_base_fecha, COL_FILTRO_TECNICO, s_tec)
                    df_c = apply_filter(df_c, COL_FILTRO_EMPRESA, s_emp)
                    df_c = apply_filter(df_c, COL_FILTRO_MOTIVO, s_mot)
                    df_c = apply_filter(df_c, COL_FILTRO_GPON, s_gpon)
                    df_c = apply_filter(df_c, COL_FILTRO_PRESENTACION, s_pres)
                    df_c = apply_filter(df_c, COL_FILTRO_CONTROL, s_ctrl)
                    op_ciu = get_multiselect_options(df_c, COL_FILTRO_CIUDAD)

                    df_t = apply_filter(df_base_fecha, COL_FILTRO_CIUDAD, s_ciu)
                    df_t = apply_filter(df_t, COL_FILTRO_EMPRESA, s_emp)
                    df_t = apply_filter(df_t, COL_FILTRO_MOTIVO, s_mot)
                    df_t = apply_filter(df_t, COL_FILTRO_GPON, s_gpon)
                    df_t = apply_filter(df_t, COL_FILTRO_PRESENTACION, s_pres)
                    df_t = apply_filter(df_t, COL_FILTRO_CONTROL, s_ctrl)
                    op_tec = get_multiselect_options(df_t, COL_FILTRO_TECNICO)

                    df_em = apply_filter(df_base_fecha, COL_FILTRO_CIUDAD, s_ciu)
                    df_em = apply_filter(df_em, COL_FILTRO_TECNICO, s_tec)
                    df_em = apply_filter(df_em, COL_FILTRO_MOTIVO, s_mot)
                    df_em = apply_filter(df_em, COL_FILTRO_GPON, s_gpon)
                    df_em = apply_filter(df_em, COL_FILTRO_PRESENTACION, s_pres)
                    df_em = apply_filter(df_em, COL_FILTRO_CONTROL, s_ctrl)
                    op_emp = get_multiselect_options(df_em, COL_FILTRO_EMPRESA)

                    df_m = apply_filter(df_base_fecha, COL_FILTRO_CIUDAD, s_ciu)
                    df_m = apply_filter(df_m, COL_FILTRO_TECNICO, s_tec)
                    df_m = apply_filter(df_m, COL_FILTRO_EMPRESA, s_emp)
                    df_m = apply_filter(df_m, COL_FILTRO_GPON, s_gpon)
                    df_m = apply_filter(df_m, COL_FILTRO_PRESENTACION, s_pres)
                    df_m = apply_filter(df_m, COL_FILTRO_CONTROL, s_ctrl)
                    op_mot = get_multiselect_options(df_m, COL_FILTRO_MOTIVO)

                    df_g = apply_filter(df_base_fecha, COL_FILTRO_CIUDAD, s_ciu)
                    df_g = apply_filter(df_g, COL_FILTRO_TECNICO, s_tec)
                    df_g = apply_filter(df_g, COL_FILTRO_EMPRESA, s_emp)
                    df_g = apply_filter(df_g, COL_FILTRO_MOTIVO, s_mot)
                    df_g = apply_filter(df_g, COL_FILTRO_PRESENTACION, s_pres)
                    df_g = apply_filter(df_g, COL_FILTRO_CONTROL, s_ctrl)
                    op_gpon = get_multiselect_options(df_g, COL_FILTRO_GPON)

                    df_pr = apply_filter(df_base_fecha, COL_FILTRO_CIUDAD, s_ciu)
                    df_pr = apply_filter(df_pr, COL_FILTRO_TECNICO, s_tec)
                    df_pr = apply_filter(df_pr, COL_FILTRO_EMPRESA, s_emp)
                    df_pr = apply_filter(df_pr, COL_FILTRO_MOTIVO, s_mot)
                    df_pr = apply_filter(df_pr, COL_FILTRO_GPON, s_gpon)
                    df_pr = apply_filter(df_pr, COL_FILTRO_CONTROL, s_ctrl)
                    op_pres = get_multiselect_options(df_pr, COL_FILTRO_PRESENTACION)

                    df_ct = apply_filter(df_base_fecha, COL_FILTRO_CIUDAD, s_ciu)
                    df_ct = apply_filter(df_ct, COL_FILTRO_TECNICO, s_tec)
                    df_ct = apply_filter(df_ct, COL_FILTRO_EMPRESA, s_emp)
                    df_ct = apply_filter(df_ct, COL_FILTRO_MOTIVO, s_mot)
                    df_ct = apply_filter(df_ct, COL_FILTRO_GPON, s_gpon)
                    df_ct = apply_filter(df_ct, COL_FILTRO_PRESENTACION, s_pres)
                    op_ctrl = get_multiselect_options(df_ct, COL_FILTRO_CONTROL)

                    valid_ciu = [v for v in s_ciu if v in op_ciu]
                    valid_tec = [v for v in s_tec if v in op_tec]
                    valid_emp = [v for v in s_emp if v in op_emp]
                    valid_mot = [v for v in s_mot if v in op_mot]
                    valid_gpon = [v for v in s_gpon if v in op_gpon]
                    valid_pres = [v for v in s_pres if v in op_pres]
                    valid_ctrl = [v for v in s_ctrl if v in op_ctrl]

                    st.session_state['f_ciu'] = valid_ciu
                    st.session_state['f_tec'] = valid_tec
                    st.session_state['f_emp'] = valid_emp
                    st.session_state['f_mot'] = valid_mot
                    st.session_state['f_gpon'] = valid_gpon
                    st.session_state['f_pres'] = valid_pres
                    st.session_state['f_ctrl'] = valid_ctrl

                    with c3:
                        filtro_ciudad = st.multiselect(f"**{COL_CIUDAD_DESCRIPTIVA}**:", options=op_ciu, key='f_ciu', placeholder="Ciudad")
                    with c4:
                        filtro_tecnico = st.multiselect(f"**{COL_TECNICO_DESCRIPTIVA}**:", options=op_tec, key='f_tec', placeholder="Técnico")
                    with c5:
                        filtro_empresa = st.multiselect(f"**{COL_EMPRESA_DESCRIPTIVA}**:", options=op_emp, key='f_emp', placeholder="Empresa")

                    c6, c7, c8, c9 = st.columns([1.3, 1.2, 1.3, 1.2])
                    with c6:
                        filtro_motivo = st.multiselect(f"**{COL_MOTIVO_DESCRIPTIVA}**:", options=op_mot, key='f_mot', placeholder="Motivo")
                    with c7:
                        filtro_gpon = st.multiselect(f"**{COL_GPON_DESCRIPTIVA}**:", options=op_gpon, key='f_gpon', placeholder="GPON")
                    with c8:
                        filtro_presentacion = st.multiselect(f"**{COL_TIPO_PRESENTACION_DESCRIPTIVA}**:", options=op_pres, key='f_pres', placeholder="Presentación")
                    with c9:
                        filtro_control = st.multiselect(f"**{COL_CONTROL_DESCRIPTIVA}**:", options=op_ctrl, key='f_ctrl', placeholder="Control")

                    df_final = df_base_fecha.copy()
                    df_final = apply_filter(df_final, COL_FILTRO_CIUDAD, filtro_ciudad)
                    df_final = apply_filter(df_final, COL_FILTRO_TECNICO, filtro_tecnico)
                    df_final = apply_filter(df_final, COL_FILTRO_EMPRESA, filtro_empresa)
                    df_final = apply_filter(df_final, COL_FILTRO_MOTIVO, filtro_motivo)
                    df_final = apply_filter(df_final, COL_FILTRO_GPON, filtro_gpon)
                    df_final = apply_filter(df_final, COL_FILTRO_PRESENTACION, filtro_presentacion)
                    df_final = apply_filter(df_final, COL_FILTRO_CONTROL, filtro_control)
                    
                    datos_filtrados = df_final

                # -----------------------------------------------------------------------------
                # --- MÉTRICAS CLAVE (KPIs) CON PORCENTAJES EN SLA --- 
                # -----------------------------------------------------------------------------
                with st.container(border=True):
                    st.markdown("#### 🎯 Métricas Clave (KPIs Individuales)") 
                    
                    total_registros = len(datos_filtrados)
                    total_gpon_si = len(datos_filtrados[datos_filtrados[COL_FILTRO_GPON] == 'SI']) if COL_FILTRO_GPON in datos_filtrados.columns else 0
                    total_gpon_no = len(datos_filtrados[datos_filtrados[COL_FILTRO_GPON] == 'NO']) if COL_FILTRO_GPON in datos_filtrados.columns else 0
                    
                    total_dentro_sla = len(datos_filtrados[datos_filtrados[COL_FILTRO_CONTROL].str.contains('DENTRO', case=False, na=False)]) if COL_FILTRO_CONTROL in datos_filtrados.columns else 0
                    total_fuera_sla = len(datos_filtrados[datos_filtrados[COL_FILTRO_CONTROL].str.contains('FUERA', case=False, na=False)]) if COL_FILTRO_CONTROL in datos_filtrados.columns else 0
                    
                    pct_dentro_sla = (total_dentro_sla / total_registros * 100) if total_registros > 0 else 0.0
                    pct_fuera_sla = (total_fuera_sla / total_registros * 100) if total_registros > 0 else 0.0
                    
                    total_tecnicos = datos_filtrados[COL_FILTRO_TECNICO].nunique() if COL_FILTRO_TECNICO in datos_filtrados.columns else 0

                    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
                    with kpi1:
                        st.markdown('<div class="metric-compact-container-total">', unsafe_allow_html=True)
                        st.metric("Total Registros", f"{total_registros:,}")
                        st.markdown('</div>', unsafe_allow_html=True)
                    with kpi2:
                        st.markdown('<div class="metric-compact-container">', unsafe_allow_html=True)
                        st.metric("GPON Sí", f"{total_gpon_si:,}")
                        st.markdown('</div>', unsafe_allow_html=True)
                    with kpi3:
                        st.markdown('<div class="metric-compact-container">', unsafe_allow_html=True)
                        st.metric("GPON No", f"{total_gpon_no:,}")
                        st.markdown('</div>', unsafe_allow_html=True)
                    with kpi4:
                        st.markdown('<div class="metric-compact-container">', unsafe_allow_html=True)
                        st.metric("Dentro SLA", f"{pct_dentro_sla:.1f}%")
                        st.markdown('</div>', unsafe_allow_html=True)
                    with kpi5:
                        st.markdown('<div class="metric-compact-container">', unsafe_allow_html=True)
                        st.metric("Fuera SLA", f"{pct_fuera_sla:.1f}%")
                        st.markdown('</div>', unsafe_allow_html=True)
                    with kpi6:
                        st.markdown('<div class="metric-compact-container">', unsafe_allow_html=True)
                        st.metric("Técnicos Activos", f"{total_tecnicos:,}")
                        st.markdown('</div>', unsafe_allow_html=True)

                st.markdown("---")

                # -----------------------------------------------------------------------------
                # --- LAYOUT PRINCIPAL: SECCIÓN DE GRÁFICOS INTUITIVOS --- 
                # -----------------------------------------------------------------------------
                st.markdown("#### 📈 Gráficos Intuitivos y Dinámicos")
                
                if not datos_filtrados.empty:
                    g1, g2 = st.columns(2)
                    
                    with g1:
                        with st.container(border=True):
                            st.markdown("##### ⏱️ Control de SLA")
                            df_ctrl = datos_filtrados[COL_CONTROL_KEY].value_counts().reset_index()
                            df_ctrl.columns = ['Control', 'Cantidad']
                            fig_c = px.bar(df_ctrl, x='Control', y='Cantidad', text='Cantidad', color='Control', color_discrete_sequence=['#4CAF50', '#F44336', '#9E9E9E'])
                            fig_c.update_layout(xaxis_title=None, yaxis_title=None, showlegend=False, margin=dict(t=25, b=25, l=15, r=15), height=320)
                            fig_c.update_traces(textposition='outside')
                            st.plotly_chart(fig_c, use_container_width=True)

                    with g2:
                        with st.container(border=True):
                            st.markdown("##### 🌐 Distribución GPON")
                            df_gpon = datos_filtrados[COL_GPON_KEY].value_counts().reset_index()
                            df_gpon.columns = ['GPON', 'Cantidad']
                            fig_g = px.pie(df_gpon, values='Cantidad', names='GPON', hole=0.4, color_discrete_sequence=['#2196F3', '#FF9800'])
                            fig_g.update_layout(margin=dict(t=25, b=15, l=15, r=15), height=320)
                            fig_g.update_traces(textposition='inside', textinfo='percent+label')
                            st.plotly_chart(fig_g, use_container_width=True)

                    g3, g4 = st.columns(2)
                    with g3:
                        with st.container(border=True):
                            st.markdown("##### 🏙️ Top Ciudades por Producción")
                            df_ciu = datos_filtrados[COL_CIUDAD_KEY].value_counts().head(7).reset_index()
                            df_ciu.columns = ['Ciudad', 'Cantidad']
                            fig_ci = px.bar(df_ciu, x='Ciudad', y='Cantidad', text='Cantidad', color='Ciudad', color_discrete_sequence=px.colors.qualitative.Pastel)
                            fig_ci.update_layout(xaxis_title=None, yaxis_title=None, showlegend=False, margin=dict(t=25, b=25, l=15, r=15), height=320)
                            fig_ci.update_traces(textposition='outside')
                            st.plotly_chart(fig_ci, use_container_width=True)

                    with g4:
                        with st.container(border=True):
                            st.markdown("##### 📌 Motivos Principales")
                            df_mot = datos_filtrados[COL_MOTIVO_KEY].value_counts().head(7).reset_index()
                            df_mot.columns = ['Motivo', 'Cantidad']
                            fig_m = px.bar(df_mot, x='Motivo', y='Cantidad', text='Cantidad', color='Motivo', color_discrete_sequence=px.colors.qualitative.Safe)
                            fig_m.update_layout(xaxis_title=None, yaxis_title=None, showlegend=False, margin=dict(t=25, b=25, l=15, r=15), height=320)
                            fig_m.update_traces(textposition='outside')
                            st.plotly_chart(fig_m, use_container_width=True)

                    with st.container(border=True):
                        st.markdown("##### 📅 Resumen de Producción por Periodo / Semana")
                        df_fecha = datos_filtrados.copy()
                        df_fecha['Semana'] = df_fecha[COL_TEMP_DATETIME].dt.to_period('W').dt.start_time.dt.strftime('%d %b %Y')
                        df_ev = df_fecha.groupby('Semana', sort=False).size().reset_index(name='Total Trabajos')
                        
                        fig_ev = px.bar(
                            df_ev, 
                            x='Semana', 
                            y='Total Trabajos', 
                            text='Total Trabajos', 
                            color_discrete_sequence=['#1E88E5']
                        )
                        fig_ev.update_layout(
                            xaxis_title="Semana / Periodo", 
                            yaxis_title="Trabajos Completados", 
                            margin=dict(t=25, b=30, l=25, r=25), 
                            height=350,
                            xaxis={'tickangle': -30}
                        )
                        fig_ev.update_traces(textposition='outside')
                        st.plotly_chart(fig_ev, use_container_width=True)

                else:
                    st.info("No hay datos disponibles para graficar.")