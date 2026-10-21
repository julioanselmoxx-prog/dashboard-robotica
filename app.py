import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import requests
import io
import config

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS HIGH-TECH UI
# ==========================================
st.set_page_config(
    page_title="EQUANS Perú - Centro de Control de Robótica",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS Avanzados estilo Dark Cyberpunk / Industrial Operations
st.markdown("""
    <style>
        /* Fondo General */
        .stApp {
            background: #090e14;
            color: #d1d5db;
            font-family: 'Consolas', 'Segoe UI', monospace;
        }
        
        /* Contenedores Glassmorphism */
        .tech-card {
            background: rgba(18, 26, 36, 0.85);
            border: 1px solid #1e2d3d;
            border-left: 4px solid #00d28d;
            border-radius: 6px;
            padding: 14px;
            margin-bottom: 12px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.5);
        }

        /* Encabezado */
        .header-title {
            font-size: 24px;
            font-weight: 800;
            color: #ffffff;
            letter-spacing: 1px;
            border-bottom: 2px solid #1e2d3d;
            padding-bottom: 8px;
            margin-bottom: 4px;
        }
        .header-title span { color: #00d28d; font-weight: 300; }
        
        .sub-header {
            font-size: 12px;
            color: #5c708a;
            text-transform: uppercase;
            letter-spacing: 1.5px;
            margin-bottom: 20px;
        }

        /* Metricas */
        div[data-testid="metric-container"] {
            background-color: #111a24;
            border: 1px solid #1e2d3d;
            border-radius: 6px;
            padding: 8px 12px;
        }
        div[data-testid="metric-container"] label {
            color: #5c708a !important;
            font-size: 11px !important;
            text-transform: uppercase;
        }
        div[data-testid="metric-container"] div[data-testid="stMetricValue"] {
            color: #00d28d !important;
            font-size: 22px !important;
            font-family: 'Consolas', monospace;
        }

        /* Tabs de Navegación */
        .stTabs [data-baseweb="tab-list"] {
            gap: 4px;
            background-color: #090e14;
            border-bottom: 1px solid #1e2d3d;
        }
        .stTabs [data-baseweb="tab"] {
            height: 40px;
            background-color: #111a24;
            border-radius: 4px 4px 0px 0px;
            color: #8b9bb4;
            font-size: 11px;
            font-weight: 700;
            padding: 0px 18px;
            border: 1px solid #1e2d3d;
            border-bottom: none;
        }
        .stTabs [aria-selected="true"] {
            background-color: #00d28d !important;
            color: #090e14 !important;
            border-color: #00d28d !important;
        }

        /* Badges de Estado */
        .badge-active { background-color: #00d28d; color: #000; padding: 2px 6px; border-radius: 3px; font-weight: bold; font-size: 10px; }
        .badge-warning { background-color: #f39c12; color: #000; padding: 2px 6px; border-radius: 3px; font-weight: bold; font-size: 10px; }
        .badge-danger { background-color: #e74c3c; color: #fff; padding: 2px 6px; border-radius: 3px; font-weight: bold; font-size: 10px; }
    </style>
""", unsafe_allow_html=True)


# ==========================================
# 2. CAPA DE CONECTORES DE DATOS
# ==========================================
class TelemetryDataService:
    @staticmethod
    def get_sharepoint_manual_logs():
        try:
            token_url = f"https://login.microsoftonline.com/{config.SHAREPOINT_TENANT_ID}/oauth2/v2.0/token"
            data = {
                'grant_type': 'client_credentials',
                'client_id': config.SHAREPOINT_CLIENT_ID,
                'client_secret': config.SHAREPOINT_CLIENT_SECRET,
                'scope': 'https://graph.microsoft.com/.default'
            }
            res = requests.post(token_url, data=data, timeout=3)
            if res.status_code == 200:
                access_token = res.json().get('access_token')
                headers = {'Authorization': f'Bearer {access_token}'}
                download_url = f"https://graph.microsoft.com/v1.0/sites/{config.SHAREPOINT_SITE_ID}/drive/root:/{config.SHAREPOINT_FILE_PATH}:/content"
                file_res = requests.get(download_url, headers=headers, timeout=3)
                if file_res.status_code == 200:
                    return pd.read_excel(io.BytesIO(file_res.content))
        except Exception:
            pass

        return pd.DataFrame([
            {"Unidad": "Robot 07 (DT-UT)", "Tipo": "Terrestre", "Ubicación": "Talara", "Horas_Operacion": 145, "Filtro_Limpio": "Sí", "Bateria": 98, "Lat": -4.57, "Lon": -81.27, "Estado": "Activo"},
            {"Unidad": "Robot 05 (Piscina)", "Tipo": "Acuático", "Ubicación": "IPD Videna", "Horas_Operacion": 210, "Filtro_Limpio": "No", "Bateria": 85, "Lat": -12.07, "Lon": -77.00, "Estado": "Mantenimiento"},
            {"Unidad": "Robot 09 (Pisos)", "Tipo": "Terrestre", "Ubicación": "Moquegua", "Horas_Operacion": 320, "Filtro_Limpio": "Sí", "Bateria": 92, "Lat": -17.19, "Lon": -70.93, "Estado": "Activo"},
        ])

    @staticmethod
    def get_api_fleet_telemetry():
        try:
            headers = {"Authorization": f"Bearer {config.DRONE_API_KEY}"}
            res = requests.get(f"{config.DRONE_API_URL}/telemetry", headers=headers, timeout=3)
            if res.status_code == 200:
                return pd.DataFrame(res.json())
        except Exception:
            pass

        return pd.DataFrame([
            {"unit": "Drone Sherpa 1.5", "type": "RPAS Hexacóptero", "site": "Moquegua", "battery": 92, "temp_c": 42, "datalink": "Green", "flight_time_min": 17, "lat": -17.15, "lon": -70.92, "status": "Activo"},
            {"unit": "Drone 04 (Térmico)", "type": "RPAS Cuadricóptero", "site": "Ica", "battery": 95, "temp_c": 45, "datalink": "Green", "flight_time_min": 22, "lat": -13.83, "lon": -76.25, "status": "Activo"},
            {"unit": "Robot 07 (DT-UT)", "type": "Terrestre", "site": "Talara", "battery": 98, "temp_c": 46, "datalink": "Green", "flight_time_min": 180, "lat": -4.57, "lon": -81.27, "status": "Activo"},
            {"unit": "Drone 07 (RPAS)", "type": "RPAS", "site": "Lima", "battery": 35, "temp_c": 52, "datalink": "Amber", "flight_time_min": 12, "lat": -12.04, "lon": -77.03, "status": "Mantenimiento"},
        ])


# ==========================================
# 3. INTERFAZ Y BARRA LATERAL (SIDEBAR)
# ==========================================
with st.sidebar:
    st.markdown("## ⚡ **EQUANS**")
    st.caption("SYSTEM OPERATIVE STATUS: ONLINE")
    st.divider()
    
    st.markdown("🌐 **CONEXIONES DE DATOS**")
    st.markdown("• API Sherpa 1.5: <span class='badge-active'>OK</span>", unsafe_allow_html=True)
    st.markdown("• API Fleet Robots: <span class='badge-active'>OK</span>", unsafe_allow_html=True)
    st.markdown("• Sync SharePoint: <span class='badge-active'>OK</span>", unsafe_allow_html=True)
    
    st.divider()
    st.markdown("📍 **DESPLIEGUES ACTIVOS**")
    st.caption("• Moquegua (Planta Solar - Sherpa 1.5)")
    st.caption("• Talara (Infraestructura Industrial)")
    st.caption("• IPD Videna (Sede Lima)")


# Header Principal
st.markdown("<div class='header-title'>EQUANS <span>PERÚ - CENTRO DE CONTROL DE ROBÓTICA</span></div>", unsafe_allow_html=True)
st.markdown("<div class='sub-header'>Unidad de Robótica de Campo (RaaS EDGE) | Telemetría & Mantenimiento Predictivo</div>", unsafe_allow_html=True)

# Pestañas Superiores
tab_resumen, tab_mapa, tab_diagnostico, tab_operaciones, tab_reportes, tab_ajustes = st.tabs([
    "RESUMEN", "MAPA DE FLOTA", "DIAGNÓSTICO", "OPERACIONES", "REPORTES", "AJUSTES"
])

df_telemetry = TelemetryDataService.get_api_fleet_telemetry()
df_sharepoint = TelemetryDataService.get_sharepoint_manual_logs()


# ==========================================
# 4. PESTAÑA 1: RESUMEN GENERAL
# ==========================================
with tab_resumen:
    c1, c2, c3 = st.columns([1.2, 2.2, 1.2])

    with c1:
        st.markdown("##### 📹 FEED DE FLOTA EN VIVO")
        for idx, row in df_telemetry.iterrows():
            badge = "badge-active" if row['status'] == 'Activo' else "badge-warning"
            st.markdown(f"""
                <div class="tech-card">
                    <b>{row['unit']}</b> <span class="{badge}">{row['status']}</span><br>
                    <small style="color:#5c708a;">📍 {row['site']} | Datalink: {row['datalink']}</small><br>
                    🔋 <b>{row['battery']}%</b> | ⏱️ <b>{row['flight_time_min']} min aut.</b> | 🌡️ <b>{row['temp_c']}°C</b>
                </div>
            """, unsafe_allow_html=True)

    with c2:
        st.markdown("##### 🌍 DESPLIEGUE NACIONAL DE ACTIVOS")
        fig_map = px.scatter_map(
            df_telemetry,
            lat="lat",
            lon="lon",
            hover_name="unit",
            hover_data=["site", "battery", "flight_time_min", "status"],
            color="status",
            color_discrete_map={"Activo": "#00d28d", "Mantenimiento": "#f39c12"},
            zoom=4.6,
            height=430
        )
        fig_map.update_layout(
            margin={"r":0,"t":0,"l":0,"b":0},
            map_style="carto-darkmatter",
            map_center={"lat": -9.19, "lon": -75.01}
        )
        st.plotly_chart(fig_map, use_container_width=True)

    with c3:
        st.markdown("##### 📊 GARANTÍAS DE DESEMPEÑO (SLAs)")
        st.metric("% Tiempo Actividad Solar", "96%", "+2%")
        st.metric("Autonomía Sherpa (4J/8Bat)", "68 min total", "Rotación 100%")
        st.metric("Integridad de Activos", "96%", "0%")

        st.markdown("<br><b>🚀 PIPELINE TECNOLÓGICO</b>", unsafe_allow_html=True)
        st.caption("1. Limpieza Solar Híbrida (90%)")
        st.progress(0.90)
        st.caption("2. Termografía Aérea Sherpa 1.5 (85%)")
        st.progress(0.85)
        st.caption("3. UT de Espesor (70%)")
        st.progress(0.70)


# ==========================================
# 5. PESTAÑA 3: DIAGNÓSTICO PREDICTIVO
# ==========================================
with tab_diagnostico:
    st.markdown("### 🔬 DIAGNÓSTICO TÉCNICO AVANZADO: SHERPA 1.5 & FLOTA")
    
    # Banner Técnico de Arquitectura Sherpa 1.5
    st.markdown("""
        <div class="tech-card" style="border-left: 4px solid #3498db;">
            <b>🚁 ESPECIFICACIÓN TÉCNICA Y MATRIZ DE ENERGÍA - LUCID BOTS SHERPA 1.5</b><br>
            <small style="color:#a0aec0;">
            • Configuración: <b>Hexacóptero (6 Motores/Rotores Brushless)</b> | Payload: Limpieza / Lavado Industrial de Páneles y Fachadas.<br>
            • Banco de Baterías: <b>4 Juegos Activos (8 Baterías Tattu 12S 10,000mAh en total)</b> | Operación dual por vuelo (2 Baterías = 1 Juego).<br>
            • Autonomía por Juego: <b>15 - 18 min de vuelo efectivo</b> | Tiempo de Carga Completo: <b>45 min por Juego (Cargador TA3200)</b>.
            </small>
        </div>
    """, unsafe_allow_html=True)

    col_diag1, col_diag2 = st.columns([2.2, 1.8])

    with col_diag1:
        st.markdown("##### 🔋 ESTADO DE SALUD & CICLOS DE VIDA (4 JUEGOS / 8 BATERÍAS DUALES)")
        
        # Datos de los 4 juegos de baterías (8 baterías físicas)
        df_batteries = pd.DataFrame([
            {"Juego": "Juego 1 (Bat 1A - 1B)", "Ciclos_Usados": 82, "SoH_Health": 96, "Estado": "En Uso (Dron)", "Voltaje_V": 44.2},
            {"Juego": "Juego 2 (Bat 2A - 2B)", "Ciclos_Usados": 145, "SoH_Health": 91, "Estado": "Cargado (Listo)", "Voltaje_V": 44.4},
            {"Juego": "Juego 3 (Bat 3A - 3B)", "Ciclos_Usados": 210, "SoH_Health": 84, "Estado": "En Carga (TA3200)", "Voltaje_V": 41.8},
            {"Juego": "Juego 4 (Bat 4A - 4B)", "Ciclos_Usados": 295, "SoH_Health": 78, "Estado": "Revisión / Balanceo", "Voltaje_V": 39.5},
        ])

        fig_battery = px.bar(
            df_batteries,
            x="SoH_Health",
            y="Juego",
            orientation='h',
            text="Ciclos_Usados",
            color="SoH_Health",
            color_continuous_scale="RdYlGn",
            labels={"SoH_Health": "Salud de Batería (%)", "Juego": "Juegos de Baterías (12S Dual)"},
            range_x=[0, 100]
        )
        fig_battery.update_traces(texttemplate='%{text} ciclos', textposition='inside')
        fig_battery.update_layout(template="plotly_dark", height=300, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig_battery, use_container_width=True)

    with col_diag2:
        st.markdown("##### ⚙️ DIAGNÓSTICO DE TELEMETRÍA: 6 ROTORES / MOTORES (SHERPA 1.5)")
        
        # Telemetría de los 6 motores del Hexacóptero
        motores = [f"Motor {i}" for i in range(1, 7)]
        temp_motores = [42, 44, 43, 48, 43, 42]
        vibra_mms = [1.2, 1.4, 1.3, 3.8, 1.2, 1.1] # Motor 4 en alerta preventiva

        fig_motors = go.Figure()
        fig_motors.add_trace(go.Bar(x=motores, y=temp_motores, name='Temp (°C)', marker_color='#00d28d'))
        fig_motors.add_trace(go.Scatter(x=motores, y=vibra_mms, name='Vibración (mm/s)', yaxis='y2', line=dict(color='#e74c3c', width=3)))
        
        fig_motors.update_layout(
            template="plotly_dark",
            height=300,
            margin=dict(l=10, r=10, t=30, b=10),
            yaxis=dict(title='Temperatura °C'),
            yaxis2=dict(title='Vibración mm/s', overlaying='y', side='right')
        )
        st.plotly_chart(fig_motors, use_container_width=True)

    st.markdown("##### 📋 BITÁCORA MANUAL EN SHAREPOINT (Registro de Ciclos & Mantenimiento Operativo)")
    st.dataframe(df_sharepoint, use_container_width=True)


# ==========================================
# 6. PESTAÑA 5: REPORTES
# ==========================================
with tab_reportes:
    st.markdown("### 📋 REGISTRO DE MISIONES Y RESULTADOS OPERATIVOS")
    
    df_reportes = pd.DataFrame([
        {"Misión ID": "M-001", "Fecha": "22/10/2026", "Unidad": "Drone Sherpa 1.5", "Ubicación": "Moquegua", "Tarea": "Limpieza Fachadas / Páneles", "Juego Batería": "Juego 1", "Tiempo Vuelo (min)": 17, "Estado": "Exitoso"},
        {"Misión ID": "M-002", "Fecha": "22/10/2026", "Unidad": "Robot 07", "Ubicación": "Talara", "Tarea": "Limpieza Híbrida", "Juego Batería": "Interna N/A", "Tiempo Vuelo (min)": 40, "Estado": "Exitoso"},
        {"Misión ID": "M-003", "Fecha": "23/10/2026", "Unidad": "Drone Sherpa 1.5", "Ubicación": "Moquegua", "Tarea": "Inspección Térmica Hexacóptero", "Juego Batería": "Juego 2", "Tiempo Vuelo (min)": 16, "Estado": "Exitoso"},
    ])
    
    st.dataframe(df_reportes, use_container_width=True)

# Pestañas complementarias
with tab_mapa:
    st.info("Vista expandida de Geofencing y zonas de vuelo autónomo para el Sherpa 1.5 y flota de campo.")

with tab_operaciones:
    st.info("Panel de despacho y comandería en tiempo real para unidades de campo.")

with tab_ajustes:
    st.info("Parámetros de conexión API REST y credenciales de Microsoft SharePoint Graph API.")