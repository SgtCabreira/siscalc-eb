import streamlit as st
from core.database import init_db
from core.auth import (
    gerenciar_sessao,
    render_login,
    logout,
    obter_usuario_atual,
    render_sidebar_usuario,
    obter_menus_autorizados
)
import views.dashboard as v_dashboard
import views.notas_credito as v_nc
import views.notas_empenho as v_ne
import views.processos as v_processos
import views.recebimento as v_recebimento
import views.estoque as v_estoque
import views.relatorios as v_relatorios
import views.admin as v_admin

# Configuração Principal da Página
st.set_page_config(
    page_title="SisCalc - Exército Brasileiro",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="auto"
)

# Inicializa banco de dados e migrações automáticas
init_db()

# Controle de sessão persistente
gerenciar_sessao()

# Tela de Login caso não esteja autenticado
if not st.session_state.get('usuario'):
    st.title("🛡️ Sistema Integrado SALC & Almoxarifado - EB")
    st.info("Insira suas credenciais na barra lateral para acessar.")
    render_login()
    st.stop()

# Carrega militar e OM atual
user = obter_usuario_atual()
cor_om = user.get('cor_primaria') or '#1B4332'

# CSS Visual Tático Militar de Alto Contraste
st.markdown(f"""
<style>
/* =========================================================
   TEMA TÁTICO MILITAR DE ALTO CONTRASTE - 5ª CIA PE
   Fundo Cinza-Escuro Acetinado + Textos Nítidos em Branco
========================================================= */
html, body, [class*="css"], .stApp {{
    background: radial-gradient(circle at 50% 0%, #1a222d 0%, #11161f 55%, #0b0e14 100%) !important;
    color: #f8fafc !important;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
}}

.stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"], [data-testid="stToolbar"] {{
    background: radial-gradient(circle at 50% 0%, #1a222d 0%, #11161f 55%, #0b0e14 100%) !important;
    color: #f8fafc !important;
}}

[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, #131821 0%, #0c1017 100%) !important;
    border-right: 1px solid rgba(197, 160, 89, 0.3) !important;
    box-shadow: 4px 0 15px rgba(0, 0, 0, 0.5) !important;
}}

div[data-testid="stWidgetLabel"] *, 
label[data-testid="stWidgetLabel"] *, 
.stSelectbox label, .stTextInput label, .stNumberInput label, .stTextArea label, 
.stDateInput label, .stMultiSelect label, .stCheckbox label,
[data-testid="stWidgetLabel"] p, [data-testid="stWidgetLabel"] label,
.stMarkdown p, .stMarkdown span {{
    color: #f8fafc !important;
    font-weight: 700 !important;
    font-size: 13.5px !important;
    text-shadow: 0 1px 3px rgba(0, 0, 0, 0.7) !important;
}}

h1, h2, h3, h4, h5, [data-testid="stHeading"] *, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {{
    color: #ffffff !important;
    font-weight: 800 !important;
    text-shadow: 0 2px 4px rgba(0, 0, 0, 0.6) !important;
}}

input, textarea, [data-baseweb="input"], [data-baseweb="select"] > div, 
.stTextInput input, .stNumberInput input, .stTextArea textarea {{
    background-color: #1a2432 !important;
    color: #ffffff !important;
    border: 1px solid rgba(255, 255, 255, 0.25) !important;
    border-radius: 7px !important;
    font-weight: 600 !important;
}}

.stMetric {{
    border-left: 4px solid #C5A059;
    padding-left: 12px;
}}
.stMetric label {{
    color: #cbd5e1 !important;
    font-weight: 700 !important;
    font-size: 12px !important;
    text-transform: uppercase;
}}
.stMetric div[data-testid="stMetricValue"] {{
    color: #ffffff !important;
    font-weight: 900 !important;
}}

div[role="radiogroup"] > label {{
    background: linear-gradient(145deg, #1e2634, #141a24) !important;
    border: 1px solid rgba(255, 255, 255, 0.12) !important;
    border-left: 4px solid #3b485d !important;
    border-radius: 8px;
    padding: 10px 14px !important;
    margin-bottom: 8px !important;
    box-shadow: 0 4px 10px rgba(0, 0, 0, 0.4) !important;
    transition: all 0.2s ease;
    cursor: pointer;
}}
div[role="radiogroup"] > label:hover {{
    transform: translateY(-2px);
    border-left: 4px solid #C5A059 !important;
    box-shadow: 0 6px 14px rgba(0, 0, 0, 0.6) !important;
}}
div[role="radiogroup"] > label > div:first-child {{
    display: none !important;
}}
div[role="radiogroup"] > label div p {{
    font-size: 15px !important;
    font-weight: 700 !important;
    letter-spacing: 0.5px;
    margin: 0 !important;
    color: #e2e8f0 !important;
}}
div[role="radiogroup"] > label:has(input:checked) {{
    background: linear-gradient(145deg, #0A2240, #173e72) !important;
    border-left: 5px solid #C5A059 !important;
    box-shadow: inset 0 2px 5px rgba(0,0,0,0.4), 0 3px 10px rgba(10,34,64,0.5) !important;
}}
div[role="radiogroup"] > label:has(input:checked) div p {{
    color: #ffffff !important;
    font-weight: 900 !important;
}}

[data-testid="stSidebar"] div.stButton > button {{
    background: linear-gradient(145deg, #881337, #4c0519) !important;
    color: #ffffff !important;
    border: 1.5px solid #f43f5e !important;
    border-radius: 8px !important;
    font-weight: 800 !important;
    width: 100% !important;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.45) !important;
}}
[data-testid="stSidebar"] div.stButton > button:hover {{
    background: linear-gradient(145deg, #9f1239, #881337) !important;
    border-color: #fda4af !important;
    transform: translateY(-1px);
}}

.stTabs [data-baseweb="tab-list"] {{
    background: transparent;
    gap: 8px;
}}
.stTabs [data-baseweb="tab"] {{
    background: rgba(255, 255, 255, 0.06) !important;
    border-radius: 6px 6px 0 0;
    border: 1px solid rgba(255, 255, 255, 0.12);
    border-bottom: none;
    padding: 9px 18px;
}}
.stTabs [data-baseweb="tab"] p, .stTabs [data-baseweb="tab"] span, .stTabs [data-baseweb="tab"] div {{
    color: #cbd5e1 !important;
    font-weight: 700 !important;
    font-size: 14px !important;
}}
.stTabs [aria-selected="true"] {{
    background: linear-gradient(145deg, #243142, #18222e) !important;
    border-top: 3.5px solid #C5A059 !important;
}}
.stTabs [aria-selected="true"] p, .stTabs [aria-selected="true"] span, .stTabs [aria-selected="true"] div {{
    color: #ffffff !important;
    font-weight: 900 !important;
}}
</style>
""", unsafe_allow_html=True)

# Renderiza militar conectado, distintivo da OM e seletor de perfil
perfil_ativo = render_sidebar_usuario(user)

# Renderiza menu lateral com permissões de acesso
menu_opcoes = obter_menus_autorizados(perfil_ativo)
menu_selecionado = st.sidebar.radio("Navegação", menu_opcoes)

if st.sidebar.button("Sair do Sistema", key="btn_logout_sidebar"):
    logout()

st.sidebar.markdown('''
<div style="margin-top: 25px; padding-top: 10px; border-top: 1px solid rgba(255,255,255,0.1); text-align: center;">
    <div style="font-size: 11px; font-weight: bold; color: #aaa;">SisLog & Orç EB — v2.0.0 Modular</div>
    <div style="font-size: 10px; color: #888; margin-top: 2px;">Desenvolvido por: <b>2º Sgt Cabreira</b></div>
    <div style="font-size: 9px; color: #666;">5ª Cia PE — Curitiba/PR</div>
</div>
''', unsafe_allow_html=True)

# Roteamento Dinâmico para as Views
if menu_selecionado == "📊 Dashboard":
    v_dashboard.render(user)
elif menu_selecionado == "📑 Notas de Crédito":
    v_nc.render(user)
elif menu_selecionado == "📋 Notas de Empenho":
    v_ne.render(user)
elif menu_selecionado == "📝 Elaboração de Processos":
    v_processos.render(user)
elif menu_selecionado == "📦 Recebimento":
    v_recebimento.render(user)
elif menu_selecionado == "🏢 Gestão de Estoque":
    v_estoque.render(user)
elif menu_selecionado == "📁 Relatórios":
    v_relatorios.render(user)
elif menu_selecionado == "⚙️ Admin":
    v_admin.render(user)
