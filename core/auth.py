import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from core.database import get_connection

def gerenciar_sessao():
    """
    Controla a persistência e segurança da sessão do militar.
    Recupera logins recentes via parâmetros de URL seguros com tolerância de até 30 minutos.
    """
    if 'usuario' not in st.session_state:
        st.session_state['usuario'] = None

    if not st.session_state['usuario']:
        user_id_param = st.query_params.get("session_user")
        last_time_param = st.query_params.get("session_time")
        if user_id_param and last_time_param:
            try:
                last_dt = datetime.fromtimestamp(float(last_time_param))
                if datetime.now() - last_dt <= timedelta(minutes=30):
                    conn = get_connection()
                    c = conn.cursor()
                    c.execute("""
                    SELECT u.*, o.sigla as om_sigla, o.logo_base64, o.cor_primaria, o.lema 
                    FROM usuarios u JOIN oms o ON u.om_id = o.id WHERE u.id = ?
                    """, (int(user_id_param),))
                    u_row = c.fetchone()
                    conn.close()
                    if u_row:
                        st.session_state['usuario'] = dict(u_row)
                        perf_rst = [p.strip() for p in u_row['perfil'].split(",") if p.strip()]
                        st.session_state['perfil_ativo'] = perf_rst[0] if perf_rst else 'SALC'
                        st.query_params["session_time"] = str(int(datetime.now().timestamp()))
            except Exception:
                st.query_params.clear()

    if st.session_state.get('usuario'):
        st.query_params["session_user"] = str(st.session_state['usuario']['id'])
        st.query_params["session_time"] = str(int(datetime.now().timestamp()))

def render_login():
    """Exibe o formulário de login seguro na barra lateral."""
    st.sidebar.markdown("### 🔐 Acesso ao Sistema")
    conn = get_connection()
    oms = pd.read_sql_query("SELECT id, sigla FROM oms", conn)
    conn.close()

    lembrado_id = st.query_params.get("lembrar_id", "")
    lembrado_senha = st.query_params.get("lembrar_senha", "")

    with st.sidebar.form("form_login"):
        om_escolhida = st.selectbox("Organização Militar (OM)", oms['sigla'].tolist())
        identidade = st.text_input("Identidade Militar", value=lembrado_id)
        senha = st.text_input("Senha", type="password", value=lembrado_senha)
        salvar_senha = st.checkbox("💾 Salvar credenciais neste navegador", value=True)
        btn_entrar = st.form_submit_button("Entrar no Sistema", type="primary")

        if btn_entrar:
            conn = get_connection()
            c = conn.cursor()
            c.execute("""
            SELECT u.*, o.sigla as om_sigla, o.logo_base64, o.cor_primaria, o.lema 
            FROM usuarios u JOIN oms o ON u.om_id = o.id 
            WHERE u.identidade_militar = ? AND u.senha = ? AND o.sigla = ?
            """, (identidade.strip(), senha.strip(), om_escolhida))
            user = c.fetchone()
            conn.close()

            if user:
                st.session_state['usuario'] = dict(user)
                perfis_u = [p.strip() for p in user['perfil'].split(",") if p.strip()]
                st.session_state['perfil_ativo'] = perfis_u[0] if perfis_u else 'SALC'
                st.query_params["session_user"] = str(user['id'])
                st.query_params["session_time"] = str(int(datetime.now().timestamp()))
                if salvar_senha:
                    st.query_params["lembrar_id"] = identidade
                    st.query_params["lembrar_senha"] = senha
                st.rerun()
            else:
                st.error("Credenciais inválidas. Verifique sua OM, Identidade e Senha.")

def logout():
    """Encerra a sessão atual e limpa os cookies e parâmetros de URL."""
    st.session_state['usuario'] = None
    st.session_state['perfil_ativo'] = None
    st.query_params.clear()
    st.rerun()

def obter_usuario_atual():
    """Recarrega do banco os dados mais recentes do usuário e da OM (logo, cor e lema)."""
    user = st.session_state.get('usuario')
    if not user:
        return None

    conn = get_connection()
    c = conn.cursor()
    c.execute("""
    SELECT u.*, o.sigla as om_sigla, o.logo_base64, o.cor_primaria, o.lema 
    FROM usuarios u JOIN oms o ON u.om_id = o.id WHERE u.id = ?
    """, (user['id'],))
    user_atualizado = c.fetchone()
    conn.close()

    if user_atualizado:
        st.session_state['usuario'] = dict(user_atualizado)
    return st.session_state['usuario']

def render_sidebar_usuario(user):
    """Renderiza o brasão da OM, crachá do militar e o seletor de perfil ativo."""
    logo_b64 = user.get('logo_base64')
    lema_om = user.get('lema') or ''
    cor_om = user.get('cor_primaria') or '#1B4332'

    # 1. Brasão da OM
    if logo_b64:
        st.sidebar.markdown(f'''
        <div style="text-align: center; margin-bottom: 8px;">
            <img src="data:image/png;base64,{logo_b64}" style="height: 80px; width: auto; max-width: 60px; object-fit: cover; filter: drop-shadow(0px 3px 6px rgba(0,0,0,0.5));">
            <div style="font-size: 10px; font-style: italic; color: #C5A059; margin-top: 4px; font-weight: 600;">{lema_om}</div>
        </div>
        ''', unsafe_allow_html=True)
    else:
        st.sidebar.markdown(f'''
        <div style="margin: 0 auto 8px auto; width: 55px; height: 78px; background: linear-gradient(180deg, #ffffff 0%, #f0f0f0 100%); border: 1.5px solid #C5A059; border-radius: 2px 2px 28% 28%; display: flex; flex-direction: column; align-items: center; justify-content: space-between; padding: 0 0 3px 0;">
            <div style="background: {cor_om}; color: white; font-weight: 900; font-size: 7px; width: 100%; text-align: center; padding: 2px 0;">{user['om_sigla']}</div>
            <div style="font-size: 20px; margin: auto 0;">⚔️</div>
            <div style="font-size: 6px; color: #555;">{lema_om if lema_om else 'EB'}</div>
        </div>
        ''', unsafe_allow_html=True)

    # 2. Crachá Tático do Militar
    st.sidebar.markdown(f'''
    <div style="background: rgba(255,255,255,0.06); border: 1px solid rgba(197,160,89,0.35); border-radius: 8px; padding: 10px 14px; margin-bottom: 12px; box-shadow: 0 2px 6px rgba(0,0,0,0.3);">
        <div style="font-size: 11px; font-weight: 700; color: #C5A059; text-transform: uppercase;">Militar Conectado</div>
        <div style="font-size: 16px; font-weight: 900; color: #ffffff; margin-top: 2px;">{user.get('posto_grad', '')} {user.get('nome_guerra', '')}</div>
    </div>
    ''', unsafe_allow_html=True)

    # 3. Gestão de Perfis de Acesso Autorizados
    perfis_autorizados = [p.strip() for p in user.get('perfil', '').split(",") if p.strip()]
    if not perfis_autorizados:
        perfis_autorizados = ['SALC']

    if 'perfil_ativo' not in st.session_state or st.session_state['perfil_ativo'] not in perfis_autorizados:
        st.session_state['perfil_ativo'] = perfis_autorizados[0]

    if len(perfis_autorizados) > 1:
        idx_ativo = perfis_autorizados.index(st.session_state['perfil_ativo']) if st.session_state['perfil_ativo'] in perfis_autorizados else 0
        perfil_selecionado = st.sidebar.selectbox("🔄 Perfil Ativo:", perfis_autorizados, index=idx_ativo, key="sb_perfil_escolhido")
        if perfil_selecionado != st.session_state['perfil_ativo']:
            st.session_state['perfil_ativo'] = perfil_selecionado
            st.rerun()
    else:
        st.sidebar.markdown(f"**Perfil:** `{perfis_autorizados[0]}`")

    return st.session_state.get('perfil_ativo', perfis_autorizados[0])

def obter_menus_autorizados(perfil_ativo):
    """Retorna os menus de navegação liberados conforme o perfil ativo do militar."""
    menu_opcoes = ["📊 Dashboard"]

    if perfil_ativo in ['Administrador', 'Cmt OM', 'Ch 4ª Seção', 'SALC']:
        menu_opcoes.extend(["📑 Notas de Crédito", "📋 Notas de Empenho", "📝 Elaboração de Processos"])

    if perfil_ativo in ['Administrador', 'Cmt OM', 'Ch 4ª Seção', 'SALC', 'Almoxarife', 'Op Almoxarifado']:
        menu_opcoes.append("📦 Recebimento")

    if perfil_ativo in ['Administrador', 'Cmt OM', 'Ch 4ª Seção', 'Almoxarife', 'Op Almoxarifado']:
        menu_opcoes.append("🏢 Gestão de Estoque")

    menu_opcoes.append("📁 Relatórios")

    if perfil_ativo == 'Administrador':
        menu_opcoes.append("⚙️ Admin")

    return menu_opcoes

if __name__ == "__main__":
    print("Módulo core/auth.py compilado com sucesso!")
