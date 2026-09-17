import streamlit as st
import sqlite3
import pandas as pd
import io
import base64
from datetime import datetime, timedelta
from core.database import get_connection
from core.documentos import formatar_data_br, normalizar_data, exportar_excel_seguro, gerar_pdf_tabela_seguro
from core.api_cnpj import buscar_dados_cnpj
from core.gestor_fluxo import GestorFluxoMilitar


def render(user):
    st.title("⚙️ Painel de Administração - Forte Pinheirinho")
    conn = get_connection()

    col_om, col_user = st.columns(2)

    with col_om:
        st.subheader("🏢 Gestão de Organizações Militares")
        tab_om_cad, tab_om_edit, tab_om_theme, tab_om_del = st.tabs(["➕ Nova OM", "✏️ Editar OM", "🎨 Identidade Visual", "🗑️ Excluir OM"])

        with tab_om_cad:
            with st.form("form_om"):
                sigla = st.text_input("Sigla da OM (ex.: 5ª Cia PE)")
                nome = st.text_input("Nome da OM")
                ug = st.text_input("Código de UG")
                if st.form_submit_button("Cadastrar OM"):
                    c = conn.cursor()
                    try:
                        c.execute("INSERT INTO oms (sigla, nome, ug) VALUES (?, ?, ?)", (sigla, nome, ug))
                        conn.commit()
                        st.success(f"OM {sigla} cadastrada!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("OM já cadastrada.")

        with tab_om_edit:
            df_oms = pd.read_sql_query("SELECT * FROM oms", conn)
            if not df_oms.empty:
                dict_om_ed = {f"{row['sigla']} - {row['nome']} (UG: {row['ug']})": row['id'] for _, row in df_oms.iterrows()}
                om_sel_edit = st.selectbox("Selecione a OM para editar:", list(dict_om_ed.keys()), key="sel_om_edit")
                id_om_edit = dict_om_ed[om_sel_edit]

                c = conn.cursor()
                c.execute("SELECT * FROM oms WHERE id = ?", (id_om_edit,))
                dados_om = dict(c.fetchone())

                with st.form("form_edit_om"):
                    nova_sigla = st.text_input("Sigla da OM:", value=dados_om['sigla'], key="ed_sigla")
                    novo_nome_om = st.text_input("Nome da OM:", value=dados_om['nome'], key="ed_nome_om")
                    nova_ug = st.text_input("Código de UG:", value=dados_om['ug'] if dados_om['ug'] else "", key="ed_ug")

                    if st.form_submit_button("Salvar Alterações da OM"):
                        try:
                            c.execute("UPDATE oms SET sigla = ?, nome = ?, ug = ? WHERE id = ?",
                                      (nova_sigla, novo_nome_om, nova_ug, id_om_edit))
                            conn.commit()
                            st.success(f"OM '{nova_sigla}' atualizada com sucesso!")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("Já existe outra OM cadastrada com essa sigla.")
            else:
                st.info("Nenhuma OM cadastrada.")

        with tab_om_theme:
            df_oms = pd.read_sql_query("SELECT * FROM oms", conn)
            if not df_oms.empty:
                dict_om_th = {f"{row['sigla']} - {row['nome']}": row['id'] for _, row in df_oms.iterrows()}
                om_th_sel = st.selectbox("Selecione a OM para personalizar:", list(dict_om_th.keys()), key="sel_om_th")
                id_om_th = dict_om_th[om_th_sel]

                c = conn.cursor()
                c.execute("SELECT * FROM oms WHERE id = ?", (id_om_th,))
                dados_om_th = dict(c.fetchone())

                with st.form("form_theme_om"):
                    st.markdown(f"#### Personalização Visual: **{dados_om_th['sigla']}**")
                    arquivo_logo = st.file_uploader("Upload do Brasão / Distintivo da OM (PNG ou JPG):", type=["png", "jpg", "jpeg"])

                    col_c1, col_c2 = st.columns(2)
                    cor_atual = dados_om_th.get('cor_primaria') or '#1B4332'
                    nova_cor = col_c1.color_picker("Cor Primária da OM:", value=cor_atual)
                    col_c2.info("Dica: Verde-Oliva (#1B4332), Azul-Marinho PE (#0A2240), Cinza Blindado (#2C3E50), Vermelho (#800000)")

                    lema_atual = dados_om_th.get('lema') or ''
                    novo_lema = st.text_input("Lema ou Frase da OM (opcional):", value=lema_atual)

                    if st.form_submit_button("Salvar Identidade Visual"):
                        c = conn.cursor()
                        if arquivo_logo is not None:
                            logo_b64_new = base64.b64encode(arquivo_logo.read()).decode('utf-8')
                            c.execute("UPDATE oms SET cor_primaria = ?, lema = ?, logo_base64 = ? WHERE id = ?",
                                      (nova_cor, novo_lema, logo_b64_new, id_om_th))
                        else:
                            c.execute("UPDATE oms SET cor_primaria = ?, lema = ? WHERE id = ?",
                                      (nova_cor, novo_lema, id_om_th))
                        conn.commit()
                        st.success(f"Identidade visual da OM {dados_om_th['sigla']} salva com sucesso!")
                        st.rerun()
            else:
                st.info("Nenhuma OM cadastrada.")

        with tab_om_del:
            df_oms = pd.read_sql_query("SELECT * FROM oms", conn)
            if not df_oms.empty:
                om_para_excluir = st.selectbox("Selecione a OM para excluir:", df_oms['sigla'].tolist(), key="del_om_select")
                id_om_del = int(df_oms[df_oms['sigla'] == om_para_excluir]['id'].values[0])

                c = conn.cursor()
                c.execute("SELECT COUNT(*) FROM usuarios WHERE om_id = ?", (id_om_del,))
                qtd_u = c.fetchone()[0]
                c.execute("SELECT COUNT(*) FROM notas_credito WHERE om_id = ?", (id_om_del,))
                qtd_nc = c.fetchone()[0]

                if qtd_u > 0 or qtd_nc > 0:
                    st.warning(f"⚠️ Esta OM possui {qtd_u} usuário(s) e {qtd_nc} NC(s) vinculadas. Remova os vínculos antes de excluir.")
                else:
                    if st.button(f"Confirmar Exclusão de {om_para_excluir}", type="primary", key="btn_del_om"):
                        c.execute("DELETE FROM oms WHERE id = ?", (id_om_del,))
                        conn.commit()
                        st.success(f"OM '{om_para_excluir}' excluída com sucesso!")
                        st.rerun()
            else:
                st.info("Nenhuma OM cadastrada.")

        st.markdown("---")
        st.write("**OMs Cadastradas no Sistema:**")
        st.dataframe(pd.read_sql_query("SELECT id, sigla, nome, ug, cor_primaria, lema FROM oms", conn), use_container_width=True, hide_index=True)

    with col_user:
        st.subheader("👤 Gestão de Usuários e Senhas")
        df_oms_user = pd.read_sql_query("SELECT id, sigla FROM oms", conn)
        dict_oms = {row['sigla']: row['id'] for _, row in df_oms_user.iterrows()}
        dict_oms_rev = {row['id']: row['sigla'] for _, row in df_oms_user.iterrows()}

        df_usuarios = pd.read_sql_query('''
        SELECT u.id, o.sigla as OM, u.posto_grad, u.nome_guerra, u.perfil, u.identidade_militar, u.om_id 
        FROM usuarios u JOIN oms o ON u.om_id = o.id
        ''', conn)

        tab_u_cad, tab_u_edit, tab_u_del = st.tabs(["➕ Novo Usuário", "✏️ Editar Tudo / Senha", "🗑️ Excluir Usuário"])
        perfis_disponiveis_todos = ["Administrador", "Cmt OM", "Ch 4ª Seção", "SALC", "Almoxarife", "Op Almoxarifado"]

        with tab_u_cad:
            if not df_oms_user.empty:
                with st.form("form_u"):
                    om_u = st.selectbox("OM:", list(dict_oms.keys()), key="cad_user_om")
                    col_g, col_n = st.columns(2)
                    p_grad = col_g.selectbox("Posto/Graduação", ["Cel", "Ten Cel", "Maj", "Cap", "1º Ten", "2º Ten", "Asp", "Subten", "1º Sgt", "2º Sgt", "3º Sgt", "Cb", "Sd", "SC"], key="cad_user_p")
                    n_guerra = col_n.text_input("Nome de Guerra", key="cad_user_ng")
                    ident = st.text_input("Identidade Militar (Login)", key="cad_user_id")

                    perfis_escolhidos = st.multiselect(
                        "Perfis de Acesso Autorizados:",
                        perfis_disponiveis_todos,
                        default=["SALC"],
                        key="cad_user_perf_multi"
                    )

                    sen = st.text_input("Senha Inicial", type="password", value="1234", key="cad_user_sen")
                    if st.form_submit_button("Cadastrar Militar"):
                        if not perfis_escolhidos:
                            st.error("Selecione pelo menos um perfil de acesso para o militar.")
                        else:
                            c = conn.cursor()
                            perf_str = ", ".join(perfis_escolhidos)
                            try:
                                c.execute("INSERT INTO usuarios (nome_guerra, posto_grad, identidade_militar, om_id, perfil, senha) VALUES (?, ?, ?, ?, ?, ?)",
                                          (n_guerra, p_grad, ident, dict_oms[om_u], perf_str, sen))
                                conn.commit()
                                st.success(f"Usuário {n_guerra} criado com perfis: {perf_str}!")
                                st.rerun()
                            except sqlite3.IntegrityError:
                                st.error("Identidade Militar já cadastrada.")
            else:
                st.warning("Cadastre uma OM antes de criar usuários.")

        with tab_u_edit:
            if not df_usuarios.empty:
                opcoes_edit = {
                    f"{row['posto_grad']} {row['nome_guerra']} - OM: {row['OM']} (ID: {row['identidade_militar']})": row['id']
                    for _, row in df_usuarios.iterrows()
                }
                user_selecionado = st.selectbox("Selecione o militar para editar:", list(opcoes_edit.keys()), key="edit_user_sel")
                user_id_edit = opcoes_edit[user_selecionado]

                c = conn.cursor()
                c.execute("SELECT * FROM usuarios WHERE id = ?", (user_id_edit,))
                dados_u = dict(c.fetchone())

                postos_lista = ["Cel", "Ten Cel", "Maj", "Cap", "1º Ten", "2º Ten", "Asp", "Subten", "1º Sgt", "2º Sgt", "3º Sgt", "Cb", "Sd", "SC"]
                lista_siglas_om = list(dict_oms.keys())
                perfis_atuais_u = [p.strip() for p in dados_u['perfil'].split(",") if p.strip() and p.strip() in perfis_disponiveis_todos]

                with st.form("form_edit_user_tab"):
                    sigla_atual_om = dict_oms_rev.get(dados_u['om_id'], lista_siglas_om[0])
                    idx_om = lista_siglas_om.index(sigla_atual_om) if sigla_atual_om in lista_siglas_om else 0
                    nova_om_u = st.selectbox("OM de Lotação:", lista_siglas_om, index=idx_om, key="ed_om_u")

                    col_e1, col_e2 = st.columns(2)
                    posto_idx = postos_lista.index(dados_u['posto_grad']) if dados_u['posto_grad'] in postos_lista else 0
                    novo_posto = col_e1.selectbox("Posto/Graduação:", postos_lista, index=posto_idx, key="ed_p")
                    novo_nome = col_e2.text_input("Nome de Guerra:", value=dados_u['nome_guerra'], key="ed_ng")

                    col_e3, col_e4 = st.columns(2)
                    nova_ident = col_e3.text_input("Identidade Militar (Login):", value=dados_u['identidade_militar'], key="ed_id_m")

                    novos_perfis_sel = st.multiselect(
                        "Perfis de Acesso Autorizados:",
                        perfis_disponiveis_todos,
                        default=perfis_atuais_u,
                        key="ed_perf_multi"
                    )

                    st.markdown("🔑 **Redefinição de Senha:**")
                    nova_senha = st.text_input("Senha (altere ou mantenha a atual):", value=dados_u['senha'], key="ed_sen")

                    st.markdown("---")
                    confirmar_alteracao = st.checkbox("⚠️ Confirmo que conferi os dados e desejo salvar estas alterações.", key="chk_conf_u")

                    if st.form_submit_button("Salvar Todas as Alterações do Usuário"):
                        if not confirmar_alteracao:
                            st.warning("Por favor, marque a caixa de confirmação acima para autorizar as alterações.")
                        elif not novos_perfis_sel:
                            st.error("O militar deve possuir pelo menos um perfil de acesso selecionado.")
                        else:
                            try:
                                nova_perf_str = ", ".join(novos_perfis_sel)
                                c.execute('''
                                UPDATE usuarios 
                                SET posto_grad = ?, nome_guerra = ?, identidade_militar = ?, om_id = ?, perfil = ?, senha = ?
                                WHERE id = ?
                                ''', (novo_posto, novo_nome, nova_ident, dict_oms[nova_om_u], nova_perf_str, nova_senha, user_id_edit))
                                conn.commit()

                                if user_id_edit == user['id']:
                                    c.execute('''
                                    SELECT u.*, o.sigla as om_sigla, o.logo_base64, o.cor_primaria, o.lema FROM usuarios u 
                                    JOIN oms o ON u.om_id = o.id WHERE u.id = ?
                                    ''', (user_id_edit,))
                                    st.session_state['usuario'] = dict(c.fetchone())
                                    if st.session_state.get('perfil_ativo') not in novos_perfis_sel:
                                        st.session_state['perfil_ativo'] = novos_perfis_sel[0]

                                st.success(f"✅ Confirmação: Dados e perfis de {novo_posto} {novo_nome} atualizados com sucesso!")
                                st.rerun()
                            except sqlite3.IntegrityError:
                                st.error("Erro: Já existe outro usuário cadastrado com essa Identidade Militar.")
            else:
                st.info("Nenhum usuário cadastrado.")

        with tab_u_del:
            if not df_usuarios.empty:
                opcoes_usuarios = {
                    f"{row['posto_grad']} {row['nome_guerra']} - OM: {row['OM']} (ID: {row['identidade_militar']})": row['id']
                    for _, row in df_usuarios.iterrows()
                }
                user_para_excluir = st.selectbox("Selecione o usuário para excluir:", list(opcoes_usuarios.keys()), key="del_u_sel")
                id_user_del = opcoes_usuarios[user_para_excluir]

                if id_user_del == user['id']:
                    st.error("⛔ Você não pode excluir o seu próprio usuário enquanto estiver conectado ao sistema.")
                else:
                    if st.button("Confirmar Exclusão de Usuário", type="primary", key="btn_del_u"):
                        c = conn.cursor()
                        c.execute("DELETE FROM usuarios WHERE id = ?", (id_user_del,))
                        conn.commit()
                        st.success("Usuário excluído com sucesso!")
                        st.rerun()
            else:
                st.info("Nenhum usuário cadastrado.")

        st.markdown("---")
        st.write("**Militares Cadastrados no Sistema:**")
        st.dataframe(df_usuarios[['OM', 'posto_grad', 'nome_guerra', 'perfil', 'identidade_militar']], use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("💾 Backup e Restauração do Banco de Dados da OM")
    st.caption("🛡️ **Importante no Streamlit Cloud:** Como os servidores em nuvem podem reiniciar ou redefinir arquivos locais ao atualizar o código no GitHub, utilize esta ferramenta para baixar cópias de segurança (.db) e restaurar todos os seus dados a qualquer momento em 1 clique.")

    col_bk1, col_bk2 = st.columns(2)
    with col_bk1:
        st.markdown("##### 📥 Exportar / Baixar Banco Atual")
        import os
        if os.path.exists(DB_FILE):
            with open(DB_FILE, "rb") as f_db:
                bytes_db = f_db.read()
            st.download_button(
                label="📥 Baixar Cópia Completa do Banco (sistema_militar.db)",
                data=bytes_db,
                file_name=f"sistema_militar_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db",
                mime="application/x-sqlite3",
                use_container_width=True,
                help="Baixa o arquivo do banco com todas as suas NCs, NEs, estoque, materiais e usuários cadastrados."
            )
            st.success("✅ O banco de dados está online e pronto para download de backup.")

    with col_bk2:
        st.markdown("##### 📤 Restaurar / Importar Banco Salvo")
        up_db_file = st.file_uploader("Selecione um arquivo de backup (.db):", type=["db", "sqlite", "sqlite3"], key="up_sqlite_db_restore")
        if up_db_file is not None:
            if st.button("⚠️ Confirmar Restauração Completa do Banco de Dados", type="primary", use_container_width=True):
                with open(DB_FILE, "wb") as f_out:
                    f_out.write(up_db_file.read())
                st.success("✅ Banco de dados restaurado com sucesso! Atualizando sistema...")
                st.rerun()

    conn.close()
