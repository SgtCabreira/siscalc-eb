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
    st.title("📋 Notas de Empenho da OM")
    conn = get_connection()

    # Migração automática de colunas para notas de empenho divididas
    cur_mig_ne = conn.cursor()
    for col, col_def in [
        ("fornecedor_email", "TEXT"),
        ("fornecedor_telefone", "TEXT"),
        ("fornecedor_cidade", "TEXT"),
        ("fornecedor_uf", "TEXT"),
        ("fornecedor_situacao", "TEXT"),
        ("nc_id_2", "INTEGER"),
        ("valor_nc_1", "REAL"),
        ("valor_nc_2", "REAL DEFAULT 0.0")
    ]:
        try:
            cur_mig_ne.execute(f"ALTER TABLE notas_empenho ADD COLUMN {col} {col_def}")
        except sqlite3.OperationalError:
            pass
    conn.commit()

    col_ne_top1, col_ne_top2, col_ne_top3 = st.columns(3)

    with col_ne_top1:
        with st.expander("➕ Cadastrar Nova Nota de Empenho", expanded=False):
            df_nc_opcoes = pd.read_sql_query(f'''
            SELECT id, numero_nc, valor_total,
                   (valor_total - COALESCE(valor_recolhido, 0.0) - 
                    (COALESCE((SELECT SUM(COALESCE(valor_nc_1, valor_ne)) FROM notas_empenho WHERE nc_id = notas_credito.id), 0.0) +
                     COALESCE((SELECT SUM(COALESCE(valor_nc_2, 0.0)) FROM notas_empenho WHERE nc_id_2 = notas_credito.id), 0.0))
                   ) as saldo
            FROM notas_credito WHERE om_id = {user['om_id']}
            ''', conn)

            if df_nc_opcoes.empty:
                st.warning("Cadastre primeiro uma Nota de Crédito com saldo disponível.")
            else:
                opcoes_nc_dict = {f"{row['numero_nc']} (Saldo Disponível: R$ {row['saldo']:,.2f})": row['id'] for _, row in df_nc_opcoes.iterrows()}

                col_sel_nc1, col_sel_nc2 = st.columns(2)
                with col_sel_nc1:
                    nc_selecionada = st.selectbox("1ª Nota de Crédito (Principal):", list(opcoes_nc_dict.keys()), key="nc_sel_empenho")
                    nc_id = opcoes_nc_dict[nc_selecionada]

                with col_sel_nc2:
                    opcoes_nc2_dict = {k: v for k, v in opcoes_nc_dict.items() if v != nc_id}
                    lista_nc2 = ["❌ Nenhuma (Empenho vinculado a apenas 1 NC)"] + list(opcoes_nc2_dict.keys())
                    nc2_selecionada = st.selectbox("2ª Nota de Crédito (Opcional - caso utilize duas NCs):", lista_nc2, key="nc_sel_empenho_2")
                    usa_segunda_nc = not nc2_selecionada.startswith("❌")
                    nc_id_2 = opcoes_nc2_dict[nc2_selecionada] if usa_segunda_nc else None

                st.markdown("##### 🔍 Busca Automática da Empresa na Receita Federal")
                col_b_cnpj, col_b_btn = st.columns([3, 1])
                cnpj_digitado = col_b_cnpj.text_input("Digite o CNPJ da Empresa:", key="cnpj_busca_ne_input")

                if col_b_btn.button("Buscar Empresa", key="btn_buscar_empresa_ne", use_container_width=True):
                    dados_emp = buscar_dados_cnpj(cnpj_digitado, conn)
                    if dados_emp:
                        st.session_state['ne_cnpj_auto'] = cnpj_digitado
                        st.session_state['ne_nome_auto'] = dados_emp.get('razao_social', '')
                        st.session_state['ne_email_auto'] = dados_emp.get('email', '')
                        st.session_state['ne_tel_auto'] = dados_emp.get('telefone', '')
                        st.session_state['ne_cidade_auto'] = dados_emp.get('cidade', '')
                        st.session_state['ne_uf_auto'] = dados_emp.get('uf', '')
                        st.session_state['ne_sit_auto'] = dados_emp.get('situacao', 'ATIVA')
                        st.success(f"✅ Empresa Localizada: **{st.session_state['ne_nome_auto']}**")
                    else:
                        st.warning("CNPJ não localizado na Receita Federal ou sem conexão no momento.")

                with st.form("form_ne_limpo"):
                    if usa_segunda_nc:
                        col1, col2 = st.columns(2)
                        numero_ne = col1.text_input("Número do Empenho (ex.: 2026NE000456)")
                        data_emissao_ne = col2.date_input("Data de Emissão da NE", value=datetime.now(), format="DD/MM/YYYY")

                        col_v1, col_v2, col_v3 = st.columns(3)
                        nc1_nom = nc_selecionada.split(" (Saldo:")[0]
                        nc2_nom = nc2_selecionada.split(" (Saldo:")[0]
                        val_nc1 = col_v1.number_input(f"Valor da 1ª NC ({nc1_nom}) - R$:", min_value=0.01, step=50.0, format="%.2f")
                        val_nc2 = col_v2.number_input(f"Valor da 2ª NC ({nc2_nom}) - R$:", min_value=0.01, step=50.0, format="%.2f")
                        valor_ne = val_nc1 + val_nc2
                        col_v3.markdown(f'''
                        <div style="background: rgba(255,255,255,0.05); border: 1px solid rgba(197,160,89,0.3); border-radius: 8px; padding: 8px 12px; margin-top: 15px;">
                            <div style="font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase;">Valor Total da NE</div>
                            <div style="font-size: 18px; font-weight: 900; color: #38bdf8;">R$ {valor_ne:,.2f}</div>
                        </div>
                        '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
                    else:
                        col1, col2, col3 = st.columns(3)
                        numero_ne = col1.text_input("Número do Empenho (ex.: 2026NE000456)")
                        data_emissao_ne = col2.date_input("Data de Emissão da NE", value=datetime.now(), format="DD/MM/YYYY")
                        valor_ne = col3.number_input("Valor da NE (R$)", min_value=0.01, step=50.0, format="%.2f")
                        val_nc1 = valor_ne
                        val_nc2 = 0.0

                    col4, col5, col6 = st.columns(3)
                    tipo_empenho = col4.selectbox("Tipo de Empenho", ["Ordinário", "Global", "Estimativo"])
                    fornecedor_nome = col5.text_input("Nome Empresarial (Razão Social)", value=st.session_state.get('ne_nome_auto', ''))
                    fornecedor_cnpj = col6.text_input("CNPJ do Fornecedor", value=st.session_state.get('ne_cnpj_auto', ''))

                    col7, col8 = st.columns(2)
                    data_envio = col7.date_input("Data de Recebimento pela Empresa (Marco Zero)", value=datetime.now(), format="DD/MM/YYYY")
                    prazo_dias = col8.number_input("Prazo de Entrega (Dias Corridos)", value=30, step=5)

                    data_limite_calc = data_envio + timedelta(days=prazo_dias)
                    st.info(f"📅 Data Limite Regulamentar Automática: **{data_limite_calc.strftime('%d/%m/%Y')}**")

                    prorrogado = st.checkbox("Houve Prorrogação de Prazo Solicitada?", key="chk_prorr_ne")
                    nova_data = None
                    justificativa = None
                    if prorrogado:
                        col_p1, col_p2 = st.columns(2)
                        nova_data = col_p1.date_input("Nova Data Limite de Entrega", value=data_limite_calc + timedelta(days=15), format="DD/MM/YYYY")
                        justificativa = col_p2.text_area("Justificativa da Empresa / Despacho")

                    if st.form_submit_button("Salvar Empenho", type="primary"):
                        if not fornecedor_cnpj:
                            st.error("O CNPJ do fornecedor é obrigatório.")
                        elif not numero_ne:
                            st.error("O número da Nota de Empenho é obrigatório.")
                        else:
                            c = conn.cursor()
                            email_salvar = st.session_state.get('ne_email_auto', '')
                            tel_salvar = st.session_state.get('ne_tel_auto', '')
                            cid_salvar = st.session_state.get('ne_cidade_auto', '')
                            uf_salvar = st.session_state.get('ne_uf_auto', '')
                            sit_salvar = st.session_state.get('ne_sit_auto', 'ATIVA')

                            c.execute('''
                            INSERT INTO notas_empenho (nc_id, nc_id_2, valor_nc_1, valor_nc_2, om_id, numero_ne, data_emissao, valor_ne, tipo_empenho,
                                                       fornecedor_nome, fornecedor_cnpj, data_envio_empresa, prazo_dias,
                                                       data_limite, prorrogado, nova_data_limite, justificativa_prorrogacao,
                                                       fornecedor_email, fornecedor_telefone, fornecedor_cidade, fornecedor_uf, fornecedor_situacao)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            ''', (nc_id, nc_id_2, val_nc1, val_nc2, user['om_id'], numero_ne, str(data_emissao_ne), valor_ne, tipo_empenho,
                                  fornecedor_nome, fornecedor_cnpj, str(data_envio), prazo_dias, str(data_limite_calc),
                                  1 if prorrogado else 0, str(nova_data) if nova_data else None, justificativa,
                                  email_salvar, tel_salvar, cid_salvar, uf_salvar, sit_salvar))
                            conn.commit()

                            for k in ['ne_cnpj_auto', 'ne_nome_auto', 'ne_email_auto', 'ne_tel_auto', 'ne_cidade_auto', 'ne_uf_auto', 'ne_sit_auto']:
                                if k in st.session_state:
                                    del st.session_state[k]

                            st.success(f"✅ Empenho {numero_ne} (R$ {valor_ne:,.2f}) cadastrado com sucesso!")
                            st.rerun()

    with col_ne_top2:
        with st.expander("✏️ Editar Nota de Empenho", expanded=False):
            df_nes_all_top = pd.read_sql_query(f"SELECT * FROM notas_empenho WHERE om_id = {user['om_id']} ORDER BY id DESC", conn)
            if not df_nes_all_top.empty:
                dict_edit_ne = {f"{r['numero_ne']} - {r['fornecedor_nome']} (R$ {r['valor_ne']:,.2f})": r['id'] for _, r in df_nes_all_top.iterrows()}
                sel_ne_ed = st.selectbox("Selecione a NE para editar:", list(dict_edit_ne.keys()), key="sel_ne_to_edit_top")
                id_ne_to_edit = dict_edit_ne[sel_ne_ed]
                ne_dados_atual = dict(df_nes_all_top[df_nes_all_top['id'] == id_ne_to_edit].iloc[0])

                with st.form(f"form_ed_ne_top_{id_ne_to_edit}"):
                    df_ncs_vinc_opts = pd.read_sql_query(f"SELECT id, numero_nc FROM notas_credito WHERE om_id = {user['om_id']}", conn)
                    dict_ncs_v = {r['numero_nc']: r['id'] for _, r in df_ncs_vinc_opts.iterrows()}
                    nc_atual_num = next((k for k, v in dict_ncs_v.items() if v == ne_dados_atual['nc_id']), list(dict_ncs_v.keys())[0] if dict_ncs_v else "")
                    idx_nc_v = list(dict_ncs_v.keys()).index(nc_atual_num) if nc_atual_num in dict_ncs_v else 0

                    col_ed_nc1, col_ed_nc2 = st.columns(2)
                    with col_ed_nc1:
                        ed_nc_vinculada = st.selectbox("1ª Nota de Crédito (Principal):", list(dict_ncs_v.keys()), index=idx_nc_v, key=f"top_ed_nc1_{id_ne_to_edit}")
                        novo_nc_id = dict_ncs_v[ed_nc_vinculada]

                    with col_ed_nc2:
                        dict_ncs_v2 = {k: v for k, v in dict_ncs_v.items() if v != novo_nc_id}
                        lista_v2 = ["❌ Nenhuma (Apenas 1 NC)"] + list(dict_ncs_v2.keys())
                        nc2_atual_num = next((k for k, v in dict_ncs_v2.items() if v == ne_dados_atual.get('nc_id_2')), lista_v2[0])
                        idx_nc_v2 = lista_v2.index(nc2_atual_num) if nc2_atual_num in lista_v2 else 0
                        ed_nc_vinculada_2 = st.selectbox("2ª Nota de Crédito (Opcional):", lista_v2, index=idx_nc_v2, key=f"top_ed_nc2_{id_ne_to_edit}")
                        novo_nc_id_2 = dict_ncs_v2[ed_nc_vinculada_2] if not ed_nc_vinculada_2.startswith("❌") else None

                    c_ne1, c_ne2, c_ne3 = st.columns(3)
                    ed_num_ne = c_ne1.text_input("Número do Empenho (NE):", value=ne_dados_atual['numero_ne'])
                    ed_forn_nome = c_ne2.text_input("Razão Social do Fornecedor:", value=ne_dados_atual['fornecedor_nome'])
                    ed_forn_cnpj = c_ne3.text_input("CNPJ do Fornecedor:", value=ne_dados_atual['fornecedor_cnpj'])

                    if novo_nc_id_2:
                        c_ev1, c_ev2 = st.columns(2)
                        val1_init = float(ne_dados_atual.get('valor_nc_1') or ne_dados_atual['valor_ne'])
                        val2_init = float(ne_dados_atual.get('valor_nc_2') or 0.0)
                        ed_val_nc1 = c_ev1.number_input("Valor retirado da 1ª NC (R$):", value=val1_init, step=50.0, format="%.2f", key=f"top_v1_{id_ne_to_edit}")
                        ed_val_nc2 = c_ev2.number_input("Valor retirado da 2ª NC (R$):", value=val2_init, step=50.0, format="%.2f", key=f"top_v2_{id_ne_to_edit}")
                        ed_val_ne = ed_val_nc1 + ed_val_nc2
                        st.info(f"💰 Valor Total Atualizado da NE: **R$ {ed_val_ne:,.2f}**")
                    else:
                        c_ev1, = st.columns(1)
                        ed_val_ne = c_ev1.number_input("Valor Total da NE (R$):", value=float(ne_dados_atual['valor_ne']), step=50.0, format="%.2f", key=f"top_vtot_{id_ne_to_edit}")
                        ed_val_nc1 = ed_val_ne
                        ed_val_nc2 = 0.0

                    c_ne4, c_ne5, c_ne6 = st.columns(3)
                    tipos_emp_lista = ["Ordinário", "Global", "Estimativo"]
                    idx_tipo_e = tipos_emp_lista.index(ne_dados_atual['tipo_empenho']) if ne_dados_atual['tipo_empenho'] in tipos_emp_lista else 0
                    ed_tipo_emp = c_ne5.selectbox("Tipo de Empenho:", tipos_emp_lista, index=idx_tipo_e)
                    status_lista = ["Aguardando Entrega", "Liquidado", "Pago"]
                    idx_st_e = status_lista.index(ne_dados_atual['status']) if ne_dados_atual['status'] in status_lista else 0
                    ed_status_ne = c_ne6.selectbox("Status:", status_lista, index=idx_st_e)

                    c_ne7, c_ne8, c_ne9 = st.columns(3)
                    dt_em_ne = pd.to_datetime(ne_dados_atual['data_emissao']).date() if pd.notna(ne_dados_atual['data_emissao']) else datetime.now().date()
                    ed_dt_emissao_ne = c_ne7.date_input("Data de Emissão da NE:", value=dt_em_ne, format="DD/MM/YYYY")

                    dt_env_ne = pd.to_datetime(ne_dados_atual['data_envio_empresa']).date() if pd.notna(ne_dados_atual['data_envio_empresa']) else datetime.now().date()
                    ed_dt_envio_ne = c_ne8.date_input("Data de Envio (Marco Zero):", value=dt_env_ne, format="DD/MM/YYYY")

                    ed_prazo_dias = c_ne9.number_input("Prazo de Entrega (Dias):", value=int(ne_dados_atual['prazo_dias'] or 30), step=5)
                    nova_data_calc = ed_dt_envio_ne + timedelta(days=ed_prazo_dias)

                    c_ne10, c_ne11 = st.columns(2)
                    ed_forn_email = c_ne10.text_input("E-mail Oficial da Empresa:", value=ne_dados_atual['fornecedor_email'] or "")
                    ed_forn_tel = c_ne11.text_input("Telefone da Empresa:", value=ne_dados_atual['fornecedor_telefone'] or "")

                    ed_prorrogado = st.checkbox("Houve Prorrogação de Prazo?", value=bool(ne_dados_atual['prorrogado']), key=f"chk_ed_prorr_top_{id_ne_to_edit}")
                    ed_nova_data_limite = None
                    ed_just_prorr = None
                    if ed_prorrogado:
                        cp1, cp2 = st.columns(2)
                        dt_prorr_val = pd.to_datetime(ne_dados_atual['nova_data_limite']).date() if pd.notna(ne_dados_atual['nova_data_limite']) else (nova_data_calc + timedelta(days=15))
                        ed_nova_data_limite = cp1.date_input("Nova Data Limite de Entrega:", value=dt_prorr_val, format="DD/MM/YYYY")
                        ed_just_prorr = cp2.text_area("Justificativa da Prorrogação:", value=ne_dados_atual['justificativa_prorrogacao'] or "")

                    if st.form_submit_button("💾 Salvar Todas as Alterações da NE", type="primary"):
                        c = conn.cursor()
                        c.execute('''
                        UPDATE notas_empenho
                        SET nc_id = ?, nc_id_2 = ?, valor_nc_1 = ?, valor_nc_2 = ?, numero_ne = ?, fornecedor_nome = ?, fornecedor_cnpj = ?,
                            valor_ne = ?, tipo_empenho = ?, status = ?, data_emissao = ?,
                            data_envio_empresa = ?, prazo_dias = ?, data_limite = ?,
                            prorrogado = ?, nova_data_limite = ?, justificativa_prorrogacao = ?,
                            fornecedor_email = ?, fornecedor_telefone = ?
                        WHERE id = ?
                        ''', (novo_nc_id, novo_nc_id_2, ed_val_nc1, ed_val_nc2, ed_num_ne.strip(), ed_forn_nome.strip(), ed_forn_cnpj.strip(),
                              ed_val_ne, ed_tipo_emp, ed_status_ne, str(ed_dt_emissao_ne), str(ed_dt_envio_ne),
                              ed_prazo_dias, str(nova_data_calc), 1 if ed_prorrogado else 0,
                              str(ed_nova_data_limite) if ed_nova_data_limite else None,
                              ed_just_prorr, ed_forn_email.strip(), ed_forn_tel.strip(), id_ne_to_edit))
                        conn.commit()
                        st.success(f"✅ Nota de Empenho {ed_num_ne} atualizada com sucesso!")
                        st.rerun()
            else:
                st.info("Nenhuma NE cadastrada para editar.")

    with col_ne_top3:
        with st.expander("🗑️ Excluir Nota de Empenho", expanded=False):
            df_nes_para_del = pd.read_sql_query(f'''
            SELECT id, numero_ne, valor_ne, fornecedor_nome FROM notas_empenho 
            WHERE om_id = {user['om_id']}
            ORDER BY id DESC
            ''', conn)

            if not df_nes_para_del.empty:
                dict_del_ne = {
                    f"{row['numero_ne']} - {row['fornecedor_nome']} (R$ {row['valor_ne']:,.2f})": row['id']
                    for _, row in df_nes_para_del.iterrows()
                }
                ne_del_sel = st.selectbox("Selecione o Empenho para excluir:", list(dict_del_ne.keys()), key="sel_ne_del_box")
                id_ne_del = dict_del_ne[ne_del_sel]

                c = conn.cursor()
                c.execute("SELECT COUNT(*) FROM notas_fiscais WHERE ne_id = ?", (id_ne_del,))
                qtd_nfs_vinculadas = c.fetchone()[0]

                if qtd_nfs_vinculadas > 0:
                    st.warning(f"⚠️ Não é possível excluir este empenho porque ele possui {qtd_nfs_vinculadas} Nota(s) Fiscal(is) vinculadas no Almoxarifado.")
                else:
                    if st.button("Confirmar Exclusão Definitiva do Empenho", type="primary", key="btn_del_ne_confirm"):
                        c.execute("DELETE FROM notas_empenho WHERE id = ?", (id_ne_del,))
                        conn.commit()
                        st.success("🗑️ Nota de Empenho excluída com sucesso! O saldo retornou para a Nota de Crédito.")
                        st.rerun()
            else:
                st.info("Nenhum empenho cadastrado para exclusão.")

    col_t_ne, col_ord_ne, col_flt_ne = st.columns([6, 1, 1])
    with col_t_ne:
        st.markdown("### 🗂️ Painel Visual de Notas de Empenho")

    with col_ord_ne:
        with st.popover("🔀 Ordenar"):
            ordem_ne_escolhida = st.radio(
                "Critério de Ordenação:",
                [
                    "🚨 Mais urgente ao menos urgente (Padrão)",
                    "💰 Maior valor",
                    "💵 Menor valor",
                    "📅 Data de emissão mais recente"
                ],
                key="radio_ordem_ne"
            )

    with col_flt_ne:
        with st.popover("🌪️ Filtrar"):
            f_ne_atraso = st.checkbox("🔴 Atrasados", value=True, key="chk_ne_atraso")
            f_ne_prazo = st.checkbox("🟢 No Prazo de Entrega", value=True, key="chk_ne_prazo")
            f_ne_liq = st.checkbox("🟡 Liquidados", value=True, key="chk_ne_liq")
            f_ne_pago = st.checkbox("🔵 Pagos", value=True, key="chk_ne_pago")

    df_nes = pd.read_sql_query(f'''
    SELECT ne.id, 
           nc1.numero_nc as NC_Origem, 
           nc2.numero_nc as NC_Origem_2,
           ne.nc_id, ne.nc_id_2, ne.valor_nc_1, ne.valor_nc_2,
           ne.numero_ne, ne.tipo_empenho, 
           ne.fornecedor_nome, ne.fornecedor_cnpj, ne.valor_ne, ne.data_emissao, 
           COALESCE(ne.nova_data_limite, ne.data_limite) as data_final,
           ne.status
    FROM notas_empenho ne
    JOIN notas_credito nc1 ON ne.nc_id = nc1.id
    LEFT JOIN notas_credito nc2 ON ne.nc_id_2 = nc2.id
    WHERE ne.om_id = {user['om_id']}
    ''', conn)

    if not df_nes.empty:
        hoje = datetime.now().date()

        def classificar_ne(row):
            st_ne = row['status']
            if st_ne == 'Pago':
                return 4, "PAGO", "🔵 PAGO", "#1e3a8a", "border-left: 6px solid #3b82f6;"
            elif st_ne == 'Liquidado':
                return 3, "LIQUIDADO", "🟡 LIQUIDADO", "#854d0e", "border-left: 6px solid #eab308;"
            else:
                dias = None
                if pd.notna(row['data_final']):
                    try:
                        dias = (pd.to_datetime(row['data_final']).date() - hoje).days
                    except Exception:
                        pass
                if dias is not None and dias < 0:
                    return 0, "ATRASADO", f"🔴 ATRASADO ({abs(dias)}d)", "#7f1d1d", "border-left: 6px solid #ef4444;"
                else:
                    return 1, "NO_PRAZO", "🟢 NO PRAZO", "#14532d", "border-left: 6px solid #22c55e;"

        df_nes[['peso_urgencia', 'cat_status', 'badge_txt', 'badge_bg', 'borda_ne']] = df_nes.apply(classificar_ne, axis=1, result_type='expand')

        status_permitidos_ne = []
        if f_ne_atraso: status_permitidos_ne.append("ATRASADO")
        if f_ne_prazo: status_permitidos_ne.append("NO_PRAZO")
        if f_ne_liq: status_permitidos_ne.append("LIQUIDADO")
        if f_ne_pago: status_permitidos_ne.append("PAGO")

        df_nes_filtrado = df_nes[df_nes['cat_status'].isin(status_permitidos_ne)].copy()

        if "Mais urgente" in ordem_ne_escolhida:
            df_nes_filtrado = df_nes_filtrado.sort_values(by=['peso_urgencia', 'valor_ne'], ascending=[True, False])
        elif "Maior valor" in ordem_ne_escolhida:
            df_nes_filtrado = df_nes_filtrado.sort_values(by='valor_ne', ascending=False)
        elif "Menor valor" in ordem_ne_escolhida:
            df_nes_filtrado = df_nes_filtrado.sort_values(by='valor_ne', ascending=True)
        elif "mais recente" in ordem_ne_escolhida:
            df_nes_filtrado = df_nes_filtrado.sort_values(by='data_emissao', ascending=False)

        cols_ne = st.columns(3)
        for idx, row in df_nes_filtrado.reset_index(drop=True).iterrows():
            with cols_ne[idx % 3]:
                st.markdown(f'''
                <div style="background: linear-gradient(145deg, #243142, #1a232e); border: 1.5px solid rgba(255,255,255,0.18); {row['borda_ne']} border-radius: 10px; padding: 15px; margin-bottom: 10px; box-shadow: 0 6px 16px rgba(0,0,0,0.45);">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="font-weight: 900; font-size: 16px; color: #ffffff; letter-spacing: 0.5px;">{row['numero_ne']}</span>
                        <span style="font-size: 11px; padding: 3px 8px; border-radius: 4px; background: {row['badge_bg']}; color: #fff; font-weight: 800;">{row['badge_txt']}</span>
                    </div>
                    <div style="font-size: 14px; color: #f8fafc; margin: 8px 0; font-weight: 700; min-height: 38px; line-height: 1.35;">
                        {row['fornecedor_nome'][:60]}
                    </div>
                    <div style="display: flex; justify-content: space-between; font-size: 12.5px; border-top: 1px solid rgba(255,255,255,0.12); padding-top: 8px;">
                        <span style="color: #cbd5e1;">Valor: <b style="color: #38bdf8; font-size: 14px;">R$ {row['valor_ne']:,.2f}</b></span>
                        <span style="color: #cbd5e1;">Origem: <b style="color: #ffffff;">{f"{row['NC_Origem']} + {row['NC_Origem_2']}" if pd.notna(row['NC_Origem_2']) and row['NC_Origem_2'] else row['NC_Origem']}</b></span>
                    </div>
                </div>
                ''', unsafe_allow_html=True)
                with st.popover("🏢 Ficha da Empresa & Detalhes", key=f"pop_ne_{row['id']}", use_container_width=True):
                    c = conn.cursor()
                    c.execute('''
                    SELECT ne.*, nc1.numero_nc as NC_Origem, nc2.numero_nc as NC_Origem_2
                    FROM notas_empenho ne
                    JOIN notas_credito nc1 ON ne.nc_id = nc1.id
                    LEFT JOIN notas_credito nc2 ON ne.nc_id_2 = nc2.id
                    WHERE ne.id = ?
                    ''', (row['id'],))
                    ne_info = dict(c.fetchone())

                    st.markdown(f'''
                    <div style="border-bottom: 2px solid #C5A059; padding-bottom: 8px; margin-bottom: 12px;">
                        <div style="font-size: 11px; font-weight: 800; color: #C5A059; text-transform: uppercase; letter-spacing: 0.5px;">Dossiê do Empenho & Fornecedor</div>
                        <div style="font-size: 19px; font-weight: 900; color: #ffffff;">{ne_info['numero_ne']}</div>
                        <div style="font-size: 15px; font-weight: 700; color: #f8fafc; margin-top: 4px;">{ne_info['fornecedor_nome']}</div>
                        <div style="font-size: 13px; color: #cbd5e1;">
    CNPJ: <b>{ne_info['fornecedor_cnpj']}</b> | 
    {f"NCs Origem: <b style='color: #38bdf8;'>{ne_info['NC_Origem']} (R$ {ne_info['valor_nc_1'] or 0.0:,.2f})</b> + <b style='color: #fbbf24;'>{ne_info['NC_Origem_2']} (R$ {ne_info['valor_nc_2'] or 0.0:,.2f})</b>" if pd.notna(ne_info.get('NC_Origem_2')) and ne_info.get('NC_Origem_2') else f"NC Origem: <b style='color: #38bdf8;'>{ne_info['NC_Origem']}</b>"}
    </div>
                    </div>
                    ''', unsafe_allow_html=True)

                    st.markdown(f'''
                    <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin: 10px 0;">
                        <div style="background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08); border-left: 3px solid #38bdf8; padding: 7px 10px; border-radius: 6px;">
                            <div style="font-size: 10px; color: #94a3b8; font-weight: 700; text-transform: uppercase;">Valor do Empenho</div>
                            <div style="font-size: 15px; font-weight: 800; color: #38bdf8; white-space: nowrap; margin-top: 2px;">R$ {ne_info['valor_ne']:,.2f}</div>
                        </div>
                        <div style="background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08); border-left: 3px solid #ffffff; padding: 7px 10px; border-radius: 6px;">
                            <div style="font-size: 10px; color: #94a3b8; font-weight: 700; text-transform: uppercase;">Data de Emissão</div>
                            <div style="font-size: 14px; font-weight: 800; color: #ffffff; white-space: nowrap; margin-top: 2px;">{formatar_data_br(ne_info['data_emissao'])}</div>
                        </div>
                        <div style="background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08); border-left: 3px solid #fbbf24; padding: 7px 10px; border-radius: 6px;">
                            <div style="font-size: 10px; color: #94a3b8; font-weight: 700; text-transform: uppercase;">Limite de Entrega</div>
                            <div style="font-size: 14px; font-weight: 800; color: #fbbf24; white-space: nowrap; margin-top: 2px;">{formatar_data_br(ne_info['data_limite'])}</div>
                        </div>
                    </div>
                    '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)

                    st.markdown("---")
                    st.markdown(f"📧 **E-mail Oficial:** [{ne_info['fornecedor_email']}](mailto:{ne_info['fornecedor_email']})" if ne_info['fornecedor_email'] else "📧 **E-mail Oficial:** *Não informado*")
                    st.markdown(f"📞 **Telefone:** `{ne_info['fornecedor_telefone']}`" if ne_info['fornecedor_telefone'] else "📞 **Telefone:** *Não informado*")

                    with st.expander("✏️ Editar Todos os Dados desta NE", expanded=False):
                        with st.form(f"form_ed_ne_card_{row['id']}"):
                            # Inicialização defensiva obrigatória (garante que nenhuma variável fique indefinida)
                            novo_nc_id_card = ne_info.get('nc_id')
                            novo_nc_id_card_2 = None
                            ed_val = float(ne_info.get('valor_ne') or 0.0)
                            ed_val_nc1_c = float(ne_info.get('valor_nc_1') or ed_val)
                            ed_val_nc2_c = float(ne_info.get('valor_nc_2') or 0.0)

                            df_ncs_vinc_card = pd.read_sql_query(f"SELECT id, numero_nc FROM notas_credito WHERE om_id = {user['om_id']}", conn)
                            dict_ncs_card = {r['numero_nc']: r['id'] for _, r in df_ncs_vinc_card.iterrows()}

                            if dict_ncs_card:
                                nc_card_atual_num = next((k for k, v in dict_ncs_card.items() if v == ne_info.get('nc_id')), list(dict_ncs_card.keys())[0])
                                idx_nc_c = list(dict_ncs_card.keys()).index(nc_card_atual_num) if nc_card_atual_num in dict_ncs_card else 0

                                col_c_nc1, col_c_nc2 = st.columns(2)
                                with col_c_nc1:
                                    c_ed_nc = st.selectbox("1ª Nota de Crédito (Principal):", list(dict_ncs_card.keys()), index=idx_nc_c, key=f"c_ed_nc_{row['id']}")
                                    novo_nc_id_card = dict_ncs_card.get(c_ed_nc)

                                with col_c_nc2:
                                    dict_ncs_c2 = {k: v for k, v in dict_ncs_card.items() if v != novo_nc_id_card}
                                    lista_c2 = ["❌ Nenhuma (Apenas 1 NC)"] + list(dict_ncs_c2.keys())
                                    nc2_c_atual_num = next((k for k, v in dict_ncs_c2.items() if v == ne_info.get('nc_id_2')), lista_c2[0])
                                    idx_nc_c2 = lista_c2.index(nc2_c_atual_num) if nc2_c_atual_num in lista_c2 else 0
                                    c_ed_nc_2 = st.selectbox("2ª Nota de Crédito (Opcional):", lista_c2, index=idx_nc_c2, key=f"c_ed_nc2_{row['id']}")
                                    novo_nc_id_card_2 = dict_ncs_c2.get(c_ed_nc_2) if (c_ed_nc_2 and not c_ed_nc_2.startswith("❌")) else None
                            else:
                                st.warning("Nenhuma Nota de Crédito cadastrada na OM.")

                            ed_c1, ed_c2, ed_c3 = st.columns(3)
                            ed_num = ed_c1.text_input("Número NE:", value=ne_info['numero_ne'], key=f"ed_num_{row['id']}")
                            ed_nome = ed_c2.text_input("Fornecedor:", value=ne_info['fornecedor_nome'], key=f"ed_nome_{row['id']}")
                            ed_cnpj = ed_c3.text_input("CNPJ:", value=ne_info['fornecedor_cnpj'], key=f"ed_cnpj_{row['id']}")

                            if novo_nc_id_card_2:
                                c_ev1_c, c_ev2_c = st.columns(2)
                                val1_c_init = float(ne_info.get('valor_nc_1') or ne_info['valor_ne'])
                                val2_c_init = float(ne_info.get('valor_nc_2') or 0.0)
                                ed_val_nc1_c = c_ev1_c.number_input("Valor retirado da 1ª NC (R$):", value=val1_c_init, step=50.0, format="%.2f", key=f"ed_v1_c_{row['id']}")
                                ed_val_nc2_c = c_ev2_c.number_input("Valor retirado da 2ª NC (R$):", value=val2_c_init, step=50.0, format="%.2f", key=f"ed_v2_c_{row['id']}")
                                ed_val = ed_val_nc1_c + ed_val_nc2_c
                                st.info(f"💰 Valor Total Atualizado da NE: **R$ {ed_val:,.2f}**")
                            else:
                                c_ev1_c = st.columns(1)[0]
                                ed_val = c_ev1_c.number_input("Valor Total da NE (R$):", value=float(ne_info['valor_ne']), step=50.0, format="%.2f", key=f"ed_val_{row['id']}")
                                ed_val_nc1_c = ed_val
                                ed_val_nc2_c = 0.0

                            ed_c5, ed_c6 = st.columns(2)
                            tipos_e = ["Ordinário", "Global", "Estimativo"]
                            ed_tipo = ed_c5.selectbox("Tipo:", tipos_e, index=tipos_e.index(ne_info['tipo_empenho']) if ne_info['tipo_empenho'] in tipos_e else 0, key=f"ed_tipo_{row['id']}")
                            sts_e = ["Aguardando Entrega", "Liquidado", "Pago"]
                            ed_st = ed_c6.selectbox("Status:", sts_e, index=sts_e.index(ne_info['status']) if ne_info['status'] in sts_e else 0, key=f"ed_st_{row['id']}")

                            ed_c7, ed_c8, ed_c9 = st.columns(3)
                            dt_em_c = pd.to_datetime(ne_info['data_emissao']).date() if pd.notna(ne_info['data_emissao']) else datetime.now().date()
                            ed_dt_em = ed_c7.date_input("Data de Emissão:", value=dt_em_c, format="DD/MM/YYYY", key=f"ed_dtem_{row['id']}")

                            dt_env_c = pd.to_datetime(ne_info['data_envio_empresa']).date() if pd.notna(ne_info['data_envio_empresa']) else datetime.now().date()
                            ed_dt_env = ed_c8.date_input("Data Envio (Marco Zero):", value=dt_env_c, format="DD/MM/YYYY", key=f"ed_dtenv_{row['id']}")

                            ed_pz = ed_c9.number_input("Prazo de Entrega (Dias):", value=int(ne_info['prazo_dias'] or 30), step=5, key=f"ed_pz_{row['id']}")
                            calc_lim_card = ed_dt_env + timedelta(days=ed_pz)

                            ed_email = st.text_input("E-mail Oficial:", value=ne_info['fornecedor_email'] or "", key=f"ed_em_{row['id']}")
                            ed_tel = st.text_input("Telefone de Contato:", value=ne_info['fornecedor_telefone'] or "", key=f"ed_tl_{row['id']}")

                            ed_prorr_c = st.checkbox("Prorrogação de Prazo Solicitada?", value=bool(ne_info['prorrogado']), key=f"chk_prorr_card_{row['id']}")
                            nova_dt_c = None
                            just_c = None
                            if ed_prorr_c:
                                cp1, cp2 = st.columns(2)
                                dt_prorr_def = pd.to_datetime(ne_info['nova_data_limite']).date() if pd.notna(ne_info['nova_data_limite']) else (calc_lim_card + timedelta(days=15))
                                nova_dt_c = cp1.date_input("Nova Data Limite:", value=dt_prorr_def, format="DD/MM/YYYY", key=f"novadt_{row['id']}")
                                just_c = cp2.text_area("Justificativa:", value=ne_info['justificativa_prorrogacao'] or "", key=f"just_{row['id']}")

                            if st.form_submit_button("💾 Salvar Todas as Alterações da NE", type="primary"):
                                c.execute('''
                                UPDATE notas_empenho
                                SET nc_id = ?, nc_id_2 = ?, valor_nc_1 = ?, valor_nc_2 = ?, numero_ne = ?, fornecedor_nome = ?, fornecedor_cnpj = ?, valor_ne = ?,
                                    tipo_empenho = ?, status = ?, data_emissao = ?, data_envio_empresa = ?,
                                    prazo_dias = ?, data_limite = ?, prorrogado = ?, nova_data_limite = ?,
                                    justificativa_prorrogacao = ?, fornecedor_email = ?, fornecedor_telefone = ?
                                WHERE id = ?
                                ''', (novo_nc_id_card, novo_nc_id_card_2, ed_val_nc1_c, ed_val_nc2_c, ed_num.strip(), ed_nome.strip(), ed_cnpj.strip(), ed_val,
                                      ed_tipo, ed_st, str(ed_dt_em), str(ed_dt_env), ed_pz, str(calc_lim_card),
                                      1 if ed_prorr_c else 0, str(nova_dt_c) if nova_dt_c else None, just_c,
                                      ed_email.strip(), ed_tel.strip(), row['id']))
                                conn.commit()
                                st.success("✅ Nota de Empenho atualizada com sucesso!")
                                st.rerun()
    conn.close()
