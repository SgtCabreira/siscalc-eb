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
    st.title("📑 Notas de Crédito da OM")
    conn = get_connection()

    col_nc_top1, col_nc_top2, col_nc_top3 = st.columns(3)
    with col_nc_top1:
        with st.expander("➕ Cadastrar Nova Nota de Crédito", expanded=False):
            with st.form("form_nc_rapido"):
                c1, c2, c3 = st.columns(3)
                ug_emit = c1.text_input("UG Emitente", value="160220 - Cmt 5ª DE")
                ug_fav = c2.text_input("UG Favorecida", value=f"{user['om_sigla']}")
                num_nc = c3.text_input("Número da NC (ex.: 2026NC000123)")

                c4, c5, c6 = st.columns(3)
                dt_emissao = c4.date_input("Data de Emissão", value=datetime.now(), format="DD/MM/YYYY")
                dt_limite = c5.date_input("Data Limite do Empenho", value=datetime.now() + timedelta(days=30), format="DD/MM/YYYY")
                val_total = c6.number_input("Valor Total (R$)", min_value=0.01, step=100.0, format="%.2f")

                c7, c8, c9 = st.columns(3)
                nd = c7.text_input("Natureza de Despesa (ND)", value="339030 - Material de Consumo")
                pi = c8.text_input("Plano Interno (PI)", value="I2B0000000")
                enq = c9.selectbox("Enquadramento", ["DISCRICIONÁRIA", "ALIMENTAÇÃO", "FARDAMENTO", "SAÚDE", "ENGENHARIA", "OUTRO"])

                fin = st.text_input("Finalidade / Objeto (Sucinto)", value="Aquisição de material")
                obs = st.text_area("Observações")

                if st.form_submit_button("Salvar Nova NC", type="primary"):
                    cur = conn.cursor()
                    cur.execute('''
                    INSERT INTO notas_credito (om_id, ug_emitente, ug_favorecida, numero_nc, data_emissao, data_limite_empenho,
                                               valor_total, valor_recolhido, finalidade, natureza_despesa, pi, enquadramento, observacao)
                    VALUES (?, ?, ?, ?, ?, ?, ?, 0.0, ?, ?, ?, ?, ?)
                    ''', (user['om_id'], ug_emit, ug_fav, num_nc, str(dt_emissao), str(dt_limite), val_total,
                          fin, nd, pi, enq, obs))
                    conn.commit()
                    st.success(f"Nota de Crédito {num_nc} cadastrada com sucesso!")
                    st.rerun()

    with col_nc_top2:
        with st.expander("✏️ Editar Nota de Crédito", expanded=False):
            df_ncs_edit_top = pd.read_sql_query(f"SELECT * FROM notas_credito WHERE om_id = {user['om_id']} ORDER BY numero_nc", conn)
            if not df_ncs_edit_top.empty:
                dict_edit_nc = {f"{r['numero_nc']} - R$ {r['valor_total']:,.2f} | {r['finalidade'][:35]}...": r['id'] for _, r in df_ncs_edit_top.iterrows()}
                sel_nc_ed = st.selectbox("Selecione a NC para editar:", list(dict_edit_nc.keys()), key="sel_nc_to_edit_top")
                id_nc_to_edit = dict_edit_nc[sel_nc_ed]
                nc_dados_atual = dict(df_ncs_edit_top[df_ncs_edit_top['id'] == id_nc_to_edit].iloc[0])

                with st.form(f"form_ed_nc_top_{id_nc_to_edit}"):
                    ec1, ec2, ec3 = st.columns(3)
                    ed_num_nc = ec1.text_input("Número da NC:", value=nc_dados_atual['numero_nc'])
                    ed_ug_emit = ec2.text_input("UG Emitente:", value=nc_dados_atual['ug_emitente'] or "")
                    ed_ug_fav = ec3.text_input("UG Favorecida:", value=nc_dados_atual['ug_favorecida'] or "")

                    ec4, ec5, ec6 = st.columns(3)
                    dt_em_val = pd.to_datetime(nc_dados_atual['data_emissao']).date() if pd.notna(nc_dados_atual['data_emissao']) else datetime.now().date()
                    ed_dt_emissao = ec4.date_input("Data de Emissão:", value=dt_em_val, format="DD/MM/YYYY")

                    dt_lim_val = pd.to_datetime(nc_dados_atual['data_limite_empenho']).date() if pd.notna(nc_dados_atual['data_limite_empenho']) else (datetime.now().date() + timedelta(days=30))
                    ed_dt_limite = ec5.date_input("Data Limite do Empenho:", value=dt_lim_val, format="DD/MM/YYYY")

                    ed_val_total = ec6.number_input("Valor Total (R$):", value=float(nc_dados_atual['valor_total']), step=100.0, format="%.2f")

                    ec7, ec8, ec9 = st.columns(3)
                    ed_nd = ec7.text_input("Natureza de Despesa (ND):", value=nc_dados_atual['natureza_despesa'] or "")
                    ed_pi = ec8.text_input("Plano Interno (PI):", value=nc_dados_atual['pi'] or "")
                    enq_lista = ["DISCRICIONÁRIA", "ALIMENTAÇÃO", "FARDAMENTO", "SAÚDE", "ENGENHARIA", "OUTRO"]
                    idx_enq = enq_lista.index(nc_dados_atual['enquadramento']) if nc_dados_atual['enquadramento'] in enq_lista else 0
                    ed_enq = ec9.selectbox("Enquadramento:", enq_lista, index=idx_enq)

                    ed_fin = st.text_input("Finalidade / Objeto Completo:", value=nc_dados_atual['finalidade'] or "")
                    ed_rec = st.number_input("Valor Recolhido / Devolvido (R$):", value=float(nc_dados_atual['valor_recolhido'] or 0.0), step=10.0, format="%.2f")
                    ed_obs = st.text_area("Observações:", value=nc_dados_atual['observacao'] or "")

                    if st.form_submit_button("💾 Salvar Todas as Alterações da NC", type="primary"):
                        c = conn.cursor()
                        c.execute('''
                        UPDATE notas_credito
                        SET numero_nc = ?, finalidade = ?, valor_total = ?, valor_recolhido = ?,
                            data_emissao = ?, data_limite_empenho = ?, enquadramento = ?,
                            natureza_despesa = ?, pi = ?, ug_emitente = ?, ug_favorecida = ?, observacao = ?
                        WHERE id = ?
                        ''', (ed_num_nc.strip(), ed_fin.strip(), ed_val_total, ed_rec,
                              str(ed_dt_emissao), str(ed_dt_limite), ed_enq, ed_nd.strip(),
                              ed_pi.strip(), ed_ug_emit.strip(), ed_ug_fav.strip(), ed_obs.strip(), id_nc_to_edit))
                        conn.commit()
                        st.success(f"✅ Nota de Crédito {ed_num_nc} atualizada com sucesso!")
                        st.rerun()
            else:
                st.info("Nenhuma NC cadastrada para editar.")

    with col_nc_top3:
        with st.expander("🗑️ Excluir Nota de Crédito", expanded=False):
            df_ncs_del = pd.read_sql_query(f"SELECT id, numero_nc, valor_total, finalidade FROM notas_credito WHERE om_id = {user['om_id']}", conn)
            if not df_ncs_del.empty:
                dict_del = {f"{row['numero_nc']} - R$ {row['valor_total']:,.2f} | {row['finalidade'][:40]}...": row['id'] for _, row in df_ncs_del.iterrows()}
                nc_del_escolha = st.selectbox("Selecione a NC para excluir:", list(dict_del.keys()), key="nc_del_sel_quick")
                id_nc_del = dict_del[nc_del_escolha]

                cur = conn.cursor()
                cur.execute("SELECT COUNT(*) FROM notas_empenho WHERE nc_id = ? OR nc_id_2 = ?", (id_nc_del, id_nc_del))
                qtd_ne = cur.fetchone()[0]

                if qtd_ne > 0:
                    st.warning(f"⚠️ Não é possível excluir: existem {qtd_ne} Nota(s) de Empenho vinculadas a esta NC.")
                else:
                    if st.button("Confirmar Exclusão Definitiva", type="primary", key="btn_del_nc_exec"):
                        cur.execute("DELETE FROM notas_credito WHERE id = ?", (id_nc_del,))
                        conn.commit()
                        st.success("Nota de Crédito excluída com sucesso!")
                        st.rerun()
            else:
                st.info("Nenhuma NC cadastrada para exclusão.")

    # Controles de Ordenação e Filtro
    col_t_nc, col_ord_nc, col_flt_nc = st.columns([6, 1, 1])
    with col_t_nc:
        st.markdown("### 🗂️ Painel Visual de Notas de Crédito")

    with col_ord_nc:
        with st.popover("🔀 Ordenar"):
            ordem_nc_escolhida = st.radio(
                "Critério de Ordenação:",
                [
                    "🚨 Mais urgente ao menos urgente (Padrão)",
                    "💰 Maior valor",
                    "💵 Menor valor",
                    "📅 Data de emissão mais recente",
                    "⏳ Prazo limite mais próximo"
                ],
                key="radio_ordem_nc"
            )

    with col_flt_nc:
        with st.popover("🌪️ Filtrar"):
            f_venc = st.checkbox("🔴 Vencidos", value=True, key="chk_venc")
            f_urg = st.checkbox("🟠 Urgentes (≤ 2 dias)", value=True, key="chk_urg")
            f_aten = st.checkbox("🟡 Atenção (≤ 7 dias)", value=True, key="chk_aten")
            f_prazo = st.checkbox("🟢 No Prazo", value=True, key="chk_prazo")
            f_baix = st.checkbox("✅ Concluídas / Baixadas", value=True, key="chk_baix")

    query_cards_nc = f'''
    SELECT nc.id, nc.numero_nc, nc.data_emissao, nc.data_limite_empenho, nc.enquadramento, 
           nc.valor_total, COALESCE(nc.valor_recolhido, 0.0) as valor_recolhido,
           (COALESCE((SELECT SUM(COALESCE(valor_nc_1, valor_ne)) FROM notas_empenho WHERE nc_id = nc.id), 0.0) +
            COALESCE((SELECT SUM(COALESCE(valor_nc_2, 0.0)) FROM notas_empenho WHERE nc_id_2 = nc.id), 0.0)
           ) as total_empenhado,
           (nc.valor_total - COALESCE(nc.valor_recolhido, 0.0) - 
            (COALESCE((SELECT SUM(COALESCE(valor_nc_1, valor_ne)) FROM notas_empenho WHERE nc_id = nc.id), 0.0) +
             COALESCE((SELECT SUM(COALESCE(valor_nc_2, 0.0)) FROM notas_empenho WHERE nc_id_2 = nc.id), 0.0))
           ) as saldo_restante,
           nc.finalidade
    FROM notas_credito nc
    WHERE nc.om_id = {user['om_id']}
    GROUP BY nc.id
    '''
    df_ncs = pd.read_sql_query(query_cards_nc, conn)

    if not df_ncs.empty:
        hoje = datetime.now().date()
        df_ncs['saldo_calc'] = df_ncs['saldo_restante'].apply(lambda s: max(s, 0.0))
        df_ncs['data_limite_dt'] = pd.to_datetime(df_ncs['data_limite_empenho']).dt.date

        def atribuir_prioridade(row):
            if row['saldo_calc'] <= 0.01:
                return 5, "BAIXADA", "✅ CONCLUÍDA", "#334155", "border-left: 6px solid #64748b;"
            if pd.isna(row['data_limite_dt']):
                return 4, "SEM_PRAZO", "⚪ SEM PRAZO", "#334155", "border-left: 6px solid #64748b;"
            dias = (row['data_limite_dt'] - hoje).days
            if dias < 0:
                return 0, "VENCIDO", f"🔴 VENCIDO ({abs(dias)}d)", "#7f1d1d", "border-left: 6px solid #ef4444;"
            elif dias <= 2:
                return 1, "URGENTE", f"🟠 URGENTE ({dias}d)", "#9a3412", "border-left: 6px solid #f97316;"
            elif dias <= 7:
                return 2, "ATENCAO", f"🟡 ATENÇÃO ({dias}d)", "#854d0e", "border-left: 6px solid #eab308;"
            else:
                return 3, "NO_PRAZO", f"🟢 NO PRAZO ({dias}d)", "#14532d", "border-left: 6px solid #22c55e;"

        df_ncs[['peso_urgencia', 'cat_prazo', 'status_txt', 'status_bg', 'borda_cor']] = df_ncs.apply(atribuir_prioridade, axis=1, result_type='expand')

        categorias_permitidas = []
        if f_venc: categorias_permitidas.append("VENCIDO")
        if f_urg: categorias_permitidas.append("URGENTE")
        if f_aten: categorias_permitidas.append("ATENCAO")
        if f_prazo: categorias_permitidas.append("NO_PRAZO")
        if f_baix: categorias_permitidas.extend(["BAIXADA", "SEM_PRAZO"])

        df_filtrado = df_ncs[df_ncs['cat_prazo'].isin(categorias_permitidas)].copy()

        if "Mais urgente" in ordem_nc_escolhida:
            df_filtrado = df_filtrado.sort_values(by=['peso_urgencia', 'saldo_calc'], ascending=[True, False])
        elif "Maior valor" in ordem_nc_escolhida:
            df_filtrado = df_filtrado.sort_values(by='valor_total', ascending=False)
        elif "Menor valor" in ordem_nc_escolhida:
            df_filtrado = df_filtrado.sort_values(by='valor_total', ascending=True)
        elif "mais recente" in ordem_nc_escolhida:
            df_filtrado = df_filtrado.sort_values(by='data_emissao', ascending=False)
        elif "Prazo limite" in ordem_nc_escolhida:
            df_filtrado = df_filtrado.sort_values(by='data_limite_dt', ascending=True)

        cols = st.columns(3)
        for idx, row in df_filtrado.reset_index(drop=True).iterrows():
            with cols[idx % 3]:
                st.markdown(f'''
                <div style="background: linear-gradient(145deg, #243142, #1a232e); border: 1.5px solid rgba(255,255,255,0.18); {row['borda_cor']} border-radius: 10px; padding: 15px; margin-bottom: 10px; box-shadow: 0 6px 16px rgba(0,0,0,0.45);">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="font-weight: 900; font-size: 16px; color: #ffffff; letter-spacing: 0.5px;">{row['numero_nc']}</span>
                        <span style="font-size: 11px; padding: 3px 8px; border-radius: 4px; background: {row['status_bg']}; color: #fff; font-weight: 800;">{row['status_txt']}</span>
                    </div>
                    <div style="font-size: 13.5px; color: #f8fafc; margin: 8px 0; font-weight: 600; line-height: 1.35; min-height: 38px;">
                        {row['finalidade'][:75]}
                    </div>
                    <div style="display: flex; justify-content: space-between; font-size: 12.5px; border-top: 1px solid rgba(255,255,255,0.12); padding-top: 8px;">
                        <span style="color: #cbd5e1;">Crédito: <b style="color: #ffffff;">R$ {row['valor_total']:,.2f}</b></span>
                        <span style="color: #cbd5e1;">Saldo: <b style="color: #4ade80; font-size: 14px;">R$ {row['saldo_calc']:,.2f}</b></span>
                    </div>
                </div>
                ''', unsafe_allow_html=True)
                with st.popover("🔍 Ver Detalhes", key=f"pop_nc_{row['id']}", use_container_width=True):
                    c = conn.cursor()
                    c.execute('''
                    SELECT nc.*, 
                           (COALESCE((SELECT SUM(COALESCE(valor_nc_1, valor_ne)) FROM notas_empenho WHERE nc_id = nc.id), 0.0) +
                            COALESCE((SELECT SUM(COALESCE(valor_nc_2, 0.0)) FROM notas_empenho WHERE nc_id_2 = nc.id), 0.0)
                           ) as empenhado_real
                    FROM notas_credito nc
                    WHERE nc.id = ?
                    ''', (row['id'],))
                    nc_info = dict(c.fetchone())
                    saldo_det = max(nc_info['valor_total'] - nc_info['empenhado_real'] - (nc_info['valor_recolhido'] or 0.0), 0.0)

                    st.markdown(f'''
                    <div style="border-bottom: 2px solid #C5A059; padding-bottom: 8px; margin-bottom: 12px;">
                        <div style="font-size: 11px; font-weight: 800; color: #C5A059; text-transform: uppercase; letter-spacing: 0.5px;">Dossiê da Nota de Crédito</div>
                        <div style="font-size: 19px; font-weight: 900; color: #ffffff;">{nc_info['numero_nc']}</div>
                        <div style="font-size: 14px; font-weight: 600; color: #f1f5f9; margin-top: 4px;">{nc_info['finalidade']}</div>
                    </div>
                    ''', unsafe_allow_html=True)

                    st.markdown(f'''
                    <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; margin: 10px 0;">
                        <div style="background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08); border-left: 3px solid #C5A059; padding: 7px 10px; border-radius: 6px;">
                            <div style="font-size: 10px; color: #94a3b8; font-weight: 700; text-transform: uppercase;">Valor Total</div>
                            <div style="font-size: 15px; font-weight: 800; color: #ffffff; white-space: nowrap; margin-top: 2px;">R$ {nc_info['valor_total']:,.2f}</div>
                        </div>
                        <div style="background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08); border-left: 3px solid #60a5fa; padding: 7px 10px; border-radius: 6px;">
                            <div style="font-size: 10px; color: #94a3b8; font-weight: 700; text-transform: uppercase;">Empenhado</div>
                            <div style="font-size: 15px; font-weight: 800; color: #60a5fa; white-space: nowrap; margin-top: 2px;">R$ {nc_info['empenhado_real']:,.2f}</div>
                        </div>
                        <div style="background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08); border-left: 3px solid #fbbf24; padding: 7px 10px; border-radius: 6px;">
                            <div style="font-size: 10px; color: #94a3b8; font-weight: 700; text-transform: uppercase;">Recolhido</div>
                            <div style="font-size: 15px; font-weight: 800; color: #fbbf24; white-space: nowrap; margin-top: 2px;">R$ {nc_info['valor_recolhido'] or 0.0:,.2f}</div>
                        </div>
                        <div style="background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08); border-left: 3px solid #22c55e; padding: 7px 10px; border-radius: 6px;">
                            <div style="font-size: 10px; color: #94a3b8; font-weight: 700; text-transform: uppercase;">Disponível</div>
                            <div style="font-size: 15px; font-weight: 800; color: #22c55e; white-space: nowrap; margin-top: 2px;">R$ {saldo_det:,.2f}</div>
                        </div>
                    </div>
                    '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)

                    st.markdown("---")
                    ci1, ci2, ci3 = st.columns(3)
                    ci1.markdown(f"**Emissão:** {formatar_data_br(nc_info['data_emissao'])}")
                    ci1.markdown(f"**Limite:** {formatar_data_br(nc_info['data_limite_empenho'])}")
                    ci2.markdown(f"**Enquadramento:** {nc_info['enquadramento']}")
                    ci2.markdown(f"**ND:** {nc_info['natureza_despesa']}")
                    ci3.markdown(f"**PI:** {nc_info['pi']}")
                    ci3.markdown(f"**UG:** {nc_info['ug_emitente']}")

                    # Notas de Empenho vinculadas - Renderizado limpo via DataFrame nativo (sem código HTML visível)
                    st.markdown("---")
                    df_nes_vinc = pd.read_sql_query(f'''
                    SELECT numero_ne as "Número NE", data_emissao as "Emissão", fornecedor_nome as "Fornecedor",
                           fornecedor_cnpj as "CNPJ", valor_ne as "Valor Total NE (R$)",
                           CASE WHEN nc_id = {row['id']} THEN COALESCE(valor_nc_1, valor_ne) ELSE COALESCE(valor_nc_2, 0.0) END as "Valor desta NC (R$)",
                           status as "Status"
                    FROM notas_empenho WHERE nc_id = {row['id']} OR nc_id_2 = {row['id']} ORDER BY id DESC
                    ''', conn)

                    if not df_nes_vinc.empty:
                        df_nes_vinc['Emissão'] = df_nes_vinc['Emissão'].apply(formatar_data_br)
                        df_nes_vinc['Valor Total NE (R$)'] = df_nes_vinc['Valor Total NE (R$)'].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                        df_nes_vinc['Valor desta NC (R$)'] = df_nes_vinc['Valor desta NC (R$)'].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                        st.markdown(f"**📋 Notas de Empenho Vinculadas ({len(df_nes_vinc)} registro(s)):**")
                        st.dataframe(df_nes_vinc, use_container_width=True, hide_index=True)
                    else:
                        st.info("ℹ️ Nenhuma Nota de Empenho vinculada a esta NC. O saldo está 100% disponível.")

                    with st.expander("✏️ Editar Todos os Dados desta Nota de Crédito", expanded=False):
                        with st.form(f"form_ed_nc_card_{row['id']}"):
                            dc1, dc2 = st.columns(2)
                            d_num_nc = dc1.text_input("Número da NC:", value=nc_info['numero_nc'], key=f"d_num_{row['id']}")
                            d_val_tot = dc2.number_input("Valor Total (R$):", value=float(nc_info['valor_total']), step=100.0, format="%.2f", key=f"d_val_{row['id']}")

                            d_fin = st.text_input("Finalidade / Objeto:", value=nc_info['finalidade'] or "", key=f"d_fin_{row['id']}")

                            dc3, dc4, dc5 = st.columns(3)
                            d_dt_em = pd.to_datetime(nc_info['data_emissao']).date() if pd.notna(nc_info['data_emissao']) else datetime.now().date()
                            d_emissao = dc3.date_input("Data de Emissão:", value=d_dt_em, format="DD/MM/YYYY", key=f"d_em_{row['id']}")

                            d_dt_lim = pd.to_datetime(nc_info['data_limite_empenho']).date() if pd.notna(nc_info['data_limite_empenho']) else (datetime.now().date() + timedelta(days=30))
                            d_limite = dc4.date_input("Data Limite Empenho:", value=d_dt_lim, format="DD/MM/YYYY", key=f"d_lim_{row['id']}")

                            d_enq_opts = ["DISCRICIONÁRIA", "ALIMENTAÇÃO", "FARDAMENTO", "SAÚDE", "ENGENHARIA", "OUTRO"]
                            d_enq_idx = d_enq_opts.index(nc_info['enquadramento']) if nc_info['enquadramento'] in d_enq_opts else 0
                            d_enq = dc5.selectbox("Enquadramento:", d_enq_opts, index=d_enq_idx, key=f"d_enq_{row['id']}")

                            dc6, dc7 = st.columns(2)
                            d_nd = dc6.text_input("Natureza da Despesa (ND):", value=nc_info['natureza_despesa'] or "", key=f"d_nd_{row['id']}")
                            d_pi = dc7.text_input("Plano Interno (PI):", value=nc_info['pi'] or "", key=f"d_pi_{row['id']}")

                            d_rec = st.number_input("Valor de Saldo Recolhido (R$):", value=float(nc_info['valor_recolhido'] or 0.0), step=10.0, format="%.2f", key=f"d_rec_{row['id']}")

                            if st.form_submit_button("💾 Salvar Alterações desta NC", type="primary"):
                                cur = conn.cursor()
                                cur.execute('''
                                UPDATE notas_credito
                                SET numero_nc = ?, finalidade = ?, valor_total = ?, valor_recolhido = ?,
                                    data_emissao = ?, data_limite_empenho = ?, enquadramento = ?,
                                    natureza_despesa = ?, pi = ?
                                WHERE id = ?
                                ''', (d_num_nc.strip(), d_fin.strip(), d_val_tot, d_rec, str(d_emissao), str(d_limite), d_enq, d_nd.strip(), d_pi.strip(), row['id']))
                                conn.commit()
                                st.success(f"✅ NC {d_num_nc} atualizada!")
                                st.rerun()

                    st.markdown("---")
                    col_rec1, col_rec2 = st.columns([2, 1])
                    novo_rec = col_rec1.number_input("Informar Saldo Recolhido (R$):", value=float(nc_info['valor_recolhido'] or 0.0), step=10.0, format="%.2f", key=f"in_rec_{row['id']}")
                    if col_rec2.button("Gravar Baixa", type="primary", key=f"btn_rec_{row['id']}", use_container_width=True):
                        c.execute("UPDATE notas_credito SET valor_recolhido = ? WHERE id = ?", (novo_rec, row['id']))
                        conn.commit()
                        st.success("Valor recolhido gravado com sucesso!")
                        st.rerun()
    conn.close()
