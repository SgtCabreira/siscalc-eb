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
    st.title("📁 Central de Relatórios Oficiais com Filtros")
    conn = get_connection()

    tipo_relatorio = st.selectbox("Selecione o Relatório a Emitir:", ["Notas de Crédito (com Empenhos Vinculados)", "Notas de Empenho", "Notas Fiscais (Recebimento)"])

    if "Notas de Crédito" in tipo_relatorio:
        st.markdown("##### 🔍 Filtros de Consulta das Notas de Crédito")
        fc1, fc2, fc3 = st.columns(3)
        df_ncs_base = pd.read_sql_query(f"SELECT DISTINCT enquadramento FROM notas_credito WHERE om_id = {user['om_id']}", conn)

        enq_opts = ["Todos"] + sorted(list(set(df_ncs_base['enquadramento'].dropna().tolist())))
        f_enq = fc1.selectbox("Enquadramento:", enq_opts, key="f_enq_rel")

        tipo_opts = ["Todos", "Consumo (33)", "Permanente (44)"]
        f_tipo = fc2.selectbox("Natureza da Despesa:", tipo_opts, key="f_tipo_rel")
        busca_texto = fc3.text_input("Palavra-chave na Finalidade:", key="f_txt_rel")

        q_rel_nc = f"""
        SELECT nc.numero_nc as "Número NC",
               nc.data_emissao,
               nc.data_limite_empenho,
               nc.enquadramento as "Enquadramento",
               nc.natureza_despesa as "ND",
               nc.pi as "PI",
               nc.valor_total as "Valor Total (R$)",
               (COALESCE((SELECT SUM(COALESCE(valor_nc_1, valor_ne)) FROM notas_empenho WHERE nc_id = nc.id), 0.0) +
                COALESCE((SELECT SUM(COALESCE(valor_nc_2, 0.0)) FROM notas_empenho WHERE nc_id_2 = nc.id), 0.0)
               ) as "Total Empenhado (R$)",
               (nc.valor_total - COALESCE(nc.valor_recolhido, 0.0) - 
                (COALESCE((SELECT SUM(COALESCE(valor_nc_1, valor_ne)) FROM notas_empenho WHERE nc_id = nc.id), 0.0) +
                 COALESCE((SELECT SUM(COALESCE(valor_nc_2, 0.0)) FROM notas_empenho WHERE nc_id_2 = nc.id), 0.0))
               ) as "Saldo Disponível (R$)",
               COALESCE((SELECT GROUP_CONCAT(numero_ne, ', ') FROM notas_empenho WHERE nc_id = nc.id OR nc_id_2 = nc.id), 'Sem empenho') as "Empenhos Vinculados",
               nc.finalidade as "Finalidade"
        FROM notas_credito nc
        WHERE nc.om_id = {user['om_id']}
        """
        if f_enq != "Todos":
            q_rel_nc += f" AND nc.enquadramento = '{f_enq}'"
        if f_tipo == "Consumo (33)":
            q_rel_nc += " AND nc.natureza_despesa LIKE '%33%'"
        elif f_tipo == "Permanente (44)":
            q_rel_nc += " AND nc.natureza_despesa LIKE '%44%'"
        if busca_texto:
            q_rel_nc += f" AND nc.finalidade LIKE '%{busca_texto}%'"

        q_rel_nc += " ORDER BY nc.data_emissao DESC"

        df_rel_nc = pd.read_sql_query(q_rel_nc, conn)
        if not df_rel_nc.empty:
            df_rel_nc['Data Emissão'] = df_rel_nc['data_emissao'].apply(formatar_data_br)
            df_rel_nc['Dt Lim Empenho'] = df_rel_nc['data_limite_empenho'].apply(formatar_data_br)

            # Formata colunas para exibição amigável
            cols_ordem = [
                "Número NC", "Data Emissão", "Dt Lim Empenho", "Enquadramento", "ND", "PI",
                "Valor Total (R$)", "Total Empenhado (R$)", "Saldo Disponível (R$)", 
                "Empenhos Vinculados", "Finalidade"
            ]
            df_exib_nc = df_rel_nc[cols_ordem].copy()

            # Formata valores em R$
            df_exib_formatado = df_exib_nc.copy()
            df_exib_formatado['Valor Total (R$)'] = df_exib_formatado['Valor Total (R$)'].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            df_exib_formatado['Total Empenhado (R$)'] = df_exib_formatado['Total Empenhado (R$)'].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            df_exib_formatado['Saldo Disponível (R$)'] = df_exib_formatado['Saldo Disponível (R$)'].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            st.dataframe(df_exib_formatado, use_container_width=True, hide_index=True)

            col_csv_nc, col_xlsx_nc, col_pdf_nc = st.columns(3)
            with col_csv_nc:
                st.download_button("📥 Baixar CSV", data=df_exib_nc.to_csv(index=False).encode('utf-8'), file_name="relatorio_notas_credito.csv", mime="text/csv", use_container_width=True)

            with col_xlsx_nc:
                buf_excel_nc = exportar_excel_seguro(df_exib_nc, 'Notas de Crédito')
                if buf_excel_nc:
                    st.download_button("📊 Baixar Planilha (.xlsx / Calc)", data=buf_excel_nc, file_name="relatorio_notas_credito.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            with col_pdf_nc:
                buf_pdf_nc = gerar_pdf_tabela_seguro("RELATÓRIO GERAL DE NOTAS DE CRÉDITO E EMPENHOS VINCULADOS", df_exib_formatado)
                if buf_pdf_nc:
                    st.download_button("📄 Baixar em PDF (.pdf)", data=buf_pdf_nc, file_name="relatorio_notas_credito.pdf", mime="application/pdf", use_container_width=True)
        else:
            st.info("Nenhuma Nota de Crédito encontrada com os filtros selecionados.")

    elif "Notas de Empenho" in tipo_relatorio:
        st.markdown("##### 🔍 Filtros de Consulta das Notas de Empenho")
        fe1, fe2, fe3 = st.columns(3)
        df_nes_base = pd.read_sql_query(f"SELECT DISTINCT status, tipo_empenho FROM notas_empenho WHERE om_id = {user['om_id']}", conn)

        status_opts = ["Todos"] + sorted(list(set(df_nes_base['status'].dropna().tolist())))
        f_status = fe1.selectbox("Status de Execução:", status_opts, key="f_st_ne_rel")

        tipo_emp_opts = ["Todos"] + sorted(list(set(df_nes_base['tipo_empenho'].dropna().tolist())))
        f_tipo_emp = fe2.selectbox("Tipo de Empenho:", tipo_emp_opts, key="f_tipo_ne_rel")
        busca_forn = fe3.text_input("Filtrar por Fornecedor (Nome ou CNPJ):", key="f_forn_ne_rel")

        q_rel_ne = f"""
        SELECT ne.numero_ne as "Número NE",
               CASE WHEN nc2.numero_nc IS NOT NULL THEN (nc1.numero_nc || ' + ' || nc2.numero_nc) ELSE nc1.numero_nc END as "NC Origem",
               ne.tipo_empenho as "Tipo",
               ne.fornecedor_nome as "Fornecedor",
               ne.fornecedor_cnpj as "CNPJ",
               ne.valor_ne as "Valor NE (R$)",
               ne.data_emissao,
               COALESCE(ne.nova_data_limite, ne.data_limite) as data_limite,
               ne.status as "Status"
        FROM notas_empenho ne
        JOIN notas_credito nc1 ON ne.nc_id = nc1.id
        LEFT JOIN notas_credito nc2 ON ne.nc_id_2 = nc2.id
        WHERE ne.om_id = {user['om_id']}
        """
        if f_status != "Todos":
            q_rel_ne += f" AND ne.status = '{f_status}'"
        if f_tipo_emp != "Todos":
            q_rel_ne += f" AND ne.tipo_empenho = '{f_tipo_emp}'"
        if busca_forn:
            q_rel_ne += f" AND (ne.fornecedor_nome LIKE '%{busca_forn}%' OR ne.fornecedor_cnpj LIKE '%{busca_forn}%')"

        q_rel_ne += " ORDER BY ne.data_emissao DESC"

        df_rel_ne = pd.read_sql_query(q_rel_ne, conn)
        if not df_rel_ne.empty:
            df_rel_ne['Data Emissão'] = df_rel_ne['data_emissao'].apply(formatar_data_br)
            df_rel_ne['Data Limite Entrega'] = df_rel_ne['data_limite'].apply(formatar_data_br)

            cols_ne_ordem = ["Número NE", "NC Origem", "Tipo", "Fornecedor", "CNPJ", "Valor NE (R$)", "Data Emissão", "Data Limite Entrega", "Status"]
            df_exib_ne = df_rel_ne[cols_ne_ordem].copy()

            df_exib_ne_fmt = df_exib_ne.copy()
            df_exib_ne_fmt['Valor NE (R$)'] = df_exib_ne_fmt['Valor NE (R$)'].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            st.dataframe(df_exib_ne_fmt, use_container_width=True, hide_index=True)

            col_csv_ne, col_xlsx_ne, col_pdf_ne = st.columns(3)
            with col_csv_ne:
                st.download_button("📥 Baixar CSV", data=df_exib_ne.to_csv(index=False).encode('utf-8'), file_name="relatorio_notas_empenho.csv", mime="text/csv", use_container_width=True)
            with col_xlsx_ne:
                buf_excel_ne = exportar_excel_seguro(df_exib_ne, 'Notas de Empenho')
                if buf_excel_ne:
                    st.download_button("📊 Baixar Planilha (.xlsx / Calc)", data=buf_excel_ne, file_name="relatorio_notas_empenho.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with col_pdf_ne:
                buf_pdf_ne = gerar_pdf_tabela_seguro("RELATÓRIO GERAL DE NOTAS DE EMPENHO", df_exib_ne_fmt)
                if buf_pdf_ne:
                    st.download_button("📄 Baixar em PDF (.pdf)", data=buf_pdf_ne, file_name="relatorio_notas_empenho.pdf", mime="application/pdf", use_container_width=True)
        else:
            st.info("Nenhum Empenho encontrado com os filtros selecionados.")

    elif "Notas Fiscais" in tipo_relatorio:
        q_rel_nf = f"""
        SELECT nf.numero_nf as "Número NF",
               ne.numero_ne as "Empenho (NE)",
               nc.numero_nc as "NC Origem",
               nf.empresa_cnpj as "CNPJ Fornecedor",
               nf.data_entrada_almox,
               nf.valor_nf as "Valor NF (R$)",
               nf.tipo_liquidacao as "Tipo Liquidação",
               nf.situacao as "Situação"
        FROM notas_fiscais nf
        JOIN notas_empenho ne ON nf.ne_id = ne.id
        JOIN notas_credito nc ON ne.nc_id = nc.id
        WHERE nf.om_id = {user['om_id']}
        ORDER BY nf.id DESC
        """
        df_rel_nf = pd.read_sql_query(q_rel_nf, conn)
        if not df_rel_nf.empty:
            df_rel_nf['Data Entrada Almox'] = df_rel_nf['data_entrada_almox'].apply(formatar_data_br)
            cols_nf = ["Número NF", "Empenho (NE)", "NC Origem", "CNPJ Fornecedor", "Data Entrada Almox", "Valor NF (R$)", "Tipo Liquidação", "Situação"]
            df_exib_nf = df_rel_nf[cols_nf].copy()

            df_exib_nf_fmt = df_exib_nf.copy()
            df_exib_nf_fmt['Valor NF (R$)'] = df_exib_nf_fmt['Valor NF (R$)'].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            st.dataframe(df_exib_nf_fmt, use_container_width=True, hide_index=True)

            col_csv_nf, col_xlsx_nf, col_pdf_nf = st.columns(3)
            with col_csv_nf:
                st.download_button("📥 Baixar CSV", data=df_exib_nf.to_csv(index=False).encode('utf-8'), file_name="relatorio_notas_fiscais.csv", mime="text/csv", use_container_width=True)
            with col_xlsx_nf:
                buf_excel_nf = exportar_excel_seguro(df_exib_nf, 'Notas Fiscais')
                if buf_excel_nf:
                    st.download_button("📊 Baixar Planilha (.xlsx / Calc)", data=buf_excel_nf, file_name="relatorio_notas_fiscais.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with col_pdf_nf:
                buf_pdf_nf = gerar_pdf_tabela_seguro("RELATÓRIO GERAL DE NOTAS FISCAIS RECEBIDAS", df_exib_nf_fmt)
                if buf_pdf_nf:
                    st.download_button("📄 Baixar em PDF (.pdf)", data=buf_pdf_nf, file_name="relatorio_notas_fiscais.pdf", mime="application/pdf", use_container_width=True)
        else:
            st.info("Nenhuma Nota Fiscal encontrada.")

    conn.close()
