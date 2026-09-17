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
    st.title("📦 Recebimento de Materiais e Notas Fiscais")
    conn = get_connection()
    df_nfs = pd.read_sql_query(f"SELECT nf.id, ne.numero_ne, nf.numero_nf, nf.empresa_cnpj, nf.data_entrada_almox, nf.valor_nf, nf.situacao FROM notas_fiscais nf JOIN notas_empenho ne ON nf.ne_id = ne.id WHERE nf.om_id = {user['om_id']}", conn)
    if not df_nfs.empty:
        df_nfs['Data Entrada Almox'] = df_nfs['data_entrada_almox'].apply(formatar_data_br)
        st.dataframe(df_nfs[['numero_ne', 'numero_nf', 'empresa_cnpj', 'Data Entrada Almox', 'valor_nf', 'situacao']], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma Nota Fiscal registrada.")
    conn.close()
