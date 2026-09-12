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
    # Banner Institucional com Cores Oficiais da 5ª Cia PE
    st.markdown(f'''
    <div style="background: linear-gradient(135deg, #0A2240 0%, #173863 60%, #1B4332 100%); padding: 20px 25px; border-radius: 12px; margin-bottom: 22px; box-shadow: 0 4px 15px rgba(10,34,64,0.25); border-left: 6px solid #C5A059; display: flex; justify-content: space-between; align-items: center;">
        <div>
            <div style="font-size: 22px; font-weight: 900; color: #ffffff; letter-spacing: 0.5px;">🛡️ Painel de Gestão Orçamentária, Prazos & Alertas</div>
            <div style="font-size: 13px; color: #C5A059; font-weight: 600; margin-top: 3px;">5ª Companhia de Polícia do Exército — Base Major Agostinho José Rodrigues</div>
        </div>
        <div style="text-align: right;">
            <span style="background: rgba(255,255,255,0.15); color: #ffffff; padding: 5px 14px; border-radius: 20px; font-size: 12px; font-weight: 800; border: 1px solid rgba(197,160,89,0.5);">
                👁️ VISÃO DE COMANDO
            </span>
        </div>
    </div>
    ''', unsafe_allow_html=True)

    conn = get_connection()
    hoje = datetime.now().date()

    total_nc = pd.read_sql_query(f"SELECT SUM(valor_total) as val FROM notas_credito WHERE om_id = {user['om_id']}", conn)['val'].values[0] or 0.0
    total_ne = pd.read_sql_query(f"SELECT SUM(valor_ne) as val FROM notas_empenho WHERE om_id = {user['om_id']}", conn)['val'].values[0] or 0.0
    total_rec = pd.read_sql_query(f"SELECT SUM(valor_recolhido) as val FROM notas_credito WHERE om_id = {user['om_id']}", conn)['val'].values[0] or 0.0
    total_liq = pd.read_sql_query(f"SELECT SUM(valor_nf) as val FROM notas_fiscais WHERE om_id = {user['om_id']} AND situacao IN ('Liquidado (Atestado no SIAFI)', 'Pago (Ordem Bancária Emitida)')", conn)['val'].values[0] or 0.0
    total_pago = pd.read_sql_query(f"SELECT SUM(valor_nf) as val FROM notas_fiscais WHERE om_id = {user['om_id']} AND situacao = 'Pago (Ordem Bancária Emitida)'", conn)['val'].values[0] or 0.0
    saldo_empenhar = max(total_nc - total_ne - total_rec, 0.0)

    # 5 Cards Executivos de Métricas com Design Metálico 3D
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(f'''
        <div style="background: linear-gradient(145deg, #1c232e, #131821); border: 1px solid rgba(255,255,255,0.08); border-top: 4px solid #38bdf8; border-radius: 10px; padding: 14px 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.35);">
            <div style="font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px;">Créditos Recebidos (NC)</div>
            <div style="font-size: 17px; font-weight: 800; white-space: nowrap; color: #ffffff; margin-top: 4px;">R$ {total_nc:,.2f}</div>
        </div>
        '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
    with c2:
        st.markdown(f'''
        <div style="background: linear-gradient(145deg, #1c232e, #131821); border: 1px solid rgba(255,255,255,0.08); border-top: 4px solid #60a5fa; border-radius: 10px; padding: 14px 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.35);">
            <div style="font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px;">Total Empenhado (NE)</div>
            <div style="font-size: 17px; font-weight: 800; white-space: nowrap; color: #60a5fa; margin-top: 4px;">R$ {total_ne:,.2f}</div>
        </div>
        '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
    with c3:
        st.markdown(f'''
        <div style="background: linear-gradient(145deg, #1c232e, #131821); border: 1px solid rgba(255,255,255,0.08); border-top: 4px solid #22c55e; border-radius: 10px; padding: 14px 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.35);">
            <div style="font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px;">Saldo Disponível SALC</div>
            <div style="font-size: 17px; font-weight: 800; white-space: nowrap; color: #22c55e; margin-top: 4px;">R$ {saldo_empenhar:,.2f}</div>
        </div>
        '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
    with c4:
        st.markdown(f'''
        <div style="background: linear-gradient(145deg, #1c232e, #131821); border: 1px solid rgba(255,255,255,0.08); border-top: 4px solid #f59e0b; border-radius: 10px; padding: 14px 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.35);">
            <div style="font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px;">Material Liquidado</div>
            <div style="font-size: 17px; font-weight: 800; white-space: nowrap; color: #f59e0b; margin-top: 4px;">R$ {total_liq:,.2f}</div>
        </div>
        '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
    with c5:
        st.markdown(f'''
        <div style="background: linear-gradient(145deg, #1c232e, #131821); border: 1px solid rgba(255,255,255,0.08); border-top: 4px solid #a855f7; border-radius: 10px; padding: 14px 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.35);">
            <div style="font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px;">Ordem Bancária (Pago)</div>
            <div style="font-size: 17px; font-weight: 800; white-space: nowrap; color: #c084fc; margin-top: 4px;">R$ {total_pago:,.2f}</div>
        </div>
        '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)

    st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)

    # Barras de Progresso da Execução Orçamentária Global
    with st.container():
        st.markdown("#### 📈 Execução Orçamentária da OM")
        cp1, cp2, cp3 = st.columns(3)
        p_emp = (total_ne / total_nc) if total_nc > 0 else 0.0
        p_liq = (total_liq / total_ne) if total_ne > 0 else 0.0
        p_pag = (total_pago / total_ne) if total_ne > 0 else 0.0

        with cp1:
            st.markdown(f"**1. Crédito $\\rightarrow$ Empenho:** `{p_emp*100:.1f}%`")
            st.progress(min(p_emp, 1.0))
        with cp2:
            st.markdown(f"**2. Empenho $\\rightarrow$ Liquidação:** `{p_liq*100:.1f}%`")
            st.progress(min(p_liq, 1.0))
        with cp3:
            st.markdown(f"**3. Empenho $\\rightarrow$ Pagamento:** `{p_pag*100:.1f}%`")
            st.progress(min(p_pag, 1.0))

    st.markdown("---")

    # =========================================================
    # RADAR DE ALERTAS DO COMANDO (PRAZOS DE NCs E NEs)
    # =========================================================
    st.markdown("### 🚨 Radar de Prazos Críticos & Alertas do Comando")
    st.caption("Visão centralizada de pendências operacionais para tomada de decisão imediata do Comando e da 4ª Seção.")

    col_alert_nc, col_alert_ne = st.columns(2)

    # 1. ALERTA DE NOTAS DE CRÉDITO (RISCO DE PERDA DE RECURSO POR PRAZO)
    with col_alert_nc:
        st.markdown("##### ⏳ Notas de Crédito no Ponto Limite de Empenho")

        q_nc_alertas = f'''
        SELECT nc.id, nc.numero_nc, nc.enquadramento, nc.data_limite_empenho, nc.finalidade, nc.valor_total,
               (nc.valor_total - COALESCE(nc.valor_recolhido, 0.0) - 
                (COALESCE((SELECT SUM(COALESCE(valor_nc_1, valor_ne)) FROM notas_empenho WHERE nc_id = nc.id), 0.0) +
                 COALESCE((SELECT SUM(COALESCE(valor_nc_2, 0.0)) FROM notas_empenho WHERE nc_id_2 = nc.id), 0.0))
               ) as saldo
        FROM notas_credito nc
        WHERE nc.om_id = {user['om_id']}
        ORDER BY nc.data_limite_empenho ASC
        '''
        df_nc_radar = pd.read_sql_query(q_nc_alertas, conn)

        ncs_com_urgencia = []
        if not df_nc_radar.empty:
            for _, row_nc in df_nc_radar.iterrows():
                saldo_val = float(row_nc['saldo'] or 0.0)
                if saldo_val > 0.01 and pd.notna(row_nc['data_limite_empenho']):
                    try:
                        limite_dt = pd.to_datetime(row_nc['data_limite_empenho']).date()
                        dias = (limite_dt - hoje).days
                        if dias <= 7: # Alerta se vence em até 7 dias ou já venceu
                            ncs_com_urgencia.append({
                                "numero_nc": row_nc['numero_nc'],
                                "enquadramento": row_nc['enquadramento'],
                                "finalidade": row_nc['finalidade'],
                                "saldo": saldo_val,
                                "limite": limite_dt,
                                "dias": dias
                            })
                    except Exception:
                        pass

        if ncs_com_urgencia:
            for item in ncs_com_urgencia:
                dias = item['dias']
                if dias < 0:
                    badge = f"<span style='background: #ef4444; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;'>🔴 VENCIDA ({abs(dias)}d atrás)</span>"
                    borda_color = "#ef4444"
                elif dias <= 2:
                    badge = f"<span style='background: #f97316; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;'>🟠 URGENTE ({dias} dias)</span>"
                    borda_color = "#f97316"
                else:
                    badge = f"<span style='background: #eab308; color: black; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;'>🟡 ATENÇÃO ({dias} dias)</span>"
                    borda_color = "#eab308"

                st.markdown(f'''
                <div style="background: linear-gradient(145deg, #1c232e, #131821); border: 1px solid rgba(255,255,255,0.08); border-left: 5px solid {borda_color}; border-radius: 10px; padding: 13px 15px; margin-bottom: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.35);">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="font-weight: 800; font-size: 15px; color: #ffffff;">{item['numero_nc']}</span>
                        {badge}
                    </div>
                    <div style="font-size: 13px; color: #cbd5e1; margin: 6px 0; font-weight: 600;">{item['finalidade'][:75]}</div>
                    <div style="display: flex; justify-content: space-between; font-size: 12px; border-top: 1px solid rgba(255,255,255,0.06); padding-top: 6px;">
                        <span style="color: #94a3b8;">Limite: <b>{formatar_data_br(item['limite'])}</b></span>
                        <span style="color: #ef4444; font-weight: 800;">Saldo em Risco: R$ {item['saldo']:,.2f}</span>
                    </div>
                </div>
                '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
        else:
            st.markdown('''
            <div style="background: rgba(34, 197, 94, 0.1); border: 1.5px solid #22c55e; border-radius: 8px; padding: 16px; text-align: center; color: #166534; font-weight: 700;">
                🟢 Nenhuma Nota de Crédito com prazo crítico de empenho ou recurso represado.
            </div>
            ''', unsafe_allow_html=True)

    # 2. ALERTA DE NOTAS DE EMPENHO (ENTREGA DE MATERIAIS PELAS EMPRESAS)
    with col_alert_ne:
        st.markdown("##### 🚚 Prazos de Entrega dos Fornecedores (NE)")

        q_ne_alertas = f'''
        SELECT ne.numero_ne, ne.fornecedor_nome, ne.fornecedor_telefone, ne.valor_ne, ne.status,
               COALESCE(ne.nova_data_limite, ne.data_limite) as prazo_final
        FROM notas_empenho ne
        WHERE ne.om_id = {user['om_id']} AND ne.status NOT IN ('Liquidado', 'Pago')
        ORDER BY prazo_final ASC
        '''
        df_ne_radar = pd.read_sql_query(q_ne_alertas, conn)

        nes_com_urgencia = []
        if not df_ne_radar.empty:
            for _, row_ne in df_ne_radar.iterrows():
                if pd.notna(row_ne['prazo_final']):
                    try:
                        p_dt = pd.to_datetime(row_ne['prazo_final']).date()
                        dias_e = (p_dt - hoje).days
                        if dias_e <= 5: # Alerta se vence em até 5 dias ou já atrasou
                            nes_com_urgencia.append({
                                "numero_ne": row_ne['numero_ne'],
                                "fornecedor": row_ne['fornecedor_nome'],
                                "telefone": row_ne['fornecedor_telefone'] or "Não cadastrado",
                                "valor": float(row_ne['valor_ne'] or 0.0),
                                "prazo": p_dt,
                                "dias": dias_e
                            })
                    except Exception:
                        pass

        if nes_com_urgencia:
            for item in nes_com_urgencia:
                dias_e = item['dias']
                if dias_e < 0:
                    badge_ne = f"<span style='background: #ef4444; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;'>🔴 ATRASADO ({abs(dias_e)}d)</span>"
                    borda_ne_cor = "#ef4444"
                elif dias_e <= 2:
                    badge_ne = f"<span style='background: #f97316; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;'>🟠 VENCE EM {dias_e}d</span>"
                    borda_ne_cor = "#f97316"
                else:
                    badge_ne = f"<span style='background: #eab308; color: black; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;'>🟡 VENCE EM {dias_e}d</span>"
                    borda_ne_cor = "#eab308"

                st.markdown(f'''
                <div style="background: linear-gradient(145deg, #1c232e, #131821); border: 1px solid rgba(255,255,255,0.08); border-left: 5px solid {borda_ne_cor}; border-radius: 10px; padding: 13px 15px; margin-bottom: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.35);">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="font-weight: 800; font-size: 15px; color: #ffffff;">{item['numero_ne']}</span>
                        {badge_ne}
                    </div>
                    <div style="font-size: 13px; color: #cbd5e1; margin: 6px 0; font-weight: 700;">{item['fornecedor'][:55]}</div>
                    <div style="display: flex; justify-content: space-between; font-size: 12px; border-top: 1px solid rgba(255,255,255,0.06); padding-top: 6px;">
                        <span style="color: #94a3b8;">📞 <b>{item['telefone']}</b> | Limite: <b>{formatar_data_br(item['prazo'])}</b></span>
                        <span style="color: #60a5fa; font-weight: 800;">R$ {item['valor']:,.2f}</span>
                    </div>
                </div>
                '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
        else:
            st.markdown('''
            <div style="background: rgba(34, 197, 94, 0.1); border: 1.5px solid #22c55e; border-radius: 8px; padding: 16px; text-align: center; color: #166534; font-weight: 700;">
                🟢 Todas as empresas e materiais estão com prazos de entrega rigorosamente em dia.
            </div>
            ''', unsafe_allow_html=True)

    conn.close()
