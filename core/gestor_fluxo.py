from datetime import datetime, timedelta
import pandas as pd
try:
    from core.database import get_connection
except ImportError:
    from database import get_connection

class GestorFluxoMilitar:
    """
    Centraliza a inteligência e o fluxo operacional do SisCalc / SisLog da 5ª Cia PE.
    Gerencia regras de negócio entre Créditos (NC), Empenhos (NE), Recebimento (NF) e Estoque.
    """

    @staticmethod
    def obter_saldo_nc(nc_id, conn=None):
        """
        Calcula o saldo real e auditável de uma Nota de Crédito,
        suportando empenhos com fonte única ou dupla (nc_id e nc_id_2).
        """
        close_conn = False
        if conn is None:
            conn = get_connection()
            close_conn = True

        c = conn.cursor()
        c.execute('''
        SELECT id, numero_nc, valor_total, COALESCE(valor_recolhido, 0.0) as valor_recolhido,
               data_limite_empenho, finalidade, enquadramento
        FROM notas_credito WHERE id = ?
        ''', (nc_id,))
        row = c.fetchone()
        if not row:
            if close_conn: conn.close()
            return None

        nc_dados = dict(row)

        # Soma valores onde a NC é a fonte primária (nc_id)
        c.execute('''
        SELECT COALESCE(SUM(COALESCE(valor_nc_1, valor_ne)), 0.0)
        FROM notas_empenho WHERE nc_id = ?
        ''', (nc_id,))
        emp_1 = c.fetchone()[0] or 0.0

        # Soma valores onde a NC é a fonte secundária (nc_id_2)
        c.execute('''
        SELECT COALESCE(SUM(COALESCE(valor_nc_2, 0.0)), 0.0)
        FROM notas_empenho WHERE nc_id_2 = ?
        ''', (nc_id,))
        emp_2 = c.fetchone()[0] or 0.0

        total_empenhado = emp_1 + emp_2
        recolhido = float(nc_dados['valor_recolhido'])
        saldo_disponivel = max(nc_dados['valor_total'] - total_empenhado - recolhido, 0.0)

        res = {
            "id": nc_dados['id'],
            "numero_nc": nc_dados['numero_nc'],
            "valor_total": nc_dados['valor_total'],
            "total_empenhado": total_empenhado,
            "valor_recolhido": recolhido,
            "saldo_disponivel": saldo_disponivel,
            "finalidade": nc_dados['finalidade'],
            "enquadramento": nc_dados['enquadramento'],
            "data_limite": nc_dados['data_limite_empenho']
        }

        if close_conn: conn.close()
        return res

    @staticmethod
    def listar_alertas_ncs_criticas(om_id, dias_limite=7, conn=None):
        """
        Localiza todas as NCs com saldo em risco que vencem nos próximos 'dias_limite' ou já venceram.
        """
        close_conn = False
        if conn is None:
            conn = get_connection()
            close_conn = True

        hoje = datetime.now().date()
        c = conn.cursor()
        c.execute("SELECT id FROM notas_credito WHERE om_id = ? ORDER BY data_limite_empenho ASC", (om_id,))
        ncs = c.fetchall()

        alertas = []
        for r in ncs:
            saldo_info = GestorFluxoMilitar.obter_saldo_nc(r[0], conn)
            if saldo_info and saldo_info['saldo_disponivel'] > 0.01 and saldo_info['data_limite']:
                try:
                    dt_lim = pd.to_datetime(saldo_info['data_limite']).date()
                    dias = (dt_lim - hoje).days
                    if dias <= dias_limite:
                        if dias < 0:
                            gravidade = "VENCIDA"
                            badge_cor = "#ef4444"
                            badge_txt = f"🔴 VENCIDA ({abs(dias)}d atrás)"
                        elif dias <= 2:
                            gravidade = "URGENTE"
                            badge_cor = "#f97316"
                            badge_txt = f"🟠 URGENTE ({dias} dias)"
                        else:
                            gravidade = "ATENÇÃO"
                            badge_cor = "#eab308"
                            badge_txt = f"🟡 ATENÇÃO ({dias} dias)"

                        saldo_info['dias_restantes'] = dias
                        saldo_info['gravidade'] = gravidade
                        saldo_info['badge_cor'] = badge_cor
                        saldo_info['badge_txt'] = badge_txt
                        alertas.append(saldo_info)
                except Exception:
                    pass

        if close_conn: conn.close()
        return alertas

    @staticmethod
    def listar_alertas_nes_atrasadas(om_id, dias_limite=5, conn=None):
        """
        Localiza fornecedores em atraso de entrega ou com prazo expirando nos próximos dias.
        """
        close_conn = False
        if conn is None:
            conn = get_connection()
            close_conn = True

        hoje = datetime.now().date()
        c = conn.cursor()
        c.execute('''
        SELECT id, numero_ne, fornecedor_nome, fornecedor_telefone, valor_ne, status,
               COALESCE(nova_data_limite, data_limite) as prazo_final
        FROM notas_empenho
        WHERE om_id = ? AND status NOT IN ('Liquidado', 'Pago')
        ORDER BY prazo_final ASC
        ''', (om_id,))
        rows = c.fetchall()

        alertas = []
        for r in rows:
            ne_item = dict(r)
            if ne_item['prazo_final']:
                try:
                    p_dt = pd.to_datetime(ne_item['prazo_final']).date()
                    dias = (p_dt - hoje).days
                    if dias <= dias_limite:
                        if dias < 0:
                            gravidade = "ATRASADA"
                            badge_cor = "#ef4444"
                            badge_txt = f"🔴 ATRASADA ({abs(dias)}d)"
                        elif dias <= 2:
                            gravidade = "URGENTE"
                            badge_cor = "#f97316"
                            badge_txt = f"🟠 VENCE EM {dias}d"
                        else:
                            gravidade = "ATENÇÃO"
                            badge_cor = "#eab308"
                            badge_txt = f"🟡 VENCE EM {dias}d"

                        ne_item['dias_restantes'] = dias
                        ne_item['gravidade'] = gravidade
                        ne_item['badge_cor'] = badge_cor
                        ne_item['badge_txt'] = badge_txt
                        alertas.append(ne_item)
                except Exception:
                    pass

        if close_conn: conn.close()
        return alertas

    @staticmethod
    def calcular_metricas_estoque(consumo_anual, saldo_atual):
        """
        Calcula as métricas regulamentares de suprimento para a 5ª Cia PE:
        - Estoque Mínimo: 20% do Consumo Anual (~2,4 meses para trâmite da SALC)
        - Estoque Ideal: 50% do Consumo Anual (cobertura para 6 meses operacionais)
        """
        if consumo_anual > 0:
            est_min = round(consumo_anual * 0.20, 1)
            est_id = round(consumo_anual * 0.50, 1)
        else:
            est_min = round(saldo_atual * 0.20, 1) if saldo_atual > 0 else 5.0
            est_id = round(saldo_atual, 1) if saldo_atual > 0 else 20.0

        if saldo_atual <= est_min:
            status_txt = "🔴 CRÍTICO / REPOR"
            status_bg = "#7f1d1d"
            borda = "border-left: 6px solid #ef4444;"
        elif saldo_atual <= (est_min * 1.3):
            status_txt = "🟡 ATENÇÃO"
            status_bg = "#854d0e"
            borda = "border-left: 6px solid #eab308;"
        else:
            status_txt = "🟢 NORMAL"
            status_bg = "#14532d"
            borda = "border-left: 6px solid #22c55e;"

        return {
            "estoque_minimo": est_min,
            "estoque_ideal": est_id,
            "status_txt": status_txt,
            "status_bg": status_bg,
            "borda_css": borda,
            "qtd_a_comprar": max(est_id - saldo_atual, 0.0)
        }

if __name__ == "__main__":
    print("Gestor de Fluxo Militar compilado e pronto para uso!")
