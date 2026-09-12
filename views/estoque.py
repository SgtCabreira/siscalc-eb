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
    st.title("🏢 Almoxarifado & Estoque da 5ª Cia PE")
    conn = get_connection()

    # Migração automática de colunas para bancos existentes
    cur_mig = conn.cursor()
    for col, col_def in [
        ("tipo_embalagem", "TEXT DEFAULT 'UNIDADE'"),
        ("fator_embalagem", "REAL DEFAULT 1.0")
    ]:
        try:
            cur_mig.execute(f"ALTER TABLE estoque_itens ADD COLUMN {col} {col_def}")
        except sqlite3.OperationalError:
            pass
    conn.commit()

    aba_est1, aba_est2, aba_est3, aba_est4 = st.tabs([
        "📦 Posição de Estoque",
        "📥 Registrar Entrada",
        "📤 Registrar Saída / Cautela",
        "🧮 Histórico & Memória de Cálculo"
    ])

    def calcular_metricas_item(consumo_anual, saldo_atual):
        # Métrica de Suprimento Militar:
        # Estoque Mínimo (Ponto de Pedido) = 20% do Consumo Anual (~2,4 meses para processo da SALC)
        # Estoque Ideal (Capacidade de Operação) = 50% do Consumo Anual (cobertura semestral)
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

        return est_min, est_id, status_txt, status_bg, borda

    # -------------------------------------------------------------
    # ABA 1: POSIÇÃO DE ESTOQUE
    # -------------------------------------------------------------
    with aba_est1:
        st.markdown("### 📊 Posição do Estoque Físico")
        df_itens = pd.read_sql_query(f'''
        SELECT i.id, i.nome_material, i.categoria, COALESCE(i.tipo_embalagem, 'UNIDADE') as tipo_embalagem,
               COALESCE(i.fator_embalagem, 1.0) as fator_embalagem, i.unidade_medida,
               i.quantidade_atual, i.valor_unitario_estimado,
               (SELECT COALESCE(SUM(m.quantidade), 0.0)
                FROM estoque_movimentacoes m
                WHERE m.item_id = i.id AND m.tipo_movimentacao IN ('SAÍDA', 'HISTÓRICO_RETROATIVO')) as consumo_anual
        FROM estoque_itens i
        WHERE i.om_id = {user['om_id']}
        ORDER BY i.nome_material
        ''', conn)

        if not df_itens.empty:
            df_itens['valor_total_item'] = df_itens['quantidade_atual'] * df_itens['valor_unitario_estimado']
            metricas = [calcular_metricas_item(r['consumo_anual'], r['quantidade_atual']) for _, r in df_itens.iterrows()]
            df_itens['est_minimo'] = [m[0] for m in metricas]
            df_itens['est_ideal'] = [m[1] for m in metricas]
            df_itens['status_txt'] = [m[2] for m in metricas]
            df_itens['status_bg'] = [m[3] for m in metricas]
            df_itens['borda'] = [m[4] for m in metricas]

            total_itens = len(df_itens)
            criticos = len(df_itens[df_itens['quantidade_atual'] <= df_itens['est_minimo']])
            valor_total_deposito = df_itens['valor_total_item'].sum()

            m1, m2, m3 = st.columns(3)
            m1.metric("Materiais Cadastrados", f"{total_itens} itens")
            m2.metric("No Ponto de Pedido (Crítico)", f"{criticos} itens", delta="- Atenção SALC" if criticos > 0 else "Abastecido", delta_color="inverse")
            m3.metric("Patrimônio em Estoque", f"R$ {valor_total_deposito:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            st.markdown("---")
            col_f1, col_f2 = st.columns([3, 1])
            busca_mat = col_f1.text_input("🔍 Buscar Material por Nome:", key="busca_mat_posicao")
            cats = ["Todas"] + sorted(list(set(df_itens['categoria'].dropna().tolist())))
            filtro_c = col_f2.selectbox("Categoria:", cats, key="filtro_cat_posicao")

            df_filtrado = df_itens.copy()
            if busca_mat:
                df_filtrado = df_filtrado[df_filtrado['nome_material'].str.contains(busca_mat, case=False, na=False)]
            if filtro_c != "Todas":
                df_filtrado = df_filtrado[df_filtrado['categoria'] == filtro_c]

            cols_cards = st.columns(3)
            for idx, r in df_filtrado.reset_index(drop=True).iterrows():
                with cols_cards[idx % 3]:
                    fator = r['fator_embalagem']
                    emb_str = f"{r['quantidade_atual'] / fator:.1f} {r['tipo_embalagem']}(s)" if fator > 1 else f"{r['quantidade_atual']:.0f} {r['unidade_medida']}"
                    st.markdown(f'''
                    <div style="background: linear-gradient(145deg, #1e222a, #15181e); border: 1px solid rgba(255,255,255,0.07); {r['borda']} border-radius: 10px; padding: 14px; margin-bottom: 8px; box-shadow: 2px 4px 10px rgba(0,0,0,0.35);">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <span style="font-weight: 800; font-size: 15px; color: #ffffff;">{r['nome_material'][:30]}</span>
                            <span style="font-size: 10px; padding: 2px 7px; border-radius: 4px; background: {r['status_bg']}; color: #fff; font-weight: bold;">{r['status_txt']}</span>
                        </div>
                        <div style="font-size: 12px; color: #94a3b8; margin: 5px 0;">
                            Categoria: <b style="color: #cbd5e1;">{r['categoria']}</b>
                        </div>
                        <div style="background: rgba(255,255,255,0.03); border-radius: 6px; padding: 8px; margin: 6px 0;">
                            <div style="display: flex; justify-content: space-between; font-size: 13px;">
                                <span style="color: #cbd5e1;">Saldo Físico:</span>
                                <b style="color: #22c55e; font-size: 15px;">{r['quantidade_atual']:.0f} {r['unidade_medida']} ({emb_str})</b>
                            </div>
                            <div style="display: flex; justify-content: space-between; font-size: 11px; color: #94a3b8; margin-top: 3px;">
                                <span>Mínimo Calculado (20%): <b>{r['est_minimo']:.0f} {r['unidade_medida']}</b></span>
                                <span>Ideal (50%): <b>{r['est_ideal']:.0f} {r['unidade_medida']}</b></span>
                            </div>
                            <div style="font-size: 11px; color: #60a5fa; margin-top: 2px;">
                                Consumo Registrado no Ano: <b>{r['consumo_anual']:.0f} {r['unidade_medida']}</b>
                            </div>
                        </div>
                        <div style="display: flex; justify-content: space-between; font-size: 12px; border-top: 1px solid rgba(255,255,255,0.06); padding-top: 6px;">
                            <span style="color: #94a3b8;">Base: R$ {r['valor_unitario_estimado']:,.2f}/{r['unidade_medida']} ({f"R$ {r['valor_unitario_estimado']*fator:,.2f}/{r['tipo_embalagem']}" if fator > 1 else ""})</span>
                            <span style="color: #4ade80; font-weight: bold;">Subtotal: R$ {r['valor_total_item']:,.2f}</span>
                        </div>
                    </div>
                    ''', unsafe_allow_html=True)
        else:
            st.info("Nenhum material em estoque. Utilize as abas ao lado para lançar as entradas ou saídas do ano.")

    # -------------------------------------------------------------
    # ABA 2: REGISTRAR ENTRADA
    # -------------------------------------------------------------
    with aba_est2:
        st.markdown("### 📥 Entrada de Material no Depósito")
        modo_entrada = st.radio(
            "Origem do Material:",
            [
                "📦 Saldo Inicial / Já em Estoque (Inventário Físico da OM)",
                "🚚 Nova Compra (Vinculada a Nota de Empenho)"
            ],
            horizontal=True,
            key="radio_modo_entrada_v3"
        )

        ne_origem_txt = "INVENTÁRIO INICIAL"
        if "Nova Compra" in modo_entrada:
            df_nes_disp = pd.read_sql_query(f"SELECT id, numero_ne, fornecedor_nome, valor_ne FROM notas_empenho WHERE om_id = {user['om_id']} ORDER BY id DESC", conn)
            if not df_nes_disp.empty:
                opcoes_ne_dict = {f"{r['numero_ne']} - {r['fornecedor_nome']} (R$ {r['valor_ne']:,.2f})": r['numero_ne'] for _, r in df_nes_disp.iterrows()}
                ne_escolhida = st.selectbox("Selecione a Nota de Empenho:", list(opcoes_ne_dict.keys()), key="sel_ne_entrada_v3")
                ne_origem_txt = opcoes_ne_dict[ne_escolhida]
            else:
                st.warning("Nenhuma NE cadastrada. Registrando como avulso.")

        with st.form("form_entrada_mat_v3"):
            col_sel1, col_sel2 = st.columns(2)
            df_cadastrados = pd.read_sql_query(f"SELECT id, nome_material, unidade_medida, quantidade_atual FROM estoque_itens WHERE om_id = {user['om_id']} ORDER BY nome_material", conn)
            opcoes_itens_disp = ["➕ Cadastrar Novo Material"] + [f"{r['nome_material']} (Saldo: {r['quantidade_atual']:.0f} {r['unidade_medida']})" for _, r in df_cadastrados.iterrows()]

            escolha_item_base = col_sel1.selectbox("Material:", opcoes_itens_disp, key="box_item_entrada_v3")
            eh_novo = (escolha_item_base == "➕ Cadastrar Novo Material")

            if eh_novo:
                nome_material_final = col_sel2.text_input("Nome do Material (ex.: Ração Canina, Detergente Líquido):", value="")
            else:
                nome_material_final = escolha_item_base.split(" (Saldo:")[0]
                col_sel2.info(f"Material Selecionado: **{nome_material_final}**")

            col_cat, col_emb, col_fat, col_uni = st.columns(4)
            cat_escolhida = col_cat.selectbox("Categoria:", [
                "Material de Escritório / Expediente",
                "Material de Limpeza e Higiene",
                "Fardamento, Calçados e EPI",
                "Suprimento de Informática e TI",
                "Canil e Material Veterinário",
                "Peças e Manutenção Automotiva",
                "Material de Construção e Instalações",
                "Material de Polícia do Exército / Bélico"
            ])

            tipo_emb_sel = col_emb.selectbox("Tipo de Embalagem / Apresentação:", ["UNIDADE", "GALÃO", "SACO", "CAIXA", "PACOTE", "RESMA", "LATA", "TAMBOR", "FRASCO"])
            fator_emb = col_fat.number_input("Conteúdo por Embalagem (Fator):", min_value=1.0, value=1.0, step=1.0, help="Exemplo: se for galão de 5 Litros, coloque 5. Se for saco de 15 Kg de ração, coloque 15. Se for resma de 500 folhas, coloque 500.")
            unidade_base = col_uni.selectbox("Unidade de Medida Base:", ["UN", "L", "KG", "M", "PAR"])

            col_q_ent1, col_q_ent2, col_q_ent3 = st.columns(3)
            qtd_embalagens_ent = col_q_ent1.number_input(f"Quantidade de {tipo_emb_sel}(S) Recebidos:", min_value=1.0, value=1.0, step=1.0)
            total_unidades_base = qtd_embalagens_ent * fator_emb
            col_q_ent2.markdown(f"<div style='padding-top: 25px; font-weight: bold; color: #4ade80;'>= {total_unidades_base:.0f} {unidade_base} no total</div>", unsafe_allow_html=True)

            preco_embalagem_fechada = col_q_ent3.number_input(f"Preço da Embalagem Fechada (R$ por {tipo_emb_sel}):", min_value=0.01, value=10.0, step=1.0, format="%.2f")
            # O preço unitário base (por Litro ou Kg) é o preço da embalagem dividido pelo conteúdo
            val_unit_base = preco_embalagem_fechada / fator_emb if fator_emb > 0 else preco_embalagem_fechada
            st.caption(f"💡 Preço da Embalagem: **R$ {preco_embalagem_fechada:,.2f}** por {tipo_emb_sel} (equivale a **R$ {val_unit_base:,.2f}** por {unidade_base}) | Valor Total do Lote: **R$ {(qtd_embalagens_ent * preco_embalagem_fechada):,.2f}**")

            col_dt_e, col_obs_e = st.columns(2)
            data_recebimento = col_dt_e.date_input("Data da Entrada:", value=datetime.now(), format="DD/MM/YYYY")
            obs_entrada = col_obs_e.text_input("Observação:", value=f"Recebido conf. {ne_origem_txt}")

            if st.form_submit_button("✅ Gravar Entrada no Estoque", type="primary"):
                if not nome_material_final.strip():
                    st.error("Informe o nome do material.")
                else:
                    cur = conn.cursor()
                    cur.execute("SELECT id, quantidade_atual FROM estoque_itens WHERE om_id = ? AND nome_material = ?", (user['om_id'], nome_material_final.strip()))
                    item_existente = cur.fetchone()

                    if item_existente:
                        id_item = item_existente[0]
                        saldo_ant = float(item_existente[1])
                        saldo_novo = saldo_ant + total_unidades_base
                        cur.execute('''
                        UPDATE estoque_itens 
                        SET quantidade_atual = ?, valor_unitario_estimado = ?, tipo_embalagem = ?, fator_embalagem = ?, unidade_medida = ?
                        WHERE id = ?
                        ''', (saldo_novo, val_unit_base, tipo_emb_sel, fator_emb, unidade_base, id_item))
                    else:
                        saldo_ant = 0.0
                        saldo_novo = total_unidades_base
                        cur.execute('''
                        INSERT INTO estoque_itens (om_id, nome_material, categoria, tipo_embalagem, fator_embalagem, 
                                                   unidade_medida, quantidade_atual, estoque_minimo, estoque_ideal, valor_unitario_estimado)
                        VALUES (?, ?, ?, ?, ?, ?, ?, 0.0, 0.0, ?)
                        ''', (user['om_id'], nome_material_final.strip(), cat_escolhida, tipo_emb_sel, fator_emb,
                              unidade_base, saldo_novo, val_unit_base))
                        id_item = cur.lastrowid

                    cur.execute('''
                    INSERT INTO estoque_movimentacoes (om_id, item_id, tipo_movimentacao, quantidade, saldo_anterior,
                                                       saldo_novo, data_movimentacao, secao_destino, militar_responsavel,
                                                       ne_origem, observacao)
                    VALUES (?, ?, 'ENTRADA', ?, ?, ?, ?, 'ALMOXARIFADO CENTRAL', ?, ?, ?)
                    ''', (user['om_id'], id_item, total_unidades_base, saldo_ant, saldo_novo, str(data_recebimento),
                          f"{user['posto_grad']} {user['nome_guerra']}", ne_origem_txt, obs_entrada))
                    conn.commit()

                    st.success(f"✅ Entrada registrada: {total_unidades_base:.0f} {unidade_base} ({qtd_embalagens_ent:.0f} {tipo_emb_sel}) de '{nome_material_final}'. Novo saldo: {saldo_novo:.0f} {unidade_base}.")
                    st.rerun()

    # -------------------------------------------------------------
    # ABA 3: REGISTRAR SAÍDA (COM RETROATIVO DO ANO & 17 SEÇÕES)
    # -------------------------------------------------------------
    with aba_est3:
        st.markdown("### 📤 Registro de Saída / Consumo das Seções")

        is_retroativo = st.checkbox(
            "🗓️ Modo Histórico / Retroativo do Ano: Lançar consumo ocorrido em 2026 sem travar por falta de saldo atual (para levantar a média anual)",
            value=False,
            key="chk_modo_retroativo_v3"
        )
        if is_retroativo:
            st.info("💡 **Modo Histórico Ativado**: Lance as saídas e consumos das seções deste ano. O sistema computará tudo no cálculo anual do Estoque Mínimo sem exigir estoque prévio.")

        query_mat_saida = f"SELECT id, nome_material, unidade_medida, quantidade_atual, tipo_embalagem, fator_embalagem FROM estoque_itens WHERE om_id = {user['om_id']}"
        if not is_retroativo:
            query_mat_saida += " AND quantidade_atual > 0"
        query_mat_saida += " ORDER BY nome_material"

        df_mats_saida = pd.read_sql_query(query_mat_saida, conn)

        # Seções da OM
        c = conn.cursor()
        c.execute("SELECT nome_secao FROM estoque_secoes WHERE om_id = ? ORDER BY nome_secao", (user['om_id'],))
        secoes_disp = [r[0] for r in c.fetchall()]

        if not secoes_disp:
            secoes_iniciais_padrao = [
                "1ª Seção", "2ª Seção", "3ª Seção", "4ª Seção",
                "Comandante", "Subcomandante", "1º Pelotão", "2º Pelotão",
                "Pelotão de Escolta e Guarda", "Seção de Cães de Guerra",
                "Garagem de Auto", "Garagem de Motos", "Reserva de Armamento",
                "Reserva de Material", "RP (Relações Públicas)",
                "Seção de Informática / Com", "Secretaria"
            ]
            for sec in secoes_iniciais_padrao:
                c.execute("INSERT OR IGNORE INTO estoque_secoes (om_id, nome_secao) VALUES (?, ?)", (user['om_id'], sec))
            conn.commit()
            secoes_disp = secoes_iniciais_padrao

        if not df_mats_saida.empty or is_retroativo:
            with st.form("form_saida_material_v3"):
                if not df_mats_saida.empty:
                    dict_mats_saida = {f"{r['nome_material']} (Saldo Atual: {r['quantidade_atual']:.0f} {r['unidade_medida']})": r for _, r in df_mats_saida.iterrows()}
                    escolha_mat = st.selectbox("Selecione o Material:", list(dict_mats_saida.keys()), key="box_sel_mat_saida_v3")
                    mat_dados = dict_mats_saida[escolha_mat]
                    id_mat_escolhido = int(mat_dados['id'])
                    unidade_saida = mat_dados['unidade_medida']
                    fator_s = float(mat_dados['fator_embalagem'] or 1.0)
                    tipo_emb_s = mat_dados['tipo_embalagem'] or 'UNIDADE'
                    saldo_disp = float(mat_dados['quantidade_atual'])
                else:
                    nome_mat_avulso = st.text_input("Nome do Material Consumido:", value="Material de Consumo do Ano")
                    id_mat_escolhido = None
                    unidade_saida = "UN"
                    fator_s = 1.0
                    tipo_emb_s = "UNIDADE"
                    saldo_disp = 0.0

                col_saida_tipo, col_saida_qtd = st.columns(2)
                opcoes_modo_saida = [f"Unidade Base ({unidade_saida})"]
                if fator_s > 1:
                    opcoes_modo_saida.append(f"Embalagem Fechada ({tipo_emb_s})")
                modo_qtd_saida = col_saida_tipo.radio("Informar Retirada em:", opcoes_modo_saida, horizontal=True)

                max_permitido = None if is_retroativo else (saldo_disp if saldo_disp > 0 else 1.0)
                if "Embalagem" in modo_qtd_saida and fator_s > 1:
                    qtd_digitada = col_saida_qtd.number_input(f"Quantidade de {tipo_emb_s}(s) Retirados:", min_value=1.0, value=1.0, step=1.0)
                    qtd_base_final = qtd_digitada * fator_s
                    st.caption(f"Equivalente a: **{qtd_base_final:.0f} {unidade_saida}**")
                else:
                    qtd_base_final = col_saida_qtd.number_input(
                        f"Quantidade Retirada ({unidade_saida}):",
                        min_value=1.0,
                        max_value=max_permitido,
                        value=1.0,
                        step=1.0
                    )

                col_sec1, col_sec2 = st.columns(2)
                opcoes_secoes_dropdown = secoes_disp + ["➕ Cadastrar Nova Seção da OM..."]
                secao_selecionada = col_sec1.selectbox("Seção / Pelotão Solicitante:", opcoes_secoes_dropdown, key="sel_secao_destino_v3")

                secao_dest_final = secao_selecionada
                if secao_selecionada == "➕ Cadastrar Nova Seção da OM...":
                    nova_secao_input = st.text_input("Digite o Nome da Nova Seção:", key="input_nova_secao_dinamica_v3")
                    if nova_secao_input.strip():
                        secao_dest_final = nova_secao_input.strip()

                data_saida_reg = col_sec2.date_input("Data da Saída / Consumo:", value=datetime.now(), format="DD/MM/YYYY")

                col_resp1, col_resp2 = st.columns(2)
                militar_resp = col_resp1.text_input("Militar que Retirou (Posto/Grad e Nome):", value="3º Sgt Aluno / Cb Furriel")
                motivo_saida = col_resp2.text_input("Finalidade / Aplicação:", value="Consumo de expediente / rotina da seção")

                if st.form_submit_button("📤 Confirmar Saída no Estoque", type="primary"):
                    if not secao_dest_final or secao_dest_final.startswith("➕"):
                        st.error("Selecione ou digite uma seção de destino válida.")
                    else:
                        cur = conn.cursor()
                        cur.execute("INSERT OR IGNORE INTO estoque_secoes (om_id, nome_secao) VALUES (?, ?)", (user['om_id'], secao_dest_final))

                        if id_mat_escolhido is None:
                            cur.execute('''
                            INSERT INTO estoque_itens (om_id, nome_material, categoria, tipo_embalagem, fator_embalagem, unidade_medida, quantidade_atual, estoque_minimo, estoque_ideal, valor_unitario_estimado)
                            VALUES (?, ?, 'Material do Ano', 'UNIDADE', 1.0, 'UN', 0.0, 0.0, 0.0, 10.0)
                            ''', (user['om_id'], nome_mat_avulso.strip()))
                            id_mat_escolhido = cur.lastrowid
                            saldo_disp = 0.0

                        tipo_op = 'HISTÓRICO_RETROATIVO' if is_retroativo else 'SAÍDA'
                        novo_saldo = max(saldo_disp - qtd_base_final, 0.0) if not is_retroativo else saldo_disp

                        if not is_retroativo:
                            cur.execute("UPDATE estoque_itens SET quantidade_atual = ? WHERE id = ?", (novo_saldo, id_mat_escolhido))

                        cur.execute('''
                        INSERT INTO estoque_movimentacoes (om_id, item_id, tipo_movimentacao, quantidade, saldo_anterior,
                                                           saldo_novo, data_movimentacao, secao_destino, militar_responsavel,
                                                           ne_origem, observacao)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, '-', ?)
                        ''', (user['om_id'], id_mat_escolhido, tipo_op, qtd_base_final, saldo_disp, novo_saldo,
                              str(data_saida_reg), secao_dest_final, militar_resp, motivo_saida))
                        conn.commit()

                        msg_tipo = "Consumo histórico de 2026 registrado com sucesso!" if is_retroativo else f"Saída de {qtd_base_final:.0f} {unidade_saida} confirmada! Novo saldo: {novo_saldo:.0f}."
                        st.success(f"✅ {msg_tipo}")
                        st.rerun()
        else:
            st.warning("Não há materiais em estoque para saída. Ative o 'Modo Histórico / Retroativo' para lançar os consumos do ano.")

    # -------------------------------------------------------------
    # ABA 4: HISTÓRICO & MEMÓRIA DE CÁLCULO
    # -------------------------------------------------------------
    with aba_est4:
        sub1, sub2 = st.tabs(["📜 Livro de Movimentações", "🧮 Memória de Cálculo Automática para Compras (SALC)"])

        with sub1:
            st.markdown("#### 📜 Livro Geral de Entradas e Saídas")
            df_movs_todas = pd.read_sql_query(f'''
            SELECT m.id, m.data_movimentacao, m.tipo_movimentacao, i.nome_material, m.quantidade,
                   i.unidade_medida, m.saldo_anterior, m.saldo_novo, m.secao_destino,
                   m.militar_responsavel, m.observacao
            FROM estoque_movimentacoes m
            JOIN estoque_itens i ON m.item_id = i.id
            WHERE m.om_id = {user['om_id']}
            ORDER BY m.id DESC
            ''', conn)

            if not df_movs_todas.empty:
                df_movs_todas['Data'] = df_movs_todas['data_movimentacao'].apply(formatar_data_br)
                df_exib_mov = df_movs_todas[['Data', 'tipo_movimentacao', 'nome_material', 'quantidade', 'unidade_medida', 'saldo_anterior', 'saldo_novo', 'secao_destino', 'militar_responsavel', 'observacao']]
                st.dataframe(df_exib_mov, use_container_width=True, hide_index=True)

                col_dl_mv1, col_dl_mv2, col_dl_mv3 = st.columns(3)
                with col_dl_mv1:
                    st.download_button("📥 Baixar CSV", data=df_exib_mov.to_csv(index=False).encode('utf-8'), file_name="livro_movimentacoes_5ciape.csv", mime="text/csv", use_container_width=True)
                with col_dl_mv2:
                    buf_mv_excel = exportar_excel_seguro(df_exib_mov, 'Movimentações')
                    if buf_mv_excel:
                        st.download_button("📊 Baixar Planilha (.xlsx)", data=buf_mv_excel, file_name="livro_movimentacoes_5ciape.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                with col_dl_mv3:
                    buf_pdf_mov = gerar_pdf_tabela_seguro("LIVRO REGISTRO DE MOVIMENTAÇÕES DE MATERIAL", df_exib_mov)
                    if buf_pdf_mov:
                        st.download_button("📄 Baixar em PDF (.pdf)", data=buf_pdf_mov, file_name="livro_movimentacoes_5ciape.pdf", mime="application/pdf", use_container_width=True)
            else:
                st.info("Nenhuma movimentação registrada no histórico ainda.")

        with sub2:
            st.markdown("#### 🧮 Memória de Cálculo Dinâmica para o Plano de Contratações")
            st.caption("Base regulamentar: Estoque Mínimo (20% do Consumo Anual) | Estoque Ideal (50% do Consumo Anual para 6 meses de suprimento).")

            df_base_calc = pd.read_sql_query(f'''
            SELECT i.id, i.nome_material, i.categoria, i.unidade_medida, i.quantidade_atual, i.valor_unitario_estimado,
                   (SELECT COALESCE(SUM(m.quantidade), 0.0)
                    FROM estoque_movimentacoes m
                    WHERE m.item_id = i.id AND m.tipo_movimentacao IN ('SAÍDA', 'HISTÓRICO_RETROATIVO')) as consumo_anual_total
            FROM estoque_itens i
            WHERE i.om_id = {user['om_id']}
            ORDER BY i.nome_material
            ''', conn)

            if not df_base_calc.empty:
                resultados = []
                for _, r in df_base_calc.iterrows():
                    c_ano = r['consumo_anual_total']
                    saldo = r['quantidade_atual']
                    preco = r['valor_unitario_estimado']

                    est_min = round(c_ano * 0.20, 1) if c_ano > 0 else round(saldo * 0.20, 1)
                    est_id = round(c_ano * 0.50, 1) if c_ano > 0 else round(saldo, 1)

                    qtd_comprar = max(est_id - saldo, 0.0)
                    custo_previsto = qtd_comprar * preco

                    resultados.append({
                        "Material": r['nome_material'],
                        "Categoria": r['categoria'],
                        "UN": r['unidade_medida'],
                        "Consumo Registrado no Ano": f"{c_ano:,.0f}".replace(",", "."),
                        "Estoque Atual em Prateleira": f"{saldo:,.0f}".replace(",", "."),
                        "Estoque Mínimo (20%)": f"{est_min:,.0f}".replace(",", "."),
                        "Estoque Ideal (50%)": f"{est_id:,.0f}".replace(",", "."),
                        "Preço Unitário Estimado": f"R$ {preco:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                        "Qtd a Comprar (Ressuprimento)": f"{qtd_comprar:,.0f}".replace(",", "."),
                        "Custo Estimado da Compra": f"R$ {custo_previsto:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                        "valor_bruto_compra": custo_previsto
                    })

                df_calc_final = pd.DataFrame(resultados)
                total_geral_licitacao = sum(r['valor_bruto_compra'] for r in resultados)

                st.markdown(f'''
                <div style="background: rgba(34, 197, 94, 0.15); border: 1.5px solid #22c55e; border-radius: 8px; padding: 14px 20px; display: flex; justify-content: space-between; align-items: center; margin: 15px 0;">
                    <div>
                        <span style="font-size: 14px; font-weight: bold; color: #86efac; text-transform: uppercase;">Valor Total Previsto para a Próxima Licitação / Compra:</span><br>
                        <span style="font-size: 12px; color: #cbd5e1;">Base de cálculo: Cobertura semestral (6 meses) para as frações da 5ª Cia PE</span>
                    </div>
                    <span style="font-size: 26px; font-weight: 900; color: #22c55e;">R$ {total_geral_licitacao:,.2f}</span>
                </div>
                '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)

                st.dataframe(df_calc_final.drop(columns=['valor_bruto_compra']), use_container_width=True, hide_index=True)

                col_dl_m1, col_dl_m2, col_dl_m3 = st.columns(3)
                df_export_salc = df_calc_final.drop(columns=['valor_bruto_compra'])
                with col_dl_m1:
                    st.download_button(
                        "📥 Baixar CSV",
                        data=df_export_salc.to_csv(index=False).encode('utf-8'),
                        file_name="memoria_calculo_compras_5ciape.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                with col_dl_m2:
                    buf_calc_salc = exportar_excel_seguro(df_export_salc, 'Memória de Cálculo')
                    if buf_calc_salc:
                        st.download_button(
                            "📊 Baixar Planilha (.xlsx)",
                            data=buf_calc_salc,
                            file_name="memoria_calculo_compras_5ciape.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                with col_dl_m3:
                    buf_pdf_salc = gerar_pdf_tabela_seguro("MEMÓRIA DE CÁLCULO E PROJEÇÃO DE COMPRAS DA OM", df_export_salc)
                    if buf_pdf_salc:
                        st.download_button(
                            "📄 Baixar em PDF (.pdf)",
                            data=buf_pdf_salc,
                            file_name="memoria_calculo_compras_5ciape.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
            else:
                st.info("Lance as saídas de materiais ao longo do ano para gerar a memória de cálculo.")

    conn.close()
