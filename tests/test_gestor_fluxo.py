import unittest
import sqlite3
import os
import sys
from datetime import datetime, timedelta

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from core.gestor_fluxo import GestorFluxoMilitar

class TestGestorFluxo(unittest.TestCase):
    """Testes unitários para as regras de negócio orçamentárias e fluxo operacional."""

    def setUp(self):
        self.conn = sqlite3.connect(":memory:")
        self.conn.row_factory = sqlite3.Row
        c = self.conn.cursor()

        # Cria esquema isolado em memória
        c.execute('''
        CREATE TABLE notas_credito (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            om_id INTEGER NOT NULL,
            numero_nc TEXT NOT NULL,
            valor_total REAL NOT NULL,
            valor_recolhido REAL DEFAULT 0.0,
            data_limite_empenho DATE,
            finalidade TEXT,
            enquadramento TEXT
        )
        ''')

        c.execute('''
        CREATE TABLE notas_empenho (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nc_id INTEGER NOT NULL,
            nc_id_2 INTEGER,
            valor_nc_1 REAL,
            valor_nc_2 REAL DEFAULT 0.0,
            om_id INTEGER NOT NULL,
            numero_ne TEXT NOT NULL,
            valor_ne REAL NOT NULL,
            fornecedor_nome TEXT NOT NULL,
            fornecedor_telefone TEXT,
            data_limite DATE,
            nova_data_limite DATE,
            status TEXT DEFAULT 'Aguardando Entrega'
        )
        ''')
        self.conn.commit()

    def tearDown(self):
        self.conn.close()

    def test_saldo_nc_fonte_unica(self):
        """Verifica o cálculo de saldo disponível para empenhos com uma única NC de origem."""
        c = self.conn.cursor()
        c.execute("INSERT INTO notas_credito (om_id, numero_nc, valor_total, valor_recolhido, data_limite_empenho, finalidade, enquadramento) VALUES (1, '2026NC001', 100000.0, 10000.0, '2026-12-31', 'Material', 'DISCRICIONÁRIA')")
        nc_id = c.lastrowid

        c.execute("INSERT INTO notas_empenho (nc_id, om_id, numero_ne, valor_ne, valor_nc_1, fornecedor_nome) VALUES (?, 1, '2026NE001', 30000.0, 30000.0, 'Empresa A')", (nc_id,))
        self.conn.commit()

        info = GestorFluxoMilitar.obter_saldo_nc(nc_id, self.conn)
        self.assertIsNotNone(info)
        self.assertEqual(info['valor_total'], 100000.0)
        self.assertEqual(info['total_empenhado'], 30000.0)
        self.assertEqual(info['valor_recolhido'], 10000.0)
        # 100.000 - 30.000 - 10.000 = 60.000
        self.assertEqual(info['saldo_disponivel'], 60000.0)

    def test_saldo_nc_fonte_dupla(self):
        """Verifica o cálculo correto de saldo quando um empenho utiliza duas NCs simultaneamente."""
        c = self.conn.cursor()
        c.execute("INSERT INTO notas_credito (om_id, numero_nc, valor_total, valor_recolhido, data_limite_empenho, finalidade, enquadramento) VALUES (1, 'NC_PRINCIPAL', 50000.0, 0.0, '2026-12-31', 'Objeto 1', 'DISCRICIONÁRIA')")
        nc1_id = c.lastrowid

        c.execute("INSERT INTO notas_credito (om_id, numero_nc, valor_total, valor_recolhido, data_limite_empenho, finalidade, enquadramento) VALUES (1, 'NC_COMPLEMENTAR', 20000.0, 0.0, '2026-12-31', 'Objeto 2', 'DISCRICIONÁRIA')")
        nc2_id = c.lastrowid

        # Empenho misto de R$ 35.000: R$ 25.000 da NC 1 e R$ 10.000 da NC 2
        c.execute("""
        INSERT INTO notas_empenho (nc_id, nc_id_2, valor_nc_1, valor_nc_2, om_id, numero_ne, valor_ne, fornecedor_nome)
        VALUES (?, ?, 25000.0, 10000.0, 1, 'NE_MISTA', 35000.0, 'Fornecedor Beta')
        """, (nc1_id, nc2_id))
        self.conn.commit()

        info_nc1 = GestorFluxoMilitar.obter_saldo_nc(nc1_id, self.conn)
        info_nc2 = GestorFluxoMilitar.obter_saldo_nc(nc2_id, self.conn)

        # NC 1: 50.000 - 25.000 = 25.000
        self.assertEqual(info_nc1['saldo_disponivel'], 25000.0)
        # NC 2: 20.000 - 10.000 = 10.000
        self.assertEqual(info_nc2['saldo_disponivel'], 10000.0)

    def test_radar_prazos_nc(self):
        """Testa o semáforo de urgência de NCs a vencer ou vencidas com saldo."""
        hoje = datetime.now().date()
        c = self.conn.cursor()

        # NC vencida com saldo
        dt_vencida = str(hoje - timedelta(days=3))
        c.execute("INSERT INTO notas_credito (om_id, numero_nc, valor_total, valor_recolhido, data_limite_empenho, finalidade, enquadramento) VALUES (1, 'NC_VENCIDA', 10000.0, 0.0, ?, 'Teste Vencido', 'DISC')", (dt_vencida,))

        # NC no prazo normal (30 dias)
        dt_longe = str(hoje + timedelta(days=30))
        c.execute("INSERT INTO notas_credito (om_id, numero_nc, valor_total, valor_recolhido, data_limite_empenho, finalidade, enquadramento) VALUES (1, 'NC_TRANQUILA', 10000.0, 0.0, ?, 'Teste Tranquilo', 'DISC')", (dt_longe,))
        self.conn.commit()

        alertas = GestorFluxoMilitar.listar_alertas_ncs_criticas(1, dias_limite=7, conn=self.conn)
        self.assertEqual(len(alertas), 1)
        self.assertEqual(alertas[0]['numero_nc'], 'NC_VENCIDA')
        self.assertEqual(alertas[0]['gravidade'], 'VENCIDA')

    def test_metricas_suprimento_militar(self):
        """Testa o cálculo regulamentar de estoque mínimo (20%) e estoque ideal (50%)."""
        # Consumo anual de 500 unidades e saldo atual de 80
        consumo_anual = 500.0
        saldo_atual = 80.0

        metricas = GestorFluxoMilitar.calcular_metricas_estoque(consumo_anual, saldo_atual)

        # Estoque mínimo esperado = 20% de 500 = 100
        self.assertEqual(metricas['estoque_minimo'], 100.0)
        # Estoque ideal esperado = 50% de 500 = 250
        self.assertEqual(metricas['estoque_ideal'], 250.0)
        # Como saldo atual (80) <= mínimo (100) -> CRÍTICO / REPOR
        self.assertIn("CRÍTICO", metricas['status_txt'])
        # Quantidade a comprar para atingir o ideal = 250 - 80 = 170
        self.assertEqual(metricas['qtd_a_comprar'], 170.0)

if __name__ == '__main__':
    unittest.main()
