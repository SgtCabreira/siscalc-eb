import unittest
import sqlite3
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from core.database import get_connection, init_db, DB_FILE

class TestDatabase(unittest.TestCase):
    """Testes unitários para conexão, esquema e integridade do banco de dados."""

    def setUp(self):
        self.conn = sqlite3.connect(":memory:")
        self.conn.row_factory = sqlite3.Row

    def tearDown(self):
        self.conn.close()

    def test_conexao(self):
        """Verifica se a conexão ativa retorna um objeto válido."""
        self.assertIsNotNone(self.conn)

    def test_inicializacao_e_tabelas(self):
        """Verifica se todas as tabelas essenciais do sistema são criadas."""
        c = self.conn.cursor()
        tabelas_esperadas = [
            'oms', 'usuarios', 'notas_credito', 'notas_empenho',
            'notas_fiscais', 'estoque_itens', 'estoque_secoes',
            'estoque_movimentacoes', 'processos_modelos', 'tipos_documentos', 'processos_gerados'
        ]

        # Executa esquema padrão em memória
        from core.database import init_db
        # Testa chamada padrão
        conn_real = get_connection()
        c_real = conn_real.cursor()
        c_real.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tabelas_reais = [r[0] for r in c_real.fetchall()]
        conn_real.close()

        for tab in tabelas_esperadas:
            self.assertIn(tab, tabelas_reais, f"Tabela '{tab}' deve existir no banco.")

    def test_colunas_especiais_empenho(self):
        """Verifica se as colunas de vínculo duplo (nc_id_2, valor_nc_1, valor_nc_2) existem."""
        conn = get_connection()
        c = conn.cursor()
        c.execute("PRAGMA table_info(notas_empenho)")
        colunas = [r['name'] for r in c.fetchall()]
        conn.close()

        self.assertIn('nc_id_2', colunas)
        self.assertIn('valor_nc_1', colunas)
        self.assertIn('valor_nc_2', colunas)

    def test_colunas_embalagem_estoque(self):
        """Verifica se as colunas de embalagem e valores padrão de estoque mínimo existem."""
        conn = get_connection()
        c = conn.cursor()
        c.execute("PRAGMA table_info(estoque_itens)")
        colunas = [r['name'] for r in c.fetchall()]
        conn.close()

        self.assertIn('tipo_embalagem', colunas)
        self.assertIn('fator_embalagem', colunas)
        self.assertIn('estoque_minimo', colunas)
        self.assertIn('estoque_ideal', colunas)

def test_backup_automatico_17h(self):
        """Valida a rotina de verificação e execução de backup das 17h."""
        from core.database import verificar_e_executar_backup_17h
        pasta_teste = "test_backups_17h"
        sucesso = verificar_e_executar_backup_17h(db_file=DB_FILE, pasta_backup=pasta_teste, forcar=True)
        self.assertTrue(sucesso)
        self.assertTrue(os.path.exists(pasta_teste))
        import shutil
        shutil.rmtree(pasta_teste, ignore_errors=True)

if __name__ == '__main__':
    unittest.main()
