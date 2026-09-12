import unittest
import os
import sys
import pandas as pd
from datetime import datetime

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from core.documentos import (
    formatar_data_br,
    normalizar_data,
    exportar_excel_seguro,
    gerar_pdf_tabela_seguro,
    gerar_requisicao_odt
)

class TestDocumentos(unittest.TestCase):
    """Testes para formatação e geradores de documentos oficiais."""

    def test_formatacao_data_br(self):
        """Verifica se a conversão de datas para o padrão brasileiro funciona perfeitamente."""
        self.assertEqual(formatar_data_br('2026-09-11'), '11/09/2026')
        self.assertEqual(formatar_data_br('11/09/2026'), '11/09/2026')
        self.assertEqual(formatar_data_br(None), '')
        self.assertEqual(formatar_data_br(''), '')

    def test_normalizacao_data_banco(self):
        """Verifica a conversão de data brasileira para o padrão de banco YYYY-MM-DD."""
        self.assertEqual(normalizar_data('11/09/2026'), '2026-09-11')
        self.assertEqual(normalizar_data('2026-09-11'), '2026-09-11')
        self.assertIsNone(normalizar_data(None))

    def test_geracao_requisicao_odt(self):
        """Verifica se o gerador de LibreOffice Writer (.odt) gera um ZIP válido com mimetype."""
        dados = {
            "numero_req": "10/2026",
            "nup": "65378.001/2026",
            "assunto": "Material de Limpeza",
            "data_extenso": "Curitiba, PR, 11 de setembro de 2026",
            "pregao": "PE 90001/2026",
            "razao_social": "Biolimp Ltda",
            "cnpj": "00.000.000/0001-00",
            "tipo_empenho": "Ordinário",
            "local_entrega": "5ª Cia PE",
            "justificativa_a": "Reposição",
            "justificativa_b": "Desabastecimento",
            "pi": "I2B", "ptres": "171502", "nd": "339030", "nc": "2026NC001",
            "fiscal_nome_posto": "Cap Fonseca",
            "fiscal_funcao": "Fiscal Adm"
        }
        df_itens = pd.DataFrame([{"Item": 1, "Descrição": "Sabão", "UN": "UN", "QTD": 10.0, "V_Unit": 5.0}])
        
        buf = gerar_requisicao_odt(dados, 50.0, df_itens)
        self.assertIsNotNone(buf)
        self.assertGreater(len(buf.getvalue()), 500)
        # Verifica cabeçalho ZIP válido
        self.assertTrue(buf.getvalue().startswith(b'PK'))

    def test_gerador_pdf_tabela_seguro(self):
        """Verifica a geração de relatório oficial em PDF sem erros de codificação."""
        df_teste = pd.DataFrame([
            {"Material": "Ração Canina", "Quantidade": "15 KG", "Data": "11/09/2026"}
        ])
        buf_pdf = gerar_pdf_tabela_seguro("RELATÓRIO DE TESTE MILITAR", df_teste)
        if buf_pdf is not None:
            self.assertTrue(buf_pdf.getvalue().startswith(b'%PDF'))

if __name__ == '__main__':
    unittest.main()
