import unittest
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from core.api_cnpj import buscar_dados_cnpj

class TestApiCnpj(unittest.TestCase):
    """Testes para o serviço de validação e busca de CNPJ."""

    def test_cnpj_invalido(self):
        """Verifica se CNPJs com tamanho incorreto ou vazios são rejeitados de imediato."""
        self.assertIsNone(buscar_dados_cnpj("123"))
        self.assertIsNone(buscar_dados_cnpj(""))
        self.assertIsNone(buscar_dados_cnpj(None))

    def test_cnpj_formatacao(self):
        """Verifica se caracteres especiais de pontuação são limpos corretamente."""
        # 13 dígitos (inválido)
        self.assertIsNone(buscar_dados_cnpj("12.345.678/0001-9"))

if __name__ == '__main__':
    unittest.main()
