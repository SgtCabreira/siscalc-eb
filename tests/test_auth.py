import unittest
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from core.auth import obter_menus_autorizados

class TestAuth(unittest.TestCase):
    """Testes para permissões de perfis de acesso do sistema."""

    def test_permissoes_administrador(self):
        """Administrador deve ter acesso irrestrito a todos os menus, inclusive Admin."""
        menus = obter_menus_autorizados("Administrador")
        self.assertIn("📊 Dashboard", menus)
        self.assertIn("📑 Notas de Crédito", menus)
        self.assertIn("📋 Notas de Empenho", menus)
        self.assertIn("📝 Elaboração de Processos", menus)
        self.assertIn("🏢 Gestão de Estoque", menus)
        self.assertIn("⚙️ Admin", menus)

    def test_permissoes_salc(self):
        """Operador da SALC não deve ter acesso ao menu Admin do sistema."""
        menus = obter_menus_autorizados("SALC")
        self.assertIn("📑 Notas de Crédito", menus)
        self.assertIn("📋 Notas de Empenho", menus)
        self.assertIn("📝 Elaboração de Processos", menus)
        self.assertNotIn("⚙️ Admin", menus)

    def test_permissoes_almoxarifado(self):
        """Almoxarife deve acessar Estoque e Recebimento, mas não Notas de Crédito nem Admin."""
        menus = obter_menus_autorizados("Almoxarife")
        self.assertIn("📦 Recebimento", menus)
        self.assertIn("🏢 Gestão de Estoque", menus)
        self.assertNotIn("📑 Notas de Crédito", menus)
        self.assertNotIn("⚙️ Admin", menus)

if __name__ == '__main__':
    unittest.main()
