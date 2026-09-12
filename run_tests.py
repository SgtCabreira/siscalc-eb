import unittest
import sys
import os
import time

def executar_bateria_testes():
    print("=" * 70)
    print("🛡️  SISCALC / SISLOG EB — BATERIA DE TESTES AUTOMATIZADOS")
    print("5ª Companhia de Polícia do Exército — Base Major Agostinho José Rodrigues")
    print("=" * 70)
    print(f"Data e Hora: {time.strftime('%d/%m/%Y %H:%M:%S')}")
    print("Iniciando varredura na pasta 'tests/'...\n")

    inicio = time.time()
    loader = unittest.TestLoader()
    suite = loader.discover(start_dir=os.path.join(os.path.dirname(__file__), 'tests'), pattern='test_*.py')

    runner = unittest.TextTestRunner(verbosity=2)
    resultado = runner.run(suite)

    duracao = time.time() - inicio
    print("\n" + "=" * 70)
    print("📊 RELATÓRIO FINAL DA BATERIA DE TESTES:")
    print(f" • Total de Testes Executados: {resultado.testsRun}")
    print(f" • Falhas: {len(resultado.failures)}")
    print(f" • Erros: {len(resultado.errors)}")
    print(f" • Tempo de Execução: {duracao:.3f} segundos")

    if resultado.wasSuccessful():
        print("\n✅ STATUS: TODOS OS TESTES PASSARAM COM 100% DE SUCESSO!")
        print("A integridade das regras de negócio, banco e documentos está garantida.")
        print("=" * 70)
        return 0
    else:
        print("\n❌ STATUS: FORAM DETECTADAS INCONFORMIDADES!")
        print("=" * 70)
        return 1

if __name__ == '__main__':
    sys.exit(executar_bateria_testes())
