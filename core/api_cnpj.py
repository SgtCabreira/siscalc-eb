import re
import urllib.request
import json

def buscar_dados_cnpj(cnpj, conn=None):
    """
    Busca os dados cadastrais da empresa pelo CNPJ.
    1. Procura primeiro no histórico local do banco de dados (rápido e offline).
    2. Se não encontrar, consulta a API pública oficial da BrasilAPI (Receita Federal).
    """
    cnpj_limpo = re.sub(r'\D', '', str(cnpj))
    if len(cnpj_limpo) != 14:
        return None

    # 1. Consulta no cache local se a conexão com o banco foi fornecida
    if conn is not None:
        try:
            cur = conn.cursor()
            cur.execute('''
            SELECT fornecedor_nome, fornecedor_email, fornecedor_telefone, fornecedor_cidade, fornecedor_uf, fornecedor_situacao 
            FROM notas_empenho 
            WHERE fornecedor_cnpj LIKE ? AND fornecedor_nome IS NOT NULL AND fornecedor_nome != ''
            LIMIT 1
            ''', (f"%{cnpj_limpo}%",))
            local = cur.fetchone()
            if local and local[0]:
                return {
                    "razao_social": local[0],
                    "email": local[1] or "",
                    "telefone": local[2] or "",
                    "cidade": local[3] or "",
                    "uf": local[4] or "",
                    "situacao": local[5] or "ATIVA"
                }
        except Exception:
            pass

    # 2. Consulta externa na BrasilAPI (Receita Federal)
    try:
        url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj_limpo}"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (SisCalc-EB)'})
        with urllib.request.urlopen(req, timeout=4) as response:
            dados = json.loads(response.read().decode('utf-8'))
            tel_bruto = str(dados.get('ddd_telefone_1') or dados.get('ddd_telefone_2') or '')
            tel_limpo = re.sub(r'\D', '', tel_bruto)
            tel_fmt = tel_limpo
            if len(tel_limpo) == 10:
                tel_fmt = f"({tel_limpo[:2]}) {tel_limpo[2:6]}-{tel_limpo[6:]}"
            elif len(tel_limpo) == 11:
                tel_fmt = f"({tel_limpo[:2]}) {tel_limpo[2:7]}-{tel_limpo[7:]}"

            return {
                "razao_social": dados.get('razao_social') or dados.get('nome_fantasia') or '',
                "email": str(dados.get('email') or '').lower(),
                "telefone": tel_fmt,
                "cidade": dados.get('municipio') or '',
                "uf": dados.get('uf') or '',
                "situacao": dados.get('descricao_situacao_cadastral') or 'ATIVA'
            }
    except Exception:
        return None

if __name__ == "__main__":
    print("Módulo core/api_cnpj.py compilado com sucesso!")
