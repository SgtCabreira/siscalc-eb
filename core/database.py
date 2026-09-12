import sqlite3
import os

DB_FILE = "sistema_militar.db"

def get_connection():
    """Retorna conexão com o banco SQLite com row_factory ativado."""
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Inicializa as tabelas do sistema e aplica migrações seguras."""
    conn = get_connection()
    c = conn.cursor()

    # 1. ORGANIZAÇÕES MILITARES
    c.execute('''
    CREATE TABLE IF NOT EXISTS oms (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sigla TEXT NOT NULL UNIQUE,
        nome TEXT NOT NULL,
        ug TEXT,
        logo_base64 TEXT,
        cor_primaria TEXT DEFAULT '#1B4332',
        lema TEXT DEFAULT ''
    )
    ''')

    for col, col_def in [
        ("logo_base64", "TEXT"),
        ("cor_primaria", "TEXT DEFAULT '#1B4332'"),
        ("lema", "TEXT DEFAULT ''")
    ]:
        try:
            c.execute(f"ALTER TABLE oms ADD COLUMN {col} {col_def}")
        except sqlite3.OperationalError:
            pass

    # 2. USUÁRIOS DO SISTEMA
    c.execute('''
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome_guerra TEXT NOT NULL,
        posto_grad TEXT NOT NULL,
        identidade_militar TEXT UNIQUE NOT NULL,
        om_id INTEGER,
        perfil TEXT NOT NULL,
        senha TEXT NOT NULL,
        FOREIGN KEY (om_id) REFERENCES oms (id)
    )
    ''')

    # 3. NOTAS DE CRÉDITO (NC)
    c.execute('''
    CREATE TABLE IF NOT EXISTS notas_credito (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        om_id INTEGER NOT NULL,
        ug_emitente TEXT NOT NULL,
        ug_favorecida TEXT NOT NULL,
        numero_nc TEXT NOT NULL,
        data_emissao DATE NOT NULL,
        data_limite_empenho DATE,
        valor_total REAL NOT NULL,
        valor_recolhido REAL DEFAULT 0.0,
        finalidade TEXT,
        natureza_despesa TEXT,
        pi TEXT,
        enquadramento TEXT,
        observacao TEXT,
        FOREIGN KEY (om_id) REFERENCES oms (id)
    )
    ''')

    for col, col_def in [
        ("data_limite_empenho", "DATE"),
        ("valor_recolhido", "REAL DEFAULT 0.0")
    ]:
        try:
            c.execute(f"ALTER TABLE notas_credito ADD COLUMN {col} {col_def}")
        except sqlite3.OperationalError:
            pass

    # 4. NOTAS DE EMPENHO (NE - SUPORTA VÍNCULO DUPLO COM NC 1 E NC 2)
    c.execute('''
    CREATE TABLE IF NOT EXISTS notas_empenho (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nc_id INTEGER NOT NULL,
        nc_id_2 INTEGER,
        valor_nc_1 REAL,
        valor_nc_2 REAL DEFAULT 0.0,
        om_id INTEGER NOT NULL,
        numero_ne TEXT NOT NULL,
        data_emissao DATE NOT NULL,
        valor_ne REAL NOT NULL,
        tipo_empenho TEXT NOT NULL,
        fornecedor_nome TEXT NOT NULL,
        fornecedor_cnpj TEXT NOT NULL,
        data_envio_empresa DATE,
        prazo_dias INTEGER DEFAULT 30,
        data_limite DATE,
        prorrogado INTEGER DEFAULT 0,
        nova_data_limite DATE,
        justificativa_prorrogacao TEXT,
        status TEXT DEFAULT 'Aguardando Entrega',
        fornecedor_email TEXT,
        fornecedor_telefone TEXT,
        fornecedor_cidade TEXT,
        fornecedor_uf TEXT,
        fornecedor_situacao TEXT,
        FOREIGN KEY (nc_id) REFERENCES notas_credito (id),
        FOREIGN KEY (nc_id_2) REFERENCES notas_credito (id),
        FOREIGN KEY (om_id) REFERENCES oms (id)
    )
    ''')

    for col, col_def in [
        ("nc_id_2", "INTEGER"),
        ("valor_nc_1", "REAL"),
        ("valor_nc_2", "REAL DEFAULT 0.0"),
        ("fornecedor_email", "TEXT"),
        ("fornecedor_telefone", "TEXT"),
        ("fornecedor_cidade", "TEXT"),
        ("fornecedor_uf", "TEXT"),
        ("fornecedor_situacao", "TEXT")
    ]:
        try:
            c.execute(f"ALTER TABLE notas_empenho ADD COLUMN {col} {col_def}")
        except sqlite3.OperationalError:
            pass

    # 5. NOTAS FISCAIS (NF)
    c.execute('''
    CREATE TABLE IF NOT EXISTS notas_fiscais (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ne_id INTEGER NOT NULL,
        om_id INTEGER NOT NULL,
        numero_nf TEXT NOT NULL,
        empresa_cnpj TEXT NOT NULL,
        data_emissao_nf DATE,
        data_entrada_almox DATE,
        valor_nf REAL NOT NULL,
        tipo_liquidacao TEXT NOT NULL,
        situacao TEXT DEFAULT 'Em Conferência',
        FOREIGN KEY (ne_id) REFERENCES notas_empenho (id),
        FOREIGN KEY (om_id) REFERENCES oms (id)
    )
    ''')

    # 6. ESTOQUE & ALMOXARIFADO
    c.execute('''
    CREATE TABLE IF NOT EXISTS estoque_itens (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        om_id INTEGER NOT NULL,
        nome_material TEXT NOT NULL,
        categoria TEXT,
        tipo_embalagem TEXT DEFAULT 'UNIDADE',
        fator_embalagem REAL DEFAULT 1.0,
        unidade_medida TEXT DEFAULT 'UN',
        quantidade_atual REAL NOT NULL,
        estoque_minimo REAL DEFAULT 0.0,
        estoque_ideal REAL DEFAULT 0.0,
        valor_unitario_estimado REAL NOT NULL,
        FOREIGN KEY (om_id) REFERENCES oms (id)
    )
    ''')

    for col, col_def in [
        ("tipo_embalagem", "TEXT DEFAULT 'UNIDADE'"),
        ("fator_embalagem", "REAL DEFAULT 1.0"),
        ("estoque_minimo", "REAL DEFAULT 0.0"),
        ("estoque_ideal", "REAL DEFAULT 0.0")
    ]:
        try:
            c.execute(f"ALTER TABLE estoque_itens ADD COLUMN {col} {col_def}")
        except sqlite3.OperationalError:
            pass

    # 7. SEÇÕES DA OM
    c.execute('''
    CREATE TABLE IF NOT EXISTS estoque_secoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        om_id INTEGER NOT NULL,
        nome_secao TEXT NOT NULL,
        UNIQUE(om_id, nome_secao)
    )
    ''')

    # 8. MOVIMENTAÇÕES DE ESTOQUE
    c.execute('''
    CREATE TABLE IF NOT EXISTS estoque_movimentacoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        om_id INTEGER NOT NULL,
        item_id INTEGER NOT NULL,
        tipo_movimentacao TEXT NOT NULL,
        quantidade REAL NOT NULL,
        saldo_anterior REAL NOT NULL,
        saldo_novo REAL NOT NULL,
        data_movimentacao DATE NOT NULL,
        secao_destino TEXT,
        militar_responsavel TEXT,
        ne_origem TEXT,
        observacao TEXT,
        FOREIGN KEY (item_id) REFERENCES estoque_itens (id),
        FOREIGN KEY (om_id) REFERENCES oms (id)
    )
    ''')

    # 9. PROCESSOS DE AQUISIÇÃO
    c.execute('''
    CREATE TABLE IF NOT EXISTS processos_modelos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        om_id INTEGER NOT NULL,
        tipo_processo TEXT NOT NULL,
        nome_modelo TEXT NOT NULL,
        arquivo_docx BLOB,
        formato TEXT DEFAULT 'docx',
        data_upload DATE,
        FOREIGN KEY (om_id) REFERENCES oms (id)
    )
    ''')

    c.execute('''
    CREATE TABLE IF NOT EXISTS tipos_documentos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        om_id INTEGER NOT NULL,
        nome_tipo TEXT NOT NULL,
        UNIQUE(om_id, nome_tipo)
    )
    ''')

    c.execute('''
    CREATE TABLE IF NOT EXISTS processos_gerados (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        om_id INTEGER NOT NULL,
        tipo_processo TEXT NOT NULL,
        numero_processo TEXT NOT NULL,
        objeto TEXT NOT NULL,
        fornecedor_nome TEXT,
        fornecedor_cnpj TEXT,
        valor_total REAL NOT NULL,
        data_criacao DATE,
        dados_json TEXT,
        FOREIGN KEY (om_id) REFERENCES oms (id)
    )
    ''')

    # População inicial padrão (5ª Cia PE) se vazio
    c.execute("SELECT COUNT(*) FROM oms")
    if c.fetchone()[0] == 0:
        c.execute("INSERT INTO oms (sigla, nome, ug, cor_primaria, lema) VALUES ('5ª Cia PE', '5ª Companhia de Polícia do Exército', '160222', '#0A2240', 'Uma vez PE, sempre PE!')")
        c.execute("INSERT INTO oms (sigla, nome, ug) VALUES ('5º B Sup', '5º Batalhão de Suprimento', '160223', '#800000', 'Suprir para Vencer!')")
        c.execute("INSERT INTO oms (sigla, nome, ug) VALUES ('Cmt 5ª DE', 'Comando da 5ª Divisão de Exército', '160220')")

    c.execute("SELECT COUNT(*) FROM usuarios")
    if c.fetchone()[0] == 0:
        c.execute("INSERT INTO usuarios (nome_guerra, posto_grad, identidade_militar, om_id, perfil, senha) VALUES ('Cabreira', '2º Sgt', '', 1, 'Administrador, Cmt OM, Ch 4ª Seção, SALC, Almoxarife, Op Almoxarifado', '')")

    conn.commit()
    conn.close()

if __name__ == "__main__":
    init_db()
    print("Banco de dados inicializado com sucesso via core/database.py!")

