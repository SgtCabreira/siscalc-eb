import streamlit as st
import sqlite3
import pandas as pd
import base64
import urllib.request
import json
import re
import io
from datetime import datetime, timedelta

try:
    import docx
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    docx = None

st.set_page_config(
    page_title="SisCalc - Exército Brasileiro",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="auto"
)

DB_FILE = "sistema_militar.db"


def gerar_pdf_tabela_seguro(titulo, df_dados):
    try:
        from fpdf import FPDF
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.set_auto_page_break(auto=True, margin=12)
        pdf.add_page()
        pdf.set_font('Helvetica', 'B', 13)
        
        def safe_latin(txt):
            if txt is None: return ''
            s = str(txt).replace('—', '-').replace('–', '-').replace('“', '"').replace('”', '"')
            return s.encode('latin-1', 'replace').decode('latin-1')

        pdf.cell(0, 8, safe_latin('EXÉRCITO BRASILEIRO'), ln=1, align='C')
        pdf.set_font('Helvetica', 'B', 10)
        pdf.cell(0, 6, safe_latin('5ª COMPANHIA DE POLÍCIA DO EXÉRCITO - FORTE PINHEIRINHO'), ln=1, align='C')
        pdf.set_font('Helvetica', 'I', 9)
        pdf.cell(0, 5, safe_latin(titulo), ln=1, align='C')
        pdf.ln(3)
        
        cols = list(df_dados.columns)
        pdf.set_font('Helvetica', 'B', 8)
        largura = (297 - 24) / max(len(cols), 1)
        pdf.set_fill_color(10, 34, 64) # Azul PE
        pdf.set_text_color(255, 255, 255)
        for c in cols:
            pdf.cell(largura, 6.5, safe_latin(c)[:18], border=1, align='C', fill=True)
        pdf.ln()
        
        pdf.set_font('Helvetica', '', 7)
        pdf.set_text_color(0, 0, 0)
        zebra = False
        for _, row in df_dados.iterrows():
            pdf.set_fill_color(240, 243, 246) if zebra else pdf.set_fill_color(255, 255, 255)
            for c in cols:
                pdf.cell(largura, 5.5, safe_latin(row[c])[:24], border=1, align='C', fill=True)
            pdf.ln()
            zebra = not zebra
            
        buf = io.BytesIO()
        pdf.output(buf)
        buf.seek(0)
        return buf
    except Exception:
        return None


def exportar_excel_seguro(df_dados, nome_aba="Dados"):
    try:
        import openpyxl
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            df_dados.to_excel(writer, index=False, sheet_name=nome_aba[:31])
        buf.seek(0)
        return buf
    except Exception:
        pass
    
    # Fallback puro em Python sem dependência externa
    try:
        import zipfile, xml.sax.saxutils as saxutils
        output = io.BytesIO()
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zf:
            content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
            zf.writestr('[Content_Types].xml', content_types)
            rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
            zf.writestr('_rels/.rels', rels)
            wb_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
            zf.writestr('xl/_rels/workbook.xml.rels', wb_rels)
            wb = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="{saxutils.escape(nome_aba[:31])}" sheetId="1" r:id="rId1"/></sheets></workbook>'
            zf.writestr('xl/workbook.xml', wb)

            sheet_rows = []
            cols = list(df_dados.columns)
            header_cells = "".join([f'<c t="inlineStr"><is><t>{saxutils.escape(str(c))}</t></is></c>' for c in cols])
            sheet_rows.append(f'<row r="1">{header_cells}</row>')
            for r_idx, (_, row) in enumerate(df_dados.iterrows(), start=2):
                cells = []
                for val in row:
                    s_val = saxutils.escape(str(val) if pd.notna(val) else '')
                    cells.append(f'<c t="inlineStr"><is><t>{s_val}</t></is></c>')
                sheet_rows.append(f'<row r="{r_idx}">{"".join(cells)}</row>')

            sheet_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{"".join(sheet_rows)}</sheetData></worksheet>'
            zf.writestr('xl/worksheets/sheet1.xml', sheet_xml)
        output.seek(0)
        return output
    except Exception:
        return None
def get_connection():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def formatar_data_br(val):
    if val is None or str(val).strip() == '' or str(val).lower() == 'nan':
        return ''
    s = str(val).strip()
    if len(s) == 10 and s[2] == '/' and s[5] == '/':
        return s
    try:
        return pd.to_datetime(s).strftime('%d/%m/%Y')
    except Exception:
        return s

def normalizar_data(val):
    if val is None or str(val).strip() == '' or str(val).lower() == 'nan':
        return None
    s = str(val).strip()
    if len(s) == 10 and s[2] == '/' and s[5] == '/':
        partes = s.split('/')
        return f"{partes[2]}-{partes[1]}-{partes[0]}"
    try:
        return pd.to_datetime(s).strftime('%Y-%m-%d')
    except Exception:
        return s

def buscar_dados_cnpj(cnpj, conn):
    cnpj_limpo = re.sub(r'\D', '', str(cnpj))
    if len(cnpj_limpo) != 14:
        return None
    cur = conn.cursor()
    cur.execute('''
    SELECT fornecedor_nome, fornecedor_email, fornecedor_telefone, fornecedor_cidade, fornecedor_uf, fornecedor_situacao 
    FROM notas_empenho WHERE fornecedor_cnpj LIKE ? AND fornecedor_nome IS NOT NULL AND fornecedor_nome != ''
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
    try:
        url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj_limpo}"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
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

def init_db():
    conn = get_connection()
    c = conn.cursor()
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

    c.execute('''
    CREATE TABLE IF NOT EXISTS notas_empenho (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nc_id INTEGER NOT NULL,
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
        FOREIGN KEY (om_id) REFERENCES oms (id)
    )
    ''')
    for col, col_def in [
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
        valor_unitario_estimado REAL NOT NULL,
        FOREIGN KEY (om_id) REFERENCES oms (id)
    )
    ''')
    for col, col_def in [
        ("tipo_embalagem", "TEXT DEFAULT 'UNIDADE'"),
        ("fator_embalagem", "REAL DEFAULT 1.0")
    ]:
        try:
            c.execute(f"ALTER TABLE estoque_itens ADD COLUMN {col} {col_def}")
        except sqlite3.OperationalError:
            pass

    # Tabelas para Gestão de Estoque e Seções da OM
    c.execute('''
    CREATE TABLE IF NOT EXISTS estoque_secoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        om_id INTEGER NOT NULL,
        nome_secao TEXT NOT NULL,
        UNIQUE(om_id, nome_secao)
    )
    ''')

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

    secoes_iniciais_padrao = [
        "1ª Seção", "2ª Seção", "3ª Seção", "4ª Seção",
        "Comandante", "Subcomandante", "1º Pelotão", "2º Pelotão",
        "Pelotão de Escolta e Guarda", "Seção de Cães de Guerra",
        "Garagem de Auto", "Garagem de Motos", "Reserva de Armamento",
        "Reserva de Material", "RP (Relações Públicas)",
        "Seção de Informática / Com", "Secretaria"
    ]
    for sec in secoes_iniciais_padrao:
        c.execute("INSERT OR IGNORE INTO estoque_secoes (om_id, nome_secao) VALUES (1, ?)", (sec,))

    # Tabelas para Processos de Aquisição
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
    try:
        c.execute("ALTER TABLE processos_modelos ADD COLUMN formato TEXT DEFAULT 'docx'")
    except sqlite3.OperationalError:
        pass

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

    c.execute("SELECT COUNT(*) FROM oms")
    if c.fetchone()[0] == 0:
        c.execute("INSERT INTO oms (sigla, nome, ug, cor_primaria, lema) VALUES ('5ª Cia PE', '5ª Companhia de Polícia do Exército', '160222', '#0A2240', 'Uma vez PE, sempre PE!')")
        c.execute("INSERT INTO oms (sigla, nome, ug) VALUES ('5º B Sup', '5º Batalhão de Suprimento', '160223', '#800000', 'Suprir para Vencer!')")
        c.execute("INSERT INTO oms (sigla, nome, ug) VALUES ('Cmt 5ª DE', 'Comando da 5ª Divisão de Exército', '160220')")

    c.execute("SELECT COUNT(*) FROM usuarios")
    if c.fetchone()[0] == 0:
        c.execute("INSERT INTO usuarios (nome_guerra, posto_grad, identidade_militar, om_id, perfil, senha) VALUES ('Cabreira', '2º Sgt', '0318377744', 1, 'Administrador, Cmt OM, Ch 4ª Seção, SALC, Almoxarife, Op Almoxarifado', '1234')")

    conn.commit()
    conn.close()

init_db()

# Gestão Segura de Sessão
if 'usuario' not in st.session_state:
    st.session_state['usuario'] = None

if not st.session_state['usuario']:
    user_id_param = st.query_params.get("session_user")
    last_time_param = st.query_params.get("session_time")
    if user_id_param and last_time_param:
        try:
            last_dt = datetime.fromtimestamp(float(last_time_param))
            if datetime.now() - last_dt <= timedelta(minutes=30):
                conn = get_connection()
                c = conn.cursor()
                c.execute("SELECT u.*, o.sigla as om_sigla FROM usuarios u JOIN oms o ON u.om_id = o.id WHERE u.id = ?", (int(user_id_param),))
                u_row = c.fetchone()
                conn.close()
                if u_row:
                    st.session_state['usuario'] = dict(u_row)
                    perf_rst = [p.strip() for p in u_row['perfil'].split(",") if p.strip()]
                    st.session_state['perfil_ativo'] = perf_rst[0] if perf_rst else 'SALC'
                    st.query_params["session_time"] = str(int(datetime.now().timestamp()))
        except Exception:
            st.query_params.clear()

if st.session_state.get('usuario'):
    st.query_params["session_user"] = str(st.session_state['usuario']['id'])
    st.query_params["session_time"] = str(int(datetime.now().timestamp()))

def login():
    st.sidebar.markdown("### 🔐 Acesso ao Sistema")
    conn = get_connection()
    oms = pd.read_sql_query("SELECT id, sigla FROM oms", conn)
    conn.close()

    lembrado_id = st.query_params.get("lembrar_id", "0401320478")
    lembrado_senha = st.query_params.get("lembrar_senha", "06109121")

    with st.sidebar.form("form_login"):
        om_escolhida = st.selectbox("Organização Militar (OM)", oms['sigla'].tolist())
        identidade = st.text_input("Identidade Militar", value=lembrado_id)
        senha = st.text_input("Senha", type="password", value=lembrado_senha)
        salvar_senha = st.checkbox("💾 Salvar minha senha neste computador", value=True)
        btn_entrar = st.form_submit_button("Entrar no Sistema")

        if btn_entrar:
            conn = get_connection()
            c = conn.cursor()
            c.execute("SELECT u.*, o.sigla as om_sigla FROM usuarios u JOIN oms o ON u.om_id = o.id WHERE u.identidade_militar = ? AND u.senha = ? AND o.sigla = ?", (identidade, senha, om_escolhida))
            user = c.fetchone()
            conn.close()

            if user:
                st.session_state['usuario'] = dict(user)
                perfis_u = [p.strip() for p in user['perfil'].split(",") if p.strip()]
                st.session_state['perfil_ativo'] = perfis_u[0] if perfis_u else 'SALC'
                st.query_params["session_user"] = str(user['id'])
                st.query_params["session_time"] = str(int(datetime.now().timestamp()))
                if salvar_senha:
                    st.query_params["lembrar_id"] = identidade
                    st.query_params["lembrar_senha"] = senha
                st.rerun()
            else:
                st.error("Credenciais inválidas.")

def logout():
    st.session_state['usuario'] = None
    st.session_state['perfil_ativo'] = None
    st.query_params.clear()
    st.rerun()

if not st.session_state['usuario']:
    st.title("🛡️ Sistema Integrado SALC & Almoxarifado - EB")
    st.info("Insira suas credenciais na barra lateral para acessar.")
    login()
    st.stop()

conn = get_connection()
c = conn.cursor()
c.execute("SELECT u.*, o.sigla as om_sigla, o.logo_base64, o.cor_primaria, o.lema FROM usuarios u JOIN oms o ON u.om_id = o.id WHERE u.id = ?", (st.session_state['usuario']['id'],))
user_atualizado = c.fetchone()
conn.close()
if user_atualizado:
    st.session_state['usuario'] = dict(user_atualizado)

user = st.session_state['usuario']
cor_om = user.get('cor_primaria') or '#1B4332'
lema_om = user.get('lema') or ''
logo_b64 = user.get('logo_base64')

# CSS 3D
st.markdown(f"""
<style>
/* =========================================================
   TEMA TÁTICO MILITAR DE ALTO CONTRASTE - 5ª CIA PE
   Fundo Cinza-Escuro Acetinado + Textos Nítidos em Branco
========================================================= */
html, body, [class*="css"], .stApp {{
    background: radial-gradient(circle at 50% 0%, #1a222d 0%, #11161f 55%, #0b0e14 100%) !important;
    color: #f8fafc !important;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
}}

/* Forçar Fundo Escuro em Todos os Containers do Streamlit */
.stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"], [data-testid="stToolbar"] {{
    background: radial-gradient(circle at 50% 0%, #1a222d 0%, #11161f 55%, #0b0e14 100%) !important;
    color: #f8fafc !important;
}}

/* BARRA LATERAL TÁTICA */
[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, #131821 0%, #0c1017 100%) !important;
    border-right: 1px solid rgba(197, 160, 89, 0.3) !important;
    box-shadow: 4px 0 15px rgba(0, 0, 0, 0.5) !important;
}}

/* FORÇAR TODOS OS RÓTULOS (LABELS), TEXTOS E PARÁGRAFOS A FICAREM BRANCOS E NÍTIDOS */
div[data-testid="stWidgetLabel"] *, 
label[data-testid="stWidgetLabel"] *, 
.stSelectbox label, 
.stTextInput label, 
.stNumberInput label, 
.stTextArea label, 
.stDateInput label, 
.stMultiSelect label,
.stCheckbox label,
[data-testid="stWidgetLabel"] p,
[data-testid="stWidgetLabel"] label,
.stMarkdown p, 
.stMarkdown span {{
    color: #f8fafc !important;
    font-weight: 700 !important;
    font-size: 13.5px !important;
    text-shadow: 0 1px 3px rgba(0, 0, 0, 0.7) !important;
}}

/* Títulos Sempre Nítidos e Imponentes */
h1, h2, h3, h4, h5, [data-testid="stHeading"] *, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {{
    color: #ffffff !important;
    font-weight: 800 !important;
    text-shadow: 0 2px 4px rgba(0, 0, 0, 0.6) !important;
}}

/* FORÇAR CAMPOS DE ENTRADA (INPUTS, SELECTS, TEXTAREAS) EM TEMA ESCURO METÁLICO */
input, textarea, [data-baseweb="input"], [data-baseweb="select"] > div, 
.stTextInput input, .stNumberInput input, .stTextArea textarea {{
    background-color: #1a2432 !important;
    color: #ffffff !important;
    border: 1px solid rgba(255, 255, 255, 0.25) !important;
    border-radius: 7px !important;
    font-weight: 600 !important;
}}

/* Rótulos e Métricas do Topo */
.stMetric {{
    border-left: 4px solid #C5A059;
    padding-left: 12px;
}}
.stMetric label {{
    color: #cbd5e1 !important;
    font-weight: 700 !important;
    font-size: 12px !important;
    text-transform: uppercase;
}}
.stMetric div[data-testid="stMetricValue"] {{
    color: #ffffff !important;
    font-weight: 900 !important;
}}

/* Botões do Menu Lateral 3D Acetinados */
div[role="radiogroup"] > label {{
    background: linear-gradient(145deg, #1e2634, #141a24) !important;
    border: 1px solid rgba(255, 255, 255, 0.12) !important;
    border-left: 4px solid #3b485d !important;
    border-radius: 8px;
    padding: 10px 14px !important;
    margin-bottom: 8px !important;
    box-shadow: 0 4px 10px rgba(0, 0, 0, 0.4) !important;
    transition: all 0.2s ease;
    cursor: pointer;
}}
div[role="radiogroup"] > label:hover {{
    transform: translateY(-2px);
    border-left: 4px solid #C5A059 !important;
    box-shadow: 0 6px 14px rgba(0, 0, 0, 0.6) !important;
}}
div[role="radiogroup"] > label > div:first-child {{
    display: none !important;
}}
div[role="radiogroup"] > label div p {{
    font-size: 15px !important;
    font-weight: 700 !important;
    letter-spacing: 0.5px;
    margin: 0 !important;
    color: #e2e8f0 !important;
}}
div[role="radiogroup"] > label:has(input:checked) {{
    background: linear-gradient(145deg, #0A2240, #173e72) !important;
    border-left: 5px solid #C5A059 !important;
    box-shadow: inset 0 2px 5px rgba(0,0,0,0.4), 0 3px 10px rgba(10,34,64,0.5) !important;
}}
div[role="radiogroup"] > label:has(input:checked) div p {{
    color: #ffffff !important;
    font-weight: 900 !important;
}}

/* Botão Sair do Sistema */
[data-testid="stSidebar"] div.stButton > button {{
    background: linear-gradient(145deg, #881337, #4c0519) !important;
    color: #ffffff !important;
    border: 1.5px solid #f43f5e !important;
    border-radius: 8px !important;
    font-weight: 800 !important;
    width: 100% !important;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.45) !important;
}}
[data-testid="stSidebar"] div.stButton > button:hover {{
    background: linear-gradient(145deg, #9f1239, #881337) !important;
    border-color: #fda4af !important;
    transform: translateY(-1px);
}}

/* Abas Customizadas em Alto Contraste */
.stTabs [data-baseweb="tab-list"] {{
    background: transparent;
    gap: 8px;
}}
.stTabs [data-baseweb="tab"] {{
    background: rgba(255, 255, 255, 0.06) !important;
    border-radius: 6px 6px 0 0;
    border: 1px solid rgba(255, 255, 255, 0.12);
    border-bottom: none;
    padding: 9px 18px;
}}
.stTabs [data-baseweb="tab"] p, .stTabs [data-baseweb="tab"] span, .stTabs [data-baseweb="tab"] div {{
    color: #cbd5e1 !important;
    font-weight: 700 !important;
    font-size: 14px !important;
}}
.stTabs [aria-selected="true"] {{
    background: linear-gradient(145deg, #243142, #18222e) !important;
    border-top: 3.5px solid #C5A059 !important;
}}
.stTabs [aria-selected="true"] p, .stTabs [aria-selected="true"] span, .stTabs [aria-selected="true"] div {{
    color: #ffffff !important;
    font-weight: 900 !important;
}}
</style>
""", unsafe_allow_html=True)

# Logo Lateral 80px e Atualização do Favicon
if logo_b64:
    st.markdown(f'''
        <HTML><head>
            <link rel="icon" type="image/png" href="data:image/png;base64,{logo_b64}">
        </head></HTML>
    ''', unsafe_allow_html=True)
    st.sidebar.markdown(f'''
    <div style="text-align: center; margin-bottom: 8px;">
        <img src="data:image/png;base64,{logo_b64}" style="height: 80px; width: auto; max-width: 60px; object-fit: cover; filter: drop-shadow(0px 3px 5px rgba(0,0,0,0.4));">
        <div style="font-size: 9px; font-style: italic; color: #888; margin-top: 2px;">{lema_om}</div>
    </div>
    ''', unsafe_allow_html=True)
else:
    st.sidebar.markdown(f'''
    <div style="margin: 0 auto 8px auto; width: 55px; height: 78px; background: linear-gradient(180deg, #ffffff 0%, #f0f0f0 100%); border: 1.5px solid #C5A059; border-radius: 2px 2px 28% 28%; display: flex; flex-direction: column; align-items: center; justify-content: space-between; padding: 0 0 3px 0;">
        <div style="background: {cor_om}; color: white; font-weight: 900; font-size: 7px; width: 100%; text-align: center; padding: 2px 0;">{user['om_sigla']}</div>
        <div style="font-size: 20px; margin: auto 0;">⚔️</div>
        <div style="font-size: 6px; color: #555;">{lema_om if lema_om else 'EB'}</div>
    </div>
    ''', unsafe_allow_html=True)

# Gestão de Perfis Autorizados
perfis_autorizados = [p.strip() for p in user.get('perfil', '').split(",") if p.strip()]
if not perfis_autorizados:
    perfis_autorizados = ['SALC']

if 'perfil_ativo' not in st.session_state or st.session_state['perfil_ativo'] not in perfis_autorizados:
    st.session_state['perfil_ativo'] = perfis_autorizados[0]

st.sidebar.markdown(f'''
<div style="background: rgba(255,255,255,0.06); border: 1px solid rgba(197,160,89,0.35); border-radius: 8px; padding: 10px 14px; margin-bottom: 12px; box-shadow: 0 2px 6px rgba(0,0,0,0.3);">
    <div style="font-size: 11px; font-weight: 700; color: #C5A059; text-transform: uppercase;">Militar Conectado</div>
    <div style="font-size: 16px; font-weight: 900; color: #ffffff; margin-top: 2px;">{user.get('posto_grad', '')} {user.get('nome_guerra', '')}</div>
</div>
''', unsafe_allow_html=True)

if len(perfis_autorizados) > 1:
    idx_ativo = perfis_autorizados.index(st.session_state['perfil_ativo']) if st.session_state['perfil_ativo'] in perfis_autorizados else 0
    perfil_selecionado = st.sidebar.selectbox("🔄 Perfil Ativo:", perfis_autorizados, index=idx_ativo, key="sb_perfil_escolhido")
    if perfil_selecionado != st.session_state['perfil_ativo']:
        st.session_state['perfil_ativo'] = perfil_selecionado
        st.rerun()
else:
    st.sidebar.markdown(f"**Perfil:** `{perfis_autorizados[0]}`")

perfil_ativo = st.session_state.get('perfil_ativo', perfis_autorizados[0])

# Menus de Navegação com Permissões
menu_opcoes = ["📊 Dashboard"]
if perfil_ativo in ['Administrador', 'Cmt OM', 'Ch 4ª Seção', 'SALC']:
    menu_opcoes.extend(["📑 Notas de Crédito", "📋 Notas de Empenho", "📝 Elaboração de Processos"])
if perfil_ativo in ['Administrador', 'Cmt OM', 'Ch 4ª Seção', 'SALC', 'Almoxarife', 'Op Almoxarifado']:
    menu_opcoes.append("📦 Recebimento")
if perfil_ativo in ['Administrador', 'Cmt OM', 'Ch 4ª Seção', 'Almoxarife', 'Op Almoxarifado']:
    menu_opcoes.append("🏢 Gestão de Estoque")
menu_opcoes.append("📁 Relatórios")
if perfil_ativo == 'Administrador':
    menu_opcoes.append("⚙️ Admin")

menu_selecionado = st.sidebar.radio("Navegação", menu_opcoes)

if st.sidebar.button("Sair do Sistema", key="btn_logout_sidebar"):
    logout()

st.sidebar.markdown('''
<div style="margin-top: 25px; padding-top: 10px; border-top: 1px solid rgba(255,255,255,0.1); text-align: center;">
    <div style="font-size: 11px; font-weight: bold; color: #aaa;">SisLog & Orç EB — v1.3.0</div>
    <div style="font-size: 10px; color: #888; margin-top: 2px;">Desenvolvido por: <b>2º Sgt Cabreira</b></div>
    <div style="font-size: 9px; color: #666;">5ª Cia PE — Curitiba/PR</div>
</div>
''', unsafe_allow_html=True)

# ==========================================
# 1. DASHBOARD
# ==========================================
if menu_selecionado == "📊 Dashboard":
    # Banner Institucional com Cores Oficiais da 5ª Cia PE
    st.markdown(f'''
    <div style="background: linear-gradient(135deg, #0A2240 0%, #173863 60%, #1B4332 100%); padding: 20px 25px; border-radius: 12px; margin-bottom: 22px; box-shadow: 0 4px 15px rgba(10,34,64,0.25); border-left: 6px solid #C5A059; display: flex; justify-content: space-between; align-items: center;">
        <div>
            <div style="font-size: 22px; font-weight: 900; color: #ffffff; letter-spacing: 0.5px;">🛡️ Painel de Gestão Orçamentária, Prazos & Alertas</div>
            <div style="font-size: 13px; color: #C5A059; font-weight: 600; margin-top: 3px;">5ª Companhia de Polícia do Exército — Base Major Agostinho José Rodrigues</div>
        </div>
        <div style="text-align: right;">
            <span style="background: rgba(255,255,255,0.15); color: #ffffff; padding: 5px 14px; border-radius: 20px; font-size: 12px; font-weight: 800; border: 1px solid rgba(197,160,89,0.5);">
                👁️ VISÃO DE COMANDO
            </span>
        </div>
    </div>
    ''', unsafe_allow_html=True)

    conn = get_connection()
    hoje = datetime.now().date()

    total_nc = pd.read_sql_query(f"SELECT SUM(valor_total) as val FROM notas_credito WHERE om_id = {user['om_id']}", conn)['val'].values[0] or 0.0
    total_ne = pd.read_sql_query(f"SELECT SUM(valor_ne) as val FROM notas_empenho WHERE om_id = {user['om_id']}", conn)['val'].values[0] or 0.0
    total_rec = pd.read_sql_query(f"SELECT SUM(valor_recolhido) as val FROM notas_credito WHERE om_id = {user['om_id']}", conn)['val'].values[0] or 0.0
    total_liq = pd.read_sql_query(f"SELECT SUM(valor_nf) as val FROM notas_fiscais WHERE om_id = {user['om_id']} AND situacao IN ('Liquidado (Atestado no SIAFI)', 'Pago (Ordem Bancária Emitida)')", conn)['val'].values[0] or 0.0
    total_pago = pd.read_sql_query(f"SELECT SUM(valor_nf) as val FROM notas_fiscais WHERE om_id = {user['om_id']} AND situacao = 'Pago (Ordem Bancária Emitida)'", conn)['val'].values[0] or 0.0
    saldo_empenhar = max(total_nc - total_ne - total_rec, 0.0)

    # 5 Cards Executivos de Métricas com Design Metálico 3D
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(f'''
        <div style="background: linear-gradient(145deg, #1c232e, #131821); border: 1px solid rgba(255,255,255,0.08); border-top: 4px solid #38bdf8; border-radius: 10px; padding: 14px 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.35);">
            <div style="font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px;">Créditos Recebidos (NC)</div>
            <div style="font-size: 17px; font-weight: 800; white-space: nowrap; color: #ffffff; margin-top: 4px;">R$ {total_nc:,.2f}</div>
        </div>
        '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
    with c2:
        st.markdown(f'''
        <div style="background: linear-gradient(145deg, #1c232e, #131821); border: 1px solid rgba(255,255,255,0.08); border-top: 4px solid #60a5fa; border-radius: 10px; padding: 14px 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.35);">
            <div style="font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px;">Total Empenhado (NE)</div>
            <div style="font-size: 17px; font-weight: 800; white-space: nowrap; color: #60a5fa; margin-top: 4px;">R$ {total_ne:,.2f}</div>
        </div>
        '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
    with c3:
        st.markdown(f'''
        <div style="background: linear-gradient(145deg, #1c232e, #131821); border: 1px solid rgba(255,255,255,0.08); border-top: 4px solid #22c55e; border-radius: 10px; padding: 14px 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.35);">
            <div style="font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px;">Saldo Disponível SALC</div>
            <div style="font-size: 17px; font-weight: 800; white-space: nowrap; color: #22c55e; margin-top: 4px;">R$ {saldo_empenhar:,.2f}</div>
        </div>
        '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
    with c4:
        st.markdown(f'''
        <div style="background: linear-gradient(145deg, #1c232e, #131821); border: 1px solid rgba(255,255,255,0.08); border-top: 4px solid #f59e0b; border-radius: 10px; padding: 14px 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.35);">
            <div style="font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px;">Material Liquidado</div>
            <div style="font-size: 17px; font-weight: 800; white-space: nowrap; color: #f59e0b; margin-top: 4px;">R$ {total_liq:,.2f}</div>
        </div>
        '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
    with c5:
        st.markdown(f'''
        <div style="background: linear-gradient(145deg, #1c232e, #131821); border: 1px solid rgba(255,255,255,0.08); border-top: 4px solid #a855f7; border-radius: 10px; padding: 14px 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.35);">
            <div style="font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px;">Ordem Bancária (Pago)</div>
            <div style="font-size: 17px; font-weight: 800; white-space: nowrap; color: #c084fc; margin-top: 4px;">R$ {total_pago:,.2f}</div>
        </div>
        '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)

    st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)

    # Barras de Progresso da Execução Orçamentária Global
    with st.container():
        st.markdown("#### 📈 Execução Orçamentária da OM")
        cp1, cp2, cp3 = st.columns(3)
        p_emp = (total_ne / total_nc) if total_nc > 0 else 0.0
        p_liq = (total_liq / total_ne) if total_ne > 0 else 0.0
        p_pag = (total_pago / total_ne) if total_ne > 0 else 0.0

        with cp1:
            st.markdown(f"**1. Crédito $\\rightarrow$ Empenho:** `{p_emp*100:.1f}%`")
            st.progress(min(p_emp, 1.0))
        with cp2:
            st.markdown(f"**2. Empenho $\\rightarrow$ Liquidação:** `{p_liq*100:.1f}%`")
            st.progress(min(p_liq, 1.0))
        with cp3:
            st.markdown(f"**3. Empenho $\\rightarrow$ Pagamento:** `{p_pag*100:.1f}%`")
            st.progress(min(p_pag, 1.0))

    st.markdown("---")

    # =========================================================
    # RADAR DE ALERTAS DO COMANDO (PRAZOS DE NCs E NEs)
    # =========================================================
    st.markdown("### 🚨 Radar de Prazos Críticos & Alertas do Comando")
    st.caption("Visão centralizada de pendências operacionais para tomada de decisão imediata do Comando e da 4ª Seção.")

    col_alert_nc, col_alert_ne = st.columns(2)

    # 1. ALERTA DE NOTAS DE CRÉDITO (RISCO DE PERDA DE RECURSO POR PRAZO)
    with col_alert_nc:
        st.markdown("##### ⏳ Notas de Crédito no Ponto Limite de Empenho")
        
        q_nc_alertas = f'''
        SELECT nc.id, nc.numero_nc, nc.enquadramento, nc.data_limite_empenho, nc.finalidade, nc.valor_total,
               (nc.valor_total - COALESCE(nc.valor_recolhido, 0.0) - COALESCE((SELECT SUM(valor_ne) FROM notas_empenho WHERE nc_id = nc.id), 0.0)) as saldo
        FROM notas_credito nc
        WHERE nc.om_id = {user['om_id']}
        ORDER BY nc.data_limite_empenho ASC
        '''
        df_nc_radar = pd.read_sql_query(q_nc_alertas, conn)
        
        ncs_com_urgencia = []
        if not df_nc_radar.empty:
            for _, row_nc in df_nc_radar.iterrows():
                saldo_val = float(row_nc['saldo'] or 0.0)
                if saldo_val > 0.01 and pd.notna(row_nc['data_limite_empenho']):
                    try:
                        limite_dt = pd.to_datetime(row_nc['data_limite_empenho']).date()
                        dias = (limite_dt - hoje).days
                        if dias <= 7: # Alerta se vence em até 7 dias ou já venceu
                            ncs_com_urgencia.append({
                                "numero_nc": row_nc['numero_nc'],
                                "enquadramento": row_nc['enquadramento'],
                                "finalidade": row_nc['finalidade'],
                                "saldo": saldo_val,
                                "limite": limite_dt,
                                "dias": dias
                            })
                    except Exception:
                        pass

        if ncs_com_urgencia:
            for item in ncs_com_urgencia:
                dias = item['dias']
                if dias < 0:
                    badge = f"<span style='background: #ef4444; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;'>🔴 VENCIDA ({abs(dias)}d atrás)</span>"
                    borda_color = "#ef4444"
                elif dias <= 2:
                    badge = f"<span style='background: #f97316; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;'>🟠 URGENTE ({dias} dias)</span>"
                    borda_color = "#f97316"
                else:
                    badge = f"<span style='background: #eab308; color: black; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;'>🟡 ATENÇÃO ({dias} dias)</span>"
                    borda_color = "#eab308"

                st.markdown(f'''
                <div style="background: linear-gradient(145deg, #1c232e, #131821); border: 1px solid rgba(255,255,255,0.08); border-left: 5px solid {borda_color}; border-radius: 10px; padding: 13px 15px; margin-bottom: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.35);">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="font-weight: 800; font-size: 15px; color: #ffffff;">{item['numero_nc']}</span>
                        {badge}
                    </div>
                    <div style="font-size: 13px; color: #cbd5e1; margin: 6px 0; font-weight: 600;">{item['finalidade'][:75]}</div>
                    <div style="display: flex; justify-content: space-between; font-size: 12px; border-top: 1px solid rgba(255,255,255,0.06); padding-top: 6px;">
                        <span style="color: #94a3b8;">Limite: <b>{formatar_data_br(item['limite'])}</b></span>
                        <span style="color: #ef4444; font-weight: 800;">Saldo em Risco: R$ {item['saldo']:,.2f}</span>
                    </div>
                </div>
                '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
        else:
            st.markdown('''
            <div style="background: rgba(34, 197, 94, 0.1); border: 1.5px solid #22c55e; border-radius: 8px; padding: 16px; text-align: center; color: #166534; font-weight: 700;">
                🟢 Nenhuma Nota de Crédito com prazo crítico de empenho ou recurso represado.
            </div>
            ''', unsafe_allow_html=True)

    # 2. ALERTA DE NOTAS DE EMPENHO (ENTREGA DE MATERIAIS PELAS EMPRESAS)
    with col_alert_ne:
        st.markdown("##### 🚚 Prazos de Entrega dos Fornecedores (NE)")
        
        q_ne_alertas = f'''
        SELECT ne.numero_ne, ne.fornecedor_nome, ne.fornecedor_telefone, ne.valor_ne, ne.status,
               COALESCE(ne.nova_data_limite, ne.data_limite) as prazo_final
        FROM notas_empenho ne
        WHERE ne.om_id = {user['om_id']} AND ne.status NOT IN ('Liquidado', 'Pago')
        ORDER BY prazo_final ASC
        '''
        df_ne_radar = pd.read_sql_query(q_ne_alertas, conn)
        
        nes_com_urgencia = []
        if not df_ne_radar.empty:
            for _, row_ne in df_ne_radar.iterrows():
                if pd.notna(row_ne['prazo_final']):
                    try:
                        p_dt = pd.to_datetime(row_ne['prazo_final']).date()
                        dias_e = (p_dt - hoje).days
                        if dias_e <= 5: # Alerta se vence em até 5 dias ou já atrasou
                            nes_com_urgencia.append({
                                "numero_ne": row_ne['numero_ne'],
                                "fornecedor": row_ne['fornecedor_nome'],
                                "telefone": row_ne['fornecedor_telefone'] or "Não cadastrado",
                                "valor": float(row_ne['valor_ne'] or 0.0),
                                "prazo": p_dt,
                                "dias": dias_e
                            })
                    except Exception:
                        pass

        if nes_com_urgencia:
            for item in nes_com_urgencia:
                dias_e = item['dias']
                if dias_e < 0:
                    badge_ne = f"<span style='background: #ef4444; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;'>🔴 ATRASADO ({abs(dias_e)}d)</span>"
                    borda_ne_cor = "#ef4444"
                elif dias_e <= 2:
                    badge_ne = f"<span style='background: #f97316; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;'>🟠 VENCE EM {dias_e}d</span>"
                    borda_ne_cor = "#f97316"
                else:
                    badge_ne = f"<span style='background: #eab308; color: black; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;'>🟡 VENCE EM {dias_e}d</span>"
                    borda_ne_cor = "#eab308"

                st.markdown(f'''
                <div style="background: linear-gradient(145deg, #1c232e, #131821); border: 1px solid rgba(255,255,255,0.08); border-left: 5px solid {borda_ne_cor}; border-radius: 10px; padding: 13px 15px; margin-bottom: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.35);">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="font-weight: 800; font-size: 15px; color: #ffffff;">{item['numero_ne']}</span>
                        {badge_ne}
                    </div>
                    <div style="font-size: 13px; color: #cbd5e1; margin: 6px 0; font-weight: 700;">{item['fornecedor'][:55]}</div>
                    <div style="display: flex; justify-content: space-between; font-size: 12px; border-top: 1px solid rgba(255,255,255,0.06); padding-top: 6px;">
                        <span style="color: #94a3b8;">📞 <b>{item['telefone']}</b> | Limite: <b>{formatar_data_br(item['prazo'])}</b></span>
                        <span style="color: #60a5fa; font-weight: 800;">R$ {item['valor']:,.2f}</span>
                    </div>
                </div>
                '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
        else:
            st.markdown('''
            <div style="background: rgba(34, 197, 94, 0.1); border: 1.5px solid #22c55e; border-radius: 8px; padding: 16px; text-align: center; color: #166534; font-weight: 700;">
                🟢 Todas as empresas e materiais estão com prazos de entrega rigorosamente em dia.
            </div>
            ''', unsafe_allow_html=True)

    conn.close()

# ==========================================
# 2. NOTAS DE CRÉDITO (CARDS VISUAIS, POP-OVER ORDENAÇÃO E FILTRO)
# ==========================================
elif menu_selecionado == "📑 Notas de Crédito":
    st.title("📑 Notas de Crédito da OM")
    conn = get_connection()

    col_nc_top1, col_nc_top2, col_nc_top3 = st.columns(3)
    with col_nc_top1:
        with st.expander("➕ Cadastrar Nova Nota de Crédito", expanded=False):
            with st.form("form_nc_rapido"):
                c1, c2, c3 = st.columns(3)
                ug_emit = c1.text_input("UG Emitente", value="160220 - Cmt 5ª DE")
                ug_fav = c2.text_input("UG Favorecida", value=f"{user['om_sigla']}")
                num_nc = c3.text_input("Número da NC (ex.: 2026NC000123)")

                c4, c5, c6 = st.columns(3)
                dt_emissao = c4.date_input("Data de Emissão", value=datetime.now(), format="DD/MM/YYYY")
                dt_limite = c5.date_input("Data Limite do Empenho", value=datetime.now() + timedelta(days=30), format="DD/MM/YYYY")
                val_total = c6.number_input("Valor Total (R$)", min_value=0.01, step=100.0, format="%.2f")

                c7, c8, c9 = st.columns(3)
                nd = c7.text_input("Natureza de Despesa (ND)", value="339030 - Material de Consumo")
                pi = c8.text_input("Plano Interno (PI)", value="I2B0000000")
                enq = c9.selectbox("Enquadramento", ["DISCRICIONÁRIA", "ALIMENTAÇÃO", "FARDAMENTO", "SAÚDE", "ENGENHARIA", "OUTRO"])

                fin = st.text_input("Finalidade / Objeto (Sucinto)", value="Aquisição de material")
                obs = st.text_area("Observações")

                if st.form_submit_button("Salvar Nova NC", type="primary"):
                    cur = conn.cursor()
                    cur.execute('''
                    INSERT INTO notas_credito (om_id, ug_emitente, ug_favorecida, numero_nc, data_emissao, data_limite_empenho,
                                               valor_total, valor_recolhido, finalidade, natureza_despesa, pi, enquadramento, observacao)
                    VALUES (?, ?, ?, ?, ?, ?, ?, 0.0, ?, ?, ?, ?, ?)
                    ''', (user['om_id'], ug_emit, ug_fav, num_nc, str(dt_emissao), str(dt_limite), val_total,
                          fin, nd, pi, enq, obs))
                    conn.commit()
                    st.success(f"Nota de Crédito {num_nc} cadastrada com sucesso!")
                    st.rerun()

    with col_nc_top2:
        with st.expander("✏️ Editar Nota de Crédito", expanded=False):
            df_ncs_edit_top = pd.read_sql_query(f"SELECT * FROM notas_credito WHERE om_id = {user['om_id']} ORDER BY numero_nc", conn)
            if not df_ncs_edit_top.empty:
                dict_edit_nc = {f"{r['numero_nc']} - R$ {r['valor_total']:,.2f} | {r['finalidade'][:35]}...": r['id'] for _, r in df_ncs_edit_top.iterrows()}
                sel_nc_ed = st.selectbox("Selecione a NC para editar:", list(dict_edit_nc.keys()), key="sel_nc_to_edit_top")
                id_nc_to_edit = dict_edit_nc[sel_nc_ed]
                nc_dados_atual = dict(df_ncs_edit_top[df_ncs_edit_top['id'] == id_nc_to_edit].iloc[0])

                with st.form(f"form_ed_nc_top_{id_nc_to_edit}"):
                    ec1, ec2, ec3 = st.columns(3)
                    ed_num_nc = ec1.text_input("Número da NC:", value=nc_dados_atual['numero_nc'])
                    ed_ug_emit = ec2.text_input("UG Emitente:", value=nc_dados_atual['ug_emitente'] or "")
                    ed_ug_fav = ec3.text_input("UG Favorecida:", value=nc_dados_atual['ug_favorecida'] or "")

                    ec4, ec5, ec6 = st.columns(3)
                    dt_em_val = pd.to_datetime(nc_dados_atual['data_emissao']).date() if pd.notna(nc_dados_atual['data_emissao']) else datetime.now().date()
                    ed_dt_emissao = ec4.date_input("Data de Emissão:", value=dt_em_val, format="DD/MM/YYYY")

                    dt_lim_val = pd.to_datetime(nc_dados_atual['data_limite_empenho']).date() if pd.notna(nc_dados_atual['data_limite_empenho']) else (datetime.now().date() + timedelta(days=30))
                    ed_dt_limite = ec5.date_input("Data Limite do Empenho:", value=dt_lim_val, format="DD/MM/YYYY")

                    ed_val_total = ec6.number_input("Valor Total (R$):", value=float(nc_dados_atual['valor_total']), step=100.0, format="%.2f")

                    ec7, ec8, ec9 = st.columns(3)
                    ed_nd = ec7.text_input("Natureza de Despesa (ND):", value=nc_dados_atual['natureza_despesa'] or "")
                    ed_pi = ec8.text_input("Plano Interno (PI):", value=nc_dados_atual['pi'] or "")
                    enq_lista = ["DISCRICIONÁRIA", "ALIMENTAÇÃO", "FARDAMENTO", "SAÚDE", "ENGENHARIA", "OUTRO"]
                    idx_enq = enq_lista.index(nc_dados_atual['enquadramento']) if nc_dados_atual['enquadramento'] in enq_lista else 0
                    ed_enq = ec9.selectbox("Enquadramento:", enq_lista, index=idx_enq)

                    ed_fin = st.text_input("Finalidade / Objeto Completo:", value=nc_dados_atual['finalidade'] or "")
                    ed_rec = st.number_input("Valor Recolhido / Devolvido (R$):", value=float(nc_dados_atual['valor_recolhido'] or 0.0), step=10.0, format="%.2f")
                    ed_obs = st.text_area("Observações:", value=nc_dados_atual['observacao'] or "")

                    if st.form_submit_button("💾 Salvar Todas as Alterações da NC", type="primary"):
                        c = conn.cursor()
                        c.execute('''
                        UPDATE notas_credito
                        SET numero_nc = ?, finalidade = ?, valor_total = ?, valor_recolhido = ?,
                            data_emissao = ?, data_limite_empenho = ?, enquadramento = ?,
                            natureza_despesa = ?, pi = ?, ug_emitente = ?, ug_favorecida = ?, observacao = ?
                        WHERE id = ?
                        ''', (ed_num_nc.strip(), ed_fin.strip(), ed_val_total, ed_rec,
                              str(ed_dt_emissao), str(ed_dt_limite), ed_enq, ed_nd.strip(),
                              ed_pi.strip(), ed_ug_emit.strip(), ed_ug_fav.strip(), ed_obs.strip(), id_nc_to_edit))
                        conn.commit()
                        st.success(f"✅ Nota de Crédito {ed_num_nc} atualizada com sucesso!")
                        st.rerun()
            else:
                st.info("Nenhuma NC cadastrada para editar.")

    with col_nc_top3:
        with st.expander("🗑️ Excluir Nota de Crédito", expanded=False):
            df_ncs_del = pd.read_sql_query(f"SELECT id, numero_nc, valor_total, finalidade FROM notas_credito WHERE om_id = {user['om_id']}", conn)
            if not df_ncs_del.empty:
                dict_del = {f"{row['numero_nc']} - R$ {row['valor_total']:,.2f} | {row['finalidade'][:40]}...": row['id'] for _, row in df_ncs_del.iterrows()}
                nc_del_escolha = st.selectbox("Selecione a NC para excluir:", list(dict_del.keys()), key="nc_del_sel_quick")
                id_nc_del = dict_del[nc_del_escolha]

                cur = conn.cursor()
                cur.execute("SELECT COUNT(*) FROM notas_empenho WHERE nc_id = ?", (id_nc_del,))
                qtd_ne = cur.fetchone()[0]

                if qtd_ne > 0:
                    st.warning(f"⚠️ Não é possível excluir: existem {qtd_ne} Nota(s) de Empenho vinculadas a esta NC.")
                else:
                    if st.button("Confirmar Exclusão Definitiva", type="primary", key="btn_del_nc_exec"):
                        cur.execute("DELETE FROM notas_credito WHERE id = ?", (id_nc_del,))
                        conn.commit()
                        st.success("Nota de Crédito excluída com sucesso!")
                        st.rerun()
            else:
                st.info("Nenhuma NC cadastrada para exclusão.")

    # Controles de Ordenação e Filtro
    col_t_nc, col_ord_nc, col_flt_nc = st.columns([6, 1, 1])
    with col_t_nc:
        st.markdown("### 🗂️ Painel Visual de Notas de Crédito")

    with col_ord_nc:
        with st.popover("🔀 Ordenar"):
            ordem_nc_escolhida = st.radio(
                "Critério de Ordenação:",
                [
                    "🚨 Mais urgente ao menos urgente (Padrão)",
                    "💰 Maior valor",
                    "💵 Menor valor",
                    "📅 Data de emissão mais recente",
                    "⏳ Prazo limite mais próximo"
                ],
                key="radio_ordem_nc"
            )

    with col_flt_nc:
        with st.popover("🌪️ Filtrar"):
            f_venc = st.checkbox("🔴 Vencidos", value=True, key="chk_venc")
            f_urg = st.checkbox("🟠 Urgentes (≤ 2 dias)", value=True, key="chk_urg")
            f_aten = st.checkbox("🟡 Atenção (≤ 7 dias)", value=True, key="chk_aten")
            f_prazo = st.checkbox("🟢 No Prazo", value=True, key="chk_prazo")
            f_baix = st.checkbox("✅ Concluídas / Baixadas", value=True, key="chk_baix")

    query_cards_nc = f'''
    SELECT nc.id, nc.numero_nc, nc.data_emissao, nc.data_limite_empenho, nc.enquadramento, 
           nc.valor_total, COALESCE(nc.valor_recolhido, 0.0) as valor_recolhido,
           COALESCE(SUM(ne.valor_ne), 0.0) as total_empenhado,
           (nc.valor_total - COALESCE(SUM(ne.valor_ne), 0.0) - COALESCE(nc.valor_recolhido, 0.0)) as saldo_restante,
           nc.finalidade
    FROM notas_credito nc
    LEFT JOIN notas_empenho ne ON nc.id = ne.nc_id
    WHERE nc.om_id = {user['om_id']}
    GROUP BY nc.id
    '''
    df_ncs = pd.read_sql_query(query_cards_nc, conn)

    if not df_ncs.empty:
        hoje = datetime.now().date()
        df_ncs['saldo_calc'] = df_ncs['saldo_restante'].apply(lambda s: max(s, 0.0))
        df_ncs['data_limite_dt'] = pd.to_datetime(df_ncs['data_limite_empenho']).dt.date

        def atribuir_prioridade(row):
            if row['saldo_calc'] <= 0.01:
                return 5, "BAIXADA", "✅ CONCLUÍDA", "#334155", "border-left: 6px solid #64748b;"
            if pd.isna(row['data_limite_dt']):
                return 4, "SEM_PRAZO", "⚪ SEM PRAZO", "#334155", "border-left: 6px solid #64748b;"
            dias = (row['data_limite_dt'] - hoje).days
            if dias < 0:
                return 0, "VENCIDO", f"🔴 VENCIDO ({abs(dias)}d)", "#7f1d1d", "border-left: 6px solid #ef4444;"
            elif dias <= 2:
                return 1, "URGENTE", f"🟠 URGENTE ({dias}d)", "#9a3412", "border-left: 6px solid #f97316;"
            elif dias <= 7:
                return 2, "ATENCAO", f"🟡 ATENÇÃO ({dias}d)", "#854d0e", "border-left: 6px solid #eab308;"
            else:
                return 3, "NO_PRAZO", f"🟢 NO PRAZO ({dias}d)", "#14532d", "border-left: 6px solid #22c55e;"

        df_ncs[['peso_urgencia', 'cat_prazo', 'status_txt', 'status_bg', 'borda_cor']] = df_ncs.apply(atribuir_prioridade, axis=1, result_type='expand')

        categorias_permitidas = []
        if f_venc: categorias_permitidas.append("VENCIDO")
        if f_urg: categorias_permitidas.append("URGENTE")
        if f_aten: categorias_permitidas.append("ATENCAO")
        if f_prazo: categorias_permitidas.append("NO_PRAZO")
        if f_baix: categorias_permitidas.extend(["BAIXADA", "SEM_PRAZO"])

        df_filtrado = df_ncs[df_ncs['cat_prazo'].isin(categorias_permitidas)].copy()

        if "Mais urgente" in ordem_nc_escolhida:
            df_filtrado = df_filtrado.sort_values(by=['peso_urgencia', 'saldo_calc'], ascending=[True, False])
        elif "Maior valor" in ordem_nc_escolhida:
            df_filtrado = df_filtrado.sort_values(by='valor_total', ascending=False)
        elif "Menor valor" in ordem_nc_escolhida:
            df_filtrado = df_filtrado.sort_values(by='valor_total', ascending=True)
        elif "mais recente" in ordem_nc_escolhida:
            df_filtrado = df_filtrado.sort_values(by='data_emissao', ascending=False)
        elif "Prazo limite" in ordem_nc_escolhida:
            df_filtrado = df_filtrado.sort_values(by='data_limite_dt', ascending=True)

        cols = st.columns(3)
        for idx, row in df_filtrado.reset_index(drop=True).iterrows():
            with cols[idx % 3]:
                st.markdown(f'''
                <div style="background: linear-gradient(145deg, #243142, #1a232e); border: 1.5px solid rgba(255,255,255,0.18); {row['borda_cor']} border-radius: 10px; padding: 15px; margin-bottom: 10px; box-shadow: 0 6px 16px rgba(0,0,0,0.45);">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="font-weight: 900; font-size: 16px; color: #ffffff; letter-spacing: 0.5px;">{row['numero_nc']}</span>
                        <span style="font-size: 11px; padding: 3px 8px; border-radius: 4px; background: {row['status_bg']}; color: #fff; font-weight: 800;">{row['status_txt']}</span>
                    </div>
                    <div style="font-size: 13.5px; color: #f8fafc; margin: 8px 0; font-weight: 600; line-height: 1.35; min-height: 38px;">
                        {row['finalidade'][:75]}
                    </div>
                    <div style="display: flex; justify-content: space-between; font-size: 12.5px; border-top: 1px solid rgba(255,255,255,0.12); padding-top: 8px;">
                        <span style="color: #cbd5e1;">Crédito: <b style="color: #ffffff;">R$ {row['valor_total']:,.2f}</b></span>
                        <span style="color: #cbd5e1;">Saldo: <b style="color: #4ade80; font-size: 14px;">R$ {row['saldo_calc']:,.2f}</b></span>
                    </div>
                </div>
                ''', unsafe_allow_html=True)
                with st.popover("🔍 Ver Detalhes", key=f"pop_nc_{row['id']}", use_container_width=True):
                    c = conn.cursor()
                    c.execute('''
                    SELECT nc.*, COALESCE(SUM(ne.valor_ne), 0.0) as empenhado_real
                    FROM notas_credito nc
                    LEFT JOIN notas_empenho ne ON nc.id = ne.nc_id
                    WHERE nc.id = ?
                    GROUP BY nc.id
                    ''', (row['id'],))
                    nc_info = dict(c.fetchone())
                    saldo_det = max(nc_info['valor_total'] - nc_info['empenhado_real'] - (nc_info['valor_recolhido'] or 0.0), 0.0)

                    st.markdown(f'''
                    <div style="border-bottom: 2px solid #C5A059; padding-bottom: 8px; margin-bottom: 12px;">
                        <div style="font-size: 11px; font-weight: 800; color: #C5A059; text-transform: uppercase; letter-spacing: 0.5px;">Dossiê da Nota de Crédito</div>
                        <div style="font-size: 19px; font-weight: 900; color: #ffffff;">{nc_info['numero_nc']}</div>
                        <div style="font-size: 14px; font-weight: 600; color: #f1f5f9; margin-top: 4px;">{nc_info['finalidade']}</div>
                    </div>
                    ''', unsafe_allow_html=True)

                    st.markdown(f'''
                    <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; margin: 10px 0;">
                        <div style="background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08); border-left: 3px solid #C5A059; padding: 7px 10px; border-radius: 6px;">
                            <div style="font-size: 10px; color: #94a3b8; font-weight: 700; text-transform: uppercase;">Valor Total</div>
                            <div style="font-size: 15px; font-weight: 800; color: #ffffff; white-space: nowrap; margin-top: 2px;">R$ {nc_info['valor_total']:,.2f}</div>
                        </div>
                        <div style="background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08); border-left: 3px solid #60a5fa; padding: 7px 10px; border-radius: 6px;">
                            <div style="font-size: 10px; color: #94a3b8; font-weight: 700; text-transform: uppercase;">Empenhado</div>
                            <div style="font-size: 15px; font-weight: 800; color: #60a5fa; white-space: nowrap; margin-top: 2px;">R$ {nc_info['empenhado_real']:,.2f}</div>
                        </div>
                        <div style="background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08); border-left: 3px solid #fbbf24; padding: 7px 10px; border-radius: 6px;">
                            <div style="font-size: 10px; color: #94a3b8; font-weight: 700; text-transform: uppercase;">Recolhido</div>
                            <div style="font-size: 15px; font-weight: 800; color: #fbbf24; white-space: nowrap; margin-top: 2px;">R$ {nc_info['valor_recolhido'] or 0.0:,.2f}</div>
                        </div>
                        <div style="background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08); border-left: 3px solid #22c55e; padding: 7px 10px; border-radius: 6px;">
                            <div style="font-size: 10px; color: #94a3b8; font-weight: 700; text-transform: uppercase;">Disponível</div>
                            <div style="font-size: 15px; font-weight: 800; color: #22c55e; white-space: nowrap; margin-top: 2px;">R$ {saldo_det:,.2f}</div>
                        </div>
                    </div>
                    '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)

                    st.markdown("---")
                    ci1, ci2, ci3 = st.columns(3)
                    ci1.markdown(f"**Emissão:** {formatar_data_br(nc_info['data_emissao'])}")
                    ci1.markdown(f"**Limite:** {formatar_data_br(nc_info['data_limite_empenho'])}")
                    ci2.markdown(f"**Enquadramento:** {nc_info['enquadramento']}")
                    ci2.markdown(f"**ND:** {nc_info['natureza_despesa']}")
                    ci3.markdown(f"**PI:** {nc_info['pi']}")
                    ci3.markdown(f"**UG:** {nc_info['ug_emitente']}")

                    # Notas de Empenho vinculadas - Renderizado limpo via DataFrame nativo (sem código HTML visível)
                    st.markdown("---")
                    df_nes_vinc = pd.read_sql_query(f'''
                    SELECT numero_ne as "Número NE", data_emissao as "Emissão", fornecedor_nome as "Fornecedor",
                           fornecedor_cnpj as "CNPJ", valor_ne as "Valor (R$)", status as "Status"
                    FROM notas_empenho WHERE nc_id = {row['id']} ORDER BY id DESC
                    ''', conn)

                    if not df_nes_vinc.empty:
                        df_nes_vinc['Emissão'] = df_nes_vinc['Emissão'].apply(formatar_data_br)
                        df_nes_vinc['Valor (R$)'] = df_nes_vinc['Valor (R$)'].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                        st.markdown(f"**📋 Notas de Empenho Vinculadas ({len(df_nes_vinc)} registro(s)):**")
                        st.dataframe(df_nes_vinc, use_container_width=True, hide_index=True)
                    else:
                        st.info("ℹ️ Nenhuma Nota de Empenho vinculada a esta NC. O saldo está 100% disponível.")

                    with st.expander("✏️ Editar Todos os Dados desta Nota de Crédito", expanded=False):
                        with st.form(f"form_ed_nc_card_{row['id']}"):
                            dc1, dc2 = st.columns(2)
                            d_num_nc = dc1.text_input("Número da NC:", value=nc_info['numero_nc'], key=f"d_num_{row['id']}")
                            d_val_tot = dc2.number_input("Valor Total (R$):", value=float(nc_info['valor_total']), step=100.0, format="%.2f", key=f"d_val_{row['id']}")

                            d_fin = st.text_input("Finalidade / Objeto:", value=nc_info['finalidade'] or "", key=f"d_fin_{row['id']}")

                            dc3, dc4, dc5 = st.columns(3)
                            d_dt_em = pd.to_datetime(nc_info['data_emissao']).date() if pd.notna(nc_info['data_emissao']) else datetime.now().date()
                            d_emissao = dc3.date_input("Data de Emissão:", value=d_dt_em, format="DD/MM/YYYY", key=f"d_em_{row['id']}")
                            
                            d_dt_lim = pd.to_datetime(nc_info['data_limite_empenho']).date() if pd.notna(nc_info['data_limite_empenho']) else (datetime.now().date() + timedelta(days=30))
                            d_limite = dc4.date_input("Data Limite Empenho:", value=d_dt_lim, format="DD/MM/YYYY", key=f"d_lim_{row['id']}")

                            d_enq_opts = ["DISCRICIONÁRIA", "ALIMENTAÇÃO", "FARDAMENTO", "SAÚDE", "ENGENHARIA", "OUTRO"]
                            d_enq_idx = d_enq_opts.index(nc_info['enquadramento']) if nc_info['enquadramento'] in d_enq_opts else 0
                            d_enq = dc5.selectbox("Enquadramento:", d_enq_opts, index=d_enq_idx, key=f"d_enq_{row['id']}")

                            dc6, dc7 = st.columns(2)
                            d_nd = dc6.text_input("Natureza da Despesa (ND):", value=nc_info['natureza_despesa'] or "", key=f"d_nd_{row['id']}")
                            d_pi = dc7.text_input("Plano Interno (PI):", value=nc_info['pi'] or "", key=f"d_pi_{row['id']}")

                            d_rec = st.number_input("Valor de Saldo Recolhido (R$):", value=float(nc_info['valor_recolhido'] or 0.0), step=10.0, format="%.2f", key=f"d_rec_{row['id']}")

                            if st.form_submit_button("💾 Salvar Alterações desta NC", type="primary"):
                                cur = conn.cursor()
                                cur.execute('''
                                UPDATE notas_credito
                                SET numero_nc = ?, finalidade = ?, valor_total = ?, valor_recolhido = ?,
                                    data_emissao = ?, data_limite_empenho = ?, enquadramento = ?,
                                    natureza_despesa = ?, pi = ?
                                WHERE id = ?
                                ''', (d_num_nc.strip(), d_fin.strip(), d_val_tot, d_rec, str(d_emissao), str(d_limite), d_enq, d_nd.strip(), d_pi.strip(), row['id']))
                                conn.commit()
                                st.success(f"✅ NC {d_num_nc} atualizada!")
                                st.rerun()

                    st.markdown("---")
                    col_rec1, col_rec2 = st.columns([2, 1])
                    novo_rec = col_rec1.number_input("Informar Saldo Recolhido (R$):", value=float(nc_info['valor_recolhido'] or 0.0), step=10.0, format="%.2f", key=f"in_rec_{row['id']}")
                    if col_rec2.button("Gravar Baixa", type="primary", key=f"btn_rec_{row['id']}", use_container_width=True):
                        c.execute("UPDATE notas_credito SET valor_recolhido = ? WHERE id = ?", (novo_rec, row['id']))
                        conn.commit()
                        st.success("Valor recolhido gravado com sucesso!")
                        st.rerun()
    conn.close()

# ==========================================
# 3. NOTAS DE EMPENHO (CARDS VISUAIS & ORDENAÇÃO)
# ==========================================
elif menu_selecionado == "📋 Notas de Empenho":
    st.title("📋 Notas de Empenho da OM")
    conn = get_connection()

    col_ne_top1, col_ne_top2, col_ne_top3 = st.columns(3)

    with col_ne_top1:
        with st.expander("➕ Cadastrar Nova Nota de Empenho", expanded=False):
            df_nc_opcoes = pd.read_sql_query(f'''
            SELECT id, numero_nc, valor_total,
                   (valor_total - COALESCE(valor_recolhido, 0.0) - (SELECT COALESCE(SUM(valor_ne), 0) FROM notas_empenho WHERE nc_id = notas_credito.id)) as saldo
            FROM notas_credito WHERE om_id = {user['om_id']}
            ''', conn)

            if df_nc_opcoes.empty:
                st.warning("Cadastre primeiro uma Nota de Crédito com saldo disponível.")
            else:
                opcoes_nc_dict = {f"{row['numero_nc']} (Saldo Disponível: R$ {row['saldo']:,.2f})": row['id'] for _, row in df_nc_opcoes.iterrows()}
                nc_selecionada = st.selectbox("Vincular à Nota de Crédito:", list(opcoes_nc_dict.keys()), key="nc_sel_empenho")
                nc_id = opcoes_nc_dict[nc_selecionada]

                st.markdown("##### 🔍 Busca Automática da Empresa na Receita Federal")
                col_b_cnpj, col_b_btn = st.columns([3, 1])
                cnpj_digitado = col_b_cnpj.text_input("Digite o CNPJ da Empresa:", key="cnpj_busca_ne_input")
                
                if col_b_btn.button("Buscar Empresa", key="btn_buscar_empresa_ne", use_container_width=True):
                    dados_emp = buscar_dados_cnpj(cnpj_digitado, conn)
                    if dados_emp:
                        st.session_state['ne_cnpj_auto'] = cnpj_digitado
                        st.session_state['ne_nome_auto'] = dados_emp.get('razao_social', '')
                        st.session_state['ne_email_auto'] = dados_emp.get('email', '')
                        st.session_state['ne_tel_auto'] = dados_emp.get('telefone', '')
                        st.session_state['ne_cidade_auto'] = dados_emp.get('cidade', '')
                        st.session_state['ne_uf_auto'] = dados_emp.get('uf', '')
                        st.session_state['ne_sit_auto'] = dados_emp.get('situacao', 'ATIVA')
                        st.success(f"✅ Empresa Localizada: **{st.session_state['ne_nome_auto']}**")
                    else:
                        st.warning("CNPJ não localizado na Receita Federal ou sem conexão no momento.")

                with st.form("form_ne_limpo"):
                    col1, col2, col3 = st.columns(3)
                    numero_ne = col1.text_input("Número do Empenho (ex.: 2026NE000456)")
                    data_emissao_ne = col2.date_input("Data de Emissão da NE", value=datetime.now(), format="DD/MM/YYYY")
                    valor_ne = col3.number_input("Valor da NE (R$)", min_value=0.01, step=50.0, format="%.2f")

                    col4, col5, col6 = st.columns(3)
                    tipo_empenho = col4.selectbox("Tipo de Empenho", ["Ordinário", "Global", "Estimativo"])
                    fornecedor_nome = col5.text_input("Nome Empresarial (Razão Social)", value=st.session_state.get('ne_nome_auto', ''))
                    fornecedor_cnpj = col6.text_input("CNPJ do Fornecedor", value=st.session_state.get('ne_cnpj_auto', ''))

                    col7, col8 = st.columns(2)
                    data_envio = col7.date_input("Data de Recebimento pela Empresa (Marco Zero)", value=datetime.now(), format="DD/MM/YYYY")
                    prazo_dias = col8.number_input("Prazo de Entrega (Dias Corridos)", value=30, step=5)

                    data_limite_calc = data_envio + timedelta(days=prazo_dias)
                    st.info(f"📅 Data Limite Regulamentar Automática: **{data_limite_calc.strftime('%d/%m/%Y')}**")

                    prorrogado = st.checkbox("Houve Prorrogação de Prazo Solicitada?", key="chk_prorr_ne")
                    nova_data = None
                    justificativa = None
                    if prorrogado:
                        col_p1, col_p2 = st.columns(2)
                        nova_data = col_p1.date_input("Nova Data Limite de Entrega", value=data_limite_calc + timedelta(days=15), format="DD/MM/YYYY")
                        justificativa = col_p2.text_area("Justificativa da Empresa / Despacho")

                    if st.form_submit_button("Salvar Empenho", type="primary"):
                        if not fornecedor_cnpj:
                            st.error("O CNPJ do fornecedor é obrigatório.")
                        elif not numero_ne:
                            st.error("O número da Nota de Empenho é obrigatório.")
                        else:
                            c = conn.cursor()
                            email_salvar = st.session_state.get('ne_email_auto', '')
                            tel_salvar = st.session_state.get('ne_tel_auto', '')
                            cid_salvar = st.session_state.get('ne_cidade_auto', '')
                            uf_salvar = st.session_state.get('ne_uf_auto', '')
                            sit_salvar = st.session_state.get('ne_sit_auto', 'ATIVA')

                            c.execute('''
                            INSERT INTO notas_empenho (nc_id, om_id, numero_ne, data_emissao, valor_ne, tipo_empenho,
                                                       fornecedor_nome, fornecedor_cnpj, data_envio_empresa, prazo_dias,
                                                       data_limite, prorrogado, nova_data_limite, justificativa_prorrogacao,
                                                       fornecedor_email, fornecedor_telefone, fornecedor_cidade, fornecedor_uf, fornecedor_situacao)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            ''', (nc_id, user['om_id'], numero_ne, str(data_emissao_ne), valor_ne, tipo_empenho,
                                  fornecedor_nome, fornecedor_cnpj, str(data_envio), prazo_dias, str(data_limite_calc),
                                  1 if prorrogado else 0, str(nova_data) if nova_data else None, justificativa,
                                  email_salvar, tel_salvar, cid_salvar, uf_salvar, sit_salvar))
                            conn.commit()

                            for k in ['ne_cnpj_auto', 'ne_nome_auto', 'ne_email_auto', 'ne_tel_auto', 'ne_cidade_auto', 'ne_uf_auto', 'ne_sit_auto']:
                                if k in st.session_state:
                                    del st.session_state[k]

                            st.success(f"✅ Empenho {numero_ne} (R$ {valor_ne:,.2f}) cadastrado com sucesso!")
                            st.rerun()

    with col_ne_top2:
        with st.expander("✏️ Editar Nota de Empenho", expanded=False):
            df_nes_all_top = pd.read_sql_query(f"SELECT * FROM notas_empenho WHERE om_id = {user['om_id']} ORDER BY id DESC", conn)
            if not df_nes_all_top.empty:
                dict_edit_ne = {f"{r['numero_ne']} - {r['fornecedor_nome']} (R$ {r['valor_ne']:,.2f})": r['id'] for _, r in df_nes_all_top.iterrows()}
                sel_ne_ed = st.selectbox("Selecione a NE para editar:", list(dict_edit_ne.keys()), key="sel_ne_to_edit_top")
                id_ne_to_edit = dict_edit_ne[sel_ne_ed]
                ne_dados_atual = dict(df_nes_all_top[df_nes_all_top['id'] == id_ne_to_edit].iloc[0])

                with st.form(f"form_ed_ne_top_{id_ne_to_edit}"):
                    df_ncs_vinc_opts = pd.read_sql_query(f"SELECT id, numero_nc FROM notas_credito WHERE om_id = {user['om_id']}", conn)
                    dict_ncs_v = {r['numero_nc']: r['id'] for _, r in df_ncs_vinc_opts.iterrows()}
                    nc_atual_num = next((k for k, v in dict_ncs_v.items() if v == ne_dados_atual['nc_id']), list(dict_ncs_v.keys())[0] if dict_ncs_v else "")
                    idx_nc_v = list(dict_ncs_v.keys()).index(nc_atual_num) if nc_atual_num in dict_ncs_v else 0
                    
                    ed_nc_vinculada = st.selectbox("Nota de Crédito Vinculada:", list(dict_ncs_v.keys()), index=idx_nc_v)
                    novo_nc_id = dict_ncs_v[ed_nc_vinculada]

                    c_ne1, c_ne2, c_ne3 = st.columns(3)
                    ed_num_ne = c_ne1.text_input("Número do Empenho (NE):", value=ne_dados_atual['numero_ne'])
                    ed_forn_nome = c_ne2.text_input("Razão Social do Fornecedor:", value=ne_dados_atual['fornecedor_nome'])
                    ed_forn_cnpj = c_ne3.text_input("CNPJ do Fornecedor:", value=ne_dados_atual['fornecedor_cnpj'])

                    c_ne4, c_ne5, c_ne6 = st.columns(3)
                    ed_val_ne = c_ne4.number_input("Valor da NE (R$):", value=float(ne_dados_atual['valor_ne']), step=50.0, format="%.2f")
                    tipos_emp_lista = ["Ordinário", "Global", "Estimativo"]
                    idx_tipo_e = tipos_emp_lista.index(ne_dados_atual['tipo_empenho']) if ne_dados_atual['tipo_empenho'] in tipos_emp_lista else 0
                    ed_tipo_emp = c_ne5.selectbox("Tipo de Empenho:", tipos_emp_lista, index=idx_tipo_e)
                    status_lista = ["Aguardando Entrega", "Liquidado", "Pago"]
                    idx_st_e = status_lista.index(ne_dados_atual['status']) if ne_dados_atual['status'] in status_lista else 0
                    ed_status_ne = c_ne6.selectbox("Status:", status_lista, index=idx_st_e)

                    c_ne7, c_ne8, c_ne9 = st.columns(3)
                    dt_em_ne = pd.to_datetime(ne_dados_atual['data_emissao']).date() if pd.notna(ne_dados_atual['data_emissao']) else datetime.now().date()
                    ed_dt_emissao_ne = c_ne7.date_input("Data de Emissão da NE:", value=dt_em_ne, format="DD/MM/YYYY")

                    dt_env_ne = pd.to_datetime(ne_dados_atual['data_envio_empresa']).date() if pd.notna(ne_dados_atual['data_envio_empresa']) else datetime.now().date()
                    ed_dt_envio_ne = c_ne8.date_input("Data de Envio (Marco Zero):", value=dt_env_ne, format="DD/MM/YYYY")

                    ed_prazo_dias = c_ne9.number_input("Prazo de Entrega (Dias):", value=int(ne_dados_atual['prazo_dias'] or 30), step=5)
                    nova_data_calc = ed_dt_envio_ne + timedelta(days=ed_prazo_dias)

                    c_ne10, c_ne11 = st.columns(2)
                    ed_forn_email = c_ne10.text_input("E-mail Oficial da Empresa:", value=ne_dados_atual['fornecedor_email'] or "")
                    ed_forn_tel = c_ne11.text_input("Telefone da Empresa:", value=ne_dados_atual['fornecedor_telefone'] or "")

                    ed_prorrogado = st.checkbox("Houve Prorrogação de Prazo?", value=bool(ne_dados_atual['prorrogado']), key=f"chk_ed_prorr_top_{id_ne_to_edit}")
                    ed_nova_data_limite = None
                    ed_just_prorr = None
                    if ed_prorrogado:
                        cp1, cp2 = st.columns(2)
                        dt_prorr_val = pd.to_datetime(ne_dados_atual['nova_data_limite']).date() if pd.notna(ne_dados_atual['nova_data_limite']) else (nova_data_calc + timedelta(days=15))
                        ed_nova_data_limite = cp1.date_input("Nova Data Limite de Entrega:", value=dt_prorr_val, format="DD/MM/YYYY")
                        ed_just_prorr = cp2.text_area("Justificativa da Prorrogação:", value=ne_dados_atual['justificativa_prorrogacao'] or "")

                    if st.form_submit_button("💾 Salvar Todas as Alterações da NE", type="primary"):
                        c = conn.cursor()
                        c.execute('''
                        UPDATE notas_empenho
                        SET nc_id = ?, numero_ne = ?, fornecedor_nome = ?, fornecedor_cnpj = ?,
                            valor_ne = ?, tipo_empenho = ?, status = ?, data_emissao = ?,
                            data_envio_empresa = ?, prazo_dias = ?, data_limite = ?,
                            prorrogado = ?, nova_data_limite = ?, justificativa_prorrogacao = ?,
                            fornecedor_email = ?, fornecedor_telefone = ?
                        WHERE id = ?
                        ''', (novo_nc_id, ed_num_ne.strip(), ed_forn_nome.strip(), ed_forn_cnpj.strip(),
                              ed_val_ne, ed_tipo_emp, ed_status_ne, str(ed_dt_emissao_ne), str(ed_dt_envio_ne),
                              ed_prazo_dias, str(nova_data_calc), 1 if ed_prorrogado else 0,
                              str(ed_nova_data_limite) if ed_nova_data_limite else None,
                              ed_just_prorr, ed_forn_email.strip(), ed_forn_tel.strip(), id_ne_to_edit))
                        conn.commit()
                        st.success(f"✅ Nota de Empenho {ed_num_ne} atualizada com sucesso!")
                        st.rerun()
            else:
                st.info("Nenhuma NE cadastrada para editar.")

    with col_ne_top3:
        with st.expander("🗑️ Excluir Nota de Empenho", expanded=False):
            df_nes_para_del = pd.read_sql_query(f'''
            SELECT id, numero_ne, valor_ne, fornecedor_nome FROM notas_empenho 
            WHERE om_id = {user['om_id']}
            ORDER BY id DESC
            ''', conn)

            if not df_nes_para_del.empty:
                dict_del_ne = {
                    f"{row['numero_ne']} - {row['fornecedor_nome']} (R$ {row['valor_ne']:,.2f})": row['id']
                    for _, row in df_nes_para_del.iterrows()
                }
                ne_del_sel = st.selectbox("Selecione o Empenho para excluir:", list(dict_del_ne.keys()), key="sel_ne_del_box")
                id_ne_del = dict_del_ne[ne_del_sel]

                c = conn.cursor()
                c.execute("SELECT COUNT(*) FROM notas_fiscais WHERE ne_id = ?", (id_ne_del,))
                qtd_nfs_vinculadas = c.fetchone()[0]

                if qtd_nfs_vinculadas > 0:
                    st.warning(f"⚠️ Não é possível excluir este empenho porque ele possui {qtd_nfs_vinculadas} Nota(s) Fiscal(is) vinculadas no Almoxarifado.")
                else:
                    if st.button("Confirmar Exclusão Definitiva do Empenho", type="primary", key="btn_del_ne_confirm"):
                        c.execute("DELETE FROM notas_empenho WHERE id = ?", (id_ne_del,))
                        conn.commit()
                        st.success("🗑️ Nota de Empenho excluída com sucesso! O saldo retornou para a Nota de Crédito.")
                        st.rerun()
            else:
                st.info("Nenhum empenho cadastrado para exclusão.")

    col_t_ne, col_ord_ne, col_flt_ne = st.columns([6, 1, 1])
    with col_t_ne:
        st.markdown("### 🗂️ Painel Visual de Notas de Empenho")

    with col_ord_ne:
        with st.popover("🔀 Ordenar"):
            ordem_ne_escolhida = st.radio(
                "Critério de Ordenação:",
                [
                    "🚨 Mais urgente ao menos urgente (Padrão)",
                    "💰 Maior valor",
                    "💵 Menor valor",
                    "📅 Data de emissão mais recente"
                ],
                key="radio_ordem_ne"
            )

    with col_flt_ne:
        with st.popover("🌪️ Filtrar"):
            f_ne_atraso = st.checkbox("🔴 Atrasados", value=True, key="chk_ne_atraso")
            f_ne_prazo = st.checkbox("🟢 No Prazo de Entrega", value=True, key="chk_ne_prazo")
            f_ne_liq = st.checkbox("🟡 Liquidados", value=True, key="chk_ne_liq")
            f_ne_pago = st.checkbox("🔵 Pagos", value=True, key="chk_ne_pago")

    df_nes = pd.read_sql_query(f'''
    SELECT ne.id, nc.numero_nc as NC_Origem, ne.numero_ne, ne.tipo_empenho, 
           ne.fornecedor_nome, ne.fornecedor_cnpj, ne.valor_ne, ne.data_emissao, 
           COALESCE(ne.nova_data_limite, ne.data_limite) as data_final,
           ne.status
    FROM notas_empenho ne
    JOIN notas_credito nc ON ne.nc_id = nc.id
    WHERE ne.om_id = {user['om_id']}
    ''', conn)

    if not df_nes.empty:
        hoje = datetime.now().date()

        def classificar_ne(row):
            st_ne = row['status']
            if st_ne == 'Pago':
                return 4, "PAGO", "🔵 PAGO", "#1e3a8a", "border-left: 6px solid #3b82f6;"
            elif st_ne == 'Liquidado':
                return 3, "LIQUIDADO", "🟡 LIQUIDADO", "#854d0e", "border-left: 6px solid #eab308;"
            else:
                dias = None
                if pd.notna(row['data_final']):
                    try:
                        dias = (pd.to_datetime(row['data_final']).date() - hoje).days
                    except Exception:
                        pass
                if dias is not None and dias < 0:
                    return 0, "ATRASADO", f"🔴 ATRASADO ({abs(dias)}d)", "#7f1d1d", "border-left: 6px solid #ef4444;"
                else:
                    return 1, "NO_PRAZO", "🟢 NO PRAZO", "#14532d", "border-left: 6px solid #22c55e;"

        df_nes[['peso_urgencia', 'cat_status', 'badge_txt', 'badge_bg', 'borda_ne']] = df_nes.apply(classificar_ne, axis=1, result_type='expand')

        status_permitidos_ne = []
        if f_ne_atraso: status_permitidos_ne.append("ATRASADO")
        if f_ne_prazo: status_permitidos_ne.append("NO_PRAZO")
        if f_ne_liq: status_permitidos_ne.append("LIQUIDADO")
        if f_ne_pago: status_permitidos_ne.append("PAGO")

        df_nes_filtrado = df_nes[df_nes['cat_status'].isin(status_permitidos_ne)].copy()

        if "Mais urgente" in ordem_ne_escolhida:
            df_nes_filtrado = df_nes_filtrado.sort_values(by=['peso_urgencia', 'valor_ne'], ascending=[True, False])
        elif "Maior valor" in ordem_ne_escolhida:
            df_nes_filtrado = df_nes_filtrado.sort_values(by='valor_ne', ascending=False)
        elif "Menor valor" in ordem_ne_escolhida:
            df_nes_filtrado = df_nes_filtrado.sort_values(by='valor_ne', ascending=True)
        elif "mais recente" in ordem_ne_escolhida:
            df_nes_filtrado = df_nes_filtrado.sort_values(by='data_emissao', ascending=False)

        cols_ne = st.columns(3)
        for idx, row in df_nes_filtrado.reset_index(drop=True).iterrows():
            with cols_ne[idx % 3]:
                st.markdown(f'''
                <div style="background: linear-gradient(145deg, #243142, #1a232e); border: 1.5px solid rgba(255,255,255,0.18); {row['borda_ne']} border-radius: 10px; padding: 15px; margin-bottom: 10px; box-shadow: 0 6px 16px rgba(0,0,0,0.45);">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="font-weight: 900; font-size: 16px; color: #ffffff; letter-spacing: 0.5px;">{row['numero_ne']}</span>
                        <span style="font-size: 11px; padding: 3px 8px; border-radius: 4px; background: {row['badge_bg']}; color: #fff; font-weight: 800;">{row['badge_txt']}</span>
                    </div>
                    <div style="font-size: 14px; color: #f8fafc; margin: 8px 0; font-weight: 700; min-height: 38px; line-height: 1.35;">
                        {row['fornecedor_nome'][:60]}
                    </div>
                    <div style="display: flex; justify-content: space-between; font-size: 12.5px; border-top: 1px solid rgba(255,255,255,0.12); padding-top: 8px;">
                        <span style="color: #cbd5e1;">Valor: <b style="color: #38bdf8; font-size: 14px;">R$ {row['valor_ne']:,.2f}</b></span>
                        <span style="color: #cbd5e1;">Origem: <b style="color: #ffffff;">{row['NC_Origem']}</b></span>
                    </div>
                </div>
                ''', unsafe_allow_html=True)
                with st.popover("🏢 Ficha da Empresa & Detalhes", key=f"pop_ne_{row['id']}", use_container_width=True):
                    c = conn.cursor()
                    c.execute('''
                    SELECT ne.*, nc.numero_nc as NC_Origem
                    FROM notas_empenho ne
                    JOIN notas_credito nc ON ne.nc_id = nc.id
                    WHERE ne.id = ?
                    ''', (row['id'],))
                    ne_info = dict(c.fetchone())

                    st.markdown(f'''
                    <div style="border-bottom: 2px solid #C5A059; padding-bottom: 8px; margin-bottom: 12px;">
                        <div style="font-size: 11px; font-weight: 800; color: #C5A059; text-transform: uppercase; letter-spacing: 0.5px;">Dossiê do Empenho & Fornecedor</div>
                        <div style="font-size: 19px; font-weight: 900; color: #ffffff;">{ne_info['numero_ne']}</div>
                        <div style="font-size: 15px; font-weight: 700; color: #f8fafc; margin-top: 4px;">{ne_info['fornecedor_nome']}</div>
                        <div style="font-size: 13px; color: #cbd5e1;">CNPJ: <b>{ne_info['fornecedor_cnpj']}</b> | NC Origem: <b style="color: #38bdf8;">{ne_info['NC_Origem']}</b></div>
                    </div>
                    ''', unsafe_allow_html=True)

                    st.markdown(f'''
                    <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin: 10px 0;">
                        <div style="background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08); border-left: 3px solid #38bdf8; padding: 7px 10px; border-radius: 6px;">
                            <div style="font-size: 10px; color: #94a3b8; font-weight: 700; text-transform: uppercase;">Valor do Empenho</div>
                            <div style="font-size: 15px; font-weight: 800; color: #38bdf8; white-space: nowrap; margin-top: 2px;">R$ {ne_info['valor_ne']:,.2f}</div>
                        </div>
                        <div style="background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08); border-left: 3px solid #ffffff; padding: 7px 10px; border-radius: 6px;">
                            <div style="font-size: 10px; color: #94a3b8; font-weight: 700; text-transform: uppercase;">Data de Emissão</div>
                            <div style="font-size: 14px; font-weight: 800; color: #ffffff; white-space: nowrap; margin-top: 2px;">{formatar_data_br(ne_info['data_emissao'])}</div>
                        </div>
                        <div style="background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08); border-left: 3px solid #fbbf24; padding: 7px 10px; border-radius: 6px;">
                            <div style="font-size: 10px; color: #94a3b8; font-weight: 700; text-transform: uppercase;">Limite de Entrega</div>
                            <div style="font-size: 14px; font-weight: 800; color: #fbbf24; white-space: nowrap; margin-top: 2px;">{formatar_data_br(ne_info['data_limite'])}</div>
                        </div>
                    </div>
                    '''.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)

                    st.markdown("---")
                    st.markdown(f"📧 **E-mail Oficial:** [{ne_info['fornecedor_email']}](mailto:{ne_info['fornecedor_email']})" if ne_info['fornecedor_email'] else "📧 **E-mail Oficial:** *Não informado*")
                    st.markdown(f"📞 **Telefone:** `{ne_info['fornecedor_telefone']}`" if ne_info['fornecedor_telefone'] else "📞 **Telefone:** *Não informado*")

                    with st.expander("✏️ Editar Todos os Dados desta NE", expanded=False):
                        with st.form(f"form_ed_ne_card_{row['id']}"):
                            df_ncs_vinc_card = pd.read_sql_query(f"SELECT id, numero_nc FROM notas_credito WHERE om_id = {user['om_id']}", conn)
                            dict_ncs_card = {r['numero_nc']: r['id'] for _, r in df_ncs_vinc_card.iterrows()}
                            nc_card_atual_num = next((k for k, v in dict_ncs_card.items() if v == ne_info['nc_id']), list(dict_ncs_card.keys())[0] if dict_ncs_card else "")
                            idx_nc_c = list(dict_ncs_card.keys()).index(nc_card_atual_num) if nc_card_atual_num in dict_ncs_card else 0
                            
                            c_ed_nc = st.selectbox("Vincular à Nota de Crédito:", list(dict_ncs_card.keys()), index=idx_nc_c, key=f"c_ed_nc_{row['id']}")
                            novo_nc_id_card = dict_ncs_card[c_ed_nc]

                            ed_c1, ed_c2, ed_c3 = st.columns(3)
                            ed_num = ed_c1.text_input("Número NE:", value=ne_info['numero_ne'], key=f"ed_num_{row['id']}")
                            ed_nome = ed_c2.text_input("Fornecedor:", value=ne_info['fornecedor_nome'], key=f"ed_nome_{row['id']}")
                            ed_cnpj = ed_c3.text_input("CNPJ:", value=ne_info['fornecedor_cnpj'], key=f"ed_cnpj_{row['id']}")

                            ed_c4, ed_c5, ed_c6 = st.columns(3)
                            ed_val = ed_c4.number_input("Valor (R$):", value=float(ne_info['valor_ne']), step=50.0, format="%.2f", key=f"ed_val_{row['id']}")
                            tipos_e = ["Ordinário", "Global", "Estimativo"]
                            ed_tipo = ed_c5.selectbox("Tipo:", tipos_e, index=tipos_e.index(ne_info['tipo_empenho']) if ne_info['tipo_empenho'] in tipos_e else 0, key=f"ed_tipo_{row['id']}")
                            sts_e = ["Aguardando Entrega", "Liquidado", "Pago"]
                            ed_st = ed_c6.selectbox("Status:", sts_e, index=sts_e.index(ne_info['status']) if ne_info['status'] in sts_e else 0, key=f"ed_st_{row['id']}")

                            ed_c7, ed_c8, ed_c9 = st.columns(3)
                            dt_em_c = pd.to_datetime(ne_info['data_emissao']).date() if pd.notna(ne_info['data_emissao']) else datetime.now().date()
                            ed_dt_em = ed_c7.date_input("Data de Emissão:", value=dt_em_c, format="DD/MM/YYYY", key=f"ed_dtem_{row['id']}")
                            
                            dt_env_c = pd.to_datetime(ne_info['data_envio_empresa']).date() if pd.notna(ne_info['data_envio_empresa']) else datetime.now().date()
                            ed_dt_env = ed_c8.date_input("Data Envio (Marco Zero):", value=dt_env_c, format="DD/MM/YYYY", key=f"ed_dtenv_{row['id']}")
                            
                            ed_pz = ed_c9.number_input("Prazo de Entrega (Dias):", value=int(ne_info['prazo_dias'] or 30), step=5, key=f"ed_pz_{row['id']}")
                            calc_lim_card = ed_dt_env + timedelta(days=ed_pz)

                            ed_email = st.text_input("E-mail Oficial:", value=ne_info['fornecedor_email'] or "", key=f"ed_em_{row['id']}")
                            ed_tel = st.text_input("Telefone de Contato:", value=ne_info['fornecedor_telefone'] or "", key=f"ed_tl_{row['id']}")

                            ed_prorr_c = st.checkbox("Prorrogação de Prazo Solicitada?", value=bool(ne_info['prorrogado']), key=f"chk_prorr_card_{row['id']}")
                            nova_dt_c = None
                            just_c = None
                            if ed_prorr_c:
                                cp1, cp2 = st.columns(2)
                                dt_prorr_def = pd.to_datetime(ne_info['nova_data_limite']).date() if pd.notna(ne_info['nova_data_limite']) else (calc_lim_card + timedelta(days=15))
                                nova_dt_c = cp1.date_input("Nova Data Limite:", value=dt_prorr_def, format="DD/MM/YYYY", key=f"novadt_{row['id']}")
                                just_c = cp2.text_area("Justificativa:", value=ne_info['justificativa_prorrogacao'] or "", key=f"just_{row['id']}")

                            if st.form_submit_button("💾 Salvar Todas as Alterações da NE", type="primary"):
                                c.execute('''
                                UPDATE notas_empenho
                                SET nc_id = ?, numero_ne = ?, fornecedor_nome = ?, fornecedor_cnpj = ?, valor_ne = ?,
                                    tipo_empenho = ?, status = ?, data_emissao = ?, data_envio_empresa = ?,
                                    prazo_dias = ?, data_limite = ?, prorrogado = ?, nova_data_limite = ?,
                                    justificativa_prorrogacao = ?, fornecedor_email = ?, fornecedor_telefone = ?
                                WHERE id = ?
                                ''', (novo_nc_id_card, ed_num.strip(), ed_nome.strip(), ed_cnpj.strip(), ed_val,
                                      ed_tipo, ed_st, str(ed_dt_em), str(ed_dt_env), ed_pz, str(calc_lim_card),
                                      1 if ed_prorr_c else 0, str(nova_dt_c) if nova_dt_c else None, just_c,
                                      ed_email.strip(), ed_tel.strip(), row['id']))
                                conn.commit()
                                st.success("✅ Nota de Empenho atualizada com sucesso!")
                                st.rerun()
    conn.close()

# ==========================================
# 4. ELABORAÇÃO DE PROCESSOS (CABECALHO FIXO B ADM AP / 5ª RM & LIBREOFFICE/WORD)
# ==========================================
elif menu_selecionado == "📝 Elaboração de Processos":
    st.title("📝 Elaboração de Processos de Aquisição")
    st.info("Selecione o tipo de processo e o modelo para carregar os campos específicos e gerar o documento com formatação oficial.")
    conn = get_connection()

    aba_proc1, aba_proc2 = st.tabs(["📄 Elaborar Processo por Modelo", "📁 Banco de Modelos (LibreOffice / Word)"])

    with aba_proc1:
        st.markdown("### 1. Seleção do Processo e Modelo")
        col_m1, col_m2 = st.columns(2)
        
        tipo_proc_sel = col_m1.selectbox(
            "Modalidade / Tipo de Processo:",
            [
                "Pregão Eletrônico (Próprio)",
                "Pregão Eletrônico (Participante)",
                "Adesão à Ata de Registro de Preços (Carona)",
                "Dispensa de Licitação (Art. 75 Lei 14.133/21)",
                "Inexigibilidade de Licitação"
            ],
            key="sel_tipo_proc_box"
        )

        c = conn.cursor()
        c.execute('''
        SELECT id, nome_modelo, tipo_processo, formato, arquivo_docx 
        FROM processos_modelos 
        WHERE om_id = ? AND (tipo_processo = ? OR tipo_processo LIKE '%Requisição%' OR tipo_processo LIKE '%Pregão%')
        ''', (user['om_id'], tipo_proc_sel))
        modelos_banco = c.fetchall()

        opcoes_modelos_dict = {"02. Requisição de Despesa (Modelo Oficial - B Adm Ap / 5ª RM)": None}
        for m_row in modelos_banco:
            opcoes_modelos_dict[f"{m_row['nome_modelo']} ({m_row['formato'].upper()})"] = m_row['id']

        modelo_escolhido_nome = col_m2.selectbox(
            "Documento / Modelo a Elaborar:",
            list(opcoes_modelos_dict.keys()),
            key="sel_modelo_doc_box"
        )
        modelo_id_selecionado = opcoes_modelos_dict[modelo_escolhido_nome]

        st.markdown("---")

        st.markdown("### 2. Dados da Requisição")
        st.caption("🔒 O cabeçalho oficial da Base de Administração e Apoio da 5ª RM é fixo e inalterável.")
        
        col_r1, col_r2, col_r3 = st.columns(3)
        num_req = col_r1.text_input("Número da Requisição:", value="5-Seção de Contratações/B Adm Ap")
        nup_req = col_r2.text_input("NUP / Número do EB:", value="65378.006979/2026-34")
        assunto_req = col_r3.text_input("Assunto da Requisição:", value="aquisição de material de expediente (PRÓPRIO)")
        
        data_extenso_req = st.text_input("Local e Data por Extenso:", value="Curitiba, PR, 22 de junho de 2026")

        st.markdown("### 3. Demonstrativo de Necessidades & Fornecedor")
        col_p1, col_p2 = st.columns(2)
        pregao_req = col_p1.text_input("Identificação do Pregão / Processo de Origem:", value="Pregão Eletrônico 90001/2026, da B Adm Ap /5ª RM (UG 160192)")
        tipo_emp_req = col_p2.selectbox("Tipo de Empenho:", ["Ordinário", "Estimativo", "Global"])

        col_cnpj_p1, col_cnpj_p2 = st.columns([3, 1])
        cnpj_forn_req = col_cnpj_p1.text_input("CNPJ do Fornecedor:", value="09.247.343/0001-20")
        if col_cnpj_p2.button("🔍 Buscar Receita", use_container_width=True):
            d_emp = buscar_dados_cnpj(cnpj_forn_req, conn)
            if d_emp:
                st.session_state['req_forn_nome'] = d_emp.get('razao_social', '')
                st.success(f"Empresa: {st.session_state['req_forn_nome']}")
        
        razao_social_req = st.text_input("Razão Social do Fornecedor:", value=st.session_state.get('req_forn_nome', "BIOLIMP SERVICOS ESPECIALIZADOS DE HIGIENIZACAO TEXTIL EIRELI"))

        st.markdown("### 4. Tabela de Itens da Requisição")
        if 'itens_req_df' not in st.session_state:
            st.session_state['itens_req_df'] = pd.DataFrame([
                {"Item": 1, "Descrição": "COPO DESCARTAVEL 200 ML", "UN": "CX", "QTD": 2.0, "V_Unit": 92.00},
                {"Item": 2, "Descrição": "DETERGENTE LIQUIDO NEUTRO", "UN": "UN", "QTD": 78.0, "V_Unit": 1.18}
            ])

        df_edit_req = st.data_editor(
            st.session_state['itens_req_df'],
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Item": st.column_config.NumberColumn("Item Nº", min_value=1, step=1),
                "Descrição": st.column_config.TextColumn("Descrição Detalhada", width="large", required=True),
                "UN": st.column_config.TextColumn("UN", width="small", default="UN"),
                "QTD": st.column_config.NumberColumn("QTD", min_value=1.0, step=1.0),
                "V_Unit": st.column_config.NumberColumn("Valor Unitário (R$)", min_value=0.01, step=0.1, format="R$ %.2f")
            }
        )
        st.session_state['itens_req_df'] = df_edit_req
        st.caption("ℹ️ * Inserir os dados conforme consta no Termo de Homologação.")

        df_calc_req = df_edit_req.copy()
        df_calc_req['Total'] = df_calc_req['QTD'] * df_calc_req['V_Unit']
        total_requisicao = df_calc_req['Total'].sum()

        st.markdown(f'''
        <div style="background: rgba(34, 197, 94, 0.15); border: 1.5px solid #22c55e; border-radius: 8px; padding: 12px 18px; display: flex; justify-content: space-between; align-items: center; margin: 15px 0;">
            <span style="font-size: 15px; font-weight: bold; color: #fff;">VALOR TOTAL DA REQUISIÇÃO:</span>
            <span style="font-size: 22px; font-weight: 900; color: #22c55e;">R$ {total_requisicao:,.2f}</span>
        </div>
        ''', unsafe_allow_html=True)

        st.markdown("### 5. Local de Entrega e Justificativas da Despesa")
        local_entrega_req = st.text_area(
            "3. Local de Entrega dos Materiais / Prestação dos Serviços:",
            value="Base de Administração e Apoio da 5ª Região Militar, situado à Rua 31 de Março, S/Nr, CEP 81.150-900, Bairro Pinheirinho, em Curitiba, PR.",
            height=70
        )

        col_j1, col_j2 = st.columns(2)
        just_a = col_j1.text_area(
            "4. a. Necessidade (Problema a ser resolvido):",
            value="a contratação/aquisição é necessária para garantir o funcionamento administrativo e operacional desta OM, visto que os estoques atuais se esgotaram.",
            height=100
        )
        just_b = col_j2.text_area(
            "4. b. Consequência da não emissão (Prejuízo potencial):",
            value="a não emissão desta NE acarretará prejuízo ao funcionamento regular da OM e à execução de suas atividades administrativas e operacionais.",
            height=100
        )

        st.markdown("### 6. Classificação Orçamentária")
        df_ncs_disp = pd.read_sql_query(f"SELECT id, numero_nc, pi, natureza_despesa FROM notas_credito WHERE om_id = {user['om_id']}", conn)
        nc_selec_req = None
        pi_padrao = "IXAPFUNADOM"
        nd_padrao = "339030"
        
        if not df_ncs_disp.empty:
            dict_nc_opcoes = {f"{r['numero_nc']} (PI: {r['pi']})": r for _, r in df_ncs_disp.iterrows()}
            escolha_nc = st.selectbox("Puxar Dados da Nota de Crédito da OM:", list(dict_nc_opcoes.keys()))
            nc_escolhida_row = dict_nc_opcoes[escolha_nc]
            nc_selec_req = nc_escolhida_row['numero_nc']
            pi_padrao = nc_escolhida_row['pi'] or pi_padrao
            nd_padrao = nc_escolhida_row['natureza_despesa'] or nd_padrao

        col_orc1, col_orc2, col_orc3, col_orc4 = st.columns(4)
        pi_req = col_orc1.text_input("Plano Interno (PI):", value=pi_padrao)
        ptres_req = col_orc2.text_input("PTRES:", value="171502")
        nd_req = col_orc3.text_input("Natureza da Despesa (ND):", value=nd_padrao)
        nc_num_req = col_orc4.text_input("Número da NC:", value=nc_selec_req if nc_selec_req else "2026NC000672")

        st.markdown("### 7. Assinatura do Fiscal Administrativo")
        col_fisc1, col_fisc2 = st.columns(2)
        fisc_nome = col_fisc1.text_input("Nome Completo e Posto/Graduação:", value="MARCO AURÉLIO GOBETTI DA FONSECA - Cap")
        fisc_funcao = col_fisc2.text_input("Função do Fiscal:", value="Fiscal Administrativo - B Adm Ap/5ª RM")

        st.markdown("---")
        
        st.markdown("### 👁️ Pré-Visualização da Folha Oficial (Padrão A4)")
        
        linhas_tabela_html = ""
        for it in df_calc_req.to_dict('records'):
            linhas_tabela_html += f"""
            <tr style="border-bottom: 1px solid #111;">
                <td style="padding: 6px; text-align: center; border-right: 1px solid #111;">{it['Item']}</td>
                <td style="padding: 6px; border-right: 1px solid #111;">{it['Descrição']}</td>
                <td style="padding: 6px; text-align: center; border-right: 1px solid #111;">{it['UN']}</td>
                <td style="padding: 6px; text-align: center; border-right: 1px solid #111;">{it['QTD']:.0f}</td>
                <td style="padding: 6px; text-align: right; border-right: 1px solid #111;">R$ {it['V_Unit']:.2f}</td>
                <td style="padding: 6px; text-align: right;">R$ {it['Total']:.2f}</td>
            </tr>
            """

        st.markdown(f"""
        <div style="background: #ffffff; color: #111827; padding: 45px 50px; border-radius: 4px; box-shadow: 0 10px 25px rgba(0,0,0,0.5); max-width: 820px; margin: 15px auto; font-family: 'Times New Roman', Times, serif; font-size: 14px; line-height: 1.4; border: 1px solid #d1d5db;">
            <div style="text-align: center; margin-bottom: 15px;">
                <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/b/bf/Coat_of_arms_of_Brazil.svg/150px-Coat_of_arms_of_Brazil.svg.png" style="height: 72px; width: auto; margin-bottom: 6px;">
                <div style="font-weight: bold; font-size: 13px; line-height: 1.25; color: #000;">
                    MINISTÉRIO DA DEFESA<br>
                    EXÉRCITO BRASILEIRO<br>
                    BASE DE ADMINISTRAÇÃO E APOIO DA 5ª REGIÃO MILITAR<br>
                    <span style="font-weight: normal; font-size: 12px;">(Companhia do QG da 5ª RM/DI)</span><br>
                    <span style="font-weight: bold; font-size: 12px;">BASE MAJOR AGOSTINHO JOSÉ RODRIGUES</span>
                </div>
            </div>
            
            <div style="margin-top: 20px;">
                <b>Requisição Nº</b> {num_req}<br>
                <b>EB:</b> {nup_req}<br>
                <b>Assunto:</b> {assunto_req}<br>
                <b>Anexos:</b> Termo de Homologação, Nota de Crédito, CADIN, SICAF, TCU
            </div>
            
            <div style="text-align: right; margin: 15px 0;">
                Curitiba, PR, {data_extenso_req}.
            </div>
            
            <p style="text-align: justify; text-indent: 30px;">
                1. Nos termos contidos no Art. 13 das IG 12-02, aprovadas pela Portaria Ministerial nº 305, de 24 MAIO 95, solicito providências no sentido de aprovar a seguinte despesa com aquisição do material abaixo especificado, a fim de atender necessidades desta OM.
            </p>
            
            <p><b>2. DEMONSTRATIVO DE NECESSIDADES:</b><br>
            {pregao_req}<br>
            <b>RAZÃO SOCIAL:</b> {razao_social_req} — <b>CNPJ:</b> {cnpj_forn_req}</p>
            
            <table style="width: 100%; border-collapse: collapse; border: 1.5px solid #111; font-size: 12px; margin: 10px 0;">
                <thead>
                    <tr style="background-color: #f3f4f6; border-bottom: 1.5px solid #111;">
                        <th style="padding: 6px; border-right: 1px solid #111; width: 8%;">Item</th>
                        <th style="padding: 6px; border-right: 1px solid #111; text-align: left;">Descrição</th>
                        <th style="padding: 6px; border-right: 1px solid #111; width: 8%;">UN</th>
                        <th style="padding: 6px; border-right: 1px solid #111; width: 8%;">QTD</th>
                        <th style="padding: 6px; border-right: 1px solid #111; width: 18%; text-align: right;">VALOR UNITÁRIO</th>
                        <th style="padding: 6px; width: 18%; text-align: right;">TOTAL</th>
                    </tr>
                </thead>
                <tbody>
                    {linhas_tabela_html}
                </tbody>
            </table>
            
            <div style="text-align: right; font-weight: bold; font-size: 13px; margin: 5px 0 10px 0;">
                VALOR TOTAL: R$ {total_requisicao:,.2f}
            </div>
            
            <p><b>TIPO DE EMPENHO:</b> ({'X' if tipo_emp_req=='Ordinário' else ' '}) Ordinário &nbsp;&nbsp;&nbsp;&nbsp; ({'X' if tipo_emp_req=='Estimativo' else ' '}) Estimativo &nbsp;&nbsp;&nbsp;&nbsp; ({'X' if tipo_emp_req=='Global' else ' '}) Global</p>
            
            <p><b>3. ENTREGA DO(S) BEM(NS)/PRESTAÇÃO DO(S) SERVIÇO(S):</b><br>
            Os materiais serão entregues no seguinte local: {local_entrega_req}</p>
            
            <p><b>4. JUSTIFICATIVA DA DESPESA:</b><br>
            <b>a. Necessidade:</b> {just_a}<br>
            <b>b. Consequência da não emissão:</b> {just_b}</p>
            
            <p><b>5. CLASSIFICAÇÃO ORÇAMENTÁRIA:</b></p>
            <table style="width: 100%; border-collapse: collapse; border: 1px solid #111; font-size: 12px; margin-bottom: 15px;">
                <tr style="background: #f3f4f6; border-bottom: 1px solid #111; text-align: center; font-weight: bold;">
                    <td style="padding: 5px; border-right: 1px solid #111;">PI</td>
                    <td style="padding: 5px; border-right: 1px solid #111;">PTRES</td>
                    <td style="padding: 5px; border-right: 1px solid #111;">ND</td>
                    <td style="padding: 5px;">NC</td>
                </tr>
                <tr style="text-align: center;">
                    <td style="padding: 5px; border-right: 1px solid #111;">{pi_req}</td>
                    <td style="padding: 5px; border-right: 1px solid #111;">{ptres_req}</td>
                    <td style="padding: 5px; border-right: 1px solid #111;">{nd_req}</td>
                    <td style="padding: 5px;">{nc_num_req}</td>
                </tr>
            </table>
            
            <div style="text-align: center; margin-top: 35px;">
                Aquisição aprovada por:<br><br>
                Curitiba, PR, {data_extenso_req}.<br><br><br>
                ____________________________________________<br>
                <b>{fisc_nome}</b><br>
                {fisc_funcao}
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### 8. Download do Documento Pronto")
        col_dl1, col_dl2 = st.columns(2)

        dados_geracao = {
            "numero_req": num_req, "nup": nup_req, "assunto": assunto_req,
            "data_extenso": data_extenso_req, "pregao": pregao_req, "tipo_empenho": tipo_emp_req,
            "razao_social": razao_social_req, "cnpj": cnpj_forn_req, "local_entrega": local_entrega_req,
            "justificativa_a": just_a, "justificativa_b": just_b, "pi": pi_req, "ptres": ptres_req,
            "nd": nd_req, "nc": nc_num_req, "fiscal_nome_posto": fisc_nome, "fiscal_funcao": fisc_funcao
        }

        # Geração em LibreOffice (.odt)
        import zipfile
        buffer_odt = io.BytesIO()
        content_odt_xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" 
                         xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" 
                         xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" 
                         xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" 
                         office:version="1.2">
  <office:body>
    <office:text>
      <text:h text:outline-level="1">MINISTÉRIO DA DEFESA - EXÉRCITO BRASILEIRO</text:h>
      <text:h text:outline-level="2">BASE DE ADMINISTRAÇÃO E APOIO DA 5ª REGIÃO MILITAR</text:h>
      <text:p>(Companhia do QG da 5ª RM/DI) - BASE MAJOR AGOSTINHO JOSÉ RODRIGUES</text:p>
      <text:p/>
      <text:p>Requisição Nº {num_req} | EB: {nup_req}</text:p>
      <text:p>Assunto: {assunto_req}</text:p>
      <text:p>Curitiba, PR, {data_extenso_req}.</text:p>
      <text:p/>
      <text:p>1. Solicito providências para aprovar a despesa abaixo especificada.</text:p>
      <text:p>2. DEMONSTRATIVO DE NECESSIDADES: {pregao_req}</text:p>
      <text:p>FORNECEDOR: {razao_social_req} (CNPJ: {cnpj_forn_req})</text:p>
      <text:p>VALOR TOTAL: R$ {total_requisicao:,.2f}</text:p>
      <text:p>3. LOCAL DE ENTREGA: {local_entrega_req}</text:p>
      <text:p>4. JUSTIFICATIVA: {just_a}</text:p>
      <text:p>5. CLASSIFICAÇÃO: PI {pi_req} | PTRES {ptres_req} | ND {nd_req} | NC {nc_num_req}</text:p>
      <text:p/>
      <text:p>Aprovado por: {fisc_nome} - {fisc_funcao}</text:p>
    </office:text>
  </office:body>
</office:document-content>'''
        manifest_odt_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>'''
        with zipfile.ZipFile(buffer_odt, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr('mimetype', 'application/vnd.oasis.opendocument.text', compress_type=zipfile.ZIP_STORED)
            zf.writestr('content.xml', content_odt_xml.encode('utf-8'))
            zf.writestr('META-INF/manifest.xml', manifest_odt_xml.encode('utf-8'))
        buffer_odt.seek(0)

        col_dl1.download_button(
            "📥 Baixar em LibreOffice Writer (.odt)",
            data=buffer_odt,
            file_name=f"Requisicao_{num_req.replace('/', '_').replace(' ', '_')}.odt",
            mime="application/vnd.oasis.opendocument.text",
            use_container_width=True
        )

        # Geração em Word (.docx)
        if docx is not None:
            doc_final = None
            if modelo_id_selecionado:
                c.execute("SELECT arquivo_docx FROM processos_modelos WHERE id = ?", (modelo_id_selecionado,))
                res_bin = c.fetchone()
                if res_bin and res_bin[0]:
                    try:
                        doc_final = docx.Document(io.BytesIO(res_bin[0]))
                    except Exception:
                        doc_final = None
            
            if doc_final is None:
                doc_final = docx.Document()
                p_c = doc_final.add_paragraph()
                p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_c = p_c.add_run("MINISTÉRIO DA DEFESA\nEXÉRCITO BRASILEIRO\nBASE DE ADMINISTRAÇÃO E APOIO DA 5ª REGIÃO MILITAR\n(Companhia do QG da 5ª RM/DI)\nBASE MAJOR AGOSTINHO JOSÉ RODRIGUES\n")
                r_c.bold = True
                doc_final.add_paragraph(f"Requisição Nº {num_req}\nEB: {nup_req}\nAssunto: {assunto_req}\nCuritiba, PR, {data_extenso_req}.\n")
                doc_final.add_paragraph(f"DEMONSTRATIVO: {pregao_req}\nFornecedor: {razao_social_req} ({cnpj_forn_req})\nValor: R$ {total_requisicao:,.2f}")
                doc_final.add_paragraph(f"Local de Entrega: {local_entrega_req}\nJustificativa: {just_a}\nClassificação: PI {pi_req} | ND {nd_req} | NC {nc_num_req}")
                p_a = doc_final.add_paragraph(f"\n\n____________________________________________\n{fisc_nome}\n{fisc_funcao}")
                p_a.alignment = WD_ALIGN_PARAGRAPH.CENTER

            buf_docx = io.BytesIO()
            doc_final.save(buf_docx)
            buf_docx.seek(0)

            col_dl2.download_button(
                "📥 Baixar em Word (.docx)",
                data=buf_docx,
                file_name=f"Requisicao_{num_req.replace('/', '_').replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

    with aba_proc2:
        st.markdown("### 📁 Modelos Oficiais Vigentes (LibreOffice e Word)")
        st.caption("Faça o upload dos modelos em formato .odt ou .docx para mantê-los centralizados no banco de dados da OM.")

        c = conn.cursor()
        c.execute("SELECT nome_tipo FROM tipos_documentos WHERE om_id = ? ORDER BY nome_tipo", (user['om_id'],))
        tipos_banco = [r[0] for r in c.fetchall()]

        if not tipos_banco:
            tipos_padrao = [
                "Requisição de Despesa (Pregão Próprio)",
                "Termo de Referência (TR)",
                "Estudo Técnico Preliminar (ETP)",
                "Documento de Formalização da Demanda (DFD)",
                "Pesquisa de Preços / Mapa Comparativo",
                "Justificativa de Adesão (Carona)",
                "Despacho do Ordenador de Despesas"
            ]
            for tp in tipos_padrao:
                c.execute("INSERT OR IGNORE INTO tipos_documentos (om_id, nome_tipo) VALUES (?, ?)", (user['om_id'], tp))
            conn.commit()
            tipos_banco = tipos_padrao

        col_up_m1, col_up_m2 = st.columns(2)
        arq_mod = col_up_m1.file_uploader("Carregar Arquivo de Modelo (.odt ou .docx):", type=["odt", "docx"], key="up_model_file")

        opcoes_tipo_doc = tipos_banco + ["➕ Adicionar Novo Tipo de Documento..."]
        tipo_mod_sel = col_up_m2.selectbox("Tipo de Documento:", opcoes_tipo_doc, key="sel_tipo_mod_box")

        tipo_documento_final = tipo_mod_sel
        if tipo_mod_sel == "➕ Adicionar Novo Tipo de Documento...":
            novo_tipo_digitado = st.text_input("Digite o Nome do Novo Tipo de Documento:", key="input_novo_tipo_doc")
            if novo_tipo_digitado.strip():
                tipo_documento_final = novo_tipo_digitado.strip()

        nome_modelo_custom = st.text_input(
            "Nome de Identificação do Modelo:", 
            value=f"Modelo Padrão - {tipo_documento_final}" if not tipo_documento_final.startswith("➕") else "Novo Modelo",
            key="nome_mod_custom_input"
        )

        if st.button("💾 Salvar Modelo no Banco de Dados", type="primary"):
            if arq_mod is not None:
                if not tipo_documento_final or tipo_documento_final.startswith("➕"):
                    st.error("Por favor, informe um tipo de documento válido.")
                else:
                    bytes_mod = arq_mod.read()
                    extensao = arq_mod.name.split('.')[-1].lower()
                    
                    c.execute("INSERT OR IGNORE INTO tipos_documentos (om_id, nome_tipo) VALUES (?, ?)", (user['om_id'], tipo_documento_final))
                    
                    try:
                        c.execute("ALTER TABLE processos_modelos ADD COLUMN formato TEXT DEFAULT 'odt'")
                    except sqlite3.OperationalError:
                        pass

                    c.execute('''
                    INSERT INTO processos_modelos (om_id, tipo_processo, nome_modelo, arquivo_docx, formato, data_upload)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ''', (user['om_id'], tipo_documento_final, nome_modelo_custom, bytes_mod, extensao, str(datetime.now().date())))
                    conn.commit()
                    
                    st.success(f"✅ Modelo '{nome_modelo_custom}' ({extensao.upper()}) cadastrado com sucesso!")
                    st.rerun()
            else:
                st.warning("Selecione um arquivo .odt ou .docx primeiro.")

        st.markdown("---")
        st.markdown("#### 📑 Modelos Já Cadastrados na OM")
        
        try:
            c.execute("ALTER TABLE processos_modelos ADD COLUMN formato TEXT DEFAULT 'docx'")
            conn.commit()
        except sqlite3.OperationalError:
            pass

        df_modelos = pd.read_sql_query(f"SELECT id, tipo_processo, nome_modelo, COALESCE(formato, 'docx') as formato, data_upload FROM processos_modelos WHERE om_id = {user['om_id']}", conn)
        if not df_modelos.empty:
            df_modelos['Data Upload'] = df_modelos['data_upload'].apply(formatar_data_br)
            df_modelos['Formato'] = df_modelos['formato'].apply(lambda f: f"📄 LibreOffice ({f.upper()})" if f == 'odt' else f"📄 Word ({f.upper()})")
            st.dataframe(df_modelos[['tipo_processo', 'nome_modelo', 'Formato', 'Data Upload']], use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum modelo cadastrado até o momento.")

    conn.close()

# ==========================================
# 5. RECEBIMENTO (NFs)
# ==========================================
elif menu_selecionado == "📦 Recebimento":
    st.title("📦 Recebimento de Materiais e Notas Fiscais")
    conn = get_connection()
    df_nfs = pd.read_sql_query(f"SELECT nf.id, ne.numero_ne, nf.numero_nf, nf.empresa_cnpj, nf.data_entrada_almox, nf.valor_nf, nf.situacao FROM notas_fiscais nf JOIN notas_empenho ne ON nf.ne_id = ne.id WHERE nf.om_id = {user['om_id']}", conn)
    if not df_nfs.empty:
        df_nfs['Data Entrada Almox'] = df_nfs['data_entrada_almox'].apply(formatar_data_br)
        st.dataframe(df_nfs[['numero_ne', 'numero_nf', 'empresa_cnpj', 'Data Entrada Almox', 'valor_nf', 'situacao']], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma Nota Fiscal registrada.")
    conn.close()

# ==========================================
# 6. GESTÃO DE ESTOQUE (MÉTRICAS MILITARES AUTOMÁTICAS & HISTÓRICO ANUAL)
# ==========================================
elif menu_selecionado == "🏢 Gestão de Estoque":
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

elif menu_selecionado == "📁 Relatórios":
    st.title("📁 Central de Relatórios Oficiais com Filtros")
    conn = get_connection()

    tipo_relatorio = st.selectbox("Selecione o Relatório a Emitir:", ["Notas de Crédito (com Empenhos Vinculados)", "Notas de Empenho", "Notas Fiscais (Recebimento)"])

    if "Notas de Crédito" in tipo_relatorio:
        st.markdown("##### 🔍 Filtros de Consulta das Notas de Crédito")
        fc1, fc2, fc3 = st.columns(3)
        df_ncs_base = pd.read_sql_query(f"SELECT DISTINCT enquadramento FROM notas_credito WHERE om_id = {user['om_id']}", conn)
        
        enq_opts = ["Todos"] + sorted(list(set(df_ncs_base['enquadramento'].dropna().tolist())))
        f_enq = fc1.selectbox("Enquadramento:", enq_opts, key="f_enq_rel")
        
        tipo_opts = ["Todos", "Consumo (33)", "Permanente (44)"]
        f_tipo = fc2.selectbox("Natureza da Despesa:", tipo_opts, key="f_tipo_rel")
        busca_texto = fc3.text_input("Palavra-chave na Finalidade:", key="f_txt_rel")

        q_rel_nc = f"""
        SELECT nc.numero_nc as "Número NC",
               nc.data_emissao,
               nc.data_limite_empenho,
               nc.enquadramento as "Enquadramento",
               nc.natureza_despesa as "ND",
               nc.pi as "PI",
               nc.valor_total as "Valor Total (R$)",
               COALESCE((SELECT SUM(valor_ne) FROM notas_empenho WHERE nc_id = nc.id), 0.0) as "Total Empenhado (R$)",
               (nc.valor_total - COALESCE(nc.valor_recolhido, 0.0) - COALESCE((SELECT SUM(valor_ne) FROM notas_empenho WHERE nc_id = nc.id), 0.0)) as "Saldo Disponível (R$)",
               COALESCE((SELECT GROUP_CONCAT(numero_ne, ', ') FROM notas_empenho WHERE nc_id = nc.id), 'Sem empenho') as "Empenhos Vinculados",
               nc.finalidade as "Finalidade"
        FROM notas_credito nc
        WHERE nc.om_id = {user['om_id']}
        """
        if f_enq != "Todos":
            q_rel_nc += f" AND nc.enquadramento = '{f_enq}'"
        if f_tipo == "Consumo (33)":
            q_rel_nc += " AND nc.natureza_despesa LIKE '%33%'"
        elif f_tipo == "Permanente (44)":
            q_rel_nc += " AND nc.natureza_despesa LIKE '%44%'"
        if busca_texto:
            q_rel_nc += f" AND nc.finalidade LIKE '%{busca_texto}%'"

        q_rel_nc += " ORDER BY nc.data_emissao DESC"

        df_rel_nc = pd.read_sql_query(q_rel_nc, conn)
        if not df_rel_nc.empty:
            df_rel_nc['Data Emissão'] = df_rel_nc['data_emissao'].apply(formatar_data_br)
            df_rel_nc['Dt Lim Empenho'] = df_rel_nc['data_limite_empenho'].apply(formatar_data_br)
            
            # Formata colunas para exibição amigável
            cols_ordem = [
                "Número NC", "Data Emissão", "Dt Lim Empenho", "Enquadramento", "ND", "PI",
                "Valor Total (R$)", "Total Empenhado (R$)", "Saldo Disponível (R$)", 
                "Empenhos Vinculados", "Finalidade"
            ]
            df_exib_nc = df_rel_nc[cols_ordem].copy()

            # Formata valores em R$
            df_exib_formatado = df_exib_nc.copy()
            df_exib_formatado['Valor Total (R$)'] = df_exib_formatado['Valor Total (R$)'].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            df_exib_formatado['Total Empenhado (R$)'] = df_exib_formatado['Total Empenhado (R$)'].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            df_exib_formatado['Saldo Disponível (R$)'] = df_exib_formatado['Saldo Disponível (R$)'].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            st.dataframe(df_exib_formatado, use_container_width=True, hide_index=True)
            
            col_csv_nc, col_xlsx_nc, col_pdf_nc = st.columns(3)
            with col_csv_nc:
                st.download_button("📥 Baixar CSV", data=df_exib_nc.to_csv(index=False).encode('utf-8'), file_name="relatorio_notas_credito.csv", mime="text/csv", use_container_width=True)
            
            with col_xlsx_nc:
                buf_excel_nc = exportar_excel_seguro(df_exib_nc, 'Notas de Crédito')
                if buf_excel_nc:
                    st.download_button("📊 Baixar Planilha (.xlsx / Calc)", data=buf_excel_nc, file_name="relatorio_notas_credito.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            
            with col_pdf_nc:
                buf_pdf_nc = gerar_pdf_tabela_seguro("RELATÓRIO GERAL DE NOTAS DE CRÉDITO E EMPENHOS VINCULADOS", df_exib_formatado)
                if buf_pdf_nc:
                    st.download_button("📄 Baixar em PDF (.pdf)", data=buf_pdf_nc, file_name="relatorio_notas_credito.pdf", mime="application/pdf", use_container_width=True)
        else:
            st.info("Nenhuma Nota de Crédito encontrada com os filtros selecionados.")

    elif "Notas de Empenho" in tipo_relatorio:
        st.markdown("##### 🔍 Filtros de Consulta das Notas de Empenho")
        fe1, fe2, fe3 = st.columns(3)
        df_nes_base = pd.read_sql_query(f"SELECT DISTINCT status, tipo_empenho FROM notas_empenho WHERE om_id = {user['om_id']}", conn)
        
        status_opts = ["Todos"] + sorted(list(set(df_nes_base['status'].dropna().tolist())))
        f_status = fe1.selectbox("Status de Execução:", status_opts, key="f_st_ne_rel")
        
        tipo_emp_opts = ["Todos"] + sorted(list(set(df_nes_base['tipo_empenho'].dropna().tolist())))
        f_tipo_emp = fe2.selectbox("Tipo de Empenho:", tipo_emp_opts, key="f_tipo_ne_rel")
        busca_forn = fe3.text_input("Filtrar por Fornecedor (Nome ou CNPJ):", key="f_forn_ne_rel")

        q_rel_ne = f"""
        SELECT ne.numero_ne as "Número NE",
               nc.numero_nc as "NC Origem",
               ne.tipo_empenho as "Tipo",
               ne.fornecedor_nome as "Fornecedor",
               ne.fornecedor_cnpj as "CNPJ",
               ne.valor_ne as "Valor NE (R$)",
               ne.data_emissao,
               COALESCE(ne.nova_data_limite, ne.data_limite) as data_limite,
               ne.status as "Status"
        FROM notas_empenho ne
        JOIN notas_credito nc ON ne.nc_id = nc.id
        WHERE ne.om_id = {user['om_id']}
        """
        if f_status != "Todos":
            q_rel_ne += f" AND ne.status = '{f_status}'"
        if f_tipo_emp != "Todos":
            q_rel_ne += f" AND ne.tipo_empenho = '{f_tipo_emp}'"
        if busca_forn:
            q_rel_ne += f" AND (ne.fornecedor_nome LIKE '%{busca_forn}%' OR ne.fornecedor_cnpj LIKE '%{busca_forn}%')"

        q_rel_ne += " ORDER BY ne.data_emissao DESC"

        df_rel_ne = pd.read_sql_query(q_rel_ne, conn)
        if not df_rel_ne.empty:
            df_rel_ne['Data Emissão'] = df_rel_ne['data_emissao'].apply(formatar_data_br)
            df_rel_ne['Data Limite Entrega'] = df_rel_ne['data_limite'].apply(formatar_data_br)
            
            cols_ne_ordem = ["Número NE", "NC Origem", "Tipo", "Fornecedor", "CNPJ", "Valor NE (R$)", "Data Emissão", "Data Limite Entrega", "Status"]
            df_exib_ne = df_rel_ne[cols_ne_ordem].copy()

            df_exib_ne_fmt = df_exib_ne.copy()
            df_exib_ne_fmt['Valor NE (R$)'] = df_exib_ne_fmt['Valor NE (R$)'].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            st.dataframe(df_exib_ne_fmt, use_container_width=True, hide_index=True)
            
            col_csv_ne, col_xlsx_ne, col_pdf_ne = st.columns(3)
            with col_csv_ne:
                st.download_button("📥 Baixar CSV", data=df_exib_ne.to_csv(index=False).encode('utf-8'), file_name="relatorio_notas_empenho.csv", mime="text/csv", use_container_width=True)
            with col_xlsx_ne:
                buf_excel_ne = exportar_excel_seguro(df_exib_ne, 'Notas de Empenho')
                if buf_excel_ne:
                    st.download_button("📊 Baixar Planilha (.xlsx / Calc)", data=buf_excel_ne, file_name="relatorio_notas_empenho.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with col_pdf_ne:
                buf_pdf_ne = gerar_pdf_tabela_seguro("RELATÓRIO GERAL DE NOTAS DE EMPENHO", df_exib_ne_fmt)
                if buf_pdf_ne:
                    st.download_button("📄 Baixar em PDF (.pdf)", data=buf_pdf_ne, file_name="relatorio_notas_empenho.pdf", mime="application/pdf", use_container_width=True)
        else:
            st.info("Nenhum Empenho encontrado com os filtros selecionados.")

    elif "Notas Fiscais" in tipo_relatorio:
        q_rel_nf = f"""
        SELECT nf.numero_nf as "Número NF",
               ne.numero_ne as "Empenho (NE)",
               nc.numero_nc as "NC Origem",
               nf.empresa_cnpj as "CNPJ Fornecedor",
               nf.data_entrada_almox,
               nf.valor_nf as "Valor NF (R$)",
               nf.tipo_liquidacao as "Tipo Liquidação",
               nf.situacao as "Situação"
        FROM notas_fiscais nf
        JOIN notas_empenho ne ON nf.ne_id = ne.id
        JOIN notas_credito nc ON ne.nc_id = nc.id
        WHERE nf.om_id = {user['om_id']}
        ORDER BY nf.id DESC
        """
        df_rel_nf = pd.read_sql_query(q_rel_nf, conn)
        if not df_rel_nf.empty:
            df_rel_nf['Data Entrada Almox'] = df_rel_nf['data_entrada_almox'].apply(formatar_data_br)
            cols_nf = ["Número NF", "Empenho (NE)", "NC Origem", "CNPJ Fornecedor", "Data Entrada Almox", "Valor NF (R$)", "Tipo Liquidação", "Situação"]
            df_exib_nf = df_rel_nf[cols_nf].copy()

            df_exib_nf_fmt = df_exib_nf.copy()
            df_exib_nf_fmt['Valor NF (R$)'] = df_exib_nf_fmt['Valor NF (R$)'].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            st.dataframe(df_exib_nf_fmt, use_container_width=True, hide_index=True)
            
            col_csv_nf, col_xlsx_nf, col_pdf_nf = st.columns(3)
            with col_csv_nf:
                st.download_button("📥 Baixar CSV", data=df_exib_nf.to_csv(index=False).encode('utf-8'), file_name="relatorio_notas_fiscais.csv", mime="text/csv", use_container_width=True)
            with col_xlsx_nf:
                buf_excel_nf = exportar_excel_seguro(df_exib_nf, 'Notas Fiscais')
                if buf_excel_nf:
                    st.download_button("📊 Baixar Planilha (.xlsx / Calc)", data=buf_excel_nf, file_name="relatorio_notas_fiscais.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with col_pdf_nf:
                buf_pdf_nf = gerar_pdf_tabela_seguro("RELATÓRIO GERAL DE NOTAS FISCAIS RECEBIDAS", df_exib_nf_fmt)
                if buf_pdf_nf:
                    st.download_button("📄 Baixar em PDF (.pdf)", data=buf_pdf_nf, file_name="relatorio_notas_fiscais.pdf", mime="application/pdf", use_container_width=True)
        else:
            st.info("Nenhuma Nota Fiscal encontrada.")

    conn.close()

# ==========================================
# 8. ADMIN (OMS, IDENTIDADE VISUAL E USUÁRIOS)
# ==========================================
elif menu_selecionado == "⚙️ Admin":
    st.title("⚙️ Painel de Administração - Forte Pinheirinho")
    conn = get_connection()

    col_om, col_user = st.columns(2)
    
    with col_om:
        st.subheader("🏢 Gestão de Organizações Militares")
        tab_om_cad, tab_om_edit, tab_om_theme, tab_om_del = st.tabs(["➕ Nova OM", "✏️ Editar OM", "🎨 Identidade Visual", "🗑️ Excluir OM"])
        
        with tab_om_cad:
            with st.form("form_om"):
                sigla = st.text_input("Sigla da OM (ex.: 5ª Cia PE)")
                nome = st.text_input("Nome da OM")
                ug = st.text_input("Código de UG")
                if st.form_submit_button("Cadastrar OM"):
                    c = conn.cursor()
                    try:
                        c.execute("INSERT INTO oms (sigla, nome, ug) VALUES (?, ?, ?)", (sigla, nome, ug))
                        conn.commit()
                        st.success(f"OM {sigla} cadastrada!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("OM já cadastrada.")
        
        with tab_om_edit:
            df_oms = pd.read_sql_query("SELECT * FROM oms", conn)
            if not df_oms.empty:
                dict_om_ed = {f"{row['sigla']} - {row['nome']} (UG: {row['ug']})": row['id'] for _, row in df_oms.iterrows()}
                om_sel_edit = st.selectbox("Selecione a OM para editar:", list(dict_om_ed.keys()), key="sel_om_edit")
                id_om_edit = dict_om_ed[om_sel_edit]

                c = conn.cursor()
                c.execute("SELECT * FROM oms WHERE id = ?", (id_om_edit,))
                dados_om = dict(c.fetchone())

                with st.form("form_edit_om"):
                    nova_sigla = st.text_input("Sigla da OM:", value=dados_om['sigla'], key="ed_sigla")
                    novo_nome_om = st.text_input("Nome da OM:", value=dados_om['nome'], key="ed_nome_om")
                    nova_ug = st.text_input("Código de UG:", value=dados_om['ug'] if dados_om['ug'] else "", key="ed_ug")
                    
                    if st.form_submit_button("Salvar Alterações da OM"):
                        try:
                            c.execute("UPDATE oms SET sigla = ?, nome = ?, ug = ? WHERE id = ?",
                                      (nova_sigla, novo_nome_om, nova_ug, id_om_edit))
                            conn.commit()
                            st.success(f"OM '{nova_sigla}' atualizada com sucesso!")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("Já existe outra OM cadastrada com essa sigla.")
            else:
                st.info("Nenhuma OM cadastrada.")

        with tab_om_theme:
            df_oms = pd.read_sql_query("SELECT * FROM oms", conn)
            if not df_oms.empty:
                dict_om_th = {f"{row['sigla']} - {row['nome']}": row['id'] for _, row in df_oms.iterrows()}
                om_th_sel = st.selectbox("Selecione a OM para personalizar:", list(dict_om_th.keys()), key="sel_om_th")
                id_om_th = dict_om_th[om_th_sel]

                c = conn.cursor()
                c.execute("SELECT * FROM oms WHERE id = ?", (id_om_th,))
                dados_om_th = dict(c.fetchone())

                with st.form("form_theme_om"):
                    st.markdown(f"#### Personalização Visual: **{dados_om_th['sigla']}**")
                    arquivo_logo = st.file_uploader("Upload do Brasão / Distintivo da OM (PNG ou JPG):", type=["png", "jpg", "jpeg"])
                    
                    col_c1, col_c2 = st.columns(2)
                    cor_atual = dados_om_th.get('cor_primaria') or '#1B4332'
                    nova_cor = col_c1.color_picker("Cor Primária da OM:", value=cor_atual)
                    col_c2.info("Dica: Verde-Oliva (#1B4332), Azul-Marinho PE (#0A2240), Cinza Blindado (#2C3E50), Vermelho (#800000)")
                    
                    lema_atual = dados_om_th.get('lema') or ''
                    novo_lema = st.text_input("Lema ou Frase da OM (opcional):", value=lema_atual)

                    if st.form_submit_button("Salvar Identidade Visual"):
                        c = conn.cursor()
                        if arquivo_logo is not None:
                            logo_b64_new = base64.b64encode(arquivo_logo.read()).decode('utf-8')
                            c.execute("UPDATE oms SET cor_primaria = ?, lema = ?, logo_base64 = ? WHERE id = ?",
                                      (nova_cor, novo_lema, logo_b64_new, id_om_th))
                        else:
                            c.execute("UPDATE oms SET cor_primaria = ?, lema = ? WHERE id = ?",
                                      (nova_cor, novo_lema, id_om_th))
                        conn.commit()
                        st.success(f"Identidade visual da OM {dados_om_th['sigla']} salva com sucesso!")
                        st.rerun()
            else:
                st.info("Nenhuma OM cadastrada.")

        with tab_om_del:
            df_oms = pd.read_sql_query("SELECT * FROM oms", conn)
            if not df_oms.empty:
                om_para_excluir = st.selectbox("Selecione a OM para excluir:", df_oms['sigla'].tolist(), key="del_om_select")
                id_om_del = int(df_oms[df_oms['sigla'] == om_para_excluir]['id'].values[0])
                
                c = conn.cursor()
                c.execute("SELECT COUNT(*) FROM usuarios WHERE om_id = ?", (id_om_del,))
                qtd_u = c.fetchone()[0]
                c.execute("SELECT COUNT(*) FROM notas_credito WHERE om_id = ?", (id_om_del,))
                qtd_nc = c.fetchone()[0]

                if qtd_u > 0 or qtd_nc > 0:
                    st.warning(f"⚠️ Esta OM possui {qtd_u} usuário(s) e {qtd_nc} NC(s) vinculadas. Remova os vínculos antes de excluir.")
                else:
                    if st.button(f"Confirmar Exclusão de {om_para_excluir}", type="primary", key="btn_del_om"):
                        c.execute("DELETE FROM oms WHERE id = ?", (id_om_del,))
                        conn.commit()
                        st.success(f"OM '{om_para_excluir}' excluída com sucesso!")
                        st.rerun()
            else:
                st.info("Nenhuma OM cadastrada.")

        st.markdown("---")
        st.write("**OMs Cadastradas no Sistema:**")
        st.dataframe(pd.read_sql_query("SELECT id, sigla, nome, ug, cor_primaria, lema FROM oms", conn), use_container_width=True, hide_index=True)

    with col_user:
        st.subheader("👤 Gestão de Usuários e Senhas")
        df_oms_user = pd.read_sql_query("SELECT id, sigla FROM oms", conn)
        dict_oms = {row['sigla']: row['id'] for _, row in df_oms_user.iterrows()}
        dict_oms_rev = {row['id']: row['sigla'] for _, row in df_oms_user.iterrows()}
        
        df_usuarios = pd.read_sql_query('''
        SELECT u.id, o.sigla as OM, u.posto_grad, u.nome_guerra, u.perfil, u.identidade_militar, u.om_id 
        FROM usuarios u JOIN oms o ON u.om_id = o.id
        ''', conn)

        tab_u_cad, tab_u_edit, tab_u_del = st.tabs(["➕ Novo Usuário", "✏️ Editar Tudo / Senha", "🗑️ Excluir Usuário"])
        perfis_disponiveis_todos = ["Administrador", "Cmt OM", "Ch 4ª Seção", "SALC", "Almoxarife", "Op Almoxarifado"]

        with tab_u_cad:
            if not df_oms_user.empty:
                with st.form("form_u"):
                    om_u = st.selectbox("OM:", list(dict_oms.keys()), key="cad_user_om")
                    col_g, col_n = st.columns(2)
                    p_grad = col_g.selectbox("Posto/Graduação", ["Cel", "Ten Cel", "Maj", "Cap", "1º Ten", "2º Ten", "Asp", "Subten", "1º Sgt", "2º Sgt", "3º Sgt", "Cb", "Sd", "SC"], key="cad_user_p")
                    n_guerra = col_n.text_input("Nome de Guerra", key="cad_user_ng")
                    ident = st.text_input("Identidade Militar (Login)", key="cad_user_id")
                    
                    perfis_escolhidos = st.multiselect(
                        "Perfis de Acesso Autorizados:",
                        perfis_disponiveis_todos,
                        default=["SALC"],
                        key="cad_user_perf_multi"
                    )
                    
                    sen = st.text_input("Senha Inicial", type="password", value="1234", key="cad_user_sen")
                    if st.form_submit_button("Cadastrar Militar"):
                        if not perfis_escolhidos:
                            st.error("Selecione pelo menos um perfil de acesso para o militar.")
                        else:
                            c = conn.cursor()
                            perf_str = ", ".join(perfis_escolhidos)
                            try:
                                c.execute("INSERT INTO usuarios (nome_guerra, posto_grad, identidade_militar, om_id, perfil, senha) VALUES (?, ?, ?, ?, ?, ?)",
                                          (n_guerra, p_grad, ident, dict_oms[om_u], perf_str, sen))
                                conn.commit()
                                st.success(f"Usuário {n_guerra} criado com perfis: {perf_str}!")
                                st.rerun()
                            except sqlite3.IntegrityError:
                                st.error("Identidade Militar já cadastrada.")
            else:
                st.warning("Cadastre uma OM antes de criar usuários.")

        with tab_u_edit:
            if not df_usuarios.empty:
                opcoes_edit = {
                    f"{row['posto_grad']} {row['nome_guerra']} - OM: {row['OM']} (ID: {row['identidade_militar']})": row['id']
                    for _, row in df_usuarios.iterrows()
                }
                user_selecionado = st.selectbox("Selecione o militar para editar:", list(opcoes_edit.keys()), key="edit_user_sel")
                user_id_edit = opcoes_edit[user_selecionado]

                c = conn.cursor()
                c.execute("SELECT * FROM usuarios WHERE id = ?", (user_id_edit,))
                dados_u = dict(c.fetchone())

                postos_lista = ["Cel", "Ten Cel", "Maj", "Cap", "1º Ten", "2º Ten", "Asp", "Subten", "1º Sgt", "2º Sgt", "3º Sgt", "Cb", "Sd", "SC"]
                lista_siglas_om = list(dict_oms.keys())
                perfis_atuais_u = [p.strip() for p in dados_u['perfil'].split(",") if p.strip() and p.strip() in perfis_disponiveis_todos]

                with st.form("form_edit_user_tab"):
                    sigla_atual_om = dict_oms_rev.get(dados_u['om_id'], lista_siglas_om[0])
                    idx_om = lista_siglas_om.index(sigla_atual_om) if sigla_atual_om in lista_siglas_om else 0
                    nova_om_u = st.selectbox("OM de Lotação:", lista_siglas_om, index=idx_om, key="ed_om_u")

                    col_e1, col_e2 = st.columns(2)
                    posto_idx = postos_lista.index(dados_u['posto_grad']) if dados_u['posto_grad'] in postos_lista else 0
                    novo_posto = col_e1.selectbox("Posto/Graduação:", postos_lista, index=posto_idx, key="ed_p")
                    novo_nome = col_e2.text_input("Nome de Guerra:", value=dados_u['nome_guerra'], key="ed_ng")

                    col_e3, col_e4 = st.columns(2)
                    nova_ident = col_e3.text_input("Identidade Militar (Login):", value=dados_u['identidade_militar'], key="ed_id_m")
                    
                    novos_perfis_sel = st.multiselect(
                        "Perfis de Acesso Autorizados:",
                        perfis_disponiveis_todos,
                        default=perfis_atuais_u,
                        key="ed_perf_multi"
                    )

                    st.markdown("🔑 **Redefinição de Senha:**")
                    nova_senha = st.text_input("Senha (altere ou mantenha a atual):", value=dados_u['senha'], key="ed_sen")

                    st.markdown("---")
                    confirmar_alteracao = st.checkbox("⚠️ Confirmo que conferi os dados e desejo salvar estas alterações.", key="chk_conf_u")

                    if st.form_submit_button("Salvar Todas as Alterações do Usuário"):
                        if not confirmar_alteracao:
                            st.warning("Por favor, marque a caixa de confirmação acima para autorizar as alterações.")
                        elif not novos_perfis_sel:
                            st.error("O militar deve possuir pelo menos um perfil de acesso selecionado.")
                        else:
                            try:
                                nova_perf_str = ", ".join(novos_perfis_sel)
                                c.execute('''
                                UPDATE usuarios 
                                SET posto_grad = ?, nome_guerra = ?, identidade_militar = ?, om_id = ?, perfil = ?, senha = ?
                                WHERE id = ?
                                ''', (novo_posto, novo_nome, nova_ident, dict_oms[nova_om_u], nova_perf_str, nova_senha, user_id_edit))
                                conn.commit()
                                
                                if user_id_edit == user['id']:
                                    c.execute('''
                                    SELECT u.*, o.sigla as om_sigla, o.logo_base64, o.cor_primaria, o.lema FROM usuarios u 
                                    JOIN oms o ON u.om_id = o.id WHERE u.id = ?
                                    ''', (user_id_edit,))
                                    st.session_state['usuario'] = dict(c.fetchone())
                                    if st.session_state.get('perfil_ativo') not in novos_perfis_sel:
                                        st.session_state['perfil_ativo'] = novos_perfis_sel[0]

                                st.success(f"✅ Confirmação: Dados e perfis de {novo_posto} {novo_nome} atualizados com sucesso!")
                                st.rerun()
                            except sqlite3.IntegrityError:
                                st.error("Erro: Já existe outro usuário cadastrado com essa Identidade Militar.")
            else:
                st.info("Nenhum usuário cadastrado.")

        with tab_u_del:
            if not df_usuarios.empty:
                opcoes_usuarios = {
                    f"{row['posto_grad']} {row['nome_guerra']} - OM: {row['OM']} (ID: {row['identidade_militar']})": row['id']
                    for _, row in df_usuarios.iterrows()
                }
                user_para_excluir = st.selectbox("Selecione o usuário para excluir:", list(opcoes_usuarios.keys()), key="del_u_sel")
                id_user_del = opcoes_usuarios[user_para_excluir]

                if id_user_del == user['id']:
                    st.error("⛔ Você não pode excluir o seu próprio usuário enquanto estiver conectado ao sistema.")
                else:
                    if st.button("Confirmar Exclusão de Usuário", type="primary", key="btn_del_u"):
                        c = conn.cursor()
                        c.execute("DELETE FROM usuarios WHERE id = ?", (id_user_del,))
                        conn.commit()
                        st.success("Usuário excluído com sucesso!")
                        st.rerun()
            else:
                st.info("Nenhum usuário cadastrado.")

        st.markdown("---")
        st.write("**Militares Cadastrados no Sistema:**")
        st.dataframe(df_usuarios[['OM', 'posto_grad', 'nome_guerra', 'perfil', 'identidade_militar']], use_container_width=True, hide_index=True)

    conn.close()

