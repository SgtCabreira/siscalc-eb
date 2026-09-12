import io
import zipfile
import pandas as pd
from datetime import datetime

try:
    import docx
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    docx = None

def formatar_data_br(val):
    """Converte qualquer data para o padrão oficial brasileiro DD/MM/AAAA."""
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
    """Converte datas brasileiras ou ISO para o formato de banco YYYY-MM-DD."""
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

def exportar_excel_seguro(df_dados, nome_aba="Dados"):
    """
    Gera um buffer de planilha Excel (.xlsx) de forma segura.
    Se openpyxl não estiver instalado no servidor, retorna None sem quebrar a aplicação.
    """
    try:
        import openpyxl
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            df_dados.to_excel(writer, index=False, sheet_name=nome_aba[:31])
        buf.seek(0)
        return buf
    except Exception:
        return None

def gerar_pdf_tabela_seguro(titulo, df_dados):
    """
    Gera relatório em PDF oficial em formato A4 paisagem com cabeçalho da 5ª Cia PE.
    Compatível com FPDF e FPDF2 com proteção contra falhas de codificação Latin-1.
    """
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

def gerar_requisicao_odt(dados, total_requisicao, df_itens=None):
    """
    Gera o documento oficial de Requisição de Despesa da Base de Apoio da 5ª RM em formato nativo LibreOffice Writer (.odt).
    """
    buffer_odt = io.BytesIO()
    itens_xml = ""
    if df_itens is not None and not df_itens.empty:
        for idx, it in df_itens.iterrows():
            subtotal = float(it.get('QTD', 0)) * float(it.get('V_Unit', 0))
            itens_xml += f"<text:p>Item {it.get('Item', idx+1)}: {it.get('Descrição', '')} | UN: {it.get('UN', 'UN')} | QTD: {it.get('QTD', 0)} | Unit: R$ {it.get('V_Unit', 0):.2f} | Total: R$ {subtotal:.2f}</text:p>"

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
      <text:p>Requisição Nº {dados.get('numero_req', '')} | EB: {dados.get('nup', '')}</text:p>
      <text:p>Assunto: {dados.get('assunto', '')}</text:p>
      <text:p>{dados.get('data_extenso', '')}</text:p>
      <text:p/>
      <text:p>1. Solicito providências no sentido de aprovar a despesa abaixo especificada.</text:p>
      <text:p>2. DEMONSTRATIVO DE NECESSIDADES: {dados.get('pregao', '')}</text:p>
      <text:p>FORNECEDOR: {dados.get('razao_social', '')} (CNPJ: {dados.get('cnpj', '')})</text:p>
      <text:p>TIPO DE EMPENHO: {dados.get('tipo_empenho', 'Ordinário')}</text:p>
      {itens_xml}
      <text:p>VALOR TOTAL: R$ {total_requisicao:,.2f}</text:p>
      <text:p/>
      <text:p>3. LOCAL DE ENTREGA: {dados.get('local_entrega', '')}</text:p>
      <text:p>4. JUSTIFICATIVAS:</text:p>
      <text:p>a. Necessidade: {dados.get('justificativa_a', '')}</text:p>
      <text:p>b. Consequência da não emissão: {dados.get('justificativa_b', '')}</text:p>
      <text:p/>
      <text:p>5. CLASSIFICAÇÃO ORÇAMENTÁRIA: PI {dados.get('pi', '')} | PTRES {dados.get('ptres', '')} | ND {dados.get('nd', '')} | NC {dados.get('nc', '')}</text:p>
      <text:p/>
      <text:p>Aquisição aprovada por: {dados.get('fiscal_nome_posto', '')} - {dados.get('fiscal_funcao', '')}</text:p>
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
    return buffer_odt

def gerar_requisicao_docx(dados, total_requisicao, df_itens=None, modelo_bytes=None):
    """
    Gera o documento oficial de Requisição em formato Microsoft Word (.docx).
    Se houver um modelo carregado no banco, preserva o cabeçalho oficial e insere os dados.
    """
    if docx is None:
        return None

    if modelo_bytes:
        try:
            doc = docx.Document(io.BytesIO(modelo_bytes))
        except Exception:
            doc = docx.Document()
    else:
        doc = docx.Document()

    p_c = doc.add_paragraph()
    p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_c = p_c.add_run("MINISTÉRIO DA DEFESA\nEXÉRCITO BRASILEIRO\nBASE DE ADMINISTRAÇÃO E APOIO DA 5ª REGIÃO MILITAR\n(Companhia do QG da 5ª RM/DI)\nBASE MAJOR AGOSTINHO JOSÉ RODRIGUES\n")
    r_c.bold = True

    doc.add_paragraph(f"Requisição Nº {dados.get('numero_req', '')}\nEB: {dados.get('nup', '')}\nAssunto: {dados.get('assunto', '')}\n{dados.get('data_extenso', '')}\n")
    doc.add_paragraph(f"DEMONSTRATIVO: {dados.get('pregao', '')}\nFornecedor: {dados.get('razao_social', '')} ({dados.get('cnpj', '')})\nValor Total: R$ {total_requisicao:,.2f}")
    doc.add_paragraph(f"Local de Entrega: {dados.get('local_entrega', '')}\nJustificativa: {dados.get('justificativa_a', '')}\nClassificação: PI {dados.get('pi', '')} | ND {dados.get('nd', '')} | NC {dados.get('nc', '')}")

    p_a = doc.add_paragraph(f"\n\n_____________________________________________\n{dados.get('fiscal_nome_posto', '')}\n{dados.get('fiscal_funcao', '')}")
    p_a.alignment = WD_ALIGN_PARAGRAPH.CENTER

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

if __name__ == "__main__":
    print("Módulo core/documentos.py compilado com sucesso!")
