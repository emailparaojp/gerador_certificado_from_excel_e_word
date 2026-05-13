import os
from docx import Document
import pandas as pd
from docx2pdf import convert
from PyPDF2 import PdfReader, PdfWriter
import zipfile
import tempfile
import shutil

# Caminhos
template_word_path = 'docs_modelo/doc_word.docx'
excel_file = 'docs_modelo/xls_para_certificados.xlsx'
pdf_verso = 'docs_modelo/pdf_verso.pdf'
output_word_dir = 'Certificados_word'  # Pasta para os arquivos Word
output_pdf_dir = 'Certificados_pdf'    # Pasta para os PDFs

# Criar diretórios principais de saída
if not os.path.exists(output_word_dir):
    os.makedirs(output_word_dir)

if not os.path.exists(output_pdf_dir):
    os.makedirs(output_pdf_dir)

# Substituir texto em runs (robusto para placeholder dividido em vários runs)
def replace_text_in_runs(paragraph, placeholder, replacement):
    # Tenta substituir diretamente em cada run (caso comum)
    for run in paragraph.runs:
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, replacement)
            return
    # Se não encontrou no mesmo run, o placeholder pode estar dividido entre runs.
    # Nesse caso, faz fallback substituindo no texto do parágrafo inteiro.
    if placeholder in paragraph.text:
        paragraph.text = paragraph.text.replace(placeholder, replacement)

# Função para substituir placeholder diretamente no arquivo .docx (document.xml)
def replace_placeholder_in_docx_file(docx_path, replacements):
    """Substitui placeholders em todos os arquivos XML dentro de word/ no pacote .docx.
    replacements: dict mapping placeholder -> replacement string
    """
    try:
        with zipfile.ZipFile(docx_path, 'r') as zin:
            names = zin.namelist()
            file_data = {name: zin.read(name) for name in names}

        any_changed = False
        for name in list(file_data.keys()):
            if name.startswith('word/') and name.endswith('.xml'):
                try:
                    xml = file_data[name].decode('utf-8')
                except Exception:
                    continue
                # Normalizar NBSP para evitar diferenças invisíveis
                xml_norm = xml.replace('\u00A0', ' ')
                file_changed = False
                for ph, rep in replacements.items():
                    if ph in xml_norm:
                        xml_norm = xml_norm.replace(ph, rep)
                        file_changed = True
                if file_changed:
                    file_data[name] = xml_norm.encode('utf-8')
                    any_changed = True

        if any_changed:
            dirpath = os.path.dirname(docx_path) or '.'
            fd, temp_path = tempfile.mkstemp(suffix='.docx', dir=dirpath)
            os.close(fd)
            with zipfile.ZipFile(temp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for name, data in file_data.items():
                    zout.writestr(name, data)
            shutil.move(temp_path, docx_path)
    except Exception as e:
        print(f"Falha ao substituir placeholder no DOCX '{docx_path}': {e}")

# Função para gerar certificados Word e PDFs
def generate_certificates(template_path, xls_data, output_word_dir, output_pdf_dir):
    for sheet_name, sheet_data in xls_data.items():
        # Criar subpasta para os PDFs da planilha atual
        sheet_pdf_dir = os.path.join(output_pdf_dir, sheet_name)
        sheet_pdf_dir_word = os.path.join(output_word_dir, sheet_name)
        if not os.path.exists(sheet_pdf_dir):
            os.makedirs(sheet_pdf_dir)

        if not os.path.exists(sheet_pdf_dir_word):
            os.makedirs(sheet_pdf_dir_word)

        # Filtrar os nomes da planilha atual
        if 'NOME' in sheet_data.columns:
            names = sheet_data['NOME'].dropna().tolist()
            
            for name in names:
                # Carregar o modelo Word
                doc = Document(template_path)
                for paragraph in doc.paragraphs:
                    replace_text_in_runs(paragraph, "NNmunicipioNN", sheet_name)
                    replace_text_in_runs(paragraph, "NNnomeNN", name.upper())
                
                # Salvar o arquivo Word gerado
                sanitized_name = name.replace(' ', '_').replace('/', '_')
                word_file_path = os.path.join(sheet_pdf_dir_word, f"certificado_{sanitized_name}.docx")
                doc.save(word_file_path)

                # Garantir substituição também no XML do .docx (cobre textboxes/caixas de texto)
                # Substituir ambos os placeholders: nome (em maiúsculas para manter consistência) e município
                replace_placeholder_in_docx_file(word_file_path, {
                    "NNnomeNN": name.upper(),
                    "NNmunicipioNN": sheet_name
                })

                # Converter para PDF na subpasta da planilha
                pdf_file_path = os.path.join(sheet_pdf_dir, f"certificado_{sanitized_name}.pdf")
                convert(word_file_path, pdf_file_path)
                pdf_verso = False
                # Se existir pdf_verso, anexar suas páginas ao PDF recém-gerado (não alterar o .docx)
                if pdf_verso and os.path.exists(pdf_verso):
                    try:
                        reader_main = PdfReader(pdf_file_path)
                        reader_verso = PdfReader(pdf_verso)
                        writer = PdfWriter()
                        for page in reader_main.pages:
                            writer.add_page(page)
                        for page in reader_verso.pages:
                            writer.add_page(page)
                        with open(pdf_file_path, 'wb') as f:
                            writer.write(f)
                    except Exception as e:
                        print(f"Falha ao anexar '{pdf_verso}' em '{pdf_file_path}': {e}")

# Ler planilhas do Excel
xls_data = pd.read_excel(excel_file, sheet_name=None)

# Gerar certificados
generate_certificates(template_word_path, xls_data, output_word_dir, output_pdf_dir)

print(f"Certificados Word gerados em: {output_word_dir}")
print(f"Certificados PDF organizados em: {output_pdf_dir}")
