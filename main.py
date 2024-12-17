import os
from docx import Document
import pandas as pd
from docx2pdf import convert

# Caminhos
template_word_path = 'mod_certificado_word.docx'
excel_file = 'xls_para_certificados.xlsx'
output_word_dir = 'Certificados_word'  # Pasta para os arquivos Word
output_pdf_dir = 'Certificados_pdf'    # Pasta para os PDFs

# Criar diretórios principais de saída
if not os.path.exists(output_word_dir):
    os.makedirs(output_word_dir)

if not os.path.exists(output_pdf_dir):
    os.makedirs(output_pdf_dir)

# Substituir texto em runs
def replace_text_in_runs(paragraph, placeholder, replacement):
    for run in paragraph.runs:
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, replacement)

# Função para gerar certificados Word e PDFs
def generate_certificates(template_path, xls_data, output_word_dir, output_pdf_dir):
    for sheet_name, sheet_data in xls_data.items():
        # Criar subpasta para os PDFs da planilha atual
        sheet_pdf_dir = os.path.join(output_pdf_dir, sheet_name)
        if not os.path.exists(sheet_pdf_dir):
            os.makedirs(sheet_pdf_dir)
        
        # Filtrar os nomes da planilha atual
        if 'NOME' in sheet_data.columns:
            names = sheet_data['NOME'].dropna().tolist()
            
            for name in names:
                # Carregar o modelo Word
                doc = Document(template_path)
                for paragraph in doc.paragraphs:
                    replace_text_in_runs(paragraph, "NNnomeNN", name)
                
                # Salvar o arquivo Word gerado
                sanitized_name = name.replace(' ', '_').replace('/', '_')
                word_file_path = os.path.join(output_word_dir, f"certificado_{sanitized_name}.docx")
                doc.save(word_file_path)
                
                # Converter para PDF na subpasta da planilha
                pdf_file_path = os.path.join(sheet_pdf_dir, f"certificado_{sanitized_name}.pdf")
                convert(word_file_path, pdf_file_path)

# Ler planilhas do Excel
xls_data = pd.read_excel(excel_file, sheet_name=None)

# Gerar certificados
generate_certificates(template_word_path, xls_data, output_word_dir, output_pdf_dir)

print(f"Certificados Word gerados em: {output_word_dir}")
print(f"Certificados PDF organizados em: {output_pdf_dir}")
