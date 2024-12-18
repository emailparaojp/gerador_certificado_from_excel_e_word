from flask import Flask, request, render_template, send_file, jsonify
import os
from werkzeug.utils import secure_filename
import pandas as pd
from docx import Document
from docx2pdf import convert
import shutil

app = Flask(__name__)

# Diretórios
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = os.path.join(UPLOAD_FOLDER, "output")
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# Rota principal para upload
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            # Receber arquivos
            excel_file = request.files["excel_file"]
            word_file = request.files["word_file"]

            # Salvar arquivos no servidor
            excel_path = os.path.join(UPLOAD_FOLDER, "excel", secure_filename(excel_file.filename))
            word_path = os.path.join(UPLOAD_FOLDER, "word", secure_filename(word_file.filename))

            os.makedirs(os.path.dirname(excel_path), exist_ok=True)
            os.makedirs(os.path.dirname(word_path), exist_ok=True)

            excel_file.save(excel_path)
            word_file.save(word_path)

            # Processar arquivos
            generate_certificates(word_path, excel_path, OUTPUT_FOLDER)

            # Compactar resultados
            zip_path = os.path.join("static", "downloads", "certificados.zip")
            os.makedirs(os.path.dirname(zip_path), exist_ok=True)
            shutil.make_archive(zip_path.replace(".zip", ""), "zip", OUTPUT_FOLDER)

            # Retornar link para download
            return jsonify({"download_link": f"/{zip_path}"})
        except Exception as e:
            # Retornar erro em caso de falha
            return jsonify({"error": str(e)}), 500
    return render_template("index.html")

# Função para gerar certificados
def generate_certificates(template_path, excel_path, output_dir):
    xls_data = pd.read_excel(excel_path, sheet_name=None)

    for sheet_name, sheet_data in xls_data.items():
        sheet_output_dir = os.path.join(output_dir, sheet_name)
        os.makedirs(sheet_output_dir, exist_ok=True)

        if "NOME" in sheet_data.columns:
            for name in sheet_data["NOME"].dropna():
                # Gerar Word
                sanitized_name = name.replace(" ", "_").replace("/", "_")
                word_file_path = os.path.join(sheet_output_dir, f"{sanitized_name}.docx")

                doc = Document(template_path)
                for paragraph in doc.paragraphs:
                    if "NNnomeNN" in paragraph.text:
                        paragraph.text = paragraph.text.replace("NNnomeNN", name)
                doc.save(word_file_path)

                # Converter para PDF
                pdf_file_path = os.path.join(sheet_output_dir, f"{sanitized_name}.pdf")
                convert(word_file_path, pdf_file_path)

# Rota para baixar o ZIP
@app.route("/static/downloads/<filename>", methods=["GET"])
def download_file(filename):
    return send_file(f"static/downloads/{filename}", as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
