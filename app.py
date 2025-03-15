import os
import comtypes.client
import PyPDF2
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configurações de upload
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')  # Caminho absoluto
PDF_FOLDER = os.path.join(os.getcwd(), 'pdfs')  # Caminho absoluto para PDFs
ALLOWED_EXTENSIONS = {'docx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PDF_FOLDER'] = PDF_FOLDER

# Garante que as pastas existem
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PDF_FOLDER, exist_ok=True)

# Verifica se o arquivo é permitido
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Função para converter arquivos Word para PDF
def converter_word_para_pdf():
    comtypes.CoInitialize()
    arquivos_pdf = []

    try:
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        for arquivo in os.listdir(UPLOAD_FOLDER):
            if arquivo.endswith(".docx"):
                arquivo_word = os.path.join(UPLOAD_FOLDER, arquivo)
                arquivo_pdf = os.path.join(PDF_FOLDER, arquivo.replace(".docx", ".pdf"))
                arquivos_pdf.append(arquivo_pdf)

                try:
                    document = word.Documents.Open(arquivo_word)
                    document.SaveAs(arquivo_pdf, FileFormat=17)
                    document.Close()
                except Exception as e:
                    print(f"Erro ao processar o documento {arquivo_word}: {e}")

        word.Quit()

        if arquivos_pdf:
            arquivo_pdf_final = os.path.join(PDF_FOLDER, "documento_final.pdf")
            pdf_writer = PyPDF2.PdfWriter()
            for pdf in arquivos_pdf:
                with open(pdf, 'rb') as f:
                    pdf_reader = PyPDF2.PdfReader(f)
                    for page in range(len(pdf_reader.pages)):
                        pdf_writer.add_page(pdf_reader.pages[page])

            with open(arquivo_pdf_final, 'wb') as f:
                pdf_writer.write(f)

            for pdf in arquivos_pdf:
                os.remove(pdf)

            return arquivo_pdf_final

        return ""

    except Exception as e:
        print(f"Erro ao inicializar o Word: {e}")
        return ""

    finally:
        comtypes.CoUninitialize()

# Página inicial
@app.route('/')
def index():
    return render_template('index.html')

# Rota para upload de arquivos
@app.route('/upload', methods=['POST'])
def upload():
    if 'files' not in request.files:
        return "Nenhum arquivo enviado", 400

    files = request.files.getlist('files')
    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

    arquivo_pdf_final = converter_word_para_pdf()
    if arquivo_pdf_final:
        return send_file(arquivo_pdf_final, as_attachment=True)

    return "Erro na conversão", 500

# Rota para apagar arquivos
@app.route('/delete_all', methods=['POST'])
def delete_all():
    try:
        for pasta in [UPLOAD_FOLDER, PDF_FOLDER]:
            for arquivo in os.listdir(pasta):
                os.remove(os.path.join(pasta, arquivo))
        return jsonify({"message": "Todos os arquivos foram apagados com sucesso!"}), 200
    except Exception as e:
        return jsonify({"error": f"Erro ao apagar arquivos: {e}"}), 500

if __name__ == '__main__':
    app.run(debug=True)
