import os
import comtypes.client
import PyPDF2
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configurações de upload
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')  # Caminho absoluto
ALLOWED_EXTENSIONS = {'docx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Verifica se o arquivo é permitido
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Função para converter arquivos Word para PDF
def converter_word_para_pdf():
    # Inicializa o COM
    comtypes.CoInitialize()

    arquivos_pdf = []
    pdf_folder = os.path.join(os.getcwd(), 'pdfs')  # Caminho absoluto para PDFs

    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)

    try:
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False  # Não exibe o Word na interface
        word.DisplayAlerts = 0  # Desabilita alertas

        # Depuração: Verifica o conteúdo do diretório de uploads
        print(f"Conteúdo do diretório de uploads: {os.listdir(UPLOAD_FOLDER)}")

        # Lê todos os arquivos enviados
        for arquivo in os.listdir(UPLOAD_FOLDER):
            if arquivo.endswith(".docx"):
                # Caminho completo para o arquivo
                arquivo_word = os.path.join(UPLOAD_FOLDER, arquivo)
                print(f"Arquivo encontrado para conversão: {arquivo_word}")  # Depuração

                # Tenta garantir que o nome do arquivo não tenha caracteres especiais
                arquivo_word = os.path.normpath(arquivo_word)

                arquivo_pdf = os.path.join(pdf_folder, arquivo.replace(".docx", ".pdf"))
                arquivos_pdf.append(arquivo_pdf)

                try:
                    # Tenta abrir o documento Word
                    document = word.Documents.Open(arquivo_word)
                    print(f"Documento aberto com sucesso: {arquivo_word}")

                    # Converte o Word para PDF
                    document.SaveAs(arquivo_pdf, FileFormat=17)  # 17 é o formato PDF
                    print(f"Documento convertido para PDF: {arquivo_pdf}")
                    document.Close()

                except Exception as e:
                    print(f"Erro ao processar o documento {arquivo_word}: {e}")

        word.Quit()  # Fecha o Word após terminar

        # Mesclar os arquivos PDF em um único arquivo
        if arquivos_pdf:
            arquivo_pdf_final = os.path.join(pdf_folder, "documento_final.pdf")
            pdf_writer = PyPDF2.PdfWriter()

            for pdf in arquivos_pdf:
                with open(pdf, 'rb') as f:
                    pdf_reader = PyPDF2.PdfReader(f)
                    for page in range(len(pdf_reader.pages)):
                        pdf_writer.add_page(pdf_reader.pages[page])

            # Salva o PDF final
            with open(arquivo_pdf_final, 'wb') as f:
                pdf_writer.write(f)

            # Apaga os PDFs temporários
            for pdf in arquivos_pdf:
                os.remove(pdf)

            return arquivo_pdf_final

        return []

    except Exception as e:
        print(f"Erro ao inicializar o Word: {e}")
        return []

    finally:
        # Desfaz a inicialização do COM após o processo
        comtypes.CoUninitialize()

# Página inicial com o formulário de upload
@app.route('/')
def index():
    return render_template('index.html')

# Rota para upload de arquivos
@app.route('/upload', methods=['POST'])
def upload():
    if 'files' not in request.files:
        return "No file part", 400

    files = request.files.getlist('files')
    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            # Depuração: Verifica se o arquivo foi salvo corretamente
            print(f"Arquivo salvo: {file_path}")

    # Converte os arquivos Word para PDF
    arquivo_pdf_final = converter_word_para_pdf()

    # Se houver um PDF final gerado, retorna o arquivo para o usuário
    if arquivo_pdf_final:
        return send_file(arquivo_pdf_final, as_attachment=True)

    return "Erro na conversão", 500

if __name__ == '__main__':
    app.run(debug=True)
