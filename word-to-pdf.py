import os
import comtypes.client
import PyPDF2
from tqdm import tqdm
import time

def converter_pasta(origem, destino):
    """Converte todos os arquivos .docx da pasta origem para .pdf na pasta destino."""
    if not os.path.exists(destino):
        os.makedirs(destino)

    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False  

    arquivos_convertidos = []  # Lista para armazenar os PDFs gerados
    arquivos_word = []  # Lista para armazenar os arquivos Word a serem apagados

    for arquivo in os.listdir(origem):
        if arquivo.endswith(".docx"):  
            caminho_docx = os.path.join(origem, arquivo)
            caminho_pdf = os.path.join(destino, arquivo.replace(".docx", ".pdf"))

            try:
                doc = word.Documents.Open(os.path.abspath(caminho_docx))
                doc.SaveAs(os.path.abspath(caminho_pdf), FileFormat=17)
                doc.Close()
                arquivos_convertidos.append(caminho_pdf)  # Adiciona à lista de PDFs
                arquivos_word.append(caminho_docx)  # Adiciona à lista de arquivos Word
                print(f"✔ Convertido: {arquivo} -> {caminho_pdf}")
            except Exception as e:
                print(f"❌ Erro ao converter {arquivo}: {e}")

    word.Quit()
    return arquivos_convertidos, arquivos_word  # Retorna PDFs e arquivos Word

def juntar_pdfs(arquivos, saida):
    """Junta múltiplos arquivos PDF em um único arquivo."""
    mesclador = PyPDF2.PdfMerger()

    print("\n📌 Iniciando a junção dos PDFs...\n")

    for arquivo in tqdm(arquivos, desc="Processando", unit="arquivo"):
        try:
            mesclador.append(arquivo)
            time.sleep(0.2)  
        except Exception as e:
            print(f"❌ Erro ao adicionar {arquivo}: {e}")

    mesclador.write(saida)
    mesclador.close()
    print(f"\n✅ PDFs mesclados com sucesso em '{saida}'!")

def apagar_arquivos(lista_arquivos, tipo=""):
    """Apaga os arquivos da lista após a mesclagem."""
    print(f"\n🗑 Apagando arquivos {tipo}...\n")
    for arquivo in lista_arquivos:
        try:
            os.remove(arquivo)
            print(f"🗑 Arquivo apagado: {arquivo}")
        except Exception as e:
            print(f"❌ Erro ao apagar {arquivo}: {e}")

# 📂 Defina as pastas
pasta_origem = r"C:\Users\PC 02\Desktop\word-to-pdf\arquivos_word"
pasta_destino = r"C:\Users\PC 02\Desktop\word-to-pdf\arquivos_pdf"
arquivo_saida = r"C:\Users\PC 02\Desktop\word-to-pdf\final.pdf"

# 🔹 Etapa 1: Converter os arquivos Word para PDF
arquivos_pdf, arquivos_word = converter_pasta(pasta_origem, pasta_destino)

# 🔹 Etapa 2: Mesclar todos os PDFs gerados
if arquivos_pdf:
    juntar_pdfs(arquivos_pdf, arquivo_saida)

    # 🔹 Etapa 3: Apagar arquivos PDF individuais após a mesclagem
    apagar_arquivos(arquivos_pdf, "PDFs convertidos")

    # 🔹 Etapa 4: Apagar arquivos Word após a conversão
    apagar_arquivos(arquivos_word, "Word originais")
else:
    print("❌ Nenhum PDF foi gerado para mesclagem.")
