# Usa uma imagem oficial do Python
FROM python:3.10

# Define o diretório de trabalho dentro do contêiner
WORKDIR /app

# Copia os arquivos do projeto para dentro do contêiner
COPY . .

# Instala as dependências do projeto
RUN pip install --no-cache-dir -r requirements.txt

# Expõe a porta 5000 para acessar o Flask
EXPOSE 5000

# Comando para rodar a aplicação Flask
CMD ["python", "app.py"]
