from dotenv import load_dotenv
import pyodbc
from functions.colors import red, green
import os

# carregar variáveis de ambiente
load_dotenv()

# Variáveis para acessar banco de dados
dsn = os.getenv('NOME_BANCO_DE_DADOS')
user = os.getenv('USUARIO_BANCO_DE_DADOS')
password = os.getenv('SENHA_BANCO_DE_DADOS')

# Função para conectar ao banco de dados
def db_connection():
    try:
        data_connection = f"DSN={dsn};UID={user};PWD={password}"
        connection = pyodbc.connect(data_connection)
        print(green("Banco de dados conectado com sucesso!"))
        return connection
    except:
        print(red("Erro ao se conectar com banco de dados!"))
        raise
