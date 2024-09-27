import pyodbc
import pandas as pd
import sys

# importando as funções globais para reaproveitamento de código
sys.path.append("C:\\data-integration\\Automatizacao\\Python\\Yorton\\Python_automations_CC\\functions")
from data_base import db_connection  # função para conectar com banco de dados
from colors import green, red  # função para utilizar cores nos prints para melhor visualização

# Função para converter valores para float com precisão estendida
def convert_to_float(value):
    if isinstance(value, str):
        # Remove separadores de milhar e substitui a vírgula decimal por ponto
        value = value.replace('.', '').replace(',', '.')
    return float(value)

# Caminho para o arquivo Excel
excel_file = 'C:/temp/IMPORTAR BANCO RESULTADO_OPERADORES.xlsx'

# Lendo arquivo excel
try:
    data = pd.read_excel(excel_file)
except:
    print(red(f"Erro ao ler o arquivo Excel"))
    exit()

# Substitui valores NaN e strings vazias por None
data = data.applymap(lambda x: None if pd.isna(x) or x == '' else x)

# Convertendo os valores para o formato padrão da tabela do banco
for column in data.columns:
    if pd.api.types.is_numeric_dtype(data[column]):
        data[column] = data[column].apply(convert_to_float)

try:
    with db_connection() as connection:
        with connection.cursor() as cursor:
            table = 'DADOS_RESULTADOS_OPERACAO'
            
            # Obter a estrutura da tabela Oracle
            cursor.execute(f"SELECT COLUMN_NAME FROM ALL_TAB_COLUMNS WHERE TABLE_NAME = '{table}'")
            columns = [row[0] for row in cursor.fetchall()]
            num_columns = len(columns)
            
            print(f"Número de colunas na tabela Oracle: {num_columns}")
            
            # Adicionar colunas faltantes ao DataFrame com valores None
            # Na planilha tem umas colunas utilizadas apenas para separar algumas coisas...
            # ... tive que fazer essas linhas de código porque o banco não estava...
            # ... identificando a ultima coluna (a coluna não tem valores)
            for column in columns:
                if column not in data.columns:
                    data[column] = None

            # Reordenar as colunas para garantir que estão na mesma ordem que a tabela
            data = data[columns]
            
            # Inserir dados na tabela Oracle (sem formatação adicional)
            for index, row in data.iterrows():
                values = row.tolist()  # Converte a linha em lista de valores
                
                # Substituir strings vazias por None para garantir que sejam tratados como NULL
                values = [None if v == '' or pd.isna(v) else v for v in values]

                # Criar a query de inserção com reservas de espaço
                placeholders = ', '.join(['?' for _ in range(num_columns)])
                command = f"INSERT INTO {table} ({', '.join(columns)}) VALUES ({placeholders})"
                
                # print(f"Comando SQL: {command}")
                # print(f"Valores fornecidos para a linha {index + 1}: {values}")

                # inserindo os dados no banco
                try:
                    cursor.execute(command, values)
                except pyodbc.Error as e:
                    print(red(f"Erro ao inserir dados: {e}"))
                    exit()

            connection.commit()
            print(green("Dados importados com sucesso!"))

except pyodbc.Error as e:
    print(red(f"Erro ao conectar ou interagir com o banco de dados: {e}"))
