import pyodbc
import pandas as pd
from functions.data_base import db_connection # função para conectar com banco de dados
from functions.colors import green, red # função para utilizar cores nos prints

# Caminho para o arquivo Excel
excel_file = 'H:/Tecnologia/EQUIPE - DADOS/1 - Relatorios Recorrentes/Diario/1 - INSERT_DADOS_X5_PERFORMANCE_AGENTES.xlsx'

# Lendo arquivo excel
try:
    data = pd.read_excel(excel_file)
except Exception as e:
    print(red(f"Erro ao ler o arquivo Excel: {e}"))
    exit()

# Substitui valores Nan e strings vazias por None
data = data.applymap(lambda x: None if pd.isna(x) or x == '' else x)

try:
    with db_connection() as connection:  # conectando com banco de dados
        with connection.cursor() as cursor:  # abrindo central para rodar a query
            table = 'DADOS_X5_PERFORMANCE_AGENTES'
            
            # Obter a estrutura da tabela Oracle
            cursor.execute(f"SELECT COLUMN_NAME FROM ALL_TAB_COLUMNS WHERE TABLE_NAME = '{table}'")
            columns = [row[0] for row in cursor.fetchall()]
            num_columns = len(columns)
            print(f"Número de colunas na tabela Oracle: {num_columns}")

            # Reordenar as colunas para garantir que estão na mesma ordem que a tabela
            data = data[columns]
            
            # Inserir dados na tabela Oracle (sem formatação adicional)
            for index, row in data.iterrows():
                values = row.tolist()  # Converte a linha em lista de valores
                
                # Substituir strings vazias por None para garantir que sejam tratados como NULL
                values = [None if v == '' or pd.isna(v) else v for v in values]

                # Criar a query de inserção com placeholders
                placeholder = ', '.join(['?' for _ in range(num_columns)])
                command = f"INSERT INTO {table} ({', '.join(columns)}) VALUES ({placeholder})"
                
                # Debug
                # print(f"Comando SQL: {comando}")
                # print(f"Valores fornecidos para a linha {index + 1}: {valores}")

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
