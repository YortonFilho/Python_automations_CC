import pyodbc
import pandas as pd
import sys
from datetime import datetime

# Importando as funções globais para reaproveitamento de código
sys.path.append("C:\\data-integration\\Automatizacao\\Python\\Yorton\\Python_automations_CC\\functions")
from data_base import db_connection  # função para conectar com banco de dados
from colors import green, red  # função para utilizar cores nos prints para melhor visualização

# Caminho para o arquivo Excel
excel_file = 'C:\\temp\\IMPORTS\\IMPORT DRC NOTAS PREFEITURA.xlsx'

# Lendo arquivo excel
try:
    data = pd.read_excel(excel_file)
except Exception as e:
    print(red(f"Erro ao ler o arquivo Excel: {e}"))
    exit()

# Substitui valores NaN e strings vazias por None
data = data.applymap(lambda x: None if (pd.isna(x) or x == '') else x)

# Convertendo colunas de data para datetime
date_columns = ['DATA_EMISSAO', 'DATA_PRESTACAO_SERVICO']
for column in date_columns:
    if column in data.columns:
        data[column] = pd.to_datetime(data[column], format='%d/%m/%Y', errors='coerce')

try:
    with db_connection() as connection:  # conectando com banco de dados
        with connection.cursor() as cursor:  # abrindo central para rodar a query
            table = 'DRC_NOTAS_PREFEITURA'
            
            # Obter a estrutura da tabela Oracle
            cursor.execute(f"SELECT COLUMN_NAME FROM ALL_TAB_COLUMNS WHERE TABLE_NAME = '{table}'")
            columns = [row[0] for row in cursor.fetchall()]
            num_columns = len(columns)
            print(f"Número de colunas na tabela Oracle: {num_columns}")

            # Reordenar as colunas para garantir que estão na mesma ordem que a tabela
            data = data[columns]

            # Váriaveis da data para ajustar a query com o mes e ano atual
            date = datetime.now()
            month = f'{date.month:02d}'
            year = date.year

            # Criando query de exclusao de dados
            delete_command = f"""
            DELETE {table}
            WHERE TO_CHAR(DATA_EMISSAO, 'MM/RRRR') = '{month}/{year}'
            """

            # Excluindo dados do banco
            try:
                cursor.execute(delete_command)
                print(green("Dados deletados com sucesso!"))
            except pyodbc.Error as e:
                print(red("Erro ao deletar dados!"))
                exit()

            # Inserir dados na tabela Oracle
            for index, row in data.iterrows():
                values = row.tolist()  # Converte a linha em lista de valores

                # Criar a query de inserção com TO_DATE para datas
                insert_command = f"""
                INSERT INTO {table} (
                    {', '.join(columns)}
                ) VALUES (
                    {values[0]},
                    TO_DATE('{values[1].strftime('%d/%m/%Y')}', 'DD/MM/YYYY'),
                    TO_DATE('{values[2].strftime('%d/%m/%Y')}', 'DD/MM/YYYY'),
                    {', '.join('?' for _ in values[3:])}
                )
                """

                # Debug
                # print(f"Comando SQL: {command}")
                # print(f"Valores fornecidos para a linha {index + 1}: {values}")

                # Inserindo os dados no banco
                try:
                    cursor.execute(insert_command, values[3:])  # Passar apenas os valores restantes
                except pyodbc.Error as e:
                    print(red(f"Erro ao inserir dados: {e}"))
                    exit()

            connection.commit()
            print(green("Dados importados com sucesso!"))

except pyodbc.Error as e:
    print(red(f"Erro ao conectar ou interagir com o banco de dados: {e}"))
