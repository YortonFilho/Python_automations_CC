import pandas as pd
import os
import pyodbc
from functions.data_base import db_connection  # função para conectar com banco de dados
from functions.colors import green, red  # função para utilizar cores nos prints para melhor visualização
from datetime import datetime

# variáveis da data para nomear o arquivo a ser salvo, formatando com 2 casas decimais
date = datetime.now()
day = f'{date.day:02d}'
month = f'{date.month:02d}'
year = f'{date.year}'

try:
    with db_connection() as connection:  # conectando com banco de dados
        with connection.cursor() as cursor:  # abrindo central para rodar query
            try:
                cursor.execute("""
                SELECT
                    DISTINCT UPPER(T.DEPENDENTE)                                           AS NOME,
                    REPLACE(REPLACE(REPLACE(T.CPF_DEPENDENTE, '.', ''), '-', ''), '/', '') AS CPF_DEP
                FROM
                    DADOS_DEPENDENTES_E_TITULARES T
                WHERE
                    T.CPF_DEPENDENTE IS NOT NULL""")
            except:
                print(red("Erro ao executar comando!"))
                exit()
            
            # armazenando nome das colunas do banco, na variável columns
            columns = [desc[0] for desc in cursor.description]

            # armazenando os dados das colunas na variável data
            data = cursor.fetchall()

            # criando "tabela" e armazenado na variável dataFrame
            dataFrame = pd.DataFrame.from_records(data, columns=columns)

            # armazenando o caminho em que o arquivo será salvo
            folder = 'H:/Tecnologia/EQUIPE - DADOS/1 - Relatorios Recorrentes/Diario/05 - Envio Farmacias'

            # verificando se o caminho existe, senão irá criar com o mesmo nome
            if not os.path.exists(folder):
                os.makedirs(folder)

            # definindo local que será salvo o arquivo e o nome do arquivo seguindo o padrão da empresa
            excel_file = os.path.join(folder, f'19013906000179 - DR_CENTRAL_FARMACIAS {day}{month}{year}.xlsx')

            # gerando arquivo excel
            try:
                dataFrame.to_excel(excel_file, index=False, engine='openpyxl')
                print(green(f"Arquivo gerado com sucesso!"))
            except:
                print(red("Erro ao gerar arquivo!"))
                exit()

except pyodbc.Error as e:
    print(red(f"Erro ao conectar ou interagir com o banco de dados: {e}"))