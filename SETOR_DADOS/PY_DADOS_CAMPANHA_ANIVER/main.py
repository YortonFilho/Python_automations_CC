import pyodbc
import pandas as pd
import os
from datetime import datetime
import requests
import json # fiz um print na linha 122 para debugging, deixei comentado por não precisar mais por enquanto
import sys

# importando as funções globais para reaproveitamento de código
sys.path.append("C:\\data-integration\\Automatizacao\\Python\\Yorton\\Python_automations_CC\\functions")
from data_base import db_connection  # função para conectar com banco de dados
from colors import green, red  # função para utilizar cores nos prints para melhor visualização

# Variáveis da data formatando com 2 casas decimais
date = datetime.now()
day = f'{date.day:02d}'
month = f'{date.month:02d}'
year = date.year

try:
    with db_connection() as connection:  # conectando com banco de dados pela função da pasta functions
        with connection.cursor() as cursor:  # abrindo central para rodar o comando query
            # Definindo comando da query
            command = f"""
                SELECT 
                    FONE_PRIMARIO_AJUSTADO AS telefone,
                    COD_PACIENTE AS cod_paciente,
                    DD_MM_ANIVER AS dd_nasc
                FROM 
                    PACIENTES_ANIVER
                WHERE 
                    DD_MM_ANIVER = '{day}/{month}'"""

            # Executando comando
            try:
                cursor.execute(command)
                print(green("Comando executado com sucesso!"))
            except:
                print(red("Erro ao buscar dados!"))
                exit()

            # Armazenando os nomes das colunas do SQL na variável columns
            columns = [desc[0] for desc in cursor.description]

            # Armazenando as linhas de dados na variável data
            data = cursor.fetchall()

            # Criando "tabela" e armazenando-a na variável dataFrame
            dataFrame = pd.DataFrame.from_records(data, columns=columns)

            # Formatando o cabeçalho das colunas para letras minúsculas
            dataFrame.columns = [col.lower() for col in dataFrame.columns]

            # Definindo o nome do arquivo
            csv_file = os.path.join(f'{year}_{month}_{day}.csv')

            # Criando arquivo CSV
            try:
                dataFrame.to_csv(csv_file, index=False, sep=';')
                print(green(f"Dados exportados para {csv_file} com sucesso!"))
            except:
                print(red("Erro ao exportar dados para o CSV!"))
                exit()
                
except pyodbc.Error as e:
    print(red(f"Erro ao conectar ou interagir com o banco de dados: {e}"))

#-------------------------------------------------------------------DEFININDO CAMPANHA A SER ATUALIZADA
# Verificando o dia da semana
date = datetime.now()
week_day = date.weekday()  # 0=segunda ... 6=domingo

# Se for dia de semana, vai ser uma campanha específica...
# se for sábado vai ser outra campanha e se for domingo será outra... 
if week_day < 5:
    campaign_id = 123
elif week_day == 5:
    campaign_id = 124
elif week_day == 6:
    campaign_id = 125
else:
    print(red("Erro ao definir ID da campanha!"))

#-------------------------------------------------------------------DELETANDO E IMPORTANDO DADOS PELA API
# Variáveis com as informações necessárias para acessar os endpoints
url_delete = f'http://192.168.1.252:8001/api/v2/campanha/{campaign_id}/contatos'
url_import = f'http://192.168.1.252:8001/api/v2/campanha/{campaign_id}/contatos_import'
token = os.getenv('CHAVE_API_X5')
header = {
    'Authorization': f'Bearer {token}',
    'Content-Type': 'application/json'
}

# deletando dados da campanha
try:
    response = requests.delete(url_delete, headers=header)
    if response.status_code == 200:
        print(green("Dados deletados com sucesso!"))
    else:
        print(red(f"Erro ao deletar dados! CÓDIGO: {response.status_code}, MOTIVO: {response.text}"))
        exit()
except Exception as e:
    print(red(f"Erro ao deletar dados! {e}"))
    exit()

# Lendo o arquivo CSV
try:
    df = pd.read_csv(csv_file, sep=';', dtype=str)
    
    # Garantir que todos os valores da coluna de telefone não venham com .0 no final
    if 'telefone' in df.columns:
        df['telefone'] = df['telefone'].str.replace('.0', '', regex=False)
    
    # Substituindo valores NaN por None
    df = df.where(pd.notnull(df), None)
    
    # Convertendo o DataFrame para uma lista de dicionários (JSON array)
    json_data = df.to_dict(orient='records')
    
    # Imprimir JSON para depuração
    # print("Dados JSON a serem enviados:", json.dumps(json_data, indent=2))
except:
    print(red("Erro ao ler o arquivo CSV e converter para JSON!"))
    exit()

# Verificando se json_data é uma lista não vazia
if not json_data:
    print(red("Nenhum dado foi convertido para JSON."))
    exit()

# Enviando o JSON para a API usando o método POST
try:
    response = requests.post(url_import, headers=header, json=json_data)
    if response.status_code == 200:
        print(green("Dados enviados com sucesso!"))
    else:
        print(red(f"Erro ao enviar dados! CÓDIGO: {response.status_code}, MOTIVO: {response.text}"))
except Exception as e:
    print(red(f"Erro ao acessar a API! Erro: {e}"))