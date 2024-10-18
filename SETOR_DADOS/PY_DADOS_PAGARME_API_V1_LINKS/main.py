import base64
import requests
import pandas as pd
import pyodbc
import sys
from datetime import datetime, timedelta
import os

# Importando as funções globais para reaproveitamento de código
sys.path.append("C:\\data-integration\\Automatizacao\\Python\\Yorton\\Python_automations_CC\\functions")
from data_base import db_connection  # função para conectar com banco de dados
from colors import green, red  # função para utilizar cores nos prints para melhor visualização

# Função para coletar dados da API
def data_colect():
    username = os.getenv('CHAVE_API_PAGARME')
    password = os.getenv('SENHA_API_PAGARME')
    url = 'http://api.pagar.me/1/payment_links?count=1000'

    # Formata e codifica as credenciais
    credentials = f'{username}:{password}'
    b64_credentials = base64.b64encode(credentials.encode()).decode()

    # Cria o cabeçalho
    headers = {
        'Authorization': f'Basic {b64_credentials}'
    }

    # Inicializa uma lista para armazenar todos os dados
    all_data = []
    page_id = None
    count = 1

    # Laço de repetição para puxar dados até a última página
    while True:
        params = {'cursor': page_id} if page_id else {}
        response = requests.get(url, headers=headers, params=params)

        if response.status_code == 200:
            data = response.json()  # Converte a resposta para JSON
            all_data.extend(data)  # Adiciona os dados à lista
            print(green(f"Dados extraídos da página {count} com sucesso!"))

            count =+ 1
            page_id = response.headers.get('x-cursor-nextpage')  # verifica o id da próxima página
            if not page_id:  # Se não houver mais páginas, sai do loop
                break
        else:
            print(red(f"Erro na requisição: {response.status_code}"))
            break

    # Cria um DataFrame a partir dos dados
    df = pd.DataFrame(all_data)

    # Filtra colunas específicas
    columns_to_keep = ['id', 'amount', 'url', 'date_created', 'status', 'name', 'orders_paid']
    filtered_df = df[columns_to_keep]

    # Reordena e renomeia as colunas
    final_df = filtered_df.rename(columns={
        'status': 'STATUS',
        'id': 'ID_DO_LINK',
        'name': 'NOME',
        'date_created': 'DATA_DE_CRIACAO',
        'url': 'LINK',
        'amount': 'VALOR_PAGO',
        'orders_paid': 'ORDERS_PAID'
    })

    # Filtra e formata o DataFrame
    final_df['STATUS'] = final_df['STATUS'].replace({'active': 'ATIVO', 'canceled': 'CANCELADO', 'expired': 'INATIVO'})
    final_df['VALOR_PAGO'] = final_df['VALOR_PAGO'].astype(float) / 100
    final_df = final_df[final_df['ORDERS_PAID'] == 1]

    # Extrair os primeiros 10 caracteres da coluna de data e converter para DD/MM/AAAA
    final_df['DATA_DE_CRIACAO'] = final_df['DATA_DE_CRIACAO'].str[:10]
    final_df['DATA_DE_CRIACAO'] = pd.to_datetime(final_df['DATA_DE_CRIACAO']).dt.strftime('%d/%m/%Y')

    return final_df

# Função para inserir dados no banco de dados
def data_updating(final_df):
    try:
        with db_connection() as connection:
            with connection.cursor() as cursor:
                table = 'DADOS_PAGARME_API_V1_LINKS'

                insert_command = f"""
                    INSERT INTO {table} (
                        STATUS,
                        ID_DO_LINK,
                        NOME,
                        DATA_DE_CRIACAO,
                        LINK,
                        VALOR_PAGO
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """
                
                delete_command = f"""DELETE {table}"""
                
                # Listar colunas da tabela para depuração
                cursor.execute(f"SELECT COLUMN_NAME FROM ALL_TAB_COLUMNS WHERE TABLE_NAME = '{table}'")
                columns = [row[0] for row in cursor.fetchall()]
                print(green(f"Colunas na tabela {table}: {columns}"))

                # Verifique os nomes das colunas no DataFrame
                print("Colunas no DataFrame:", final_df.columns.tolist())

                try:
                        cursor.execute(delete_command)
                        print(green("Dados deletados com sucesso!"))
                except pyodbc.Error as e:
                    print(red(f"Erro ao deletar dados! {e}"))
                    exit()

                for index, row in final_df.iterrows():
                    # Converte a data para o tipo datetime
                    formatted_date = row['DATA_DE_CRIACAO']
                    
                    # Usar loc para acessar as colunas
                    values = row.loc[['STATUS', 'ID_DO_LINK', 'NOME', 'LINK', 'VALOR_PAGO']].tolist()
                    values.insert(3, formatted_date)  # Insere a data convertida na posição correta

                    # Debug: Mostre os valores
                    print("Valores a serem inseridos:", values)

                    try:
                        cursor.execute(insert_command, values)
                    except pyodbc.Error as e:
                        print(red(f"Erro ao inserir dados: {e}"))
                        exit()

                connection.commit()
                print(green("Dados importados com sucesso!"))

    except pyodbc.Error as e:
        print(red(f"Erro ao conectar ou interagir com o banco de dados: {e}"))

# Execução principal
if __name__ == "__main__":
    final_df = data_colect()
    
    if final_df.empty:
        print(red("Nenhum dado encontrado para inserir no banco."))
    else:
        data_updating(final_df)