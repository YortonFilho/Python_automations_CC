import pandas as pd
import requests
import base64
import sys
import pyodbc
import os
from unidecode import unidecode

# Importando as funções globais para reaproveitamento de código
sys.path.append("C:\\data-integration\\Automatizacao\\Python\\Yorton\\Python_automations_CC\\functions")
from data_base import db_connection  # Função para conectar com banco de dados
from colors import green, red  # Função para utilizar cores nos prints para melhor visualização

# Função para coletar dados da API
def data_colect():
    username = os.getenv('CHAVE_API_PAGARME')
    password = os.getenv('SENHA_API_PAGARME')
    url = 'https://api.pagar.me/1/orders?count=1000'

    # Formata e codifica as credenciais para autenticação
    credentials = f'{username}:{password}'
    b64_credentials = base64.b64encode(credentials.encode()).decode()

    headers = {
        'Authorization': f'Basic {b64_credentials}'
    }

    data_list = [] 
    page_id = None
    count = 1

    # Laço para coletar dados enquanto houver páginas
    while True:
        try:
            # Configura parâmetros para a requisição
            params = {'cursor': page_id} if page_id else {}
            response = requests.get(url, headers=headers, params=params)

            if response.status_code == 200:
                data = response.json()
                data_list.extend(data)
                print(green(f"Dados da página {count} importados com sucesso!"))

                count += 1
                page_id = response.headers.get('x-cursor-nextpage')  # Obtém o ID da próxima página

                if not page_id:
                    break
            else:
                print(red("Erro ao fazer requisição para API!"))

        except Exception as e:
            print(red(f"Erro ao extrair dados: {e}"))

    # Convertendo a lista de dados em um DataFrame
    df = pd.DataFrame(data_list)

    # Expandindo a coluna 'items' se ela existir
    if 'items' in df.columns:
        items_expanded = df['items'].explode()
        items_normalized = pd.json_normalize(items_expanded)
        items_normalized.columns = [f'items_{col}' for col in items_normalized.columns] 
        df = df.drop(columns='items').join(items_normalized)

    # Selecionando colunas necessárias
    required_columns = ['id', 'status', 'amount', 'payment_link_id', 'date_created',
                        'items_id', 'items_title', 'items_unit_price']

    df_filtered = df.loc[:, required_columns]  # Filtra as colunas no DataFrame

    # Convertendo as colunas de data para o formato 'dd/mm/aaaa'
    df_filtered['date_created'] = df_filtered['date_created'].str[:10]
    df_filtered['date_created'] = pd.to_datetime(df_filtered['date_created']).dt.strftime('%d/%m/%Y')

    df_filtered['status'] = df_filtered['status'].replace({'created': 'CRIADO', 'paid': 'PAGO'})

    # Convertendo valores para float
    df_filtered['amount'] = df_filtered['amount'].astype(float) / 100
    df_filtered['items_unit_price'] = df_filtered['items_unit_price'].astype(float) / 100

    # Renomeando colunas
    df_filtered = df_filtered.rename(columns={
        'id': 'ID',
        'status': 'STATUS',
        'amount': 'VALOR',
        'payment_link_id': 'ID_LINK_PAGAMENTO',
        'date_created': 'DATA_CRIACAO',
        'items_id': 'ID_ITEM',
        'items_title': 'ITEM',
        'items_unit_price': 'VALOR_UNITARIO_ITEM'
    })

    return df_filtered

def data_updating(final_df):
    try:
        with db_connection() as connection:
            with connection.cursor() as cursor:
                table = 'DADOS_PAGARME_V1_ORDERS' 
                
                insert_command = f"""
                    INSERT INTO {table} (
                        STATUS,
                        ID,
                        VALOR,
                        ID_LINK_PAGAMENTO,
                        DATA_CRIACAO,
                        ID_ITEM,
                        ITEM,
                        VALOR_UNITARIO_ITEM
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """

                delete_command = f"DELETE FROM {table}"

                try:
                    cursor.execute(delete_command)
                    print(green("Dados deletados com sucesso!"))
                except pyodbc.Error as e:
                    print(red(f"Erro ao deletar dados! {e}"))
                    exit()

                for index, row in final_df.iterrows():
                    try:
                        amount = float(row['VALOR'])
                        unit_price = float(row['VALOR_UNITARIO_ITEM'])
                    except ValueError:
                        print(red(f"Valor inválido encontrado na linha {index}: {row}"))
                        continue

                    # Normaliza o nome do item
                    item_normalized = unidecode(row['ITEM'])

                    values = row.loc[['STATUS', 'ID', 'VALOR', 'ID_LINK_PAGAMENTO',
                                      'DATA_CRIACAO', 'ID_ITEM', 
                                      'VALOR_UNITARIO_ITEM']].tolist()

                    # Adiciona o ITEM normalizado à lista de valores
                    values.insert(6, item_normalized)

                    try:
                        cursor.execute(insert_command, values)
                    except pyodbc.Error as e:
                        print(red(f"Erro ao inserir dados: {e} - Valores: {values}"))
                    except UnicodeEncodeError:
                        print(red(f"Erro de codificação ao inserir dados: {values}"))
                    except Exception as e:
                        print(red(f"Erro inesperado ao inserir dados: {e} - Valores: {values}"))

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
