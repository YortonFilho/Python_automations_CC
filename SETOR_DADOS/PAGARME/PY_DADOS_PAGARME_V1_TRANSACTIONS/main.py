import pandas as pd
import requests
import base64
import sys
import pyodbc
import os

# Importando as funções globais para reaproveitamento de código
sys.path.append("C:\\data-integration\\Automatizacao\\Python\\Yorton\\Python_automations_CC\\functions")
from data_base import db_connection  # Função para conectar com banco de dados
from colors import green, red  # Função para utilizar cores nos prints para melhor visualização

# Função para coletar dados da API
def data_colect():
    username = os.getenv('CHAVE_API_PAGARME')
    password = os.getenv('SENHA_API_PAGARME')
    url = 'https://api.pagar.me/1/transactions?count=1000'

    # Formata e codifica as credenciais para autenticação
    credentials = f'{username}:{password}'
    b64_credentials = base64.b64encode(credentials.encode()).decode()

    headers = {
        'Authorization': f'Basic {b64_credentials}'
    }

    data_list = []  # Lista para armazenar os dados coletados
    page_id = None  # Variável para controlar a paginação
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

                # if count > 200:
                #     break
                if not page_id:  # Se não houver mais páginas, sai do loop
                    print(green("Dados coletados com sucesso!"))
                    break
            else:
                print(red("Erro ao fazer requisição para API!"))
                break

        except Exception as e:
            print(red(f"Erro ao extrair dados: {e}"))

    # Convertendo a lista de dados em um DataFrame
    df = pd.DataFrame(data_list)

    # Expandindo JSONs e arrays dentro das colunas
    for column in df.columns:
        if df[column].apply(lambda x: isinstance(x, list) or isinstance(x, dict)).any():
            if df[column].apply(lambda x: isinstance(x, list)).any():
                df = df.explode(column) 
            if df[column].apply(lambda x: isinstance(x, dict)).any():
                temp_df = df[column].apply(pd.Series)  
                temp_df.columns = [f"{column}_{subcol}" for subcol in temp_df.columns] 
                df = df.join(temp_df, rsuffix='_expanded')
                df.drop(columns=[column], inplace=True)

    # Filtra apenas as colunas necessárias
    required_columns = [
        'status', 'date_created', 'date_updated', 'amount', 'authorized_amount',
        'paid_amount', 'id', 'payment_method', 
        'customer_id', 'customer_document_number', 'customer_name',
        'order_id', 'discount', 'payment', 'card_brand'
    ]

    df_filtered = df.loc[:, required_columns]  # Filtra as colunas no DataFrame

    # Convertendo as colunas de data para o formato 'dd/mm/aaaa'
    df_filtered['date_created'] = df_filtered['date_created'].str[:10]  
    df_filtered['date_created'] = pd.to_datetime(df_filtered['date_created']).dt.strftime('%d/%m/%Y')
    df_filtered['date_updated'] = df_filtered['date_updated'].str[:10]  
    df_filtered['date_updated'] = pd.to_datetime(df_filtered['date_updated']).dt.strftime('%d/%m/%Y')

    # Convertendo status e métodos de pagamento para português
    df_filtered['status'] = df_filtered['status'].replace({'paid': 'PAGO', 'waiting_payment': 'AGUARDANDO PAGAMENTO', 'refused': 'RECUSADO'})
    df_filtered['payment_method'] = df_filtered['payment_method'].replace({'credit_card': 'cartao de credito'})

    # Convertendo as colunas de valor para float
    df_filtered['amount'] = df_filtered['amount'].astype(float) / 100
    df_filtered['authorized_amount'] = df_filtered['authorized_amount'].astype(float) / 100
    df_filtered['paid_amount'] = df_filtered['paid_amount'].astype(float) / 100

    # Renomeando as colunas
    df_filtered.rename(columns={
        'order_id': 'ID_PEDIDO',
        'status': 'STATUS',
        'date_created': 'DATA_CRIACAO',
        'date_updated': 'DATA_ATUALIZACAO',
        'amount': 'VALOR',
        'authorized_amount': 'VALOR_AUTORIZADO',
        'paid_amount': 'VALOR_PAGO',
        'id': 'ID',
        'payment_method': 'METODO_PAGAMENTO',
        'customer_id': 'ID_CLIENTE',
        'customer_document_number': 'CPF',
        'customer_name': 'NOME_CLIENTE',
        'discount': 'DESCONTO',
        'payment': 'PAGAMENTO',
        'card_brand': 'BANDEIRA_CARTAO'
    }, inplace=True)

    return df_filtered 

# Função para inserir dados no banco de dados
def data_updating(final_df):
    try:
        with db_connection() as connection:  # Conecta ao banco de dados
            with connection.cursor() as cursor:
                table = 'DADOS_PAGARME_V1_TRANSACTIONS'

                insert_command = f"""
                    INSERT INTO {table} (
                        STATUS,
                        ID_PEDIDO,
                        NOME_CLIENTE,
                        DATA_CRIACAO,
                        DATA_ATUALIZACAO,
                        VALOR,
                        VALOR_AUTORIZADO,
                        VALOR_PAGO,
                        ID_CLIENTE,
                        CPF,
                        METODO_PAGAMENTO,
                        DESCONTO,
                        PAGAMENTO,
                        BANDEIRA_CARTAO
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """
                
                delete_command = f"DELETE FROM {table}" 

                # Listar colunas da tabela para depuração
                # cursor.execute(f"SELECT COLUMN_NAME FROM ALL_TAB_COLUMNS WHERE TABLE_NAME = '{table}'")
                # columns = [row[0] for row in cursor.fetchall()]  # Obtém os nomes das colunas
                # print(f"Colunas na tabela {table}: {columns}")

                # Verifique os nomes das colunas no DataFrame
                # print("Colunas no DataFrame:", final_df.columns.tolist())

                try:
                    cursor.execute(delete_command)
                    print(green("Dados deletados com sucesso!"))
                except pyodbc.Error as e:
                    print(red(f"Erro ao deletar dados! {e}"))
                    exit()

                # Inserindo dados no banco de dados
                for index, row in final_df.iterrows():
                    values = row.loc[['STATUS', 'ID_PEDIDO', 'NOME_CLIENTE', 'DATA_CRIACAO', 
                                      'DATA_ATUALIZACAO', 'VALOR', 'VALOR_AUTORIZADO', 
                                      'VALOR_PAGO', 'ID_CLIENTE', 'CPF', 
                                      'METODO_PAGAMENTO', 'DESCONTO', 'PAGAMENTO', 
                                      'BANDEIRA_CARTAO']].tolist()  # Extrai valores da linha

                    # Debug: Mostre os valores
                    # print("Valores a serem inseridos:", values)

                    try:
                        cursor.execute(insert_command, values)
                    except pyodbc.Error as e:
                        print(red(f"Erro ao inserir dados: {e}"))

                connection.commit()
                print(green("Dados importados com sucesso!"))

    except pyodbc.Error as e:
        print(red(f"Erro ao conectar ou interagir com o banco de dados: {e}"))

# Execução principal
if __name__ == "__main__":
    final_df = data_colect()
    
    if final_df.empty:  # Verifica se o DataFrame está vazio
        print(red("Nenhum dado encontrado para inserir no banco."))
    else:
        data_updating(final_df)
