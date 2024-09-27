import requests
import pandas as pd
import re
from datetime import datetime, timedelta
import pyodbc
import os
import sys

sys.path.append("C:\\Users\\yorton.filho\\Desktop\\Python_automations_CC\\functions")
from data_base import db_connection  # função para conectar com banco de dados
from colors import green, red  # função para utilizar cores nos prints para melhor visualização

# Função recursiva para normalizar colunas com JSONs ou arrays
def normalize_columns(df):
    while True:
        columns_to_expand = [col for col in df.columns if df[col].apply(lambda x: isinstance(x, dict) or isinstance(x, list)).any()]
        if not columns_to_expand:
            break
        for column in columns_to_expand:
            if df[column].apply(lambda x: isinstance(x, dict)).any():
                json_df = pd.json_normalize(df[column])
                json_df.columns = [f"{column}_{subcol}" for subcol in json_df.columns]
                df = pd.concat([df.drop(columns=[column]), json_df], axis=1)
            elif df[column].apply(lambda x: isinstance(x, list)).any():
                array_df = df[column].apply(lambda x: pd.Series(x) if isinstance(x, list) else pd.Series([x]))
                array_df.columns = [f"{column}_{i}" for i in array_df.columns]
                df = pd.concat([df.drop(columns=[column]), array_df], axis=1)
    return df

# Função para converter timestamps para datas e formatar como DD/MM/AAAA
def convert_and_format_dates(df, columns):
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], unit='ms', errors='coerce').dt.strftime('%d/%m/%Y')
    return df

# Função para renomear as colunas
def rename_columns(df, column_mapping):
    df = df.rename(columns=column_mapping)
    return df

# Função para adicionar a coluna MES_REF
def add_mes_ref_column(df):
    def get_mes_ref(date_str):
        if pd.isna(date_str):
            return None
        
        date = pd.to_datetime(date_str, format='%d/%m/%Y')
        if date.day in [29, 30, 31]:
            if date.month == 12:
                next_month = 1
                next_year = date.year + 1
            else:
                next_month = date.month + 1
                next_year = date.year
            return f'01/{next_month:02d}/{next_year}'
        else:
            return f'01/{date.month:02d}/{date.year}'
    
    df['MES_REF'] = df['DATA_CRIACAO'].apply(get_mes_ref)
    return df

# Função para ajustar a coluna SLA e renomear para SLA_AJUSTADO
def adjust_sla_column(df):
    df['SLA_AJUSTADO'] = df['SLA'].apply(lambda x: ''.join(re.findall(r'\d+', str(x))))
    return df

# Função para adicionar a coluna FILTRO com base na coluna SLA_AJUSTADO
def add_filtro_column(df):
    df['FILTRO'] = df['SLA_AJUSTADO'].apply(lambda x: 'OK' if int(x) < 3 else 'FORA')
    return df

# Função para calcular a data de início com base no dia atual
def calculate_start_date():
    today = datetime.now()
    if today.day > 29:
        start_date = today.replace(day=29).strftime('%Y-%m-%d')
    else:
        if today.month == 1:
            start_date = today.replace(year=today.year - 1, month=12, day=29).strftime('%Y-%m-%d')
        else:
            start_date = today.replace(month=today.month - 1, day=29).strftime('%Y-%m-%d')
    
    return start_date

# Variáveis da API
token = os.getenv('CHAVE_API_CLICKUP')
url = 'https://api.clickup.com/api/v2/view/12zuj6-6773/task'
headers = { 'Authorization': token }

# Calcula a data de início e a data de ontem
start_date = calculate_start_date()
yesterday = datetime.now() - timedelta(1)
end_date = yesterday.strftime('%Y-%m-%d')

# Função para fazer a requisição e coletar todos os dados
def fetch_all_tasks(url, headers, start_date, end_date):
    all_tasks = []
    page = 0
    while True:
        response = requests.get(f"{url}?page={page}&start_date={start_date}&end_date={end_date}", headers=headers)
        if response.status_code == 200:
            data = response.json()
            tasks = data.get('tasks', [])
            all_tasks.extend(tasks)
            if len(tasks) == 0:
                break
            page += 1
        else:
            print(f"Falha na requisição. Status code: {response.status_code}")
            print(f"Mensagem de erro: {response.text}")
            break
    return all_tasks

# Fazendo a requisição para coletar todos os dados
tasks = fetch_all_tasks(url, headers, start_date, end_date)

# Convertendo os dados para um DataFrame do pandas
df = pd.DataFrame(tasks)

# Normalizando as colunas com JSONs ou arrays
df = normalize_columns(df)

# Selecionando apenas as colunas desejadas
desired_columns = [
    'name', 
    'assignees_0_username', 
    'status_status', 
    'date_created', 
    'date_closed', 
    'custom_fields_3_value'
]

available_columns = [col for col in desired_columns if col in df.columns]
df_filtered = df[available_columns]
df_filtered = df_filtered[desired_columns]

# Converte os timestamps para datas e formata como DD/MM/AAAA
df_filtered = convert_and_format_dates(df_filtered, ['date_created', 'date_closed'])

# Mapeamento dos novos nomes das colunas
column_mapping = {
    'name': 'NOME_TAREFA',
    'assignees_0_username': 'RESPONSAVEL',
    'status_status': 'STATUS',
    'date_created': 'DATA_CRIACAO',
    'date_closed': 'DATA_FECHAMENTO',
    'custom_fields_3_value': 'SLA'
}

# Renomeia as colunas
df_filtered = rename_columns(df_filtered, column_mapping)

# Adiciona a coluna MES_REF
df_filtered = add_mes_ref_column(df_filtered)

# Ajusta a coluna SLA para incluir apenas números e renomeia para SLA_AJUSTADO
df_filtered = adjust_sla_column(df_filtered)

# Adiciona a coluna FILTRO com base na coluna SLA_AJUSTADO
df_filtered = add_filtro_column(df_filtered)

# Reordena as colunas conforme a ordem desejada, incluindo a nova coluna
final_columns = [
    'NOME_TAREFA', 
    'RESPONSAVEL', 
    'STATUS', 
    'DATA_CRIACAO', 
    'DATA_FECHAMENTO', 
    'SLA',
    'MES_REF',
    'SLA_AJUSTADO',
    'FILTRO'
]
df_filtered = df_filtered[final_columns]

# Salvando o DataFrame filtrado, reordenado e renomeado em um arquivo Excel
df_filtered.to_excel('dados_clickup.xlsx', index=False)
print("Dados exportados com sucesso para 'dados_clickup.xlsx'")

# Processando o DataFrame e inserindo no banco de dados
try:
    with db_connection() as connection:
        with connection.cursor() as cursor:
            table = 'DADOS_OUVIDORIA'
            
            # Obter a estrutura da tabela Oracle
            cursor.execute(f"SELECT COLUMN_NAME FROM ALL_TAB_COLUMNS WHERE TABLE_NAME = '{table}'")
            columns = [row[0] for row in cursor.fetchall()]
            num_columns = len(columns)
            
            # Verifica se o número de colunas na tabela Oracle corresponde ao número de colunas no DataFrame
            if len(df_filtered.columns) != num_columns:
                print(red(f"Número de colunas no DataFrame ({len(df_filtered.columns)}) não corresponde ao número de colunas na tabela Oracle ({num_columns})."))
                exit()
            
            # Reordenar as colunas do DataFrame para corresponder à ordem da tabela Oracle
            df_filtered = df_filtered[columns]
            
            # Converter os tipos de dados, se necessário
            for col in columns:
                if 'DATA' in col.upper():
                    continue
                elif 'NUM' in col.upper():
                    df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce')

            # Obter a data atual
            date = datetime.now()
            month = date.month
            year = date.year

            # Verifica se o dia atual é 29, 30 ou 31
            if date.day >= 29:
                if month == 12:
                    next_month = 1
                    next_year = year + 1
                else:
                    next_month = month + 1
                    next_year = year
            else:
                next_month = month
                next_year = year

            # Formata o primeiro dia do próximo mês
            next_month_start = f'01/{next_month:02d}/{next_year}'

            # Comando para deletar dados do banco
            delete_command = f"DELETE FROM {table} WHERE MES_REF = '{next_month_start}'"
            
            # Deletar dados do banco
            try:
                cursor.execute(delete_command)
                print(green("Dados deletados com sucesso!"))
            except pyodbc.Error as e:
                print(red(f"Erro ao deletar dados: {e}"))
                exit()

            # Criar a query de inserção
            placeholders = ', '.join(['?' for _ in range(num_columns)])
            insert_command = f"INSERT INTO {table} ({', '.join(columns)}) VALUES ({placeholders})"
            
            # Inserir dados na tabela Oracle
            for index, row in df_filtered.iterrows():
                values = row.tolist()
                values = [None if v == '' or pd.isna(v) else v for v in values]
                try:
                    cursor.execute(insert_command, values)
                except pyodbc.Error as e:
                    print(red(f"Erro ao inserir dados: {e}"))
                    exit()

            connection.commit()
            print(green("Dados importados com sucesso!"))

except pyodbc.Error as e:
    print(red(f"Erro ao conectar ou interagir com o banco de dados: {e}"))
