import pandas as pd
import pyodbc
import tkinter as tk
from tkinter import filedialog
import sys

# Importando as funções globais para reaproveitamento de código
sys.path.append("C:\\data-integration\\Automatizacao\\Python\\Yorton\\Python_automations_CC\\functions")
from data_base import db_connection  # função para conectar com banco de dados
from colors import green, red  # função para utilizar cores nos prints para melhor visualização

# função para selecionar os arquivos a serem analisados
def select_file():
    root = tk.Tk()
    root.withdraw()
    file = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    return file

# Selecionar arquivo Excel
file = select_file()
excel_file = file[0]

# Ler os dados do arquivo Excel
try:
    data = pd.read_excel(excel_file)
    print(green("Arquivo Excel lido com sucesso!"))
except Exception as e:
    print(red(f"Erro ao ler o arquivo Excel: {e}"))
    exit()

# Limpar caracteres especiais usando applymap para cada coluna
data = data.apply(lambda col: col.map(lambda x: x.encode('ascii', 'ignore').decode('ascii') if isinstance(x, str) else x))

# Converter data para o formato DD/MM/YYYY
if 'DATA_AGENDA' in data.columns:
    data['DATA_AGENDA'] = pd.to_datetime(data['DATA_AGENDA'], errors='coerce').dt.strftime('%d/%m/%Y')

# Truncar values na coluna "RESERVA" se necessário
if 'RESERVA' in data.columns:
    data['RESERVA'] = data['RESERVA'].apply(lambda x: x[:100] if isinstance(x, str) else x)

# Inserir dados na tabela existente
try:
    with db_connection() as connection:  # conectando com banco de dados
        with connection.cursor() as cursor:  # abrindo central para rodar a query
            tabel = 'DADOS_AGENDAS_MEDICAS'
            
            # Obter a estrutura da tabela
            cursor.execute(f"SELECT COLUMN_NAME FROM ALL_TAB_COLUMNS WHERE TABLE_NAME = '{tabel.upper()}'")
            colunas = [row[0] for row in cursor.fetchall()]
            num_colunas = len(colunas)
            
            print(f"Número de colunas na tabela Oracle: {num_colunas}")

            # Reordenar as colunas
            data = data[colunas]

            # Inserir dados na tabela
            for index, row in data.iterrows():
                values = row.tolist()
                values = [None if pd.isna(v) or v == '' else v for v in values]
                espaco_reservado = ', '.join(['?' for _ in range(num_colunas)])
                command = f"INSERT INTO {tabel} ({', '.join(colunas)}) VALUES ({espaco_reservado})"
            
                try:
                    cursor.execute(command, values)
                except pyodbc.Error as e:
                    print(red(f"Erro ao inserir dados na linha {index + 1}: {e}"))

            connection.commit()
            print(green("Dados importados com sucesso!"))

except pyodbc.Error as e:
    print(red(f"Erro ao conectar ou interagir com o banco de dados: {e}"))