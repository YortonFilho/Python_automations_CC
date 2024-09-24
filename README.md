
# README para automações em Python

Este README fornece uma visão geral das automações que eu desenvolvi para a minha empresa, operações de importação e exportação de dados. Os scripts interagem com um banco de dados Oracle e manipulam arquivos Excel e CSV, além de utilizar APIs.

## Descrição

As automações que começam com "PY_", são automações implementadas no Jenkins, programadas para acontecerem de forma 100% automática. 

Os scripts que começam com "IMPORT_", são scripts feitos para facilitarem a importação de dados no banco de dados.


## Informações sobre cada automação

#### PY_CAMPANHA_ANIVER:

Função: Exporta dados de aniversariantes do dia atual, do banco de dados Oracle para um arquivo CSV (para ter salvo nos arquivos da empresa), depois converte esse CSV para um array para importar e atualizar os dados em uma campanha via API. Temos 3 campanhas disponíveis em relação aos aniversáriantes, a campanha certa é definida pelo seu ID, sendo validada por qual dia da semana está sendo rodado o script.

#### PY_DADOS_OUVIDORIA

Função: Primeiro, filtra dados de um espaço específico no ClickUp, se baseando no dia 29 do mes passado até um dia anterior do dia atual, e os extrai usando a API. Em seguida, trata esses dados, realizando as conversões necessárias, e os salva em um arquivo Excel. Após isso, o script remove os dados referentes ao mês atual no banco de dados Oracle e importa os dados extraídos para atualizar o banco. 

#### PY_DADOS_FARMACIA

Função: Exporta dados de assinantes de um plano da empresa pelo banco de dados Oracle, depois importa esses dados em um arquivo excel, atualizando a 1º aba. Esse arquivo tem 2 abas, uma com todos assinantes, e a outra com todos assinantes já enviados para as farmácias. Então faz um cruzamento entre essas 2 abas para gerar um arquivo com os novos assinantes que não foram enviados ainda. No meio do código tem uma condição, que envia um dos 3 tipos de emails específicos para cada situação, sendo um deles, o email com os novos assinantes para atualizar as farmácias. Por fim, o script atualiza a 2º aba com a data atual para no dia seguinte, fazer todo o processo de novo.

#### IMPORT_DADOS_X5_PERFORMANCE_AGENTES:

Função: Importa dados de um arquivo Excel para uma tabela do banco de dados Oracle. 

#### IMPORT_DADOS_RESULTADO_OPERADORES:

Função: Importa dados de um arquivo Excel para uma tabela do banco de dados Oracle, com conversão de valores numéricos. 

## Funções

As funções "db_connection" e "colors", localizadas na pasta functions, foram criadas para promover a reutilização de código e tornar as automações mais limpas e fáceis de manter.

## Variáveis de Ambiente

Para rodar os scripts, você vai precisar adicionar as seguintes variáveis de ambiente no seu .env

'NOME_BANCO_DE_DADOS'

'USUARIO_BANCO_DE_DADOS'

'SENHA_BANCO_DE_DADOS'

'CHAVE_API_X5'
## Como Rodar o Projeto

1 - Clone o Repositório

```bash
  git clone https://github.com/YortonFilho/Python_automations_CC.git
```
    
2 - Instale as Dependências:

```bash
  cd Python_automations_CC
  pip install -r requirements.txt
```
Crie um arquivo .env para fornecer as credenciais do seu banco de dados (Nome, Usuário e Senha) via variáveis de ambiente. 

Instale a biblioteca de variáveis de ambiente "dotenv"

Ajuste o caminho que deseja salvar os arquivos excel.

Ajuste as URLs e a chave da API.
