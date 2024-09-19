
# README para automações em Python

Este README fornece uma visão geral das automações que eu desenvolvi para a minha empresa, operações de importação e exportação de dados. Os scripts interagem com um banco de dados Oracle e manipulam arquivos Excel e CSV, além de utilizar APIs.

## Descrição

As automações que começam com "PY_", são automações implementadas no Jenkins, programadas para acontecerem de forma 100% automática. 

Os scripts que começam com "IMPORT_", são scripts feitos para facilitarem a importação de dados no banco de dados.


## Informações sobre cada automação

#### PY_CAMPANHA_ANIVER:

Função: Exporta dados de aniversariantes do banco de dados Oracle para um arquivo CSV, depois converte esse CSV para um array para importar e atualizar uma campanha via API. Determina o ID da campanha com base no dia da semana (são três campanhas para atualizar, cada uma em dias diferentes).

#### PY_DADOS_OUVIDORIA

Função: Primeiro, filtra dados de um espaço específico no ClickUp, se baseando no dia 29 do mes passado até um dia anterior do dia atual, e os extrai usando a API. Em seguida, trata esses dados, realizando as conversões necessárias, e os salva em um arquivo Excel. Após isso, o script remove os dados referentes ao mês atual no banco de dados Oracle e importa os dados extraídos para atualizar o banco. 

#### PY_DADOS_FARMACIA

Função: Exporta dados de novos assinantes pelo banco de dados Oracle para um arquivo CSV, para descontos em farmácias (Não está 100% automatizada pois preciso fazer cruzamentos com outras tabelas usando filtros antes de enviar por email, passos esses que não consegui automatizar ainda, por enquanto a única coisa que essa automação faz é deixar o arquivo atualizado pra mim toda manhã, 5 minutos antes de eu chegar)

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
