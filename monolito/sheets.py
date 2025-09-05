# %% [markdown] 
### ENTREGAS 2025

from google.oauth2 import service_account
from googleapiclient.discovery import build
import pandas as pd
import datetime
import locale
from dotenv import load_dotenv
import os

dotenv_path = r"..\config\.env"
load_dotenv(dotenv_path)

SERVICE_ACCOUNT_FILE = r"..\config\credentials.json"

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']  # leitura + escrita


# Pega as variáveis
SPREADSHEET_ID = os.getenv("SHEETS-ENTREGA2025")
SHEET_NAME = 'ENTREGAS'  


credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
service = build('sheets', 'v4', credentials=credentials)

result = service.spreadsheets().values().get(
    spreadsheetId=SPREADSHEET_ID,
    range=SHEET_NAME 
).execute()

values = result.get('values', [])
df = pd.DataFrame(values)

df
# %% [markdown] 
### ATUALIZAÇÃO

atualizacao = df

# Filtrar a Primeira Linha
atualizacao = atualizacao.iloc[[0]]

# selecionar colunas 
atualizacao = atualizacao[[
    atualizacao.columns[1], 
    atualizacao.columns[2]
]].copy()
atualizacao = atualizacao.reset_index(drop=True)

# Renomear colunas
atualizacao = atualizacao.rename(columns={
    atualizacao.columns[0]: "Horas",
    atualizacao.columns[1]: "Data"
})


atualizacao

# %% [markdown] 
### GANTT

gantt = df

# Filtrar Linhas Iniciais
gantt = gantt.iloc[2:].copy()

# Cabeçalhos Promovidos
gantt.columns = gantt.iloc[0]
gantt = gantt.drop(gantt.index[0])
gantt = gantt.reset_index(drop=True)

# Filtrar linhas não vazias
gantt = gantt[gantt[gantt.columns[1]].notna()].copy()

# Selecionar colunas
gantt = gantt[[
    gantt.columns[3], 
    gantt.columns[4]
]].copy()

# renomear colunas imediatamente após selecionar
gantt.columns = ['TAREFA', 'ENTREGA']

# Converter datas para o formato "DD/MM/YYYY" sem função
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR')
    except locale.Error:
        pass

# Remove o nome do dia da semana e converte para "DD/MM/YYYY"
gantt['ENTREGA'] = gantt['ENTREGA'].astype(str)
gantt['ENTREGA'] = gantt['ENTREGA'].str.split(',', n=1).str[-1].str.strip()
gantt['ENTREGA'] = gantt['ENTREGA'].apply(
    lambda x: datetime.datetime.strptime(x, "%d de %B de %Y").strftime("%d/%m/%Y") if 'de' in x else x
)

# Filtrar linhas em branco
gantt = gantt[
    (gantt['ENTREGA'] != "") & 
    (gantt['ENTREGA'] != "-")
].copy()
gantt = gantt.reset_index(drop=True)

gantt


# %% [markdown] 
# _GANTT: Inserir dados na planilha de tempo_

# Transformar o DataFrame em lista de listas
gantt_values = [gantt.columns.to_list()] + gantt.values.tolist()

dotenv_path = r"..\config\.env"
load_dotenv(dotenv_path)

NEW_SPREADSHEET_ID = os.getenv("SHEETS-GANTT")
SHEET_NAME = 'GANTT'  # nome da aba que você quer inserir

range_start = 'A5'  # pode ser 'B2', 'C5', etc.

range_to_write = f'{SHEET_NAME}!{range_start}'

body = {
    'values': gantt_values
}

result = service.spreadsheets().values().update(
    spreadsheetId=NEW_SPREADSHEET_ID,
    range=range_to_write,
    valueInputOption='RAW',
    body=body
).execute()

print(f"{result.get('updatedCells')} células atualizadas.")

# %%
