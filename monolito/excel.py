# %% [markdown] 
### EXCEL
import pandas as pd
import re
from dotenv import load_dotenv
import os

dotenv_path = r"..\config\.env"
load_dotenv(dotenv_path)

caminho = os.getenv("EXCEL")

# fonte
df = pd.read_excel(caminho, sheet_name="MOBILIÁRIO")

# Coluna Condicional Adicionada
df["I1"] = None
df.loc[df[df.columns[1]] == "Qtd.", "I1"] = "1"
df.loc[df[df.columns[3]] == "TOTAL DOS PRODUTOS ", "I1"] = "2"

# Preenchido Abaixo
df["I1"] = df["I1"].ffill()

# Linhas Filtradas
df = df[df["I1"] == "1"].copy()

# Cabeçalhos Promovidos
df.columns = df.iloc[0]
df = df.drop(df.index[0])
df.reset_index(drop=True, inplace=True)
df.columns = df.columns.str.strip()


# Outras Colunas Removidas
colunas_para_deletar = [
    df.columns[0],
    df.columns[4],
    df.columns[5],
    df.columns[6],
    df.columns[8],
]
df.drop(columns=colunas_para_deletar, inplace=True, errors='ignore')  


# Extrair apenas o número, tratando floats e NaN
df["Cód/ Imagem"] = df["Cód/ Imagem"].apply(
    lambda x: re.search(r"\d+", str(x)).group() if re.search(r"\d+", str(x)) else pd.NA
)

df
# %% [markdown] 
### PISOS

# Consultas
pisos = df

# Linhas Filtradas
pisos = pisos[pisos["Qtd."].isnull()].copy()

# Coluna Condicional Adicionada
pisos["I1"] = None
pisos.loc[pisos[pisos.columns[2]].str.contains("TÉRREO|PAVIMENTO|SUBSOLO", na=False), "I1"] = "1"

pisos = pisos[pisos["I1"] == "1"].copy()

# Outras Colunas Removidas
colunas_para_deletar = [
    pisos.columns[0],
    pisos.columns[1],
    pisos.columns[3]
]
pisos.drop(columns=colunas_para_deletar, inplace=True, errors='ignore')

# Renomear Coluna 
pisos = pisos.rename(columns={pisos.columns[0]: "Pisos"})

pisos

# %% [markdown] 
### SALAS

# Consultas
salas = df

# Linhas Filtradas
salas = salas[salas["Qtd."].isnull()].copy()

# Outras Colunas Removidas
colunas_para_deletar = [
    salas.columns[0],
    salas.columns[1],
    salas.columns[3]
]
salas.drop(columns=colunas_para_deletar, inplace=True, errors='ignore')

# Índice Adicionado
salas["Index"] = range(1, len(salas) + 1)
salas = salas.rename(columns={salas.columns[0]: "Salas"})
salas
# %% [markdown] 
### MOVEIS

# Consultas
moveis = df

# Adiciona coluna Pisos e Salas
moveis["Pisos"] = None
moveis["Salas"] = None 

# Mesclar tabelas pisos e salas
moveis['Pisos'] = None
moveis.loc[moveis[moveis.columns[2]].str.contains("TÉRREO|PAVIMENTO|SUBSOLO", na=False), 'Pisos'] = moveis[moveis.columns[2]]
moveis['Pisos'] = moveis['Pisos'].ffill()

moveis['Salas'] = None
moveis.loc[moveis[moveis.columns[2]].notnull() & moveis['Qtd.'].isnull(), 'Salas'] = moveis[moveis.columns[2]]
moveis['Salas'] = moveis['Salas'].ffill()

# Linhas Filtradas
moveis = moveis[moveis[moveis.columns[1]].notna()].copy()
moveis

# %% [markdown] 
### TRELLO

# Consultas
trello = moveis

#  Adiciona coluna Dimensoes
trello["Dimensoes"] = trello["Produto"].apply(
    lambda txt: "; ".join(
        re.findall(r"\bDIM(?:ENSÕES?|ENSÃO|\.?)\b[^\n]*", txt, flags=re.IGNORECASE)
    )
)

# Adiciona coluna Primeira Linha
trello["Primeira Linha"] = trello["Produto"].apply(
    lambda txt: next((linha.strip() for linha in txt.splitlines() if linha.strip()), "")
)

# Selecionar colunas
"""
trello = trello[[
    "Pisos",
    "Salas",
    "Cód/ Imagem", 
    "Primeira Linha",
    "Dimensoes",
    "Qtd.",
    "Produto"
]].copy()
"""


trello.to_excel("../planilhas/trello.xlsx", index=False)

trello
# %%

from google.oauth2 import service_account
from googleapiclient.discovery import build
import polars as pl
import pandas as pd
from dotenv import load_dotenv
import os

dotenv_path = r"..\config\.env"
load_dotenv(dotenv_path)

SERVICE_ACCOUNT_FILE = r"..\config\credentials.json"

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']  # leitura + escrita


# Pega as variáveis
SPREADSHEET_ID = os.getenv("SHEETS-LISTA-ATIVIDADES")
SHEET_NAME = 'ADEMILSON'  


credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
service = build('sheets', 'v4', credentials=credentials)

result = service.spreadsheets().values().get(
    spreadsheetId=SPREADSHEET_ID,
    range=SHEET_NAME 
).execute()


values = result.get("values", [])
df = pl.from_pandas(pd.DataFrame(values))

df.to_pandas()

# %%
