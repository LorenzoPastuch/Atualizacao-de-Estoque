import os.path
import pandas as pd
import datetime
import json
from time import sleep
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

def find_cell_row(values, search_value):
    for row_idx, row in enumerate(values):
        if search_value in row:
            return row_idx + 1  # +1 para converter índice de 0-base para 1-base
    return None


# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = "1Sa9E7b8-e7jqRGcMo1uDQTqhW3mTkqBO_OSs-nTb4Nc"

creds = None
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists("token.json"):
  creds = Credentials.from_authorized_user_file("token.json", SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
  if creds and creds.expired and creds.refresh_token:
    creds.refresh(Request())
  else:
    flow = InstalledAppFlow.from_client_secrets_file(
        "credentials.json", SCOPES
    )
    creds = flow.run_local_server(port=0)
  # Save the credentials for the next run
  with open("token.json", "w") as token:
    token.write(creds.to_json())

service = build("sheets", "v4", credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()
planilha = (
    sheet.values()
    .get(spreadsheetId=SPREADSHEET_ID, range="PCP PRODUTOS LISOS!A:A")
    .execute()
)
valores_produtos = planilha.get("values", [])

diretorio = 'Downloads'

downloads = os.listdir(diretorio)

with open('SKU-Produtos.json') as produtos:
  produtos = json.load(produtos)


data = [[datetime.date.today().strftime('%d-%m-%Y')]]

for download in downloads:
  sleep(1)
  ok = False

  sku = os.path.splitext(download)[0]
  excel = pd.read_excel(os.path.join(diretorio, download))
  excel = excel.fillna("")

  codigo = excel["Código"]
  codigo = [[valor] for valor in codigo]

  prod = excel["Produto"]
  prod = [[valor] for valor in prod]
  
  atributos = excel["Atributos"]
  atributos = [[valor] for valor in atributos]

  estoque = excel["Estoque"]
  estoque = [[valor] for valor in estoque]

  reservado = excel["Reservado"]
  reservado = [[valor] for valor in reservado]
  
  disponivel = excel["Disponível"]
  disponivel = [[valor] for valor in disponivel]

  minimo = excel["Mínimo"]
  minimo = [[valor] for valor in minimo]

  batch_update_body = {
    'valueInputOption':'USER_ENTERED',
    'data':[
      {'range': f"PCP PRODUTOS LISOS!A{find_cell_row(valores_produtos, produtos[sku])+2}", 'values': codigo},
      {'range': f"PCP PRODUTOS LISOS!B{find_cell_row(valores_produtos, produtos[sku])+2}", 'values': prod},
      {'range': f"PCP PRODUTOS LISOS!D{find_cell_row(valores_produtos, produtos[sku])+2}", 'values': atributos},
      {'range': f"PCP PRODUTOS LISOS!F{find_cell_row(valores_produtos, produtos[sku])+2}", 'values': estoque},
      {'range': f"PCP PRODUTOS LISOS!G{find_cell_row(valores_produtos, produtos[sku])+2}", 'values': reservado},
      {'range': f"PCP PRODUTOS LISOS!H{find_cell_row(valores_produtos, produtos[sku])+2}", 'values': disponivel},
      {'range': f"PCP PRODUTOS LISOS!I{find_cell_row(valores_produtos, produtos[sku])+2}", 'values': minimo}
    ]
  }
  
  for produto in valores_produtos:
    if produto == [f'{produtos[sku]}'] and len(prod)<=60:
      sheet.values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body=batch_update_body
      ).execute()

  
      print(f"{produtos[sku]} atualizado com sucesso.")
      ok = True

  if not ok:
    print(f"{sku}: Produto não encontrado/arquivo com erro! Apague o arquivo e refaça o download.")
        
sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                            range=f"PCP PRODUTOS LISOS!I14",
                            valueInputOption="RAW",
                            body={"values": data}).execute()