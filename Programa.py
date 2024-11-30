from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import os
import time
import json
import tkinter as tk
from tkinter.filedialog import askdirectory, askopenfilename
from pathlib import Path
import os.path
import pandas as pd
import datetime
import json
from time import sleep
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

 # If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = "1Sa9E7b8-e7jqRGcMo1uDQTqhW3mTkqBO_OSs-nTb4Nc"

creds = None

# Lista de produtos para download
with open('SKU-Produtos.json') as produtos:
    produtos = json.load(produtos)

chromedriver_path = 'chromedriver.exe'
service = Service(executable_path=chromedriver_path)

 # Configurar opções do Chrome
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
prefs = {
    "profile.default_content_settings.popups": 0,
    "safebrowsing.enabled": False,
    "profile.default_content_setting_values.automatic_downloads": 1,
    "download_restrictions": 0,
    "profile.content_settings.exceptions.download.*.setting": 1,
    "profile.content_settings.exceptions.plugins.*.setting": 1,
    }
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("--safebrowsing-disable-download-protection")
chrome_options.add_argument("--safebrowsing-disable-extension-blacklist")


class Diretorio:

    def __init__(self):
        self.diretorio_download = tk.StringVar()

    def selecionar_diretorio(self):
        diretorio = askdirectory(title="Selecionar diretório")
        self.diretorio_download.set(diretorio)
        self.diretorio_download = Path(self.diretorio_download.get())
        if diretorio:
            texto["text"] = diretorio


class Download:

    def __init__(self, instancia_download):
        self.instancia_download = instancia_download
    
    def executar(self):
        # Função para aguardar o download e renomear o arquivo
        def wait_for_download_and_rename(diretorio_download, new_file_name, timeout=30):
            start_time = time.time()
            while True:
                # Verificar arquivos na pasta de downloads
                files = os.listdir(diretorio_download)
                if files:
                    # Considerar o arquivo mais recente
                    files = [f for f in files if f.endswith('.xls')]
                    if files:
                        latest_file = max([os.path.join(diretorio_download, f) for f in files], key=os.path.getmtime)
                        # Renomear o arquivo baixado
                        new_file_path = os.path.join(diretorio_download, f"{new_file_name}.xls")
                        os.rename(latest_file, new_file_path)
                        return
                # Verificar se o tempo de espera foi excedido
                elapsed_time = time.time() - start_time
                if elapsed_time > timeout:
                    print("Tempo de espera excedido. O download não foi concluído a tempo.")
                    break
                time.sleep(1)

        try:
            # Inicializar o driver do Chrome com opções configuradas
            navegador = webdriver.Chrome(service=service, options=chrome_options)
            navegador.execute_cdp_cmd('Page.setDownloadBehavior', {
                'behavior': 'allow',
                'downloadPath': str(self.instancia_download.diretorio_download) 
            })

            # Acessar a página de login
            navegador.get('http://168.90.16.122:59478/One/login.jsf')
            navegador.find_element(By.ID, 'j_username').send_keys("FRANCARMO")
            navegador.find_element(By.ID, 'j_password').send_keys("lorenzo3151")
            navegador.find_element(By.ID, 'cl_login').click()

            # Acessar a página de estoque
            navegador.get('http://168.90.16.122:59478/One/pages/estoque/estoqueList.jsf')
            navegador.find_element(By.ID, 'sbb_somentecestoque').click()
            navegador.find_element(By.ID, 'sbb_somentesemproducaoautomatica').click()

            # Baixar os arquivos de cada produto
            for i in produtos:
                arquivos = [os.path.splitext(arquivo)[0] for arquivo in os.listdir(self.instancia_download.diretorio_download)]

                if i not in arquivos:
                    navegador.find_element(By.ID, 'dpc_Produtos:dpcIT_IDNdpc_Produtos').send_keys(f"{i}")
                    time.sleep(1)
                    navegador.find_element(By.ID, 'cl_filtro').click()
                    time.sleep(1)
                    navegador.find_element(By.ID, 'cl_filtro').click()
                    time.sleep(1)
                    navegador.find_element(By.XPATH, '//*[@id="f_filtro"]/legend').click()
                    time.sleep(2)
                    navegador.find_element(By.ID, 'cl_exportar').click()
                    time.sleep(3)  # Esperar um pouco para garantir que o download foi iniciado
                    wait_for_download_and_rename(self.instancia_download.diretorio_download, i, timeout=30)
                    time.sleep(3)  # Esperar um pouco antes de iniciar o próximo download
                    navegador.find_element(By.ID, 'dpc_Produtos:dpcIT_IDNdpc_Produtos').clear()
                    time.sleep(1)
                    print(f"Download de {i} concluído.")
            # Fechar o navegador
            navegador.quit()
        except Exception as e:
            print(f"Erro ocorrido: {e}.")
            navegador.quit()
            

class Atualizar:

    def __init__(self, instancia_download):
        self.instancia_download = instancia_download
        

    def atualizar(self):
        
        def find_cell_row(values, search_value):
            for row_idx, row in enumerate(values):
                if search_value in row:
                    return row_idx + 1  # +1 para converter índice de 0-base para 1-base
        
        diretorio = self.instancia_download.diretorio_download

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

        downloads = os.listdir(diretorio)

        data = [[datetime.date.today().strftime('%d-%m-%Y')]]

        for download in downloads:
            sleep(1)
            ok = False

            sku = os.path.splitext(download)[0]
            excel = pd.read_excel(os.path.join(diretorio, download))

            codigo = excel["Código"]
            codigo = [[valor] for valor in codigo]

            prod = excel["Produto"]
            prod = [[valor] for valor in prod]

            atributos = excel["Atributos"]
            atributos = [[valor] for valor in atributos]

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
                {'range': f"PCP PRODUTOS LISOS!F{find_cell_row(valores_produtos, produtos[sku])+2}", 'values': disponivel},
                {'range': f"PCP PRODUTOS LISOS!G{find_cell_row(valores_produtos, produtos[sku])+2}", 'values': minimo}
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

janela =tk.Tk()

diretorio = Diretorio()


janela.title("Atualização de Estoque")
janela.geometry("600x525+100+100")

titulo = tk.Label(
    text="Atualização de Estoque PCP Produtos Lisos",
    width=40,
    height=10
)
titulo.grid(
    row=0,
    column=0,
    columnspan=2,
    sticky="NSEW",
    padx=10,
    pady=10
)

texto = tk.Label(
    text="Nenhum diretório de download selecionado",
    width=40,
)
texto.grid(
    row=1,
    column=1,
    columnspan=1,
    sticky="NSEW",
    padx=10,
    pady=10
)

selectdrt = tk.Button(
    text="Selecionar",
    command=diretorio.selecionar_diretorio
)
selectdrt.grid(
    row=1,
    column=0,
    columnspan=1,
    sticky="NSEW",
    padx=10,
    pady=10
)
download = Download(diretorio)
atualizar= Atualizar(diretorio)
download = tk.Button(
    text="Fazer download",
    command=download.executar
)
download.grid(
    row=2,
    column=0,
    columnspan=1,
    sticky="NSEW",
    padx=10,
    pady=10
)

atualizar = tk.Button(
    text="Atualizar planilha",
    command=atualizar.atualizar
)
atualizar.grid(
    row=3,
    column=0,
    columnspan=1,
    sticky="NSEW",
    padx=10,
    pady=10
)

janela.mainloop()

