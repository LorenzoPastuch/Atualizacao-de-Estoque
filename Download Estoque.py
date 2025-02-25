from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import os
import time
import json

# Função para aguardar o download e renomear o arquivo
def wait_for_download_and_rename(download_dir, new_file_name, timeout=30):
    start_time = time.time()
    while True:
        # Verificar arquivos na pasta de downloads
        files = os.listdir(download_dir)
        if files:
            # Considerar o arquivo mais recente
            files = [f for f in files if f.endswith('.xls')]
            if files:
                latest_file = max([os.path.join(download_dir, f) for f in files], key=os.path.getmtime)
                # Renomear o arquivo baixado
                new_file_path = os.path.join(download_dir, f"{new_file_name}.xls")
                os.rename(latest_file, new_file_path)
                return
        # Verificar se o tempo de espera foi excedido
        elapsed_time = time.time() - start_time
        if elapsed_time > timeout:
            print("Tempo de espera excedido. O download não foi concluído a tempo.")
            break
        time.sleep(1)

# Lista de produtos para download
with open('SKU-Produtos.json') as produtos:
    produtos = json.load(produtos)

# Diretórios
download_dir = 'Downloads'
chromedriver_path = 'chromedriver.exe'
service = Service(executable_path=chromedriver_path)

def executar():
    try:
        # Configurar opções do Chrome
        chrome_options = Options()
        prefs = {
            "detach": True,
            "profile.default_content_settings.popups": 0,
            "download.default_directory": os.path.abspath(download_dir),
            "download.prompt_for_download": False,
            "safebrowsing.enabled": False,
            "download.directory_upgrade": True,
            "profile.default_content_setting_values.automatic_downloads": 1,
            "download_restrictions": 0,
            "profile.content_settings.exceptions.download.*.setting": 1,
            "profile.content_settings.exceptions.plugins.*.setting": 1,
            }
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--safebrowsing-disable-download-protection")
        chrome_options.add_argument("--safebrowsing-disable-extension-blacklist")
        chrome_options.add_argument("--allow-running-insecure-content")
        chrome_options.add_argument("--disable-features=InsecureDownloadWarnings")
        chrome_options.add_argument("--unsafely-treat-insecure-origin-as-secure=http://168.90.16.122:59478/")
        # Inicializar o driver do Chrome com opções configuradas
        navegador = webdriver.Chrome(service=service, options=chrome_options)

        # Acessar a página de login
        navegador.get('http://168.90.16.122:59478/One/login.jsf')
        navegador.find_element(By.ID, 'j_username').send_keys("lorenzo")
        navegador.find_element(By.ID, 'j_password').send_keys("supercopo3151")
        navegador.find_element(By.ID, 'cl_login').click()

        # Acessar a página de estoque
        navegador.get('http://168.90.16.122:59478/One/pages/estoque/estoqueList.jsf')
        navegador.find_element(By.ID, 'sbb_somentecestoque').click()
        navegador.find_element(By.ID, 'sbb_somentesemproducaoautomatica').click()

        # Baixar os arquivos de cada produto
        for i in produtos:
            arquivos = [os.path.splitext(arquivo)[0] for arquivo in os.listdir(download_dir)]

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
                wait_for_download_and_rename(download_dir, i, timeout=30)
                time.sleep(3)  # Esperar um pouco antes de iniciar o próximo download
                navegador.find_element(By.ID, 'dpc_Produtos:dpcIT_IDNdpc_Produtos').clear()
                time.sleep(1)
                print(f"Download de {i} concluído.")
        # Fechar o navegador
        navegador.quit()
    except Exception as e:
        print(f"Erro ocorrido: {e}. Reiniciando o processo...")
        navegador.quit()
        executar()

executar()