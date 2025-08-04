import os
import logging
import requests
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import shutil





# Configuração do logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")



def obter_diretorio_download():
    """
    Cria a pasta 'Gestta - Tarefas' no diretório onde o script está, se não existir.
    Retorna o caminho completo da pasta sem apagar nenhum arquivo.
    """
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    diretorio_download = os.path.join(diretorio_atual, "Gestta - Tarefas")

    if not os.path.exists(diretorio_download):
        os.makedirs(diretorio_download)
        print(f" Diretório de download criado: {diretorio_download}")
    else:
        print(f" Diretório de download já existe: {diretorio_download}")

    return diretorio_download


# Diretório de download
DOWNLOAD_DIR = obter_diretorio_download()

        
        
