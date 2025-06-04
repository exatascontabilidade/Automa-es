import glob
import sys
import time
import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from datetime import datetime
import os

inscricao_municipal = sys.argv[1]
nome_empresa = sys.argv[2]
mMes = sys.argv[3]
mANO = sys.argv[4]
formato_arquivo = sys.argv[5]

# Configurações do Chrome
options = Options()
options.add_argument("--disable-gpu")
options.add_argument("--mute-audio")
options.add_argument("--no-sandbox")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("useAutomationExtension", False)
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_argument("--headless=new")

# Caminho absoluto da pasta "temp" no mesmo diretório do script
download_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp")
os.makedirs(download_dir, exist_ok=True)

# ⬇️ Preferências para baixar PDFs automaticamente ⬇️
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True  # 👈 Faz o Chrome BAIXAR o PDF em vez de abrir
}
options.add_experimental_option("prefs", prefs)

# Inicialização do navegador
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico, options=options)
driver = navegador
navegador.get("https://www.sefaz.se.gov.br/SitePages/acesso_usuario.aspx")
wait = WebDriverWait(navegador, 3)




#--------------------------------------------------------------------------------LOGIN NA PAGINA--------------------------------------------------------------------------------------------------------------------------------------------------
try:
    try:
        accept_button = wait.until(EC.element_to_be_clickable((By.ID, 'accept-button')))
        accept_button.click()
    except:
        print("Botão 'Aceitar' não encontrado. Continuando sem clicar.")
    time.sleep(1)
    iframes = navegador.find_elements(By.TAG_NAME, "iframe")
    navegador.switch_to.frame(iframes[0])
    dropdown = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'acessoRapido')))
    dropdown.click()
    time.sleep(1)
    option_contabilista = wait.until(EC.element_to_be_clickable((By.XPATH, "//option[@value='https://security.sefaz.se.gov.br/internet/portal/contabilista/atoAcessoContabilista.jsp']")))
    option_contabilista.click()
    time.sleep(1)
    body = navegador.find_element(By.TAG_NAME, "body")
    body.click()
    time.sleep(1)
    try:
        iframe_login = wait.until(EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'atoAcessoContribuinte.jsp')]")))
        navegador.switch_to.frame(iframe_login)
    except:
        print("❌ Erro: O iframe do login NÃO foi encontrado!")
        raise Exception("Iframe do login não localizado!")
    try:
        tabela_login = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tabelaVerde")))
    except:
        raise Exception("Tabela de login não localizada!")
    try:
        usuario = tabela_login.find_element(By.NAME, "UserName")
        senha = tabela_login.find_element(By.NAME, "Password")
        botao_login = tabela_login.find_element(By.NAME, "submit")  # Botão "OK"
    except:
        raise Exception("Campos de login não localizados!")
    usuario.click()
    usuario.send_keys("SE007829")
    senha.click()
    senha.send_keys("Exatas2024@")
    botao_login.click()
    print("🎉 Login realizado com sucesso!")
except Exception as e:
    print(f"Erro ao localizar os campos de login: {e}")






def scroll_ate_fim_pagina(navegador, timeout=3):
    scroll_pause = 1
    altura_anterior = navegador.execute_script("return document.body.scrollHeight")

    while True:
        navegador.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(scroll_pause)

        nova_altura = navegador.execute_script("return document.body.scrollHeight")
        if nova_altura == altura_anterior:
            break
        altura_anterior = nova_altura

    try:
        # Espera por qualquer novo conteúdo no final da página (ajuste o seletor se quiser algo específico)
        WebDriverWait(navegador, timeout).until(
            EC.presence_of_element_located((By.TAG_NAME, "footer"))
        )
    except:
        pass  # Se não tiver footer, apenas continue

    print("✅ Rolagem concluída.")

scroll_ate_fim_pagina(navegador)















#-----------------------------------------------------------------------------------SOLICITAR XML-------------------------------------------------------------------------------------------------------------------------------------------------
try:
    elementos = navegador.find_elements(By.TAG_NAME, "a")
    if elementos:
        elementos[0].click()  
    else:
        print("❌ Nenhum link encontrado para clicar.")
except Exception as e:
    print(f"❌ Erro ao clicar em um campo aleatório: {e}")
try:
    menu_nfe = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'DIA')]")))

    menu_nfe.click()
    print("✅ 'DIA' acessada com sucesso!")
    time.sleep(5)  
    solicitar_xml = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Demonstrativo ICMS' )]")))   
    solicitar_xml.click()
    print("✅ 'Demonstrativo' acessada com sucesso!")
    
except Exception as e:
    print(f"❌ Erro ao localizar e clicar na opção: {e}")
#-----------------------------------------------------------------------------------SELEÇÃO EMPRESA------------------------------------------------------------------------------------------------------------------------------------------
try:
    select_empresas = wait.until(EC.presence_of_element_located((By.ID, "cdPessoaLookup")))
    select = Select(select_empresas)
    select.select_by_value(str(inscricao_municipal))
    print(f"✅ Empresa '{inscricao_municipal}' selecionada!")
except Exception as e:
    print(f"❌ Erro ao selecionar empresa: {e}")
    navegador.quit()
    sys.exit(1)
    

#-------------------------------------- SELEÇÃO De ANO E MES --------------------------------------#
try:
    tipo_arquivo = mMes.strip().capitalize()  # Corrige capitalização para bater com o texto do dropdown

    # Localiza o <select> de meses pelo ID
    select_mes_element = wait.until(EC.presence_of_element_located((By.ID, "nrMesDia")))
    select_mes = Select(select_mes_element)

    # Lista os nomes dos meses disponíveis (textos das opções)
    opcoes_disponiveis = [op.text.strip() for op in select_mes.options if op.text.strip()]

    if tipo_arquivo in opcoes_disponiveis:
        # Seleciona o mês pelo texto visível
        select_mes.select_by_visible_text(tipo_arquivo)
        print(f"✅ Mês '{tipo_arquivo}' selecionado com sucesso!")
    else:
        print(f"❌ Mês '{tipo_arquivo}' não encontrado. Opções disponíveis: {opcoes_disponiveis}")
        sys.exit(1)
        
    # --- Seleção do ano ---
    data = mANO.strip()  # Ex: '2025'

    select_ano_element = wait.until(EC.presence_of_element_located((By.ID, "nrAno")))
    select_ano = Select(select_ano_element)

    opcoes_anos = [op.text.strip() for op in select_ano.options if op.text.strip()]
    
    if data in opcoes_anos:
        select_ano.select_by_visible_text(data)
        print(f"✅ Ano '{data}' selecionado com sucesso!")
    else:
        print(f"❌ Ano '{data}' não encontrado. Opções disponíveis: {opcoes_anos}")
        sys.exit(1)

    formato_arquivo = formato_arquivo.strip().upper()  # Ex: 'PDF' ou 'EXCEL'

    select_formato_element = wait.until(EC.presence_of_element_located((By.ID, "tpFormato")))
    select_formato = Select(select_formato_element)

    opcoes_formatos = [op.text.strip().upper() for op in select_formato.options if op.text.strip()]
    
    if formato_arquivo in opcoes_formatos:
        select_formato.select_by_visible_text(formato_arquivo)
        print(f"✅ Formato '{formato_arquivo}' selecionado com sucesso!")
    else:
        print(f"❌ Formato '{formato_arquivo}' não encontrado. Opções disponíveis: {opcoes_formatos}")
        sys.exit(1)

    # --- Botão OK ---
    botao_ok = wait.until(EC.element_to_be_clickable((By.ID, "okButton")))
    botao_ok.click()



except Exception as e:
    print(f"❌ Erro ao selecionar mês ou ano: {e}")
    sys.exit(1)

#---------------------------MENSAGEM DE ERRO ------------------------------------------------
try:
    # Aguarda e verifica se alguma mensagem de erro aparece na página
    mensagem_erro_element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "fontMessageError")))
    mensagem_erro = mensagem_erro_element.text.strip()
    if mensagem_erro:
        print(f"⚠️ Mensagem de erro detectada: {mensagem_erro}")
except:
    print("✅ Nenhuma mensagem de erro detectada.")


        
        