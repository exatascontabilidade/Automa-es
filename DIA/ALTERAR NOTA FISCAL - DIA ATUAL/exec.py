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
from selenium.common.exceptions import StaleElementReferenceException

inscricao_municipal           = sys.argv[1]  # Ex: "275434569"
nome_empresa                  = sys.argv[2]  # Ex: "EMPRESA TESTE LTDA"
ETIQUETA                      = sys.argv[3]  # Ex: "10053414501604"
ICMSnovo                      = sys.argv[4]  # Ex: "300,00"
ICMSatual                     = sys.argv[5].replace(",", ".")  # Ex: "212,75" → "212.75"
Formaderecolhimentonovo       = sys.argv[6]  # Ex: "8"
Formaderecolhimentoatual      = sys.argv[7]  # Ex: "33"
ADIAR                         = sys.argv[8]  # Ex: "SIM" ou "NÃO"

# Configurações do Chrome
options = Options()
options.add_argument("--disable-gpu")
options.add_argument("--mute-audio")
options.add_argument("--no-sandbox")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("useAutomationExtension", False)
options.add_experimental_option("excludeSwitches", ["enable-automation"])
#options.add_argument("--headless=new")

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
     
    solicitar_xml = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Alterar Nota Fiscal - DIA Atual' )]")))   
    solicitar_xml.click()
    print("✅ 'Alterar Nota Fiscal - DIA Atual' acessada com sucesso!")
    
except Exception as e:
    print(f"❌ Erro ao localizar e clicar na opção: {e}")
    
#-----------------------------------------------------------------------------------SELEÇÃO EMPRESA------------------------------------------------------------------------------------------------------------------------------------------
try:
    select_element = wait.until(EC.presence_of_element_located((By.ID, "cdPessoaContribuinte")))
    select = Select(select_element)

    valores = [opt.get_attribute("value") for opt in select.options]

    if str(inscricao_municipal) in valores:
        select.select_by_value(str(inscricao_municipal))
        print(f"✅ Empresa '{inscricao_municipal}' selecionada!")
        botao_ok = wait.until(EC.element_to_be_clickable((By.ID, "okButton")))
        botao_ok.click()
        print("✅ 'Botão Ok")
    else:
        print(f"❌ Empresa com inscrição '{inscricao_municipal}' não encontrada.")
        navegador.quit()
        sys.exit(1)
except Exception as e:
    print(f"❌ Erro ao selecionar empresa: {e}")
    navegador.quit()
    sys.exit(1)
#---------------------------------------------------------------------------------- ALTERAÇÃO DE NOTAS FISCAIS---------------------------------------------------------------------------------------------------------------------------------------


#-------- Selecionando Campos Vazios 
try:
    select_exibicao = wait.until(EC.presence_of_element_located((By.ID, "exibicaoConsulta")))
    select = Select(select_exibicao)
    select.select_by_value("")  # valor vazio
    print("✅ Campo 'Exibição da Consulta' selecionado com valor vazio!")
except Exception as e:
    print(f"❌ Erro ao selecionar o campo 'Exibição da Consulta': {e}")
    navegador.quit()
    sys.exit(1) 
try:
    select_exibicao = wait.until(EC.presence_of_element_located((By.ID, "AnoReferencia")))
    select = Select(select_exibicao)
    select.select_by_value("")  # valor vazio
    print("✅ Campo 'AnoReferencia' selecionado com valor vazio!")
except Exception as e:
    print(f"❌ Erro ao selecionar o campo 'AnoReferencia': {e}")
    navegador.quit()
    sys.exit(1)  
try:
    select_exibicao = wait.until(EC.presence_of_element_located((By.ID, "MesReferencia")))
    select = Select(select_exibicao)
    select.select_by_value("")  # valor vazio
    print("✅ Campo 'MesReferencia' selecionado com valor vazio!")
except Exception as e:
    print(f"❌ Erro ao selecionar o campo 'MesReferencia': {e}")
    navegador.quit()
    sys.exit(1)  
    
#--------- Preenchendo a Etiqueta e Consultando
try:
    ETIQUETA = sys.argv[3]  # Captura o valor da etiqueta da linha de comando

    input_etiqueta = wait.until(EC.presence_of_element_located((By.ID, "ETQ_nrEtiqueta")))
    input_etiqueta.clear()  # limpa o valor atual (evita concatenação)
    input_etiqueta.send_keys(ETIQUETA)

    print(f"✅ Campo 'Etiqueta' preenchido com: {ETIQUETA}")
    
    botao_ok = wait.until(EC.element_to_be_clickable((By.ID, "okButton")))
    botao_ok.click()
    
except Exception as e:
    print(f"❌ Erro ao preencher o campo 'Etiqueta': {e}")
    navegador.quit()
    sys.exit(1)    

#------------------------- Acessando a Etiqueta Correspondente 

try:
    tabela = wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
    linha_encontrada = False

    linhas = tabela.find_elements(By.XPATH, ".//tr[contains(@class, 'trTableImpar') or contains(@class, 'trTablePar')]")
    for i in range(len(linhas)):
        try:
            # Reobtem a tabela e as linhas para garantir que são atuais
            tabela = wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            linhas = tabela.find_elements(By.XPATH, ".//tr[contains(@class, 'trTableImpar') or contains(@class, 'trTablePar')]")
            linha = linhas[i]
            colunas = linha.find_elements(By.TAG_NAME, "td")
            if len(colunas) < 10:
                continue

            etiqueta_valor = colunas[0].text.strip()
            icms_valor = colunas[2].text.strip()
            forma_recolhimento_valor = colunas[9].text.strip()

            etiqueta_valor_limpa = etiqueta_valor.lstrip("0")
            etiqueta_recebida_limpa = ETIQUETA.lstrip("0")
            icms_valor_normalizado = icms_valor.replace(",", ".").strip()
            icms_recebido_normalizado = ICMSatual.replace(",", ".").strip()
            forma_pagina = forma_recolhimento_valor.strip().upper()
            forma_recebida = Formaderecolhimentoatual.strip().upper()

            if (
                etiqueta_valor_limpa == etiqueta_recebida_limpa and
                icms_valor_normalizado == icms_recebido_normalizado and
                forma_pagina == forma_recebida
            ):
                # Aguarda o link atual e clica nele
                link = WebDriverWait(linha, 10).until(
                    EC.element_to_be_clickable((By.TAG_NAME, "a"))
                )
                navegador.execute_script("arguments[0].scrollIntoView(true);", link)
                link.click()
                linha_encontrada = True
                print(f"\n✅ Linha correspondente encontrada e clicada: Etiqueta {ETIQUETA}")
                break

        except StaleElementReferenceException:
            print("⚠️ Elemento ficou obsoleto (stale), tentando novamente...")

        except Exception as e:
            print(f"⚠️ Erro ao processar linha {i}: {e}")

    if not linha_encontrada:
        print("\n❌ Nenhuma linha correspondente encontrada com os valores informados.")
        navegador.quit()
        sys.exit(1)

except Exception as e:
    print(f"\n❌ Erro geral: {e}")
    navegador.quit()
    sys.exit(1)
    

#------------------------------------------- Preenchendo Nota com novos valores 

try:
    ICMSnovo = sys.argv[4].strip().replace(".", ",")
    Formaderecolhimentonovo = sys.argv[6].strip()

    preenchimento_icms_ok = False
    preenchimento_forma_ok = False

    # Preenchimento do campo ICMS
    if ICMSnovo:
        try:
            input_icms = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "inputNumber")))
            input_icms.clear()
            input_icms.send_keys(ICMSnovo)
            print(f"✅ Campo 'ICMS' preenchido com: {ICMSnovo}")
            preenchimento_icms_ok = True
        except Exception as e:
            print(f"⚠️ Erro ao preencher ICMS: {e}")
    else:
        print("⚠️ ICMSnovo não fornecido. Pulando preenchimento do campo ICMS.")

    # Seleção da forma de recolhimento
    if Formaderecolhimentonovo:
        try:
            select_element = wait.until(EC.presence_of_element_located((By.ID, "FRC_cdFormaRecolhimento")))
            select = Select(select_element)
            textos_disponiveis = [opt.text.strip() for opt in select.options]

            if Formaderecolhimentonovo in textos_disponiveis:
                select.select_by_visible_text(Formaderecolhimentonovo)
                print(f"✅ Forma de recolhimento '{Formaderecolhimentonovo}' selecionada!")
                preenchimento_forma_ok = True
            else:
                print(f"❌ Forma de recolhimento '{Formaderecolhimentonovo}' não encontrada entre os valores disponíveis.")
        except Exception as e:
            print(f"⚠️ Erro ao selecionar forma de recolhimento: {e}")
    else:
        print("⚠️ Formaderecolhimentonovo não fornecido. Pulando seleção.")

    # Clique no botão OK apenas se ambos os preenchimentos foram realizados
    if preenchimento_icms_ok and preenchimento_forma_ok:
        try:
            botao_ok = wait.until(EC.element_to_be_clickable((By.ID, "okButton")))
            botao_ok.click()
            print("🔘 Botão 'OK' clicado com sucesso.")
        except Exception as e:
            print(f"⚠️ Não foi possível clicar no botão 'OK': {e}")
    else:
        print("⏭️ Botão 'OK' não clicado pois nem todos os campos foram preenchidos.")

except Exception as e:
    print(f"\n❌ Erro inesperado durante o preenchimento: {e}")
    navegador.quit()
    sys.exit(1)
    