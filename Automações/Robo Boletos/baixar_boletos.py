import json
import time
import random
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from entrar_omie import baixar_boletos_atrasados, trocar_para_nova_janela


def obter_diretorio_download():
    """Garante que a pasta 'temp' exista e retorna seu caminho."""
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    diretorio_download = os.path.join(diretorio_atual, "temp")

    if not os.path.exists(diretorio_download):
        os.makedirs(diretorio_download)

    return diretorio_download


# Obtém o diretório de download (pasta "temp")
DOWNLOAD_DIR = obter_diretorio_download()

# Caminho do arquivo JSON para armazenar os e-mails encontrados na pasta "temp"
ARQUIVO_EMAILS = os.path.join(DOWNLOAD_DIR, "emails_encontrados.json")


def carregar_emails():
    """Carrega a lista de e-mails armazenados."""
    if not os.path.exists(ARQUIVO_EMAILS):
        print("❌ Nenhum e-mail armazenado para processar!")
        return []
    
    with open(ARQUIVO_EMAILS, "r", encoding="utf-8") as f:
        return json.load(f)

def carregar_navegador(navegador):
    """Aguarda o carregamento da página do Gmail antes de interagir com elementos."""
    try:
        WebDriverWait(navegador, 15).until(
            EC.presence_of_element_located((By.NAME, "q"))
        )
    except:
        print("⚠️ Campo de pesquisa do Gmail não encontrado! Listando elementos...")
        elementos = navegador.find_elements(By.XPATH, "//*")
        for elem in elementos[:20]:  # Lista os 20 primeiros elementos encontrados
            print(f"➡️ {elem.tag_name} | Classe: {elem.get_attribute('class')}")
        navegador.quit()

def fechar_aba_atual(navegador):
    """Fecha a aba atual e retorna para a aba anterior."""
    if len(navegador.window_handles) > 1:
        navegador.close()
        navegador.switch_to.window(navegador.window_handles[0])
        print("✅ Aba do Omie fechada, voltando ao Gmail.")
        time.sleep(2)  # Pequeno tempo para garantir que a aba foi fechada antes de processar o próximo

def baixar_boletos(navegador):
    """Processa os e-mails armazenados **um de cada vez** para evitar sobrecarga."""
    carregar_navegador(navegador)
    
    emails = carregar_emails()
    if not emails:
        return

    wait = WebDriverWait(navegador, 15)

    for email in emails:
        remetente = email["remetente"]
        assunto = email["assunto"]

        print(f"🔍 Pesquisando e-mail de {remetente} sobre '{assunto}'")

        try:
            search_box = WebDriverWait(navegador, 15).until(
                EC.presence_of_element_located((By.NAME, "q"))
            )
        except:
            print("⚠️ Campo de pesquisa não encontrado! Pulando e-mail...")
            continue  # Pula para o próximo e-mail se não encontrar a caixa de pesquisa

        # Pesquisar pelo e-mail específico no Gmail
        search_box.clear()
        search_box.send_keys(f"from:{remetente} {assunto}")
        search_box.send_keys(Keys.RETURN)
        time.sleep(random.uniform(2, 3))

        try:
            # Aguardar a exibição do primeiro e-mail na lista
            email_item = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "[role='main'] [role='row'].zA")))
            email_item.click()
            time.sleep(random.uniform(2, 4))

            # Tentar encontrar o botão "Visualizar o Documento no Portal Omie"
            botao_omie = wait.until(
                EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Visualizar o Documento no Portal Omie')]"))
            )
            navegador.execute_script("arguments[0].scrollIntoView();", botao_omie)
            time.sleep(random.uniform(1, 2))

            action = ActionChains(navegador)
            action.move_to_element(botao_omie).click().perform()
            time.sleep(random.uniform(3, 5))

            print(f"✅ Botão do Omie clicado! Abrindo o Portal Omie...")

            # 🔄 **Troca para a nova aba antes de tentar baixar o documento**
            trocar_para_nova_janela(navegador)

            # 🟢 **Chama a função de download do documento já na nova aba**
            baixar_boletos_atrasados(navegador)
            

        except Exception as e:
            print(f"⚠️ Erro ao tentar clicar no botão do Omie: {e}")

        print("⏳ Aguardando antes de processar o próximo e-mail...")
        time.sleep(random.uniform(2, 5))  # Pequeno intervalo para evitar sobrecarga
        
    print("✅ Todos os e-mails processados")