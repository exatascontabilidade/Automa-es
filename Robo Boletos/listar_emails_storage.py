from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import random
from selenium.webdriver.common.action_chains import ActionChains
import json
import os
import re


def random_sleep(min_seconds=1, max_seconds=2):
    """Aguarda um tempo aleatório para simular comportamento humano."""
    time.sleep(random.uniform(min_seconds, max_seconds))


def obter_diretorio_download():
    """Garante que a pasta 'temp' exista e retorna seu caminho."""
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    diretorio_download = os.path.join(diretorio_atual, "temp")

    if not os.path.exists(diretorio_download):
        os.makedirs(diretorio_download)

    return diretorio_download


def buscar_emails(navegador):
    """
    Busca e-mails no Gmail no mesmo navegador já logado.
    """
    filtro_busca = obter_datas()
    print(f"[INFO] - Buscando e-mails com o filtro: {filtro_busca}")

    wait = WebDriverWait(navegador, 20)
    search_box = wait.until(EC.presence_of_element_located((By.NAME, "q")))

    search_box.clear()
    search_box.send_keys(filtro_busca)
    random_sleep()
    search_box.send_keys(Keys.RETURN)
    
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "[role='main'] [role='row']")))


    random_sleep(2)  # Aguarda os resultados carregarem
    
    print("[INFO] - Resultados carregados com sucesso!")


def obter_diretorio_download():
    """Garante que a pasta 'temp' exista e retorna seu caminho."""
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    diretorio_download = os.path.join(diretorio_atual, "temp")

    if not os.path.exists(diretorio_download):
        os.makedirs(diretorio_download)

    return diretorio_download


def validar_data(mensagem):
    """
    Solicita ao usuário uma data no formato correto (AAAA/MM/DD).
    Repete a solicitação até que a entrada seja válida.
    """
    padrao_data = re.compile(r"^\d{4}/\d{2}/\d{2}$")  # Verifica formato AAAA/MM/DD
    while True:
        data = input(mensagem).strip()
        if padrao_data.match(data):  # Confirma que o formato está correto
            return data
        else:
            print("❌ Formato inválido! Digite a data no formato correto (AAAA/MM/DD).")


def obter_datas():
    """
    Pergunta ao usuário as datas para a busca.
    Retorna as datas formatadas para a pesquisa do Gmail e salva o mês/ano em um arquivo JSON.
    """
    data_inicio = validar_data("📅 Digite a data inicial (AAAA/MM/DD): ")
    data_fim = validar_data("📅 Digite a data final (AAAA/MM/DD): ")

    # Extrai apenas o mês e o ano da data inicial
    ano_mes = data_inicio[:7]  # Formato: "AAAA/MM"

    # Define o caminho do arquivo JSON onde os dados serão salvos
    diretorio_download = obter_diretorio_download()
    caminho_arquivo = os.path.join(diretorio_download, "mes_ano.json")

    # Dados a serem salvos
    dados = {
        "ano_mes": ano_mes
    }

    # Salva os dados no formato JSON
    with open(caminho_arquivo, "w", encoding="utf-8") as file:
        json.dump(dados, file, indent=4, ensure_ascii=False)

    print(f"✅ Mês e ano ({ano_mes}) armazenados em: {caminho_arquivo}")

    # Retorna o filtro para busca no Gmail
    filtro_busca = f"label:financeiroexatas@exatascontabilidade.com.br recibo de after:{data_inicio} before:{data_fim}"
    return filtro_busca


def verificar_fim_paginacao(navegador):
    """Verifica se já estamos na última página antes de iniciar a busca e se o botão de próxima página existe."""
    wait = WebDriverWait(navegador, 10)
    
    # 🔍 Verifica se o botão de próxima página está desativado na classe pai
    try:
        botao_proxima_pagina = wait.until(
            EC.presence_of_element_located((By.CLASS_NAME, "amJ"))
        )
        botao_pai = botao_proxima_pagina.find_element(By.XPATH, "ancestor::div[contains(@class, 'T-I')]")

        if botao_pai.get_attribute("aria-disabled") == "true":
            print("✅ Botão de próxima página está desativado. Última página alcançada.")
            return True
        else:
            print("➡️ Botão de próxima página encontrado e está ativo.")

    except Exception:
        print("🚫 Erro ao localizar o botão de próxima página. Considerando como última página.")
        return True

#------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Obtém o diretório de download (pasta "temp")
DOWNLOAD_DIR = obter_diretorio_download()

# Caminho do arquivo JSON para armazenar os e-mails encontrados na pasta "temp"
ARQUIVO_EMAILS = os.path.join(DOWNLOAD_DIR, "emails_encontrados.json")

# Lista global para armazenar os e-mails encontrados
emails_encontrados = []

def verificar_arquivo_existente():
    """Verifica se o arquivo de e-mails existe e deleta antes de iniciar uma nova listagem."""
    if os.path.exists(ARQUIVO_EMAILS):
        os.remove(ARQUIVO_EMAILS)
        print("🗑️ Arquivo de e-mails encontrado e deletado para nova listagem.")

def salvar_emails():
    """Salva a lista de e-mails encontrados em um arquivo JSON."""
    with open(ARQUIVO_EMAILS, "w", encoding="utf-8") as f:
        json.dump(emails_encontrados, f, ensure_ascii=False, indent=4)
        print(f"✅ E-mails salvos em: {ARQUIVO_EMAILS}")

def carregar_emails():
    """Carrega os e-mails armazenados do arquivo JSON."""
    global emails_encontrados
    if os.path.exists(ARQUIVO_EMAILS):
        with open(ARQUIVO_EMAILS, "r", encoding="utf-8") as f:
            emails_encontrados = json.load(f)
        print(f"📂 E-mails carregados de: {ARQUIVO_EMAILS}")

def listar_todos_emails(navegador):
    """Lista todos os e-mails paginando até o final e armazena os dados sem exibir no terminal."""
    verificar_arquivo_existente()
    wait = WebDriverWait(navegador, 15)

    while True:
        try:
            print("\n🔍 Buscando e-mails na página atual...")
            # Aguarda os e-mails carregarem
            email_elementos = wait.until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "[role='main'] [role='row'].zA"))
            )

            if email_elementos:
                print(f"✅ {len(email_elementos)} e-mail(s) encontrado(s) nesta página!\n")
                for email_elemento in email_elementos:
                    try:
                        remetente_elemento = email_elemento.find_element(By.CLASS_NAME, "yX")
                        remetente_texto = remetente_elemento.text if remetente_elemento.text else "Desconhecido"
                        remetente = re.sub(r"[\n\r].*|[\d]+$", "", remetente_texto).strip()
                        
                        if remetente != "EXATAS CONTABILIDADE":
                            continue
                        assunto_elemento = email_elemento.find_element(By.CLASS_NAME, "bog")
                        assunto = assunto_elemento.text if assunto_elemento.text else "Sem assunto"

                        data_elemento = email_elemento.find_element(By.CLASS_NAME, "xW")
                        data = data_elemento.text if data_elemento.text else "Data não disponível"

                        emails_encontrados.append({
                            "remetente": remetente,
                            "assunto": assunto,
                            "data": data
                        })
                    except Exception:
                        pass
            
            # Salva os e-mails após cada página carregada
            salvar_emails()

            # 🔍 Verifica se já estamos na última página antes de tentar avançar
            if verificar_fim_paginacao(navegador):
                print("✅ Última página alcançada. Finalizando listagem.")
                break

            # 🛑 Verifica se há uma próxima página de e-mails
            try:
                botao_proxima_pagina = wait.until(
                    EC.presence_of_element_located((By.CLASS_NAME, "amJ"))
                )
                
                if botao_proxima_pagina.is_displayed() and botao_proxima_pagina.is_enabled():
                    print("➡️ Indo para a próxima página...\n")
                    navegador.execute_script("arguments[0].scrollIntoView();", botao_proxima_pagina)
                    time.sleep(random.uniform(1, 2))
                    
                    action = ActionChains(navegador)
                    action.move_to_element(botao_proxima_pagina).click().perform()
                    time.sleep(random.uniform(2, 4))
                else:
                    print("✅ Todas as páginas foram percorridas.")
                    break
            
            except Exception:
                break

        except Exception:
            break
    print("✅ Listagem de e-mails concluída e armazenada com sucesso.")

def obter_emails_encontrados():
    """Retorna a lista de e-mails encontrados."""
    carregar_emails()
    return emails_encontrados



