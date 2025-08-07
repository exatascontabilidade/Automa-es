import requests
import base64
import os
import time
import datetime
import re
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# Carrega as variáveis de ambiente do arquivo .env
load_dotenv()

# --- CONFIGURAÇÕES GLOBAIS ---
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
REPO = "exatascontabilidade/automacao-remota"
ARQUIVO = "comando.json"
API_URL = f"https://api.github.com/repos/{REPO}/contents/{ARQUIVO}"
INTERVALO_VERIFICACAO_SEGUNDOS = 10

HEADERS = {
    "Authorization": f"Bearer {GITHUB_TOKEN}",
    "Accept": "application/vnd.github.v3+json"
}

# --- FUNÇÕES AUXILIARES ---

def extrair_id_video(url: str) -> str | None:
    """Extrai o ID do vídeo de uma URL do YouTube."""
    padroes = [
        r"(?:v=|\/)([0-9A-Za-z_-]{11}).*",
        r"youtu\.be\/([0-9A-Za-z_-]{11}).*",
        r"youtube\.com\/embed\/([0-9A-Za-z_-]{11}).*"
    ]
    for padrao in padroes:
        match = re.search(padrao, url)
        if match:
            return match.group(1)
    return None

# --- FUNÇÃO PRINCIPAL DA AUTOMAÇÃO (LÓGICA DO NAVEGADOR) ---

def executar_tarefa_navegador(url_do_video: str):
    """
    Automatiza a navegação em modo 'disfarçado':
    1. Pesquisa por 'Mercado Livre' no Google.
    2. Acessa o site do Mercado Livre.
    3. Abre o vídeo do YouTube em modo embed.
    """
    print("🚀 Iniciando automação do navegador (em modo 'disfarçado')...")
    driver = None
    try:
        options = Options()

        # --- PARÂMETROS PARA 'DISFARÇAR' A AUTOMAÇÃO ---
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_experimental_option("detach", True) # Mantém o navegador aberto no final
        options.add_argument("--disable-infobars")
        options.add_argument("--start-maximized")

        # Gerenciador automático do WebDriver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)

        print("   1. Abrindo o Google...")
        driver.get("https://www.google.com")
        time.sleep(1)

        # Tenta aceitar os cookies se o botão aparecer
        try:
            botao_cookies = driver.find_element(By.XPATH, "//button[.//div[contains(text(), 'Aceitar tudo')]]")
            botao_cookies.click()
            time.sleep(1)
        except:
            print("      (Não foi necessário aceitar cookies)")

        print("   2. Pesquisando por 'Mercado Livre'...")
        barra_pesquisa = driver.find_element(By.NAME, "q")
        barra_pesquisa.send_keys("Mercado Livre")
        barra_pesquisa.send_keys(Keys.RETURN)
        time.sleep(2)

        print("   3. Acessando o site do Mercado Livre...")
        link_mercado_livre = driver.find_element(By.XPATH, "//a[contains(@href, 'mercadolivre.com.br')]")
        link_mercado_livre.click()
        time.sleep(10)

        print("   4. Abrindo o vídeo do YouTube...")
        video_id = extrair_id_video(url_do_video)
        if not video_id:
            print("      URL do YouTube inválida.")
            return

        url_embed = f"https://www.youtube.com/embed/{video_id}?autoplay=1"
        driver.execute_script(f"window.open('{url_embed}', '_blank');")
        print(f"   ✅ Vídeo aberto em uma nova aba: {url_embed}")
        
        print("\n🎉 Automação finalizada com sucesso!")

    except Exception as e:
        print(f"❌ Ocorreu um erro durante a automação do navegador: {e}")

# --- LOOP PRINCIPAL DO ROBÔ (OUVINTE) ---

def iniciar_robo_ouvinte():
    print(">>> Robô iniciado. Aguardando comandos... (Pressione Ctrl+C para parar) <<<")
    while True:
        try:
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            print(f"[{timestamp}] Verificando comando no GitHub...")

            res = requests.get(API_URL, headers=HEADERS)
            
            if res.status_code != 200:
                print(f"  -> ❌ Erro ao buscar comando.json (Status: {res.status_code}).")
                time.sleep(INTERVALO_VERIFICACAO_SEGUNDOS)
                continue

            conteudo_json = res.json()
            sha_atual = conteudo_json.get("sha")
            conteudo_decodificado = base64.b64decode(conteudo_json.get("content")).decode("utf-8")

            if '"executar": true' in conteudo_decodificado:
                print("  -> 🟢 Comando 'executar: true' encontrado!")

                url_do_video_exemplo = "https://www.youtube.com/watch?v=tN3qmAzTrX4"
                executar_tarefa_navegador(url_do_video_exemplo)

                print("  -> 🔵 Resetando o comando para 'false' no GitHub...")
                novo_conteudo = base64.b64encode(b'{"executar": false}').decode("utf-8")
                payload = {
                    "message": "Robô: Resetando comando para false",
                    "content": novo_conteudo,
                    "sha": sha_atual
                }
                res_update = requests.put(API_URL, headers=HEADERS, json=payload)

                if res_update.status_code in [200, 201]:
                    print("  -> ✅ Comando resetado com sucesso!")
                else:
                    print(f"  -> ❌ Erro ao resetar o comando (Status: {res_update.status_code})")
            else:
                print("  -> ⚪ Nenhuma ação pendente.")

            time.sleep(INTERVALO_VERIFICACAO_SEGUNDOS)

        except KeyboardInterrupt:
            print("\n>>> Robô desligado manualmente pelo usuário. <<<")
            break
        except requests.exceptions.RequestException as e:
            print(f"\n❌ Erro de conexão: {e}. Tentando novamente em 60 segundos...")
            time.sleep(60)
        except Exception as e:
            print(f"\nOcorreu um erro inesperado no loop principal: {e}. Reiniciando ciclo...")
            time.sleep(30)

# --- PONTO DE PARTIDA DO SCRIPT ---
if __name__ == "__main__":
    iniciar_robo_ouvinte()