from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import random
import os
import sys
from baixar_pdf import obter_diretorio_download

sys.path.append(os.path.dirname(os.path.abspath(__file__)))  # Adiciona o diretório do script ao path

# ⚠️ PEGANDO E-MAIL E SENHA DE VARIÁVEIS DE AMBIENTE ⚠️
EMAIL_GMAIL = "financeiroexatas136@gmail.com" # Defina antes de rodar o script: export EMAIL_GMAIL="seu_email@gmail.com"
SENHA_GMAIL = "Exatas1010@" # Defina antes de rodar: export SENHA_GMAIL="sua_senha"

def random_sleep(min_seconds=1, max_seconds=2):
    """Aguarda um tempo aleatório para simular      comportamento humano."""
    time.sleep(random.uniform(min_seconds, max_seconds))


def iniciar_navegador(headless=False):
    options = webdriver.ChromeOptions()
    
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-gpu")
    options.add_argument("--mute-audio")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-extensions")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-infobars")
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--disable-features=ChromeWhatsNewUI")
    
    options.add_argument("--disable-sync")  # Impede o pop-up "Fazer login no Chrome"
    options.add_argument("--no-first-run")  # Remove tela inicial de configuração
    options.add_argument("--no-default-browser-check")  # Evita verificação de navegador padrão
    options.add_argument("--disable-popup-blocking")  # Pode ajudar a evitar redirecionamentos forçados
    options.add_argument("--disable-features=ChromeWhatsNewUI,AccountConsistency")  # Impede avisos de login
    options.add_argument("--force-device-scale-factor=1")  # Evita redimensionamento que pode ativar o pop-up
    options.add_argument("--disable-component-update")  # Impede que o Chrome peça configurações iniciais

    


    if headless:
        options.add_argument("--headless=new")

    # 📂 Diretório de downloads
    download_dir = obter_diretorio_download()
    os.makedirs(download_dir, exist_ok=True)

    prefs = {
        "download.default_directory": download_dir,  # Define o diretório de download
        "download.prompt_for_download": False,  # Não pergunta onde salvar
        "download.directory_upgrade": True,  # Garante que o Chrome respeite as configurações de download
        "safebrowsing.disable_download_protection": True,  # Evita bloqueios ao baixar arquivos
        "plugins.always_open_pdf_externally": True,  # Faz o Chrome baixar PDFs ao invés de abrir
        "profile.default_content_setting_values.automatic_downloads": 1,  # Permite múltiplos downloads sem confirmação
        "profile.default_content_setting_values.popups": 0,  # Bloqueia pop-ups que poderiam atrapalhar
        "profile.default_content_setting_values.notifications": 2,  # Bloqueia notificações que podem interferir
        "profile.default_content_setting_values.pdf_documents": 1,  # Força o download de PDFs
        "download.extensions_to_open": "",  # Impede o Chrome de abrir qualquer arquivo automaticamente
    }

    options.add_experimental_option("prefs", prefs)

    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico, options=options)

    return navegador


def fazer_login(navegador):
    """Realiza login no Gmail."""
    navegador.get("https://mail.google.com")

    wait = WebDriverWait(navegador, 30)

    try:
        # 🔑 Insere o e-mail
        email_elem = wait.until(EC.presence_of_element_located((By.ID, "identifierId")))
        email_elem.send_keys(EMAIL_GMAIL)
        random_sleep()
        email_elem.send_keys(Keys.RETURN)
        time.sleep(10)

        # 🔒 Aguarda a entrada da senha
        password_elem = wait.until(EC.presence_of_element_located((By.NAME, "Passwd")))
        password_elem.send_keys(SENHA_GMAIL)
        random_sleep()
        password_elem.send_keys(Keys.RETURN)

        random_sleep(3, 5)

        # ✅ Confirma que o login foi bem-sucedido
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "z0")))  # Botão "Escrever" do Gmail
        print("[INFO] - Login realizado com sucesso!")
        time.sleep(5)

        return True

    except Exception as e:
        print(f"⚠️ Erro ao fazer login: {e}")
        return False


def aguardar_fechamento():
    """Aguarda o usuário pressionar Enter antes de fechar o navegador."""
    input("🔴 Pressione ENTER para fechar o navegador...")
    
if __name__ == "__main__":
    navegador = iniciar_navegador()
    if fazer_login(navegador):
        from listar_emails_storage import listar_todos_emails,buscar_emails, verificar_fim_paginacao
        from baixar_boletos import baixar_boletos
        from renomear_pdfs import renomear_arquivos
        from pdfs_zip import extrair_zips_em_lote
        from login_onvio import abrir_nova_aba, entrar_no_portal, acessar_aba_documentos
        
    
        


        # 🔍 Inicia a busca no mesmo navegador logad
        #buscar_emails(navegador)
        #listar_todos_emails(navegador)
        #verificar_fim_paginacao(navegador)
        #print("⏳ Aguardando 10 segundos antes de iniciar o download dos boletos...")
        #time.sleep(5) 
        
        #baixar_boletos(navegador)
        #print("⏳ Aguardando 5 segundos antes de renomear boletos...")
        #time.sleep(2)
        #pasta_origem = obter_diretorio_download()           # Altere para o caminho da sua pasta com os .zip
        #pasta_destino = obter_diretorio_download()    
        #extrair_zips_em_lote(pasta_origem, pasta_destino)
        #renomear_arquivos()
        #🔗 Acesso ao Onvio
        url_destino = "https://onvio.com.br/#/"
        usuario = "automacao.gestta@exatascontabilidade.com.br"
        senha = "Exatas@1010"
        abrir_nova_aba(navegador, url_destino)
        entrar_no_portal(navegador, usuario, senha)
        acessar_aba_documentos(navegador)
        
        aguardar_fechamento()   

    else:
        print("❌ Não foi possível fazer login. Verifique suas credenciais ou autenticação 2FA.")
        navegador.quit()
