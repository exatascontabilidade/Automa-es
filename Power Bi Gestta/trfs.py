from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import sys
import logging
from datetime import datetime, timedelta
from baixar_tarefas import obter_diretorio_download
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from selenium.common.exceptions import TimeoutException
from email_verificacao import extrair_codigo_do_email




# 🎯 Arquivo de log unificado para simular terminal
PRONT_LOG_PATH = "relatorio_execucao.txt"
def registrar_execucao(texto):
    with open(PRONT_LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {texto}\n")
    print(texto)

# 🎯 Logging básico (arquivo rotativo de falhas)
log_filename = datetime.now().strftime("log_execucao_%Y-%m-%d.log")
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# 🎯 Configurações iniciais
URL_DESTINO = "https://onvio.com.br/#/"
USUARIO = "exatas.contabilidade136@gmail.com"
SENHA = "Exatas@dominio!10"

# 🎯 Checkpoint
def salvar_checkpoint():
    with open("checkpoint.txt", "w") as f:
        f.write(datetime.now().isoformat())

def carregar_checkpoint():
    try:
        with open("checkpoint.txt", "r") as f:
            return datetime.fromisoformat(f.read())
    except:
        return datetime.now()

# 🎯 Utilitários

def tentar_executar(funcao, *args, tentativas=3, espera=5, **kwargs):
    for tentativa in range(1, tentativas + 1):
        try:
            return funcao(*args, **kwargs)
        except Exception as e:
            registrar_execucao(f"[{funcao.__name__}] Erro na tentativa {tentativa}: {e}")
            time.sleep(espera)
    raise Exception(f"Falha após {tentativas} tentativas em {funcao.__name__}")

def navegador_ativo(navegador):
    try:
        navegador.title
        return True
    except:
        return False

def aguardar_download_e_renomear(diretorio_download, timeout=120):
    tempo_inicial = time.time()
    momento_antes = tempo_inicial
    arquivo_final = None
    while time.time() - tempo_inicial < timeout:
        arquivos = [f for f in os.listdir(diretorio_download) if f.lower().endswith(".xlsx")]
        crdownloads = [f for f in os.listdir(diretorio_download) if f.endswith(".crdownload")]
        arquivos_validos = [
            os.path.join(diretorio_download, f)
            for f in arquivos
            if os.path.getmtime(os.path.join(diretorio_download, f)) > momento_antes
        ]
        if arquivos_validos and not crdownloads:
            arquivo_final = max(arquivos_validos, key=os.path.getmtime)
            break
        time.sleep(1)
    if not arquivo_final:
        registrar_execucao("ERRO - Timeout: download não finalizado.")
        return
    agora = datetime.now()
    novo_nome = f"Relatorio_{agora.strftime('%Y-%m-%dT%H-%M-%S')}.xlsx"
    novo_caminho = os.path.join(diretorio_download, novo_nome)
    os.rename(arquivo_final, novo_caminho)
    registrar_execucao(f"INFO - Arquivo salvo como: {novo_nome}")


def iniciar_navegador():
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-gpu")
    options.add_argument("--mute-audio")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-sync")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-extensions")
    #options.add_argument("--headless=new")

    download_dir = obter_diretorio_download()
    os.makedirs(download_dir, exist_ok=True)

    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.disable_download_protection": True,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "profile.default_content_setting_values.popups": 0,
        "profile.default_content_setting_values.notifications": 2,
        "profile.default_content_setting_values.pdf_documents": 1,
    }
    options.add_experimental_option("prefs", prefs)

    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico, options=options)

    navegador.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            })
        """
    })

    return navegador

def abrir_nova_aba(navegador, url_destino):
    navegador.execute_script("window.open('');")
    time.sleep(1)
    navegador.switch_to.window(navegador.window_handles[-1])
    navegador.get(url_destino)
    time.sleep(5)
    if len(navegador.window_handles) > 1:
        navegador.switch_to.window(navegador.window_handles[0])
        navegador.close()
        navegador.switch_to.window(navegador.window_handles[-1])

def entrar_no_portal(navegador, usuario, senha):
    try:
        registrar_execucao("INFO - Aguardando botão de continuar login...")
        WebDriverWait(navegador, 5).until(EC.element_to_be_clickable((By.ID, "trauth-continue-signin-btn"))).click()
        registrar_execucao("INFO - Botão de continuar login clicado.")
    except Exception as e:
        registrar_execucao(f"ERRO - Falha ao clicar no botão de continuar login: {e}")

    try:
        registrar_execucao("INFO - Aguardando campo de e-mail...")
        campo_usuario = WebDriverWait(navegador, 15).until(EC.presence_of_element_located((By.NAME, "username")))
        campo_usuario.clear()
        campo_usuario.send_keys(usuario)
        registrar_execucao("INFO - Campo de e-mail preenchido.")
    except Exception as e:
        registrar_execucao(f"ERRO - Falha ao preencher campo de e-mail: {e}")

    try:
        navegador.find_element(By.XPATH, "//*[@type='submit']").click()
        registrar_execucao("INFO - Botão de envio do e-mail clicado.")
    except Exception as e:
        registrar_execucao(f"ERRO - Falha ao clicar no botão de envio do e-mail: {e}")

    try:
        registrar_execucao("INFO - Aguardando campo de senha...")
        campo_senha = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.ID, "password")))
        campo_senha.clear()
        campo_senha.send_keys(senha)
        registrar_execucao("INFO - Campo de senha preenchido.")
    except Exception as e:
        registrar_execucao(f"ERRO - Falha ao preencher a senha: {e}")

    try:
        navegador.find_element(By.XPATH, "//*[@type='submit']").click()
        registrar_execucao("INFO - Botão de envio da senha clicado.")
    except Exception as e:
        registrar_execucao(f"ERRO - Falha ao clicar no botão de envio da senha: {e}")

# ➕ VERIFICAÇÃO DE AUTENTICAÇÃO EM DUAS ETAPAS
    try:
        WebDriverWait(navegador, 5).until(
            EC.presence_of_element_located(
                (By.XPATH, "//h1[text()='Selecione um método para verificar sua identidade']")
            )
        )
        registrar_execucao("ALERTA - Tela de verificação em duas etapas detectada.")

        # Espera aparecer o campo para inserir o código
        campo_codigo = WebDriverWait(navegador, 15).until(
            EC.presence_of_element_located((By.ID, "otp-code-input"))
        )

        registrar_execucao("INFO - Aguardando recebimento do código por e-mail...")
        codigo = extrair_codigo_do_email()

        if codigo:
            campo_codigo.clear()
            campo_codigo.send_keys(codigo)
            registrar_execucao(f"INFO - Código de verificação inserido: {codigo}")

            # Clica no botão de envio do código (ajuste se o XPath for diferente)
            navegador.find_element(By.XPATH, "//*[@type='submit']").click()
            registrar_execucao("INFO - Código enviado.")
        else:
            raise Exception("Código de verificação não recebido.")
    except TimeoutException:
        registrar_execucao("INFO - Nenhuma verificação em duas etapas solicitada.")



def acessar_aba_processos(navegador, max_tentativas=5):
    for tentativa in range(1, max_tentativas + 1):
        registrar_execucao(f"INFO - Tentativa {tentativa} para acessar aba de processos...")

        try:
            registrar_execucao("INFO - Aguardando carregamento do dashboard...")
            WebDriverWait(navegador, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "dashboard-section__content"))
            )
            registrar_execucao("INFO - Dashboard carregado com sucesso.")
        except Exception as e:
            registrar_execucao(f"ERRO - Falha ao carregar o dashboard: {e}")
            navegador.refresh()
            continue

        try:
            registrar_execucao("INFO - Aguardando botão de menu hamburguer...")
            menu = WebDriverWait(navegador, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "bento-icon-hamburger-menu"))
            )
            menu.click()
            registrar_execucao("INFO - Menu hamburguer clicado.")
        except Exception as e:
            registrar_execucao(f"ERRO - Falha ao clicar no menu hamburguer: {e}")
            navegador.refresh()
            continue

        try:
            registrar_execucao("INFO - Aguardando link para os processos...")
            processos = WebDriverWait(navegador, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'gestta.com.br/dominio/auth/redirect')]"))
            )
            processos.click()
            registrar_execucao("INFO - Link para os processos clicado.")
            return  # Sucesso, sai da função
        except Exception as e:
            registrar_execucao(f"ERRO - Falha ao clicar no link dos processos: {e}")
            navegador.refresh()
            continue

    # Se chegou aqui, nenhuma tentativa teve sucesso
    registrar_execucao("ERRO - Todas as tentativas de acessar a aba de processos falharam.")
        

def acessar_menu_relatorios(navegador, max_tentativas=5):
    for tentativa in range(1, max_tentativas + 1):
        registrar_execucao(f"INFO - Tentativa {tentativa} para acessar o menu de relatórios...")

        try:
            registrar_execucao("INFO - Aguardando botão de relatórios ficar clicável...")
            relatorios = WebDriverWait(navegador, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "[data-testid='gestta_menu-relatorios']"))
            )
            relatorios.click()
            registrar_execucao("INFO - Botão de relatórios clicado com sucesso.")
            return  # Sucesso
        except Exception as e:
            registrar_execucao(f"ERRO - Tentativa {tentativa}: Falha ao clicar no botão de relatórios: {e}")
            navegador.refresh()
            registrar_execucao("INFO - Página recarregada após falha.")
            continue

    registrar_execucao("ERRO - Todas as tentativas de acessar o menu de relatórios falharam.")
        
        

def alternar_para_aba_com_url(navegador, url_parcial_alvo="https://app.gestta.com.br"):
    try:
        registrar_execucao("INFO - Verificando abas abertas para localizar URL alvo...")
        abas = navegador.window_handles
        for aba in abas:
            navegador.switch_to.window(aba)
            url_atual = navegador.current_url
            registrar_execucao(f"INFO - Verificando aba com URL: {url_atual}")
            if url_parcial_alvo in url_atual:
                registrar_execucao("INFO - Aba com URL alvo encontrada.")
                acessar_menu_relatorios(navegador)
                alternar_para_ultima_aba(navegador)
                clicar_em_tarefa_bi(navegador)
                return
        registrar_execucao("ERRO - Nenhuma aba com a URL alvo foi encontrada.")
    except Exception as e:
        registrar_execucao(f"ERRO - Falha ao alternar para aba com URL alvo: {e}")

def alternar_para_ultima_aba(navegador):
    try:
        registrar_execucao("INFO - Alternando para a última aba do navegador...")
        abas = navegador.window_handles
        navegador.switch_to.window(abas[-1])
        registrar_execucao("INFO - Última aba ativada com sucesso.")
    except Exception as e:
        registrar_execucao(f"ERRO - Falha ao alternar para a última aba: {e}")

def clicar_em_tarefa_bi(navegador):
    try:
        registrar_execucao("INFO - Aguardando botão 'Tarefas - BI' ficar clicável...")
        elemento_tarefa_bi = WebDriverWait(navegador, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//a[normalize-space(text())='Tarefas - BI']"))
        )
        elemento_tarefa_bi.click()
        registrar_execucao("INFO - Botão 'Tarefas - BI' clicado com sucesso.")
        time.sleep(3)
    except Exception as e:
        registrar_execucao(f"ERRO - Falha ao clicar em 'Tarefas - BI': {e}")

def limpar_pasta_download(diretorio, extensao=".xlsx"):
    try:
        registrar_execucao(f"INFO - Verificando arquivos com extensão '{extensao}' no diretório: {diretorio}")
        arquivos = [arq for arq in os.listdir(diretorio) if arq.lower().endswith(extensao)]

        if len(arquivos) <= 1:
            registrar_execucao("INFO - Nenhum ou apenas um arquivo encontrado. Nada será apagado.")
            return

        # Obter caminho completo e data de modificação
        caminhos_arquivos = [
            (arquivo, os.path.getmtime(os.path.join(diretorio, arquivo)))
            for arquivo in arquivos
        ]

        # Ordenar por data de modificação (mais recente por último)
        caminhos_arquivos.sort(key=lambda x: x[1], reverse=True)

        # Manter apenas o mais novo (primeiro da lista após reverse)
        arquivos_a_manter = caminhos_arquivos[:1]
        arquivos_a_apagar = caminhos_arquivos[1:]

        for arquivo, _ in arquivos_a_apagar:
            caminho = os.path.join(diretorio, arquivo)
            try:
                os.remove(caminho)
                registrar_execucao(f"INFO - Arquivo deletado: {arquivo}")
            except Exception as e:
                registrar_execucao(f"ERRO - Não foi possível deletar {arquivo}: {e}")
    except Exception as e:
        registrar_execucao(f"ERRO - Falha ao acessar o diretório {diretorio}: {e}")

def exportar_para_excel(navegador):
    try:
        download_dir = obter_diretorio_download()
        registrar_execucao("INFO - Diretório de download obtido.")
        wait = WebDriverWait(navegador, 15)

        registrar_execucao("INFO - Limpando arquivos .xlsx antes de iniciar exportação.")

        registrar_execucao("INFO - Aguardando carregamento do título da coluna 'Cliente - Código'...")
        wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='tooltip' and @title='Cliente - Código']")))
        registrar_execucao("INFO - Coluna 'Cliente - Código' carregada.")


        registrar_execucao("INFO - Aguardando botão de exportação ficar clicável...")
        botao_exportar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "li.custom-flexmonster-action.export-action")))
        botao_exportar.click()
        registrar_execucao("INFO - Botão de exportação clicado.")

        registrar_execucao("INFO - Aguardando botão 'EXCEL' na exportação...")
        botao_excel = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'export-to-excel-action')]/span[contains(text(), 'EXCEL')]")))
        botao_excel.click()
        registrar_execucao("INFO - Botão 'EXCEL' clicado. Aguardando download...")
        aguardar_download_e_renomear(download_dir)
    except Exception as e:
        registrar_execucao(f"ERRO - Falha durante a exportação para Excel: {e}")
        

def aguardar_download_e_renomear(diretorio_download, timeout=60):
    try:
        registrar_execucao("INFO - Iniciando verificação de download de arquivo .xlsx...")
        tempo_inicial = time.time()
        momento_antes = tempo_inicial
        arquivo_final = None

        while time.time() - tempo_inicial < timeout:
            arquivos = [f for f in os.listdir(diretorio_download) if f.lower().endswith(".xlsx")]
            crdownloads = [f for f in os.listdir(diretorio_download) if f.endswith(".crdownload")]

            arquivos_validos = [
                os.path.join(diretorio_download, f)
                for f in arquivos
                if os.path.getmtime(os.path.join(diretorio_download, f)) > momento_antes
            ]

            if arquivos_validos and not crdownloads:
                arquivo_final = max(arquivos_validos, key=os.path.getmtime)
                break

            time.sleep(1)

        if not arquivo_final:
            registrar_execucao("ERRO - Timeout: download não finalizado.")
            return

        agora = datetime.now()
        novo_nome = f"Relatorio_{agora.strftime('%Y-%m-%dT%H-%M-%S')}.xlsx"
        novo_caminho = os.path.join(diretorio_download, novo_nome)
        os.rename(arquivo_final, novo_caminho)
        registrar_execucao(f"INFO - Arquivo salvo como: {novo_nome}")
        download_dir = obter_diretorio_download()
        limpar_pasta_download(download_dir, extensao=".xlsx")

    except Exception as e:
        registrar_execucao(f"ERRO - Falha durante verificação ou renomeação do download: {e}")


def ciclo_download_relatorio(navegador):
    # 💾 Atualiza o arquivo txt sobrescrevendo o conteúdo
    with open(PRONT_LOG_PATH, "w", encoding="utf-8") as f:
        f.write(f"=== RELATÓRIO DE EXECUÇÃO ===\n")

    try:
        registrar_execucao(f"INFO Iniciando ciclo de download único...")
        navegador.refresh()
        time.sleep(5)
        exportar_para_excel(navegador)
        salvar_checkpoint()
        registrar_execucao("INFO Ciclo único finalizado com sucesso.")

    except Exception as erro:
        registrar_execucao(f"ERRO - Erro durante o ciclo: {erro}")
        if not navegador_ativo(navegador):
            navegador = iniciar_navegador()
            abrir_nova_aba(navegador, URL_DESTINO)
            entrar_no_portal(navegador, USUARIO, SENHA)
            acessar_aba_processos(navegador)
            alternar_para_aba_com_url(navegador)

def enviar_email_erro(mensagem, caminho_print="screenshot.png"):
    remetente = "ealoisio16@gmail.com"
    senha_app = "kxdwcxpfiaqlfzmm"
    destinatario = "ealoisio16@gmail.com"

    msg = MIMEMultipart()
    msg['Subject'] = '❌ Erro na execução do script Power BI Gestta'
    msg['From'] = remetente
    msg['To'] = destinatario

    msg.attach(MIMEText(mensagem, 'plain'))

    if os.path.exists(caminho_print):
        with open(caminho_print, 'rb') as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(caminho_print)}')
            msg.attach(part)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(remetente, senha_app)
            server.send_message(msg)
        print("INFO - E-mail de erro enviado com print.")
    except Exception as e:
        print(f"ERRO - Falha ao enviar e-mail: {e}")

  
  
        
# ==== EXECUÇÃO FINAL ====
if __name__ == "__main__":
    try:
        registrar_execucao("INFO - Iniciando execução principal do script...")
        navegador = iniciar_navegador()
        abrir_nova_aba(navegador, URL_DESTINO)
        entrar_no_portal(navegador, USUARIO, SENHA)
        acessar_aba_processos(navegador)
        alternar_para_aba_com_url(navegador)
        ciclo_download_relatorio(navegador)
        navegador.quit()

        mensagem_exe = "Execução do script Power BI Gestta finalizada com sucesso."

    except Exception as erro:
        mensagem_erro = f"Erro na execução do script Power BI Gestta:\n\n{erro}"
        registrar_execucao(f"ERRO - Erro geral na execução: {erro}")

        # Tentativa de captura de screenshot
        try:
            navegador.save_screenshot("screenshot.png")
            registrar_execucao("INFO - Screenshot de erro salvo como 'screenshot.png'.")
        except Exception as e:
            registrar_execucao(f"ERRO - Falha ao capturar screenshot: {e}")

        enviar_email_erro(mensagem_erro, caminho_print="screenshot.png")

        try:
            navegador.quit()
        except:
            pass

        sys.exit(1)