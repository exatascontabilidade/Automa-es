import time
import threading
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium_config.driver_config import configurar_driver
from automacoes.download import baixar_arquivos_com_blocos

from utils.state import executando
import utils.state as state

usar_headless = False


def checar_parada(navegador):
    global executando
    if not state.executando:
        print("🛑 Execução interrompida.")
        navegador.quit()
        return True
    return False

def executar_codigo_completo():
    global executando
    if not state.executando:
        print("🛑 Execução cancelada antes de iniciar.")
        return

    navegador = configurar_driver(headless=usar_headless)
    wait = WebDriverWait(navegador, 3)
    navegador.get("https://www.sefaz.se.gov.br/SitePages/acesso_usuario.aspx")
    navegador.maximize_window()

    try:
        if checar_parada(navegador): return
        try:
            wait.until(EC.element_to_be_clickable((By.ID, 'accept-button'))).click()
        except:
            print("Botão 'Aceitar' não encontrado. Continuando.")
        time.sleep(0.5)

        if checar_parada(navegador): return
        navegador.switch_to.frame(navegador.find_elements(By.TAG_NAME, "iframe")[0])
        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'acessoRapido'))).click()
        time.sleep(0.5)

        if checar_parada(navegador): return
        wait.until(EC.element_to_be_clickable((By.XPATH, "//option[contains(@value,'contabilista')]"))).click()
        navegador.find_element(By.TAG_NAME, "body").click()
        time.sleep(0.5)

        if checar_parada(navegador): return
        navegador.switch_to.frame(wait.until(EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'atoAcessoContribuinte.jsp')]"))))
        tabela_login = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tabelaVerde")))
        tabela_login.find_element(By.NAME, "UserName").send_keys("SE007829")
        tabela_login.find_element(By.NAME, "Password").send_keys("Exatas2024@")
        tabela_login.find_element(By.NAME, "submit").click()
        print("🎉 Login realizado com sucesso!")

        if checar_parada(navegador): return
        navegador.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        navegador.find_element(By.TAG_NAME, "body").send_keys(Keys.TAB)
        navegador.find_element(By.TAG_NAME, "body").send_keys(Keys.ENTER)
        time.sleep(0.5)

        if checar_parada(navegador): return
        navegador.find_elements(By.TAG_NAME, "a")[0].click()

        if checar_parada(navegador): return
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'NFE/DOCUMENTOS ELETRONICOS')]"))).click()
        print("✅ Acessado menu NFE/DOCUMENTOS ELETRONICOS")
        time.sleep(0.5)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Solicitar Arquivos XML')]"))).click()
        time.sleep(0.5)

        print("📥 Iniciando verificação de arquivos para download...")
        baixar_arquivos_com_blocos(navegador)
        print("✅ Concluído o processo de download e renomeação.")

    except Exception as e:
        print(f"❌ Erro durante o processo: {e}")

def iniciar_thread():
    global executando
    state.executando = True
    thread = threading.Thread(target=executar_codigo_completo)
    thread.start()

def parar_automacao():
    global executando
    if executando:
        state.executando = False
        print("🛑 Automação interrompida pelo usuário.")
    else:
        print("⚠️ Nenhuma automação em execução.")