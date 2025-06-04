import time
import threading
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

from scripts.dolowd.baixar import baixar_arquivos_com_blocos
import scripts.dolowd.state as state

usar_headless = False

def configurar_driver(headless=False):
    import os
    pasta_download = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp")
    os.makedirs(pasta_download, exist_ok=True)

    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--mute-audio")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    prefs = {
        "download.default_directory": pasta_download,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    servico = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=servico, options=options)

def checar_parada():
    if not state.get_estado():
        if state.navegador:
            state.navegador.quit()
        return True
    return False

def executar_codigo_completo():
    print("[THREAD] Entrou na função executar_codigo_completo")
    if not state.get_estado():
        print("[THREAD] Execução não iniciada. Abortando.")
        return

    state.navegador = configurar_driver(headless=usar_headless)
    wait = WebDriverWait(state.navegador, 3)

    try:
        navegador = state.navegador
        navegador.get("https://www.sefaz.se.gov.br/SitePages/acesso_usuario.aspx")
        navegador.maximize_window()

        if checar_parada(): return
        try:
            wait.until(EC.element_to_be_clickable((By.ID, 'accept-button'))).click()
        except:
            pass
        time.sleep(0.5)

        if checar_parada(): return
        navegador.switch_to.frame(navegador.find_elements(By.TAG_NAME, "iframe")[0])
        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'acessoRapido'))).click()
        time.sleep(0.5)

        if checar_parada(): return
        wait.until(EC.element_to_be_clickable((By.XPATH, "//option[contains(@value,'contabilista')]"))).click()
        navegador.find_element(By.TAG_NAME, "body").click()
        time.sleep(0.5)

        if checar_parada(): return
        navegador.switch_to.frame(wait.until(EC.presence_of_element_located(
            (By.XPATH, "//iframe[contains(@src, 'atoAcessoContribuinte.jsp')]")
        )))
        tabela_login = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tabelaVerde")))
        tabela_login.find_element(By.NAME, "UserName").send_keys("SE007829")
        tabela_login.find_element(By.NAME, "Password").send_keys("Exatas2024@")
        time.sleep(1)
        tabela_login.find_element(By.NAME, "submit").click()

        if checar_parada(): return
        navegador.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        navegador.find_element(By.TAG_NAME, "body").send_keys(Keys.ENTER)
        time.sleep(0.5)

        if checar_parada(): return
        navegador.find_elements(By.TAG_NAME, "a")[0].click()

        if checar_parada(): return
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'NFE/DOCUMENTOS ELETRONICOS')]"))).click()
        time.sleep(0.5)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Solicitar Arquivos XML')]"))).click()
        time.sleep(0.5)

        baixar_arquivos_com_blocos(navegador)

    except Exception as e:
        print(f"[ERRO] Erro inesperado durante a execução: {str(e)}")

    finally:
        try:
            if state.navegador:
                state.navegador.quit()
                print("[EXIT] Navegador encerrado no finally.")
        except Exception as e:
            print(f"[ERRO] Falha ao encerrar navegador: {e}")
            
            state.navegador = None
            state.remover_estado() 

def iniciar_thread():
    print("[THREAD] Chamando iniciar_thread()")
    state.set_estado(True)
    thread = threading.Thread(target=executar_codigo_completo)
    thread.start()

def parar_automacao():
    if state.get_estado():
        state.remover_estado()  # ou state.set_estado(False)
        return "⏹ Consulta marcada para parar."
    else:
        return "⚠️ Nenhuma consulta está em andamento."
