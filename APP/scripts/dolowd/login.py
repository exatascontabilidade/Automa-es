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
    options = Options()
    if state.usar_headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--mute-audio")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    servico = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=servico, options=options)

def checar_parada(navegador):
    if not state.executando:
        navegador.quit()
        return True
    return False

def executar_codigo_completo():
    if not state.executando:
        return

    navegador = configurar_driver(headless=usar_headless)
    wait = WebDriverWait(navegador, 3)

    try:
        navegador.get("https://www.sefaz.se.gov.br/SitePages/acesso_usuario.aspx")
        navegador.maximize_window()

        if checar_parada(navegador): return
        try:
            wait.until(EC.element_to_be_clickable((By.ID, 'accept-button'))).click()
        except:
            pass
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
        navegador.switch_to.frame(wait.until(EC.presence_of_element_located(
            (By.XPATH, "//iframe[contains(@src, 'atoAcessoContribuinte.jsp')]")
        )))
        tabela_login = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tabelaVerde")))
        tabela_login.find_element(By.NAME, "UserName").send_keys("SE007829")
        tabela_login.find_element(By.NAME, "Password").send_keys("Exatas2024@")
        tabela_login.find_element(By.NAME, "submit").click()

        if checar_parada(navegador): return
        navegador.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        navegador.find_element(By.TAG_NAME, "body").send_keys(Keys.TAB)
        navegador.find_element(By.TAG_NAME, "body").send_keys(Keys.ENTER)
        time.sleep(0.5)

        if checar_parada(navegador): return
        navegador.find_elements(By.TAG_NAME, "a")[0].click()

        if checar_parada(navegador): return
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'NFE/DOCUMENTOS ELETRONICOS')]"))).click()
        time.sleep(0.5)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Solicitar Arquivos XML')]"))).click()
        time.sleep(0.5)

        baixar_arquivos_com_blocos(navegador)

    except Exception as e:
        print(f"[ERRO] Erro inesperado durante a execução: {str(e)}")

    finally:
        try:
            navegador.quit()
        except:
            pass

def iniciar_thread():
    state.executando = True
    thread = threading.Thread(target=executar_codigo_completo)
    thread.start()

def parar_automacao():
    if state.executando:
        state.executando = False
