
import sys
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

def iniciar_navegador():
    options = Options()
    options.add_argument("--disable-gpu")
    options.add_argument("--headless=new")
    options.add_argument("--mute-audio")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=options)
    return driver

def login(driver):
    wait = WebDriverWait(driver, 10)
    driver.get("https://www.sefaz.se.gov.br/SitePages/acesso_usuario.aspx")
    try:
        try:
            accept_button = wait.until(EC.element_to_be_clickable((By.ID, 'accept-button')))
            accept_button.click()
        except:
            pass
        time.sleep(1)
        driver.switch_to.frame(driver.find_elements(By.TAG_NAME, "iframe")[0])
        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'acessoRapido'))).click()
        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.XPATH,
            "//option[@value='https://security.sefaz.se.gov.br/internet/portal/contabilista/atoAcessoContabilista.jsp']"))).click()
        driver.find_element(By.TAG_NAME, "body").click()
        time.sleep(1)
        driver.switch_to.frame(wait.until(EC.presence_of_element_located(
            (By.XPATH, "//iframe[contains(@src, 'atoAcessoContribuinte.jsp')]"))))
        tabela_login = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tabelaVerde")))
        usuario = tabela_login.find_element(By.NAME, "UserName")
        senha = tabela_login.find_element(By.NAME, "Password")
        botao_login = tabela_login.find_element(By.NAME, "submit")
        usuario.send_keys("SE007829")
        senha.send_keys("Exatas2024@")
        botao_login.click()
        return True
    except Exception as e:
        print(f"Erro ao realizar login: {e}")
        return False

def navegar_para_solicitacao_xml(driver):
    wait = WebDriverWait(driver, 10)
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys(Keys.ENTER)
        time.sleep(2)
        driver.find_elements(By.TAG_NAME, "a")[0].click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'NFE/DOCUMENTOS ELETRONICOS')]"))).click()
        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Solicitar Arquivos XML')]"))).click()
        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
            "body > table > tbody > tr:nth-child(2) > td:nth-child(2) > table > tbody > tr > td > table:nth-child(6) > tbody > tr > td:nth-child(1) > a"))).click()
    except Exception as e:
        print(f"Erro na navegação inicial: {e}")
        return False
    return True

def selecionar_empresa(driver, inscricao_municipal):
    wait = WebDriverWait(driver, 10)
    try:
        select_empresas = wait.until(EC.presence_of_element_located((By.ID, "cdPessoaContribuinte")))
        select = Select(select_empresas)
        select.select_by_value(str(inscricao_municipal))
        wait.until(EC.element_to_be_clickable((By.ID, "okButton"))).click()
        return True
    except Exception as e:
        print(f"Erro ao selecionar empresa: {e}")
        return False

def selecionar_tipo_arquivo(driver, tipo_arquivo):
    wait = WebDriverWait(driver, 10)
    try:
        tipo_arquivo = tipo_arquivo.strip().upper()
        select_el = wait.until(EC.presence_of_element_located((By.ID, "tipoArquivo")))
        select = Select(select_el)
        opcoes = [op.text.strip().upper() for op in select.options]
        if tipo_arquivo in opcoes:
            select.select_by_visible_text(tipo_arquivo)
            wait.until(EC.element_to_be_clickable((By.ID, "okButton"))).click()
            return True
        else:
            print(f"Tipo de arquivo inválido. Opções: {opcoes}")
            return False
    except Exception as e:
        print(f"Erro ao selecionar tipo de arquivo: {e}")
        return False

def selecionar_tipo_pesquisa(driver, pesquisar_por, tipo_arquivo):
    wait = WebDriverWait(driver, 10)
    try:
        pesquisar_por = pesquisar_por.strip()
        if tipo_arquivo in ["NFE", "NFC"]:
            tipo_pesquisa_el = wait.until(EC.presence_of_element_located((By.ID, "tipoPesquisa")))
            select = Select(tipo_pesquisa_el)
            opcoes = [op.text.strip() for op in select.options]
            if pesquisar_por in opcoes:
                select.select_by_visible_text(pesquisar_por)
                return True
            else:
                print(f"Tipo de pesquisa inválido: {pesquisar_por}. Opções: {opcoes}")
                return False
        elif tipo_arquivo == "CTE":
            mapeamento = {
                "Remetente": "Remetente",
                "Expedidor": "Expedidor",
                "Recebedor": "Recebedor",
                "Destinatário": "Destinatario",
                "Emitente": "Emitente",
                "Outros": "Outros"
            }
            if pesquisar_por in mapeamento:
                campo = mapeamento[pesquisar_por]
                campo_el = wait.until(EC.presence_of_element_located((By.ID, campo)))
                if not campo_el.is_selected():
                    campo_el.click()
                return True
            else:
                print(f"Tipo de pesquisa inválido: {pesquisar_por}")
                return False
    except Exception as e:
        print(f"Erro ao selecionar tipo de pesquisa: {e}")
        return False

def formatar_data(data):
    try:
        if isinstance(data, pd.Timestamp):
            return data.strftime("%d/%m/%Y")
        data_obj = pd.to_datetime(data, format="%d/%m/%Y", errors='coerce')
        if pd.isna(data_obj):
            raise ValueError("Data inválida")
        return data_obj.strftime("%d/%m/%Y")
    except Exception as e:
        print(f"Erro ao formatar data '{data}': {e}")
        return None

def preencher_datas(driver, data_inicial, data_final):
    wait = WebDriverWait(driver, 10)
    try:
        data_inicial_fmt = formatar_data(data_inicial)
        data_final_fmt = formatar_data(data_final)
        if not (data_inicial_fmt and data_final_fmt):
            print("Erro no formato das datas")
            return False

        campo_ini = wait.until(EC.presence_of_element_located((By.ID, "dtInicio")))
        campo_fim = wait.until(EC.presence_of_element_located((By.ID, "dtFinal")))

        for campo, valor in [(campo_ini, data_inicial_fmt), (campo_fim, data_final_fmt)]:
            campo.clear()
            campo.send_keys(Keys.CONTROL + "a")
            campo.send_keys(Keys.DELETE)
            campo.send_keys(valor)

        wait.until(EC.element_to_be_clickable((By.ID, "okButton"))).click()
        return True
    except Exception as e:
        print(f"Erro ao preencher datas: {e}")
        return False

def verificar_mensagem_erro(driver):
    wait = WebDriverWait(driver, 5)
    try:
        msg_el = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "fontMessageError")))
        msg = msg_el.text.strip()
        if msg:
            print(f"Mensagem de erro detectada: {msg}")
    except:
        print("Nenhuma mensagem de erro detectada.")

def main():
    inscricao_municipal = sys.argv[1]
    tipo_arquivo = sys.argv[2]
    pesquisar_por = sys.argv[3]
    data_inicial = sys.argv[4]
    data_final = sys.argv[5]

    driver = iniciar_navegador()

    if not login(driver): sys.exit(1)
    if not navegar_para_solicitacao_xml(driver): sys.exit(1)
    if not selecionar_empresa(driver, inscricao_municipal): sys.exit(1)
    if not selecionar_tipo_arquivo(driver, tipo_arquivo): sys.exit(1)
    if not selecionar_tipo_pesquisa(driver, pesquisar_por, tipo_arquivo): sys.exit(1)
    if not preencher_datas(driver, data_inicial, data_final): sys.exit(1)

    verificar_mensagem_erro(driver)
    print(" Processo finalizado com sucesso!")
    driver.quit()

if __name__ == "__main__":
    main()
