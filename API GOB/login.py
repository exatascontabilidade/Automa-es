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
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from selenium.common.exceptions import TimeoutException
from email_verificacao import extrair_codigo_do_email
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
import pandas as pd
import re
import os
import json
import vg 
import pandas as pd
import os
import json
from selenium.common.exceptions import TimeoutException
import re
from selenium.webdriver.common.keys import Keys
import vg 
import os
import time
import json
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    NoSuchElementException,
    WebDriverException,
)# sua variável global, mantenha conforme já usada

#Func de establecer o diretorio de download dos arquivos 

def obter_diretorio_download():
    """
    Cria a pasta 'Gestta - Tarefas' no diretório onde o script está, se não existir.
    Retorna o caminho completo da pasta sem apagar nenhum arquivo.
    """
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    diretorio_download = os.path.join(diretorio_atual, "Gestta - Tarefas")

    if not os.path.exists(diretorio_download):
        os.makedirs(diretorio_download)
        print(f" Diretório de download criado: {diretorio_download}")
    else:
        print(f" Diretório de download já existe: {diretorio_download}")

    return diretorio_download


DOWNLOAD_DIR = obter_diretorio_download()

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
USUARIO = "automacao.gestta@exatascontabilidade.com.br"
SENHA = "Exatas@1010"

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
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--start-maximized")
    options.add_argument("--disable-infobars")
    options.add_argument("--mute-audio")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-sync")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

    download_dir = obter_diretorio_download()
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
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
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


# FUNC - FEITA DIA 24/07/2025
def entrar_no_portal(navegador, usuario, senha):
    max_tentativas = 4
    tentativa = 0

    while tentativa < max_tentativas:
        tentativa += 1
        try:
            registrar_execucao(f"INFO - Tentativa {tentativa}: Aguardando botão de continuar login...")
            WebDriverWait(navegador, 5).until(
                EC.element_to_be_clickable((By.ID, "trauth-continue-signin-btn"))
            ).click()
            registrar_execucao("INFO - Botão de continuar login clicado.")
            break  # sai do loop se funcionar
        except Exception as e:
            registrar_execucao(f"⚠️ Tentativa {tentativa}: Falha ao encontrar botão de continuar login: {e}")
            if tentativa < max_tentativas:
                registrar_execucao("🔄 Recarregando a página...")
                navegador.refresh()
                time.sleep(3)
            else:
                registrar_execucao("❌ Erro definitivo: Não foi possível encontrar o botão de login após múltiplas tentativas.")
                return False  # ou raise Exception(...)

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
        WebDriverWait(navegador, 10).until(
            lambda d: d.find_element(By.XPATH, "//h1[contains(text(), 'verificar sua identidade')]").is_displayed()
        )
        registrar_execucao("ALERTA - Tela de verificação em duas etapas detectada.")

        xpath_botao_email = "//button[@name='action' and contains(@value, 'email')]"

        WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.XPATH, xpath_botao_email))
        )
        botao_email = WebDriverWait(navegador, 5).until(
            EC.element_to_be_clickable((By.XPATH, xpath_botao_email))
        )
        botao_email.click()
        registrar_execucao("INFO - Método de verificação por e-mail selecionado.")

        campo_codigo = WebDriverWait(navegador, 60).until(
            EC.presence_of_element_located((By.ID, "code"))
        )

        registrar_execucao("INFO - Aguardando recebimento do código por e-mail...")
        time.sleep(20)
        codigo = extrair_codigo_do_email()

        if codigo:
            campo_codigo.clear()
            campo_codigo.send_keys(codigo)
            registrar_execucao(f"INFO - Código de verificação inserido: {codigo}")

            botao_continuar = WebDriverWait(navegador, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@name='action' and @value='default']"))
            )
            botao_continuar.click()
            registrar_execucao("INFO - Código confirmado com sucesso.")
        else:
            raise Exception("Código de verificação não recebido.")
    except TimeoutException:
        registrar_execucao("INFO - Nenhuma verificação em duas etapas solicitada.")


# FUNC - FEITA DIA 21/07/2025

def obter_codigo_empresa(caminho_planilha_codigos="Relação Empresas - Nome - CNPJ.xls", caminho_json_baixados="parcelamentos_baixados.json"):
    """
    Retorna a lista de códigos das empresas com base nos CNPJs encontrados no JSON dos parcelamentos baixados.
    Também armazena os CNPJs correspondentes.
    Compatível com planilhas .xls contendo colunas 'Cód.' e 'CNPJ'.
    """

    print(f"📂 Pasta atual do script: {os.getcwd()}")
    caminho_planilha_codigos = os.path.join(os.getcwd(), caminho_planilha_codigos)
    print(f"🔍 Procurando arquivo: {caminho_planilha_codigos}")

    # 1. Verifica JSON
    if not os.path.exists(caminho_json_baixados):
        print("❌ Arquivo de parcelamentos não encontrado.")
        vg.lista_codigos_empresas = []
        vg.lista_cnpjs_empresas = []
        return []

    with open(caminho_json_baixados, "r", encoding="utf-8") as f:
        dados_baixados = json.load(f)

    cnpjs_baixados = {entry["cnpj"] for entry in dados_baixados if "cnpj" in entry}
    cnpjs_baixados_limpos = {re.sub(r"\D", "", cnpj) for cnpj in cnpjs_baixados}

    if not cnpjs_baixados_limpos:
        print("⚠️ Nenhum CNPJ encontrado no JSON.")
        vg.lista_codigos_empresas = []
        vg.lista_cnpjs_empresas = []
        return []

    # 2. Verifica planilha
    if not os.path.exists(caminho_planilha_codigos):
        print("❌ Planilha de códigos não encontrada.")
        vg.lista_codigos_empresas = []
        vg.lista_cnpjs_empresas = []
        return []

    try:
        df = pd.read_excel(caminho_planilha_codigos, dtype=str, engine="xlrd")
        print(f"🧾 Colunas da planilha: {df.columns.tolist()}")
    except Exception as e:
        print(f"❌ Erro ao ler a planilha .xls: {e}")
        vg.lista_codigos_empresas = []
        vg.lista_cnpjs_empresas = []
        return []

    # 3. Limpa e cruza CNPJs
    df["CNPJ"] = df["CNPJ"].str.replace(r"\D", "", regex=True)
    df_filtrado = df[df["CNPJ"].isin(cnpjs_baixados_limpos)]

    codigos_encontrados = df_filtrado["Cód."].dropna().unique().tolist()
    cnpjs_encontrados = df_filtrado["CNPJ"].dropna().unique().tolist()

    vg.lista_codigos_empresas = codigos_encontrados
    vg.lista_cnpjs_empresas = cnpjs_encontrados

    print(f"📌 Códigos atribuídos à variável global: {codigos_encontrados}")
    print(f"📌 CNPJs atribuídos à variável global: {cnpjs_encontrados}")




# FUNC - FEITA DIA 21/07/2025

# Acesso a parte de documentos 
def acessar_aba_documentos(navegador, caminho_planilha_codigos, max_tentativas=3):
    """
    Acessa diretamente a URL da aba 'Documentos' do Onvio.
    Agora associa o CNPJ correspondente ao código da empresa processada,
    e armazena em vg.cnpj_empresa_selecionada.
    """
    tentativa = 0
    url_documentos = "https://onvio.com.br/staff/#/documents/client"

    while tentativa < max_tentativas:
        tentativa += 1
        try:
            print(f"⏳ Tentativa {tentativa}/{max_tentativas} para acessar a aba 'Documentos'...")

            navegador.get(url_documentos)

            WebDriverWait(navegador, 15).until(lambda d: d.current_url.startswith(url_documentos))
            print(f"✅ URL da aba 'Documentos' carregada: {navegador.current_url}")

            WebDriverWait(navegador, 20).until(
                EC.presence_of_element_located((By.XPATH, "//ul[contains(@class, 'nav-tabs')]//span[contains(text(), 'Documentos do Cliente')]"))
            )
            print("✅ Elementos da aba 'Documentos' carregados com sucesso!")

            print("🔄 Recarregando a URL diretamente")
            navegador.get(url_documentos)
            time.sleep(3)

            obter_codigo_empresa(caminho_planilha_codigos)
            lista_empresas = vg.lista_codigos_empresas
            lista_cnpjs = vg.lista_cnpjs_empresas

            if not lista_empresas:
                print("❌ Nenhuma empresa encontrada para processar.")
                return False

            mapeamento_cod_cnpj = dict(zip(lista_empresas, lista_cnpjs))

            for codigo_empresa in lista_empresas:
                cnpj_empresa = mapeamento_cod_cnpj.get(codigo_empresa, "CNPJ não encontrado")
                print(f"\n🔄 Iniciando o processamento da empresa com código: {codigo_empresa} | CNPJ: {cnpj_empresa}")

                try:
                    WebDriverWait(navegador, 20).until(
                        EC.presence_of_element_located((By.XPATH, '//span[@data-qe-id="Bluemoon.DMS.SpecialFolders.MyDocuments" and contains(text(), "Meus Documentos")]'))
                    )
                    print("✅ Elemento 'Meus Documentos' carregado com sucesso.")
                except Exception as e:
                    print(f"❌ Erro ao aguardar carregamento do elemento 'Meus Documentos': {e}")

                res = selecionar_empresa(navegador, codigo_empresa)

                if res["status"]:
                    nome_empresa = res["nome_empresa"]

                    # ✅ Armazenar dados em vg
                    vg.nome_empresa_selecionada = nome_empresa
                    vg.codigo_empresa_selecionada = codigo_empresa
                    vg.cnpj_empresa_selecionada = cnpj_empresa

                    print(f"✅ Empresa '{nome_empresa}' (Código: {codigo_empresa}) processada com sucesso!")

                    try:
                        verificar_pastas_fiscais(
                            navegador,
                            res["dados_json"],
                            codigo_empresa,
                            nome_empresa,
                            cnpj_empresa
                        )
                    except Exception as e:
                        print(f"⚠️ Erro ao verificar pastas fiscais para {nome_empresa} (Código: {codigo_empresa}): {e}")
                    time.sleep(2)
                else:
                    print(f"⚠ Erro ao selecionar empresa {codigo_empresa}: {res.get('erro', 'Erro desconhecido')}")

                # ♻️ Reset ao final da empresa
                if vg.nome_empresa_selecionada:
                    print(f"♻️ Resetando empresa '{vg.nome_empresa_selecionada}' (Código: {vg.codigo_empresa_selecionada})...")
                vg.nome_empresa_selecionada = None
                vg.codigo_empresa_selecionada = None
                vg.cnpj_empresa_selecionada = None

            print("\n✅ Todas as empresas foram processadas com sucesso!")
            return True

        except Exception as e:
            print(f"❌ Erro ao acessar a aba 'Documentos' na tentativa {tentativa}: {e}")
            if tentativa < max_tentativas:
                print("🔄 Recarregando a página e tentando novamente...")
                time.sleep(3)
                navegador.refresh()

    print("🚫 Não foi possível acessar a aba 'Documentos' após várias tentativas.")
    return False





# FUNC - FEITA DIA 24/07/2025
def aguardar_preloader(navegador, tempo=15):
    WebDriverWait(navegador, tempo).until(
        lambda d: d.find_element(By.CSS_SELECTOR, "div.preloader").get_attribute("innerHTML").strip() == ""
    )


def carregar_json_seguro(caminho):
    if not os.path.exists(caminho):
        return None, f"Arquivo '{caminho}' não encontrado."
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            return json.load(f), None
    except Exception as e:
        return None, f"Erro ao ler JSON: {e}"


def selecionar_empresa(navegador, codigo_empresa, tempo_max=20):
    """
    Localiza e seleciona a empresa na lista de sugestões do Onvio com base no código informado.
    Utiliza a estrutura <li> > <bento-combobox-row-template> > <span> para identificar e clicar.
    """

    def rolar_lista_completa(ul_element, navegador, delay_scroll=0.3):
        from selenium.webdriver.common.action_chains import ActionChains
        import time

        print("🔁 Rolando a lista até o final para carregar todas as empresas...")

        ultimo_altura = 0
        mesma_altura_repetida = 0

        while True:
            navegador.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", ul_element)
            time.sleep(delay_scroll)

            nova_altura = navegador.execute_script("return arguments[0].scrollHeight", ul_element)
            if nova_altura == ultimo_altura:
                mesma_altura_repetida += 1
                if mesma_altura_repetida >= 2:
                    break  # duas repetições iguais = fim do scroll
            else:
                mesma_altura_repetida = 0
                ultimo_altura = nova_altura


    def aguardar_lista_estavel(driver, timeout=10, intervalo=0.5):
        import time
        from datetime import datetime

        fim = datetime.now().timestamp() + timeout
        tamanho_anterior = -1

        while datetime.now().timestamp() < fim:
            uls = driver.find_elements(By.CSS_SELECTOR, "ul.bento-combobox-container-list")
            ul_element = next((ul for ul in uls if ul.is_displayed()), None)
            if not ul_element:
                continue
            li_elements = ul_element.find_elements(By.TAG_NAME, "li")
            tamanho_atual = len(li_elements)
            if tamanho_atual == tamanho_anterior:
                return ul_element, li_elements
            tamanho_anterior = tamanho_atual
            time.sleep(intervalo)
        raise Exception("Lista de empresas não estabilizou.")
    
    


    try:
        print("🟡 Aguardando campo de busca da empresa...")
        campo_input = WebDriverWait(navegador, tempo_max).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Selecione um cliente']"))
        )

        aguardar_preloader(navegador, tempo_max)
        print("✅ Preloader limpo. Focando no campo de busca...")

        # Verifica se já há conteúdo preenchido e limpa se necessário
        classe_input = campo_input.get_attribute("class")
        if "ng-not-empty" in classe_input:
            print("🔁 Campo já preenchido. Limpando valor atual...")
            campo_input.send_keys(Keys.CONTROL + "a")
            campo_input.send_keys(Keys.BACKSPACE)
            time.sleep(0.5)

        codigo_empresa = str(codigo_empresa).strip()
        campo_input.send_keys(codigo_empresa)
        print(f"🔍 Código {codigo_empresa} digitado.")
        aguardar_preloader(navegador, tempo_max)

        print("🟡 Aguardando a lista de empresas carregar e estabilizar...")
        ul_element, li_elements = aguardar_lista_estavel(navegador, timeout=tempo_max)
        aguardar_preloader(navegador, tempo_max)
        
        
        rolar_lista_completa(ul_element, navegador)  # 🔥 Aqui é o lugar certo!
        aguardar_preloader(navegador, tempo_max) 
        
        # Atualiza os <li> após rolagem
        ul_element, li_elements = aguardar_lista_estavel(navegador, timeout=tempo_max)# Garante carregamento dos novos <li>
        
    
        
        print(f"🔎 Total de itens na lista: {len(li_elements)}")
       
        for i, li in enumerate(li_elements, 1):
            try:
                template = li.find_element(By.TAG_NAME, "bento-combobox-row-template")
                spans = template.find_elements(By.TAG_NAME, "span")

                codigo_extraido = spans[0].text.strip() if len(spans) >= 1 else "N/A"
                nome_empresa = spans[1].text.strip() if len(spans) > 1 else "Desconhecido"

                print(f"Item {i}: Código='{codigo_extraido}', Nome='{nome_empresa}'")

                if codigo_extraido == codigo_empresa:
                    navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", li)
                    WebDriverWait(navegador, 5).until(EC.element_to_be_clickable((By.TAG_NAME, "li")))

                    try:
                        li.click()
                    except:
                        navegador.execute_script("arguments[0].click();", li)

                    vg.nome_empresa_selecionada = nome_empresa
                    vg.codigo_empresa_selecionada = codigo_empresa
                    print(f"   ➤ Nome:   {vg.nome_empresa_selecionada}")
                    print(f"   ➤ Código: {vg.codigo_empresa_selecionada}")
                    print(f"   ➤ CNPJ:   {vg.cnpj_empresa_selecionada}")
                    

                    dados_json, erro_json = carregar_json_seguro("parcelamentos_baixados.json")
                    if erro_json:
                        return {"status": False, "erro": erro_json}

                    return {
                        "status": True,
                        "nome_empresa": nome_empresa,
                        "codigo_empresa": codigo_empresa,
                        "dados_json": dados_json
                    }

            except Exception as e:
                print(f"⚠️ Erro ao processar item {i}: {e}")
                continue

        print(f"❌ Código {codigo_empresa} não encontrado na lista.")
        return {"status": False, "erro": f"Código {codigo_empresa} não encontrado na lista."}

    except Exception as e:
        print(f"❌ Erro ao selecionar empresa {codigo_empresa}: {e}")
        return {"status": False, "erro": str(e)}





#Verificando acessoa a pasta FISCAL 

# Mapeamento dos nomes internos para os nomes exibidos no Onvio
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Mapeamento do tipo_parcelamento → nome da pasta no Onvio
MAPPER_NOMES_PASTAS = {
    "FEDERAL_SIMPLIFICADO": "PARCELAMENTO SIMPLIFICADO",
    "PGFN": "PARCELAMENTO PGFN",
    "SIMPLES_NACIONAL": "PARCELAMENTO SIMPLES NACIONAL",
    "PREVIDENCIARIO": "PARCELAMENTO PREVIDENCIARIO",
    "NAO_PREVIDENCIARIO": "PARCELAMENTO NAO PREVIDENCIARIO"
}

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def acessar_pasta_fiscal(navegador, max_espera=30):
    try:
        print("🟡 Aguardando carregamento do painel esquerdo (<aside>)...")
        aside = WebDriverWait(navegador, max_espera).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "aside.bento-splitter-group-left"))
        )

        print("🟡 Aguardando elemento <bm-tree> dentro do painel esquerdo...")
        host = WebDriverWait(aside, max_espera).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "bm-tree"))
        )

        print("🧩 Obtendo shadowRoot do <bm-tree>...")
        shadow_root = navegador.execute_script("return arguments[0].shadowRoot", host)

        print("🔍 Buscando item com title='Fiscal' dentro do shadowRoot...")
        fiscal_item = WebDriverWait(shadow_root, max_espera).until(
            lambda d: d.find_element(By.CSS_SELECTOR, "bm-tree-item[title='Fiscal']")
        )

        href = fiscal_item.get_attribute("href")
        if href:
            print(f"🌐 Redirecionando para: {href}")
            navegador.get(href)

            print("🕓 Aguardando carregamento dos elementos da pasta (linhas ou aviso de pasta vazia)...")
            WebDriverWait(navegador, max_espera).until(
                EC.any_of(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "li.paginate_info")),
                    EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'A pasta selecionada está vazia.')]"))
                )
            )

            print("✅ Acesso à pasta Fiscal realizado com sucesso.")
            return True
        else:
            print("❌ Atributo href não encontrado no item Fiscal.")
            return False

    except TimeoutException as e:
        print(f"❌ Tempo excedido ao aguardar os elementos da pasta Fiscal: {e}")
        return False
    except Exception as e:
        print(f"❌ Erro ao acessar a pasta Fiscal: {e}")
        return False




import json
import os
import re

def verificar_pastas_fiscais(navegador, dados_json, codigo_empresa, nome_empresa, max_espera=20):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import time
    import json
    import os

    def esperar_paginate_info_e_tabela(navegador, timeout=20):
        WebDriverWait(navegador, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "li.paginate_info"))
        )
        fim = time.time() + timeout
        while time.time() < fim:
            linhas = navegador.find_elements(By.CSS_SELECTOR, 'div.wj-cells[wj-part="cells"] div.wj-row')
            if linhas:
                return linhas
            time.sleep(0.5)
        return []

    def criar_pasta_parcelamentos(navegador, max_espera=20):
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.common.keys import Keys

        try:
            print("🟢 Criando a pasta 'PARCELAMENTOS'...")

            botao_novo = WebDriverWait(navegador, max_espera).until(
                EC.element_to_be_clickable((By.ID, "dms-fe-legacy-components-client-documents-new-menu-button"))
            )
            botao_novo.click()

            botao_pasta = WebDriverWait(navegador, max_espera).until(
                EC.element_to_be_clickable((By.ID, "dms-fe-legacy-components-client-documents-new-folder-button"))
            )
            botao_pasta.click()

            campo_nome = WebDriverWait(navegador, max_espera).until(
                EC.presence_of_element_located((By.ID, "containerName"))
            )
            campo_nome.clear()
            campo_nome.send_keys("PARCELAMENTOS")

            botao_salvar = WebDriverWait(navegador, max_espera).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-qe-id='dmsNewContainerModal-saveButton']"))
            )
            botao_salvar.click()

            # Confirmação da criação
            WebDriverWait(navegador, max_espera).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'div.bottom-alerts-pane div[ng-if="operations.length === 1"]')
                )
            )
            print("✅ Pasta 'PARCELAMENTOS' criada com sucesso.")
            return True

        except Exception as e:
            print(f"❌ Erro ao criar pasta: {e}")
            return False

    def executar_verificacao():
        print("🔍 Verificando pastas dentro da aba Fiscal...\n")
        nomes_visiveis = []
        sucesso = acessar_pasta_fiscal(navegador, max_espera)
        if not sucesso:
            print("❌ Não foi possível acessar a aba Fiscal.")
            return []

        aguardar_preloader(navegador, max_espera)

        try:
            linhas = esperar_paginate_info_e_tabela(navegador, timeout=max_espera)
            for linha in linhas:
                celulas = linha.find_elements(By.CSS_SELECTOR, 'div.wj-cell[aria-colindex="2"] a')
                for cel in celulas:
                    texto = cel.text.strip()
                    if texto:
                        nomes_visiveis.append(texto)

        except Exception:
            print("⚠️ Pasta Fiscal vazia (elementos não renderizados).\n")

        print("📂 Pastas visíveis na aba Fiscal:")
        try:
            msg_vazia = navegador.find_elements(By.XPATH, "//div[contains(text(),'A pasta selecionada está vazia.')]")
            if msg_vazia:
                print("   ⚠️ Nenhuma pasta visível (mensagem do sistema: Pasta vazia).")
            elif nomes_visiveis:
                for nome in nomes_visiveis:
                    print(f"   ➤ {nome}")
            else:
                print("   ⚠️ Nenhuma pasta visível (sem mensagem do sistema).")
        except Exception as e:
            print(f"   ⚠️ Erro ao verificar visibilidade das pastas: {e}")

        return nomes_visiveis

    # Executa a primeira verificação
    nomes_visiveis = executar_verificacao()

    # Verifica se a pasta "PARCELAMENTOS" já existe
    ja_existe_pasta_parcelamentos = "PARCELAMENTOS" in [n.upper() for n in nomes_visiveis]

    # Cria se não existir
    if not ja_existe_pasta_parcelamentos:
        print("📦 Pasta 'PARCELAMENTOS' não encontrada. Criando agora...")
        sucesso_criacao = criar_pasta_parcelamentos(navegador, max_espera)
        if sucesso_criacao:
            print("🔁 Recarregando a página para nova verificação...")
            navegador.refresh()
            time.sleep(3)
            aguardar_preloader(navegador, max_espera)
            nomes_visiveis = executar_verificacao()

    # Salva resultado
    caminho_json = "pastas_fiscais.json"
    if os.path.exists(caminho_json):
        with open(caminho_json, "r", encoding="utf-8") as f:
            dados_totais = json.load(f)
    else:
        dados_totais = []

    dados_totais = [d for d in dados_totais if d["codigo_empresa"] != codigo_empresa]
    dados_totais.append({
        "codigo_empresa": codigo_empresa,
        "nome_empresa": nome_empresa,
        "pastas_visiveis": nomes_visiveis,
        "pasta_parcelamentos_ja_existia": ja_existe_pasta_parcelamentos
    })

    with open(caminho_json, "w", encoding="utf-8") as f:
        json.dump(dados_totais, f, indent=2, ensure_ascii=False)

    print(f"\n💾 Resultado salvo em '{caminho_json}' com sucesso.")




import json
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def criar_subpastas_parcelamentos_por_cnpj(navegador, max_espera=20):
    """
    Cria subpastas dentro da pasta 'PARCELAMENTOS' com base no tipo de parcelamento
    associado ao CNPJ selecionado, definido em vg.cnpj_empresa_selecionada.
    """

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.keys import Keys
    import json
    import time
    import vg  # <- usa as variáveis globais definidas no módulo

    caminho_json = "parcelamentos_baixados.json"

    if not vg.cnpj_empresa_selecionada:
        print("⚠️ CNPJ da empresa selecionada não está definido em vg.cnpj_empresa_selecionada.")
        return

    try:
        with open(caminho_json, "r", encoding="utf-8") as f:
            dados = json.load(f)
    except Exception as e:
        print(f"❌ Erro ao ler o arquivo JSON: {e}")
        return

    dados_filtrados = [item for item in dados if item["cnpj"] == vg.cnpj_empresa_selecionada]
    if not dados_filtrados:
        print(f"⚠️ Nenhum dado encontrado para o CNPJ {vg.cnpj_empresa_selecionada}")
        return

    tipos_parcelamento = sorted(set(item["tipo_parcelamento"].strip().upper() for item in dados_filtrados))
    if not tipos_parcelamento:
        print("⚠️ Nenhum tipo de parcelamento identificado no JSON.")
        return

    print(f"📁 Criando subpastas para o CNPJ {vg.cnpj_empresa_selecionada}...\n")
    print(f"📌 Subpastas a serem criadas: {tipos_parcelamento}")

    try:
        pasta_parcelamentos = WebDriverWait(navegador, max_espera).until(
            EC.element_to_be_clickable((By.XPATH, f"//a[contains(text(), 'PARCELAMENTOS')]"))
        )
        pasta_parcelamentos.click()
        time.sleep(2)
    except Exception as e:
        print(f"❌ Não foi possível acessar a pasta 'PARCELAMENTOS': {e}")
        return

    for tipo in tipos_parcelamento:
        try:
            print(f"🛠️ Criando subpasta: {tipo}")

            botao_novo = WebDriverWait(navegador, max_espera).until(
                EC.element_to_be_clickable((By.ID, "dms-fe-legacy-components-client-documents-new-menu-button"))
            )
            botao_novo.click()

            botao_pasta = WebDriverWait(navegador, max_espera).until(
                EC.element_to_be_clickable((By.ID, "dms-fe-legacy-components-client-documents-new-folder-button"))
            )
            botao_pasta.click()

            campo_nome = WebDriverWait(navegador, max_espera).until(
                EC.presence_of_element_located((By.ID, "containerName"))
            )
            campo_nome.clear()
            campo_nome.send_keys(tipo)

            botao_salvar = WebDriverWait(navegador, max_espera).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-qe-id='dmsNewContainerModal-saveButton']"))
            )
            botao_salvar.click()

            WebDriverWait(navegador, max_espera).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'div.bottom-alerts-pane div[ng-if="operations.length === 1"]')
                )
            )
            print(f"✅ Subpasta '{tipo}' criada com sucesso.")

        except Exception as e:
            print(f"❌ Erro ao criar subpasta '{tipo}': {e}")

    print("\n🏁 Finalizado o processo de criação de subpastas.\n")




    
#########################################################################################
def executar_automacao_onvio():
    try:
        registrar_execucao("🚀 Iniciando a automação no Onvio...")

        # Iniciar navegador
        navegador = iniciar_navegador()

        # Acessar site
        navegador.get(URL_DESTINO)
        time.sleep(5)

        # Login
        entrar_no_portal(navegador, USUARIO, SENHA)

        # Caminho para a planilha com CNPJ e código
        caminho_planilha_codigos = os.path.join(os.path.dirname(__file__), "Relação Empresas - Nome - CNPJ.xls")

        # Acessar aba de documentos e processar empresas
        sucesso = acessar_aba_documentos(navegador, caminho_planilha_codigos)

        if sucesso:
            registrar_execucao("✅ Processo finalizado com sucesso.")
        else:
            registrar_execucao("⚠️ Processo finalizado com falhas.")

    except Exception as e:
        registrar_execucao(f"❌ Erro geral na automação: {e}")

    finally:
        try:
            navegador.quit()
            registrar_execucao("🧹 Navegador fechado.")
        except:
            pass
         
if __name__ == "__main__":
    executar_automacao_onvio()
         