# Versão: 3.8 - Inserção de data com JavaScript (sem send_keys)
# --- BIBLIOTECAS NECESSÁRIAS ---
# pip install selenium webdriver-manager pandas openpyxl pyautoit pygetwindow python-dotenv

import os
import sys
import json
import time
import logging
from datetime import datetime
import autoit
import pygetwindow as gw
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
load_dotenv()

try:
    from email_verificacao import extrair_codigo_do_email
except ImportError:
    print("AVISO: Módulo 'email_verificacao' não encontrado. A autenticação de dois fatores pode falhar.")
    def extrair_codigo_do_email():
        return input("Por favor, insira o código MFA recebido: ")


# --- CONFIGURAÇÃO ---
CONFIG = {
    "urls": {
        "login": "https://onvio.com.br/#/",
        "documentos": "https://onvio.com.br/staff/#/documents",
        "api_notes": "https://onvio.com.br/api/notes/v1/comments/search",
    },
    "credenciais": {
        "usuario": os.getenv("ONVIO_USER"),
        "senha": os.getenv("ONVIO_PASSWORD"),
    },
    "arquivos": {
        "arquivo_de_entrada": "download.json",
        "log_execucao": "log_envio_boletos.txt",
        "resultados_upload": "resultados_upload_boletos.json",
        "caminho_local_dos_boletos": "temp",
    },
    "pastas": {
        "principal": "Financeiro",
        "subpasta": "Boletos",
    },
    "seletores": {
        # Login & MFA
        "login_continuar_btn": (By.ID, "trauth-continue-signin-btn"),
        "login_usuario_input": (By.NAME, "username"),
        "login_senha_input": (By.ID, "password"),
        "login_submit_btn": (By.XPATH, "//*[@type='submit']"),
        "mfa_titulo": (By.XPATH, "//h1[contains(text(), 'verificar sua identidade')]"),
        "mfa_email_btn": (By.XPATH, "//button[@name='action' and contains(@value, 'email')]"),
        "mfa_codigo_input": (By.ID, "code"),
        "mfa_continuar_btn": (By.XPATH, "//button[@name='action' and @value='default']"),
        "dashboard_content": (By.CLASS_NAME, "dashboard-section__content"),
        
        # Seletores da Área de Documentos
        "docs_pagina_confirmacao": (By.XPATH, "//ul[contains(@class, 'nav-tabs')]//span[contains(text(), 'Documentos do Cliente')]"),
        "docs_combobox_clicavel": (By.CSS_SELECTOR, ".clients-combobox .bento-combobox"),
        "docs_selecionar_cliente_input": (By.CSS_SELECTOR, "input[placeholder='Selecione um cliente']"),
        "docs_container_rolagem": (By.CLASS_NAME, "bento-combobox-container-body"),
        "financeiro_aside_panel": (By.CSS_SELECTOR, "aside.bento-splitter-group-left"),
        "financeiro_bm_tree": (By.CSS_SELECTOR, "bm-tree"),
        "financeiro_tree_item_shadow": (By.CSS_SELECTOR, "bm-tree-item[title='Financeiro']"),
        "tree_item_loading": (By.CSS_SELECTOR, "bm-tree-item[title='Carregando...']"),
        "folder_grid_ready": (By.CSS_SELECTOR, "li.paginate_info"),
        "folder_grid_empty": (By.XPATH, "//div[contains(text(), 'A pasta selecionada está vazia.')]"),
        
        # Botões e Modais de Criação/Upload
        "novo_menu_btn": (By.ID, "dms-fe-legacy-components-client-documents-new-menu-button"),
        "nova_pasta_btn": (By.ID, "dms-fe-legacy-components-client-documents-new-folder-button"),
        "nova_pasta_nome_input": (By.ID, "containerName"),
        "nova_pasta_salvar_btn": (By.CSS_SELECTOR, "button[data-qe-id='dmsNewContainerModal-saveButton']"),
        "upload_button_geral": (By.ID, "dms-fe-legacy-components-client-documents-upload-button"),
        "voltar_nivel_btn": (By.XPATH, "//a[i[@class='dms-icon-up-one-level']]"),

        # Confirmação e Alertas
        "confirmacao_operacao_div": (By.CSS_SELECTOR, 'div.bottom-alerts-pane div[ng-if="operations.length === 1"]'),
        "upload_carregando_msg": (By.CSS_SELECTOR, "span.file-alert-text"),
        "popup_erro_texto": (By.CSS_SELECTOR, "div.alert-error .file-alert-text"),
        "popup_erro_fechar_btn": (By.CSS_SELECTOR, "button.bento-alert-close"),

        # --- SELETORES PARA DEFINIÇÃO DE DATA DE VENCIMENTO ---
        "file_selection_checkbox_xpath": (By.XPATH, "//div[@wj-part='rh']//div[@class='wj-row' and @aria-rowindex='{row_index}']//i[contains(@class, 'bento-flex-grid-checkbox')]"),
        "manage_menu_btn": (By.ID, "dms-fe-legacy-components-client-documents-manage-docs-menu-button"),
        "manage_menu_li_parent_open": (By.CSS_SELECTOR, "li[data-button-id='3'].open"),
        "set_due_date_option": (By.XPATH, "//ul[@ng-if='button.dropdown']//a[normalize-space()='Definir data de vencimento']"),
        "due_date_modal": (By.CSS_SELECTOR, "div.modal-content"),
        "due_date_modal_input": (By.CSS_SELECTOR, "input[data-qe-id='dueDate-date']"),
        "due_date_modal_save_btn": (By.XPATH, "//button[contains(@class, 'btn-primary') and contains(text(), 'Salvar')]"),
    },
}

def setup_logger():
    # ... (código do logger inalterado) ...
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_path = os.path.join(script_dir, CONFIG["arquivos"]["log_execucao"])
    logger = logging.getLogger("OnvioAutomator")
    logger.setLevel(logging.INFO)
    if logger.hasHandlers(): logger.handlers.clear()
    file_handler = logging.FileHandler(log_path, encoding="utf-8", mode='w')
    file_formatter = logging.Formatter("[%(asctime)s] [%(levelname)s] - %(message)s", "%Y-%m-%d %H:%M:%S")
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)
    stream_handler = logging.StreamHandler(sys.stdout)
    stream_formatter = logging.Formatter("[%(asctime)s] - %(message)s", "%H:%M:%S")
    stream_handler.setFormatter(stream_formatter)
    logger.addHandler(stream_handler)
    return logger

logger = setup_logger()

# --- PROCESSAMENTO ---
class DataProcessor:
    # ... (código da classe DataProcessor inalterado) ...
    def __init__(self, config):
        self.config = config
        script_dir = os.path.dirname(os.path.abspath(__file__))
        nome_arquivo_entrada = self.config["arquivos"]["arquivo_de_entrada"]
        self.caminho_arquivo_entrada = os.path.join(script_dir, nome_arquivo_entrada)
        nome_pasta_boletos = self.config["arquivos"]["caminho_local_dos_boletos"]
        self.pasta_local_boletos = os.path.join(script_dir, nome_pasta_boletos)
    def carregar_dados_de_entrada(self):
        logger.info(f"Procurando arquivo de entrada em: {self.caminho_arquivo_entrada}")
        if not os.path.exists(self.caminho_arquivo_entrada):
            logger.error(f"Arquivo de entrada '{os.path.basename(self.caminho_arquivo_entrada)}' não encontrado na pasta do script!")
            return []
        logger.info(f"Procurando pasta de boletos em: {self.pasta_local_boletos}")
        if not os.path.isdir(self.pasta_local_boletos):
            logger.error(f"A pasta de boletos '{os.path.basename(self.pasta_local_boletos)}' não foi encontrada no diretório do script.")
            return []
        with open(self.caminho_arquivo_entrada, "r", encoding="utf-8") as f:
            dados_brutos = json.load(f)
        boletos_para_processar = []
        boletos_validos = dados_brutos.get("boletos", [])
        logger.info(f"Encontrados {len(boletos_validos)} registros na seção 'boletos' do JSON.")
        for boleto in boletos_validos:
            codigo_empresa = boleto.get("codigo_empresa_dominio")
            nome_pdf = boleto.get("nome_pdf")
            data_vencimento = boleto.get("data_vencimento")
            if codigo_empresa and nome_pdf and str(codigo_empresa).upper() != "NÃO ENCONTRADO":
                caminho_completo = os.path.join(self.pasta_local_boletos, nome_pdf)
                if os.path.exists(caminho_completo):
                    boletos_para_processar.append({
                        "codigo_empresa": str(codigo_empresa),
                        "nome_empresa": boleto.get("nome_empresa"),
                        "nome_arquivo": nome_pdf,
                        "caminho_completo_arquivo": caminho_completo,
                        "data_vencimento": data_vencimento
                    })
                else:
                    logger.warning(f"PDF '{nome_pdf}' para empresa '{codigo_empresa}' não encontrado na pasta 'temp'. Pulando item.")
            else:
                logger.warning(f"Registro inválido ou sem código de empresa no JSON: {boleto}. Pulando item.")
        logger.info(f"Total de {len(boletos_para_processar)} boletos válidos e com arquivos encontrados para processamento.")
        return boletos_para_processar

# --- AUTOMAÇÃO ---
class  OnvioAutomator:
    # ... (código de __init__ até _salvar_resultado inalterado) ...
    def __init__(self, config):
        self.config = config
        self.logger = logger
        self.driver = self._iniciar_navegador()
        if not self.driver:
            raise Exception("Falha na inicialização do WebDriver.")
        self.item_atual_info = {}
    def _iniciar_navegador(self):
        self.logger.info("Iniciando navegador Chrome...")
        options = webdriver.ChromeOptions()
        options.add_experimental_option("prefs", {"plugins.always_open_pdf_externally": True})
        options.add_argument("--start-maximized"); options.add_argument("--no-sandbox"); options.add_argument("--disable-dev-shm-usage")
        try:
            servico = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=servico, options=options)
            driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"})
            return driver
        except Exception as e:
            self.logger.error(f"Falha ao iniciar o navegador: {e}")
            return None
    def _aguardar_elemento(self, seletor, tempo=20):
        return WebDriverWait(self.driver, tempo).until(EC.element_to_be_clickable(seletor))
    def _wait_for_navigation_panel_to_load(self, tempo=20):
        self.logger.info("Aguardando painel de navegação lateral estabilizar...")
        try:
            WebDriverWait(self.driver, tempo).until(EC.invisibility_of_element_located(self.config["seletores"]["tree_item_loading"]))
            self.logger.info("Painel de navegação estabilizado.")
            return True
        except TimeoutException:
            self.logger.warning("Painel de navegação não estabilizou no tempo esperado.")
            return False
    def _wait_for_api_request(self, url_api, status_code=202, settle_time=3, timeout=30):
        self.logger.info(f"Aguardando API '{os.path.basename(url_api)}' estabilizar com status {status_code}...")
        script = f"return window.performance.getEntriesByType('resource').filter(req => req.name.startsWith('{url_api}') && req.responseStatus === {status_code}).length;"
        start_time = time.time()
        last_count = -1
        while time.time() - start_time < timeout:
            try:
                current_count = self.driver.execute_script(script)
                if last_count != -1 and current_count == last_count:
                    break
                last_count = current_count
                time.sleep(settle_time)
            except Exception:
                time.sleep(settle_time)
        final_count = self.driver.execute_script(script)
        if final_count > 0:
            self.logger.info(f"Total de {final_count} requisição(ões) com status {status_code} encontradas.")
        else:
            self.logger.warning(f"Nenhuma requisição para '{os.path.basename(url_api)}' com status {status_code} foi concluída.")
    def _wait_for_grid_to_load(self, tempo=20):
        self.logger.info("Aguardando o conteúdo do grid carregar (API, Painel e UI)...")
        try:
            self._wait_for_api_request(self.config["urls"]["api_notes"], status_code=202)
            self._wait_for_navigation_panel_to_load()
            WebDriverWait(self.driver, tempo).until(
                EC.any_of(
                    EC.presence_of_element_located(self.config["seletores"]["folder_grid_ready"]),
                    EC.presence_of_element_located(self.config["seletores"]["folder_grid_empty"])
                )
            )
            self.logger.info("Grid carregado e estável.")
            return True
        except Exception as e:
            self.logger.error(f"Falha na espera pelo grid: {e}")
            return False
    def fazer_login(self):
        self.logger.info("Iniciando processo de login.")
        if not self.config["credenciais"]["usuario"] or not self.config["credenciais"]["senha"]:
            self.logger.error("Credenciais de usuário ou senha não encontradas. Verifique seu arquivo .env ou a configuração.")
            return False
        self.driver.get(self.config["urls"]["login"])
        try:
            try: WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable(self.config["seletores"]["login_continuar_btn"])).click()
            except TimeoutException: self.logger.info("Botão 'Continuar' inicial não foi necessário.")
            self._aguardar_elemento(self.config["seletores"]["login_usuario_input"]).send_keys(self.config["credenciais"]["usuario"])
            self._aguardar_elemento(self.config["seletores"]["login_submit_btn"]).click()
            self._aguardar_elemento(self.config["seletores"]["login_senha_input"]).send_keys(self.config["credenciais"]["senha"])
            self._aguardar_elemento(self.config["seletores"]["login_submit_btn"]).click()
            self.logger.info("Login e senha enviados.")
            return self._handle_mfa()
        except Exception as e:
            self.logger.error(f"Erro inesperado durante o login: {e}")
            return False
    def _handle_mfa(self):
        try:
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(self.config["seletores"]["mfa_titulo"]))
            self.logger.info("Tela de MFA detectada.")
            self._aguardar_elemento(self.config["seletores"]["mfa_email_btn"]).click()
            self.logger.info("Aguardando 20s pelo código do e-mail...")
            time.sleep(20)
            codigo = extrair_codigo_do_email()
            if not codigo: raise Exception("Código de verificação não foi extraído.")
            self._aguardar_elemento(self.config["seletores"]["mfa_codigo_input"]).send_keys(codigo)
            self._aguardar_elemento(self.config["seletores"]["mfa_continuar_btn"]).click()
            self.logger.info(f"Código MFA inserido com sucesso.")
        except TimeoutException:
            self.logger.info("Nenhuma tela de MFA foi solicitada.")
        except Exception as e:
            self.logger.error(f"Falha fatal durante o processo de MFA: {e}")
            return False
        return True
    def _navegar_para_documentos(self):
        self.logger.info(f"Navegando para a área de Documentos: {self.config['urls']['documentos']}")
        try:
            WebDriverWait(self.driver, 30).until(EC.presence_of_element_located(self.config["seletores"]["dashboard_content"]))
            self.logger.info("Dashboard carregado. Redirecionando...")
            self.driver.get(self.config['urls']['documentos'])
            WebDriverWait(self.driver, 30).until(EC.presence_of_element_located(self.config["seletores"]["docs_pagina_confirmacao"]))
            self.logger.info("Página de Documentos carregada com sucesso!")
            return True
        except TimeoutException:
            self.logger.error("Não foi possível carregar a página de documentos.")
            return False
    def selecionar_empresa(self, codigo_empresa):
        self.logger.info(f"Iniciando seleção (v3.2) para empresa: {codigo_empresa}")
        try:
            self.logger.info("Ativando o combobox de cliente...")
            combobox_container = self._aguardar_elemento(self.config["seletores"]["docs_combobox_clicavel"])
            combobox_container.click()
            time.sleep(1)
            self.logger.info("Encontrando o campo de busca...")
            campo_busca = self._aguardar_elemento(self.config["seletores"]["docs_selecionar_cliente_input"])
            self.logger.info("Limpando o campo de busca...")
            campo_busca.send_keys(Keys.CONTROL + "a", Keys.BACKSPACE)
            time.sleep(0.5)
            self.logger.info(f"Digitando código '{codigo_empresa}'...")
            campo_busca.send_keys(codigo_empresa)
            container_de_rolagem = WebDriverWait(self.driver, 15).until(
                EC.visibility_of_element_located(self.config["seletores"]["docs_container_rolagem"])
            )
            self.logger.info("Lista de resultados carregada.")
            time.sleep(1)
            max_tentativas_rolagem = 30
            for tentativa in range(max_tentativas_rolagem):
                xpath_clicavel = f".//li[.//span[1][text()='{codigo_empresa}']]//bento-combobox-row-template"
                try:
                    item_para_clicar = container_de_rolagem.find_element(By.XPATH, xpath_clicavel)
                    self.logger.info(f"Empresa com código '{codigo_empresa}' encontrada! Clicando...")
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", item_para_clicar)
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click();", item_para_clicar)
                    return self._wait_for_grid_to_load()
                except NoSuchElementException:
                    if tentativa < max_tentativas_rolagem - 1:
                        self.logger.info(f"Tentativa {tentativa + 1}: Não visível. Rolando a lista...")
                        self.driver.execute_script("arguments[0].scrollTop += 150;", container_de_rolagem)
                        time.sleep(0.4)
            self.logger.error(f"Empresa '{codigo_empresa}' não encontrada após {max_tentativas_rolagem} tentativas de rolagem.")
            return False
        except Exception as e:
            self.logger.error(f"Ocorreu um erro inesperado durante a seleção da empresa: {e}", exc_info=True)
            return False
    def _acessar_pasta_principal(self, nome_pasta):
        self.logger.info(f"Acessando a pasta principal '{nome_pasta}' via painel lateral (Shadow DOM)...")
        try:
            aside = self._aguardar_elemento(self.config["seletores"]["financeiro_aside_panel"])
            host = WebDriverWait(aside, 20).until(EC.presence_of_element_located(self.config["seletores"]["financeiro_bm_tree"]))
            shadow_root = self.driver.execute_script("return arguments[0].shadowRoot", host)
            seletor_pasta_shadow = (By.CSS_SELECTOR, f"bm-tree-item[title='{nome_pasta}']")
            pasta_item = WebDriverWait(shadow_root, 20).until(lambda d: d.find_element(*seletor_pasta_shadow))
            href = pasta_item.get_attribute("href")
            if href:
                self.logger.info(f"Navegando para o link da pasta: {href}")
                self.driver.get(href)
            else:
                self.logger.warning("Atributo 'href' não encontrado. Tentando clicar diretamente.")
                pasta_item.click()
            return self._wait_for_grid_to_load()
        except Exception as e:
            self.logger.error(f"Erro ao acessar a pasta '{nome_pasta}' via Shadow DOM: {e}", exc_info=True)
            return False
    def _navegar_para_subpasta(self, nome_subpasta):
        self.logger.info(f"Entrando na subpasta '{nome_subpasta}'...")
        try:
            xpath_link_subpasta = f"//dms-grid-text-cell[@text='{nome_subpasta}']//a"
            link_subpasta = self._aguardar_elemento((By.XPATH, xpath_link_subpasta))
            link_subpasta.click()
            return self._wait_for_grid_to_load()
        except Exception as e:
            self.logger.error(f"Erro ao entrar na subpasta '{nome_subpasta}': {e}")
            return False
    def _item_existe_no_grid(self, nome_item):
        self.logger.info(f"Verificando no grid a existência de '{nome_item}'...")
        try:
            xpath_preciso = f"//dms-grid-text-cell[@text='{nome_item}']"
            self.driver.find_element(By.XPATH, xpath_preciso)
            self.logger.info(f"Item '{nome_item}' encontrado no grid.")
            return True
        except NoSuchElementException:
            self.logger.info(f"Item '{nome_item}' NÃO encontrado no grid.")
            return False
    def _criar_pasta(self, nome_pasta):
        self.logger.info(f"Iniciando criação da pasta '{nome_pasta}'...")
        try:
            self._aguardar_elemento(self.config["seletores"]["novo_menu_btn"]).click()
            self._aguardar_elemento(self.config["seletores"]["nova_pasta_btn"]).click()
            self._aguardar_elemento(self.config["seletores"]["nova_pasta_nome_input"]).send_keys(nome_pasta)
            self._aguardar_elemento(self.config["seletores"]["nova_pasta_salvar_btn"]).click()
            self.logger.info("Aguardando confirmação da operação...")
            WebDriverWait(self.driver, 20).until(EC.presence_of_element_located(self.config["seletores"]["confirmacao_operacao_div"]))
            WebDriverWait(self.driver, 10).until(EC.invisibility_of_element_located(self.config["seletores"]["confirmacao_operacao_div"]))
            self.logger.info("Verificação pós-criação: Recarregando a página...")
            self.driver.refresh()
            if not self._wait_for_grid_to_load():
                self.logger.error("Falha ao recarregar a grade de pastas após a criação.")
                return False
            if self._item_existe_no_grid(nome_pasta):
                self.logger.info(f"VERIFICADO: Pasta '{nome_pasta}' existe no grid após recarregar.")
                return True
            else:
                self.logger.error(f"FALHA NA VERIFICAÇÃO: Pasta '{nome_pasta}' não foi encontrada após recarregar.")
                return False
        except Exception as e:
            self.logger.error(f"Falha durante o processo de criação da pasta '{nome_pasta}': {e}", exc_info=True)
            return False
    def _upload_arquivo(self, caminho_arquivo):
        if not os.path.exists(caminho_arquivo):
            self.logger.error(f"Arquivo para upload não encontrado em: {caminho_arquivo}")
            return False
        try:
            self.logger.info(f"Iniciando upload do arquivo: {os.path.basename(caminho_arquivo)}")
            self._aguardar_elemento(self.config["seletores"]["upload_button_geral"]).click()
            time.sleep(2)
            janela_titulo = "Abrir"
            if not autoit.win_wait(janela_titulo, timeout=10):
                self.logger.error(f"Janela de upload '{janela_titulo}' não foi encontrada!")
                return False
            self.logger.info(f"Janela '{janela_titulo}' detectada. Enviando caminho do arquivo...")
            autoit.win_activate(janela_titulo)
            autoit.control_set_text(janela_titulo, "Edit1", f'"{caminho_arquivo}"')
            time.sleep(1)
            autoit.control_click(janela_titulo, "Button1")
            self.logger.info("Aguardando progresso do upload (mensagem de carregamento)...")
            WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(self.config["seletores"]["upload_carregando_msg"]))
            self.logger.info("Progresso detectado. Aguardando finalização...")
            WebDriverWait(self.driver, 300).until(EC.invisibility_of_element_located(self.config["seletores"]["upload_carregando_msg"]))
            self.logger.info("Mensagem de progresso desapareceu.")
            return self._verificar_e_registrar_resultado_upload(os.path.basename(caminho_arquivo))
        except Exception as e:
            self.logger.error(f"Ocorreu um erro inesperado durante o upload: {e}", exc_info=True)
            if 'janela_titulo' in locals() and autoit.win_exists(janela_titulo):
                autoit.win_close(janela_titulo)
            return False
    def _verificar_e_registrar_resultado_upload(self, nome_arquivo):
        resultado_info = {
            "codigo_empresa": self.item_atual_info.get("codigo_empresa"),
            "nome_empresa": self.item_atual_info.get("nome_empresa"),
            "arquivo_enviado": nome_arquivo,
            "timestamp": datetime.now().isoformat()
        }
        try:
            popup_erro = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(self.config["seletores"]["popup_erro_texto"]))
            mensagem_erro = popup_erro.text.strip()
            self.logger.error(f"Erro de upload detectado pelo Onvio: {mensagem_erro}")
            resultado_info.update({"status": "Erro", "detalhes": mensagem_erro})
            try:
                self.driver.find_element(*self.config["seletores"]["popup_erro_fechar_btn"]).click()
            except Exception: pass
            self._salvar_resultado(resultado_info)
            return False
        except TimeoutException:
            self.logger.info(f"Upload do arquivo '{nome_arquivo}' concluído com sucesso.")
            resultado_info.update({"status": "Sucesso", "detalhes": "Arquivo enviado com sucesso."})
            self._salvar_resultado(resultado_info)
            return True
    def _salvar_resultado(self, resultado_info):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        caminho_json = os.path.join(script_dir, self.config["arquivos"]["resultados_upload"])
        try:
            dados = []
            if os.path.exists(caminho_json):
                with open(caminho_json, "r", encoding="utf-8") as f:
                    dados = json.load(f)
            dados.append(resultado_info)
            with open(caminho_json, "w", encoding="utf-8") as f:
                json.dump(dados, f, indent=4, ensure_ascii=False)
        except Exception as e:
            self.logger.error(f"Falha ao salvar o resultado no arquivo JSON: {e}")

    def _selecionar_checkbox_do_arquivo(self, nome_arquivo):
        self.logger.info(f"Tentando selecionar o checkbox para o arquivo '{nome_arquivo}'.")
        try:
            self.logger.info("Localizando a linha do arquivo para obter seu índice...")
            celula_arquivo_xpath = f"//dms-grid-text-cell[@text='{nome_arquivo}']/ancestor::div[contains(@class, 'wj-row')]"
            linha_do_arquivo = self._aguardar_elemento((By.XPATH, celula_arquivo_xpath), 30)
            
            row_index = linha_do_arquivo.get_attribute("aria-rowindex")
            if not row_index:
                self.logger.error("Não foi possível encontrar o 'aria-rowindex' para a linha do arquivo.")
                return False
            
            self.logger.info(f"Índice da linha encontrado: {row_index}.")
            seletor_checkbox_dinamico = self.config["seletores"]["file_selection_checkbox_xpath"][1].format(row_index=row_index)
            checkbox_para_clicar = self._aguardar_elemento((By.XPATH, seletor_checkbox_dinamico))
            
            checkbox_para_clicar.click()
            self.logger.info(f"Checkbox para '{nome_arquivo}' selecionado com sucesso.")
            time.sleep(1)
            return True
        except Exception as e:
            self.logger.error(f"Ocorreu um erro ao tentar selecionar o checkbox do arquivo: {e}", exc_info=True)
            return False

    def _set_input_value_with_js(self, seletor, valor):
        """
        Define o valor de um campo de input usando JavaScript e dispara o evento 'input'.
        Isso é mais robusto para frameworks como AngularJS.
        """
        self.logger.info(f"Definindo valor '{valor}' no campo via JavaScript...")
        try:
            elemento = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(seletor)
            )
            script = """
            var element = arguments[0];
            var value = arguments[1];
            element.value = value;
            element.dispatchEvent(new Event('input', { bubbles: true }));
            element.dispatchEvent(new Event('change', { bubbles: true }));
            """
            self.driver.execute_script(script, elemento, valor)
            self.logger.info("Valor definido com sucesso via JavaScript.")
            return True
        except Exception as e:
            self.logger.error(f"Falha ao definir valor com JavaScript: {e}")
            return False

    def _definir_data_vencimento_via_menu_gerenciar(self, nome_arquivo, data_vencimento):
        self.logger.info(f"Iniciando definição de data de vencimento para '{nome_arquivo}' via menu 'Gerenciar'.")
        try:
            self.logger.info("Aguardando a UI estabilizar...")
            if not self._wait_for_grid_to_load(): return False
            if not self._selecionar_checkbox_do_arquivo(nome_arquivo): return False
            
            self.logger.info("Clicando no botão 'Gerenciar'...")
            self._aguardar_elemento(self.config["seletores"]["manage_menu_btn"]).click()
            
            self.logger.info("Aguardando o menu dropdown ser renderizado e aberto...")
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(self.config["seletores"]["manage_menu_li_parent_open"]))
            self.logger.info("Menu 'Gerenciar' confirmado como aberto.")
            
            self.logger.info("Localizando e clicando em 'Definir data de vencimento' com JavaScript...")
            seletor_item = (By.XPATH, "//ul[@ng-if='button.dropdown']//a[normalize-space()='Definir data de vencimento']")
            item_para_clicar = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(seletor_item))
            self.driver.execute_script("arguments[0].click();", item_para_clicar)

            self.logger.info("Aguardando o pop-up de definição de data...")
            WebDriverWait(self.driver, 20).until(EC.visibility_of_element_located(self.config["seletores"]["due_date_modal"]))
            
            self.logger.info("Usando JavaScript para definir o valor do campo de data...")
            if not self._set_input_value_with_js(self.config["seletores"]["due_date_modal_input"], data_vencimento):
                raise Exception("Falha ao usar JS para definir a data.")
            time.sleep(1)

            self.logger.info("Localizando botão 'Salvar' e clicando com JavaScript...")
            botao_salvar = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(self.config["seletores"]["due_date_modal_save_btn"]))
            self.driver.execute_script("arguments[0].click();", botao_salvar)

            self.logger.info("Aguardando confirmação do salvamento (pop-up fechar)...")
            WebDriverWait(self.driver, 20).until(EC.invisibility_of_element_located(self.config["seletores"]["due_date_modal"]))
            self.logger.info("Data de vencimento definida com sucesso!")
            return True
            
        except Exception as e:
            self.logger.error(f"Ocorreu um erro ao tentar definir a data de vencimento via menu: {e}", exc_info=True)
            return False

    def processar_envio_boleto(self, item_info):
        self.item_atual_info = item_info
        codigo_empresa = self.item_atual_info["codigo_empresa"]
        nome_arquivo = self.item_atual_info["nome_arquivo"]
        caminho_completo_arquivo = self.item_atual_info["caminho_completo_arquivo"]
        self.logger.info(f"--- Iniciando processamento para empresa {codigo_empresa}: {self.item_atual_info['nome_empresa']} ---")
        if not self.selecionar_empresa(codigo_empresa): return
        pasta_principal = self.config["pastas"]["principal"]
        if not self._acessar_pasta_principal(pasta_principal): return
        subpasta = self.config["pastas"]["subpasta"]
        if not self._item_existe_no_grid(subpasta):
            self.logger.info(f"Subpasta '{subpasta}' não encontrada. Criando...")
            if not self._criar_pasta(subpasta): return
        if not self._navegar_para_subpasta(subpasta): return
        if self._item_existe_no_grid(nome_arquivo):
            self.logger.info(f"Arquivo '{nome_arquivo}' já existe nesta pasta. Nenhuma ação necessária.")
        else:
            self.logger.info(f"Arquivo '{nome_arquivo}' não encontrado. Iniciando upload...")
            upload_sucesso = self._upload_arquivo(caminho_completo_arquivo)
            data_vencimento = self.item_atual_info.get("data_vencimento")
            if upload_sucesso and data_vencimento:
                self._definir_data_vencimento_via_menu_gerenciar(nome_arquivo, data_vencimento)
            elif upload_sucesso and not data_vencimento:
                self.logger.warning(f"Upload do arquivo '{nome_arquivo}' concluído, mas nenhuma data de vencimento foi fornecida no JSON.")
        self.logger.info(f"--- Finalizado processamento para empresa {codigo_empresa} ---")

    def run(self, dados_para_processar):
        self.logger.info("Iniciando a execução da automação.")
        if not self.fazer_login():
            self.logger.error("Falha no login. A automação não pode continuar.")
            return
        if not self._navegar_para_documentos():
            self.logger.error("Falha ao navegar para a área de documentos. A automação não pode continuar.")
            return
        total_itens = len(dados_para_processar)
        for i, item in enumerate(dados_para_processar):
            self.logger.info(f"\n[ Processando item {i+1} de {total_itens} ]")
            try:
                self.processar_envio_boleto(item)
                if i < total_itens - 1:
                    self.logger.info("Resetando para o próximo cliente...")
                    self.driver.get(self.config['urls']['documentos'])
                    self._aguardar_elemento(self.config["seletores"]["docs_selecionar_cliente_input"], 30)
            except Exception as e:
                self.logger.error(f"Ocorreu um erro não tratado ao processar o item {item['codigo_empresa']}: {e}", exc_info=True)
                self.logger.info("Tentando resetar a página para continuar com o próximo item...")
                self.driver.get(self.config['urls']['documentos'])

    def fechar(self):
        if self.driver:
            self.logger.info("Fechando o navegador.")
            self.driver.quit()

# --- PONTO DE ENTRADA DA EXECUÇÃO ---
if __name__ == "__main__":
    logger.info("="*60)
    logger.info(" INICIANDO ROBÔ DE UPLOAD DE BOLETOS PARA O ONVIO ")
    logger.info("="*60)
    processor = DataProcessor(CONFIG)
    boletos_a_processar = processor.carregar_dados_de_entrada()
    if boletos_a_processar:
        automator = OnvioAutomator(CONFIG)
        try:
            automator.run(boletos_a_processar)
        except Exception as e:
            logger.error(f" Ocorreu um erro fatal na execução principal: {e}", exc_info=True)
        finally:
            automator.fechar()
    else:
        logger.error("Nenhum boleto válido para processar. Verifique o arquivo 'download.json' e a pasta 'temp'.")
    logger.info("="*60)
    logger.info(" EXECUÇÃO FINALIZADA ")
    logger.info("="*60)