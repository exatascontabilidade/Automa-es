# -*- coding: utf-8 -*-

import os
import sys
import json
import time
import logging
import re
from datetime import datetime
from collections import OrderedDict

# Instalação de dependências (descomente se for a primeira vez)
# os.system(f"{sys.executable} -m pip install selenium webdriver-manager pandas openpyxl")

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

try:
    from email_verificacao import extrair_codigo_do_email
except ImportError:
    print("AVISO: Módulo 'email_verificacao' não encontrado. A autenticação de dois fatores pode falhar.")
    def extrair_codigo_do_email():
        return None

# --- CONFIGURAÇÃO CENTRALIZADA ---

CONFIG = {
    "urls": {
        "login": "https://onvio.com.br/#/",
        "documentos": "https://onvio.com.br/staff/#/documents/client",
        "api_notes": "https://onvio.com.br/api/notes/v1/comments/search",
    },
    "credenciais": {
        "usuario": "automacao.gestta@exatascontabilidade.com.br",
        "senha": "Exatas@1010",
    },
    "arquivos": {
        "relacao_empresas": "Relação Empresas - Nome - CNPJ.xls",
        "parcelamentos": "parcelamentos_baixados.json",
        "log_execucao": "relatorio_execucao.txt",
    },
    "seletores": {
        "login_continuar_btn": (By.ID, "trauth-continue-signin-btn"),
        "login_usuario_input": (By.NAME, "username"),
        "login_senha_input": (By.ID, "password"),
        "login_submit_btn": (By.XPATH, "//*[@type='submit']"),
        "mfa_titulo": (By.XPATH, "//h1[contains(text(), 'verificar sua identidade')]"),
        "mfa_email_btn": (By.XPATH, "//button[@name='action' and contains(@value, 'email')]"),
        "mfa_codigo_input": (By.ID, "code"),
        "mfa_continuar_btn": (By.XPATH, "//button[@name='action' and @value='default']"),
        "docs_selecionar_cliente_input": (By.CSS_SELECTOR, "input[placeholder='Selecione um cliente']"),
        "docs_lista_empresas_ul": (By.CSS_SELECTOR, "ul.bento-combobox-container-list"),
        "fiscal_aside_panel": (By.CSS_SELECTOR, "aside.bento-splitter-group-left"),
        "fiscal_bm_tree": (By.CSS_SELECTOR, "bm-tree"),
        "fiscal_tree_item_shadow": (By.CSS_SELECTOR, "bm-tree-item[title='Fiscal']"),
        "tree_item_loading": (By.CSS_SELECTOR, "bm-tree-item[title='Carregando...']"), # <-- NOVO SELETOR
        "folder_grid_ready": (By.CSS_SELECTOR, "li.paginate_info"),
        "folder_grid_empty": (By.XPATH, "//div[contains(text(), 'A pasta selecionada está vazia.')]"),
        "novo_menu_btn": (By.ID, "dms-fe-legacy-components-client-documents-new-menu-button"),
        "nova_pasta_btn": (By.ID, "dms-fe-legacy-components-client-documents-new-folder-button"),
        "nova_pasta_nome_input": (By.ID, "containerName"),
        "nova_pasta_salvar_btn": (By.CSS_SELECTOR, "button[data-qe-id='dmsNewContainerModal-saveButton']"),
        "confirmacao_operacao_div": (By.CSS_SELECTOR, 'div.bottom-alerts-pane div[ng-if="operations.length === 1"]'),
    },
    "pastas": {
        "principal": "PARCELAMENTOS",
        "mapeamento": {
            "FEDERAL_SIMPLIFICADO": "PARCELAMENTO SIMPLIFICADO",
            "PGFN": "PARCELAMENTO PGFN",
            "SIMPLES_NACIONAL": "PARCELAMENTO SIMPLES NACIONAL",
            "PREVIDENCIARIO": "PARCELAMENTO PREVIDENCIARIO",
            "NAO_PREVIDENCIARIO": "PARCELAMENTO NAO PREVIDENCIARIO"
        }
    }
}

# --- GERENCIADOR DE LOGS ---
def setup_logger():
    log_path = CONFIG["arquivos"]["log_execucao"]
    logger = logging.getLogger("OnvioAutomator")
    logger.setLevel(logging.INFO)
    if logger.hasHandlers():
        logger.handlers.clear()
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

# --- PROCESSAMENTO DE DADOS ---
class DataProcessor:
    def __init__(self, config):
        self.config = config
        self.caminho_planilha = self.config["arquivos"]["relacao_empresas"]
        self.caminho_json = self.config["arquivos"]["parcelamentos"]

    def _carregar_planilha_codigos(self):
        if not os.path.exists(self.caminho_planilha):
            logger.error(f"Planilha de códigos não encontrada em: {self.caminho_planilha}")
            return None
        try:
            try:
                df = pd.read_excel(self.caminho_planilha, dtype=str, engine="xlrd")
            except Exception:
                df = pd.read_excel(self.caminho_planilha, dtype=str, engine="openpyxl")
            df["CNPJ_LIMPO"] = df["CNPJ"].str.replace(r"\D", "", regex=True)
            return df[["CNPJ_LIMPO", "Cód."]].set_index("CNPJ_LIMPO").to_dict()["Cód."]
        except Exception as e:
            logger.error(f"Erro ao ler a planilha de códigos: {e}")
            return None

    def preparar_dados(self):
        logger.info("Iniciando preparação de dados...")
        mapa_codigos = self._carregar_planilha_codigos()
        if mapa_codigos is None: return None
        if not os.path.exists(self.caminho_json):
            logger.error(f"Arquivo JSON de parcelamentos não encontrado: {self.caminho_json}")
            return None
        with open(self.caminho_json, "r", encoding="utf-8") as f:
            dados_json = json.load(f)
        dados_atualizados, total_atualizados = [], 0
        for item in dados_json:
            cnpj_limpo = re.sub(r"\D", "", item.get("cnpj", ""))
            codigo = mapa_codigos.get(cnpj_limpo)
            novo_item = OrderedDict(item)
            novo_item["codigo"] = codigo
            dados_atualizados.append(novo_item)
            if codigo: total_atualizados += 1
        with open(self.caminho_json, "w", encoding="utf-8") as f:
            json.dump(dados_atualizados, f, ensure_ascii=False, indent=4)
        logger.info(f"{total_atualizados} de {len(dados_json)} registros foram atualizados com 'codigo'.")
        empresas_para_processar = [item for item in dados_atualizados if item.get("codigo")]
        if not empresas_para_processar:
            logger.warning("Nenhuma empresa com código válido encontrada para processar.")
            return []
        return empresas_para_processar

# --- CLASSE DE AUTOMAÇÃO ---
class OnvioAutomator:
    def __init__(self, config):
        self.config = config
        self.logger = logger
        self.driver = self._iniciar_navegador()
        self.empresa_atual = {}

    def _iniciar_navegador(self):
        self.logger.info("Iniciando navegador Chrome...")
        options = webdriver.ChromeOptions()
        options.add_experimental_option("prefs", {"plugins.always_open_pdf_externally": True})
        options.add_argument("--start-maximized")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        try:
            servico = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=servico, options=options)
            driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            })
            return driver
        except Exception as e:
            self.logger.error(f"Falha ao iniciar o navegador: {e}")
            raise

    def _aguardar_elemento(self, seletor, tempo=20):
        return WebDriverWait(self.driver, tempo).until(EC.element_to_be_clickable(seletor))
        
    def _wait_for_navigation_panel_to_load(self, tempo=20):
        """Espera o painel de navegação lateral parar de mostrar 'Carregando...'."""
        self.logger.info("Aguardando painel de navegação lateral estabilizar...")
        try:
            WebDriverWait(self.driver, tempo).until(
                EC.invisibility_of_element_located(self.config["seletores"]["tree_item_loading"])
            )
            self.logger.info("Painel de navegação estabilizado.")
            return True
        except TimeoutException:
            self.logger.warning("Painel de navegação não estabilizou (ainda mostra 'Carregando...').")
            return False
            
    def _wait_for_api_request(self, url_api, status_code=202, settle_time=3, timeout=30):
        self.logger.info(f"Aguardando API '{url_api}' estabilizar com status {status_code}...")
        script = f"""
            return window.performance.getEntriesByType('resource')
                .filter(req => req.name.startsWith('{url_api}') && req.responseStatus === {status_code})
                .length;
        """
        start_time = time.time()
        last_count = -1
        is_stable = False
        while time.time() - start_time < timeout:
            try:
                current_count = self.driver.execute_script(script)
                if last_count != -1 and current_count == last_count:
                    self.logger.info(f"Contagem de requisições estabilizou em {current_count}.")
                    is_stable = True
                    break
                last_count = current_count
                self.logger.info(f"Contagem atual de requisições (status {status_code}): {current_count}. Aguardando {settle_time}s...")
                time.sleep(settle_time)
            except Exception as e:
                self.logger.error(f"Erro ao executar script para verificar API: {e}")
                time.sleep(settle_time)
        
        if not is_stable:
            self.logger.warning(f"API não estabilizou no tempo de {timeout}s.")
        
        final_count = self.driver.execute_script(script)
        if final_count > 0:
            self.logger.info(f"Total de {final_count} requisições com status {status_code} encontradas.")
            self.logger.info("Aguardando 2 segundos extras de segurança...")
            time.sleep(2)
        else:
            self.logger.warning(f"Nenhuma requisição para '{url_api}' com status {status_code} foi concluída.")
        return True

    def _wait_for_grid_to_load(self, tempo=20):
        self.logger.info("Aguardando o conteúdo do grid de pastas carregar (API, Painel e UI)...")
        try:
            self._wait_for_api_request(self.config["urls"]["api_notes"], status_code=202)
            self._wait_for_navigation_panel_to_load()
            WebDriverWait(self.driver, tempo).until(
                EC.any_of(
                    EC.presence_of_element_located(self.config["seletores"]["folder_grid_ready"]),
                    EC.presence_of_element_located(self.config["seletores"]["folder_grid_empty"])
                )
            )
            self.logger.info("Grid carregado e dados recebidos.")
            return True
        except TimeoutException:
            self.logger.error("Conteúdo do grid (UI) não carregou no tempo esperado.")
            return False

    def fazer_login(self):
        self.logger.info("Iniciando processo de login.")
        self.driver.get(self.config["urls"]["login"])
        try:
            self._aguardar_elemento(self.config["seletores"]["login_continuar_btn"], tempo=10).click()
        except TimeoutException:
            self.logger.info("Botão 'Continuar' não encontrado, prosseguindo.")
        self._aguardar_elemento(self.config["seletores"]["login_usuario_input"]).send_keys(self.config["credenciais"]["usuario"])
        self._aguardar_elemento(self.config["seletores"]["login_submit_btn"]).click()
        self._aguardar_elemento(self.config["seletores"]["login_senha_input"]).send_keys(self.config["credenciais"]["senha"])
        self._aguardar_elemento(self.config["seletores"]["login_submit_btn"]).click()
        self.logger.info("Login e senha enviados.")
        self._handle_mfa()

    def _handle_mfa(self):
        try:
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(self.config["seletores"]["mfa_titulo"]))
            self.logger.info("Tela de verificação de identidade (MFA) detectada.")
            self._aguardar_elemento(self.config["seletores"]["mfa_email_btn"]).click()
            self.logger.info("Aguardando 20s pelo código...")
            time.sleep(20)
            codigo = extrair_codigo_do_email()
            if not codigo: raise Exception("Código de verificação não foi extraído do e-mail.")
            self._aguardar_elemento(self.config["seletores"]["mfa_codigo_input"]).send_keys(codigo)
            self._aguardar_elemento(self.config["seletores"]["mfa_continuar_btn"]).click()
            self.logger.info(f"Código '{codigo}' inserido com sucesso.")
        except TimeoutException:
            self.logger.info("Nenhuma verificação de identidade (MFA) foi solicitada.")
        except Exception as e:
            self.logger.error(f"Falha durante o processo de MFA: {e}")
            raise

    def selecionar_empresa(self, codigo_empresa):
        self.logger.info(f"Selecionando empresa com código: {codigo_empresa}")
        self.driver.execute_script("window.performance.clearResourceTimings();")
        campo_busca = self._aguardar_elemento(self.config["seletores"]["docs_selecionar_cliente_input"])
        if "ng-not-empty" in campo_busca.get_attribute("class"):
            campo_busca.send_keys(Keys.CONTROL + "a", Keys.BACKSPACE)
            time.sleep(1)
        campo_busca.send_keys(codigo_empresa)
        try:
            lista_ul = WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located(self.config["seletores"]["docs_lista_empresas_ul"]))
            xpath_empresa = f".//li[.//span[text()='{codigo_empresa}']]//span[2]"
            item_empresa = WebDriverWait(lista_ul, 10).until(EC.element_to_be_clickable((By.XPATH, xpath_empresa)))
            nome_empresa = item_empresa.text.strip()
            self.logger.info(f"Empresa encontrada: '{nome_empresa}'. Clicando...")
            self.empresa_atual['nome'] = nome_empresa
            item_empresa.click()
            return True
        except TimeoutException:
            self.logger.error(f"Não foi possível encontrar a empresa com código '{codigo_empresa}' na lista.")
            return False

    def _acessar_pasta_fiscal(self, max_espera=30):
        self.logger.info("Acessando a pasta 'Fiscal' via painel lateral (Shadow DOM)...")
        self.driver.execute_script("window.performance.clearResourceTimings();")
        try:
            aside = WebDriverWait(self.driver, max_espera).until(EC.presence_of_element_located(self.config["seletores"]["fiscal_aside_panel"]))
            host = WebDriverWait(aside, max_espera).until(EC.presence_of_element_located(self.config["seletores"]["fiscal_bm_tree"]))
            shadow_root = self.driver.execute_script("return arguments[0].shadowRoot", host)
            fiscal_item = WebDriverWait(shadow_root, max_espera).until(lambda d: d.find_element(*self.config["seletores"]["fiscal_tree_item_shadow"]))
            href = fiscal_item.get_attribute("href")
            if not href:
                self.logger.warning("Atributo 'href' não encontrado. Tentando clicar diretamente no item 'Fiscal'.")
                fiscal_item.click()
            else:
                self.logger.info(f"Navegando para o link da pasta Fiscal: {href}")
                self.driver.get(href)
            
            return self._wait_for_grid_to_load(max_espera)

        except Exception as e:
            self.logger.error(f"Erro ao acessar a pasta Fiscal via Shadow DOM: {e}", exc_info=True)
            return False
            
    def _item_existe_no_grid(self, nome_item):
        self.logger.info(f"Verificando no grid a existência de '{nome_item.upper()}'...")
        try:
            xpath_preciso = f"//dms-grid-text-cell[@text='{nome_item.upper()}']"
            elementos_encontrados = self.driver.find_elements(By.XPATH, xpath_preciso)
            
            if len(elementos_encontrados) > 0:
                self.logger.info(f"Item '{nome_item.upper()}' ENCONTRADO no grid.")
                return True
            else:
                self.logger.info(f"Item '{nome_item.upper()}' NÃO foi encontrado no grid.")
                return False
        except Exception as e:
            self.logger.error(f"Ocorreu um erro ao verificar a existência do item no grid: {e}")
            return False

    def _criar_pasta(self, nome_pasta):
        self.logger.info(f"Iniciando criação da pasta '{nome_pasta}'...")
        self.driver.execute_script("window.performance.clearResourceTimings();")
        try:
            self._aguardar_elemento(self.config["seletores"]["novo_menu_btn"]).click()
            self._aguardar_elemento(self.config["seletores"]["nova_pasta_btn"]).click()
            campo_nome = self._aguardar_elemento(self.config["seletores"]["nova_pasta_nome_input"])
            campo_nome.send_keys(nome_pasta)
            botao_salvar = self._aguardar_elemento(self.config["seletores"]["nova_pasta_salvar_btn"])
            self.driver.execute_script("arguments[0].click();", botao_salvar)
            
            WebDriverWait(self.driver, 20).until(EC.presence_of_element_located(self.config["seletores"]["confirmacao_operacao_div"]))
            self.logger.info(f"Confirmação de criação da pasta '{nome_pasta}' recebida.")
            WebDriverWait(self.driver, 10).until(EC.invisibility_of_element_located(self.config["seletores"]["confirmacao_operacao_div"]))
            
            self.logger.info("Atualizando o navegador para sincronizar a interface...")
            self.driver.refresh()
            
            self.logger.info("Aguardando página recarregar e VERIFICANDO a existência da nova pasta...")
            self._wait_for_grid_to_load()
            
            if self._item_existe_no_grid(nome_pasta):
                 self.logger.info(f"VERIFICADO: Pasta '{nome_pasta}' existe após a atualização.")
                 return True
            else:
                 self.logger.error(f"FALHA NA VERIFICAÇÃO: A pasta '{nome_pasta}' não foi encontrada após a criação.")
                 return False

        except Exception as e:
            self.logger.error(f"Falha durante o processo de criação da pasta '{nome_pasta}': {e}", exc_info=True)
            return False

    def processar_pastas_da_empresa(self, empresa_info):
        self.empresa_atual = empresa_info.copy()
        codigo = self.empresa_atual['codigo']
        self.logger.info(f"--- Processando pastas para: {self.empresa_atual.get('empresa', 'Nome não encontrado')} (Código: {codigo}) ---")
        if not self.selecionar_empresa(codigo): return
        if not self._acessar_pasta_fiscal():
            self.logger.error(f"Não foi possível continuar para a empresa {codigo} por falha ao acessar a pasta Fiscal.")
            return
        
        if not self._item_existe_no_grid(self.config["pastas"]["principal"]):
            if not self._criar_pasta(self.config["pastas"]["principal"]):
                self.logger.error("Não foi possível criar a pasta principal 'PARCELAMENTOS'. Abortando para esta empresa.")
                return
        
        try:
            self.logger.info("Acessando a pasta 'PARCELAMENTOS' via link direto (href)...")
            self.driver.execute_script("window.performance.clearResourceTimings();")
            
            xpath_link_pasta = f"//dms-grid-text-cell[@text='{self.config['pastas']['principal']}']//a"
            link_da_pasta = self._aguardar_elemento((By.XPATH, xpath_link_pasta))
            
            href = link_da_pasta.get_attribute('href')
            if href:
                self.logger.info(f"Navegando diretamente para: {href}")
                self.driver.get(href)
            else:
                self.logger.warning("Não foi possível encontrar o href, tentando clicar como alternativa...")
                link_da_pasta.click()

            self._wait_for_grid_to_load()
        except Exception as e:
            self.logger.error(f"Erro ao entrar na pasta 'PARCELAMENTOS': {e}")
            return
        
        tipos_parcelamento = set(p["tipo_parcelamento"] for p in self.empresa_atual["parcelamentos"])
        for tipo in tipos_parcelamento:
            nome_subpasta = self.config["pastas"]["mapeamento"].get(tipo.upper())
            if nome_subpasta:
                if not self._item_existe_no_grid(nome_subpasta):
                    self._criar_pasta(nome_subpasta)
            else:
                self.logger.warning(f"Tipo de parcelamento '{tipo}' não possui mapeamento de pasta.")

    def run(self, dados_empresas):
        if not dados_empresas:
            self.logger.warning("Nenhuma empresa para processar. Encerrando.")
            return
        try:
            self.fazer_login()
            empresas_agrupadas = {}
            for item in dados_empresas:
                codigo = item["codigo"]
                if codigo not in empresas_agrupadas:
                    empresas_agrupadas[codigo] = {"codigo": codigo, "empresa": item.get("empresa"), "cnpj": item.get("cnpj"), "parcelamentos": []}
                empresas_agrupadas[codigo]["parcelamentos"].append(item)
            for codigo, info in empresas_agrupadas.items():
                self.driver.get(self.config["urls"]["documentos"])
                self._aguardar_elemento(self.config["seletores"]["docs_selecionar_cliente_input"])
                self.processar_pastas_da_empresa(info)
                self.logger.info(f"Processo para a empresa {codigo} finalizado.")
            self.logger.info("✅ Automação concluída com sucesso para todas as empresas.")
        except Exception as e:
            self.logger.error(f"❌ Ocorreu um erro fatal durante a execução: {e}", exc_info=True)
        finally:
            self.fechar()

    def fechar(self):
        if self.driver:
            self.logger.info("Fechando o navegador.")
            self.driver.quit()

# --- PONTO DE ENTRADA DA EXECUÇÃO ---
if __name__ == "__main__":
    logger.info("="*50)
    logger.info("🚀 INICIANDO ROBÔ DE CRIAÇÃO DE PASTAS NO ONVIO (v15) 🚀")
    logger.info("="*50)
    processor = DataProcessor(CONFIG)
    empresas_a_processar = processor.preparar_dados()
    if empresas_a_processar:
        automator = OnvioAutomator(CONFIG)
        automator.run(empresas_a_processar)
    else:
        logger.error("Execução interrompida devido à falha na preparação dos dados.")
    logger.info("="*50)
    logger.info("🏁 EXECUÇÃO FINALIZADA 🏁")
    logger.info("="*50)