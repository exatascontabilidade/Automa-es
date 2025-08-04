# -*- coding: utf-8 -*-

import os
import sys
import json
import time
import logging
import re
from datetime import datetime
from collections import OrderedDict

# --- NOVAS BIBLIOTECAS ---
# pip install pyautoit
# pip install pygetwindow
import autoit
import pygetwindow as gw

# --- BIBLIOTECAS PADRÃO ---
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
        "documentos": "https://onvio.com.br/staff/#/documents",
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
        "resultados_upload": "upload_resultados.json",
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
        "dashboard_content": (By.CLASS_NAME, "dashboard-section__content"),
        "documentos_pagina_confirmacao": (By.XPATH, "//ul[contains(@class, 'nav-tabs')]//span[contains(text(), 'Documentos do Cliente')]"),
        "docs_selecionar_cliente_input": (By.CSS_SELECTOR, "input[placeholder='Selecione um cliente']"),
        "docs_lista_empresas_ul": (By.CSS_SELECTOR, "ul.bento-combobox-container-list"),
        "fiscal_aside_panel": (By.CSS_SELECTOR, "aside.bento-splitter-group-left"),
        "fiscal_bm_tree": (By.CSS_SELECTOR, "bm-tree"),
        "fiscal_tree_item_shadow": (By.CSS_SELECTOR, "bm-tree-item[title='Fiscal']"),
        "tree_item_loading": (By.CSS_SELECTOR, "bm-tree-item[title='Carregando...']"),
        "folder_grid_ready": (By.CSS_SELECTOR, "li.paginate_info"),
        "folder_grid_empty": (By.XPATH, "//div[contains(text(), 'A pasta selecionada está vazia.')]"),
        "novo_menu_btn": (By.ID, "dms-fe-legacy-components-client-documents-new-menu-button"),
        "nova_pasta_btn": (By.ID, "dms-fe-legacy-components-client-documents-new-folder-button"),
        "nova_pasta_nome_input": (By.ID, "containerName"),
        "nova_pasta_salvar_btn": (By.CSS_SELECTOR, "button[data-qe-id='dmsNewContainerModal-saveButton']"),
        "confirmacao_operacao_div": (By.CSS_SELECTOR, 'div.bottom-alerts-pane div[ng-if="operations.length === 1"]'),
        "upload_button_geral": (By.ID, "dms-fe-legacy-components-client-documents-upload-button"),
        "upload_carregando_msg": (By.CSS_SELECTOR, "span.file-alert-text"),
        "popup_erro_texto": (By.CSS_SELECTOR, "div.alert-error .file-alert-text"),
        "popup_erro_fechar_btn": (By.CSS_SELECTOR, "button.bento-alert-close"),
        "voltar_nivel_btn": (By.XPATH, "//a[i[@class='dms-icon-up-one-level']]"),
    },
    "pastas": {
        "principal": "PARCELAMENTOS",
        "mapeamento": {
            "FEDERAL_SIMPLIFICADO": "PARCELAMENTO SIMPLIFICADO", "PGFN": "PARCELAMENTO PGFN",
            "SIMPLES_NACIONAL": "PARCELAMENTO SIMPLES NACIONAL", "PREVIDENCIARIO": "PARCELAMENTO PREVIDENCIARIO",
            "NAO_PREVIDENCIARIO": "PARCELAMENTO NAO PREVIDENCIARIO"
        }
    }
}

# --- GERENCIADOR DE LOGS ---
def setup_logger():
    log_path = CONFIG["arquivos"]["log_execucao"]
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

# --- PROCESSAMENTO DE DADOS ---
class DataProcessor:
    def __init__(self, config):
        self.config = config
        self.caminho_planilha = self.config["arquivos"]["relacao_empresas"]
        self.caminho_json = self.config["arquivos"]["parcelamentos"]

    def _carregar_planilha_codigos(self):
        if not os.path.exists(self.caminho_planilha):
            logger.error(f"Planilha de códigos não encontrada: {self.caminho_planilha}")
            return None
        try:
            try: df = pd.read_excel(self.caminho_planilha, dtype=str, engine="xlrd")
            except Exception: df = pd.read_excel(self.caminho_planilha, dtype=str, engine="openpyxl")
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
        with open(self.caminho_json, "r", encoding="utf-8") as f: dados_json = json.load(f)
        dados_atualizados, total_atualizados = [], 0
        for item in dados_json:
            cnpj_limpo = re.sub(r"\D", "", item.get("cnpj", ""))
            codigo = mapa_codigos.get(cnpj_limpo)
            novo_item = OrderedDict(item); novo_item["codigo"] = codigo
            dados_atualizados.append(novo_item)
            if codigo: total_atualizados += 1
        with open(self.caminho_json, "w", encoding="utf-8") as f: json.dump(dados_atualizados, f, ensure_ascii=False, indent=4)
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
        options.add_argument("--start-maximized"); options.add_argument("--no-sandbox"); options.add_argument("--disable-dev-shm-usage")
        try:
            servico = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=servico, options=options)
            driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"})
            return driver
        except Exception as e:
            self.logger.error(f"Falha ao iniciar o navegador: {e}")
            raise

    def _aguardar_elemento(self, seletor, tempo=20):
        return WebDriverWait(self.driver, tempo).until(EC.element_to_be_clickable(seletor))
        
    def _wait_for_navigation_panel_to_load(self, tempo=20):
        self.logger.info("Aguardando painel de navegação lateral estabilizar...")
        try:
            WebDriverWait(self.driver, tempo).until(EC.invisibility_of_element_located(self.config["seletores"]["tree_item_loading"]))
            self.logger.info("Painel de navegação estabilizado.")
            return True
        except TimeoutException:
            self.logger.warning("Painel de navegação não estabilizou.")
            return False
            
    def _wait_for_api_request(self, url_api, status_code=202, settle_time=3, timeout=30):
        self.logger.info(f"Aguardando API '{url_api}' estabilizar com status {status_code}...")
        script = f"return window.performance.getEntriesByType('resource').filter(req => req.name.startsWith('{url_api}') && req.responseStatus === {status_code}).length;"
        start_time = time.time()
        last_count = -1
        while time.time() - start_time < timeout:
            try:
                current_count = self.driver.execute_script(script)
                if last_count != -1 and current_count == last_count: break
                last_count = current_count
                time.sleep(settle_time)
            except Exception: time.sleep(settle_time)
        final_count = self.driver.execute_script(script)
        if final_count > 0:
            self.logger.info(f"Total de {final_count} requisição(ões) com status {status_code} encontradas.")
            time.sleep(2)
        else:
            self.logger.warning(f"Nenhuma requisição para '{url_api}' com status {status_code} foi concluída.")
        return True

    def _wait_for_grid_to_load(self, tempo=20):
        self.logger.info("Aguardando o conteúdo do grid de pastas carregar (API, Painel e UI)...")
        try:
            self._wait_for_api_request(self.config["urls"]["api_notes"], status_code=202)
            self._wait_for_navigation_panel_to_load()
            WebDriverWait(self.driver, tempo).until(EC.any_of(EC.presence_of_element_located(self.config["seletores"]["folder_grid_ready"]), EC.presence_of_element_located(self.config["seletores"]["folder_grid_empty"])))
            self.logger.info("Grid carregado e dados recebidos.")
            return True
        except TimeoutException:
            self.logger.error("Conteúdo do grid (UI) não carregou no tempo esperado.")
            return False

    def fazer_login(self):
        self.logger.info("Iniciando processo de login.")
        self.driver.get(self.config["urls"]["login"])
        try: self._aguardar_elemento(self.config["seletores"]["login_continuar_btn"], tempo=10).click()
        except TimeoutException: self.logger.info("Botão 'Continuar' não encontrado.")
        self._aguardar_elemento(self.config["seletores"]["login_usuario_input"]).send_keys(self.config["credenciais"]["usuario"])
        self._aguardar_elemento(self.config["seletores"]["login_submit_btn"]).click()
        self._aguardar_elemento(self.config["seletores"]["login_senha_input"]).send_keys(self.config["credenciais"]["senha"])
        self._aguardar_elemento(self.config["seletores"]["login_submit_btn"]).click()
        self.logger.info("Login e senha enviados."); self._handle_mfa()

    def _handle_mfa(self):
        try:
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(self.config["seletores"]["mfa_titulo"]))
            self.logger.info("Tela de MFA detectada.")
            self._aguardar_elemento(self.config["seletores"]["mfa_email_btn"]).click()
            self.logger.info("Aguardando 20s pelo código..."); time.sleep(20)
            codigo = extrair_codigo_do_email()
            if not codigo: raise Exception("Código de verificação não foi extraído.")
            self._aguardar_elemento(self.config["seletores"]["mfa_codigo_input"]).send_keys(codigo)
            self._aguardar_elemento(self.config["seletores"]["mfa_continuar_btn"]).click()
            self.logger.info(f"Código '{codigo}' inserido.")
        except TimeoutException: self.logger.info("Nenhuma tela de MFA foi solicitada.")
        except Exception as e: self.logger.error(f"Falha durante o processo de MFA: {e}"); raise

    def _navegar_para_documentos(self, max_tentativas=2):
        """Espera o dashboard carregar, com retentativa, e navega para a URL de documentos."""
        self.logger.info("Iniciando navegação para a área de Documentos...")
        for tentativa in range(1, max_tentativas + 1):
            try:
                self.logger.info(f"Tentativa {tentativa}/{max_tentativas}: Aguardando dashboard carregar (até 4s)...")
                WebDriverWait(self.driver, 4).until(EC.presence_of_element_located(self.config["seletores"]["dashboard_content"]))
                
                self.logger.info("Dashboard carregado. Navegando para a área de Documentos...")
                self.driver.get(self.config["urls"]["documentos"])
                
                WebDriverWait(self.driver, 20).until(EC.presence_of_element_located(self.config["seletores"]["documentos_pagina_confirmacao"]))
                self.logger.info("Página de Documentos carregada com sucesso!")
                return True
            except Exception as e:
                self.logger.warning(f"Dashboard não carregou na tentativa {tentativa}. Erro: {e}")
                if tentativa < max_tentativas:
                    self.logger.info("Atualizando a página para nova tentativa...")
                    self.driver.refresh()
                else:
                    self.logger.error("Não foi possível carregar o dashboard após múltiplas tentativas.")
                    return False

    def selecionar_empresa(self, codigo_empresa):
        self.logger.info(f"Selecionando empresa com código: {codigo_empresa}")
        self.driver.execute_script("window.performance.clearResourceTimings();")
        campo_busca = self._aguardar_elemento(self.config["seletores"]["docs_selecionar_cliente_input"])
        if "ng-not-empty" in campo_busca.get_attribute("class"):
            campo_busca.send_keys(Keys.CONTROL + "a", Keys.BACKSPACE); time.sleep(1)
        campo_busca.send_keys(codigo_empresa)
        try:
            lista_ul = WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located(self.config["seletores"]["docs_lista_empresas_ul"]))
            xpath_empresa = f".//li[.//span[text()='{codigo_empresa}']]//span[2]"
            item_empresa = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath_empresa)))
            nome_empresa = item_empresa.text.strip()
            self.logger.info(f"Empresa encontrada: '{nome_empresa}'. Clicando...")
            self.empresa_atual['nome'] = nome_empresa; item_empresa.click()
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
            if not href: self.logger.warning("Atributo 'href' não encontrado. Tentando clicar."); fiscal_item.click()
            else: self.logger.info(f"Navegando para o link da pasta Fiscal: {href}"); self.driver.get(href)
            return self._wait_for_grid_to_load(max_espera)
        except Exception as e:
            self.logger.error(f"Erro ao acessar a pasta Fiscal via Shadow DOM: {e}", exc_info=True)
            return False
            
    def _item_existe_no_grid(self, nome_item):
        self.logger.info(f"Verificando no grid a existência de '{nome_item}'...")
        try:
            xpath_preciso = f"//dms-grid-text-cell[@text='{nome_item}']"
            elementos_encontrados = self.driver.find_elements(By.XPATH, xpath_preciso)
            return len(elementos_encontrados) > 0
        except Exception as e:
            self.logger.error(f"Ocorreu um erro ao verificar a existência do item '{nome_item}' no grid: {e}")
            return False

    def _criar_pasta(self, nome_pasta):
        self.logger.info(f"Iniciando criação da pasta '{nome_pasta}'...")
        self.driver.execute_script("window.performance.clearResourceTimings();")
        try:
            self._aguardar_elemento(self.config["seletores"]["novo_menu_btn"]).click()
            self._aguardar_elemento(self.config["seletores"]["nova_pasta_btn"]).click()
            campo_nome = self._aguardar_elemento(self.config["seletores"]["nova_pasta_nome_input"]); campo_nome.send_keys(nome_pasta)
            botao_salvar = self._aguardar_elemento(self.config["seletores"]["nova_pasta_salvar_btn"])
            self.driver.execute_script("arguments[0].click();", botao_salvar)
            WebDriverWait(self.driver, 20).until(EC.presence_of_element_located(self.config["seletores"]["confirmacao_operacao_div"]))
            WebDriverWait(self.driver, 10).until(EC.invisibility_of_element_located(self.config["seletores"]["confirmacao_operacao_div"]))
            self.logger.info("Atualizando o navegador..."); self.driver.refresh(); self._wait_for_grid_to_load()
            if self._item_existe_no_grid(nome_pasta): self.logger.info(f"VERIFICADO: Pasta '{nome_pasta}' existe."); return True
            else: self.logger.error(f"FALHA NA VERIFICAÇÃO: A pasta '{nome_pasta}' não foi encontrada."); return False
        except Exception as e:
            self.logger.error(f"Falha durante o processo de criação da pasta '{nome_pasta}': {e}", exc_info=True)
            return False
    
    def _salvar_resultado_upload(self, resultado_info):
        caminho_json = self.config["arquivos"]["resultados_upload"]
        try:
            dados = []
            if os.path.exists(caminho_json):
                with open(caminho_json, "r", encoding="utf-8") as f: dados = json.load(f)
            dados.append(resultado_info)
            with open(caminho_json, "w", encoding="utf-8") as f: json.dump(dados, f, indent=4, ensure_ascii=False)
        except Exception as e:
            self.logger.error(f"Não foi possível salvar o resultado no JSON '{caminho_json}': {e}")
    
    def _verificar_e_registrar_resultado_upload(self):
        try:
            self.logger.info("Verificando resultado do upload (sucesso ou pop-up de erro)...")
            popup_erro = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(self.config["seletores"]["popup_erro_texto"]))
            mensagem_erro = popup_erro.text.strip()
            self.logger.error(f"Erro de upload detectado: {mensagem_erro}")
            resultado = {"codigo": self.empresa_atual.get("codigo"), "nome": self.empresa_atual.get("nome"), "status": "Erro", "detalhes": mensagem_erro, "timestamp": datetime.now().isoformat()}
            self._salvar_resultado_upload(resultado)
            try:
                botao_fechar = self.driver.find_element(*self.config["seletores"]["popup_erro_fechar_btn"])
                self.driver.execute_script("arguments[0].click();", botao_fechar)
                self.logger.info("Pop-up de erro fechado.")
            except Exception: self.logger.warning("Não foi possível fechar o pop-up de erro automaticamente.")
        except TimeoutException:
            self.logger.info("Nenhum pop-up de erro detectado. Upload considerado sucesso.")
            resultado = {"codigo": self.empresa_atual.get("codigo"), "nome": self.empresa_atual.get("nome"), "status": "Sucesso", "detalhes": "Arquivos enviados com sucesso.", "timestamp": datetime.now().isoformat()}
            self._salvar_resultado_upload(resultado)

    def _upload_arquivos_via_gui(self, arquivos_para_upload):
        if not arquivos_para_upload: return True
        try:
            self.logger.info(f"Iniciando upload de {len(arquivos_para_upload)} arquivo(s)...")
            botao_upload = self._aguardar_elemento(self.config["seletores"]["upload_button_geral"])
            self.driver.execute_script("arguments[0].scrollIntoView(true);", botao_upload); botao_upload.click(); time.sleep(2)
            
            self.logger.info("Procurando pela janela 'Abrir'..."); janela_titulo = next((w.title for w in gw.getWindowsWithTitle('Abrir')), None)
            if not janela_titulo: self.logger.error("Janela de upload 'Abrir' não foi encontrada."); return False

            self.logger.info(f"Janela '{janela_titulo}' detectada. Enviando arquivos..."); autoit.win_wait_active(janela_titulo, timeout=10)
            caminhos_formatados = [f'"{os.path.join(info["pasta"], info["nome_arquivo"])}"' for info in arquivos_para_upload]
            string_de_arquivos = " ".join(caminhos_formatados)
            autoit.control_set_text(janela_titulo, "Edit1", string_de_arquivos); time.sleep(1)
            autoit.control_send(janela_titulo, "Edit1", "{ENTER}")
            
            try:
                self.logger.info("Aguardando mensagem de progresso do upload aparecer...")
                WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(self.config["seletores"]["upload_carregando_msg"]))
                self.logger.info("Mensagem de progresso detectada. Aguardando desaparecer...")
                WebDriverWait(self.driver, 300).until(EC.invisibility_of_element_located(self.config["seletores"]["upload_carregando_msg"]))
                self.logger.info("Mensagem de progresso desapareceu.")
            except TimeoutException: self.logger.warning("A mensagem de progresso do upload não foi detectada ou não desapareceu a tempo.")
            
            self._verificar_e_registrar_resultado_upload()
            self._wait_for_grid_to_load()
            return True
        except Exception as e:
            self.logger.error(f"Ocorreu um erro durante o upload via GUI: {e}", exc_info=True)
            try:
                if janela_titulo and autoit.win_exists(janela_titulo): autoit.win_close(janela_titulo); self.logger.warning("Janela 'Abrir' fechada após erro.")
            except: pass
            return False

    def processar_pastas_da_empresa(self, empresa_info):
        self.empresa_atual = empresa_info.copy(); codigo = self.empresa_atual['codigo']
        self.logger.info(f"--- Processando pastas e arquivos para: {self.empresa_atual.get('empresa', 'N/A')} (Código: {codigo}) ---")
        
        if not self.selecionar_empresa(codigo): return
        if not self._acessar_pasta_fiscal(): self.logger.error(f"Falha ao acessar a pasta Fiscal da empresa {codigo}."); return
        
        if not self._item_existe_no_grid(self.config["pastas"]["principal"]):
            if not self._criar_pasta(self.config["pastas"]["principal"]): self.logger.error("Falha ao criar a pasta 'PARCELAMENTOS'."); return
        
        try:
            self.logger.info("Acessando a pasta 'PARCELAMENTOS' via link direto...")
            self.driver.execute_script("window.performance.clearResourceTimings();")
            xpath_link_pasta = f"//dms-grid-text-cell[@text='{self.config['pastas']['principal']}']//a"
            link_da_pasta = self._aguardar_elemento((By.XPATH, xpath_link_pasta))
            href = link_da_pasta.get_attribute('href')
            if href: self.driver.get(href)
            else: link_da_pasta.click()
            self._wait_for_grid_to_load()
        except Exception as e: self.logger.error(f"Erro ao entrar na pasta 'PARCELAMENTOS': {e}"); return
        
        tipos_de_parcelamento = sorted(list(set(p["tipo_parcelamento"] for p in self.empresa_atual["parcelamentos"])))
        for tipo in tipos_de_parcelamento:
            nome_subpasta = self.config["pastas"]["mapeamento"].get(tipo.upper())
            if not nome_subpasta: self.logger.warning(f"Tipo de parcelamento '{tipo}' não possui mapeamento de pasta."); continue

            if not self._item_existe_no_grid(nome_subpasta):
                if not self._criar_pasta(nome_subpasta): self.logger.error(f"Falha ao criar subpasta '{nome_subpasta}'."); continue
            
            try:
                self.logger.info(f"Acessando a subpasta '{nome_subpasta}'...")
                self.driver.execute_script("window.performance.clearResourceTimings();")
                xpath_link_subpasta = f"//dms-grid-text-cell[@text='{nome_subpasta}']//a"
                link_subpasta = self._aguardar_elemento((By.XPATH, xpath_link_subpasta))
                href_sub = link_subpasta.get_attribute('href')
                if href_sub: self.driver.get(href_sub)
                else: link_subpasta.click()
                if not self._wait_for_grid_to_load(): self.logger.error(f"Grid da subpasta '{nome_subpasta}' não carregou. Pulando uploads."); continue
            except Exception as e: self.logger.error(f"Erro ao entrar na subpasta '{nome_subpasta}': {e}"); continue
            
            arquivos_do_tipo = [p for p in self.empresa_atual["parcelamentos"] if p["tipo_parcelamento"] == tipo]
            arquivos_para_enviar = [p for p in arquivos_do_tipo if not self._item_existe_no_grid(p['nome_arquivo'])]
            
            if arquivos_para_enviar:
                self.logger.info(f"Encontrados {len(arquivos_para_enviar)} arquivo(s) novo(s) para upload.")
                self._upload_arquivos_via_gui(arquivos_para_enviar)
            else:
                self.logger.info("Nenhum arquivo novo para enviar nesta pasta.")
            
            try:
                self.logger.info("Navegando para o nível anterior ('PARCELAMENTOS')..."); self._aguardar_elemento(self.config["seletores"]["voltar_nivel_btn"]).click(); self._wait_for_grid_to_load()
            except Exception as e:
                self.logger.error(f"Não foi possível voltar para a pasta 'PARCELAMENTOS'. Erro: {e}"); break
    
    def run(self, dados_empresas):
        if not dados_empresas: self.logger.warning("Nenhuma empresa para processar. Encerrando."); return
        try:
            self.fazer_login()
            if not self._navegar_para_documentos():
                raise Exception("Não foi possível carregar a página de documentos.")

            empresas_agrupadas = {}
            for item in dados_empresas:
                codigo = item["codigo"]
                if codigo not in empresas_agrupadas: empresas_agrupadas[codigo] = {"codigo": codigo, "empresa": item.get("empresa"), "cnpj": item.get("cnpj"), "parcelamentos": []}
                empresas_agrupadas[codigo]["parcelamentos"].append(item)
            
            for codigo, info in empresas_agrupadas.items():
                self.logger.info(f"\n==== INICIANDO PROCESSAMENTO PARA EMPRESA: {info.get('empresa')} (Cód: {codigo}) ====")
                self.selecionar_empresa(codigo)
                self.processar_pastas_da_empresa(info)
                self.logger.info(f"==== PROCESSAMENTO FINALIZADO PARA EMPRESA: {info.get('empresa')} (Cód: {codigo}) ====")

            self.logger.info("✅ Automação concluída com sucesso para todas as empresas.")
        except Exception as e:
            self.logger.error(f"❌ Ocorreu um erro fatal durante a execução: {e}", exc_info=True)
        finally:
            self.fechar()

    def fechar(self):
        if self.driver: self.logger.info("Fechando o navegador."); self.driver.quit()

# --- PONTO DE ENTRADA DA EXECUÇÃO ---
if __name__ == "__main__":
    logger.info("="*50)
    logger.info("🚀 INICIANDO ROBÔ DE CRIAÇÃO E UPLOAD NO ONVIO (v27) 🚀")
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