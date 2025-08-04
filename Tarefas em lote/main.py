import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import os
import time
import re

# === CONFIGURAÇÕES DE CAMINHO LOCAL ===
DIRETORIO_ATUAL = os.path.dirname(os.path.abspath(__file__))
CAMINHO_PLANILHA = os.path.join(DIRETORIO_ATUAL, 'gestta-tarefas-recorrentes.xlsx')
PASTA_SAIDA = os.path.join(DIRETORIO_ATUAL, 'Tarefas')
COLUNA_TAREFAS = 'Nome da Tarefa'
COLUNA_ID = 'ID da Tarefa' 
TEMPO_ESPERA = 10  # segundos

# === CONFIGURA OPÇÕES DO CHROME PARA NÃO SER DETECTADO COMO AUTOMATIZAÇÃO ===
chrome_options = Options()
chrome_options.add_argument('--start-maximized')
chrome_options.add_argument('--disable-blink-features=AutomationControlled')
chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
chrome_options.add_experimental_option('useAutomationExtension', False)

# === ABRE NAVEGADOR COM CONFIGURAÇÕES OCULTAS ===
driver = webdriver.Chrome(options=chrome_options)
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": """
        Object.defineProperty(navigator, 'webdriver', {
          get: () => undefined
        })
    """
})

# === ACESSA A PÁGINA INICIAL DO ONVIO ===
driver.get("https://onvio.com.br/#/")

# === CONFIGURA TEMPO DE ESPERA EXPLÍCITO ===
TEMPO_ESPERA = 20  # ou o tempo que preferir
wait = WebDriverWait(driver, TEMPO_ESPERA)

print("🔐 Faça login manualmente no navegador (incluindo captcha, se houver).")
print("Login : eduardo@exatascontabilidade.com.br")
print("Senha : Exatas1010")
input("✅ Pressione ENTER após completar o login e estar na tela de busca.")

# === TENTA CAPTURAR UMA URL VÁLIDA ===
driver.switch_to.window(driver.window_handles[0])

try:
    tentativa = 0
    while tentativa < 10:
        url_base = driver.current_url
        if url_base and not url_base.startswith("chrome://") and "about:blank" not in url_base:
            break
        tentativa += 1
        print("⏳ Aguardando o usuário abrir uma página válida...")
        time.sleep(2)

    if not url_base or url_base.strip() == "" or url_base.startswith("chrome://") or "about:blank" in url_base:
        raise Exception("Nenhuma página válida foi aberta na aba atual.")

    print(f"🌐 URL base capturada: {url_base}")

except Exception as e:
    print(f"❌ Erro ao capturar a URL base: {e}")
    print("🚫 O navegador foi fechado ou a página ainda não carregou. Reinicie o script.")
    driver.quit()
    exit()

# === LÊ PLANILHA ===
df = pd.read_excel(CAMINHO_PLANILHA)

def limpar_nome_arquivo(nome):
    # Remove ou substitui caracteres inválidos no nome de arquivos
    nome_limpo = re.sub(r'[\\/*?:"<>|]', '_', nome)
    return nome_limpo

# === LOOP DE TAREFAS ===
for index, row in df.iterrows():
    nome_tarefa = str(row[COLUNA_TAREFAS]).strip()
    id_tarefa = str(row.get(COLUNA_ID, '')).strip()
    print(f"\n🔍 Processando tarefa: {nome_tarefa} (ID: {id_tarefa})")

    try:
        # 1. VOLTA PARA PÁGINA BASE
        driver.get(url_base)

        # 2. ESPERA CAMPO DE BUSCA APARECER
        campo_busca = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'input[placeholder="Pesquisar por nome"]')
        ))
        campo_busca.clear()
        campo_busca.send_keys(nome_tarefa)
        time.sleep(5)

        # 3. ESPERA TABELA DE RESULTADOS
        wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, '.ag-body-container .ag-row')
        ))
        time.sleep(2)

        # 4. LOCALIZA TODOS OS LINKS DA TABELA
        links = driver.find_elements(By.CSS_SELECTOR, 'a.link-to-edit')

        encontrou = False
        for link in links:
            texto_link = link.text.strip()
            href = link.get_attribute("href")
            if texto_link == nome_tarefa:
                if id_tarefa == "" or id_tarefa in href:
                    print(f"✅ Encontrado: {texto_link} com ID no href: {href}")
                    link.click()
                    encontrou = True
                    break

        if not encontrou:
            raise Exception("Tarefa não encontrada com ID correspondente ou nome exato.")

        # 5. ESPERA FORMULÁRIO CARREGAR
        wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'form[name="detailsCtrl.detailsForm"]')
        ))
        # 5. ESPERA FORMULÁRIO CARREGAR
        wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'form[name="detailsCtrl.detailsForm"]')
        ))    
        # 6. EXTRAI DADOS DO FORMULÁRIO COM ESPERA E TRATAMENTO DE ERROS
        try:
            nome_tarefa_form = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'input#name'))).get_attribute('value')
        except Exception as e:
            nome_tarefa_form = "[Erro ao capturar Nome da Tarefa]"
            print(f"⚠️ Erro ao capturar Nome da Tarefa: {e}")

        try:
            departamento = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'span[ng-bind="$select.selected.name"]'))).text
        except Exception as e:
            departamento = "[Erro ao capturar Departamento]"
            print(f"⚠️ Erro ao capturar Departamento: {e}")

        try:
            frequencia = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'span[ng-bind*="TASK_FREQUENCIES"]'))).text
        except Exception as e:
            frequencia = "[Erro ao capturar Frequência]"
            print(f"⚠️ Erro ao capturar Frequência: {e}")

        try:
            competencia = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//label[contains(text(), "Competência")]/following::span[@ng-bind="$select.selected"][1]'))).text
        except Exception as e:
            competencia = "[Erro ao capturar Competência]"
            print(f"⚠️ Erro ao capturar Competência: {e}")

        try:
            data_meta = wait.until(EC.presence_of_element_located((
                By.XPATH, '//label[contains(text(), "Data meta")]/following::span[@ng-bind="$select.selected"][1]'
            ))).text
        except Exception as e:
            data_meta = "[Erro ao capturar Data Meta]"
            print(f"⚠️ Erro ao capturar Data Meta: {e}")

        try:
            esfera = wait.until(EC.presence_of_element_located((
                By.XPATH, '//span[contains(@ng-bind, "REGIME_SPHERE.") and contains(@class, "ng-binding")]'
            ))).text
        except Exception as e:
            esfera = "[Erro ao capturar Esfera]"
            print(f"⚠️ Erro ao capturar Esfera: {e}")

        try:
            data_legal = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//label[contains(text(), "Data legal")]/following::span[@ng-bind="$select.selected"][1]'))).text
        except Exception as e:
            data_legal = "[Erro ao capturar Data Legal]"
            print(f"⚠️ Erro ao capturar Data Legal: {e}")

        try:
            checkbox_multa = driver.find_element(By.CSS_SELECTOR, 'input[name="fine"]')
            gera_multa = "Sim" if checkbox_multa.is_selected() else "Não"
        except Exception as e:
            gera_multa = "[Erro ao verificar Geração de Multa]"
            print(f"⚠️ Erro ao verificar Geração de Multa: {e}")
        try:
            ativo_checkbox = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input#active')))
            ativo = "Sim" if ativo_checkbox.is_selected() else "Não"
        except Exception as e:
            ativo = "[Erro ao capturar status de Ativo]"
            print(f"⚠️ Erro ao capturar Ativo: {e}") 
            
        try:
            atividades_anexo_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//a[span[text()="Atividades com anexo"]]')))
            atividades_anexo_btn.click()
            print("✅ Aba 'Atividades com anexo' acessada com sucesso.")
        except Exception as e:
            print(f"⚠️ Erro ao acessar aba 'Atividades com anexo': {e}")
            
        try:
            documentos = []
            container = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.ag-body-container')))
            linhas = container.find_elements(By.CSS_SELECTOR, 'div[role="row"]')
            
            for i, linha in enumerate(linhas, start=1):
                try:
                    nome_doc = linha.find_element(By.CSS_SELECTOR, 'div[col-id="name"] a span').text.strip()
                    documentos.append(f"{i}. {nome_doc}")
                except Exception as e:
                    documentos.append(f"{i}. [Erro ao capturar documento: {e}]")
            
            print("✅ Documentos capturados com sucesso.")
        except Exception as e:
            documentos = ["[Erro ao capturar documentos da aba 'Atividades com anexo']"]
            print(f"⚠️ Erro ao capturar documentos: {e}")
  
        
        try:
            checklist_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//a[span[text()="Checklist"]]')))
            checklist_btn.click()
            print("✅ Aba 'Checklist' acessada com sucesso.")
        except Exception as e:
            print(f"⚠️ Erro ao acessar aba 'Checklist': {e}")
            
        try:
            checklist_itens = wait.until(EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, 'div.row[ng-repeat*="step in stepCtrl.model.steps"]')
            ))

            lista_checklist = []
            for item in checklist_itens:
                try:
                    ordem = item.find_element(By.CSS_SELECTOR, 'input[data-qe-id*="step-order"]').get_attribute('value').strip()
                    nome = item.find_element(By.CSS_SELECTOR, 'input[data-qe-id*="step-name"]').get_attribute('value').strip()
                    checkbox = item.find_element(By.CSS_SELECTOR, 'input[data-qe-id*="step-required"]')
                    obrigatorio = checkbox.is_selected()
                    obrigatorio_str = "Sim" if obrigatorio else "Não"
                    lista_checklist.append(f"{ordem}. {nome} – Obrigatório: {obrigatorio_str}")
                except Exception as item_error:
                    print(f"⚠️ Erro ao capturar item do checklist: {item_error}")
        except Exception as e:
            lista_checklist = ["[Erro ao capturar checklist]"]
            print(f"⚠️ Erro ao capturar checklist: {e}")    

        # 7. SALVA RESULTADO EM .TXT
        nome_tarefa_arq = limpar_nome_arquivo(nome_tarefa_form)
        with open(os.path.join(PASTA_SAIDA, f'{nome_tarefa_arq} - {id_tarefa}.txt'), 'w', encoding='utf-8') as f:
            f.write(f"Tarefa: {nome_tarefa_form}\n")
            f.write(f"Esfera: {esfera}\n")
            f.write(f"Frequência: {frequencia}\n")
            f.write(f"Data legal: {data_legal}\n")
            f.write(f"Data Meta: {data_meta}\n")
            f.write(f"Departamento: {departamento}\n")
            f.write(f"Competência: {competencia}\n")
            f.write(f"Gera Multa: {gera_multa}\n")
            f.write(f"Tarefa Ativa: {ativo}\n")
            f.write("\nChecklist da Tarefa:\n")
            for item in lista_checklist:
                f.write(f"{item}\n")
            f.write("Ativiadades com anexo:\n")
            for doc in documentos:
                f.write(f"{doc}\n")    
            

        print("✅ Tarefa extraída com sucesso.")

    except Exception as e:
        print(f"❌ Erro ao processar '{nome_tarefa}': {e}")
            
        # Garante que o nome usado no arquivo seja válido
        nome_tarefa_limpa = limpar_nome_arquivo(nome_tarefa)
        
        # Cria a pasta de saída se não existir
        os.makedirs(PASTA_SAIDA, exist_ok=True)
        
        with open(os.path.join(PASTA_SAIDA, f'{nome_tarefa_limpa}_ERRO.txt'), 'w', encoding='utf-8') as f:
            f.write(f"Erro ao processar a tarefa '{nome_tarefa}':\n{str(e)}\n")
            
# === ENCERRA ===
driver.quit()
print("\n🏁 Processamento de todas as tarefas concluído.")
