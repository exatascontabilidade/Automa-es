import os
import time
import base64
import re
import requests
from urllib.parse import urljoin
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# Caminho absoluto do diretório do script
diretorio = os.path.abspath(os.path.dirname(__file__))

# Configurações de preferências do Chrome
prefs = {
    "download.default_directory": diretorio,  # Salva os arquivos no mesmo diretório do script
    "download.prompt_for_download": False,  # Não perguntar onde salvar
    "download.directory_upgrade": True,  # Permitir sobrepor diretório
    "plugins.always_open_pdf_externally": True,  # Baixa em vez de abrir no navegador
    "profile.default_content_setting_values.automatic_downloads": 1,  # Permite múltiplos downloads automáticos
}

# Configurações do ChromeDriver
chrome_options = Options()
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-print-preview")  # Impede abertura do visualizador de impressão
chrome_options.add_argument("--kiosk-printing")  # Impressão silenciosa (geralmente útil só com .print())

# Inicializa o driver
driver = webdriver.Chrome(options=chrome_options)

def verificar_e_baixar_anexo(driver):
    try:
        print("🔍 Verificando se há anexo disponível para download...")

        # Arquivos existentes antes do clique
        arquivos_antes = set(os.listdir(diretorio))

        # Aguarda botão de anexo aparecer
        botao = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@title='Baixar o arquivo']"))
        )
        nome_pdf = botao.text.strip()

        if not nome_pdf.lower().endswith(".pdf"):
            print("ℹ️ O elemento encontrado não é um PDF.")
            return False

        print(f"📎 Anexo detectado: {nome_pdf}")
        botao.click()

        # Aguarda novo arquivo ser criado no diretório
        timeout = 20
        for i in range(timeout):
            time.sleep(1)
            arquivos_depois = set(os.listdir(diretorio))
            novos = arquivos_depois - arquivos_antes
            pdfs = [f for f in novos if f.lower().endswith(".pdf")]
            if pdfs:
                print(f"✅ PDF salvo: {pdfs[0]}")
                return True

        print("⏱️ Timeout ao aguardar download do anexo.")
        return False

    except NoSuchElementException:
        print("📎 Nenhum anexo encontrado na página.")
        return False
    except Exception as e:
        print(f"❌ Erro ao tentar baixar o anexo: {e}")
        return False

    

def obter_numero_pagina(driver):
    try:
        WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "visualizacao-sequencial"))
        )
        contador = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((
                By.XPATH,
                "//div[contains(@class, 'visualizacao-sequencial')]//div[contains(@class, 'contador')]"
            ))
        )
        texto = contador.text.strip()
        numeros = re.findall(r'\d+', texto)
        if len(numeros) == 2:
            return int(numeros[0]), int(numeros[1]), texto
        else:
            raise ValueError(f"Formato inesperado do contador: '{texto}'")
    except Exception as e:
        print(f"❌ Erro ao obter número da página: {e}")
        with open("pagina_debug.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print("💾 HTML salvo como 'pagina_debug.html' para depuração.")
        return None, None, None

def salvar_pdf_impressao(driver, numero_pagina, id_norma):
    try:
        url_impressao = f"https://normasinternet2.receita.fazenda.gov.br/#/consulta/externa/imprimir/{id_norma}/visao/multivigente"
        print(f"🖨️ Acessando URL de impressão: {url_impressao}")

        aba_original = driver.current_window_handle
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])

        # Abrir a URL da impressão
        driver.get("about:blank")

        # Bloquear o print antes da página carregar
        driver.execute_script("window.print = () => {};")

        # Agora sim acessar a página de impressão
        driver.get(url_impressao)

        # Aguarda o título da página carregar corretamente
        WebDriverWait(driver, 20).until(
            lambda d: d.title.strip() != "" and d.title.strip().lower() != "normas"
        )

        titulo = driver.title.strip()
        nome_arquivo = re.sub(r'[^a-zA-Z0-9\s]+', '', titulo).strip().replace(' ', '_')
        if not nome_arquivo:
            nome_arquivo = f"pagina_{numero_pagina:03d}"
        arquivo_pdf = os.path.join(diretorio, f"{nome_arquivo}.pdf")

        # Gera o PDF via DevTools
        result = driver.execute_cdp_cmd("Page.printToPDF", {
            'landscape': False,
            'displayHeaderFooter': False,
            'printBackground': True,
            'preferCSSPageSize': True,
            'scale': 1.0,
            'paperWidth': 8.5,
            'paperHeight': 11,
            'marginTop': 0.4,
            'marginBottom': 0.4,
            'marginLeft': 0.4,
            'marginRight': 0.4,
        })

        with open(arquivo_pdf, 'wb') as f:
            f.write(base64.b64decode(result['data']))

        print(f"✅ PDF salvo: {arquivo_pdf}")

        driver.close()
        driver.switch_to.window(aba_original)
        return True

    except Exception as e:
        print(f"❌ Erro ao salvar PDF de impressão: {e}")
        return False


# === EXECUÇÃO PRINCIPAL ===
try:
    url = "https://normasinternet2.receita.fazenda.gov.br/#/consulta/externa/124275/vs/MTI0Mjc1LDEyMzc4NSwxMTMwNDksOTI0MzksODUzNTEsODQwNjUsODQwNjYsNzkwMjIsNjAzMjIsNTc4MjUsMzkzOTIsMzcxMjcsMTU4ODYsMTU4NzMsMTU4NjMsMTU4MzYsMTU3MjYsMzc4NzEsMTU3MTMsMTU3MTAsMTU3MTEsMTU3MDIsMTU2OTcsMTQ0ODgzLDE0NDgyMywxNDQ4MTksMTQ0ODIxLDE0NDgyNCwxNDQ3NjYsMTQ0NzA4LDE0NDY4MCwxNDQ1ODcsMTQ0NTc5LDE0NDU2NywxNDM4MDEsMTQzNTI3LDE0MzUzMiwxNDMzODQsMTQzMjUyLDE0MzE5MCwxNDMxMjQsMTQyODgzLDE0MjQ0OSwxNDI0NDcsMTQyNDQzLDE0MjM2NSwxNDIzNDMsMTQyMTI1LDE0MjE1MiwxNDIxMTUsMTQxMTUyLDE0MTE1MSwxNDA5NDQsMTQwOTQ1LDE0MDc0NCwxNDA3NTgsMTQwNzYwLDE0MDY2MCwxNDA2NTcsMTQwNjc4LDE0MDU4MiwxNDA1NzYsMTQwNDQxLDE0MDQyOSwxNDAzMjAsMTQwMzMzLDE0MDIzNiwxNDAyNDMsMTQwMDI5LDEzOTg3OCwxMzk4MzIsMTM5ODI5LDEzOTc2MywxMzk1MTksMTM5NTI3LDEzOTQ4NCwxMzk0ODMsMTM5NDYyLDEzOTM3NiwxMzkzMDgsMTM5Mjk0LDEzOTI4MywxMzkyNzYsMTM5MjAwLDEzODk4NCwxMzg4NjMsMTM4ODQxLDEzODc5MywxMzg3ODksMTM4NzYzLDEzODUzMiwxMzg0MjEsMTM3ODE1LDEzNzY2MSwxMzc2NjIsMTM3NTgzLDEzNzQ4OSwxMzcyMTAsMTM3MTY0LDEzNjk3Mw=="
    driver.get(url)
    print("🔗 Acessando página inicial:", url)

    while True:
        try:
            pagina_atual, total_paginas, texto_contador = obter_numero_pagina(driver)
            if not pagina_atual or not total_paginas:
                break
            print(f"\n📄 Página atual: {pagina_atual}/{total_paginas}")
            verificar_e_baixar_anexo(driver)
            url_atual = driver.current_url
            id_match = re.search(r'/consulta/externa/(\d+)', url_atual)
            if not id_match:
                print("❌ Não foi possível extrair o ID da norma.")
                break
            id_norma = id_match.group(1)
            sucesso = salvar_pdf_impressao(driver, pagina_atual, id_norma)
            if not sucesso:
                break
            if pagina_atual >= total_paginas:
                print("🏁 Todas as páginas foram salvas.")
                break
            botao_proxima_xpath = "//i[text()='keyboard_arrow_right' and contains(@class, 'material-icons icon-24')]/.."
            botao_proxima = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, botao_proxima_xpath))
            )
            print("▶️ Indo para a próxima página...")
            botao_proxima.click()
            WebDriverWait(driver, 15).until(
                lambda d: d.find_element(
                    By.XPATH,
                    "//div[contains(@class, 'visualizacao-sequencial')]//div[contains(@class, 'contador')]"
                ).text.strip() != texto_contador
            )
        except Exception as e:
            print(f"❌ Erro durante a navegação: {e}")
            break
except Exception as e:
    print(f"❌ Erro geral: {e}")
finally:
    driver.quit()
    print("✅ Script finalizado.")
