import time
import os
import threading
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import WebDriverException

is_capturing = False  # Flag de controle global

def configurar_chrome_options():
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--disable-gpu")
    options.add_argument("--mute-audio")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-sync")
    options.add_argument("--disable-extensions")
    return options

def aplicar_stealth(driver):
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', {
              get: () => undefined
            });
        """
    })

def monitorar_entrada_terminal():
    global is_capturing
    print("\n--- DIGITE 'c' + Enter para alternar captura ON/OFF ---")
    print("--- DIGITE 'q' + Enter para ENCERRAR o script ---")
    while True:
        comando = input().strip().lower()
        if comando == 'c':
            is_capturing = not is_capturing
            print(f"Captura {'ativada' if is_capturing else 'pausada'}.")
        elif comando == 'q':
            print("Encerrando por comando do usuário.")
            os._exit(0)  # Encerra imediatamente todos os threads

def capturar_html_com_tecla_terminal(url, interval=2):
    global is_capturing

    options = configurar_chrome_options()
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    aplicar_stealth(driver)

    print(f"Abrindo navegador na página: {url}")
    driver.get(url)

    # Diretório atual do script
    output_dir = os.path.dirname(os.path.abspath(__file__))

    print(f"Salvando arquivos HTML no diretório: {output_dir}")

    # Inicia thread para escutar comandos no terminal
    thread_entrada = threading.Thread(target=monitorar_entrada_terminal, daemon=True)
    thread_entrada.start()

    try:
        while True:
            if is_capturing:
                timestamp = int(time.time() * 1000)
                html_content = driver.page_source
                filename = os.path.join(output_dir, f"captura_{timestamp}.html")
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(html_content)
                print(f"[{time.strftime('%H:%M:%S')}] Captura salva: {filename}")

            time.sleep(interval)

    except WebDriverException as e:
        print(f"\nNavegador fechado. Erro: {e}")
    except KeyboardInterrupt:
        print("\n--- Script interrompido pelo usuário. ---")
    finally:
        print("Fechando navegador...")
        driver.quit()

# --- Execução ---
if __name__ == "__main__":
    pagina_alvo_inicial = "https://www.google.com"
    capturar_html_com_tecla_terminal(pagina_alvo_inicial, interval=2)
