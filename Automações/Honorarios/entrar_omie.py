import time
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup

# Configuração do logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def trocar_para_nova_janela(navegador):
    """Troca o foco do Selenium para a última janela aberta."""
    time.sleep(3)
    janelas = navegador.window_handles
    navegador.switch_to.window(janelas[-1])
    logging.info("✅ Alternado para a nova aba com sucesso!")

def fechar_popup(navegador):
    """Fecha o popup se ele estiver presente."""
    try:
        popup = WebDriverWait(navegador, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "MuiDialog-container"))
        )
        if popup.is_displayed():
            logging.info("🔍 Popup detectado! Tentando fechar...")
            try:
                botao_fechar = popup.find_element(By.XPATH, ".//button[contains(@class, 'MuiButtonBase-root')]")
                botao_fechar.click()
                logging.info("✅ Popup fechado com sucesso!")
                time.sleep(2)
            except:
                logging.warning("⚠️ Nenhum botão de fechar encontrado, tentando clicar fora.")
                navegador.execute_script("document.querySelector('.MuiDialog-container').click();")
                time.sleep(2)
    except Exception:
        logging.info("⚠️ Nenhum popup detectado ou erro ao tentar fechá-lo.")

def fechar_abas_omie(navegador, aba_omie):
    """Fecha todas as abas abertas e retorna à aba principal."""
    janelas = navegador.window_handles
    while len(janelas) > 1:
        navegador.switch_to.window(janelas[-1])
        navegador.close()
        janelas = navegador.window_handles
    navegador.switch_to.window(navegador.window_handles[0])

def verificar_disponibilidade_boleto(navegador):
    """Verifica se o boleto está disponível ou se há um aviso indicando que precisa se conectar."""
    try:
        mensagem_erro = WebDriverWait(navegador, 3).until(
            EC.presence_of_element_located((By.XPATH, "//p[contains(text(), 'Conecte para Visualizar o Link')]"))
        )
        if mensagem_erro.is_displayed():
            logging.warning("⚠️ O boleto não está disponível para download. É necessário conectar-se para visualizar o link.")
            return False
    except:
        logging.info("✅ Nenhuma mensagem de bloqueio encontrada. Boleto disponível para download.")
    return True

def baixar_boletos_atrasados(navegador):
    """Baixa todas as parcelas disponíveis, verificando sequências (1/5, 2/5, etc.)."""
    trocar_para_nova_janela(navegador)
    aba_omie = navegador.current_window_handle

    if not verificar_disponibilidade_boleto(navegador):
        logging.info("🚫 Pulando o processo de download, pois o boleto não está acessível.")
        fechar_abas_omie(navegador, aba_omie)
        return

    try:
        fechar_popup(navegador)
        logging.info("🔍Baixando todos os Boletos...")
        
        botao_baixar_todas = WebDriverWait(navegador, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[.//p[text()='Baixar todas']]"))
            )
        
        navegador.execute_script("arguments[0].click();", botao_baixar_todas)
        logging.info("📥 Botão 'Baixar todas' clicado com sucesso! Iniciando download dos boletos.")
        time.sleep(5)  # tempo para garantir o download

    except Exception as e:
        logging.error(f"❌ Erro ao localizar ou processar a seção de cobranças: {e}")
        fechar_abas_omie(navegador, aba_omie)

    logging.info("✅ Processo concluído!")
    fechar_abas_omie(navegador, aba_omie)
