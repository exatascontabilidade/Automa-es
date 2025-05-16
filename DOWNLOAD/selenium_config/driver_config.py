### selenium_config/driver_config.py
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import utils.state as state

def configurar_driver(headless=False):
    options = Options()
    if state.usar_headless:
        options.add_argument("--headless=new")
        print("🔍 Modo headless ativado.")
    options.add_argument("--disable-gpu")
    options.add_argument("--mute-audio")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])

    servico = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=servico, options=options)