import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import variaveis_globais as vg
from comparar_pdfs import obter_codigo_empresa
from upload_boleto import acessar_pasta_financeiro
from selenium.common.exceptions import TimeoutException
from email_verificacao import extrair_codigo_do_email

# Configuração do logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def abrir_nova_aba(navegador, url_destino):
    """
    Abre uma nova aba, acessa o site desejado e fecha a aba do Gmail.
    """
    try:
        # 🆕 Abrir nova aba
        navegador.execute_script("window.open('');")
        time.sleep(1)  # Pequeno delay para garantir que a aba foi aberta

        # 🔄 Alternar para a nova aba (última aberta)
        navegador.switch_to.window(navegador.window_handles[-1])
        logging.info(f"✅ Nova aba aberta! Acessando: {url_destino}")

        # 🌐 Acessar o site desejado
        navegador.get(url_destino)
        time.sleep(5)  # Aguarda um tempo para garantir o carregamento do site

        # 🔴 Fechar a aba do Gmail (Primeira aba)
        if len(navegador.window_handles) > 1:
            logging.info("❌ Fechando a aba do Gmail...")
            navegador.switch_to.window(navegador.window_handles[0])  # Voltar para o Gmail
            navegador.close()  # Fechar o Gmail
            
            # 🔄 Garantir que estamos na aba correta após o fechamento
            navegador.switch_to.window(navegador.window_handles[-1])
            logging.info("✅ Aba do Gmail fechada com sucesso!")

    except Exception as e:
        logging.error(f"❌ Erro ao abrir nova aba e fechar Gmail: {e}")


def entrar_no_portal(navegador, usuario, senha):
    try:
        # 🔘 Clica no botão de início
        botao_entrar = WebDriverWait(navegador, 5).until(
            EC.element_to_be_clickable((By.ID, "trauth-continue-signin-btn"))
        )
        botao_entrar.click()
        logging.info("✅ Entrando no portal")

        wait = WebDriverWait(navegador, 15)

        # 🧑 Campo de usuário
        logging.info("🔍 Aguardando campo de usuário...")
        campo_usuario = wait.until(EC.presence_of_element_located((By.NAME, "username")))
        campo_usuario.clear()
        campo_usuario.send_keys(usuario)

        # 👉 Botão após usuário
        botao_usuario = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//*[@type='submit']"))
        )
        botao_usuario.click()

        # 🔒 Campo de senha
        logging.info("🔐 Aguardando campo de senha...")
        campo_senha = wait.until(EC.presence_of_element_located((By.NAME, "password")))
        campo_senha.clear()
        campo_senha.send_keys(senha)

        # 👉 Botão após senha
        botao_senha = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//*[@type='submit']"))
        )
        botao_senha.click()

        # 🔄 Verificação em duas etapas
        try:
            wait_2fa = WebDriverWait(navegador, 10)
            wait_2fa.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//h1[text()='Selecione um método para verificar sua identidade']")
                )
            )
            logging.info("🔐 Autenticação de dois fatores detectada.")

            # 📧 Clica em "E-mail"
            botao_email = wait_2fa.until(
                EC.element_to_be_clickable((By.XPATH, "//button[@name='action' and @value='email::3']"))
            )
            botao_email.click()

            # Espera campo de código
            campo_codigo = WebDriverWait(navegador, 60).until(
                EC.presence_of_element_located((By.ID, "code"))
            )

            # Espera o e-mail chegar
            logging.info("📩 Aguardando código por e-mail...")
            time.sleep(20)  # opcional, se o e-mail costuma demorar
            codigo = extrair_codigo_do_email()

            if codigo:
                campo_codigo.clear()
                campo_codigo.send_keys(codigo)

                # Clica em "Continuar"
                botao_continuar = WebDriverWait(navegador, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[@name='action' and @value='default']"))
                )
                botao_continuar.click()
                logging.info("✅ Código inserido e confirmado.")
            else:
                raise Exception("Código de verificação não recebido.")
        except TimeoutException:
            logging.info("ℹ️ Nenhuma autenticação de dois fatores foi solicitada.")

        logging.info("✅ Login realizado com sucesso!")

    except Exception as e:
        logging.error(f"❌ Erro ao tentar fazer login no portal: {e}")
    

def acessar_aba_documentos(navegador, max_tentativas=3):
    """
    Percorre a lista de empresas e acessa a aba 'Documentos' para cada uma, garantindo que cada empresa seja processada antes de continuar.
    
    Parâmetros:
    - navegador: Instância do WebDriver.
    - max_tentativas: Número máximo de tentativas para carregar o site (padrão: 3).
    
    Retorna:
    - True se todas as empresas forem processadas com sucesso, False em caso de erro.
    """
    tentativa = 0
    
    while tentativa < max_tentativas:
        tentativa += 1
        try:
            print(f"⏳ Tentativa {tentativa}/{max_tentativas} para carregar o site...")

            # 🔄 Aguarda o carregamento completo do site
            WebDriverWait(navegador, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "dashboard-section__content"))
            )
            print("✅ Site carregado com sucesso!")

            # 🔍 Aguarda o menu lateral onde está a opção 'Documentos'
            aba_documentos = WebDriverWait(navegador, 30).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='documents']/a"))
            )
            
            # 📌 Clica na aba "Documentos"
            print("📂 Acessando a aba 'Documentos'...")
            navegador.execute_script("arguments[0].click();", aba_documentos)

            # 🔄 Espera dinamicamente a nova página carregar
            WebDriverWait(navegador, 10).until(
                EC.presence_of_element_located((By.XPATH, "//ul[contains(@class, 'nav-tabs')]//span[contains(text(), 'Documentos do Cliente')]"))
            )
            print("✅ Aba 'Documentos' carregada com sucesso!")

            # 🔄 Obtém a lista de códigos das empresas
            lista_empresas = obter_codigo_empresa()

            if not lista_empresas or not isinstance(lista_empresas, list):
                print("❌ Erro: Nenhuma empresa encontrada para processar.")
                return False

            # 🔄 Loop pelas empresas
            for codigo_empresa in lista_empresas:
                print(f"\n🔄 Iniciando o processamento da empresa com código: {codigo_empresa}")
                
                # Seleciona a empresa na lista e obtém o nome da empresa selecionada
                selecionado = selecionar_empresa(navegador, codigo_empresa)
                nome_empresa = vg.nome_empresa_selecionada
                codigo_empresa = vg.codigo_empresa_selecionada
                
                
                if selecionado and nome_empresa:
                    print(f"✅ Empresa '{nome_empresa}' (Código: {codigo_empresa}) processada com sucesso!")
                    time.sleep(2)  # Pequeno delay para finalizar o processamento atual
                else:
                    print(f"⚠ Empresa com código {codigo_empresa} não foi encontrada. Pulando para a próxima.")

                # ♻️ Reseta a variável global após o processamento de cada empresa
                if vg.nome_empresa_selecionada:
                    print(f"♻️ Resetando a variável 'nome_empresa_selecionada' de '{vg.nome_empresa_selecionada}' para None...")
                    vg.nome_empresa_selecionada = None
                    vg.codigo_empresa_selecionada = None
                else:
                    print("🔍 Nenhuma empresa estava selecionada anteriormente. Nenhum reset necessário.")
                    
            print("\n✅ Todas as empresas foram processadas com sucesso!")
            return True
        
        except Exception as e:
            print(f"❌ Erro ao carregar o site ou acessar a aba 'Documentos' na tentativa {tentativa}: {e}")
            if tentativa < max_tentativas:
                print("🔄 Recarregando a página...")
                navegador.refresh()
                time.sleep(5)  # Pausa para evitar loops rápidos demais

    print("🚫 Não foi possível acessar a aba 'Documentos' após várias tentativas.")
    return False


def selecionar_empresa(navegador, codigo_empresa):
    """
    Localiza e seleciona a empresa na lista de sugestões do Onvio.
    Armazena dinamicamente o nome da empresa selecionada.
    """
    try:
        campo_busca = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='text' and @placeholder='Selecione um cliente' and @aria-label='Cliente']"))
        )

        navegador.execute_script("arguments[0].value = '';", campo_busca)
        time.sleep(1)

        campo_busca.send_keys(str(codigo_empresa))
        print(f"✅ Código da empresa inserido no campo de busca: {codigo_empresa}")
        time.sleep(5)

        combobox_containers = navegador.find_elements(By.CLASS_NAME, "bento-combobox-container")
        if not combobox_containers:
            print("⚠️ Nenhuma combobox encontrada! Recarregando página e tentando novamente...")
            navegador.refresh()
            time.sleep(6)

            campo_busca = WebDriverWait(navegador, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='text' and @placeholder='Selecione um cliente' and @aria-label='Cliente']"))
            )
            navegador.execute_script("arguments[0].value = '';", campo_busca)
            time.sleep(1)
            campo_busca.send_keys(str(codigo_empresa))
            print(f"🔁 Repetindo busca pelo código: {codigo_empresa}")
            time.sleep(5)

            combobox_containers = navegador.find_elements(By.CLASS_NAME, "bento-combobox-container")
            if not combobox_containers:
                print("❌ Nenhuma combobox encontrada mesmo após recarregar!")
                return False

        anterior = -1

        while True:
            opcoes = navegador.find_elements(By.XPATH, "//li[contains(@class, 'bento-combobox-container-item')]")
            atual = len(opcoes)

            if atual == 0:
                print(f"⚠️ Nenhuma empresa foi carregada na lista. Pulando o código {codigo_empresa}...")
                break

            for opcao in opcoes:
                spans = opcao.find_elements(By.TAG_NAME, "span")
                if spans and spans[0].text.strip() == str(codigo_empresa):
                    nome_empresa = spans[1].text.strip() if len(spans) > 1 else "Nome não identificado"
                    print(f"✅ Empresa encontrada: {nome_empresa} - Selecionando...")

                    vg.nome_empresa_selecionada = nome_empresa
                    vg.codigo_empresa_selecionada = codigo_empresa

                    navegador.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center'});", opcao)

                    WebDriverWait(navegador, 5).until(
                        EC.element_to_be_clickable(opcao)
                    )

                    try:
                        opcao.click()
                    except Exception:
                        navegador.execute_script("arguments[0].click();", opcao)

                    print(f"✅ Empresa {codigo_empresa} selecionada com sucesso!")
                    acessar_pasta_financeiro(navegador)
                    return True

            if atual == anterior:
                print(f"❌ Empresa com código {codigo_empresa} não encontrada após carregar todas as opções.")
                break

            anterior = atual

            try:
                corpo_lista = navegador.find_element(By.CLASS_NAME, "bento-combobox-container-body")
                navegador.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", corpo_lista)
                print("🔄 Rolando para carregar mais empresas...")
                time.sleep(2)
            except Exception as e:
                print(f"⚠️ Erro ao tentar rolar lista: {e}")
                break

        print(f"❌ Empresa {codigo_empresa} não encontrada na lista!")
        return False

    except Exception as e:
        print(f"❌ Erro ao selecionar a empresa {codigo_empresa}: {str(e)}")
        return False

