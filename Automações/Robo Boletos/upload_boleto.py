import datetime
from datetime import datetime
import json
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from subir_arquivos import clicar_botao_upload
from vencimento import definir_vencimento_para_boletos
import variaveis_globais as vg  # Importa o módulo de variáveis globais

# Garante que estamos pegando o diretório correto do script em execução
DIRETORIO_PROJETO = os.path.dirname(os.path.abspath(__file__))  # Obtém o diretório do script
DIRETORIO = os.path.join(DIRETORIO_PROJETO, "temp")  # Define a pasta "temp"
ARQUIVO_LOG = os.path.join(DIRETORIO, "processamento.json")

def acessar_pasta_financeiro(navegador):
    """
    Acessa a pasta 'Financeiro' no menu lateral dentro do Shadow DOM.
    Se encontrar a pasta, clica nela e continua o fluxo para verificar 'Boletos'.
    """
    time.sleep(5)
    try:
        shadow_host = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "bm-tree"))
        )

        shadow_root = navegador.execute_script("return arguments[0].shadowRoot", shadow_host)

        financeiro_element = WebDriverWait(shadow_root, 10).until(
            lambda sr: sr.find_element(By.CSS_SELECTOR, "bm-tree-item[name='Financeiro']")
        )

        navegador.execute_script("arguments[0].click();", financeiro_element)
        print("✅ Pasta 'Financeiro' acessada com sucesso!")

        time.sleep(3)

        if not encontrar_e_clicar_pasta_boletos(navegador, shadow_root):
            criar_pasta_boletos(navegador, shadow_root)

        return shadow_root

    except Exception as e:
        print(f"❌ Erro ao acessar a pasta 'Financeiro': {str(e)}")
        return None

def encontrar_e_clicar_pasta_boletos(navegador, shadow_root):
    """
    Verifica se existe uma pasta com nome semelhante a "Boletos" dentro de "Financeiro" e clica nela.
    Retorna True se encontrou e acessou, False caso contrário.
    """
    try:
        pastas = shadow_root.find_elements(By.CSS_SELECTOR, "bm-tree-item")

        for pasta in pastas:
            nome_pasta = pasta.get_attribute("name").strip().lower()
            if nome_pasta in ["boletos", "boleto"]:
                print(f"✅ Pasta encontrada: {pasta.get_attribute('name')} - Clicando...")
                navegador.execute_script("arguments[0].click();", pasta)
                
                # Verifica o nome da empresa selecionada na variável global
                if vg.nome_empresa_selecionada:
                    print(f"📂 Empresa selecionada: {vg.nome_empresa_selecionada}")
                else:
                    print("⚠️ Nenhuma empresa selecionada. A variável global está vazia.")
                
            
                time.sleep(5)
        
                clicar_botao_upload(navegador)
                verificar_popup_erro(navegador)
                definir_vencimento_para_boletos(navegador)
                return True

        print("⚠️ Pasta 'Boletos' não encontrada.")
        return False

    except Exception as e:
        print(f"❌ Erro ao acessar a pasta 'Boletos': {str(e)}")
        return False

def criar_pasta_boletos(navegador, shadow_root):
    """
    Cria a pasta 'Boletos' dentro de 'Financeiro' caso ela não exista.
    """
    try:
        print("📂 Criando a pasta 'Boletos'...")

        botao_novo = WebDriverWait(navegador, 5).until(
            EC.element_to_be_clickable((By.ID, "dms-fe-legacy-components-client-documents-new-menu-button"))
        )
        navegador.execute_script("arguments[0].click();", botao_novo)
        time.sleep(2)

        botao_pasta = WebDriverWait(navegador, 5).until(
            EC.element_to_be_clickable((By.ID, "dms-fe-legacy-components-client-documents-new-folder-button"))
        )
        navegador.execute_script("arguments[0].click();", botao_pasta)

        modal_input = WebDriverWait(navegador, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[data-qe-id='dmsNewContainerModal-nameField']"))
        )

        modal_input.send_keys("Boletos")

        botao_salvar = WebDriverWait(navegador, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-qe-id='dmsNewContainerModal-saveButton']"))
        )
        navegador.execute_script("arguments[0].click();", botao_salvar)

        print("✅ Pasta 'Boletos' criada com sucesso!")
        time.sleep(5)
        
        return acessar_pasta_boletos_novo(navegador)

    except Exception as e:
        print(f"❌ Erro ao criar a pasta 'Boletos': {str(e)}")
        return False

def acessar_pasta_boletos_novo(navegador):
    """
    Acessa a pasta 'Boletos' clicando no primeiro link dentro da estrutura fornecida.
    """
    try:
        print("📂 Acessando a pasta 'Boletos'...")

        pasta_boletos_link = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'Folder') and contains(text(), 'Boletos')]"))
        )
        
        navegador.execute_script("arguments[0].click();", pasta_boletos_link)
        
        if vg.nome_empresa_selecionada:
            print(f"📂 Empresa selecionada: {vg.nome_empresa_selecionada}")
        else:
            print("⚠️ Nenhuma empresa selecionada. A variável global está vazia.")
        
        print("✅ Pasta 'Boletos' acessada com sucesso!")
        time.sleep(5)
        
        clicar_botao_upload(navegador)
        verificar_popup_erro(navegador)
        return True

    except Exception as e:
        print(f"❌ Erro ao acessar a pasta 'Boletos': {str(e)}")
        return False


def salvar_resultado_upload(info_upload):
    """
    Salva o resultado do upload no JSON de processamento.
    """
    try:
        if os.path.exists(ARQUIVO_LOG):
            with open(ARQUIVO_LOG, "r", encoding="utf-8") as file:
                dados = json.load(file)
        else:
            dados = []
            
        info_upload["data_execucao"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        dados.append(info_upload)

        with open(ARQUIVO_LOG, "w", encoding="utf-8") as file:
            json.dump(dados, file, ensure_ascii=False, indent=4)

    except Exception as e:
        print(f"❌ Erro ao salvar no JSON: {str(e)}")


def verificar_popup_erro(navegador):
    """
    Verifica se um pop-up de erro aparece após o upload do arquivo.
    Salva no JSON se houve sucesso ou erro e fecha o alerta, se necessário.
    """
    try:
        print("🔍 Verificando se há erro no upload...")

        # Obtém a empresa atualmente processada
        codigo_empresa = vg.codigo_empresa_selecionada if hasattr(vg, "codigo_empresa_selecionada") else "Desconhecido"
        nome_empresa = vg.nome_empresa_selecionada if hasattr(vg, "nome_empresa_selecionada") else "Desconhecido"
        

        try:
            popup_erro = WebDriverWait(navegador, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.alert-error .file-alert-text"))
            )

            mensagem_erro = popup_erro.text.strip()

            if mensagem_erro:
                print(f"❌ Erro detectado: {mensagem_erro}")

                # Salva no JSON como erro
                erro_info = {
                    "codigo": codigo_empresa,
                    "nome": nome_empresa,
                    "status": "Erro",
                    "erro": mensagem_erro,
                }
                salvar_resultado_upload(erro_info)

                # Fecha o pop-up de erro
                try:
                    botao_fechar = navegador.find_element(By.CSS_SELECTOR, "button.bento-alert-close")
                    navegador.execute_script("arguments[0].click();", botao_fechar)
                    print("✅ Pop-up de erro fechado.")
                except:
                    print("⚠️ Não foi possível fechar o pop-up automaticamente.")

                return mensagem_erro

        except:
            print("✅ Nenhum erro detectado no upload.")

            # Salva no JSON como sucesso
            sucesso_info = {
                "codigo": codigo_empresa,
                "nome": nome_empresa,
                "status": "Sucesso",
                "erro": None
            }
            salvar_resultado_upload(sucesso_info)

            return None

    except Exception as e:
        print(f"❌ Erro inesperado ao verificar pop-up de erro: {str(e)}")