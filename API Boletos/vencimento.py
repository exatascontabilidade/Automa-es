from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import variaveis_globais as vg
import time

def definir_vencimento_para_boletos(navegador):
    """
    Para cada PDF na variável global, localiza sua linha na interface,
    marca a checkbox, abre o menu Gerenciar e define a data de vencimento.
    """
    if not vg.nomes_pdfs_enviados:
        print("❌ Nenhum nome de PDF armazenado.")
        return

    nomes = [n.strip() for n in vg.nomes_pdfs_enviados.split("/////")]
    for nome_pdf in nomes:
        try:
            print(f"\n🔍 Processando: {nome_pdf}")

            # Espera o carregamento do nome do PDF
            elemento_nome = WebDriverWait(navegador, 15).until(
                EC.presence_of_element_located((By.XPATH, f"//a[contains(text(), '{nome_pdf.strip()}')]"))
            )

            # Localiza a linha pai do PDF
            linha = elemento_nome.find_element(By.XPATH, "./ancestor::div[contains(@class, 'wj-row')]")

            # Clica na checkbox correspondente
            checkbox = linha.find_element(By.CSS_SELECTOR, "i.bento-flex-grid-checkbox")
            navegador.execute_script("arguments[0].click();", checkbox)
            print("☑️ Checkbox marcada.")

            # Clica no botão 'Gerenciar'
            botao_gerenciar = WebDriverWait(navegador, 10).until(
                EC.element_to_be_clickable((By.ID, "dms-fe-legacy-components-client-documents-manage-docs-menu-button"))
            )
            botao_gerenciar.click()
            print("📂 Menu 'Gerenciar' aberto.")

            # Clica em "Definir data de vencimento"
            definir_vencimento = WebDriverWait(navegador, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Definir data de vencimento')]"))
            )
            definir_vencimento.click()
            print("📅 Opção 'Definir data de vencimento' clicada.")

            # Preenche o campo de data com a variável global
            campo_data = WebDriverWait(navegador, 10).until(
                EC.presence_of_element_located((By.ID, "dueDate"))
            )
            campo_data.clear()
            campo_data.send_keys(vg.vencimento_arquivo_atual)
            print(f"📆 Data preenchida: {vg.vencimento_arquivo_atual}")

            # Clica em Salvar (assumindo botão com texto ou ID identificável)
            botao_salvar = WebDriverWait(navegador, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Salvar')]"))
            )
            botao_salvar.click()
            print("💾 Data de vencimento salva com sucesso!")

            # Aguarda um pouco para seguir para o próximo
            time.sleep(2)

        except Exception as e:
            print(f"❌ Erro ao processar '{nome_pdf}': {e}")
