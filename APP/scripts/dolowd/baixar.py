import time
import os
from pathlib import Path
from selenium.webdriver.common.by import By
import scripts.dolowd.state as state

def renomear_ultimo_arquivo(tipo_prefixo):
    pasta_download = os.path.join(os.path.dirname(__file__), "temp")
    os.makedirs(pasta_download, exist_ok=True)
    time.sleep(2)
    arquivos = list(Path(pasta_download).glob("*.zip"))
    if not arquivos:
        print("[INFO] Nenhum arquivo encontrado para renomear.")
        return
    ultimo = max(arquivos, key=os.path.getctime)
    nome_original = ultimo.name
    if not nome_original.startswith(("NFE_", "NFC_")):
        novo_nome = f"{tipo_prefixo}_{nome_original}"
        destino = Path(pasta_download) / novo_nome
        ultimo.rename(destino)
        print(f"[INFO] Arquivo renomeado para: {novo_nome}")

total_baixados = 0

def baixar_arquivos_na_pagina(navegador):
    global total_baixados
    linhas = navegador.find_elements(By.XPATH, "//tr[contains(@class, 'tr')]")

    prontos = 0
    for linha in linhas:
        try:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            if len(colunas) < 3:
                continue
            status = colunas[-1].text.strip()
            if status == "PRONTO PARA DOWNLOAD":
                prontos += 1
        except Exception:
            continue

    print(f"[INFO] Arquivos 'PRONTO PARA DOWNLOAD' nesta página: {prontos}")

    baixados_nessa_pagina = 0
    for linha in linhas:
        if not state.get_estado():
            break
        try:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            if len(colunas) < 3:
                continue
            status = colunas[-1].text.strip()
            tipo = colunas[1].text.strip()
            if status == "PRONTO PARA DOWNLOAD":
                link = linha.find_element(By.XPATH, ".//a[contains(text(), 'PRONTO PARA DOWNLOAD')]")
                navegador.execute_script("arguments[0].scrollIntoView();", link)
                link.click()
                print(f"[INFO] Download iniciado para tipo {tipo}")
                time.sleep(0.5)
                baixados_nessa_pagina += 1
                print(f"[INFO] Baixando arquivos {baixados_nessa_pagina}/{prontos}")
                navegador.back()
                renomear_ultimo_arquivo(tipo.upper())
                time.sleep(0.5)
        except Exception as e:
            print(f"[ERRO] Erro ao processar linha de download: {e}")

    total_baixados += baixados_nessa_pagina

def baixar_arquivos_com_blocos(navegador):
    global total_baixados
    total_baixados = 0

    while state.get_estado():
        try:
            print("[INFO] Verificando arquivos para download na página atual")
            baixar_arquivos_na_pagina(navegador)

            botoes = navegador.find_elements(By.XPATH, "//a[not(contains(text(),'Próximo')) and not(contains(text(),'|'))]")
            numeros_paginas = [botao.text.strip() for botao in botoes if botao.text.strip().isdigit()]

            for numero in numeros_paginas:
                if not state.get_estado():
                    print("[INFO] Interrupção antes de acessar próxima página.")
                    return
                print(f"[INFO] Acessando página {numero}")
                botao_pagina = navegador.find_element(By.XPATH, f"//a[text()='{numero}']")
                navegador.execute_script("arguments[0].scrollIntoView();", botao_pagina)
                botao_pagina.click()
                time.sleep(0.5)
                baixar_arquivos_na_pagina(navegador)

            if not state.get_estado():
                print("[INFO] Execução interrompida antes de avançar bloco.")
                return

            try:
                botao_proximo = navegador.find_element(By.XPATH, "//a[contains(text(), 'Próximo')]")
                navegador.execute_script("arguments[0].scrollIntoView();", botao_proximo)
                print("[INFO] Avançando para o próximo bloco de páginas...")
                botao_proximo.click()
                time.sleep(0.5)
            except:
                print("[INFO] Fim da paginação. Todos os blocos foram verificados.")
                break

        except Exception as e:
            print(f"[ERRO] Erro inesperado no bloco: {e}")
            break

    print(f"[INFO] Total de arquivos baixados: {total_baixados}")
    state.redirector.total_baixados = total_baixados
    state.redirector.gerar_relatorio_final()
    state.remover_estado()
