import os
import re
import time
import json
import glob
import pandas as pd# ✅ Necessária para leitura do empresas.json

# Bibliotecas de terceiros
import pygetwindow as gw
import autoit
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Módulo do seu projeto
import variaveis_globais as vg  

def clicar_botao_upload(navegador):
    """
    Clica no botão de upload na interface web e detecta a janela 'Abrir' usando PyAutoIt.
    """
    try:
        print("📂 Tentando clicar no botão de upload...")

        # Aguarda até que o botão esteja visível e clicável no navegador Chrome via Selenium
        botao_upload = WebDriverWait(navegador, 15).until(
            EC.element_to_be_clickable((By.ID, "dms-fe-legacy-components-client-documents-upload-button"))
        )

        # Garante que o botão esteja visível e clica nele
        navegador.execute_script("arguments[0].scrollIntoView();", botao_upload)
        botao_upload.click()
        print("✅ Botão de upload clicado com sucesso!")

        # Atraso adicional para a janela de diálogo abrir corretamente
        time.sleep(2)

        # Detecta a janela "Abrir" usando pygetwindow
        print("🔍 Procurando pela janela 'Abrir' no sistema...")
        janela_abrir = next((w for w in gw.getAllTitles() if "Abrir" in w), None)

        if not janela_abrir:
            print("❌ Nenhuma janela com o título 'Abrir' encontrada.")
            return

        print(f"🪟 Janela 'Abrir' detectada com o título: {janela_abrir}")

        # Chama a função para selecionar os arquivos sequenciais e únicos na janela detectada
        selecionar_todos_arquivos(janela_abrir)

    except Exception as e:
        print(f"❌ Erro ao clicar no botão de upload ou detectar a janela: {str(e)}")


def selecionar_todos_arquivos(titulo_janela: str):
    """
    Seleciona automaticamente todos os PDFs, incluindo sequenciais e únicos,
    com base no CNPJ correspondente ao código da empresa atual (usando planilha Excel).
    """
    codigo_empresa = vg.codigo_empresa_selecionada
    vg.vencimento_arquivo_atual = None  # ♻️ Reset
    vg.nomes_pdfs_enviados = None       # ♻️ Reset

    if not codigo_empresa:
        print("❌ Nenhum código de empresa selecionado.")
        return

    diretorio_script = os.path.dirname(os.path.abspath(__file__))
    pasta_temp = os.path.join(diretorio_script, "temp")

    # 🧾 Localiza o primeiro arquivo Excel (.xls ou .xlsx)
    arquivos_excel = glob.glob(os.path.join(pasta_temp, "*.xls")) + \
                     glob.glob(os.path.join(pasta_temp, "*.xlsx"))

    if not arquivos_excel:
        print("❌ Nenhum arquivo Excel encontrado na pasta 'temp'.")
        return

    excel_path = arquivos_excel[0]
    extensao = os.path.splitext(excel_path)[-1].lower()
    engine = "xlrd" if extensao == ".xls" else "openpyxl"

    try:
        df = pd.read_excel(excel_path, engine=engine)
        if "Cód." not in df.columns or "CNPJ" not in df.columns:
            print("❌ A planilha deve conter as colunas 'Cód.' e 'CNPJ'.")
            return

        df["Cód."] = df["Cód."].astype(str)
        df["CNPJ"] = df["CNPJ"].astype(str).str.replace(r'\D', '', regex=True).str.zfill(14)

        cnpj_empresa = df.loc[df["Cód."] == str(codigo_empresa), "CNPJ"].squeeze()

        if pd.isna(cnpj_empresa):
            print(f"❌ CNPJ não encontrado para o código {codigo_empresa}.")
            return

    except Exception as e:
        print(f"❌ Erro ao ler a planilha: {e}")
        return

    print(f"📌 Código selecionado: {codigo_empresa} → CNPJ: {cnpj_empresa}")

    padrao_sequencial = re.compile(r'\((\d+) de (\d+)\)')
    padrao_cnpj_no_nome = re.compile(rf'cnpj_{cnpj_empresa}', re.IGNORECASE)
    arquivos = os.listdir(pasta_temp)

    pdfs_sequenciais = [
        f for f in arquivos
        if f.lower().endswith(".pdf")
        and padrao_cnpj_no_nome.search(f)
        and padrao_sequencial.search(f)
    ]

    pdfs_unicos = [
        f for f in arquivos
        if f.lower().endswith(".pdf")
        and padrao_cnpj_no_nome.search(f)
        and not padrao_sequencial.search(f)
    ]

    if not pdfs_sequenciais and not pdfs_unicos:
        print(f"❌ Nenhum PDF encontrado com CNPJ '{cnpj_empresa}' na pasta 'temp'.")
        autoit.win_close(titulo_janela)
        print("🛑 Janela 'Abrir' fechada.")
        return

    pdfs_ordenados = sorted(
        pdfs_sequenciais, key=lambda f: int(padrao_sequencial.search(f).group(1))
    ) + pdfs_unicos

    print(f"📂 Arquivos encontrados para upload: {pdfs_ordenados}")

    # 🔍 Extração do vencimento antes do envio
    for nome_arquivo in pdfs_ordenados:
        if extrair_vencimento_do_nome(nome_arquivo):
            break

    print(f"📆 Vencimento armazenado na variável global: {vg.vencimento_arquivo_atual}")

    # 🗂️ Armazena os nomes dos arquivos enviados na variável global
    armazenar_nomes_pdfs_enviados(pdfs_ordenados)
    print(f"📄 Nomes dos PDFs armazenados na variável global: {vg.nomes_pdfs_enviados}")

    try:
        autoit.win_wait_active(titulo_janela, timeout=10)

        caminhos_arquivos = [
            f'"{os.path.join(pasta_temp, nome)}"' for nome in pdfs_ordenados
        ]
        caminho_arquivos = " ".join(caminhos_arquivos)

        autoit.control_set_text(titulo_janela, "Edit1", caminho_arquivos)
        time.sleep(1)
        autoit.control_send(titulo_janela, "Edit1", "{ENTER}")
        time.sleep(2)

        print("⏳ Aguardando envio dos PDFs...")
        print("✅ Upload realizado com sucesso!")

    except Exception as e:
        print(f"❌ Erro ao enviar arquivos: {e}")
        autoit.win_close(titulo_janela)
        print("🛑 Janela 'Abrir' fechada.")
        
def extrair_vencimento_do_nome(nome_arquivo):
    """
    Extrai o vencimento no padrão VENC_XX-XX-XXXX do nome do arquivo e armazena na variável global.
    """
    match = re.search(r'venc[_\-]?(\d{2}[-/]\d{2}[-/]\d{4})', nome_arquivo, re.IGNORECASE)
    if match:
        vencimento = match.group(1).replace("/", "-")
        vg.vencimento_arquivo_atual = vencimento
        print(f"📆 Vencimento extraído do nome do arquivo: {vencimento}")
        return vencimento
    else:
        print(f"⚠️ Nenhum vencimento encontrado no nome: {nome_arquivo}")
        vg.vencimento_arquivo_atual = None
        
        return None

def armazenar_nomes_pdfs_enviados(lista_de_pdfs):
    """
    Armazena os nomes dos PDFs enviados em uma variável global, separados por '/////'.
    """
    vg.nomes_pdfs_enviados = None  #
    if not lista_de_pdfs:
        vg.nomes_pdfs_enviados = None
        print("⚠️ Nenhum PDF armazenado na variável global.")
        return

    nomes_formatados = ' ///// '.join(lista_de_pdfs)
    vg.nomes_pdfs_enviados = nomes_formatados
    print(f"📄 Nomes dos PDFs enviados armazenados:\n{vg.nomes_pdfs_enviados}")    