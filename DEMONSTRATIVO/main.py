import os
import sys
import pandas as pd
import subprocess
from datetime import datetime
import shutil

# Configurações iniciais
os.environ["PYTHONUTF8"] = "1"

parar_execucao = False

def log(mensagem):
    hora = datetime.now().strftime("[%H:%M:%S] ")
    print(hora + mensagem)

def selecionar_planilha_terminal():
    caminho = input("Digite o caminho completo da planilha Excel (.xlsx ou .xls): ").strip()
    if not os.path.isfile(caminho):
        log("❌ Caminho inválido ou arquivo não encontrado.")
        return None
    return caminho

def carregar_empresas(arquivo):
    colunas_necessarias = ["Inscrição Municipal", "Nome da Empresa", "MES", "ANO", "FORMATO"]
    try:
        df = pd.read_excel(arquivo, usecols=colunas_necessarias, dtype=str).fillna("")
        return df if not df.empty else None
    except Exception as e:
        log(f"❌ Erro ao carregar a planilha: {e}")
        return None

def processar_consultas(df_empresas):
    global parar_execucao
    script_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    caminho_script = os.path.join(script_dir, "demonstrativo.py")

    if not os.path.exists(caminho_script):
        log("❌ Arquivo consulta_sefaz.py não encontrado!")
        return

    log("🔄 Iniciando consultas... Pressione CTRL+C para interromper.")

    for index, row in df_empresas.iterrows():
        log(f"▶️ Processando empresa {index + 1} de {len(df_empresas)}")

        if parar_execucao:
            log("⛔ Execução interrompida pelo usuário.")
            return

        inscricao_municipal = row["Inscrição Municipal"].strip()
        nome_empresa = row["Nome da Empresa"].strip()

        if not inscricao_municipal:
            log(f"⚠️ Linha {index + 2}: Inscrição Municipal ausente. Pulando...")
            continue

        log(f"🔍 Consultando Empresa: {nome_empresa} com inscrição: {inscricao_municipal}...")

        mMES = row["MES"].strip()
        mANO = row["ANO"].strip()
        formato_arquivo = row["FORMATO"].strip()

        try:
            resultado = subprocess.run([
                sys.executable, caminho_script,
                inscricao_municipal, nome_empresa, mMES, mANO, formato_arquivo
            ], capture_output=True, text=True, encoding="utf-8", timeout=60)

            log(f"📜 Info:\n{resultado.stdout}")
            if resultado.stderr:
                log(f"⚠️ Erros:\n{resultado.stderr}")

            separar_arquivos_em_pdf_e_excel()
        except subprocess.TimeoutExpired:
            log(f"⏳ Tempo excedido para {inscricao_municipal}. Pulando...")
        except Exception as e:
            log(f"❌ Erro: {e}")

    log("✅ Todas as consultas foram concluídas!")

def separar_arquivos_em_pdf_e_excel():
    script_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    pasta_temp = os.path.join(script_dir, "temp")
    pasta_pdfs = os.path.join(script_dir, "pdfs")
    pasta_excel = os.path.join(script_dir, "excel")

    os.makedirs(pasta_pdfs, exist_ok=True)
    os.makedirs(pasta_excel, exist_ok=True)

    for arquivo in os.listdir(pasta_temp):
        caminho_origem = os.path.join(pasta_temp, arquivo)
        if arquivo.lower().endswith(".pdf"):
            shutil.move(caminho_origem, os.path.join(pasta_pdfs, arquivo))
        elif arquivo.lower().endswith((".xls", ".xlsx")):
            shutil.move(caminho_origem, os.path.join(pasta_excel, arquivo))

if __name__ == "__main__":
    try:
        caminho = selecionar_planilha_terminal()
        if caminho:
            df_empresas = carregar_empresas(caminho)
            if df_empresas is not None:
                processar_consultas(df_empresas)
            else:
                log("⚠️ Nenhuma empresa encontrada para processar.")
    except KeyboardInterrupt:
        parar_execucao = True
        log("⏹️ Interrupção solicitada. Finalizando...")
