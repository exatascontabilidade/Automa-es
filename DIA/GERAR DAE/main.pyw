import pandas as pd
import subprocess
import os
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import sys
from datetime import datetime

os.environ["PYTHONUTF8"] = "1"
root = tk.Tk()
root.title("Gearar Dae - Sefaz")
root.geometry("600x700")
arquivo_planilha = tk.StringVar()
parar_execucao = threading.Event()
def selecionar_planilha():
    caminho = filedialog.askopenfilename(
        title="Selecione a planilha", filetypes=[("Arquivos Excel", "*.xlsx;*.xls")]
    )
    if caminho:
        arquivo_planilha.set(caminho)
        log_output.insert("end", f"📂 Planilha selecionada: {caminho}\n")
        log_output.see("end")
def carregar_empresas(arquivo):
    colunas_necessarias = ["Inscrição Municipal", "Nome da Empresa","MES", "ANO"]
    try:
        df = pd.read_excel(arquivo, usecols=colunas_necessarias, dtype=str).fillna("")
        return df if not df.empty else None
    except Exception as e:
        log_output.insert("end", f"❌ Erro ao carregar a planilha: {e}\n")
        log_output.see("end")
        return None
def executar_consulta():
    if not arquivo_planilha.get():
        messagebox.showwarning("Erro", "Nenhuma planilha selecionada!")
        return
    df_empresas = carregar_empresas(arquivo_planilha.get())
    if df_empresas is None:
        messagebox.showerror("Erro", "Nenhuma empresa encontrada para processar. Verifique a planilha.")
        return
    threading.Thread(target=processar_consultas, args=(df_empresas,), daemon=True).start()



def processar_consultas(df_empresas):
    def log(mensagem):
        hora = datetime.now().strftime("[%H:%M:%S] ")
        log_output.insert("end", hora + mensagem + "\n")
        log_output.see("end")

    parar_execucao.clear()
    script_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    caminho_script = os.path.join(script_dir, "exec.py")

    if not os.path.exists(caminho_script):
        messagebox.showerror("Erro", "❌ Arquivo exec.py não encontrado!")
        return

    log("🔄 Iniciando consultas...")

    for index, row in df_empresas.iterrows():
        log(f"▶️ Processando empresa {index + 1} de {len(df_empresas)}")

        if parar_execucao.is_set():
            log("⛔ Execução interrompida pelo usuário.")
            return

        inscricao_municipal = row["Inscrição Municipal"].strip()
        nome_empresa = row["Nome da Empresa"].strip()

        if not inscricao_municipal:
            log(f"⚠️ Linha {index + 2}: Inscrição Municipal ausente. Pulando...")
            continue

        log(f"🔍 Consultando Empresa: {nome_empresa} com inscrição:{inscricao_municipal}...")

        mMES = row["MES"].strip()
        mANO = row["ANO"].strip()

        try:
            resultado = subprocess.run([
                sys.executable, caminho_script,
                inscricao_municipal, nome_empresa, mMES, mANO
            ], capture_output=True, text=True, encoding="utf-8", timeout=60)

            log(f"📜 Info:\n{resultado.stdout}")
            if resultado.stderr:
                log(f"⚠️ Erros: {resultado.stderr}")
                
            # ✅ Separa arquivos baixados   
            separar_arquivos_em_pdf_e_excel(nome_empresa)
        except subprocess.TimeoutExpired:
            log(f"⏳ Tempo excedido para {inscricao_municipal}. Pulando...")
        except Exception as e:
            log(f"❌ Erro: {e}")

    log("✅ Todas as consultas foram concluídas!")
    messagebox.showinfo("Concluído", "Todas as consultas foram processadas!")
    
def separar_arquivos_em_pdf_e_excel(nome_empresa):
    import shutil
    import os
    import sys

    script_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    pasta_temp = os.path.join(script_dir, "temp")
    os.makedirs(pasta_temp, exist_ok=True)

    for arquivo in os.listdir(pasta_temp):
        caminho_origem = os.path.join(pasta_temp, arquivo)

        if not os.path.isfile(caminho_origem):
            continue

        extensao = os.path.splitext(arquivo)[1].lower()
        nome_limpo = nome_empresa.replace(" ", "_").replace("/", "_")  # evita caracteres inválidos
        novo_nome = f"{nome_limpo}{extensao}"
        caminho_destino = os.path.join(pasta_temp, novo_nome)

        try:
            shutil.move(caminho_origem, caminho_destino)
            print(f"✅ Arquivo renomeado para {novo_nome} na pasta temp.")
        except Exception as e:
            print(f"❌ Erro ao renomear {arquivo}: {e}")

    
    
def encerrar_consulta():
    parar_execucao.set()
    log_output.insert("end", "\n⏹️ Interrompendo consultas...\n")
    log_output.see("end")
    messagebox.showinfo("Interrompido", "As consultas estão sendo interrompidas!")
frame = tk.Frame(root)
frame.pack(pady=10)
btn_selecionar = tk.Button(frame, text="Selecionar Planilha", command=selecionar_planilha)
btn_selecionar.pack(side=tk.LEFT, padx=5)
btn_executar = tk.Button(frame, text="Executar Consultas", command=executar_consulta)
btn_executar.pack(side=tk.LEFT, padx=5)
btn_encerrar = tk.Button(frame, text="Encerrar Consultas", command=encerrar_consulta, bg="red", fg="white")
btn_encerrar.pack(side=tk.LEFT, padx=5)
log_output = scrolledtext.ScrolledText(root, height=35, width=70)
log_output.pack(pady=10)
root.mainloop()
