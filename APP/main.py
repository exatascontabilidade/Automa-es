import webview
import os
import sys
import threading
import pandas as pd
import subprocess
from datetime import datetime
from tkinter import filedialog, Tk

class Backend:
    def __init__(self):
        self.planilhas = {}
        self.parar_execucao = threading.Event()

    def selecionar_planilha(self, nome_script):
        root = Tk()
        root.withdraw()
        caminho = filedialog.askopenfilename(
            title=f"Selecione a planilha para {nome_script}",
            filetypes=[("Arquivos Excel", "*.xlsx;*.xls")]
        )
        root.destroy()
        if caminho:
            self.planilhas[nome_script] = caminho
            return f" Planilha selecionada para {nome_script}: {caminho}"
        return " Nenhuma planilha selecionada."

    def executar_script(self, nome_script):
        if nome_script not in self.planilhas:
            return f" Nenhuma planilha selecionada para {nome_script}!"
        
        caminho_script = os.path.join(os.path.dirname(__file__), "scripts", f"{nome_script}.py")
        if not os.path.exists(caminho_script):
            return f" Script {nome_script}.py não encontrado."

        try:
            processo = subprocess.Popen(
                [sys.executable, caminho_script, self.planilhas[nome_script]],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="cp1252"
            )

            for linha in processo.stdout:
                linha_formatada = linha.strip()  # <-- Removido o prefixo do nome do script
                print(linha_formatada)
                webview.windows[0].evaluate_js(
                    f'appendLog("{nome_script}", `{linha_formatada}`)'
                )

            processo.wait()
            final_msg = f"✅ Script finalizado com código {processo.returncode}"
            webview.windows[0].evaluate_js(
                f'appendLog("{nome_script}", `{final_msg}`)'
            )

        except Exception as e:
            erro_msg = f"❌ Erro ao executar: {e}"
            webview.evaluate_js(f'appendLog("{nome_script}", `{erro_msg}`)')
            return erro_msg

if __name__ == '__main__':
    api = Backend()
    webview.create_window("Painel de Automação", "index.html", js_api=api, width=1024, height=768)
    webview.start(debug=False)
