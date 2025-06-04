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
        self.executando = False
        self.processo = None

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
        if nome_script != "mensal" and nome_script not in self.planilhas:
            return f" Nenhuma planilha selecionada para {nome_script}!"
        
        caminho_script = os.path.join(os.path.dirname(__file__), "scripts", f"{nome_script}.py")
        if not os.path.exists(caminho_script):
            return f" Script {nome_script}.py não encontrado."

        try:
            self.executando = True
            if nome_script == "download_xml":
                # Executa sem argumentos
                comando = [sys.executable, caminho_script]
            else:
                # Executa com argumento da planilha
                comando = [sys.executable, caminho_script, self.planilhas[nome_script]]
                
            self.processo = subprocess.Popen(
                comando,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="cp1252"
            )

            for linha in self.processo.stdout:
                if not self.executando:
                    self.processo.terminate()
                    break

                linha_formatada = linha.strip()
                print(linha_formatada)
                webview.windows[0].evaluate_js(
                    f'appendLog("{nome_script}", `{linha_formatada}`)'
                )

            self.processo.wait()
            self.executando = False
            final_msg = f"✅ Script finalizado com código {self.processo.returncode}"
            webview.windows[0].evaluate_js(
                f'appendLog("{nome_script}", `{final_msg}`)'
            )

        except Exception as e:
            erro_msg = f"❌ Erro ao executar: {e}"
            self.executando = False
            webview.windows[0].evaluate_js(
                f'appendLog("{nome_script}", `{erro_msg}`)'
            )
            return erro_msg

# ✅ Método chamado pelo botão "⛔ Encerrar Consulta"
    
    def encerrar_consulta(self, nome_script):
        print(f"[EXIT] Pedido de encerramento recebido para: {nome_script}")
        
        if nome_script == 'download_xml':
            return self.encerrar_dolowd()
        elif nome_script == 'relatorio':
            return self.encerrar_relatorio()
        elif nome_script == 'consulta_xml':
            return self.encerrar_sefaz()
        elif nome_script == 'demonstrativo':
            return self.encerrar_demo()
        else:
            return f"[ERRO] Script '{nome_script}' não reconhecido."

    def encerrar_dolowd(self):
        from scripts.dolowd.login import parar_automacao
        return parar_automacao()

    def encerrar_relatorio(self):
        return "⚠️ Encerramento do relatório ainda não implementado."

    def encerrar_sefaz(self):
        return "⚠️ Encerramento do SEFAZ ainda não implementado."
    def encerrar_demo(self):
        from scripts.demonstrativo import parar_consulta
        return parar_consulta()
if __name__ == '__main__':
    api = Backend()
    webview.create_window(
        "Controlla SEFAZ",
        "index.html",
        js_api=api,
        width=1024,
        height=768
    )

    webview.start(debug=False)
