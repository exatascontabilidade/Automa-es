import os
import json

# Garante que os arquivos fiquem no mesmo diretório deste script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATE_PATH = os.path.join(BASE_DIR, "automacao.state")
RELATORIO_PATH = os.path.join(BASE_DIR, "relatorio.json")

# === Estado persistente em arquivo ===

def set_estado(executando: bool):
    with open(STATE_PATH, "w") as f:
        f.write("on" if executando else "off")

def get_estado() -> bool:
    return os.path.exists(STATE_PATH) and open(STATE_PATH).read().strip() == "on"

def remover_estado():
    if os.path.exists(STATE_PATH):
        os.remove(STATE_PATH)

def salvar_relatorio(total_baixados: int):
    with open(RELATORIO_PATH, "w") as f:
        json.dump({"total_baixados": total_baixados}, f)

def ler_relatorio():
    if not os.path.exists(RELATORIO_PATH):
        return {"total_baixados": 0}
    with open(RELATORIO_PATH) as f:
        return json.load(f)

# === Variáveis em tempo real (em memória) ===

navegador = None

class Redirector:
    def __init__(self):
        self.total_baixados = 0

    def gerar_relatorio_final(self):
        print(f"[INFO] Relatório gerado. Total baixados: {self.total_baixados}")
        salvar_relatorio(self.total_baixados)

redirector = Redirector()
