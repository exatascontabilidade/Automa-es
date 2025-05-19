executando = True
usar_headless = False

class Redirector:
    def __init__(self):
        self.total_baixados = 0

    def gerar_relatorio_final(self):
        print(f" Relatório gerado. Total baixados: {self.total_baixados}")

redirector = Redirector()