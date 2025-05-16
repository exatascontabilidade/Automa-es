from datetime import datetime
from pathlib import Path
import os

class RedirectText:
    def __init__(self, widget):
        self.widget = widget

        # Garante que a pasta 'log/' existe
        log_dir = Path("log")
        log_dir.mkdir(exist_ok=True)

        # Define o caminho do log
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.log_file_path = log_dir / f"logs_automacao_{timestamp}.txt"
        self.error_log_path = log_dir / f"erros_automacao_{timestamp}.txt"
        self.total_baixados = 0
        self.inicio = datetime.now()

        with open(self.log_file_path, "w", encoding="utf-8") as f:
            f.write("=== LOG DE EXECUÇÃO ===\n")

    def write(self, message):
        self.widget.insert("end", message)
        self.widget.see("end")
        with open(self.log_file_path, "a", encoding="utf-8") as f:
            f.write(message)
        if "❌" in message or "⚠️" in message:
            with open(self.error_log_path, "a", encoding="utf-8") as f:
                f.write(message)

    def flush(self):
        pass
    def gerar_relatorio_final(self):
        fim = datetime.now()
        duracao = fim - self.inicio
        with open(self.log_file_path, "a", encoding="utf-8") as f:
            f.write("\n=== RELATÓRIO FINAL ===\n")
            f.write(f"Total de arquivos baixados: {self.total_baixados}\n")
            f.write(f"Duração da operação: {duracao}\n")