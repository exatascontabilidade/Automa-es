import tkinter as tk
from tkinter import scrolledtext
import sys
import threading
from utils.logger import RedirectText
from automacoes.login import executar_codigo_completo, parar_automacao, iniciar_thread
import automacoes.state as state

def iniciar_interface():
    janela = tk.Tk()
    janela.title("Automação SEFAZ - Interface Intuitiva")
    janela.geometry("850x650")

    headless_var = tk.BooleanVar(value=False)

    def atualizar_headless():
        state.usar_headless = headless_var.get()
        print(f"🔧 Modo headless definido como: {state.usar_headless}")

    frame_topo = tk.Frame(janela)
    frame_topo.pack(pady=10)

    botao_iniciar = tk.Button(frame_topo, text="▶ Iniciar Automação", command=iniciar_thread,
                              font=("Segoe UI", 12, "bold"), bg="green", fg="white", width=18)
    botao_iniciar.grid(row=0, column=0, padx=10)

    botao_parar = tk.Button(frame_topo, text="⛔ Parar Automação", command=parar_automacao,
                            font=("Segoe UI", 12, "bold"), bg="red", fg="white", width=18)
    botao_parar.grid(row=0, column=1, padx=10)

    check_headless = tk.Checkbutton(
        frame_topo,
        text="Modo Headless (oculto)",
        variable=headless_var,
        command=atualizar_headless,
        font=("Segoe UI", 11)
    )
    check_headless.grid(row=0, column=2, padx=10)

    log_area = scrolledtext.ScrolledText(janela, wrap=tk.WORD, font=("Consolas", 10))
    log_area.pack(expand=True, fill='both', padx=10, pady=10)

    sys.stdout = RedirectText(log_area)
    state.redirector = sys.stdout
    janela.mainloop()
