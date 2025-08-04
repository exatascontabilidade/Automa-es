import sys
import os
import pandas as pd
import requests
import re
from datetime import datetime, timedelta
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLabel, QComboBox, QMessageBox
)
from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import QLabel, QWidget, QVBoxLayout, QFileDialog, QPushButton
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QTextEdit, QScrollArea


# Credenciais da API Omie
APP_KEY = "4138937861058"
APP_SECRET = "ad9bf03a6aa1f6bc11570a1640c7c090"

def buscar_clientes():
    clientes = []
    pagina = 1
    while True:
        payload = {
            "call": "ListarClientes",
            "app_key": APP_KEY,
            "app_secret": APP_SECRET,
            "param": [{"pagina": pagina, "registros_por_pagina": 50, "apenas_importado_api": "N"}]
        }
        r = requests.post("https://app.omie.com.br/api/v1/geral/clientes/", json=payload)
        data = r.json()
        clientes += data.get("clientes_cadastro", [])
        if pagina * 50 >= data.get("total_de_registros", 0):
            break
        pagina += 1
    return clientes

def buscar_ultimo_codigo_pedido_integracao():
    pagina = 1
    maior_num = 0
    padrao = re.compile(r"PED_(\d+)")

    while True:
        payload = {
            "call": "ListarPedidos",
            "app_key": APP_KEY,
            "app_secret": APP_SECRET,
            "param": [{
                "pagina": pagina,
                "registros_por_pagina": 100,
                "apenas_importado_api": "N"
            }]
        }
        r = requests.post("https://app.omie.com.br/api/v1/produtos/pedido/", json=payload)
        data = r.json()

        pedidos = data.get("pedido_venda_produto", [])
        if not pedidos:
            break

        for pedido in pedidos:
            codigo = pedido.get("cabecalho", {}).get("codigo_pedido_integracao", "")
            match = padrao.match(codigo)
            if match:
                numero = int(match.group(1))
                maior_num = max(maior_num, numero)

        if pagina * 100 >= data.get("total_de_registros", 0):
            break
        pagina += 1

    return maior_num

def encontrar_cliente_por_cpf(clientes, cpf):
    cpf = ''.join(filter(str.isdigit, str(cpf)))
    for cli in clientes:
        if 'cnpj_cpf' in cli and ''.join(filter(str.isdigit, cli['cnpj_cpf'])) == cpf:
            return cli
    return None

def enviar_pedido(cliente, email, valor, numero):
    data_previsao = (datetime.today() + timedelta(days=1)).strftime("%d/%m/%Y")
    payload = {
        "call": "IncluirPedido",
        "app_key": APP_KEY,
        "app_secret": APP_SECRET,
        "param": [
            {
                "cabecalho": {
                    "codigo_cliente": cliente["codigo_cliente_omie"],
                    "codigo_pedido": "",
                    "codigo_pedido_integracao": f"PED_{str(numero).zfill(2)}",
                    "data_previsao": data_previsao,
                    "etapa": "10"
                },
                "det": [
                    {
                        "ide": {
                            "codigo_item_integracao": "PRD00001",
                            "simples_nacional": "N"
                        },
                        "produto": {
                            "codigo_produto": 4938559998,
                            "descricao": "FORMAÇÃO EM HARMONIZAÇÃO FACIAL",
                            "unidade": "UND",
                            "quantidade": 1,
                            "valor_unitario": valor,
                            "tipo_desconto": "P",
                            "percentual_desconto": 0,
                            "valor_total": valor
                        },
                        "inf_adic": {
                            "codigo_categoria_item": "1.01.01",
                            "codigo_cenario_impostos_item": "4938562213",
                            "codigo_local_estoque": 4936863631,
                            "nao_gerar_financeiro": "S",
                            "nao_movimentar_estoque": "S",
                            "nao_somar_total": "N"
                        },
                        "observacao": {}
                    }
                ],
                "frete": {"modalidade": "9"},
                "informacoes_adicionais": {
                    "codigo_categoria": "1.01.01",
                    "codigo_conta_corrente": 4936863626,
                    "consumidor_final": "S",
                    "enviar_email": "S",
                    "dados_adicionais_nf": "PRODUTO COM IMUNIDADE TRIBUTARIA CONFORME ALINEA D, DO INCISO VI, DO ARTIGO 150 DA CF/88|Produto destinado a Consumidor Final.",
                    "utilizar_emails": email or "financeiro@vortex.com.br"
                }
            }
        ]
    }
    r = requests.post("https://app.omie.com.br/api/v1/produtos/pedido/", json=payload)
    return r.json()

class ProcessarThread(QThread):
    finished = Signal(str)
    erro = Signal(str)
    progresso = Signal(int)


    def __init__(self, arquivo, aba):
        super().__init__()
        self.arquivo = arquivo
        self.aba = aba



# Processo Principal 
    def run(self):
        try:
            self.progresso.emit(5)
            df = pd.read_excel(self.arquivo, sheet_name=self.aba, header=[0, 1])
            df.columns = [' '.join([str(i).strip() for i in col if pd.notna(i)]) for col in df.columns]
            df.columns = [col.strip().upper() for col in df.columns]

            coluna_turma = 'CONTROLE, TRANSFERÊNCIA E EMISSÃO DE NFS-E TURMA'
            coluna_cpf = 'CONTROLE, TRANSFERÊNCIA E EMISSÃO DE NFS-E CPF'
            coluna_obs = 'CONTROLE, TRANSFERÊNCIA E EMISSÃO DE NFS-E OBSERVAÇÃO'
            coluna_valor = 'CONTROLE, TRANSFERÊNCIA E EMISSÃO DE NFS-E FISCAI NORMAIS'
            coluna_cliente = 'CONTROLE, TRANSFERÊNCIA E EMISSÃO DE NFS-E CLIENTE'
            
            
            self.progresso.emit(10)
            for col in [coluna_turma, coluna_cpf, coluna_obs, coluna_valor]:
                if col not in df.columns:
                    self.erro.emit(f"A coluna '{col}' não foi encontrada na planilha.")
                    return

            # Ignorar linhas com observação preenchida
            self.progresso.emit(20)
            linhas_ignoradas_obs = df[df[coluna_obs].notna()]
            df = df[df[coluna_obs].isna()]
            
            # Ignorar linhas com valor em branco
            linhas_ignoradas_valor = df[df[coluna_valor].isna()]
            df = df[df[coluna_valor].notna()]

            self.progresso.emit(30)
            clientes = buscar_clientes()
            
            self.progresso.emit(40)
            numero_atual = buscar_ultimo_codigo_pedido_integracao()
            total_enviados = 0
            msg = ""

            erros_por_turma = {}
            enviados_por_turma = {}

            turmas = df[coluna_turma].dropna().unique()
            total_turmas = len(turmas)
            turma_atual = 0

            for turma in turmas:
                turma_df = df[df[coluna_turma] == turma]
                enviados = 0
                erros = []
                turma_atual += 1
                progresso_percentual = int((turma_atual / total_turmas) * 95)
                self.progresso.emit(min(progresso_percentual, 95))

                for idx, row in turma_df.iterrows():
                    cpf = row[coluna_cpf]
                    valor = row[coluna_valor]
                    cliente = encontrar_cliente_por_cpf(clientes, cpf)
                    if cliente:
                        email = cliente.get("email", "")
                        numero = numero_atual + total_enviados + 1
                        try:
                            resposta = enviar_pedido(cliente, email, valor, numero)
                            if resposta.get("codigo_status") == "0":
                                enviados += 1
                                total_enviados += 1
                            else:
                                erros.append((row.get(coluna_cliente, f"Linha {idx+2}"), resposta.get("descricao_status")))
                        except Exception as e:
                            erros.append((row.get(coluna_cliente, f"Linha {idx+2}"), str(e)))
                    else:
                        erros.append((row.get(coluna_cliente, f"Linha {idx+2}"), "Cliente não encontrado"))

                enviados_por_turma[turma] = enviados
                erros_por_turma[turma] = erros

            # Mensagem final
            self.progresso.emit(100)
            msg += f"✅ Total de pedidos enviados: {total_enviados}\n"

            for turma in turmas:
                msg += f"\n📘 Turma {turma}:\n"
                msg += f"   - Pedidos enviados com sucesso: {enviados_por_turma.get(turma, 0)}\n"
                if erros_por_turma.get(turma):
                    msg += "   - Erros:\n"
                    for c, e in erros_por_turma[turma]:
                        msg += f"     • {c}: {e}\n"

            if not linhas_ignoradas_obs.empty:
                msg += f"\n⚠️ Linhas ignoradas por preenchimento na coluna OBSERVAÇÃO:\n"
                for i in linhas_ignoradas_obs.index:
                    cliente = linhas_ignoradas_obs.loc[i].get(coluna_cliente, f"Linha {i+2}")
                    msg += f"- {cliente} (linha {i+2})\n"

            if not linhas_ignoradas_valor.empty:
                msg += f"\n⚠️ Linhas ignoradas por VALOR em branco:\n"
                for i in linhas_ignoradas_valor.index:
                    cliente = linhas_ignoradas_valor.loc[i].get(coluna_cliente, f"Linha {i+2}")
                    msg += f"- {cliente} (linha {i+2})\n"

            self.finished.emit(msg)

        except Exception as e:
            self.erro.emit(str(e))





class Faturador(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Faturador Vortex 2025")
        self.setGeometry(500, 400, 520, 200)
        self.layout = QVBoxLayout()
        self.arquivo = None

        # 🔷 Título centralizado
        self.titulo = QLabel("🧾 Importar Pedidos de Venda - Vortex")
        self.titulo.setAlignment(Qt.AlignCenter)  # Centraliza o texto
        self.titulo.setStyleSheet("font-size: 16px; font-weight: bold;")
        self.layout.addWidget(self.titulo)

        self.botao_arquivo = QPushButton("Selecionar Planilha Excel")
        self.botao_arquivo.clicked.connect(self.selecionar_arquivo)
        self.layout.addWidget(self.botao_arquivo)

        self.combo_abas = QComboBox()
        self.layout.addWidget(self.combo_abas)

        self.botao_processar = QPushButton("Importar Pedidos")
        self.botao_processar.clicked.connect(self.processar)
        self.layout.addWidget(self.botao_processar)
        from PySide6.QtWidgets import QProgressBar
        
        self.barra_progresso = QProgressBar()
        self.barra_progresso.setValue(0)
        self.layout.addWidget(self.barra_progresso)

        self.setLayout(self.layout)

    def selecionar_arquivo(self):
        nome_arquivo, _ = QFileDialog.getOpenFileName(self, "Abrir Planilha", "", "Arquivos Excel (*.xlsx *.xls)")
        if nome_arquivo:
            self.arquivo = nome_arquivo
            xls = pd.ExcelFile(nome_arquivo)
            self.combo_abas.clear()
            self.combo_abas.addItems(xls.sheet_names)

    def processar(self):
        if not self.arquivo:
            QMessageBox.warning(self, "Erro", "Nenhuma planilha selecionada.")
            return

        aba = self.combo_abas.currentText()
        self.botao_processar.setEnabled(False)
        self.botao_processar.setText("Processando...")
        
        
        self.thread = ProcessarThread(self.arquivo, aba)
        self.thread.progresso.connect(self.atualizar_progresso)
        self.thread.finished.connect(self.exibir_resultado)
        self.thread.erro.connect(self.exibir_erro)
        self.thread.start()
    
    def atualizar_progresso(self, valor):
        self.barra_progresso.setValue(valor)


    def exibir_resultado(self, msg):
        dialogo = QDialog(self)
        dialogo.setWindowTitle("Resultado do Faturamento")
        dialogo.resize(600, 500)  # ⬅️ Tamanho maior para melhor leitura

        layout = QVBoxLayout()

        texto = QTextEdit()
        texto.setReadOnly(True)
        texto.setPlainText(msg)  # ⬅️ você pode usar setHtml(msg_formatado) se quiser usar cores, negrito, etc.
        layout.addWidget(texto)

        botao_ok = QPushButton("OK")
        botao_ok.clicked.connect(dialogo.accept)
        layout.addWidget(botao_ok)

        dialogo.setLayout(layout)
        dialogo.exec()

        self.botao_processar.setEnabled(True)
        self.botao_processar.setText("Faturar Pedidos")
        self.barra_progresso.setValue(0)

    def exibir_erro(self, msg):
        QMessageBox.critical(self, "Erro", msg)
        self.botao_processar.setEnabled(True)
        self.botao_processar.setText("Faturar Pedidos")
        self.barra_progresso.setValue(0)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Faturador()
    window.show()
    sys.exit(app.exec())
