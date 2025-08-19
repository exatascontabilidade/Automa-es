import pdfplumber
import re
import os
import argparse
from typing import List, Optional, Tuple

def formatar_para_txt_final(lista_lancamentos: list) -> str:
    """Formata os lançamentos em uma tabela de 4 colunas."""
    if not lista_lancamentos: return "Nenhum lançamento para formatar."
    
    tabela_para_formatar = [['Data', 'Descrição', 'Tipo', 'Valor']] + lista_lancamentos
    larguras_colunas = [0, 0, 0, 0]
    max_largura_desc = 120
    
    for linha in tabela_para_formatar:
        try:
            larguras_colunas[0] = max(larguras_colunas[0], len(str(linha[0])))
            larguras_colunas[1] = max(larguras_colunas[1], len(str(linha[1])[:max_largura_desc]))
            larguras_colunas[2] = max(larguras_colunas[2], len(str(linha[2])))
            larguras_colunas[3] = max(larguras_colunas[3], len(str(linha[3])))
        except (IndexError, TypeError):
            continue

    string_final = ""
    cabecalho = tabela_para_formatar[0]
    string_final += f"{cabecalho[0].center(larguras_colunas[0])} | {cabecalho[1].center(larguras_colunas[1])} | {cabecalho[2].center(larguras_colunas[2])} | {cabecalho[3].center(larguras_colunas[3])}\n"
    string_final += f"{'-' * larguras_colunas[0]}-+-{'-' * larguras_colunas[1]}-+-{'-' * larguras_colunas[2]}-+-{'-' * larguras_colunas[3]}\n"
    
    for linha in tabela_para_formatar[1:]:
        data, desc, tipo, valor = linha
        desc_formatada = str(desc)[:max_largura_desc]
        string_final += f"{str(data).ljust(larguras_colunas[0])} | {desc_formatada.ljust(larguras_colunas[1])} | {str(tipo).center(larguras_colunas[2])} | {str(valor).rjust(larguras_colunas[3])}\n"
        
    return string_final

def extrair_linhas_brutas(caminho_pdf: str) -> list[str]:
    todas_as_linhas = []
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for i, pagina in enumerate(pdf.pages):
                texto_pagina = pagina.extract_text(x_tolerance=2, y_tolerance=2)
                if texto_pagina:
                    todas_as_linhas.extend(texto_pagina.split('\n'))
    except Exception as e:
        print(f"ERRO ao ler o PDF: {e}")
    return todas_as_linhas

def extrair_agencia_conta(linhas_bloco: List[str]) -> Tuple[str, str]:
    regex_ag_cc = re.compile(r"AGENCIA:\s+(\d+)\s+CONTA\s+([\d.-]+)")
    for linha in linhas_bloco:
        match = regex_ag_cc.search(linha)
        if match:
            agencia = match.group(1)
            conta = match.group(2)
            return agencia, conta
    return "N/A", "N/A"

def extrair_mes_ano(linhas_brutas: List[str]) -> Optional[str]:
    regex_mes_ano = re.compile(r"Mês:\s*(\w+)/(\d{4})")
    for linha in linhas_brutas:
        match = regex_mes_ano.search(linha)
        if match:
            mes_nome = match.group(1).lower()
            ano = match.group(2)
            meses = {
                'janeiro': '01', 'fevereiro': '02', 'março': '03', 'abril': '04', 
                'maio': '05', 'junho': '06', 'julho': '07', 'agosto': '08', 
                'setembro': '09', 'outubro': '10', 'novembro': '11', 'dezembro': '12'
            }
            mes_num = meses.get(mes_nome, "MM")
            return f"{mes_num}/{ano}"
    return None

def extrair_bloco_de_lancamentos_bnb(linhas_do_bloco_conta: List[str]) -> List[str]:
    MARCADOR_INICIO = "> DEMONSTRATIVO DA MOVIMENTACAO DE CONTA CORRENTE"
    MARCADOR_FIM = "> "

    linhas_da_tabela = []
    capturando = False
    for linha in linhas_do_bloco_conta:
        if MARCADOR_FIM in linha:
            capturando = False
            break

        if MARCADOR_INICIO in linha:
            capturando = True
            continue

        if capturando:
            if linha.strip().startswith('*') or "DIA HISTORICO" in linha or "___" in linha:
                continue
            if linha.strip():
                if not re.search(r'https://nel\.bnb\.gov\.br', linha):
                     linhas_da_tabela.append(linha)
            
    return linhas_da_tabela

def processar_bloco_bnb(linhas_do_bloco: List[str], mes_ano: str) -> List[List[str]]:
    lancamentos_finais = []
    regex_lancamento = re.compile(r'^(\d{1,2})?\s*(.+?)\s+([\d.,]+)([+-])\s+[\d.,]+$')
    
    dia_atual = ""
    for linha in linhas_do_bloco:
        match = regex_lancamento.search(linha)
        if match:
            dia_grupo, descricao, valor_str, sinal = match.groups()
            
            if dia_grupo:
                dia_atual = dia_grupo.zfill(2)

            if not dia_atual:
                continue

            data_completa = f"{dia_atual}/{mes_ano}"
            tipo = "CRÉDITO" if sinal == '+' else "DÉBITO"
            valor_final = valor_str.replace(".", "").replace(",", ".")
            descricao = re.sub(r'\s+\d+$', '', descricao.strip())
            
            lancamentos_finais.append([data_completa, descricao, tipo, valor_final])
            
    return lancamentos_finais

# ==============================================================================
# DISPARADOR)
# ==============================================================================
def main():
    parser = argparse.ArgumentParser(description="Processador para extratos do banco BNB com multiplas contas.")
    parser.add_argument("caminho_pdf", help="O caminho para o arquivo PDF do extrato consolidado.")
    args = parser.parse_args()
    
    linhas_brutas = extrair_linhas_brutas(args.caminho_pdf)
    if not linhas_brutas:
        print("Nenhum texto foi extraído do PDF.")
        return

    regex_ag_cc = re.compile(r"AGENCIA:\s+(\d+)\s+CONTA\s+([\d.-]+)") # <<< CORRIGIDO >>>
    indices_inicio_contas = [i for i, linha in enumerate(linhas_brutas) if regex_ag_cc.search(linha)]

    if not indices_inicio_contas:
        print("Nenhuma secao de Agencia/Conta foi encontrada no documento.")
        return

    print(f"Mapeamento concluido. Encontrado(s) {len(indices_inicio_contas)} extrato(s) de conta no arquivo.\n")
    
    mes_ano_global = extrair_mes_ano(linhas_brutas)
    if not mes_ano_global:
        print("ERRO: Nao foi possivel encontrar o Mes/Ano de referencia no extrato.")
        return
    
    for i, start_index in enumerate(indices_inicio_contas):
        end_index = indices_inicio_contas[i + 1] if i + 1 < len(indices_inicio_contas) else len(linhas_brutas)
        bloco_conta_atual = linhas_brutas[start_index:end_index]
        
        agencia, conta = extrair_agencia_conta(bloco_conta_atual)

        print(f"Agencia: {agencia}")
        print(f"Conta: {conta}")

        bloco_lancamentos = extrair_bloco_de_lancamentos_bnb(bloco_conta_atual)
        if not bloco_lancamentos:
            print("Nenhum bloco de lancamentos encontrado para esta conta.\n")
            continue
            
        todos_lancamentos = processar_bloco_bnb(bloco_lancamentos, mes_ano_global)
        
        if todos_lancamentos:
            print(f"\n> {len(todos_lancamentos)} lancamentos extraidos.")
            
            texto_formatado = formatar_para_txt_final(todos_lancamentos)
            
            nome_banco_fmt = "BNB"
            script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()
            pasta_processamento = os.path.join(script_dir, "Processamento")
            pasta_banco = os.path.join(pasta_processamento, nome_banco_fmt.upper())
            os.makedirs(pasta_banco, exist_ok=True)
            
            base_name = os.path.basename(args.caminho_pdf).replace('.pdf', '')
            
            nome_arquivo = f"{nome_banco_fmt}_AG{agencia}_CC{conta}.txt"
            caminho_arquivo_saida = os.path.join(pasta_banco, nome_arquivo)
            
            cabecalho_arquivo = (
                f"Banco: {nome_banco_fmt.upper()}\n"
                f"Agência: {agencia}\n"
                f"Conta: {conta}\n\n"
            )
            
            with open(caminho_arquivo_saida, 'w', encoding='utf-8') as f:
                f.write(cabecalho_arquivo + texto_formatado)   
        else:
            print("Nenhum lancamento valido foi extraido para esta conta.\n")

if __name__ == "__main__":
    main()