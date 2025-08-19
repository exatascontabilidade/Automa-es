import pdfplumber
import argparse
import re
import os
from typing import List, Tuple

def formatar_para_txt_final(lista_lancamentos: list) -> str:
    if not lista_lancamentos:
        return "Nenhum lançamento para formatar."
    
    tabela_para_formatar = [['Data', 'Descrição', 'Tipo', 'Valor']] + lista_lancamentos
    larguras_colunas = [0, 0, 0, 0]
    
    max_largura_desc = 80
    
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

# ==============================================================================
# FUNÇÕES DE EXTRAÇÃO PARA O BANESE
# ==============================================================================

def extrair_linhas_brutas(caminho_pdf: str) -> List[str]:
    todas_as_linhas = []
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for pagina in pdf.pages:
                texto_pagina = pagina.extract_text(x_tolerance=2, y_tolerance=2)
                if texto_pagina:
                    todas_as_linhas.extend(texto_pagina.split('\n'))
    except Exception as e:
        print(f"Erro ao ler o arquivo PDF: {e}")
    return todas_as_linhas

def extrair_dados_conta(linhas_brutas: List[str]) -> Tuple[str, str]:

    regex_ag_cc = re.compile(r'Agência\s+([\d-]+)\s+Tipo\s+\d+\s+Conta\s+([\d-]+)', re.IGNORECASE)

    for linha in linhas_brutas:
        match = regex_ag_cc.search(linha)
        if match:
            agencia_encontrada = match.group(1).strip().replace('-', '')
            conta_encontrada = match.group(2).strip().replace('-', '')
            return agencia_encontrada, conta_encontrada
    

    return "Agencia_Nao_Definida", "Conta_Nao_Definida"

def extrair_bloco_de_lancamentos(linhas_brutas: List[str]) -> List[str]:

    bloco_lancamentos = []
    capturando = False
    

    MARCADOR_FIM = "Alô Banese" 
    
    PALAVRAS_CHAVE_INICIO = ["Data", "Histórico", "Valor", "Saldo"]

    for linha in linhas_brutas:
    
        if capturando and MARCADOR_FIM and MARCADOR_FIM in linha:
            break

        if not capturando and sum(keyword in linha for keyword in PALAVRAS_CHAVE_INICIO) >= 3:
            capturando = True
            continue # Pula a linha do cabeçalho


        if capturando and linha.strip():
            bloco_lancamentos.append(linha)
    
    return bloco_lancamentos

def processar_bloco_de_lancamentos(linhas_lancamentos: List[str]) -> List[List[str]]:

    lancamentos_finais = []
    regex_lancamento = re.compile(
        r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d.,]+)\s*([+-])\s+[\d.,]+\s*[+-]$"
    )

    for linha in linhas_lancamentos:
        match = regex_lancamento.search(linha)
        if match:
            # Extrai os dados usando os grupos do regex
            data = match.group(1)
            descricao = match.group(2).strip()
            valor_str = match.group(3)
            sinal = match.group(4)

            # Determina o tipo e formata o valor
            tipo = "CRÉDITO" if sinal == '+' else "DÉBITO"
            valor_final = valor_str.replace(".", "").replace(",", ".")

            lancamentos_finais.append([data, descricao, tipo, valor_final])
        
        elif lancamentos_finais:
            # Se a linha não corresponde a um novo lançamento,
            # considera-se que é uma continuação da descrição anterior.
            lancamentos_finais[-1][1] += " " + linha.strip()
            
    return lancamentos_finais

# ==============================================================================
# (DISPARADOR)
# ==============================================================================
def main():
    parser = argparse.ArgumentParser(description="Processa extratos bancários do Banese em formato PDF.")
    parser.add_argument("caminho_pdf", help="O caminho para o arquivo PDF do extrato.")
    args = parser.parse_args()

    # 1. Extrai todas as linhas do PDF
    linhas_brutas = extrair_linhas_brutas(args.caminho_pdf)
    if not linhas_brutas:
        print("Nenhum texto foi extraido do PDF. Verifique o arquivo.")
        return
    
    agencia, conta = extrair_dados_conta(linhas_brutas)
    print(f"Agencia: {agencia}")
    print(f"Conta: {conta}")

    bloco_lancamentos = extrair_bloco_de_lancamentos(linhas_brutas)
    if not bloco_lancamentos:
        print("Nenhum bloco de lancamentos foi encontrado no extrato.")
        return


    lancamentos_processados = processar_bloco_de_lancamentos(bloco_lancamentos)
    
    if lancamentos_processados:
        print(f"\n> {len(lancamentos_processados)} lancamentos extraidos.")
        texto_formatado = formatar_para_txt_final(lancamentos_processados)
        nome_banco_fmt = "BANESE"
        script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()
        pasta_processamento = os.path.join(script_dir, "Processamento")
        pasta_banco = os.path.join(pasta_processamento, nome_banco_fmt.upper())
        os.makedirs(pasta_banco, exist_ok=True)
        
        base_name = os.path.basename(args.caminho_pdf).replace('.pdf', '')
        
        nome_arquivo = f"{nome_banco_fmt}_AG{agencia}_CC{conta}'.txt"
        caminho_arquivo_saida = os.path.join(pasta_banco, nome_arquivo)
        
        cabecalho_arquivo = (
            f"Banco: {nome_banco_fmt.upper()}\n"
            f"Agência: {agencia}\n"
            f"Conta: {conta}\n\n"
        )
        
        with open(caminho_arquivo_saida, 'w', encoding='utf-8') as f:
            f.write(cabecalho_arquivo + texto_formatado)
    else:
        print("Nenhum lancamento valido foi extraido apos o processamento.")

if __name__ == "__main__":
    main()
