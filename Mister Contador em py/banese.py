import pdfplumber
import argparse
import re
import os
from typing import List, Tuple

def formatar_para_txt_final(lista_lancamentos: list) -> str:
    """
    Formata os lançamentos em uma tabela de 4 colunas.
    """
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

def extrair_dados_conta(linhas_brutas: List[str]) -> Tuple[str, str]:
    """
    Extrai a Agência e a Conta do PDF, procurando pelos respectivos campos.
    Também imprime as 10 primeiras linhas para depuração.
    """
    # --- DEBUG: Imprime as 10 primeiras linhas do PDF ---
    print("\n--- [DEBUG] 10 primeiras linhas do PDF ---")
    for i, linha in enumerate(linhas_brutas[:10]):
        print(f"Linha {i}: '{linha}'")
    print("------------------------------------------\n")
    # --- FIM DO DEBUG ---

    agencia_encontrada = "Agencia_Nao_Definida"
    conta_encontrada = "Conta_Nao_Definida"

    # Regex para encontrar 'Agência' e 'Conta' (case-insensitive)
    regex_agencia = re.compile(r'Agência\s*:?\s*([\d-]+)', re.IGNORECASE)
    regex_conta = re.compile(r'Conta\s*:?\s*([\d-]+)', re.IGNORECASE)

    for linha in linhas_brutas:
        match_agencia = regex_agencia.search(linha)
        if match_agencia:
            agencia_encontrada = match_agencia.group(1).strip()

        match_conta = regex_conta.search(linha)
        if match_conta:
            conta_encontrada = match_conta.group(1).strip()
    
    return agencia_encontrada, conta_encontrada

def extrair_bloco_de_lancamentos(linhas_brutas: List[str]) -> List[str]:
    """
    Localiza e extrai o bloco de texto contendo os lançamentos.
    """
    bloco_lancamentos = []
    capturando = False
    for linha in linhas_brutas:
        # O cabeçalho da tabela de lançamentos marca o início da captura
        if all(keyword in linha for keyword in ["Data", "Local", "Histórico", "Valor", "Saldo"]):
            capturando = True
            continue # Pula a linha do cabeçalho

        # Adiciona a linha se estivermos no modo de captura e a linha não for vazia
        if capturando and linha.strip():
            bloco_lancamentos.append(linha)
    
    return bloco_lancamentos

def processar_bloco_de_lancamentos(linhas_lancamentos: List[str]) -> List[List[str]]:
    """
    Processa cada linha do bloco de lançamentos para extrair os dados.
    """
    lancamentos_finais = []
    # Regex para capturar: Data, Descrição, Valor da transação e o sinal (+ ou -)
    # Exemplo: 02/06/2025 CRED BANESE CARD 00000001 1.225,79 + 5.419,71 +
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
            
    return lancamentos_finais

def extrair_linhas_brutas(caminho_pdf: str) -> List[str]:
    """
    Abre um arquivo PDF e extrai todas as linhas de texto de todas as páginas.
    """
    todas_as_linhas = []
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for i, pagina in enumerate(pdf.pages):
                texto_pagina = pagina.extract_text(x_tolerance=2, y_tolerance=2)
                if texto_pagina:
                    linhas_pagina = texto_pagina.split('\n')
                    todas_as_linhas.extend(linhas_pagina)
    except Exception as e:
        print(f"Erro ao ler o arquivo PDF: {e}")
    return todas_as_linhas

def main():
    """
    Função principal para extrair, processar e formatar um extrato do Banese.
    """
    parser = argparse.ArgumentParser(description="Processa extratos bancários do Banese em formato PDF.")
    parser.add_argument("caminho_pdf", help="O caminho para o arquivo PDF do extrato.")
    args = parser.parse_args()

    # 1. Extrai todas as linhas do PDF
    linhas_brutas = extrair_linhas_brutas(args.caminho_pdf)
    if not linhas_brutas:
        print("Nenhum texto foi extraído do PDF. Verifique o arquivo.")
        return

    # 2. Extrai a agência e a conta do cabeçalho do extrato
    agencia, conta = extrair_dados_conta(linhas_brutas)
    print(f"Agência extraída: {agencia}")
    print(f"Conta extraída: {conta}")

    # 3. Isola apenas as linhas que contêm os lançamentos
    bloco_lancamentos = extrair_bloco_de_lancamentos(linhas_brutas)
    if not bloco_lancamentos:
        print("Nenhum bloco de lançamentos foi encontrado no extrato.")
        return

    # 4. Processa cada linha para extrair os dados
    lancamentos_processados = processar_bloco_de_lancamentos(bloco_lancamentos)
    
    # 5. Formata os dados em uma tabela para exibição e salvamento
    if lancamentos_processados:
        print(f"\nProcessamento concluído. {len(lancamentos_processados)} lançamentos extraídos.")
        texto_formatado = formatar_para_txt_final(lancamentos_processados)
        print("\n--- VISUALIZAÇÃO DO RESULTADO EM TABELA ---\n")
        print(texto_formatado)

        # --- LÓGICA DE SALVAMENTO DO ARQUIVO ---
        nome_banco_fmt = "BANESE"
        
        # Define o diretório onde o script está sendo executado
        script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()
        
        # Define o caminho para a pasta 'Processamento' e a subpasta do banco
        pasta_processamento = os.path.join(script_dir, "Processamento")
        pasta_banco = os.path.join(pasta_processamento, nome_banco_fmt.upper())
        
        # Cria as pastas se elas não existirem
        os.makedirs(pasta_banco, exist_ok=True)
        
        # Define o nome do arquivo usando o número da conta
        nome_arquivo = f"Extrato_{nome_banco_fmt}_{conta}.txt"
        caminho_arquivo_saida = os.path.join(pasta_banco, nome_arquivo)
        
        # Salva o arquivo com os novos dados no cabeçalho
        with open(caminho_arquivo_saida, 'w', encoding='utf-8') as f:
            f.write(f"Banco: {nome_banco_fmt.upper()}\n")
            f.write(f"Agência: {agencia}\n")
            f.write(f"Conta: {conta}\n\n")
            f.write(texto_formatado)
        
        print(f"\nResultado salvo em: '{caminho_arquivo_saida}'")
    else:
        print("Nenhum lançamento válido foi extraído após o processamento.")

if __name__ == "__main__":
    main()
