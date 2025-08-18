import pdfplumber
import re
import os
import json
import argparse
from typing import List, Dict, Tuple

# ==============================================================================
# FUNÇÃO DE FORMATAÇÃO PARA TABELA (Reintroduzida)
# ==============================================================================
def formatar_para_txt_final(lista_lancamentos: list) -> str:
    """Formata os lançamentos em uma tabela de 4 colunas."""
    if not lista_lancamentos: return "Nenhum lançamento para formatar."
    
    tabela_para_formatar = [['Data', 'Descrição', 'Tipo', 'Valor']] + lista_lancamentos
    larguras_colunas = [0, 0, 0, 0]
    
    # Define a largura máxima da descrição para não quebrar a tabela
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
        desc_formatada = str(desc)[:max_largura_desc] # Garante que a descrição seja string
        string_final += f"{str(data).ljust(larguras_colunas[0])} | {desc_formatada.ljust(larguras_colunas[1])} | {str(tipo).center(larguras_colunas[2])} | {str(valor).rjust(larguras_colunas[3])}\n"
        
    return string_final


def carregar_config_bradesco() -> Dict:
    """Carrega a configuração específica para o banco Bradesco."""
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        script_dir = os.getcwd()
    
    caminho_config = os.path.join(script_dir, 'config.json')
    try:
        with open(caminho_config, 'r', encoding='utf-8') as f:
            config_total = json.load(f)
            for banco_info in config_total.get('bancos', []):
                if banco_info.get('nome_banco') == 'bradesco':
                    print("Configuração do modelo de banco 'BRADESCO' carregada.")
                    return banco_info.get('config_extracao', {})
    except Exception as e:
        print(f"Erro ao carregar configuração do BRADESCO: {e}")
    return {}

def extrair_bloco_de_lancamentos(pdf: pdfplumber.PDF) -> List[str]:
    """
    Extrai todas as linhas de texto entre o cabeçalho 'Data Lançamento...' 
    e a linha 'Total...'.
    """
    linhas_da_tabela = []
    capturando = False
    for pagina in pdf.pages:
        texto_pagina = pagina.extract_text(x_tolerance=2, y_tolerance=2)
        if not texto_pagina: continue
        linhas = texto_pagina.split('\n')
        for linha in linhas:
            if capturando and linha.strip().startswith("Total"):
                capturando = False
                break
            if "Data" in linha and "Lançamento" in linha and "Saldo (R$)" in linha:
                capturando = True
                continue
            if capturando and linha.strip():
                linhas_da_tabela.append(linha)
        if not capturando and len(linhas_da_tabela) > 0:
            break
    return linhas_da_tabela

def processar_bloco_de_lancamentos(linhas_brutas: List[str]) -> List[List[str]]:
    """
    Processa as linhas brutas com a lógica refinada e junta a descrição
    em uma única linha.
    """
    lancamentos_finais = []
    linhas_filtradas = [linha for linha in linhas_brutas if "SALDO ANTERIOR" not in linha]

    regex_ancora = re.compile(r'(-?[\d.,]+)\s+(-?[\d.,]+)$')
    regex_data = re.compile(r'(\d{2}/\d{2}/\d{4})')
    
    indices_ancora = [i for i, linha in enumerate(linhas_filtradas) if regex_ancora.search(linha)]
    
    grupos = []
    inicio_grupo = 0
    for i, indice_atual in enumerate(indices_ancora):
        fim_grupo = indice_atual
        if (indice_atual + 1) < len(linhas_filtradas):
            proximo_indice = indice_atual + 1
            if proximo_indice not in indices_ancora:
                fim_grupo = proximo_indice
        grupos.append(linhas_filtradas[inicio_grupo : fim_grupo + 1])
        inicio_grupo = fim_grupo + 1

    data_recente = ""
    for grupo in grupos:
        linha_ancora = ""
        descricao_parts = []
        for linha in grupo:
            if regex_ancora.search(linha):
                linha_ancora = linha
            else:
                descricao_parts.append(linha)

        if not linha_ancora: continue

        valor_transacao_str = regex_ancora.search(linha_ancora).group(1)
        match_data = regex_data.search(linha_ancora)
        if match_data:
            data_recente = match_data.group(1)

        texto_ancora = regex_ancora.sub('', linha_ancora).strip()
        texto_ancora_limpo = regex_data.sub('', texto_ancora).strip()
        descricao_parts.append(texto_ancora_limpo)

        # --- MUDANÇA PRINCIPAL AQUI ---
        # Juntamos as partes da descrição com ' ' (espaço) para formar uma única linha
        descricao_final = ' '.join(p.strip() for p in descricao_parts if p.strip())
        
        tipo = "DÉBITO" if valor_transacao_str.startswith('-') else "CRÉDITO"
        valor_final = valor_transacao_str.replace("-", "").replace(".", "").replace(",", ".")

        lancamentos_finais.append([data_recente, descricao_final, tipo, valor_final])
        
    return lancamentos_finais

def main():
    """Função principal do especialista Bradesco."""
    parser = argparse.ArgumentParser(description="Processador para extratos do banco BRADESCO.")
    parser.add_argument("caminho_pdf", help="O caminho para o arquivo PDF do BRADESCO.")
    args = parser.parse_args()
    
    carregar_config_bradesco()

    try:
        with pdfplumber.open(args.caminho_pdf) as pdf:
            bloco_bruto = extrair_bloco_de_lancamentos(pdf)
            
            if not bloco_bruto:
                print("Nenhuma tabela de lançamentos foi encontrada.")
                return

            todos_lancamentos = processar_bloco_de_lancamentos(bloco_bruto)
            
            if todos_lancamentos:
                print(f"Processamento concluído. {len(todos_lancamentos)} lançamentos extraídos.")
                
                # --- VOLTAMOS A USAR O FORMATADOR DE TABELA ---
                texto_formatado = formatar_para_txt_final(todos_lancamentos)

                print("\n--- VISUALIZAÇÃO DO RESULTADO EM TABELA ---\n")
                print(texto_formatado)
                
                # Salva o resultado no arquivo
                cnpj = "CNPJ_Nao_Definido"
                nome_banco_fmt = "BRADESCO"
                script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()
                pasta_processamento = os.path.join(script_dir, "Processamento")
                pasta_banco = os.path.join(pasta_processamento, nome_banco_fmt.upper())
                os.makedirs(pasta_banco, exist_ok=True)
                nome_arquivo = f"Extrato_{nome_banco_fmt}_{cnpj}.txt"
                caminho_arquivo_saida = os.path.join(pasta_banco, nome_arquivo)
                
                with open(caminho_arquivo_saida, 'w', encoding='utf-8') as f:
                    f.write(f"Banco: {nome_banco_fmt.upper()}\n\n{texto_formatado}")
                
                print(f"\nResultado salvo em: '{caminho_arquivo_saida}'")
            else:
                print("Nenhum lançamento válido foi extraído após o processamento.")
            
    except Exception as e:
        print(f"Ocorreu um erro no processamento: {e}")

if __name__ == "__main__":
    main()