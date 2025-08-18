# É necessário instalar as bibliotecas:
# pip install PyMuPDF
# pip install pdfplumber
import fitz  # PyMuPDF
import pdfplumber
import re
import os
import json
import argparse
from typing import List, Dict, Tuple

# ==============================================================================
# FUNÇÃO DE FORMATAÇÃO PARA TABELA (Sem alterações)
# ==============================================================================
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

# ==============================================================================
# FUNÇÃO PARA EXTRAIR CNPJ (USANDO PyMuPDF)
# ==============================================================================
def extrair_cnpj(pdf: fitz.Document) -> str:
    """
    Extrai o CNPJ do PDF, procurando em todas as páginas pelo padrão 'CNPJ:'.
    """
    cnpj_encontrado = "CNPJ_Nao_Definido"
    regex_cnpj = re.compile(r'CNPJ:\s*([\d./-]+)')

    for pagina in pdf.pages():
        texto_pagina = pagina.get_text("text")
        if not texto_pagina:
            continue
        
        for linha in texto_pagina.split('\n'):
            match = regex_cnpj.search(linha)
            if match:
                cnpj_limpo = re.sub(r'[./-]', '', match.group(1))
                return cnpj_limpo
    
    return cnpj_encontrado

# ==============================================================================
# FUNÇÃO PARA EXTRAIR BLOCO DE LANÇAMENTOS (USANDO pdfplumber)
# ==============================================================================
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
    Processa as linhas brutas com uma lógica híbrida:
    1. Prioriza o padrão de 3 linhas para âncoras "puras" (sem descrição).
    2. Trata âncoras "descritivas" (com texto) como lançamentos de 1 linha.
    """
    lancamentos_finais = []
    linhas_filtradas = [linha for linha in linhas_brutas if "SALDO ANTERIOR" not in linha]

    regex_ancora = re.compile(r'(-?[\d.,]+)\s+(-?[\d.,]+)$')
    regex_data = re.compile(r'(\d{2}/\d{2}/\d{4})')
    regex_data_curta = re.compile(r'\b\d{2}/\d{2}\b')

    indices_ancora = {i for i, linha in enumerate(linhas_filtradas) if regex_ancora.search(linha)}
    
    grupos_com_indices = []
    indices_processados = set()

    for i in sorted(list(indices_ancora)):
        if i in indices_processados:
            continue

        texto_antes_valor = regex_ancora.sub('', linhas_filtradas[i]).strip()
        is_pure_anchor = not re.search(r'[a-zA-Z]', texto_antes_valor)

        indice_anterior = i - 1
        indice_posterior = i + 1

        if (is_pure_anchor and
            indice_anterior >= 0 and indice_anterior not in indices_ancora and
            indice_posterior < len(linhas_filtradas) and indice_posterior not in indices_ancora):
            
            grupo = [linhas_filtradas[indice_anterior], linhas_filtradas[i], linhas_filtradas[indice_posterior]]
            grupos_com_indices.append({'indice': indice_anterior, 'grupo': grupo})
            
            indices_processados.add(indice_anterior)
            indices_processados.add(i)
            indices_processados.add(indice_posterior)

    for i in sorted(list(indices_ancora)):
        if i not in indices_processados:
            grupo = [linhas_filtradas[i]]
            grupos_com_indices.append({'indice': i, 'grupo': grupo})
            indices_processados.add(i)

    grupos_com_indices.sort(key=lambda x: x['indice'])
    grupos = [item['grupo'] for item in grupos_com_indices]

    data_recente = ""
    for grupo in grupos:
        linha_ancora = ""
        descricao_bruta = ""

        if len(grupo) == 3:
            linha_ancora = grupo[1]
            descricao_bruta = grupo[0] + ' ' + grupo[2]
        elif len(grupo) == 1:
            linha_ancora = grupo[0]
            descricao_bruta = grupo[0]
        else:
            continue

        data_transacao = None
        for linha in grupo:
            match_data = regex_data.search(linha)
            if match_data:
                data_transacao = match_data.group(1)
                break
        
        if not data_transacao:
            data_transacao = data_recente
        else:
            data_recente = data_transacao

        valor_match = regex_ancora.search(linha_ancora)
        if not valor_match: continue
        valor_transacao_str = valor_match.group(1)
        
        descricao_limpa = regex_ancora.sub('', descricao_bruta).strip()
        descricao_limpa = regex_data.sub('', descricao_limpa).strip()
        descricao_limpa = regex_data_curta.sub('', descricao_limpa).strip()
        
        palavras_a_remover = ['REM:']
        for palavra in palavras_a_remover:
            descricao_limpa = re.sub(r'\b' + re.escape(palavra) + r'\b', '', descricao_limpa, flags=re.IGNORECASE)
            
        descricao_final = re.sub(r'\s+', ' ', descricao_limpa).strip()
        
        tipo = "DÉBITO" if valor_transacao_str.startswith('-') else "CRÉDITO"
        valor_final = valor_transacao_str.replace("-", "").replace(".", "").replace(",", ".")

        lancamentos_finais.append([data_transacao, descricao_final, tipo, valor_final])
        
    return lancamentos_finais

def main():
    """Função principal do especialista Bradesco."""
    parser = argparse.ArgumentParser(description="Processador para extratos do banco BRADESCO.")
    parser.add_argument("caminho_pdf", help="O caminho para o arquivo PDF do BRADESCO.")
    args = parser.parse_args()
    
    carregar_config_bradesco()

    try:
        # Abre o PDF com PyMuPDF (fitz) para extrair o CNPJ
        with fitz.open(args.caminho_pdf) as pdf_fitz:
            cnpj = extrair_cnpj(pdf_fitz)
            print(f"CNPJ extraído: {cnpj}")
        
        # Abre o PDF com pdfplumber para extrair os lançamentos
        with pdfplumber.open(args.caminho_pdf) as pdf_plumber:
            bloco_bruto = extrair_bloco_de_lancamentos(pdf_plumber)
            
            if not bloco_bruto:
                print("Nenhuma tabela de lançamentos foi encontrada.")
                return

            todos_lancamentos = processar_bloco_de_lancamentos(bloco_bruto)
            
            if todos_lancamentos:
                print(f"\nProcessamento concluído. {len(todos_lancamentos)} lançamentos extraídos.")
                
                texto_formatado = formatar_para_txt_final(todos_lancamentos)

                print("\n--- VISUALIZAÇÃO DO RESULTADO EM TABELA ---\n")
                print(texto_formatado)
                
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
