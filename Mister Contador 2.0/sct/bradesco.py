import pdfplumber
import re
import os
import argparse
from typing import List, Dict, Tuple

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

# ==============================================================================
# FUNÇÕES DE EXTRAÇÃO
# ==============================================================================

def extrair_agencia_conta(pdf: pdfplumber.PDF) -> Tuple[str, str]:

    regex_ag_cc = re.compile(r"Extrato de:\s*Ag:\s*(\d+)\s*\|\s*CC:\s*([\d-]+)")

    for pagina in pdf.pages:
        texto_pagina = pagina.extract_text(x_tolerance=2, y_tolerance=2)
        if not texto_pagina:
            continue
        
        for linha in texto_pagina.split('\n'):
            match = regex_ag_cc.search(linha)
            if match:
                agencia = match.group(1)
                conta_bruta = match.group(2)
                conta_limpa = conta_bruta.replace('-', '')
                
                return agencia, conta_limpa
    
    return "NA", "NA"

def extrair_bloco_de_lancamentos(pdf: pdfplumber.PDF) -> List[str]:
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
    lancamentos_finais = []
    linhas_filtradas = [linha for linha in linhas_brutas if "SALDO ANTERIOR" not in linha]
    regex_ancora = re.compile(r'(-?[\d.,]+)\s+(-?[\d.,]+)$')
    regex_data = re.compile(r'(\d{2}/\d{2}/\d{4})')
    regex_data_curta = re.compile(r'\b\d{2}/\d{2}\b')
    indices_ancora = {i for i, linha in enumerate(linhas_filtradas) if regex_ancora.search(linha)}
    grupos_com_indices = []
    indices_processados = set()
    for i in sorted(list(indices_ancora)):
        if i in indices_processados: continue
        texto_antes_valor = regex_ancora.sub('', linhas_filtradas[i]).strip()
        is_pure_anchor = not re.search(r'[a-zA-Z]', texto_antes_valor)
        indice_anterior = i - 1
        indice_posterior = i + 1
        if (is_pure_anchor and indice_anterior >= 0 and indice_anterior not in indices_ancora and indice_posterior < len(linhas_filtradas) and indice_posterior not in indices_ancora):
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
        else: continue
        data_transacao = None
        for linha in grupo:
            match_data = regex_data.search(linha)
            if match_data:
                data_transacao = match_data.group(1)
                break
        if not data_transacao: data_transacao = data_recente
        else: data_recente = data_transacao
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

# ==============================================================================
# (DISPARADOR)
# ==============================================================================
def main():
    parser = argparse.ArgumentParser(description="Processador para extratos do banco BRADESCO.")
    parser.add_argument("caminho_pdf", help="O caminho para o arquivo PDF do BRADESCO.")
    args = parser.parse_args()
    
    try:
        with pdfplumber.open(args.caminho_pdf) as pdf_plumber:
            
            agencia, conta = extrair_agencia_conta(pdf_plumber)
            print(f"Agencia: {agencia}")
            print(f"Conta: {conta}")
            
            bloco_bruto = extrair_bloco_de_lancamentos(pdf_plumber)
            
            if not bloco_bruto:
                print("Nenhuma tabela de lancamentos foi encontrada.")
                return

            todos_lancamentos = processar_bloco_de_lancamentos(bloco_bruto)
            
            if todos_lancamentos:
                print(f"\n> {len(todos_lancamentos)} lancamentos extraidos.")
                
                texto_formatado = formatar_para_txt_final(todos_lancamentos)

                nome_banco_fmt = "BRADESCO"
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
                print("Nenhum lançamento valido foi extraído apos o processamento.")
            
    except Exception as e:
        print(f"Ocorreu um erro no processamento: {e}")

if __name__ == "__main__":
    main()