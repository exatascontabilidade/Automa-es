import pdfplumber
import argparse
import re
import os
from typing import List, Tuple, Optional, Dict

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
# FUNÇÕES DE LÓGICA DE CORES
# ==============================================================================
def definir_tipo_por_cor(cor: tuple) -> str:
    if not cor or len(cor) < 3: return "Padrão"
    r, g, b = cor[:3]
    if r > 0.5 and g < 0.4: return "Débito"  # Cor avermelhada
    if g > 0.5 and r < 0.4: return "Crédito" # Cor esverdeada
    return "Padrão"

def encontrar_cor_do_texto(pagina, texto_procurado: str) -> Optional[tuple]:
    resultados = pagina.search(texto_procurado, case=False)
    if not resultados: return None

    bbox = (resultados[0]['x0'], resultados[0]['top'], resultados[0]['x1'], resultados[0]['bottom'])
   
    chars_na_caixa = [c for c in pagina.chars if c['x0'] >= bbox[0] and c['x1'] <= bbox[2] and c['top'] >= bbox[1] and c['bottom'] <= bbox[3]]

    return chars_na_caixa[0].get('non_stroking_color') if chars_na_caixa else None

# ==============================================================================
# FUNÇÕES DE EXTRAÇÃO
# ==============================================================================

def extrair_dados_conta_efi(pdf: pdfplumber.PDF) -> Tuple[str, str]:

    agencia_encontrada = "NA"
    conta_encontrada = "NA"

    texto_primeira_pagina = ""
    if pdf.pages:
        texto_primeira_pagina = pdf.pages[0].extract_text()
    regex_ag_cc = re.compile(r"Agência\s+(\d+)\s*•\s*Conta\s+([\d-]+)", re.IGNORECASE)

    match = regex_ag_cc.search(texto_primeira_pagina)
    if match:
        agencia_encontrada = match.group(1).strip().replace('-', '')
        conta_encontrada = match.group(2).strip().replace('-', '')
    
    return agencia_encontrada, conta_encontrada

def processar_lancamentos_por_pagina(pagina_obj: pdfplumber.page.Page) -> List[List[str]]:
    PALAVRA_CHAVE_INICIO = "data" 
    PADRAO_DATA = re.compile(r"^\d{2}/\d{2}/\d{4}$")

    lancamentos_extraidos = []
    
    todas_as_palavras = pagina_obj.extract_words(x_tolerance=2, y_tolerance=2)
    
    pos_inicio = 0
    for i, palavra in enumerate(todas_as_palavras):
        if PALAVRA_CHAVE_INICIO in palavra['text'].lower():
            pos_inicio = i + 1
            break
            
    palavras_relevantes = todas_as_palavras[pos_inicio:]

    blocos_de_transacao = []
    bloco_atual = []
    for palavra in palavras_relevantes:
        if PADRAO_DATA.match(palavra['text']):
            if bloco_atual: blocos_de_transacao.append(bloco_atual)
            bloco_atual = [palavra]
        elif bloco_atual:
            bloco_atual.append(palavra)
    if bloco_atual: blocos_de_transacao.append(bloco_atual)

    for bloco in blocos_de_transacao:
        if not bloco or len(bloco) < 2: continue
        
        data_str = bloco[0]['text']
        palavras_do_bloco_sem_data = bloco[1:]
        
        valor_str = None
        valor_index = -1

        for i in range(len(palavras_do_bloco_sem_data) - 1, -1, -1):
            palavra_candidata = palavras_do_bloco_sem_data[i]['text']
            try:
                float(palavra_candidata.replace('.', '').replace(',', '.'))
                valor_str = palavra_candidata
                valor_index = i
                break
            except ValueError:
                continue

        if valor_str is None: continue

        descricao_palavras = [p['text'] for i, p in enumerate(palavras_do_bloco_sem_data) if i != valor_index]
        descricao = ' '.join(descricao_palavras)
        descricao = re.sub(r'\b\d{8,}\b', '', descricao) 
        descricao = ' '.join(descricao.split())

        cor_valor = encontrar_cor_do_texto(pagina_obj, valor_str)
        tipo_transacao = definir_tipo_por_cor(cor_valor)
        
        lancamentos_extraidos.append([data_str, descricao, tipo_transacao.upper(), valor_str])
            
    return lancamentos_extraidos

# ==============================================================================
# (DISPARADOR)
# ==============================================================================
def main():

    parser = argparse.ArgumentParser(description="Processa extratos bancários do EFI em formato PDF.")
    parser.add_argument("caminho_pdf", help="O caminho para o arquivo PDF do extrato.")
    args = parser.parse_args()

    todos_lancamentos = []
    try:
        with pdfplumber.open(args.caminho_pdf) as pdf:

            agencia, conta = extrair_dados_conta_efi(pdf)
            print(f"Agencia: {agencia}")
            print(f"Conta: {conta}")

        
            for pagina_obj in pdf.pages:
                lancamentos_da_pagina = processar_lancamentos_por_pagina(pagina_obj)
                todos_lancamentos.extend(lancamentos_da_pagina)
        
    
        if todos_lancamentos:
            print(f"\n> {len(todos_lancamentos)} lançamentos extraidos.")
            texto_formatado = formatar_para_txt_final(todos_lancamentos)
            nome_banco_fmt = "EFI"
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
            print("Nenhum lancamento valido foi extraido apos o processamento.")

    except Exception as e:
        print(f"Ocorreu um erro no processamento do EFI: {e}")

if __name__ == "__main__":
    main()
