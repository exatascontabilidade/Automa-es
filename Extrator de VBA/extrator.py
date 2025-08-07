import os
import sys
from oletools.olevba import VBA_Parser

def extracao_forcada_direta(caminho_arquivo_excel, diretorio_saida):
    """
    Tenta uma extração forçada lendo o arquivo diretamente com o método
    extract_macros() do olevba, que não depende da automação do Excel.
    Este método é ideal quando o acesso ao projeto VBA é bloqueado pelo Excel.
    """
    if not os.path.exists(caminho_arquivo_excel):
        print(f"\nErro: O arquivo '{caminho_arquivo_excel}' não foi encontrado.")
        return

    print(f"\n[*] Analisando o arquivo: {os.path.basename(caminho_arquivo_excel)}")

    vba_parser = None
    try:
        vba_parser = VBA_Parser(caminho_arquivo_excel)

        if vba_parser.detect_vba_macros():
            print("[+] Macros VBA detectadas.")
            print("[*] Tentando extração forçada e direta do código-fonte (sem usar o Excel)...")
            
            nome_base_excel = os.path.splitext(os.path.basename(caminho_arquivo_excel))[0]
            arquivos_salvos = 0
            
            # Usando extract_macros() para uma abordagem mais direta
            resultados_extracao = vba_parser.extract_macros()

            for (nome_arquivo_origem, stream_path, nome_modulo_vba, codigo) in resultados_extracao:
                
                # O código pode vir em bytes, então decodificamos com tratamento de erros.
                try:
                    codigo_decodificado = codigo.decode('utf-8', errors='replace')
                except AttributeError:
                    # Caso o código já seja uma string
                    codigo_decodificado = codigo

                if codigo_decodificado and codigo_decodificado.strip():
                    print(f"  -> Código encontrado no módulo: '{nome_modulo_vba}'")
                    nome_arquivo_saida = f"{nome_modulo_vba}.vba"
                    caminho_completo_saida = os.path.join(diretorio_saida, nome_arquivo_saida)
                    
                    try:
                        with open(caminho_completo_saida, 'w', encoding='utf-8') as f:
                            f.write(codigo_decodificado)
                        print(f"    -> SUCESSO! Código salvo em: {caminho_completo_saida}")
                        arquivos_salvos += 1
                    except IOError as e:
                        print(f"    -> ERRO ao salvar o arquivo {nome_arquivo_saida}: {e}")
                else:
                    print(f"  -> Módulo '{nome_modulo_vba}' está vazio ou não contém código-fonte extraível.")

            if arquivos_salvos > 0:
                print(f"\n[+] Processo concluído! {arquivos_salvos} arquivo(s) de código-fonte salvo(s).")
            else:
                print("\n[-] A extração forçada não encontrou nenhum código-fonte visível. Isso é comum em arquivos com 'VBA Stomping', onde o código-fonte foi apagado.")

        else:
            print("[-] Nenhuma macro VBA foi encontrada neste arquivo.")

    except Exception as e:
        print(f"\n[!] Ocorreu um erro inesperado durante a análise direta do arquivo: {e}")

    finally:
        if vba_parser:
            vba_parser.close()

# --- Bloco de Execução ---
if __name__ == "__main__":
    try:
        diretorio_do_script = os.path.dirname(os.path.abspath(sys.argv[0]))
    except NameError:
        diretorio_do_script = os.getcwd()

    caminho_do_arquivo = input("Insira o caminho para o arquivo Excel (.xlsb) para extração forçada: ")
    caminho_do_arquivo = caminho_do_arquivo.strip().strip('"')

    extracao_forcada_direta(caminho_do_arquivo, diretorio_do_script)