import os
import time
import pythoncom
from oletools.olevba import VBA_Parser
from win32com.client import gencache, constants
from pywintypes import com_error

def extrair_vba_de_arquivos(pasta):
    modulos_extraidos = []
    arquivos = [f for f in os.listdir(pasta) if f.lower().endswith(('.xlsm', '.xlsb', '.xlam')) and not f.startswith('~$')]

    for arquivo in arquivos:
        caminho = os.path.join(pasta, arquivo)
        print(f"📦 Lendo: {arquivo}")
        try:
            parser = VBA_Parser(caminho)
            if parser.detect_vba_macros():
                for (_, _, vba_filename, vba_code) in parser.extract_macros():
                    nome_modulo = f"{os.path.splitext(arquivo)[0]}_{vba_filename}".replace(" ", "_").replace(":", "_")
                    modulos_extraidos.append((nome_modulo[:31], vba_code))
                    print(f"✅ Extraído: {nome_modulo}")
            parser.close()
        except Exception as e:
            print(f"❌ Erro ao extrair de {arquivo}: {e}")
    return modulos_extraidos

def safe_adicionar_modulo(vb_components, nome, codigo, max_retry=3):
    for attempt in range(max_retry):
        try:
            componente = vb_components.Add(1)  # 1 = vbext_ct_StdModule
            componente.Name = nome
            componente.CodeModule.AddFromString(codigo)
            return True
        except com_error as e:
            if attempt == max_retry - 1:
                print(f"❌ Falha ao adicionar {nome}: {e.excepinfo[2]}")
                return False
            time.sleep(0.5 * (2 ** attempt))
            pythoncom.PumpWaitingMessages()

def criar_arquivo_unico(pasta, modulos):
    print("\n🚀 Criando arquivo final...")

    xl = gencache.EnsureDispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 3  # msoAutomationSecurityForceDisable

    wb = xl.Workbooks.Add()
    destino = os.path.join(pasta, "Projeto_Unificado.xlsm")
    wb.SaveAs(Filename=destino, FileFormat=52)  # xlOpenXMLWorkbookMacroEnabled
    vb = wb.VBProject.VBComponents

    for i, (nome, codigo) in enumerate(modulos, 1):
        safe_adicionar_modulo(vb, nome, codigo)
        if i % 20 == 0:
            time.sleep(1)
            pythoncom.PumpWaitingMessages()

    wb.Save()
    wb.Close(SaveChanges=True)
    xl.Quit()

    print(f"\n✅ Arquivo final salvo como: {destino}")

if __name__ == "__main__":
    pasta = os.path.dirname(os.path.abspath(__file__))
    vbas = extrair_vba_de_arquivos(pasta)
    if vbas:
        criar_arquivo_unico(pasta, vbas)
    else:
        print("⚠️ Nenhum módulo encontrado para importar.")
