import sys
import os
import json
import webview

# Permite imports absolutos da pasta APP
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from scripts.dolowd.login import iniciar_thread



# ✅ Executado diretamente
if __name__ == "__main__":  # Sinaliza que está rodando
    iniciar_thread()  # Executa sua automação
