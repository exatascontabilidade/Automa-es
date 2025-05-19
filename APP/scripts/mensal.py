import sys
import os

# Adiciona a pasta APP no sys.path para permitir importação absoluta
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from scripts.dolowd.login import iniciar_thread

if __name__ == "__main__":
    iniciar_thread()