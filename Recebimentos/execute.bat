@echo off
setlocal
cd /d "%~dp0"

echo 🔍 Verificando Python...
where python >nul 2>nul
if errorlevel 1 (
    echo ❗ Python não encontrado. Iniciando instalação...
    curl -o python_installer.exe https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe
    python_installer.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
    del python_installer.exe
    echo 🔁 Atualizando variáveis de ambiente...
    setx PATH "%PATH%;C:\Program Files\Python312;C:\Program Files\Python312\Scripts"
    set PATH=%PATH%;C:\Program Files\Python312;C:\Program Files\Python312\Scripts
)

echo 🔍 Verificando pip...
python -m pip --version >nul 2>nul
if errorlevel 1 (
    echo 📥 Instalando PIP...
    curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py
    python get-pip.py
    del get-pip.py
)

echo 📦 Instalando bibliotecas necessárias...
python -m pip install --upgrade pip >nul 2>&1
python -m pip install pandas openpyxl pillow >nul 2>&1

echo 🚀 Iniciando o script...
start "" pythonw main.py

exit
