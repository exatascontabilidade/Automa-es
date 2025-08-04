@echo off
setlocal

echo [🛠] Verificando instalação do Python...
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [⚠️] Python não encontrado. Baixando e instalando Python 3.11.9...
    curl -o python-installer.exe https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe
    python-installer.exe /quiet InstallAllUsers=1 PrependPath=1 Include_pip=1
    set PATH=%PATH%;C:\Program Files\Python311\Scripts;C:\Program Files\Python311\
    echo [✅] Python instalado.
) else (
    echo [✅] Python já está instalado.
)

echo.
echo [📦] Instalando dependências do requirements.txt via pip...
pip install --upgrade pip
pip install -r requirements.txt

echo.
echo [✅] Instalação finalizada com sucesso!
pause