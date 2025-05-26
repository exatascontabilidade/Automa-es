@echo off
echo =====================================
echo Instalando dependências do projeto...
echo =====================================

REM Instala bibliotecas padrão
pip install selenium
pip install webdriver-manager
pip install requests
pip install beautifulsoup4
pip install pygetwindow
pip install pyautoit

echo =====================================
echo Todas as dependências foram instaladas com sucesso!
echo Pressione qualquer tecla para sair.
pause > nul
